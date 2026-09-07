using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Infrastructure
{
    /// <summary>Cookie-based retrieval for standard tables, intended initially for solutioncomponent.</summary>
    public sealed class DataversePagedReader
    {
        private readonly IOrganizationService service;

        public DataversePagedReader(IOrganizationService service)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
        }

        /// <summary>
        /// Reads from page one without mutating the supplied query. primaryIdAttribute must be the
        /// root table's actual primary key. It is appended ascending only if not already ordered.
        /// TopCount/distinct/joined queries and multi-page queries without cookies are deliberately rejected.
        /// Progress executes on the calling thread; a future UI caller must use XrmToolBox WorkAsync.
        /// Returns only on complete success. Faults, callback errors and cancellation propagate.
        /// Cancellation is cooperative between SDK calls, not an abort of an in-flight request.
        /// </summary>
        public IReadOnlyList<Entity> ReadAll(QueryExpression query, string primaryIdAttribute,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null, int pageSize = 5000)
        {
            if (query == null) throw new ArgumentNullException(nameof(query));
            if (string.IsNullOrWhiteSpace(primaryIdAttribute) ||
                primaryIdAttribute.Any(c => !(char.IsLetterOrDigit(c) || c == '_')))
                throw new ArgumentException("Supply the root table primary key logical name.", nameof(primaryIdAttribute));
            if (pageSize < 1 || pageSize > 5000) throw new ArgumentOutOfRangeException(nameof(pageSize));
            if (query.TopCount.HasValue || query.Distinct)
                throw new ArgumentException("TopCount and distinct queries are not supported by this reader.", nameof(query));
            if (query.LinkEntities.Count != 0)
                throw new ArgumentException("Joined queries may duplicate root IDs; this reader requires a root-table query.", nameof(query));
            cancellationToken.ThrowIfCancellationRequested();

            QueryExpression request;
            var serializer = new DataContractSerializer(typeof(QueryExpression));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, query);
                stream.Position = 0;
                request = (QueryExpression)serializer.ReadObject(stream);
            }
            // Preserve an explicit primary-key order, including its direction and position.
            if (!request.Orders.Any(order => string.Equals(order.AttributeName, primaryIdAttribute, StringComparison.OrdinalIgnoreCase)))
            {
                request.AddOrder(primaryIdAttribute, OrderType.Ascending);
            }
            request.PageInfo = new PagingInfo { Count = pageSize, PageNumber = 1 };
            var records = new List<Entity>();
            var seenCookies = new HashSet<string>(StringComparer.Ordinal);

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var response = service.RetrieveMultiple(request);
                cancellationToken.ThrowIfCancellationRequested();
                if (response == null) throw new InvalidOperationException("Dataverse returned no page response.");
                records.AddRange(response.Entities);
                progress?.Invoke(new RetrievalProgress(request.PageInfo.PageNumber, records.Count));
                cancellationToken.ThrowIfCancellationRequested();
                if (!response.MoreRecords) return records.AsReadOnly();
                if (string.IsNullOrWhiteSpace(response.PagingCookie))
                    throw new NotSupportedException("Dataverse did not return a paging cookie. Complete retrieval cannot be guaranteed by this reader.");
                if (response.Entities.Count == 0 || !seenCookies.Add(response.PagingCookie))
                    throw new InvalidOperationException("Dataverse paging did not advance.");
                request.PageInfo.PagingCookie = response.PagingCookie;
                request.PageInfo.PageNumber = checked(request.PageInfo.PageNumber + 1);
            }
        }
    }
}
