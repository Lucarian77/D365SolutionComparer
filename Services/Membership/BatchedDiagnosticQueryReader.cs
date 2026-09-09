using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Performs operation-scoped, bounded diagnostic reads and validates correlation without
    /// applying component-family identity or classification policy.
    /// </summary>
    internal sealed class BatchedDiagnosticQueryReader
    {
        internal const int BatchSize = 200;
        private readonly DataverseReadContext context;

        public BatchedDiagnosticQueryReader(DataverseReadContext context)
        {
            this.context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public DiagnosticQueryResult Read(IReadOnlyList<SolutionComponentRecord> records,
            string entityName, string primaryIdAttribute, IReadOnlyList<string> columns,
            string incompleteResultDiagnostic, string conflictingPrimaryKeyDiagnostic,
            Func<FaultException, string> faultDiagnostic, CancellationToken cancellationToken)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (string.IsNullOrWhiteSpace(entityName)) throw new ArgumentException(
                "An entity logical name is required.", nameof(entityName));
            if (string.IsNullOrWhiteSpace(primaryIdAttribute)) throw new ArgumentException(
                "A primary-key attribute is required.", nameof(primaryIdAttribute));
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            if (faultDiagnostic == null) throw new ArgumentNullException(nameof(faultDiagnostic));

            var objectIds = CollectDistinctNonemptyObjectIds(records);
            var returnedById = objectIds.ToDictionary(item => item, item => new List<Entity>());
            var failures = new Dictionary<Guid, string>();
            var unassociatedRows = objectIds.ToDictionary(item => item, item => new List<Entity>());
            int returnedCount = 0;
            bool countsAvailable = true;

            foreach (var batch in Batch(objectIds))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var query = new QueryExpression(entityName)
                    {
                        ColumnSet = new ColumnSet(columns.ToArray())
                    };
                    query.Criteria.AddCondition(new ConditionExpression(primaryIdAttribute,
                        ConditionOperator.In, batch.Select(item => (object)item).ToArray()));
                    var rows = context.Query(query);
                    cancellationToken.ThrowIfCancellationRequested();
                    returnedCount += rows.Entities.Count;
                    bool invalidResponse = false;
                    foreach (var row in rows.Entities)
                    {
                        Guid primaryId;
                        if (string.Equals(row.LogicalName, entityName, StringComparison.OrdinalIgnoreCase) &&
                            TryReadGuid(row, primaryIdAttribute, out primaryId) &&
                            batch.Contains(primaryId) && (row.Id == Guid.Empty || row.Id == primaryId))
                        {
                            returnedById[primaryId].Add(row);
                            continue;
                        }

                        invalidResponse = true;
                        Guid conflictingId;
                        var affectedIds = TryReadGuid(row, primaryIdAttribute, out conflictingId) &&
                            batch.Contains(conflictingId) ? new[] { conflictingId } : batch;
                        foreach (var objectId in affectedIds) unassociatedRows[objectId].Add(row);
                    }

                    if (rows.MoreRecords)
                    {
                        countsAvailable = false;
                        SetFailures(batch, incompleteResultDiagnostic, failures);
                    }
                    else if (invalidResponse)
                    {
                        countsAvailable = false;
                        SetFailures(batch, conflictingPrimaryKeyDiagnostic, failures);
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (FaultException ex)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    countsAvailable = false;
                    SetFailures(batch, faultDiagnostic(ex), failures);
                }
            }

            return new DiagnosticQueryResult(objectIds, returnedById, failures, unassociatedRows,
                returnedCount, countsAvailable);
        }

        private static IReadOnlyList<Guid> CollectDistinctNonemptyObjectIds(
            IEnumerable<SolutionComponentRecord> records)
        {
            return records.Where(item => item.ObjectId.HasValue && item.ObjectId.Value != Guid.Empty)
                .Select(item => item.ObjectId.Value).Distinct().OrderBy(item => item).ToList();
        }

        private static IEnumerable<IReadOnlyList<Guid>> Batch(IReadOnlyList<Guid> items)
        {
            for (int offset = 0; offset < items.Count; offset += BatchSize)
                yield return items.Skip(offset).Take(Math.Min(BatchSize, items.Count - offset)).ToList();
        }

        private static bool TryReadGuid(Entity entity, string attributeName, out Guid value)
        {
            object raw;
            if (entity.Attributes.TryGetValue(attributeName, out raw) && raw is Guid)
            {
                value = (Guid)raw;
                return value != Guid.Empty;
            }
            value = Guid.Empty;
            return false;
        }

        private static void SetFailures(IEnumerable<Guid> objectIds, string diagnostic,
            IDictionary<Guid, string> failures)
        {
            foreach (var objectId in objectIds) failures[objectId] = diagnostic;
        }
    }

    internal sealed class DiagnosticQueryResult
    {
        private readonly IDictionary<Guid, List<Entity>> returnedById;
        private readonly IDictionary<Guid, string> failures;
        private readonly IDictionary<Guid, List<Entity>> unassociatedRowsById;

        public DiagnosticQueryResult(IReadOnlyList<Guid> objectIds,
            IDictionary<Guid, List<Entity>> returnedById, IDictionary<Guid, string> failures,
            IDictionary<Guid, List<Entity>> unassociatedRowsById, int returnedRowCount,
            bool countsAvailable)
        {
            ObjectIds = objectIds;
            this.returnedById = returnedById;
            this.failures = failures;
            this.unassociatedRowsById = unassociatedRowsById;
            ReturnedRowCount = returnedRowCount;
            CountsAvailable = countsAvailable;
        }

        public IReadOnlyList<Guid> ObjectIds { get; }
        public IEnumerable<Entity> ReturnedRows => ObjectIds.SelectMany(item => returnedById[item]);
        public int ReturnedRowCount { get; }
        public bool CountsAvailable { get; }

        public DiagnosticRowCorrelation GetCorrelation(Guid objectId)
        {
            List<Entity> rows;
            List<Entity> unassociatedRows;
            string failure;
            if (!returnedById.TryGetValue(objectId, out rows) ||
                !unassociatedRowsById.TryGetValue(objectId, out unassociatedRows))
                throw new ArgumentOutOfRangeException(nameof(objectId),
                    "The object ID was not part of this diagnostic query.");
            var status = failures.TryGetValue(objectId, out failure)
                ? DiagnosticCorrelationStatus.Failed
                : rows.Count == 0 ? DiagnosticCorrelationStatus.Missing
                : rows.Count == 1 ? DiagnosticCorrelationStatus.Unique
                : DiagnosticCorrelationStatus.Duplicate;
            return new DiagnosticRowCorrelation(objectId, status, rows, failure, unassociatedRows);
        }
    }

    internal enum DiagnosticCorrelationStatus
    {
        Failed,
        Missing,
        Unique,
        Duplicate
    }

    internal sealed class DiagnosticRowCorrelation
    {
        public DiagnosticRowCorrelation(Guid objectId, DiagnosticCorrelationStatus status,
            IReadOnlyList<Entity> rows, string failure, IReadOnlyList<Entity> unassociatedRows)
        {
            ObjectId = objectId;
            Status = status;
            Rows = rows;
            Failure = failure;
            UnassociatedRows = unassociatedRows;
        }

        public Guid ObjectId { get; }
        public DiagnosticCorrelationStatus Status { get; }
        public IReadOnlyList<Entity> Rows { get; }
        public string Failure { get; }
        public IReadOnlyList<Entity> UnassociatedRows { get; }
    }
}
