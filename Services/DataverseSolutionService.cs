using System;
using System.Collections.Generic;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ModelSolutionInfo = D365SolutionComparer.Models.SolutionInfo;

namespace D365SolutionComparer.Services
{
    public class DataverseSolutionService
    {
        public List<ModelSolutionInfo> GetSolutions(IOrganizationService service)
        {
            if (service == null)
            {
                throw new ArgumentNullException(nameof(service));
            }

            var results = new List<ModelSolutionInfo>();

            var query = new QueryExpression("solution")
            {
                ColumnSet = new ColumnSet(
                    "uniquename",
                    "friendlyname",
                    "version",
                    "publisherid",
                    "ismanaged"),
                NoLock = true
            };

            query.Criteria.AddCondition("isvisible", ConditionOperator.Equal, true);

            // Publish an inventory only after terminal retrieval; preserve the existing visibility rule.
            var response = new DataversePagedReader(service).ReadAll(query, "solutionid", CancellationToken.None);

            foreach (var entity in response)
            {
                var publisherName = string.Empty;

                var publisherRef = entity.GetAttributeValue<EntityReference>("publisherid");
                if (publisherRef != null)
                {
                    publisherName = publisherRef.Name ?? string.Empty;
                }

                bool? isManagedValue = null;
                if (entity.Attributes.Contains("ismanaged"))
                {
                    isManagedValue = entity.GetAttributeValue<bool>("ismanaged");
                }

                results.Add(new ModelSolutionInfo
                {
                    UniqueName = entity.GetAttributeValue<string>("uniquename") ?? string.Empty,
                    DisplayName = entity.GetAttributeValue<string>("friendlyname") ?? string.Empty,
                    Version = entity.GetAttributeValue<string>("version") ?? string.Empty,
                    Publisher = publisherName,
                    IsManaged = isManagedValue
                });
            }

            return results;
        }
    }
}