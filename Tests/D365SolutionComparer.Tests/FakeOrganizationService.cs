using System;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    internal sealed class FakeOrganizationService : IOrganizationService
    {
        public Func<QueryExpression, EntityCollection> RetrievePage { get; set; }
        public Func<OrganizationRequest, OrganizationResponse> ExecuteRequest { get; set; }
        public int ExecuteCalls { get; private set; }
        public int Calls { get; private set; }
        public int WriteCalls { get; private set; }
        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            Calls++;
            return RetrievePage((QueryExpression)query);
        }
        public Guid Create(Entity entity) { WriteCalls++; throw new NotSupportedException(); }
        public void Update(Entity entity) { WriteCalls++; throw new NotSupportedException(); }
        public void Delete(string entityName, Guid id) { WriteCalls++; throw new NotSupportedException(); }
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) => throw new NotSupportedException();
        public OrganizationResponse Execute(OrganizationRequest request)
        {
            ExecuteCalls++;
            var metadata = request as RetrieveMetadataChangesRequest;
            if (metadata != null && (HasConditions(metadata.Query.AttributeQuery?.Criteria) ||
                HasConditions(metadata.Query.RelationshipQuery?.Criteria)))
                throw new FaultException("Unable to evaluate query: child metadata conditions are not supported by the production bulk reader.");
            return ExecuteRequest != null ? ExecuteRequest(request) : throw new NotSupportedException();
        }
        private static bool HasConditions(MetadataFilterExpression filter) => filter != null &&
            (filter.Conditions.Count != 0 || filter.Filters.Any(HasConditions));

        public void Associate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) { WriteCalls++; throw new NotSupportedException(); }
        public void Disassociate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) { WriteCalls++; throw new NotSupportedException(); }
    }
}
