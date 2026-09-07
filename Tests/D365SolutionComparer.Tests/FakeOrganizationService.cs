using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    internal sealed class FakeOrganizationService : IOrganizationService
    {
        public Func<QueryExpression, EntityCollection> RetrievePage { get; set; }
        public int Calls { get; private set; }
        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            Calls++;
            return RetrievePage((QueryExpression)query);
        }
        public Guid Create(Entity entity) => throw new NotSupportedException();
        public void Update(Entity entity) => throw new NotSupportedException();
        public void Delete(string entityName, Guid id) => throw new NotSupportedException();
        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) => throw new NotSupportedException();
        public OrganizationResponse Execute(OrganizationRequest request) => throw new NotSupportedException();
        public void Associate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) => throw new NotSupportedException();
        public void Disassociate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) => throw new NotSupportedException();
    }
}
