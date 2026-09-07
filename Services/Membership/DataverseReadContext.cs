using System;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>Operation-local service validation; no credentials, global cache or client ownership.</summary>
    internal sealed class DataverseReadContext
    {
        private readonly IOrganizationService service;
        private readonly CancellationToken cancellationToken;
        private readonly DataverseRequestCounter requestCounter;

        public DataverseReadContext(IOrganizationService service, EnvironmentIdentity environment, CancellationToken cancellationToken,
            DataverseRequestCounter requestCounter = null)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            if (environment == null) throw new ArgumentNullException(nameof(environment));
            Environment = environment;
            this.cancellationToken = cancellationToken;
            this.requestCounter = requestCounter;
            var response = Execute(new WhoAmIRequest()) as WhoAmIResponse;
            if (response == null || response.OrganizationId != environment.OrganizationId)
                throw new InvalidOperationException("The service does not identify the captured environment.");
        }

        public EnvironmentIdentity Environment { get; }
        public IOrganizationService Service => requestCounter == null
            ? service : new InstrumentedOrganizationService(service, requestCounter);

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            cancellationToken.ThrowIfCancellationRequested();
            requestCounter?.RecordExecute(request.RequestName);
            var response = service.Execute(request);
            cancellationToken.ThrowIfCancellationRequested();
            return response ?? throw new InvalidOperationException("Dataverse returned no response.");
        }

        public EntityCollection Query(QueryExpression query)
        {
            cancellationToken.ThrowIfCancellationRequested();
            requestCounter?.RecordQuery(query?.EntityName);
            var response = service.RetrieveMultiple(query);
            cancellationToken.ThrowIfCancellationRequested();
            return response ?? throw new InvalidOperationException("Dataverse returned no query response.");
        }


        private sealed class InstrumentedOrganizationService : IOrganizationService
        {
            private readonly IOrganizationService inner;
            private readonly DataverseRequestCounter counter;

            public InstrumentedOrganizationService(IOrganizationService inner, DataverseRequestCounter counter)
            {
                this.inner = inner;
                this.counter = counter;
            }

            public EntityCollection RetrieveMultiple(QueryBase query)
            {
                var expression = query as QueryExpression;
                counter.RecordQuery(expression?.EntityName);
                return inner.RetrieveMultiple(query);
            }

            public OrganizationResponse Execute(OrganizationRequest request)
            {
                counter.RecordExecute(request?.RequestName);
                return inner.Execute(request);
            }

            public Guid Create(Entity entity) => inner.Create(entity);
            public void Update(Entity entity) => inner.Update(entity);
            public void Delete(string entityName, Guid id) => inner.Delete(entityName, id);
            public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) => inner.Retrieve(entityName, id, columnSet);
            public void Associate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) =>
                inner.Associate(entityName, id, relationship, relatedEntities);
            public void Disassociate(string entityName, Guid id, Relationship relationship, EntityReferenceCollection relatedEntities) =>
                inner.Disassociate(entityName, id, relationship, relatedEntities);
        }
    }
}
