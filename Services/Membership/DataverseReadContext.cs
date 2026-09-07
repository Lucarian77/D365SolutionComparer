using System;
using System.Threading;
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

        public DataverseReadContext(IOrganizationService service, EnvironmentIdentity environment, CancellationToken cancellationToken)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            if (environment == null) throw new ArgumentNullException(nameof(environment));
            this.cancellationToken = cancellationToken;
            var response = Execute(new WhoAmIRequest()) as WhoAmIResponse;
            if (response == null || response.OrganizationId != environment.OrganizationId)
                throw new InvalidOperationException("The service does not identify the captured environment.");
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var response = service.Execute(request);
            cancellationToken.ThrowIfCancellationRequested();
            return response ?? throw new InvalidOperationException("Dataverse returned no response.");
        }

        public EntityCollection Query(QueryExpression query)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var response = service.RetrieveMultiple(query);
            cancellationToken.ThrowIfCancellationRequested();
            return response ?? throw new InvalidOperationException("Dataverse returned no query response.");
        }
    }
}
