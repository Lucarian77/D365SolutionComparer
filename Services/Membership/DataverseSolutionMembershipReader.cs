using System;
using System.Collections.Generic;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Contracts;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Services.Membership
{
    public sealed class DataverseSolutionMembershipReader : ISolutionMembershipReader
    {
        public MembershipSnapshot Read(IOrganizationService service, SolutionIdentity solution,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null)
        {
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            var context = new DataverseReadContext(service, solution.Environment, cancellationToken);
            return ReadCore(context, solution.Environment, solution.UniqueName, solution.SolutionId, cancellationToken, progress);
        }

        /// <summary>Look up the same Unique Name independently on each side, including absent solutions.</summary>
        public MembershipSnapshot Read(IOrganizationService service, EnvironmentIdentity environment, string uniqueName,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null)
        {
            ValidateSelection(environment, uniqueName);
            var context = new DataverseReadContext(service, environment, cancellationToken);
            return ReadCore(context, environment, uniqueName, null, cancellationToken, progress);
        }

        /// <summary>
        /// Explicit state-oriented alternative: errors become Unavailable, never Complete or Absent.
        /// Cancellation still propagates. The strict Read methods always propagate retrieval errors.
        /// </summary>
        public MembershipSnapshot ReadOrUnavailable(IOrganizationService service, EnvironmentIdentity environment, string uniqueName,
            CancellationToken cancellationToken)
        {
            ValidateSelection(environment, uniqueName);
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                return Read(service, environment, uniqueName, cancellationToken);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return MembershipSnapshot.Unavailable(environment, uniqueName, DateTimeOffset.UtcNow,
                    ex.GetType().Name + ": " + ex.Message);
            }
        }

        internal MembershipSnapshot Read(DataverseReadContext context, SolutionIdentity solution,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            EnsureEnvironment(context, solution.Environment);
            return ReadCore(context, solution.Environment, solution.UniqueName, solution.SolutionId, cancellationToken, progress);
        }

        internal MembershipSnapshot Read(DataverseReadContext context, EnvironmentIdentity environment, string uniqueName,
            CancellationToken cancellationToken, Action<RetrievalProgress> progress = null)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            ValidateSelection(environment, uniqueName);
            EnsureEnvironment(context, environment);
            return ReadCore(context, environment, uniqueName, null, cancellationToken, progress);
        }

        private static MembershipSnapshot ReadCore(DataverseReadContext context, EnvironmentIdentity environment, string uniqueName,
            Guid? expectedSolutionId, CancellationToken cancellationToken, Action<RetrievalProgress> progress)
        {
            ValidateSelection(environment, uniqueName);
            // TopCount=2 distinguishes a unique selection from ambiguity; it is not an inventory query.
            var lookup = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "uniquename"), TopCount = 2 };
            lookup.Criteria.AddCondition("uniquename", ConditionOperator.Equal, uniqueName);
            var solutions = context.Query(lookup);
            if (solutions.MoreRecords || solutions.Entities.Count > 1)
                throw new InvalidOperationException("The solution Unique Name lookup is ambiguous.");
            if (solutions.Entities.Count == 0)
                return MembershipSnapshot.Absent(environment, uniqueName, DateTimeOffset.UtcNow);
            var row = solutions.Entities[0];
            if (row.Id == Guid.Empty || !string.Equals(row.GetAttributeValue<string>("uniquename"), uniqueName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Dataverse returned an invalid solution identity.");
            if (expectedSolutionId.HasValue && expectedSolutionId.Value != row.Id)
                throw new InvalidOperationException("The selected solution ID has changed. Reload the solution selection.");
            var solution = new SolutionIdentity(environment, row.Id, uniqueName);
            var query = new QueryExpression("solutioncomponent")
            {
                ColumnSet = new ColumnSet("solutioncomponentid", "componenttype", "objectid", "rootcomponentbehavior",
                    "rootsolutioncomponentid", "ismetadata", "solutionid")
            };
            query.Criteria.AddCondition("solutionid", ConditionOperator.Equal, solution.SolutionId);
            var rows = new DataversePagedReader(context.Service).ReadAll(query, "solutioncomponentid", cancellationToken, progress);
            var components = new List<ComponentIdentity>();
            foreach (var entity in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var type = entity.GetAttributeValue<OptionSetValue>("componenttype");
                var owner = entity.GetAttributeValue<EntityReference>("solutionid");
                if (entity.Id == Guid.Empty || type == null || (owner != null && owner.Id != solution.SolutionId))
                    throw new InvalidOperationException("A solutioncomponent record has missing identifiers or belongs to a different solution.");
                var record = new SolutionComponentRecord(entity.Id, type.Value, entity.GetAttributeValue<Guid?>("objectid"),
                    entity.GetAttributeValue<OptionSetValue>("rootcomponentbehavior")?.Value,
                    entity.GetAttributeValue<Guid?>("rootsolutioncomponentid"), entity.GetAttributeValue<bool?>("ismetadata"));
                components.Add(new ComponentIdentity(record, IdentityResolutionStatus.Unresolved, diagnostic: "Identity resolution has not run."));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return MembershipSnapshot.Complete(solution, components, DateTimeOffset.UtcNow);
        }

        private static void EnsureEnvironment(DataverseReadContext context, EnvironmentIdentity environment)
        {
            if (context.Environment.OrganizationId != environment.OrganizationId)
                throw new InvalidOperationException("The verified context belongs to a different environment.");
        }

        private static void ValidateSelection(EnvironmentIdentity environment, string uniqueName)
        {
            if (environment == null) throw new ArgumentNullException(nameof(environment));
            if (string.IsNullOrWhiteSpace(uniqueName)) throw new ArgumentException("A solution Unique Name is required.", nameof(uniqueName));
        }
    }
}
