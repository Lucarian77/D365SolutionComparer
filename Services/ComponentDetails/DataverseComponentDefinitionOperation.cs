using System;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.ComponentDetails
{
    /// <summary>
    /// Production bulk path. Environment verification, membership retrieval, identity
    /// resolution, and definition retrieval share one operation context and metadata cache.
    /// </summary>
    internal sealed class DataverseComponentDefinitionOperation
    {
        private readonly DataverseSolutionMembershipReader membershipReader =
            new DataverseSolutionMembershipReader();
        private readonly DataverseComponentIdentityResolver identityResolver =
            new DataverseComponentIdentityResolver();
        private readonly DataverseComponentDefinitionReader definitionReader =
            new DataverseComponentDefinitionReader();
        internal DataverseComponentDefinitionOperation() { }
        private readonly object invocationIdentity = new object();

        public ComponentDefinitionSnapshot ReadAndResolve(IOrganizationService service,
            SolutionIdentity solution, CancellationToken cancellationToken,
            Action<RetrievalProgress> progress = null, DataverseRequestCounter requestCounter = null)
        {
            if (solution == null) throw new ArgumentNullException(nameof(solution));
            var context = new DataverseReadContext(service, solution.Environment,
                cancellationToken, requestCounter);
            var membership = membershipReader.Read(context, solution, cancellationToken, progress);
            var resolved = identityResolver.ResolveSnapshot(context, membership, cancellationToken);
            return definitionReader.Read(context, resolved, cancellationToken);
        }

        /// <summary>
        /// Live UI path. Membership, identity resolution, and definition retrieval share one
        /// verified environment context, request counter, and operation-scoped metadata cache.
        /// </summary>
        public ComponentDefinitionSnapshot ReadAndResolve(IOrganizationService service,
            string environmentDisplayName, string solutionUniqueName,
            CancellationToken cancellationToken, Action<MembershipOperationProgress> progress,
            DataverseRequestCounter requestCounter = null)
        {
            return ReadAndResolveWithCache(service, environmentDisplayName, solutionUniqueName,
                cancellationToken, progress, requestCounter).Definitions;
        }

        internal DefinitionReadWithCacheResult ReadAndResolveWithCache(IOrganizationService service,
            string environmentDisplayName, string solutionUniqueName,
            CancellationToken cancellationToken, Action<MembershipOperationProgress> progress,
            DataverseRequestCounter requestCounter = null)
            => ReadAndResolveCore(service, environmentDisplayName, solutionUniqueName, cancellationToken,
                progress, requestCounter, null);

        internal DefinitionReadWithCacheResult ReadAndResolveWithCache(IOrganizationService service,
            string environmentDisplayName, string solutionUniqueName, CancellationToken cancellationToken,
            Action<MembershipOperationProgress> progress, DataverseRequestCounter requestCounter, DisplayEnvironmentSide side)
            => ReadAndResolveCore(service, environmentDisplayName, solutionUniqueName, cancellationToken,
                progress, requestCounter, side);

        private DefinitionReadWithCacheResult ReadAndResolveCore(IOrganizationService service,
            string environmentDisplayName, string solutionUniqueName, CancellationToken cancellationToken,
            Action<MembershipOperationProgress> progress, DataverseRequestCounter requestCounter, DisplayEnvironmentSide? side)
        {
            if (side.HasValue && !Enum.IsDefined(typeof(DisplayEnvironmentSide), side.Value))
                throw new ArgumentOutOfRangeException(nameof(side));
            var readOwner = new object();
            if (string.IsNullOrWhiteSpace(solutionUniqueName))
                throw new ArgumentException("A solution Unique Name is required.",
                    nameof(solutionUniqueName));

            Report(progress, MembershipOperationStage.ValidatingEnvironment,
                "Validating the Dataverse environment...");
            var context = new DataverseReadContext(service, environmentDisplayName,
                cancellationToken, requestCounter);
            Report(progress, MembershipOperationStage.ReadingMembership,
                "Reading solution membership...");
            var membership = membershipReader.Read(context, context.Environment,
                solutionUniqueName, cancellationToken, item =>
                {
                    Report(progress, MembershipOperationStage.ReadingMembership,
                        "Reading solution membership: " + item.RecordsRetrieved +
                        " record(s) from " + item.PagesRetrieved + " page(s)...",
                        item.PagesRetrieved, item.RecordsRetrieved);
                });
            if (membership.State == Models.Membership.MembershipSnapshotState.Complete)
                Report(progress, MembershipOperationStage.ResolvingIdentities,
                    "Resolving portable component identities...");
            AppSettingParentEvidenceEmission emission = null;
            if (side.HasValue && membership.State == MembershipSnapshotState.Complete)
            {
                try { emission = new AppSettingParentEvidenceEmission(membership, side.Value, invocationIdentity, readOwner); }
                catch (Exception) { cancellationToken.ThrowIfCancellationRequested(); }
            }
            var resolved = identityResolver.ResolveSnapshot(context, membership,
                cancellationToken, emission);
            if (resolved.State == Models.Membership.MembershipSnapshotState.Complete)
                Report(progress, MembershipOperationStage.ReadingDefinitions,
                    "Reading component definitions...");
            var definitions = definitionReader.Read(context, resolved, cancellationToken);
            Report(progress, MembershipOperationStage.Completed,
                "Membership and component definition comparison data completed.");
            AppSettingParentCapture capture = null;
            try { capture = emission?.SealAgainstFinalSnapshot(definitions.Membership, cancellationToken); }
            catch (Exception) { cancellationToken.ThrowIfCancellationRequested(); }
            return new DefinitionReadWithCacheResult(definitions, context.MetadataCache,
                invocationIdentity, context.Environment.OrganizationId, solutionUniqueName, readOwner, side, capture);
        }

        /// <summary>Worker-local handoff only. No context or service escapes the read.</summary>
        internal sealed class DefinitionReadWithCacheResult
        {
            private readonly object invocationIdentity;
            private readonly Guid organizationId;
            private readonly string solutionUniqueName;

            internal DefinitionReadWithCacheResult(ComponentDefinitionSnapshot definitions,
                DataverseComponentMetadataCache cache, object invocationIdentity,
                Guid organizationId, string solutionUniqueName, object readOwner = null,
                DisplayEnvironmentSide? side = null, AppSettingParentCapture parentCapture = null)
            {
                Definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
                Cache = cache ?? throw new ArgumentNullException(nameof(cache));
                this.invocationIdentity = invocationIdentity;
                this.organizationId = organizationId;
                this.solutionUniqueName = solutionUniqueName;
                ReadOwner = readOwner; Side = side; ParentCapture = parentCapture;
            }

            internal ComponentDefinitionSnapshot Definitions { get; }
            internal DataverseComponentMetadataCache Cache { get; }
            internal object ReadOwner { get; }
            internal DisplayEnvironmentSide? Side { get; }
            internal AppSettingParentCapture ParentCapture { get; }
            internal object ComparisonInvocationOwner => invocationIdentity;
            internal bool EvidenceBelongsTo(DataverseComponentDefinitionOperation operation,
                MembershipSnapshot expectedSnapshot, DisplayEnvironmentSide expectedSide) =>
                BelongsTo(operation, expectedSnapshot) && Side == expectedSide && ParentCapture != null &&
                ParentCapture.BelongsTo(expectedSnapshot, expectedSide, operation.invocationIdentity, ReadOwner);

            internal bool BelongsTo(DataverseComponentDefinitionOperation operation,
                MembershipSnapshot expectedSnapshot)
            {
                return operation != null && expectedSnapshot != null &&
                    ReferenceEquals(invocationIdentity, operation.invocationIdentity) &&
                    ReferenceEquals(Definitions.Membership, expectedSnapshot) &&
                    organizationId == expectedSnapshot.Environment.OrganizationId &&
                    string.Equals(solutionUniqueName, expectedSnapshot.SolutionUniqueName,
                        StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void Report(Action<MembershipOperationProgress> progress,
            MembershipOperationStage stage, string message, int pagesRetrieved = 0,
            int recordsRetrieved = 0)
        {
            progress?.Invoke(new MembershipOperationProgress(stage, message,
                pagesRetrieved, recordsRetrieved));
        }
    }
}
