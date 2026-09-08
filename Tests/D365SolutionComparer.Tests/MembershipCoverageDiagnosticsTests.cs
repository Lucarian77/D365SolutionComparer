using System;
using System.Linq;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class MembershipCoverageDiagnosticsTests
    {
        private readonly MembershipCoverageDiagnosticsBuilder builder =
            new MembershipCoverageDiagnosticsBuilder();

        [TestMethod]
        public void CompleteSemanticKindReportsCountsAndCompleteCoverage()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.name"),
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.number")));

            var column = Kind(diagnostics, ComponentSemanticKinds.Column);
            Assert.AreEqual(2, column.TotalCandidates);
            Assert.AreEqual(2, column.Resolved);
            Assert.AreEqual(0, column.Unsupported);
            Assert.AreEqual(0, column.Unresolved);
            Assert.AreEqual(0, column.Ambiguous);
            Assert.AreEqual(MembershipCoverageStatus.Complete, column.CoverageStatus);
            Assert.AreEqual(MembershipCoverageBucketType.SemanticKind, column.BucketType);
        }

        [TestMethod]
        public void UnresolvedCandidateMakesOnlyItsSemanticKindIncomplete()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.name"),
                Candidate(IdentityResolutionStatus.Unresolved, 2, ComponentSemanticKinds.Column,
                    diagnostic: "Column metadata missing"),
                Candidate(IdentityResolutionStatus.Resolved, 61, ComponentSemanticKinds.WebResource,
                    "new_/script.js")));

            var column = Kind(diagnostics, ComponentSemanticKinds.Column);
            Assert.AreEqual(2, column.TotalCandidates);
            Assert.AreEqual(1, column.Resolved);
            Assert.AreEqual(1, column.Unresolved);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, column.CoverageStatus);
            Assert.AreEqual(MembershipCoverageStatus.Complete,
                Kind(diagnostics, ComponentSemanticKinds.WebResource).CoverageStatus);
        }

        [TestMethod]
        public void UnsupportedRelationshipFamilyCandidateIsASameKindBlocker()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Unsupported, 3, diagnostic: "Unsupported relationship role")));

            var relationship = Kind(diagnostics, ComponentSemanticKinds.Relationship);
            Assert.AreEqual(MembershipCoverageBucketType.SemanticKind, relationship.BucketType);
            Assert.AreEqual(1, relationship.Unsupported);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, relationship.CoverageStatus);
        }

        [TestMethod]
        public void KnownUnsupportedTypeUsesAnIsolatedBucketWithoutBlockingColumns()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Unsupported, 9, diagnostic: "Option set unsupported"),
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.name")));

            var isolated = Kind(diagnostics, "unsupported:componenttype:9");
            Assert.AreEqual(MembershipCoverageBucketType.KnownUnsupportedIsolatedType, isolated.BucketType);
            Assert.AreEqual(1, isolated.Unsupported);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, isolated.CoverageStatus);
            Assert.AreEqual(MembershipCoverageStatus.Complete,
                Kind(diagnostics, ComponentSemanticKinds.Column).CoverageStatus);
            Assert.AreEqual(0, diagnostics.BroadRawComponentTypes.Count);
        }

        [TestMethod]
        public void UnclassifiableCandidateIsCountedSeparatelyAndBlocksEveryKind()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Unsupported, 98765, diagnostic: "Unknown type"),
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.name")));

            Assert.IsTrue(diagnostics.HasBroadUnclassifiableBlockers);
            Assert.AreEqual(1, diagnostics.BroadUnclassifiable.TotalCandidates);
            Assert.AreEqual(1, diagnostics.BroadUnclassifiable.Unsupported);
            Assert.AreEqual(MembershipCoverageBucketType.BroadUnclassifiable,
                diagnostics.BroadUnclassifiable.BucketType);
            Assert.IsTrue(diagnostics.SemanticKinds.All(item =>
                item.CoverageStatus == MembershipCoverageStatus.Incomplete));
            Assert.AreEqual(1, diagnostics.BroadRawComponentTypes.Count);
            Assert.AreEqual(98765, diagnostics.BroadRawComponentTypes[0].ComponentType);
            Assert.AreEqual(1, diagnostics.BroadRawComponentTypes[0].Count);
        }

        [TestMethod]
        public void BroadBlockersGroupByRawTypeWithoutMergingSharedDiagnosticText()
        {
            const string exact = "Unsupported component type.";
            var candidates = new[]
            {
                Candidate(IdentityResolutionStatus.Unsupported, 80, diagnostic: exact),
                Candidate(IdentityResolutionStatus.Unsupported, 80, diagnostic: exact),
                Candidate(IdentityResolutionStatus.Unsupported, 511, diagnostic: exact)
            };
            var diagnostics = builder.Build(Snapshot(candidates));

            Assert.AreEqual(3, diagnostics.BroadUnclassifiable.TotalCandidates);
            Assert.AreEqual(2, diagnostics.BroadRawComponentTypes.Count);
            var type80 = diagnostics.BroadRawComponentTypes.Single(group => group.ComponentType == 80);
            var type511 = diagnostics.BroadRawComponentTypes.Single(group => group.ComponentType == 511);
            Assert.AreEqual(2, type80.Count);
            Assert.AreEqual(1, type511.Count);
            Assert.AreEqual(2, type80.DiagnosticGroups.Single().Count);
            Assert.AreEqual(1, type511.DiagnosticGroups.Single().Count);
            Assert.AreEqual(exact, type80.DiagnosticGroups.Single().Diagnostic);
            Assert.AreEqual(exact, type511.DiagnosticGroups.Single().Diagnostic);
            Assert.AreEqual(2, type80.Evidence.Count);
            Assert.AreEqual(1, type511.Evidence.Count);
            Assert.IsTrue(type80.Evidence.Concat(type511.Evidence).All(item =>
                item.ResolutionStatus == IdentityResolutionStatus.Unsupported && item.Diagnostic == exact));
            CollectionAssert.AreEquivalent(candidates.Where(item => item.Record.ComponentType == 80)
                    .Select(item => item.Record.SolutionComponentId).ToArray(),
                type80.Evidence.Select(item => item.SolutionComponentId).ToArray());
            CollectionAssert.AreEquivalent(candidates.Where(item => item.Record.ComponentType == 80)
                    .Select(item => item.Record.ObjectId.Value).ToArray(),
                type80.Evidence.Select(item => item.ObjectId.Value).ToArray());
            CollectionAssert.AreEquivalent(candidates.Where(item => item.Record.ComponentType == 511)
                    .Select(item => item.Record.SolutionComponentId).ToArray(),
                type511.Evidence.Select(item => item.SolutionComponentId).ToArray());
            CollectionAssert.AreEquivalent(candidates.Where(item => item.Record.ComponentType == 511)
                    .Select(item => item.Record.ObjectId.Value).ToArray(),
                type511.Evidence.Select(item => item.ObjectId.Value).ToArray());
            Assert.IsFalse(type80.Evidence.Select(item => item.SolutionComponentId)
                .Intersect(type511.Evidence.Select(item => item.SolutionComponentId)).Any());
            Assert.IsFalse(type80.Evidence.Select(item => item.ObjectId)
                .Intersect(type511.Evidence.Select(item => item.ObjectId)).Any());
            CollectionAssert.AreEquivalent(candidates.Select(item => item.Record.SolutionComponentId).ToArray(),
                diagnostics.BroadRawComponentTypes.SelectMany(group => group.Evidence)
                    .Select(item => item.SolutionComponentId).ToArray());
            CollectionAssert.AreEquivalent(candidates.Select(item => item.Record.ObjectId.Value).ToArray(),
                diagnostics.BroadRawComponentTypes.SelectMany(group => group.Evidence)
                    .Select(item => item.ObjectId.Value).ToArray());
            Assert.AreEqual(diagnostics.BroadUnclassifiable.TotalCandidates,
                diagnostics.BroadRawComponentTypes.Sum(group => group.Count));
            Assert.AreEqual(diagnostics.BroadUnclassifiable.TotalCandidates,
                diagnostics.BroadRawComponentTypes.Sum(group => group.Evidence.Count));

            Assert.ThrowsException<ArgumentException>(() => new MembershipCoverageDiagnostics(
                diagnostics.SnapshotState, diagnostics.SemanticKinds, diagnostics.BroadUnclassifiable,
                new MembershipCoverageRawComponentTypeGroup[0], diagnostics.DynamicComponentTypes));
        }

        [TestMethod]
        public void DynamicallyClassifiedFamilyIsIsolatedAndCoverageTotalsReconcile()
        {
            var firstDefinition = new SolutionComponentDefinitionIdentity(10266,
                "Contoso.Education.Assessment", "contoso_assessment");
            var secondDefinition = new SolutionComponentDefinitionIdentity(10267,
                "contoso.education.assessment", "contoso_assessment");
            var items = new[]
            {
                DynamicCandidate(firstDefinition),
                DynamicCandidate(firstDefinition),
                DynamicCandidate(secondDefinition),
                Candidate(IdentityResolutionStatus.Resolved, 2, ComponentSemanticKinds.Column, "account.name")
            };

            var diagnostics = builder.Build(Snapshot(items));
            var dynamicBucket = diagnostics.SemanticKinds.Single(item =>
                item.BucketType == MembershipCoverageBucketType.DynamicallyClassifiedIsolatedFamily);

            Assert.AreEqual(3, dynamicBucket.TotalCandidates);
            Assert.AreEqual(3, dynamicBucket.Unsupported);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, dynamicBucket.CoverageStatus);
            Assert.AreEqual(MembershipCoverageStatus.Complete,
                Kind(diagnostics, ComponentSemanticKinds.Column).CoverageStatus);
            Assert.AreEqual(0, diagnostics.BroadUnclassifiable.TotalCandidates);
            Assert.AreEqual(2, diagnostics.DynamicComponentTypes.Count);
            Assert.AreEqual(3, diagnostics.DynamicComponentTypes.Sum(item => item.Count));
            Assert.AreEqual(items.Length, diagnostics.TotalCandidates);
            Assert.AreEqual(items.Length, diagnostics.SemanticKinds.Sum(item => item.TotalCandidates) +
                diagnostics.BroadUnclassifiable.TotalCandidates);
        }

        [TestMethod]
        public void AmbiguousCandidateIsCountedAsASameKindBlocker()
        {
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Ambiguous, 2, ComponentSemanticKinds.Column,
                    diagnostic: "Duplicate portable key")));

            var column = Kind(diagnostics, ComponentSemanticKinds.Column);
            Assert.AreEqual(1, column.Ambiguous);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, column.CoverageStatus);
        }

        [TestMethod]
        public void DiagnosticGroupingPreservesStatusCaseWhitespaceAndOriginalText()
        {
            const string exact = "  Metadata lookup failed  ";
            var diagnostics = builder.Build(Snapshot(
                Candidate(IdentityResolutionStatus.Unresolved, 2, ComponentSemanticKinds.Column,
                    diagnostic: exact),
                Candidate(IdentityResolutionStatus.Unresolved, 2, ComponentSemanticKinds.Column,
                    diagnostic: exact),
                Candidate(IdentityResolutionStatus.Unresolved, 2, ComponentSemanticKinds.Column,
                    diagnostic: "  metadata lookup failed  "),
                Candidate(IdentityResolutionStatus.Unsupported, 2, ComponentSemanticKinds.Column,
                    diagnostic: exact)));

            var groups = Kind(diagnostics, ComponentSemanticKinds.Column).DiagnosticGroups;
            Assert.AreEqual(3, groups.Count);
            var repeated = groups.Single(item => item.ResolutionStatus == IdentityResolutionStatus.Unresolved &&
                item.Diagnostic == exact);
            Assert.AreEqual(2, repeated.Count);
            Assert.AreEqual(exact, repeated.Diagnostic);
            Assert.AreEqual(1, groups.Single(item => item.ResolutionStatus == IdentityResolutionStatus.Unresolved &&
                item.Diagnostic == "  metadata lookup failed  ").Count);
            Assert.AreEqual(1, groups.Single(item => item.ResolutionStatus == IdentityResolutionStatus.Unsupported &&
                item.Diagnostic == exact).Count);
        }

        [TestMethod]
        public void AbsentAndUnavailableSnapshotsExposeOppositeCoverageStates()
        {
            var absent = builder.Build(MembershipSnapshot.Absent(Solution().Environment, "sample",
                DateTimeOffset.UtcNow));
            var unavailable = builder.BuildUnavailable();

            Assert.IsTrue(absent.SemanticKinds.All(item =>
                item.CoverageStatus == MembershipCoverageStatus.Complete));
            Assert.IsTrue(unavailable.SemanticKinds.All(item =>
                item.CoverageStatus == MembershipCoverageStatus.Incomplete));
        }

        private static MembershipCoverageBucket Kind(MembershipCoverageDiagnostics diagnostics, string kind) =>
            diagnostics.SemanticKinds.Single(item => string.Equals(item.SemanticKind, kind,
                StringComparison.OrdinalIgnoreCase));

        private static ComponentIdentity Candidate(IdentityResolutionStatus status, int componentType,
            string semanticKind = null, string key = null, string diagnostic = null)
        {
            return new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(), componentType, Guid.NewGuid()),
                status, status == IdentityResolutionStatus.Resolved ? key : null,
                diagnostic, semanticKind, semanticKind);
        }

        private static ComponentIdentity DynamicCandidate(SolutionComponentDefinitionIdentity definition)
        {
            return new ComponentIdentity(new SolutionComponentRecord(Guid.NewGuid(),
                definition.ObjectTypeCode, Guid.NewGuid()), IdentityResolutionStatus.Unsupported,
                diagnostic: "No portable identity resolver supports this registered family.",
                registeredDefinition: definition);
        }
    }
}
