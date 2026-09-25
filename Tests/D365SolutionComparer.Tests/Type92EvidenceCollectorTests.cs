using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class Type92EvidenceCollectorTests
    {
        private const string Category = "Phase2G11A1";
        private const string ProductionCategory = "Phase2G11B2";

        [TestMethod, TestCategory(ProductionCategory)]
        public void PublicType92CatalogAndStandaloneResolverSupportStepsInEveryBuild()
        {
            Assert.AreEqual(ComponentSemanticKinds.SdkMessageProcessingStep,
                ComponentSemanticKinds.FromRawComponentType(92));
            Assert.AreEqual("unsupported:componenttype:90", ComponentSemanticKinds.FromRawComponentType(90));
            var fixture = new Fixture(); fixture.AddCompleteStep(Guid.NewGuid());
            var identity = new DataverseComponentIdentityResolver().Resolve(fixture.Service,
                fixture.Solution.Environment, fixture.Snapshot().Components.Single().Record,
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, identity.Status);
            Assert.AreEqual(ComponentSemanticKinds.SdkMessageProcessingStep, identity.SemanticKind);
            Assert.IsFalse(identity.Diagnostic.Contains("Guarded"));
            Assert.IsFalse(identity.DiagnosticEvidence.Any(item => item.Contains("Type92ValidationEnabled")));
            Assert.AreEqual(0, fixture.Service.WriteCalls);
#if !DEBUG
            Assert.IsNull(typeof(DataverseComponentIdentityResolver).Assembly.GetType(
                "D365SolutionComparer.Services.Membership.Type92EvidenceCollector"));
#endif
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void DefaultResolverAndApplicationCompositionResolveType92WithoutValidationSwitch()
        {
            var fixture = new Fixture(); fixture.AddCompleteStep(Guid.NewGuid());
            var defaultSnapshot = new DataverseComponentIdentityResolver().ResolveSnapshot(
                fixture.Service, fixture.Snapshot(), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, defaultSnapshot.Components.Single().Status);
            Assert.IsNull(typeof(DataverseComponentDefinitionOperation).GetProperty("Type92ValidationEnabled",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
            Assert.IsFalse(typeof(DataverseComponentDefinitionOperation)
                .GetConstructors(System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public)
                .Any(constructor => constructor.GetParameters().Any(parameter =>
                    parameter.ParameterType == typeof(bool))));
            Assert.IsNull(typeof(DataverseComponentDefinitionOperation).Assembly.GetType(
                "D365SolutionComparer.Services.Membership.Type92GuardedMembershipResolver"));
            var composed = SolutionComparerControl.CreateMembershipComparisonOperation()
                .ReadAndResolve(fixture.Service, fixture.Solution, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, composed.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentSemanticKinds.SdkMessageProcessingStep,
                composed.Membership.Components.Single().SemanticKind);
            Assert.AreEqual(ComponentDefinitionReadStatus.Unsupported, composed.Definitions.Single().Status);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void ProductionPortableKeyIsLengthFramedAndIgnoresLocalGuidsExportKeyAndCase()
        {
            var dev = new Fixture(); dev.AddCompleteStep(Guid.NewGuid(), exportKey: "DEV-key",
                typeName: " MOE.PCLookup.Plugin.GetToken ");
            var uat = new Fixture(); uat.AddCompleteStep(Guid.NewGuid(), exportKey: "UAT-key",
                typeName: "moe.pclookup.plugin.gettoken");
            var source = dev.ResolveSupported().Components.Single();
            var target = uat.ResolveSupported().Components.Single();
            StringAssert.StartsWith(source.ComparisonKey, "sdkmessageprocessingstep:v1:");
            Assert.IsFalse(source.ComparisonKey.Contains("DEV-key"));
            Assert.AreNotEqual(source.Record.ObjectId, target.Record.ObjectId);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(source.ComparisonKey, target.ComparisonKey));
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                new SolutionMembershipComparer().Compare(dev.ResolvedSnapshot, uat.ResolvedSnapshot).Single().Presence);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void MatchedProductionStepHasUnsupportedDefinitionStatus()
        {
            var source = new Fixture(); source.AddCompleteStep(Guid.NewGuid());
            var target = new Fixture(); target.AddCompleteStep(Guid.NewGuid());
            var operation = SolutionComparerControl.CreateMembershipComparisonOperation();
            var sourceDefinitions = operation.ReadAndResolve(source.Service, source.Solution,
                CancellationToken.None);
            var targetDefinitions = operation.ReadAndResolve(target.Service, target.Solution,
                CancellationToken.None);
            var membership = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", sourceDefinitions.Membership, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", targetDefinitions.Membership, 0, TimeSpan.Zero));
            var presentation = new ComponentDefinitionResultPresenter().Apply(membership,
                sourceDefinitions, targetDefinitions);
            Assert.AreEqual("Present in Both", presentation.Rows.Single().MembershipStatus);
            Assert.AreEqual("Unsupported", presentation.Rows.Single().DefinitionStatus);
            Assert.AreEqual(0, source.Service.WriteCalls);
            Assert.AreEqual(0, target.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void NoFilterAndValidNoneNoneFilterMatchButEntityScopeDiffers()
        {
            var dev = new Fixture(); dev.AddCompleteStep(Guid.NewGuid());
            var uat = new Fixture(); var id = Guid.NewGuid(); var filter = Guid.NewGuid();
            uat.AddCompleteStep(id, filterId: filter); uat.AddFilterForStep(id, filter, "none", "none");
            var left = dev.ResolveSupported(); var right = uat.ResolveSupported();
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                new SolutionMembershipComparer().Compare(left, right).Single().Presence);
            var scoped = new Fixture(); var scopedId = Guid.NewGuid(); var scopedFilter = Guid.NewGuid();
            scoped.AddCompleteStep(scopedId, filterId: scopedFilter);
            scoped.AddFilterForStep(scopedId, scopedFilter, "account", "none");
            Assert.AreNotEqual(left.Components.Single().ComparisonKey,
                scoped.ResolveSupported().Components.Single().ComparisonKey);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void MissingOrConflictingFilterBlocksUnsafeAbsence()
        {
            var complete = new Fixture(); complete.AddCompleteStep(Guid.NewGuid());
            foreach (var conflicting in new[] { false, true })
            {
                var uncertain = new Fixture(); var id = Guid.NewGuid(); var filter = Guid.NewGuid();
                uncertain.AddCompleteStep(id, filterId: filter);
                if (conflicting) uncertain.AddFilterForStep(id, filter, "none", "none", mismatch: true);
                var incomplete = uncertain.ResolveSupported();
                Assert.AreEqual(IdentityResolutionStatus.Unresolved, incomplete.Components.Single().Status);
                var compared = new SolutionMembershipComparer().Compare(complete.ResolveSupported(), incomplete);
                Assert.IsTrue(compared.All(item => item.Presence == MembershipPresence.Indeterminate));
            }
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void DuplicateOrPagedFilterAndMissingHandlerRemainUnresolved()
        {
            var duplicate = new Fixture(); var stepId = Guid.NewGuid(); var filterId = Guid.NewGuid();
            duplicate.AddCompleteStep(stepId, filterId: filterId);
            duplicate.AddFilterForStep(stepId, filterId, "none", "none");
            duplicate.Rows.Add(duplicate.Rows.Single(item => item.LogicalName == "sdkmessagefilter"));
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                duplicate.ResolveSupported().Components.Single().Status);
            var paged = new Fixture(); stepId = Guid.NewGuid(); filterId = Guid.NewGuid();
            paged.AddCompleteStep(stepId, filterId: filterId);
            paged.AddFilterForStep(stepId, filterId, "none", "none");
            paged.MoreRecordsEntity = "sdkmessagefilter";
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                paged.ResolveSupported().Components.Single().Status);
            var missingHandler = new Fixture(); stepId = Guid.NewGuid();
            missingHandler.AddCompleteStep(stepId);
            missingHandler.Rows.Single(item => item.LogicalName == "sdkmessageprocessingstep")
                .Attributes.Remove("eventhandler");
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                missingHandler.ResolveSupported().Components.Single().Status);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void AssemblyIdentityAndBothEntityScopesParticipateInTheTuple()
        {
            var first = new Fixture(); var firstId = Guid.NewGuid(); var firstFilter = Guid.NewGuid();
            first.AddCompleteStep(firstId, filterId: firstFilter);
            first.AddFilterForStep(firstId, firstFilter, "account", "contact");
            var alternateAssembly = new Fixture(); var secondId = Guid.NewGuid(); var secondFilter = Guid.NewGuid();
            alternateAssembly.AddCompleteStep(secondId, filterId: secondFilter);
            alternateAssembly.AddFilterForStep(secondId, secondFilter, "account", "contact");
            alternateAssembly.Rows.Single(item => item.LogicalName == "pluginassembly")["publickeytoken"] = "other-token";
            var alternateScope = new Fixture(); var thirdId = Guid.NewGuid(); var thirdFilter = Guid.NewGuid();
            alternateScope.AddCompleteStep(thirdId, filterId: thirdFilter);
            alternateScope.AddFilterForStep(thirdId, thirdFilter, "account", "lead");
            var baseline = first.ResolveSupported().Components.Single().ComparisonKey;
            Assert.AreNotEqual(baseline, alternateAssembly.ResolveSupported().Components.Single().ComparisonKey);
            Assert.AreNotEqual(baseline, alternateScope.ResolveSupported().Components.Single().ComparisonKey);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void IncompleteHandlerAssemblyOrMessageHasNoPortableKey()
        {
            var missingHandler = new Fixture(); missingHandler.AddCompleteStep(Guid.NewGuid(),
                handlerType: "serviceendpoint");
            var blankType = new Fixture(); blankType.AddCompleteStep(Guid.NewGuid(), typeName: " ");
            var missingAssembly = new Fixture(); missingAssembly.AddCompleteStep(Guid.NewGuid());
            missingAssembly.Rows.RemoveAll(item => item.LogicalName == "pluginassembly");
            var missingMessage = new Fixture(); missingMessage.AddCompleteStep(Guid.NewGuid());
            missingMessage.Rows.RemoveAll(item => item.LogicalName == "sdkmessage");
            foreach (var fixture in new[] { missingHandler, blankType, missingAssembly, missingMessage })
            {
                var result = fixture.ResolveSupported().Components.Single();
                Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
                Assert.IsNull(result.ComparisonKey);
            }
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void DistinctStepsWithSameKeyBlockOnlyThatKeyAndNotTimeentry()
        {
            var dev = new Fixture(); dev.AddCompleteStep(Guid.NewGuid(), typeName: "GetToken");
            var uat = new Fixture();
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "Timeentry");
            var source = dev.ResolveSupported(); var target = uat.ResolveSupported();
            var compared = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(1, compared.Count(item => item.Presence == MembershipPresence.OnlyInTarget));
            Assert.AreEqual(MembershipPresence.OnlyInTarget, compared.Single(item =>
                item.Target != null && item.Target.ComparisonKey != null &&
                item.Target.ComparisonKey.Contains("Timeentry")).Presence);
            Assert.AreEqual(2, compared.Count(item =>
                (item.Source ?? item.Target).Status == IdentityResolutionStatus.Ambiguous &&
                item.Presence == MembershipPresence.Indeterminate));
            Assert.AreEqual(MembershipPresence.Indeterminate, compared.Single(item =>
                item.Source != null && item.Source.ComparisonKey != null &&
                item.Source.ComparisonKey.Contains("GetToken")).Presence);
            Assert.AreEqual(1, target.Components.Where(item => item.Status == IdentityResolutionStatus.Ambiguous)
                .Select(item => item.BlockerPortableIdentity).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void RankFilteringConfigurationAndStateCannotSplitDuplicateSteps()
        {
            var fixture = new Fixture(); var a = Guid.NewGuid(); var b = Guid.NewGuid();
            fixture.AddCompleteStep(a, exportKey: "first");
            fixture.AddCompleteStep(b, exportKey: "second", configuration: "private value");
            var row = fixture.Rows.Single(item => item.LogicalName == "sdkmessageprocessingstep" && item.Id == b);
            row["rank"] = 99; row["filteringattributes"] = "other";
            row["statecode"] = new OptionSetValue(1); row["componentstate"] = new OptionSetValue(2);
            row["ismanaged"] = true;
            var result = fixture.ResolveSupported();
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Ambiguous));
            Assert.AreEqual(1, result.Components.Select(item => item.BlockerPortableIdentity)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void RepeatedRawReferencesCollapseAndDefinitionUsesCachedEvidenceOnly()
        {
            var fixture = new Fixture(); var id = Guid.NewGuid();
            fixture.AddCompleteStep(id); fixture.AddRaw(id);
            var snapshot = fixture.Snapshot();
            var context = new DataverseReadContext(fixture.Service, snapshot.Environment, CancellationToken.None);
            var resolved = new DataverseComponentIdentityResolver().ResolveSnapshot(context, snapshot, CancellationToken.None);
            Assert.IsNotNull(context.MetadataCache.Type92Evidence);
            var requests = fixture.Service.Calls;
            var definitions = new DataverseComponentDefinitionReader().Read(context, resolved, CancellationToken.None);
            Assert.AreEqual(requests, fixture.Service.Calls);
            Assert.IsTrue(definitions.Definitions.All(item => item.Status == ComponentDefinitionReadStatus.Unsupported));
            Assert.AreEqual(1, new SolutionMembershipComparer().Compare(resolved, resolved).Count);
            Assert.AreEqual(1, fixture.Service.ExecuteCalls);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void CapturedEightAndTenShapeReconcilesWithoutHeuristicMatching()
        {
            var dev = new Fixture(); var uat = new Fixture();
            for (int i = 0; i < 7; i++)
            {
                dev.AddCompleteStep(Guid.NewGuid(), typeName: "Common." + i);
                uat.AddCompleteStep(Guid.NewGuid(), typeName: "Common." + i);
            }
            dev.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "Moe.Plugin.TimeentryExcelGenerator.GenerateExcel");
            var source = dev.ResolveSupported(); var target = uat.ResolveSupported();
            var compared = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(8, source.Components.Count);
            Assert.AreEqual(10, target.Components.Count);
            Assert.AreEqual(7, compared.Count(item => item.Presence == MembershipPresence.PresentInBoth));
            Assert.AreEqual(0, compared.Count(item => item.Presence == MembershipPresence.OnlyInSource));
            Assert.AreEqual(1, compared.Count(item => item.Presence == MembershipPresence.OnlyInTarget));
            Assert.AreEqual(1, compared.Where(item => (item.Source ?? item.Target).Status ==
                IdentityResolutionStatus.Ambiguous).Select(item =>
                (item.Source ?? item.Target).BlockerPortableIdentity)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void AmbiguousParentAssemblyOrBlankSdkMessageCannotResolve()
        {
            var ambiguousAssembly = new Fixture(); ambiguousAssembly.AddCompleteStep(Guid.NewGuid());
            var parent = ambiguousAssembly.Rows.Single(item => item.LogicalName == "pluginassembly");
            ambiguousAssembly.Rows.Add(parent);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                ambiguousAssembly.ResolveSupported().Components.Single().Status);
            var blankMessage = new Fixture(); blankMessage.AddCompleteStep(Guid.NewGuid());
            blankMessage.Rows.Single(item => item.LogicalName == "sdkmessage")["name"] = " ";
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                blankMessage.ResolveSupported().Components.Single().Status);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void ProductionCompositionAddsOneBatchedReadPerFamilyWithoutAnotherWhoAmI()
        {
            var fixture = new Fixture(); fixture.AddCompleteStep(Guid.NewGuid());
            var result = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(
                fixture.Service, fixture.Solution, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.AreEqual(1, fixture.Service.ExecuteCalls);
            foreach (var entity in new[] { "solution", "solutioncomponent", "sdkmessageprocessingstep",
                "plugintype", "pluginassembly", "sdkmessage" })
                Assert.AreEqual(1, fixture.Queries.Count(query => query.EntityName == entity), entity);
            Assert.AreEqual(0, fixture.Queries.Count(query => query.EntityName == "sdkmessagefilter"));
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void ProductionResolutionBatches201DistinctStepsWithoutSecureConfigurationOrExtraWhoAmI()
        {
            var fixture = new Fixture();
            for (var i = 0; i < 201; i++) fixture.AddRaw(Guid.NewGuid());
            var resolved = fixture.ResolveSupported();
            Assert.AreEqual(201, resolved.Components.Count);
            Assert.IsTrue(resolved.Components.All(item => item.Status == IdentityResolutionStatus.Unresolved));
            var stepQueries = fixture.Queries.Where(query =>
                query.EntityName == "sdkmessageprocessingstep").ToList();
            Assert.AreEqual(2, stepQueries.Count);
            CollectionAssert.AreEquivalent(new[] { 200, 1 }, stepQueries.Select(query =>
                query.Criteria.Conditions.Single().Values.Count).ToArray());
            Assert.IsTrue(stepQueries.All(query => query.Criteria.Conditions.Single().Values
                .All(value => value is Guid) && !query.ColumnSet.Columns.Contains(
                    "sdkmessageprocessingstepsecureconfigid")));
            Assert.AreEqual(1, fixture.Service.ExecuteCalls);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void CoordinatedType91ParentAssemblyRowIsReusedByType92WithoutAnotherQuery()
        {
            var fixture = new Fixture(); fixture.AddCompleteStep(Guid.NewGuid());
            var assemblyId = fixture.Rows.Single(item => item.LogicalName == "pluginassembly").Id;
            fixture.AddRaw(assemblyId, 91);
            var result = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(
                fixture.Service, fixture.Solution, CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single(item =>
                item.Record.ComponentType == 92).Status);
            Assert.AreEqual(1, fixture.Queries.Count(query => query.EntityName == "pluginassembly"));
            Assert.AreEqual(1, fixture.Service.ExecuteCalls);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void ProductionFailureAndCancellationNeverReturnCompletedPartialInventory()
        {
            var failed = new Fixture(); failed.AddCompleteStep(Guid.NewGuid());
            failed.FaultEntity = "sdkmessage";
            Assert.ThrowsException<FaultException>(() => failed.ResolveSupported());
            var cancelled = new Fixture(); cancelled.AddCompleteStep(Guid.NewGuid());
            var source = new CancellationTokenSource(); source.Cancel();
            Assert.ThrowsException<OperationCanceledException>(() => cancelled.ResolveSupported(source.Token));
            Assert.AreEqual(0, cancelled.Service.Calls);
        }

        [TestMethod, TestCategory(ProductionCategory)]
        public void ProductionKindAppearsInPresentationAndSupportedCoverageSeed()
        {
            var dev = new Fixture(); dev.AddCompleteStep(Guid.NewGuid());
            var uat = new Fixture(); uat.AddCompleteStep(Guid.NewGuid());
            var source = dev.ResolveSupported(); var target = uat.ResolveSupported();
            var presentation = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", source, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", target, 0, TimeSpan.Zero));
            Assert.AreEqual("SDK Message Processing Step", presentation.Rows.Single().ComponentKind);
            var bucket = new MembershipCoverageDiagnosticsBuilder().Build(source).SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.SdkMessageProcessingStep);
            Assert.AreEqual("SDK Message Processing Step", bucket.DisplayName);
            Assert.AreEqual(1, bucket.Resolved);
            var empty = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(dev.Solution, new ComponentIdentity[0], DateTimeOffset.UtcNow));
            Assert.IsTrue(empty.SemanticKinds.Any(item =>
                item.SemanticKind == ComponentSemanticKinds.SdkMessageProcessingStep && item.TotalCandidates == 0));
            Assert.AreEqual(ComponentSemanticKinds.SdkMessageProcessingStep,
                ComponentSemanticKinds.FromRawComponentType(92));
        }

#if DEBUG
        [TestMethod, TestCategory(Category)]
        public void BatchesAt200And201WithGuidFiltersAndNoWhoAmI()
        {
            var fixture = new Fixture();
            for (var i = 0; i < 201; i++) fixture.AddRaw(Guid.NewGuid());
            var report = fixture.Capture();
            Assert.AreEqual(201, report.Raw.Count);
            Assert.AreEqual(2, report.Count("sdkmessageprocessingstep"));
            Assert.AreEqual(2, fixture.Service.Calls);
            Assert.AreEqual(0, fixture.Service.ExecuteCalls);
            Assert.AreEqual(0, fixture.Service.WriteCalls);
            CollectionAssert.AreEquivalent(new[] { 200, 1 }, fixture.Queries.Select(query =>
                query.Criteria.Conditions.Single().Values.Count).ToArray());
            Assert.IsTrue(fixture.Queries.All(query => query.Criteria.Conditions.Single().Values
                .All(value => value is Guid)));
            Assert.IsTrue(fixture.Queries.All(query => !query.ColumnSet.Columns.Contains(
                "sdkmessageprocessingstepsecureconfigid")));
        }

        [TestMethod, TestCategory(Category)]
        public void Exactly200UsesOneBatch()
        {
            var fixture = new Fixture();
            for (var i = 0; i < 200; i++) fixture.AddRaw(Guid.NewGuid());
            Assert.AreEqual(1, fixture.Capture().Count("sdkmessageprocessingstep"));
        }

        [TestMethod, TestCategory(Category)]
        public void RepeatedRawMembershipCollapsesToOneBackingStep()
        {
            var fixture = new Fixture();
            var id = Guid.NewGuid();
            fixture.AddCompleteStep(id);
            fixture.AddRaw(id);
            var report = fixture.Capture();
            Assert.AreEqual(2, report.Raw.Count);
            Assert.AreEqual(1, report.Steps.Count);
            Assert.AreEqual(2, report.Steps[0].MembershipCount);
            Assert.AreEqual("Candidate", report.Steps[0].CandidateStatus);
            Assert.AreEqual(1, report.Count("sdkmessageprocessingstep"));
        }

        [TestMethod, TestCategory(Category)]
        public void DifferentBackingStepsWithSameTupleAreAmbiguous()
        {
            var fixture = new Fixture();
            var first = Guid.NewGuid();
            var second = Guid.NewGuid();
            fixture.AddCompleteStep(first);
            fixture.AddCompleteStep(second, exportKey: "Different.Type");
            var secondStep = fixture.Rows.Single(item => item.LogicalName == "sdkmessageprocessingstep" && item.Id == second);
            secondStep["rank"] = 7;
            secondStep["filteringattributes"] = "differentfield";
            secondStep["configuration"] = "different configuration";
            var report = fixture.Capture();
            Assert.AreEqual(2, report.Steps.Count);
            Assert.IsTrue(report.Steps.All(item => item.CandidateStatus == "AmbiguousCandidate"));
            StringAssert.Contains(Type92EvidenceReport.Build(report, report), "Ambiguous identity group | handlerAssembly=");
        }

        [TestMethod, TestCategory(Category)]
        public void BlankExportKeyIsAuditOnlyWhenAssemblyAndTypeNameAreKnown()
        {
            var fixture = new Fixture();
            fixture.AddCompleteStep(Guid.NewGuid(), exportKey: "");
            var step = fixture.Capture().Steps[0];
            Assert.AreEqual("Candidate", step.CandidateStatus);
            Assert.AreEqual("", step.ExportKey);
            Assert.IsNotNull(step.Candidate);
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateExportKeyDoesNotDisambiguateDistinctBackingSteps()
        {
            var fixture = new Fixture();
            fixture.AddCompleteStep(Guid.NewGuid(), handlerId: Guid.NewGuid());
            fixture.AddCompleteStep(Guid.NewGuid(), handlerId: Guid.NewGuid());
            Assert.IsTrue(fixture.Capture().Steps.All(item => item.CandidateStatus == "AmbiguousCandidate"));
        }

        [TestMethod, TestCategory(Category)]
        public void AbsentFilterIsNoFilterButMissingReferencedFilterIsUnresolved()
        {
            var fixture = new Fixture();
            var noFilter = Guid.NewGuid();
            fixture.AddCompleteStep(noFilter);
            var missingFilter = Guid.NewGuid();
            fixture.AddCompleteStep(missingFilter, exportKey: "Other.Type", filterId: Guid.NewGuid());
            var report = fixture.Capture();
            Assert.AreEqual("NoFilter", report.Steps.Single(item => item.StepId == noFilter).FilterStatus);
            Assert.AreEqual("Candidate", report.Steps.Single(item => item.StepId == noFilter).CandidateStatus);
            Assert.AreEqual("UnresolvedFilter", report.Steps.Single(item => item.StepId == missingFilter).CandidateStatus);
            Assert.AreEqual(1, report.Count("sdkmessagefilter"));
        }

        [TestMethod, TestCategory(Category)]
        public void NonPluginHandlerDoesNotAcquireFallbackIdentity()
        {
            var fixture = new Fixture();
            var id = Guid.NewGuid();
            fixture.AddCompleteStep(id, handlerType: "serviceendpoint");
            var report = fixture.Capture();
            Assert.AreEqual("serviceendpoint", report.Steps[0].HandlerCategory);
            Assert.AreEqual("IncompleteHandler", report.Steps[0].CandidateStatus);
            Assert.AreEqual(0, report.Count("plugintype"));
            Assert.AreEqual(0, report.Count("pluginassembly"));
        }

        [TestMethod, TestCategory(Category)]
        public void FilteringAttributesAreCanonicalAndConfigurationContentNeverEmitted()
        {
            Assert.AreEqual("accountid,name", Type92EvidenceCollector.CanonicalFilteringAttributes(
                " Name,accountid, name , , ACCOUNTID"));
            var fixture = new Fixture();
            fixture.AddCompleteStep(Guid.NewGuid(), configuration: "SECRET_PRIVATE_TOKEN");
            var report = fixture.Capture();
            var text = Type92EvidenceReport.Build(report, report);
            StringAssert.Contains(text, "configurationPresent=True");
            StringAssert.Contains(text, "configurationLength=20");
            Assert.IsFalse(text.Contains("SECRET_PRIVATE_TOKEN"));
            Assert.AreEqual("accountid,name", report.Steps[0].FilteringCanonical);
        }

        [TestMethod, TestCategory(Category)]
        public void CancellationPropagatesWithoutReturningPartialEvidence()
        {
            var fixture = new Fixture();
            fixture.AddCompleteStep(Guid.NewGuid());
            var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.ThrowsException<OperationCanceledException>(() => fixture.Capture(cancellation.Token));
            Assert.AreEqual(0, fixture.Service.Calls);
        }

        [TestMethod, TestCategory(Category)]
        public void QueryFaultPropagatesWithoutReturningPartialEvidence()
        {
            var fixture = new Fixture();
            fixture.AddCompleteStep(Guid.NewGuid());
            fixture.FaultEntity = "sdkmessage";
            Assert.ThrowsException<FaultException>(() => fixture.Capture());
            Assert.AreEqual(0, fixture.Service.WriteCalls);
        }

        [TestMethod, TestCategory(Category)]
        public void MissingDuplicateConflictingAndIncompleteStepCorrelationsRemainUntrusted()
        {
            var missing = new Fixture();
            missing.AddRaw(Guid.NewGuid());
            Assert.AreEqual("Missing", missing.Capture().Steps[0].Correlation);
            var duplicate = new Fixture();
            var id = Guid.NewGuid(); duplicate.AddRaw(id);
            duplicate.Rows.Add(Step(id)); duplicate.Rows.Add(Step(id));
            Assert.AreEqual("DuplicateReturnedRow", duplicate.Capture().Steps[0].Correlation);
            var conflict = new Fixture();
            var conflictingId = Guid.NewGuid(); conflict.AddRaw(conflictingId);
            var row = Step(conflictingId); row.Id = Guid.NewGuid(); conflict.Rows.Add(row);
            Assert.AreEqual("ConflictingPrimaryKey", conflict.Capture().Steps[0].Correlation);
            var incomplete = new Fixture();
            incomplete.AddRaw(Guid.NewGuid()); incomplete.MoreRecordsEntity = "sdkmessageprocessingstep";
            Assert.AreEqual("IncompleteBatch", incomplete.Capture().Steps[0].Correlation);
        }

        [TestMethod, TestCategory(Category)]
        public void CaseInsensitiveTupleMatchingAndDifferentLocalGuidsNeedNoGuidEquality()
        {
            var source = new Fixture("DEV"); source.AddCompleteStep(Guid.NewGuid(),
                exportKey: "EXAMPLE.Type", typeName: " Example.Type ");
            var target = new Fixture("UAT"); target.AddCompleteStep(Guid.NewGuid(),
                exportKey: "example.type", typeName: "example.type");
            var left = source.Capture(); var right = target.Capture();
            Assert.AreNotEqual(left.Steps[0].StepId, right.Steps[0].StepId);
            Assert.AreEqual(left.Steps[0].Candidate, right.Steps[0].Candidate);
            StringAssert.Contains(Type92EvidenceReport.Build(left, right), "Unique match | handlerAssembly=");
            Assert.IsTrue(left.Snapshot.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
            Assert.IsTrue(right.Snapshot.Components.All(item => item.Status == IdentityResolutionStatus.Unsupported));
        }

        [TestMethod, TestCategory(Category)]
        public void DifferentExportKeysAndPluginTypeGuidsMatchThroughAssemblyAndTypeName()
        {
            var source = new Fixture("DEV"); source.AddCompleteStep(Guid.NewGuid(), exportKey: "DEV-export");
            var target = new Fixture("UAT"); target.AddCompleteStep(Guid.NewGuid(), exportKey: "UAT-export");
            var sourceEvidence = source.Capture(); var targetEvidence = target.Capture();
            var left = sourceEvidence.Steps[0]; var right = targetEvidence.Steps[0];
            Assert.AreNotEqual(left.PluginTypeId, right.PluginTypeId);
            Assert.AreNotEqual(left.AssemblyId, right.AssemblyId);
            Assert.AreNotEqual(left.ExportKey, right.ExportKey);
            Assert.AreEqual(left.Candidate, right.Candidate);
            Assert.AreEqual(left.HandlerSemanticIdentity, right.HandlerSemanticIdentity);
            var report = Type92EvidenceReport.Build(sourceEvidence, targetEvidence);
            StringAssert.Contains(report, "plugintypeid=");
            StringAssert.Contains(report, "exportkey=DEV-export");
            StringAssert.Contains(report, "handlerSemanticIdentity=pluginassembly:v1:");
            StringAssert.Contains(report, "semanticScope=Global/Unbound");
            Assert.AreEqual(1, sourceEvidence.Count("plugintype"));
            Assert.AreEqual(1, sourceEvidence.Count("pluginassembly"));
        }

        [TestMethod, TestCategory(Category)]
        public void NoFilterAndResolvedNoneNoneFilterHaveTheSameGlobalScope()
        {
            var source = new Fixture("DEV"); source.AddCompleteStep(Guid.NewGuid());
            var target = new Fixture("UAT");
            var targetStep = Guid.NewGuid(); var filterId = Guid.NewGuid();
            target.AddCompleteStep(targetStep, filterId: filterId);
            target.AddFilterForStep(targetStep, filterId, "none", " NONE ");
            var left = source.Capture().Steps[0]; var right = target.Capture().Steps[0];
            Assert.AreEqual("NoFilter", left.FilterStatus);
            Assert.AreEqual("Filter", right.FilterStatus);
            Assert.AreEqual("Global/Unbound", left.SemanticScope);
            Assert.AreEqual("Global/Unbound", right.SemanticScope);
            Assert.AreEqual(left.Candidate, right.Candidate);
        }

        [TestMethod, TestCategory(Category)]
        public void FilterMessageMismatchRemainsIncomplete()
        {
            var fixture = new Fixture();
            var stepId = Guid.NewGuid(); var filterId = Guid.NewGuid();
            fixture.AddCompleteStep(stepId, filterId: filterId);
            fixture.AddFilterForStep(stepId, filterId, "none", "none", mismatch: true);
            var result = fixture.Capture().Steps[0];
            Assert.AreEqual("ConflictingFilter", result.CandidateStatus);
            Assert.IsNull(result.Candidate);
            Assert.IsNull(result.SemanticScope);
        }

        [TestMethod, TestCategory(Category)]
        public void OneDevAndTwoUatSameTupleRemainOneAmbiguousGroup()
        {
            var dev = new Fixture("DEV"); dev.AddCompleteStep(Guid.NewGuid(), exportKey: "DEV-export");
            var uat = new Fixture("UAT");
            uat.AddCompleteStep(Guid.NewGuid(), exportKey: "UAT-first");
            uat.AddCompleteStep(Guid.NewGuid(), exportKey: "UAT-second");
            var left = dev.Capture(); var right = uat.Capture();
            Assert.AreEqual(1, left.Steps.Count);
            Assert.AreEqual(2, right.Steps.Select(item => item.StepId).Distinct().Count());
            Assert.IsTrue(right.Steps.All(item => item.CandidateStatus == "AmbiguousCandidate"));
            var report = Type92EvidenceReport.Build(left, right);
            StringAssert.Contains(report, "Candidate totals: matched=0; DEV-only=0; UAT-only=0; ambiguous=1");
        }

        [TestMethod, TestCategory(Category)]
        public void DistinctTimeentryHandlerIsUatOnlyBesideAmbiguousGetToken()
        {
            var dev = new Fixture("DEV"); dev.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            var uat = new Fixture("UAT");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "MOE.PCLookup.Plugin.GetToken");
            uat.AddCompleteStep(Guid.NewGuid(), typeName: "Moe.Plugin.TimeentryExcelGenerator");
            var report = Type92EvidenceReport.Build(dev.Capture(), uat.Capture());
            StringAssert.Contains(report, "Candidate totals: matched=0; DEV-only=0; UAT-only=1; ambiguous=1");
            StringAssert.Contains(report, "UAT only | handlerAssembly=");
        }

        [TestMethod, TestCategory(Category)]
        public void MissingTypeNameOrAssemblyKeepsHandlerIncomplete()
        {
            var blankType = new Fixture(); blankType.AddCompleteStep(Guid.NewGuid(), typeName: " ");
            Assert.AreEqual("IncompleteHandler", blankType.Capture().Steps[0].CandidateStatus);
            var missingAssembly = new Fixture(); missingAssembly.AddCompleteStep(Guid.NewGuid());
            missingAssembly.Rows.RemoveAll(item => item.LogicalName == "pluginassembly");
            Assert.AreEqual("IncompleteHandler", missingAssembly.Capture().Steps[0].CandidateStatus);
        }
#endif

        private static Entity Step(Guid id) => new Entity("sdkmessageprocessingstep", id)
        { ["sdkmessageprocessingstepid"] = id };

        private sealed class Fixture
        {
            private readonly List<ComponentIdentity> components = new List<ComponentIdentity>();
            private readonly D365SolutionComparer.Models.Identity.SolutionIdentity solution;
            internal readonly List<Entity> Rows = new List<Entity>();
            internal readonly List<QueryExpression> Queries = new List<QueryExpression>();
            internal string FaultEntity;
            internal string MoreRecordsEntity;
            internal FakeOrganizationService Service;

            internal Fixture(string name = "DEV")
            {
                solution = MembershipTestData.Solution("PCLookupTokenPlugin");
                Service = MembershipTestData.Service(solution, query =>
                {
                    Queries.Add(query);
                    if (query.EntityName == FaultEntity) throw new FaultException("Read failed.");
                    if (query.EntityName == "solution")
                        return MembershipTestData.Rows(MembershipTestData.SolutionRow(solution));
                    if (query.EntityName == "solutioncomponent")
                        return MembershipTestData.Rows(components.Select(item =>
                        {
                            var record = item.Record;
                            return new Entity("solutioncomponent", record.SolutionComponentId)
                            {
                                ["solutioncomponentid"] = record.SolutionComponentId,
                                ["solutionid"] = new EntityReference("solution", solution.SolutionId),
                                ["componenttype"] = new OptionSetValue(record.ComponentType),
                                ["objectid"] = record.ObjectId
                            };
                        }).ToArray());
                    var ids = query.Criteria.Conditions.Single().Values.Cast<Guid>().ToHashSet();
                    var collection = MembershipTestData.Rows(Rows.Where(row =>
                        row.LogicalName == query.EntityName && ids.Contains((Guid)row[
                            query.Criteria.Conditions.Single().AttributeName])).ToArray());
                    collection.MoreRecords = query.EntityName == MoreRecordsEntity;
                    return collection;
                });
                EnvironmentName = name;
            }
            private string EnvironmentName { get; }
            internal D365SolutionComparer.Models.Identity.SolutionIdentity Solution => solution;
            internal void AddRaw(Guid id, int componentType = 92) => components.Add(new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), componentType, id), IdentityResolutionStatus.Unsupported));
            internal void AddCompleteStep(Guid id, Guid? handlerId = null, string exportKey = "Example.Type",
                string handlerType = "plugintype", Guid? filterId = null, string configuration = null,
                string typeName = "Example.Type")
            {
                AddRaw(id);
                var typeId = handlerId ?? Guid.NewGuid();
                var messageId = Guid.NewGuid();
                var assemblyId = Guid.NewGuid();
                var step = Step(id);
                step["eventhandler"] = new EntityReference(handlerType, typeId);
                step["sdkmessageid"] = new EntityReference("sdkmessage", messageId);
                if (filterId.HasValue) step["sdkmessagefilterid"] = new EntityReference("sdkmessagefilter", filterId.Value);
                step["stage"] = new OptionSetValue(20);
                step["mode"] = new OptionSetValue(0);
                step["supporteddeployment"] = new OptionSetValue(0);
                step["rank"] = 1;
                step["filteringattributes"] = " Name,accountid, name , , ACCOUNTID";
                step["configuration"] = configuration;
                Rows.Add(step);
                if (handlerType != "plugintype") return;
                Rows.Add(new Entity("plugintype", typeId)
                {
                    ["plugintypeid"] = typeId, ["plugintypeexportkey"] = exportKey,
                    ["pluginassemblyid"] = new EntityReference("pluginassembly", assemblyId),
                    ["typename"] = typeName
                });
                Rows.Add(new Entity("pluginassembly", assemblyId)
                {
                    ["pluginassemblyid"] = assemblyId, ["name"] = "Example.Assembly",
                    ["publickeytoken"] = "abcd", ["culture"] = "neutral", ["version"] = "1.0.0.0"
                });
                Rows.Add(new Entity("sdkmessage", messageId)
                { ["sdkmessageid"] = messageId, ["name"] = "Update" });
            }
            internal void AddFilterForStep(Guid stepId, Guid filterId, string primary, string secondary,
                bool mismatch = false)
            {
                var step = Rows.Single(item => item.LogicalName == "sdkmessageprocessingstep" && item.Id == stepId);
                var message = ((EntityReference)step["sdkmessageid"]).Id;
                Rows.Add(new Entity("sdkmessagefilter", filterId)
                {
                    ["sdkmessagefilterid"] = filterId,
                    ["sdkmessageid"] = new EntityReference("sdkmessage", mismatch ? Guid.NewGuid() : message),
                    ["primaryobjecttypecode"] = primary,
                    ["secondaryobjecttypecode"] = secondary
                });
            }
#if DEBUG
            internal Type92EnvironmentEvidence Capture(CancellationToken token = default(CancellationToken))
            {
                var snapshot = Snapshot();
                return new Type92EvidenceCollector().Capture(Service, snapshot, "1.0.0.0", token);
            }
#endif
            internal MembershipSnapshot Snapshot() =>
                MembershipSnapshot.Complete(solution, components, DateTimeOffset.UtcNow);
            internal MembershipSnapshot ResolvedSnapshot { get; private set; }
            internal MembershipSnapshot ResolveSupported(CancellationToken token = default(CancellationToken))
            {
                var snapshot = Snapshot();
                var context = new DataverseReadContext(Service, snapshot.Environment, token);
                ResolvedSnapshot = new DataverseComponentIdentityResolver().ResolveSnapshot(context,
                    snapshot, token);
                return ResolvedSnapshot;
            }
        }
    }
}
