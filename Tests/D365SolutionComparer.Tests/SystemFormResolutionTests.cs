using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using D365SolutionComparer.Infrastructure;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.ComponentDetails;
using D365SolutionComparer.Services.Membership;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class SystemFormResolutionTests
    {
        private const string Category = "Phase2G8B1";

        [TestMethod, TestCategory(Category)]
        public void EntityBackedUniqueNameMatchesAcrossLocalIdsAndCase()
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var sourceIdentity = ResolveOne(source, Form(Guid.NewGuid(), "new_Main", "account", 2, 1));
            var targetIdentity = ResolveOne(target, Form(Guid.NewGuid(), "NEW_MAIN", "ACCOUNT", 2, 1));

            Assert.AreEqual(IdentityResolutionStatus.Resolved, sourceIdentity.Status);
            Assert.AreEqual("systemform:v1:entity:7:account:8:new_Main", sourceIdentity.ComparisonKey);
            Assert.AreEqual(MembershipPresence.PresentInBoth, Compare(source, target,
                sourceIdentity, targetIdentity).Single().Presence);
            StringAssert.Contains(sourceIdentity.DiagnosticEvidence.Last(item =>
                item.Contains("identityPath=")), "identityPath=EntityScopedUniqueName");
            StringAssert.Contains(sourceIdentity.DiagnosticEvidence.Last(item =>
                item.Contains("identityPath=")),
                "System Form correlation evidence. Portable identity is selected according to " +
                "the reported identityPath and inventoryAbsencePolicy.");
        }

        [TestMethod, TestCategory(Category)]
        public void SameUniqueNameOnDifferentParentsProducesDifferentPortableIdentities()
        {
            var solution = Solution("DEV");
            var account = ResolveOne(solution, Form(Guid.NewGuid(), "new_Main", "account", 2, 1));
            var contact = ResolveOne(solution, Form(Guid.NewGuid(), "new_Main", "contact", 2, 1));
            Assert.AreNotEqual(account.ComparisonKey, contact.ComparisonKey);
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateEntityScopedIdentityIsAmbiguous()
        {
            var solution = Solution("DEV");
            var result = Resolve(solution, new[]
            {
                Form(Guid.NewGuid(), "new_Main", "account", 2, 1),
                Form(Guid.NewGuid(), "NEW_MAIN", "ACCOUNT", 7, 1)
            });
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(result.Components.All(item => item.ComparisonKey == null));
        }

        [TestMethod, TestCategory(Category)]
        public void TablelessUniqueNameMatchesWithoutInventedParent()
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var left = ResolveOne(source, Form(Guid.NewGuid(), "new_Dashboard", "none", 8, 1));
            var right = ResolveOne(target, Form(Guid.NewGuid(), "NEW_DASHBOARD", "NONE", 8, 1));
            Assert.AreEqual("systemform:v1:tableless:13:new_Dashboard", left.ComparisonKey);
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                Compare(source, target, left, right).Single().Presence);
            StringAssert.Contains(left.DiagnosticEvidence.Last(item => item.Contains("identityPath=")),
                "identityPath=TablelessUniqueName");
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateTablelessUniqueNameIsAmbiguous()
        {
            var solution = Solution("DEV");
            var result = Resolve(solution, new[]
            {
                Form(Guid.NewGuid(), "new_Dashboard", "none", 8, 1),
                Form(Guid.NewGuid(), "NEW_DASHBOARD", "NONE", 9, 1)
            });
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Ambiguous));
        }

        [TestMethod, TestCategory(Category)]
        public void BlankUniqueNameMatchingFormIdEstablishesSharedPresence()
        {
            var formId = Guid.NewGuid();
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var left = ResolveOne(source, Form(formId, " ", "account", 2, 1));
            var right = ResolveOne(target, Form(formId, null, "account", 2, 1));
            Assert.AreEqual(InventoryAbsencePolicy.MatchOnly, left.InventoryAbsencePolicy);
            Assert.AreEqual("systemform:v1:formid:36:" + formId.ToString("D"), left.ComparisonKey);
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                Compare(source, target, left, right).Single().Presence);
            StringAssert.Contains(left.DiagnosticEvidence.Last(item => item.Contains("identityPath=")),
                "identityPath=FormIdFallback");
        }

        [TestMethod, TestCategory(Category)]
        public void BlankUniqueNameDifferingFormIdsRemainIndeterminate()
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var results = Compare(source, target,
                ResolveOne(source, Form(Guid.NewGuid(), "", "account", 2, 1)),
                ResolveOne(target, Form(Guid.NewGuid(), "", "account", 2, 1)));
            Assert.AreEqual(2, results.Count);
            Assert.IsTrue(results.All(item => item.Presence == MembershipPresence.Indeterminate &&
                item.AbsenceEvidence == MembershipAbsenceEvidence.None));
            var sourceSnapshot = MembershipSnapshot.Complete(source,
                new[] { results.Single(item => item.Source != null).Source }, DateTimeOffset.UtcNow);
            var targetSnapshot = MembershipSnapshot.Complete(target,
                new[] { results.Single(item => item.Target != null).Target }, DateTimeOffset.UtcNow);
            var presentation = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", sourceSnapshot, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", targetSnapshot, 0, TimeSpan.Zero));
            Assert.IsTrue(presentation.Rows.All(item => item.ComponentKind == "System Form"));
            Assert.IsTrue(presentation.Rows.All(item => item.Diagnostic.Contains(
                "can prove a shared match but cannot use complete inventory")));
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(sourceSnapshot)
                .SemanticKinds.Single(item => item.SemanticKind == ComponentSemanticKinds.SystemForm);
            Assert.AreEqual("System Form", coverage.DisplayName);
            Assert.AreEqual(MembershipCoverageStatus.Incomplete, coverage.CoverageStatus);
        }

        [TestMethod, TestCategory(Category)]
        public void FormIdFallbackBlocksCrossPathSystemFormAbsence()
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var fallback = ResolveOne(source, Form(Guid.NewGuid(), "", "account", 2, 1));
            var semantic = ResolveOne(target, Form(Guid.NewGuid(), "new_Main", "account", 2, 1));
            var results = Compare(source, target, fallback, semantic);
            Assert.AreEqual(2, results.Count);
            Assert.IsTrue(results.All(item => item.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory(Category)]
        public void SemanticNameIdentityCanUseCompleteOppositeInventoryForAbsence()
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var sourceIdentity = ResolveOne(source,
                Form(Guid.NewGuid(), "new_Main", "account", 2, 1));
            var result = new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(source, new[] { sourceIdentity }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(target, new ComponentIdentity[0], DateTimeOffset.UtcNow)).Single();
            Assert.AreEqual(MembershipPresence.OnlyInSource, result.Presence);
            Assert.AreEqual(MembershipAbsenceEvidence.CompleteResolvedInventory, result.AbsenceEvidence);
        }

        [TestMethod, TestCategory(Category)]
        public void SystemFormObjectIdMustCorrelateToFormId()
        {
            var solution = Solution("DEV"); var requested = Guid.NewGuid();
            var result = Resolve(solution, new[] { Raw(requested) }, query =>
                Rows(Form(Guid.NewGuid(), "new_Main", "account", 2, 1)));
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void MissingBackingRecordIsUnresolved()
        {
            var solution = Solution("DEV");
            Assert.AreEqual(IdentityResolutionStatus.Unresolved,
                Resolve(solution, new[] { Raw(Guid.NewGuid()) }, query => Rows())
                    .Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void MultipleBackingCandidatesAreAmbiguous()
        {
            var solution = Solution("DEV"); var id = Guid.NewGuid();
            var result = Resolve(solution, new[] { Raw(id) }, query => Rows(
                Form(id, "new_First", "account", 2, 1),
                Form(id, "new_Second", "account", 2, 1)));
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Components.Single().Status);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow(2)]
        [DataRow(7)]
        [DataRow(8)]
        [DataRow(9)]
        public void SupportedLiveFormTypesDoNotChangeIdentityConstruction(int formType)
        {
            var result = ResolveOne(Solution("DEV"),
                Form(Guid.NewGuid(), "new_Form", "account", formType, 1));
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            StringAssert.Contains(result.ComparisonKey, "new_Form");
        }

        [TestMethod, TestCategory(Category)]
        public void DefinitionMatchUsesOnlyTypeAndActivationState()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1,
                    name: "DEV", isManaged: false, formXml: "<form>DEV</form>"),
                Form(Guid.NewGuid(), "NEW_MAIN", "ACCOUNT", 2, 1,
                    name: "UAT", isManaged: true, formXml: "<form>UAT</form>"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            Assert.AreEqual(2, result.Source.ComparableProperties.Count);
        }

        [TestMethod, TestCategory(Category)]
        public void TypeDifferenceIsDifferent()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1),
                Form(Guid.NewGuid(), "new_Main", "account", 7, 1));
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            CollectionAssert.AreEqual(new[] { "type" }, result.Differences.Select(item => item.PropertyName).ToArray());
        }

        [TestMethod, TestCategory(Category)]
        public void ActivationStateDifferenceIsDifferent()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1),
                Form(Guid.NewGuid(), "new_Main", "account", 2, 0));
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            CollectionAssert.AreEqual(new[] { "formactivationstate" },
                result.Differences.Select(item => item.PropertyName).ToArray());
        }

        [TestMethod, TestCategory(Category)]
        public void ManagedStateDifferenceIsIgnored()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1,
                    isManaged: false),
                Form(Guid.NewGuid(), "new_Main", "account", 2, 1,
                    isManaged: true));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            CollectionAssert.Contains(ComponentDefinitionContractCatalog.For(
                ComponentSemanticKinds.SystemForm).AuditOnlyProperties.ToArray(), "ismanaged");
        }

        [TestMethod, TestCategory(Category)]
        public void DisplayNameDifferenceIsIgnored()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1,
                    name: "First"),
                Form(Guid.NewGuid(), "new_Main", "account", 2, 1, name: "Second"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            CollectionAssert.Contains(ComponentDefinitionContractCatalog.For(
                ComponentSemanticKinds.SystemForm).AuditOnlyProperties.ToArray(), "name");
        }

        [TestMethod, TestCategory(Category)]
        public void FormXmlDifferenceIsExcluded()
        {
            var result = CompareDefinitions(Form(Guid.NewGuid(), "new_Main", "account", 2, 1,
                    formXml: "<first />"),
                Form(Guid.NewGuid(), "new_Main", "account", 2, 1, formXml: "<second />"));
            Assert.AreEqual(ComponentDetailComparisonStatus.Match, result.Status);
            CollectionAssert.Contains(ComponentDefinitionContractCatalog.For(
                ComponentSemanticKinds.SystemForm).AuditOnlyProperties.ToArray(), "formxml");
        }

        [TestMethod, TestCategory(Category)]
        public void IncompleteDefinitionCannotProduceMatch()
        {
            var source = Form(Guid.NewGuid(), "new_Main", "account", 2, 1);
            source.Attributes.Remove("formactivationstate");
            Assert.AreEqual(ComponentDetailComparisonStatus.Unresolved,
                CompareDefinitions(source, Form(Guid.NewGuid(), "new_Main", "account", 2, 1)).Status);
        }

        [TestMethod, TestCategory(Category)]
        public void RetrievalFaultIsUnresolved()
        {
            var solution = Solution("DEV"); var raw = Raw(Guid.NewGuid());
            var fault = Resolve(solution, new[] { raw }, query => throw new FaultException("denied"));
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, fault.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void CancellationPropagates()
        {
            var solution = Solution("DEV"); var raw = Raw(Guid.NewGuid());
            using (var cancellation = new CancellationTokenSource())
            {
                var service = Service(solution, query => { cancellation.Cancel(); return Rows(); });
                Assert.ThrowsException<OperationCanceledException>(() => Resolver().ResolveSnapshot(service,
                    MembershipSnapshot.Complete(solution, new[] { raw }, DateTimeOffset.UtcNow),
                    cancellation.Token));
            }
        }

        [TestMethod, TestCategory(Category)]
        public void BatchedRetrievalHasNoPerFormRequestPatternAndFeedsDefinitionCache()
        {
            var solution = Solution("DEV");
            var rows = Enumerable.Range(0, 201).Select(index =>
                Form(Guid.NewGuid(), "new_Form" + index, "account", 2, 1)).ToArray();
            var counter = new DataverseRequestCounter();
            var service = Service(solution, query => Rows(rows.Where(row => Requested(query).Contains(row.Id)).ToArray()));
            var context = new DataverseReadContext(service, solution.Environment, CancellationToken.None, counter);
            var raw = MembershipSnapshot.Complete(solution, rows.Select(row => Raw(row.Id)), DateTimeOffset.UtcNow);
            var resolved = Resolver().ResolveSnapshot(context, raw, CancellationToken.None);
            var definitions = new DataverseComponentDefinitionReader().Read(context, resolved, CancellationToken.None);
            Assert.AreEqual(2, counter.GetQueryCount("systemform"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.IsTrue(definitions.Definitions.All(item => item.Status == ComponentDefinitionReadStatus.Available));
            Assert.AreEqual(0, service.WriteCalls);
        }

        [TestMethod, TestCategory(Category)]
        public void RepeatedRawMembershipForSameFormProducesOneComparisonFinding()
        {
            var formId = Guid.NewGuid();
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var left = Resolve(source, new[] { Raw(formId), Raw(formId) }, query =>
                Rows(Form(formId, " ", "account", 2, 1)));
            var right = Resolve(target, new[] { Raw(formId), Raw(formId) }, query =>
                Rows(Form(formId, " ", "account", 2, 1)));
            var result = new SolutionMembershipComparer().Compare(left, right);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Single().Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void ExistingSupportedFamiliesRemainUnchanged()
        {
            Assert.AreEqual(ComponentSemanticKinds.Table, ComponentSemanticKinds.FromRawComponentType(1));
            Assert.AreEqual(ComponentSemanticKinds.SystemForm, ComponentSemanticKinds.FromRawComponentType(60));
            Assert.IsNotNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Table));
            Assert.IsNotNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Column));
            Assert.IsNotNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.Relationship));
            Assert.IsNotNull(ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.EntityKey));
        }

        [TestMethod, TestCategory(Category)]
        public void PublicResolverSupportsSystemForm()
        {
            var solution = Solution("DEV");
            var form = Form(Guid.NewGuid(), "new_Main", "account", 2, 1);
            var resolved = new DataverseComponentIdentityResolver().Resolve(Service(solution, q => Rows(form)),
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 60, form.Id), CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, resolved.Status);
            Assert.AreEqual(ComponentSemanticKinds.SystemForm, resolved.SemanticKind);
            Assert.IsNotNull(resolved.ComparisonKey);
        }

        [TestMethod, TestCategory(Category)]
        public void DefaultDefinitionOperationSupportsSystemForm()
        {
            var solution = Solution("DEV"); var form = Form(Guid.NewGuid(), "new_Main", "account", 2, 1);
            var defaultResult = new DataverseComponentDefinitionOperation().ReadAndResolve(
                MembershipService(solution, form), solution.Environment.DisplayName, solution.UniqueName,
                CancellationToken.None, null);
            Assert.AreEqual(IdentityResolutionStatus.Resolved,
                defaultResult.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available,
                defaultResult.Definitions.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void XrmToolBoxCompareMembershipCompositionSupportsSystemFormInEveryBuild()
        {
            var source = Solution("DEV");
            var target = Solution("UAT", source.UniqueName);
            var sourceForm = Form(Guid.NewGuid(), "new_Main", "account", 2, 1);
            var targetForm = Form(Guid.NewGuid(), "NEW_MAIN", "ACCOUNT", 2, 1);
            var operation = SolutionComparerControl.CreateMembershipComparisonOperation();

            var sourceResult = operation.ReadAndResolve(MembershipService(source, sourceForm),
                source.Environment.DisplayName, source.UniqueName, CancellationToken.None, null);
            var targetResult = operation.ReadAndResolve(MembershipService(target, targetForm),
                target.Environment.DisplayName, target.UniqueName, CancellationToken.None, null);
            var presentation = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", sourceResult.Membership, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", targetResult.Membership, 0, TimeSpan.Zero));

            Assert.AreEqual(IdentityResolutionStatus.Resolved,
                sourceResult.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentSemanticKinds.SystemForm,
                sourceResult.Membership.Components.Single().SemanticKind);
            Assert.AreEqual(IdentityResolutionStatus.Resolved,
                targetResult.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentSemanticKinds.SystemForm,
                targetResult.Membership.Components.Single().SemanticKind);
            Assert.AreEqual(1, presentation.Summary.PresentInBoth);
            Assert.AreEqual(0, presentation.Summary.Unsupported);
            Assert.AreEqual("System Form", presentation.Rows.Single().ComponentKind);
            Assert.AreEqual(1, new MembershipCoverageDiagnosticsBuilder().Build(
                sourceResult.Membership).SemanticKinds.Single(item =>
                    item.SemanticKind == ComponentSemanticKinds.SystemForm).Resolved);
        }

        [TestMethod, TestCategory(Category)]
        public void SupportedCoverageSeedIncludesEmptySystemFormBucket()
        {
            var solution = Solution("DEV");
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(solution, new ComponentIdentity[0], DateTimeOffset.UtcNow));
            var systemForms = coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.SystemForm);
            Assert.AreEqual("System Form", systemForms.DisplayName);
            Assert.AreEqual(0, systemForms.TotalCandidates);
            Assert.AreEqual(MembershipCoverageStatus.Complete, systemForms.CoverageStatus);
        }

        private static DataverseComponentIdentityResolver Resolver() =>
            new DataverseComponentIdentityResolver();

        private static SolutionIdentity Solution(string name, string uniqueName = "test_solution") =>
            new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), name), Guid.NewGuid(), uniqueName);

        private static ComponentIdentity Raw(Guid objectId) => new ComponentIdentity(
            new SolutionComponentRecord(Guid.NewGuid(), 60, objectId), IdentityResolutionStatus.Unresolved);

        private static ComponentIdentity ResolveOne(SolutionIdentity solution, Entity row) =>
            Resolve(solution, new[] { Raw(row.Id) }, query => Rows(row)).Components.Single();

        private static MembershipSnapshot Resolve(SolutionIdentity solution, IEnumerable<Entity> rows)
        {
            var list = rows.ToList();
            return Resolve(solution, list.Select(row => Raw(row.Id)), query =>
                Rows(list.Where(row => Requested(query).Contains(row.Id)).ToArray()));
        }

        private static MembershipSnapshot Resolve(SolutionIdentity solution,
            IEnumerable<ComponentIdentity> records, Func<QueryExpression, EntityCollection> query) =>
            Resolver().ResolveSnapshot(Service(solution, query),
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow), CancellationToken.None);

        private static IReadOnlyList<MembershipCompareResult> Compare(SolutionIdentity source,
            SolutionIdentity target, ComponentIdentity sourceIdentity, ComponentIdentity targetIdentity) =>
            new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(source, new[] { sourceIdentity }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(target, new[] { targetIdentity }, DateTimeOffset.UtcNow));

        private static ComponentDetailCompareResult CompareDefinitions(Entity sourceRow, Entity targetRow)
        {
            var source = Solution("DEV"); var target = Solution("UAT", source.UniqueName);
            var sourceFixture = Definitions(source, sourceRow);
            var targetFixture = Definitions(target, targetRow);
            var membership = new SolutionMembershipComparer().Compare(sourceFixture.Membership,
                targetFixture.Membership);
            return new ComponentDetailComparer().Compare(membership, sourceFixture, targetFixture).Single();
        }

        private static ComponentDefinitionSnapshot Definitions(SolutionIdentity solution, Entity row)
        {
            var context = new DataverseReadContext(Service(solution, query => Rows(row)),
                solution.Environment, CancellationToken.None);
            var raw = MembershipSnapshot.Complete(solution, new[] { Raw(row.Id) }, DateTimeOffset.UtcNow);
            var resolved = Resolver().ResolveSnapshot(context, raw, CancellationToken.None);
            return new DataverseComponentDefinitionReader().Read(context, resolved, CancellationToken.None);
        }

        private static Entity Form(Guid id, string uniqueName, object objectTypeCode,
            int type, int activationState, string name = "Form", bool isManaged = false,
            string formXml = "<form />")
        {
            var row = new Entity("systemform", id)
            {
                ["formid"] = id,
                ["uniquename"] = uniqueName,
                ["name"] = name,
                ["objecttypecode"] = objectTypeCode,
                ["type"] = new OptionSetValue(type),
                ["formactivationstate"] = new OptionSetValue(activationState),
                ["formidunique"] = Guid.NewGuid(),
                ["componentstate"] = new OptionSetValue(0),
                ["ismanaged"] = isManaged,
                ["formxml"] = formXml
            };
            return row;
        }

        private static FakeOrganizationService Service(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> query)
        {
            return new FakeOrganizationService
            {
                RetrievePage = expression =>
                {
                    Assert.AreEqual("systemform", expression.EntityName);
                    CollectionAssert.AreEquivalent(new[] { "formid", "uniquename", "name", "objecttypecode",
                        "type", "formactivationstate", "formidunique", "componentstate", "ismanaged" },
                        expression.ColumnSet.Columns.ToArray());
                    return query(expression);
                },
                ExecuteRequest = request => request is WhoAmIRequest
                    ? (OrganizationResponse)new WhoAmIResponse
                    {
                        Results = new ParameterCollection { ["OrganizationId"] = solution.Environment.OrganizationId }
                    }
                    : throw new NotSupportedException(request.RequestName)
            };
        }

        private static FakeOrganizationService MembershipService(SolutionIdentity solution, Entity form)
        {
            return new FakeOrganizationService
            {
                RetrievePage = query =>
                {
                    if (query.EntityName == "solution")
                        return Rows(new Entity("solution", solution.SolutionId)
                        {
                            ["solutionid"] = solution.SolutionId,
                            ["uniquename"] = solution.UniqueName,
                            ["friendlyname"] = solution.UniqueName
                        });
                    if (query.EntityName == "solutioncomponent")
                        return Rows(new Entity("solutioncomponent", Guid.NewGuid())
                        {
                            ["solutioncomponentid"] = Guid.NewGuid(),
                            ["componenttype"] = new OptionSetValue(60),
                            ["objectid"] = form.Id
                        });
                    if (query.EntityName == "systemform") return Rows(form);
                    throw new InvalidOperationException(query.EntityName);
                },
                ExecuteRequest = request => request is WhoAmIRequest
                    ? (OrganizationResponse)new WhoAmIResponse
                    {
                        Results = new ParameterCollection { ["OrganizationId"] = solution.Environment.OrganizationId }
                    }
                    : throw new NotSupportedException(request.RequestName)
            };
        }

        private static HashSet<Guid> Requested(QueryExpression query) => new HashSet<Guid>(
            query.Criteria.Conditions.Single().Values.Cast<Guid>());

        private static EntityCollection Rows(params Entity[] rows) =>
            new EntityCollection(rows.ToList());
    }
}
