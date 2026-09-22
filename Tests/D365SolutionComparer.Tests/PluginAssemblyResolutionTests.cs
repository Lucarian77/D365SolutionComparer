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
using Microsoft.Xrm.Sdk.Query;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class PluginAssemblyResolutionTests
    {
        private const string Category = "Phase2G10B1";

        [TestMethod, TestCategory(Category)]
        public void DifferentLocalIdsAndManagementStateUseSamePortableIdentity()
        {
            var source = ResolveOne(Assembly(Guid.NewGuid(), "PCLookupGetToken", "31bf3856ad364e35", "neutral",
                "1.0.0.0", 2, 0, false));
            var target = ResolveOne(Assembly(Guid.NewGuid(), "PCLookupGetToken", "31BF3856AD364E35", "NEUTRAL",
                "1.0.0.0", 2, 0, true));
            Assert.AreNotEqual(source.Record.ObjectId, target.Record.ObjectId);
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(source.ComparisonKey, target.ComparisonKey));
            Assert.AreEqual("pluginassembly:v1:16:PCLookupGetToken:16:31bf3856ad364e35:7:neutral",
                source.ComparisonKey);
            Assert.AreEqual(MembershipPresence.PresentInBoth,
                Compare(source, target).Single().Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void IdentityPartsAreTrimmedAndComparedOrdinalIgnoreCase()
        {
            var first = ResolveOne(Assembly(Guid.NewGuid(), "  Example.Plugin  ", "  ABCD  ", " neutral "));
            var second = ResolveOne(Assembly(Guid.NewGuid(), "example.plugin", "abcd", "NEUTRAL"));
            Assert.IsTrue(StringComparer.OrdinalIgnoreCase.Equals(first.ComparisonKey, second.ComparisonKey));
            Assert.AreEqual("pluginassembly:v1:14:Example.Plugin:4:ABCD:7:neutral", first.ComparisonKey);
        }

        [TestMethod, TestCategory(Category)]
        public void DifferentTokenOrCultureProducesDistinctIdentity()
        {
            var baseline = ResolveOne(Assembly(Guid.NewGuid(), "Example", "token-a", "neutral"));
            var token = ResolveOne(Assembly(Guid.NewGuid(), "Example", "token-b", "neutral"));
            var culture = ResolveOne(Assembly(Guid.NewGuid(), "Example", "token-a", "fr-FR"));
            Assert.AreNotEqual(baseline.ComparisonKey, token.ComparisonKey);
            Assert.AreNotEqual(baseline.ComparisonKey, culture.ComparisonKey);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("", "token", "neutral")]
        [DataRow("Example", "", "neutral")]
        [DataRow("Example", "token", "")]
        public void BlankRequiredIdentityPartRemainsUnresolved(string name, string token, string culture)
        {
            var result = ResolveOne(Assembly(Guid.NewGuid(), name, token, culture));
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.PluginAssembly, result.SemanticKind);
            Assert.IsNull(result.ComparisonKey);
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateCanonicalIdentityIsAmbiguousAndVersionDoesNotDisambiguate()
        {
            var a = Assembly(Guid.NewGuid(), "Microsoft.CDS.PowerAIExtensions.Plugins", "token", "neutral", "1.2");
            var b = Assembly(Guid.NewGuid(), "Microsoft.CDS.PowerAIExtensions.Plugins", "token", "neutral", "1.4");
            var result = Resolve(new[] { a, b });
            Assert.IsTrue(result.Components.All(item => item.Status == IdentityResolutionStatus.Ambiguous));
            Assert.IsTrue(result.Components.All(item => item.BlockerScope == ResolutionBlockerScope.PortableIdentity));
            Assert.AreEqual(1, result.Components.Select(item => item.BlockerPortableIdentity)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicateAmbiguityIsScopedAndDoesNotBlockUnrelatedAbsence()
        {
            var duplicateA = Assembly(Guid.NewGuid(), "BusinessCopilot", "token", "neutral", "1");
            var duplicateB = Assembly(Guid.NewGuid(), "BusinessCopilot", "token", "neutral", "2");
            var unique = Assembly(Guid.NewGuid(), "Unique.Plugin", "token", "neutral", "1");
            var source = Resolve(new[] { duplicateA, duplicateB, unique });
            var target = Resolve(new[] { Assembly(Guid.NewGuid(), "Other.Plugin", "token", "neutral", "1") },
                "UAT", source.SolutionUniqueName);
            var compared = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(MembershipPresence.OnlyInSource, compared.Single(item =>
                item.Source != null && item.Source.ComparisonKey != null &&
                item.Source.ComparisonKey.Contains("Unique.Plugin")).Presence);
            Assert.IsTrue(compared.Where(item => (item.Source ?? item.Target).Status ==
                IdentityResolutionStatus.Ambiguous).All(item => item.Presence == MembershipPresence.Indeterminate));
        }

        [TestMethod, TestCategory(Category)]
        public void RepeatedRawMembershipForSameAssemblyProducesOneFinding()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var source = Resolve(new[] { row, row });
            var target = Resolve(new[] { Clone(row) }, "UAT", source.SolutionUniqueName);
            var compared = new SolutionMembershipComparer().Compare(source, target);
            Assert.AreEqual(1, compared.Count);
            Assert.AreEqual(MembershipPresence.PresentInBoth, compared.Single().Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void MissingBackingRowRemainsUnresolved()
        {
            var id = Guid.NewGuid();
            var result = ResolveRecords(new[] { Raw(id) }, query => Rows());
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void DuplicatePrimaryKeyCorrelationIsAmbiguous()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var result = ResolveRecords(new[] { Raw(row.Id) }, query => Rows(row, Clone(row)));
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void ConflictingPrimaryKeyCorrelationIsAmbiguous()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            row["pluginassemblyid"] = Guid.NewGuid();
            var result = ResolveRecords(new[] { Raw(row.Id) }, query => Rows(row));
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void UnexpectedPagingIsAmbiguous()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var result = ResolveRecords(new[] { Raw(row.Id) }, query =>
                new EntityCollection(new List<Entity> { row }) { MoreRecords = true });
            Assert.AreEqual(IdentityResolutionStatus.Ambiguous, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void QueryFaultRemainsUnresolved()
        {
            var id = Guid.NewGuid();
            var result = ResolveRecords(new[] { Raw(id) }, query => throw new FaultException("denied"));
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, result.Components.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void CancellationPropagates()
        {
            var id = Guid.NewGuid();
            using (var source = new CancellationTokenSource())
            {
                var solution = Solution("DEV");
                var service = Service(solution, query => { source.Cancel(); return Rows(); });
                Assert.ThrowsException<OperationCanceledException>(() => EnabledResolver().ResolveSnapshot(service,
                    MembershipSnapshot.Complete(solution, new[] { Raw(id) }, DateTimeOffset.UtcNow), source.Token));
            }
        }

        [TestMethod, TestCategory(Category)]
        public void TwoHundredIdsUseOneBatchAndTwoHundredOneUseTwoBatches()
        {
            Assert.AreEqual(1, BatchQueryCount(200));
            Assert.AreEqual(2, BatchQueryCount(201));
        }

        [TestMethod, TestCategory(Category)]
        public void QueryShapeUsesGuidValuesAndNeverRequestsBinaryContent()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            QueryExpression captured = null;
            ResolveRecords(new[] { Raw(row.Id) }, query => { captured = query; return Rows(row); });
            Assert.AreEqual("pluginassembly", captured.EntityName);
            CollectionAssert.AreEquivalent(new[] { "pluginassemblyid", "pluginassemblyidunique", "name",
                "publickeytoken", "culture", "version", "isolationmode", "sourcetype", "ismanaged",
                "componentstate" }, captured.ColumnSet.Columns.ToArray());
            Assert.IsFalse(captured.ColumnSet.Columns.Contains("content", StringComparer.OrdinalIgnoreCase));
            Assert.IsTrue(captured.Criteria.Conditions.Single().Values.All(value => value is Guid));
            Assert.IsTrue(captured.Criteria.Conditions.Single().Values.Count <= 200);
        }

        [TestMethod, TestCategory(Category)]
        public void DefinitionReadUsesCacheAndAddsNoRequestOrWhoAmI()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral", "1.0", 2, 0);
            var solution = Solution("DEV"); var counter = new DataverseRequestCounter();
            var service = Service(solution, query => Rows(row));
            var context = new DataverseReadContext(service, solution.Environment, CancellationToken.None, counter);
            var membership = EnabledResolver().ResolveSnapshot(context,
                MembershipSnapshot.Complete(solution, new[] { Raw(row.Id) }, DateTimeOffset.UtcNow),
                CancellationToken.None);
            int requests = counter.TotalRequests;
            var definitions = new DataverseComponentDefinitionReader().Read(context, membership,
                CancellationToken.None);
            Assert.AreEqual(requests, counter.TotalRequests);
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
            Assert.AreEqual(1, counter.GetQueryCount("pluginassembly"));
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, definitions.Definitions.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void IdenticalDefinitionMatchesAndManagementOrLocalFieldsAreIgnored()
        {
            var source = Assembly(Guid.NewGuid(), "Example", "token", "neutral", "1.0", 2, 0, false);
            var target = Assembly(Guid.NewGuid(), "EXAMPLE", "TOKEN", "NEUTRAL", "1.0", 2, 0, true);
            target["pluginassemblyidunique"] = Guid.NewGuid();
            target["componentstate"] = new OptionSetValue(3);
            Assert.AreEqual(ComponentDetailComparisonStatus.Match,
                CompareDefinitions(source, target).Status);
        }

        [DataTestMethod, TestCategory(Category)]
        [DataRow("version", "2.0")]
        [DataRow("isolationmode", 3)]
        [DataRow("sourcetype", 1)]
        public void ComparablePropertyDifferenceProducesDifferent(string property, object value)
        {
            var source = Assembly(Guid.NewGuid(), "Example", "token", "neutral", "1.0", 2, 0);
            var target = Assembly(Guid.NewGuid(), "Example", "token", "neutral", "1.0", 2, 0);
            target[property] = value is int ? (object)new OptionSetValue((int)value) : value;
            var result = CompareDefinitions(source, target);
            Assert.AreEqual(ComponentDetailComparisonStatus.Different, result.Status);
            CollectionAssert.Contains(result.Differences.Select(item => item.PropertyName).ToArray(), property);
            Assert.AreEqual(MembershipPresence.PresentInBoth, result.Membership.Presence);
        }

        [TestMethod, TestCategory(Category)]
        public void IncompleteDefinitionNeverProducesMatch()
        {
            var source = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var target = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            target.Attributes.Remove("sourcetype");
            Assert.AreEqual(ComponentDetailComparisonStatus.Unresolved,
                CompareDefinitions(source, target).Status);
        }

        [TestMethod, TestCategory(Category)]
        public void ProductionCatalogAndDefaultResolverSupportType91()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var solution = Solution("DEV");
            var service = Service(solution, query => Rows(row));
            var result = new DataverseComponentIdentityResolver().Resolve(service,
                solution.Environment, new SolutionComponentRecord(Guid.NewGuid(), 91, row.Id),
                CancellationToken.None);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Status);
            Assert.AreEqual(ComponentSemanticKinds.PluginAssembly, result.SemanticKind);
            Assert.AreEqual(ComponentSemanticKinds.PluginAssembly,
                ComponentSemanticKinds.FromRawComponentType(91));
            Assert.AreEqual(1, service.Calls);
        }

        [TestMethod, TestCategory(Category)]
        public void DefaultDefinitionOperationSupportsType91AndReusesBackingRow()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var solution = Solution("DEV"); var counter = new DataverseRequestCounter();
            var result = new DataverseComponentDefinitionOperation().ReadAndResolve(
                MembershipService(solution, row), solution.Environment.DisplayName, solution.UniqueName,
                CancellationToken.None, null, counter);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, result.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available, result.Definitions.Single().Status);
            Assert.AreEqual(1, counter.GetQueryCount("pluginassembly"));
            Assert.AreEqual(1, counter.GetExecuteCount("WhoAmI"));
        }

        [TestMethod, TestCategory(Category)]
        public void XrmToolBoxCompositionSupportsType91InEveryBuild()
        {
            var row = Assembly(Guid.NewGuid(), "PCLookupGetToken", "31bf3856ad364e35", "neutral");
            var solution = Solution("DEV");
            var composed = SolutionComparerControl.CreateMembershipComparisonOperation().ReadAndResolve(
                MembershipService(solution, row), solution.Environment.DisplayName, solution.UniqueName,
                CancellationToken.None, null);
            Assert.AreEqual(IdentityResolutionStatus.Resolved,
                composed.Membership.Components.Single().Status);
            Assert.AreEqual(ComponentSemanticKinds.PluginAssembly,
                composed.Membership.Components.Single().SemanticKind);
            Assert.AreEqual(ComponentDefinitionReadStatus.Available,
                composed.Definitions.Single().Status);
        }

        [TestMethod, TestCategory(Category)]
        public void TemporaryType91ValidationPlumbingIsRemoved()
        {
            Assert.IsNull(typeof(DataverseComponentDefinitionOperation)
                .GetProperty("PluginAssemblyValidationEnabled",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
            Assert.IsNull(typeof(SolutionComparerControl)
                .GetMethod("CreateMembershipCompositionDiagnostic",
                    System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic));
        }

        [TestMethod, TestCategory(Category)]
        public void PresentationAndCoverageUsePlugInAssemblyLabel()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var source = Resolve(new[] { row });
            var target = Resolve(new[] { Clone(row) }, "UAT", source.SolutionUniqueName);
            var presentation = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("DEV", source, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("UAT", target, 0, TimeSpan.Zero));
            Assert.AreEqual("Plug-in Assembly", presentation.Rows.Single().ComponentKind);
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(source)
                .SemanticKinds.Single(item => item.SemanticKind == ComponentSemanticKinds.PluginAssembly);
            Assert.AreEqual("Plug-in Assembly", coverage.DisplayName);
            Assert.AreEqual(1, coverage.Resolved);
            Assert.AreEqual(0, coverage.Unsupported);
            Assert.AreEqual(0, coverage.Unresolved);
            Assert.AreEqual(0, coverage.Ambiguous);
            Assert.AreEqual(MembershipCoverageStatus.Complete, coverage.CoverageStatus);
        }

        [TestMethod, TestCategory(Category)]
        public void SupportedCoverageSeedIncludesEmptyPlugInAssemblyBucket()
        {
            var coverage = new MembershipCoverageDiagnosticsBuilder().Build(
                MembershipSnapshot.Complete(Solution("DEV"), new ComponentIdentity[0],
                    DateTimeOffset.UtcNow));
            var assemblies = coverage.SemanticKinds.Single(item =>
                item.SemanticKind == ComponentSemanticKinds.PluginAssembly);
            Assert.AreEqual("Plug-in Assembly", assemblies.DisplayName);
            Assert.AreEqual(0, assemblies.TotalCandidates);
            Assert.AreEqual(MembershipCoverageStatus.Complete, assemblies.CoverageStatus);
        }

        [TestMethod, TestCategory(Category)]
        public void ExistingSupportedFamiliesAndDefinitionContractsRemainAvailable()
        {
            foreach (var kind in new[] { ComponentSemanticKinds.Table, ComponentSemanticKinds.Column,
                ComponentSemanticKinds.Relationship, ComponentSemanticKinds.WebResource,
                ComponentSemanticKinds.GlobalChoice, ComponentSemanticKinds.EnvironmentVariableDefinition,
                ComponentSemanticKinds.ConnectionReference, ComponentSemanticKinds.AppModule,
                ComponentSemanticKinds.EntityKey, ComponentSemanticKinds.SystemForm })
                Assert.IsNotNull(ComponentDefinitionContractCatalog.For(kind), kind);
            CollectionAssert.AreEquivalent(new[] { "version", "isolationmode", "sourcetype" },
                ComponentDefinitionContractCatalog.For(ComponentSemanticKinds.PluginAssembly)
                    .ComparableProperties.ToArray());
        }

        [TestMethod, TestCategory(Category)]
        public void ResolverIssuesNoDataverseWrites()
        {
            var row = Assembly(Guid.NewGuid(), "Example", "token", "neutral");
            var solution = Solution("DEV"); var service = Service(solution, query => Rows(row));
            EnabledResolver().ResolveSnapshot(service,
                MembershipSnapshot.Complete(solution, new[] { Raw(row.Id) }, DateTimeOffset.UtcNow),
                CancellationToken.None);
            Assert.AreEqual(0, service.WriteCalls);
        }

        private static int BatchQueryCount(int count)
        {
            var rows = Enumerable.Range(0, count).Select(index =>
                Assembly(Guid.NewGuid(), "Assembly" + index, "token", "neutral")).ToList();
            int requests = 0;
            ResolveRecords(rows.Select(row => Raw(row.Id)), query =>
            {
                requests++;
                var ids = Requested(query);
                Assert.IsTrue(ids.Count <= 200);
                return Rows(rows.Where(row => ids.Contains(row.Id)).ToArray());
            });
            return requests;
        }

        private static ComponentIdentity ResolveOne(Entity row) =>
            Resolve(new[] { row }).Components.Single();

        private static MembershipSnapshot Resolve(IEnumerable<Entity> rows, string environment = "DEV",
            string uniqueName = "plugin_solution")
        {
            var list = rows.ToList();
            return ResolveRecords(list.Select(row => Raw(row.Id)), query =>
                Rows(list.Where(row => Requested(query).Contains(row.Id)).GroupBy(row => row.Id)
                    .Select(group => Clone(group.First())).ToArray()),
                environment, uniqueName);
        }

        private static MembershipSnapshot ResolveRecords(IEnumerable<ComponentIdentity> records,
            Func<QueryExpression, EntityCollection> query, string environment = "DEV",
            string uniqueName = "plugin_solution")
        {
            var solution = Solution(environment, uniqueName);
            return EnabledResolver().ResolveSnapshot(Service(solution, query),
                MembershipSnapshot.Complete(solution, records, DateTimeOffset.UtcNow),
                CancellationToken.None);
        }

        private static IReadOnlyList<MembershipCompareResult> Compare(ComponentIdentity source,
            ComponentIdentity target)
        {
            var left = Solution("DEV"); var right = Solution("UAT", left.UniqueName);
            return new SolutionMembershipComparer().Compare(
                MembershipSnapshot.Complete(left, new[] { source }, DateTimeOffset.UtcNow),
                MembershipSnapshot.Complete(right, new[] { target }, DateTimeOffset.UtcNow));
        }

        private static ComponentDetailCompareResult CompareDefinitions(Entity sourceRow, Entity targetRow)
        {
            var sourceSolution = Solution("DEV");
            var targetSolution = Solution("UAT", sourceSolution.UniqueName);
            var source = ResolveDefinitions(sourceSolution, sourceRow);
            var target = ResolveDefinitions(targetSolution, targetRow);
            var membership = new SolutionMembershipComparer().Compare(source.Membership, target.Membership);
            return new ComponentDetailComparer().Compare(membership, source, target).Single();
        }

        private static ComponentDefinitionSnapshot ResolveDefinitions(SolutionIdentity solution, Entity row)
        {
            var service = Service(solution, query => Rows(Clone(row)));
            var context = new DataverseReadContext(service, solution.Environment, CancellationToken.None);
            var membership = EnabledResolver().ResolveSnapshot(context,
                MembershipSnapshot.Complete(solution, new[] { Raw(row.Id) }, DateTimeOffset.UtcNow),
                CancellationToken.None);
            return new DataverseComponentDefinitionReader().Read(context, membership, CancellationToken.None);
        }

        private static DataverseComponentIdentityResolver EnabledResolver() =>
            new DataverseComponentIdentityResolver();

        private static ComponentIdentity Raw(Guid objectId) => new ComponentIdentity(
            new SolutionComponentRecord(Guid.NewGuid(), 91, objectId), IdentityResolutionStatus.Unresolved);

        private static SolutionIdentity Solution(string displayName, string uniqueName = "plugin_solution") =>
            new SolutionIdentity(new EnvironmentIdentity(Guid.NewGuid(), displayName), Guid.NewGuid(), uniqueName);

        private static Entity Assembly(Guid id, string name, string token, string culture,
            string version = "1.0.0.0", int isolationMode = 2, int sourceType = 0,
            bool isManaged = false)
        {
            return new Entity("pluginassembly", id)
            {
                ["pluginassemblyid"] = id,
                ["pluginassemblyidunique"] = Guid.NewGuid(),
                ["name"] = name,
                ["publickeytoken"] = token,
                ["culture"] = culture,
                ["version"] = version,
                ["isolationmode"] = new OptionSetValue(isolationMode),
                ["sourcetype"] = new OptionSetValue(sourceType),
                ["ismanaged"] = isManaged,
                ["componentstate"] = new OptionSetValue(0)
            };
        }

        private static Entity Clone(Entity row)
        {
            var clone = new Entity(row.LogicalName, row.Id);
            foreach (var attribute in row.Attributes) clone[attribute.Key] = attribute.Value;
            return clone;
        }

        private static FakeOrganizationService Service(SolutionIdentity solution,
            Func<QueryExpression, EntityCollection> query) => new FakeOrganizationService
        {
            RetrievePage = query,
            ExecuteRequest = request => request is WhoAmIRequest
                ? (OrganizationResponse)new WhoAmIResponse
                {
                    Results = new ParameterCollection { ["OrganizationId"] = solution.Environment.OrganizationId }
                }
                : throw new NotSupportedException(request.RequestName)
        };

        private static FakeOrganizationService MembershipService(SolutionIdentity solution, Entity assembly) =>
            new FakeOrganizationService
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
                            ["componenttype"] = new OptionSetValue(91),
                            ["objectid"] = assembly.Id
                        });
                    if (query.EntityName == "pluginassembly") return Rows(Clone(assembly));
                    throw new InvalidOperationException(query.EntityName);
                },
                ExecuteRequest = request => request is WhoAmIRequest
                    ? (OrganizationResponse)new WhoAmIResponse
                    {
                        Results = new ParameterCollection { ["OrganizationId"] = solution.Environment.OrganizationId }
                    }
                    : throw new NotSupportedException(request.RequestName)
            };

        private static HashSet<Guid> Requested(QueryExpression query) => new HashSet<Guid>(
            query.Criteria.Conditions.Single().Values.Cast<Guid>());

        private static EntityCollection Rows(params Entity[] rows) =>
            new EntityCollection(rows.ToList());
    }
}
