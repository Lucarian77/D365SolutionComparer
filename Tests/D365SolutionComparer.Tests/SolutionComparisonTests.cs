using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models;
using D365SolutionComparer.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class SolutionComparisonTests
    {
        private static SolutionInfo Solution(string key = "sample") => new SolutionInfo
        {
            UniqueName = key, DisplayName = "Sample", Version = "1.0.0.0", Publisher = "Publisher", IsManaged = true
        };

        private static SolutionCompareResult Compare(SolutionInfo source, SolutionInfo target) =>
            new SolutionComparisonService().Compare(new List<SolutionInfo> { source }, new List<SolutionInfo> { target }).Single();

        [DataTestMethod]
        [DataRow(0, "Match")]
        [DataRow(1, "Version Mismatch")]
        [DataRow(2, "Publisher Mismatch")]
        [DataRow(3, "Multiple Differences")]
        [DataRow(4, "Display Name Mismatch")]
        [DataRow(5, "Multiple Differences")]
        [DataRow(6, "Multiple Differences")]
        [DataRow(7, "Multiple Differences")]
        [DataRow(8, "Managed/Unmanaged Mismatch")]
        [DataRow(9, "Multiple Differences")]
        [DataRow(10, "Multiple Differences")]
        [DataRow(11, "Multiple Differences")]
        [DataRow(12, "Multiple Differences")]
        [DataRow(13, "Multiple Differences")]
        [DataRow(14, "Multiple Differences")]
        [DataRow(15, "Multiple Differences")]
        public void AllFieldDifferenceCombinationsPreserveBaselineStatuses(int mask, string expected)
        {
            var target = Solution();
            if ((mask & 1) != 0) target.Version = "2.0.0.0";
            if ((mask & 2) != 0) target.Publisher = "Different";
            if ((mask & 4) != 0) target.DisplayName = "Different";
            if ((mask & 8) != 0) target.IsManaged = false;
            var result = Compare(Solution(), target);
            Assert.AreEqual(expected, result.Status);
            Assert.AreEqual((mask & 8) != 0 ? "Managed/Unmanaged Mismatch" : "Match", result.PackageTypeStatus);
            Assert.AreEqual((mask & 8) != 0, result.IsManagedUnmanagedMismatch);
        }

        [TestMethod]
        public void UniqueNamesIgnoreCaseAndPreserveSourceSpelling()
        {
            var result = Compare(Solution("Sample"), Solution("sAMPLE"));
            Assert.AreEqual("Sample", result.UniqueName);
            Assert.AreEqual("Match", result.Status);
        }

        [TestMethod]
        public void UniqueNamesAreNotTrimmedAndDisplayNamesDoNotEstablishIdentity()
        {
            var results = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution("sample ") }, new List<SolutionInfo> { Solution("sample") });
            CollectionAssert.AreEquivalent(new[] { "Missing in Source", "Missing in Target" }, results.Select(r => r.Status).ToArray());
        }

        [TestMethod]
        public void ComparedFieldsIgnoreCaseAndSurroundingWhitespaceWithoutChangingDisplayedValues()
        {
            var target = Solution();
            target.Version = " 1.0.0.0 "; target.Publisher = " PUBLISHER "; target.DisplayName = " SAMPLE ";
            var result = Compare(Solution(), target);
            Assert.AreEqual("Match", result.Status);
            Assert.AreEqual(" PUBLISHER ", result.TargetPublisher);
            Assert.AreEqual(" SAMPLE ", result.TargetDisplayName);
            Assert.AreEqual(" 1.0.0.0 ", result.TargetVersion);
        }

        [TestMethod]
        public void VersionsRemainStringComparisons()
        {
            var target = Solution(); target.Version = "1.0";
            Assert.AreEqual("Version Mismatch", Compare(Solution(), target).Status);
        }

        [TestMethod]
        public void NullAndWhitespaceMetadataCompareEqually()
        {
            var source = Solution(); var target = Solution();
            source.Version = null; source.Publisher = null; source.DisplayName = null;
            target.Version = " "; target.Publisher = "\t"; target.DisplayName = "";
            var result = Compare(source, target);
            Assert.AreEqual("Match", result.Status);
            Assert.AreEqual(string.Empty, result.SourcePublisher);
        }

        [TestMethod]
        public void DuplicateUniqueNamesKeepFirstRecordOnEachSide()
        {
            var source = Solution("sample"); source.Version = "source first";
            var target = Solution("SAMPLE"); target.Version = "target first";
            var result = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { source, Solution("SAMPLE") },
                new List<SolutionInfo> { target, Solution("sample") }).Single();
            Assert.AreEqual("source first", result.SourceVersion);
            Assert.AreEqual("target first", result.TargetVersion);
        }

        [TestMethod]
        public void NullAndEmptyUniqueNamesCollapseToTheSameKey()
        {
            var result = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution(null), Solution("") }, new List<SolutionInfo> { Solution("") }).Single();
            Assert.AreEqual(string.Empty, result.UniqueName);
            Assert.AreEqual("Match", result.Status);
        }

        [TestMethod]
        public void NullOrEmptyInventoriesProduceNoRows()
        {
            var comparer = new SolutionComparisonService();
            Assert.AreEqual(0, comparer.Compare(null, null).Count);
            Assert.AreEqual(0, comparer.Compare(new List<SolutionInfo>(), new List<SolutionInfo>()).Count);
        }

        [DataTestMethod]
        [DataRow(true, "Missing in Target")]
        [DataRow(false, "Missing in Source")]
        public void OneSidedInventoriesRetainValuesAndBothMissingStatuses(bool sourcePresent, string status)
        {
            var items = new List<SolutionInfo> { Solution() };
            var result = new SolutionComparisonService().Compare(sourcePresent ? items : null, sourcePresent ? null : items).Single();
            Assert.AreEqual(status, result.Status);
            Assert.AreEqual(status, result.PackageTypeStatus);
            Assert.AreEqual(sourcePresent ? "Sample" : "", result.SourceDisplayName);
            Assert.AreEqual(sourcePresent ? "" : "Sample", result.TargetDisplayName);
            Assert.AreEqual(sourcePresent ? "Managed" : "", result.SourcePackageType);
            Assert.AreEqual(sourcePresent ? "" : "Managed", result.TargetPackageType);
        }

        [TestMethod]
        public void UnionIsSortedOrdinalIgnoreCase()
        {
            var results = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution("z"), Solution("B") }, new List<SolutionInfo> { Solution("a") });
            CollectionAssert.AreEqual(new[] { "a", "B", "z" }, results.Select(r => r.UniqueName).ToArray());
        }

        [TestMethod]
        public void NullableManagedStatePreservesExistingUnknownBehavior()
        {
            var source = Solution(); source.IsManaged = null;
            Assert.AreEqual("", source.PackageType);
            Assert.AreEqual("Managed/Unmanaged Mismatch", Compare(source, Solution()).Status);
            var target = Solution(); target.IsManaged = null;
            Assert.AreEqual("Match", Compare(source, target).Status);
            target.IsManaged = false;
            Assert.AreEqual("Unmanaged", target.PackageType);
            Assert.AreEqual("Managed/Unmanaged Mismatch", Compare(source, target).Status);
        }

        [DataTestMethod]
        [DataRow("managed/unmanaged mismatch", true)]
        [DataRow(" Managed/Unmanaged Mismatch ", false)]
        [DataRow("Package Type Mismatch", false)]
        [DataRow(null, false)]
        public void ModelMismatchAliasesKeepExactBaselineSemantics(string status, bool expected)
        {
            var result = new SolutionCompareResult { PackageTypeStatus = status };
            Assert.AreEqual(expected, result.IsManagedUnmanagedMismatch);
            Assert.AreEqual(expected, result.IsPackageTypeMismatch);
        }
        // D4-T6B: synthetic selected fixtures; no live service or environment-variable values.
        [DataTestMethod, TestCategory("D4T6BSolutionVisibility"), DataRow(false), DataRow(true)]
        public void T6BAllEligibleSolutionPagesSurviveWithNamesPublisherAndManagedState(bool terminalCookie)
        {
            var first = T6BSolutionRow("synthetic_first", "Different friendly label", true);
            var second = T6BSolutionRow("EnvironmentVariables", "Synthetic Variables", false);
            var pages = new List<int>();
            var fake = new FakeOrganizationService { RetrievePage = query =>
            {
                T6BAssertQuery(query); pages.Add(query.PageInfo.PageNumber);
                if (query.PageInfo.PageNumber == 1)
                {
                    Assert.IsNull(query.PageInfo.PagingCookie);
                    return new Microsoft.Xrm.Sdk.EntityCollection(new[] { first }) { MoreRecords = true, PagingCookie = "synthetic-1" };
                }
                Assert.AreEqual("synthetic-1", query.PageInfo.PagingCookie);
                return new Microsoft.Xrm.Sdk.EntityCollection(new[] { second }) { MoreRecords = false, PagingCookie = terminalCookie ? "synthetic-terminal" : null };
            }};
            var rows = new DataverseSolutionService().GetSolutions(fake);
            CollectionAssert.AreEqual(new[] { 1, 2 }, pages);
            CollectionAssert.AreEqual(new[] { "synthetic_first", "EnvironmentVariables" }, rows.Select(r => r.UniqueName).ToArray());
            CollectionAssert.AreEqual(new[] { "Different friendly label", "Synthetic Variables" }, rows.Select(r => r.DisplayName).ToArray());
            Assert.AreEqual("1.2.3.4", rows[1].Version); Assert.AreEqual("Synthetic publisher", rows[1].Publisher);
            Assert.AreEqual("Managed", rows[0].PackageType); Assert.AreEqual("Unmanaged", rows[1].PackageType);
            Assert.AreEqual(2, fake.Calls); Assert.AreEqual(0, fake.ExecuteCalls + fake.WriteCalls);
        }

        [DataTestMethod, TestCategory("D4T6BSolutionVisibility")]
        [DataRow("Fault"), DataRow("MissingCookie"), DataRow("RepeatedCookie"), DataRow("EmptyContinuingPage"), DataRow("NullResponse")]
        public void T6BIncompleteSolutionRetrievalThrowsInsteadOfReturningSuccessfulPartialInventory(string problem)
        {
            var fake = new FakeOrganizationService { RetrievePage = query =>
            {
                T6BAssertQuery(query);
                if (problem == "NullResponse") return null;
                if (query.PageInfo.PageNumber == 1)
                    return new Microsoft.Xrm.Sdk.EntityCollection(new[] { T6BSolutionRow("first", "First", true) })
                        { MoreRecords = true, PagingCookie = problem == "MissingCookie" ? null : "synthetic-cookie" };
                if (problem == "Fault") throw new System.ServiceModel.FaultException("Synthetic page fault");
                return new Microsoft.Xrm.Sdk.EntityCollection(problem == "EmptyContinuingPage" ? new Microsoft.Xrm.Sdk.Entity[0] :
                    new[] { T6BSolutionRow("second", "Second", false) }) { MoreRecords = true, PagingCookie = "synthetic-cookie" };
            }};
            Exception error = null;
            try { new DataverseSolutionService().GetSolutions(fake); } catch (Exception ex) { error = ex; }
            Assert.IsNotNull(error); Assert.IsTrue(fake.Calls <= 2); Assert.AreEqual(0, fake.ExecuteCalls + fake.WriteCalls);
        }

        [TestMethod, TestCategory("D4T6BSolutionVisibility")]
        public void T6BVisibilityIsServerPredicateAndAbsentRecordsAreNeverFabricated()
        {
            var hidden = T6BSolutionRow("EnvironmentVariables", "Synthetic hidden", true); hidden["isvisible"] = false;
            var visible = T6BSolutionRow("different_unique_name", "EnvironmentVariables", false); visible["isvisible"] = true;
            var fixture = new[] { hidden, visible };
            var fake = new FakeOrganizationService { RetrievePage = query =>
            {
                T6BAssertQuery(query);
                return new Microsoft.Xrm.Sdk.EntityCollection(fixture.Where(r => (bool)r["isvisible"]).ToArray());
            }};
            var rows = new DataverseSolutionService().GetSolutions(fake);
            Assert.AreEqual(1, rows.Count); Assert.AreEqual("different_unique_name", rows[0].UniqueName);
            Assert.AreEqual("EnvironmentVariables", rows[0].DisplayName);
            Assert.IsFalse(rows.Any(r => r.UniqueName == "EnvironmentVariables"));
            var compared = new SolutionComparisonService().Compare(rows, new List<SolutionInfo>());
            Assert.AreEqual("Missing in Target", compared.Single().Status);
            fake.RetrievePage = query => { T6BAssertQuery(query); return new Microsoft.Xrm.Sdk.EntityCollection(); };
            Assert.AreEqual(0, new DataverseSolutionService().GetSolutions(fake).Count);
            Assert.AreEqual(0, fake.ExecuteCalls + fake.WriteCalls);
        }

        [DataTestMethod, TestCategory("D4T6BComponentCoverage")]
        [DataRow("Valid"), DataRow("BlankSchema"), DataRow("MissingBacking"), DataRow("DuplicateBacking")]
        public void T6BEdULocalDefinitionsAndRawValuesRemainDistinctWithoutValueReads(string scenario)
        {
            var leftSolution = MembershipTestData.Solution("SAA"); var rightSolution = MembershipTestData.Solution("SAA");
            var leftId = Guid.NewGuid(); var rightId = Guid.NewGuid(); var valueId = Guid.NewGuid();
            Func<int, Guid, D365SolutionComparer.Models.Membership.ComponentIdentity> member = (type, id) =>
                new D365SolutionComparer.Models.Membership.ComponentIdentity(new D365SolutionComparer.Models.Membership.SolutionComponentRecord(Guid.NewGuid(), type, id),
                    D365SolutionComparer.Models.Membership.IdentityResolutionStatus.Unresolved);
            var leftRaw = D365SolutionComparer.Models.Membership.MembershipSnapshot.Complete(leftSolution,
                new[] { member(380, leftId), member(381, valueId), member(381, valueId) }, DateTimeOffset.UtcNow);
            var rightRaw = D365SolutionComparer.Models.Membership.MembershipSnapshot.Complete(rightSolution,
                new[] { member(380, rightId), member(381, Guid.NewGuid()) }, DateTimeOffset.UtcNow);
            var source = MembershipTestData.Service(leftSolution, query =>
            {
                T6BAssertDefinitionIdentityQuery(query, leftId);
                var row = new Microsoft.Xrm.Sdk.Entity("environmentvariabledefinition", leftId)
                    { ["environmentvariabledefinitionid"] = leftId, ["schemaname"] = scenario == "BlankSchema" ? " " : "synthetic_setting" };
                return scenario == "MissingBacking" ? MembershipTestData.Rows() : scenario == "DuplicateBacking" ? MembershipTestData.Rows(row, row) : MembershipTestData.Rows(row);
            });
            var target = MembershipTestData.Service(rightSolution, query =>
            {
                T6BAssertDefinitionIdentityQuery(query, rightId);
                return MembershipTestData.Rows(new Microsoft.Xrm.Sdk.Entity("environmentvariabledefinition", rightId)
                    { ["environmentvariabledefinitionid"] = rightId, ["schemaname"] = "synthetic_setting" });
            });
            var resolver = new D365SolutionComparer.Services.Membership.DataverseComponentIdentityResolver();
            var left = resolver.ResolveSnapshot(source, leftRaw, System.Threading.CancellationToken.None);
            var right = resolver.ResolveSnapshot(target, rightRaw, System.Threading.CancellationToken.None);
            var definition = left.Components.Single(c => c.Record.ComponentType == 380);
            Assert.AreEqual(scenario == "Valid" ? D365SolutionComparer.Models.Membership.IdentityResolutionStatus.Resolved :
                scenario == "DuplicateBacking" ? D365SolutionComparer.Models.Membership.IdentityResolutionStatus.Ambiguous :
                D365SolutionComparer.Models.Membership.IdentityResolutionStatus.Unresolved, definition.Status);
            Assert.IsTrue(left.Components.Concat(right.Components).Where(c => c.Record.ComponentType == 381).All(c =>
                c.Status == D365SolutionComparer.Models.Membership.IdentityResolutionStatus.Unsupported && c.ComparisonKey == null));
            var rows = new D365SolutionComparer.Services.Membership.SolutionMembershipComparer().Compare(left, right);
            Assert.AreEqual(scenario == "Valid" ? 1 : 0, rows.Count(r => r.Presence == D365SolutionComparer.Models.Membership.MembershipPresence.PresentInBoth));
            Assert.IsTrue(rows.Where(r => (r.Source ?? r.Target).Record.ComponentType == 381)
                .All(r => r.Presence == D365SolutionComparer.Models.Membership.MembershipPresence.Indeterminate));
            Assert.IsTrue(rows.All(r => r.Presence != D365SolutionComparer.Models.Membership.MembershipPresence.OnlyInSource &&
                r.Presence != D365SolutionComparer.Models.Membership.MembershipPresence.OnlyInTarget));
            Assert.AreEqual(3, new D365SolutionComparer.Services.Membership.MembershipCoverageDiagnosticsBuilder().Build(left).TotalCandidates);
            Assert.AreEqual(2, new D365SolutionComparer.Services.Membership.MembershipCoverageDiagnosticsBuilder().Build(right).TotalCandidates);
            Assert.AreEqual(1, source.Calls); Assert.AreEqual(1, target.Calls);
            Assert.AreEqual(2, source.ExecuteCalls + target.ExecuteCalls); Assert.AreEqual(0, source.WriteCalls + target.WriteCalls);
        }

        [DataTestMethod, TestCategory("D4T6BSolutionVisibility"), DataRow(false), DataRow(true)]
        public void T6BLoadedSideCollectionsBindingReloadAndStatusFiltersPreserveEligibleSolutions(bool targetFirst)
        {
            T6BSta(() =>
            {
                var source = new FakeOrganizationService { RetrievePage = query =>
                {
                    T6BAssertQuery(query);
                    return query.PageInfo.PageNumber == 1
                        ? new Microsoft.Xrm.Sdk.EntityCollection(new[] { T6BSolutionRow("shared", "Shared", true) }) { MoreRecords = true, PagingCookie = "source-page" }
                        : new Microsoft.Xrm.Sdk.EntityCollection(new[] { T6BSolutionRow("EnvironmentVariables", "Synthetic Variables", false) });
                }};
                var target = new FakeOrganizationService { RetrievePage = query =>
                { T6BAssertQuery(query); return new Microsoft.Xrm.Sdk.EntityCollection(new[] { T6BSolutionRow("shared", "Shared", true), T6BSolutionRow("target_only", "Target only", false) }); }};
                using (var host = new System.Windows.Forms.Form())
                using (var control = new SolutionComparerControl())
                {
                    host.Controls.Add(control); host.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    host.Location = new System.Drawing.Point(-4000, -4000); host.Show(); System.Windows.Forms.Application.DoEvents();
                    control.UpdateConnection(source, new T6BConnection { ConnectionName = "Synthetic source", ConnectionId = Guid.NewGuid() }, string.Empty, null);
                    control.UpdateConnection(target, new T6BConnection { ConnectionName = "Synthetic target", ConnectionId = Guid.NewGuid() }, "AdditionalOrganization", null);
                    T6BSet(control, "targetService", target); T6BSet(control, "suppressSettingsSave", true);
                    if (targetFirst) T6BInvoke(control, "BtnLoadTarget_Click", null, EventArgs.Empty);
                    T6BInvoke(control, "BtnLoadSource_Click", null, EventArgs.Empty);
                    T6BInvoke(control, "BtnLoadTarget_Click", null, EventArgs.Empty);
                    var left = T6BField<List<SolutionInfo>>(control, "sourceSolutions"); var right = T6BField<List<SolutionInfo>>(control, "targetSolutions");
                    CollectionAssert.AreEquivalent(new[] { "shared", "EnvironmentVariables" }, left.Select(r => r.UniqueName).ToArray());
                    CollectionAssert.AreEquivalent(new[] { "shared", "target_only" }, right.Select(r => r.UniqueName).ToArray());
                    Assert.AreSame(source, control.Service); Assert.AreSame(target, T6BField<Microsoft.Xrm.Sdk.IOrganizationService>(control, "targetService"));
                    T6BInvoke(control, "BtnLoadSource_Click", null, EventArgs.Empty); Assert.AreSame(right, T6BField<List<SolutionInfo>>(control, "targetSolutions"));
                    var refreshedLeft = T6BField<List<SolutionInfo>>(control, "sourceSolutions");
                    T6BInvoke(control, "BtnLoadTarget_Click", null, EventArgs.Empty); Assert.AreSame(refreshedLeft, T6BField<List<SolutionInfo>>(control, "sourceSolutions"));
                    T6BInvoke(control, "SetNonAllItemsChecked", false);
                    T6BField<System.Windows.Forms.ToolStripMenuItem>(control, "miAll").Checked = true;
                    T6BField<System.Windows.Forms.CheckBox>(control, "chkChangedOnly").Checked = false;
                    T6BField<System.Windows.Forms.CheckBox>(control, "chkPackageTypeMismatchOnly").Checked = false;
                    T6BInvoke(control, "BtnCompare_Click", null, EventArgs.Empty);
                    var grid = T6BField<System.Windows.Forms.DataGridView>(control, "dgvResults");
                    var compared = (List<SolutionCompareResult>)grid.DataSource;
                    CollectionAssert.AreEqual(new[] { "EnvironmentVariables", "shared", "target_only" }, compared.Select(r => r.UniqueName).ToArray());
                    Assert.AreEqual("Synthetic Variables", compared[0].SourceDisplayName); Assert.AreEqual("Missing in Target", compared[0].Status);
                    T6BField<System.Windows.Forms.ToolStripMenuItem>(control, "miMissingInTarget").Checked = true;
                    Assert.AreEqual("EnvironmentVariables", ((List<SolutionCompareResult>)grid.DataSource).Single().UniqueName);
                    T6BField<System.Windows.Forms.ToolStripMenuItem>(control, "miAll").Checked = true;
                    Assert.AreEqual(3, ((List<SolutionCompareResult>)grid.DataSource).Count);
                    T6BField<System.Windows.Forms.CheckBox>(control, "chkChangedOnly").Checked = true;
                    Assert.AreEqual(2, ((List<SolutionCompareResult>)grid.DataSource).Count);
                    Assert.AreEqual(4, source.Calls); Assert.AreEqual(targetFirst ? 3 : 2, target.Calls);
                    Assert.AreEqual(0, source.ExecuteCalls + target.ExecuteCalls + source.WriteCalls + target.WriteCalls);
                    host.Close();
                }
            });
        }

        private static Microsoft.Xrm.Sdk.Entity T6BSolutionRow(string unique, string friendly, bool managed) =>
            new Microsoft.Xrm.Sdk.Entity("solution", Guid.NewGuid()) { ["uniquename"] = unique, ["friendlyname"] = friendly,
                ["version"] = "1.2.3.4", ["ismanaged"] = managed,
                ["publisherid"] = new Microsoft.Xrm.Sdk.EntityReference("publisher", Guid.NewGuid()) { Name = "Synthetic publisher" } };
        private static void T6BAssertQuery(Microsoft.Xrm.Sdk.Query.QueryExpression query)
        {
            Assert.AreEqual("solution", query.EntityName); Assert.IsTrue(query.NoLock);
            CollectionAssert.AreEquivalent(new[] { "uniquename", "friendlyname", "version", "publisherid", "ismanaged" }, query.ColumnSet.Columns.ToArray());
            var filter = query.Criteria.Conditions.Single(); Assert.AreEqual("isvisible", filter.AttributeName);
            Assert.AreEqual(Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, filter.Operator); Assert.AreEqual(true, filter.Values.Single());
            Assert.AreEqual(1, query.Orders.Count); Assert.AreEqual("solutionid", query.Orders[0].AttributeName);
            Assert.AreEqual(Microsoft.Xrm.Sdk.Query.OrderType.Ascending, query.Orders[0].OrderType);
            Assert.AreEqual(5000, query.PageInfo.Count);
        }
        private static void T6BAssertDefinitionIdentityQuery(Microsoft.Xrm.Sdk.Query.QueryExpression query, Guid id)
        {
            Assert.AreEqual("environmentvariabledefinition", query.EntityName);
            CollectionAssert.AreEquivalent(new[] { "environmentvariabledefinitionid", "schemaname",
                "type", "defaultvalue", "valueschema", "isrequired", "secretstore", "displayname", "description" },
                query.ColumnSet.Columns.ToArray());
            Assert.AreEqual(id, query.Criteria.Conditions.Single().Values.Single());
        }
        private sealed class T6BConnection : McTools.Xrm.Connection.ConnectionDetail
        { public override Microsoft.Xrm.Tooling.Connector.CrmServiceClient GetCrmServiceClient(bool forceNewService = false) => null; }
        private static T T6BField<T>(SolutionComparerControl control, string name) => (T)typeof(SolutionComparerControl)
            .GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(control);
        private static void T6BSet(SolutionComparerControl control, string name, object value) => typeof(SolutionComparerControl)
            .GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(control, value);
        private static object T6BInvoke(SolutionComparerControl control, string name, params object[] args) => typeof(SolutionComparerControl)
            .GetMethod(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(control, args);
        private static void T6BSta(Action action)
        {
            Exception error = null;
            var thread = new System.Threading.Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true };
            thread.SetApartmentState(System.Threading.ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(30)), "Synthetic solution-list UI did not complete.");
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
        // End D4-T6B additions.
    }
}
