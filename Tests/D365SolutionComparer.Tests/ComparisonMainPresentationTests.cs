using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using D365SolutionComparer.Models;
using D365SolutionComparer.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComparisonMainPresentationTests
    {
        [DataTestMethod]
        [DataRow("Match", 1), DataRow("Version Mismatch", 2), DataRow("Publisher Mismatch", 2)]
        [DataRow("Display Name Mismatch", 2), DataRow("Managed/Unmanaged Mismatch", 2)]
        [DataRow("Package Type Mismatch", 2), DataRow("Multiple Differences", 2)]
        [DataRow("Missing in Source", 3), DataRow("Missing in Target", 3)]
        public void EveryEstablishedOverallStatusOccupiesExactlyOneBucket(string status, int bucket)
        {
            var row = new SolutionCompareResult { Status = status };
            var counts = SolutionComparerControl.MainComparisonDashboardCounts(new[] { row });
            Assert.AreEqual(1, counts[0]); Assert.AreEqual(1, counts[bucket]); Assert.AreEqual(0, counts[4]);
            Assert.AreEqual(counts[0], counts[1] + counts[2] + counts[3]); Assert.AreEqual(status, row.Status);
        }

        [DataTestMethod, DataRow("mAtCh", 1), DataRow("MULTIPLE DIFFERENCES", 2), DataRow("missing in target", 3)]
        public void DashboardUsesExistingCaseInsensitiveOverallStatusContract(string status, int bucket)
        { Assert.AreEqual(1, SolutionComparerControl.MainComparisonDashboardCounts(new[] { new SolutionCompareResult { Status = status } })[bucket]); }

        [DataTestMethod, DataRow("Match", 1), DataRow("Multiple Differences", 2)]
        public void OverlappingPackageAndPublisherEvidenceCannotDoubleCount(string status, int bucket)
        {
            var row = new SolutionCompareResult { Status = status, PackageTypeStatus = "Managed/Unmanaged Mismatch",
                SourcePackageType = "Unmanaged", TargetPackageType = "Managed", SourcePublisher = "A", TargetPublisher = "B" };
            var before = Snapshot(row); var counts = SolutionComparerControl.MainComparisonDashboardCounts(new[] { row });
            Assert.AreEqual(1, counts[bucket]); Assert.AreEqual(1, counts.Skip(1).Sum()); Assert.AreEqual(before, Snapshot(row));
        }

        [DataTestMethod, DataRow(null), DataRow(""), DataRow("Unrecognized status")]
        public void UnestablishedStatusIsUnavailableRatherThanInventedDifferent(string status)
        {
            var row = new SolutionCompareResult { Status = status }; var counts = SolutionComparerControl.MainComparisonDashboardCounts(new[] { row });
            CollectionAssert.AreEqual(new[] { 1, 0, 0, 0, 1 }, counts); Assert.AreEqual(status, row.Status);
        }

        [TestMethod]
        public void NullRowIsUnavailableAndNeverRemovedFromTheReportedInventory()
        { CollectionAssert.AreEqual(new[] { 1, 0, 0, 0, 1 }, SolutionComparerControl.MainComparisonDashboardCounts(new SolutionCompareResult[] { null })); }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void NoResultsProducesFourZeroIndicators(bool nullResults)
        { CollectionAssert.AreEqual(new int[5], SolutionComparerControl.MainComparisonDashboardCounts(nullResults ? null : new SolutionCompareResult[0])); }

        [TestMethod]
        public void DashboardEnumeratesTheAuthoritativeCollectionOnlyOnce()
        {
            var input = new SinglePassRows(); CollectionAssert.AreEqual(new[] { 3, 1, 1, 1, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(input));
            Assert.AreEqual(1, input.Enumerations);
        }

        [TestMethod]
        public void FilteredViewsDoNotMutateUnderlyingDashboardTotalsOrStatusText()
        {
            var all = new[] { "Match", "Version Mismatch", "Multiple Differences", "Missing in Source", "Missing in Target" }
                .Select((s, i) => new SolutionCompareResult { UniqueName = "row" + i, Status = s }).ToList();
            var original = all.Select(Snapshot).ToArray();
            var full = SolutionComparerControl.MainComparisonDashboardCounts(all);
            foreach (var subset in new[] { all.Where(r => r.Status == "Match").ToList(), all.Where(r => r.Status.StartsWith("Missing")).ToList(), all.Take(0).ToList() })
            {
                SolutionComparerControl.MainComparisonDashboardCounts(subset);
                CollectionAssert.AreEqual(full, SolutionComparerControl.MainComparisonDashboardCounts(all));
            }
            CollectionAssert.AreEqual(new[] { 5, 1, 2, 2, 0 }, full); CollectionAssert.AreEqual(original, all.Select(Snapshot).ToArray());
        }

        [TestMethod]
        public void ExistingComparisonServiceCountsAndExportInputsSurviveDashboardProjection()
        {
            var source = Enumerable.Range(0, 10000).Select(i => new SolutionInfo { UniqueName = "solution" + i, DisplayName = "Name" + i,
                Version = "1", Publisher = "A", IsManaged = false }).ToList();
            var target = source.Select(s => new SolutionInfo { UniqueName = s.UniqueName, DisplayName = s.DisplayName,
                Version = s.Version, Publisher = s.Publisher, IsManaged = s.IsManaged }).ToList();
            target[0].Version = "2"; target[1].Publisher = "B"; target[2].IsManaged = true;
            target[3].Version = "2"; target[3].Publisher = "B"; target.RemoveAt(4);
            target.Add(new SolutionInfo { UniqueName = "targetOnly", DisplayName = "Target", Version = "1", Publisher = "B" });
            var rows = new SolutionComparisonService().Compare(source, target); var before = rows.Select(Snapshot).ToArray();
            CollectionAssert.AreEqual(new[] { 10001, 9995, 4, 2, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(rows));
            CollectionAssert.AreEqual(before, rows.Select(Snapshot).ToArray());
            CollectionAssert.AreEqual(before, new SolutionComparisonService().Compare(source, target).Select(Snapshot).ToArray());
        }

        [DataTestMethod]
        [DataRow(false, false, 0, "Not connected  |  Solutions not loaded")]
        [DataRow(true, false, 0, "Connected  |  Solutions not loaded")]
        [DataRow(true, true, 0, "Connected  |  0 loaded solutions")]
        [DataRow(true, true, 195, "Connected  |  195 loaded solutions")]
        [DataRow(false, true, 184, "Not connected  |  184 loaded solutions")]
        [DataRow(false, false, 195, "Not connected  |  Solutions not loaded")]
        public void EnvironmentStateDoesNotSuggestUnloadedOrDisconnectedSuccess(bool connected, bool loaded, int count, string expected)
        {
            var result = SolutionComparerControl.MainEnvironmentPresentation("Owner environment", connected, loaded, count);
            Assert.AreEqual("Owner environment", result[0]); Assert.AreEqual(expected, result[1]);
        }

        [DataTestMethod, DataRow(null), DataRow("")]
        public void UnavailableEnvironmentNameIsHonest(string name)
        { Assert.AreEqual("Environment name unavailable", SolutionComparerControl.MainEnvironmentPresentation(name, false, false, 0)[0]); }

        [TestMethod]
        public void EnvironmentProjectionPreservesDistinctSideOwnedNamesCountsAndState()
        {
            var source = SolutionComparerControl.MainEnvironmentPresentation("SYN-THET-IAB", true, true, 195);
            var target = SolutionComparerControl.MainEnvironmentPresentation("SYN-THET-IAC", true, false, 184);
            Assert.AreEqual("SYN-THET-IAB", source[0]); Assert.AreEqual("Connected  |  195 loaded solutions", source[1]);
            Assert.AreEqual("SYN-THET-IAC", target[0]); Assert.AreEqual("Connected  |  Solutions not loaded", target[1]);
        }

        [DataTestMethod, DataRow(420, 46), DataRow(419, 84), DataRow(280, 84), DataRow(279, 122), DataRow(140, 122), DataRow(0, 122)]
        public void ToolbarWrapsWithoutDroppingActionsAtExactBoundaries(int width, int expected)
        { Assert.AreEqual(expected, SolutionComparerControl.MainToolbarWrappedHeight(width, new[] { new Size(140, 38), new Size(140, 38), new Size(140, 38) }, 8)); }

        [DataTestMethod, DataRow(1), DataRow(2), DataRow(3)]
        public void ToolbarHeightUsesScaledItemMeasurementsRatherThanAFixedTwoRowLimit(int scale)
        {
            var items = Enumerable.Repeat(new Size(140 * scale, 38 * scale), 11).ToArray();
            Assert.AreEqual((38 * 6 + 8) * scale, SolutionComparerControl.MainToolbarWrappedHeight(280 * scale, items, 8 * scale));
        }

        [TestMethod]
        public void ToolbarAccountsForTallWrappedCheckboxAndPadding()
        {
            var items = new[] { new Size(120, 38), new Size(120, 38), new Size(240, 58) };
            Assert.AreEqual(104, SolutionComparerControl.MainToolbarWrappedHeight(240, items, 8));
            Assert.AreEqual(66, SolutionComparerControl.MainToolbarWrappedHeight(480, items, 8));
        }

        [TestMethod]
        public void EmptyToolbarRequiresOnlyItsPadding()
        { Assert.AreEqual(8, SolutionComparerControl.MainToolbarWrappedHeight(800, new Size[0], 8)); }

        private static string Snapshot(SolutionCompareResult r) => string.Join("|", new[] { r.UniqueName, r.SourceDisplayName, r.TargetDisplayName,
            r.SourceVersion, r.TargetVersion, r.SourcePublisher, r.TargetPublisher, r.SourcePackageType, r.TargetPackageType, r.PackageTypeStatus, r.Status });

        private sealed class SinglePassRows : IEnumerable<SolutionCompareResult>
        {
            internal int Enumerations;
            public IEnumerator<SolutionCompareResult> GetEnumerator()
            {
                Assert.AreEqual(1, ++Enumerations);
                yield return new SolutionCompareResult { Status = "Match" };
                yield return new SolutionCompareResult { Status = "Multiple Differences" };
                yield return new SolutionCompareResult { Status = "Missing in Source" };
            }
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
