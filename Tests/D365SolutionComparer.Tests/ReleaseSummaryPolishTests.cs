using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using D365SolutionComparer.Models;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ReleaseSummaryPolishTests
    {
        [TestMethod]
        public void DisclosureStartsCollapsedAndReopensWithoutChangingCounts()
        {
            var state = new MainComparisonSummaryPresentation(); var rows = SyAeFixture();
            Assert.IsFalse(state.Expanded); Assert.AreEqual("Detailed breakdown ▾", state.Caption);
            var counts = SolutionComparerControl.MainDetailedSummaryCounts(rows);
            state.Toggle(); Assert.IsTrue(state.Expanded); Assert.AreEqual("Detailed breakdown ▴", state.Caption);
            state.Toggle(); Assert.IsFalse(state.Expanded); Assert.AreEqual(0, state.DetailsHeight(120));
            state.Toggle(); CollectionAssert.AreEqual(counts, SolutionComparerControl.MainDetailedSummaryCounts(rows));
        }

        [DataTestMethod, DataRow(0), DataRow(18), DataRow(60), DataRow(180)]
        public void CollapsedDetailsReturnTheirEntireMeasuredHeightToTheGrid(int measured)
        { Assert.AreEqual(0, new MainComparisonSummaryPresentation().DetailsHeight(measured)); }

        [DataTestMethod, DataRow(1), DataRow(24), DataRow(120)]
        public void ExpandedDetailsUseTheMeasuredWrappingAllocation(int measured)
        { var state = new MainComparisonSummaryPresentation(); state.Toggle(); Assert.AreEqual(measured, state.DetailsHeight(measured)); }

        [TestMethod]
        public void AcceptedSummaryPartitionScenariosRemainConsistent()
        {
            var rows = SyAeFixture(); var dashboard = SolutionComparerControl.MainComparisonDashboardCounts(rows);
            var detail = SolutionComparerControl.MainDetailedSummaryCounts(rows);
            CollectionAssert.AreEqual(new[] { 186, 129, 29, 28, 0 }, dashboard);
            CollectionAssert.AreEqual(new[] { 186, 129, 24, 2, 0, 3, 3, 1, 14, 14 }, detail);
            Assert.AreEqual(30, detail[2] + detail[3] + detail[4] + detail[5] + detail[7]);
            Assert.AreEqual(1, rows.Count(r => r.Status == "Multiple Differences" && Predicate("IsAnyPackageTypeDifference", r)));
            Assert.AreEqual(2, rows.Count(r => r.Status == "Managed/Unmanaged Mismatch"));
            Assert.AreEqual(29, rows.Count(r => r.Status != "Match" && !r.Status.StartsWith("Missing", StringComparison.Ordinal)));
        }

        [DataTestMethod, DataRow(9), DataRow(10), DataRow(12), DataRow(15)]
        public void MultipleRowsAlsoCountInPackageIndicatorsWithoutBeingCountedTwiceInDifferent(int mask)
        {
            var row = One(mask); var counts = SolutionComparerControl.MainDetailedSummaryCounts(new[] { row });
            Assert.AreEqual("Multiple Differences", row.Status); Assert.AreEqual(1, counts[5]);
            Assert.AreEqual(1, counts[6]); Assert.AreEqual(1, counts[7]);
            Assert.AreEqual(0, counts[2] + counts[3] + counts[4]);
            CollectionAssert.AreEqual(new[] { 1, 0, 1, 0, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(new[] { row }));
        }

        [DataTestMethod, DataRow(1, 2), DataRow(2, 3), DataRow(4, 4), DataRow(8, 5)]
        public void SingleDifferenceStatusCountsKeepTheirOriginalMeaning(int mask, int index)
        {
            var row = One(mask); var detail = SolutionComparerControl.MainDetailedSummaryCounts(new[] { row });
            Assert.AreEqual(1, detail[index]); Assert.AreEqual(0, detail[7]);
            Assert.AreEqual(1, SolutionComparerControl.MainComparisonDashboardCounts(new[] { row })[2]);
            Assert.AreEqual(mask == 8 ? 1 : 0, detail[5]);
        }

        [DataTestMethod, DataRow("Match"), DataRow("Version Mismatch"), DataRow("Multiple Differences"), DataRow("Missing in Target"), DataRow("EMPTY")]
        public void ProvidedFilteredViewAndAllResultCardsKeepTheirDistinctScopes(string status)
        {
            var all = SyAeFixture(); var before = all.Select(Snapshot).ToArray(); var full = SolutionComparerControl.MainComparisonDashboardCounts(all);
            var filtered = all.Where(r => status != "EMPTY" && r.Status == status).ToList();
            Assert.AreEqual(filtered.Count, SolutionComparerControl.MainDetailedSummaryCounts(filtered)[0]);
            StringAssert.Contains(MainComparisonSummaryPresentation.ScopeText(all.Count, filtered.Count), "Cards: all 186");
            StringAssert.Contains(MainComparisonSummaryPresentation.ScopeText(all.Count, filtered.Count), "(" + filtered.Count + " rows)");
            CollectionAssert.AreEqual(full, SolutionComparerControl.MainComparisonDashboardCounts(all));
            CollectionAssert.AreEqual(before, all.Select(Snapshot).ToArray());
        }

        [TestMethod]
        public void ChangedAndManagedPredicatesRetainCountsAndResetToAllWithoutDisclosureSideEffects()
        {
            var all = SyAeFixture(); var state = new MainComparisonSummaryPresentation(); state.Toggle();
            var changed = all.Where(r => Predicate("IsChangedResult", r)).ToArray();
            var managed = changed.Where(r => Predicate("IsManagedUnmanagedDifference", r)).ToArray();
            Assert.AreEqual(57, changed.Length); Assert.AreEqual(3, managed.Length);
            Assert.AreEqual(3, SolutionComparerControl.MainDetailedSummaryCounts(managed)[0]);
            Assert.IsTrue(state.Expanded);
            CollectionAssert.AreEqual(new[] { 186, 129, 29, 28, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(all));
            Assert.AreEqual(186, SolutionComparerControl.MainDetailedSummaryCounts(all)[0]);
            state.Toggle(); Assert.IsFalse(state.Expanded); Assert.AreEqual(57, changed.Length);
        }

        [DataTestMethod, DataRow(0, false), DataRow(1, true), DataRow(8, true), DataRow(9, true)]
        public void ChangedOnlyKeepsExistingPredicateForEveryDisclosureState(int mask, bool changed)
        {
            var row = One(mask); var state = new MainComparisonSummaryPresentation();
            for (int i = 0; i < 3; i++) { Assert.AreEqual(changed, Predicate("IsChangedResult", row)); state.Toggle(); }
        }

        [TestMethod]
        public void ExpandedTextExplainsTheNonAdditiveIndicatorsInsteadOfHidingTheDiscrepancy()
        {
            string text = MainComparisonSummaryPresentation.Breakdown(SolutionComparerControl.MainDetailedSummaryCounts(SyAeFixture()));
            StringAssert.Contains(text, "Current filtered view: Total=186 | Match=129");
            StringAssert.Contains(text, "Version differences=24"); StringAssert.Contains(text, "Publisher differences=2");
            StringAssert.Contains(text, "Display-name differences=0"); StringAssert.Contains(text, "Multiple differences=1");
            StringAssert.Contains(text, "Package-type differences=3"); StringAssert.Contains(text, "Managed/unmanaged differences=3");
            StringAssert.Contains(text, "Missing in Source=14"); StringAssert.Contains(text, "Missing in Target=14");
            StringAssert.Contains(text, "single-difference overall statuses"); StringAssert.Contains(text, "subset");
            StringAssert.Contains(text, "must not be summed to Different");
        }

        [TestMethod]
        public void StatusHelpRetainsStatusWordingAndSelectionMeaning()
        {
            foreach (string status in new[] { "Match", "Version Mismatch", "Publisher Mismatch", "Display Name Mismatch", "Managed/Unmanaged Mismatch", "Package Type Mismatch", "Multiple Differences", "Missing in Source", "Missing in Target" })
                StringAssert.Contains(MainComparisonSummaryPresentation.StatusHelp, status);
            StringAssert.Contains(MainComparisonSummaryPresentation.StatusHelp, "blue with white text");
            StringAssert.Contains(MainComparisonSummaryPresentation.StatusHelp, "neutral");
            StringAssert.Contains(MainComparisonSummaryPresentation.StatusHelp, "Colours do not change comparison meaning");
        }

        [TestMethod]
        public void AssemblyFileAndInformationalVersionsAreAlignedForHeaderAboutAndDiscovery()
        {
            var assembly = typeof(SolutionComparerControl).Assembly;
            Assert.AreEqual("1.2026.2.0", assembly.GetName().Version.ToString());
            Assert.AreEqual("1.2026.2.0", assembly.GetCustomAttribute<AssemblyFileVersionAttribute>().Version);
            Assert.AreEqual("1.2026.2.0", assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion);
            Assert.AreEqual("1.2026.2.0", FileVersionInfo.GetVersionInfo(assembly.Location).ProductVersion);
        }

        [DataTestMethod, DataRow(360, 96), DataRow(800, 96), DataRow(1200, 144), DataRow(1920, 192)]
        public void CardLayoutIsIndependentOfDisclosureAndKeepsExistingResizeBehavior(int width, int dpi)
        {
            var before = SolutionComparerControl.MainDashboardCardBounds(width, dpi, 40);
            var state = new MainComparisonSummaryPresentation(); state.Toggle(); Assert.AreEqual(120, state.DetailsHeight(120));
            CollectionAssert.AreEqual(before, SolutionComparerControl.MainDashboardCardBounds(width, dpi, 40));
            state.Toggle(); Assert.AreEqual(0, state.DetailsHeight(120));
            foreach (var rect in before) Assert.IsTrue(rect.Left >= 0 && rect.Right <= width);
        }

        [TestMethod]
        public void SummaryStateNeverMutatesExportInputsOrderOrOriginalReferences()
        {
            var rows = SyAeFixture(); var references = rows.ToArray(); var fields = rows.Select(Snapshot).ToArray();
            var membership = ExplorerFixture.Saa(); string coverage = ExplorerFixture.Csv(membership);
            var state = new MainComparisonSummaryPresentation();
            for (int i = 0; i < 4; i++)
            {
                MainComparisonSummaryPresentation.Breakdown(SolutionComparerControl.MainDetailedSummaryCounts(rows));
                state.Toggle(); Assert.AreEqual(coverage, ExplorerFixture.Csv(membership));
                CollectionAssert.AreEqual(fields, rows.Select(Snapshot).ToArray());
                for (int r = 0; r < rows.Count; r++) Assert.AreSame(references[r], rows[r]);
            }
        }

        [TestMethod]
        public void EmptyFilteredDetailsStayZeroWithoutChangingAllResultCounts()
        {
            var all = SyAeFixture(); var empty = SolutionComparerControl.MainDetailedSummaryCounts(new SolutionCompareResult[0]);
            CollectionAssert.AreEqual(new int[10], empty);
            StringAssert.Contains(MainComparisonSummaryPresentation.Breakdown(empty), "Total=0 | Match=0");
            Assert.AreEqual(186, SolutionComparerControl.MainComparisonDashboardCounts(all)[0]);
        }

        private static bool Predicate(string name, SolutionCompareResult row)
        {
            var method = typeof(SolutionComparerControl).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method); return (bool)method.Invoke(null, new object[] { row });
        }
        private static SolutionInfo Info(string name) => new SolutionInfo { UniqueName = name, DisplayName = name, Version = "1", Publisher = "Publisher", IsManaged = false };
        private static SolutionCompareResult One(int mask)
        {
            var a = Info("offline"); var b = Info("offline");
            if ((mask & 1) != 0) b.Version = "2"; if ((mask & 2) != 0) b.Publisher = "Other";
            if ((mask & 4) != 0) b.DisplayName = "Other"; if ((mask & 8) != 0) b.IsManaged = true;
            return new SolutionComparisonService().Compare(new List<SolutionInfo> { a }, new List<SolutionInfo> { b }).Single();
        }
        private static List<SolutionCompareResult> SyAeFixture()
        {
            var source = new List<SolutionInfo>(); var target = new List<SolutionInfo>(); int index = 0;
            Action<int, int> add = (count, mask) => { for (int i = 0; i < count; i++) {
                string key = "offline-summary-" + index++; var a = Info(key); var b = Info(key);
                if ((mask & 1) != 0) b.Version = "2"; if ((mask & 2) != 0) b.Publisher = "Other";
                if ((mask & 8) != 0) b.IsManaged = true;
                if (mask != 16) source.Add(a); if (mask != 32) target.Add(b);
            }};
            add(129, 0); add(24, 1); add(2, 2); add(2, 8); add(1, 9); add(14, 16); add(14, 32);
            return new SolutionComparisonService().Compare(source, target);
        }
        private static string Snapshot(SolutionCompareResult r) => string.Join("|", new[] { r.UniqueName, r.SourceDisplayName, r.TargetDisplayName,
            r.SourceVersion, r.TargetVersion, r.SourcePublisher, r.TargetPublisher, r.SourcePackageType, r.TargetPackageType, r.PackageTypeStatus, r.Status });
    }
}
