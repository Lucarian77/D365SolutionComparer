using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models;
using D365SolutionComparer.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComparisonGridStyleTests
    {
        public TestContext TestContext { get; set; }
        [DataTestMethod]
        [DataRow("Match", 0x166534, 0xF0FDF4)]
        [DataRow("Version Mismatch", 0x92400E, 0xFFFBEB)]
        [DataRow("Publisher Mismatch", 0x1D4ED8, 0xEFF6FF)]
        [DataRow("Display Name Mismatch", 0x1D4ED8, 0xEFF6FF)]
        [DataRow("Package Type Mismatch", 0x9A3412, 0xFFF7ED)]
        [DataRow("Managed/Unmanaged Mismatch", 0x9A3412, 0xFFF7ED)]
        [DataRow("Multiple Differences", 0x6D28D9, 0xF5F3FF)]
        [DataRow("Missing in Source", 0xB91C1C, 0xFEF2F2)]
        [DataRow("Missing in Target", 0xB91C1C, 0xFEF2F2)]
        public void OverallStatusUsesRequestedPalette(string status, int foreground, int background)
        {
            var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, "Status", status, false);
            Assert.AreEqual(Rgb(foreground), style.ForeColor); Assert.AreEqual(Rgb(background), style.BackColor);
        }

        [DataTestMethod]
        [DataRow("Version Mismatch", 0x92400E, 0xFFFBEB)]
        [DataRow("Publisher Mismatch", 0x1D4ED8, 0xEFF6FF)]
        [DataRow("Display Name Mismatch", 0x1D4ED8, 0xEFF6FF)]
        [DataRow("Package Type Mismatch", 0x9A3412, 0xFFF7ED)]
        [DataRow("Managed/Unmanaged Mismatch", 0x9A3412, 0xFFF7ED)]
        [DataRow("Multiple Differences", 0x6D28D9, 0xF5F3FF)]
        [DataRow("Missing in Source", 0xB91C1C, 0xFEF2F2)]
        [DataRow("Missing in Target", 0xB91C1C, 0xFEF2F2)]
        public void PackageDifferenceUsesRequestedPalette(string status, int foreground, int background)
        {
            var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, "PackageTypeStatus", status, true);
            Assert.AreEqual(Rgb(foreground), style.ForeColor); Assert.AreEqual(Rgb(background), style.BackColor);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void PackageMatchKeepsNeutralRowBackground(bool alternate)
        {
            var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, "PackageTypeStatus", "Match", alternate);
            Assert.AreEqual(Rgb(0x475569), style.ForeColor); Assert.AreEqual(Rgb(alternate ? 0xF8FAFC : 0xFFFFFF), style.BackColor);
        }

        [TestMethod]
        public void EveryColumnAndStatusKeepsReadableSelectionColors()
        {
            foreach (string column in new[] { "UniqueName", "SourceDisplayName", "TargetDisplayName", "SourceVersion", "TargetVersion", "SourcePublisher", "TargetPublisher", "SourcePackageType", "TargetPackageType", "PackageTypeStatus", "Status" })
                foreach (string status in Statuses.Concat(new string[] { null, "", "Unknown" }))
                    foreach (bool alternate in new[] { false, true })
                    {
                        var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, column, status, alternate);
                        Assert.AreEqual(Rgb(0x1D4ED8), style.SelectionBackColor); Assert.AreEqual(Color.White, style.SelectionForeColor);
                    }
        }

        [DataTestMethod, DataRow("Status"), DataRow("PackageTypeStatus"), DataRow("SourceDisplayName")]
        public void UnknownOrMissingStatusResetsStaleColors(string column)
        {
            var style = new DataGridViewCellStyle { BackColor = Color.Red, ForeColor = Color.Yellow };
            SolutionComparerControl.ApplyComparisonCellStyle(style, column, null, true);
            Assert.AreEqual(Rgb(0xF8FAFC), style.BackColor); Assert.AreEqual(Rgb(0x0F172A), style.ForeColor);
        }

        [DataTestMethod, DataRow("Match"), DataRow(" match "), DataRow("Unknown"), DataRow(null), DataRow("")]
        public void StylingPreservesTheOriginalStatusDescription(string status)
        {
            var result = new SolutionCompareResult { Status = status, PackageTypeStatus = status }; string before = Snapshot(result);
            var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, "Status", result.Status, false);
            SolutionComparerControl.ApplyComparisonCellStyle(style, "PackageTypeStatus", result.PackageTypeStatus, false);
            Assert.AreEqual(before, Snapshot(result)); Assert.AreEqual(status, result.Status); Assert.AreEqual(status, result.PackageTypeStatus);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void StandardRowAndHeaderMetadataIsRestrained(bool alternate)
        {
            var row = SolutionComparerControl.CreateComparisonRowStyle(alternate); var header = SolutionComparerControl.CreateComparisonHeaderStyle();
            Assert.AreEqual(Rgb(alternate ? 0xF8FAFC : 0xFFFFFF), row.BackColor); Assert.AreEqual(DataGridViewTriState.False, row.WrapMode);
            Assert.AreEqual(new Padding(4, 2, 4, 2), row.Padding); Assert.IsNull(row.Font);
            Assert.AreEqual(Rgb(0xF1F5F9), header.BackColor); Assert.AreEqual(Rgb(0x334155), header.ForeColor);
        }

        [DataTestMethod, DataRow(15, 96, 24), DataRow(23, 144, 36), DataRow(30, 192, 48)]
        public void CompactRowMetricsScaleWithDpiAndFont(int fontHeight, int dpi, int expected)
        { Assert.AreEqual(expected, SolutionComparerControl.ComparisonGridRowHeight(fontHeight, dpi)); }

        [TestMethod]
        public void FormattingDoesNotAllocateOrReplaceAFontPerCell()
        {
            using (var font = new Font("Segoe UI", 9F))
            {
                var style = new DataGridViewCellStyle { Font = font };
                foreach (string status in Statuses) { SolutionComparerControl.ApplyComparisonCellStyle(style, "Status", status, false); Assert.AreSame(font, style.Font); }
            }
        }

        [TestMethod]
        public void LargeResultFormattingHasBoundedHeadlessCost()
        {
            var style = new DataGridViewCellStyle(); var clock = Stopwatch.StartNew();
            for (int row = 0; row < 50000; row++)
                for (int column = 0; column < 11; column++) SolutionComparerControl.ApplyComparisonCellStyle(style,
                    column == 10 ? "Status" : column == 9 ? "PackageTypeStatus" : "UniqueName", Statuses[row % Statuses.Length], (row & 1) != 0);
            clock.Stop(); TestContext.WriteLine("Synthetic rows=50000; formatted cells=550000; elapsed(ms)=" + clock.ElapsedMilliseconds);
            Assert.IsTrue(clock.Elapsed < TimeSpan.FromSeconds(5), "Headless styling should remain bounded without per-cell fonts or row scans.");
        }

        [TestMethod]
        public void ComparisonResultsExportInputsAndCountsAreUnchangedByStyling()
        {
            var source = Enumerable.Range(0, 10000).Select(i => new SolutionInfo { UniqueName = "synthetic" + i, DisplayName = "Source " + i,
                Version = "1", Publisher = "Source Publisher", IsManaged = true }).ToList();
            var target = source.Select(s => new SolutionInfo { UniqueName = s.UniqueName, DisplayName = s.DisplayName, Version = s.Version,
                Publisher = s.Publisher, IsManaged = s.IsManaged }).ToList(); target[0].Version = "2"; target[1].Publisher = "Different Publisher";
            var rows = new SolutionComparisonService().Compare(source, target); var before = rows.Select(Snapshot).ToArray();
            var counts = rows.GroupBy(r => r.Status).ToDictionary(g => g.Key, g => g.Count()); var style = new DataGridViewCellStyle();
            foreach (var row in rows) { SolutionComparerControl.ApplyComparisonCellStyle(style, "Status", row.Status, false); SolutionComparerControl.ApplyComparisonCellStyle(style, "PackageTypeStatus", row.PackageTypeStatus, false); }
            CollectionAssert.AreEqual(before, rows.Select(Snapshot).ToArray()); Assert.AreEqual(10000, rows.Count);
            foreach (var pair in counts) Assert.AreEqual(pair.Value, rows.Count(r => r.Status == pair.Key));
        }

        private static readonly string[] Statuses = { "Match", "Version Mismatch", "Publisher Mismatch", "Display Name Mismatch", "Package Type Mismatch", "Managed/Unmanaged Mismatch", "Multiple Differences", "Missing in Source", "Missing in Target" };
        private static Color Rgb(int rgb) => Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
        private static string Snapshot(SolutionCompareResult r) => string.Join("|", new[] { r.UniqueName, r.SourceDisplayName, r.TargetDisplayName, r.SourceVersion, r.TargetVersion, r.SourcePublisher, r.TargetPublisher, r.SourcePackageType, r.TargetPackageType, r.PackageTypeStatus, r.Status });
    }
}
