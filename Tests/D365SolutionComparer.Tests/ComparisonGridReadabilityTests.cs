using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComparisonGridReadabilityTests
    {
        [DataTestMethod, DataRow(96, 260), DataRow(144, 390), DataRow(192, 520)]
        public void StatusDefaultWidthsScaleWithoutSettingAMinimum(int dpi, int expected)
        { Assert.AreEqual(expected, SolutionComparerControl.ComparisonStatusDefaultWidth(dpi)); }

        [TestMethod]
        public void CommonStatusTextFitsTheDefaultWithPaddingAndSortGlyphAllowance()
        {
            using (var font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point))
            using (var bitmap = new Bitmap(1, 1))
                foreach (int dpi in new[] { 96, 144, 192 })
                {
                    bitmap.SetResolution(dpi, dpi);
                    using (var graphics = Graphics.FromImage(bitmap))
                        foreach (string text in Statuses)
                        {
                            var width = graphics.MeasureString(text, font, PointF.Empty, StringFormat.GenericTypographic).Width;
                            Assert.IsTrue(width + 32d * dpi / 96d <= SolutionComparerControl.ComparisonStatusDefaultWidth(dpi), text + " at DPI " + dpi);
                        }
                }
        }

        [DataTestMethod, DataRow(true, false), DataRow(true, true), DataRow(false, false), DataRow(false, true)]
        public void WidthCacheRecordsOnlyComparisonChangesOutsideLayout(bool comparison, bool layout)
        {
            var saved = new Dictionary<string, int>();
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 91, comparison, layout);
            Assert.AreEqual(comparison && !layout ? 1 : 0, saved.Count);
            if (comparison && !layout) Assert.AreEqual(91, saved["Status"]);
        }

        [DataTestMethod, DataRow("", 90), DataRow("Status", 0), DataRow("Status", -1)]
        public void InvalidWidthsCannotPolluteTheSessionCache(string name, int width)
        {
            var saved = new Dictionary<string, int>();
            SolutionComparerControl.RememberComparisonColumnWidth(saved, name, width, true, false); Assert.AreEqual(0, saved.Count);
        }

        [DataTestMethod, DataRow("Status", 80), DataRow("PackageTypeStatus", 95), DataRow("SourceDisplayName", 310), DataRow("UniqueName", 180)]
        public void SavedWidthsWinOverNewDefaultsIncludingNarrowUserWidths(string name, int width)
        {
            var saved = new Dictionary<string, int>(); SolutionComparerControl.RememberComparisonColumnWidth(saved, name, width, true, false);
            Assert.AreEqual(width, SolutionComparerControl.ComparisonColumnWidth(name, 260, saved));
            Assert.AreEqual(260, SolutionComparerControl.ComparisonColumnWidth("UnseenColumn", 260, saved));
        }

        [TestMethod]
        public void SessionWidthsSurviveRefreshReloadAndInternalDefaultAssignments()
        {
            var saved = new Dictionary<string, int>();
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 123, true, false);
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "PackageTypeStatus", 345, true, false);
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 260, true, true);
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 180, false, false);
            Assert.AreEqual(123, SolutionComparerControl.ComparisonColumnWidth("Status", 520, saved));
            Assert.AreEqual(345, SolutionComparerControl.ComparisonColumnWidth("PackageTypeStatus", 520, saved));
            SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 149, true, false);
            Assert.AreEqual(149, SolutionComparerControl.ComparisonColumnWidth("Status", 260, saved));
        }

        [DataTestMethod]
        [DataRow("Match"), DataRow("Managed/Unmanaged Mismatch"), DataRow("Multiple Differences")]
        [DataRow("Publisher Mismatch"), DataRow("Display Name Mismatch"), DataRow("Version Mismatch")]
        [DataRow("Missing in Source"), DataRow("Missing in Target"), DataRow("Package Type Mismatch")]
        public void BothStatusColumnsAlwaysExposeTheCompleteTooltip(string text)
        {
            foreach (string column in new[] { "Status", "PackageTypeStatus" })
                Assert.AreEqual(text, SolutionComparerControl.ComparisonCellToolTip(column, text));
        }

        [DataTestMethod]
        [DataRow("UniqueName", "short", "")]
        [DataRow("UniqueName", "A synthetic name longer than twenty-four characters", "A synthetic name longer than twenty-four characters")]
        [DataRow("Status", null, "")]
        [DataRow("PackageTypeStatus", "", "")]
        [DataRow("Status", "unknown short", "unknown short")]
        public void TooltipPolicyKeepsOtherColumnsAndUnavailableValuesHonest(string column, string value, string expected)
        { Assert.AreEqual(expected, SolutionComparerControl.ComparisonCellToolTip(column, value)); }

        [DataTestMethod, DataRow("Match"), DataRow("Missing in Target"), DataRow("Managed/Unmanaged Mismatch"), DataRow("Multiple Differences")]
        public void WidthAndTooltipPoliciesDoNotAlterApprovedSelectionColors(string status)
        {
            foreach (string column in new[] { "Status", "PackageTypeStatus" })
            {
                var style = new DataGridViewCellStyle(); SolutionComparerControl.ApplyComparisonCellStyle(style, column, status, false);
                var before = new[] { style.BackColor, style.ForeColor, style.SelectionBackColor, style.SelectionForeColor };
                SolutionComparerControl.ComparisonCellToolTip(column, status);
                SolutionComparerControl.ComparisonColumnWidth(column, 260, new Dictionary<string, int> { [column] = 80 });
                CollectionAssert.AreEqual(before, new[] { style.BackColor, style.ForeColor, style.SelectionBackColor, style.SelectionForeColor });
                Assert.AreEqual(Color.White, style.SelectionForeColor); Assert.AreEqual(Color.FromArgb(29, 78, 216), style.SelectionBackColor);
            }
        }

        [TestMethod]
        public void TooltipAndWidthMetadataCannotChangeComparisonExportValues()
        {
            var rows = Statuses.Select((text, i) => new SolutionCompareResult { UniqueName = "synthetic" + i, Status = text, PackageTypeStatus = text }).ToList();
            var before = rows.Select(r => r.UniqueName + "|" + r.Status + "|" + r.PackageTypeStatus).ToArray(); var saved = new Dictionary<string, int>();
            foreach (var row in rows)
            {
                SolutionComparerControl.ComparisonCellToolTip("Status", row.Status); SolutionComparerControl.ComparisonCellToolTip("PackageTypeStatus", row.PackageTypeStatus);
                SolutionComparerControl.RememberComparisonColumnWidth(saved, "Status", 90, true, false);
            }
            CollectionAssert.AreEqual(before, rows.Select(r => r.UniqueName + "|" + r.Status + "|" + r.PackageTypeStatus).ToArray());
            Assert.AreEqual(Statuses.Length, rows.Count);
        }
        private static readonly string[] Statuses = { "Match", "Managed/Unmanaged Mismatch", "Multiple Differences", "Publisher Mismatch", "Display Name Mismatch", "Version Mismatch", "Missing in Source", "Missing in Target", "Package Type Mismatch" };
    }
}
