using System;
using System.Drawing;
using System.Linq;
using D365SolutionComparer.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ComparisonUiRefinementTests
    {
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void UnavailableCompareHasNeutralStartupAndInteractionColors(bool focused)
        {
            var colors = SolutionComparerControl.MainCompareButtonColors(false, focused);
            Assert.AreEqual(Color.FromArgb(241, 245, 249), colors[0]); Assert.AreEqual(Color.FromArgb(71, 85, 105), colors[1]);
            Assert.AreEqual(colors[0], colors[2]); Assert.AreEqual(colors[0], colors[3]);
            Assert.AreNotEqual(Color.FromArgb(29, 78, 216), colors[0]);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ReadyCompareRetainsPrimaryBlueAndWhiteAcrossInteractionStates(bool focused)
        {
            var colors = SolutionComparerControl.MainCompareButtonColors(true, focused);
            Assert.AreEqual(Color.FromArgb(29, 78, 216), colors[0]); Assert.AreEqual(Color.White, colors[1]);
            Assert.AreEqual(Color.FromArgb(30, 64, 175), colors[2]); Assert.AreEqual(Color.FromArgb(30, 58, 138), colors[3]);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ComparePaletteHasAtLeastFourPointFiveContrastForNormalHoverAndPressed(bool enabled)
        {
            foreach (bool focused in new[] { false, true })
            {
                var colors = SolutionComparerControl.MainCompareButtonColors(enabled, focused);
                foreach (int background in new[] { 0, 2, 3 }) Assert.IsTrue(Contrast(colors[1], colors[background]) >= 4.5);
            }
        }

        [TestMethod]
        public void FocusChangesOnlyTheBorderAndDoesNotChooseAPressedBackground()
        {
            var plain = SolutionComparerControl.MainCompareButtonColors(true, false); var focus = SolutionComparerControl.MainCompareButtonColors(true, true);
            CollectionAssert.AreEqual(plain.Take(4).ToArray(), focus.Take(4).ToArray());
            Assert.AreNotEqual(plain[4], focus[4]); Assert.AreEqual(Color.FromArgb(15, 23, 42), focus[4]);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void EnableAndFocusTransitionsCannotRetainPreviousPaletteMetadata(bool initial)
        {
            foreach (bool enabled in new[] { initial, !initial, initial, false, true, false })
            {
                var colors = SolutionComparerControl.MainCompareButtonColors(enabled, true);
                colors[0] = Color.Red; colors[3] = Color.Magenta;
                var restored = SolutionComparerControl.MainCompareButtonColors(enabled, false);
                Assert.AreEqual(enabled ? Color.FromArgb(29, 78, 216) : Color.FromArgb(241, 245, 249), restored[0]);
                Assert.AreNotEqual(Color.Magenta, restored[3]);
            }
        }

        [DataTestMethod, DataRow(740), DataRow(1024), DataRow(1920), DataRow(3840)]
        public void WideDashboardKeepsCompactLeftAlignedCardsAndTwelvePixelGaps(int available)
        {
            var bounds = SolutionComparerControl.MainDashboardCardBounds(available, 96, 40);
            for (int i = 0; i < 4; i++) Assert.AreEqual(new Rectangle(i * 188, 0, 176, 40), bounds[i]);
            Assert.AreEqual(740, bounds[3].Right); Assert.AreEqual(0, bounds[0].Left);
        }

        [DataTestMethod, DataRow(96, 176, 12), DataRow(144, 264, 18), DataRow(192, 352, 24)]
        public void DashboardCardWidthsAndSpacingScaleConsistently(int dpi, int width, int gap)
        {
            var bounds = SolutionComparerControl.MainDashboardCardBounds(5000, dpi, 80);
            foreach (var b in bounds) Assert.AreEqual(width, b.Width);
            for (int i = 1; i < 4; i++) { Assert.AreEqual(gap, bounds[i].Left - bounds[i - 1].Right); Assert.AreEqual(0, bounds[i].Top); }
        }

        [DataTestMethod, DataRow(739, 3), DataRow(552, 3), DataRow(551, 2), DataRow(364, 2)]
        [DataRow(363, 1), DataRow(176, 1), DataRow(120, 1), DataRow(0, 1)]
        public void DashboardWrapsAllFourCardsWithoutClippingOrOverlapInItsAllocation(int available, int columns)
        {
            var bounds = SolutionComparerControl.MainDashboardCardBounds(available, 96, 40);
            Assert.AreEqual(4, bounds.Length);
            for (int i = 0; i < bounds.Length; i++)
            {
                Assert.AreEqual(i / columns * 52, bounds[i].Top);
                Assert.IsTrue(bounds[i].Left >= 0 && bounds[i].Right <= Math.Max(1, available));
                Assert.AreEqual(bounds[0].Width, bounds[i].Width);
                for (int j = 0; j < i; j++) Assert.IsFalse(bounds[i].IntersectsWith(bounds[j]));
            }
        }

        [DataTestMethod, DataRow(40), DataRow(68), DataRow(100)]
        public void TallerTextAllocationsKeepAllCardsAlignedAndSeparated(int height)
        {
            var bounds = SolutionComparerControl.MainDashboardCardBounds(364, 96, height);
            Assert.AreEqual(0, bounds[0].Top); Assert.AreEqual(0, bounds[1].Top);
            Assert.AreEqual(height + 12, bounds[2].Top); Assert.AreEqual(bounds[2].Top, bounds[3].Top);
            foreach (var b in bounds) Assert.AreEqual(height, b.Height);
        }

        [TestMethod]
        public void ZeroSizeLayoutInputsStillProduceNonnegativeCardBounds()
        {
            foreach (var b in SolutionComparerControl.MainDashboardCardBounds(0, 0, 0))
                Assert.IsTrue(b.X >= 0 && b.Y >= 0 && b.Width == 1 && b.Height == 1);
        }

        [TestMethod]
        public void OwnerSampleTextFitsCompactCardWidthAtCommonDpiInOffscreenFontEstimate()
        {
            using (var font = new Font("Segoe UI", 10F, FontStyle.Bold))
            using (var bitmap = new Bitmap(1, 1))
                foreach (int dpi in new[] { 96, 144, 192 })
                {
                    bitmap.SetResolution(dpi, dpi);
                    using (var graphics = Graphics.FromImage(bitmap))
                        foreach (string text in new[] { "Total   209", "Match   99", "Different   63", "Missing   47" })
                            Assert.IsTrue(graphics.MeasureString(text, font, PointF.Empty, StringFormat.GenericTypographic).Width + 14d * dpi / 96d
                                <= SolutionComparerControl.MainDashboardCardBounds(5000, dpi, 40)[0].Width, text);
                }
        }

        [TestMethod]
        public void HistoricalOwnerCountsReconcileAsSyntheticEvidenceWithoutHardcodingProduction()
        {
            var rows = OwnerFixture(); var text = rows.Select(r => r.Status).ToArray();
            CollectionAssert.AreEqual(new[] { 209, 99, 63, 47, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(rows));
            CollectionAssert.AreEqual(text, rows.Select(r => r.Status).ToArray());
        }

        [TestMethod]
        public void FilteredSubsetAndFullComparisonRetainTheirExistingDistinctMeanings()
        {
            var all = OwnerFixture(); var filtered = all.Where(r => r.Status == "Match").ToArray();
            CollectionAssert.AreEqual(new[] { 99, 99, 0, 0, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(filtered));
            SolutionComparerControl.MainDashboardCardBounds(364, 144, 60);
            CollectionAssert.AreEqual(new[] { 209, 99, 63, 47, 0 }, SolutionComparerControl.MainComparisonDashboardCounts(all));
            Assert.AreEqual(99, filtered.Length); Assert.AreEqual(209, all.Length);
        }

        [DataTestMethod, DataRow(640, 20, 620), DataRow(480, 20, 460), DataRow(1920, 20, 1900), DataRow(20, 20, 1), DataRow(0, 20, 1)]
        public void ToolbarUsesItsClientWidthAndPaddingWithASafeLowerBound(int client, int padding, int expected)
        { Assert.AreEqual(expected, SolutionComparerControl.MainToolbarAvailableWidth(client, padding)); }

        [DataTestMethod, DataRow(640, 122), DataRow(1024, 84), DataRow(1920, 46), DataRow(480, 198)]
        public void ElevenToolbarItemsIncludingBothCheckboxesReceiveEnoughWrappedHeight(int client, int expected)
        { Assert.AreEqual(expected, SolutionComparerControl.MainToolbarWrappedHeight(SolutionComparerControl.MainToolbarAvailableWidth(client, 20), ToolbarItems(), 8)); }

        [TestMethod]
        public void DockedToolbarWidthAvoidsUnderallocationFromAWiderParent()
        {
            int actual = SolutionComparerControl.MainToolbarAvailableWidth(640, 20), parent = SolutionComparerControl.MainToolbarAvailableWidth(1000, 20);
            int required = SolutionComparerControl.MainToolbarWrappedHeight(actual, ToolbarItems(), 8);
            Assert.AreEqual(122, required); Assert.IsTrue(required > SolutionComparerControl.MainToolbarWrappedHeight(parent, ToolbarItems(), 8));
        }

        [DataTestMethod, DataRow(96), DataRow(144), DataRow(192)]
        public void ToolbarAllocationIncludesScaledMarginsPaddingAndAllElevenItems(int dpi)
        {
            double scale = dpi / 96d;
            var items = ToolbarItems().Select(s => new Size((int)(s.Width * scale), (int)(s.Height * scale))).ToArray();
            int available = SolutionComparerControl.MainToolbarAvailableWidth((int)(640 * scale), (int)(20 * scale));
            Assert.AreEqual((int)(122 * scale), SolutionComparerControl.MainToolbarWrappedHeight(available, items, (int)(8 * scale)));
        }

        [TestMethod]
        public void TallFinalManagedCheckboxGetsAdditionalRowHeight()
        {
            var items = ToolbarItems(); items[10] = new Size(items[10].Width, 96);
            Assert.AreEqual(180, SolutionComparerControl.MainToolbarWrappedHeight(620, items, 8));
        }

        private static Size[] ToolbarItems() => new[] { 138, 138, 138, 190, 138, 182, 138, 138, 138, 134, 310 }.Select(w => new Size(w, 38)).ToArray();
        private static SolutionCompareResult[] OwnerFixture() => Enumerable.Repeat("Match", 99).Concat(Enumerable.Repeat("Multiple Differences", 63))
            .Concat(Enumerable.Repeat("Missing in Source", 23)).Concat(Enumerable.Repeat("Missing in Target", 24)).Select(s => new SolutionCompareResult { Status = s }).ToArray();
        private static double Luminance(Color color)
        {
            Func<byte, double> channel = value => { double s = value / 255d; return s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4); };
            return 0.2126 * channel(color.R) + 0.7152 * channel(color.G) + 0.0722 * channel(color.B);
        }
        private static double Contrast(Color a, Color b) { double x = Luminance(a), y = Luminance(b); return (Math.Max(x, y) + 0.05) / (Math.Min(x, y) + 0.05); }
    }
}
