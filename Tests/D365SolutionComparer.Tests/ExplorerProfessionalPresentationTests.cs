using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class ExplorerProfessionalPresentationTests
    {
        [DataTestMethod, DataRow("Match", 1), DataRow("Different", 2), DataRow("Needs Review", 3), DataRow("Target Only", 4)]
        [DataRow("Source Only", 0), DataRow("Unknown", 0), DataRow(null, 0)]
        public void HeaderMetricsProjectExactExistingStatusGroupsWithoutReclassification(string status, int metric)
        {
            var values = new[] { status }; var counts = SolutionObjectsExplorerControl.ProfessionalExplorerCounts(values);
            Assert.AreEqual(1, counts[0]); Assert.AreEqual(metric == 0 ? 0 : 1, counts.Skip(1).Sum());
            if (metric != 0) Assert.AreEqual(1, counts[metric]); Assert.AreEqual(status, values[0]);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void EmptyOrUnboundMetricInputDoesNotInventComponents(bool nullInput)
        { CollectionAssert.AreEqual(new int[5], SolutionObjectsExplorerControl.ProfessionalExplorerCounts(nullInput ? null : new string[0])); }

        [DataTestMethod, DataRow((int)ObjectsWorkspaceScope.Compare), DataRow((int)ObjectsWorkspaceScope.Source), DataRow((int)ObjectsWorkspaceScope.Target)]
        public void HeaderMetricsAndStylingKeepExistingSideScopeObjectInventory(int scopeValue)
        {
            var scope = (ObjectsWorkspaceScope)scopeValue;
            var f = Fixture(2, true); var projection = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(projection);
            var items = model.Items(scope); var originalRows = f.Presentation.Rows.ToArray(); string csv = ExplorerFixture.Csv(f);
            var before = items.Select(i => i.Key + "|" + i.Status(projection)).ToArray();
            var metrics = SolutionObjectsExplorerControl.ProfessionalExplorerCounts(items.Select(i => i.Status(projection)));
            Assert.AreEqual(items.Count, metrics[0]);
            foreach (var item in items) SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(new DataGridViewCellStyle(), item.Status(projection), true, false);
            CollectionAssert.AreEqual(before, model.Items(scope).Select(i => i.Key + "|" + i.Status(projection)).ToArray());
            for (int i = 0; i < originalRows.Length; i++) Assert.AreSame(originalRows[i], projection.Rows[i].Original.AuthoritativeRow);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
            if (scope == ObjectsWorkspaceScope.Source) Assert.IsTrue(items.All(i => i.Has(DisplayEnvironmentSide.Source)));
            if (scope == ObjectsWorkspaceScope.Target) Assert.IsTrue(items.All(i => i.Has(DisplayEnvironmentSide.Target)));
        }

        [DataTestMethod]
        [DataRow("Match", 0x166534, 0xF0FDF4), DataRow("Different", 0x92400E, 0xFFFBEB)]
        [DataRow("Needs Review", 0x92400E, 0xFFFBEB), DataRow("Source Only", 0xB91C1C, 0xFEF2F2)]
        [DataRow("Target Only", 0xB91C1C, 0xFEF2F2), DataRow("Unavailable", 0x475569, 0xF1F5F9)]
        [DataRow("Unsupported", 0x92400E, 0xFFFBEB), DataRow("Indeterminate", 0x92400E, 0xFFFBEB)]
        public void ObjectStatusPaletteIsSubduedAndPreservesStatusText(string status, int foreground, int background)
        {
            var style = new DataGridViewCellStyle(); SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(style, status, true, false);
            Assert.AreEqual(Rgb(foreground), style.ForeColor); Assert.AreEqual(Rgb(background), style.BackColor);
            Assert.AreEqual(Rgb(0x1D4ED8), style.SelectionBackColor); Assert.AreEqual(Color.White, style.SelectionForeColor);
        }

        [TestMethod]
        public void EveryColumnAndStatusKeepsBlueWhiteSelectionMetadata()
        {
            foreach (string status in new[] { "Match", "Different", "Needs Review", "Source Only", "Target Only", null })
                foreach (bool statusCell in new[] { false, true })
                    foreach (bool alternate in new[] { false, true })
                    {
                        var style = new DataGridViewCellStyle(); SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(style, status, statusCell, alternate);
                        Assert.AreEqual(Rgb(0x1D4ED8), style.SelectionBackColor); Assert.AreEqual(Color.White, style.SelectionForeColor);
                    }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void NonStatusCellsRetainSubtleAlternatingRows(bool alternate)
        {
            var style = new DataGridViewCellStyle(); SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(style, "Target Only", false, alternate);
            Assert.AreEqual(Rgb(alternate ? 0xF8FAFC : 0xFFFFFF), style.BackColor); Assert.AreEqual(Rgb(0x0F172A), style.ForeColor);
        }

        [DataTestMethod, DataRow(null, null), DataRow(null, "value"), DataRow("value", null), DataRow(null, "")]
        public void UnavailablePropertyEvidenceNeverReceivesConfirmedMatchGreen(string source, string target)
        {
            var property = new ComponentPropertyPresentation("field", source, target, false, "Match"); var style = new DataGridViewCellStyle();
            SolutionObjectsExplorerControl.ProfessionalPropertyCellStyle(style, property, true, false);
            Assert.AreEqual("Unavailable", SolutionObjectsExplorerRow.PropertyComparison(property));
            Assert.AreEqual(Rgb(0xF1F5F9), style.BackColor); Assert.AreEqual("Match", property.Comparison);
            Assert.AreEqual(source, property.SourceValue); Assert.AreEqual(target, property.TargetValue);
        }

        [DataTestMethod, DataRow("Different", true, 0xFFFBEB), DataRow("Match", false, 0xF0FDF4)]
        [DataRow("Unsupported", false, 0xFFFBEB), DataRow("Not Compared", false, 0xF1F5F9), DataRow("Indeterminate", false, 0xFFFBEB)]
        public void PropertyStatusColorsKeepAvailabilityAndComparisonMeaningsSeparate(string comparison, bool changed, int background)
        {
            var property = new ComponentPropertyPresentation("field", "source", "target", changed, comparison); var style = new DataGridViewCellStyle();
            SolutionObjectsExplorerControl.ProfessionalPropertyCellStyle(style, property, true, true);
            Assert.AreEqual(Rgb(background), style.BackColor); Assert.AreEqual(comparison, property.Comparison); Assert.AreEqual(changed, property.Changed);
        }

        [TestMethod]
        public void PropertyOrderingFullValuesAndExplanationsRemainExact()
        {
            var values = new[] { new ComponentPropertyPresentation("matching", "same", "same", false, "Match"),
                new ComponentPropertyPresentation("defaultvalue", "source\r\nfull\tvalue", "target\\full", true, "Different"),
                new ComponentPropertyPresentation("unavailable", null, null, false, "Match") };
            var ordered = SolutionObjectsExplorerRow.OrderProperties(values); var explanation = SolutionObjectsExplorerRow.PropertyExplanation("Different", values, true);
            foreach (var p in ordered) foreach (bool result in new[] { false, true }) SolutionObjectsExplorerControl.ProfessionalPropertyCellStyle(new DataGridViewCellStyle(), p, result, false);
            Assert.AreSame(values[1], ordered[0]); Assert.AreEqual("source\r\nfull\tvalue", values[1].SourceValue); Assert.IsNull(values[2].SourceValue);
            Assert.AreEqual(explanation, SolutionObjectsExplorerRow.PropertyExplanation("Different", values, true));
            StringAssert.Contains(explanation, "defaultvalue");
        }

        [TestMethod]
        public void TreeSelectionUsesApprovedBlueWhiteWithoutChangingNodeLabels()
        {
            var selected = SolutionObjectsExplorerControl.ProfessionalTreeColors(true); var idle = SolutionObjectsExplorerControl.ProfessionalTreeColors(false);
            Assert.AreEqual(Rgb(0x1D4ED8), selected[0]); Assert.AreEqual(Color.White, selected[1]);
            Assert.AreEqual(Color.White, idle[0]); Assert.AreEqual(Rgb(0x334155), idle[1]);
        }

        [DataTestMethod, DataRow(300, 90, 45, true), DataRow(160, 59, 45, true), DataRow(115, 36, 23, true)]
        [DataRow(114, 36, 23, false), DataRow(80, 36, 23, false)]
        public void DetailsBudgetPreservesACompletePropertyLineWhenReadableAreasFit(int available, int identity, int explanation, bool fits)
        {
            var result = SolutionObjectsExplorerControl.ProfessionalDetailsAllocation(available, 56, 90, 45, 36, 23);
            CollectionAssert.AreEqual(new[] { identity, explanation }, result);
            Assert.AreEqual(fits, result.Sum() + 56 <= available);
            Assert.IsTrue(result[0] >= 36 && result[1] >= 23, "Too-short panes preserve readable areas and the existing explicit Expand route.");
        }

        [DataTestMethod, DataRow(96), DataRow(144), DataRow(192)]
        public void DetailsAllocationScalesWithoutChangingTheFirstPropertyRequirement(int dpi)
        {
            Func<int, int> scale = value => (int)Math.Ceiling(value * dpi / 96d);
            var result = SolutionObjectsExplorerControl.ProfessionalDetailsAllocation(scale(115), scale(56), scale(90), scale(45), scale(36), scale(23));
            Assert.IsTrue(result.Sum() + scale(56) <= scale(115));
        }

        [DataTestMethod, DataRow(900, 20, 872), DataRow(1400, 20, 1372), DataRow(600, 20, 572), DataRow(200, 20, 180)]
        public void HeaderWrappingUsesAvailableWindowWidth(int client, int padding, int expected)
        { Assert.AreEqual(expected, SimplifiedSolutionExplorerPanel.ProfessionalHeaderContentWidth(client, padding)); }

        [TestMethod]
        public void Synthetic159ObjectReferenceKeepsAuthoritativeRowsOccurrencesAndCsv()
        {
            var f = Fixture(159, false); var projection = ExplorerFixture.Project(f); var model = new SolutionObjectNavigation(projection);
            var original = f.Presentation.Rows.ToArray(); var items = model.Items(ObjectsWorkspaceScope.Compare); var csv = ExplorerFixture.Csv(f);
            Assert.AreEqual(159, items.Count); Assert.AreEqual(159, projection.Rows.Count); Assert.AreEqual(159, projection.Occurrences.Count);
            var before = items.Select(i => i.Key + "|" + i.Status(projection)).ToArray();
            Assert.AreEqual(159, SolutionObjectsExplorerControl.ProfessionalExplorerCounts(items.Select(i => i.Status(projection)))[0]);
            foreach (var item in items) SolutionObjectsExplorerControl.ProfessionalExplorerCellStyle(new DataGridViewCellStyle(), item.Status(projection), true, false);
            CollectionAssert.AreEqual(before, model.Items(ObjectsWorkspaceScope.Compare).Select(i => i.Key + "|" + i.Status(projection)).ToArray());
            Assert.AreEqual(csv, ExplorerFixture.Csv(f)); for (int i = 0; i < original.Length; i++) Assert.AreSame(original[i], projection.Rows[i].Original.AuthoritativeRow);
        }

        [TestMethod]
        public void SubsetMetricsCannotMutateTheFullScopeInventory()
        {
            var full = new[] { "Match", "Different", "Source Only", "Target Only", "Needs Review" };
            var original = full.ToArray(); var counts = SolutionObjectsExplorerControl.ProfessionalExplorerCounts(full);
            SolutionObjectsExplorerControl.ProfessionalExplorerCounts(full.Where(s => s == "Match"));
            CollectionAssert.AreEqual(original, full); CollectionAssert.AreEqual(counts, SolutionObjectsExplorerControl.ProfessionalExplorerCounts(full));
            CollectionAssert.AreEqual(new[] { 5, 1, 1, 1, 1 }, counts); // Source Only remains in the unchanged full narrative, not a fabricated extra metric.
        }

        private static ExplorerFixture Fixture(int sourceCount, bool matchedTarget)
        {
            var source = Enumerable.Range(0, sourceCount).Select(i => ExplorerFixture.Identity(61, "retained" + i)).ToArray();
            var f = ExplorerFixture.Make(source, matchedTarget ? new[] { ExplorerFixture.Identity(61, "retained0") } : null);
            Func<ComponentIdentity, DisplayEnvironmentSide, ComponentSideDisplayEvidence> display = (identity, side) => ExplorerFixture.Display(identity, side,
                friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Name " + identity?.ComparisonKey, "Synthetic offline presentation"),
                name: new ComponentDisplayValue(DisplayAvailability.Available, identity?.ComparisonKey, "Synthetic offline presentation"));
            f.Completed = new CompletedMembershipComparison(f.Presentation, f.Presentation.Rows.Select((r, i) => new ComponentDisplayEvidence(i,
                display(r.SourceIdentity, DisplayEnvironmentSide.Source), display(r.TargetIdentity, DisplayEnvironmentSide.Target))), DisplayProjectionOutcome.Completed);
            return f;
        }
        private static Color Rgb(int value) => Color.FromArgb((value >> 16) & 255, (value >> 8) & 255, value & 255);
    }
}
