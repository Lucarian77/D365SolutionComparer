using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static D365SolutionComparer.Tests.MembershipTestData;

namespace D365SolutionComparer.Tests
{
    // Synthetic presentation fixtures; no live Dataverse observations are created by these tests.
    [TestClass, TestCategory("P2AComponentPresentation")]
    public class MembershipGridPresentationTests
    {
        public TestContext TestContext { get; set; }

        [DataTestMethod]
        [DataRow("ComponentKind"), DataRow("PortableKey"), DataRow("SourcePresence"), DataRow("TargetPresence")]
        [DataRow("MembershipStatus"), DataRow("DefinitionStatus"), DataRow("ChangedProperties")]
        [DataRow("SourceResolutionStatus"), DataRow("TargetResolutionStatus"), DataRow("Diagnostic")]
        [DataRow("SourceRawComponentType"), DataRow("TargetRawComponentType")]
        public void AllExistingFieldsDelegateExactly(string property)
        {
            var f = Fixture();
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            var expected = typeof(MembershipResultRow).GetProperty(property);
            var actual = typeof(MembershipGridRow).GetProperty(property);
            Assert.AreEqual(expected.PropertyType, actual.PropertyType);
            Assert.AreEqual(expected.GetValue(f.Presentation.Rows[0]), actual.GetValue(row));
            Assert.AreSame(f.Presentation.Rows[0], row.AuthoritativeRow);
            Assert.AreSame(f.Presentation.Rows[0].DefinitionDetail, row.AuthoritativeRow.DefinitionDetail);
        }

        [TestMethod]
        public void DuplicateNamesKeysAndCategoriesKeepExactOriginalRowsAndIndices()
        {
            var f = Fixture(3);
            var rows = MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            Assert.AreEqual(3, rows.Count);
            foreach (var row in rows.Reverse())
            {
                Assert.AreSame(f.Presentation.Rows[row.OriginalRowIndex], row.AuthoritativeRow);
                Assert.AreSame(f.Completed.DisplayEvidence[row.OriginalRowIndex].SourceDisplayEvidence, row.SourceDisplayEvidence);
                Assert.AreSame(f.Completed.DisplayEvidence[row.OriginalRowIndex].TargetDisplayEvidence, row.TargetDisplayEvidence);
            }
            Assert.AreEqual(1, rows.Select(r => r.FriendlyNameText).Distinct().Count());
            Assert.AreEqual(1, rows.Select(r => r.PortableKey).Distinct().Count());
        }

        [TestMethod]
        public void EmptyCompletedComparisonIsValid()
        {
            var f = Fixture(0);
            Assert.AreEqual(0, MembershipGridRow.CreateRows(f.Presentation, f.Completed).Count);
            Assert.AreEqual(DisplayProjectionOutcome.Completed, f.Completed.DisplayProjectionOutcome);
            Assert.ThrowsException<ArgumentNullException>(() => MembershipGridRow.CreateRows(null, null));
        }

        [DataTestMethod]
        [DataRow("Legacy"), DataRow("WrongPresentation"), DataRow("Failed"), DataRow("Count")]
        [DataRow("Index"), DataRow("NullEntry"), DataRow("SwappedSides"), DataRow("NullSide")]
        [DataRow("RejectedSource"), DataRow("RejectedTarget"), DataRow("NullField")]
        public void InvalidOptionalContextNeverAttachesPartialEvidence(string problem)
        {
            var f = Fixture(2);
            var completed = f.Completed;
            if (problem == "Legacy") completed = null;
            else if (problem == "WrongPresentation") completed = Fixture(2).Completed;
            else if (problem == "Failed")
                completed = new CompletedMembershipComparison(f.Presentation, new ComponentDisplayEvidence[0], DisplayProjectionOutcome.Failed);
            else if (problem.StartsWith("Rejected", StringComparison.Ordinal))
                completed = new CompletedMembershipComparison(f.Presentation, completed.DisplayEvidence, DisplayProjectionOutcome.Completed,
                    problem == "RejectedSource" ? DisplayCacheAssociation.Rejected : DisplayCacheAssociation.Verified,
                    problem == "RejectedTarget" ? DisplayCacheAssociation.Rejected : DisplayCacheAssociation.Verified);
            else
            {
                // Corrupt only the optional synthetic wrapper to exercise defensive UI validation.
                var evidence = completed.DisplayEvidence.ToArray();
                if (problem == "Count") evidence = evidence.Take(1).ToArray();
                if (problem == "Index") evidence[1] = new ComponentDisplayEvidence(0, evidence[1].SourceDisplayEvidence, evidence[1].TargetDisplayEvidence);
                if (problem == "NullEntry") evidence[1] = null;
                if (problem == "SwappedSides") evidence[1] = new ComponentDisplayEvidence(1, evidence[1].TargetDisplayEvidence, evidence[1].SourceDisplayEvidence);
                if (problem == "NullSide") evidence[1] = new ComponentDisplayEvidence(1, null, evidence[1].TargetDisplayEvidence);
                if (problem == "NullField")
                {
                    var side = evidence[1].SourceDisplayEvidence;
                    evidence[1] = new ComponentDisplayEvidence(1, new ComponentSideDisplayEvidence(side.EnvironmentSide,
                        side.SolutionComponentId, side.LocalObjectId, side.RawComponentType, side.ComponentCategory,
                        null, side.InternalName, side.BackingCorrelationStatus, side.ParentDisplayContext), evidence[1].TargetDisplayEvidence);
                }
                SetBacking(completed, "DisplayEvidence", Array.AsReadOnly(evidence));
            }
            var rows = MembershipGridRow.CreateRows(f.Presentation, completed);
            Assert.AreEqual(2, rows.Count);
            for (int i = 0; i < rows.Count; i++)
            {
                Assert.AreSame(f.Presentation.Rows[i], rows[i].AuthoritativeRow);
                Assert.AreEqual(i, rows[i].OriginalRowIndex);
                Assert.IsNull(rows[i].SourceDisplayEvidence);
                Assert.IsNull(rows[i].TargetDisplayEvidence);
                Assert.AreEqual("[Display evidence unavailable]", rows[i].FriendlyNameText);
                Assert.IsFalse(string.IsNullOrWhiteSpace(rows[i].DisplayContextDiagnostic));
            }
        }

        [DataTestMethod]
        [DataRow("Friendly"), DataRow("Internal"), DataRow("Parent"), DataRow("Category")]
        public void EqualActualFieldsCombineWithoutLosingSideProvenance(string field)
        {
            var f = Fixture();
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            Assert.AreEqual("Both: " + FieldValue(field), Text(row, field));
            Assert.AreSame(f.Completed.DisplayEvidence[0].SourceDisplayEvidence, row.SourceDisplayEvidence);
            Assert.AreSame(f.Completed.DisplayEvidence[0].TargetDisplayEvidence, row.TargetDisplayEvidence);
        }

        [DataTestMethod]
        [DataRow("Name A", "Name B"), DataRow("Name", "name"), DataRow(" Name", "Name")]
        [DataRow("Name ", "Name"), DataRow("Name\r\nTarget: injected", "Name"), DataRow("A\tB", "A B")]
        public void DifferencesAndEmbeddedLineBreaksRemainSideSpecific(string source, string target)
        {
            var row = One(Available(source), Available(target));
            Assert.AreEqual("Source: " + MembershipGridRow.EscapeForGrid(source) + Environment.NewLine +
                "Target: " + MembershipGridRow.EscapeForGrid(target), row.FriendlyNameText);
            Assert.AreEqual(row.FriendlyNameText, row.ComponentText);
            Assert.AreEqual(source, row.SourceDisplayEvidence.FriendlyName.Value);
            Assert.AreEqual(target, row.TargetDisplayEvidence.FriendlyName.Value);
        }

        [DataTestMethod]
        [DataRow(DisplayAvailability.NotRetrieved, "[Not retrieved]")]
        [DataRow(DisplayAvailability.AttributeOmitted, "[Attribute omitted]")]
        [DataRow(DisplayAvailability.Unavailable, "[Unavailable]")]
        [DataRow(DisplayAvailability.Ambiguous, "[Ambiguous display evidence]")]
        [DataRow(DisplayAvailability.RetrievalFailed, "[Display retrieval failed]")]
        [DataRow(DisplayAvailability.NotApplicable, "[Not applicable]")]
        public void NonValuesStayDistinctForEveryDisplayField(DisplayAvailability state, string marker)
        {
            var f = Fixture(value: new ComponentDisplayValue(state));
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            foreach (var field in new[] { "Friendly", "Internal", "Parent", "Category" })
                Assert.AreEqual("Both: " + marker, Text(row, field));
            Assert.AreEqual(state, row.SourceDisplayEvidence.FriendlyName.Availability);
            Assert.IsNull(row.SourceDisplayEvidence.FriendlyName.Value);
        }

        [DataTestMethod]
        [DataRow(null, "[Retrieved null]"), DataRow("", "[Retrieved empty]"), DataRow(" \t", "[Retrieved whitespace]")]
        public void NullEmptyAndWhitespaceAreNotOmittedOrBorrowed(string value, string marker)
        {
            var row = One(Available(value), Available("Target Name"));
            Assert.AreEqual("Source: " + marker + Environment.NewLine + "Target: Target Name", row.FriendlyNameText);
            Assert.AreEqual(DisplayAvailability.Available, row.SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual(value, row.SourceDisplayEvidence.FriendlyName.Value);
        }

        [DataTestMethod]
        [DataRow(true), DataRow(false)]
        public void NoLocalObservationCannotChangePresence(bool sourceMissing)
        {
            var f = Fixture();
            var e = f.Completed.DisplayEvidence[0];
            var empty = Side(sourceMissing ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target, null, Available("Must not appear"));
            var evidence = new ComponentDisplayEvidence(0, sourceMissing ? empty : e.SourceDisplayEvidence,
                sourceMissing ? e.TargetDisplayEvidence : empty);
            var c = new CompletedMembershipComparison(f.Presentation, new[] { evidence }, DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(f.Presentation, c).Single();
            StringAssert.Contains(row.FriendlyNameText, (sourceMissing ? "Source" : "Target") + ": [No local observation]");
            Assert.IsFalse(row.FriendlyNameText.Contains("Must not appear"));
            Assert.AreEqual(f.Presentation.Rows[0].MembershipStatus, row.MembershipStatus);
        }

        [TestMethod]
        public void FutureDisplayStateIsVisibleAndCannotManufactureName()
        {
            var value = new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            SetBacking(value, "Availability", (DisplayAvailability)999);
            var row = One(value, Available("Target"));
            StringAssert.Contains(row.FriendlyNameText, "[Unsupported display state: 999]");
            Assert.AreEqual("Target", row.TargetDisplayEvidence.FriendlyName.Value);
        }

        [DataTestMethod]
        [DataRow("Friendly"), DataRow("Internal"), DataRow("Parent")]
        public void UnavailableFieldCannotUseOtherFieldsOrKeysAsFallback(string field)
        {
            var f = Fixture();
            var evidence = f.Completed.DisplayEvidence[0];
            var side = evidence.SourceDisplayEvidence;
            var unknown = new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            var changed = new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source, side.SolutionComponentId,
                side.LocalObjectId, side.RawComponentType, side.ComponentCategory,
                field == "Friendly" ? unknown : side.FriendlyName,
                field == "Internal" ? unknown : side.InternalName, side.BackingCorrelationStatus,
                field == "Parent" ? unknown : side.ParentDisplayContext);
            var completed = new CompletedMembershipComparison(f.Presentation,
                new[] { new ComponentDisplayEvidence(0, changed, evidence.TargetDisplayEvidence) }, DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(f.Presentation, completed).Single();
            StringAssert.StartsWith(Text(row, field), "Source: [Not retrieved]");
            Assert.IsFalse(Text(row, field).Contains(row.PortableKey));
            Assert.IsFalse(Text(row, field).Contains(side.LocalObjectId.ToString()));
        }

        [TestMethod]
        public void ApprovedExamplesRemainDisplayOnlySyntheticFixtures()
        {
            var app = Fixture(value: Available("I.C.M.S System"));
            var row = MembershipGridRow.CreateRows(app.Presentation, app.Completed).Single();
            Assert.AreEqual("Both: I.C.M.S System", row.FriendlyNameText);
            var column = Fixture(type: 2, value: new ComponentDisplayValue(DisplayAvailability.NotRetrieved),
                internalValue: Available("accountnumber"), parentValue: Available("account"));
            row = MembershipGridRow.CreateRows(column.Presentation, column.Completed).Single();
            Assert.AreEqual("Both: [Not retrieved]", row.FriendlyNameText);
            Assert.AreEqual("Both: accountnumber", row.InternalNameText);
            Assert.AreEqual("Both: account", row.ParentContextText);
            var setting = Fixture(type: 10075, value: new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
            row = MembershipGridRow.CreateRows(setting.Presentation, setting.Completed).Single();
            Assert.AreEqual("Both: [Not retrieved]", row.InternalNameText);
            Assert.AreEqual("Resolved", row.SourceResolutionStatus);
        }

        [DataTestMethod]
        [DataRow(300), DataRow(10072), DataRow(10266), DataRow(10267), DataRow(36)]
        [DataRow(74), DataRow(511), DataRow(10276), DataRow(31)]
        public void RecognizableUnsupportedComponentsCannotBePromoted(int type)
        {
            var f = Fixture(type: type, presence: MembershipPresence.Indeterminate, status: IdentityResolutionStatus.Unsupported);
            var rows = MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            Assert.AreEqual(1, rows.Count);
            Assert.AreEqual("Unsupported", rows[0].SourceResolutionStatus);
            Assert.AreEqual("Indeterminate - Unsupported", rows[0].MembershipStatus);
            Assert.AreEqual("(identity not resolved)", rows[0].PortableKey);
        }

        [DataTestMethod]
        [DataRow(MembershipPresence.OnlyInSource, IdentityResolutionStatus.Resolved)]
        [DataRow(MembershipPresence.OnlyInTarget, IdentityResolutionStatus.Resolved)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Ambiguous)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Unresolved)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Resolved)]
        public void AuthoritativeAccountingAndCsvNeverDependOnNames(MembershipPresence presence, IdentityResolutionStatus status)
        {
            var f = Fixture(presence: presence, status: status);
            var summary = MembershipResultsForm.FormatSummary(f.Presentation.Summary);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
            MembershipGridRow.CreateRows(f.Presentation, f.Completed);
            MembershipGridRow.CreateRows(f.Presentation, null);
            Assert.AreEqual(summary, MembershipResultsForm.FormatSummary(f.Presentation.Summary));
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(f.Presentation));
            var s = f.Presentation.Summary;
            Assert.AreEqual(s.TotalComparisonRows, s.PresentInBoth + s.SourceOnly + s.TargetOnly +
                s.IndeterminateUnsupported + s.IndeterminateUnresolved + s.IndeterminateAmbiguous +
                s.IndeterminateResolved + s.IndeterminateOtherOrUnavailable);
        }

        [DataTestMethod]
        [DataRow(false), DataRow(true)]
        public void ConstructorsLayoutDetailsAndRequestsPreserveAuthority(bool completed)
        {
            OnSta(() =>
            {
                var f = Fixture(hasDetail: false);
                var service = new RequestSpy();
                int captures = 0;
                using (var form = completed ? new MembershipResultsForm(f.Completed, captureAppSettingEvidence: () => captures++, sourceService: service, targetService: service)
                    : new MembershipResultsForm(f.Presentation, captureAppSettingEvidence: () => captures++, sourceService: service, targetService: service))
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    var columns = new[] { "ComponentText", "TypeText", "MembershipStatus", "DefinitionStatusText",
                        "ManagedText", "CustomizedText", "LastModifiedText", "OwnerText", "ObjectStatusText" };
                    CollectionAssert.AreEqual(columns, grid.Columns.Cast<DataGridViewColumn>().Select(c => c.DataPropertyName).ToArray());
                    CollectionAssert.AreEqual(new[] { "Component", "Type", "Membership Status", "Definition Status",
                        "Managed", "Customized", "Last Modified", "Owner", "Status" },
                        grid.Columns.Cast<DataGridViewColumn>().Select(c => c.HeaderText).ToArray());
                    var weights = new[] { 100F };
                    var widths = new[] { 300, 130, 175, 190, 110, 140, 200, 200, 170 };
                    for (int i = 0; i < columns.Length; i++)
                    {
                        var c = grid.Columns[i];
                        Assert.AreEqual(i < 4, c.Visible);
                        Assert.AreEqual(i < 1 ? DataGridViewAutoSizeColumnMode.Fill : DataGridViewAutoSizeColumnMode.None, c.AutoSizeMode);
                        if (i < 1) Assert.AreEqual(weights[i], c.FillWeight, 0.01F);
                        Assert.AreEqual(MembershipResultsForm.ScaleColumnWidth(i < 1 ? widths[i] : Math.Min(100, widths[i]), form.DeviceDpi), c.MinimumWidth);
                        if (i >= 1) Assert.AreEqual(MembershipResultsForm.ScaleColumnWidth(widths[i], form.DeviceDpi), c.Width);
                        Assert.IsTrue(c.ReadOnly);
                        Assert.AreEqual(DataGridViewTriState.True, c.Resizable);
                        Assert.IsFalse(c.Frozen);
                        Assert.AreEqual(DataGridViewColumnSortMode.NotSortable, c.SortMode);
                    }
                    Assert.IsFalse(grid.VirtualMode);
                    Assert.IsTrue(grid.AllowUserToResizeColumns);
                    Assert.AreEqual(DataGridViewAutoSizeRowsMode.None, grid.AutoSizeRowsMode);
                    var adapter = (MembershipGridRow)grid.Rows[0].DataBoundItem;
                    Assert.AreSame(f.Presentation.Rows[0], MembershipResultsForm.GetAuthoritativeRow(adapter));
                    Assert.IsNull(MembershipResultsForm.GetAuthoritativeRow(new object()));
                    Assert.AreSame(completed ? f.Completed : null, form.CompletedComparison);
                    Assert.AreEqual(completed ? "Both: Recognizable" : "[Display evidence unavailable]", adapter.FriendlyNameText);
                    var button = form.Controls.OfType<Panel>().Single().Controls.OfType<Panel>().Single()
                        .Controls.OfType<Button>().Single();
                    Assert.IsTrue(button.Enabled);
                    Assert.AreEqual("Display evidence...", button.AccessibleName);
                    Assert.IsTrue(button.UseMnemonic);
                    Assert.IsTrue(button.Text.Contains("&"));
                    using (var details = form.CreateDisplayEvidenceDetails(adapter))
                    {
                        details.Show(form); Application.DoEvents();
                        var text = details.Controls.OfType<TextBox>().Single();
                        Assert.IsTrue(text.ReadOnly);
                        Assert.AreEqual(ScrollBars.Both, text.ScrollBars);
                        Assert.IsFalse(text.WordWrap);
                        StringAssert.Contains(text.Text, "Original comparison row index: 0");
                        StringAssert.Contains(text.Text, "Complete portable identity: opaque:key");
                        StringAssert.Contains(text.Text, "Membership status: " + adapter.MembershipStatus);
                        StringAssert.Contains(text.Text, "Definition status: " + adapter.DefinitionStatus);
                        StringAssert.Contains(text.Text, "Source resolution: " + adapter.SourceResolutionStatus);
                        StringAssert.Contains(text.Text, "Target resolution: " + adapter.TargetResolutionStatus);
                        StringAssert.Contains(text.Text, "Source raw component type: " + adapter.SourceRawComponentType);
                        StringAssert.Contains(text.Text, "Target raw component type: " + adapter.TargetRawComponentType);
                        StringAssert.Contains(text.Text, "Component Kind (authoritative): " + adapter.ComponentKind);
                        StringAssert.Contains(text.Text, "Changed properties: " + adapter.ChangedProperties);
                        StringAssert.Contains(text.Text, "Diagnostic / Reason: " + adapter.Diagnostic);
                        text.SelectAll();
                        StringAssert.Contains(text.SelectedText, adapter.PortableKey);
                        StringAssert.Contains(text.Text, "Source (Source)");
                        StringAssert.Contains(text.Text, "Target (Target)");
                        if (completed)
                        {
                            StringAssert.Contains(text.Text, "Friendly name availability: Available");
                            StringAssert.Contains(text.Text, "source.friendly");
                            StringAssert.Contains(text.Text, "target.friendly");
                        }
                        Assert.AreSame(form, details.Owner);
                    }
                    Assert.AreEqual(0, service.RetrieveCalls + service.RetrieveMultipleCalls + service.ExecuteCalls + service.Writes + captures);
                    var summary = form.Controls.OfType<Panel>().Single().Controls.OfType<System.Windows.Forms.Label>()
                        .Single(l => l.Name == "MembershipSummary");
                    Assert.AreEqual(MembershipResultsForm.FormatSummary(f.Presentation.Summary), summary.Text);
                    Assert.IsTrue(grid.Rows[0].Selected);
                    // Inspect the native copy payload without changing the operator's clipboard.
                    var clipboard = grid.GetClipboardContent();
                    Assert.IsNotNull(clipboard);
                    StringAssert.Contains((string)clipboard.GetData(DataFormats.UnicodeText), adapter.ComponentText);
                    Assert.IsFalse(((string)clipboard.GetData(DataFormats.UnicodeText)).Contains(adapter.PortableKey));
                    StringAssert.Contains(grid.Rows[0].AccessibilityObject.Description, adapter.MembershipStatus);
                    StringAssert.Contains(grid.Rows[0].AccessibilityObject.Description, adapter.DefinitionStatus);
                    var selected = grid.Rows[0].Cells["FriendlyNameText"].InheritedStyle;
                    Assert.AreEqual(SystemColors.Highlight, selected.SelectionBackColor);
                    Assert.AreEqual(SystemColors.HighlightText, selected.SelectionForeColor);
                }
            });
        }

        [TestMethod]
        public void DefinitionDoubleClickUsesTheExactRetainedDetail()
        {
            OnSta(() =>
            {
                var f = Fixture();
                using (var form = new MembershipResultsForm(f.Completed))
                using (var timer = new System.Windows.Forms.Timer { Interval = 25 })
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    var adapter = (MembershipGridRow)form.Controls.OfType<DataGridView>().Single().Rows[0].DataBoundItem;
                    Assert.AreSame(f.Presentation.Rows[0].DefinitionDetail, adapter.AuthoritativeRow.DefinitionDetail);
                    bool opened = false;
                    timer.Tick += (s, e) =>
                    {
                        var dialog = Application.OpenForms.Cast<Form>().OfType<ComponentDefinitionDetailsForm>().FirstOrDefault();
                        if (dialog == null) return;
                        opened = true;
                        Assert.IsTrue(dialog.Controls.OfType<SplitContainer>().Any());
                        dialog.Close();
                    };
                    timer.Start();
                    typeof(MembershipResultsForm).GetMethod("ResultsGrid_CellDoubleClick", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(form, new object[] { null, new DataGridViewCellEventArgs(0, 0) });
                    Assert.IsTrue(opened);
                }
            });
        }

        [DataTestMethod]
        [DataRow(96), DataRow(120), DataRow(144), DataRow(192)]
        public void DesignWidthsScaleWithoutChangingEvidence(int dpi)
        {
            Assert.AreEqual((int)Math.Round(200 * dpi / 96.0), MembershipResultsForm.ScaleColumnWidth(200, dpi));
            Assert.AreEqual("Both: Recognizable", One(Available("Recognizable"), Available("Recognizable")).FriendlyNameText);
        }

        [TestMethod]
        public void ActualDisplayActionWorksWithoutDefinitionDetailAndNeverInvokesCollectors()
        {
            OnSta(() =>
            {
                var f = Fixture(hasDetail: false, status: IdentityResolutionStatus.Unsupported,
                    presence: MembershipPresence.Indeterminate);
                var service = new RequestSpy(); int captures = 0;
                using (var form = new MembershipResultsForm(f.Completed,
                    captureAppSettingEvidence: () => captures++, sourceService: service, targetService: service))
                using (var timer = new System.Windows.Forms.Timer { Interval = 25 })
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    // Existing double-click remains a no-op when there is no definition detail.
                    typeof(MembershipResultsForm).GetMethod("ResultsGrid_CellDoubleClick", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(form, new object[] { null, new DataGridViewCellEventArgs(0, 0) });
                    bool opened = false;
                    timer.Tick += (s, e) =>
                    {
                        var dialog = Application.OpenForms.Cast<Form>().FirstOrDefault(d => d.Text == "Display Evidence");
                        if (dialog == null) return;
                        opened = true;
                        Assert.AreSame(form, dialog.Owner);
                        Assert.IsTrue(dialog.Controls.OfType<TextBox>().Single().ReadOnly);
                        dialog.Close();
                    };
                    timer.Start();
                    var action = form.Controls.OfType<Panel>().Single().Controls.OfType<Panel>().Single().Controls.OfType<Button>().Single();
                    action.PerformClick();
                    Assert.IsTrue(opened);
                    Assert.AreEqual(0, captures + service.Total);
                }
            });
        }

        [TestMethod]
        public void DetailsKeepOriginalNewlinesAndProvenanceWithoutMutatingEvidence()
        {
            OnSta(() =>
            {
                var f = Fixture(value: Available("Name\r\nTarget: original\t\\text"));
                using (var form = new MembershipResultsForm(f.Completed))
                {
                    var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
                    StringAssert.Contains(row.FriendlyNameText, "\\r\\nTarget: original\\t\\\\text");
                    StringAssert.Contains(form.FormatDisplayEvidenceDetails(row), "Name\r\nTarget: original\t\\text");
                    Assert.AreEqual("Name\r\nTarget: original\t\\text", row.SourceDisplayEvidence.FriendlyName.Value);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Friendly"), DataRow("Internal"), DataRow("Category")]
        public void ThreeColumnValuesRetainIndependentSourceTargetDifferences(string field)
        {
            var row = One(Available("Source value"), Available("Target value"));
            Assert.AreEqual("Source: Source value" + Environment.NewLine + "Target: Target value", Text(row, field));
            Assert.AreEqual("Present in Both", row.MembershipStatus);
        }

        [DataTestMethod]
        [DataRow(MembershipPresence.PresentInBoth, IdentityResolutionStatus.Resolved)]
        [DataRow(MembershipPresence.OnlyInSource, IdentityResolutionStatus.Resolved)]
        [DataRow(MembershipPresence.OnlyInTarget, IdentityResolutionStatus.Resolved)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Unsupported)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Ambiguous)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Unresolved)]
        [DataRow(MembershipPresence.Indeterminate, IdentityResolutionStatus.Resolved)]
        public void ComparisonStatusIsAccessibleWithoutATechnicalColumn(MembershipPresence presence, IdentityResolutionStatus status)
        {
            OnSta(() =>
            {
                var f = Fixture(presence: presence, status: status);
                var spy = new RequestSpy();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    var row = (MembershipGridRow)grid.Rows[0].DataBoundItem;
                    Assert.AreEqual(9, grid.Columns.Count);
                    StringAssert.Contains(grid.Rows[0].AccessibilityObject.Description, row.MembershipStatus);
                    StringAssert.Contains(grid.Rows[0].AccessibilityObject.Description, row.DefinitionStatus);
                    var tip = (DataGridViewCellToolTipTextNeededEventArgs)Activator.CreateInstance(
                        typeof(DataGridViewCellToolTipTextNeededEventArgs), BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object[] { 0, 0, "" }, CultureInfo.InvariantCulture);
                    typeof(MembershipResultsForm).GetMethod("ResultsGrid_CellToolTipTextNeeded",
                        BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { grid, tip });
                    StringAssert.Contains(tip.ToolTipText, row.MembershipStatus);
                    StringAssert.Contains(tip.ToolTipText, row.DefinitionStatus);
                    var formatting = new DataGridViewCellFormattingEventArgs(0, 0, row.FriendlyNameText,
                        typeof(string), new DataGridViewCellStyle());
                    typeof(MembershipResultsForm).GetMethod("ResultsGrid_CellFormatting",
                        BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { grid, formatting });
                    Assert.AreEqual(presence == MembershipPresence.PresentInBoth ? Color.Honeydew :
                        presence == MembershipPresence.Indeterminate ? Color.LemonChiffon : Color.MistyRose,
                        formatting.CellStyle.BackColor);
                    StringAssert.Contains(form.FormatDisplayEvidenceDetails(row), "Membership status: " + row.MembershipStatus);
                    Assert.AreEqual(0, spy.Total);
                }
            });
        }

        [DataTestMethod]
        [DataRow(840), DataRow(1500), DataRow(10000)]
        public void ActualAdapterAndGridPerformanceIsMeasuredWithZeroRequests(int count)
        {
            OnSta(() =>
            {
                // Warm WinForms/JIT in this STA before timing, independently of test execution order.
                using (var warm = new MembershipResultsForm(Fixture().Completed))
                {
                    warm.StartPosition = FormStartPosition.Manual; warm.Location = new Point(-4000, -4000);
                    warm.Show(); Application.DoEvents();
                }
                var f = Fixture(count, value: Available(new string('n', 300)),
                    internalValue: new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
                var csvBefore = new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
                var service = new RequestSpy();
                GC.Collect(); var before = GC.GetTotalMemory(true);
                var watch = Stopwatch.StartNew();
                var adapterCount = MeasureAdapters(f, count);
                var adapterMs = watch.Elapsed.TotalMilliseconds;
                Assert.AreEqual(count, adapterCount);
                before = GC.GetTotalMemory(true); // Standalone timing adapters are not retained with the form.
                watch.Restart();
                using (var form = new MembershipResultsForm(f.Completed, sourceService: service, targetService: service))
                {
                    var constructMs = watch.Elapsed.TotalMilliseconds;
                    form.StartPosition = FormStartPosition.Manual; form.Location = new Point(-4000, -4000);
                    watch.Restart(); SelectFlat(form); form.Show(); Application.DoEvents();
                    var paintMs = watch.Elapsed.TotalMilliseconds;
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    watch.Restart(); grid.FirstDisplayedScrollingRowIndex = count - 1; Application.DoEvents();
                    var endMs = watch.Elapsed.TotalMilliseconds;
                    Assert.AreEqual(Rectangle.Empty, grid.Rows[0].AccessibilityObject.Bounds);
                    Assert.IsFalse(grid.Rows[count - 1].AccessibilityObject.Bounds.IsEmpty);
                    watch.Restart(); grid.FirstDisplayedScrollingRowIndex = 0; Application.DoEvents();
                    var backMs = watch.Elapsed.TotalMilliseconds;
                    watch.Restart(); grid.CurrentCell = grid.Rows[count - 1].Cells["FriendlyNameText"]; Application.DoEvents();
                    var selectMs = watch.Elapsed.TotalMilliseconds;
                    Assert.AreEqual(count - 1, ((MembershipGridRow)grid.CurrentRow.DataBoundItem).OriginalRowIndex);
                    watch.Restart();
                    using (var details = form.CreateDisplayEvidenceDetails((MembershipGridRow)grid.CurrentRow.DataBoundItem))
                    { details.Show(form); Application.DoEvents(); details.Close(); }
                    var detailMs = watch.Elapsed.TotalMilliseconds;
                    watch.Restart();
                    using (var closeDefinition = new System.Windows.Forms.Timer { Interval = 25 })
                    {
                        bool opened = false;
                        closeDefinition.Tick += (s, e) =>
                        {
                            var dialog = Application.OpenForms.Cast<Form>().OfType<ComponentDefinitionDetailsForm>().FirstOrDefault();
                            if (dialog == null) return;
                            opened = true; dialog.Close();
                        };
                        closeDefinition.Start();
                        typeof(MembershipResultsForm).GetMethod("ResultsGrid_CellDoubleClick",
                            BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form,
                            new object[] { grid, new DataGridViewCellEventArgs(0, count - 1) });
                        Assert.IsTrue(opened);
                    }
                    var definitionMs = watch.Elapsed.TotalMilliseconds;
                    watch.Restart();
                    using (var closeCoverage = new System.Windows.Forms.Timer { Interval = 25 })
                    {
                        bool opened = false;
                        closeCoverage.Tick += (s, e) =>
                        {
                            var dialog = Application.OpenForms.Cast<Form>().OfType<MembershipCoverageDetailsForm>().FirstOrDefault();
                            if (dialog == null) return;
                            opened = true; dialog.Close();
                        };
                        closeCoverage.Start();
                        typeof(MembershipResultsForm).GetMethod("CoverageDetails_LinkClicked", BindingFlags.Instance | BindingFlags.NonPublic)
                            .Invoke(form, new object[] { null, null });
                        Assert.IsTrue(opened);
                    }
                    var coverageMs = watch.Elapsed.TotalMilliseconds;
                    watch.Restart(); form.Size = new Size(900, 760); Application.DoEvents(); form.Size = new Size(1920, 760); Application.DoEvents();
                    var resizeMs = watch.Elapsed.TotalMilliseconds;
                    var uncollected = GC.GetTotalMemory(false); var retained = GC.GetTotalMemory(true);
                    Assert.AreEqual(count, grid.Rows.Count);
                    Assert.AreEqual(csvBefore, new MembershipCoverageCsvExporter().CreateCsv(f.Presentation));
                    Assert.AreEqual(0, service.Total);
                    TestContext.WriteLine(string.Format(CultureInfo.InvariantCulture,
                        "P2A rows={0};adapter_ms={1:F3};construct_bind_ms={2:F3};first_paint_ms={3:F3};scroll_end_ms={4:F3};scroll_back_ms={5:F3};selection_ms={6:F3};display_dialog_ms={7:F3};resize_ms={8:F3};retained_heap_delta_bytes={9};temporary_heap_proxy_bytes={10};working_set_bytes={11};requests={12};definition_dialog_ms={13:F3};coverage_dialog_ms={14:F3}",
                        count, adapterMs, constructMs, paintMs, endMs, backMs, selectMs, detailMs, resizeMs,
                        Math.Max(0, retained - before), Math.Max(0, uncollected - retained), Process.GetCurrentProcess().WorkingSet64, service.Total, definitionMs, coverageMs));
                }
            });
        }

        [TestMethod]
        public void CoverageNavigationStillReceivesOriginalPresentationAndSnapshots()
        {
            OnSta(() =>
            {
                var f = Fixture(); var spy = new RequestSpy(); int captures = 0;
                MembershipComparisonPresentation received = null;
                using (var form = new MembershipResultsForm(f.Completed,
                    captureAppSettingEvidence: () => captures++, sourceService: spy, targetService: spy))
                using (var timer = new System.Windows.Forms.Timer { Interval = 25 })
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    timer.Tick += (s, e) =>
                    {
                        var dialog = Application.OpenForms.Cast<Form>().OfType<MembershipCoverageDetailsForm>().FirstOrDefault();
                        if (dialog == null) return;
                        received = (MembershipComparisonPresentation)typeof(MembershipCoverageDetailsForm)
                            .GetField("presentation", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dialog);
                        dialog.Close();
                    };
                    timer.Start();
                    typeof(MembershipResultsForm).GetMethod("CoverageDetails_LinkClicked",
                        BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { null, null });
                    Assert.AreSame(f.Presentation, received);
                    Assert.AreSame(f.Presentation.Source.Snapshot, received.Source.Snapshot);
                    Assert.AreSame(f.Presentation.Target.Snapshot, received.Target.Snapshot);
                    Assert.AreEqual(0, captures + spy.Total);
                }
            });
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(DisplayAvailability.NotRetrieved, "[Not retrieved]")]
        [DataRow(DisplayAvailability.Unavailable, "[Unavailable]")]
        [DataRow(DisplayAvailability.AttributeOmitted, "[Attribute omitted]")]
        [DataRow(DisplayAvailability.RetrievalFailed, "[Display retrieval failed]")]
        [DataRow(DisplayAvailability.Ambiguous, "[Ambiguous display evidence]")]
        [DataRow(DisplayAvailability.NotApplicable, "[Not applicable]")]
        public void FiveObjectFieldsPreserveAvailabilityWithoutFabricatingValues(DisplayAvailability state, string indicator)
        {
            var source = Objects(DisplayEnvironmentSide.Source, state);
            var target = Objects(DisplayEnvironmentSide.Target, state);
            var row = ObjectRow(source, target);
            foreach (var text in new[] { row.ManagedText, row.CustomizedText, row.LastModifiedText, row.OwnerText, row.ObjectStatusText })
                Assert.AreEqual("Both: " + indicator, text);
            Assert.AreEqual("opaque:key", row.PortableKey);
            Assert.AreEqual("Not Compared", row.DefinitionStatus);
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(true, true), DataRow(false, false), DataRow(true, false), DataRow(false, true)]
        public void VerifiedBooleanRenderingDoesNotAffectIdentityOrDefinition(bool source, bool target)
        {
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source,
                managed: Observed<bool?>(source)), new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Target,
                managed: Observed<bool?>(target)));
            Assert.AreEqual(source == target ? "Both: " + (source ? "Yes" : "No") :
                "Source: " + (source ? "Yes" : "No") + Environment.NewLine + "Target: " + (target ? "Yes" : "No"), row.ManagedText);
            Assert.AreEqual("Both: [Not retrieved]", row.CustomizedText); // Managed=True is not Customized=False.
            Assert.AreEqual("Present in Both", row.MembershipStatus);
            Assert.AreEqual("Not Compared", row.DefinitionStatus);
            Assert.AreEqual("opaque:key", row.PortableKey);
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void AvailableNullMetadataDoesNotBecomeFalseDateOwnerOrState()
        {
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source,
                Observed<bool?>(null), Observed<bool?>(null), Observed<ComponentTimestampEvidence>(null),
                Observed<ComponentOwnerEvidence>(null), Observed<ComponentObjectStatusEvidence>(null)),
                ComponentObjectMetadataEvidence.NotRetrieved(DisplayEnvironmentSide.Target));
            foreach (var text in new[] { row.ManagedText, row.CustomizedText, row.LastModifiedText, row.OwnerText, row.ObjectStatusText })
                Assert.AreEqual("Source: [Retrieved null]" + Environment.NewLine + "Target: [Not retrieved]", text);
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(MembershipPresence.OnlyInSource), DataRow(MembershipPresence.OnlyInTarget)]
        public void MissingLocalObjectNeverBorrowsOppositeMetadata(MembershipPresence presence)
        {
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source, managed: Observed<bool?>(true)),
                new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Target, managed: Observed<bool?>(false)), presence);
            StringAssert.Contains(row.ManagedText, presence == MembershipPresence.OnlyInSource ?
                "Target: [No local observation]" : "Source: [No local observation]");
            Assert.IsFalse(row.OwnerText.Contains("opaque:key"));
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void MissingProjectionIsDifferentFromLegacyObjectNotRetrievedDefaults()
        {
            var f = Fixture();
            foreach (var completed in new[] { (CompletedMembershipComparison)null,
                new CompletedMembershipComparison(f.Presentation, new ComponentDisplayEvidence[0], DisplayProjectionOutcome.Failed),
                new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed,
                    DisplayCacheAssociation.Rejected) })
            {
                var row = MembershipGridRow.CreateRows(f.Presentation, completed).Single();
                Assert.AreEqual("[Display evidence unavailable]", row.ManagedText);
                Assert.AreEqual("[Display evidence unavailable]", row.ObjectStatusText);
                Assert.AreSame(f.Presentation.Rows[0], row.AuthoritativeRow);
            }
            Assert.AreEqual("Both: [Not retrieved]", MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single().ManagedText);
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(DateTimeKind.Utc), DataRow(DateTimeKind.Local), DataRow(DateTimeKind.Unspecified)]
        public void UninterpretedTimestampIsRenderedWithoutConversion(DateTimeKind kind)
        {
            var raw = new DateTime(2025, 11, 2, 1, 30, 0, kind);
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source,
                lastModified: Observed(new ComponentTimestampEvidence(raw))), ComponentObjectMetadataEvidence.NotRetrieved(DisplayEnvironmentSide.Target));
            StringAssert.Contains(row.LastModifiedText, "2025-11-02 01:30:00.0000000");
            StringAssert.Contains(row.LastModifiedText, "Kind=" + kind);
            StringAssert.Contains(row.LastModifiedText, "timezone not interpreted");
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(""), DataRow(" "), DataRow("Alice")]
        public void OwnerLabelsRemainExactAndLocalGuidsAreAuditOnly(string name)
        {
            var id = Guid.NewGuid();
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source,
                owner: Observed(new ComponentOwnerEvidence("systemuser", id, name, true, DisplayBackingCorrelationStatus.Unique))),
                new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Target,
                    owner: Observed(new ComponentOwnerEvidence("systemuser", Guid.NewGuid()))));
            StringAssert.Contains(row.OwnerText, "[Owner display name not verified]");
            Assert.IsFalse(row.OwnerText.Contains(id.ToString()));
            if (name == "") StringAssert.Contains(row.OwnerText, "[Retrieved empty]");
            else if (name == " ") StringAssert.Contains(row.OwnerText, "[Retrieved whitespace]");
            else StringAssert.Contains(row.OwnerText, name);
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void ObjectStatusHasItsOwnVerifiedLabelAndNoUniversalMapping()
        {
            var row = ObjectRow(new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source,
                objectStatus: Observed(new ComponentObjectStatusEvidence("statecode", 2, "Suspended", true, "Synthetic workflow category"))),
                new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Target,
                    objectStatus: Observed(new ComponentObjectStatusEvidence("statecode", 99))));
            StringAssert.Contains(row.ObjectStatusText, "Suspended (code 2)");
            StringAssert.Contains(row.ObjectStatusText, "[Object status label not verified] (code 99)");
            Assert.IsFalse(row.ObjectStatusText.Contains("Present in Both"));
            Assert.IsFalse(row.ObjectStatusText.Contains("Not Compared"));
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void ObjectDetailsRetainTypedInterpretationAndIndependentSidesWithZeroRequests()
        {
            OnSta(() =>
            {
                var f = Fixture(); var spy = new RequestSpy(); var raw = new DateTime(2025, 11, 2, 1, 30, 0, DateTimeKind.Unspecified);
                var ownerId = Guid.NewGuid();
                var source = new ComponentObjectMetadataEvidence(DisplayEnvironmentSide.Source, Observed<bool?>(true),
                    lastModified: Observed(new ComponentTimestampEvidence(raw, "UserLocal", "Synthetic offset proof", new DateTimeOffset(raw, TimeSpan.FromHours(-5)))),
                    owner: Observed(new ComponentOwnerEvidence("team", ownerId, "Synthetic Team", true, DisplayBackingCorrelationStatus.Unique)),
                    objectStatus: Observed(new ComponentObjectStatusEvidence("statecode", 1, "Active", true, "Synthetic object state")));
                var completed = WithObjects(f, source, ComponentObjectMetadataEvidence.NotRetrieved(DisplayEnvironmentSide.Target));
                using (var form = new MembershipResultsForm(completed, sourceService: spy, targetService: spy))
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    var grid = form.Controls.OfType<DataGridView>().Single(); var row = (MembershipGridRow)grid.Rows[0].DataBoundItem;
                    StringAssert.Contains(grid.Columns["ObjectStatusText"].ToolTipText, "not Membership Status");
                    Assert.AreEqual(row.MembershipStatus, grid.Rows[0].Cells["MembershipStatus"].Value);
                    Assert.AreEqual(row.DefinitionStatus, grid.Rows[0].Cells["DefinitionStatus"].Value);
                    var text = form.FormatDisplayEvidenceDetails(row);
                    foreach (var field in new[] { "Managed", "Customized", "Last Modified", "Owner", "Object Status" })
                    { StringAssert.Contains(text, field + " availability:"); StringAssert.Contains(text, field + " requested:"); StringAssert.Contains(text, field + " provenance:"); }
                    StringAssert.Contains(text, "Timezone provenance: Synthetic offset proof");
                    StringAssert.Contains(text, "Timestamp kind: Unspecified");
                    StringAssert.Contains(text, "Owner ID (local audit): " + ownerId);
                    StringAssert.Contains(text, "Object status field: statecode");
                    StringAssert.Contains(text, "Object status raw numeric value: 1");
                    StringAssert.Contains(text, "Detached object metadata side: Target");
                    Assert.AreEqual(0, spy.Total);
                }
            });
        }

        // These tests exercise native Flat-grid layout, accessibility and actions.
        private static void SelectFlat(MembershipResultsForm form) => form.Controls.OfType<ToolStrip>()
            .Single(t => t.Name == "MembershipViewBar").Items.OfType<ToolStripComboBox>()
            .Single(c => c.Name == "MembershipViewSelector").SelectedIndex = 0;

        private static ComponentObjectMetadataValue<T> Observed<T>(T value) =>
            new ComponentObjectMetadataValue<T>(DisplayAvailability.Available, value, "synthetic.object.field", true);

        private static ComponentObjectMetadataEvidence Objects(DisplayEnvironmentSide side, DisplayAvailability state)
        {
            bool requested = state == DisplayAvailability.AttributeOmitted || state == DisplayAvailability.RetrievalFailed;
            return new ComponentObjectMetadataEvidence(side,
                new ComponentObjectMetadataValue<bool?>(state, provenance: "synthetic.field", wasRequested: requested),
                new ComponentObjectMetadataValue<bool?>(state, provenance: "synthetic.field", wasRequested: requested),
                new ComponentObjectMetadataValue<ComponentTimestampEvidence>(state, provenance: "synthetic.field", wasRequested: requested),
                new ComponentObjectMetadataValue<ComponentOwnerEvidence>(state, provenance: "synthetic.field", wasRequested: requested),
                new ComponentObjectMetadataValue<ComponentObjectStatusEvidence>(state, provenance: "synthetic.field", wasRequested: requested));
        }

        private static MembershipGridRow ObjectRow(ComponentObjectMetadataEvidence source, ComponentObjectMetadataEvidence target,
            MembershipPresence presence = MembershipPresence.PresentInBoth)
        {
            var f = Fixture(presence: presence);
            return MembershipGridRow.CreateRows(f.Presentation, WithObjects(f, source, target)).Single();
        }

        private static CompletedMembershipComparison WithObjects(Data f, ComponentObjectMetadataEvidence source, ComponentObjectMetadataEvidence target)
        {
            var row = f.Presentation.Rows[0];
            return new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                Side(DisplayEnvironmentSide.Source, row.SourceIdentity, Available("Recognizable"), objectMetadata: source),
                Side(DisplayEnvironmentSide.Target, row.TargetIdentity, Available("Recognizable"), objectMetadata: target)) }, DisplayProjectionOutcome.Completed);
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static int MeasureAdapters(Data fixture, int count) =>
            MembershipGridRow.CreateRows(fixture.Presentation, fixture.Completed).Count;

        [DataTestMethod, TestCategory("MembershipUiPolish")]
        [DataRow("Friendly", "Friendly")]
        [DataRow(null, "logical")]
        [DataRow("", "logical")]
        [DataRow("  ", "logical")]
        public void ComponentPrefersVerifiedDisplayThenVerifiedLogicalName(string friendly, string expected)
        {
            var f = Fixture(value: Available(friendly), internalValue: Available("logical"));
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            Assert.AreEqual(expected, row.ComponentText);
            Assert.IsFalse(row.ComponentText.StartsWith("Both:", StringComparison.Ordinal));
            StringAssert.Contains(row.ComponentToolTip, "provenance:");
            Assert.IsFalse(row.ComponentToolTip.Contains("Both: "));
        }

        [TestMethod, TestCategory("MembershipUiPolish")]
        public void NameFallbackRemainsIndependentOnSourceAndTarget()
        {
            var f = Fixture(); var original = f.Presentation.Rows[0];
            var completed = new CompletedMembershipComparison(f.Presentation, new[] {
                new ComponentDisplayEvidence(0,
                    Side(DisplayEnvironmentSide.Source, original.SourceIdentity, new ComponentDisplayValue(DisplayAvailability.NotRetrieved), Available("logical")),
                    Side(DisplayEnvironmentSide.Target, original.TargetIdentity, Available("Display"), Available("logical"))) }, DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(f.Presentation, completed).Single();
            Assert.AreEqual("Source: logical" + Environment.NewLine + "Target: Display", row.ComponentText);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, row.SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual("logical", row.TargetDisplayEvidence.InternalName.Value);
        }

        [DataTestMethod, TestCategory("MembershipUiPolish")]
        [DataRow(IdentityResolutionStatus.Unsupported)]
        [DataRow(IdentityResolutionStatus.Unresolved)]
        [DataRow(IdentityResolutionStatus.Ambiguous)]
        public void UnverifiedAppSettingCannotManufactureCompositeOrNames(IdentityResolutionStatus status)
        {
            var identity = Identity(null, 10075, status, ComponentSemanticKinds.AppSetting);
            var p = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(identity), 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(), 0, TimeSpan.Zero));
            var completed = new CompletedMembershipComparison(p, ComponentDisplayEvidenceProjector.Project(p), DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(p, completed).Single();
            Assert.IsFalse(row.ComponentText.Contains("appsetting:v1:"));
            Assert.AreEqual(status.ToString(), row.SourceResolutionStatus);
            Assert.AreEqual(MembershipPresence.Indeterminate, row.AuthoritativeRow.Presence);
        }

        [DataTestMethod, TestCategory("MembershipUiPolish")]
        [DataRow(DisplayAvailability.NotRetrieved)]
        [DataRow(DisplayAvailability.AttributeOmitted)]
        [DataRow(DisplayAvailability.Ambiguous)]
        [DataRow(DisplayAvailability.RetrievalFailed)]
        public void MissingFriendlyNameUsesOnlyVerifiedLogicalObservation(DisplayAvailability availability)
        {
            var f = Fixture(value: new ComponentDisplayValue(availability), internalValue: Available("verified.logical"));
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            Assert.AreEqual("verified.logical", row.ComponentText);
            Assert.AreEqual(availability, row.SourceDisplayEvidence.FriendlyName.Availability);
            StringAssert.Contains(row.ComponentToolTip, availability.ToString());
            var missing = Fixture(value: new ComponentDisplayValue(availability));
            var missingRow = MembershipGridRow.CreateRows(missing.Presentation, missing.Completed).Single();
            Assert.IsFalse(missingRow.ComponentText.Contains(missingRow.PortableKey));
        }

        [TestMethod, TestCategory("MembershipUiPolish")]
        public void IdenticalDisplayNamesRetainDifferingVerifiedLogicalNames()
        {
            var f = Fixture(); var original = f.Presentation.Rows[0];
            var completed = new CompletedMembershipComparison(f.Presentation, new[] {
                new ComponentDisplayEvidence(0,
                    Side(DisplayEnvironmentSide.Source, original.SourceIdentity, Available("Shared display"), Available("source.logical")),
                    Side(DisplayEnvironmentSide.Target, original.TargetIdentity, Available("Shared display"), Available("target.logical"))) }, DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(f.Presentation, completed).Single();
            Assert.AreEqual("Source: Shared display (source.logical)" + Environment.NewLine +
                "Target: Shared display (target.logical)", row.ComponentText);
            Assert.AreSame(original, row.AuthoritativeRow);
        }

        [DataTestMethod, TestCategory("MembershipUiPolish")]
        [DataRow(MembershipPresence.OnlyInSource)]
        [DataRow(MembershipPresence.OnlyInTarget)]
        public void OneSidedComponentRetainsExplicitEnvironmentLabels(MembershipPresence presence)
        {
            var f = Fixture(presence: presence);
            var row = MembershipGridRow.CreateRows(f.Presentation, f.Completed).Single();
            StringAssert.Contains(row.ComponentText, "Source: " + (presence == MembershipPresence.OnlyInSource ? "Recognizable" : "[No local observation]"));
            StringAssert.Contains(row.ComponentText, "Target: " + (presence == MembershipPresence.OnlyInTarget ? "Recognizable" : "[No local observation]"));
            Assert.AreEqual(f.Presentation.Rows[0].MembershipStatus, row.MembershipStatus);
        }

        [TestMethod, TestCategory("MembershipUiPolish")]
        public void AppSettingUsesExistingCompositeWithoutInventingNamesOrComparingValues()
        {
            var source = Identity(AppSettingResolutionOperation.PortableKey("app.internal", "setting.internal"), 10075,
                kind: ComponentSemanticKinds.AppSetting);
            var target = Identity(source.ComparisonKey, 10075, kind: ComponentSemanticKinds.AppSetting);
            var p = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(source), 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(target), 0, TimeSpan.Zero));
            var completed = new CompletedMembershipComparison(p, ComponentDisplayEvidenceProjector.Project(p), DisplayProjectionOutcome.Completed);
            var row = MembershipGridRow.CreateRows(p, completed).Single();
            Assert.AreEqual(source.ComparisonKey, row.ComponentText);
            Assert.AreEqual("Present in Both", row.MembershipStatus);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, row.SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, row.SourceDisplayEvidence.ParentDisplayContext.Availability);
            StringAssert.Contains(row.ComponentToolTip, "not retrieved display names");
            Assert.AreSame(p.Rows[0], row.AuthoritativeRow);
        }

        [DataTestMethod, TestCategory("MembershipUiPolish")]
        [DataRow("Match", "Match (compared properties)")]
        [DataRow("Different", "Different")]
        [DataRow("Not Compared", "Not Compared")]
        [DataRow("Unresolved", "Unresolved")]
        [DataRow("Unsupported", "Unsupported")]
        public void DefinitionLabelQualifiesOnlyActualPropertyMatch(string status, string expected)
        {
            var f = Fixture(); var original = f.Presentation.Rows[0].WithDefinition(status, "", f.Presentation.Rows[0].DefinitionDetail);
            var p = new MembershipComparisonPresentation(f.Presentation.SolutionUniqueName, f.Presentation.Source, f.Presentation.Target, new[] { original });
            var completed = new CompletedMembershipComparison(p, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            var row = MembershipGridRow.CreateRows(p, completed).Single();
            Assert.AreEqual(expected, row.DefinitionStatusText);
            Assert.AreEqual(status, row.DefinitionStatus);
            Assert.AreSame(original.DefinitionDetail, row.AuthoritativeRow.DefinitionDetail);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p));
        }

        [TestMethod, TestCategory("MembershipUiPolish")]
        public void MetadataIsOptInAndFormattingDoesNotHideRetrievalLimitationsOrChangeResults()
        {
            OnSta(() =>
            {
                var f = Fixture(); var spy = new RequestSpy();
                var csv = new MembershipCoverageCsvExporter().CreateCsv(f.Presentation);
                using (var form = new MembershipResultsForm(f.Completed, sourceService: spy, targetService: spy))
                {
                    SelectFlat(form); form.Show(); Application.DoEvents();
                    var grid = form.Controls.OfType<DataGridView>().Single();
                    var toggle = form.Controls.OfType<Panel>().Single().Controls.OfType<Panel>().Single().Controls.OfType<CheckBox>().Single();
                    CollectionAssert.AreEqual(new[] { "Component", "Type", "Membership Status", "Definition Status" },
                        grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).Select(c => c.HeaderText).ToArray());
                    Assert.AreEqual("Recognizable", grid.Rows[0].Cells["FriendlyNameText"].FormattedValue);
                    Assert.AreEqual("Category", grid.Rows[0].Cells["ComponentCategoryText"].FormattedValue);
                    toggle.Checked = true; Application.DoEvents();
                    Assert.AreEqual(9, grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible));
                    foreach (var name in new[] { "ManagedText", "CustomizedText", "LastModifiedText", "OwnerText", "ObjectStatusText" })
                        Assert.AreEqual("[Not retrieved]", grid.Rows[0].Cells[name].FormattedValue);
                    var row = (MembershipGridRow)grid.Rows[0].DataBoundItem;
                    StringAssert.Contains(form.FormatDisplayEvidenceDetails(row), "Managed availability: NotRetrieved");
                    StringAssert.Contains(form.FormatDisplayEvidenceDetails(row), "Managed requested: False");
                    toggle.Checked = false; Application.DoEvents();
                    Assert.AreEqual(4, grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible));
                    Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(f.Presentation));
                    Assert.AreSame(f.Presentation.Rows[0], row.AuthoritativeRow);
                    Assert.AreEqual(0, spy.Total);
                }
            });
        }

        [TestMethod, TestCategory("MembershipUiPolish")]
        public void SyntheticEduAccountingShapeRetainsAll840AuthoritativeRowsAndThreeResolvedIndeterminate()
        {
            // Aggregate regression only. These generated records do not identify the three live SAA occurrences.
            var source = new List<ComponentIdentity>(); var target = new List<ComponentIdentity>();
            for (int i = 0; i < 672; i++) { source.Add(Identity("shared" + i)); target.Add(Identity("shared" + i)); }
            for (int i = 0; i < 103; i++) source.Add(Identity("source-only" + i));
            for (int i = 0; i < 32; i++) target.Add(Identity("target-only" + i));
            for (int i = 0; i < 13; i++) { source.Add(Identity(null, 300, IdentityResolutionStatus.Unsupported)); target.Add(Identity(null, 300, IdentityResolutionStatus.Unsupported)); }
            for (int i = 0; i < 2; i++) { source.Add(Identity(null, 92, IdentityResolutionStatus.Ambiguous)); target.Add(Identity(null, 92, IdentityResolutionStatus.Ambiguous)); }
            Func<string, ComponentIdentity> matchOnly = key => new ComponentIdentity(
                new SolutionComponentRecord(Guid.NewGuid(), 60, Guid.NewGuid()), IdentityResolutionStatus.Resolved, key,
                componentTypeKey: ComponentSemanticKinds.SystemForm, inventoryAbsencePolicy: InventoryAbsencePolicy.MatchOnly);
            source.Add(matchOnly("synthetic-source-indeterminate"));
            target.Add(matchOnly("synthetic-target-indeterminate-one")); target.Add(matchOnly("synthetic-target-indeterminate-two"));
            var p = new MembershipResultPresenter().Create(
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(source.ToArray()), 27, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(target.ToArray()), 27, TimeSpan.Zero));
            var completed = new CompletedMembershipComparison(p, ComponentDisplayEvidenceProjector.Project(p), DisplayProjectionOutcome.Completed);
            var csv = new MembershipCoverageCsvExporter().CreateCsv(p);
            var rows = MembershipGridRow.CreateRows(p, completed);
            Assert.AreEqual(840, rows.Count); Assert.AreEqual(672, p.Summary.PresentInBoth);
            Assert.AreEqual(103, p.Summary.SourceOnly); Assert.AreEqual(32, p.Summary.TargetOnly);
            Assert.AreEqual(26, p.Summary.IndeterminateUnsupported); Assert.AreEqual(4, p.Summary.IndeterminateAmbiguous);
            Assert.AreEqual(3, p.Summary.IndeterminateResolved);
            Assert.AreEqual(776, p.Source.Diagnostics.ResolvedCount); Assert.AreEqual(706, p.Target.Diagnostics.ResolvedCount);
            foreach (var row in rows) Assert.AreSame(p.Rows[row.OriginalRowIndex], row.AuthoritativeRow);
            foreach (var row in rows.Where(r => r.AuthoritativeRow.Presence == MembershipPresence.Indeterminate && r.SourceResolutionStatus == "Resolved" ||
                r.AuthoritativeRow.Presence == MembershipPresence.Indeterminate && r.TargetResolutionStatus == "Resolved"))
            {
                Assert.AreEqual("Indeterminate", row.MembershipStatus);
                Assert.AreEqual(MembershipAbsenceEvidence.None, row.AuthoritativeRow.Comparison.AbsenceEvidence);
                StringAssert.Contains(row.Diagnostic, "cannot use complete inventory");
            }
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(p));
        }

        private static void SetBacking(object value, string property, object replacement) =>
            value.GetType().GetField("<" + property + ">k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(value, replacement);
        private static ComponentDisplayValue Available(string value, string provenance = "synthetic.field") =>
            new ComponentDisplayValue(DisplayAvailability.Available, value, provenance);
        private static string FieldValue(string field) => field == "Internal" ? "internal" : field == "Parent" ? "parent" : field == "Category" ? "Category" : "Recognizable";
        private static string Text(MembershipGridRow row, string field) => field == "Internal" ? row.InternalNameText :
            field == "Parent" ? row.ParentContextText : field == "Category" ? row.ComponentCategoryText : row.FriendlyNameText;
        private static MembershipGridRow One(ComponentDisplayValue source, ComponentDisplayValue target)
        {
            var f = Fixture();
            var row = f.Presentation.Rows[0];
            var completed = new CompletedMembershipComparison(f.Presentation,
                new[] { new ComponentDisplayEvidence(0, Side(DisplayEnvironmentSide.Source, row.SourceIdentity, source),
                    Side(DisplayEnvironmentSide.Target, row.TargetIdentity, target)) }, DisplayProjectionOutcome.Completed);
            return MembershipGridRow.CreateRows(f.Presentation, completed).Single();
        }
        private static ComponentSideDisplayEvidence Side(DisplayEnvironmentSide side, ComponentIdentity identity, ComponentDisplayValue value,
            ComponentDisplayValue internalValue = null, ComponentDisplayValue parent = null, ComponentDisplayValue category = null,
            ComponentObjectMetadataEvidence objectMetadata = null) =>
            new ComponentSideDisplayEvidence(side, identity?.Record.SolutionComponentId, identity?.Record.ObjectId,
                identity?.Record.ComponentType, category ?? value, value, internalValue ?? value,
                DisplayBackingCorrelationStatus.Unique, parent ?? value, objectMetadata);
        private sealed class Data { internal MembershipComparisonPresentation Presentation; internal CompletedMembershipComparison Completed; }
        private static Data Fixture(int count = 1, int type = 80, ComponentDisplayValue value = null,
            ComponentDisplayValue internalValue = null, ComponentDisplayValue parentValue = null,
            MembershipPresence presence = MembershipPresence.PresentInBoth, IdentityResolutionStatus status = IdentityResolutionStatus.Resolved,
            bool hasDetail = true)
        {
            var rows = new List<MembershipResultRow>(); var evidence = new List<ComponentDisplayEvidence>();
            for (int i = 0; i < count; i++)
            {
                var source = presence == MembershipPresence.OnlyInTarget ? null : Identity("opaque:key", type, status);
                var target = presence == MembershipPresence.OnlyInSource ? null : Identity("opaque:key", type, status);
                var membership = presence == MembershipPresence.OnlyInSource ? "Source Only" : presence == MembershipPresence.OnlyInTarget ?
                    "Target Only" : presence == MembershipPresence.Indeterminate ? "Indeterminate" + (status == IdentityResolutionStatus.Resolved ? "" : " - " + status) : "Present in Both";
                var detail = hasDetail ? new ComponentDefinitionDetailPresentation("Kind", "opaque:key", membership, "Not Compared",
                    new ComponentPropertyPresentation[0], "Diagnostic", new[] { "Source evidence" }, new[] { "Target evidence" }) : null;
                rows.Add(new MembershipResultRow(presence, source, target, "Kind", status == IdentityResolutionStatus.Resolved ? "opaque:key" : "(identity not resolved)",
                    source == null ? "Missing" : "Present", target == null ? "Missing" : "Present", membership,
                    source?.Status.ToString() ?? "Not present", target?.Status.ToString() ?? "Not present", "Diagnostic",
                    definitionStatus: "Not Compared", changedProperties: "changed", definitionDetail: detail));
                evidence.Add(new ComponentDisplayEvidence(i,
                    Side(DisplayEnvironmentSide.Source, source, value ?? Available("Recognizable", "source.friendly"),
                        internalValue ?? (value == null ? Available("internal", "source.internal") : value), parentValue ?? (value == null ? Available("parent") : value), value ?? Available("Category")),
                    Side(DisplayEnvironmentSide.Target, target, value ?? Available("Recognizable", "target.friendly"),
                        internalValue ?? (value == null ? Available("internal", "target.internal") : value), parentValue ?? (value == null ? Available("parent") : value), value ?? Available("Category"))));
            }
            var presentation = new MembershipComparisonPresentation("sample",
                MembershipEnvironmentResult.FromSnapshot("Source", Snapshot(rows.Where(r => r.SourceIdentity != null).Select(r => r.SourceIdentity).ToArray()), 27, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Target", Snapshot(rows.Where(r => r.TargetIdentity != null).Select(r => r.TargetIdentity).ToArray()), 27, TimeSpan.Zero), rows);
            return new Data { Presentation = presentation, Completed = new CompletedMembershipComparison(presentation, evidence,
                DisplayProjectionOutcome.Completed, DisplayCacheAssociation.Verified, DisplayCacheAssociation.Verified) };
        }
        private static void OnSta(Action action)
        {
            Exception error = null;
            var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(120)), "Presentation UI did not complete.");
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
        private sealed class RequestSpy : IOrganizationService
        {
            internal int RetrieveCalls, RetrieveMultipleCalls, ExecuteCalls, Writes;
            internal int Total => RetrieveCalls + RetrieveMultipleCalls + ExecuteCalls + Writes;
            public Entity Retrieve(string name, Guid id, ColumnSet columns) { RetrieveCalls++; throw new NotSupportedException(); }
            public EntityCollection RetrieveMultiple(QueryBase query) { RetrieveMultipleCalls++; throw new NotSupportedException(); }
            public OrganizationResponse Execute(OrganizationRequest request) { ExecuteCalls++; throw new NotSupportedException(); }
            public Guid Create(Entity entity) { Writes++; throw new NotSupportedException(); }
            public void Update(Entity entity) { Writes++; throw new NotSupportedException(); }
            public void Delete(string name, Guid id) { Writes++; throw new NotSupportedException(); }
            public void Associate(string name, Guid id, Relationship relationship, EntityReferenceCollection entities) { Writes++; throw new NotSupportedException(); }
            public void Disassociate(string name, Guid id, Relationship relationship, EntityReferenceCollection entities) { Writes++; throw new NotSupportedException(); }
        }
    }
}
