using System;
using System.Linq;
using D365SolutionComparer.Models.Identity;
using D365SolutionComparer.Models.ComponentDetails;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Models.Presentation;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass, TestCategory("H2DExplorerPresentation")]
    public class SolutionObjectsExplorerPresentationTests
    {
        [DataTestMethod]
        [DataRow(DisplayAvailability.AttributeOmitted), DataRow(DisplayAvailability.Unavailable), DataRow(DisplayAvailability.NotRetrieved),
         DataRow(DisplayAvailability.RetrievalFailed), DataRow(DisplayAvailability.Ambiguous), DataRow(DisplayAvailability.NotApplicable)]
        public void AvailabilityIsNeverInventedMetadata(DisplayAvailability availability)
        {
            var f = Named(new ComponentDisplayValue(availability)); var row = ExplorerFixture.Project(f).Rows[0];
            Assert.AreEqual("[" + availability + "]", row.DisplayName); StringAssert.Contains(row.Name, "Identity:");
        }
        [DataTestMethod]
        [DataRow(null, "[Retrieved null]"), DataRow("", "[Retrieved empty]"), DataRow("Case", "Case")]
        public void ActualValuesRemainDistinctFromOmission(string value, string expected)
        {
            var row = ExplorerFixture.Project(Named(new ComponentDisplayValue(DisplayAvailability.Available, value, "Synthetic label"))).Rows[0];
            Assert.AreEqual(expected, row.DisplayName); Assert.IsFalse(row.DisplayName.StartsWith("Both:"));
        }
        [TestMethod]
        public void DifferentSideNamesAndPresenceRemainExplicit()
        {
            var f = Named(new ComponentDisplayValue(DisplayAvailability.Available, "Left", "Synthetic label"),
                new ComponentDisplayValue(DisplayAvailability.Available, "Right", "Synthetic label")); var row = ExplorerFixture.Project(f).Rows[0];
            StringAssert.Contains(row.DisplayName, "Source: Left"); StringAssert.Contains(row.DisplayName, "Target: Right");
            Assert.AreEqual("Present", row.Source); Assert.AreEqual("Present", row.Target);
        }
        [TestMethod]
        public void DefinitionAndMembershipAreNotAnAggregateMatch()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(26, "savedquery-key") }, new[] { ExplorerFixture.Identity(26, "savedquery-key") });
            var p = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target,
                f.Presentation.Rows.Select(r => r.WithDefinition("Not Compared", "", null)));
            var row = SolutionObjectsExplorerProjector.Project(p, null, ComponentHierarchyProjector.Project(p)).Rows[0];
            StringAssert.Contains(row.ComparisonStatus, "Membership: Present in Both"); StringAssert.Contains(row.ComparisonStatus, "Definition: Not Compared");
        }
        [TestMethod]
        public void QualifiedColumnsAndAppSettingNamesReuseExistingFormatting()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(2, "appointment.regardingobjectid"), ExplorerFixture.Identity(10075, ExplorerFixture.Key("app", "OnlineMode")) });
            var p = ExplorerFixture.Project(f);
            Assert.IsTrue(p.Rows.Any(r => r.Name == "Source: Identity: appointment.regardingobjectid" + Environment.NewLine + "Target: [No local row]"));
            Assert.IsTrue(p.Rows.Any(r => r.Name == "Source: Identity: OnlineMode" + Environment.NewLine + "Target: [No local row]"));
            Assert.IsTrue(p.Rows.All(r => r.DisplayName == "Source: [NotRetrieved]" + Environment.NewLine + "Target: [No local row]"));
        }
        [TestMethod]
        public void MismatchedLocalDisplayEvidenceCannotSupplyNames()
        {
            var f = Named(new ComponentDisplayValue(DisplayAvailability.Available, "Case", "Synthetic label"));
            var i = f.Presentation.Rows[0]; var wrong = ExplorerFixture.Identity(1, "different");
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                ExplorerFixture.Display(wrong, DisplayEnvironmentSide.Source, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Wrong", "Synthetic label")),
                ExplorerFixture.Display(i.TargetIdentity, DisplayEnvironmentSide.Target)) }, DisplayProjectionOutcome.Completed);
            var row = ExplorerFixture.Project(f).Rows[0]; Assert.IsFalse(row.DisplayName.Contains("Wrong")); Assert.IsNull(row.SourceDisplay);
        }
        [TestMethod]
        public void AuthoritativeFamilyRemainsReadableWithoutCategoryMetadata()
        {
            Assert.AreEqual("Table", ExplorerFixture.Project(ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "account") })).Rows[0].Type);
        }
        [DataTestMethod, TestCategory("D4T8R2C")]
        [DataRow("Friendly", "logical", "Friendly", "Retained friendly")]
        [DataRow(null, "logical", "logical", "Internal/logical name fallback")]
        [DataRow("", "logical", "logical", "Internal/logical name fallback")]
        [DataRow(null, null, null, "Name unavailable")]
        [DataRow("", "", null, "Name unavailable")]
        [DataRow("11111111-1111-1111-1111-111111111111", "11111111-1111-1111-1111-111111111111", null, "Name unavailable")]
        public void R2CNamesPreferVerifiedFriendlyThenExplainLogicalFallback(string friendly, string name, string expected, string basis)
        {
            var f = R2CNamed(new ComponentDisplayValue(DisplayAvailability.Available, friendly, "Synthetic friendly"),
                new ComponentDisplayValue(DisplayAvailability.Available, name, "Synthetic logical"));
            var original = f.Presentation.Rows[0]; var csv = ExplorerFixture.Csv(f); var row = ExplorerFixture.Project(f).Rows[0];
            Assert.AreEqual(expected, row.PrimaryName(DisplayEnvironmentSide.Source, true));
            Assert.AreEqual(name == null || name.Length == 0 || Guid.TryParse(name, out _) ? null : name, row.PrimaryName(DisplayEnvironmentSide.Target, false));
            StringAssert.Contains(row.NameBasis(DisplayEnvironmentSide.Source), basis);
            Assert.AreSame(original, row.Original.AuthoritativeRow); Assert.AreEqual(0, row.OriginalRowIndex);
            Assert.AreEqual(csv, ExplorerFixture.Csv(f));
        }
        [DataTestMethod, TestCategory("D4T8R2C")]
        [DataRow(DisplayAvailability.NotRetrieved), DataRow(DisplayAvailability.AttributeOmitted), DataRow(DisplayAvailability.Unavailable),
         DataRow(DisplayAvailability.RetrievalFailed), DataRow(DisplayAvailability.Ambiguous), DataRow(DisplayAvailability.NotApplicable)]
        public void R2CUnavailableNamesKeepTheirExactAvailability(DisplayAvailability availability)
        {
            var field = new ComponentDisplayValue(availability); var row = ExplorerFixture.Project(R2CNamed(field, field)).Rows[0];
            Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Source, true)); Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Target, false));
            StringAssert.Contains(row.NameBasis(DisplayEnvironmentSide.Source), "[" + availability + "]");
            Assert.AreEqual("[" + availability + "]", row.DisplayName, "The existing technical adapter stays compatible.");
        }
        [DataTestMethod, TestCategory("D4T8R2C"), DataRow(31), DataRow(36), DataRow(300)]
        public void R2CRetainedDuplicateReportTemplateAndCanvasNamesNeverPairRows(int type)
        {
            var f = ExplorerFixture.Make(Enumerable.Range(0, 2).Select(_ => ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported)),
                Enumerable.Range(0, 2).Select(_ => ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported)));
            var originals = f.Presentation.Rows.ToArray(); var csv = ExplorerFixture.Csv(f);
            Func<ComponentIdentity, DisplayEnvironmentSide, ComponentSideDisplayEvidence> display = (id, side) => ExplorerFixture.Display(id, side,
                friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Same retained title", "Synthetic retained title"),
                name: new ComponentDisplayValue(DisplayAvailability.Available, "same_internal", "Synthetic retained name"));
            f.Completed = new CompletedMembershipComparison(f.Presentation, originals.Select((r, i) => new ComponentDisplayEvidence(i,
                r.SourceIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Source) : display(r.SourceIdentity, DisplayEnvironmentSide.Source),
                r.TargetIdentity == null ? ExplorerFixture.Display(null, DisplayEnvironmentSide.Target) : display(r.TargetIdentity, DisplayEnvironmentSide.Target))), DisplayProjectionOutcome.Completed);
            var projection = ExplorerFixture.Project(f); Assert.AreEqual(4, projection.Rows.Count); Assert.AreEqual(4, projection.Occurrences.Count);
            foreach (var row in projection.Rows)
            {
                bool source = row.Original.AuthoritativeRow.SourceIdentity != null;
                Assert.AreEqual("Same retained title", row.PrimaryName(source ? DisplayEnvironmentSide.Source : DisplayEnvironmentSide.Target, true));
                Assert.IsNull(row.PrimaryName(source ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source, true));
                StringAssert.Contains(row.MembershipPresentation, "Unsupported");
                Assert.AreEqual("", row.Original.AuthoritativeRow.DefinitionStatus, "The fixture does not supply a definition result.");
                Assert.AreEqual("Not compared — result not supplied", row.DefinitionPresentation);
                Assert.AreSame(originals[row.OriginalRowIndex], row.Original.AuthoritativeRow);
                if (type == 300) Assert.AreEqual("Canvas App/library (unclassified)", row.PrimaryType);
            }
            Assert.AreEqual(csv, ExplorerFixture.Csv(f)); CollectionAssert.AreEqual(originals, f.Presentation.Rows.ToArray());
        }
        [TestMethod, TestCategory("D4T8R2C")]
        public void R2CNameHelpersRejectWrongSideAndRetainIndependentOwners()
        {
            var f = R2CNamed(new ComponentDisplayValue(DisplayAvailability.Available, "Source title", "Synthetic title"),
                new ComponentDisplayValue(DisplayAvailability.Available, "logical", "Synthetic name")); var owner = f.Presentation.Rows[0];
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                ExplorerFixture.Display(owner.SourceIdentity, DisplayEnvironmentSide.Target, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Wrong side", "Synthetic wrong owner")),
                ExplorerFixture.Display(owner.TargetIdentity, DisplayEnvironmentSide.Target, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Target title", "Synthetic correct owner"))) }, DisplayProjectionOutcome.Completed);
            var row = ExplorerFixture.Project(f).Rows[0]; Assert.IsNull(row.SourceDisplay); Assert.IsNull(row.TargetDisplay);
            Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Source, true)); Assert.IsNull(row.PrimaryName(DisplayEnvironmentSide.Target, true));
            Assert.AreSame(owner, row.Original.AuthoritativeRow);
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                ExplorerFixture.Display(owner.SourceIdentity, DisplayEnvironmentSide.Source, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Source title", "Synthetic correct owner")),
                ExplorerFixture.Display(owner.TargetIdentity, DisplayEnvironmentSide.Target, friendly: new ComponentDisplayValue(DisplayAvailability.Available, "Target title", "Synthetic correct owner"))) }, DisplayProjectionOutcome.Completed);
            var validRow = ExplorerFixture.Project(f).Rows[0];
            Assert.AreEqual("Source title", validRow.PrimaryName(DisplayEnvironmentSide.Source, true));
            Assert.AreEqual("Target title", validRow.PrimaryName(DisplayEnvironmentSide.Target, true));
            Assert.AreSame(owner, validRow.Original.AuthoritativeRow);
        }
        [DataTestMethod, TestCategory("D4T8R2C")]
        [DataRow("Match"), DataRow("Different"), DataRow("Not Compared"), DataRow("Unsupported"), DataRow("Ambiguous"), DataRow("Unresolved"),
         DataRow("NotRetrieved"), DataRow("NotAttempted"), DataRow("NotApplicable"), DataRow("NotEstablished"), DataRow("Unavailable"), DataRow("RetrievalFailed")]
        public void R2CStatusesKeepMembershipAndDefinitionIndependent(string status)
        {
            var f = R2CNamed(new ComponentDisplayValue(DisplayAvailability.NotRetrieved), new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
            var rows = f.Presentation.Rows.Select(r => r.WithDefinition(status, "", null)).ToArray();
            f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
            f.Completed = new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed);
            var row = ExplorerFixture.Project(f).Rows[0]; Assert.AreEqual("Present in Both", row.MembershipPresentation);
            Assert.AreEqual(status == "Match" ? "Match (compared properties)" : status, row.DefinitionPresentation);
            Assert.AreEqual(status, row.Original.DefinitionStatus); Assert.AreEqual(0, row.RetainedComparedProperties.Count);
            StringAssert.Contains(row.ComparedPropertiesExplanation, "not retained / not compared");
            Assert.IsFalse(row.DefinitionPresentation.Contains("Objects match")); Assert.AreSame(rows[0], row.Original.AuthoritativeRow);
        }
        [TestMethod, TestCategory("D4T8R2C")]
        public void R2CRetainedPropertyAdapterNeverInventsNullOrEmptyEvidence()
        {
            var f = R2CNamed(new ComponentDisplayValue(DisplayAvailability.NotRetrieved), new ComponentDisplayValue(DisplayAvailability.NotRetrieved));
            var properties = new[] { new ComponentPropertyPresentation("unchanged", "", "", false, "Match"),
                new ComponentPropertyPresentation("changed", null, "[NotRetrieved]", true, "Different") };
            var detail = new ComponentDefinitionDetailPresentation("Web resource", "retained", "Present in Both", "Different", properties, "Synthetic context");
            var rows = f.Presentation.Rows.Select(r => r.WithDefinition("Different", "changed", detail)).ToArray();
            f.Presentation = new MembershipComparisonPresentation("SAA", f.Presentation.Source, f.Presentation.Target, rows);
            f.Completed = new CompletedMembershipComparison(f.Presentation, f.Completed.DisplayEvidence, DisplayProjectionOutcome.Completed);
            var row = ExplorerFixture.Project(f).Rows[0]; Assert.AreSame(detail.Properties, row.RetainedComparedProperties);
            Assert.AreSame(properties[0], row.RetainedComparedProperties[0]); Assert.AreSame(properties[1], row.RetainedComparedProperties[1]);
            Assert.AreEqual("[No value supplied; availability not established]", SolutionObjectsExplorerRow.PropertyValue(null));
            Assert.AreEqual("[Blank value supplied]", SolutionObjectsExplorerRow.PropertyValue(""));
            Assert.AreEqual("[NotRetrieved]", SolutionObjectsExplorerRow.PropertyValue("[NotRetrieved]"));
        }
        [DataTestMethod, TestCategory("D4T8R2C"), DataRow(true), DataRow(false)]
        public void R2COneSidedMembershipLabelsDoNotBorrowOppositeNames(bool source)
        {
            var identity = ExplorerFixture.Identity(61, "retained");
            var f = ExplorerFixture.Make(source ? new[] { identity } : new ComponentIdentity[0], source ? new ComponentIdentity[0] : new[] { identity });
            var row = ExplorerFixture.Project(f).Rows[0];
            Assert.AreEqual(source ? "Source only" : "Target only", row.MembershipPresentation);
            StringAssert.Contains(row.NameBasis(source ? DisplayEnvironmentSide.Target : DisplayEnvironmentSide.Source), "No local observation");
            Assert.AreSame(f.Presentation.Rows[0], row.Original.AuthoritativeRow);
        }
        [DataTestMethod, TestCategory("D4T8R2C"), DataRow(IdentityResolutionStatus.Unsupported), DataRow(IdentityResolutionStatus.Ambiguous), DataRow(IdentityResolutionStatus.Unresolved)]
        public void R2CUncertainMembershipLabelsRetainTheOriginalReason(IdentityResolutionStatus status)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(300, null, status: status) }); var row = ExplorerFixture.Project(f).Rows[0];
            StringAssert.Contains(row.MembershipPresentation, "Needs review"); StringAssert.Contains(row.MembershipPresentation, row.Original.MembershipStatus);
            Assert.AreSame(f.Presentation.Rows[0], row.Original.AuthoritativeRow);
        }
        [DataTestMethod, TestCategory("D4T8R2C"), DataRow(29, "Process/workflow (unclassified)"), DataRow(60, "Form/dashboard (unclassified)"), DataRow(300, "Canvas App/library (unclassified)")]
        public void R2CTypeDescriptionsDoNotInventSubtypes(int type, string expected)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(type, null, status: IdentityResolutionStatus.Unsupported) });
            Assert.AreEqual(expected, ExplorerFixture.Project(f).Rows[0].PrimaryType);
        }
        [TestMethod, TestCategory("D4T8R2C")]
        public void R2CAppSettingFallbackUsesOnlyAnEstablishedNameContract()
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(10075, "raw-portable-key"), ExplorerFixture.Identity(10075, ExplorerFixture.Key("app", "OnlineMode")) });
            var p = ExplorerFixture.Project(f);
            var raw = p.Rows.Single(r => r.Original.AuthoritativeRow.SourceIdentity.ComparisonKey == "raw-portable-key");
            Assert.IsNull(raw.PrimaryName(DisplayEnvironmentSide.Source, true));
            var verified = p.Rows.Single(r => r.Original.AuthoritativeRow.SourceIdentity.ComparisonKey != "raw-portable-key");
            Assert.AreEqual("OnlineMode", verified.PrimaryName(DisplayEnvironmentSide.Source, true));
            Assert.IsTrue(p.Occurrences.All(o => o.Children.Count == 0), "A formatted name is not structural parent evidence.");
        }
        private static ExplorerFixture R2CNamed(ComponentDisplayValue friendly, ComponentDisplayValue name)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(61, "resource") }, new[] { ExplorerFixture.Identity(61, "resource") }); var row = f.Presentation.Rows[0];
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                ExplorerFixture.Display(row.SourceIdentity, DisplayEnvironmentSide.Source, friendly: friendly, name: name),
                ExplorerFixture.Display(row.TargetIdentity, DisplayEnvironmentSide.Target, friendly: friendly, name: name)) }, DisplayProjectionOutcome.Completed); return f;
        }


        private static ExplorerFixture Named(ComponentDisplayValue left, ComponentDisplayValue right = null)
        {
            var f = ExplorerFixture.Make(new[] { ExplorerFixture.Identity(1, "incident") }, new[] { ExplorerFixture.Identity(1, "incident") }); var row = f.Presentation.Rows[0];
            f.Completed = new CompletedMembershipComparison(f.Presentation, new[] { new ComponentDisplayEvidence(0,
                ExplorerFixture.Display(row.SourceIdentity, DisplayEnvironmentSide.Source, friendly: left),
                ExplorerFixture.Display(row.TargetIdentity, DisplayEnvironmentSide.Target, friendly: right ?? left)) }, DisplayProjectionOutcome.Completed);
            return f;
        }
    }
}
