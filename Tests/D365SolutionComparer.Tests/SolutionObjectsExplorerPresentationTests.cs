using System;
using System.Linq;
using D365SolutionComparer.Models.Identity;
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
