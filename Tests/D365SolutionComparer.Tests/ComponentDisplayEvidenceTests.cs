using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using D365SolutionComparer.Models.Membership;
using D365SolutionComparer.Services.Membership;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace D365SolutionComparer.Tests
{
    /// <summary>All observations here are synthetic, not captured live SAA evidence.</summary>
    [TestClass]
    public class ComponentDisplayEvidenceTests
    {
        [DataTestMethod]
        [DataRow(1, "Table")]
        [DataRow(2, "Column")]
        [DataRow(9, "Global Choice (subset)")]
        [DataRow(10, "Relationship")]
        [DataRow(26, "Saved Query")]
        [DataRow(29, "Process / Workflow")]
        [DataRow(31, "Report (subset)")]
        [DataRow(36, "Email Template")]
        [DataRow(59, "Chart / Saved Query Visualization")]
        [DataRow(60, "Form")]
        [DataRow(61, "Web Resource")]
        [DataRow(62, "Site Map")]
        [DataRow(80, "Model-Driven App")]
        [DataRow(91, "Plug-in Assembly")]
        [DataRow(92, "SDK Message Processing Step")]
        [DataRow(300, "Canvas App family")]
        [DataRow(380, "Environment Variable Definition")]
        [DataRow(74, "Unknown")]
        [DataRow(511, "Unknown")]
        [DataRow(10072, "Unknown")]
        [DataRow(10075, "Unknown")]
        [DataRow(10121, "Unknown")]
        [DataRow(10266, "Unknown")]
        [DataRow(10267, "Unknown")]
        [DataRow(10276, "Unknown")]
        [DataRow(999999, "Unknown")]
        public void CategoriesDoNotInventLocalRegistrations(int rawType, string category)
        {
            var identity = Identity(rawType, IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new ComponentIdentity[0]);
            var side = ComponentDisplayEvidenceProjector.Project(presentation)[0].SourceDisplayEvidence;
            Assert.AreEqual(category, side.ComponentCategory.Value);
            Assert.AreEqual(rawType, side.RawComponentType);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
        }

        [DataTestMethod]
        [DataRow(74, "MaskingRule", "maskingrule")]
        [DataRow(511, "TeamTemplate", "teamtemplate")]
        [DataRow(10072, "AppElement", "appelement")]
        [DataRow(10075, "AppSetting", "appsetting")]
        [DataRow(10266, "AppAction", "appaction")]
        [DataRow(10267, "AppAction", "appaction")]
        [DataRow(10276, "AISkillConfig", "aiskillconfig")]
        public void RegisteredCategoryDoesNotImplyProductionSupport(int rawType, string name, string table)
        {
            var identity = new ComponentIdentity(Record(rawType), IdentityResolutionStatus.Unsupported,
                registeredDefinition: new SolutionComponentDefinitionIdentity(rawType, name, table));
            var side = ComponentDisplayEvidenceProjector.Project(Present(new[] { identity }, new ComponentIdentity[0]))[0].SourceDisplayEvidence;
            Assert.AreEqual(name, side.ComponentCategory.Value);
            Assert.AreEqual("SolutionComponentDefinition.Name", side.ComponentCategory.DisplayEvidenceSource);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, side.FriendlyName.Availability);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status);
        }

        [TestMethod]
        public void VerifiedConnectionReferenceSemanticKindIsIndependentOfRawCode()
        {
            var item = new ComponentIdentity(Record(10121), IdentityResolutionStatus.Resolved, "synthetic-reference",
                semanticKind: ComponentSemanticKinds.ConnectionReference);
            Assert.AreEqual("Connection Reference", ComponentDisplayEvidenceProjector.Project(
                Present(new[] { item }, new ComponentIdentity[0]))[0].SourceDisplayEvidence.ComponentCategory.Value);
        }

        [DataTestMethod]
        [DataRow(true, true)]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(false, false)]
        public void FriendlyAndInternalFieldsAreIndependent(bool friendly, bool internalName)
        {
            var item = Identity(80);
            var row = new Entity("appmodule", item.Record.ObjectId.Value);
            if (friendly) row["name"] = "Synthetic readable name";
            if (internalName) row["uniquename"] = "synthetic_internal";
            var side = Project(item, Cache(row));
            Assert.AreEqual(friendly ? DisplayAvailability.Available : DisplayAvailability.AttributeOmitted, side.FriendlyName.Availability);
            Assert.AreEqual(internalName ? DisplayAvailability.Available : DisplayAvailability.AttributeOmitted, side.InternalName.Availability);
            Assert.AreEqual(friendly ? "Synthetic readable name" : null, side.FriendlyName.Value);
            Assert.AreEqual(internalName ? "synthetic_internal" : null, side.InternalName.Value);
        }

        [TestMethod]
        public void BothSidesKeepDifferentNamesWithoutChangingMatch()
        {
            var source = Identity(80);
            var target = Identity(80);
            var presentation = Present(new[] { source }, new[] { target });
            var sourceRow = App(source, "Source display", "source_internal");
            var targetRow = App(target, "Target display", "target_internal");
            var evidence = ComponentDisplayEvidenceProjector.Project(presentation, Cache(sourceRow), Cache(targetRow))[0];
            Assert.AreEqual("Source display", evidence.SourceDisplayEvidence.FriendlyName.Value);
            Assert.AreEqual("Target display", evidence.TargetDisplayEvidence.FriendlyName.Value);
            Assert.AreEqual("source_internal", evidence.SourceDisplayEvidence.InternalName.Value);
            Assert.AreEqual("target_internal", evidence.TargetDisplayEvidence.InternalName.Value);
            Assert.AreEqual("Present in Both", presentation.Rows[0].MembershipStatus);
            Assert.AreEqual(DisplayEnvironmentSide.Source, evidence.SourceDisplayEvidence.EnvironmentSide);
            Assert.AreEqual(DisplayEnvironmentSide.Target, evidence.TargetDisplayEvidence.EnvironmentSide);
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void OneSidedLocalObservationDoesNotManufactureOtherSide(bool source)
        {
            var identity = Identity(80);
            var empty = new ComponentIdentity[0];
            var present = new[] { identity };
            var presentation = Present(source ? present : empty, source ? empty : present);
            var evidence = ComponentDisplayEvidenceProjector.Project(presentation)[0];
            var observed = source ? evidence.SourceDisplayEvidence : evidence.TargetDisplayEvidence;
            var unobserved = source ? evidence.TargetDisplayEvidence : evidence.SourceDisplayEvidence;
            Assert.IsTrue(observed.HasLocalObservation);
            Assert.IsFalse(unobserved.HasLocalObservation);
            Assert.AreEqual(DisplayAvailability.NotApplicable, unobserved.FriendlyName.Availability);
            Assert.AreEqual(1, source ? presentation.Summary.SourceOnly : presentation.Summary.TargetOnly);
        }

        [TestMethod]
        public void UnsupportedWithCachedDisplayStillCannotPair()
        {
            var source = Identity(300, IdentityResolutionStatus.Unsupported);
            var target = Identity(300, IdentityResolutionStatus.Unsupported);
            var sourceCache = Cache(new Entity("canvasapp", source.Record.ObjectId.Value) { ["name"] = "same" });
            var presentation = Present(new[] { source }, new[] { target });
            var display = ComponentDisplayEvidenceProjector.Project(presentation, sourceCache, null);
            Assert.AreEqual(2, display.Count);
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(2, presentation.Summary.IndeterminateUnsupported);
            Assert.IsTrue(display.All(row => row.SourceDisplayEvidence.FriendlyName.Availability != DisplayAvailability.Available));
            Assert.IsTrue(display.All(row => (row.SourceDisplayEvidence.HasLocalObservation ? row.SourceDisplayEvidence : row.TargetDisplayEvidence)
                .ComponentCategory.Value == "Canvas App family"));
        }

        [TestMethod]
        public void UnsupportedWithAvailableMetadataRetainsLocalDisplayOnly()
        {
            var item = Identity(1, IdentityResolutionStatus.Unsupported);
            var cache = new DataverseComponentMetadataCache();
            cache.Store(new EntityMetadata { MetadataId = item.Record.ObjectId, LogicalName = "synthetic_table",
                DisplayName = new Label { UserLocalizedLabel = new LocalizedLabel("Readable table", 1033) } });
            var side = Project(item, cache);
            Assert.AreEqual("Readable table", side.FriendlyName.Value);
            Assert.AreEqual("synthetic_table", side.InternalName.Value);
            Assert.AreEqual("EntityMetadata.DisplayName", side.FriendlyName.DisplayEvidenceSource);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, item.Status);
        }

        [DataTestMethod]
        [DataRow(DisplayAvailability.Unavailable)]
        [DataRow(DisplayAvailability.NotRetrieved)]
        [DataRow(DisplayAvailability.AttributeOmitted)]
        [DataRow(DisplayAvailability.Ambiguous)]
        [DataRow(DisplayAvailability.RetrievalFailed)]
        [DataRow(DisplayAvailability.NotApplicable)]
        public void NonValuesRemainDistinctAndCannotCarryFabricatedNames(DisplayAvailability availability)
        {
            var value = new ComponentDisplayValue(availability);
            Assert.AreEqual(availability, value.Availability);
            Assert.IsNull(value.Value);
            Assert.ThrowsException<ArgumentException>(() => new ComponentDisplayValue(availability, "fake"));
        }

        [TestMethod]
        public void NullEmptyAndOmittedAreDifferentObservations()
        {
            var item = Identity(80);
            var row = new Entity("appmodule", item.Record.ObjectId.Value) { ["name"] = null, ["uniquename"] = "" };
            var side = Project(item, Cache(row));
            Assert.AreEqual(DisplayAvailability.Available, side.FriendlyName.Availability);
            Assert.IsNull(side.FriendlyName.Value);
            Assert.AreEqual(DisplayAvailability.Available, side.InternalName.Availability);
            Assert.AreEqual("", side.InternalName.Value);
            row.Attributes.Remove("name");
            Assert.AreEqual(DisplayAvailability.AttributeOmitted, Project(item, Cache(row)).FriendlyName.Availability);
        }

        [TestMethod]
        public void AvailableValueRequiresProvenance()
        {
            Assert.ThrowsException<ArgumentException>(() => new ComponentDisplayValue(DisplayAvailability.Available, "synthetic"));
        }

        [TestMethod]
        public void MissingCacheIsNotMissingBackingRecordOrSuccessfulEmptyRead()
        {
            var side = Project(Identity(80), new DataverseComponentMetadataCache());
            Assert.AreEqual(DisplayAvailability.NotRetrieved, side.FriendlyName.Availability);
            Assert.AreEqual(DisplayBackingCorrelationStatus.Unavailable, side.BackingCorrelationStatus);
        }

        [DataTestMethod]
        [DataRow("Missing", DisplayBackingCorrelationStatus.Missing)]
        [DataRow("DuplicateReturnedRow", DisplayBackingCorrelationStatus.Ambiguous)]
        [DataRow("ConflictingPrimaryKey", DisplayBackingCorrelationStatus.Ambiguous)]
        [DataRow("IncompleteBatch", DisplayBackingCorrelationStatus.Unavailable)]
        public void CompletedStepCorrelationsStayConservative(string correlation, DisplayBackingCorrelationStatus expected)
        {
            var item = Identity(92, IdentityResolutionStatus.Unresolved);
            var cache = new DataverseComponentMetadataCache();
            cache.Type92Evidence = new Type92EnvironmentEvidence(null, "", new SolutionComponentRecord[0],
                new[] { new Type92StepEvidence(item.Record.ObjectId.Value, 1, correlation, "Synthetic diagnostic") }, new Dictionary<string, int>());
            var side = Project(item, cache);
            Assert.AreEqual(expected, side.BackingCorrelationStatus);
            Assert.AreEqual(IdentityResolutionStatus.Unresolved, item.Status);
        }

        [TestMethod]
        public void DuplicateFriendlyNamesCannotAddMatchesOrRepairIdentity()
        {
            var a = Identity(80, key: "A");
            var b = Identity(80, key: "B");
            var target = Identity(80, key: "C");
            var sourceCache = Cache(App(a, "Same name", "same"), App(b, "Same name", "same"));
            var presentation = Present(new[] { a, b }, new[] { target });
            ComponentDisplayEvidenceProjector.Project(presentation, sourceCache, Cache(App(target, "Same name", "same")));
            Assert.AreEqual(0, presentation.Summary.PresentInBoth);
            Assert.AreEqual(2, presentation.Summary.SourceOnly);
            Assert.AreEqual(1, presentation.Summary.TargetOnly);
        }

        [TestMethod]
        public void OpaqueKeysAndDiagnosticTextAreNeverDisplayFallbacks()
        {
            var item = new ComponentIdentity(Record(80), IdentityResolutionStatus.Resolved,
                "opaque:v1:some_friendly_name", diagnosticEvidence: new[] { "name='unverified text'" });
            var side = Project(item, null);
            Assert.IsNull(side.InternalName.Value);
            Assert.IsNull(side.FriendlyName.Value);
        }

        [TestMethod]
        public void AppSettingSupportSummaryAndCsvRemainUnchanged()
        {
            var definition = new SolutionComponentDefinitionIdentity(10075, "AppSetting", "appsetting");
            var source = new ComponentIdentity(Record(10075), IdentityResolutionStatus.Resolved, "synthetic-appsetting",
                registeredDefinition: definition);
            var target = new ComponentIdentity(Record(10075), IdentityResolutionStatus.Resolved, "synthetic-appsetting",
                registeredDefinition: definition);
            var presentation = Present(new[] { source, Identity(36, IdentityResolutionStatus.Unsupported) }, new[] { target });
            var rows = presentation.Rows.ToArray();
            var summary = presentation.Summary;
            var csv = new MembershipCoverageCsvExporter().CreateCsv(presentation);
            var display = ComponentDisplayEvidenceProjector.Project(presentation);
            Assert.AreEqual(csv, new MembershipCoverageCsvExporter().CreateCsv(presentation));
            Assert.AreSame(summary, presentation.Summary);
            CollectionAssert.AreEqual(rows, presentation.Rows.ToArray());
            Assert.AreEqual(1, summary.PresentInBoth);
            Assert.AreEqual(1, summary.IndeterminateUnsupported);
            Assert.AreEqual(2, summary.TotalComparisonRows);
            Assert.IsTrue(display.Any(row => row.SourceDisplayEvidence.ComponentCategory.Value == "AppSetting"));
        }

        [TestMethod]
        public void DisplayIsDetachedFromMutableSdkRowsAndDoesNotRetainPayloads()
        {
            var item = Identity(61);
            var row = new Entity("webresource", item.Record.ObjectId.Value) { ["name"] = "synthetic.js", ["displayname"] = "Friendly", ["content"] = new string('X', 1000000) };
            var cache = Cache(row);
            var result = Project(item, cache);
            row["displayname"] = "Changed after projection";
            Assert.AreEqual("Friendly", result.FriendlyName.Value);
            Assert.AreEqual("webresource.displayname", result.FriendlyName.DisplayEvidenceSource);
            Assert.IsFalse(typeof(ComponentSideDisplayEvidence).GetProperties().Any(property =>
                property.PropertyType == typeof(Entity) || property.PropertyType == typeof(ComponentIdentity) || property.CanWrite));
            Assert.IsFalse(typeof(ComponentDisplayEvidence).GetProperties().Any(property => property.CanWrite));
            Assert.IsFalse(typeof(ComponentDisplayValue).GetProperties().Any(property => property.CanWrite));
        }

        [TestMethod]
        public void VerifiedColumnHasParentContextWithoutReadingPortableKey()
        {
            var item = Identity(2);
            var cache = new DataverseComponentMetadataCache();
            var column = new StringAttributeMetadata { MetadataId = item.Record.ObjectId, LogicalName = "synthetic_column",
                DisplayName = new Label("Column label", 1033) };
            typeof(AttributeMetadata).GetProperty("EntityLogicalName").SetValue(column, "synthetic_table");
            cache.Store(column);
            var side = Project(item, cache);
            Assert.AreEqual("synthetic_table", side.ParentDisplayContext.Value);
            Assert.AreEqual("AttributeMetadata.EntityLogicalName", side.ParentDisplayContext.DisplayEvidenceSource);
            Assert.AreEqual("synthetic_column", side.InternalName.Value);
        }

        [TestMethod]
        public void WorkflowParentDefinitionNamesHaveExplicitCorrelationProvenance()
        {
            var item = Identity(29);
            var cache = new DataverseComponentMetadataCache();
            cache.WorkflowDefinitions[item.Record.ObjectId.Value] = new Entity("workflow", Guid.NewGuid())
                { ["name"] = "Synthetic definition", ["uniquename"] = "synthetic_definition" };
            Assert.AreEqual(DisplayBackingCorrelationStatus.VerifiedDefinition, Project(item, cache).BackingCorrelationStatus);
        }

        [TestMethod]
        public void GroupedParentMetadataProjectionDoesNotInvokeOrMutateResolverCorrelation()
        {
            var item = Identity(2);
            var attribute = new StringAttributeMetadata { MetadataId = item.Record.ObjectId, LogicalName = "synthetic_column" };
            var table = Parent(new[] { attribute });
            var cache = new DataverseComponentMetadataCache { ParentMetadata = new ParentEntityMetadataInventory() };
            cache.ParentMetadata.Load(new[] { table });
            cache.Store(table, true);
            var side = Project(item, cache);
            Assert.AreEqual("synthetic_column", side.InternalName.Value);
            Assert.AreEqual("synthetic_parent", side.ParentDisplayContext.Value);
            Assert.AreEqual("VerifiedParentMetadata.LogicalName", side.ParentDisplayContext.DisplayEvidenceSource);
            Assert.IsNull(attribute.EntityLogicalName);
            Assert.AreEqual("synthetic-key", item.ComparisonKey);
        }

        [TestMethod]
        public void DuplicatePublishedColumnMetadataIsDisplayAmbiguityOnly()
        {
            var item = Identity(2);
            var table = Parent(new AttributeMetadata[] { new StringAttributeMetadata { MetadataId = item.Record.ObjectId, LogicalName = "a" },
                new StringAttributeMetadata { MetadataId = item.Record.ObjectId, LogicalName = "b" } });
            var cache = new DataverseComponentMetadataCache { ParentMetadata = new ParentEntityMetadataInventory() };
            cache.Store(table, true);
            var side = Project(item, cache);
            Assert.AreEqual(DisplayAvailability.Ambiguous, side.InternalName.Availability);
            Assert.AreEqual(IdentityResolutionStatus.Resolved, item.Status);
        }

        [TestMethod]
        public void ExplicitUniqueStepEvidenceSuppliesOnlyStepDisplayName()
        {
            var item = Identity(92);
            var cache = new DataverseComponentMetadataCache();
            cache.Type92Evidence = new Type92EnvironmentEvidence(null, "", new SolutionComponentRecord[0],
                new[] { new Type92StepEvidence(item.Record.ObjectId.Value, 1, "Unique", "") { Name = "Synthetic step", ConfigurationLength = 900000 } },
                new Dictionary<string, int>());
            var side = Project(item, cache);
            Assert.AreEqual("Synthetic step", side.FriendlyName.Value);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, side.InternalName.Availability);
            Assert.AreEqual(DisplayBackingCorrelationStatus.Unique, side.BackingCorrelationStatus);
        }

        [TestMethod]
        public void UnavailableOppositeSnapshotDoesNotBecomeNotApplicableOrAbsence()
        {
            var solution = MembershipTestData.Solution();
            var source = MembershipEnvironmentResult.FromSnapshot("Synthetic Source", MembershipSnapshot.Complete(solution,
                new[] { Identity(80) }, DateTimeOffset.UtcNow), 0, TimeSpan.Zero);
            var target = MembershipEnvironmentResult.Unavailable("Synthetic Target", solution.UniqueName, 0, TimeSpan.Zero, "Synthetic fault");
            var presentation = new MembershipResultPresenter().Create(source, target);
            var result = ComponentDisplayEvidenceProjector.Project(presentation);
            Assert.AreEqual(DisplayAvailability.Unavailable, result[0].TargetDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual(0, presentation.Summary.SourceOnly);
            Assert.AreEqual(1, presentation.Summary.IndeterminateResolved);
        }

        [TestMethod]
        public void MalformedDisplayFieldIsNonfatal()
        {
            var item = Identity(80);
            var row = App(item, "Readable", "internal");
            row["name"] = new OptionSetValue(1);
            var presentation = Present(new[] { item }, new ComponentIdentity[0]);
            var display = ComponentDisplayEvidenceProjector.Project(presentation, Cache(row), null);
            Assert.AreEqual(DisplayAvailability.Unavailable, display[0].SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual(1, presentation.Summary.SourceOnly);
        }

        [TestMethod]
        public void CancellationDiscardsOnlyOptionalProjection()
        {
            var presentation = Present(new[] { Identity(80) }, new ComponentIdentity[0]);
            var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.AreEqual(0, ComponentDisplayEvidenceProjector.Project(presentation, cancellation.Token).Count);
            Assert.AreEqual(1, presentation.Summary.SourceOnly);
        }

        [TestMethod]
        public void EmptyProjectionIsEmpty()
        {
            Assert.AreEqual(0, ComponentDisplayEvidenceProjector.Project(Present(new ComponentIdentity[0], new ComponentIdentity[0])).Count);
        }

        [TestMethod]
        public void LargeSyntheticProjectionIsDeterministicAndKeepsEveryAuthoritativeRow()
        {
            var source = Enumerable.Range(0, 792).Select(i => Identity(80, key: "synthetic-source-" + i)).ToArray();
            var target = Enumerable.Range(0, 720).Select(i => Identity(80, key: "synthetic-target-" + i)).ToArray();
            var presentation = Present(source, target);
            var timer = Stopwatch.StartNew();
            var first = ComponentDisplayEvidenceProjector.Project(presentation);
            timer.Stop();
            var second = ComponentDisplayEvidenceProjector.Project(presentation);
            Assert.AreEqual(1512, first.Count); // Synthetic disjoint identities, not the live 840-row capture.
            CollectionAssert.AreEqual(first.Select(row => row.SourceDisplayEvidence.SolutionComponentId).ToArray(),
                second.Select(row => row.SourceDisplayEvidence.SolutionComponentId).ToArray());
            CollectionAssert.AreEqual(Enumerable.Range(0, 1512).ToArray(), first.Select(row => row.RowIndex).ToArray());
            Assert.IsTrue(timer.Elapsed < TimeSpan.FromSeconds(5), "Display-only processing exceeded the generous synthetic regression bound.");
            TestContext.WriteLine("Synthetic 1512-row projection elapsed ms: " + timer.Elapsed.TotalMilliseconds);
        }

        [TestMethod]
        public void ProjectorHasNoServiceRequestOrIdentityMutationBoundary()
        {
            Assert.IsFalse(typeof(ComponentDisplayEvidenceProjector).GetMethods().SelectMany(method => method.GetParameters())
                .Any(parameter => typeof(IOrganizationService).IsAssignableFrom(parameter.ParameterType)));
            var fake = new FakeOrganizationService();
            var presentation = Present(new[] { Identity(80) }, new ComponentIdentity[0]);
            ComponentDisplayEvidenceProjector.Project(presentation);
            Assert.AreEqual(0, fake.Calls);
            Assert.AreEqual(0, fake.ExecuteCalls);
            Assert.AreEqual(0, presentation.Source.Diagnostics.RequestCount);
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(true), DataRow(false)]
        public void TypedBooleanRequiresObservationAndPreservesFalse(bool value)
        {
            var observed = new ComponentObjectMetadataValue<bool?>(DisplayAvailability.Available, value, "synthetic.ismanaged", true);
            Assert.AreEqual(value, observed.Value);
            Assert.IsTrue(observed.WasRequested);
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(DisplayAvailability.Available, value));
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(DisplayAvailability.Available, value, "synthetic", false));
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(DisplayAvailability.NotRetrieved), DataRow(DisplayAvailability.Unavailable)]
        [DataRow(DisplayAvailability.AttributeOmitted), DataRow(DisplayAvailability.RetrievalFailed)]
        [DataRow(DisplayAvailability.Ambiguous), DataRow(DisplayAvailability.NotApplicable)]
        public void ObjectAvailabilityDoesNotSupplyFalseOrFabricatedValues(DisplayAvailability state)
        {
            bool requested = state == DisplayAvailability.AttributeOmitted || state == DisplayAvailability.RetrievalFailed;
            var observation = new ComponentObjectMetadataValue<bool?>(state, provenance: "synthetic.applicability", wasRequested: requested);
            Assert.IsNull(observation.Value);
            Assert.AreEqual(state, observation.Availability);
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(state, false, "synthetic", requested));
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void AvailableNullIsObservedButNotOmittedOrNotRetrieved()
        {
            var observed = new ComponentObjectMetadataValue<bool?>(DisplayAvailability.Available, null, "synthetic.null", true);
            Assert.IsNull(observed.Value);
            Assert.IsTrue(observed.WasRequested);
            Assert.AreEqual(DisplayAvailability.Available, observed.Availability);
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(DisplayAvailability.AttributeOmitted));
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(DisplayAvailability.NotRetrieved, wasRequested: true));
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectMetadataValue<bool?>(DisplayAvailability.NotApplicable));
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(1), DataRow(2), DataRow(29), DataRow(60), DataRow(80), DataRow(10075), DataRow(300)]
        public void B1DoesNotEnrichObjectMetadataFromPopulatedCaches(int type)
        {
            var identity = Identity(type);
            var cache = new DataverseComponentMetadataCache();
            var row = new Entity(type == 29 ? "workflow" : type == 60 ? "systemform" : "appmodule", identity.Record.ObjectId.Value)
            {
                ["name"] = "Synthetic name", ["uniquename"] = "synthetic_unique", ["ismanaged"] = true,
                ["iscustomizable"] = new BooleanManagedProperty(true), ["componentstate"] = new OptionSetValue(1),
                ["statecode"] = new OptionSetValue(1), ["modifiedon"] = DateTime.UtcNow,
                ["ownerid"] = new EntityReference("systemuser", Guid.NewGuid()) { Name = "Do not enrich" }
            };
            cache.StoreEntityRow(row, row.Id);
            if (type == 29) cache.WorkflowDefinitions[row.Id] = row;
            if (type == 1)
            {
                var metadata = new EntityMetadata { MetadataId = row.Id, LogicalName = "synthetic_table" };
                typeof(EntityMetadata).GetProperty("IsManaged").SetValue(metadata, true);
                cache.Store(metadata);
            }
            if (type == 2)
            {
                var column = new StringAttributeMetadata { MetadataId = row.Id, LogicalName = "synthetic_column" };
                typeof(AttributeMetadata).GetProperty("EntityLogicalName").SetValue(column, "synthetic_parent");
                typeof(AttributeMetadata).GetProperty("IsManaged").SetValue(column, false);
                cache.Store(column);
                var parent = Parent(new AttributeMetadata[] { column });
                typeof(EntityMetadata).GetProperty("IsManaged").SetValue(parent, true);
                cache.Store(parent);
            }
            var side = Project(identity, cache);
            AssertObjectNotRetrieved(side.ObjectMetadata);
            Assert.AreEqual(DisplayEnvironmentSide.Source, side.ObjectMetadata.EnvironmentSide);
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void LegacySideDefaultsAreDetachedAndEnvironmentSpecific()
        {
            var completed = ComponentDisplayEvidenceProjector.Project(Present(new[] { Identity(80) }, new[] { Identity(80) }));
            foreach (var item in completed)
            {
                AssertObjectNotRetrieved(item.SourceDisplayEvidence.ObjectMetadata);
                AssertObjectNotRetrieved(item.TargetDisplayEvidence.ObjectMetadata);
                Assert.AreEqual(DisplayEnvironmentSide.Source, item.SourceDisplayEvidence.ObjectMetadata.EnvironmentSide);
                Assert.AreEqual(DisplayEnvironmentSide.Target, item.TargetDisplayEvidence.ObjectMetadata.EnvironmentSide);
            }
            var value = new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
            Assert.ThrowsException<ArgumentException>(() => new ComponentSideDisplayEvidence(DisplayEnvironmentSide.Source,
                Guid.NewGuid(), Guid.NewGuid(), 80, value, value, value, DisplayBackingCorrelationStatus.Unique, value,
                ComponentObjectMetadataEvidence.NotRetrieved(DisplayEnvironmentSide.Target)));
            foreach (var type in new[] { typeof(ComponentObjectMetadataEvidence), typeof(ComponentTimestampEvidence),
                typeof(ComponentOwnerEvidence), typeof(ComponentObjectStatusEvidence), typeof(ComponentObjectMetadataValue<bool?>) })
                Assert.IsFalse(type.GetProperties().Any(p => p.CanWrite || typeof(IOrganizationService).IsAssignableFrom(p.PropertyType) ||
                    p.PropertyType == typeof(Entity) || p.PropertyType == typeof(EntityReference)));
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow(DateTimeKind.Unspecified), DataRow(DateTimeKind.Local), DataRow(DateTimeKind.Utc)]
        public void TimestampRetainsKindWithoutInventingTimezone(DateTimeKind kind)
        {
            var raw = new DateTime(2025, 11, 2, 1, 30, 0, kind);
            var evidence = new ComponentTimestampEvidence(raw, "Synthetic behavior");
            Assert.AreEqual(raw, evidence.RawTimestamp);
            Assert.AreEqual(kind, evidence.Kind);
            Assert.IsNull(evidence.VerifiedInstant);
            Assert.IsNull(evidence.TimezoneProvenance);
            Assert.ThrowsException<ArgumentException>(() => new ComponentTimestampEvidence(raw, verifiedInstant: new DateTimeOffset(raw.Ticks, TimeSpan.Zero)));
        }

        [TestMethod, TestCategory("P2AObjectMetadata")]
        public void TimestampIndependentProofAndOwnerAuditRemainTyped()
        {
            var raw = new DateTime(2025, 11, 2, 1, 30, 0, DateTimeKind.Unspecified);
            var instant = new DateTimeOffset(raw, TimeSpan.FromHours(-5));
            var evidence = new ComponentTimestampEvidence(raw, "UserLocal", "Synthetic independent offset proof", instant);
            Assert.AreEqual(instant, evidence.VerifiedInstant);
            var id = Guid.NewGuid();
            var owner = new ComponentOwnerEvidence("team", id);
            Assert.AreEqual(id, owner.LocalOwnerId);
            Assert.IsNull(owner.DisplayName);
            Assert.IsFalse(owner.DisplayNameVerified);
            Assert.ThrowsException<ArgumentException>(() => new ComponentOwnerEvidence("team", id, "Unverified label"));
        }

        [DataTestMethod, TestCategory("P2AObjectMetadata")]
        [DataRow("componentstate"), DataRow("MembershipStatus"), DataRow("DefinitionStatus")]
        public void ComparisonAndLayerStatusesCannotSupplyObjectActivation(string field)
        {
            Assert.ThrowsException<ArgumentException>(() => new ComponentObjectStatusEvidence(field, 1));
        }

        internal static void AssertObjectNotRetrieved(ComponentObjectMetadataEvidence evidence)
        {
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.Managed.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.Customized.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.LastModified.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.Owner.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.ObjectStatus.Availability);
            Assert.IsFalse(evidence.Managed.WasRequested || evidence.Customized.WasRequested || evidence.LastModified.WasRequested ||
                evidence.Owner.WasRequested || evidence.ObjectStatus.WasRequested);
            Assert.IsNull(evidence.Managed.Value); Assert.IsNull(evidence.Customized.Value); Assert.IsNull(evidence.LastModified.Value);
            Assert.IsNull(evidence.Owner.Value); Assert.IsNull(evidence.ObjectStatus.Value);
        }

        [DataTestMethod, TestCategory("H2E2C3FA1")]
        [DataRow(31, "Nonblank"), DataRow(36, "Nonblank")]
        [DataRow(31, "Omitted"), DataRow(36, "Omitted")]
        [DataRow(31, "Null"), DataRow(36, "Null")]
        [DataRow(31, "Empty"), DataRow(36, "Empty")]
        [DataRow(31, "Blank"), DataRow(36, "Blank")]
        [DataRow(31, "Malformed"), DataRow(36, "Malformed")]
        [DataRow(31, "Missing"), DataRow(36, "Missing")]
        [DataRow(31, "Duplicate"), DataRow(36, "Duplicate")]
        [DataRow(31, "Conflict"), DataRow(36, "Conflict")]
        [DataRow(31, "FailedOrIncomplete"), DataRow(36, "FailedOrIncomplete")]
        public void A1TypedNormalNamesPreserveFieldAvailabilityWithoutIdentity(int rawType, string state)
        {
            var identity = Identity(rawType, IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new ComponentIdentity[0]);
            string field = rawType == 31 ? "name" : "title", primary = rawType == 31 ? "reportid" : "templateid";
            var row = A1NameRow(rawType, identity.Record.ObjectId.Value, "Synthetic friendly");
            if (state == "Omitted") row.Attributes.Remove(field);
            if (state == "Null") row[field] = null;
            if (state == "Empty") row[field] = string.Empty;
            if (state == "Blank") row[field] = "  ";
            if (state == "Malformed") row[field] = new OptionSetValue(1);
            if (state == "Conflict") row[primary] = Guid.NewGuid();
            var returned = state == "Missing" ? new Entity[0] : state == "Duplicate" ? new[] { row, row } : new[] { row };
            var cache = A1NameCache(presentation.Source.Snapshot, rawType, returned, state == "FailedOrIncomplete");
            var evidence = ComponentDisplayEvidenceProjector.Project(presentation, cache, null)[0].SourceDisplayEvidence;
            var expected = state == "Omitted" ? DisplayAvailability.AttributeOmitted :
                state == "Duplicate" || state == "Conflict" ? DisplayAvailability.Ambiguous :
                state == "Malformed" || state == "Missing" || state == "FailedOrIncomplete" ? DisplayAvailability.Unavailable : DisplayAvailability.Available;
            Assert.AreEqual(expected, evidence.FriendlyName.Availability);
            Assert.AreEqual(expected == DisplayAvailability.Available ? (string)row[field] : null, evidence.FriendlyName.Value);
            Assert.AreEqual("NormalMembershipDiagnostic." + (rawType == 31 ? "report.name" : "template.title"),
                evidence.FriendlyName.DisplayEvidenceSource);
            Assert.AreEqual(state == "Missing" ? DisplayBackingCorrelationStatus.Missing :
                state == "Duplicate" || state == "Conflict" ? DisplayBackingCorrelationStatus.Ambiguous :
                state == "FailedOrIncomplete" ? DisplayBackingCorrelationStatus.Unavailable : DisplayBackingCorrelationStatus.Unique,
                evidence.BackingCorrelationStatus);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.InternalName.Availability);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(MembershipPresence.Indeterminate, presentation.Rows.Single().Presence);
            Assert.AreEqual(1, presentation.Summary.IndeterminateUnsupported);
        }

        [DataTestMethod, TestCategory("H2E2C3FA1"), DataRow(31), DataRow(36)]
        public void A1NamesRejectCopiedSnapshotsRecordsAndOppositeSideEvenWithSameLocalIds(int rawType)
        {
            var identity = Identity(rawType, IdentityResolutionStatus.Unsupported);
            var target = new ComponentIdentity(new SolutionComponentRecord(identity.Record.SolutionComponentId,
                rawType, identity.Record.ObjectId), IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new[] { target });
            var cache = A1NameCache(presentation.Source.Snapshot, rawType,
                new[] { A1NameRow(rawType, identity.Record.ObjectId.Value, "Source only label") });
            var correct = ComponentDisplayEvidenceProjector.Project(presentation, cache, cache);
            Assert.AreEqual("Source only label", correct.Single(x => x.SourceDisplayEvidence.HasLocalObservation).SourceDisplayEvidence.FriendlyName.Value);
            Assert.IsTrue(correct.All(x => x.TargetDisplayEvidence.FriendlyName.Availability != DisplayAvailability.Available));
            var copied = MembershipSnapshot.Complete(presentation.Source.Snapshot.Solution,
                presentation.Source.Snapshot.Components, presentation.Source.Snapshot.CapturedAt);
            DataverseComponentMetadataCache.NormalDisplayNameObservation observation;
            Assert.IsFalse(cache.TryGetNormalDisplayName(copied, identity, out observation));
            Assert.IsFalse(cache.TryGetNormalDisplayName(presentation.Source.Snapshot, target, out observation));
            Assert.AreEqual(2, presentation.Rows.Count); Assert.AreEqual(0, presentation.Summary.PresentInBoth);
        }

        [DataTestMethod, TestCategory("H2E2C3FA1"), DataRow(31), DataRow(36)]
        public void A1DiagnosticOrGenericEntityRowsCannotSupplyNormalFriendlyNames(int rawType)
        {
            var identity = new ComponentIdentity(Record(rawType), IdentityResolutionStatus.Unsupported,
                diagnosticEvidence: new[] { "name='Synthetic unverified diagnostic'; title='Synthetic unverified title'" });
            var cache = Cache(A1NameRow(rawType, identity.Record.ObjectId.Value, "Not a normal observation"));
            var evidence = Project(identity, cache);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.FriendlyName.Availability);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, evidence.InternalName.Availability);
            Assert.IsNull(evidence.FriendlyName.Value); Assert.IsNull(identity.ComparisonKey);
        }

        [TestMethod, TestCategory("H2E2C3FA1")]
        public void A1NormalNameCopiesOnlyImmutableScalarAndCannotRebindReadOwner()
        {
            var identity = Identity(31, IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new ComponentIdentity[0]);
            var row = A1NameRow(31, identity.Record.ObjectId.Value, "Original synthetic name");
            var cache = A1NameCache(presentation.Source.Snapshot, 31, new[] { row });
            row["name"] = "Changed SDK payload";
            var retained = ComponentDisplayEvidenceProjector.Project(presentation, cache, null)[0].SourceDisplayEvidence;
            Assert.AreEqual("Original synthetic name", retained.FriendlyName.Value);
            cache.BeginNormalDisplayNames(presentation.Source.Snapshot);
            Assert.AreEqual(DisplayAvailability.NotRetrieved,
                ComponentDisplayEvidenceProjector.Project(presentation, cache, null)[0].SourceDisplayEvidence.FriendlyName.Availability);
            Assert.AreEqual("Original synthetic name", retained.FriendlyName.Value);
        }

        internal static Entity A1NameRow(int rawType, Guid id, string name, bool managed = false, Guid? signature = null)
        {
            var row = new Entity(rawType == 31 ? "report" : "template", id)
            { [rawType == 31 ? "reportid" : "templateid"] = id,
              [rawType == 31 ? "name" : "title"] = name,
              [rawType == 31 ? "reportidunique" : "templateidunique"] = Guid.NewGuid(),
              ["componentstate"] = new OptionSetValue(0), ["ismanaged"] = managed };
            if (rawType == 31)
            {
                row["filename"] = "synthetic.rdl"; row["reporttypecode"] = new OptionSetValue(1);
                if (signature.HasValue) { row["signatureid"] = signature.Value; row["signaturelcid"] = 1033; }
            }
            else { row["templatetypecode"] = "contact"; row["languagecode"] = 1033; row["ispersonal"] = false; }
            return row;
        }

        internal static DataverseComponentMetadataCache A1NameCache(MembershipSnapshot owner, int rawType,
            Entity[] returned, bool failedOrIncomplete = false)
        {
            var ids = owner.Components.Where(i => i.Record.ComponentType == rawType && i.Record.ObjectId.HasValue &&
                i.Record.ObjectId.Value != Guid.Empty).Select(i => i.Record.ObjectId.Value).Distinct().ToList();
            var rows = ids.ToDictionary(id => id, id => returned.Where(row => row.Id == id).ToList());
            var failures = new Dictionary<Guid, string>();
            if (failedOrIncomplete) foreach (var id in ids) failures.Add(id, "Synthetic failure/incomplete; not parsed");
            var retrieval = new DiagnosticQueryResult(ids, rows, failures,
                ids.ToDictionary(id => id, id => new List<Entity>()), returned.Length, !failedOrIncomplete);
            var cache = new DataverseComponentMetadataCache(); cache.BeginNormalDisplayNames(owner);
            cache.ObserveNormalDisplayNames(rawType, retrieval); cache.SealNormalDisplayNames(owner); return cache;
        }

        // D4-T6A-R1: synthetic side-local Canvas App display evidence only.
        [DataTestMethod, TestCategory("D4T6AType300Display")]
        [DataRow("Nonblank"), DataRow("Omitted"), DataRow("Null"), DataRow("Empty"), DataRow("Blank")]
        [DataRow("Malformed"), DataRow("Missing"), DataRow("Duplicate"), DataRow("Conflict")]
        [DataRow("WrongTable"), DataRow("FailedOrIncomplete")]
        public void T6ACanvasDisplayCopiesOnlyVerifiedLocalScalarsWithExactAvailability(string state)
        {
            var identity = Identity(300, IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new ComponentIdentity[0]);
            var row = T6ACanvasRow(identity.Record.ObjectId.Value, "Synthetic display", "synthetic_internal");
            if (state == "Omitted") { row.Attributes.Remove("name"); row.Attributes.Remove("displayname"); }
            if (state == "Null") { row["name"] = null; row["displayname"] = null; }
            if (state == "Empty") { row["name"] = ""; row["displayname"] = ""; }
            if (state == "Blank") { row["name"] = "  "; row["displayname"] = "  "; }
            if (state == "Malformed") { row["name"] = new OptionSetValue(1); row["displayname"] = new OptionSetValue(1); }
            if (state == "Conflict") row["canvasappid"] = Guid.NewGuid();
            if (state == "WrongTable") row.LogicalName = "appmodule";
            var returned = state == "Missing" ? new Entity[0] : state == "Duplicate" ? new[] { row, row } : new[] { row };
            var cache = T6ACanvasCache(presentation.Source.Snapshot, returned, state == "FailedOrIncomplete");
            var display = ComponentDisplayEvidenceProjector.Project(presentation, cache, null).Single().SourceDisplayEvidence;
            var expected = state == "Omitted" ? DisplayAvailability.AttributeOmitted :
                state == "Duplicate" || state == "Conflict" || state == "WrongTable" ? DisplayAvailability.Ambiguous :
                state == "Malformed" || state == "Missing" || state == "FailedOrIncomplete" ? DisplayAvailability.Unavailable : DisplayAvailability.Available;
            Assert.AreEqual(expected, display.FriendlyName.Availability); Assert.AreEqual(expected, display.InternalName.Availability);
            Assert.AreEqual(expected == DisplayAvailability.Available ? (string)row["displayname"] : null, display.FriendlyName.Value);
            Assert.AreEqual(expected == DisplayAvailability.Available ? (string)row["name"] : null, display.InternalName.Value);
            Assert.AreEqual("NormalMembershipDiagnostic.canvasapp.displayname", display.FriendlyName.DisplayEvidenceSource);
            Assert.AreEqual("NormalMembershipDiagnostic.canvasapp.name", display.InternalName.DisplayEvidenceSource);
            Assert.AreEqual(state == "Missing" ? DisplayBackingCorrelationStatus.Missing :
                state == "Duplicate" || state == "Conflict" || state == "WrongTable" ? DisplayBackingCorrelationStatus.Ambiguous :
                state == "FailedOrIncomplete" ? DisplayBackingCorrelationStatus.Unavailable : DisplayBackingCorrelationStatus.Unique,
                display.BackingCorrelationStatus);
            AssertObjectNotRetrieved(display.ObjectMetadata);
            Assert.AreEqual(IdentityResolutionStatus.Unsupported, identity.Status); Assert.IsNull(identity.ComparisonKey);
            Assert.AreEqual(MembershipPresence.Indeterminate, presentation.Rows.Single().Presence);
            Assert.AreEqual(1, presentation.Summary.IndeterminateUnsupported);
        }

        [TestMethod, TestCategory("D4T6AType300Display")]
        public void T6ACanvasFieldsAreIndependentImmutableAndCannotImportUnrequestedKindOrSubtype()
        {
            var identity = Identity(300, IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { identity }, new ComponentIdentity[0]);
            var row = T6ACanvasRow(identity.Record.ObjectId.Value, null, "synthetic_internal");
            row.Attributes.Remove("displayname");
            // Extra synthetic SDK values are not part of the approved normal display contract.
            row["canvasapptype"] = new OptionSetValue(1); row["canvasappsubtype"] = "invented";
            var cache = T6ACanvasCache(presentation.Source.Snapshot, new[] { row });
            row["name"] = "Mutated SDK payload";
            var side = ComponentDisplayEvidenceProjector.Project(presentation, cache, null).Single().SourceDisplayEvidence;
            Assert.AreEqual(DisplayAvailability.AttributeOmitted, side.FriendlyName.Availability);
            Assert.AreEqual("synthetic_internal", side.InternalName.Value);
            AssertObjectNotRetrieved(side.ObjectMetadata);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, side.ParentDisplayContext.Availability);
            Assert.AreEqual("Canvas App family", side.ComponentCategory.Value);
            Assert.IsNull(identity.ComparisonKey);
            var completed = new CompletedMembershipComparison(presentation, ComponentDisplayEvidenceProjector.Project(presentation, cache, null), DisplayProjectionOutcome.Completed);
            StringAssert.Contains(D365SolutionComparer.Models.Presentation.MembershipGridRow.CreateRows(presentation, completed).Single().ComponentText, "synthetic_internal");
            cache.BeginNormalDisplayNames(presentation.Source.Snapshot);
            Assert.AreEqual(DisplayAvailability.NotRetrieved,
                ComponentDisplayEvidenceProjector.Project(presentation, cache, null).Single().SourceDisplayEvidence.InternalName.Availability);
        }

        [TestMethod, TestCategory("D4T6AType300Display")]
        public void T6ACanvasDisplayRejectsWrongSnapshotOppositeSideAndGenericCacheEvenForSameLocalIds()
        {
            var source = Identity(300, IdentityResolutionStatus.Unsupported);
            var target = new ComponentIdentity(new SolutionComponentRecord(source.Record.SolutionComponentId, 300, source.Record.ObjectId), IdentityResolutionStatus.Unsupported);
            var presentation = Present(new[] { source }, new[] { target });
            var row = T6ACanvasRow(source.Record.ObjectId.Value, "Source synthetic", "source_internal");
            var cache = T6ACanvasCache(presentation.Source.Snapshot, new[] { row });
            var projected = ComponentDisplayEvidenceProjector.Project(presentation, cache, cache);
            Assert.AreEqual("Source synthetic", projected.Single(d => d.SourceDisplayEvidence.HasLocalObservation).SourceDisplayEvidence.FriendlyName.Value);
            Assert.IsTrue(projected.All(d => d.TargetDisplayEvidence.FriendlyName.Availability != DisplayAvailability.Available));
            DataverseComponentMetadataCache.NormalDisplayNameObservation observation;
            var copy = MembershipSnapshot.Complete(presentation.Source.Snapshot.Solution, presentation.Source.Snapshot.Components, presentation.Source.Snapshot.CapturedAt);
            Assert.IsFalse(cache.TryGetNormalDisplayName(copy, source, out observation));
            Assert.IsFalse(cache.TryGetNormalDisplayName(presentation.Source.Snapshot, target, out observation));
            var generic = Cache(row);
            Assert.AreEqual(DisplayAvailability.NotRetrieved, ComponentDisplayEvidenceProjector.Project(presentation, generic, null)
                .Single(d => d.SourceDisplayEvidence.HasLocalObservation).SourceDisplayEvidence.InternalName.Availability);
            Assert.AreEqual(2, presentation.Rows.Count); Assert.AreEqual(0, presentation.Summary.PresentInBoth);
        }

        internal static Entity T6ACanvasRow(Guid id, string display, string name) => new Entity("canvasapp", id)
        { ["canvasappid"] = id, ["displayname"] = display, ["name"] = name,
          ["uniquecanvasappid"] = "synthetic_local_installation", ["componentstate"] = new OptionSetValue(0), ["ismanaged"] = false };

        private static DataverseComponentMetadataCache T6ACanvasCache(MembershipSnapshot owner, Entity[] returned, bool failed = false)
        {
            var ids = owner.Components.Select(i => i.Record.ObjectId.Value).Distinct().ToList();
            var failures = new Dictionary<Guid, string>();
            if (failed) foreach (var id in ids) failures.Add(id, "Synthetic fault or incomplete batch; not parsed");
            var retrieval = new DiagnosticQueryResult(ids, ids.ToDictionary(id => id, id => returned.Where(r => r.Id == id).ToList()),
                failures, ids.ToDictionary(id => id, id => new List<Entity>()), returned.Length, !failed);
            var cache = new DataverseComponentMetadataCache(); cache.BeginNormalDisplayNames(owner);
            cache.ObserveCanvasAppDisplayNames(retrieval); cache.SealNormalDisplayNames(owner); return cache;
        }
        // End D4-T6A-R1 additions.

        public TestContext TestContext { get; set; }

        private static EntityMetadata Parent(AttributeMetadata[] attributes)
        {
            var parent = new EntityMetadata { MetadataId = Guid.NewGuid(), LogicalName = "synthetic_parent" };
            typeof(EntityMetadata).GetProperty("Attributes").SetValue(parent, attributes);
            typeof(EntityMetadata).GetProperty("OneToManyRelationships").SetValue(parent, new OneToManyRelationshipMetadata[0]);
            typeof(EntityMetadata).GetProperty("ManyToOneRelationships").SetValue(parent, new OneToManyRelationshipMetadata[0]);
            typeof(EntityMetadata).GetProperty("ManyToManyRelationships").SetValue(parent, new ManyToManyRelationshipMetadata[0]);
            return parent;
        }

        private static SolutionComponentRecord Record(int type) => new SolutionComponentRecord(Guid.NewGuid(), type, Guid.NewGuid());
        private static ComponentIdentity Identity(int type, IdentityResolutionStatus status = IdentityResolutionStatus.Resolved, string key = "synthetic-key") =>
            new ComponentIdentity(Record(type), status, status == IdentityResolutionStatus.Resolved ? key : null);
        private static Entity App(ComponentIdentity identity, string friendly, string internalName) =>
            new Entity("appmodule", identity.Record.ObjectId.Value) { ["name"] = friendly, ["uniquename"] = internalName };
        private static DataverseComponentMetadataCache Cache(params Entity[] rows)
        {
            var cache = new DataverseComponentMetadataCache();
            foreach (var row in rows) cache.StoreEntityRow(row, row.Id);
            return cache;
        }
        private static ComponentSideDisplayEvidence Project(ComponentIdentity identity, DataverseComponentMetadataCache cache) =>
            ComponentDisplayEvidenceProjector.Project(Present(new[] { identity }, new ComponentIdentity[0]), cache, null)[0].SourceDisplayEvidence;
        private static MembershipComparisonPresentation Present(ComponentIdentity[] source, ComponentIdentity[] target)
        {
            var sourceSnapshot = MembershipSnapshot.Complete(MembershipTestData.Solution(), source, DateTimeOffset.UtcNow);
            var targetSnapshot = MembershipSnapshot.Complete(MembershipTestData.Solution(), target, DateTimeOffset.UtcNow);
            return new MembershipResultPresenter().Create(MembershipEnvironmentResult.FromSnapshot("Synthetic Source", sourceSnapshot, 0, TimeSpan.Zero),
                MembershipEnvironmentResult.FromSnapshot("Synthetic Target", targetSnapshot, 0, TimeSpan.Zero));
        }
    }
}
