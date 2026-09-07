using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models;
using D365SolutionComparer.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365SolutionComparer.Tests
{
    [TestClass]
    public class SolutionComparisonTests
    {
        private static SolutionInfo Solution(string key = "sample") => new SolutionInfo
        {
            UniqueName = key, DisplayName = "Sample", Version = "1.0.0.0", Publisher = "Publisher", IsManaged = true
        };

        private static SolutionCompareResult Compare(SolutionInfo source, SolutionInfo target) =>
            new SolutionComparisonService().Compare(new List<SolutionInfo> { source }, new List<SolutionInfo> { target }).Single();

        [DataTestMethod]
        [DataRow(0, "Match")]
        [DataRow(1, "Version Mismatch")]
        [DataRow(2, "Publisher Mismatch")]
        [DataRow(3, "Multiple Differences")]
        [DataRow(4, "Display Name Mismatch")]
        [DataRow(5, "Multiple Differences")]
        [DataRow(6, "Multiple Differences")]
        [DataRow(7, "Multiple Differences")]
        [DataRow(8, "Managed/Unmanaged Mismatch")]
        [DataRow(9, "Multiple Differences")]
        [DataRow(10, "Multiple Differences")]
        [DataRow(11, "Multiple Differences")]
        [DataRow(12, "Multiple Differences")]
        [DataRow(13, "Multiple Differences")]
        [DataRow(14, "Multiple Differences")]
        [DataRow(15, "Multiple Differences")]
        public void AllFieldDifferenceCombinationsPreserveBaselineStatuses(int mask, string expected)
        {
            var target = Solution();
            if ((mask & 1) != 0) target.Version = "2.0.0.0";
            if ((mask & 2) != 0) target.Publisher = "Different";
            if ((mask & 4) != 0) target.DisplayName = "Different";
            if ((mask & 8) != 0) target.IsManaged = false;
            var result = Compare(Solution(), target);
            Assert.AreEqual(expected, result.Status);
            Assert.AreEqual((mask & 8) != 0 ? "Managed/Unmanaged Mismatch" : "Match", result.PackageTypeStatus);
            Assert.AreEqual((mask & 8) != 0, result.IsManagedUnmanagedMismatch);
        }

        [TestMethod]
        public void UniqueNamesIgnoreCaseAndPreserveSourceSpelling()
        {
            var result = Compare(Solution("Sample"), Solution("sAMPLE"));
            Assert.AreEqual("Sample", result.UniqueName);
            Assert.AreEqual("Match", result.Status);
        }

        [TestMethod]
        public void UniqueNamesAreNotTrimmedAndDisplayNamesDoNotEstablishIdentity()
        {
            var results = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution("sample ") }, new List<SolutionInfo> { Solution("sample") });
            CollectionAssert.AreEquivalent(new[] { "Missing in Source", "Missing in Target" }, results.Select(r => r.Status).ToArray());
        }

        [TestMethod]
        public void ComparedFieldsIgnoreCaseAndSurroundingWhitespaceWithoutChangingDisplayedValues()
        {
            var target = Solution();
            target.Version = " 1.0.0.0 "; target.Publisher = " PUBLISHER "; target.DisplayName = " SAMPLE ";
            var result = Compare(Solution(), target);
            Assert.AreEqual("Match", result.Status);
            Assert.AreEqual(" PUBLISHER ", result.TargetPublisher);
            Assert.AreEqual(" SAMPLE ", result.TargetDisplayName);
            Assert.AreEqual(" 1.0.0.0 ", result.TargetVersion);
        }

        [TestMethod]
        public void VersionsRemainStringComparisons()
        {
            var target = Solution(); target.Version = "1.0";
            Assert.AreEqual("Version Mismatch", Compare(Solution(), target).Status);
        }

        [TestMethod]
        public void NullAndWhitespaceMetadataCompareEqually()
        {
            var source = Solution(); var target = Solution();
            source.Version = null; source.Publisher = null; source.DisplayName = null;
            target.Version = " "; target.Publisher = "\t"; target.DisplayName = "";
            var result = Compare(source, target);
            Assert.AreEqual("Match", result.Status);
            Assert.AreEqual(string.Empty, result.SourcePublisher);
        }

        [TestMethod]
        public void DuplicateUniqueNamesKeepFirstRecordOnEachSide()
        {
            var source = Solution("sample"); source.Version = "source first";
            var target = Solution("SAMPLE"); target.Version = "target first";
            var result = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { source, Solution("SAMPLE") },
                new List<SolutionInfo> { target, Solution("sample") }).Single();
            Assert.AreEqual("source first", result.SourceVersion);
            Assert.AreEqual("target first", result.TargetVersion);
        }

        [TestMethod]
        public void NullAndEmptyUniqueNamesCollapseToTheSameKey()
        {
            var result = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution(null), Solution("") }, new List<SolutionInfo> { Solution("") }).Single();
            Assert.AreEqual(string.Empty, result.UniqueName);
            Assert.AreEqual("Match", result.Status);
        }

        [TestMethod]
        public void NullOrEmptyInventoriesProduceNoRows()
        {
            var comparer = new SolutionComparisonService();
            Assert.AreEqual(0, comparer.Compare(null, null).Count);
            Assert.AreEqual(0, comparer.Compare(new List<SolutionInfo>(), new List<SolutionInfo>()).Count);
        }

        [DataTestMethod]
        [DataRow(true, "Missing in Target")]
        [DataRow(false, "Missing in Source")]
        public void OneSidedInventoriesRetainValuesAndBothMissingStatuses(bool sourcePresent, string status)
        {
            var items = new List<SolutionInfo> { Solution() };
            var result = new SolutionComparisonService().Compare(sourcePresent ? items : null, sourcePresent ? null : items).Single();
            Assert.AreEqual(status, result.Status);
            Assert.AreEqual(status, result.PackageTypeStatus);
            Assert.AreEqual(sourcePresent ? "Sample" : "", result.SourceDisplayName);
            Assert.AreEqual(sourcePresent ? "" : "Sample", result.TargetDisplayName);
            Assert.AreEqual(sourcePresent ? "Managed" : "", result.SourcePackageType);
            Assert.AreEqual(sourcePresent ? "" : "Managed", result.TargetPackageType);
        }

        [TestMethod]
        public void UnionIsSortedOrdinalIgnoreCase()
        {
            var results = new SolutionComparisonService().Compare(
                new List<SolutionInfo> { Solution("z"), Solution("B") }, new List<SolutionInfo> { Solution("a") });
            CollectionAssert.AreEqual(new[] { "a", "B", "z" }, results.Select(r => r.UniqueName).ToArray());
        }

        [TestMethod]
        public void NullableManagedStatePreservesExistingUnknownBehavior()
        {
            var source = Solution(); source.IsManaged = null;
            Assert.AreEqual("", source.PackageType);
            Assert.AreEqual("Managed/Unmanaged Mismatch", Compare(source, Solution()).Status);
            var target = Solution(); target.IsManaged = null;
            Assert.AreEqual("Match", Compare(source, target).Status);
            target.IsManaged = false;
            Assert.AreEqual("Unmanaged", target.PackageType);
            Assert.AreEqual("Managed/Unmanaged Mismatch", Compare(source, target).Status);
        }

        [DataTestMethod]
        [DataRow("managed/unmanaged mismatch", true)]
        [DataRow(" Managed/Unmanaged Mismatch ", false)]
        [DataRow("Package Type Mismatch", false)]
        [DataRow(null, false)]
        public void ModelMismatchAliasesKeepExactBaselineSemantics(string status, bool expected)
        {
            var result = new SolutionCompareResult { PackageTypeStatus = status };
            Assert.AreEqual(expected, result.IsManagedUnmanagedMismatch);
            Assert.AreEqual(expected, result.IsPackageTypeMismatch);
        }
    }
}
