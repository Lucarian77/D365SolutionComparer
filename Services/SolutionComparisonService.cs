using System;
using System.Collections.Generic;
using System.Linq;
using D365SolutionComparer.Models;

namespace D365SolutionComparer.Services
{
    public class SolutionComparisonService
    {
        public List<SolutionCompareResult> Compare(List<SolutionInfo> sourceSolutions, List<SolutionInfo> targetSolutions)
        {
            sourceSolutions = sourceSolutions ?? new List<SolutionInfo>();
            targetSolutions = targetSolutions ?? new List<SolutionInfo>();

            var results = new List<SolutionCompareResult>();

            var sourceLookup = sourceSolutions
                .GroupBy(s => s.UniqueName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var targetLookup = targetSolutions
                .GroupBy(s => s.UniqueName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var allUniqueNames = sourceLookup.Keys
                .Union(targetLookup.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var uniqueName in allUniqueNames)
            {
                sourceLookup.TryGetValue(uniqueName, out var source);
                targetLookup.TryGetValue(uniqueName, out var target);

                var packageTypeStatus = GetPackageTypeStatus(source, target);

                results.Add(new SolutionCompareResult
                {
                    UniqueName = uniqueName,
                    SourceDisplayName = source?.DisplayName ?? string.Empty,
                    TargetDisplayName = target?.DisplayName ?? string.Empty,
                    SourceVersion = source?.Version ?? string.Empty,
                    TargetVersion = target?.Version ?? string.Empty,
                    SourcePublisher = source?.Publisher ?? string.Empty,
                    TargetPublisher = target?.Publisher ?? string.Empty,
                    SourcePackageType = source?.PackageType ?? string.Empty,
                    TargetPackageType = target?.PackageType ?? string.Empty,
                    PackageTypeStatus = packageTypeStatus,
                    Status = GetStatus(source, target, packageTypeStatus)
                });
            }

            return results;
        }

        private string GetStatus(SolutionInfo source, SolutionInfo target, string packageTypeStatus)
        {
            if (source == null)
                return "Missing in Source";

            if (target == null)
                return "Missing in Target";

            bool versionMatches = AreEqual(source.Version, target.Version);
            bool publisherMatches = AreEqual(source.Publisher, target.Publisher);
            bool displayNameMatches = AreEqual(source.DisplayName, target.DisplayName);
            bool packageTypeMatches = string.Equals(packageTypeStatus, "Match", StringComparison.OrdinalIgnoreCase);

            int mismatchCount = 0;

            if (!versionMatches) mismatchCount++;
            if (!publisherMatches) mismatchCount++;
            if (!displayNameMatches) mismatchCount++;
            if (!packageTypeMatches) mismatchCount++;

            if (mismatchCount == 0)
                return "Match";

            if (mismatchCount > 1)
                return "Multiple Differences";

            if (!versionMatches)
                return "Version Mismatch";

            if (!publisherMatches)
                return "Publisher Mismatch";

            if (!displayNameMatches)
                return "Display Name Mismatch";

            return packageTypeStatus;
        }

        private string GetPackageTypeStatus(SolutionInfo source, SolutionInfo target)
        {
            if (source == null)
                return "Missing in Source";

            if (target == null)
                return "Missing in Target";

            if (AreEqual(source.PackageType, target.PackageType))
                return "Match";

            return "Managed/Unmanaged Mismatch";
        }

        private bool AreEqual(string left, string right)
        {
            return string.Equals(
                left?.Trim() ?? string.Empty,
                right?.Trim() ?? string.Empty,
                StringComparison.OrdinalIgnoreCase);
        }
    }
}