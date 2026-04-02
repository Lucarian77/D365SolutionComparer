namespace D365SolutionComparer.Models
{
    public class SolutionCompareResult
    {
        public string UniqueName { get; set; }

        public string SourceDisplayName { get; set; }
        public string TargetDisplayName { get; set; }

        public string SourceVersion { get; set; }
        public string TargetVersion { get; set; }

        public string SourcePublisher { get; set; }
        public string TargetPublisher { get; set; }

        public string SourcePackageType { get; set; }
        public string TargetPackageType { get; set; }
        public string PackageTypeStatus { get; set; }

        public string Status { get; set; }

        public bool IsManagedUnmanagedMismatch
        {
            get
            {
                return string.Equals(
                    PackageTypeStatus,
                    "Managed/Unmanaged Mismatch",
                    System.StringComparison.OrdinalIgnoreCase);
            }
        }

        public bool IsPackageTypeMismatch
        {
            get
            {
                return IsManagedUnmanagedMismatch;
            }
        }
    }
}