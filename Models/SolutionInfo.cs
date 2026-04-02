namespace D365SolutionComparer.Models
{
    public class SolutionInfo
    {
        public string UniqueName { get; set; }

        public string DisplayName { get; set; }

        public string Version { get; set; }

        public string Publisher { get; set; }

        public bool? IsManaged { get; set; }

        public string PackageType
        {
            get
            {
                if (!IsManaged.HasValue)
                {
                    return string.Empty;
                }

                return IsManaged.Value ? "Managed" : "Unmanaged";
            }
        }
    }
}