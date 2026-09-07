namespace D365SolutionComparer.Infrastructure
{
    public enum MembershipOperationStage
    {
        ValidatingEnvironment,
        ReadingMembership,
        ResolvingIdentities,
        Completed
    }

    /// <summary>Operation progress without implying a server-side total.</summary>
    public sealed class MembershipOperationProgress
    {
        internal MembershipOperationProgress(MembershipOperationStage stage, string message,
            int pagesRetrieved = 0, int recordsRetrieved = 0)
        {
            Stage = stage;
            Message = message ?? string.Empty;
            PagesRetrieved = pagesRetrieved;
            RecordsRetrieved = recordsRetrieved;
        }

        public MembershipOperationStage Stage { get; }
        public string Message { get; }
        public int PagesRetrieved { get; }
        public int RecordsRetrieved { get; }
    }
}
