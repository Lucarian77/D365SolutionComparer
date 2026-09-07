namespace D365SolutionComparer.Infrastructure
{
    /// <summary>Counts retrieved so far; no total or completion guarantee is implied.</summary>
    public sealed class RetrievalProgress
    {
        internal RetrievalProgress(int pagesRetrieved, int recordsRetrieved)
        {
            PagesRetrieved = pagesRetrieved;
            RecordsRetrieved = recordsRetrieved;
        }

        public int PagesRetrieved { get; }
        public int RecordsRetrieved { get; }
    }
}
