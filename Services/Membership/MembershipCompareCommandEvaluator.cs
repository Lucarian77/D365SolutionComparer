namespace D365SolutionComparer.Services.Membership
{
    /// <summary>UI-independent gating for the live membership command.</summary>
    public static class MembershipCompareCommandEvaluator
    {
        public static bool CanExecute(bool hasSelectedComparisonRow, bool sourceConnected, bool targetConnected,
            bool sourceDataLoaded, bool targetDataLoaded, bool sourceSolutionPresent, bool targetSolutionPresent,
            bool operationInProgress)
        {
            return !operationInProgress && hasSelectedComparisonRow && sourceConnected && targetConnected &&
                sourceDataLoaded && targetDataLoaded && (sourceSolutionPresent || targetSolutionPresent);
        }
    }
}
