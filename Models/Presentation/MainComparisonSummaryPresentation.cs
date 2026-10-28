using System;
using System.Collections.Generic;

namespace D365SolutionComparer.Models.Presentation
{
    // Presentation state only. Counts, filters and authoritative rows remain owned by the control.
    internal sealed class MainComparisonSummaryPresentation
    {
        internal bool Expanded { get; private set; }
        internal string Caption => Expanded ? "Detailed breakdown ▴" : "Detailed breakdown ▾";
        internal void Toggle() { Expanded = !Expanded; }
        internal int DetailsHeight(int measuredHeight) => Expanded ? Math.Max(0, measuredHeight) : 0;

        internal const string StatusHelp = "Status colours: Match — green (Package Type Status Match is neutral); " +
            "Version Mismatch — amber; Publisher Mismatch / Display Name Mismatch — blue; " +
            "Managed/Unmanaged Mismatch / Package Type Mismatch — orange; Multiple Differences — purple; " +
            "Missing in Source / Missing in Target — red. Selected rows use blue with white text. " +
            "Colours do not change comparison meaning.";

        internal static string ScopeText(int all, int visible) =>
            "Cards: all " + all + " comparison results. Detailed breakdown: current filtered view (" + visible + " rows).";

        internal static string Breakdown(IReadOnlyList<int> counts)
        {
            if (counts == null || counts.Count != 10) throw new ArgumentException("Ten retained summary counts are required.", nameof(counts));
            return "Current filtered view: Total=" + counts[0] + " | Match=" + counts[1] + Environment.NewLine +
                "Overall statuses: Version differences=" + counts[2] + " | Publisher differences=" + counts[3] +
                " | Display-name differences=" + counts[4] + " | Multiple differences=" + counts[7] +
                " | Missing in Source=" + counts[8] + " | Missing in Target=" + counts[9] + Environment.NewLine +
                "Overlapping package indicators: Package-type differences=" + counts[5] +
                " | Managed/unmanaged differences=" + counts[6] + " (subset of package-type differences)." + Environment.NewLine +
                "Version, publisher and display-name counts are single-difference overall statuses. " +
                "A Multiple Differences row can also count as a package indicator; these indicators must not be summed to Different.";
        }
    }
}
