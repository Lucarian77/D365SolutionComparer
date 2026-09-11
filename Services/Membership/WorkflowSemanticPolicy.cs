using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using D365SolutionComparer.Models.Membership;
using Microsoft.Xrm.Sdk;

namespace D365SolutionComparer.Services.Membership
{
    /// <summary>
    /// Phase 2G.4 solution-scoped semantic candidate (not a Microsoft alternate key).
    /// Name alone is never identity. Length framing preserves punctuation and whitespace.
    /// Mode/subprocess/state/deployment values never disambiguate identity collisions.
    /// </summary>
    internal static class WorkflowSemanticPolicy
    {
        internal const string Prefix = "workflow-semantic:v1:";
        internal static readonly string[] Columns = { "workflowid", "uniquename", "name", "type",
            "category", "primaryentity", "mode", "parentworkflowid", "workflowidunique", "statecode",
            "statuscode", "componentstate", "ismanaged", "subprocess", "businessprocesstype",
            "modernflowtype", "uiflowtype" };
        internal const string Coverage = "Configuration-only definition coverage; XAML, JSON, clientdata and execution logic are not compared.";

        internal static int? Option(Entity row, string field) =>
            (row.GetAttributeValue<object>(field) as OptionSetValue)?.Value;
        internal static string Text(Entity row, string field) => row.GetAttributeValue<object>(field) as string;

        internal static string Candidate(Entity row, out string reason)
        {
            reason = null;
            if (Option(row, "type") != 1) { reason = "Missing or unsupported workflow definition type."; return null; }
            var category = Option(row, "category");
            if (!category.HasValue || category < 0 || category > 7)
            { reason = "Missing or unsupported workflow category."; return null; }
            if (string.IsNullOrWhiteSpace(Text(row, "name")))
            { reason = "Missing required semantic field: name."; return null; }
            if (string.IsNullOrWhiteSpace(Text(row, "primaryentity")))
            { reason = "Missing required semantic field: primaryentity."; return null; }
            string subtype;
            if (!Subtype(row, category.Value, out subtype, out reason)) return null;
            return Prefix + string.Concat(new[] { Text(row, "name"), "1", category.Value.ToString(CultureInfo.InvariantCulture),
                Text(row, "primaryentity"), subtype }.Select(value => value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value));
        }

        private static bool Subtype(Entity row, int category, out string value, out string reason)
        {
            value = ""; reason = null;
            string field = category == 4 ? "businessprocesstype" : category == 5 ? "modernflowtype" :
                category == 6 ? "uiflowtype" : null;
            if (field == null) return true;
            var option = Option(row, field);
            if (!option.HasValue || option < 0 || option > 1)
            { reason = "Missing or unsupported workflow subtype: " + field + "."; return false; }
            value = field + "=" + option.Value.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        internal static bool Configuration(Entity row, out Dictionary<string, string> properties, out string reason)
        {
            properties = null;
            reason = "Missing or unsupported workflow configuration evidence.";
            var category = Option(row, "category");
            if (Option(row, "type") != 1 || !category.HasValue || category < 0 || category > 7 ||
                string.IsNullOrWhiteSpace(Text(row, "primaryentity")) ||
                (Option(row, "mode") != 0 && Option(row, "mode") != 1) ||
                !(row.GetAttributeValue<object>("subprocess") is bool)) return false;
            string subtype;
            if (!Subtype(row, category.Value, out subtype, out reason)) return false;
            properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in new[] { "type", "category", "mode" })
                properties[field] = Option(row, field).Value.ToString(CultureInfo.InvariantCulture);
            properties["primaryentity"] = Text(row, "primaryentity").ToLowerInvariant();
            properties["subprocess"] = ((bool)row["subprocess"]).ToString();
            // Inapplicable subtype fields are null by contract, not fabricated defaults.
            properties["businessprocesstype"] = category == 4 ? Option(row, "businessprocesstype").Value.ToString(CultureInfo.InvariantCulture) : null;
            properties["modernflowtype"] = category == 5 ? Option(row, "modernflowtype").Value.ToString(CultureInfo.InvariantCulture) : null;
            properties["uiflowtype"] = category == 6 ? Option(row, "uiflowtype").Value.ToString(CultureInfo.InvariantCulture) : null;
            reason = null;
            return true;
        }
    }
}
