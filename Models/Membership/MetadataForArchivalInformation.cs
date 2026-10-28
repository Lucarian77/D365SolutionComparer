using System;
using System.Collections.Generic;
using System.Linq;

namespace D365SolutionComparer.Models.Membership
{
    // Detached, side-owned information. None of these values is a definition comparison.
    internal sealed class MetadataForArchivalInformation
    {
        private readonly IReadOnlyDictionary<string, ComponentDisplayValue> values;
        internal MetadataForArchivalInformation(DisplayEnvironmentSide side, SolutionComponentRecord record,
            IDictionary<string, ComponentDisplayValue> fields)
        {
            Side = side; SolutionComponentId = record.SolutionComponentId; ObjectId = record.ObjectId;
            values = new System.Collections.ObjectModel.ReadOnlyDictionary<string, ComponentDisplayValue>(
                new Dictionary<string, ComponentDisplayValue>(fields, StringComparer.Ordinal));
        }
        internal DisplayEnvironmentSide Side { get; }
        internal Guid SolutionComponentId { get; }
        internal Guid? ObjectId { get; }
        internal ComponentDisplayValue Field(string name) => values.ContainsKey(name) ? values[name] :
            new ComponentDisplayValue(DisplayAvailability.NotRetrieved);
        internal string Text => Side + " MetadataForArchival information (not compared): SolutionComponentId=" +
            SolutionComponentId + "; ObjectId=" + ObjectId + "; " + string.Join("; ",
                new[] { "Name", "Related Table", "Managed", "Archival availability", "Archival readiness", "Backing solution ID", "Component unique ID" }
                .Select(name => name + "=" + Value(Field(name))));
        private static string Value(ComponentDisplayValue field) => field.Availability == DisplayAvailability.Available
            ? field.Value ?? "[Retrieved null]" : "[" + field.Availability + "]";
    }
}
