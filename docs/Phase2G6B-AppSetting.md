# AppSetting identity and structural comparison

AppSetting support relies on verified registration metadata and side-owned backing records. Registration names and expected table relationships must be verified rather than inferred from a numeric component type alone.

## Identity

The portable key uses the related parent app unique name and setting-definition unique name with explicit length framing and ordinal case-insensitive equality. Local identifiers establish within-side correlation; they are not portable keys. Display names, setting values and numeric type identifiers do not resolve duplicate identities.

Missing or conflicting references retain unresolved or ambiguous outcomes. Duplicate framed keys prevent unsafe absence conclusions. Identity resolution and structural definition availability remain separate.

## Structural definition evidence

Only verified, readable metadata of the expected types supports comparison of datatype, isoverridable, overridablelevel and releaselevel. Missing, inaccessible or wrong-type metadata and unavailable values do not become Match through defaulting. Structural evidence can remain unresolved while identity is resolved.

Setting values and defaults, local audit/reference identifiers, owners and timestamps are outside this structural comparison contract. Secret or default values must not be used to disambiguate identities. Diagnostics and exported evidence must preserve the established allowlist rather than formatting arbitrary server rows.

## Reading results

Review available side-owned identity and supported property evidence. Treat unavailable or unsupported metadata honestly, retain cancellation and coverage safeguards, and do not force results to match an earlier snapshot. This document contains no environment-specific examples, comparison totals or capture instructions.

[Microsoft settings documentation](https://learn.microsoft.com/en-us/power-apps/maker/data-platform/create-edit-configure-settings)
