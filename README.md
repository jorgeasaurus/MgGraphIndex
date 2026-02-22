# Microsoft Graph PowerShell Reference

Interactive reference for Microsoft Graph PowerShell cmdlets. Dynamically discovers all installed `Microsoft.Graph.*` modules, extracts metadata, and serves a searchable frontend with lazy-loaded detail.

## Quick Start

```bash
# Serve locally
python3 -m http.server 8000 -d public
# Open http://localhost:8000
```

Also deployed via Vercel (`vercel.json` points at `public/`).

## Generating Data

### Python parser (primary)

Parses markdown documentation from the [microsoftgraph-docs-powershell](https://github.com/MicrosoftDocs/microsoftgraph-docs-powershell) repo. No module installation required.

```bash
# Clone the docs repo
git clone --depth 1 --filter=blob:none --sparse \
  https://github.com/MicrosoftDocs/microsoftgraph-docs-powershell.git docs
cd docs && git sparse-checkout set microsoftgraph/graph-powershell-1.0 microsoftgraph/graph-powershell-beta
cd ..

# Run the parser
python3 scripts/parse_docs.py docs/microsoftgraph
```

### PowerShell script (alternative)

Requires the Microsoft.Graph PowerShell modules to be installed. Slower but extracts live module data.

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
.\scripts\get-graphcmdlets.ps1 -IncludeBeta -Verbose
```

### Output

The script produces three data tiers:

| File | Purpose | Size |
|------|---------|------|
| `public/data/manifest.json` | Lightweight summary for search/filter/cards | ~200-400KB |
| `public/data/modules/*.json` | Per-module detail (syntax, examples, permissions) | Loaded lazily on card expand |
| `public/data/cmdlets.json` | Flat file with all fields (backward compat) | ~2MB+ |

### Manifest entry shape

```json
{
  "name": "Get-MgUser",
  "verb": "Get",
  "category": "User Management",
  "module": "Microsoft.Graph.Users",
  "apiVersion": "v1.0",
  "description": "Retrieve properties and relationships of user object.",
  "permissionCount": 3,
  "hasExamples": true
}
```

### Module detail shape (`modules/Microsoft.Graph.Users.json`)

```json
{
  "module": "Microsoft.Graph.Users",
  "cmdlets": {
    "Get-MgUser": {
      "syntax": "Get-MgUser [-UserId <String>] ...",
      "examples": ["Get-MgUser -All | ..."],
      "permissions": ["User.Read.All", "User.ReadWrite.All"]
    }
  }
}
```

## Project Structure

```
public/
  index.html              Main SPA (~1K lines)
  data/
    manifest.json          Cmdlet summaries (loaded on init)
    cmdlets.json           Full flat file (backward compat)
    modules/               Per-module detail files (lazy loaded)
scripts/
  parse_docs.py            Python parser (primary, stdlib only)
  get-graphcmdlets.ps1     PowerShell extraction (alternative)
vercel.json                Deployment config + cache headers
```

## Features

- **Dynamic module discovery** -- scans all installed `Microsoft.Graph.*` modules, not a hardcoded list
- **API version tracking** -- each cmdlet tagged as `v1.0` or `beta`, filterable in the UI
- **Two-tier loading** -- manifest loads on init for fast search; detail fetched per-module on card expand
- **Fuzzy search** -- weighted scoring across name, description, category
- **Filters** -- category, module, API version, verb, permission count, sort order
- **URL hash state** -- all filters serialized in the URL for shareable links
- **Keyboard navigation** -- `/` to search, `j`/`k` to navigate, `Enter` to expand
- **Export** -- filtered results to JSON or CSV
- **Presets** -- save and restore filter combinations

## Troubleshooting

**No data loads**: Run `parse_docs.py` (or `get-graphcmdlets.ps1`) first to generate the data files, then serve via HTTP (not `file://`).

**Module not found**: `Install-Module Microsoft.Graph -Scope CurrentUser`

**Permission errors on extraction**: Run with `-Verbose` to identify which modules fail. Some modules require `Connect-MgGraph` first for permission lookups.
