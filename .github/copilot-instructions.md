# Copilot Instructions — PsGraphReference

## Architecture

This is a single-page web app (`public/index.html`) that serves as an interactive reference for Microsoft Graph PowerShell cmdlets. It is a self-contained HTML file (~700 lines) with inline CSS and JavaScript — no build system, no framework, no dependencies. Web assets live in `public/` which is the Vercel `outputDirectory`.

**Data flow:**
1. On page load, `loadCmdlets()` fetches `data/cmdlets.json` (relative URL) via the Fetch API
2. If the fetch fails (e.g., opened as a local file without an HTTP server), it falls back to hardcoded embedded data inside `loadEmbeddedData()`
3. Cmdlet data is normalized into a flat array of `{name, category, description, syntax, examples[], permissions[]}` objects
4. Filtering/search operates client-side against this in-memory array

**Data generation:** `scripts/get-graphcmdlets.ps1` extracts cmdlet metadata from installed `Microsoft.Graph.*` PowerShell modules and outputs JSON to `public/data/cmdlets.json`. It requires the `PowerShellForGitHub` and `Powershell-Yaml` modules.

## Running Locally

Serve the `public/` directory via any local HTTP server to enable JSON loading (direct file:// access triggers CORS fallback):

```bash
python -m http.server 8000 -d public      # then open http://localhost:8000
```

## Key Conventions

- **Single-file frontend**: All HTML, CSS, and JS live in `public/index.html`. Do not extract into separate files unless explicitly asked.
- **No build step**: Changes to `public/index.html` are immediately testable — just reload the browser.
- **Dual data sources**: Any changes to the cmdlet data schema must be reflected in both `public/data/cmdlets.json` AND the embedded fallback data in `loadEmbeddedData()`.
- **Cmdlet object shape**: `{name, category, description, syntax, examples: string[], permissions: string[]}`. The JSON loader normalizes various input shapes (array, `{cmdlets: [...]}`, single object) into this format.
- **Requirements doc**: `claude_requirements_md.md` contains the full enhancement roadmap (automated scraping, GitHub Actions, schema validation, etc.). Reference it for planned features and data structure goals.

## Frontend Design Guidelines

When building or modifying UI in this project, follow these principles to produce distinctive, production-grade interfaces — not generic "AI slop."

### Design Thinking

Before coding, commit to a clear aesthetic direction:
- **Purpose**: What problem does this interface solve? Who uses it? (Here: endpoint engineers and IT pros looking up Graph cmdlets quickly.)
- **Tone**: Pick a deliberate direction — brutally minimal, retro-futuristic, editorial, industrial, luxury, etc. — and execute it with precision. Bold maximalism and refined minimalism both work; the key is intentionality.
- **Differentiation**: What makes this memorable? What's the one thing someone will remember?

### Aesthetics Rules

- **Typography**: Choose distinctive, characterful fonts — never default to Inter, Roboto, Arial, or system fonts. Pair a display font with a refined body font.
- **Color & Theme**: Commit to a cohesive palette using CSS variables. Dominant colors with sharp accents outperform timid, evenly-distributed palettes. Avoid cliché purple-gradient-on-white.
- **Motion**: Prioritize CSS-only animations. Focus on high-impact moments — one well-orchestrated page load with staggered `animation-delay` creates more delight than scattered micro-interactions. Surprise with scroll-triggered and hover states.
- **Spatial Composition**: Use asymmetry, overlap, grid-breaking elements, generous negative space, or controlled density — not predictable symmetric layouts.
- **Backgrounds & Depth**: Create atmosphere with gradient meshes, noise textures, geometric patterns, layered transparencies, dramatic shadows, or grain overlays — not flat solid colors.

### Anti-Patterns (Never Do)

- Generic font stacks (Inter, Roboto, Arial, system fonts)
- Purple gradients on white backgrounds
- Cookie-cutter card grids with no personality
- Converging on the same "safe" choices across generations (e.g., always picking Space Grotesk)

Match implementation complexity to the aesthetic vision: maximalist designs need elaborate animations; minimalist designs need precise spacing and typography.
