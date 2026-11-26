# Peptide structure panel

## Scope

The synthesis tab displays a neutral, fully deprotected 2D structure of the final
peptide. The existing sequence, dose calculations, progress and stored synthesis
records are not changed by this panel. Neutral form is a depiction convention,
not a prediction of pH-dependent charge, salt composition or counterions.

Supported building blocks: the 20 proteinogenic L amino acids, glycine, Aib and
L-norleucine. Explicit D names for supported chiral amino acids are also mapped;
D-Thr and D-Ile invert both stereocentres, not the allo isomers. Exact aliases
from the chemical catalog are accepted. Known protecting groups are removed in
the displayed final product, including Alloc and OAll. Unknown residues are not
replaced with generic atoms or guessed from their mass. No automatic structure
lookup is inferred from a molecular-weight result.

Paired asterisk markers use the application's existing interpretation: the first
and last residue inside each pair are the side-chain bridge endpoints. Cys/Cys
adds an S-S bond; Lys/Asp or Lys/Glu adds the side-chain amide bond and removes
the side-chain carboxylic OH. Terminal amidation follows the selected resin.
Other marked endpoints or malformed markers produce an unavailable state.
The current preview limit is 100 residues. This is a viewer, not an interactive
chemical editor. Other terminal modifications and arbitrary user-defined
residue structures are not implemented.

## Rendering and deployment

RDKit 2026.9.1 (BSD-3-Clause) is vendored in `vendor/rdkit/`. Preserve the
directory structure, the JavaScript, WebAssembly binary, and license when
publishing. RDKit validates molecular graphs and generates coordinates and
stereobonds in a worker. Formulas are derived from the resulting graph.
No peptide or inventory data is sent to a remote depiction service.

Primary technical references:

- https://rdkitjs.com/
- https://www.rdkit.org/rdkitjs/beta/GettingStartedInJS.html
- https://www.rcsb.org/ligand/ALA (example CCD stereochemical identity)
- https://pubchem.ncbi.nlm.nih.gov/compound/L-alanine

Controls: zoom, fit, fullscreen, SVG and MOL export. Print waits for the same
depiction, so screen and print use the identical graph. A drawing failure does
not prevent printing calculations; an explanatory message replaces the image.

Use GitHub Pages or another HTTP(S) static host. Opening `index.html` directly
with a file URL may block workers/WASM. For local preview run
`node dev-server.cjs 8080` and open `http://127.0.0.1:8080`.
Browser storage for this local URL is separate from GitHub Pages storage.

## Tests

- `node tests/peptide-structure.test.cjs`: actual RDKit graph validation,
  residue/terminal masses, CIP stereochemistry, bridge mass deltas and ring
  counts, unsupported input rejection, worker SVG/MOL/formula generation.
- `node tests/peptide-structure-ui.test.cjs`: actual WASM rendering in Edge,
  zoom/fit/export, print parity, desktop/mobile screenshots, dark theme and
  unavailable state. Synthetic isolated browser storage is used.
- Existing chemistry and chemistry UI tests also remain applicable.

The three supplemental toolbar icons use Lucide 0.468.0 (ISC), matching the
application's existing icon format; see `vendor/lucide-LICENSE`.
