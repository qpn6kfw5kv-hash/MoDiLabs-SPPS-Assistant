# Chemical data and calculation policy

Review date: 2026-10-06. Runtime catalog: `chemistry.js`.

## Identity and molar masses

The catalog contains 20 standard protected building blocks and Aib, Nle,
Lys(Alloc), and Asp(OAll). Each record includes its formula, CAS number and
manufacturer reference. Molar masses are average molar masses in g/mol,
not monoisotopic masses. Display rounding does not change the dose calculation.

Aliases are matched without case or separator differences. Stereochemistry,
protecting groups and positional numbers are not discarded. Keyword-based
inventory search is not evidence for a molecular weight.

Uncatalogued compounds are queried through PubChem using their full protected
name. A unique CID, matching synonym and compatible formula are required.
There is no fallback to the free amino acid. Failed or ambiguous lookups show
an unresolved value, not an estimated protected mass. A manually confirmed
molar mass is recorded as manual. Previous unverified cached numbers require
confirmation. The protected reagent MW and peptide residue MW are separate;
unsupported residue chemistry leaves peptide MW unresolved.

Examples of primary references:

- Fmoc-Aib-OH, 325.36: https://shop.bachem.com/product/4013932/
- Fmoc-Nle-OH, 353.41: https://www.sigmaaldrich.com/GI/en/product/mm/852014
- Fmoc-Lys(Alloc)-OH, 452.50: https://www.sigmaaldrich.com/US/en/product/aldrich/47583
- Fmoc-Asp(OAll)-OH, 395.41: https://www.sigmaaldrich.com/US/en/substance/fmocaspoalloh39541146982243

## Individual coupling reagents

HBTU, HATU, PyBOP, PyAOP, HCTU, DIC, Oxyma, HOBt (anhydrous),
HOBt monohydrate and DIPEA are separate selections. Equivalents are editable.
New selections do not silently add a base or additive. Existing legacy recipes
retain their original components and per-step alternatives.

- DIC: 126.20 g/mol; reference density 0.814 g/mL at 20 C. This is the midpoint
  of the manufacturer range 0.813-0.815, not an exact density for every bottle.
  https://www.sigmaaldrich.com/DE/de/product/mm/803649
- DIPEA: 129.24 g/mol; density 0.742 g/mL at 25 C.
  https://www.sigmaaldrich.com/US/en/product/sial/d125806
- Oxyma: 142.11 g/mol, solid.
  https://b2b.sigmaaldrich.com/US/en/product/mm/851086
- HOBt monohydrate: 153.15 g/mol.
  https://www.tcichemicals.com/US/en/p/H0468
- PyAOP: 521.38 g/mol.
  https://b2b.sigmaaldrich.com/AT/en/product/aldrich/535303

Generic wet HOBt is not necessarily stoichiometric monohydrate. Use the exact
product specification and lot certificate for hydration, purity and assay.
The app does not automatically correct weights for unspecified purity or water.
Density depends on temperature; the optional density override takes precedence
and is labelled manual. Solutions must use their actual molarity, not the
density of the neat reagent.

## Units

Scale is mmol. For each coupling:

```
amount (mmol) = scale * equivalents
mass (mg) = amount * molar mass
neat liquid volume (uL) = mass (mg) / density (g/mL)
solution volume (uL) = amount (mmol) / concentration (mol/L) * 1000
```

An amino acid repeated in a sequence is printed once at its single-coupling
dose, without multiplying by occurrence count. Different reagent recipes are
listed separately with their sequence positions. Preloaded Wang residues do
not receive an additional coupling dose.

## Verification

Run `node tests/chemistry.test.cjs` and `node tests/chemistry-ui.test.cjs`.
The latter uses isolated browser storage and mocked PubChem responses; it does
not prove network availability or guarantee identification of every compound.
Catalog formulas, protected variants, aliases, invalid values, liquid/solution
units, legacy compatibility, persistence and bench/print parity are covered.
