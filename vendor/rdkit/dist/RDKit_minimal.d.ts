// TypeScript bindings for emscripten-generated code.  Automatically generated at compile time.
interface WasmModule {
}

type EmbindString = ArrayBuffer|Uint8Array|Uint8ClampedArray|Int8Array|string;
export interface ClassHandle {
  isAliasOf(other: ClassHandle): boolean;
  delete(): void;
  deleteLater(): this;
  isDeleted(): boolean;
  // @ts-ignore - If targeting lower than ESNext, this symbol might not exist.
  [Symbol.dispose](): void;
  clone(): this;
}
export interface StringList extends ClassHandle, Iterable<string> {
  push_back(_0: EmbindString): void;
  resize(_0: number, _1: EmbindString): void;
  size(): number;
  get(_0: number): string | undefined;
  set(_0: number, _1: EmbindString): boolean;
}

export interface JSMolListList extends ClassHandle, Iterable<MolList | null> {
  size(): number;
  get(_0: number): MolList | undefined;
  push_back(_0: MolList | null): void;
  resize(_0: number, _1: MolList | null): void;
  set(_0: number, _1: MolList | null): boolean;
}

export interface Mol extends ClassHandle {
  is_valid(): boolean;
  has_coords(): number;
  get_smiles(): string;
  get_smiles(details: EmbindString): string;
  get_cxsmiles(): string;
  get_cxsmiles(details: EmbindString): string;
  get_smarts(): string;
  get_smarts(details: EmbindString): string;
  get_cxsmarts(): string;
  get_cxsmarts(details: EmbindString): string;
  get_molblock(): string;
  get_molblock(details: EmbindString): string;
  get_v3Kmolblock(): string;
  get_v3Kmolblock(details: EmbindString): string;
  get_v2Kmolblock(): string;
  get_v2Kmolblock(details: EmbindString): string;
  get_as_uint8array(details: EmbindString): any;
  get_as_uint8array(): any;
  get_inchi(options: EmbindString): string;
  get_inchi(): string;
  get_json(): string;
  get_svg(): string;
  get_svg(width: number, height: number): string;
  get_svg_with_highlights(details: EmbindString): string;
  combine_with(other: Mol): string;
  combine_with(other: Mol, details: EmbindString): string;
  draw_to_canvas_with_offset(canvas: any, offsetx: number, offsety: number, width: number, height: number): string;
  draw_to_canvas(canvas: any, width: number, height: number): string;
  draw_to_canvas_with_highlights(canvas: any, details: EmbindString): string;
  generate_aligned_coords(templateMol: Mol, param: any): string;
  get_morgan_fp_as_uint8array(): any;
  get_morgan_fp_as_uint8array(details: EmbindString): any;
  get_pattern_fp(param: any): string;
  get_pattern_fp_as_uint8array(): any;
  get_pattern_fp_as_uint8array(param: any): any;
  get_topological_torsion_fp_as_uint8array(): any;
  get_topological_torsion_fp_as_uint8array(details: EmbindString): any;
  get_rdkit_fp_as_uint8array(): any;
  get_rdkit_fp_as_uint8array(details: EmbindString): any;
  get_atom_pair_fp_as_uint8array(): any;
  get_atom_pair_fp_as_uint8array(details: EmbindString): any;
  get_maccs_fp_as_uint8array(): any;
  get_frags(details: EmbindString): any;
  get_frags(): any;
  add_to_png_blob(pngString: EmbindString, details: EmbindString): any;
  add_to_png_blob(pngString: EmbindString): any;
  get_coords(): any;
  get_substruct_match(mol: Mol): string;
  get_substruct_match(mol: Mol, details: EmbindString): string;
  get_substruct_matches(mol: Mol): string;
  get_substruct_matches(mol: Mol, details: EmbindString): string;
  get_descriptors(): string;
  get_morgan_fp(): string;
  get_morgan_fp(details: EmbindString): string;
  get_pattern_fp(): string;
  get_topological_torsion_fp(): string;
  get_topological_torsion_fp(details: EmbindString): string;
  get_rdkit_fp(): string;
  get_rdkit_fp(details: EmbindString): string;
  get_atom_pair_fp(): string;
  get_atom_pair_fp(details: EmbindString): string;
  get_maccs_fp(): string;
  get_stereo_tags(): string;
  get_aromatic_form(): string;
  convert_to_aromatic_form(): void;
  get_kekule_form(): string;
  convert_to_kekule_form(): void;
  set_new_coords(): boolean;
  get_new_coords(): string;
  set_new_coords(useCoordGen: boolean): boolean;
  get_new_coords(useCoordGen: boolean): string;
  has_prop(key: EmbindString): boolean;
  get_prop_list(includePrivate: boolean, includeComputed: boolean): StringList;
  get_prop_list(includePrivate: boolean): StringList;
  get_prop_list(): StringList;
  set_prop(key: EmbindString, val: EmbindString, computed: boolean): boolean;
  set_prop(key: EmbindString, val: EmbindString): boolean;
  get_prop(key: EmbindString): string;
  clear_prop(key: EmbindString): boolean;
  condense_abbreviations(): string;
  condense_abbreviations(maxCoverage: number, useLinkers: boolean): string;
  add_hs(): string;
  add_hs_in_place(): boolean;
  remove_hs(details: EmbindString): string;
  remove_hs(): string;
  remove_hs_in_place(details: EmbindString): boolean;
  remove_hs_in_place(): boolean;
  normalize_depiction(): number;
  normalize_depiction(canonicalize: number): number;
  normalize_depiction(canonicalize: number, scaleFactor: number): number;
  straighten_depiction(): void;
  straighten_depiction(minimizeRotation: boolean): void;
  get_num_atoms(heavyOnly: boolean): number;
  get_num_atoms(): number;
  get_num_bonds(): number;
  copy(): Mol | null;
  get_mmpa_frags(minCuts: number, maxCuts: number, maxCutBonds: number): any;
}

export interface MolList extends ClassHandle {
  append(mol: Mol): number;
  insert(idx: number, mol: Mol): number;
  at(idx: number): Mol | null;
  pop(idx: number): Mol | null;
  next(): Mol | null;
  reset(): void;
  at_end(): boolean;
  size(): number;
}

export interface Reaction extends ClassHandle {
  run_reactants(reactants: MolList, maxProducts: number): JSMolListList;
  draw_to_canvas_with_offset(canvas: any, offsetx: number, offsety: number, width: number, height: number): string;
  draw_to_canvas(canvas: any, width: number, height: number): string;
  draw_to_canvas_with_highlights(canvas: any, details: EmbindString): string;
  get_svg(): string;
  get_svg(width: number, height: number): string;
  get_svg_with_highlights(details: EmbindString): string;
}

export interface SubstructLibrary extends ClassHandle {
  add_mol(m: Mol): number;
  add_smiles(smi: EmbindString): number;
  add_trusted_smiles(smi: EmbindString): number;
  get_trusted_smiles(i: number): string;
  add_trusted_smiles_and_pattern_fp(smi: EmbindString, patternFpAsUInt8Array: any): number;
  get_pattern_fp_as_uint8array(i: number): any;
  get_matches_as_uint32array(q: Mol, useChirality: boolean, numThreads: number, maxResults: number): any;
  get_matches_as_uint32array(q: Mol, maxResults: number): any;
  get_matches_as_uint32array(q: Mol): any;
  get_mol(i: number): Mol | null;
  get_matches(q: Mol, useChirality: boolean, numThreads: number, maxResults: number): string;
  get_matches(q: Mol, maxResults: number): string;
  get_matches(q: Mol): string;
  count_matches(q: Mol, useChirality: boolean, numThreads: number): number;
  count_matches(q: Mol, useChirality: boolean): number;
  count_matches(q: Mol): number;
  size(): number;
}

export interface Log extends ClassHandle {
  get_buffer(): string;
  clear_buffer(): void;
}

export interface RGroupDecomposition extends ClassHandle {
  add(mol: Mol): number;
  process(): boolean;
  get_rgroups_as_columns(): any;
  get_rgroups_as_rows(): any;
}

interface EmbindModule {
  StringList: {
    new(): StringList;
  };
  JSMolListList: {
    new(): JSMolListList;
  };
  Mol: {};
  MolList: {
    new(): MolList;
  };
  Reaction: {};
  SubstructLibrary: {
    new(): SubstructLibrary;
    new(_0: number): SubstructLibrary;
  };
  Log: {};
  version(): string;
  prefer_coordgen(prefer: boolean): void;
  use_legacy_stereo_perception(value: boolean): boolean;
  allow_non_tetrahedral_chirality(value: boolean): boolean;
  get_inchikey_for_inchi(input: EmbindString): string;
  get_mol(input: EmbindString, details_json: EmbindString): Mol | null;
  get_mol(input: EmbindString): Mol | null;
  get_mol_from_uint8array(pklAsUInt8Array: any): Mol | null;
  get_mol_copy(other: Mol): Mol | null;
  get_qmol(input: EmbindString): Mol | null;
  enable_logging(logName: EmbindString): boolean;
  enable_logging(): void;
  disable_logging(logName: EmbindString): boolean;
  disable_logging(): void;
  set_log_capture(log_name: EmbindString): Log | null;
  set_log_tee(log_name: EmbindString): Log | null;
  get_rxn(input: EmbindString, details_json: EmbindString): Reaction | null;
  get_rxn(input: EmbindString): Reaction | null;
  get_mcs_as_json(mols: MolList, details_json: EmbindString): string;
  get_mcs_as_json(mols: MolList): string;
  get_mcs_as_mol(mols: MolList, details_json: EmbindString): Mol | null;
  get_mcs_as_mol(mols: MolList): Mol | null;
  get_mcs_as_smarts(mols: MolList, details_json: EmbindString): string;
  get_mcs_as_smarts(mols: MolList): string;
  RGroupDecomposition: {};
  get_rgd(singleOrMultipleCores: any, details_json: EmbindString): RGroupDecomposition | null;
  get_rgd(singleOrMultipleCores: any): RGroupDecomposition | null;
  molzip(a: Mol, b: Mol, details_json: EmbindString): Mol | null;
  molzip(param1: any, param2: any): Mol | null;
  molzip(rgdRow: any): Mol | null;
  get_mol_from_png_blob(pngAsUInt8Array: any, details: EmbindString): Mol | null;
  get_mol_from_png_blob(pngAsUInt8Array: any): Mol | null;
  get_mols_from_png_blob(pngAsUInt8Array: any, details: EmbindString): MolList | null;
  get_mols_from_png_blob(pngAsUInt8Array: any): MolList | null;
}

export type MainModule = WasmModule & EmbindModule;
export default function MainModuleFactory (options?: unknown): Promise<MainModule>;
