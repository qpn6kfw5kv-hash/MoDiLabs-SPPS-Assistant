/* RDKit owns graph validation, stereochemical depiction and 2D coordinates. */
importScripts('vendor/rdkit/dist/RDKit_minimal.js');
let engine;
self.onmessage=async({data})=>{
  let mol;
  try{
    if(!engine)engine=initRDKitModule({locateFile:file=>new URL('vendor/rdkit/dist/'+file,self.location.href).href});
    const rdkit=await engine;
    mol=rdkit.get_mol(data.smiles);
    if(!mol||!mol.is_valid())throw new Error('La struttura non supera la validazione chimica.');
    mol.set_new_coords(true);
    const width=Math.min(2400,Math.max(900,data.count*100)),height=data.bridges?720:500;
    const svg=mol.get_svg_with_highlights(JSON.stringify({width,height,bondLineWidth:1.2,addStereoAnnotation:false,baseFontSize:.6,padding:.06}));
    const descriptors=JSON.parse(mol.get_descriptors());
    // Formula from RDKit's explicit graph plus implicit hydrogen counts.
    const json=JSON.parse(mol.get_json()),molecule=json.molecules[0],defaults=json.defaults.atom;
    const elements={1:'H',6:'C',7:'N',8:'O',16:'S'},counts={};
    for(const atom of molecule.atoms){const z=atom.z??defaults.z,symbol=elements[z];if(!symbol)throw new Error('Elemento non supportato.');counts[symbol]=(counts[symbol]||0)+1;counts.H=(counts.H||0)+(atom.impHs??defaults.impHs??0);}
    const formula=['C','H','N','O','S'].filter(e=>counts[e]).map(e=>e+(counts[e]===1?'':counts[e])).join('');
    self.postMessage({id:data.id,result:{svg,width,height,formula,molblock:mol.get_molblock(),smiles:mol.get_smiles(),mw:descriptors.amw}});
  }catch(e){engine=undefined;self.postMessage({id:data.id,error:e.message||'Motore 2D non disponibile.'});}
  finally{mol?.delete();}
};
