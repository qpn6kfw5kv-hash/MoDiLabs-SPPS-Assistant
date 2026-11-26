const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),context=vm.createContext({console});
for(const file of ['amino-acids.js','chemistry.js','peptide-structure.js'])vm.runInContext(fs.readFileSync(path.join(root,file),'utf8'),context,{filename:file});
const build=(sequence,resinType='Wang')=>vm.runInContext(`PeptideStructure.build(${JSON.stringify({sequence,resinType})})`,context);
(async()=>{
 const rdkit=await require('../vendor/rdkit/dist/RDKit_minimal.js')();
 function inspect(sequence,resin='Wang'){
  const model=build(sequence,resin),mol=rdkit.get_mol(model.smiles);assert(mol,model.smiles);assert(mol.is_valid());
  try{return {mw:JSON.parse(mol.get_descriptors()).amw,stereo:JSON.parse(mol.get_stereo_tags()),json:JSON.parse(mol.get_json()),smiles:mol.get_smiles(),rings:JSON.parse(mol.get_descriptors()).NumRings};}finally{mol.delete();}
 }
 for(const token of 'GAVLIPFWMSTCYNQDEKHR')assert(inspect(token).mw>0);
 for(const [token,expected] of [['A',['(S)']],['C',['(R)']],['T',['(S)','(R)']],['I',['(S)','(S)']],['P',['(S)']]])assert.deepEqual(inspect(token).stereo.CIP_atoms.map(a=>a[1]),expected);
 for(const token of 'GAVLIPFWMSTCYNQDEKHR'){
  const expected=vm.runInContext(`calculatePeptideMW(['${token}'],'acid')`,context);
  assert(Math.abs(inspect(token).mw-expected)<.03,token+' residue mass mismatch');
 }
 assert(Math.abs(inspect('AG').mw-146.146)<.005);
 assert(Math.abs(inspect('AG','Rink Amide').mw-145.162)<.005);
 assert(Math.abs(inspect('Aib').mw-103.121)<.005);
 assert(Math.abs(inspect('Nle').mw-131.175)<.005);
 for(const [linear,cyclic,delta] of [['CAGC','*CAGC*',2.016],['KAGD','*KAGD*',18.015],['DAGK','*DAGK*',18.015],['KAGE','*KAGE*',18.015]]){
  assert(Math.abs(inspect(linear).mw-inspect(cyclic).mw-delta)<.005);assert.equal(inspect(cyclic).rings,inspect(linear).rings+1);
 }
 assert.throws(()=>build('A(Unknown)G'),/non definita/);
 assert.throws(()=>build('*KAG*'),/non definita/);
 assert.throws(()=>build('*CAG'),/non definita/);
 assert.throws(()=>build('AG','Unknown'),/resina/);
 assert.notEqual(inspect('(Fmoc-D-Ala-OH)').smiles,inspect('A').smiles);
 // Validate the actual worker and its graph-derived formula, not just the model.
 const responses=[],worker=vm.createContext({console,URL,importScripts(){},initRDKitModule:()=>Promise.resolve(rdkit),self:{location:{href:'http://localhost/peptide-structure-worker.js'},postMessage:data=>responses.push(data)}});
 vm.runInContext(fs.readFileSync(path.join(root,'peptide-structure-worker.js'),'utf8'),worker);
 await worker.self.onmessage({data:{id:1,...build('*KAGD*')}});
 assert(!responses[0].error,responses[0].error);assert.equal(responses[0].result.formula,'C15H25N5O6');
 assert(responses[0].result.svg.includes('<svg'));assert(responses[0].result.molblock.includes('V2000'));
 console.log('PASS: 22 residue graphs, terminal chemistry, disulfides, both lactam orientations, D stereochemistry, unknown rejection and real RDKit worker rendering.');
})().catch(e=>{console.error(e);process.exitCode=1;});
