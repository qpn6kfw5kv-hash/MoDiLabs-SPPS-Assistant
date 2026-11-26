// Identity-preserving chemical records. MW is average molar mass, not exact MS mass.
const CHEMISTRY_REVIEW_DATE = '2026-10-06';
const CHEMISTRY_CATALOG = (() => {
  const standard = {
    G:['29022-11-5','C17H15NO4','47627'], A:['35661-39-3','C18H17NO4','531480'],
    V:['68858-20-8','C20H21NO4','47638'], L:['35661-60-0','C21H23NO4','47633'],
    I:['71989-23-6','C21H23NO4','47628'], P:['71989-31-6','C20H19NO4','47636'],
    F:['35661-40-6','C24H21NO4','338338'], W:['143824-78-6','C31H30N2O6','47561'],
    M:['71989-28-1','C20H21NO4S','47634'], S:['71989-33-8','C22H25NO5','47619'],
    T:['71989-35-0','C23H27NO5','47622'], C:['103213-32-7','C37H31NO4S','47695'],
    Y:['71989-38-3','C28H29NO5',null], H:['109425-51-6','C40H33N3O4','47639'],
    D:['71989-14-5','C23H25NO6','47618'], E:['71989-18-9','C24H27NO6',null],
    N:['132388-59-1','C38H32N2O5','47672'], Q:['132327-80-1','C39H34N2O5',null],
    K:['71989-26-9','C26H32N2O6',null], R:['154445-77-9','C34H40N4O7S','47349']
  };
  const records = Object.entries(AMINO_ACIDS).map(([code,aa]) => {
    const [cas,formula,product] = standard[code];
    const source = code === 'Y' ? 'https://www.sigmaaldrich.com/US/en/product/mm/852020' : code === 'E' ? 'https://www.sigmaaldrich.com/US/en/product/mm/852009' : code === 'Q' ? 'https://shop.bachem.com/product/4014949/' : code === 'K' ? 'https://shop.bachem.com/product/4003151/' :
      'https://www.sigmaaldrich.com/US/en/product/aldrich/'+product;
    return {...aa,code,cas,formula,phase:'solid',source,reviewed:CHEMISTRY_REVIEW_DATE,
      aliases:[aa.code3,aa.name,aa.fmocName,aa.fmocName.replace('Fmoc-','Fmoc-L-')]};
  });
  records.push(
    {code3:'Aib',name:'Aminoisobutyric acid',fmocName:'Fmoc-Aib-OH',fmocMW:325.36,residueMW:85.10,formula:'C19H19NO4',cas:'94744-50-0',protection:null,source:'https://shop.bachem.com/product/4013932/',aliases:['Aib','alpha-aminoisobutyric acid','2-aminoisobutyric acid','Fmoc-Aib-OH']},
    {code3:'Nle',name:'Norleucine',fmocName:'Fmoc-Nle-OH',fmocMW:353.41,residueMW:113.16,formula:'C21H23NO4',cas:'77284-32-3',protection:null,source:'https://www.sigmaaldrich.com/US/en/product/mm/852014',aliases:['Nle','Norleucine','Fmoc-Nle-OH','Fmoc-L-Nle-OH']},
    {...CYCLIZATION_PROTECTED_AMINO_ACIDS.K,residueMW:128.17,cas:'146982-27-6',formula:'C25H28N2O6',source:'https://www.sigmaaldrich.com/US/en/product/aldrich/47583',aliases:['Lys(Alloc)','Fmoc-Lys(Alloc)-OH','Fmoc-L-Lys(Alloc)-OH']},
    {...CYCLIZATION_PROTECTED_AMINO_ACIDS.D,residueMW:115.09,cas:'146982-24-3',formula:'C22H21NO6',source:'https://www.sigmaaldrich.com/US/en/substance/fmocaspoalloh39541146982243',aliases:['Asp(OAll)','Fmoc-Asp(OAll)-OH','Fmoc-L-Asp(OAll)-OH']}
  );
  return records.map(r=>({...r,aliases:[...r.aliases,r.cas],phase:'solid',reviewed:CHEMISTRY_REVIEW_DATE}));
})();

function chemicalIdentityKey(value) {
  return String(value||'').normalize('NFKC').toLowerCase().replace(/[\s\-\u2010-\u2015]/g,'');
}
function resolveAminoAcid(token,synthesis,forwardIndex) {
  const key=chemicalIdentityKey(token);
  let aa=AMINO_ACIDS[token] ? CHEMISTRY_CATALOG.find(r=>r.code===token) :
    CHEMISTRY_CATALOG.find(r=>r.aliases.some(a=>chemicalIdentityKey(a)===key));
  const override=synthesis && getCyclizationProtectedOverrides(synthesis.sequence)[forwardIndex];
  if(override)aa=CHEMISTRY_CATALOG.find(r=>r.fmocName===override.fmocName);
  if(aa)return {...aa,status:'verified'};
  const saved=synthesis?.chemicalRecords?.[token];
  if(saved && ['verified','manual'].includes(saved.status) && Number(saved.fmocMW)>0)return {...saved,code3:token};
  return {code3:token,fmocName:/fmoc|fluoren/i.test(token)?token:'Fmoc-'+token+'-OH',fmocMW:NaN,status:'unresolved'};
}
function readKnownResidueToken(seq,index) {
  // Uppercase strings remain one-letter sequences; mixed-case aliases are words.
  const tail=seq.slice(index);
  const aliases=CHEMISTRY_CATALOG.flatMap(r=>r.aliases).filter(a=>!a.includes(' ')&&!a.includes('(')).sort((a,b)=>b.length-a.length);
  for(const alias of aliases){
    const part=tail.slice(0,alias.length);
    if(part.toLowerCase()!==alias.toLowerCase() || (!/[a-z]/.test(part) && !/^fmoc/i.test(part)))continue;
    return {token:part,nextIndex:index+part.length};
  }
  return null;
}

function parseMolecularFormula(formula){
  if(!/^(?:[A-Z][a-z]?\d*)+$/.test(String(formula)))return null;
  const result={};for(const match of formula.matchAll(/([A-Z][a-z]?)(\d*)/g))result[match[1]]=(result[match[1]]||0)+Number(match[2]||1);
  return result;
}
function deriveResidueMW(formula,name){
  const counts=parseMolecularFormula(formula);
  if(!counts)return NaN;
  const groups={boc:{C:5,H:8,O:2},trt:{C:19,H:14},tbu:{C:4,H:8},otbu:{C:4,H:8},pbf:{C:13,H:16,O:3,S:1},alloc:{C:4,H:4,O:2},oall:{C:3,H:4}};
  const removals=[{C:15,H:10,O:2},{H:2,O:1}];
  for(const match of String(name).matchAll(/\(([^()]*)\)/g)){
    const group=groups[match[1].toLowerCase()];if(!group)return NaN;removals.push(group);
  }
  for(const group of removals)for(const [atom,n] of Object.entries(group)){counts[atom]=(counts[atom]||0)-n;if(counts[atom]<0)return NaN;}
  const weights={C:12.011,H:1.008,N:14.007,O:15.999,S:32.06,F:18.998403163,Cl:35.45,P:30.973761998};
  if(Object.keys(counts).some(atom=>!weights[atom]))return NaN;
  return Object.entries(counts).reduce((sum,[atom,n])=>sum+n*weights[atom],0);
}

const COUPLING_REAGENTS = [
  {name:'HBTU',mw:379.24,defaultEq:2.9,cas:'94790-37-1',formula:'C11H16F6N5OP',product:'12804'},
  {name:'HATU',mw:380.23,defaultEq:2.9,cas:'148893-10-1',formula:'C10H15F6N6OP',product:'11373'},
  {name:'PyBOP',mw:520.39,defaultEq:2.9,cas:'128625-52-5',formula:'C18H28F6N6OP2',product:'377848'},
  {name:'PyAOP',mw:521.38,defaultEq:3,cas:'156311-83-0',formula:'C17H27F6N7OP2',product:'535303'},
  {name:'HCTU',mw:413.69,defaultEq:2.9,cas:'330645-87-9',formula:'C11H15ClF6N5OP',source:'https://www.merckmillipore.com/TW/en/product/HCTU,MDA_CHEM-851012'},
  {name:'DIC',mw:126.20,defaultEq:3,cas:'693-13-0',formula:'C7H14N2',phase:'liquid',density:0.814,densityTemperature:20,source:'https://www.sigmaaldrich.com/DE/de/product/mm/803649'},
  {name:'Oxyma',mw:142.11,defaultEq:3,cas:'3849-21-6',formula:'C5H6N2O3',source:'https://b2b.sigmaaldrich.com/US/en/product/mm/851086'},
  {name:'HOBt',displayName:'HOBt (anidro)',mw:135.13,defaultEq:3,cas:'2592-95-2',formula:'C6H5N3O',source:'https://pubchem.ncbi.nlm.nih.gov/compound/1-Hydroxybenzotriazole'},
  {name:'HOBt monoidrato',mw:153.15,defaultEq:3,cas:'80029-43-2',formula:'C6H7N3O2',source:'https://www.tcichemicals.com/US/en/p/H0468'},
  {name:'DIPEA',mw:129.24,defaultEq:6,cas:'7087-68-5',formula:'C8H19N',phase:'liquid',density:0.742,densityTemperature:25,product:'D125806'}
].map(r=>({phase:'solid',...r,source:r.source||'https://www.sigmaaldrich.com/US/en/product/aldrich/'+r.product,reviewed:CHEMISTRY_REVIEW_DATE}));

function getReagent(name){return COUPLING_REAGENTS.find(r=>r.name.toLowerCase()===String(name).toLowerCase() || (r.name==='DIPEA' && String(name).toUpperCase()==='DIEA'));}
function expandLegacyReagent(name) {
  const direct=getReagent(name);
  if(name==='DIC/Oxyma')return [{name:'DIC',eq:3},{name:'Oxyma',eq:3}];
  if(name==='DIC/HOBt')return [{name:'DIC',eq:3},{name:'HOBt',eq:3}];
  if(!direct)return [];
  const rows=[{name:direct.name,eq:direct.defaultEq}];
  if(['HBTU','HATU','HCTU','PyBOP','PyAOP'].includes(name))rows.push({name:'DIPEA',eq:6});
  return rows;
}
function getCouplingReagents(s,index) {
  if(Array.isArray(s.couplingReagents)){
    const selection=index!==undefined?s.aaReagents?.[index]:null;
    return s.couplingReagents.filter(r=>!selection || selection.includes(r.name)).filter(r=>getReagent(r.name)).map(r=>({...getReagent(r.name),name:getReagent(r.name).name,eq:r.eq,concentrationM:r.concentrationM||null,densityOverride:r.densityOverride||null}));
  }
  const names=index===undefined?(s.activators||[s.activator||'HBTU']):[s.aaActivators?.[index]||(s.activators||[])[0]||s.activator||'HBTU'];
  const rows=names.flatMap(expandLegacyReagent);
  return rows.filter((r,i)=>rows.findIndex(x=>x.name===r.name)===i).map(r=>({...getReagent(r.name),...r}));
}
function calculateReagentDose(reagent,scale) {
  const eq=Number(reagent.eq??reagent.defaultEq),n=Number(scale)*eq;
  if(!Number.isFinite(n)||n<=0||!(reagent.mw>0))return {valid:false};
  const massMg=n*reagent.mw;
  const concentration=Number(reagent.concentrationM);
  const density=Number(reagent.densityOverride||reagent.density);
  if((reagent.concentrationM && (!(concentration>0)||!Number.isFinite(concentration))) || (reagent.densityOverride && (!(density>0)||!Number.isFinite(density))))return {valid:false};
  const volumeUL=concentration>0?n/concentration*1000:reagent.phase==='liquid' && density>0?massMg/density:null;
  return {valid:true,massMg,volumeUL,eq,concentrationM:concentration||null,density:concentration>0?null:density||null};
}
function formatReagentDose(reagent,scale) {
  const dose=calculateReagentDose(reagent,scale);
  if(!dose.valid)return 'Dati mancanti';
  return dose.volumeUL!==null?dose.volumeUL.toFixed(2)+' µL':dose.massMg.toFixed(2)+' mg';
}
function formatPeptideMW(value){return Number.isFinite(value)?value.toFixed(2)+' Da':'Da verificare';}
