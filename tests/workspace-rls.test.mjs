import { PGlite } from '../.cloud-test-deps/pglite/dist/index.js';
import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';

const db = new PGlite();
let checks = 0;
const A='00000000-0000-4000-8000-000000000001', B='00000000-0000-4000-8000-000000000002';
const C='00000000-0000-4000-8000-000000000003', D='00000000-0000-4000-8000-000000000004';
const E='00000000-0000-4000-8000-000000000005', F='00000000-0000-4000-8000-000000000006';

// Only Supabase Auth infrastructure is stubbed. Tables, constraints, functions,
// role privileges and RLS below run on the real embedded PostgreSQL engine.
await db.exec(`create role anon nologin; create role authenticated nologin;
 create schema auth; create table auth.users(id uuid primary key, raw_user_meta_data jsonb default '{}');
 create function auth.uid() returns uuid language sql stable as
 $$ select nullif(current_setting('request.jwt.claim.sub', true), '')::uuid $$;
 grant usage on schema auth to authenticated, anon;
 grant execute on function auth.uid() to authenticated, anon;`);
const sql = await readFile(new URL('../supabase/migrations/202609280001_workspace_foundation.sql', import.meta.url),'utf8');
await db.exec(sql);
for(const [id,name] of [[A,'Owner'],[B,'Editor'],[C,'Viewer'],[D,'Outsider'],[E,'Admin'],[F,'Other admin']]) {
  await db.query('insert into auth.users(id,raw_user_meta_data) values ($1,$2)',[id,JSON.stringify({display_name:name,role:'owner'})]);
}
async function as(user,query,args=[],role='authenticated') {
  await db.exec('begin');
  try {
    await db.exec(`set local role ${role}`);
    await db.query("select set_config('request.jwt.claim.sub',$1,true)",[user||'']);
    const result=await db.query(query,args);
    await db.exec('commit');
    return result.rows;
  } catch(error) { await db.exec('rollback'); throw error; }
}
async function denied(user,query,args=[],role) {
  await assert.rejects(()=>as(user,query,args,role),e=>e.code==='42501');checks++;
}
function equal(actual,expected){assert.deepEqual(actual,expected);checks++;}
try {
  const request='10000000-0000-4000-8000-000000000001';
  const [{id:w}]=await as(A,'select public.spps_create_workspace($1,$2) as id',['Lab A',request]);
  equal((await as(A,'select public.spps_create_workspace($1,$2) as id',['Retry',request]))[0].id,w);
  equal((await as(A,'select count(*)::int as count from public.workspace_members'))[0].count,1);
  equal((await as(B,'select * from public.workspaces')),[]);
  // Synthetic provisioning only; production phase 1 deliberately has no add-member RPC.
  for(const [id,role] of [[B,'editor'],[C,'viewer'],[E,'admin'],[F,'admin']]) {
    await db.query('insert into public.workspace_members(workspace_id,user_id,role) values ($1,$2,$3)',[w,id,role]);
  }
  const [{id:other}]=await as(D,'select public.spps_create_workspace($1,$2) as id',['Lab B',request]);
  equal((await as(C,'select id from public.workspaces')).map(r=>r.id),[w]);
  equal(await as(B,'select * from public.workspaces where id=$1',[other]),[]);
  equal(await as(D,'select * from public.workspace_members where workspace_id=$1',[w]),[]);
  equal(await as(D,'select * from public.profiles where id=$1',[A]),[]);
  equal((await as(C,'select id from public.profiles')).length,5);
  for(const roleUser of [A,B,C,D,E]) {
    await denied(roleUser,'insert into public.workspaces(name,owner_id,created_by,creation_key) values ($1,$2,$2,$3)',['Bypass',roleUser,request]);
    await denied(roleUser,'update public.workspace_members set role=\'owner\' where workspace_id=$1 and user_id=$2',[w,roleUser]);
    await denied(roleUser,'delete from public.workspace_members where workspace_id=$1',[w]);
    await denied(roleUser,'update public.profiles set display_name=\'Bypass\' where id=$1',[A]);
  }
  await denied(null,'select * from public.workspaces',[],'anon');
  await denied(null,'select public.spps_create_workspace($1,$2)',['Unauthorized',request],'anon');
  await denied(null,'select public.spps_create_workspace($1,$2)',['No uid',request]);
  for(const id of [B,C,D])await denied(id,'select public.spps_rename_workspace($1,$2,$3)',[w,'Bad name',1]);
  equal((await as(E,'select public.spps_rename_workspace($1,$2,$3) as version',[w,'Renamed',1]))[0].version,2);
  await assert.rejects(()=>as(A,'select public.spps_rename_workspace($1,$2,$3)',[w,'Stale',1]),e=>e.code==='40001');checks++;
  await as(C,'select public.spps_update_profile($1)',['Viewer renamed']);
  equal((await as(A,'select display_name from public.profiles where id=$1',[C]))[0].display_name,'Viewer renamed');
  for(const id of [B,C,D,E])await denied(id,'select public.spps_change_member($1,$2,$3,false)',[w,A,'viewer']);
  await denied(E,'select public.spps_change_member($1,$2,$3,false)',[w,B,'admin']);
  await denied(E,'select public.spps_change_member($1,$2,$3,true)',[w,F,null]);
  await denied(A,'select public.spps_change_member($1,$2,$3,true)',[w,A,null]);
  await as(E,'select public.spps_change_member($1,$2,$3,false)',[w,B,'viewer']);
  equal((await as(B,'select role from public.workspace_members where user_id=$1',[B]))[0].role,'viewer');
  await denied(E,'select public.spps_transfer_workspace($1,$2)',[w,E]);
  await assert.rejects(()=>as(A,'select public.spps_transfer_workspace($1,$2)',[w,D]),e=>e.code==='22023');checks++;
  await as(A,'select public.spps_transfer_workspace($1,$2)',[w,E]);
  equal((await as(E,'select owner_id from public.workspaces where id=$1',[w]))[0].owner_id,E);
  equal((await as(A,'select role from public.workspace_members where user_id=$1',[A]))[0].role,'admin');
  await as(E,'select public.spps_change_member($1,$2,$3,true)',[w,B,null]);
  equal(await as(B,'select * from public.workspaces where id=$1',[w]),[]);
  equal(await as(B,'select * from public.profiles where id=$1',[E]),[]);
  await denied(B,'select public.spps_rename_workspace($1,$2,$3)',[w,'Revoked',3]);
  await assert.rejects(()=>db.query('delete from public.workspace_members where workspace_id=$1 and user_id=$2',[w,E]),e=>e.code==='23514');checks++;
  const policies=await db.query("select relname,relrowsecurity from pg_class where relname in ('profiles','workspaces','workspace_members')");
  equal(policies.rows.every(r=>r.relrowsecurity),true);
  const publicDefiners=await db.query("select proname from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='public' and proname like 'spps_%' and prosecdef");
  equal(publicDefiners.rows.length,0);
  console.log(JSON.stringify({checks,engine:'PGlite PostgreSQL',result:'PASS',limitations:'Auth JWT context simulated; not a live Supabase Auth/REST test.'},null,2));
} finally { await db.close(); }
