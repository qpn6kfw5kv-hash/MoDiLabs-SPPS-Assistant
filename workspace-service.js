const SPPSWorkspaces = (() => {
  async function list(userId) {
    const client = await SPPSBackend.getClient();
    const { data, error } = await client.from('workspace_members')
      .select('workspace_id,role,workspaces(id,name,owner_id,version)')
      .eq('user_id', userId).is('revoked_at', null);
    if (error) throw error;
    return (data || []).filter(row => row.workspaces).map(row => ({ ...row.workspaces, role: row.role }));
  }
  async function members(workspaceId) {
    const client = await SPPSBackend.getClient();
    const result = await client.from('workspace_members').select('user_id,role,created_at')
      .eq('workspace_id', workspaceId).is('revoked_at', null).order('created_at');
    if (result.error) throw result.error;
    const ids = (result.data || []).map(row => row.user_id);
    if (!ids.length) return [];
    const profiles = await client.from('profiles').select('id,display_name').in('id', ids);
    if (profiles.error) throw profiles.error;
    return result.data.map(row => ({ ...row,
      display_name: profiles.data.find(p => p.id === row.user_id)?.display_name || 'Membro del laboratorio' }));
  }
  async function create(name, requestId) {
    const client = await SPPSBackend.getClient();
    const { data, error } = await client.rpc('spps_create_workspace', { p_name: name, p_request_id: requestId });
    if (error) throw error;
    return data;
  }
  function last(userId) { try { return localStorage.getItem('spps-cloud-last-workspace:' + userId); } catch { return null; } }
  function remember(userId, workspaceId) {
    try { localStorage.setItem('spps-cloud-last-workspace:' + userId, workspaceId); } catch { /* Optional preference. */ }
  }
  return Object.freeze({ list, members, create, last, remember });
})();
