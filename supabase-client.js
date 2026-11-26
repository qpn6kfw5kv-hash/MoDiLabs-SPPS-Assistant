/* The SDK is loaded only after cloud mode is explicitly enabled. */
const SPPSBackend = (() => {
  const SDK_URL = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.117.2/+esm';
  let clientPromise;
  function configured() { return window.SPPS_CONFIG?.cloudEnabled === true; }
  function configuration() {
    const config = window.SPPS_CONFIG || {};
    if (!configured()) throw new Error('cloud_disabled');
    let url;
    try { url = new URL(config.supabaseUrl); } catch { throw new Error('cloud_configuration'); }
    const local = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
    if ((url.protocol !== 'https:' && !(local && url.protocol === 'http:')) || url.username || url.password || url.search || url.hash) {
      throw new Error('cloud_configuration');
    }
    const key = String(config.supabasePublicKey || '');
    if (!key.startsWith('sb_publishable_')) {
      try {
        const payload = JSON.parse(atob(key.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')));
        if (payload.role !== 'anon') throw new Error();
      } catch { throw new Error('cloud_public_key_required'); }
    }
    return { url: url.origin, key };
  }
  async function getClient() {
    if (!clientPromise) {
      clientPromise = (async () => {
        const { url, key } = configuration();
        const { createClient } = await import(SDK_URL);
        return createClient(url, key, {
          auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false,
            storageKey: 'spps-auth-' + new URL(url).host, flowType: 'pkce' }
        });
      })().catch(error => { clientPromise = null; throw error; });
    }
    return clientPromise;
  }
  return Object.freeze({ configured, getClient });
})();
