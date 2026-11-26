const SPPSAuth = (() => {
  let client, subscription, user = null, revision = 0;
  const listeners = new Set();
  function emit() { listeners.forEach(fn => fn({ user, revision })); }
  function assign(next) { user = next; revision++; emit(); }
  async function start() {
    client = await SPPSBackend.getClient();
    if (!subscription) {
      subscription = client.auth.onAuthStateChange((event, session) => {
        // Never await a Supabase request inside this callback (Auth lock).
        if (event === 'SIGNED_OUT') assign(null);
        else if (session?.user && (user?.id !== session.user.id || event === 'USER_UPDATED')) {
          assign(session.user);
        }
      }).data.subscription;
    }
    const startRevision = revision;
    const result = await client.auth.getUser();
    if (startRevision !== revision) return user;
    if (result.error && result.error.name !== 'AuthSessionMissingError') throw result.error;
    assign(result.data?.user || null);
    return user;
  }
  async function sendCode(email) {
    client = await SPPSBackend.getClient();
    const { error } = await client.auth.signInWithOtp({ email, options: { shouldCreateUser: true } });
    if (error) throw error;
  }
  async function verifyCode(email, token) {
    const { data, error } = await client.auth.verifyOtp({ email, token, type: 'email' });
    if (error) throw error;
    if (!data.user) throw new Error('authentication_required');
    if (user?.id !== data.user.id) assign(data.user);
    return data.user;
  }
  async function signOut() {
    // Do not report a successful logout if session removal fails.
    if (client) {
      const { error } = await client.auth.signOut({ scope: 'local' });
      if (error) throw error;
    }
    if (user) assign(null);
  }
  return Object.freeze({ start, sendCode, verifyCode, signOut,
    current: () => user, revision: () => revision,
    subscribe(fn) { listeners.add(fn); return () => listeners.delete(fn); }
  });
})();
