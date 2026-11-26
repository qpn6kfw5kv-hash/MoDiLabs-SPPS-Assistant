// Public browser configuration only. Keep cloud disabled until SQL and Auth are configured.
// Never put a service-role key, database password or OpenAI key here.
window.SPPS_CONFIG = Object.freeze({
  cloudEnabled: false,
  supabaseUrl: '',
  supabasePublicKey: ''
});
