// ⚡ Supabase 前端設定檔 (供 GitHub Pages 前端使用)
// Project URL 與公開 anon public key 是安全公開於前端的

const SUPABASE_URL = "https://ioxlptzwpmczsxboggct.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_By6SrH7lHFwdOCgw8srvUg_swLAM3cz";

if (typeof window !== 'undefined') {
  window._SUPABASE_CONFIG = {
    url: SUPABASE_URL,
    anonKey: SUPABASE_ANON_KEY
  };
  window.SUPABASE_CONFIG = window._SUPABASE_CONFIG;
  if (window.supabase && typeof window.supabase.createClient === 'function') {
    window._supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  }
}

if (typeof exports !== 'undefined') {
  exports.SUPABASE_URL = SUPABASE_URL;
  exports.SUPABASE_ANON_KEY = SUPABASE_ANON_KEY;
}
