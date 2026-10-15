const SUPABASE_URL = "https://zecbwtcljgouogguhdiq.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_lrjcvQVf8X0dscjtdtAjmA_LWhYcWLw";

const supabaseClient = window.supabase.createClient(
    SUPABASE_URL,
    SUPABASE_PUBLISHABLE_KEY
);