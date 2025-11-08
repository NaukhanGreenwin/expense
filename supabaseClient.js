const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

// Prefer service role on server for storage admin ops; fall back to anon for read-only
const supabaseKey = SUPABASE_SERVICE_ROLE_KEY || SUPABASE_ANON_KEY;

let supabase = null;
if (SUPABASE_URL && supabaseKey) {
  supabase = createClient(SUPABASE_URL, supabaseKey, {
    auth: { persistSession: false },
  });
}

module.exports = { supabase };

