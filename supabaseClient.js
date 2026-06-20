// ════════════════════════════════════════════════════════════════
// FeeTrack UG — Supabase Connection
// ════════════════════════════════════════════════════════════════
// This is the ONLY file that should ever reference your Supabase project
// URL and key directly. Every other part of the app imports `supabase`
// from here rather than connecting separately — one connection, reused
// everywhere.
//
// The actual URL and key live in a separate ".env" file (see .env.example
// in this same folder) — NOT typed directly into this file. That .env
// file is excluded from GitHub via .gitignore, so your project's specific
// values never get uploaded publicly, even though this connection code
// does. This is standard practice, not a security requirement specific
// to FeeTrack — even though the "publishable" key itself is safe to be
// visible, keeping real values out of committed code is just good habit
// and makes it easy to use different values for testing vs. real use later.
// ════════════════════════════════════════════════════════════════

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_PUBLISHABLE_KEY = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY;

if (!SUPABASE_URL || !SUPABASE_PUBLISHABLE_KEY) {
  // Fails loudly and clearly rather than silently breaking later — if you
  // see this error, it means the .env file is missing or wasn't set up
  // (locally) or the Environment Variables weren't added in Vercel's
  // project settings (when deployed).
  throw new Error(
    "Missing Supabase connection details. Make sure VITE_SUPABASE_URL and " +
    "VITE_SUPABASE_PUBLISHABLE_KEY are set (in your local .env file, and in " +
    "Vercel's Project Settings → Environment Variables when deployed)."
  );
}

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);

