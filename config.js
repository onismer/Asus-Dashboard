// ============================================================
// CONFIG — paste your Supabase project credentials here.
// Supabase Dashboard > Project Settings > API
// (anon/public key is safe to expose in a browser app;
//  data is protected by Row Level Security + login)
// ============================================================
const CONFIG = {
  SUPABASE_URL: "https://hefibrlvlxxeraqfujwl.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_SURvh16qB3puPZ3wkmCzuw_GphJNMno",
  APP_TITLE: "ASUS Store Maintenance Dashboard",
  // Storage bucket for backing up original uploaded files (optional).
  // Create a private bucket with this name in Supabase Storage, or leave
  // as-is — upload still works, backup step is skipped gracefully.
  BACKUP_BUCKET: "raw-uploads",
  // Rows per upsert batch during upload
  UPSERT_CHUNK: 500,
};
