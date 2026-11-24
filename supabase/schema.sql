-- Prevent-IA · Esquema SQL para Supabase (PostgreSQL)
-- Ejecuta este script en el SQL Editor de tu proyecto Supabase.
-- Activa RLS sin políticas: solo el servidor (SERVICE_ROLE_KEY) accede a los datos.

CREATE TABLE IF NOT EXISTS documents (
  id BIGINT PRIMARY KEY,
  title TEXT NOT NULL,
  law TEXT DEFAULT '',
  cat TEXT NOT NULL DEFAULT 'doc',
  area TEXT NOT NULL DEFAULT 'all',
  group_name TEXT,
  nw BOOLEAN DEFAULT FALSE,
  section_order JSONB NOT NULL DEFAULT '[]'::jsonb,
  deleted BOOLEAN NOT NULL DEFAULT FALSE,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS document_sections (
  id TEXT PRIMARY KEY,
  doc_id BIGINT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  heading TEXT NOT NULL,
  body TEXT NOT NULL DEFAULT '',
  deleted BOOLEAN NOT NULL DEFAULT FALSE,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_document_sections_doc_id ON document_sections(doc_id);

CREATE TABLE IF NOT EXISTS section_translations (
  cache_key TEXT PRIMARY KEY,
  sid TEXT NOT NULL REFERENCES document_sections(id) ON DELETE CASCADE,
  lang TEXT NOT NULL,
  src_hash BIGINT NOT NULL,
  heading TEXT NOT NULL,
  body TEXT NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_section_translations_lang ON section_translations(lang);
CREATE UNIQUE INDEX IF NOT EXISTS uq_section_translations_sid_lang ON section_translations(sid, lang);

CREATE TABLE IF NOT EXISTS ui_translations (
  lang TEXT NOT NULL,
  source_text TEXT NOT NULL,
  translated_text TEXT NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  PRIMARY KEY (lang, source_text)
);

CREATE TABLE IF NOT EXISTS quizzes (
  doc_id BIGINT PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE,
  questions JSONB NOT NULL DEFAULT '[]'::jsonb,
  doc_hash BIGINT,
  ai_generated BOOLEAN NOT NULL DEFAULT FALSE,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS workers (
  username_key TEXT PRIMARY KEY,
  full_name TEXT NOT NULL,
  cargo TEXT DEFAULT '',
  area TEXT NOT NULL DEFAULT 'alb',
  lang TEXT NOT NULL DEFAULT 'es',
  progress_pct INTEGER NOT NULL DEFAULT 0,
  avatar_url TEXT DEFAULT '',
  assigned_items JSONB,
  deleted BOOLEAN NOT NULL DEFAULT FALSE,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS worker_progress (
  username_key TEXT PRIMARY KEY,
  started_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  items_progress JSONB NOT NULL DEFAULT '{}'::jsonb,
  signatures JSONB NOT NULL DEFAULT '{}'::jsonb,
  campaigns_progress JSONB NOT NULL DEFAULT '{}'::jsonb,
  profile_data JSONB NOT NULL DEFAULT '{}'::jsonb,
  avatar_url TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS findings (
  id BIGSERIAL PRIMARY KEY,
  reporter_key TEXT DEFAULT '',
  type_idx INTEGER NOT NULL DEFAULT 0,
  area TEXT NOT NULL,
  description TEXT NOT NULL,
  status INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE documents ENABLE ROW LEVEL SECURITY;
ALTER TABLE document_sections ENABLE ROW LEVEL SECURITY;
ALTER TABLE section_translations ENABLE ROW LEVEL SECURITY;
ALTER TABLE ui_translations ENABLE ROW LEVEL SECURITY;
ALTER TABLE quizzes ENABLE ROW LEVEL SECURITY;
ALTER TABLE workers ENABLE ROW LEVEL SECURITY;
ALTER TABLE worker_progress ENABLE ROW LEVEL SECURITY;
ALTER TABLE findings ENABLE ROW LEVEL SECURITY;
