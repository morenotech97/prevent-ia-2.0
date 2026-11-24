import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { createClient } from '@supabase/supabase-js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

function getSupabase() {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY;
  if (!url || !key) return null;
  return createClient(url, key, { auth: { persistSession: false } });
}

function isMissingTableError(err) {
  const code = err?.code || '';
  const msg = String(err?.message || '');
  return code === '42P01' || code === 'PGRST205' || msg.includes('does not exist') || msg.includes('schema cache');
}

app.get('/api/health', (_req, res) => {
  res.json({
    ok: true,
    supabaseConfigured: Boolean(getSupabase()),
  });
});

/* ---------------- Supabase Persistence Endpoints ---------------- */

app.get('/api/data/bootstrap', async (_req, res) => {
  const supabase = getSupabase();
  if (!supabase) {
    return res.json({
      configured: false,
      reason: 'Configura SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY en el archivo .env.local para activar la persistencia en Supabase.',
    });
  }

  try {
    const [docsR, uiR, qzR, teamR, progR, hallR, qztR] = await Promise.all([
      supabase.from('documents').select('*'),
      supabase.from('ui_translations').select('*'),
      supabase.from('quizzes').select('*'),
      supabase.from('workers').select('*'),
      supabase.from('worker_progress').select('*'),
      supabase.from('findings').select('*').order('created_at', { ascending: false }),
      supabase.from('quiz_translations').select('*'),
    ]);

    const firstErr = [docsR, uiR, qzR, teamR, progR, hallR, qztR].find((r) => r.error)?.error;
    if (firstErr) {
      if (isMissingTableError(firstErr)) {
        return res.json({
          configured: false,
          schemaMissing: true,
          reason: 'Tablas no encontradas en Supabase. Ejecuta el archivo supabase/schema.sql en el SQL Editor de tu proyecto.',
        });
      }
      throw firstErr;
    }

    const docs = {};
    for (const r of docsR.data || []) {
      docs['d' + r.id] = {
        id: Number(r.id),
        title: r.title,
        law: r.law || '',
        cat: r.cat || 'doc',
        area: r.area || 'all',
        g: r.group_name || undefined,
        ln: Array.isArray(r.langs) ? r.langs : undefined,
        nw: Boolean(r.nw),
        order: [],
        del: Boolean(r.deleted),
      };
    }

    const secs = {}; // el texto de los documentos ya no se usa: se muestran los archivos

    const ui = {};
    for (const r of uiR.data || []) {
      if (!ui[r.lang]) ui[r.lang] = {};
      ui[r.lang][r.source_text] = r.translated_text;
    }

    const qz = {};
    for (const r of qzR.data || []) {
      qz['q' + r.doc_id] = {
        id: /^\d+$/.test(r.doc_id) ? Number(r.doc_id) : r.doc_id,
        qs: Array.isArray(r.questions) ? r.questions : [],
        h: r.doc_hash != null ? Number(r.doc_hash) : undefined,
        ai: Boolean(r.ai_generated),
      };
    }

    const team = {};
    for (const r of teamR.data || []) {
      team[r.username_key] = {
        u: r.username_key,
        n: r.full_name,
        cargo: r.cargo || '',
        a: r.area || 'alb',
        l: r.lang || 'es',
        p: r.progress_pct || 0,
        img: r.avatar_url || '',
        asg: r.assigned_items || undefined,
        del: Boolean(r.deleted),
      };
    }

    const prog = {};
    for (const r of progR.data || []) {
      const cp = r.campaigns_progress || {};
      prog[r.username_key] = {
        u: r.username_key,
        ps: Number(r.started_at),
        now: Number(r.updated_at),
        i: r.items_progress || {},
        s: r.signatures || {},
        c: cp._pct ? cp._pct : cp,
        cd: cp._det || {},
        f: r.profile_data || {},
        av: r.avatar_url || '',
      };
    }

    const hall = (hallR.data || []).map((r) => ({
      id: r.id,
      ty: r.type_idx,
      a: r.area,
      d: r.description,
      s: r.status || 0,
      reporter: r.reporter_key || '',
    }));

    const qzt = {};
    for (const r of qztR.data || []) {
      qzt[r.doc_id + '_' + r.lang] = { id: Number(r.doc_id), lang: r.lang, qs: Array.isArray(r.questions) ? r.questions : [], src: Number(r.src_hash) };
    }

    res.json({
      configured: true,
      data: { docs, secs, ui, qz, team, prog, hall, qzt },
    });
  } catch (err) {
    console.error('Supabase bootstrap error:', err);
    res.status(500).json({ error: err?.message || 'Error cargando datos desde Supabase.' });
  }
});

app.put('/api/data/:collection/:key', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) {
    return res.status(503).json({
      error: 'Supabase no está configurado en el archivo .env.local.',
    });
  }

  const { collection, key } = req.params;
  const d = req.body || {};

  try {
    if (collection === 'docs') {
      const id = Number(d.id ?? String(key).replace(/^d/, ''));
      const { error } = await supabase.from('documents').upsert(
        {
          id,
          title: d.title || 'Sin título',
          law: d.law || '',
          cat: d.cat || 'doc',
          area: d.area || 'all',
          group_name: d.g || null,
          langs: Array.isArray(d.ln) ? d.ln.filter((l) => ['en', 'ht', 'pt', 'de'].includes(l)) : null,
          nw: Boolean(d.nw),
          section_order: Array.isArray(d.order) ? d.order : [],
          deleted: Boolean(d.del),
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'id' }
      );
      if (error) throw error;
    } else if (collection === 'secs') {
      const { error } = await supabase.from('document_sections').upsert(
        {
          id: d.id || key,
          doc_id: Number(d.doc),
          heading: d.h || '—',
          body: d.t || '',
          deleted: Boolean(d.del),
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'id' }
      );
      if (error) throw error;
    } else if (collection === 'ui') {
      const langCode = d.l || String(key).split('_')[0];
      const ks = Array.isArray(d.k) ? d.k : [];
      const vs = Array.isArray(d.v) ? d.v : [];
      const rows = ks
        .map((src, idx) => ({
          lang: langCode,
          source_text: String(src),
          translated_text: String(vs[idx] ?? src),
          updated_at: new Date().toISOString(),
        }))
        .filter((r) => r.source_text);
      if (rows.length) {
        const { error } = await supabase.from('ui_translations').upsert(rows, { onConflict: 'lang,source_text' });
        if (error) throw error;
      }
    } else if (collection === 'qz') {
      const docId = String(d.id ?? String(key).replace(/^q/, ''));
      const { error } = await supabase.from('quizzes').upsert(
        {
          doc_id: docId,
          questions: Array.isArray(d.qs) ? d.qs : [],
          doc_hash: d.h != null ? Number(d.h) : null,
          ai_generated: Boolean(d.ai),
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'doc_id' }
      );
      if (error) throw error;
    } else if (collection === 'qzt') {
      const { error } = await supabase.from('quiz_translations').upsert(
        {
          doc_id: Number(d.id),
          lang: String(d.lang),
          questions: Array.isArray(d.qs) ? d.qs : [],
          src_hash: Number(d.src || 0),
          engine: d.engine || null,
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'doc_id,lang' }
      );
      if (error) throw error;
    } else if (collection === 'team') {
      const { error } = await supabase.from('workers').upsert(
        {
          username_key: d.u || key,
          full_name: d.n || key,
          cargo: d.cargo || '',
          area: d.a || 'alb',
          lang: d.l || 'es',
          progress_pct: Number(d.p || 0),
          avatar_url: d.img || '',
          assigned_items: Array.isArray(d.asg) ? d.asg : null,
          deleted: Boolean(d.del),
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'username_key' }
      );
      if (error) throw error;
    } else if (collection === 'prog') {
      const { error } = await supabase.from('worker_progress').upsert(
        {
          username_key: d.u || key,
          started_at: Number(d.ps || Date.now()),
          updated_at: Number(d.now || Date.now()),
          items_progress: d.i || {},
          signatures: d.s || {},
          campaigns_progress: { _pct: d.c || {}, _det: d.cd || {} },
          profile_data: d.f || {},
          avatar_url: d.av || '',
        },
        { onConflict: 'username_key' }
      );
      if (error) throw error;
    } else {
      return res.status(400).json({ error: `Colección no soportada: ${collection}` });
    }

    res.json({ ok: true });
  } catch (err) {
    console.error(`Supabase PUT ${collection}/${key} error:`, err);
    res.status(500).json({ error: err?.message || 'Error al guardar en Supabase.' });
  }
});

app.post('/api/data/reset-progress', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) {
    return res.status(503).json({ error: 'Supabase no está configurado.' });
  }

  const keys = Array.isArray(req.body?.keys) ? req.body.keys : [];
  const now = Date.now();
  try {
    if (keys.length) {
      const rows = keys.map((k) => ({
        username_key: String(k),
        started_at: now,
        updated_at: now,
        items_progress: {},
        signatures: {},
        campaigns_progress: {},
        profile_data: {},
        avatar_url: '',
      }));
      const { error } = await supabase.from('worker_progress').upsert(rows, { onConflict: 'username_key' });
      if (error) throw error;
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('Supabase reset-progress error:', err);
    res.status(500).json({ error: err?.message || 'Error al reiniciar avances en Supabase.' });
  }
});

app.post('/api/data/findings', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) {
    return res.status(503).json({ error: 'Supabase no está configurado.' });
  }

  const { reporter, ty, a, d } = req.body || {};
  if (!a || !d) {
    return res.status(400).json({ error: 'Falta el área o descripción del hallazgo.' });
  }

  try {
    const { data, error } = await supabase
      .from('findings')
      .insert({
        reporter_key: String(reporter || ''),
        type_idx: Number(ty || 0),
        area: String(a),
        description: String(d),
        status: 0,
      })
      .select()
      .single();
    if (error) throw error;
    res.json({
      finding: {
        id: data.id,
        ty: data.type_idx,
        a: data.area,
        d: data.description,
        s: data.status,
        reporter: data.reporter_key,
      },
    });
  } catch (err) {
    console.error('Supabase create finding error:', err);
    res.status(500).json({ error: err?.message || 'Error al guardar hallazgo.' });
  }
});

app.patch('/api/data/findings/:id', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) {
    return res.status(503).json({ error: 'Supabase no está configurado.' });
  }

  const id = Number(req.params.id);
  const status = Math.min(2, Math.max(0, Number(req.body?.status ?? 0)));
  try {
    const { error } = await supabase.from('findings').update({ status }).eq('id', id);
    if (error) throw error;
    res.json({ ok: true });
  } catch (err) {
    console.error('Supabase update finding error:', err);
    res.status(500).json({ error: err?.message || 'Error al actualizar hallazgo.' });
  }
});

/* ===== Archivos originales y traducidos (Supabase Storage, bucket privado "documentos") ===== */
const FILE_BUCKET = 'documentos';
const ALLOWED = {
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};
const safeName = (n) => String(n || 'archivo').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9._-]+/g, '_').slice(0, 120);

/* Subida directa: el navegador sube el archivo a Supabase con un permiso temporal firmado por el servidor.
   Así el archivo no pasa por el servidor (en Vercel cada petición admite como máximo 4,5 MB). */
async function checkUpload(supabase, q) {
  const docId = Number(q.doc);
  const lang = String(q.lang || 'es').toLowerCase();
  const name = safeName(q.name);
  const kind = q.kind === 'original' ? 'original' : 'translation';
  if (!docId || !/^[a-z]{2,3}(-[a-z]{2})?$/.test(lang) || !q.name) return { status: 400, error: 'Faltan datos: doc, lang y name.' };
  const ext = (name.toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1];
  if (!ALLOWED[ext]) return { status: 415, error: 'Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).' };
  if (Number(q.size) > 50 * 1024 * 1024) return { status: 413, error: 'El archivo supera los 50 MB.' };
  if (kind === 'translation') {
    const { data: orig } = await supabase.from('document_files').select('file_name').eq('doc_id', docId).eq('kind', 'original').maybeSingle();
    if (!orig) return { status: 409, error: 'Primero sube el archivo original en español.' };
    const oext = (orig.file_name.toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1];
    if (oext !== ext) return { status: 415, error: `La traducción debe tener el mismo formato que el original (.${oext}).` };
  }
  return { docId, lang, name, kind, ext, mime: ALLOWED[ext], path: `${docId}/${lang}/${name}` };
}

app.post('/api/files/sign-upload', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) return res.status(503).json({ error: 'Supabase no está configurado.' });
  try {
    const c = await checkUpload(supabase, req.body || {});
    if (c.error) return res.status(c.status).json({ error: c.error });
    const { data, error } = await supabase.storage.from(FILE_BUCKET).createSignedUploadUrl(c.path, { upsert: true });
    if (error) throw error;
    res.json({ path: c.path, token: data.token, url: data.signedUrl, mime: c.mime });
  } catch (err) {
    console.error('Sign upload error:', err);
    res.status(500).json({ error: err?.message || 'No se pudo preparar la subida.' });
  }
});

app.post('/api/files/commit', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) return res.status(503).json({ error: 'Supabase no está configurado.' });
  const q = req.body || {};
  try {
    const c = await checkUpload(supabase, q);
    if (c.error) return res.status(c.status).json({ error: c.error });
    const dir = `${c.docId}/${c.lang}`;
    const { data: objs, error: lsErr } = await supabase.storage.from(FILE_BUCKET).list(dir, { search: c.name });
    if (lsErr) throw lsErr;
    const obj = (objs || []).find((o) => o.name === c.name);
    if (!obj) return res.status(404).json({ error: 'El archivo no llegó a Supabase. Inténtalo de nuevo.' });
    const status = ['draft', 'review', 'published'].includes(q.status) ? q.status : 'review';
    const engine = String(q.engine || '').slice(0, 40) || null;
    const { data: prev } = await supabase.from('document_files').select('storage_path').eq('doc_id', c.docId).eq('lang', c.lang).maybeSingle();
    if (prev?.storage_path && prev.storage_path !== c.path) await supabase.storage.from(FILE_BUCKET).remove([prev.storage_path]);
    const size = Number(obj.metadata?.size || q.size || 0);
    const { error: dbErr } = await supabase.from('document_files').upsert({
      doc_id: c.docId, lang: c.lang, kind: c.kind, file_name: c.name, storage_path: c.path, mime_type: c.mime,
      size_bytes: size, status: c.kind === 'original' ? 'published' : status, engine, updated_at: new Date().toISOString(),
    }, { onConflict: 'doc_id,lang' });
    if (dbErr) throw dbErr;
    res.json({ ok: true, path: c.path, size });
  } catch (err) {
    console.error('Commit upload error:', err);
    res.status(500).json({ error: err?.message || 'No se pudo registrar el archivo.' });
  }
});

app.get('/api/files', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) return res.status(503).json({ error: 'Supabase no está configurado.' });
  try {
    let q = supabase.from('document_files').select('*').order('doc_id').order('lang');
    if (req.query.doc) q = q.eq('doc_id', Number(req.query.doc));
    const { data, error } = await q;
    if (error) throw error;
    if (req.query.lite) return res.json({ files: (data || []).map(({ doc_id, lang, kind, status, file_name }) => ({ doc_id, lang, kind, status, file_name })) });
    const rows = data || [];
    const out = await Promise.all(rows.map(async (r) => {
      const { data: s } = await supabase.storage.from(FILE_BUCKET).createSignedUrl(r.storage_path, 3600, { download: r.file_name });
      return { ...r, url: s?.signedUrl || null };
    }));
    res.json({ files: out });
  } catch (err) {
    res.status(500).json({ error: err?.message || 'No se pudieron listar los archivos.' });
  }
});

async function listAllFiles(supabase, prefix) {
  const out = [];
  const { data, error } = await supabase.storage.from(FILE_BUCKET).list(prefix, { limit: 1000 });
  if (error) throw error;
  for (const it of data || []) {
    const p = `${prefix}/${it.name}`;
    if (it.id) out.push(p); else out.push(...await listAllFiles(supabase, p));
  }
  return out;
}

// Borra definitivamente un documento: sus archivos en Storage (todos los idiomas) y, por cascada,
// secciones, traducciones, preguntas y registros de archivos.
app.delete('/api/docs/:id', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) return res.status(503).json({ error: 'Supabase no está configurado.' });
  const id = Number(req.params.id);
  if (!id) return res.status(400).json({ error: 'Documento inválido.' });
  try {
    const paths = await listAllFiles(supabase, String(id));
    if (paths.length) {
      const { error: rmErr } = await supabase.storage.from(FILE_BUCKET).remove(paths);
      if (rmErr) throw rmErr;
    }
    const { error } = await supabase.from('documents').delete().eq('id', id);
    if (error) throw error;
    res.json({ ok: true, filesRemoved: paths.length });
  } catch (err) {
    console.error('Delete document error:', err);
    res.status(500).json({ error: err?.message || 'No se pudo borrar el documento.' });
  }
});

app.post('/api/files/status', async (req, res) => {
  const supabase = getSupabase();
  if (!supabase) return res.status(503).json({ error: 'Supabase no está configurado.' });
  const { doc, lang, status } = req.body || {};
  if (!doc || !lang || !['draft', 'review', 'published'].includes(status)) return res.status(400).json({ error: 'Datos inválidos.' });
  const { error } = await supabase.from('document_files').update({ status, updated_at: new Date().toISOString() }).eq('doc_id', Number(doc)).eq('lang', lang);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ ok: true });
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

export default app;

if (!process.env.VERCEL) app.listen(PORT, '0.0.0.0', () => {
  console.log(`Prevent-IA server listening on http://0.0.0.0:${PORT}`);
});
