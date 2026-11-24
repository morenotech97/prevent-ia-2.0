# Prevent-IA

Plataforma multilenguaje de prevención de riesgos: documentos en su formato original por idioma,
evaluación, firma y constancia.

## Estructura
- `server.mjs` — servidor Express (API). En Vercel se ejecuta como función.
- `public/` — la página (HTML, CSS, JS, imágenes). En Vercel se publica desde el CDN.
- `supabase/schema.sql` — esquema de la base de datos.

## Desarrollo local
1. `npm install`
2. Copia `.env.example` a `.env.local` y completa `SUPABASE_URL` y `SUPABASE_SERVICE_ROLE_KEY`.
3. `npm run dev` y abre http://localhost:3000

## Vercel
1. Importa el proyecto (Vercel detecta Express solo; no hay comando de build).
2. En *Settings > Environment Variables* agrega `SUPABASE_URL` y `SUPABASE_SERVICE_ROLE_KEY`.
3. Los archivos se suben directo del navegador a Supabase Storage (bucket privado `documentos`),
   por lo que no aplica el límite de 4,5 MB por petición de Vercel.
