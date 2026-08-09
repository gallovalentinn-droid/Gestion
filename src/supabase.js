import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !key) {
  // Falta configurar las variables de entorno — ver README.md, paso 3.
  console.error(
    "Faltan VITE_SUPABASE_URL y/o VITE_SUPABASE_ANON_KEY. " +
    "En desarrollo local, creá un archivo .env con esos valores. " +
    "En Vercel, cargalos en Project Settings → Environment Variables."
  );
}

export const supabase = createClient(url, key);
