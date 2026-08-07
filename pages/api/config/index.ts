import type { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient("https://ohxehnsxfbvdflmqlzxq.supabase.co", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM")

export const config = { api: { bodyParser: { sizeLimit: "5mb" } } }

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  // GET - retrieve a config value
  if (req.method === "GET") {
    const { clave } = req.query
    if (!clave) return res.status(400).json({ error: "clave required" })
    
    const { data, error } = await supabase
      .from("configuracion")
      .select("valor")
      .eq("clave", clave)
      .single()
    
    if (error || !data) return res.status(200).json({ valor: null })
    return res.status(200).json({ valor: data.valor })
  }

  // POST - save a config value  
  if (req.method === "POST") {
    const { clave, valor } = req.body
    if (!clave) return res.status(400).json({ error: "clave required" })

    const { error } = await supabase
      .from("configuracion")
      .upsert({ clave, valor }, { onConflict: "clave" })
    
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).json({ error: "Method not allowed" })
}
