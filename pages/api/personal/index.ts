import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {
  if (req.method === "GET") {
    const { data, error } = await supabase.from("personal").select("*").eq("activo", true).order("creado_en")
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data)
  }
  if (req.method === "POST") {
    const { nombre, color_bg, color_bd, color_fg } = req.body
    if (!nombre) return res.status(400).json({ error: "Nombre requerido" })
    const { data, error } = await supabase.from("personal").insert({ nombre, color_bg, color_bd, color_fg }).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(201).json(data)
  }
  if (req.method === "DELETE") {
    const { id } = req.query
    const { error } = await supabase.from("personal").update({ activo: false }).eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }
  return res.status(405).end()
}
