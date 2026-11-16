import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {

  // GET — listar plantillas
  // Admin (esAdmin=true) ve todas; vendedor solo las suyas
  if (req.method === "GET") {
    const { vendedor, esAdmin, id } = req.query
    if (id) {
      const { data, error } = await supabase.from("plantillas_cotizacion").select("*").eq("id", id).single()
      if (error) return res.status(404).json({ error: error.message })
      return res.status(200).json(data)
    }
    let q = supabase.from("plantillas_cotizacion").select("*").order("nombre", { ascending: true })
    if (esAdmin !== "true" && vendedor) {
      q = q.eq("vendedor", vendedor)
    }
    const { data, error } = await q.limit(200)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data || [])
  }

  // POST — crear plantilla
  if (req.method === "POST") {
    const { vendedor, nombre, partidas } = req.body
    if (!vendedor || !nombre) return res.status(400).json({ error: "vendedor y nombre son requeridos" })
    const { data, error } = await supabase
      .from("plantillas_cotizacion")
      .insert({ vendedor, nombre, partidas: partidas || [] })
      .select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(201).json(data)
  }

  // PATCH — actualizar plantilla
  if (req.method === "PATCH") {
    const { id } = req.query
    if (!id) return res.status(400).json({ error: "id requerido" })
    const { data, error } = await supabase
      .from("plantillas_cotizacion")
      .update({ ...req.body, actualizado_en: new Date().toISOString() })
      .eq("id", id).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data)
  }

  // DELETE — eliminar plantilla
  if (req.method === "DELETE") {
    const { id } = req.query
    if (!id) return res.status(400).json({ error: "id requerido" })
    const { error } = await supabase.from("plantillas_cotizacion").delete().eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).end()
}
