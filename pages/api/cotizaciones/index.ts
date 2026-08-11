import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {

  // GET
  if (req.method === "GET") {
    const { id, estado } = req.query
    if (id) {
      const { data, error } = await supabase.from("cotizaciones").select("*").eq("id", id).single()
      if (error) return res.status(404).json({ error: error.message })
      return res.status(200).json(data)
    }
    let q = supabase.from("cotizaciones").select("*").order("creado_en", { ascending: false })
    if (estado && estado !== "todos") q = q.eq("estado", estado)
    const { data, error } = await q.limit(500)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data || [])
  }

  // POST — crear cotización con folio automático + prefijo vendedor
  if (req.method === "POST") {
    const { _prefijo, ...body } = req.body
    // Contar cotizaciones existentes para número consecutivo global
    const { count } = await supabase
      .from("cotizaciones")
      .select("id", { count: "exact", head: true })
    const num = (count || 0) + 1
    const prefijo = _prefijo || "COT"
    const folio = `${prefijo}-${String(num).padStart(6,"0")}`
    const { data, error } = await supabase
      .from("cotizaciones").insert({ ...body, folio }).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(201).json(data)
  }

  // PATCH — actualizar
  if (req.method === "PATCH") {
    const { id } = req.query
    // Whitelist de columnas válidas en Supabase
    const COLS = ["cliente","cliente_nombre","archivo","estado","fecha_evento","fecha_entrega",
      "fecha_desmonte","lugar","tel","vendedor","partidas","articulos","total","subtotal",
      "descuento_pct","descuento_monto_global","aplica_iva","iva","notas","observaciones",
      "tipo","folio","a_cuenta","cobrado","pagos","aplica_descuento"]
    const clean: any = { actualizado_en: new Date().toISOString() }
    for (const k of COLS) {
      if (k in req.body) clean[k] = req.body[k]
    }
    const { data, error } = await supabase
      .from("cotizaciones").update(clean)
      .eq("id", id).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data)
  }

  // DELETE
  if (req.method === "DELETE") {
    const { id } = req.query
    const { error } = await supabase.from("cotizaciones").delete().eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).end()
}
