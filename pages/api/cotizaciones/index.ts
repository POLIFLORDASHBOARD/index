import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

// Extraer UUID limpio de cualquier string (quita prefijos excel_, _ etc)
function cleanUUID(raw: string): string {
  const m = raw.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i)
  return m ? m[0] : raw.trim()
}

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {

  // GET
  if (req.method === "GET") {
    const { id, estado } = req.query
    if (id) {
      const cleanId = cleanUUID(String(id))
      const { data, error } = await supabase.from("cotizaciones").select("*").eq("id", cleanId).maybeSingle()
      if (error) return res.status(404).json({ error: error.message })
      if (!data) return res.status(404).json({ error: "No encontrado" })
      return res.status(200).json(data)
    }
    let q = supabase.from("cotizaciones").select("*").order("creado_en", { ascending: false })
    if (estado && estado !== "todos") q = q.eq("estado", String(estado))
    const { data, error } = await q.limit(500)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data || [])
  }

  // POST — crear cotización con folio automático + prefijo vendedor
  if (req.method === "POST") {
    // Quitar todos los campos internos del frontend (prefijos _ o fromExcel)
    const { _prefijo, _fromExcel, fromExcel, id, ...bodyRaw } = req.body
    // Quitar cualquier campo que empiece con _ o no sea columna válida
    const COT_COLS = ["cliente","cliente_nombre","archivo","estado","fecha_evento","fecha_entrega",
      "fecha_desmonte","lugar","lugar_evento","tel","cliente_tel","vendedor","partidas","articulos","total","subtotal",
      "descuento_pct","descuento_monto_global","aplica_iva","iva","notas","observaciones",
      "tipo","a_cuenta","cobrado","pagos","aplica_descuento"]
    const body: any = {}
    for (const k of COT_COLS) {
      if (k in bodyRaw) body[k] = bodyRaw[k]
    }
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
    const id = cleanUUID(String(req.query.id || ""))
    if (!id) return res.status(400).json({ error: "ID requerido" })

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
      .eq("id", id).select()
    if (error) return res.status(500).json({ error: error.message })
    if (!data || data.length === 0) return res.status(404).json({ error: "Cotización no encontrada con id: " + id })
    return res.status(200).json(data[0])
  }

  // DELETE
  if (req.method === "DELETE") {
    const id = cleanUUID(String(req.query.id || ""))
    const { error } = await supabase.from("cotizaciones").delete().eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).end()
}
