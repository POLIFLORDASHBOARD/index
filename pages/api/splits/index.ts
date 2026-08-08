import type { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient("https://ohxehnsxfbvdflmqlzxq.supabase.co", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM")

export const config = { api: { bodyParser: { sizeLimit: "2mb" } } }

export default async function handler(req: NextApiRequest, res: NextApiResponse) {

  // GET — listar splits (por contrato_id o por semana)
  if (req.method === "GET") {
    const { contrato_id, fecha_desde, fecha_hasta, tipo, estado } = req.query
    let q = supabase.from("splits").select("*")
    if (contrato_id) q = q.eq("contrato_id", contrato_id)
    if (tipo) q = q.eq("tipo", tipo)
    if (estado) q = q.eq("estado", estado)
    // Filter by fecha_evento OR fecha_evento_original to catch all splits in the week
    if (fecha_desde && fecha_hasta) {
      q = q.or(`fecha_evento.gte.${fecha_desde},fecha_evento_original.gte.${fecha_desde}`)
        .or(`fecha_evento.lte.${fecha_hasta},fecha_evento_original.lte.${fecha_hasta}`)
    } else {
      if (fecha_desde) q = q.gte("fecha_evento", fecha_desde)
      if (fecha_hasta) q = q.lte("fecha_evento", fecha_hasta)
    }
    q = q.order("fecha_evento", { ascending: true }).order("tipo", { ascending: true })
    const { data, error } = await q
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data || [])
  }

  // POST — crear splits para un contrato
  if (req.method === "POST") {
    const splits = req.body
    if (!Array.isArray(splits)) return res.status(400).json({ error: "Se esperan un array de splits" })
    // Whitelist de columnas que existen en la tabla splits
    const COLS = ["contrato_id","contrato_folio","cliente","lugar","tel","fecha_evento","fecha_evento_original",
                  "fecha_preparacion","tipo","nombre","estado","notas","observaciones",
                  "responsable","articulos","mostrar_precios","mostrar_direccion","proveedor_nombre"]
    const clean = splits.map((s:any) => {
      const obj:any = {}
      COLS.forEach(k => { if (s[k] !== undefined) obj[k] = s[k] })
      return obj
    })
    // Insert individually so one failure doesn't block others
    const results: any[] = []
    const errors: any[] = []
    for (const s of clean) {
      const { data: d, error: e } = await supabase.from("splits").insert(s).select().single()
      if (e) errors.push({ tipo: s.tipo, error: e.message })
      else if (d) results.push(d)
    }
    if (errors.length > 0) console.error("Split insert errors:", errors)
    if (results.length === 0) return res.status(500).json({ error: errors[0]?.error || "No splits insertados", errors })
    return res.status(200).json(results)
  }

  // PATCH — actualizar un split (estado, notas, artículos, etc)
  if (req.method === "PATCH") {
    const { id } = req.query
    const ALLOWED = ["estado","notas","observaciones","articulos","responsable","tel",
                     "mostrar_direccion","mostrar_precios","proveedor_nombre","nombre","actualizado_en",
                     "fecha_evento","fecha_preparacion"]
    const updates: any = { actualizado_en: new Date().toISOString() }
    ALLOWED.forEach(k => { if (req.body[k] !== undefined) updates[k] = req.body[k] })
    const { data, error } = await supabase.from("splits").update(updates).eq("id", id).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data)
  }

  // DELETE — eliminar splits de un contrato (para regenerar)
  if (req.method === "DELETE") {
    const { contrato_id, id } = req.query
    if (id) {
      await supabase.from("splits").delete().eq("id", id)
    } else if (contrato_id) {
      await supabase.from("splits").delete().eq("contrato_id", contrato_id)
    }
    return res.status(200).json({ ok: true })
  }

  return res.status(405).json({ error: "Method not allowed" })
}
