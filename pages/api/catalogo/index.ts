import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {

  // GET — listar artículos con filtros
  if (req.method === "GET") {
    const { categoria, busq, activo, id } = req.query

    if (id) {
      const { data, error } = await supabase
        .from("catalogo_articulos").select("*").eq("id", id).single()
      if (error) return res.status(404).json({ error: error.message })
      return res.status(200).json(data)
    }

    let q = supabase.from("catalogo_articulos").select("*").order("nombre")
    if (categoria && categoria !== "todos") q = q.eq("categoria", categoria)
    if (activo !== undefined) q = q.eq("activo", activo === "true")
    if (busq) q = q.ilike("nombre", `%${busq}%`)

    const { data, error } = await q.limit(500)
    if (error) return res.status(500).json({ error: error.message })

    // Re-sort: exact match first, starts-with second, contains third
    if (busq && data) {
      const b = String(busq).toLowerCase().trim()
      const sorted = [...data].sort((a: any, b2: any) => {
        const na = (a.nombre || "").toLowerCase()
        const nb = (b2.nombre || "").toLowerCase()
        const aExact = na === b ? 0 : na.startsWith(b) ? 1 : 2
        const bExact = nb === b ? 0 : nb.startsWith(b) ? 1 : 2
        if (aExact !== bExact) return aExact - bExact
        return na.localeCompare(nb)
      })
      return res.status(200).json(sorted)
    }

    return res.status(200).json(data || [])
  }

  // POST — crear artículo
  if (req.method === "POST") {
    const { data, error } = await supabase
      .from("catalogo_articulos").insert(req.body).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(201).json(data)
  }

  // PATCH — actualizar artículo
  if (req.method === "PATCH") {
    const { id } = req.query
    const { data, error } = await supabase
      .from("catalogo_articulos").update(req.body).eq("id", id).select().single()
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json(data)
  }

  // DELETE — desactivar artículo (soft delete)
  if (req.method === "DELETE") {
    const { id } = req.query
    const { error } = await supabase
      .from("catalogo_articulos").update({ activo: false }).eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).end()
}
