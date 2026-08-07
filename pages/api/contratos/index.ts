import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const SUPABASE_URL = "https://ohxehnsxfbvdflmqlzxq.supabase.co"
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"

function getSupabase() {
  return createClient(SUPABASE_URL, SUPABASE_KEY)
}

export default async function handler(req: NextApiRequest, res: NextApiResponse): Promise<any> {
  const supabase = getSupabase()

  // GET — listar TODOS los contratos con artículos (paginando de a 1000)
  if (req.method === "GET") {
    // Búsqueda rápida de clientes (sin artículos para ser más rápido)
    const { busq_cliente } = req.query
    if (busq_cliente) {
      const { data, error } = await supabase
        .from("contratos")
        .select("cliente,tel,telefono,lugar")
        .ilike("cliente", `%${busq_cliente}%`)
        .order("cliente")
        .limit(20)
      if (error) return res.status(500).json({ error: error.message })
      return res.status(200).json(data || [])
    }

    let allData: any[] = []
    let from = 0
    const pageSize = 1000
    while (true) {
      const { data, error } = await supabase
        .from("contratos")
        .select("*, articulos(nombre,cantidad,pu,importe,seccion)")
        .order("fecha_evento", { ascending: true })
        .range(from, from + pageSize - 1)
      if (error) return res.status(500).json({ error: error.message })
      if (!data || data.length === 0) break
      allData = allData.concat(data)
      if (data.length < pageSize) break
      from += pageSize
    }
    const result = allData.map((c: any) => ({
      ...c,
      tel: c.telefono || c.tel || "",
      articulos: c.articulos || []
    }))
    return res.status(200).json(result)
  }

  // POST — insertar nuevo contrato
  if (req.method === "POST") {
    const { articulos, ...contratoData } = req.body
    // Map tel -> telefono
    const row = {
      archivo: contratoData.archivo,
      cliente: contratoData.cliente,
      lugar: contratoData.lugar,
      telefono: contratoData.tel || contratoData.telefono || "",
      fecha_evento: contratoData.fecha_evento,
      fecha_entrega: contratoData.fecha_entrega,
      fecha_desmonte: contratoData.fecha_desmonte,
      dia_evento: contratoData.dia_evento,
      dia_entrega: contratoData.dia_entrega,
      dia_desmonte: contratoData.dia_desmonte,
      estado_entrega: contratoData.estado_entrega || "pend",
      estado_desmonte: contratoData.estado_desmonte || "pend",
      asig_entrega: contratoData.asig_entrega || [],
      asig_desmonte: contratoData.asig_desmonte || [],
      checklist: contratoData.checklist || [],
      notas: contratoData.notas || "",
      es_duplicado: contratoData.es_duplicado || false,
      carpeta: contratoData.carpeta || "",
      vendedor: contratoData.vendedor || "",
      tipo: contratoData.tipo || "contrato",
      folio: contratoData.folio || "",
      total: contratoData.total || 0,
      a_cuenta: contratoData.a_cuenta || 0,
      cobrado: contratoData.cobrado || contratoData.a_cuenta || 0,
      pagos: contratoData.pagos || [],
    }
    const { data: newC, error } = await supabase
      .from("contratos")
      .insert(row)
      .select()
      .single()
    if (error) return res.status(500).json({ error: error.message })
    // Insert articulos
    if (articulos?.length && newC?.id) {
      await supabase.from("articulos").insert(
        articulos.map((a: any) => ({
          contrato_id: newC.id,
          nombre: a.nombre,
          cantidad: a.cantidad,
          unidad: "pza",
          categoria: a.seccion || "General",
          pu: a.pu || 0,
          importe: a.importe || 0,
          seccion: a.seccion || "General"
        }))
      )
    }
    return res.status(201).json({ ...newC, articulos: articulos || [] })
  }

  // PATCH — actualizar contrato
  if (req.method === "PATCH") {
    const { id } = req.query
    const { articulos, vendedor, tel, ...rest } = req.body

    // Only send known contrato columns to Supabase (NOT articulos — separate table)
    const ALLOWED = ["cliente","lugar","telefono","fecha_evento","fecha_entrega","fecha_desmonte",
      "dia_evento","dia_entrega","dia_desmonte","tipo","folio","total","a_cuenta","cobrado",
      "asig_entrega","asig_desmonte","estado_entrega","estado_desmonte","checklist","notas",
      "pagos","es_duplicado","carpeta","archivo",
      "descuento_pct","descuento_monto_global","aplica_iva"]
    const updates: any = {}
    ALLOWED.forEach(k => { if (rest[k] !== undefined) updates[k] = rest[k] })
    // vendedor and tel handled explicitly
    if (vendedor !== undefined) updates.vendedor = vendedor
    if (tel !== undefined) updates.telefono = tel
    if (rest.telefono !== undefined) updates.telefono = rest.telefono

    const { data, error } = await supabase
      .from("contratos")
      .update(updates)
      .eq("id", id)
      .select()
      .single()
    if (error) return res.status(500).json({ error: error.message })

    // Update articulos — delete old, insert new
    if (Array.isArray(articulos)) {
      const { error: delErr } = await supabase
        .from("articulos")
        .delete()
        .eq("contrato_id", id)
      if (delErr) console.error("Error deleting articulos:", delErr.message)

      if (articulos.length > 0) {
        const { error: insErr } = await supabase
          .from("articulos")
          .insert(
            articulos.map((a: any) => ({
              contrato_id: id,
              nombre: a.nombre || "",
              cantidad: a.cantidad || 0,
              unidad: "pza",
              categoria: a.seccion || "General",
              pu: a.pu || 0,
              importe: a.importe || 0,
              seccion: a.seccion || "General"
            }))
          )
        if (insErr) console.error("Error inserting articulos:", insErr.message)
      }
    }

    return res.status(200).json({ ...data, articulos: articulos || [], tel: data?.telefono || "" })
  }

  // DELETE
  if (req.method === "DELETE") {
    const { id } = req.query
    await supabase.from("articulos").delete().eq("contrato_id", id)
    const { error } = await supabase.from("contratos").delete().eq("id", id)
    if (error) return res.status(500).json({ error: error.message })
    return res.status(200).json({ ok: true })
  }

  return res.status(405).end()
}
