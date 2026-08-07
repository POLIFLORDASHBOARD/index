import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const SUPABASE_URL = "https://ohxehnsxfbvdflmqlzxq.supabase.co"
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"

async function parseExcel(buffer: ArrayBuffer, fileName: string, fileId: string, modified: string) {
  const XLSX = require("xlsx")
  const data = new Uint8Array(buffer)
  const wb = XLSX.read(data, { type: "array", cellDates: true })
  let wsName = wb.SheetNames.find((s: string) => s.trim().toUpperCase() === "EVENTO")
  if (!wsName) wsName = wb.SheetNames[0]
  const ws = wb.Sheets[wsName]
  if (!ws) return null
  let fe: Date | null = null
  for (const addr of ["C3","B3","D3","C2","B2"]) {
    const cell = ws[addr]
    if (cell && cell.v) {
      const d = typeof cell.v === "number" ? new Date(Math.round((cell.v-25569)*86400*1000)) : new Date(cell.v)
      if (!isNaN(d.getTime()) && d.getFullYear() > 2000) { fe = d; break }
    }
  }
  if (!fe) return null
  let cliente = ""
  const allKeys = Object.keys(ws).filter((k: string) => !k.startsWith("!"))
  for (const addr of allKeys) {
    const cell = ws[addr]
    if (cell && cell.v && String(cell.v).trim().toLowerCase().replace(/\s/g,"").startsWith("cliente")) {
      const m = addr.match(/^([A-Z]+)(\d+)$/)
      if (m) {
        const rc = ws[String.fromCharCode(m[1].charCodeAt(0)+1)+m[2]]
        if (rc && rc.v && String(rc.v).trim().length > 1) { cliente = String(rc.v).trim(); break }
      }
    }
  }
  if (!cliente) cliente = fileName.replace(/_dividido/g,"").replace(".xlsx","")
  const lugar = ws["C4"] && ws["C4"].v ? String(ws["C4"].v).trim() : ""
  const tel = ws["C6"] && ws["C6"].v ? String(ws["C6"].v).trim() : ""
  const pad = (n: number) => String(n).padStart(2,"0")
  const isoDate = (d: Date) => d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())
  const DIAS = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]
  const ent = new Date(fe); ent.setDate(ent.getDate()-1)
  const des = new Date(fe); des.setDate(des.getDate()+1)
  return {
    archivo: fileName, cliente, lugar, telefono: tel,
    fecha_evento: isoDate(fe), fecha_entrega: isoDate(ent), fecha_desmonte: isoDate(des),
    dia_evento: DIAS[fe.getDay()], dia_entrega: DIAS[ent.getDay()], dia_desmonte: DIAS[des.getDay()],
    onedrive_id: fileId, onedrive_modified: modified
  }
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") return res.status(405).end()
  const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)
  const { data: config } = await supabase.from("onedrive_config").select("*").eq("activa", true).limit(1).single()
  if (!config) return res.status(400).json({ error: "OneDrive no conectado" })

  const token = config.access_token
  const results = { nuevos: 0, actualizados: 0, errores: 0, archivos: [] as string[] }

  try {
    const sharedRes = await fetch(
      "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe?$select=id,name,remoteItem",
      { headers: { Authorization: "Bearer " + token } }
    )
    const sharedData = await sharedRes.json()
    const items = sharedData.value || []

    for (const item of items) {
      const driveId = item.remoteItem && item.remoteItem.parentReference ? item.remoteItem.parentReference.driveId : null
      const itemId = item.remoteItem ? item.remoteItem.id : item.id
      if (!driveId || !itemId) continue

      const filesRes = await fetch(
        "https://graph.microsoft.com/v1.0/drives/" + driveId + "/items/" + itemId + "/children?$select=id,name,lastModifiedDateTime",
        { headers: { Authorization: "Bearer " + token } }
      )
      const filesData = await filesRes.json()
      const files = (filesData.value || []).filter((f: any) =>
        f.name && f.name.toLowerCase().endsWith(".xlsx") && !f.name.startsWith("~")
      )

      for (const file of files) {
        try {
          const { data: existing } = await supabase.from("contratos")
            .select("id,onedrive_modified").eq("onedrive_id", file.id).single()
          if (existing && existing.onedrive_modified === file.lastModifiedDateTime) continue

          const dlRes = await fetch(
            "https://graph.microsoft.com/v1.0/drives/" + driveId + "/items/" + file.id + "/content",
            { headers: { Authorization: "Bearer " + token } }
          )
          const buffer = await dlRes.arrayBuffer()
          const contrato = await parseExcel(buffer, file.name, file.id, file.lastModifiedDateTime)
          if (!contrato) { results.errores++; continue }

          if (existing) {
            await supabase.from("contratos").update({ ...contrato, carpeta: item.name }).eq("id", existing.id)
            results.actualizados++
          } else {
            await supabase.from("contratos").insert({ ...contrato, carpeta: item.name })
            results.nuevos++
          }
          results.archivos.push(file.name)
        } catch(e) {
          results.errores++
        }
      }
    }

    await supabase.from("onedrive_config").update({ ultima_sync: new Date().toISOString() }).eq("id", config.id)
    return res.status(200).json({
      ok: true, ...results,
      mensaje: "Sync completado: " + results.nuevos + " nuevos, " + results.actualizados + " actualizados, " + results.errores + " errores"
    })
  } catch(e: any) {
    return res.status(500).json({ error: e.message })
  }
}
