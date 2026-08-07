import { NextApiRequest, NextApiResponse } from 'next'
import { requireAuth, TokenPayload } from '../../../lib/auth'
import { supabaseAdmin } from '../../../lib/supabase'

async function handler(req: NextApiRequest, res: NextApiResponse, user: TokenPayload) {
  if (req.method !== 'GET') return res.status(405).end()

  const { semana } = req.query // offset de semana: 0=esta, 1=próxima, etc.
  const offset = parseInt(semana as string || '0')

  // Calcular rango de la semana
  const hoy = new Date()
  hoy.setHours(0, 0, 0, 0)
  const dow = hoy.getDay()
  const diffLunes = dow === 0 ? -6 : 1 - dow
  const lunes = new Date(hoy)
  lunes.setDate(lunes.getDate() + diffLunes + offset * 7)
  const domingo = new Date(lunes)
  domingo.setDate(domingo.getDate() + 6)
  domingo.setHours(23, 59, 59, 999)

  const desde = lunes.toISOString().split('T')[0]
  const hasta = domingo.toISOString().split('T')[0]

  // Contratos de la semana
  const { data: contratos } = await supabaseAdmin
    .from('contratos')
    .select(`
      id, cliente, lugar, fecha_evento, fecha_entrega, fecha_desmonte,
      estado_entrega, estado_desmonte,
      articulos (nombre, cantidad, unidad, categoria)
    `)
    .gte('fecha_evento', desde)
    .lte('fecha_evento', hasta)
    .order('fecha_evento')

  if (!contratos) return res.status(500).json({ error: 'Error consultando contratos' })

  // Consolidar artículos
  const consolidado: Record<string, { nombre: string, cantidad: number, unidad: string, contratos: string[] }> = {}

  for (const c of contratos) {
    for (const art of (c.articulos as any[] || [])) {
      const key = art.nombre.toLowerCase().trim()
      if (!consolidado[key]) {
        consolidado[key] = { nombre: art.nombre, cantidad: 0, unidad: art.unidad, contratos: [] }
      }
      consolidado[key].cantidad += art.cantidad
      if (!consolidado[key].contratos.includes(c.cliente)) {
        consolidado[key].contratos.push(c.cliente)
      }
    }
  }

  // Ordenar por cantidad descendente
  const articulosOrdenados = Object.values(consolidado)
    .sort((a, b) => b.cantidad - a.cantidad)

  // Stats de la semana
  const stats = {
    total_eventos: contratos.length,
    entregas_pend: contratos.filter((c: any) => c.estado_entrega === 'pend').length,
    entregas_listo: contratos.filter((c: any) => c.estado_entrega === 'listo').length,
    desmontes_pend: contratos.filter((c: any) => c.estado_desmonte === 'pend').length,
    desmontes_listo: contratos.filter((c: any) => c.estado_desmonte === 'recog').length,
    total_piezas: articulosOrdenados.reduce((s, a) => s + a.cantidad, 0),
    articulo_top: articulosOrdenados[0]?.nombre || '—',
  }

  return res.status(200).json({
    semana: { desde, hasta, offset },
    stats,
    contratos,
    articulos: articulosOrdenados,
  })
}

export default requireAuth(handler)
