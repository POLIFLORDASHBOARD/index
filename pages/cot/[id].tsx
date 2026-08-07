import { useEffect, useState } from "react"
import { useRouter } from "next/router"
import Head from "next/head"

interface Partida {
  nombre: string; cantidad: number; precio_unitario: number; subtotal: number; notas: string
}
interface Cotizacion {
  id: string; folio: string; cliente_nombre: string; cliente_tel: string
  lugar_evento: string; fecha_evento: string; fecha_entrega: string
  fecha_desmonte: string; fecha_vigencia: string; estado: string; vendedor: string
  subtotal: number; descuento_pct: number; descuento_monto: number
  aplica_iva: boolean; iva_monto: number; total: number
  notas_cliente: string; condiciones: string; partidas: Partida[]
}

export default function PortalCotizacion() {
  const router = useRouter()
  const { id } = router.query
  const [cot, setCot] = useState<Cotizacion | null>(null)
  const [estado, setEstado] = useState<"loading"|"ok"|"error"|"aprobando"|"aprobada"|"rechazada">("loading")
  const [nombre, setNombre] = useState("")
  const [confirmar, setConfirmar] = useState<"aprobar"|"rechazar"|null>(null)

  useEffect(() => {
    if (!id) return
    fetch(`/api/cotizaciones?id=${id}`)
      .then(r => r.json())
      .then(data => {
        if (data.error) { setEstado("error"); return }
        setCot(data)
        if (data.estado === "convertida") setEstado("aprobada")
        else if (data.estado === "rechazada") setEstado("rechazada")
        else setEstado("ok")
      })
      .catch(() => setEstado("error"))
  }, [id])

  const aprobar = async () => {
    if (!nombre.trim()) { alert("Por favor escribe tu nombre para confirmar"); return }
    setEstado("aprobando")
    await fetch(`/api/cotizaciones?id=${id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ estado: "aceptada", notas_cliente: (cot?.notas_cliente || "") + `\n[Aprobada digitalmente por: ${nombre} el ${new Date().toLocaleString("es-MX")}]` })
    })
    setEstado("aprobada")
    setConfirmar(null)
  }

  const rechazar = async () => {
    setEstado("aprobando")
    await fetch(`/api/cotizaciones?id=${id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ estado: "rechazada" })
    })
    setEstado("rechazada")
    setConfirmar(null)
  }

  const fmt = (n: number) => "$" + Math.round(n || 0).toLocaleString("es-MX")
  const fmtFecha = (f: string) => f ? new Date(f + "T12:00:00").toLocaleDateString("es-MX", { weekday: "long", day: "numeric", month: "long", year: "numeric" }) : "—"

  if (estado === "loading") return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f5f4f0", fontFamily: "Arial,sans-serif" }}>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 32, marginBottom: 12 }}>🌸</div>
        <div style={{ color: "#9a9590" }}>Cargando cotización...</div>
      </div>
    </div>
  )

  if (estado === "error") return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f5f4f0", fontFamily: "Arial,sans-serif" }}>
      <div style={{ textAlign: "center", padding: 32 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>😕</div>
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>Cotización no encontrada</div>
        <div style={{ color: "#9a9590" }}>El link puede haber expirado o ser inválido.</div>
        <div style={{ marginTop: 16, color: "#9a9590", fontSize: 13 }}>Contacta a Poliflor para más información.</div>
      </div>
    </div>
  )

  if (estado === "aprobada") return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f5f4f0", fontFamily: "Arial,sans-serif" }}>
      <div style={{ textAlign: "center", padding: 40, background: "#fff", borderRadius: 20, maxWidth: 400, boxShadow: "0 8px 40px rgba(0,0,0,.1)" }}>
        <div style={{ fontSize: 56, marginBottom: 16 }}>🎉</div>
        <div style={{ fontSize: 22, fontWeight: 800, color: "#2d6a4f", marginBottom: 8 }}>¡Cotización aprobada!</div>
        <div style={{ fontSize: 14, color: "#9a9590", marginBottom: 20, lineHeight: 1.6 }}>
          Hemos recibido tu confirmación para el evento del <strong>{fmtFecha(cot?.fecha_evento || "")}</strong>.<br/>
          Nuestro equipo se pondrá en contacto contigo pronto.
        </div>
        <div style={{ background: "#edf7f2", borderRadius: 10, padding: "12px 16px", fontSize: 13, color: "#2d6a4f", fontWeight: 600 }}>
          🌸 ¡Gracias por confiar en Poliflor!
        </div>
      </div>
    </div>
  )

  if (estado === "rechazada") return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f5f4f0", fontFamily: "Arial,sans-serif" }}>
      <div style={{ textAlign: "center", padding: 40, background: "#fff", borderRadius: 20, maxWidth: 400, boxShadow: "0 8px 40px rgba(0,0,0,.1)" }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>💭</div>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Cotización no aceptada</div>
        <div style={{ fontSize: 13, color: "#9a9590" }}>Si cambiaste de opinión o necesitas ajustes, contáctanos por WhatsApp.</div>
      </div>
    </div>
  )

  if (!cot) return null
  const vigDias = cot.fecha_vigencia ? Math.ceil((new Date(cot.fecha_vigencia + "T23:59:59").getTime() - Date.now()) / 86400000) : null

  return (
    <>
      <Head>
        <title>Cotización {cot.folio} — Poliflor</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
      </Head>
      <div style={{ minHeight: "100vh", background: "#f0ece4", fontFamily: "Inter,Arial,sans-serif", color: "#1a1814" }}>
        {/* Top bar */}
        <div style={{ background: "#1a1814", padding: "10px 20px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <img src={typeof window!=="undefined"?(localStorage.getItem("pf_logo")||"/logo.png"):"/logo.png"} alt="Poliflor" style={{ height: 32, filter: "brightness(0) invert(1)", opacity: .9 }} onError={(e: any) => e.target.style.display = "none"}/>
            <span style={{ color: "rgba(255,255,255,.4)", fontSize: 12 }}>Cotización digital</span>
          </div>
          <span style={{ fontFamily: "monospace", fontSize: 12, color: "rgba(255,255,255,.6)", background: "rgba(255,255,255,.08)", padding: "3px 10px", borderRadius: 6 }}>{cot.folio}</span>
        </div>

        <div style={{ maxWidth: 720, margin: "0 auto", padding: "24px 16px 80px" }}>
          {/* Vigencia alert */}
          {vigDias !== null && vigDias >= 0 && vigDias <= 5 && (
            <div style={{ background: "#fdf5e8", border: "1px solid #e8d4b8", borderRadius: 10, padding: "10px 14px", marginBottom: 16, fontSize: 13, color: "#92580a", display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 18 }}>⏰</span>
              <span>Esta cotización <strong>vence en {vigDias} día{vigDias !== 1 ? "s" : ""}</strong>. Apruébala antes de que expire.</span>
            </div>
          )}

          {/* COTIZACIÓN CARD */}
          <div style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 4px 32px rgba(0,0,0,.08)", marginBottom: 16 }}>
            {/* Header cotización */}
            <div style={{ padding: "28px 28px 20px", borderBottom: "1px solid #f0ece4" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
                <div>
                  <img src={typeof window!=="undefined"?(localStorage.getItem("pf_logo")||"/logo.png"):"/logo.png"} alt="Poliflor" style={{ height: 48, marginBottom: 8 }} onError={(e: any) => { e.target.style.display = "none" }}/>
                  <div style={{ fontSize: 11, color: "#9a9590" }}>Renta de Mobiliario para Eventos</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 11, fontWeight: 600, color: "#9a9590", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 4 }}>Cotización</div>
                  <div style={{ fontFamily: "Playfair Display,serif", fontSize: 26, fontWeight: 800, color: "#1a1814", lineHeight: 1 }}>{cot.folio}</div>
                  {cot.fecha_vigencia && <div style={{ fontSize: 11, color: "#9a9590", marginTop: 4 }}>Vigente hasta: {new Date(cot.fecha_vigencia + "T12:00:00").toLocaleDateString("es-MX", { day: "numeric", month: "long", year: "numeric" })}</div>}
                  {cot.vendedor && <div style={{ fontSize: 11, color: "#9a9590", marginTop: 2 }}>Atendida por: {cot.vendedor}</div>}
                </div>
              </div>

              {/* Cliente y evento */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <div style={{ background: "#f8f6f2", borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 6 }}>Para</div>
                  <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 4 }}>{cot.cliente_nombre || "—"}</div>
                  {cot.cliente_tel && <div style={{ fontSize: 12, color: "#4a4640" }}>📞 {cot.cliente_tel}</div>}
                </div>
                <div style={{ background: "#f8f6f2", borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 6 }}>Evento</div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>📅 {fmtFecha(cot.fecha_evento)}</div>
                  {cot.lugar_evento && <div style={{ fontSize: 12, color: "#4a4640" }}>📍 {cot.lugar_evento}</div>}
                  {cot.fecha_entrega && <div style={{ fontSize: 11, color: "#9a9590", marginTop: 4 }}>Entrega: {cot.fecha_entrega} · Desmonte: {cot.fecha_desmonte || "—"}</div>}
                </div>
              </div>
            </div>

            {/* Tabla artículos */}
            <div style={{ padding: "0" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ background: "#1a1814" }}>
                    <th style={{ padding: "12px 20px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#fff", letterSpacing: ".04em" }}>Descripción</th>
                    <th style={{ padding: "12px 16px", textAlign: "center", fontSize: 11, fontWeight: 600, color: "#fff", width: 70 }}>Cant.</th>
                    <th style={{ padding: "12px 16px", textAlign: "right", fontSize: 11, fontWeight: 600, color: "#fff", width: 110 }}>P. Unit.</th>
                    <th style={{ padding: "12px 20px", textAlign: "right", fontSize: 11, fontWeight: 600, color: "#fff", width: 120 }}>Importe</th>
                  </tr>
                </thead>
                <tbody>
                  {(cot.partidas || []).map((p, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #f0ece4", background: i % 2 === 0 ? "#fff" : "#fafaf8" }}>
                      <td style={{ padding: "11px 20px", fontSize: 13, fontWeight: 500 }}>
                        {p.nombre}
                        {p.notas && <div style={{ fontSize: 11, color: "#9a9590", marginTop: 2 }}>{p.notas}</div>}
                      </td>
                      <td style={{ padding: "11px 16px", textAlign: "center", fontFamily: "monospace", fontWeight: 700, color: "#1a3a5c" }}>{p.cantidad}</td>
                      <td style={{ padding: "11px 16px", textAlign: "right", fontFamily: "monospace", fontSize: 13, color: "#9a9590" }}>{fmt(p.precio_unitario)}</td>
                      <td style={{ padding: "11px 20px", textAlign: "right", fontFamily: "monospace", fontSize: 13, fontWeight: 700, color: "#1a1814" }}>{fmt(p.subtotal)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Totales */}
            <div style={{ padding: "16px 20px", borderTop: "1px solid #f0ece4", display: "flex", justifyContent: "flex-end" }}>
              <div style={{ width: 280 }}>
                <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontSize: 13, color: "#9a9590" }}>
                  <span>Subtotal</span>
                  <span style={{ fontFamily: "monospace", fontWeight: 600, color: "#1a1814" }}>{fmt(cot.subtotal)}</span>
                </div>
                {cot.descuento_pct > 0 && (
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontSize: 13, color: "#2d6a4f" }}>
                    <span>Descuento ({cot.descuento_pct}%)</span>
                    <span style={{ fontFamily: "monospace", fontWeight: 600 }}>-{fmt(cot.descuento_monto)}</span>
                  </div>
                )}
                {cot.aplica_iva && (
                  <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", fontSize: 13, color: "#4a2d6e" }}>
                    <span>IVA (16%)</span>
                    <span style={{ fontFamily: "monospace", fontWeight: 600 }}>{fmt(cot.iva_monto)}</span>
                  </div>
                )}
                <div style={{ display: "flex", justifyContent: "space-between", padding: "14px 16px", background: "#1a1814", borderRadius: 10, marginTop: 8 }}>
                  <span style={{ fontFamily: "Playfair Display,serif", fontSize: 16, fontWeight: 800, color: "#fff" }}>TOTAL</span>
                  <span style={{ fontFamily: "Playfair Display,serif", fontSize: 20, fontWeight: 800, color: "#fff" }}>{fmt(cot.total)}</span>
                </div>
              </div>
            </div>

            {/* Notas y condiciones */}
            {(cot.notas_cliente || cot.condiciones) && (
              <div style={{ padding: "16px 20px", borderTop: "1px solid #f0ece4" }}>
                {cot.notas_cliente && (
                  <div style={{ background: "#fdf5e8", borderLeft: "4px solid #92580a", borderRadius: "0 8px 8px 0", padding: "12px 14px", marginBottom: 12, fontSize: 13, color: "#4a4640" }}>
                    <div style={{ fontWeight: 700, color: "#92580a", marginBottom: 4, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }}>Notas</div>
                    {cot.notas_cliente.split("\n").filter(l => !l.includes("[Aprobada")).map((l, i) => <div key={i}>{l}</div>)}
                  </div>
                )}
                {cot.condiciones && (
                  <div style={{ background: "#f8f6f2", borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontWeight: 700, color: "#9a9590", marginBottom: 6, fontSize: 11, textTransform: "uppercase", letterSpacing: ".06em" }}>Términos y condiciones</div>
                    <div style={{ fontSize: 12, color: "#4a4640", lineHeight: 1.7, whiteSpace: "pre-line" }}>{cot.condiciones}</div>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* ACCIONES */}
          {cot.estado !== "convertida" && cot.estado !== "rechazada" && (
            <div style={{ background: "#fff", borderRadius: 16, padding: "24px", boxShadow: "0 4px 32px rgba(0,0,0,.08)" }}>
              {!confirmar ? (
                <>
                  <div style={{ fontFamily: "Playfair Display,serif", fontSize: 18, fontWeight: 800, marginBottom: 6 }}>¿Apruebas esta cotización?</div>
                  <div style={{ fontSize: 13, color: "#9a9590", marginBottom: 20 }}>Al aprobar, Poliflor confirmará la reserva de tu evento y se pondrá en contacto contigo.</div>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={() => setConfirmar("aprobar")} style={{ flex: 2, padding: "14px", borderRadius: 10, background: "#2d6a4f", color: "#fff", border: "none", fontSize: 15, fontWeight: 700, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>
                      ✅ Aprobar cotización
                    </button>
                    <button onClick={() => setConfirmar("rechazar")} style={{ flex: 1, padding: "14px", borderRadius: 10, background: "#fff", color: "#9a9590", border: "1px solid #e8e5de", fontSize: 14, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>
                      No por ahora
                    </button>
                  </div>
                </>
              ) : confirmar === "aprobar" ? (
                <>
                  <div style={{ fontFamily: "Playfair Display,serif", fontSize: 18, fontWeight: 800, marginBottom: 6 }}>Confirmar aprobación</div>
                  <div style={{ fontSize: 13, color: "#9a9590", marginBottom: 16 }}>Escribe tu nombre completo como confirmación digital.</div>
                  <input value={nombre} onChange={e => setNombre(e.target.value)} placeholder="Tu nombre completo..."
                    style={{ width: "100%", padding: "12px 14px", border: "2px solid #2d6a4f", borderRadius: 10, fontSize: 14, outline: "none", marginBottom: 12, boxSizing: "border-box" as const, fontFamily: "Inter,sans-serif" }}/>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={aprobar} disabled={estado === "aprobando"}
                      style={{ flex: 2, padding: "13px", borderRadius: 10, background: estado === "aprobando" ? "#9a9590" : "#2d6a4f", color: "#fff", border: "none", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>
                      {estado === "aprobando" ? "Confirmando..." : "✅ Confirmar y aprobar"}
                    </button>
                    <button onClick={() => setConfirmar(null)} style={{ flex: 1, padding: "13px", borderRadius: 10, background: "#f5f4f0", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>Cancelar</button>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ fontFamily: "Playfair Display,serif", fontSize: 16, fontWeight: 800, marginBottom: 8 }}>¿Seguro que no te interesa?</div>
                  <div style={{ fontSize: 13, color: "#9a9590", marginBottom: 16 }}>Puedes contactarnos para hacer ajustes a la cotización.</div>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button onClick={rechazar} style={{ flex: 1, padding: "12px", borderRadius: 10, background: "#fdf0f0", color: "#8b2e2e", border: "1px solid #e8b8b8", fontSize: 13, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>Rechazar cotización</button>
                    <button onClick={() => setConfirmar(null)} style={{ flex: 2, padding: "12px", borderRadius: 10, background: "#1a1814", color: "#fff", border: "none", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "Inter,sans-serif" }}>← Volver a la cotización</button>
                  </div>
                </>
              )}
            </div>
          )}

          {/* Footer */}
          <div style={{ textAlign: "center", marginTop: 32, color: "#9a9590", fontSize: 12 }}>
            <div style={{ marginBottom: 4 }}>🌸 Poliflor — Renta de Mobiliario para Eventos</div>
            <div>Este documento fue generado digitalmente y tiene validez como cotización formal.</div>
          </div>
        </div>
      </div>
    </>
  )
}
