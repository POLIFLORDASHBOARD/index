// BUILD: 1785892031
impor${["vajilla","carpinteria","bases","flores"].includes(split.tipo)&&split.fecha_entrega_contrato?`<div style="font-size:11px;color:#9a9590;margin-top:3px">🚚 Entrega: ${new Date(split.fecha_entrega_contrato+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>`:""}\n    ${["vajilla","carpinteria","bases","flores"].includes(split.tipo)&&split.fecha_desmonte_contrato?`<div style="font-size:11px;color:#9a9590;margin-top:2px">📦 Desmonte: ${new Date(split.fecha_desmonte_contrato+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>`:""}\n  </div>\n  <!-- Hoja de Trabajo / Área -->
  <div style="text-align:right">
    <div style="font-size:9px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px">Hoja de Trabajo</div>
    <div style="display:inline-flex;align-items:center;gap:6px;background:${tipo.bg};border:2px solid ${tipo.color};border-radius:8px;padding:5px 12px">
      <span style="font-size:18px">${tipo.icono}</span>
      <span style="font-size:15px;font-weight:800;color:${tipo.color}">${tipo.nombre}</span>
    </div>
    ${split.responsable?`<div style="font-size:11px;color:#9a9590;margin-top:4px">👤 ${split.responsable}</div>`:""}
  </div>
</div>

<!-- ══ TABLA DE ARTÍCULOS ══ -->
<table style="margin-bottom:14px;border:1.5px solid #e8e5de;overflow:hidden">
  <thead>
    <tr style="background:${tipo.color};color:#fff">
      <th style="padding:8px 10px;text-align:center;font-size:11px;font-weight:700;width:56px">Cant.</th>
      <th style="padding:8px 10px;text-align:left;font-size:11px;font-weight:700">Artículo / Descripción</th>
      ${mostrarPrecio?`<th style="padding:8px 10px;text-align:right;font-size:11px;font-weight:700;width:90px">P.U.</th>`:""}
      <th style="padding:8px 10px;text-align:left;font-size:11px;font-weight:700;width:120px">Nota</th>
    </tr>
  </thead>
  <tbody>
    ${filas||`<tr><td colspan="10" style="padding:20px;text-align:center;color:#9a9590;font-style:italic">Sin artículos asignados</td></tr>`}
  </tbody>
  
</table>

${split.notas?`<div style="background:${tipo.bg};border:1px solid ${tipo.color}44;border-radius:7px;padding:8px 12px;margin-bottom:14px;font-size:11px"><span style="font-weight:700;color:${tipo.color}">Notas: </span>${split.notas}</div>`:""}

<!-- ══ FIRMAS ══ -->
<div style="border-top:1.5px solid #e8e5de;padding-top:12px;margin-top:32px;display:grid;grid-template-columns:1fr 1fr;gap:48px">
  <div style="text-align:center">
    <div style="height:40px;border-bottom:1.5px solid #1a1814;margin-bottom:7px"></div>
    <div style="font-size:10px;color:#9a9590">Responsable — ${tipo.nombre}</div>
    ${split.responsable?`<div style="font-size:11px;font-weight:700;color:#1a1814;margin-top:1px">${split.responsable}</div>`:""}
  </div>
  <div style="text-align:center">
    <div style="height:40px;border-bottom:1.5px solid #1a1814;margin-bottom:7px"></div>
    <div style="font-size:10px;color:#9a9590">Autorizado por Poliflor</div>
  </div>
</div>

</body></html>`
    const w=window.open("","_blank","width=1000,height=820")
  if(w){w.document.write(html);w.document.close()}
}

// ── SplitsSection ────────────────────────────────────────────────────
function SplitsSection({token,contratos,logoUrl}:{token:string,contratos:any[],logoUrl:string}){
  const [splits,setSplits]=useState<any[]>([])
  const [cargando,setCargando]=useState(true)
  const [semOff,setSemOff]=useState(0)
  const [vistaMode,setVistaMode]=useState<"splits"|"vendedores">("splits")
  const [busq,setBusq]=useState("")
  const [filtroEst,setFiltroEst]=useState("todos")
  const [splitEdit,setSplitEdit]=useState<any|null>(null)
  const [editNotas,setEditNotas]=useState("")
  const [editObs,setEditObs]=useState("")
  const [editResp,setEditResp]=useState("")
  const [editArts,setEditArts]=useState<any[]>([])
  const [busqArtSplit,setBusqArtSplit]=useState("")
  const [sugsArtSplit,setSugsArtSplit]=useState<any[]>([])
  const [saving,setSaving]=useState(false)
  const [contratoEditRef,setContratoEditRef]=useState<any|null>(null)
  const [editFecha,setEditFecha]=useState("")
  const [editFechaOriginal,setEditFechaOriginal]=useState("")
  const [contratoPanel,setContratoPanel]=useState<any|null>(null)
  const [modalProv,setModalProv]=useState<any|null>(null)
  const [provNombre,setProvNombre]=useState("")
  const [provArts,setProvArts]=useState<any[]>([])
  const [provBusq,setProvBusq]=useState("")
  const [provSugs,setProvSugs]=useState<any[]>([])
  const [provMostrarDir,setProvMostrarDir]=useState(false)
  const [provDireccion,setProvDireccion]=useState("")
  const [creandoProv,setCreandoProv]=useState(false)

  const hoy=new Date();hoy.setHours(0,0,0,0)
  const dow=hoy.getDay()===0?6:hoy.getDay()-1
  const lunesBase=new Date(hoy);lunesBase.setDate(hoy.getDate()-dow+semOff*7)
  const finBase=new Date(lunesBase);finBase.setDate(lunesBase.getDate()+6)
  const isoD=(d:Date)=>d.toISOString().split("T")[0]
  const MESES=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
  const semLabel=`${lunesBase.getDate()} ${MESES[lunesBase.getMonth()]} — ${finBase.getDate()} ${MESES[finBase.getMonth()]} ${finBase.getFullYear()}`

  const cargar=useCallback(async()=>{
    setCargando(true)
    const r=await fetch(`/api/splits?fecha_desde=${isoD(lunesBase)}&fecha_hasta=${isoD(finBase)}`,{headers:{Authorization:`Bearer ${token}`}})
    const data=await r.json()
    setSplits(Array.isArray(data)?data:[])
    setCargando(false)
  },[semOff,token])

  // Generar splits directo desde la vista de vendedores
  const [generandoLocal,setGenerandoLocal]=useState<string|null>(null)
  const generarSplitsLocal=async(contrato:any)=>{
    setGenerandoLocal(contrato.id)
    // Borrar splits anteriores EXCEPTO proveedores
    const existR=await fetch(`/api/splits?contrato_id=${contrato.id}`,{headers:{Authorization:`Bearer ${token}`}})
    const exist=await existR.json()
    if(Array.isArray(exist)){
      const toDel=exist.filter((s:any)=>s.tipo!=="proveedor")
      await Promise.all(toDel.map((s:any)=>
        fetch(`/api/splits?id=${s.id}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}})
      ))
    }
    const nuevos=generarSplitsDesdeContrato(contrato)
    const r=await fetch("/api/splits",{
      method:"POST",
      headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
      body:JSON.stringify(nuevos)
    })
    const data=await r.json()
    if(Array.isArray(data)){
      // Reload full splits list
      await cargar()
    }
    setGenerandoLocal(null)
  }

  useEffect(()=>{cargar()},[cargar])

  const gruposPorContrato=splits.reduce((acc:any,s:any)=>{
    const key=s.contrato_id||"sin_contrato"
    // Usar fecha_evento_original si existe, sino fecha_evento del split no-desmonte
    const fechaGrupo=s.fecha_evento_original||(s.tipo!=="desmonte"?s.fecha_evento:null)||s.fecha_evento
    if(!acc[key]) acc[key]={contrato_id:s.contrato_id,folio:s.contrato_folio||"Sin folio",cliente:s.cliente||"—",fecha:fechaGrupo||"—",splits:[]}
    if(fechaGrupo&&fechaGrupo<acc[key].fecha) acc[key].fecha=fechaGrupo
    acc[key].splits.push(s)
    return acc
  },{})
  const grupos=Object.values(gruposPorContrato) as any[]

  const gruposFilt=grupos.filter((g:any)=>{
    if(busq){
      const q=busq.toLowerCase()
      if(!(g.cliente||"").toLowerCase().includes(q)&&!(g.folio||"").toLowerCase().includes(q))return false
    }
    if(filtroEst!=="todos") return g.splits.some((s:any)=>s.estado===filtroEst)
    return true
  }).sort((a:any,b:any)=>a.fecha.localeCompare(b.fecha))

  const kpis=SPLIT_ESTADOS.map(e=>({...e,n:splits.filter((s:any)=>s.estado===e.id).length}))

  const cambiarEstado=async(s:any,estado:string)=>{
    await fetch(`/api/splits?id=${s.id}`,{method:"PATCH",headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},body:JSON.stringify({estado})})
    setSplits(prev=>prev.map((x:any)=>x.id===s.id?{...x,estado}:x))
    if(splitEdit?.id===s.id)setSplitEdit((p:any)=>({...p,estado}))
  }

  const abrirEdicion=(s:any)=>{
    setSplitEdit(s); setEditNotas(s.notas||""); setEditObs(s.observaciones||""); setEditResp(s.responsable||"")
    setEditFecha(s.fecha_evento||"")
    // For desmonte: store the contrato's original event date to compute +1/al termino
    if(s.tipo==="desmonte"){
      const ct=contratos.find((x:any)=>x.id===s.contrato_id)
      setEditFechaOriginal(ct?.fecha_evento||s.fecha_evento||"")
    }
    setEditArts(JSON.parse(JSON.stringify(s.articulos||[]))); setBusqArtSplit(""); setSugsArtSplit([]); setContratoPanel(null)
    // Load full contrato to show checklist
    const ct=contratos.find((x:any)=>x.id===s.contrato_id)
    setContratoEditRef(ct||null)
  }

  const guardarEdicion=async()=>{
    if(!splitEdit)return; setSaving(true)
    const updates={notas:editNotas,observaciones:editObs,responsable:editResp,articulos:editArts,fecha_evento:editFecha||splitEdit.fecha_evento}
    await fetch(`/api/splits?id=${splitEdit.id}`,{method:"PATCH",headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},body:JSON.stringify(updates)})
    setSplits(prev=>prev.map((x:any)=>x.id===splitEdit.id?{...x,...updates}:x))
    setSplitEdit((p:any)=>p?{...p,...updates,fecha_evento:editFecha||splitEdit.fecha_evento}:null)
    setSaving(false)
  }

  const verContrato=(contratoId:string)=>{
    const ct=contratos.find((x:any)=>x.id===contratoId)
    setContratoPanel(ct||null); setSplitEdit(null)
  }

  const abrirModalProv=(g:any)=>{
    const ct=contratos.find((x:any)=>x.id===g.contrato_id)
    setModalProv({contratoId:g.contrato_id,folio:g.folio,cliente:g.cliente,fecha:g.fecha,fechaEvento:ct?.fecha_evento||g.fecha,lugarContrato:ct?.lugar||""})
    setProvNombre(""); setProvArts([]); setProvBusq(""); setProvMostrarDir(false); setProvDireccion("")
  }

  const guardarProveedor=async()=>{
    if(!provNombre.trim()||!modalProv)return
    setCreandoProv(true)
    const nuevo={
      contrato_id:modalProv.contratoId,contrato_folio:modalProv.folio,cliente:modalProv.cliente,
      lugar:provMostrarDir?(provDireccion||modalProv.lugarContrato||""):"",
      fecha_evento:modalProv.fecha,tipo:"proveedor",nombre:`📦 ${provNombre}`,
      proveedor_nombre:provNombre,estado:"pendiente",articulos:provArts,notas:"",observaciones:"",
      mostrar_precios:false,mostrar_direccion:provMostrarDir,
    }
    const r=await fetch("/api/splits",{method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},body:JSON.stringify([nuevo])})
    const data=await r.json()
    if(Array.isArray(data))setSplits(prev=>[...prev,...data])
    setCreandoProv(false); setModalProv(null)
  }

  const imprimirProveedor=()=>{
    if(!provNombre.trim()||!modalProv)return
    abrirHojaSplit({tipo:"proveedor",nombre:provNombre,cliente:modalProv.cliente,contrato_folio:modalProv.folio,
      lugar:provMostrarDir?(provDireccion||modalProv.lugarContrato||""):"",fecha_evento:modalProv.fecha,
      fecha_preparacion:modalProv.fecha,
      articulos:provArts,notas:"",mostrar_direccion:provMostrarDir,mostrar_precios:false,responsable:""},logoUrl)
  }

  const SORT_ORD:{[k:string]:number}={rutas:0,carpinteria:1,vajilla:2,flores:3,bases:4,desmonte:5,proveedor:6}

  return(
    <div style={{display:"flex",flexDirection:"column" as const,gap:14}}>

      {/* ── MODAL PROVEEDOR ── */}
      {modalProv&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.6)",zIndex:3000,display:"flex",alignItems:"center" as const,justifyContent:"center" as const,padding:16}}
          onClick={e=>{if(e.target===e.currentTarget)setModalProv(null)}}>
          <div style={{background:"#fff",borderRadius:16,width:560,maxHeight:"88vh",display:"flex",flexDirection:"column" as const,boxShadow:"0 32px 80px rgba(0,0,0,.3)",overflow:"hidden"}}>
            <div style={{background:"#8b2e2e",padding:"16px 20px",display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:22}}>📦</span>
              <div style={{flex:1}}>
                <div style={{fontSize:15,fontWeight:800,color:"#fff"}}>Solicitar a Proveedor</div>
                <div style={{fontSize:11,color:"rgba(255,255,255,.6)"}}>
                  {modalProv.cliente} · <span style={{fontFamily:"monospace",fontWeight:700,color:"#fca5a5"}}>{modalProv.folio}</span>
                </div>
              </div>
              <button onClick={()=>setModalProv(null)} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:15}}>✕</button>
            </div>
            <div style={{padding:20,overflowY:"auto" as const,display:"flex",flexDirection:"column" as const,gap:14}}>
              <div>
                <label style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",display:"block",marginBottom:6}}>Nombre del proveedor *</label>
                <input value={provNombre} onChange={e=>setProvNombre(e.target.value)} placeholder="Ej: Flores del Valle, Mantelería Express..."
                  style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
              </div>
              <div style={{background:"#f8f6f2",borderRadius:8,padding:"10px 12px"}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <input type="checkbox" checked={provMostrarDir} onChange={e=>setProvMostrarDir(e.target.checked)} style={{width:16,height:16,cursor:"pointer"}}/>
                  <span style={{fontSize:12,fontWeight:600,color:"#4a4640",cursor:"pointer"}} onClick={()=>setProvMostrarDir(v=>!v)}>Incluir dirección de entrega en la hoja</span>
                </div>
                {provMostrarDir&&(
                  <div style={{marginTop:10}}>
                    <div style={{fontSize:10,color:"#9a9590",fontWeight:600,marginBottom:5}}>Dirección de entrega:</div>
                    <input value={provDireccion} onChange={e=>setProvDireccion(e.target.value)}
                      placeholder={modalProv?.lugarContrato||"Escribe la dirección específica..."}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}/>
                    {modalProv?.lugarContrato&&!provDireccion&&(
                      <div style={{fontSize:10,color:"#9a9590",marginTop:4}}>
                        Vacío = usa la del contrato: <strong>{modalProv.lugarContrato.slice(0,50)}</strong>
                      </div>
                    )}
                  </div>
                )}
              </div>
              <div>
                <div style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",marginBottom:8}}>
                  Artículos ({provArts.length}) · {provArts.reduce((s:number,a:any)=>s+(a.cantidad||0),0)} pzas
                </div>
                {provArts.length>0&&(
                  <div style={{border:"1px solid #e8e5de",borderRadius:8,overflow:"hidden",marginBottom:8,maxHeight:180,overflowY:"auto" as const}}>
                    {provArts.map((a:any,i:number)=>(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderBottom:i<provArts.length-1?"1px solid #f0ece4":"none",background:i%2===0?"#fff":"#fafaf8"}}>
                        <div style={{flex:1,fontSize:12,fontWeight:500}}>{a.nombre}</div>
                        <input type="number" min="1" value={a.cantidad}
                          onChange={e=>{const v=Math.max(1,parseInt(e.target.value)||1);setProvArts(prev=>prev.map((x:any,j:number)=>j===i?{...x,cantidad:v}:x))}}
                          style={{width:50,padding:"3px 4px",border:"1.5px solid #e8e5de",borderRadius:6,textAlign:"center" as const,fontFamily:"monospace",fontSize:12,fontWeight:700,outline:"none"}}/>
                        <button onClick={()=>setProvArts(prev=>prev.filter((_:any,j:number)=>j!==i))}
                          style={{width:22,height:22,borderRadius:4,background:"#fdf0f0",border:"none",cursor:"pointer",color:"#8b2e2e",fontSize:16,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                      </div>
                    ))}
                  </div>
                )}
                <div style={{position:"relative" as const}}>
                  <input value={provBusq} onChange={async e=>{
                    const v=e.target.value; setProvBusq(v)
                    if(v.trim().length>=2){
                      const r=await fetch(`/api/catalogo?busq=${encodeURIComponent(v.trim())}&activo=true`,{headers:{Authorization:`Bearer ${token}`}})
                      const data=await r.json()
                      setProvSugs(Array.isArray(data)?data.slice(0,6):[])
                    }else{setProvSugs([])}
                  }} onBlur={()=>setTimeout(()=>setProvSugs([]),200)}
                  placeholder="Buscar artículo del catálogo..."
                  style={{width:"100%",padding:"8px 12px",border:"1.5px solid #8b2e2e",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const,background:"#fdf0f0"}}/>
                  {provSugs.length>0&&(
                    <div style={{position:"absolute" as const,top:"100%",left:0,right:0,background:"#fff",border:"2px solid #8b2e2e",borderRadius:"0 0 8px 8px",boxShadow:"0 8px 24px rgba(139,46,46,.2)",zIndex:100,maxHeight:200,overflowY:"auto" as const}}>
                      {provSugs.map((a:any,ai:number)=>(
                        <div key={ai} onMouseDown={e=>{
                          e.preventDefault()
                          const exists=provArts.some((x:any)=>x.nombre===a.nombre)
                          if(exists){setProvArts(prev=>prev.map((x:any)=>x.nombre===a.nombre?{...x,cantidad:(x.cantidad||0)+1}:x))}
                          else{setProvArts(prev=>[...prev,{nombre:a.nombre,cantidad:1,pu:a.precio_renta||0,importe:0,seccion:a.categoria||""}])}
                          setProvBusq(""); setProvSugs([])
                        }} style={{padding:"7px 12px",cursor:"pointer",borderBottom:"1px solid #f1f5f9",display:"flex",gap:8,fontSize:12}}
                        onMouseEnter={e=>(e.currentTarget as HTMLElement).style.background="#fdf0f0"}
                        onMouseLeave={e=>(e.currentTarget as HTMLElement).style.background="#fff"}>
                          <span style={{flex:1,fontWeight:500}}>{a.nombre}</span>
                          <span style={{fontSize:10,color:"#9a9590"}}>{a.categoria}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                {(()=>{
                  const ct=contratos.find((x:any)=>x.id===modalProv.contratoId)
                  if(!ct||(ct.articulos||[]).length===0)return null
                  const arts=ct.articulos||[]
                  return(
                    <div style={{border:"1.5px solid #8b2e2e55",borderRadius:8,overflow:"hidden",marginTop:8}}>
                      <div style={{background:"#8b2e2e",padding:"6px 10px",display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{fontSize:11,fontWeight:700,color:"#fff",flex:1}}>📋 Artículos del contrato</span>
                        <button onMouseDown={e=>{e.preventDefault();setProvArts(arts.map((a:any)=>({nombre:a.nombre,cantidad:a.cantidad||1,pu:a.pu||0,importe:a.importe||0,seccion:a.seccion||""})))}}
                          style={{fontSize:9,padding:"2px 8px",borderRadius:4,background:"rgba(255,255,255,.2)",border:"none",color:"#fff",cursor:"pointer",fontWeight:700}}>✓ Todos</button>
                        <button onMouseDown={e=>{e.preventDefault();setProvArts([])}}
                          style={{fontSize:9,padding:"2px 8px",borderRadius:4,background:"rgba(255,255,255,.15)",border:"none",color:"rgba(255,255,255,.8)",cursor:"pointer",fontWeight:700}}>✗ Ninguno</button>
                      </div>
                      <div style={{maxHeight:200,overflowY:"auto" as const}}>
                        {arts.map((a:any,i:number)=>{
                          const selIdx=provArts.findIndex((x:any)=>x.nombre===a.nombre)
                          const sel=selIdx>=0
                          return(
                            <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderBottom:i<arts.length-1?"1px solid #f0ece4":"none",background:sel?"#fdf0f0":"#fff",cursor:"pointer"}}
                              onClick={()=>{
                                if(sel){setProvArts(prev=>prev.filter((_:any,j:number)=>j!==selIdx))}
                                else{setProvArts(prev=>[...prev,{nombre:a.nombre,cantidad:a.cantidad||1,pu:a.pu||0,importe:a.importe||0,seccion:a.seccion||""}])}
                              }}>
                              <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${sel?"#8b2e2e":"#c4bfb8"}`,background:sel?"#8b2e2e":"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                                {sel&&<span style={{color:"#fff",fontSize:12,lineHeight:1,fontWeight:900}}>✓</span>}
                              </div>
                              <div style={{flex:1,fontSize:12,fontWeight:sel?700:500,color:sel?"#8b2e2e":"#4a4640",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</div>
                              <span style={{fontSize:10,color:"#9a9590",fontFamily:"monospace",flexShrink:0}}>{a.cantidad}x</span>
                              {sel&&(
                                <input type="number" min="1" value={provArts[selIdx].cantidad}
                                  onClick={e=>e.stopPropagation()}
                                  onChange={e=>{e.stopPropagation();const v=Math.max(1,parseInt(e.target.value)||1);setProvArts(prev=>prev.map((x:any,j:number)=>j===selIdx?{...x,cantidad:v}:x))}}
                                  style={{width:42,padding:"2px 4px",border:"1.5px solid #8b2e2e",borderRadius:5,textAlign:"center" as const,fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#8b2e2e",outline:"none"}}/>
                              )}
                            </div>
                          )
                        })}
                      </div>
                    </div>
                  )
                })()}
              </div>
            </div>
            <div style={{padding:"12px 20px",borderTop:"1px solid #e8e5de",display:"flex",gap:8}}>
              <button onClick={guardarProveedor} disabled={creandoProv||!provNombre.trim()}
                style={{flex:2,padding:"11px",borderRadius:8,background:creandoProv||!provNombre.trim()?"#9a9590":"#8b2e2e",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:700}}>
                {creandoProv?"Creando...":"📦 Guardar hoja de proveedor"}
              </button>
              <button onClick={imprimirProveedor}
                style={{flex:1,padding:"11px",borderRadius:8,background:"#f5f4f0",color:"#4a4640",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                🖨️ Solo imprimir
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── HEADER ── */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"16px 20px"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap" as const,marginBottom:14}}>
          <div style={{flex:1}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800}}>✂️ Splits — Centro de Operaciones</div>
            <div style={{fontSize:12,color:"#9a9590",marginTop:2}}>{semLabel} · {splits.length} hojas</div>
          </div>
          <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap" as const}}>
            {/* Vista toggle */}
            <div style={{display:"flex",background:"#f5f4f0",borderRadius:8,padding:3,gap:3}}>
              {(["splits","vendedores"] as const).map(v=>(
                <button key={v} onClick={()=>setVistaMode(v)}
                  style={{padding:"5px 12px",borderRadius:6,border:"none",background:vistaMode===v?"#0f172a":"transparent",color:vistaMode===v?"#fff":"#4a4640",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:vistaMode===v?700:400,cursor:"pointer"}}>
                  {v==="splits"?"✂️ Splits":"👥 Por vendedor"}
                </button>
              ))}
            </div>
            <button onClick={()=>setSemOff(v=>v-1)} style={{width:32,height:32,borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16}}>‹</button>
            <button onClick={()=>setSemOff(0)} style={{padding:"6px 12px",borderRadius:8,border:"1px solid #e8e5de",background:semOff===0?"#0f172a":"#fff",color:semOff===0?"#fff":"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>Esta semana</button>
            <button onClick={()=>setSemOff(v=>v+1)} style={{width:32,height:32,borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16}}>›</button>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
          {kpis.map(e=>(
            <div key={e.id} onClick={()=>setFiltroEst(filtroEst===e.id?"todos":e.id)}
              style={{background:filtroEst===e.id?e.bg:"#f8fafc",border:`1.5px solid ${filtroEst===e.id?e.color:"#e8e5de"}`,borderRadius:8,padding:"10px 12px",cursor:"pointer"}}>
              <div style={{fontSize:9,fontWeight:700,color:e.color,textTransform:"uppercase" as const,letterSpacing:".05em"}}>{e.label}</div>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:e.color}}>{e.n}</div>
            </div>
          ))}
        </div>
      </div>

      {/* ── VISTA VENDEDORES ── */}
      {vistaMode==="vendedores"&&(()=>{
        // Contratos de la semana (con o sin splits)
        const contratosDelaPeriodo=contratos.filter((x:any)=>{
          const tipo=(x.tipo||"contrato").toLowerCase()
          if(tipo!=="contrato")return false
          const fe=x.fecha_evento||""
          return fe>=isoD(lunesBase)&&fe<=isoD(finBase)
        })
        // IDs que tienen splits esta semana
        const idsConSplits=new Set(splits.map((s:any)=>s.contrato_id))
        // Agrupa por vendedor
        const porVend:{[k:string]:any[]}={}
        contratosDelaPeriodo.forEach((x:any)=>{
          const v=x.vendedor||vendedorDesdeFolio(x.folio||"")||"Sin vendedor"
          if(!porVend[v])porVend[v]=[]
          porVend[v].push(x)
        })
        const vendedoresOrden=VENDEDORES.map(v=>v.nombre).filter(v=>porVend[v]).concat(porVend["Sin vendedor"]?["Sin vendedor"]:[])
        if(contratosDelaPeriodo.length===0) return(
          <div style={{padding:48,textAlign:"center" as const,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12,color:"#9a9590"}}>
            <div style={{fontSize:32,opacity:.2,marginBottom:8}}>📅</div>
            <div style={{fontWeight:600}}>Sin contratos esta semana</div>
          </div>
        )
        return(
          <div style={{display:"flex",flexDirection:"column" as const,gap:12}}>
            {/* Resumen semana */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
              {[
                {l:"Contratos semana",v:contratosDelaPeriodo.length,c:"#1a1814",bg:"#f8fafc"},
                {l:"Con splits listos",v:contratosDelaPeriodo.filter((x:any)=>idsConSplits.has(x.id)).length,c:"#2d6a4f",bg:"#f0fdf4"},
                {l:"Sin splits — pendiente",v:contratosDelaPeriodo.filter((x:any)=>!idsConSplits.has(x.id)).length,c:"#8b2e2e",bg:"#fdf0f0"},
              ].map((k,i)=>(
                <div key={i} style={{background:k.bg,border:`1px solid ${k.c}22`,borderRadius:10,padding:"12px 14px"}}>
                  <div style={{fontSize:9,fontWeight:700,color:k.c,textTransform:"uppercase" as const,letterSpacing:".05em"}}>{k.l}</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:26,fontWeight:800,color:k.c,marginTop:2}}>{k.v}</div>
                </div>
              ))}
            </div>
            {/* Por vendedor */}
            {vendedoresOrden.map(vend=>{
              const ctsList=porVend[vend]||[]
              const conSplit=ctsList.filter((x:any)=>idsConSplits.has(x.id))
              const sinSplit=ctsList.filter((x:any)=>!idsConSplits.has(x.id))
              const pct=ctsList.length>0?Math.round(conSplit.length/ctsList.length*100):0
              return(
                <div key={vend} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
                  {/* Header vendedor */}
                  <div style={{padding:"12px 16px",display:"flex",alignItems:"center",gap:12,borderBottom:"1px solid #f0ece4",background:"#fafaf8"}}>
                    <div style={{width:36,height:36,borderRadius:"50%",background:"#0f172a",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:15,flexShrink:0}}>
                      {vend.charAt(0)}
                    </div>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:700,fontSize:14}}>{vend}</div>
                      <div style={{fontSize:11,color:"#9a9590",marginTop:1}}>{ctsList.length} contrato{ctsList.length!==1?"s":""} esta semana</div>
                    </div>
                    {/* Barra progreso */}
                    <div style={{textAlign:"right" as const}}>
                      <div style={{fontSize:13,fontWeight:800,color:pct===100?"#2d6a4f":pct>50?"#92580a":"#8b2e2e"}}>{pct}%</div>
                      <div style={{fontSize:9,color:"#9a9590"}}>splits listos</div>
                    </div>
                    <div style={{width:60,height:6,background:"#f0ece4",borderRadius:3,overflow:"hidden"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:pct===100?"#2d6a4f":pct>50?"#92580a":"#8b2e2e",borderRadius:3,transition:"width .3s"}}/>
                    </div>
                  </div>
                  {/* Lista de contratos */}
                  <div>
                    {/* Primero los SIN split */}
                    {sinSplit.map((x:any,i:number)=>{
                      const nArts=(x.articulos||[]).reduce((s:number,a:any)=>s+(a.cantidad||0),0)
                      const diasEvento=Math.ceil((new Date(x.fecha_evento+"T12:00:00").getTime()-Date.now())/86400000)
                      const urgente=diasEvento<=3
                      return(
                        <div key={x.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px",borderBottom:"1px solid #f9f7f4",background:"#fffbfb"}}>
                          {/* Indicador NO split */}
                          <div style={{width:6,height:6,borderRadius:"50%",background:"#ef4444",flexShrink:0}}/>
                          {/* Fecha */}
                          <div style={{width:38,flexShrink:0,textAlign:"center" as const,background:"#fdf0f0",borderRadius:7,padding:"4px 3px"}}>
                            <div style={{fontSize:16,fontWeight:800,color:"#8b2e2e",lineHeight:1}}>{x.fecha_evento?.slice(8)||"—"}</div>
                            <div style={{fontSize:7,color:"#8b2e2e",textTransform:"uppercase" as const}}>
                              {["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][(parseInt(x.fecha_evento?.slice(5,7)||"1")-1)]||""}
                            </div>
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente||x.archivo}</div>
                            <div style={{display:"flex",gap:6,marginTop:2,flexWrap:"wrap" as const}}>
                              {x.folio&&<span style={{fontSize:9,fontFamily:"monospace",background:"#f5f4f0",padding:"0 5px",borderRadius:3}}>{x.folio}</span>}
                              {nArts>0&&<span style={{fontSize:9,color:"#9a9590"}}>📦 {nArts} pzas</span>}
                              <span style={{fontSize:9,color:urgente?"#8b2e2e":"#92580a",fontWeight:700,background:urgente?"#fdf0f0":"#fffbeb",padding:"0 5px",borderRadius:3}}>
                                {urgente?"🚨":"⚠️"} en {diasEvento}d
                              </span>
                            </div>
                          </div>
                          <button
                            onClick={e=>{e.stopPropagation();generarSplitsLocal(x)}}
                            disabled={generandoLocal===x.id}
                            style={{fontSize:10,padding:"5px 10px",borderRadius:7,background:generandoLocal===x.id?"#9a9590":"#0f172a",color:"#fff",border:"none",cursor:"pointer",fontWeight:700,whiteSpace:"nowrap" as const,flexShrink:0,fontFamily:"Epilogue,sans-serif"}}>
                            {generandoLocal===x.id?"Generando...":"✂️ Generar split"}
                          </button>
                        </div>
                      )
                    })}
                    {/* Luego los CON split */}
                    {conSplit.map((x:any,i:number)=>{
                      const splitsDelCont=splits.filter((s:any)=>s.contrato_id===x.id)
                      const nArts=(x.articulos||[]).reduce((s:number,a:any)=>s+(a.cantidad||0),0)
                      const pendientes=splitsDelCont.filter((s:any)=>s.estado==="pendiente").length
                      const listos=splitsDelCont.filter((s:any)=>s.estado==="listo"||s.estado==="entregado").length
                      return(
                        <div key={x.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px",borderBottom:i<conSplit.length-1?"1px solid #f9f7f4":"none",background:"#fff",opacity:.85}}>
                          <div style={{width:6,height:6,borderRadius:"50%",background:"#22c55e",flexShrink:0}}/>
                          <div style={{width:38,flexShrink:0,textAlign:"center" as const,background:"#f0fdf4",borderRadius:7,padding:"4px 3px"}}>
                            <div style={{fontSize:16,fontWeight:800,color:"#2d6a4f",lineHeight:1}}>{x.fecha_evento?.slice(8)||"—"}</div>
                            <div style={{fontSize:7,color:"#2d6a4f",textTransform:"uppercase" as const}}>
                              {["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][(parseInt(x.fecha_evento?.slice(5,7)||"1")-1)]||""}
                            </div>
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontWeight:600,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const,color:"#4a4640"}}>{x.cliente||x.archivo}</div>
                            <div style={{display:"flex",gap:6,marginTop:2,flexWrap:"wrap" as const}}>
                              {x.folio&&<span style={{fontSize:9,fontFamily:"monospace",background:"#f5f4f0",padding:"0 5px",borderRadius:3}}>{x.folio}</span>}
                              {/* Estado de cada split */}
                              {splitsDelCont.slice(0,5).map((s:any)=>{
                                const tp=SPLIT_TIPOS.find((t:any)=>t.id===s.tipo)
                                const est=SPLIT_ESTADOS.find((e:any)=>e.id===s.estado)||SPLIT_ESTADOS[0]
                                return tp?(
                                  <span key={s.id} style={{fontSize:9,padding:"1px 5px",borderRadius:3,background:est.bg,color:est.color,fontWeight:600}} title={`${tp.nombre}: ${est.label}`}>
                                    {tp.icono}
                                  </span>
                                ):null
                              })}
                              {pendientes>0&&<span style={{fontSize:9,color:"#92580a",fontWeight:700}}>· {pendientes} pend.</span>}
                            </div>
                          </div>
                          <span style={{fontSize:9,padding:"3px 8px",borderRadius:6,background:"#f0fdf4",color:"#2d6a4f",fontWeight:700,whiteSpace:"nowrap" as const,flexShrink:0}}>
                            ✓ {listos}/{splitsDelCont.length}
                          </span>
                        </div>
                      )
                    })}
                  </div>
                </div>
              )
            })}
          </div>
        )
      })()}

      {/* ── VISTA SPLITS (original) ── */}
      {vistaMode==="splits"&&<>
      {/* ── FILTROS ── */}
      <div style={{display:"flex",gap:8}}>
        <input value={busq} onChange={e=>setBusq(e.target.value)} placeholder="🔍 Buscar cliente o folio..."
          style={{flex:1,padding:"8px 12px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none"}}/>
        {filtroEst!=="todos"&&(
          <button onClick={()=>setFiltroEst("todos")}
            style={{padding:"8px 12px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,color:"#4a4640"}}>
            Todos los estados ×
          </button>
        )}
      </div>

      {/* ── LAYOUT: lista + panel lateral ── */}
      <div style={{display:"grid",gridTemplateColumns:splitEdit||contratoPanel?"1fr 420px":"1fr",gap:14,alignItems:"flex-start" as const}}>

        {/* ── LISTA POR CONTRATO ── */}
        <div style={{display:"flex",flexDirection:"column" as const,gap:10}}>
          {cargando?(
            <div style={{padding:48,textAlign:"center" as const,color:"#9a9590",background:"#fff",border:"1px solid #e8e5de",borderRadius:12}}>Cargando splits...</div>
          ):gruposFilt.length===0?(
            <div style={{padding:48,textAlign:"center" as const,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12,color:"#9a9590"}}>
              <div style={{fontSize:36,opacity:.2,marginBottom:8}}>✂️</div>
              <div style={{fontWeight:600}}>Sin splits esta semana</div>
              <div style={{fontSize:11,marginTop:6}}>Genera splits desde un contrato en el módulo de <strong>Contratos</strong></div>
            </div>
          ):(
            gruposFilt.map((g:any)=>(
              <div key={g.contrato_id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
                <div style={{background:"#0f172a",padding:"12px 16px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap" as const}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" as const}}>
                      <span style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800,color:"#fff"}}>{g.cliente}</span>
                      <span style={{fontFamily:"monospace",fontSize:11,color:"#60a5fa",background:"rgba(96,165,250,.15)",padding:"2px 8px",borderRadius:6,fontWeight:700}}>{g.folio}</span>
                    </div>
                    <div style={{fontSize:11,color:"rgba(255,255,255,.7)",marginTop:4,display:"flex",flexWrap:"wrap" as const,gap:8}}>
                      {(()=>{
                        const ct=contratos.find((x:any)=>x.id===g.contrato_id)
                        const fmt=(f:string)=>f?new Date(f+"T12:00:00").toLocaleDateString("es-MX",{weekday:"short",day:"numeric",month:"short",year:"numeric"}):""
                        return(<>
                          {ct?.fecha_evento&&<span>🎉 {fmt(ct.fecha_evento)}</span>}
                          {ct?.fecha_entrega&&<span style={{color:"#93c5fd"}}>🚚 {fmt(ct.fecha_entrega)}</span>}
                          {ct?.fecha_desmonte&&<span style={{color:"#fca5a5"}}>📦 {fmt(ct.fecha_desmonte)}</span>}
                        </>)
                      })()}
                    </div>
                  </div>
                  <button onClick={()=>{
                    const ct=contratos.find((x:any)=>x.id===g.contrato_id)
                    if(ct) generarSplitsLocal(ct)
                  }} disabled={generandoLocal===g.contrato_id}
                    style={{padding:"5px 12px",borderRadius:8,border:"1px solid rgba(255,255,255,.2)",background:generandoLocal===g.contrato_id?"rgba(255,255,255,.1)":"rgba(255,255,255,.15)",color:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,whiteSpace:"nowrap" as const}}>
                    {generandoLocal===g.contrato_id?"...":"✂️ Re-generar"}
                  </button>
                  <button onClick={()=>abrirModalProv(g)}
                    style={{padding:"5px 12px",borderRadius:8,border:"1px solid rgba(255,255,255,.2)",background:"rgba(139,46,46,.4)",color:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,whiteSpace:"nowrap" as const}}>
                    📦 + Proveedor
                  </button>
                  <button onClick={()=>verContrato(g.contrato_id)}
                    style={{padding:"5px 12px",borderRadius:8,border:"1px solid rgba(255,255,255,.2)",background:contratoPanel?.id===g.contrato_id?"rgba(255,255,255,.2)":"transparent",color:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,whiteSpace:"nowrap" as const}}>
                    📄 Ver contrato
                  </button>
                </div>
                <div style={{padding:"8px 10px",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:8}}>
                  {(filtroEst==="todos"?g.splits:g.splits.filter((s:any)=>s.estado===filtroEst))
                    .sort((a:any,b:any)=>(SORT_ORD[a.tipo]??9)-(SORT_ORD[b.tipo]??9))
                    .map((s:any)=>{
                      const tipo=SPLIT_TIPOS.find((t:any)=>t.id===s.tipo)||SPLIT_TIPOS[0]
                      const est=SPLIT_ESTADOS.find((e:any)=>e.id===s.estado)||SPLIT_ESTADOS[0]
                      const nArts=(s.articulos||[]).length
                      const nPiezas=(s.articulos||[]).reduce((sum:number,a:any)=>sum+(a.cantidad||0),0)
                      const isActive=splitEdit?.id===s.id
                      return(
                        <div key={s.id} style={{border:`2px solid ${isActive?tipo.color:tipo.color+"44"}`,borderRadius:10,overflow:"hidden",background:isActive?tipo.bg:"#fff",cursor:"pointer"}}
                          onClick={()=>isActive?setSplitEdit(null):abrirEdicion(s)}>
                          <div style={{background:tipo.color,padding:"8px 12px",display:"flex",alignItems:"center",gap:8}}>
                            <span style={{fontSize:18}}>{tipo.icono}</span>
                            <div style={{flex:1}}>
                              <div style={{fontSize:12,fontWeight:700,color:"#fff"}}>{s.nombre||tipo.nombre}</div>
                              {s.responsable&&<div style={{fontSize:9,color:"rgba(255,255,255,.6)"}}>👤 {s.responsable}</div>}
                            </div>
                            <span style={{fontSize:9,padding:"2px 8px",borderRadius:8,background:est.bg,color:est.color,fontWeight:700,whiteSpace:"nowrap" as const}}>{est.label}</span>
                          </div>
                          <div style={{padding:"8px 12px"}}>
                            {nArts===0?(
                              <div style={{fontSize:11,color:"#c4bfb8",fontStyle:"italic",marginBottom:6}}>Sin artículos</div>
                            ):(
                              <div style={{marginBottom:6}}>
                                <div style={{fontSize:10,color:"#9a9590",marginBottom:3}}>{nArts} artículos · {nPiezas} piezas</div>
                                {(s.articulos||[]).slice(0,4).map((a:any,ai:number)=>(
                                  <div key={ai} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"1px 0",borderBottom:"1px solid #f5f4f0"}}>
                                    <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const,flex:1,color:"#4a4640"}}>{a.nombre}</span>
                                    <span style={{fontWeight:700,color:tipo.color,marginLeft:6,flexShrink:0}}>{a.cantidad}</span>
                                  </div>
                                ))}
                                {nArts>4&&<div style={{fontSize:10,color:"#9a9590",marginTop:2}}>+{nArts-4} más</div>}
                              </div>
                            )}
                            {s.notas&&<div style={{fontSize:10,color:"#4a4640",background:"#f8f6f2",borderRadius:5,padding:"4px 8px",marginBottom:6,fontStyle:"italic"}}>"{s.notas.slice(0,60)}{s.notas.length>60?"...":""}"</div>}
                            <div style={{display:"flex",gap:5,marginTop:6}}>
                              <button onClick={e=>{e.stopPropagation();abrirHojaSplit(s,logoUrl)}} style={{flex:1,padding:"5px",borderRadius:6,background:tipo.color,color:"#fff",border:"none",cursor:"pointer",fontSize:10,fontWeight:700}}>🖨️ Imprimir</button>
                              <button onClick={e=>{e.stopPropagation();isActive?setSplitEdit(null):abrirEdicion(s)}} style={{flex:1,padding:"5px",borderRadius:6,background:isActive?"#0f172a":"#f5f4f0",color:isActive?"#fff":"#4a4640",border:"none",cursor:"pointer",fontSize:10,fontWeight:700}}>{isActive?"✕ Cerrar":"✏️ Editar"}</button>
                              <button onClick={async e=>{e.stopPropagation();if(!window.confirm("¿Eliminar esta hoja?"))return;await fetch(`/api/splits?id=${s.id}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}});setSplits(prev=>prev.filter((x:any)=>x.id!==s.id));if(splitEdit?.id===s.id)setSplitEdit(null)}} style={{padding:"5px 7px",borderRadius:6,background:"#fdf0f0",color:"#8b2e2e",border:"none",cursor:"pointer",fontSize:12,fontWeight:700}} title="Eliminar">×</button>
                              <button onClick={async e=>{e.stopPropagation();if(!window.confirm("¿Eliminar esta hoja?"))return;await fetch(`/api/splits?id=${s.id}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}});setSplits(prev=>prev.filter((x:any)=>x.id!==s.id));if(isActive)setSplitEdit(null)}}
                                style={{width:26,padding:"5px",borderRadius:6,background:"#fdf0f0",color:"#8b2e2e",border:"none",cursor:"pointer",fontSize:12,fontWeight:700}} title="Eliminar hoja">🗑️</button>
                            </div>
                          </div>
                        </div>
                      )
                    })
                  }
                </div>
              </div>
            ))
          )}
        </div>

        {/* ── PANEL EDITAR SPLIT ── */}
        {splitEdit&&(()=>{
          const tipo=SPLIT_TIPOS.find((t:any)=>t.id===splitEdit.tipo)||SPLIT_TIPOS[0]
          return(
            <div style={{position:"sticky" as const,top:70,background:"#fff",border:`2px solid ${tipo.color}`,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 24px rgba(0,0,0,.12)"}}>
              <div style={{background:tipo.color,padding:"14px 16px",display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:22}}>{tipo.icono}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:15,fontWeight:800,color:"#fff"}}>{tipo.nombre}</div>
                  <div style={{fontSize:11,color:"rgba(255,255,255,.6)"}}>{splitEdit.cliente} · <span style={{fontFamily:"monospace",fontWeight:700,color:"#93c5fd"}}>{splitEdit.contrato_folio}</span></div>
                </div>
                <button onClick={()=>setSplitEdit(null)} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",width:28,height:28,borderRadius:"50%",cursor:"pointer",fontSize:14}}>✕</button>
              </div>
              <div style={{padding:16,display:"flex",flexDirection:"column" as const,gap:14,maxHeight:"75vh",overflowY:"auto" as const}}>
                <div style={{background:"#f8f6f2",borderRadius:8,padding:12,fontSize:12}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    <div>
                      <div style={{fontSize:9,color:"#9a9590",fontWeight:700,textTransform:"uppercase" as const}}>
                        {splitEdit.tipo==="desmonte"?"📦 Fecha de desmonte":"📅 Fecha"}
                      </div>
                      <input type="date" value={editFecha} onChange={e=>setEditFecha(e.target.value)}
                        style={{fontWeight:600,fontSize:12,border:"1.5px solid #e8e5de",borderRadius:6,padding:"4px 7px",marginTop:4,width:"100%",outline:"none",fontFamily:"Epilogue,sans-serif",boxSizing:"border-box" as const}}
                        onFocus={e=>e.target.style.borderColor=tipo.color}
                        onBlur={e=>e.target.style.borderColor="#e8e5de"}/>
                      {splitEdit.tipo==="desmonte"&&(
                        <div style={{display:"flex",gap:4,marginTop:5}}>
                          <button onClick={()=>{
                            if(editFechaOriginal) setEditFecha(editFechaOriginal)
                          }} style={{flex:1,fontSize:9,padding:"3px 6px",borderRadius:5,background:"#f5f4f0",border:"1px solid #e8e5de",cursor:"pointer",fontWeight:700,color:"#4a4640"}}>
                            Al término
                          </button>
                          <button onClick={()=>{
                            if(!editFechaOriginal)return
                            const d=new Date(editFechaOriginal+"T12:00:00");d.setDate(d.getDate()+1)
                            setEditFecha(d.toISOString().slice(0,10))
                          }} style={{flex:1,fontSize:9,padding:"3px 6px",borderRadius:5,background:"#f1f5f9",border:"1px solid #cbd5e1",cursor:"pointer",fontWeight:700,color:"#334155"}}>
                            +1 día
                          </button>
                        </div>
                      )}
                    </div>
                    {splitEdit.mostrar_direccion&&<div><div style={{fontSize:9,color:"#9a9590",fontWeight:700,textTransform:"uppercase" as const}}>Lugar</div><div style={{fontWeight:600,marginTop:2,fontSize:11}}>{splitEdit.lugar||"—"}</div></div>}
                  </div>
                </div>
                <div>
                  <div style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",marginBottom:8}}>Estado</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                    {SPLIT_ESTADOS.map(e=>(
                      <button key={e.id} onClick={()=>cambiarEstado(splitEdit,e.id)}
                        style={{padding:"7px",borderRadius:8,border:`2px solid ${e.color}`,background:splitEdit.estado===e.id?e.color:"#fff",color:splitEdit.estado===e.id?"#fff":e.color,fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,cursor:"pointer"}}>
                        {e.label}
                      </button>
                    ))}
                  </div>
                </div>
                <div>
                  <label style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",display:"block",marginBottom:6}}>Responsable</label>
                  <input value={editResp} onChange={e=>setEditResp(e.target.value)} placeholder="Nombre del responsable..."
                    style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}/>
                </div>
                {/* ── ARTÍCULOS DEL CONTRATO — CHECKLIST ── */}
                <div>
                  <div style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",marginBottom:6,display:"flex",justifyContent:"space-between" as const,alignItems:"center"}}>
                    <span>Artículos del contrato</span>
                    <span style={{fontSize:10,color:tipo.color,fontWeight:700}}>{editArts.length} seleccionados · {editArts.reduce((s:number,a:any)=>s+(Number(a.cantidad)||0),0)} pzas</span>
                  </div>
                  {/* Lista del CONTRATO como checklist */}
                  {contratoEditRef&&(contratoEditRef.articulos||[]).length>0?(
                    <div style={{border:`1.5px solid ${tipo.color}55`,borderRadius:8,overflow:"hidden",marginBottom:8,maxHeight:260,overflowY:"auto" as const}}>
                      {/* Header */}
                      <div style={{background:tipo.color,padding:"6px 10px",display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{fontSize:11,fontWeight:700,color:"#fff",flex:1}}>📋 Artículos del contrato — palomea los que van aquí</span>
                        <button onMouseDown={e=>{e.preventDefault();setEditArts((contratoEditRef.articulos||[]).map((a:any)=>({nombre:a.nombre,cantidad:a.cantidad||1,pu:a.pu||0,importe:a.importe||0,seccion:a.seccion||""})))}}
                          style={{fontSize:9,padding:"2px 8px",borderRadius:4,background:"rgba(255,255,255,.2)",border:"none",color:"#fff",cursor:"pointer",fontWeight:700}}>
                          ✓ Todos
                        </button>
                        <button onMouseDown={e=>{e.preventDefault();setEditArts([])}}
                          style={{fontSize:9,padding:"2px 8px",borderRadius:4,background:"rgba(255,255,255,.15)",border:"none",color:"rgba(255,255,255,.8)",cursor:"pointer",fontWeight:700}}>
                          ✗ Ninguno
                        </button>
                      </div>
                      {(contratoEditRef.articulos||[]).map((a:any,i:number)=>{
                        const selIdx=editArts.findIndex((x:any)=>x.nombre===a.nombre)
                        const seleccionado=selIdx>=0
                        const cantSel=seleccionado?editArts[selIdx].cantidad:a.cantidad||1
                        return(
                          <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderBottom:i<(contratoEditRef.articulos||[]).length-1?"1px solid #f0ece4":"none",background:seleccionado?tipo.bg:"#fff",cursor:"pointer",transition:"background .1s"}}
                            onClick={()=>{
                              if(seleccionado){setEditArts(prev=>prev.filter((_:any,j:number)=>j!==selIdx))}
                              else{setEditArts(prev=>[...prev,{nombre:a.nombre,cantidad:a.cantidad||1,pu:a.pu||0,importe:a.importe||0,seccion:a.seccion||""}])}
                            }}>
                            {/* Checkbox visual */}
                            <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${seleccionado?tipo.color:"#c4bfb8"}`,background:seleccionado?tipo.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"all .1s"}}>
                              {seleccionado&&<span style={{color:"#fff",fontSize:12,lineHeight:1,fontWeight:900}}>✓</span>}
                            </div>
                            {/* Nombre */}
                            <div style={{flex:1,fontSize:12,fontWeight:seleccionado?700:500,color:seleccionado?tipo.color:"#4a4640",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</div>
                            {/* Cantidad del contrato */}
                            <span style={{fontSize:10,color:"#9a9590",flexShrink:0,fontFamily:"monospace"}}>{a.cantidad}x</span>
                            {/* Cantidad para este split (si seleccionado) */}
                            {seleccionado&&(
                              <input type="number" min="1" max={a.cantidad} value={cantSel}
                                onClick={e=>e.stopPropagation()}
                                onChange={e=>{
                                  e.stopPropagation()
                                  const v=Math.max(1,Math.min(a.cantidad,parseInt(e.target.value)||1))
                                  setEditArts(prev=>prev.map((x:any,j:number)=>j===selIdx?{...x,cantidad:v}:x))
                                }}
                                style={{width:42,padding:"2px 4px",border:`1.5px solid ${tipo.color}`,borderRadius:5,textAlign:"center" as const,fontFamily:"monospace",fontSize:11,fontWeight:700,color:tipo.color,outline:"none",background:"#fff"}}/>
                            )}
                          </div>
                        )
                      })}
                    </div>
                  ):(
                    <div style={{padding:12,background:"#f8f6f2",borderRadius:8,fontSize:11,color:"#9a9590",marginBottom:8,textAlign:"center" as const}}>
                      Sin artículos en el contrato
                    </div>
                  )}
                </div>
                <div>
                  <label style={{fontSize:10,fontWeight:700,color:tipo.color,textTransform:"uppercase" as const,letterSpacing:".05em",display:"block",marginBottom:6}}>📋 Notas del área <span style={{fontWeight:400,color:"#9a9590"}}>(se imprimen)</span></label>
                  <textarea value={editNotas} onChange={e=>setEditNotas(e.target.value)} placeholder="Notas visibles para el equipo..." rows={3}
                    style={{width:"100%",padding:"8px 10px",border:`1.5px solid ${tipo.color}66`,borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",resize:"vertical" as const,boxSizing:"border-box" as const}}/>
                </div>
                <div>
                  <label style={{fontSize:10,fontWeight:700,color:"#64748b",textTransform:"uppercase" as const,letterSpacing:".05em",display:"block",marginBottom:6}}>🔒 Observaciones internas <span style={{fontWeight:400,color:"#9a9590"}}>(no se imprimen)</span></label>
                  <textarea value={editObs} onChange={e=>setEditObs(e.target.value)} placeholder="Notas internas, alertas..." rows={2}
                    style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",resize:"vertical" as const,boxSizing:"border-box" as const}}/>
                </div>
                <button onClick={guardarEdicion} disabled={saving}
                  style={{padding:"11px",borderRadius:8,background:saving?"#9a9590":tipo.color,color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:700}}>
                  {saving?"Guardando...":"💾 Guardar cambios"}
                </button>
                <button onClick={()=>abrirHojaSplit(splitEdit,logoUrl)}
                  style={{padding:"10px",borderRadius:8,background:"#f5f4f0",color:"#1a1814",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                  🖨️ Imprimir hoja de {tipo.nombre}
                </button>
              </div>
            </div>
          )
        })()}

        {/* ── PANEL VER CONTRATO ── */}
        {contratoPanel&&!splitEdit&&(()=>{
          const ct=contratoPanel
          const totalArts=(ct.articulos||[]).reduce((s:number,a:any)=>s+(a.cantidad||0),0)
          return(
            <div style={{position:"sticky" as const,top:70,background:"#fff",border:"2px solid #0f172a",borderRadius:12,overflow:"hidden",boxShadow:"0 4px 24px rgba(0,0,0,.12)"}}>
              <div style={{background:"#0f172a",padding:"14px 16px",display:"flex",alignItems:"center",gap:10}}>
                <div style={{flex:1}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800,color:"#fff"}}>{ct.cliente||ct.archivo}</div>
                  <div style={{display:"flex",gap:8,alignItems:"center",marginTop:3,flexWrap:"wrap" as const}}>
                    <span style={{fontFamily:"monospace",fontSize:11,color:"#60a5fa",fontWeight:700}}>{ct.folio||ct.archivo}</span>
                    {ct.vendedor&&<span style={{fontSize:10,color:"rgba(255,255,255,.5)"}}>👤 {ct.vendedor}</span>}
                  </div>
                </div>
                <button onClick={()=>setContratoPanel(null)} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",width:28,height:28,borderRadius:"50%",cursor:"pointer",fontSize:14}}>✕</button>
              </div>
              <div style={{maxHeight:"75vh",overflowY:"auto" as const}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",borderBottom:"1px solid #e8e5de"}}>
                  {[{l:"Evento",v:ct.fecha_evento,ico:"🎉"},{l:"Entrega",v:ct.fecha_entrega,ico:"🚚"},{l:"Desmonte",v:ct.fecha_desmonte,ico:"📦"}].map((k,i)=>(
                    <div key={i} style={{padding:"10px 8px",textAlign:"center" as const,borderRight:i<2?"1px solid #e8e5de":"none"}}>
                      <div style={{fontSize:15,marginBottom:2}}>{k.ico}</div>
                      <div style={{fontSize:11,fontWeight:700}}>{k.v||"—"}</div>
                      <div style={{fontSize:8,color:"#9a9590",textTransform:"uppercase" as const}}>{k.l}</div>
                    </div>
                  ))}
                </div>
                <div style={{padding:"10px 14px",borderBottom:"1px solid #e8e5de"}}>
                  {ct.lugar&&<div style={{fontSize:12,color:"#4a4640",marginBottom:3}}>📍 {ct.lugar}</div>}
                  {ct.tel&&<div style={{fontSize:12,color:"#4a4640"}}>📞 {ct.tel}</div>}
                </div>
                <div style={{padding:"10px 14px"}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".05em",marginBottom:8}}>Artículos — {totalArts} piezas</div>
                  {(ct.articulos||[]).length===0?(
                    <div style={{fontSize:11,color:"#c4bfb8",fontStyle:"italic"}}>Sin artículos</div>
                  ):(ct.articulos||[]).map((a:any,i:number)=>(
                    <div key={i} style={{display:"flex",gap:8,padding:"6px 0",borderBottom:"1px solid #f5f4f0",fontSize:11}}>
                      <span style={{fontWeight:700,color:"#1a3a5c",minWidth:28,textAlign:"right" as const}}>{a.cantidad}x</span>
                      <span style={{flex:1}}>{a.nombre}</span>
                      <span style={{fontSize:10,color:"#9a9590",flexShrink:0,fontStyle:"italic"}}>{a.seccion||a.categoria||""}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )
        })()}
      </div>
      </>}
    </div>
  )
}



// ─── STYLES ───────────────────────────────────────────────────────────
const S: any = {
  loginWrap:{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f5f4f0"},
  loginBox:{background:"#fff",border:"1px solid #e8e5de",borderRadius:16,padding:36,width:360,display:"flex",flexDirection:"column" as const,gap:12,boxShadow:"0 12px 40px rgba(26,24,20,.12)"},
  loginTitle:{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,textAlign:"center" as const},
  loginSub:{fontSize:11,color:"#9a9590",letterSpacing:".08em",textTransform:"uppercase" as const,textAlign:"center" as const,marginBottom:8},
  loginBtn:{padding:"10px 20px",background:"#1a1814",color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"Epilogue,sans-serif"},
  input:{padding:"8px 12px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,background:"#fff",color:"#1a1814",width:"100%",margin:"0 0 4px",outline:"none"},
  topbar:{position:"sticky" as const,top:0,zIndex:100,background:"rgba(245,244,240,.95)",backdropFilter:"blur(16px)",borderBottom:"1px solid #e8e5de",padding:"0 32px",height:56,display:"flex",alignItems:"center",justifyContent:"space-between",gap:12},
  brandMark:{width:30,height:30,border:"1.5px solid #d4cfc4",borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,background:"#fff"},
  brandName:{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700},
  brandSub:{fontSize:10,color:"#9a9590",letterSpacing:".07em",textTransform:"uppercase" as const},
  nav:{display:"none"},
  navBtn:{padding:"5px 10px",borderRadius:6,border:"none",background:"transparent",color:"#9a9590",fontFamily:"Epilogue,sans-serif",fontSize:11.5,fontWeight:500,cursor:"pointer"},
  navBtnActive:{background:"#1a1814",color:"#fff",fontWeight:600},
  chip:{background:"#fff",border:"1px solid #e8e5de",borderRadius:20,padding:"4px 11px",fontFamily:"monospace",fontSize:10.5,color:"#9a9590"},
  iconBtn:{display:"flex",alignItems:"center",gap:5,padding:"6px 12px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",color:"#4a4640",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:500,cursor:"pointer"},
  kstrip:{display:"grid",gridTemplateColumns:"repeat(6,1fr)",background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden",marginBottom:20,boxShadow:"0 1px 3px rgba(26,24,20,.06)"},
  kpi:{padding:"16px 18px",borderRight:"1px solid #e8e5de"},
  kv:{fontFamily:"Playfair Display,serif",fontSize:32,fontWeight:800,lineHeight:1,marginBottom:3},
  kl:{fontSize:10,fontWeight:700,textTransform:"uppercase" as const,letterSpacing:".07em",color:"#9a9590"},
  ks:{fontSize:10,color:"#c4bfb8",marginTop:2},
  fset:{display:"flex",background:"#fff",border:"1px solid #e8e5de",borderRadius:8,overflow:"hidden"},
  fbtn:{padding:"6px 12px",border:"none",borderRight:"1px solid #e8e5de",background:"transparent",color:"#9a9590",fontFamily:"Epilogue,sans-serif",fontSize:11.5,fontWeight:500,cursor:"pointer"},
  fbtnActive:{background:"#1a1814",color:"#fff",fontWeight:600},
  dzbtn1:{display:"inline-block",padding:"8px 20px",borderRadius:8,fontSize:12.5,fontWeight:600,fontFamily:"Epilogue,sans-serif",cursor:"pointer",background:"#1a1814",color:"#fff",border:"none"},
}

const globalCSS=`
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Epilogue',sans-serif;background:#f5f4f0;-webkit-tap-highlight-color:transparent}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:#f5f4f0}
::-webkit-scrollbar-thumb{background:#d4cfc4;border-radius:10px}
input,select,button{font-size:16px!important}
@media(max-width:768px){
  table{display:block;overflow-x:auto;-webkit-overflow-scrolling:touch}
  .grid-2{grid-template-columns:1fr!important}
  .grid-3{grid-template-columns:1fr!important}
  .grid-4{grid-template-columns:1fr 1fr!important}
}
@media print{.topbar,button,select,input[type=text],input[type=date],input[type=file]{display:none!important}}
  @media print{body{-webkit-filter:grayscale(100%);filter:grayscale(100%)}img{-webkit-filter:grayscale(0%)!important;filter:grayscale(0%)!important}}
`
