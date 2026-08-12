// BUILD: 1785892031
import { useState, useEffect, useCallback } from "react"
import Head from "next/head"
import * as XLSX from "xlsx"

interface Usuario { id: string; nombre: string; email: string; rol: string }
interface Persona { id: string; nombre: string; color_bg: string; color_bd: string; color_fg: string }
interface Articulo { nombre: string; cantidad: number; pu: number; importe: number; seccion: string }
interface Pago { fecha: string; monto: number; nota: string; metodo?: string; folio?: string }
interface Contrato {
  id: string; archivo: string; cliente: string; lugar: string; tel: string
  fecha_evento: string; fecha_entrega: string; fecha_desmonte: string
  dia_evento: string; dia_entrega: string; dia_desmonte: string
  estado_entrega: string; estado_desmonte: string
  asig_entrega: string[]; asig_desmonte: string[]
  checklist: { txt: string; done: boolean }[]
  notas: string; es_duplicado: boolean; carpeta: string
  articulos: Articulo[]
  folio: string
  vendedor: string
  tipo: string
  total: number
  a_cuenta: number
  cobrado: number
  pagos: Pago[]
}
interface Ruta {
  id: string; nombre: string; fecha: string; tipo: string
  unidades: { id: string; nombre: string; color: string; asignados: string[] }[]
  contratos_ids: string[]
  asignaciones: Record<string,string>
}

const DIAS = ["DOMINGO","LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO"]
const MESES = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
const MESES_F = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
const ESTADOS: Record<string,{label:string;bg:string;color:string;border:string}> = {
  pend:   {label:"⏳ Pendiente",bg:"#fdf5e8",color:"#92580a",border:"#f0d49a"},
  camino: {label:"🚛 En camino",bg:"#edf3fa",color:"#1a3a5c",border:"#b8ceea"},
  listo:  {label:"✅ Entregado", bg:"#edf7f2",color:"#2d6a4f",border:"#b7deca"},
  recog:  {label:"📦 Recogido",  bg:"#f3eef8",color:"#4a2d6e",border:"#c9b8e8"},
}
const COLORES_PERSONA = [
  {bg:"#edf7f2",bd:"#b7deca",fg:"#2d6a4f"},{bg:"#edf3fa",bd:"#b8ceea",fg:"#1a3a5c"},
  {bg:"#fdf5e8",bd:"#f0d49a",fg:"#92580a"},{bg:"#f3eef8",bd:"#c9b8e8",fg:"#4a2d6e"},
  {bg:"#edf7f5",bd:"#9adace",fg:"#1a5c52"},{bg:"#fdf0f0",bd:"#e8b8b8",fg:"#8b2e2e"},
]
const UNIDAD_COLORES = ["#2d6a4f","#1a3a5c","#92580a","#4a2d6e","#8b2e2e","#1a5c52","#6e4a2d"]
const UNIDADES_DEFAULT = ["LD-36-696","LD-36-352","LF-43068","LF-63073","LG-90-921","ESTAQUITA"]
const FINANZAS_PWD = "POLIFLOR_GERENCIA_2026"
const fmt = (n:number) => "$"+Math.round(n||0).toLocaleString()

function toJSDate(v: any): Date|null {
  if (!v) return null
  if (v instanceof Date) return new Date(v)
  if (typeof v === "number") return new Date(Math.round((v-25569)*86400*1000))
  const d = new Date(v); return isNaN(d.getTime()) ? null : d
}
function addDays(d:Date,n:number):Date{const r=new Date(d);r.setDate(r.getDate()+n);return r}
function isoDate(d:Date):string{const m=String(d.getMonth()+1).padStart(2,"0");return `${d.getFullYear()}-${m}-${String(d.getDate()).padStart(2,"0")}`}
function fmtDate(s:string):string{if(!s)return "—";const d=new Date(s+"T12:00:00");return `${d.getDate()} ${MESES[d.getMonth()]} ${d.getFullYear()}`}
function sameDay(a:Date|string,b:Date|string):boolean{
  const da=typeof a==="string"?new Date(a+"T12:00:00"):a
  const db=typeof b==="string"?new Date(b+"T12:00:00"):b
  return da.getFullYear()===db.getFullYear()&&da.getMonth()===db.getMonth()&&da.getDate()===db.getDate()
}

async function parsearExcel(file:File):Promise<Contrato|null>{
  return new Promise(resolve=>{
    const reader=new FileReader()
    reader.onload=e=>{
      try{
        const bytes=new Uint8Array(e.target!.result as ArrayBuffer)
        const wb=XLSX.read(bytes,{type:"array",cellDates:true})
        let wsName=wb.SheetNames.find(s=>s.trim().toUpperCase()==="EVENTO")
        if(!wsName)wsName=wb.SheetNames[0]
        const ws=wb.Sheets[wsName]
        if(!ws){resolve(null);return}
        let fe:Date|null=null
        // First try known cells
        for(const addr of["C3","B3","D3","C2","B2","E3","F3","C4","D4","B4","E2","F2"]){
          const cell=ws[addr]
          if(cell?.v){const d=toJSDate(cell.v);if(d&&d.getFullYear()>2020){fe=d;break}}
        }
        // If not found, scan ALL cells for a date in range 2024-2030
        if(!fe){
          const allAddrs=Object.keys(ws).filter(k=>!k.startsWith("!"))
          for(const addr of allAddrs){
            const cell=ws[addr]
            if(cell?.v){
              const d=toJSDate(cell.v)
              if(d&&d.getFullYear()>=2024&&d.getFullYear()<=2030){fe=d;break}
            }
          }
        }
        // Last resort: use today
        if(!fe){fe=new Date()}
        let cliente=""
        const allKeys=Object.keys(ws).filter(k=>!k.startsWith("!"))
        for(const addr of allKeys){
          const cell=ws[addr]
          if(cell?.v&&String(cell.v).trim().toLowerCase().replace(/\s/g,"").startsWith("cliente")){
            const m=addr.match(/^([A-Z]+)(\d+)$/)
            if(m){const rc=ws[String.fromCharCode(m[1].charCodeAt(0)+1)+m[2]];if(rc?.v&&String(rc.v).trim().length>1){cliente=String(rc.v).trim();break}}
          }
        }
        if(!cliente){for(const addr of["D1","C1","D2","C2"]){const cell=ws[addr];if(cell?.v&&String(cell.v).trim().length>1&&!String(cell.v).toLowerCase().includes("cliente")){cliente=String(cell.v).trim();break}}}
        const lugar=ws["C4"]?.v?String(ws["C4"].v).trim():""
        const tel=ws["C6"]?.v?String(ws["C6"].v).trim():""
        const ent=addDays(fe,-1),des=addDays(fe,1)
        // Vendedor (seller code) from E6 — e.g. "L", "AL", "K/H"
        let vendedor_default=""
        const vendedorCell=ws["E6"]||ws["F6"]||ws["E5"]
        if(vendedorCell?.v){
          const vt=String(vendedorCell.v).trim()
          // Only use as vendedor if it looks like initials (short, no numbers, no "folio" keyword)
          if(vt.length<=6 && !vt.match(/\d{3,}/) && !vt.toUpperCase().includes("FOLIO") && !vt.toUpperCase().includes("TEL"))
            vendedor_default=vt
        }
        // Folio — scan all cells for "FOLIO" label then read adjacent cell
        let folio=""
        const allK=Object.keys(ws).filter(k=>!k.startsWith("!"))
        for(const addr of allK){
          const cell=ws[addr]
          if(!cell?.v) continue
          const sv=String(cell.v).trim().toUpperCase()
          // Look for cell with "FOLIO" as label
          if(sv==="FOLIO"||sv.startsWith("FOLIO:")||sv.startsWith("FOLIO :")){
            // Value is in same cell after colon, or in next column
            const colon=String(cell.v).indexOf(":")
            if(colon>=0){
              const after=String(cell.v).slice(colon+1).trim()
              if(after.length>0){folio=after;break}
            }
            // Try next column same row
            const m=addr.match(/^([A-Z]+)(\d+)$/)
            if(m){
              const nextCol=String.fromCharCode(m[1].charCodeAt(0)+1)
              const nextCell=ws[nextCol+m[2]]
              if(nextCell?.v){folio=String(nextCell.v).trim();break}
            }
          }
          // Also check for "No." or "NUM" pattern near top of sheet
          if((sv.startsWith("NO.")||sv==="NO"||sv==="NUM"||sv==="NÚMERO")&&Number(addr.replace(/[A-Z]/g,""))<=10){
            const m=addr.match(/^([A-Z]+)(\d+)$/)
            if(m){
              const nextCol=String.fromCharCode(m[1].charCodeAt(0)+1)
              const nextCell=ws[nextCol+m[2]]
              if(nextCell?.v&&String(nextCell.v).trim().length>0&&String(nextCell.v).trim().length<=20){
                folio=String(nextCell.v).trim();break
              }
            }
          }
        }
        // If no folio found, use a sequential from filename
        if(!folio){
          const numMatch=file.name.match(/(\d{4,})/);
          if(numMatch) folio=numMatch[1]
        }
        // Tipo from filename
        let tipo="contrato"
        const nArch=file.name.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"")
        if(nArch.includes("DECLIN"))tipo="declinado"
        else if(nArch.includes("COTIZ"))tipo="cotizacion"
        // Articulos
        const articulos:Articulo[]=[]
        let seccionActual="General"
        const range=XLSX.utils.decode_range(ws["!ref"]||"A1")
        for(let r=range.s.r;r<=range.e.r;r++){
          const cA=ws[XLSX.utils.encode_cell({r,c:0})]
          const cB=ws[XLSX.utils.encode_cell({r,c:1})]
          const cC=ws[XLSX.utils.encode_cell({r,c:3})]
          const cD=ws[XLSX.utils.encode_cell({r,c:4})]
          if((!cA?.v||typeof cA.v!=="number")&&cB?.v){
            const txt=String(cB.v).trim().toUpperCase()
            if(!txt.includes("CANT")&&!txt.includes("DESCRIPCI")&&txt.length>3&&!txt.includes("TOTAL")&&!txt.includes("RESTANTE"))seccionActual=String(cB.v).trim()
          }
          if(cA?.v&&typeof cA.v==="number"&&cA.v>0&&cA.v<10000&&cB?.v){
            const desc=String(cB.v).trim()
            if(desc.length>2&&!desc.toUpperCase().includes("CANT."))
              articulos.push({nombre:desc,cantidad:Math.round(cA.v),pu:typeof cC?.v==="number"?cC.v:0,importe:typeof cD?.v==="number"?cD.v:0,seccion:seccionActual})
          }
        }
        // TOTAL y AC - buscar en todas las celdas
        let total=0,a_cuenta=0
        for(const addr of allKeys){
          const cell=ws[addr]
          if(!cell?.v)continue
          const cellTxt=String(cell.v).toUpperCase().trim()
          if(cellTxt==="TOTAL"||cellTxt==="TOTAL:"){
            const m=addr.match(/^([A-Z]+)(\d+)$/)
            if(m){for(const col of["E","F","D","G"]){const vc=ws[col+m[2]];if(vc?.v&&typeof vc.v==="number"&&vc.v>0){total=vc.v;break}}}
          }
          if(cellTxt==="AC"||cellTxt==="AC:"||cellTxt==="A CUENTA"||cellTxt==="ANTICIPO"){
            const m=addr.match(/^([A-Z]+)(\d+)$/)
            if(m){for(const col of["E","F","D","G"]){const vc=ws[col+m[2]];if(vc?.v&&typeof vc.v==="number"&&vc.v>0){a_cuenta=vc.v;break}}}
          }
        }
        resolve({
          id:Date.now()+"_"+Math.random().toString(36).slice(2,6),
          archivo:file.name,cliente:cliente||file.name.replace(/_dividido/g,"").replace(".xlsx",""),
          lugar,tel,folio,vendedor:vendedor_default,tipo,total,a_cuenta,cobrado:a_cuenta,pagos:[],
          fecha_evento:isoDate(fe),fecha_entrega:isoDate(ent),fecha_desmonte:isoDate(des),
          dia_evento:DIAS[fe.getDay()],dia_entrega:DIAS[ent.getDay()],dia_desmonte:DIAS[des.getDay()],
          estado_entrega:"pend",estado_desmonte:"pend",asig_entrega:[],asig_desmonte:[],
          checklist:[],notas:"",es_duplicado:false,carpeta:"",articulos,
        })
      }catch(e){
        console.error("parsearExcel error:",e)
        // Try to return minimal data so file doesn't fail completely
        resolve(null)
      }
    }
    reader.readAsArrayBuffer(file)
  })
}

async function apiCall(url:string,method:string,body?:any,token?:string){
  const res=await fetch(url,{method,headers:{"Content-Type":"application/json",...(token?{Authorization:`Bearer ${token}`}:{})},body:body?JSON.stringify(body):undefined})
  return res.json()
}

export default function Dashboard(){
  const [user,setUser]=useState<Usuario|null>(null)
  const [token,setToken]=useState("")
  const [loginForm,setLoginForm]=useState({email:"admin@poliflor.com",password:""})
  const [loginError,setLoginError]=useState("")
  const [contratos,setContratos]=useState<Contrato[]>([])
  const [personal,setPersonal]=useState<Persona[]>([])
  const [seccion,setSeccion]=useState("inicio")
  const [subTab,setSubTab]=useState("")
  const [sidebarOpen,setSidebarOpen]=useState(true)
  const [isMobile,setIsMobile]=useState(false)
  const [logoUrl,setLogoUrl]=useState<string>("/logo.png")
  const [logoLoaded,setLogoLoaded]=useState(false)

  // Load logo from Supabase on mount (works on ALL devices)
  useEffect(()=>{
    fetch("/api/config?clave=logo")
      .then(r=>r.json())
      .then(data=>{
        if(data.valor){
          setLogoUrl(data.valor)
          if(typeof window!=="undefined") localStorage.setItem("pf_logo",data.valor)
        } else if(typeof window!=="undefined"){
          const local=localStorage.getItem("pf_logo")
          if(local) setLogoUrl(local)
        }
        setLogoLoaded(true)
      })
      .catch(()=>{
        if(typeof window!=="undefined"){
          const local=localStorage.getItem("pf_logo")
          if(local) setLogoUrl(local)
        }
        setLogoLoaded(true)
      })
  },[])
  const esAdmin = user?.rol==="admin"
  const vendedorActual = VENDEDORES.find(v=>v.nombre===user?.nombre)?.nombre || ""

  // Detect mobile on mount and resize
  useEffect(()=>{
    if(typeof window==="undefined") return
    const check=()=>{
      const mobile=window.innerWidth<768
      setIsMobile(mobile)
      setSidebarOpen(window.innerWidth>=768)
    }
    check()
    window.addEventListener("resize",check)
    return ()=>window.removeEventListener("resize",check)
  },[])
  const [filtro,setFiltro]=useState("todos")
  const [filtroTipo,setFiltroTipo]=useState("contrato")
  const [filtroAnoGlobal,setFiltroAnoGlobal]=useState(String(new Date().getFullYear()))
  const [filtroFolio,setFiltroFolio]=useState("")
  const [rangoDesde,setRangoDesde]=useState("")
  const [rangoHasta,setRangoHasta]=useState("")
  const [busqueda,setBusqueda]=useState("")
  const [semanaOffset,setSemanaOffset]=useState(0)
  const [artSemOffset,setArtSemOffset]=useState(0)
  const [artBusqueda,setArtBusqueda]=useState("")
  const [teamFiltro,setTeamFiltro]=useState("semana")
  const [nuevaPersona,setNuevaPersona]=useState("")
  const [ganttDesde,setGanttDesde]=useState(()=>isoDate(addDays(new Date(),-5)))
  const [ganttHasta,setGanttHasta]=useState(()=>isoDate(addDays(new Date(),35)))
  const [rutaTipo,setRutaTipo]=useState("entrega")
  const [rutaFecha,setRutaFecha]=useState("")
  const [rutas,setRutas]=useState<Ruta[]>([])
  const [showNuevaRuta,setShowNuevaRuta]=useState(false)
  const [checkInputs,setCheckInputs]=useState<Record<string,string>>({})
  const [expandedArts,setExpandedArts]=useState<Record<string,boolean>>({})
  const [syncMsg,setSyncMsg]=useState("")
  const [finanzasPwd,setFinanzasPwd]=useState(false)
  const [finanzasPwdInput,setFinanzasPwdInput]=useState("")
  const [finanzasPwdError,setFinanzasPwdError]=useState("")
  const [pagoModal,setPagoModal]=useState<{cid:string;monto:string;nota:string}|null>(null)
  const [pagoMonto,setPagoMonto]=useState("")
  const [pagoNota,setPagoNota]=useState("")

  const cargar=useCallback(async(tk:string)=>{
    const[cr,pr]=await Promise.all([apiCall("/api/contratos","GET",undefined,tk),apiCall("/api/personal","GET",undefined,tk)])
    if(Array.isArray(cr))setContratos(cr)
    if(Array.isArray(pr))setPersonal(pr)
  },[])

  useEffect(()=>{
    const tk=localStorage.getItem("pf_token")
    const usr=localStorage.getItem("pf_user")
    if(tk&&usr){setToken(tk);setUser(JSON.parse(usr));setBusqueda("");cargar(tk)}
  },[cargar])

  // Auto-refresh cada 2 minutos
  useEffect(()=>{
    if(!token)return
    const iv=setInterval(()=>cargar(token),120000)
    return ()=>clearInterval(iv)
  },[token,cargar])

  async function login(){
    setLoginError("")
    const data=await apiCall("/api/auth/login","POST",loginForm)
    if(data.error){setLoginError(data.error);return}
    localStorage.setItem("pf_token",data.token)
    localStorage.setItem("pf_user",JSON.stringify(data.user))
    setToken(data.token);setUser(data.user);setBusqueda("");cargar(data.token)
  }
  function logout(){localStorage.removeItem("pf_token");localStorage.removeItem("pf_user");setToken("");setUser(null);setContratos([]);setPersonal([])}

  async function importarArchivos(files:FileList){
    const arr=Array.from(files).filter(f=>f.name.toLowerCase().endsWith(".xlsx")&&!f.name.startsWith("~"))
    if(!arr.length){setSyncMsg("No hay archivos .xlsx válidos — selecciona archivos .xlsx");return}
    setSyncMsg(`Leyendo ${arr.length} archivo(s)...`)
    let nuevos=0,actualizados=0,errores=0
    for(const file of arr){
      setSyncMsg(`📂 Parseando: ${file.name}`)
      let parsed:any=null
      try{
        parsed=await parsearExcel(file)
        console.log("Parsed:",file.name,"->",parsed?.cliente,parsed?.fecha_evento,parsed?.articulos?.length,"arts")
      }catch(ex){
        console.error("parsearExcel threw:",ex)
        setSyncMsg(`❌ Error leyendo ${file.name}: ${ex}`)
        errores++;continue
      }
      if(!parsed){
        console.error("parsearExcel null:",file.name)
        setSyncMsg(`⚠️ No se pudo leer ${file.name} — verifica el formato`)
        errores++;continue
      }
      setSyncMsg(`💾 Guardando: ${parsed.cliente||file.name}...`)
      const exists=contratos.find((x:Contrato)=>x.archivo===parsed.archivo)
      if(exists){
        const res=await apiCall(`/api/contratos?id=${exists.id}`,"PATCH",{
          cliente:parsed.cliente,lugar:parsed.lugar,telefono:parsed.tel||"",
          fecha_evento:parsed.fecha_evento,fecha_entrega:parsed.fecha_entrega,fecha_desmonte:parsed.fecha_desmonte,
          dia_evento:parsed.dia_evento,dia_entrega:parsed.dia_entrega,dia_desmonte:parsed.dia_desmonte,
          tipo:parsed.tipo,folio:parsed.folio,vendedor:parsed.vendedor||"",
          total:parsed.total||0,a_cuenta:parsed.a_cuenta||0,
          articulos:parsed.articulos||[],
          cobrado:exists.cobrado>0?exists.cobrado:(parsed.a_cuenta||0),
          asig_entrega:exists.asig_entrega,asig_desmonte:exists.asig_desmonte,
          estado_entrega:exists.estado_entrega,estado_desmonte:exists.estado_desmonte,
          checklist:exists.checklist,pagos:exists.pagos||[],
        },token)
        if(res?.error){
          console.error("PATCH error:",res.error)
          setSyncMsg(`❌ Error guardando ${file.name}: ${res.error}`)
          errores++
        }else{
          setContratos(cs=>cs.map(x=>x.id===exists.id?{...x,...parsed,articulos:parsed.articulos||[]}:x))
          actualizados++
        }
      }else{
        const saved=await apiCall("/api/contratos","POST",parsed,token)
        if(saved?.error){
          console.error("POST error:",saved.error)
          setSyncMsg(`❌ Error creando ${file.name}: ${saved.error}`)
          errores++
        }else if(saved?.id||saved?.cliente){
          nuevos++
        }else{
          console.error("POST unexpected:",saved)
          errores++
        }
      }
    }
    const msg=`✓ ${nuevos} nuevos, ${actualizados} actualizados${errores>0?`, ${errores} errores`:""}`
    setSyncMsg(msg)
    setTimeout(()=>cargar(token),500)
  }
  async function actualizarContrato(id:string,updates:any){
    if(updates===null){
      // null = eliminar
      setContratos(cs=>cs.filter(c=>String(c.id)!==String(id)))
      return
    }
    setContratos(cs=>cs.map(c=>c.id===id?{...c,...updates}:c))
    await apiCall(`/api/contratos?id=${id}`,"PATCH",updates,token)
  }

  async function agregarPersona(){
    if(!nuevaPersona.trim())return
    const color=COLORES_PERSONA[personal.length%COLORES_PERSONA.length]
    const p=await apiCall("/api/personal","POST",{nombre:nuevaPersona.trim(),color_bg:color.bg,color_bd:color.bd,color_fg:color.fg},token)
    if(p?.id){setPersonal(ps=>[...ps,p]);setNuevaPersona("")}
  }

  // ================================================================
  // FILTROS — reescrito completo, sin dependencias externas
  // ================================================================
  const hoyStr = isoDate(new Date())
  const hoy0 = new Date(); hoy0.setHours(0,0,0,0)
  const en30 = new Date(hoy0); en30.setDate(hoy0.getDate()+30)
  // semana actual lunes-domingo
  const _dow = hoy0.getDay()===0 ? 6 : hoy0.getDay()-1
  const _lunes = new Date(hoy0); _lunes.setDate(hoy0.getDate()-_dow)
  const _dom = new Date(_lunes); _dom.setDate(_lunes.getDate()+6); _dom.setHours(23,59,59,999)

  // paso 1: filtrar por tipo (aplica a TODAS las secciones)
  // Enrich contratos with inferred vendedor from folio
  const contratosEnriquecidos: Contrato[] = contratos.map((x:Contrato)=>({
    ...x,
    vendedor: x.vendedor||vendedorDesdeFolio(x.folio||"")
  }))

  const cBase: Contrato[] = contratosEnriquecidos.filter((x:Contrato) => {
    if(filtroAnoGlobal && x.fecha_evento?.slice(0,4)!==filtroAnoGlobal) return false
    if(filtroTipo==="todos") return true
    return (x.tipo||"contrato") === filtroTipo
  })

  // paso 2: filtrar para la lista de contratos (aplica filtros adicionales)
  const lista: Contrato[] = cBase.filter((x:Contrato) => {
    // filtro período
    if(filtro==="hoy") {
      if(x.fecha_entrega!==hoyStr && x.fecha_evento!==hoyStr && x.fecha_desmonte!==hoyStr) return false
    }
    if(filtro==="semana") {
      const fe = new Date(x.fecha_evento+"T12:00:00")
      const fen = new Date(x.fecha_entrega+"T12:00:00")
      const fd = new Date(x.fecha_desmonte+"T12:00:00")
      const inSem = (d:Date) => d>=_lunes && d<=_dom
      if(!inSem(fe) && !inSem(fen) && !inSem(fd)) return false
    }
    if(filtro==="entrega") {
      const fen = new Date(x.fecha_entrega+"T12:00:00")
      if(fen < hoy0 || fen > en30) return false
    }
    if(filtro==="pend") {
      const tieneEntrega = (x.asig_entrega||[]).length > 0
      const tieneDesmonte = (x.asig_desmonte||[]).length > 0
      if(tieneEntrega && tieneDesmonte) return false
    }
    // rango fechas
    if(rangoDesde && x.fecha_evento < rangoDesde) return false
    if(rangoHasta && x.fecha_evento > rangoHasta) return false
    // folio
    if(filtroFolio.trim()) {
      if(!(x.folio||"").toLowerCase().includes(filtroFolio.toLowerCase())) return false
    }
    // búsqueda
    if(busqueda) {
      const q = busqueda.toLowerCase()
      const ok = (x.cliente||"").toLowerCase().includes(q)
        || (x.lugar||"").toLowerCase().includes(q)
        || (x.tel||"").toLowerCase().includes(q)
        || (x.folio||"").toLowerCase().includes(q)
        || Boolean(x.articulos?.some((a:Articulo)=>a.nombre.toLowerCase().includes(q)))
      if(!ok) return false
    }
    return true
  }).sort((a:Contrato,b:Contrato)=>a.fecha_evento.localeCompare(b.fecha_evento))

  // KPIs — conteos globales (sin filtro tipo) para las tarjetas de resumen
  const kpis = {
    contratos:    contratos.filter((x:Contrato)=>(x.tipo||"contrato")==="contrato").length,
    cotizaciones: contratos.filter((x:Contrato)=>x.tipo==="cotizacion").length,
    declinados:   contratos.filter((x:Contrato)=>x.tipo==="declinado").length,
    // Los siguientes respetan filtroTipo para mostrar relevancia operativa
    hoy:        cBase.filter((x:Contrato)=>x.fecha_entrega===hoyStr||x.fecha_evento===hoyStr||x.fecha_desmonte===hoyStr).length,
    entregas:   cBase.filter((x:Contrato)=>{const d=new Date(x.fecha_entrega+"T12:00:00");return d>=hoy0&&d<=en30}).length,
    sinAsignar: cBase.filter((x:Contrato)=>!(x.asig_entrega||[]).length||!(x.asig_desmonte||[]).length).length,
    listos:     cBase.filter((x:Contrato)=>x.estado_entrega==="listo"&&x.estado_desmonte==="recog").length,
  }

  function getRango(offset:number){
    if(offset===-1) return null
    const h=new Date(); h.setHours(0,0,0,0)
    const dow=h.getDay(),diff=dow===0?-6:1-dow
    const lunes=addDays(h,diff+offset*7)
    const domingo=addDays(lunes,6); domingo.setHours(23,59,59,999)
    return {desde:lunes,hasta:domingo}
  }

  if(!user)return (
    <div style={S.loginWrap}>
      <Head><title>Poliflor — Login</title><link rel="icon" type="image/png" href="/favicon.png"/><link rel="shortcut icon" href="/favicon.png"/><link rel="apple-touch-icon" href="/favicon.png"/></Head>
      <div style={S.loginBox}>
        <div style={{fontSize:36,textAlign:"center" as const}}>🪑</div>
        <div style={S.loginTitle}>Renta de Mobiliario</div>
        <div style={S.loginSub}>POLIFLOR DASHBOARD</div>
        <input style={S.input} type="email" placeholder="Email" value={loginForm.email} onChange={e=>setLoginForm(f=>({...f,email:e.target.value}))}/>
        <input style={S.input} type="password" placeholder="Contraseña" value={loginForm.password} onChange={e=>setLoginForm(f=>({...f,password:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&login()}/>
        {loginError&&<div style={{background:"#fdf0f0",border:"1px solid #e8b8b8",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#8b2e2e"}}>{loginError}</div>}
        <button style={S.loginBtn} onClick={login}>Entrar</button>
      </div>
    </div>
  )

  return (
    <div style={{fontFamily:"Epilogue,sans-serif",background:"#0f172a",minHeight:"100vh",color:"#1a1814",display:"flex"}}>
      <Head>
        <title>Poliflor Dashboard</title>
        <link rel="icon" type="image/png" href={logoUrl||"/favicon.png"}/>
        <link rel="shortcut icon" href={logoUrl||"/favicon.png"}/>
        <link rel="apple-touch-icon" href={logoUrl||"/favicon.png"}/>
        <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Epilogue:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
        <style>{globalCSS}</style>
      </Head>

      {/* ══ SIDEBAR ══ */}
      {/* Mobile overlay */}
      {isMobile&&sidebarOpen&&(
        <div onClick={()=>setSidebarOpen(false)}
          style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.5)",zIndex:199,backdropFilter:"blur(2px)"}}/>
      )}
      <aside style={{
        width:220,minWidth:220,background:"#0f172a",
        display:"flex",flexDirection:"column" as const,
        position:"fixed" as const,top:0,left:0,bottom:0,zIndex:200,
        borderRight:"1px solid rgba(255,255,255,.06)",
        transform:sidebarOpen?"translateX(0)":"translateX(-220px)",
        transition:"transform .25s cubic-bezier(.4,0,.2,1)",
      }}>
        {/* Logo + toggle */}
        <div style={{padding:"16px 14px 14px",borderBottom:"1px solid rgba(255,255,255,.06)",display:"flex",alignItems:"center",gap:8}}>
          <img src={logoUrl} alt="Poliflor" style={{height:36,width:"auto",objectFit:"contain" as const,filter:"brightness(0) invert(1)",opacity:.9,flex:1}}
            onError={(e:any)=>{e.target.style.display="none"}}/>
          <button onClick={()=>setSidebarOpen(false)}
            style={{width:28,height:28,borderRadius:7,border:"none",background:"rgba(255,255,255,.08)",color:"rgba(255,255,255,.5)",cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
            ✕
          </button>
        </div>

        {/* Nav items */}
        <nav style={{flex:1,overflowY:"auto" as const,padding:"10px 8px"}}>
          {([
            {id:"inicio",label:"Panel",icon:"home"},
            {id:"ventas",label:"Ventas",icon:"clip",subs:[["cotizaciones","Cotizaciones"],["contratos-conf","Contratos"]]},
            {id:"planeacion",label:"Planeación",icon:"cal",subs:[["agenda","Agenda"],["carga","🚚 Carga"],["gantt","Gantt"],["dias","Por Día"]]},
            {id:"catalogo",label:"Catálogo",icon:"box",subs:[["cat-articulos","Artículos"],["cat-clientes","Clientes"],["cat-busqueda","Buscar"]]},
            {id:"inventario",label:"Inventario",icon:"inv"},
            {id:"splits",label:"Splits",icon:"cut"},
            ...(esAdmin?[{id:"finanzas",label:"Finanzas",icon:"money"},{id:"rh",label:"RH",icon:"people"}]:[]),
            {id:"config",label:"Config",icon:"cog",subs:[["cfg-equipo","Equipo"],["cfg-rutas","Rutas"],["cfg-misrutas","Mis Rutas"],["cfg-logo","🖼️ Logo"],["cfg-password","🔑 Contraseña"]]},
          ] as {id:string,label:string,icon:string,subs?:string[][]}[]).map(item=>{
            const isActive=seccion===item.id
            const hasSubs=item.subs&&item.subs.length>0
            const ICONS:Record<string,string>={
              home:"M10.707 2.293a1 1 0 00-1.414 0l-7 7a1 1 0 001.414 1.414L4 10.414V17a1 1 0 001 1h2a1 1 0 001-1v-2a1 1 0 011-1h2a1 1 0 011 1v2a1 1 0 001 1h2a1 1 0 001-1v-6.586l.293.293a1 1 0 001.414-1.414l-7-7z",
              clip:"M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01",
              cal:"M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z",
              box:"M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4",
              inv:"M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01",
              money:"M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z",
              cut:"M6 2a2 2 0 00-2 2v1H3a1 1 0 000 2h1v1a2 2 0 002 2h.172l3.5 6.063A3 3 0 1014 17a3 3 0 00-2.465 1.294L8.086 12H9a2 2 0 002-2V9h1a1 1 0 000-2h-1V6a2 2 0 00-2-2H6z",
              cog:"M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z",
            }
            return(
              <div key={item.id}>
                <button onClick={()=>{
                  if((window as any).__cotFormDirty){
                    const ok=window.confirm("¿Salir sin guardar?\n\nTienes una cotización en progreso con datos sin guardar.\n\nAcepta para salir · Cancela para seguir editando")
                    if(!ok) return
                    ;(window as any).__cotFormDirty = false
                  }
                  setSeccion(item.id)
                  if(!hasSubs) setSubTab("")
                  if(isMobile) setSidebarOpen(false)
                }}
                  style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:10,border:"none",cursor:"pointer",background:isActive?"#2563eb":"transparent",color:isActive?"#fff":"rgba(255,255,255,.6)",fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:isActive?600:400,textAlign:"left" as const,transition:"all .15s",marginBottom:2}}>
                  <svg width="18" height="18" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}>
                    <path d={ICONS[item.icon]||ICONS.clip}/>
                  </svg>
                  <span style={{flex:1}}>{item.label}</span>
                  {hasSubs&&<span style={{fontSize:10,opacity:.5,transition:"transform .2s",transform:isActive?"rotate(180deg)":"rotate(0deg)"}}>▾</span>}
                </button>
                {hasSubs&&isActive&&(
                  <div style={{marginLeft:16,marginBottom:4,borderLeft:"1px solid rgba(255,255,255,.1)",paddingLeft:12}}>
                    {item.subs!.map(([s,l])=>{
                      const isSubActive=subTab===s||(subTab===""&&s===item.subs![0][0])
                      return(
                        <button key={s} onClick={()=>{setSubTab(s);if(isMobile)setSidebarOpen(false)}}
                          style={{width:"100%",display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:8,border:"none",cursor:"pointer",background:isSubActive?"rgba(255,255,255,.1)":"transparent",color:isSubActive?"#fff":"rgba(255,255,255,.45)",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:isSubActive?600:400,textAlign:"left" as const,marginBottom:1,transition:"all .12s"}}>
                          <span style={{width:5,height:5,borderRadius:"50%",background:isSubActive?"#60a5fa":"rgba(255,255,255,.2)",flexShrink:0}}/>
                          {l}
                        </button>
                      )
                    })}
                  </div>
                )}
              </div>
            )
          })}
        </nav>

        {/* User + logout */}
        <div style={{padding:"12px 14px",borderTop:"1px solid rgba(255,255,255,.06)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <div style={{width:30,height:30,borderRadius:"50%",background:"#2563eb",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:12,color:"#fff",flexShrink:0}}>
              {(user?.nombre||"U").charAt(0).toUpperCase()}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:12,fontWeight:600,color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{user?.nombre}</div>
              <div style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>{esAdmin?"Administrador":"Vendedor"}</div>
            </div>
          </div>
          <button onClick={logout} style={{width:"100%",padding:"7px",borderRadius:8,background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",color:"rgba(255,255,255,.5)",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:500}}>
            Cerrar sesión
          </button>
        </div>
      </aside>

      {/* ══ MAIN CONTENT ══ */}
      <div style={{marginLeft:sidebarOpen&&!isMobile?220:0,flex:1,minHeight:"100vh",background:"#f5f4f0",display:"flex",flexDirection:"column" as const,transition:"margin-left .25s cubic-bezier(.4,0,.2,1)"}}>
        {/* Top bar */}
        <div style={{background:"#fff",borderBottom:"1px solid #e8e5de",padding:"0 16px",height:52,display:"flex",alignItems:"center",gap:10,position:"sticky" as const,top:0,zIndex:100}}>
          {/* Hamburger / toggle */}
          <button onClick={()=>setSidebarOpen(v=>!v)}
            style={{width:36,height:36,borderRadius:8,border:"none",background:"#f5f4f0",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,color:"#1a1814",fontSize:16,transition:"background .15s"}}
            title={sidebarOpen?"Ocultar menú":"Mostrar menú"}>
            {sidebarOpen?"☰":"☰"}
          </button>
          {/* Logo en topbar — siempre renderizado, visible cuando sidebar cerrado */}
          <img
            src={logoUrl}
            alt="Poliflor"
            style={{
              height:32,width:"auto",objectFit:"contain" as const,
              maxWidth:(!sidebarOpen||isMobile)?120:0,
              opacity:(!sidebarOpen||isMobile)?1:0,
              overflow:"hidden",
              flexShrink:0,
              transition:"all .25s ease",
            }}
            onError={(e:any)=>{e.target.style.display="none"}}/>
          {/* KPIs rápidos — solo en desktop */}
          <div style={{display:isMobile?"none":"flex",gap:8,alignItems:"center",flex:1}}>
            <div style={{display:"flex",gap:6,alignItems:"center",background:"#f5f4f0",borderRadius:8,padding:"4px 10px"}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:"#2d6a4f",display:"inline-block"}}/>
              <span style={{fontSize:11,fontWeight:700,color:"#1a1814"}}>{kpis.contratos}</span>
              <span style={{fontSize:10,color:"#9a9590"}}>contratos</span>
            </div>
            <div style={{display:"flex",gap:6,alignItems:"center",background:"#f5f4f0",borderRadius:8,padding:"4px 10px"}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:"#1a3a5c",display:"inline-block"}}/>
              <span style={{fontSize:11,fontWeight:700,color:"#1a1814"}}>{kpis.cotizaciones}</span>
              <span style={{fontSize:10,color:"#9a9590"}}>cotizaciones</span>
            </div>
            <div style={{display:"flex",gap:6,alignItems:"center",background:"#f5f4f0",borderRadius:8,padding:"4px 10px"}}>
              <span style={{fontSize:10,color:"#9a9590"}}>Hoy:</span>
              <span style={{fontSize:11,fontWeight:700,color:"#1a1814"}}>{kpis.hoy}</span>
            </div>
          </div>
          {/* Año filter — oculto en mobile */}
          {!isMobile&&seccion!=="finanzas"&&(()=>{
            const anos=[...new Set(contratos.map((x:Contrato)=>x.fecha_evento?.slice(0,4)).filter(Boolean))].sort().reverse() as string[]
            return(
              <div style={{display:"flex",alignItems:"center",gap:5,background:"#f5f4f0",borderRadius:8,padding:"4px 10px"}}>
                <span style={{fontSize:10,fontWeight:700,color:"#1a1814"}}>📅</span>
                <select value={filtroAnoGlobal} onChange={e=>setFiltroAnoGlobal(e.target.value)}
                  style={{border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,color:"#1a1814",outline:"none",cursor:"pointer"}}>
                  <option value="">Todos</option>
                  {anos.map(a=><option key={a} value={a}>{a}</option>)}
                </select>
              </div>
            )
          })()}
          {/* Tipo filter — oculto en mobile */}
          {!isMobile&&seccion!=="finanzas"&&seccion!=="inicio"&&(
            <div style={{display:"flex",gap:3}}>
              {[["todos","Todos"],["contrato","Contratos"],["cotizacion","Cot."],["declinado","Dec."]] .map(([f,l])=>(
                <button key={f} onClick={()=>setFiltroTipo(f)}
                  style={{padding:"4px 10px",borderRadius:20,border:`1px solid ${filtroTipo===f?"#1a1814":"#e8e5de"}`,background:filtroTipo===f?"#1a1814":"#fff",color:filtroTipo===f?"#fff":"#4a4640",fontSize:10,fontWeight:filtroTipo===f?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>
                  {l}
                </button>
              ))}
            </div>
          )}
          {/* Search — oculto en mobile */}
          {!isMobile&&<input placeholder="🔍 Buscar..." value={busqueda} onChange={e=>{setBusqueda(e.target.value);if(e.target.value)setSeccion("contratos")}}
            style={{padding:"5px 12px",border:"1px solid #e8e5de",borderRadius:20,background:"#fff",fontFamily:"Epilogue,sans-serif",fontSize:11,width:160,outline:"none",color:"#1a1814"}}/>}
        </div>

        {/* SyncMsg */}
        {syncMsg&&<div style={{background:"#edf7f2",borderBottom:"1px solid #b7deca",padding:"6px 24px",fontSize:12,color:"#2d6a4f",fontWeight:500}}>{syncMsg}</div>}

        {/* Main content */}
        <div style={{padding:isMobile?"14px 12px 80px":"24px 28px 60px",flex:1}}>


        {/* ── SUBTABS BAR ── */}
        {(seccion==="ventas"||seccion==="planeacion"||seccion==="catalogo"||seccion==="config")&&(
          <div style={{display:"flex",gap:4,marginBottom:20,borderBottom:"2px solid #e8e5de",paddingBottom:0}}>
            {(
              seccion==="ventas"?[["cotizaciones","📋 Cotizaciones"],["contratos-conf","📄 Contratos"]] :
              seccion==="planeacion"?[["agenda","📅 Agenda"],["carga","🚚 Carga"],["gantt","Gantt"],["dias","Por Día"]] :
              seccion==="catalogo"?[["cat-articulos","📦 Artículos"],["cat-clientes","👥 Clientes"],["cat-busqueda","🔍 Buscar"]] :
              [["cfg-equipo","👥 Equipo"],["cfg-rutas","🗺️ Rutas"],["cfg-misrutas","Mis Rutas"],["cfg-logo","🖼️ Logo"]]
            ).map(([s,l])=>{
              const isActive = subTab===s || (subTab===""&&s===
                (seccion==="ventas"?"cotizaciones":
                 seccion==="planeacion"?"agenda":
                 seccion==="catalogo"?"cat-articulos":"cfg-equipo"))
              return(
                <button key={s} onClick={()=>setSubTab(s)}
                  style={{padding:"10px 18px",border:"none",borderBottom:`2.5px solid ${isActive?"#1a1814":"transparent"}`,background:"transparent",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:isActive?700:400,color:isActive?"#1a1814":"#9a9590",marginBottom:-2,whiteSpace:"nowrap" as const}}>
                  {l}
                </button>
              )
            })}
          </div>
        )}

        {/* ── INICIO ── */}
        {seccion==="inicio"&&(
          <div>
            {/* Dropzone carga Excel */}
            <div style={{background:"#fff",border:"1.5px dashed #d4cfc4",borderRadius:14,padding:20,textAlign:"center" as const,marginBottom:16}}
              onDragOver={e=>e.preventDefault()}
              onDrop={e=>{
                e.preventDefault()
                const items=Array.from(e.dataTransfer.items||[])
                const allFiles:File[]=[]
                const process=(entry:any):Promise<void>=>new Promise(res=>{
                  if(entry.isFile){entry.file((f:File)=>{if(f.name.endsWith(".xlsx")&&!f.name.startsWith("~"))allFiles.push(f);res()})}
                  else if(entry.isDirectory){const reader=entry.createReader();reader.readEntries((entries:any[])=>Promise.all(entries.map(process)).then(()=>res()))}
                  else res()
                })
                Promise.all(items.map(i=>process(i.webkitGetAsEntry()))).then(()=>{
                  if(allFiles.length){const dt=new DataTransfer();allFiles.forEach(f=>dt.items.add(f));importarArchivos(dt.files)}
                  else if(e.dataTransfer.files.length)importarArchivos(e.dataTransfer.files)
                })
              }}>
              <div style={{fontSize:28,opacity:.3,marginBottom:6}}>📂</div>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:4}}>Arrastra archivos o carpetas de contratos</div>
              <div style={{fontSize:11,color:"#9a9590",marginBottom:10}}>Archivos .xlsx individuales o carpeta completa</div>
              <div style={{display:"flex",gap:10,justifyContent:"center"}}>
                <label style={{...S.dzbtn1,cursor:"pointer"}}>
                  <input type="file" multiple style={{display:"none"}} onChange={e=>{
                    if(!e.target.files)return
                    const xlsx=Array.from(e.target.files).filter((f:File)=>f.name.toLowerCase().endsWith(".xlsx")&&!f.name.startsWith("~"))
                    if(xlsx.length){const dt=new DataTransfer();xlsx.forEach((f:File)=>dt.items.add(f));importarArchivos(dt.files)}
                    else setSyncMsg("No se encontraron archivos .xlsx válidos")
                  }}/>
                  📄 Seleccionar archivos
                </label>
                <label style={{...S.dzbtn1,background:"#fff",color:"#1a1814",border:"1.5px solid #1a1814",cursor:"pointer"}}>
                  <input type="file" style={{display:"none"}} onChange={e=>{
                    if(!e.target.files)return
                    const xlsx=Array.from(e.target.files).filter(f=>f.name.endsWith(".xlsx")&&!f.name.startsWith("~"))
                    if(xlsx.length){const dt=new DataTransfer();xlsx.forEach(f=>dt.items.add(f));importarArchivos(dt.files)}
                  }} {...{webkitdirectory:"",directory:""} as any}/>
                  📁 Seleccionar carpeta
                </label>
              </div>
            </div>
            <InicioSection contratos={contratosEnriquecidos} esAdmin={esAdmin} vendedorActual={vendedorActual} token={token}/>
          </div>
        )}

        {/* ── VENTAS ── */}
        {seccion==="ventas"&&(subTab===""||subTab==="cotizaciones")&&(
          <CotizacionesSection token={token} personal={personal} logoUrl={logoUrl} vendedorActual={vendedorActual} esAdmin={esAdmin} contratos={cBase}/>
        )}
        {seccion==="ventas"&&subTab==="contratos-conf"&&(
          <ContratosConfirmadosSection token={token} contratos={cBase} onActualizar={actualizarContrato} isMobile={isMobile} vendedorActual={vendedorActual} esAdmin={esAdmin}/>
        )}

        {/* ── PLANEACIÓN ── */}
        {seccion==="planeacion"&&(subTab===""||subTab==="agenda")&&(
          <AgendaSection contratos={cBase}/>
        )}
        {seccion==="planeacion"&&subTab==="carga"&&(
          <CargaSection contratos={cBase}/>
        )}
        {seccion==="planeacion"&&subTab==="gantt"&&(
          <GanttSection contratos={cBase} desde={ganttDesde} hasta={ganttHasta} setDesde={setGanttDesde} setHasta={setGanttHasta} personal={personal}/>
        )}
        {seccion==="planeacion"&&subTab==="dias"&&(
          <DiasSection contratos={cBase} semanaOffset={semanaOffset} setSemanaOffset={setSemanaOffset} getRango={getRango}/>
        )}

        {/* ── CATÁLOGO ── */}
        {seccion==="catalogo"&&(subTab===""||subTab==="cat-articulos")&&(
          <CatalogoSection token={token}/>
        )}
        {seccion==="catalogo"&&subTab==="cat-clientes"&&(
          <ClientesSection contratos={contratos}/>
        )}
        {seccion==="catalogo"&&subTab==="cat-busqueda"&&(
          <BusquedaSection contratos={contratos}/>
        )}

        {/* ── FINANZAS (solo admin) ── */}
        {seccion==="finanzas"&&esAdmin&&(
          <FinanzasSection
            contratos={contratos} token={token} pwd={finanzasPwd} pwdInput={finanzasPwdInput}
            setPwdInput={setFinanzasPwdInput} pwdError={finanzasPwdError}
            pagoModal={pagoModal} setPagoModal={setPagoModal}
            pagoMonto={pagoMonto} setPagoMonto={setPagoMonto}
            pagoNota={pagoNota} setPagoNota={setPagoNota}
            onUnlock={()=>{
              if(finanzasPwdInput===FINANZAS_PWD){setFinanzasPwd(true);setFinanzasPwdError("")}
              else setFinanzasPwdError("Contraseña incorrecta")
            }}
            onAgregarPago={(cid:string,monto:number,nota:string,metodo?:string,folio?:string,vendedor?:string)=>{
              if(nota.startsWith("__SET_TOTAL__:")){
                const parts=nota.split(":")
                const nuevoTotal=parseFloat(parts[1])||0
                const nuevoACuenta=parseFloat(parts[2])||0
                // Recalculate cobrado from existing pagos + new a_cuenta
                const contratoActual=contratos.find((x:Contrato)=>x.id===cid)
                const sumaPageos=(contratoActual?.pagos||[]).reduce((s:number,p:any)=>s+(p.monto||0),0)
                const nuevoCobrado=Math.max(sumaPageos,nuevoACuenta)
                actualizarContrato(cid,{total:nuevoTotal,a_cuenta:nuevoACuenta,cobrado:nuevoCobrado})
              } else {
                const fecha=new Date().toISOString().split("T")[0]
                const nuevoPago={fecha,monto,nota,metodo:metodo||"efectivo",folio:folio||"",vendedor:vendedor||""}
                actualizarContrato(cid,{pagos:[...(contratos.find((x:Contrato)=>x.id===cid)?.pagos||[]),nuevoPago],cobrado:(contratos.find((x:Contrato)=>x.id===cid)?.cobrado||0)+monto})
              }
              setPagoModal(null);setPagoMonto("");setPagoNota("")
            }}
          />
        )}

        {/* ── CONFIG ── */}
        {seccion==="inventario"&&<InventarioSection contratos={cBase} token={token}/>}
        {seccion==="splits"&&(
          <SplitsSection token={token} contratos={contratosEnriquecidos} logoUrl={logoUrl}/>
        )}
        {seccion==="rh"&&esAdmin&&(
          <div style={{padding:"20px"}}>
            <RHSection token={token} isMobile={isMobile}/>
          </div>
        )}
        {seccion==="config"&&(subTab===""||subTab==="cfg-equipo")&&esAdmin&&(
          <EquipoSection
            personal={personal} contratos={cBase} user={user}
            nuevaPersona={nuevaPersona} setNuevaPersona={setNuevaPersona}
            onAgregar={agregarPersona}
            teamFiltro={teamFiltro} setTeamFiltro={setTeamFiltro}
          />
        )}
        {seccion==="config"&&subTab==="cfg-rutas"&&(
          <RutasSection contratos={cBase} semanaOffset={semanaOffset} setSemanaOffset={setSemanaOffset}/>
        )}
        {seccion==="config"&&subTab==="cfg-misrutas"&&(
          <MisRutasSection token={token}/>
        )}
        {seccion==="config"&&subTab==="cfg-logo"&&(
          <LogoSection logoUrl={logoUrl} setLogoUrl={(u:string)=>setLogoUrl(u)}/>
        )}
        {seccion==="config"&&subTab==="cfg-password"&&(
          <CambiarPasswordSection token={token} user={user}/>
        )}
        </div>

      {/* ══ MOBILE BOTTOM NAV ══ */}
      {isMobile&&(
        <div style={{position:"fixed" as const,bottom:0,left:0,right:0,background:"#0f172a",borderTop:"1px solid rgba(255,255,255,.1)",display:"flex",zIndex:300}}>
          {([
            {id:"inicio",ico:"🏠",label:"Inicio"},
            {id:"ventas",ico:"📋",label:"Ventas"},
            {id:"planeacion",ico:"📅",label:"Plan"},
            {id:"catalogo",ico:"📦",label:"Catálogo"},
            {id:"inventario",ico:"📦",label:"Inv."},
          {id:"splits",ico:"✂️",label:"Splits"},
          ] as {id:string,ico:string,label:string}[]).map(item=>{
            const isActive=seccion===item.id
            return(
              <button key={item.id} onClick={()=>{setSeccion(item.id);setSubTab("");setSidebarOpen(false)}}
                style={{flex:1,display:"flex",flexDirection:"column" as const,alignItems:"center",justifyContent:"center",padding:"8px 2px 10px",border:"none",background:"transparent",cursor:"pointer",color:isActive?"#60a5fa":"rgba(255,255,255,.4)"}}>
                <span style={{fontSize:18,lineHeight:1}}>{item.ico}</span>
                <span style={{fontSize:9,fontWeight:isActive?700:400,marginTop:2,fontFamily:"Epilogue,sans-serif"}}>{item.label}</span>
                {isActive&&<div style={{width:14,height:2,background:"#2563eb",borderRadius:1,marginTop:2}}/>}
              </button>
            )
          })}
        </div>
      )}
      </div>
    </div>
  )
}


// ─── CONTRATO CARD ────────────────────────────────────────────────────
function ContratoCard({c,personal,busqueda,expandedArt,onToggleArt,checkInput,onCheckInput,onToggleAsig,onEstado,onCheck,onAddCheck,onDelete}: any){
  return (
    <div style={{background:"#fff",border:`1px solid ${c.es_duplicado?"#e8b8b8":"#e8e5de"}`,borderRadius:12,display:"flex",flexDirection:"column" as const,overflow:"hidden",boxShadow:"0 1px 3px rgba(26,24,20,.06)"}}>
      <div style={{padding:"14px 16px 12px",borderBottom:"1px solid #e8e5de"}}>
        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:8,marginBottom:10}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,lineHeight:1.3}}>{c.cliente||c.archivo}</div>
          <div style={{display:"flex",gap:5,alignItems:"center",flexShrink:0}}>
            {c.folio&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:8,background:"#fafaf8",border:"1px solid #e8e5de",color:"#9a9590",fontFamily:"monospace"}}>#{c.folio}</span>}
            {(c.tipo==="cotizacion")&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:8,fontWeight:800,background:"#edf3fa",color:"#1a3a5c",border:"1px solid #b8ceea"}}>COTIZ.</span>}
            {(c.tipo==="declinado")&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:8,fontWeight:800,background:"#fdf0f0",color:"#8b2e2e",border:"1px solid #e8b8b8"}}>DECLIN.</span>}
            {(!c.tipo||c.tipo==="contrato")&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:8,fontWeight:800,background:"#edf7f2",color:"#2d6a4f",border:"1px solid #b7deca"}}>CONTRATO</span>}
            {c.es_duplicado&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:10,fontWeight:800,background:"#8b2e2e",color:"#fff"}}>DUP</span>}
            <button onClick={onDelete} style={{background:"none",border:"none",color:"#c4bfb8",cursor:"pointer",fontSize:14,padding:0}}>✕</button>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
          {[{lbl:"🚚 Entrega",fecha:c.fecha_entrega,dia:c.dia_entrega,bg:"#edf7f2",bd:"#b7deca",fg:"#2d6a4f"},
            {lbl:"📅 Evento",fecha:c.fecha_evento,dia:c.dia_evento,bg:"#fdf5e8",bd:"#f0d49a",fg:"#92580a"},
            {lbl:"🔧 Desmonte",fecha:c.fecha_desmonte,dia:c.dia_desmonte,bg:"#fdf0f0",bd:"#e8b8b8",fg:"#8b2e2e"}].map((d,i)=>(
            <div key={i} style={{borderRadius:8,padding:"7px 9px",border:`1px solid ${d.bd}`,background:d.bg}}>
              <div style={{fontSize:9,fontWeight:700,textTransform:"uppercase" as const,letterSpacing:".07em",marginBottom:2,color:d.fg}}>{d.lbl}</div>
              <div style={{fontFamily:"monospace",fontSize:10.5,fontWeight:500}}>{fmtDate(d.fecha)}</div>
              <div style={{fontSize:10,color:"#4a4640",marginTop:1,fontWeight:600}}>{d.dia}</div>
            </div>
          ))}
        </div>
      </div>
      {/* Finanzas */}
      {(c.total||0)>0&&(
        <div style={{padding:"6px 16px",borderBottom:"1px solid #e8e5de",background:"#fafaf8",display:"flex",gap:10,alignItems:"center",flexWrap:"wrap" as const}}>
          <span style={{fontSize:11,fontWeight:700,color:"#1a1814"}}>Total: {fmt(c.total)}</span>
          <span style={{fontSize:11,color:"#9a9590"}}>·</span>
          <span style={{fontSize:11,color:"#1a3a5c"}}>AC: {fmt(c.a_cuenta)}</span>
          <span style={{fontSize:11,color:"#9a9590"}}>·</span>
          <span style={{fontSize:11,fontWeight:700,color:(c.total||0)-(c.cobrado||0)>0?"#92580a":"#2d6a4f"}}>
            {(c.total||0)-(c.cobrado||0)>0?`Saldo: ${fmt((c.total||0)-(c.cobrado||0))}`:"✓ Liquidado"}
          </span>
        </div>
      )}
      {/* Asignaciones */}
      <div style={{padding:"10px 16px",borderTop:"1px solid #e8e5de",background:"#fafaf8",display:"flex",flexDirection:"column" as const,gap:8}}>
        {personal.length>0?[{tipo:"asig_entrega",estTipo:"estado_entrega",label:"Entrega"},{tipo:"asig_desmonte",estTipo:"estado_desmonte",label:"Desmonte"}].map(({tipo,estTipo,label})=>(
          <div key={tipo}>
            <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" as const,marginBottom:4}}>
              <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase" as const,letterSpacing:".06em",color:"#9a9590",width:60,flexShrink:0}}>{label}</div>
              <div style={{display:"flex",gap:5,flexWrap:"wrap" as const,flex:1}}>
                {personal.map((p:Persona)=>{
                  const on=(c[tipo]||[]).includes(p.id)
                  return <span key={p.id} onClick={()=>onToggleAsig(tipo,p.id)} style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",borderRadius:20,fontSize:11,fontWeight:600,border:`1px solid ${on?p.color_bd:"#e8e5de"}`,background:on?p.color_bg:"#fff",color:on?p.color_fg:"#4a4640",cursor:"pointer"}}>{p.nombre}</span>
                })}
              </div>
            </div>
            <div style={{display:"flex",gap:4,flexWrap:"wrap" as const,paddingLeft:68}}>
              {Object.entries(ESTADOS).map(([k,e])=>{
                const isAct=c[estTipo]===k
                return <span key={k} onClick={()=>onEstado(estTipo,k)} style={{padding:"3px 9px",borderRadius:20,fontSize:10.5,fontWeight:700,border:`1px solid ${e.border}`,background:e.bg,color:e.color,cursor:"pointer",boxShadow:isAct?`0 0 0 2px ${e.color}`:undefined}}>{e.label}</span>
              })}
            </div>
          </div>
        )):<span style={{fontSize:11,color:"#c4bfb8"}}>Ve a Equipo para agregar colaboradores</span>}
      </div>
      {/* Checklist */}
      <div style={{padding:"10px 16px",borderTop:"1px solid #e8e5de",background:"#fafaf8"}}>
        <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase" as const,letterSpacing:".06em",color:"#9a9590",marginBottom:6}}>Checklist {c.checklist?.length?`(${c.checklist.filter((x:any)=>x.done).length}/${c.checklist.length})`:""}</div>
        {(c.checklist||[]).map((item:any,idx:number)=>(
          <label key={idx} style={{display:"flex",alignItems:"center",fontSize:12,cursor:"pointer",padding:"2px 0",textDecoration:item.done?"line-through":"none",color:item.done?"#9a9590":"inherit"}}>
            <input type="checkbox" checked={item.done} onChange={()=>onCheck(idx)} style={{marginRight:6}}/>
            {item.txt}
          </label>
        ))}
        <div style={{display:"flex",gap:6,marginTop:6}}>
          <input value={checkInput} onChange={e=>onCheckInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onAddCheck()} placeholder="Agregar item..." style={{flex:1,padding:"5px 9px",border:"1px solid #e8e5de",borderRadius:6,fontSize:11.5,fontFamily:"Epilogue,sans-serif",background:"#fff",outline:"none"}}/>
          <button onClick={onAddCheck} style={{padding:"5px 10px",borderRadius:6,background:"#fafaf8",border:"1px solid #e8e5de",fontSize:11.5,cursor:"pointer"}}>+</button>
        </div>
      </div>
      {/* Artículos */}
      {c.articulos?.length>0&&(
        <div>
          <div style={{padding:"6px 16px",borderTop:"1px solid #e8e5de"}}>
            <button onClick={onToggleArt} style={{background:"none",border:"none",fontSize:12,color:"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:600,padding:"2px 0"}}>
              📦 {expandedArt?"Ocultar":"Ver"} {c.articulos.length} artículos {expandedArt?"▲":"▼"}
            </button>
          </div>
          {expandedArt&&(
            <div style={{padding:"10px 16px",borderTop:"1px solid #e8e5de",background:"#fafaf8",maxHeight:280,overflowY:"auto" as const}}>
              {c.articulos.map((a:Articulo,i:number)=>(
                <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"3px 0",borderBottom:"1px solid rgba(232,229,222,.3)"}}>
                  <span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#92580a",minWidth:28,textAlign:"right" as const}}>{a.cantidad}</span>
                  <span style={{fontSize:11.5,flex:1,color:"#4a4640"}}>{a.nombre}</span>
                  {a.importe>0&&<span style={{fontFamily:"monospace",fontSize:10.5,color:"#2d6a4f",fontWeight:600}}>${a.importe.toLocaleString()}</span>}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
      {/* Footer */}
      <div style={{padding:"8px 16px",display:"flex",gap:8,alignItems:"center",borderTop:"1px solid #e8e5de"}}>
        <span style={{fontSize:11,color:"#4a4640",flex:1,whiteSpace:"nowrap" as const,overflow:"hidden",textOverflow:"ellipsis"}}>{c.lugar?`📍 ${c.lugar}${c.tel?` · 📞 ${c.tel}`:""}`:""}</span>
        <span style={{fontFamily:"monospace",fontSize:9.5,color:"#9a9590",background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:4,padding:"2px 6px",overflow:"hidden",textOverflow:"ellipsis",maxWidth:180}}>{c.archivo}</span>
      </div>
    </div>
  )
}

// ─── AGENDA SEMANAL ────────────────────────────
function AgendaSection({contratos}:{contratos:Contrato[]}){
  const [semOff,setSemOff]=useState(0)
  const DIAS=["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]
  const DIAS_FULL=["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
  const TIPO_COLOR:Record<string,string>={contrato:"#1a3a5c",cotizacion:"#2d6a4f",declinado:"#8b2e2e"}
  const TIPO_BG:Record<string,string>={contrato:"#edf3fa",cotizacion:"#edf7f2",declinado:"#fdf0f0"}
  const MOV_COLOR:Record<string,string>={entrega:"#92580a",desmonte:"#4a2d6e",evento:"#1a3a5c"}
  const MOV_BG:Record<string,string>={entrega:"#fdf5e8",desmonte:"#f5f0fc",evento:"#edf3fa"}

  // Compute week range
  const hoy=new Date();hoy.setHours(0,0,0,0)
  const dow=hoy.getDay()===0?6:hoy.getDay()-1
  const lunes=new Date(hoy);lunes.setDate(hoy.getDate()-dow+semOff*7)
  const dias=Array.from({length:7},(_,i)=>{const d=new Date(lunes);d.setDate(lunes.getDate()+i);return d})
  const isoD=(d:Date)=>d.toISOString().slice(0,10)

  // Build event map per day
  type Mov={tipo:"entrega"|"desmonte"|"evento";contrato:Contrato}
  const byDay:Record<string,Mov[]>={}
  dias.forEach(d=>{byDay[isoD(d)]=[]})
  contratos.forEach((x:Contrato)=>{
    if(x.tipo==="declinado"||x.tipo==="cotizacion") return
    const fe=x.fecha_entrega,fd=x.fecha_evento,fdes=x.fecha_desmonte
    if(byDay[fe]!==undefined) byDay[fe].push({tipo:"entrega",contrato:x})
    if(byDay[fd]!==undefined) byDay[fd].push({tipo:"evento",contrato:x})
    if(byDay[fdes]!==undefined) byDay[fdes].push({tipo:"desmonte",contrato:x})
  })

  const mesAno=lunes.toLocaleDateString("es-MX",{month:"long",year:"numeric"})
  const esHoy=(d:Date)=>isoD(d)===isoD(new Date())

  return(
    <div>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"12px 18px"}}>
        <div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,textTransform:"capitalize" as const}}>{mesAno}</div>
          <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>
            {lunes.toLocaleDateString("es-MX",{day:"numeric",month:"short"})} — {dias[6].toLocaleDateString("es-MX",{day:"numeric",month:"short"})}
            <span style={{marginLeft:12,color:"#1a3a5c",fontWeight:600}}>
              {contratos.filter((x:Contrato)=>x.tipo==="contrato"&&dias.some(d=>x.fecha_entrega===isoD(d)||x.fecha_evento===isoD(d)||x.fecha_desmonte===isoD(d))).length} eventos esta semana
            </span>
          </div>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <button onClick={()=>setSemOff(0)} style={{...S.iconBtn,fontSize:11,padding:"4px 10px",opacity:semOff===0?.4:1}}>Hoy</button>
          <button onClick={()=>setSemOff(semOff-1)} style={{...S.iconBtn,fontSize:16,padding:"2px 10px"}}>‹</button>
          <button onClick={()=>setSemOff(semOff+1)} style={{...S.iconBtn,fontSize:16,padding:"2px 10px"}}>›</button>
        </div>
      </div>

      {/* Leyenda */}
      <div style={{display:"flex",gap:12,marginBottom:12,flexWrap:"wrap" as const}}>
        {([["entrega","🚚 Entrega",MOV_COLOR.entrega,MOV_BG.entrega],["evento","🎉 Evento",MOV_COLOR.evento,MOV_BG.evento],["desmonte","📦 Desmonte",MOV_COLOR.desmonte,MOV_BG.desmonte]] as [string,string,string,string][]).map(([k,l,col,bg])=>(
          <div key={k} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:col,background:bg,padding:"3px 10px",borderRadius:8,border:`1px solid ${col}30`}}>{l}</div>
        ))}
      </div>

      {/* Calendar grid */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8}}>
        {dias.map((dia,i)=>{
          const key=isoD(dia)
          const movs=byDay[key]||[]
          const isToday=esHoy(dia)
          const isPast=dia<new Date(isoD(new Date()))
          return(
            <div key={key} style={{background:isToday?"#fff8f0":isPast?"#fafaf8":"#fff",border:`1.5px solid ${isToday?"#92580a":"#e8e5de"}`,borderRadius:10,padding:"10px 8px",minHeight:180,display:"flex",flexDirection:"column" as const}}>
              {/* Day header */}
              <div style={{marginBottom:8,textAlign:"center" as const}}>
                <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase" as const,color:isToday?"#92580a":"#9a9590",letterSpacing:".06em"}}>{DIAS[i]}</div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,color:isToday?"#92580a":"#1a1814",lineHeight:1}}>{dia.getDate()}</div>
                {movs.length>0&&<div style={{fontSize:9,color:"#9a9590",marginTop:2}}>{movs.length} mov.</div>}
              </div>
              {/* Events */}
              <div style={{display:"flex",flexDirection:"column" as const,gap:4,flex:1,overflowY:"auto" as const,maxHeight:300}}>
                {movs.sort((a,b)=>a.tipo.localeCompare(b.tipo)).map((m,j)=>(
                  <div key={j} style={{background:MOV_BG[m.tipo],border:`1px solid ${MOV_COLOR[m.tipo]}25`,borderLeft:`3px solid ${MOV_COLOR[m.tipo]}`,borderRadius:5,padding:"4px 6px"}}>
                    <div style={{fontSize:9,fontWeight:700,textTransform:"uppercase" as const,color:MOV_COLOR[m.tipo],letterSpacing:".04em",marginBottom:1}}>
                      {m.tipo==="entrega"?"🚚":m.tipo==="evento"?"🎉":"📦"} {m.tipo}
                    </div>
                    <div style={{fontSize:10,fontWeight:600,color:"#1a1814",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{m.contrato.cliente||m.contrato.archivo}</div>
                    {m.contrato.lugar&&<div style={{fontSize:9,color:"#9a9590",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{m.contrato.lugar.slice(0,30)}</div>}
                    {m.contrato.folio&&m.contrato.folio!=="L"&&m.contrato.folio!=="K/H"&&<div style={{fontSize:9,color:"#4a4640"}}>#{m.contrato.folio}</div>}
                  </div>
                ))}
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}


// ─── EQUIPO ───────────────────────────────────────────────────────────
function EquipoSection({personal,contratos,user,nuevaPersona,setNuevaPersona,onAgregar,teamFiltro,setTeamFiltro}: any){
  const hoy=new Date();hoy.setHours(0,0,0,0)
  const en7=addDays(hoy,7)
  const cFilt=contratos.filter((c:Contrato)=>{
    const teamMatch=(new Date(c.fecha_entrega+"T12:00:00")>=hoy&&new Date(c.fecha_entrega+"T12:00:00")<=en7)||(new Date(c.fecha_desmonte+"T12:00:00")>=hoy&&new Date(c.fecha_desmonte+"T12:00:00")<=en7)
    if(teamFiltro==="semana")return teamMatch
    return true
  })
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap" as const,gap:8}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Coordinación del equipo</div>
        <div style={S.fset}>
          {[["hoy","Hoy"],["semana","Esta semana"],["todo","Todo"]].map(([f,l])=>(
            <button key={f} style={{...S.fbtn,...(teamFiltro===f?S.fbtnActive:{})}} onClick={()=>setTeamFiltro(f)}>{l}</button>
          ))}
        </div>
      </div>
      {user.rol==="admin"&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:18,marginBottom:20}}>
          <div style={{fontFamily:"Playfair Display,serif",fontWeight:700,fontSize:14,marginBottom:12}}>👥 Personal</div>
          <div style={{display:"flex",flexWrap:"wrap" as const,gap:8,marginBottom:12}}>
            {personal.map((p:Persona)=>(
              <div key={p.id} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 12px",borderRadius:20,border:`1.5px solid ${p.color_bd}`,background:p.color_bg,color:p.color_fg,fontSize:12,fontWeight:600}}>{p.nombre}</div>
            ))}
          </div>
          <div style={{display:"flex",gap:6}}>
            <input style={{...S.input,maxWidth:200,margin:0}} placeholder="Nombre del colaborador" value={nuevaPersona} onChange={e=>setNuevaPersona(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onAgregar()}/>
            <button style={S.loginBtn} onClick={onAgregar}>+ Agregar</button>
          </div>
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:12}}>
        {personal.map((p:Persona)=>{
          const movs:any[]=[]
          cFilt.forEach((c:Contrato)=>{
            if((c.asig_entrega||[]).includes(p.id))movs.push({c,tipo:"entrega",fecha:c.fecha_entrega,est:c.estado_entrega||"pend"})
            if((c.asig_desmonte||[]).includes(p.id))movs.push({c,tipo:"desmonte",fecha:c.fecha_desmonte,est:c.estado_desmonte||"pend"})
          })
          movs.sort((a,b)=>a.fecha.localeCompare(b.fecha))
          return (
            <div key={p.id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"14px 16px",borderBottom:"1px solid #e8e5de"}}>
                <div style={{width:36,height:36,borderRadius:"50%",background:p.color_bg,border:`2px solid ${p.color_bd}`,color:p.color_fg,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:14}}>{p.nombre.charAt(0).toUpperCase()}</div>
                <div><div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>{p.nombre}</div><div style={{fontSize:11,color:"#9a9590"}}>{movs.length} movimiento(s)</div></div>
              </div>
              <div style={{padding:"10px 16px",display:"flex",flexDirection:"column" as const,gap:6}}>
                {movs.length?movs.map((m,i)=>{
                  const e=ESTADOS[m.est]
                  return (
                    <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:8,background:"#fafaf8",border:"1px solid #e8e5de"}}>
                      <div style={{fontSize:14}}>{m.tipo==="entrega"?"🚚":"🔧"}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{m.c.cliente||m.c.archivo}</div>
                        <div style={{fontFamily:"monospace",fontSize:10,color:"#9a9590"}}>{fmtDate(m.fecha)}</div>
                      </div>
                      <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:e.bg,color:e.color,border:`1px solid ${e.border}`,fontWeight:700}}>{e.label}</span>
                    </div>
                  )
                }):<div style={{fontSize:12,color:"#c4bfb8",fontStyle:"italic",padding:"8px 0"}}>Sin movimientos</div>}
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}

// ─── DÍAS ─────────────────────────────────────────────────────────────
function DiasSection({contratos,semanaOffset,setSemanaOffset,getRango}: any){
  const rango=getRango(semanaOffset)
  const cFilt=rango?contratos.filter((c:Contrato)=>{
    const fe=new Date(c.fecha_evento+"T12:00:00"),fen=new Date(c.fecha_entrega+"T12:00:00"),fd=new Date(c.fecha_desmonte+"T12:00:00")
    const inRange=(fe>=rango.desde&&fe<=rango.hasta)||(fen>=rango.desde&&fen<=rango.hasta)||(fd>=rango.desde&&fd<=rango.hasta); return inRange
    return inRange
  }):contratos
  const dow=Array.from({length:7},(_,i)=>({i,e:0,v:0,d:0}))
  const byDate:Record<string,any>={}
  cFilt.forEach((c:Contrato)=>{
    dow[new Date(c.fecha_entrega+"T12:00:00").getDay()].e++
    dow[new Date(c.fecha_evento+"T12:00:00").getDay()].v++
    dow[new Date(c.fecha_desmonte+"T12:00:00").getDay()].d++
    ;[[c.fecha_entrega,"e"],[c.fecha_evento,"v"],[c.fecha_desmonte,"d"]].forEach(([f,t])=>{
      const d=new Date((f as string)+"T12:00:00")
      if(rango&&(d<rango.desde||d>rango.hasta))return
      const k=f as string
      if(!byDate[k])byDate[k]={d,e:0,v:0,d2:0,rows:[]}
      if(t==="e")byDate[k].e++;else if(t==="v")byDate[k].v++;else byDate[k].d2++
      byDate[k].rows.push({...c,_t:t})
    })
  })
  const mx=Math.max(...dow.map(d=>d.e+d.v+d.d),1)
  const hoy2=new Date();hoy2.setHours(0,0,0,0)
  const fechas=Object.keys(byDate).sort()
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap" as const,gap:8}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Conteos por día</div>
        <div style={S.fset}>
          {([[0,"Esta semana"],[1,"Próx. semana"],[2,"+2 sem"],[3,"+3 sem"],[-1,"Todo"]] as const).map(([o,l])=>(
            <button key={o} style={{...S.fbtn,...(semanaOffset===o?S.fbtnActive:{})}} onClick={()=>setSemanaOffset(o)}>{l}</button>
          ))}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))",gap:10,marginBottom:24}}>
        {dow.filter(d=>d.e+d.v+d.d>0).map(d=>(
          <div key={d.i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:11,padding:16}}>
            <div style={{display:"flex",alignItems:"baseline",justifyContent:"space-between",marginBottom:12}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700}}>{DIAS[d.i]}</div>
              <div style={{fontFamily:"monospace",fontSize:11,background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:5,padding:"2px 7px",color:"#9a9590"}}>{d.e+d.v+d.d}</div>
            </div>
            {[["Entrega",d.e,"#2d6a4f"],["Evento",d.v,"#92580a"],["Desmonte",d.d,"#8b2e2e"]].map(([l,n,col])=>(
              <div key={l as string} style={{display:"flex",alignItems:"center",gap:7,marginBottom:5}}>
                <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase" as const,letterSpacing:".05em",width:58,flexShrink:0,color:"#9a9590"}}>{l}</div>
                <div style={{flex:1,background:"#f5f4f0",borderRadius:3,height:6,overflow:"hidden"}}><div style={{height:"100%",borderRadius:3,background:col as string,width:`${Math.round((n as number)/mx*100)}%`}}/></div>
                <div style={{fontFamily:"monospace",fontSize:10,width:14,textAlign:"right" as const,color:"#4a4640"}}>{n}</div>
              </div>
            ))}
          </div>
        ))}
      </div>
      {fechas.length>0&&(
        <div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700,marginBottom:10}}>Detalle por fecha</div>
          {fechas.map(k=>{
            const g=byDate[k];const d=new Date(k+"T12:00:00")
            return (
              <div key={k} style={{marginBottom:10,borderRadius:10,overflow:"hidden",border:"1px solid #e8e5de"}}>
                <div style={{display:"flex",alignItems:"center",gap:8,padding:"9px 14px",background:"#fff",borderBottom:"1px solid #e8e5de",flexWrap:"wrap" as const}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>{d.getDate()} de {MESES_F[d.getMonth()]} {d.getFullYear()}</div>
                  <div style={{fontSize:10,fontWeight:700,fontFamily:"monospace",padding:"2px 8px",borderRadius:10,background:"#fafaf8",color:"#4a4640"}}>{DIAS[d.getDay()]}</div>
                  {g.e>0&&<span style={{fontSize:10,padding:"2px 7px",borderRadius:10,fontWeight:700,background:"#edf7f2",color:"#2d6a4f",border:"1px solid #b7deca"}}>🚚 {g.e}</span>}
                  {g.v>0&&<span style={{fontSize:10,padding:"2px 7px",borderRadius:10,fontWeight:700,background:"#fdf5e8",color:"#92580a",border:"1px solid #f0d49a"}}>📅 {g.v}</span>}
                  {g.d2>0&&<span style={{fontSize:10,padding:"2px 7px",borderRadius:10,fontWeight:700,background:"#fdf0f0",color:"#8b2e2e",border:"1px solid #e8b8b8"}}>🔧 {g.d2}</span>}
                </div>
                {g.rows.map((c:any,i:number)=>{
                  const tp=c._t==="e"?"entrega":c._t==="v"?"evento":"desmonte"
                  return (
                    <div key={i} style={{display:"flex",flexDirection:"column" as const,padding:"8px 14px",background:i%2===0?"#fff":"#fafaf8",borderTop:"1px solid #e8e5de"}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap" as const}}>
                        <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,fontWeight:700,background:tp==="entrega"?"#edf7f2":tp==="evento"?"#fdf5e8":"#fdf0f0",color:tp==="entrega"?"#2d6a4f":tp==="evento"?"#92580a":"#8b2e2e",border:`1px solid ${tp==="entrega"?"#b7deca":tp==="evento"?"#f0d49a":"#e8b8b8"}`}}>{tp.toUpperCase()}</span>
                        <span style={{fontWeight:700,fontSize:13}}>{c.cliente||c.archivo}</span>
                        {c.lugar&&<span style={{color:"#9a9590",fontSize:11}}>📍 {c.lugar}</span>}
                      </div>
                    </div>
                  )
                })}
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}

// ─── GANTT ────────────────────────────────────────────────────────────
function GanttSection({contratos,desde,hasta,setDesde,setHasta,personal}: any){
  const dias:Date[]=[]
  let cur=new Date(desde+"T12:00:00")
  const end=new Date(hasta+"T12:00:00")
  while(cur<=end){dias.push(new Date(cur));cur.setDate(cur.getDate()+1)}
  const lista=contratos.filter((c:Contrato)=>c.fecha_desmonte>=desde&&c.fecha_entrega<=hasta).sort((a:Contrato,b:Contrato)=>a.fecha_evento.localeCompare(b.fecha_evento))
  const hoy=new Date();hoy.setHours(12,0,0,0)
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap" as const,gap:8}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Gantt</div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <span style={{fontSize:11,color:"#9a9590"}}>Desde</span>
          <input type="date" value={desde} onChange={e=>setDesde(e.target.value)} style={{background:"#fff",border:"1px solid #e8e5de",color:"#1a1814",borderRadius:8,padding:"6px 10px",fontFamily:"monospace",fontSize:11}}/>
          <span style={{fontSize:11,color:"#9a9590"}}>Hasta</span>
          <input type="date" value={hasta} onChange={e=>setHasta(e.target.value)} style={{background:"#fff",border:"1px solid #e8e5de",color:"#1a1814",borderRadius:8,padding:"6px 10px",fontFamily:"monospace",fontSize:11}}/>
        </div>
      </div>
      <div style={{overflowX:"auto" as const,border:"1px solid #e8e5de",borderRadius:12}}>
        <table style={{borderCollapse:"collapse" as const,minWidth:"100%"}}>
          <thead>
            <tr>
              <th style={{background:"#fafaf8",borderBottom:"1px solid #e8e5de",borderRight:"2px solid #d4cfc4",padding:"6px 12px",fontSize:9.5,fontWeight:700,textTransform:"uppercase" as const,color:"#9a9590",textAlign:"left" as const,minWidth:180,position:"sticky" as const,left:0,zIndex:3}}>Contrato</th>
              {dias.map((d,i)=>{
                const isH=sameDay(d,hoy)
                return (
                  <th key={i} style={{background:isH?"#fdf5e8":"#fafaf8",borderBottom:"1px solid #e8e5de",borderRight:"1px solid #e8e5de",padding:"5px 2px",fontSize:9,fontWeight:700,textTransform:"uppercase" as const,color:isH?"#92580a":"#9a9590",textAlign:"center" as const,fontFamily:"monospace",minWidth:36}}>
                    <div>{d.getDate()}</div>
                    <div style={{fontSize:8,opacity:.7}}>{MESES[d.getMonth()]}</div>
                    <div style={{fontSize:8}}>{DIAS[d.getDay()].slice(0,3)}</div>
                  </th>
                )
              })}
            </tr>
          </thead>
          <tbody>
            {lista.map((c:Contrato)=>{
              const asigColor=c.asig_entrega?.length?(personal.find((p:Persona)=>p.id===c.asig_entrega[0])?.color_fg||""):""
              return (
                <tr key={c.id}>
                  <td style={{padding:"0 12px",fontSize:11.5,fontWeight:600,whiteSpace:"nowrap" as const,background:"#fff",position:"sticky" as const,left:0,zIndex:2,borderRight:"2px solid #d4cfc4",maxWidth:190,overflow:"hidden",textOverflow:"ellipsis",height:34,borderLeft:asigColor?`3px solid ${asigColor}`:"none"}}>{c.cliente||c.archivo}</td>
                  {dias.map((d,i)=>{
                    const ds=isoDate(d)
                    const isE=ds===c.fecha_entrega,isV=ds===c.fecha_evento,isD=ds===c.fecha_desmonte
                    const isSp=ds>c.fecha_entrega&&ds<c.fecha_desmonte&&!isV
                    const isH=sameDay(d,hoy)
                    return (
                      <td key={i} style={{background:isH?"rgba(146,88,10,.04)":"#fff",position:"relative" as const,height:34,borderRight:"1px solid rgba(232,229,222,.3)"}}>
                        {isE&&<div style={{position:"absolute" as const,inset:"4px 2px",borderRadius:5,display:"flex",alignItems:"center",padding:"0 5px",fontSize:9,fontWeight:700,background:"#edf7f2",color:"#2d6a4f",border:"1px solid #b7deca"}}>🚚</div>}
                        {isV&&<div style={{position:"absolute" as const,inset:"4px 2px",borderRadius:5,display:"flex",alignItems:"center",padding:"0 5px",fontSize:9,fontWeight:700,background:"#fdf5e8",color:"#92580a",border:"1px solid #f0d49a"}}>📅</div>}
                        {isD&&<div style={{position:"absolute" as const,inset:"4px 2px",borderRadius:5,display:"flex",alignItems:"center",padding:"0 5px",fontSize:9,fontWeight:700,background:"#fdf0f0",color:"#8b2e2e",border:"1px solid #e8b8b8"}}>🔧</div>}
                        {isSp&&<div style={{position:"absolute" as const,inset:"11px 0",borderTop:"1.5px dashed #b8ceea",borderBottom:"1.5px dashed #b8ceea"}}/>}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ─── RUTAS ────────────────────────────────────────────────────────────
function RutasSection({contratos,personal,rutaTipo,setRutaTipo,rutaFecha,setRutaFecha}: any){
  const hoy=new Date();hoy.setHours(0,0,0,0)
  const fechas=[...new Set(contratos.flatMap((c:Contrato)=>[c.fecha_entrega,c.fecha_desmonte]))].filter(f=>{const d=new Date((f as string)+"T12:00:00");return d>=hoy}).sort() as string[]
  let lista:any[]=[]
  if(rutaFecha){
    if(rutaTipo==="ambos"){
      contratos.filter((c:Contrato)=>c.fecha_entrega===rutaFecha).forEach((c:Contrato)=>lista.push({...c,_tipo:"entrega"}))
      contratos.filter((c:Contrato)=>c.fecha_desmonte===rutaFecha).forEach((c:Contrato)=>lista.push({...c,_tipo:"desmonte"}))
    } else {
      contratos.filter((c:Contrato)=>(rutaTipo==="entrega"?c.fecha_entrega:c.fecha_desmonte)===rutaFecha).forEach((c:Contrato)=>lista.push({...c,_tipo:rutaTipo}))
    }
    lista.sort((a,b)=>a.cliente.localeCompare(b.cliente))
  }
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Rutas</div>
      </div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap" as const,marginBottom:16}}>
        <select value={rutaFecha} onChange={e=>setRutaFecha(e.target.value)} style={{...S.input,minWidth:220,margin:0,cursor:"pointer"}}>
          <option value="">— Selecciona fecha —</option>
          {fechas.map(f=>{const d=new Date(f+"T12:00:00");return<option key={f} value={f}>{d.getDate()} de {MESES_F[d.getMonth()]} — {DIAS[d.getDay()]}</option>})}
        </select>
        <div style={S.fset}>
          {[["entrega","🚚 Entregas"],["desmonte","🔧 Desmontes"],["ambos","📋 Ambos"]].map(([t,l])=>(
            <button key={t} style={{...S.fbtn,...(rutaTipo===t?S.fbtnActive:{})}} onClick={()=>setRutaTipo(t)}>{l}</button>
          ))}
        </div>
      </div>
      {rutaFecha&&lista.length>0?(
        <div style={{display:"flex",flexDirection:"column" as const,gap:7}}>
          {lista.map((c,i)=>{
            const e=ESTADOS[c[c._tipo==="entrega"?"estado_entrega":"estado_desmonte"]||"pend"]
            const asignados=(c[c._tipo==="entrega"?"asig_entrega":"asig_desmonte"]||[]).map((pid:string)=>personal.find((p:Persona)=>p.id===pid)?.nombre).filter(Boolean)
            return (
              <div key={c.id+c._tipo} style={{display:"flex",alignItems:"flex-start",gap:10,background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:"#d4cfc4",width:32,textAlign:"right" as const,lineHeight:1,flexShrink:0}}>{String(i+1).padStart(2,"0")}</div>
                <div style={{fontSize:17,flexShrink:0}}>{c._tipo==="entrega"?"🚚":"🔧"}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:700}}>{c.cliente||c.archivo}</div>
                  <div style={{fontSize:11.5,color:"#4a4640",marginTop:2}}>{c.lugar||"Sin dirección"}</div>
                  {c.tel&&<div style={{fontFamily:"monospace",fontSize:11,color:"#9a9590",marginTop:1}}>📞 {c.tel}</div>}
                  {asignados.length>0&&<div style={{fontSize:11,color:"#9a9590",marginTop:2}}>👤 {asignados.join(", ")}</div>}
                </div>
                <span style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:e.bg,color:e.color,border:`1px solid ${e.border}`,fontWeight:700,whiteSpace:"nowrap" as const,flexShrink:0}}>{e.label}</span>
              </div>
            )
          })}
        </div>
      ):(
        <div style={{textAlign:"center" as const,padding:"48px 20px",background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
          <div style={{fontSize:36,opacity:.2}}>🗺️</div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,color:"#4a4640"}}>{rutaFecha?"Sin movimientos ese día":"Selecciona una fecha"}</div>
        </div>
      )}
    </div>
  )
}

// ─── MIS RUTAS ────────────────────────────────────────────────────────
function MisRutasSection({contratos,personal,rutas,setRutas,showNueva,setShowNueva}: any){
  const [paso,setPaso]=useState(1)
  const [rutaVer,setRutaVer]=useState<string|null>(null)
  const [nuevaRuta,setNuevaRuta]=useState({
    nombre:"",fecha:"",tipo:"ambos",
    unidades:UNIDADES_DEFAULT.map((n,i)=>({id:"u"+i,nombre:n,color:UNIDAD_COLORES[i],asignados:[] as string[]})),
    contratos_ids:[] as string[],
    asignaciones:{} as Record<string,string>
  })
  const hoy=new Date();hoy.setHours(0,0,0,0)
  const getDisponibles=(fecha:string,tipo:string)=>{
    const showEnt=tipo==="ambos"||tipo==="entrega"
    const showDes=tipo==="ambos"||tipo==="desmonte"
    const ent=showEnt?contratos.filter((c:Contrato)=>c.fecha_entrega===fecha).map((c:Contrato)=>({...c,_mov:"entrega"})):[]
    const des=showDes?contratos.filter((c:Contrato)=>c.fecha_desmonte===fecha).map((c:Contrato)=>({...c,_mov:"desmonte"})):[]
    const merged:any[]=[],seen=new Set<string>()
    ent.forEach((c:any)=>{const hasDes=des.some((d:any)=>d.id===c.id);merged.push({...c,_mov:hasDes?"ambos":"entrega"});seen.add(c.id)})
    des.forEach((c:any)=>{if(!seen.has(c.id))merged.push({...c,_mov:"desmonte"})})
    return merged.sort((a:any,b:any)=>a.cliente.localeCompare(b.cliente))
  }
  const cDisp=nuevaRuta.fecha?getDisponibles(nuevaRuta.fecha,nuevaRuta.tipo):[]
  const cSel=cDisp.filter((c:any)=>nuevaRuta.contratos_ids.includes(c.id))
  function crearRuta(){
    if(!nuevaRuta.fecha)return
    setRutas((rs:any[])=>[...rs,{...nuevaRuta,id:Date.now().toString()}])
    setNuevaRuta({nombre:"",fecha:"",tipo:"ambos",unidades:UNIDADES_DEFAULT.map((n,i)=>({id:"u"+i,nombre:n,color:UNIDAD_COLORES[i],asignados:[]})),contratos_ids:[],asignaciones:{}})
    setShowNueva(false);setPaso(1)
  }
  // Vista detalle
  if(rutaVer){
    const ruta=rutas.find((r:any)=>r.id===rutaVer)
    if(!ruta){setRutaVer(null);return null}
    const cRuta=contratos.filter((c:Contrato)=>ruta.contratos_ids.includes(c.id))
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
          <button onClick={()=>setRutaVer(null)} style={{...S.iconBtn}}>← Volver</button>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>{ruta.nombre||"Ruta"}</div>
          <div style={{fontSize:12,color:"#9a9590"}}>{fmtDate(ruta.fecha)} · {cRuta.length} contratos</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
          {ruta.unidades.map((u:any)=>{
            const uC=cRuta.filter((c:Contrato)=>(ruta.asignaciones||{})[c.id]===u.id)
            const uP=(u.asignados||[]).map((pid:string)=>personal.find((p:Persona)=>p.id===pid)).filter(Boolean)
            return (
              <div key={u.id} style={{background:"#fff",border:`2px solid ${u.color}20`,borderRadius:12,overflow:"hidden"}}>
                <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 14px",background:`${u.color}12`,borderBottom:`1px solid ${u.color}30`}}>
                  <div style={{width:12,height:12,borderRadius:"50%",background:u.color}}/>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,flex:1}}>{u.nombre}</div>
                  <div style={{fontSize:10,color:u.color,fontWeight:700}}>{uC.length} paradas</div>
                </div>
                {uP.length>0&&(
                  <div style={{padding:"8px 14px",borderBottom:`1px solid ${u.color}20`,display:"flex",gap:4,flexWrap:"wrap" as const}}>
                    {uP.map((p:Persona)=>(
                      <span key={p.id} style={{fontSize:10,padding:"2px 8px",borderRadius:10,background:p.color_bg,border:`1px solid ${p.color_bd}`,color:p.color_fg,fontWeight:600}}>{p.nombre}</span>
                    ))}
                  </div>
                )}
                <div style={{padding:"10px 14px",display:"flex",flexDirection:"column" as const,gap:6}}>
                  {uC.length===0&&<div style={{fontSize:11,color:"#c4bfb8",fontStyle:"italic",textAlign:"center" as const,padding:"8px 0"}}>Sin paradas</div>}
                  {uC.map((c:Contrato)=>{
                    const secs:Record<string,Articulo[]>={}
                    ;(c.articulos||[]).forEach((a:Articulo)=>{const s=a.seccion||"General";if(!secs[s])secs[s]=[];secs[s].push(a)})
                    return (
                      <div key={c.id} style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:8,padding:"9px 11px"}}>
                        <div style={{fontFamily:"Playfair Display,serif",fontSize:13,fontWeight:700}}>{c.cliente||c.archivo}</div>
                        {c.lugar&&<div style={{fontSize:11,color:"#4a4640",marginTop:2}}>📍 {c.lugar}</div>}
                        <div style={{marginTop:6}}>
                          {Object.entries(secs).map(([sec,arts])=>(
                            <div key={sec}>
                              <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:2}}>{sec}</div>
                              {(arts as Articulo[]).map((a,j)=>(
                                <div key={j} style={{display:"flex",gap:6,padding:"1px 0"}}>
                                  <span style={{fontFamily:"monospace",fontWeight:700,color:u.color,minWidth:22,textAlign:"right" as const,fontSize:11}}>{a.cantidad}</span>
                                  <span style={{fontSize:10.5}}>{a.nombre}</span>
                                </div>
                              ))}
                            </div>
                          ))}
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            )
          })}
        </div>
      </div>
    )
  }
  // Lista de rutas
  if(!showNueva)return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Mis Rutas</div>
        <button style={{...S.loginBtn,padding:"8px 18px",fontSize:12}} onClick={()=>setShowNueva(true)}>+ Nueva ruta</button>
      </div>
      {!rutas.length?(
        <div style={{textAlign:"center" as const,padding:"60px 20px",background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
          <div style={{fontSize:48,opacity:.15,marginBottom:12}}>🚛</div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,color:"#4a4640",marginBottom:6}}>Sin rutas creadas</div>
          <button style={{...S.loginBtn,padding:"10px 24px"}} onClick={()=>setShowNueva(true)}>Crear primera ruta</button>
        </div>
      ):(
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(340px,1fr))",gap:12}}>
          {rutas.map((ruta:any)=>{
            const cRuta=contratos.filter((c:Contrato)=>ruta.contratos_ids.includes(c.id))
            return (
              <div key={ruta.id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden",cursor:"pointer"}} onClick={()=>setRutaVer(ruta.id)}>
                <div style={{padding:"14px 16px",borderBottom:"1px solid #e8e5de",background:"#fafaf8"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{fontSize:22}}>{ruta.tipo==="entrega"?"🚚":ruta.tipo==="desmonte"?"🔧":"🚚🔧"}</div>
                    <div style={{flex:1}}>
                      <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>{ruta.nombre||"Ruta"}</div>
                      <div style={{fontSize:11,color:"#9a9590"}}>{fmtDate(ruta.fecha)} · {cRuta.length} contratos</div>
                    </div>
                    <button onClick={e=>{e.stopPropagation();setRutas((rs:any[])=>rs.filter((r:any)=>r.id!==ruta.id))}} style={{background:"none",border:"none",color:"#c4bfb8",cursor:"pointer",fontSize:16,padding:0}}>✕</button>
                  </div>
                </div>
                <div style={{padding:"8px 16px",borderTop:"1px solid #e8e5de",display:"flex",justifyContent:"flex-end"}}>
                  <span style={{fontSize:11,color:"#9a9590"}}>Ver detalle →</span>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
  // Crear nueva ruta
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:20}}>
        <button onClick={()=>{setShowNueva(false);setPaso(1)}} style={{...S.iconBtn}}>✕ Cancelar</button>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,flex:1}}>Nueva ruta</div>
        <div style={{display:"flex",gap:4}}>
          {[1,2,3].map(p=>(
            <div key={p} style={{display:"flex",alignItems:"center",gap:4}}>
              <div style={{width:28,height:28,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,background:paso>=p?"#1a1814":"#e8e5de",color:paso>=p?"#fff":"#9a9590"}}>{p}</div>
              {p<3&&<div style={{width:20,height:2,background:paso>p?"#1a1814":"#e8e5de"}}/>}
            </div>
          ))}
        </div>
      </div>
      {paso===1&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:24}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:700,marginBottom:16}}>Paso 1 — Configurar ruta</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:16,marginBottom:20}}>
            <div>
              <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:6,textTransform:"uppercase" as const,letterSpacing:".06em"}}>Nombre</div>
              <input style={{...S.input,margin:0}} placeholder="ej. Ruta Norte..." value={nuevaRuta.nombre} onChange={e=>setNuevaRuta((n:any)=>({...n,nombre:e.target.value}))}/>
            </div>
            <div>
              <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:6,textTransform:"uppercase" as const,letterSpacing:".06em"}}>Fecha</div>
              <input type="date" style={{...S.input,margin:0}} value={nuevaRuta.fecha} onChange={e=>setNuevaRuta((n:any)=>({...n,fecha:e.target.value,contratos_ids:[]}))}/>
            </div>
            <div>
              <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:6,textTransform:"uppercase" as const,letterSpacing:".06em"}}>Tipo</div>
              <div style={{display:"flex",gap:6}}>
                {[["ambos","🚚🔧"],["entrega","🚚"],["desmonte","🔧"]].map(([t,l])=>(
                  <button key={t} onClick={()=>setNuevaRuta((n:any)=>({...n,tipo:t,contratos_ids:[]}))} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${nuevaRuta.tipo===t?"#1a1814":"#e8e5de"}`,background:nuevaRuta.tipo===t?"#1a1814":"#fff",color:nuevaRuta.tipo===t?"#fff":"#4a4640",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>{l}</button>
                ))}
              </div>
            </div>
          </div>
          {nuevaRuta.fecha&&(
            <div style={{marginBottom:20}}>
              <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:8,textTransform:"uppercase" as const}}>Contratos ese día ({cDisp.length})</div>
              {cDisp.length===0?(
                <div style={{padding:"12px",background:"#fafaf8",borderRadius:8,fontSize:12,color:"#9a9590",textAlign:"center" as const}}>Sin contratos</div>
              ):(
                <div style={{display:"flex",flexWrap:"wrap" as const,gap:6}}>
                  {cDisp.slice(0,8).map((cx:any)=>(
                    <div key={cx.id} style={{padding:"4px 10px",borderRadius:8,background:cx._mov==="entrega"?"#edf7f2":cx._mov==="desmonte"?"#fdf0f0":"#edf3fa",border:"1px solid #e8e5de",fontSize:11,color:"#1a1814"}}>
                      {cx._mov==="entrega"?"🚚":cx._mov==="desmonte"?"🔧":"🚚🔧"} {cx.cliente||cx.archivo}
                    </div>
                  ))}
                  {cDisp.length>8&&<div style={{fontSize:11,color:"#9a9590"}}>+{cDisp.length-8} más</div>}
                </div>
              )}
            </div>
          )}
          <div style={{display:"flex",justifyContent:"flex-end"}}>
            <button style={{...S.loginBtn,padding:"10px 28px",opacity:(!nuevaRuta.fecha||cDisp.length===0)?.5:1}} onClick={()=>nuevaRuta.fecha&&cDisp.length>0&&setPaso(2)}>Siguiente →</button>
          </div>
        </div>
      )}
      {paso===2&&(
        <div>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:24,marginBottom:16}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:700,marginBottom:4}}>Paso 2 — Seleccionar contratos</div>
            <div style={{fontSize:12,color:"#9a9590",marginBottom:12}}>{nuevaRuta.contratos_ids.length} de {cDisp.length} seleccionados</div>
            <div style={{display:"flex",gap:8,marginBottom:12}}>
              <button onClick={()=>setNuevaRuta((n:any)=>({...n,contratos_ids:cDisp.map((c:any)=>c.id)}))} style={{...S.iconBtn,fontSize:11}}>✓ Todos</button>
              <button onClick={()=>setNuevaRuta((n:any)=>({...n,contratos_ids:[]}))} style={{...S.iconBtn,fontSize:11}}>✕ Ninguno</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:8}}>
              {cDisp.map((c:any)=>{
                const sel=nuevaRuta.contratos_ids.includes(c.id)
                const pzs=(c.articulos||[]).reduce((s:number,a:Articulo)=>s+a.cantidad,0)
                return (
                  <div key={c.id} onClick={()=>setNuevaRuta((n:any)=>({...n,contratos_ids:sel?n.contratos_ids.filter((x:string)=>x!==c.id):[...n.contratos_ids,c.id]}))}
                    style={{padding:"12px 14px",borderRadius:10,border:`2px solid ${sel?"#1a1814":"#e8e5de"}`,background:sel?"#1a1814":"#fff",cursor:"pointer"}}>
                    <div style={{fontSize:13,fontWeight:700,color:sel?"#fff":"#1a1814"}}>{c.cliente||c.archivo}</div>
                    {c.lugar&&<div style={{fontSize:11,color:sel?"rgba(255,255,255,.7)":"#9a9590"}}>📍 {c.lugar}</div>}
                    {pzs>0&&<div style={{fontSize:10,color:sel?"rgba(255,255,255,.6)":"#92580a",fontWeight:600}}>📦 {pzs} pzs</div>}
                  </div>
                )
              })}
            </div>
          </div>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <button style={{...S.iconBtn}} onClick={()=>setPaso(1)}>← Atrás</button>
            <button style={{...S.loginBtn,padding:"10px 28px",opacity:nuevaRuta.contratos_ids.length===0?.5:1}} onClick={()=>nuevaRuta.contratos_ids.length>0&&setPaso(3)}>Siguiente → Asignar ({nuevaRuta.contratos_ids.length})</button>
          </div>
        </div>
      )}
      {paso===3&&(
        <div>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:24,marginBottom:16}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:700,marginBottom:16}}>Paso 3 — Asignar a unidades</div>
            {cSel.filter((c:any)=>!nuevaRuta.asignaciones[c.id]).length>0&&(
              <div style={{marginBottom:16}}>
                <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:8,textTransform:"uppercase" as const}}>Sin asignar</div>
                <div style={{display:"flex",flexWrap:"wrap" as const,gap:6}}>
                  {cSel.filter((c:any)=>!nuevaRuta.asignaciones[c.id]).map((c:any)=>(
                    <div key={c.id} style={{background:"#fafaf8",border:"1.5px dashed #d4cfc4",borderRadius:8,padding:"8px 12px"}}>
                      <div style={{fontSize:12,fontWeight:700,marginBottom:6}}>{c.cliente||c.archivo}</div>
                      <div style={{display:"flex",gap:4,flexWrap:"wrap" as const}}>
                        {nuevaRuta.unidades.map((u:any)=>(
                          <button key={u.id} onClick={()=>setNuevaRuta((n:any)=>({...n,asignaciones:{...n.asignaciones,[c.id]:u.id}}))}
                            style={{padding:"3px 9px",borderRadius:6,border:`1.5px solid ${u.color}`,background:`${u.color}10`,color:u.color,fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>{u.nombre}</button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10}}>
              {nuevaRuta.unidades.map((u:any)=>{
                const uC=cSel.filter((c:any)=>nuevaRuta.asignaciones[c.id]===u.id)
                return (
                  <div key={u.id} style={{background:"#fafaf8",border:`2px solid ${u.color}30`,borderRadius:10,overflow:"hidden"}}>
                    <div style={{padding:"10px 12px",background:`${u.color}12`,borderBottom:`1px solid ${u.color}20`,display:"flex",alignItems:"center",gap:6}}>
                      <div style={{width:10,height:10,borderRadius:"50%",background:u.color}}/>
                      <div style={{fontWeight:700,fontSize:12,flex:1}}>{u.nombre}</div>
                      <div style={{fontSize:10,color:u.color,fontWeight:700}}>{uC.length} paradas</div>
                    </div>
                    <div style={{padding:"8px 12px",display:"flex",flexDirection:"column" as const,gap:4}}>
                      {uC.length===0&&<div style={{fontSize:10,color:"#c4bfb8",fontStyle:"italic",textAlign:"center" as const,padding:"6px 0"}}>Vacía</div>}
                      {uC.map((c:any)=>(
                        <div key={c.id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:7,padding:"7px 9px"}}>
                          <div style={{fontSize:11.5,fontWeight:700}}>{c.cliente||c.archivo}</div>
                          {(c.articulos||[]).length>0&&<div style={{fontSize:9.5,color:"#92580a",fontWeight:600}}>📦 {(c.articulos||[]).reduce((s:number,a:Articulo)=>s+a.cantidad,0)} pzs</div>}
                        </div>
                      ))}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <button style={{...S.iconBtn}} onClick={()=>setPaso(2)}>← Atrás</button>
            <button style={{...S.loginBtn,padding:"10px 28px"}} onClick={crearRuta}>✓ Crear ruta</button>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── ARTÍCULOS ────────────────────────────────────────────────────────
function ArticulosSection({contratos,artSemOffset,setArtSemOffset,getRango,artBusqueda,setArtBusqueda}: any){
  const rango=getRango(artSemOffset)
  const cFilt=rango?contratos.filter((c:Contrato)=>{const fe=new Date(c.fecha_evento+"T12:00:00");return fe>=rango.desde&&fe<=rango.hasta}):contratos
  const q=artBusqueda.toLowerCase().trim()
  const consolidado:Record<string,any>={}
  cFilt.forEach((c:Contrato)=>{
    (c.articulos||[]).forEach((a:Articulo)=>{
      if(q&&!a.nombre.toLowerCase().includes(q))return
      const key=a.nombre.toLowerCase().trim()
      if(!consolidado[key])consolidado[key]={nombre:a.nombre,cantidad:0,importe:0,seccion:a.seccion||"General",contratos:[]}
      consolidado[key].cantidad+=a.cantidad
      consolidado[key].importe+=a.importe||0
      if(!consolidado[key].contratos.includes(c.cliente||c.archivo))consolidado[key].contratos.push(c.cliente||c.archivo)
    })
  })
  const lista=Object.values(consolidado).sort((a,b)=>b.cantidad-a.cantidad)
  const secciones:Record<string,any[]>={}
  lista.forEach(a=>{const s=a.seccion||"General";if(!secciones[s])secciones[s]=[];secciones[s].push(a)})
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap" as const,gap:8}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Inventario de artículos</div>
        <div style={S.fset}>
          {([[0,"Esta semana"],[1,"Próx. semana"],[2,"+2 sem"],[-1,"Todo"]] as const).map(([o,l])=>(
            <button key={o} style={{...S.fbtn,...(artSemOffset===o?S.fbtnActive:{})}} onClick={()=>setArtSemOffset(o)}>{l}</button>
          ))}
        </div>
      </div>
      <input value={artBusqueda} onChange={e=>setArtBusqueda(e.target.value)} placeholder="🔍 Buscar artículo..." style={{...S.input,maxWidth:400,marginBottom:14}}/>
      {!lista.length?(
        <div style={{textAlign:"center" as const,padding:"48px 20px",background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
          <div style={{fontSize:36,opacity:.2}}>📦</div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,color:"#4a4640"}}>Sin artículos</div>
        </div>
      ):Object.keys(secciones).map(sec=>(
        <div key={sec} style={{marginBottom:20}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:8}}>{sec}</div>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,overflow:"hidden"}}>
            {secciones[sec].map((a,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 16px",background:i%2===0?"#fff":"#fafaf8",borderBottom:"1px solid rgba(232,229,222,.4)"}}>
                <div style={{fontFamily:"monospace",fontSize:20,fontWeight:800,color:"#92580a",minWidth:50,textAlign:"right" as const}}>{a.cantidad}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:600}}>{a.nombre}</div>
                  <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>{a.contratos.slice(0,3).join(" · ")}{a.contratos.length>3?` +${a.contratos.length-3} más`:""}</div>
                </div>
                {a.importe>0&&<div style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#2d6a4f"}}>${a.importe.toLocaleString()}</div>}
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

// ─── CLIENTES ─────────────────────────────────────────────────────────
function ClientesSection({contratos}: any){
  const [busq,setBusq]=useState("")
  const [sel,setSel]=useState<string|null>(null)
  const clienteMap:Record<string,any>={}
  contratos.forEach((c:Contrato)=>{
    const key=(c.cliente||c.archivo).toLowerCase().trim()
    if(!clienteMap[key])clienteMap[key]={nombre:c.cliente||c.archivo,tel:c.tel||"",eventos:[],totalContratado:0,totalCobrado:0}
    clienteMap[key].eventos.push(c)
    clienteMap[key].totalContratado+=(c.total||0)
    clienteMap[key].totalCobrado+=(c.cobrado||0)
  })
  const clientes=Object.values(clienteMap).filter((cl:any)=>!busq||cl.nombre.toLowerCase().includes(busq.toLowerCase())).sort((a:any,b:any)=>b.totalContratado-a.totalContratado)
  if(sel){
    const cl=clienteMap[sel.toLowerCase().trim()]
    if(!cl){setSel(null);return null}
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
          <button style={{...S.iconBtn}} onClick={()=>setSel(null)}>← Volver</button>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>{cl.nombre}</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
          {[{v:cl.eventos.length,l:"Eventos"},{v:fmt(cl.totalContratado),l:"Total contratado"},{v:fmt(cl.totalCobrado),l:"Cobrado"}].map((k,i)=>(
            <div key={i} style={S.kpi}><div style={S.kl}>{k.l}</div><div style={{...S.kv,fontSize:20,marginTop:4}}>{k.v}</div></div>
          ))}
        </div>
        <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
          {cl.eventos.sort((a:Contrato,b:Contrato)=>b.fecha_evento.localeCompare(a.fecha_evento)).map((c:Contrato)=>(
            <div key={c.id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:13,fontWeight:700}}>{c.cliente||c.archivo}</div>
              <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>📅 {c.fecha_evento}{c.lugar?" · 📍 "+c.lugar:""}</div>
              {(c.total||0)>0&&<div style={{fontSize:11,marginTop:4}}>Total: <strong style={{color:"#2d6a4f"}}>{fmt(c.total)}</strong> · Saldo: <strong style={{color:"#92580a"}}>{fmt((c.total||0)-(c.cobrado||0))}</strong></div>}
            </div>
          ))}
        </div>
      </div>
    )
  }
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Clientes</div>
        <div style={{fontSize:12,color:"#9a9590"}}>{clientes.length} clientes</div>
      </div>
      <input value={busq} onChange={e=>setBusq(e.target.value)} placeholder="🔍 Buscar..." style={{...S.input,maxWidth:340,marginBottom:14}}/>
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
        {clientes.map((cl:any,i:number)=>{
          const initials=cl.nombre.split(" ").slice(0,2).map((w:string)=>w[0]?.toUpperCase()||"").join("")
          const saldo=cl.totalContratado-cl.totalCobrado
          return (
            <div key={i} onClick={()=>setSel(cl.nombre)} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",borderBottom:"1px solid #e8e5de",cursor:"pointer"}} onMouseEnter={e=>(e.currentTarget.style.background="#fafaf8")} onMouseLeave={e=>(e.currentTarget.style.background="#fff")}>
              <div style={{width:36,height:36,borderRadius:"50%",background:"#1a1814",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,flexShrink:0}}>{initials}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:600}}>{cl.nombre}</div>
                <div style={{fontSize:11,color:"#9a9590"}}>{cl.eventos.length} evento(s){cl.tel?" · 📞 "+cl.tel:""}</div>
              </div>
              {cl.totalContratado>0&&(
                <div style={{textAlign:"right" as const,flexShrink:0}}>
                  <div style={{fontSize:12,fontWeight:700}}>{fmt(cl.totalContratado)}</div>
                  {saldo>0&&<div style={{fontSize:10,color:"#92580a",fontWeight:600}}>Saldo: {fmt(saldo)}</div>}
                  {saldo===0&&<div style={{fontSize:10,color:"#2d6a4f",fontWeight:600}}>✓ Liquidado</div>}
                </div>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}

// --- FINANZAS ---------------------------------
function FinanzasSection({contratos,token,pwd,pwdInput,setPwdInput,pwdError,onUnlock,pagoModal,setPagoModal,pagoMonto,setPagoMonto,pagoNota,setPagoNota,onAgregarPago}: any){
  const [vista,setVista]=useState("gerencia")
  const [filtroPeriodo,setFiltroPeriodo]=useState("todo")
  const [cots,setCots]=useState<any[]>([])
  const [cotsCargando,setCotsCargando]=useState(false)
  const [cotsLoaded,setCotsLoaded]=useState(false)
  const [slicerDesde,setSlicerDesde]=useState("")
  const [slicerHasta,setSlicerHasta]=useState("")
  const [filtroAno,setFiltroAno]=useState(String(new Date().getFullYear()))
  const [filtroTipoFin,setFiltroTipoFin]=useState("contrato")
  const [statsSortBy,setStatsSortBy]=useState("ingreso")
  const [statsBusq,setStatsBusq]=useState("")
  const [statsTopN,setStatsTopN]=useState(25)
  const [historialModal,setHistorialModal]=useState<string|null>(null)
  const [pagoMetodo,setPagoMetodo]=useState("efectivo")
  const [pagoFolio,setPagoFolio]=useState("")
  const [pagoVendedor,setPagoVendedor]=useState("")

  const METODOS: Record<string,{label:string;icon:string;color:string}> = {
    efectivo:     {label:"Efectivo",     icon:"💵",color:"#2d6a4f"},
    transferencia:{label:"Transferencia",icon:"🏦",color:"#1a3a5c"},
    tarjeta:      {label:"Tarjeta",      icon:"💳",color:"#4a2d6e"},
    otro:         {label:"Otro",         icon:"📝",color:"#92580a"},
  }

  if(!pwd) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:400}}>
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:16,padding:36,width:340,textAlign:"center" as const}}>
        <div style={{fontSize:36,marginBottom:12}}>🔒</div>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,marginBottom:4}}>Módulo Finanzas</div>
        <div style={{fontSize:12,color:"#9a9590",marginBottom:20}}>Ingresa tu contraseña para acceder</div>
        <input type="password" placeholder="Contraseña" value={pwdInput} onChange={e=>setPwdInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onUnlock()} style={{...S.input,textAlign:"center" as const,marginBottom:8,letterSpacing:4}}/>
        {pwdError&&<div style={{fontSize:12,color:"#8b2e2e",marginBottom:8}}>{pwdError}</div>}
        <button style={{...S.loginBtn,width:"100%",justifyContent:"center" as const,display:"flex"}} onClick={onUnlock}>Entrar</button>
      </div>
    </div>
  )

  // --- Filtrado ---
  const now = new Date()
  const cFilt = contratos.filter((c:Contrato)=>{
    const fe = new Date(c.fecha_evento+"T12:00:00")
    // Filtro de tipo
    if(filtroTipoFin!=="todos" && (c.tipo||"contrato")!==filtroTipoFin) return false
    // Filtro de año
    if(filtroAno && c.fecha_evento?.slice(0,4)!==filtroAno) return false
    // Slicer de fechas tiene prioridad
    if(slicerDesde && isoDate(fe) < slicerDesde) return false
    if(slicerHasta && isoDate(fe) > slicerHasta) return false
    if(slicerDesde||slicerHasta) return true
    if(filtroAno) return true
    // Periodo rápido
    if(filtroPeriodo==="dia") return fe.toDateString()===now.toDateString()
    if(filtroPeriodo==="semana"){
      const dow=now.getDay(),diff=dow===0?-6:1-dow
      const lunes=new Date(now);lunes.setDate(now.getDate()+diff);lunes.setHours(0,0,0,0)
      const dom=new Date(lunes);dom.setDate(lunes.getDate()+6);dom.setHours(23,59,59,999)
      return fe>=lunes&&fe<=dom
    }
    if(filtroPeriodo==="mes") return fe.getMonth()===now.getMonth()&&fe.getFullYear()===now.getFullYear()
    return true
  })

  const cConTotal = cFilt.filter((c:Contrato)=>(c.total||0)>0)
  const totalVentas   = cConTotal.reduce((s:number,c:Contrato)=>s+(c.total||0),0)
  const totalCobrado  = cConTotal.reduce((s:number,c:Contrato)=>s+(c.cobrado||0),0)
  const totalPendiente= totalVentas-totalCobrado
  const cPendientes   = cConTotal.filter((c:Contrato)=>(c.cobrado||0)<(c.total||0)).sort((a:Contrato,b:Contrato)=>a.fecha_evento.localeCompare(b.fecha_evento))
  const cLiquidados   = cConTotal.filter((c:Contrato)=>(c.cobrado||0)>=(c.total||0))

  // breakdown by method across all pagos
  const byMethod:Record<string,number>={}
  cConTotal.forEach((c:Contrato)=>(c.pagos||[]).forEach((p:Pago)=>{
    const m=(p as any).metodo||"efectivo"; byMethod[m]=(byMethod[m]||0)+p.monto
  }))

  // breakdown by vendedor
  const byVendedor:Record<string,number>={}
  cConTotal.forEach((c:Contrato)=>(c.pagos||[]).forEach((p:Pago)=>{
    const v=(p as any).vendedor||"—"; byVendedor[v]=(byVendedor[v]||0)+p.monto
  }))

  const historialContrato = historialModal ? contratos.find((x:Contrato)=>x.id===historialModal) : null

  function openPagoModal(cid:string){
    setPagoModal({cid,monto:"",nota:""})
    setPagoMonto(""); setPagoNota(""); setPagoMetodo("efectivo"); setPagoFolio(""); setPagoVendedor("")
  }

  return (
    <div>
      {/* ── HEADER + FILTROS ── */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 18px",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap" as const,gap:10,marginBottom:12}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>Finanzas</div>
          <div style={S.fset}>
            {[["gerencia","📊 Gerencia"],["resumen","Resumen"],["contratos","Por contrato"],["pendientes","Pendientes"],["comparativa","Año vs Año"],["stats","📊 Artículos"]].map(([v,l])=>(
              <button key={v} style={{...S.fbtn,...(vista===v?S.fbtnActive:{})}} onClick={()=>setVista(v)}>{l}</button>
            ))}
          </div>
        </div>
        {/* Filtros en una fila */}
        <div style={{display:"flex",flexWrap:"wrap" as const,gap:10,alignItems:"center"}}>
          {/* Año */}
          {(()=>{
            const anos=[...new Set(contratos.map((c:Contrato)=>c.fecha_evento?.slice(0,4)).filter(Boolean))].sort().reverse()
            return anos.length>0?(
              <div style={{display:"flex",alignItems:"center",gap:6,background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:8,padding:"4px 10px"}}>
                <span style={{fontSize:11,color:"#9a9590",fontWeight:600}}>Año:</span>
                <select value={filtroAno} onChange={e=>{setFiltroAno(e.target.value);setSlicerDesde("");setSlicerHasta("");setFiltroPeriodo("todo")}}
                  style={{border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:12,color:"#1a1814",outline:"none",cursor:"pointer"}}>
                  <option value="">Todos</option>
                  {anos.map((a:string)=><option key={a} value={a}>{a}</option>)}
                </select>
              </div>
            ):null
          })()}
          {/* Periodo rápido */}
          <div style={S.fset}>
            {[["todo","Todo"],["dia","Hoy"],["semana","Semana"],["mes","Mes"]].map(([v,l])=>(
              <button key={v} style={{...S.fbtn,...(filtroPeriodo===v&&!slicerDesde&&!slicerHasta&&!filtroAno?S.fbtnActive:{})}}
                onClick={()=>{setFiltroPeriodo(v);setSlicerDesde("");setSlicerHasta("");setFiltroAno("")}}>{l}</button>
            ))}
          </div>
          {/* Tipo */}
          <div style={{display:"flex",gap:4}}>
            {([["todos","Todos"],["contrato","Contratos"],["cotizacion","Cotizaciones"],["declinado","Declinados"]] as [string,string][]).map(([v,l])=>{
              const n=v==="todos"?contratos.length:contratos.filter((x:Contrato)=>(x.tipo||"contrato")===v).length
              return(
                <button key={v} onClick={()=>setFiltroTipoFin(v)}
                  style={{padding:"3px 10px",borderRadius:16,border:`1.5px solid ${filtroTipoFin===v?"#1a1814":"#e8e5de"}`,background:filtroTipoFin===v?"#1a1814":"#fff",color:filtroTipoFin===v?"#fff":"#4a4640",fontSize:10,fontWeight:filtroTipoFin===v?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif",whiteSpace:"nowrap" as const}}>
                  {filtroTipoFin===v?"✓ ":""}{l} ({n})
                </button>
              )
            })}
          </div>
          {/* Slicer de fechas */}
          <div style={{display:"flex",alignItems:"center",gap:6,background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:8,padding:"5px 12px"}}>
            <span style={{fontSize:11,color:"#9a9590",fontWeight:600}}>Rango:</span>
            <input type="date" value={slicerDesde} onChange={e=>{setSlicerDesde(e.target.value);setFiltroPeriodo("todo")}}
              style={{border:"none",background:"transparent",fontFamily:"monospace",fontSize:11,color:"#1a1814",outline:"none"}}/>
            <span style={{fontSize:11,color:"#9a9590"}}>→</span>
            <input type="date" value={slicerHasta} onChange={e=>{setSlicerHasta(e.target.value);setFiltroPeriodo("todo")}}
              style={{border:"none",background:"transparent",fontFamily:"monospace",fontSize:11,color:"#1a1814",outline:"none"}}/>
            {(slicerDesde||slicerHasta)&&(
              <button onClick={()=>{setSlicerDesde("");setSlicerHasta("")}} style={{background:"none",border:"none",color:"#9a9590",cursor:"pointer",fontSize:13,padding:0}}>✕</button>
            )}
          </div>
          {/* Contador */}
          <span style={{fontSize:11,color:"#9a9590",background:"#f5f4f0",padding:"3px 10px",borderRadius:8}}>
            {cConTotal.length} contrato{cConTotal.length!==1?"s":""}
            {filtroPeriodo!=="todo"&&<span style={{color:"#92580a"}}> · {filtroPeriodo==="dia"?"hoy":filtroPeriodo==="semana"?"esta semana":"este mes"}</span>}
            {filtroTipoFin!=="todos"&&<span style={{color:"#1a3a5c"}}> · {filtroTipoFin}</span>}
            {filtroAno&&<span style={{color:"#2d6a4f"}}> · {filtroAno}</span>}
          </span>
        </div>
      </div>

      {/* ── KPIs ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
        {[
          {v:fmt(totalVentas),  l:"Total vendido", c:"#1a1814"},
          {v:fmt(totalCobrado), l:"Cobrado",        c:"#2d6a4f"},
          {v:fmt(totalPendiente),l:"Por cobrar",   c:totalPendiente>0?"#92580a":"#2d6a4f"},
          {v:cLiquidados.length+"/"+cConTotal.length,l:"Liquidados",c:"#1a3a5c"},
        ].map((k,i)=>(
          <div key={i} style={S.kpi}>
            <div style={S.kl}>{k.l}</div>
            <div style={{...S.kv,fontSize:20,color:k.c,marginTop:4}}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* ── RESUMEN ── */}
      {/* ── GERENCIA ── */}
      {vista==="gerencia"&&(()=>{
        // Load cotizaciones on first render
        if(!cotsLoaded&&!cotsCargando&&token){
          setCotsCargando(true)
          fetch("/api/cotizaciones",{headers:{Authorization:`Bearer ${token}`}})
            .then(r=>r.json()).then(data=>{
              setCots(Array.isArray(data)?data:[])
              setCotsLoaded(true)
              setCotsCargando(false)
            }).catch(()=>setCotsCargando(false))
        }

        const hoy=new Date()
        hoy.setHours(0,0,0,0)
        const hoyStr=hoy.toISOString().slice(0,10)
        const dow=hoy.getDay()===0?6:hoy.getDay()-1
        const lunesStr=new Date(hoy.getTime()-dow*86400000).toISOString().slice(0,10)
        const domStr=new Date(hoy.getTime()+(6-dow)*86400000).toISOString().slice(0,10)

        // ── CONTRATOS métricas ──
        const ctHoy=contratos.filter((x:any)=>x.creado_en?.slice(0,10)===hoyStr||x.fecha_evento===hoyStr)
        const ctSemana=contratos.filter((x:any)=>{
          const f=x.fecha_evento||""
          return f>=lunesStr&&f<=domStr
        })
        const totalSemana=ctSemana.reduce((s:number,x:any)=>s+(x.total||0),0)

        // ── COTIZACIONES métricas ──
        const cotHoy=cots.filter((x:any)=>x.creado_en?.slice(0,10)===hoyStr)
        const cotSemana=cots.filter((x:any)=>(x.creado_en?.slice(0,10)||"")>=lunesStr&&(x.creado_en?.slice(0,10)||"")<=domStr)
        const cotConvertidas=cots.filter((x:any)=>x.estado==="convertida")
        const cotPendientes=cots.filter((x:any)=>x.estado==="borrador"||x.estado==="enviada")
        const conversionRate=cots.length>0?Math.round(cotConvertidas.length/cots.length*100):0

        // ── POR VENDEDOR (cotizaciones semana) ──
        const porVend:{[k:string]:{cots:number,contratos:number,monto:number}}={}
        cotSemana.forEach((x:any)=>{
          const v=x.vendedor||"Sin asignar"
          if(!porVend[v]) porVend[v]={cots:0,contratos:0,monto:0}
          porVend[v].cots++
          porVend[v].monto+=(x.total||0)
          if(x.estado==="convertida") porVend[v].contratos++
        })
        ctSemana.forEach((x:any)=>{
          const v=x.vendedor||vendedorDesdeFolio(x.folio||"")||"Sin asignar"
          if(!porVend[v]) porVend[v]={cots:0,contratos:0,monto:0}
        })

        // ── ACTIVIDAD RECIENTE (últimas 20 cotizaciones) ──
        const recientes=[...cots]
          .sort((a:any,b:any)=>(b.creado_en||"").localeCompare(a.creado_en||""))
          .slice(0,15)

        const ESTADO_COLOR:{[k:string]:string}={
          borrador:"#9a9590",enviada:"#1a3a5c",aceptada:"#2d6a4f",
          rechazada:"#8b2e2e",expirada:"#92580a",convertida:"#1a1814"
        }
        const ESTADO_BG:{[k:string]:string}={
          borrador:"#f5f4f0",enviada:"#edf3fa",aceptada:"#f0fdf4",
          rechazada:"#fdf0f0",expirada:"#fffbeb",convertida:"#f0f0f0"
        }

        return(
          <div style={{display:"flex",flexDirection:"column" as const,gap:16}}>

            {/* ── KPIs principales ── */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
              {[
                {l:"Cotiz. hoy",v:cotHoy.length,sub:`${cotSemana.length} esta semana`,c:"#1a3a5c",bg:"#edf3fa"},
                {l:"Contratos semana",v:ctSemana.length,sub:fmt(totalSemana),c:"#2d6a4f",bg:"#f0fdf4"},
                {l:"Conversión",v:`${conversionRate}%`,sub:`${cotConvertidas.length} de ${cots.length} total`,c:"#92580a",bg:"#fffbeb"},
                {l:"Pendientes",v:cotPendientes.length,sub:"En proceso",c:"#8b2e2e",bg:"#fdf0f0"},
              ].map((k,i)=>(
                <div key={i} style={{background:k.bg,border:`1.5px solid ${k.c}22`,borderRadius:12,padding:"14px 16px"}}>
                  <div style={{fontSize:10,fontWeight:700,color:k.c,textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:4}}>{k.l}</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:28,fontWeight:800,color:k.c,lineHeight:1}}>{k.v}</div>
                  <div style={{fontSize:10,color:k.c,opacity:.7,marginTop:4}}>{k.sub}</div>
                </div>
              ))}
            </div>

            {/* ── Grid: Actividad + Vendedores ── */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 320px",gap:14}}>

              {/* Actividad reciente */}
              <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
                <div style={{padding:"14px 16px",borderBottom:"1px solid #f0ece4",display:"flex",alignItems:"center",gap:8}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:800,flex:1}}>📋 Actividad reciente</div>
                  <span style={{fontSize:10,color:"#9a9590"}}>Últimas cotizaciones</span>
                </div>
                {cotsCargando?(
                  <div style={{padding:32,textAlign:"center" as const,color:"#9a9590",fontSize:12}}>Cargando...</div>
                ):recientes.length===0?(
                  <div style={{padding:32,textAlign:"center" as const,color:"#9a9590",fontSize:12}}>Sin cotizaciones</div>
                ):(
                  <div style={{maxHeight:420,overflowY:"auto" as const}}>
                    {recientes.map((x:any,i:number)=>{
                      const esHoy=x.creado_en?.slice(0,10)===hoyStr
                      return(
                        <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px",borderBottom:"1px solid #f9f7f4",background:esHoy?"#fffbf0":"#fff"}}>
                          {/* Dot */}
                          <div style={{width:8,height:8,borderRadius:"50%",background:ESTADO_COLOR[x.estado]||"#9a9590",flexShrink:0}}/>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                              <span style={{fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente_nombre||"Sin cliente"}</span>
                              {esHoy&&<span style={{fontSize:8,background:"#f59e0b",color:"#fff",padding:"1px 5px",borderRadius:3,fontWeight:700,flexShrink:0}}>HOY</span>}
                            </div>
                            <div style={{display:"flex",gap:6,flexWrap:"wrap" as const}}>
                              {x.folio&&<span style={{fontSize:9,fontFamily:"monospace",background:"#f5f4f0",padding:"0 4px",borderRadius:3}}>{x.folio}</span>}
                              {x.vendedor&&<span style={{fontSize:9,color:"#1a3a5c",fontWeight:600}}>👤 {x.vendedor}</span>}
                              {x.fecha_evento&&<span style={{fontSize:9,color:"#9a9590"}}>📅 {x.fecha_evento}</span>}
                            </div>
                          </div>
                          <div style={{textAlign:"right" as const,flexShrink:0}}>
                            {(x.total||0)>0&&<div style={{fontSize:11,fontFamily:"monospace",fontWeight:700}}>{fmt(x.total)}</div>}
                            <span style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:ESTADO_BG[x.estado]||"#f5f4f0",color:ESTADO_COLOR[x.estado]||"#9a9590",fontWeight:700}}>
                              {x.estado||"borrador"}
                            </span>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                )}
              </div>

              {/* Por vendedor */}
              <div style={{display:"flex",flexDirection:"column" as const,gap:10}}>

                {/* Tabla vendedores semana */}
                <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
                  <div style={{padding:"12px 16px",borderBottom:"1px solid #f0ece4",fontFamily:"Playfair Display,serif",fontSize:13,fontWeight:800}}>
                    👥 Vendedores — esta semana
                  </div>
                  <div>
                    {Object.entries(porVend).length===0?(
                      <div style={{padding:20,textAlign:"center" as const,color:"#9a9590",fontSize:11}}>Sin actividad esta semana</div>
                    ):Object.entries(porVend)
                      .sort(([,a],[,b])=>(b as any).cots-(a as any).cots)
                      .map(([vend,data]:any,i:number)=>(
                        <div key={i} style={{padding:"10px 16px",borderBottom:"1px solid #f9f7f4",display:"flex",alignItems:"center",gap:10}}>
                          <div style={{width:32,height:32,borderRadius:"50%",background:"#0f172a",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,flexShrink:0}}>
                            {vend.charAt(0)}
                          </div>
                          <div style={{flex:1}}>
                            <div style={{fontWeight:700,fontSize:12}}>{vend}</div>
                            <div style={{fontSize:10,color:"#9a9590",marginTop:1}}>
                              {data.cots} cotiz · {data.contratos} contratos
                            </div>
                          </div>
                          {data.monto>0&&<div style={{fontSize:11,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{fmt(data.monto)}</div>}
                        </div>
                      ))
                    }
                  </div>
                </div>

                {/* Contratos semana */}
                <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
                  <div style={{padding:"12px 16px",borderBottom:"1px solid #f0ece4",fontFamily:"Playfair Display,serif",fontSize:13,fontWeight:800}}>
                    📄 Contratos de la semana ({ctSemana.length})
                  </div>
                  <div style={{maxHeight:200,overflowY:"auto" as const}}>
                    {ctSemana.length===0?(
                      <div style={{padding:16,textAlign:"center" as const,color:"#9a9590",fontSize:11}}>Sin contratos esta semana</div>
                    ):ctSemana
                      .sort((a:any,b:any)=>(a.fecha_evento||"").localeCompare(b.fecha_evento||""))
                      .map((x:any,i:number)=>(
                        <div key={i} style={{padding:"8px 14px",borderBottom:"1px solid #f9f7f4",display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:28,flexShrink:0,textAlign:"center" as const,background:"#f5f4f0",borderRadius:5,padding:"3px 2px"}}>
                            <div style={{fontSize:13,fontWeight:800,lineHeight:1}}>{x.fecha_evento?.slice(8)||"—"}</div>
                            <div style={{fontSize:7,color:"#9a9590",textTransform:"uppercase" as const}}>
                              {["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][(parseInt(x.fecha_evento?.slice(5,7)||"1")-1)]||""}
                            </div>
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:11,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente||x.archivo}</div>
                            {x.vendedor&&<div style={{fontSize:9,color:"#1a3a5c"}}>👤 {x.vendedor}</div>}
                          </div>
                          {(x.total||0)>0&&<span style={{fontSize:10,fontFamily:"monospace",fontWeight:700,flexShrink:0}}>{fmt(x.total)}</span>}
                        </div>
                      ))
                    }
                  </div>
                </div>
              </div>
            </div>
          </div>
        )
      })()}

      {vista==="resumen"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {/* Izquierda: distribución + método + vendedor */}
          <div style={{display:"flex",flexDirection:"column" as const,gap:12}}>
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:12}}>Distribución</div>
              {[{l:"Cobrado",v:totalCobrado,c:"#2d6a4f"},{l:"Por cobrar",v:totalPendiente,c:"#92580a"}].map((b,i)=>(
                <div key={i} style={{marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#9a9590",marginBottom:3}}>
                    <span>{b.l}</span><span style={{fontWeight:600,color:"#1a1814"}}>{fmt(b.v)}</span>
                  </div>
                  <div style={{height:8,background:"#f5f4f0",borderRadius:4,overflow:"hidden"}}>
                    <div style={{height:"100%",borderRadius:4,background:b.c,width:totalVentas>0?Math.round(b.v/totalVentas*100)+"%":"0%"}}/>
                  </div>
                </div>
              ))}
            </div>
            {/* Por método */}
            {Object.keys(byMethod).length>0&&(
              <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:12}}>Por método de pago</div>
                {Object.entries(byMethod).map(([m,v])=>{
                  const met=METODOS[m]||METODOS.otro
                  const pct=totalCobrado>0?Math.round((v as number)/totalCobrado*100):0
                  return(
                    <div key={m} style={{marginBottom:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:3}}>
                        <span style={{fontSize:16}}>{met.icon}</span>
                        <span style={{fontSize:12,flex:1,color:"#4a4640"}}>{met.label}</span>
                        <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:met.color}}>{fmt(v as number)}</span>
                        <span style={{fontSize:10,color:"#9a9590"}}>{pct}%</span>
                      </div>
                      <div style={{height:5,background:"#f5f4f0",borderRadius:3,overflow:"hidden"}}>
                        <div style={{height:"100%",borderRadius:3,background:met.color,width:pct+"%"}}/>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
            {/* Por vendedor */}
            {Object.keys(byVendedor).filter(v=>v!=="—").length>0&&(
              <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:12}}>Por vendedor</div>
                {Object.entries(byVendedor).filter(([v])=>v!=="—").sort((a,b)=>(b[1] as number)-(a[1] as number)).map(([v,total])=>{
                  const pct=totalCobrado>0?Math.round((total as number)/totalCobrado*100):0
                  return(
                    <div key={v} style={{marginBottom:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:3}}>
                        <div style={{width:28,height:28,borderRadius:"50%",background:"#1a1814",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{v.charAt(0).toUpperCase()}</div>
                        <span style={{fontSize:12,flex:1,fontWeight:600}}>{v}</span>
                        <span style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#2d6a4f"}}>{fmt(total as number)}</span>
                        <span style={{fontSize:10,color:"#9a9590"}}>{pct}%</span>
                      </div>
                      <div style={{height:5,background:"#f5f4f0",borderRadius:3,overflow:"hidden",marginLeft:36}}>
                        <div style={{height:"100%",borderRadius:3,background:"#1a3a5c",width:pct+"%"}}/>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
          </div>
          {/* Derecha: TODOS los pendientes */}
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>Pendientes de pago</div>
              <span style={{fontSize:11,color:"#9a9590"}}>{cPendientes.length} contrato{cPendientes.length!==1?"s":""}</span>
            </div>
            <div style={{maxHeight:520,overflowY:"auto" as const,display:"flex",flexDirection:"column" as const,gap:6}}>
              {!cPendientes.length&&(
                <div style={{textAlign:"center" as const,padding:"32px 0",color:"#9a9590",fontSize:12}}>
                  <div style={{fontSize:28,opacity:.2,marginBottom:6}}>✓</div>
                  Todo cobrado
                </div>
              )}
              {cPendientes.map((c:Contrato)=>{
                const saldo=(c.total||0)-(c.cobrado||0)
                const pct=(c.total||0)>0?Math.round((c.cobrado||0)/(c.total||0)*100):0
                return(
                  <div key={c.id} style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:8,padding:"10px 12px"}}>
                    <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{c.cliente||c.archivo}</div>
                        <div style={{fontSize:10,color:"#9a9590"}}>{c.fecha_evento}{c.folio?" · #"+c.folio:""}</div>
                      </div>
                      <div style={{textAlign:"right" as const,flexShrink:0}}>
                        <div style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#92580a"}}>{fmt(saldo)}</div>
                        <div style={{fontSize:10,color:"#9a9590"}}>{fmt(c.cobrado||0)} cobrado</div>
                      </div>
                    </div>
                    <div style={{height:4,background:"#e8e5de",borderRadius:2,overflow:"hidden",marginBottom:6}}>
                      <div style={{height:"100%",borderRadius:2,background:"#1a3a5c",width:pct+"%"}}/>
                    </div>
                    <div style={{display:"flex",gap:5,justifyContent:"flex-end"}}>
                      {(c.pagos||[]).length>0&&(
                        <button onClick={()=>setHistorialModal(c.id)} style={{fontSize:10,background:"#edf3fa",border:"1px solid #b8ceea",borderRadius:6,color:"#1a3a5c",padding:"2px 8px",cursor:"pointer",fontWeight:600}}>
                          📋 {c.pagos.length} pago{c.pagos.length>1?"s":""}
                        </button>
                      )}
                      <button onClick={()=>openPagoModal(c.id)} style={{...S.iconBtn,padding:"2px 8px",fontSize:11}}>+ Pago</button>
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </div>
      )}

      {/* ── POR CONTRATO ── */}
      {vista==="contratos"&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:13}}>
            <thead>
              <tr style={{background:"#fafaf8"}}>
                {["Cliente","Fecha","Total","Cobrado","Saldo","Estado","Pagos",""].map((h,i)=>(
                  <th key={i} style={{padding:"9px 12px",textAlign:"left" as const,fontSize:10,fontWeight:700,textTransform:"uppercase" as const,color:"#9a9590",letterSpacing:".05em",borderBottom:"1px solid #e8e5de"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {cConTotal.sort((a:Contrato,b:Contrato)=>a.fecha_evento.localeCompare(b.fecha_evento)).map((c:Contrato,i:number)=>{
                const saldo=(c.total||0)-(c.cobrado||0)
                const liq=saldo<=0
                return(
                  <tr key={i} style={{borderBottom:"1px solid #e8e5de"}}>
                    <td style={{padding:"10px 12px",fontWeight:600,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{c.cliente||c.archivo}</td>
                    <td style={{padding:"10px 12px",color:"#9a9590",fontFamily:"monospace",fontSize:11}}>{c.fecha_evento}</td>
                    <td style={{padding:"10px 12px",fontWeight:700}}>{fmt(c.total||0)}</td>
                    <td style={{padding:"10px 12px",color:"#2d6a4f",fontWeight:600}}>{fmt(c.cobrado||0)}</td>
                    <td style={{padding:"10px 12px",color:liq?"#2d6a4f":"#92580a",fontWeight:700}}>{liq?"✓ Liq.":fmt(saldo)}</td>
                    <td style={{padding:"10px 12px"}}>
                      <span style={{fontSize:10,padding:"2px 8px",borderRadius:10,fontWeight:700,background:liq?"#edf7f2":saldo<(c.total||0)?"#fdf5e8":"#fdf0f0",color:liq?"#2d6a4f":saldo<(c.total||0)?"#92580a":"#8b2e2e"}}>
                        {liq?"Liquidado":saldo<(c.total||0)?"Parcial":"Pendiente"}
                      </span>
                    </td>
                    <td style={{padding:"10px 12px"}}>
                      {(c.pagos||[]).length>0&&(
                        <button onClick={()=>setHistorialModal(c.id)} style={{fontSize:11,background:"#edf3fa",border:"1px solid #b8ceea",borderRadius:6,color:"#1a3a5c",padding:"2px 8px",cursor:"pointer",fontWeight:600}}>
                          {c.pagos.length} pago{c.pagos.length>1?"s":""}
                        </button>
                      )}
                    </td>
                    <td style={{padding:"10px 12px",display:"flex",gap:4}}>
                      {!liq&&<button onClick={()=>openPagoModal(c.id)} style={{...S.iconBtn,padding:"3px 8px",fontSize:11}}>+ Pago</button>}
                      <button onClick={()=>{setPagoModal({cid:c.id,monto:"EDITAR",nota:""});setPagoMonto(String(c.total||0));setPagoNota(String(c.a_cuenta||0))}} style={{...S.iconBtn,padding:"3px 8px",fontSize:11,color:"#1a3a5c"}}>✏️</button>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* ── PENDIENTES ── */}
      {vista==="pendientes"&&(
        <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
          {!cPendientes.length&&(
            <div style={{textAlign:"center" as const,padding:"48px",background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
              <div style={{fontSize:36,opacity:.2}}>✓</div>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,color:"#4a4640"}}>Todo cobrado en este período</div>
            </div>
          )}
          {cPendientes.map((c:Contrato)=>{
            const saldo=(c.total||0)-(c.cobrado||0)
            const pct=(c.total||0)>0?Math.round((c.cobrado||0)/(c.total||0)*100):0
            return(
              <div key={c.id} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"14px 16px"}}>
                <div style={{display:"flex",alignItems:"flex-start",gap:10,marginBottom:10}}>
                  <div style={{flex:1}}>
                    <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>{c.cliente||c.archivo}</div>
                    <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>{c.fecha_evento}{c.folio?" · #"+c.folio:""}</div>
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    {(c.pagos||[]).length>0&&(
                      <button onClick={()=>setHistorialModal(c.id)} style={{...S.iconBtn,padding:"6px 12px",fontSize:12,color:"#1a3a5c"}}>📋 {c.pagos.length} pago{c.pagos.length>1?"s":""}</button>
                    )}
                    <button onClick={()=>openPagoModal(c.id)} style={{...S.loginBtn,padding:"6px 14px",fontSize:12}}>+ Pago</button>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:10}}>
                  {[{l:"Total",v:fmt(c.total||0)},{l:"Cobrado",v:fmt(c.cobrado||0),c:"#2d6a4f"},{l:"Saldo",v:fmt(saldo),c:"#92580a"}].map((k,i)=>(
                    <div key={i} style={{background:"#fafaf8",borderRadius:8,padding:"8px 10px"}}>
                      <div style={{fontSize:10,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:2}}>{k.l}</div>
                      <div style={{fontSize:14,fontWeight:700,color:k.c||"#1a1814"}}>{k.v}</div>
                    </div>
                  ))}
                </div>
                <div style={{height:6,background:"#f5f4f0",borderRadius:3,overflow:"hidden",marginBottom:6}}>
                  <div style={{height:"100%",borderRadius:3,background:"#1a3a5c",width:pct+"%"}}/>
                </div>
                {(c.pagos||[]).length>0&&(
                  <div style={{display:"flex",flexWrap:"wrap" as const,gap:5,marginTop:4}}>
                    {c.pagos.map((p:Pago,i:number)=>{
                      const met=METODOS[(p as any).metodo||"efectivo"]||METODOS.otro
                      return(
                        <span key={i} style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:"#fafaf8",border:"1px solid #e8e5de",color:"#4a4640"}}>
                          {met.icon} {fmt(p.monto)} · {p.fecha}{(p as any).vendedor?" · "+((p as any).vendedor):""}
                        </span>
                      )
                    })}
                  </div>
                )}
              </div>
            )
          })}
        </div>
      )}

      {/* ── AÑO VS AÑO ── */}
      {vista==="comparativa"&&(()=>{
        const añoActual=new Date().getFullYear()
        const añoAnterior=añoActual-1
        const MESES_CORTO=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
        // Agrupar todos los contratos (sin filtro de período) por año y mes
        // Comparativa siempre muestra ambos años — filtra por tipo pero no por período/año
        const todosConTotal=(filtroTipoFin==="todos"?contratos:contratos.filter((cx:Contrato)=>(cx.tipo||"contrato")===filtroTipoFin)).filter((cx:Contrato)=>(cx.total||0)>0)
        const porMes=(año:number)=>Array.from({length:12},(_,m)=>{
          const cMes=todosConTotal.filter((cx:Contrato)=>{
            const fe=new Date(cx.fecha_evento+"T12:00:00")
            return fe.getFullYear()===año&&fe.getMonth()===m
          })
          return{mes:m,contratos:cMes.length,vendido:cMes.reduce((s:number,cx:Contrato)=>s+(cx.total||0),0),cobrado:cMes.reduce((s:number,cx:Contrato)=>s+(cx.cobrado||0),0)}
        })
        const datosActual=porMes(añoActual)
        const datosAnterior=porMes(añoAnterior)
        const totalActual=datosActual.reduce((s,m)=>s+m.vendido,0)
        const totalAnterior=datosAnterior.reduce((s,m)=>s+m.vendido,0)
        const crecimiento=totalAnterior>0?Math.round((totalActual-totalAnterior)/totalAnterior*100):0
        const maxVal=Math.max(...datosActual.map(m=>m.vendido),...datosAnterior.map(m=>m.vendido),1)
        return(
          <div>
            {/* KPIs comparativos */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
              {[
                {l:`Total ${añoActual}`,v:fmt(totalActual),c:"#1a3a5c"},
                {l:`Total ${añoAnterior}`,v:fmt(totalAnterior),c:"#9a9590"},
                {l:"Crecimiento",v:(crecimiento>=0?"+":"")+crecimiento+"%",c:crecimiento>=0?"#2d6a4f":"#8b2e2e"},
              ].map((k,i)=>(
                <div key={i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:20,textAlign:"center" as const}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>{k.l}</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:28,fontWeight:800,color:k.c}}>{k.v}</div>
                </div>
              ))}
            </div>
            {/* Gráfica de barras por mes */}
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:20,marginBottom:16}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700,marginBottom:4}}>Ventas por mes</div>
              <div style={{display:"flex",gap:8,marginBottom:16,fontSize:11}}>
                <div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:12,height:12,borderRadius:2,background:"#1a3a5c"}}/>{añoActual}</div>
                <div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:12,height:12,borderRadius:2,background:"#d4cfc4"}}/>{añoAnterior}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(12,1fr)",gap:6,alignItems:"flex-end",height:200}}>
                {MESES_CORTO.map((mes,m)=>{
                  const da=datosActual[m],dp=datosAnterior[m]
                  const hA=da.vendido>0?Math.max(Math.round(da.vendido/maxVal*180),4):0
                  const hP=dp.vendido>0?Math.max(Math.round(dp.vendido/maxVal*180),4):0
                  return(
                    <div key={m} style={{display:"flex",flexDirection:"column" as const,alignItems:"center",gap:2}}>
                      <div style={{display:"flex",gap:2,alignItems:"flex-end",width:"100%",justifyContent:"center"}}>
                        <div title={`${añoActual}: ${fmt(da.vendido)}`} style={{width:"45%",height:hA,background:"#1a3a5c",borderRadius:"3px 3px 0 0",minHeight:0,cursor:"pointer"}}/>
                        <div title={`${añoAnterior}: ${fmt(dp.vendido)}`} style={{width:"45%",height:hP,background:"#d4cfc4",borderRadius:"3px 3px 0 0",minHeight:0,cursor:"pointer"}}/>
                      </div>
                      <div style={{fontSize:8,color:"#9a9590",textAlign:"center" as const,fontWeight:600}}>{mes}</div>
                    </div>
                  )
                })}
              </div>
            </div>
            {/* Tabla detalle por mes */}
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:12}}>
                <thead>
                  <tr style={{background:"#fafaf8"}}>
                    {["Mes","Contratos "+añoActual,"Vendido "+añoActual,"Contratos "+añoAnterior,"Vendido "+añoAnterior,"Δ Ventas"].map((h,i)=>(
                      <th key={i} style={{padding:"10px 14px",textAlign:"left" as const,fontSize:10,fontWeight:700,textTransform:"uppercase" as const,color:"#9a9590",letterSpacing:".05em",borderBottom:"1px solid #e8e5de"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {MESES_CORTO.map((mes,m)=>{
                    const da=datosActual[m],dp=datosAnterior[m]
                    const delta=da.vendido-dp.vendido
                    const pct=dp.vendido>0?Math.round(delta/dp.vendido*100):da.vendido>0?100:0
                    const hasData=da.vendido>0||dp.vendido>0
                    if(!hasData)return null
                    return(
                      <tr key={m} style={{borderBottom:"1px solid #e8e5de",background:m%2===0?"#fff":"#fafaf8"}}>
                        <td style={{padding:"10px 14px",fontWeight:700}}>{mes}</td>
                        <td style={{padding:"10px 14px",fontFamily:"monospace",fontSize:11}}>{da.contratos}</td>
                        <td style={{padding:"10px 14px",fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{fmt(da.vendido)}</td>
                        <td style={{padding:"10px 14px",fontFamily:"monospace",fontSize:11,color:"#9a9590"}}>{dp.contratos}</td>
                        <td style={{padding:"10px 14px",fontFamily:"monospace",color:"#9a9590"}}>{fmt(dp.vendido)}</td>
                        <td style={{padding:"10px 14px"}}>
                          <span style={{fontFamily:"monospace",fontWeight:700,fontSize:12,color:delta>=0?"#2d6a4f":"#8b2e2e"}}>
                            {delta>=0?"+":""}{fmt(delta)} <span style={{fontSize:10,opacity:.8}}>({pct>=0?"+":""}{pct}%)</span>
                          </span>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )
      })()}

      {/* ── STATS ARTÍCULOS ── */}
      {vista==="stats"&&(()=>{
        const statsMap:Record<string,{nombre:string,veces:number,cantidad:number,ingreso:number}>={}
        const base=cFilt.filter((x:Contrato)=>(x.tipo||"contrato")==="contrato")
        base.forEach((x:Contrato)=>{
          (x.articulos||[]).forEach((a:Articulo)=>{
            const k=a.nombre.trim().toLowerCase()
            if(!k||k.length<3) return
            if(!statsMap[k]) statsMap[k]={nombre:a.nombre.trim(),veces:0,cantidad:0,ingreso:0}
            statsMap[k].veces++
            statsMap[k].cantidad+=a.cantidad||0
            statsMap[k].ingreso+=a.importe||0
          })
        })
        const all=Object.values(statsMap)
        const sorted=(statsSortBy==="veces"?[...all].sort((a,b)=>b.veces-a.veces):statsSortBy==="cantidad"?[...all].sort((a,b)=>b.cantidad-a.cantidad):[...all].sort((a,b)=>b.ingreso-a.ingreso)).filter(s=>!statsBusq||s.nombre.toLowerCase().includes(statsBusq.toLowerCase()))
        const top=sorted.slice(0,statsTopN)
        const totalIng=all.reduce((s,a)=>s+a.ingreso,0)
        const maxVal=(statsSortBy==="veces"?top[0]?.veces:statsSortBy==="cantidad"?top[0]?.cantidad:top[0]?.ingreso)||1
        return(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
              {[{l:"Artículos únicos",v:all.length},{l:"Total rentas",v:all.reduce((s,a)=>s+a.veces,0)},{l:"Ingreso artículos",v:"$"+Math.round(totalIng).toLocaleString("es-MX")}].map((k,i)=>(
                <div key={i} style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px",textAlign:"center" as const}}>
                  <div style={{fontSize:9,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:4}}>{k.l}</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800}}>{k.v.toLocaleString?.()??k.v}</div>
                </div>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap" as const,alignItems:"center"}}>
              <div style={{display:"flex",gap:4}}>
                {([["ingreso","💰 Ingreso"],["veces","🔁 Rentas"],["cantidad","📦 Cantidad"]] as [string,string][]).map(([v,l])=>(
                  <button key={v} onClick={()=>setStatsSortBy(v)} style={{padding:"4px 10px",borderRadius:8,border:`1.5px solid ${statsSortBy===v?"#1a1814":"#e8e5de"}`,background:statsSortBy===v?"#1a1814":"#fff",color:statsSortBy===v?"#fff":"#4a4640",fontSize:10,fontWeight:statsSortBy===v?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>{l}</button>
                ))}
              </div>
              <input value={statsBusq} onChange={e=>setStatsBusq(e.target.value)} placeholder="Filtrar artículo..."
                style={{flex:1,padding:"5px 10px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:11,outline:"none",minWidth:120}}/>
              <span style={{fontSize:10,color:"#9a9590"}}>{sorted.length} artículos</span>
            </div>
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:11}}>
                <thead><tr style={{background:"#fafaf8"}}>
                  {["#","Artículo","Rentas","Cantidad","Ingreso","% ingreso"].map((h,i)=>(
                    <th key={i} style={{padding:"8px 12px",textAlign:i<=1?"left" as const:"center" as const,fontSize:9,fontWeight:700,textTransform:"uppercase" as const,color:"#9a9590",letterSpacing:".05em",borderBottom:"1px solid #e8e5de"}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {top.map((s,i)=>{
                    const barVal=statsSortBy==="veces"?s.veces:statsSortBy==="cantidad"?s.cantidad:s.ingreso
                    const pct=totalIng>0?Math.round(s.ingreso/totalIng*1000)/10:0
                    return(
                      <tr key={i} style={{borderBottom:"1px solid #e8e5de",background:i%2===0?"#fff":"#fafaf8"}}>
                        <td style={{padding:"8px 12px",fontFamily:"monospace",fontSize:10,color:"#9a9590",fontWeight:700}}>#{i+1}</td>
                        <td style={{padding:"8px 12px",maxWidth:250}}>
                          <div style={{fontWeight:600,fontSize:11,marginBottom:2}}>{s.nombre}</div>
                          <div style={{height:3,background:"#f5f4f0",borderRadius:2,overflow:"hidden"}}>
                            <div style={{height:"100%",borderRadius:2,background:statsSortBy==="ingreso"?"#2d6a4f":statsSortBy==="veces"?"#1a3a5c":"#92580a",width:Math.round(barVal/maxVal*100)+"%"}}/>
                          </div>
                        </td>
                        <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{s.veces}</td>
                        <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace",color:"#4a4640"}}>{s.cantidad.toLocaleString()}</td>
                        <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#2d6a4f"}}>${Math.round(s.ingreso).toLocaleString("es-MX")}</td>
                        <td style={{padding:"8px 12px",textAlign:"center" as const}}>
                          <span style={{fontSize:10,padding:"2px 7px",borderRadius:8,background:pct>=5?"#edf3fa":"#f5f4f0",color:pct>=5?"#1a3a5c":"#9a9590",fontWeight:600}}>{pct}%</span>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
              {sorted.length>statsTopN&&(
                <div style={{padding:"10px",textAlign:"center" as const,borderTop:"1px solid #e8e5de"}}>
                  <button onClick={()=>setStatsTopN(statsTopN+25)} style={{padding:"5px 16px",borderRadius:8,border:"1px solid #e8e5de",background:"#fafaf8",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11}}>
                    Ver más ({sorted.length-statsTopN} restantes)
                  </button>
                </div>
              )}
            </div>
          </div>
        )
      })()}

      {/* ── HISTORIAL MODAL ── */}
      {historialModal&&historialContrato&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:16,padding:28,width:520,maxHeight:"85vh",overflowY:"auto" as const,boxShadow:"0 20px 60px rgba(0,0,0,.2)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
              <div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800}}>Historial de pagos</div>
                <div style={{fontSize:12,color:"#9a9590"}}>{historialContrato.cliente||historialContrato.archivo}</div>
              </div>
              <button onClick={()=>setHistorialModal(null)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#9a9590"}}>✕</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:16}}>
              {[{l:"Total",v:fmt(historialContrato.total||0)},{l:"Cobrado",v:fmt(historialContrato.cobrado||0),c:"#2d6a4f"},{l:"Saldo",v:fmt((historialContrato.total||0)-(historialContrato.cobrado||0)),c:"#92580a"}].map((k,i)=>(
                <div key={i} style={{background:"#fafaf8",borderRadius:8,padding:"10px 12px",textAlign:"center" as const}}>
                  <div style={{fontSize:10,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:2}}>{k.l}</div>
                  <div style={{fontSize:16,fontWeight:700,color:k.c||"#1a1814"}}>{k.v}</div>
                </div>
              ))}
            </div>
            {!(historialContrato.pagos||[]).length?(
              <div style={{textAlign:"center" as const,padding:"24px",color:"#9a9590",fontSize:12}}>Sin pagos registrados</div>
            ):(
              <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
                {(historialContrato.pagos||[]).map((p:Pago,i:number)=>{
                  const met=METODOS[(p as any).metodo||"efectivo"]||METODOS.otro
                  return(
                    <div key={i} style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
                        <span style={{fontSize:22}}>{met.icon}</span>
                        <div style={{flex:1}}>
                          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" as const}}>
                            <span style={{fontFamily:"monospace",fontSize:16,fontWeight:700,color:met.color}}>{fmt(p.monto)}</span>
                            <span style={{fontSize:11,padding:"2px 8px",borderRadius:8,background:`${met.color}15`,color:met.color,fontWeight:600,border:`1px solid ${met.color}30`}}>{met.label}</span>
                            <span style={{fontSize:10,color:"#9a9590"}}>#{i+1}</span>
                          </div>
                          <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>📅 {p.fecha}</div>
                        </div>
                      </div>
                      <div style={{display:"flex",flexWrap:"wrap" as const,gap:10,fontSize:11}}>
                        {(p as any).vendedor&&<span style={{color:"#4a4640"}}>👤 <strong>{(p as any).vendedor}</strong></span>}
                        {(p as any).folio&&<span style={{color:"#4a4640"}}>🔖 Ref: <strong>{(p as any).folio}</strong></span>}
                        {p.nota&&<span style={{color:"#9a9590",fontStyle:"italic"}}>💬 {p.nota}</span>}
                      </div>
                    </div>
                  )
                })}
              </div>
            )}
            <div style={{marginTop:16,paddingTop:16,borderTop:"1px solid #e8e5de",display:"flex",justifyContent:"flex-end"}}>
              <button onClick={()=>{setHistorialModal(null);openPagoModal(historialContrato.id)}} style={{...S.loginBtn,padding:"8px 20px"}}>+ Agregar pago</button>
            </div>
          </div>
        </div>
      )}

      {/* ── PAGO MODAL ── */}
      {pagoModal&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:16,padding:28,width:440,maxHeight:"90vh",overflowY:"auto" as const,boxShadow:"0 20px 60px rgba(0,0,0,.2)"}}>
            {pagoModal.monto==="EDITAR"?(
              <div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,marginBottom:4}}>Editar importe</div>
                <div style={{fontSize:12,color:"#9a9590",marginBottom:16}}>{contratos.find((x:Contrato)=>x.id===pagoModal.cid)?.cliente||""}</div>
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>Total del contrato</div>
                  <input type="number" placeholder="0" value={pagoMonto} onChange={e=>setPagoMonto(e.target.value)} style={{...S.input,margin:0,fontSize:16,fontFamily:"monospace"}}/>
                </div>
                <div style={{marginBottom:20}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>A cuenta inicial</div>
                  <input type="number" placeholder="0" value={pagoNota} onChange={e=>setPagoNota(e.target.value)} style={{...S.input,margin:0,fontSize:16,fontFamily:"monospace"}}/>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <button style={{...S.iconBtn,flex:1,justifyContent:"center" as const}} onClick={()=>setPagoModal(null)}>Cancelar</button>
                  <button style={{...S.loginBtn,flex:1,justifyContent:"center" as const,display:"flex"}} onClick={()=>{const t=parseFloat(pagoMonto),ac=parseFloat(pagoNota)||0;if(t>0)onAgregarPago(pagoModal.cid,0,"__SET_TOTAL__:"+t+":"+ac)}}>Guardar</button>
                </div>
              </div>
            ):(
              <div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,marginBottom:2}}>Registrar pago</div>
                <div style={{fontSize:12,color:"#9a9590",marginBottom:16}}>{contratos.find((x:Contrato)=>x.id===pagoModal.cid)?.cliente||""}</div>
                {/* Método */}
                <div style={{marginBottom:14}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:6,textTransform:"uppercase" as const}}>Método de pago</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6}}>
                    {Object.entries(METODOS).map(([k,v])=>(
                      <button key={k} onClick={()=>setPagoMetodo(k)}
                        style={{padding:"10px 4px",borderRadius:8,border:`2px solid ${pagoMetodo===k?v.color:"#e8e5de"}`,background:pagoMetodo===k?v.color+"18":"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",textAlign:"center" as const}}>
                        <div style={{fontSize:20}}>{v.icon}</div>
                        <div style={{fontSize:10,fontWeight:700,color:pagoMetodo===k?v.color:"#9a9590",marginTop:3}}>{v.label}</div>
                      </button>
                    ))}
                  </div>
                </div>
                {/* Monto */}
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>Monto</div>
                  <input type="number" placeholder="0" value={pagoMonto} onChange={e=>setPagoMonto(e.target.value)} style={{...S.input,margin:0,fontSize:18,fontFamily:"monospace",fontWeight:700}}/>
                </div>
                {/* Vendedor */}
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>Vendedor <span style={{fontWeight:400,color:"#c4bfb8"}}>(quién vendió)</span></div>
                  <input placeholder="ej. Ana García, Carlos..." value={pagoVendedor} onChange={e=>setPagoVendedor(e.target.value)} style={{...S.input,margin:0}}/>
                </div>
                {/* Folio */}
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>Folio / Referencia <span style={{fontWeight:400,color:"#c4bfb8"}}>(opcional)</span></div>
                  <input placeholder="ej. TRF-0042, REF-123..." value={pagoFolio} onChange={e=>setPagoFolio(e.target.value)} style={{...S.input,margin:0}}/>
                </div>
                {/* Nota */}
                <div style={{marginBottom:20}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#9a9590",marginBottom:5,textTransform:"uppercase" as const}}>Comentarios <span style={{fontWeight:400,color:"#c4bfb8"}}>(opcional)</span></div>
                  <input placeholder="ej. Segunda parcialidad, pago mixto..." value={pagoNota} onChange={e=>setPagoNota(e.target.value)} style={{...S.input,margin:0}}/>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <button style={{...S.iconBtn,flex:1,justifyContent:"center" as const}} onClick={()=>setPagoModal(null)}>Cancelar</button>
                  <button style={{...S.loginBtn,flex:1,justifyContent:"center" as const,display:"flex",opacity:!pagoMonto||parseFloat(pagoMonto)<=0?.5:1}}
                    onClick={()=>{const m=parseFloat(pagoMonto);if(m>0)onAgregarPago(pagoModal.cid,m,pagoNota,pagoMetodo,pagoFolio,pagoVendedor)}}>
                    Guardar pago
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

// ─── BÚSQUEDA GLOBAL ───────────────────────────────────────────────
function BusquedaSection({contratos}:{contratos:Contrato[]}){
  const [q,setQ]=useState("")
  const [sel,setSel]=useState<string|null>(null)
  const fmt=(n:number)=>"$"+Math.round(n).toLocaleString("es-MX")

  const resultados = q.trim().length<2 ? [] : contratos.filter((x:Contrato)=>{
    const s=q.toLowerCase()
    return (x.cliente||"").toLowerCase().includes(s)
      ||(x.lugar||"").toLowerCase().includes(s)
      ||(x.tel||"").toLowerCase().includes(s)
      ||(x.folio||"").toLowerCase().includes(s)
      ||x.articulos?.some((a:Articulo)=>a.nombre.toLowerCase().includes(s))
  }).sort((a:Contrato,b:Contrato)=>b.fecha_evento.localeCompare(a.fecha_evento))

  // Group by cliente for summary
  const porCliente = resultados.reduce((acc:any,x:Contrato)=>{
    const k=x.cliente||x.archivo
    if(!acc[k]) acc[k]={nombre:k,contratos:[],totalVendido:0,totalCobrado:0,tel:x.tel}
    acc[k].contratos.push(x)
    acc[k].totalVendido+=(x.total||0)
    acc[k].totalCobrado+=(x.cobrado||0)
    return acc
  },{})
  const clientes=Object.values(porCliente) as any[]

  const selContrato = sel ? contratos.find((x:Contrato)=>x.id===sel) : null
  const clienteSelData = selContrato ? porCliente[selContrato.cliente||selContrato.archivo] : null

  const TIPO_COL:any={contrato:"#1a3a5c",cotizacion:"#2d6a4f",declinado:"#8b2e2e"}
  const TIPO_BG:any={contrato:"#edf3fa",cotizacion:"#edf7f2",declinado:"#fdf0f0"}

  return(
    <div style={{display:"grid",gridTemplateColumns:sel?"1fr 420px":"1fr",gap:16}}>
      {/* Columna principal */}
      <div>
        {/* Search bar */}
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 18px",marginBottom:16}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,marginBottom:12}}>🔍 Búsqueda global</div>
          <div style={{position:"relative" as const}}>
            <input autoFocus value={q} onChange={e=>setQ(e.target.value)}
              placeholder="Buscar cliente, lugar, teléfono, folio, artículo..."
              style={{width:"100%",padding:"12px 40px 12px 16px",border:"2px solid #1a1814",borderRadius:10,fontFamily:"Epilogue,sans-serif",fontSize:14,outline:"none",boxSizing:"border-box" as const}}/>
            {q&&<button onClick={()=>{setQ("");setSel(null)}} style={{position:"absolute" as const,right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#9a9590"}}>✕</button>}
          </div>
          {q.length>0&&q.length<2&&<div style={{fontSize:11,color:"#9a9590",marginTop:6}}>Escribe al menos 2 caracteres</div>}
          {q.length>=2&&<div style={{fontSize:11,color:"#9a9590",marginTop:6}}>{resultados.length} contratos · {clientes.length} clientes</div>}
        </div>

        {/* Results by client */}
        {q.length>=2&&clientes.length===0&&(
          <div style={{textAlign:"center" as const,padding:48,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
            <div style={{fontSize:32,opacity:.2}}>🔍</div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:14,color:"#9a9590"}}>Sin resultados para "{q}"</div>
          </div>
        )}

        <div style={{display:"flex",flexDirection:"column" as const,gap:10}}>
          {clientes.map((cli:any)=>{
            const isSelCli=sel&&(contratos.find((x:Contrato)=>x.id===sel)?.cliente||"")===(cli.nombre)
            return(
              <div key={cli.nombre} style={{background:"#fff",border:`1.5px solid ${isSelCli?"#1a1814":"#e8e5de"}`,borderRadius:12,overflow:"hidden"}}>
                {/* Client header */}
                <div style={{padding:"12px 16px",borderBottom:"1px solid #e8e5de",display:"flex",alignItems:"center",gap:12,background:isSelCli?"#fafaf8":"#fff"}}>
                  <div style={{width:40,height:40,borderRadius:"50%",background:"#1a1814",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,flexShrink:0}}>
                    {cli.nombre.charAt(0).toUpperCase()}
                  </div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700}}>{cli.nombre}</div>
                    <div style={{fontSize:11,color:"#9a9590"}}>{cli.tel||""} · {cli.contratos.length} contrato{cli.contratos.length!==1?"s":""}</div>
                  </div>
                  <div style={{textAlign:"right" as const}}>
                    <div style={{fontFamily:"monospace",fontSize:14,fontWeight:700,color:"#1a3a5c"}}>{fmt(cli.totalVendido)}</div>
                    <div style={{fontSize:10,color:cli.totalVendido-cli.totalCobrado>0?"#92580a":"#2d6a4f"}}>
                      {cli.totalVendido-cli.totalCobrado>0?`${fmt(cli.totalVendido-cli.totalCobrado)} pendiente`:"✓ Al corriente"}
                    </div>
                  </div>
                </div>
                {/* Contract list */}
                <div style={{padding:"8px 12px",display:"flex",flexDirection:"column" as const,gap:4}}>
                  {cli.contratos.map((x:Contrato)=>(
                    <div key={x.id} onClick={()=>setSel(sel===x.id?null:x.id)}
                      style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderRadius:8,cursor:"pointer",background:sel===x.id?"#f5f4f0":"transparent",border:`1px solid ${sel===x.id?"#d4cfc4":"transparent"}`}}>
                      <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,fontWeight:700,background:TIPO_BG[x.tipo||"contrato"],color:TIPO_COL[x.tipo||"contrato"],flexShrink:0}}>
                        {x.tipo==="contrato"?"CONT":x.tipo==="cotizacion"?"COT":"DEC"}
                      </span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.lugar||x.archivo}</div>
                        <div style={{fontSize:10,color:"#9a9590"}}>{x.fecha_evento}</div>
                      </div>
                      <div style={{textAlign:"right" as const,flexShrink:0}}>
                        <div style={{fontFamily:"monospace",fontSize:12,fontWeight:700}}>{fmt(x.total||0)}</div>
                        {(x.total||0)-(x.cobrado||0)>0&&<div style={{fontSize:9,color:"#92580a"}}>debe {fmt((x.total||0)-(x.cobrado||0))}</div>}
                        {(x.total||0)>0&&(x.cobrado||0)>=(x.total||0)&&<div style={{fontSize:9,color:"#2d6a4f"}}>✓ pagado</div>}
                      </div>
                      <span style={{color:"#9a9590",fontSize:12}}>{sel===x.id?"▾":"▸"}</span>
                    </div>
                  ))}
                </div>
              </div>
            )
          })}
        </div>
      </div>

      {/* Panel detalle */}
      {sel&&selContrato&&(
        <div style={{position:"sticky" as const,top:80,alignSelf:"flex-start" as const}}>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden",maxHeight:"85vh",overflowY:"auto" as const}}>
            {/* Header */}
            <div style={{padding:"16px 18px",borderBottom:"1px solid #e8e5de",background:"#fafaf8",display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
              <div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800}}>{selContrato.cliente||selContrato.archivo}</div>
                <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>{selContrato.fecha_evento} · {selContrato.lugar?.slice(0,40)}</div>
                <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,fontWeight:700,marginTop:4,display:"inline-block",background:TIPO_BG[selContrato.tipo||"contrato"],color:TIPO_COL[selContrato.tipo||"contrato"]}}>
                  {selContrato.tipo?.toUpperCase()}
                </span>
              </div>
              <button onClick={()=>setSel(null)} style={{background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#9a9590",marginLeft:8}}>✕</button>
            </div>
            {/* KPIs */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,padding:12}}>
              {[{l:"Total",v:fmt(selContrato.total||0)},{l:"Cobrado",v:fmt(selContrato.cobrado||0),c:"#2d6a4f"},{l:"Saldo",v:fmt((selContrato.total||0)-(selContrato.cobrado||0)),c:(selContrato.total||0)-(selContrato.cobrado||0)>0?"#92580a":"#2d6a4f"}].map((k,i)=>(
                <div key={i} style={{background:"#f5f4f0",borderRadius:8,padding:"10px 12px",textAlign:"center" as const}}>
                  <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:2}}>{k.l}</div>
                  <div style={{fontSize:15,fontWeight:700,color:k.c||"#1a1814"}}>{k.v}</div>
                </div>
              ))}
            </div>
            {/* Info */}
            <div style={{padding:"0 12px 12px"}}>
              {selContrato.tel&&<div style={{fontSize:12,marginBottom:4}}>📞 {selContrato.tel}</div>}
              {selContrato.folio&&<div style={{fontSize:12,marginBottom:4}}>🔖 Folio: {selContrato.folio}</div>}
              {selContrato.vendedor&&<div style={{fontSize:12,marginBottom:4}}>👤 Vendedor: {selContrato.vendedor}</div>}
              {selContrato.notas&&<div style={{fontSize:11,color:"#9a9590",marginBottom:8,fontStyle:"italic"}}>💬 {selContrato.notas}</div>}
            </div>
            {/* Pagos */}
            {(selContrato.pagos||[]).length>0&&(
              <div style={{padding:"0 12px 12px"}}>
                <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:6}}>Historial de pagos</div>
                {selContrato.pagos.map((p:Pago,i:number)=>(
                  <div key={i} style={{background:"#f5f4f0",borderRadius:8,padding:"8px 10px",marginBottom:4,display:"flex",gap:8,alignItems:"center"}}>
                    <div style={{flex:1}}>
                      <div style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#2d6a4f"}}>${p.monto?.toLocaleString()}</div>
                      <div style={{fontSize:10,color:"#9a9590"}}>{p.fecha} · {(p as any).metodo||"efectivo"}</div>
                      {(p as any).vendedor&&<div style={{fontSize:10,color:"#4a4640"}}>👤 {(p as any).vendedor}</div>}
                      {p.nota&&<div style={{fontSize:10,color:"#9a9590",fontStyle:"italic"}}>{p.nota}</div>}
                    </div>
                  </div>
                ))}
              </div>
            )}
            {/* Artículos */}
            {(selContrato.articulos||[]).length>0&&(
              <div style={{padding:"0 12px 12px"}}>
                <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:6}}>Artículos ({selContrato.articulos.length})</div>
                <div style={{maxHeight:200,overflowY:"auto" as const,display:"flex",flexDirection:"column" as const,gap:3}}>
                  {selContrato.articulos.map((a:Articulo,i:number)=>(
                    <div key={i} style={{display:"flex",gap:8,fontSize:11,padding:"4px 6px",background:"#fafaf8",borderRadius:5}}>
                      <span style={{color:"#9a9590",minWidth:20,textAlign:"right" as const}}>{a.cantidad}x</span>
                      <span style={{flex:1}}>{a.nombre}</span>
                      <span style={{fontFamily:"monospace",color:"#1a3a5c",fontWeight:600}}>${(a.importe||0).toLocaleString()}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {/* Historial completo del cliente */}
            {clienteSelData&&clienteSelData.contratos.length>1&&(
              <div style={{padding:"0 12px 12px",borderTop:"1px solid #e8e5de",paddingTop:12}}>
                <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:6}}>
                  Todos los contratos de este cliente ({clienteSelData.contratos.length})
                </div>
                {clienteSelData.contratos.filter((x:Contrato)=>x.id!==sel).map((x:Contrato)=>(
                  <div key={x.id} onClick={()=>setSel(x.id)} style={{display:"flex",gap:8,padding:"6px 8px",borderRadius:6,cursor:"pointer",marginBottom:3,background:"#f5f4f0"}}>
                    <span style={{fontSize:10,padding:"2px 6px",borderRadius:6,background:TIPO_BG[x.tipo||"contrato"],color:TIPO_COL[x.tipo||"contrato"],fontWeight:700,flexShrink:0}}>
                      {x.tipo==="contrato"?"C":x.tipo==="cotizacion"?"Q":"D"}
                    </span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.fecha_evento} · {x.lugar?.slice(0,25)||x.archivo}</div>
                    </div>
                    <span style={{fontFamily:"monospace",fontSize:11,fontWeight:700}}>${(x.total||0).toLocaleString()}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}


// ─── ESTADÍSTICAS ARTÍCULOS ────────────────────────────────────────
function ArticulosStatsSection({contratos}:{contratos:Contrato[]}){
  const [vista,setVista]=useState("renta")
  const [busq,setBusq]=useState("")
  const [topN,setTopN]=useState(20)
  const [statsTopN,setStatsTopN]=useState(25)
  const fmt=(n:number)=>"$"+Math.round(n).toLocaleString("es-MX")

  // Consolidate articles across all contracts
  const statsMap:Record<string,{nombre:string,veces:number,cantidad:number,ingreso:number,contratos:Set<string>}>={}
  contratos.filter((x:Contrato)=>(x.tipo||"contrato")==="contrato").forEach((x:Contrato)=>{
    (x.articulos||[]).forEach((a:Articulo)=>{
      const k=a.nombre.trim().toLowerCase()
      if(!k||k.length<3) return
      if(!statsMap[k]) statsMap[k]={nombre:a.nombre.trim(),veces:0,cantidad:0,ingreso:0,contratos:new Set()}
      statsMap[k].veces++
      statsMap[k].cantidad+=a.cantidad||0
      statsMap[k].ingreso+=a.importe||0
      statsMap[k].contratos.add(x.id)
    })
  })

  const allStats=Object.values(statsMap).map(s=>({...s,contratos:s.contratos.size}))
  const sorted = (vista==="renta"
    ? [...allStats].sort((a,b)=>b.veces-a.veces)
    : vista==="ingreso"
    ? [...allStats].sort((a,b)=>b.ingreso-a.ingreso)
    : [...allStats].sort((a,b)=>b.cantidad-a.cantidad)
  ).filter(s=>!busq||s.nombre.toLowerCase().includes(busq.toLowerCase()))

  const top=sorted.slice(0,statsTopN)
  const maxVeces=top[0]?.veces||1
  const maxIngreso=top[0]?.ingreso||1
  const maxCantidad=top[0]?.cantidad||1

  const totalIngresos=allStats.reduce((s,a)=>s+a.ingreso,0)
  const totalVeces=allStats.reduce((s,a)=>s+a.veces,0)

  return(
    <div>
      {/* Header KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
        {[
          {l:"Artículos únicos",v:allStats.length.toLocaleString(),c:"#1a1814"},
          {l:"Total rentas",v:totalVeces.toLocaleString(),c:"#1a3a5c"},
          {l:"Ingreso total artículos",v:fmt(totalIngresos),c:"#2d6a4f"},
          {l:"Promedio por artículo",v:fmt(totalIngresos/(allStats.length||1)),c:"#92580a"},
        ].map((k,i)=>(
          <div key={i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"16px 18px"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:6}}>{k.l}</div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:k.c}}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Controls */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"12px 16px",marginBottom:16,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap" as const}}>
        <div style={{display:"flex",gap:4}}>
          {([["renta","🔁 Por veces rentado"],["ingreso","💰 Por ingreso"],["cantidad","📦 Por cantidad"]] as [string,string][]).map(([v,l])=>(
            <button key={v} onClick={()=>setVista(v)} style={{padding:"5px 12px",borderRadius:8,border:`1.5px solid ${vista===v?"#1a1814":"#e8e5de"}`,background:vista===v?"#1a1814":"#fff",color:vista===v?"#fff":"#4a4640",fontSize:11,fontWeight:vista===v?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>{l}</button>
          ))}
        </div>
        <div style={{flex:1,position:"relative" as const}}>
          <input value={busq} onChange={e=>setBusq(e.target.value)} placeholder="Filtrar artículo..."
            style={{width:"100%",padding:"6px 12px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}/>
          {busq&&<button onClick={()=>setBusq("")} style={{position:"absolute" as const,right:8,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#9a9590"}}>✕</button>}
        </div>
        <span style={{fontSize:11,color:"#9a9590"}}>{sorted.length} artículos</span>
      </div>

      {/* Table */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:12}}>
          <thead>
            <tr style={{background:"#fafaf8"}}>
              {["#","Artículo","Veces rentado","Cantidad total","Ingreso total","% del ingreso","Contratos"].map((h,i)=>(
                <th key={i} style={{padding:"10px 14px",textAlign:i===0||i>=2?"center" as const:"left" as const,fontSize:10,fontWeight:700,textTransform:"uppercase" as const,color:"#9a9590",letterSpacing:".05em",borderBottom:"1px solid #e8e5de"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {top.map((s,i)=>{
              const barW=(vista==="renta"?s.veces/maxVeces:vista==="ingreso"?s.ingreso/maxIngreso:s.cantidad/maxCantidad)*100
              const pct=totalIngresos>0?Math.round(s.ingreso/totalIngresos*1000)/10:0
              return(
                <tr key={i} style={{borderBottom:"1px solid #e8e5de",background:i%2===0?"#fff":"#fafaf8"}}>
                  <td style={{padding:"10px 14px",textAlign:"center" as const,fontFamily:"monospace",fontSize:11,color:"#9a9590",fontWeight:700}}>#{i+1}</td>
                  <td style={{padding:"10px 14px",maxWidth:300}}>
                    <div style={{fontWeight:600,marginBottom:3}}>{s.nombre}</div>
                    <div style={{height:4,background:"#f5f4f0",borderRadius:2,overflow:"hidden"}}>
                      <div style={{height:"100%",borderRadius:2,background:vista==="renta"?"#1a3a5c":vista==="ingreso"?"#2d6a4f":"#92580a",width:barW+"%"}}/>
                    </div>
                  </td>
                  <td style={{padding:"10px 14px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{s.veces}</td>
                  <td style={{padding:"10px 14px",textAlign:"center" as const,fontFamily:"monospace",color:"#4a4640"}}>{s.cantidad.toLocaleString()}</td>
                  <td style={{padding:"10px 14px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#2d6a4f"}}>{fmt(s.ingreso)}</td>
                  <td style={{padding:"10px 14px",textAlign:"center" as const}}>
                    <span style={{fontSize:11,padding:"2px 8px",borderRadius:8,background:pct>=5?"#edf3fa":"#f5f4f0",color:pct>=5?"#1a3a5c":"#9a9590",fontWeight:600}}>{pct}%</span>
                  </td>
                  <td style={{padding:"10px 14px",textAlign:"center" as const,color:"#9a9590",fontSize:11}}>{s.contratos}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
        {sorted.length>statsTopN&&(
          <div style={{padding:"12px 16px",textAlign:"center" as const,borderTop:"1px solid #e8e5de"}}>
            <button onClick={()=>setTopN(prev=>prev+20)} style={{padding:"6px 20px",borderRadius:8,border:"1px solid #e8e5de",background:"#fafaf8",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12}}>
              Ver más ({sorted.length-statsTopN} restantes)
            </button>
          </div>
        )}
      </div>
    </div>
  )
}


// ─── CARGA DE CAMIÓN ───────────────────────────────────────────────
const UNIDADES_CARGA = ["LD-36-696","LD-36-352","LF-43068","LF-63073","LG-90-921","ESTAQUITA"]
const COL_U = ["#1a3a5c","#2d6a4f","#92580a","#4a2d6e","#8b2e2e","#1a5c4a"]

type ModoMov = "entrega"|"desmonte"
type PendienteKey = `pend_${string}` // fecha -> ids[]

function CargaSection({contratos}:{contratos:Contrato[]}){
  const hoy = isoDate(new Date())
  const [fecha, setFecha] = useState(hoy)
  // asignaciones: { [contratoId]: { unidad, modo } }
  const [asig, setAsig] = useState<Record<string,{unidad:string,modo:ModoMov}>>({})
  const [orden, setOrden] = useState<Record<string,string[]>>({}) // unidad -> ids ordenados
  const [check, setCheck] = useState<Record<string,Record<string,boolean>>>({})
  const [notas, setNotas] = useState<Record<string,string>>({})
  const [expandido, setExpandido] = useState<string|null>(null)
  const [unidadActiva, setUnidadActiva] = useState(UNIDADES_CARGA[0])
  const [vista, setVista] = useState<"planear"|"unidad">("planear")
  const [reasigModal, setReasigModal] = useState<string|null>(null) // contratoId
  const [unidadesDisp, setUnidadesDisp] = useState<string[]>([...UNIDADES_CARGA])
  const [msgModal, setMsgModal] = useState<{unidad:string,txt:string}|null>(null)
  const [msgEdit, setMsgEdit] = useState("")
  const [copiado, setCopiado] = useState(false)
  const [guardado, setGuardado] = useState(false)
  const [pendFwd, setPendFwd] = useState<string[]>([]) // ids marcados para pasar al día siguiente

  // Load / save per date
  const STORE_KEY = `carga_${fecha}`
  useState(()=>{
    try {
      const s = localStorage.getItem(STORE_KEY)
      if(s){ const d=JSON.parse(s); setAsig(d.asig||{}); setOrden(d.orden||{}); setCheck(d.check||{}); setNotas(d.notas||{}) }
      else { setAsig({}); setOrden({}); setCheck({}); setNotas({}) }
    } catch{}
  })

  const save = () => {
    localStorage.setItem(STORE_KEY, JSON.stringify({asig,orden,check,notas}))
    setGuardado(true); setTimeout(()=>setGuardado(false),2000)
  }

  // Contratos del día por tipo de movimiento
  const delDia = (modo:ModoMov) => contratos.filter((x:Contrato)=>{
    if((x.tipo||"contrato")!=="contrato") return false
    return modo==="entrega" ? x.fecha_entrega===fecha : x.fecha_desmonte===fecha
  })
  const entregas = delDia("entrega")
  const desmontes = delDia("desmonte")
  const todos = [...entregas,...desmontes.filter(d=>!entregas.find(e=>e.id===d.id))]

  // Sin asignar
  const sinAsig = todos.filter((x:Contrato)=>!asig[x.id])

  // Por unidad ordenado
  const porUnidad = (u:string) => {
    const base = Object.entries(asig).filter(([,v])=>v.unidad===u).map(([id])=>contratos.find(x=>x.id===id)).filter(Boolean) as Contrato[]
    const ord = orden[u]||[]
    if(!ord.length) return base
    return [...base].sort((a,b)=>{const ia=ord.indexOf(a.id),ib=ord.indexOf(b.id); if(ia<0&&ib<0)return 0; if(ia<0)return 1; if(ib<0)return -1; return ia-ib})
  }

  // Progreso checklist
  const prog = (cid:string) => {
    const x = contratos.find(c=>c.id===cid); if(!x) return {done:0,total:0}
    const arts = x.articulos||[]; const done=arts.filter((_,i)=>check[cid]?.[String(i)]).length
    return {done,total:arts.length}
  }

  // Consolidado carga por unidad
  const consolidado = (u:string) => {
    const map:Record<string,{nombre:string,cantidad:number,clientes:string[],modo:string}>={}
    porUnidad(u).forEach(x=>{
      const modo=asig[x.id]?.modo||"entrega"
      ;(x.articulos||[]).forEach((a:Articulo)=>{
        const k=a.nombre.trim().toLowerCase()
        if(!map[k]) map[k]={nombre:a.nombre.trim(),cantidad:0,clientes:[],modo}
        map[k].cantidad+=a.cantidad||0
        if(!map[k].clientes.includes(x.cliente)) map[k].clientes.push(x.cliente)
      })
    })
    return Object.values(map).sort((a,b)=>b.cantidad-a.cantidad)
  }

  // Asignar
  const asignar = (id:string, unidad:string, modo:ModoMov) => setAsig(prev=>({...prev,[id]:{unidad,modo}}))
  const desasignar = (id:string) => setAsig(prev=>{const n={...prev};delete n[id];return n})
  const reasignar = (id:string, nuevaUnidad:string) => {
    setAsig(prev=>({...prev,[id]:{...prev[id],unidad:nuevaUnidad}}))
    setReasigModal(null)
  }

  // Reordenar
  const mover = (u:string, id:string, dir:1|-1) => {
    const curr = orden[u]||porUnidad(u).map(x=>x.id)
    const i=curr.indexOf(id); if(i<0) return
    const ni=i+dir; if(ni<0||ni>=curr.length) return
    const next=[...curr]; [next[i],next[ni]]=[next[ni],next[i]]
    setOrden({...orden,[u]:next})
  }

  // Pasar pendientes al día siguiente
  const pasarAlSiguiente = () => {
    const siguienteFecha = new Date(fecha+"T12:00:00"); siguienteFecha.setDate(siguienteFecha.getDate()+1)
    const sig = isoDate(siguienteFecha)
    try {
      const existing = localStorage.getItem(`carga_${sig}`)
      const d = existing ? JSON.parse(existing) : {asig:{},orden:{},check:{},notas:{}}
      // Los pendientes se pasan sin unidad asignada al día siguiente (para que los asignen)
      // Solo guardamos nota de que venían del día anterior
      pendFwd.forEach(id=>{
        const nota = `(pendiente del ${fecha})`
        d.notas[id] = nota
      })
      localStorage.setItem(`carga_${sig}`, JSON.stringify(d))
      // Quitar de hoy
      const newAsig={...asig}; pendFwd.forEach(id=>delete newAsig[id])
      setAsig(newAsig); setPendFwd([])
      alert(`✓ ${pendFwd.length} contrato(s) pasados al ${sig}`)
    } catch(e){ alert("Error al pasar pendientes") }
  }

  // Generar mensaje
  // Mensaje detallado por unidad (para el operador)
  const generarMsg = (u:string) => {
    const cs = porUnidad(u)
    const csEnt = cs.filter(x=>asig[x.id]?.modo==="entrega")
    const csDes = cs.filter(x=>asig[x.id]?.modo==="desmonte")
    const fechaStr = new Date(fecha+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long"})
    let msg = `🚚 *RUTA ${u}*
📅 ${fechaStr}
`
    if(csEnt.length>0){
      msg += `
*📦 ENTREGAS — ${csEnt.length} parada${csEnt.length>1?"s":""}*
`
      csEnt.forEach((x,i)=>{
        const arts=(x.articulos||[]).slice(0,5)
        const total=(x.articulos||[]).reduce((s:number,a:Articulo)=>s+a.cantidad,0)
        msg+=`
${i+1}️⃣ *${(x.cliente||"").toUpperCase()}*`
        if(x.folio&&x.folio.length>1) msg+=` | Folio: ${x.folio}`
        msg+=`
   📍 ${(x.lugar||"Sin dirección").slice(0,70)}
`
        if(x.tel) msg+=`   📞 ${x.tel}
`
        arts.forEach((a:Articulo)=>msg+=`   • ${a.cantidad}x ${a.nombre.slice(0,45)}
`)
        if((x.articulos||[]).length>5) msg+=`   • ...y ${(x.articulos||[]).length-5} artículos más
`
        msg+=`   ✅ *${total} piezas en total*
`
        if(notas[x.id]) msg+=`   ⚠️ ${notas[x.id]}
`
      })
    }
    if(csDes.length>0){
      msg+=`
*📤 DESMONTES — ${csDes.length} recogida${csDes.length>1?"s":""}*
`
      csDes.forEach((x,i)=>{
        const total=(x.articulos||[]).reduce((s:number,a:Articulo)=>s+a.cantidad,0)
        msg+=`
${i+1}️⃣ *${(x.cliente||"").toUpperCase()}*`
        if(x.folio&&x.folio.length>1) msg+=` | Folio: ${x.folio}`
        msg+=`
   📍 ${(x.lugar||"Sin dirección").slice(0,70)}
`
        if(x.tel) msg+=`   📞 ${x.tel}
`
        if(total>0) msg+=`   📦 *${total} piezas a recoger*
`
        if(notas[x.id]) msg+=`   ⚠️ ${notas[x.id]}
`
      })
    }
    msg+=`
— *Poliflor* 🌸`
    return msg
  }

  // Resumen general del día (todas las unidades, compacto)
  const generarResumenGeneral = () => {
    const unidadesActivas = UNIDADES_CARGA.filter(u=>porUnidad(u).length>0)
    const fechaStr = new Date(fecha+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long"})
    let msg = `📋 *RESUMEN GENERAL DE RUTAS*
📅 ${fechaStr}
`
    msg += `🚚 ${entregas.length} entregas · 📦 ${desmontes.length} desmontes · ${unidadesActivas.length} unidades
`
    msg += `${"─".repeat(30)}
`

    unidadesActivas.forEach(u=>{
      const cs = porUnidad(u)
      const csEnt = cs.filter(x=>asig[x.id]?.modo==="entrega")
      const csDes = cs.filter(x=>asig[x.id]?.modo==="desmonte")
      msg += `
*🚛 ${u}* (${cs.length} paradas)
`
      cs.forEach((x,i)=>{
        const modo = asig[x.id]?.modo==="entrega" ? "🚚" : "📦"
        const folio = x.folio&&x.folio.length>1 ? ` [${x.folio}]` : ""
        const dir = (x.lugar||"Sin dirección").slice(0,50)
        msg += `  ${i+1}. ${modo} *${x.cliente||x.archivo}*${folio}
`
        msg += `      📍 ${dir}
`
      })
    })

    if(sinAsig.length>0){
      msg += `
⚠️ *SIN ASIGNAR (${sinAsig.length})*
`
      sinAsig.forEach(x=>{
        const modo = desmontes.find(d=>d.id===x.id)?"📦":"🚚"
        msg += `  ${modo} ${x.cliente||x.archivo}
`
      })
    }
    msg += `
— *Poliflor* 🌸`
    return msg
  }

  const ui = (u:string) => COL_U[UNIDADES_CARGA.indexOf(u)]||"#1a1814"

  // Cancelar todos los cambios
  const cancelarCambios = () => {
    if(!window.confirm("¿Cancelar todos los cambios del día? Esto borrará asignaciones, orden y notas.")) return
    setAsig({}); setOrden({}); setCheck({}); setNotas({}); setPendFwd([]); setExpandido(null)
    localStorage.removeItem(STORE_KEY)
  }

  // Imprimir resumen
  const imprimirResumen = () => {
    const unidadesActivas = UNIDADES_CARGA.filter(u=>porUnidad(u).length>0)
    const fechaStr = new Date(fecha+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})
    const filas = unidadesActivas.map(u=>{
      const cs = porUnidad(u)
      const rows = cs.map((x,i)=>{
        const modo = asig[x.id]?.modo==="entrega" ? "ENTREGA" : "DESMONTE"
        const folio = x.folio&&x.folio.length>1 ? x.folio : "—"
        const total = (x.articulos||[]).reduce((s:number,a:Articulo)=>s+a.cantidad,0)
        const zona = detectarZona(x.lugar)<ZONAS_ORDEN.length ? ZONAS_ORDEN[detectarZona(x.lugar)][0] : ""
        return `<tr style="border-bottom:1px solid #e8e8e8">
          <td style="padding:7px 10px;font-weight:600;text-align:center;color:#666">${i+1}</td>
          <td style="padding:7px 10px"><span style="font-size:10px;padding:2px 6px;border-radius:4px;background:${modo==="ENTREGA"?"#edf3fa":"#f5f0fc"};color:${modo==="ENTREGA"?"#1a3a5c":"#4a2d6e"};font-weight:700">${modo==="ENTREGA"?"🚚":"📦"} ${modo}</span></td>
          <td style="padding:7px 10px;font-weight:700">${x.cliente||x.archivo}</td>
          <td style="padding:7px 10px;font-family:monospace;font-size:11px">${folio}</td>
          <td style="padding:7px 10px;font-size:11px;color:#555">${(x.lugar||"Sin dirección").slice(0,55)}</td>
          <td style="padding:7px 10px;font-size:11px;color:#555">${x.tel||"—"}</td>
          <td style="padding:7px 10px;text-align:center;font-weight:700;color:#92580a">${total}</td>
          ${notas[x.id]?`<td style="padding:7px 10px;font-size:10px;color:#92580a;font-style:italic">⚠️ ${notas[x.id]}</td>`:`<td></td>`}
        </tr>`
      }).join("")
      const uidx = UNIDADES_CARGA.indexOf(u)
      const col = COL_U[uidx]
      return `
        <div style="margin-bottom:24px;page-break-inside:avoid">
          <div style="background:${col};color:#fff;padding:8px 14px;border-radius:8px 8px 0 0;display:flex;align-items:center;gap:10px">
            <span style="font-size:16px;font-weight:800">${u}</span>
            <span style="font-size:12px;opacity:.8">${cs.length} paradas · ${cs.filter(x=>asig[x.id]?.modo==="entrega").length} entregas · ${cs.filter(x=>asig[x.id]?.modo==="desmonte").length} desmontes</span>
          </div>
          <table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-top:none">
            <thead>
              <tr style="background:#f8f8f8">
                <th style="padding:6px 10px;font-size:10px;text-align:center;color:#888;border-bottom:2px solid #e0e0e0">#</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">TIPO</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">CLIENTE</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">FOLIO</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">DIRECCIÓN</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">TEL</th>
                <th style="padding:6px 10px;font-size:10px;text-align:center;color:#888;border-bottom:2px solid #e0e0e0">PZAS</th>
                <th style="padding:6px 10px;font-size:10px;color:#888;border-bottom:2px solid #e0e0e0">NOTA</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>`
    }).join("")

    const sinAsigRows = sinAsig.length>0 ? `
      <div style="margin-top:16px;padding:10px 14px;background:#fff8f0;border:1.5px solid #e8d4b8;border-radius:8px">
        <div style="font-weight:700;color:#92580a;margin-bottom:6px">⚠️ Sin asignar (${sinAsig.length})</div>
        ${sinAsig.map(x=>`<div style="font-size:12px;padding:2px 0">${x.cliente||x.archivo} — ${(x.lugar||"").slice(0,50)}</div>`).join("")}
      </div>` : ""

    const html = `<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<title>Hoja de Carga — ${fechaStr}</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family: Arial, sans-serif; font-size: 13px; color: #1a1814; padding: 20px; }
  @media print {
    body { padding: 10px; }
    .no-print { display: none !important; }
    @page { margin: 15mm; size: A4 landscape; }
  }
</style>
</head><body>
<div class="no-print" style="text-align:center;margin-bottom:16px">
  <button onclick="window.print()" style="padding:10px 24px;background:#1a1814;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">🖨️ Imprimir</button>
  <button onclick="window.close()" style="padding:10px 24px;background:#f5f4f0;border:1px solid #ccc;border-radius:8px;font-size:14px;cursor:pointer;margin-left:8px">✕ Cerrar</button>
</div>
<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:20px;padding-bottom:12px;border-bottom:2px solid #1a1814">
  <div>
    <div style="font-size:22px;font-weight:800;font-family:Georgia,serif">🚚 Hoja de Carga — Poliflor</div>
    <div style="font-size:14px;color:#666;margin-top:3px;text-transform:capitalize">${fechaStr}</div>
  </div>
  <div style="text-align:right;font-size:12px;color:#666">
    <div>${entregas.length} entregas · ${desmontes.length} desmontes</div>
    <div>${unidadesActivas.length} unidades activas</div>
    <div style="margin-top:4px;font-size:10px">Generado: ${new Date().toLocaleString("es-MX")}</div>
  </div>
</div>
${filas}
${sinAsigRows}
</body></html>`

    const w = window.open("","_blank","width=1100,height=800")
    if(w){ w.document.write(html); w.document.close() }
  }

  // ── Zonas en orden lógico de ruta (NO-O-SO-S-SE-E-NE) ──────────
  const ZONAS_ORDEN = [
    ["INTERLOMAS","BOSQUE REAL","NAUCALPAN","HUIXQUILUCAN"],          // NW
    ["SANTA FE","CUAJIMALPA","LOMAS DE SANTA FE","ZEDEC","PASEO DE LA REFORMA"],
    ["LOMAS","BOSQUES DE LAS LOMAS","LA HERRADURA","TECAMACHALCO","BOSQUE DE LAS LOMAS","BOSQUES DE REFORMA"],
    ["POLANCO","LOS MORALES","LAGO","ANAHUAC"],
    ["PEDREGAL","JARDINES DEL PEDREGAL","LOMAS DEL PEDREGAL","VILLA OBREGON"],
    ["COYOACAN","TLALPAN","PEDREGAL DE SAN ÁNGEL","SAN ÁNGEL"],
    ["SATÉLITE","LINDAVISTA","ECATEPEC","TLALNEPANTLA"],
    ["CONDESA","ROMA","DOCTORES","JUÁREZ","CUAUHTÉMOC"],
    ["TLALPAN","XOCHIMILCO","TLAHUAC","MILPA ALTA"],
  ]

  const detectarZona = (lugar:string):number => {
    const u = (lugar||"").toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g,"")
    for(let i=0;i<ZONAS_ORDEN.length;i++){
      if(ZONAS_ORDEN[i].some(z=>u.includes(z.normalize("NFD").replace(/[̀-ͯ]/g,"")))) return i
    }
    return 99 // sin zona detectada va al final
  }

  // Auto-asignar todos los sin asignar distribuyendo por zona y carga
  const sugerirAsignacion = () => {
    if(sinAsig.length===0) return
    const nuevas = {...asig}
    // Contar carga actual por unidad
    const cargaU:Record<string,number> = {}
    const uActivas = UNIDADES_CARGA.filter(u=>unidadesDisp.includes(u))
    uActivas.forEach(u=>{ cargaU[u]=porUnidad(u).length })

    // Ordenar sin asignar por zona
    const sinAsigOrdenado = [...sinAsig].sort((a,b)=>detectarZona(a.lugar)-detectarZona(b.lugar))

    sinAsigOrdenado.forEach(x=>{
      const zonaX = detectarZona(x.lugar)
      const modo:ModoMov = desmontes.find(d=>d.id===x.id)&&!entregas.find(e=>e.id===x.id) ? "desmonte" : "entrega"

      let mejorUnidad = uActivas[0]
      let mejorScore = Infinity

      uActivas.forEach(u=>{
        const csU = porUnidad(u)
        const mismaZona = csU.some(cx=>detectarZona(cx.lugar)===zonaX)
        const carga = cargaU[u]||0
        // Score: menos carga = mejor; misma zona = bonus
        const score = carga * 10 - (mismaZona ? 8 : 0)
        if(score < mejorScore){ mejorScore=score; mejorUnidad=u }
      })

      nuevas[x.id] = {unidad:mejorUnidad, modo}
      cargaU[mejorUnidad] = (cargaU[mejorUnidad]||0) + 1
    })

    setAsig(nuevas)

    // Sugerir orden por zona para cada unidad
    const nuevoOrden:Record<string,string[]> = {...orden}
    uActivas.forEach(u=>{
      const csU = Object.entries(nuevas).filter(([,v])=>v.unidad===u).map(([id])=>contratos.find(x=>x.id===id)).filter(Boolean) as Contrato[]
      if(csU.length>1){
        const ent=csU.filter(x=>nuevas[x.id]?.modo==="entrega").sort((a,b)=>detectarZona(a.lugar)-detectarZona(b.lugar))
        const des=csU.filter(x=>nuevas[x.id]?.modo==="desmonte").sort((a,b)=>detectarZona(a.lugar)-detectarZona(b.lugar))
        const resultado:Contrato[]=[]
        const desUsados=new Set<string>()
        ent.forEach(e=>{
          resultado.push(e)
          const zE=detectarZona(e.lugar)
          const dZ=des.find(d=>!desUsados.has(d.id)&&detectarZona(d.lugar)===zE)
          if(dZ){resultado.push(dZ);desUsados.add(dZ.id)}
        })
        des.forEach(d=>{if(!desUsados.has(d.id))resultado.push(d)})
        nuevoOrden[u]=resultado.map(x=>x.id)
      }
    })
    setOrden(nuevoOrden)
  }

  const sugerirRuta = (u:string) => {
    const cs = porUnidad(u)
    if(cs.length<=1) return // nada que ordenar
    // Separar entregas y desmontes
    const ent = cs.filter(x=>asig[x.id]?.modo==="entrega")
    const des = cs.filter(x=>asig[x.id]?.modo==="desmonte")
    // Ordenar cada grupo por zona
    const sortZona = (arr:Contrato[]) => [...arr].sort((a,b)=>detectarZona(a.lugar)-detectarZona(b.lugar))
    // Estrategia: intercalar entrega+desmonte de la misma zona
    const entOrdenado = sortZona(ent)
    const desOrdenado = sortZona(des)
    // Mezclar: por cada zona, primero entrega luego desmonte cercano
    const resultado:Contrato[] = []
    const desUsados = new Set<string>()
    entOrdenado.forEach(e=>{
      resultado.push(e)
      // buscar desmonte de la misma zona
      const zonaE = detectarZona(e.lugar)
      const desZona = desOrdenado.find(d=>!desUsados.has(d.id)&&detectarZona(d.lugar)===zonaE)
      if(desZona){ resultado.push(desZona); desUsados.add(desZona.id) }
    })
    // Agregar desmontes que no se colocaron junto a una entrega
    desOrdenado.forEach(d=>{ if(!desUsados.has(d.id)) resultado.push(d) })
    // Aplicar orden
    const nuevoOrden = resultado.map(x=>x.id)
    setOrden(prev=>({...prev,[u]:nuevoOrden}))
  }

  return(
    <div>
      {/* ── HEADER ── */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 18px",marginBottom:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap" as const}}>
        <div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>🚚 Hoja de Carga</div>
          <div style={{fontSize:11,color:"#9a9590"}}>Entrega y desmonte en una sola ruta</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,background:"#f5f4f0",borderRadius:8,padding:"6px 12px"}}>
          <span style={{fontSize:12,fontWeight:700}}>📅</span>
          <input type="date" value={fecha} onChange={e=>setFecha(e.target.value)}
            style={{border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:700,color:"#1a1814",outline:"none",cursor:"pointer"}}/>
        </div>
        <div style={{display:"flex",gap:6}}>
          {(["planear","unidad"] as const).map(v=>(
            <button key={v} onClick={()=>setVista(v)} style={{padding:"6px 14px",borderRadius:8,border:`1.5px solid ${vista===v?"#1a1814":"#e8e5de"}`,background:vista===v?"#1a1814":"#fff",color:vista===v?"#fff":"#4a4640",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>
              {v==="planear"?"📋 Planear":"🚛 Por unidad"}
            </button>
          ))}
        </div>
        <div style={{marginLeft:"auto",display:"flex",gap:6,flexWrap:"wrap" as const}}>
          {/* Selector unidades disponibles */}
          {sinAsig.length>0&&(
            <div style={{display:"flex",alignItems:"center",gap:5,background:"#f5f0fc",border:"1.5px solid #4a2d6e",borderRadius:10,padding:"5px 10px",flexWrap:"wrap" as const}}>
              <span style={{fontSize:10,fontWeight:700,color:"#4a2d6e",marginRight:2}}>🚛</span>
              {UNIDADES_CARGA.map(u=>{
                const sel=unidadesDisp.includes(u)
                return(
                  <button key={u} onClick={()=>setUnidadesDisp(prev=>sel?prev.filter(x=>x!==u):[...prev,u])}
                    style={{padding:"3px 9px",borderRadius:6,border:`1.5px solid ${sel?"#4a2d6e":"#d4cfc4"}`,background:sel?"#4a2d6e":"#fff",color:sel?"#fff":"#9a9590",fontSize:9,fontWeight:sel?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif",whiteSpace:"nowrap" as const}}>
                    {sel?"✓ ":""}{u}
                  </button>
                )
              })}
              <span style={{fontSize:9,color:"#4a2d6e",opacity:.7,marginLeft:2}}>{unidadesDisp.length}/{UNIDADES_CARGA.length}</span>
            </div>
          )}
          {sinAsig.length>0&&(
            <button onClick={()=>{
              // Generate preview of what auto-assignment would look like
              const uPrev=UNIDADES_CARGA.filter(u=>unidadesDisp.includes(u))
              const cargaU:Record<string,number>={}
              uPrev.forEach(u=>{cargaU[u]=porUnidad(u).length})
              const preview:Record<string,{cliente:string,modo:string,lugar:string,folio:string}[]>={}
              uPrev.forEach(u=>{preview[u]=[]})
              // Show already assigned
              uPrev.forEach(u=>{
                porUnidad(u).forEach(x=>{
                  preview[u].push({cliente:x.cliente||x.archivo,modo:asig[x.id]?.modo||"entrega",lugar:(x.lugar||"").slice(0,40),folio:x.folio||""})
                })
              })
              // Simulate auto-assign for unassigned
              const sinOrdenado=[...sinAsig].sort((a,b)=>detectarZona(a.lugar)-detectarZona(b.lugar))
              sinOrdenado.forEach(x=>{
                const zonaX=detectarZona(x.lugar)
                const modo:ModoMov=desmontes.find(d=>d.id===x.id)&&!entregas.find(e=>e.id===x.id)?"desmonte":"entrega"
                let mejor=uPrev[0],mejorScore=Infinity
                uPrev.forEach(u=>{
                  const mismaZona=preview[u].some(p=>detectarZona(p.lugar)===zonaX)
                  const score=(cargaU[u]||0)*10-(mismaZona?8:0)
                  if(score<mejorScore){mejorScore=score;mejor=u}
                })
                preview[mejor].push({cliente:x.cliente||x.archivo,modo,lugar:(x.lugar||"").slice(0,40),folio:x.folio||""})
                cargaU[mejor]=(cargaU[mejor]||0)+1
              })
              // Build preview message
              const fechaStr=new Date(fecha+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long"})
              let txt=`👁️ *VISTA PREVIA SUGERIDA*
📅 ${fechaStr}
_(aplica "Asignación sugerida" para confirmar)
`
              uPrev.filter(u=>preview[u].length>0).forEach(u=>{
                txt+=`
*🚛 ${u}* (${preview[u].length} paradas)
`
                preview[u].forEach((p,i)=>{
                  const m=p.modo==="entrega"?"🚚":"📦"
                  const f=p.folio&&p.folio.length>1?` [${p.folio}]`:""
                  txt+=`  ${i+1}. ${m} *${p.cliente}*${f}
`
                  if(p.lugar) txt+=`      📍 ${p.lugar}
`
                })
              })
              txt+=`
— *Poliflor* 🌸`
              setMsgEdit(txt); setMsgModal({unidad:"Vista previa sugerida",txt}); setCopiado(false)
            }} style={{padding:"6px 12px",borderRadius:8,background:"#f5f0fc",color:"#4a2d6e",border:"1.5px solid #4a2d6e",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
              👁️ Ver sugerida
            </button>
          )}
          {/* Botones preview por unidad */}
          {UNIDADES_CARGA.filter(u=>porUnidad(u).length>0).map(u=>(
            <button key={u} onClick={()=>{const txt=generarMsg(u);setMsgEdit(txt);setMsgModal({unidad:u,txt});setCopiado(false)}}
              style={{padding:"6px 10px",borderRadius:8,background:"#25D366",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
              📋 {u}
            </button>
          ))}
          <button onClick={()=>{
            const txt=generarResumenGeneral()
            setMsgEdit(txt)
            setMsgModal({unidad:"General",txt})
            setCopiado(false)
          }} style={{padding:"6px 12px",borderRadius:8,background:"#4a2d6e",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
            📊 Resumen general
          </button>
          <button onClick={imprimirResumen}
            style={{padding:"6px 12px",borderRadius:8,background:"#fff",color:"#1a1814",border:"1.5px solid #1a1814",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
            🖨️ Imprimir
          </button>
          <button onClick={cancelarCambios}
            style={{padding:"6px 12px",borderRadius:8,background:"#fdf0f0",color:"#8b2e2e",border:"1.5px solid #e8b8b8",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>
            ✕ Cancelar todo
          </button>
          <button onClick={save} style={{padding:"6px 14px",borderRadius:8,background:guardado?"#2d6a4f":"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>
            {guardado?"✓ Guardado":"💾 Guardar"}
          </button>
        </div>
      </div>

      {/* ── KPIs ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:14}}>
        {[
          {l:"Entregas hoy",v:entregas.length,c:"#1a3a5c"},
          {l:"Desmontes hoy",v:desmontes.length,c:"#4a2d6e"},
          {l:"Sin asignar",v:sinAsig.length,c:sinAsig.length>0?"#8b2e2e":"#2d6a4f"},
          {l:"Unidades activas",v:UNIDADES_CARGA.filter(u=>porUnidad(u).length>0).length,c:"#92580a"},
          {l:"Piezas totales",v:todos.reduce((s:number,x:Contrato)=>s+(x.articulos||[]).reduce((a:number,ar:Articulo)=>a+ar.cantidad,0),0),c:"#1a1814"},
        ].map((k,i)=>(
          <div key={i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px",textAlign:"center" as const}}>
            <div style={{fontSize:9,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:4}}>{k.l}</div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:24,fontWeight:800,color:k.c}}>{k.v.toLocaleString()}</div>
          </div>
        ))}
      </div>

      {/* ══ PLANEAR ══ */}
      {vista==="planear"&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 300px",gap:14}}>
          <div>
            {/* Sin asignar */}
            {sinAsig.length>0&&(
              <div style={{background:"#fff",border:"1.5px solid #e8d4b8",borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap" as const,gap:6}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,color:"#92580a"}}>⚠️ Sin asignar ({sinAsig.length})</div>
                  {sinAsig.length>0&&(
                    <div style={{display:"flex",gap:5,flexWrap:"wrap" as const}}>
                      {/* Botón asignación sugerida automática */}
                      <button onClick={sugerirAsignacion}
                        style={{fontSize:10,padding:"4px 12px",borderRadius:6,border:"1.5px solid #1a1814",background:"#1a1814",color:"#fff",cursor:"pointer",fontWeight:700,fontFamily:"Epilogue,sans-serif",display:"flex",alignItems:"center",gap:4}}>
                        🤖 Sugerir ({unidadesDisp.length} unidad{unidadesDisp.length!==1?"es":""})
                      </button>
                      {UNIDADES_CARGA.map((u,i)=>(
                        <button key={u} onClick={()=>{
                          const nuevas={...asig}
                          sinAsig.forEach(x=>{
                            const modo:ModoMov = desmontes.find(d=>d.id===x.id) ? "desmonte" : "entrega"
                            nuevas[x.id]={unidad:u,modo}
                          })
                          setAsig(nuevas)
                        }} style={{fontSize:9,padding:"3px 8px",borderRadius:6,border:`1.5px solid ${COL_U[i]}`,background:COL_U[i]+"15",color:COL_U[i],cursor:"pointer",fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                          Todo → {u}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div style={{display:"flex",flexDirection:"column" as const,gap:5}}>
                  {sinAsig.map((x:Contrato)=>{
                    const esDesmonte = !!desmontes.find(d=>d.id===x.id)
                    const esEntrega = !!entregas.find(e=>e.id===x.id)
                    return(
                      <div key={x.id} style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:8,padding:"8px 10px",display:"flex",alignItems:"center",gap:8}}>
                        <div style={{display:"flex",gap:3,flexShrink:0}}>
                          {esEntrega&&<span style={{fontSize:9,padding:"2px 5px",borderRadius:5,background:"#edf3fa",color:"#1a3a5c",fontWeight:700}}>🚚ENT</span>}
                          {esDesmonte&&<span style={{fontSize:9,padding:"2px 5px",borderRadius:5,background:"#f5f0fc",color:"#4a2d6e",fontWeight:700}}>📦DES</span>}
                        </div>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente}</div>
                          <div style={{fontSize:10,color:"#9a9590"}}>{(x.lugar||"Sin dir.").slice(0,45)}</div>
                        </div>
                        <select onChange={e=>{
                          if(!e.target.value) return
                          const modo:ModoMov = esDesmonte&&!esEntrega ? "desmonte" : "entrega"
                          asignar(x.id, e.target.value, modo)
                        }} defaultValue=""
                          style={{padding:"4px 6px",borderRadius:6,border:"1.5px solid #1a1814",fontFamily:"Epilogue,sans-serif",fontSize:10,fontWeight:700,cursor:"pointer",outline:"none"}}>
                          <option value="">Asignar →</option>
                          {UNIDADES_CARGA.map(u=><option key={u} value={u}>{u}</option>)}
                        </select>
                      </div>
                    )
                  })}
                </div>
              </div>
            )}

            {/* Por unidad */}
            {UNIDADES_CARGA.map((u,uidx)=>{
              const cs = porUnidad(u)
              if(!cs.length) return null
              const csEnt=cs.filter(x=>asig[x.id]?.modo==="entrega")
              const csDes=cs.filter(x=>asig[x.id]?.modo==="desmonte")
              return(
                <div key={u} style={{background:"#fff",border:`1.5px solid ${COL_U[uidx]}35`,borderRadius:12,marginBottom:10,overflow:"hidden"}}>
                  <div style={{background:COL_U[uidx]+"15",borderBottom:`1px solid ${COL_U[uidx]}30`,padding:"8px 12px",display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:COL_U[uidx]}}/>
                    <span style={{fontWeight:800,fontSize:13,color:COL_U[uidx],flex:1}}>{u}</span>
                    <span style={{fontSize:10,color:"#9a9590"}}>🚚{csEnt.length} entregas · 📦{csDes.length} desmontes</span>
                    {cs.length>1&&(
                      <button onClick={e=>{e.stopPropagation();sugerirRuta(u)}}
                        title="Ordenar paradas por zona geográfica"
                        style={{fontSize:10,padding:"2px 8px",borderRadius:6,border:`1px solid ${COL_U[uidx]}`,background:"#fff",color:COL_U[uidx],cursor:"pointer",fontWeight:700,fontFamily:"Epilogue,sans-serif",display:"flex",alignItems:"center",gap:4}}>
                        🗺️ Ruta sugerida
                      </button>
                    )}
                  </div>
                  <div style={{padding:"6px 10px",display:"flex",flexDirection:"column" as const,gap:4}}>
                    {cs.map((x:Contrato)=>{
                      const p=prog(x.id)
                      const isExp=expandido===x.id
                      const modoX=asig[x.id]?.modo||"entrega"
                      const modoCol=modoX==="entrega"?"#1a3a5c":"#4a2d6e"
                      const modoBg=modoX==="entrega"?"#edf3fa":"#f5f0fc"
                      const isPend=pendFwd.includes(x.id)
                      return(
                        <div key={x.id} style={{border:`1px solid ${isPend?"#e8d4b8":"#e8e5de"}`,borderRadius:8,overflow:"hidden",background:isPend?"#fffbf5":"#fff"}}>
                          <div style={{display:"flex",alignItems:"center",gap:6,padding:"7px 8px",cursor:"pointer"}} onClick={()=>setExpandido(isExp?null:x.id)}>
                            <span style={{fontSize:9,padding:"2px 5px",borderRadius:5,background:modoBg,color:modoCol,fontWeight:700,flexShrink:0}}>
                              {modoX==="entrega"?"🚚":"📦"}
                            </span>
                            <div style={{flex:1,minWidth:0}}>
                              <div style={{fontWeight:700,fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente}</div>
                              <div style={{fontSize:9,color:"#9a9590",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.lugar?.slice(0,40)||"—"}</div>
                              {(()=>{const z=detectarZona(x.lugar);const zonaLabel=z<ZONAS_ORDEN.length?ZONAS_ORDEN[z][0]:null;return zonaLabel?<span style={{fontSize:8,color:"#9a9590",fontStyle:"italic"}}>{zonaLabel}</span>:null})()}
                            </div>
                            {p.total>0&&(
                              <div style={{display:"flex",alignItems:"center",gap:3,flexShrink:0}}>
                                <div style={{width:36,height:3,background:"#e8e5de",borderRadius:2,overflow:"hidden"}}>
                                  <div style={{height:"100%",background:p.done===p.total?"#2d6a4f":COL_U[uidx],width:p.total>0?Math.round(p.done/p.total*100)+"%":"0%"}}/>
                                </div>
                                <span style={{fontSize:8,color:"#9a9590"}}>{p.done}/{p.total}</span>
                              </div>
                            )}
                            <div style={{display:"flex",gap:2,flexShrink:0}}>
                              <button onClick={e=>{e.stopPropagation();mover(u,x.id,-1)}} style={{fontSize:9,background:"none",border:"1px solid #e8e5de",borderRadius:3,color:"#9a9590",cursor:"pointer",padding:"1px 4px"}}>▲</button>
                              <button onClick={e=>{e.stopPropagation();mover(u,x.id,1)}} style={{fontSize:9,background:"none",border:"1px solid #e8e5de",borderRadius:3,color:"#9a9590",cursor:"pointer",padding:"1px 4px"}}>▼</button>
                              <button onClick={e=>{e.stopPropagation();setReasigModal(x.id)}} title="Reasignar unidad"
                                style={{fontSize:9,background:"#fdf5e8",border:"1px solid #e8d4b8",borderRadius:3,color:"#92580a",cursor:"pointer",padding:"1px 5px",fontWeight:700}}>↔</button>
                              <button onClick={e=>{e.stopPropagation();setPendFwd(prev=>prev.includes(x.id)?prev.filter(i=>i!==x.id):[...prev,x.id])}} title="Pasar al siguiente día"
                                style={{fontSize:9,background:isPend?"#fdf5e8":"none",border:`1px solid ${isPend?"#e8d4b8":"#e8e5de"}`,borderRadius:3,color:isPend?"#92580a":"#9a9590",cursor:"pointer",padding:"1px 5px",fontWeight:isPend?700:400}}>
                                {isPend?"📌":"→D+1"}
                              </button>
                              <button onClick={e=>{e.stopPropagation();desasignar(x.id)}}
                                style={{fontSize:9,background:"none",border:"1px solid #e8e5de",borderRadius:3,color:"#9a9590",cursor:"pointer",padding:"1px 4px"}}>✕</button>
                            </div>
                            <span style={{color:"#9a9590",fontSize:10}}>{isExp?"▴":"▾"}</span>
                          </div>
                          {isExp&&(
                            <div style={{borderTop:"1px solid #e8e5de",padding:"8px 10px",background:"#fafaf8"}}>
                              <div style={{fontSize:9,color:"#9a9590",marginBottom:5}}>📞 {x.tel||"—"}</div>
                              <div style={{display:"flex",flexDirection:"column" as const,gap:3,marginBottom:6}}>
                                {(x.articulos||[]).map((a:Articulo,i:number)=>{
                                  const done=check[x.id]?.[String(i)]||false
                                  return(
                                    <label key={i} style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",padding:"3px 5px",borderRadius:4,background:done?"#edf7f2":"transparent"}}>
                                      <input type="checkbox" checked={done} onChange={e=>setCheck(prev=>({...prev,[x.id]:{...(prev[x.id]||{}),[String(i)]:e.target.checked}}))}
                                        style={{width:13,height:13,cursor:"pointer",accentColor:modoX==="entrega"?"#1a3a5c":"#4a2d6e"}}/>
                                      <span style={{fontSize:11,fontWeight:600,color:"#9a9590",minWidth:20,textAlign:"right" as const}}>{a.cantidad}x</span>
                                      <span style={{flex:1,fontSize:11,color:done?"#9a9590":"#1a1814",textDecoration:done?"line-through":"none"}}>{a.nombre}</span>
                                    </label>
                                  )
                                })}
                              </div>
                              <textarea value={notas[x.id]||""} onChange={e=>setNotas({...notas,[x.id]:e.target.value})}
                                placeholder="Nota para operador..." rows={2}
                                style={{width:"100%",padding:"5px 7px",border:"1px solid #e8e5de",borderRadius:5,fontFamily:"Epilogue,sans-serif",fontSize:10,resize:"none" as const,outline:"none",boxSizing:"border-box" as const}}/>
                            </div>
                          )}
                        </div>
                      )
                    })}
                  </div>
                </div>
              )
            })}

            {/* Pasar pendientes */}
            {pendFwd.length>0&&(
              <div style={{background:"#fdf5e8",border:"1.5px solid #e8d4b8",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:10}}>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:13,color:"#92580a"}}>📌 {pendFwd.length} contrato(s) marcados para mañana</div>
                  <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>{pendFwd.map(id=>contratos.find(x=>x.id===id)?.cliente||id).join(", ")}</div>
                </div>
                <button onClick={pasarAlSiguiente} style={{padding:"8px 16px",borderRadius:8,background:"#92580a",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                  Pasar al D+1 →
                </button>
              </div>
            )}
          </div>

          {/* Panel derecho — resumen por unidad */}
          <div style={{position:"sticky" as const,top:80,alignSelf:"flex-start" as const,display:"flex",flexDirection:"column" as const,gap:8}}>
            {UNIDADES_CARGA.map((u,i)=>{
              const cs=porUnidad(u)
              const cons=consolidado(u)
              if(!cs.length) return(
                <div key={u} style={{background:"#fafaf8",border:"1px dashed #e8e5de",borderRadius:8,padding:"8px 12px",opacity:.5,display:"flex",alignItems:"center",gap:6}}>
                  <div style={{width:7,height:7,borderRadius:"50%",background:"#d4cfc4"}}/>
                  <span style={{fontSize:11,color:"#9a9590"}}>{u} — vacía</span>
                </div>
              )
              return(
                <div key={u} style={{background:"#fff",border:`1.5px solid ${COL_U[i]}40`,borderRadius:10,overflow:"hidden"}}>
                  <div style={{background:COL_U[i],padding:"7px 10px",display:"flex",gap:6,alignItems:"center"}}>
                    <span style={{fontWeight:800,fontSize:11,color:"#fff",flex:1}}>{u}</span>
                    <span style={{fontSize:9,color:"rgba(255,255,255,.8)"}}>{cs.length} paradas · {cons.reduce((s,a)=>s+a.cantidad,0)} pzas</span>
                  </div>
                  <div style={{padding:"6px 8px",maxHeight:140,overflowY:"auto" as const}}>
                    {cons.slice(0,7).map((a,j)=>(
                      <div key={j} style={{display:"flex",gap:5,fontSize:9,padding:"2px 0",borderBottom:"1px solid #f5f4f0"}}>
                        <span style={{fontWeight:700,color:COL_U[i],minWidth:20,textAlign:"right" as const}}>{a.cantidad}x</span>
                        <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</span>
                      </div>
                    ))}
                    {cons.length>7&&<div style={{fontSize:8,color:"#9a9590",textAlign:"center" as const,paddingTop:2}}>+{cons.length-7} más</div>}
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      )}

      {/* ══ POR UNIDAD ══ */}
      {vista==="unidad"&&(
        <div>
          <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap" as const}}>
            {UNIDADES_CARGA.map((u,i)=>{
              const cs=porUnidad(u)
              return(
                <button key={u} onClick={()=>setUnidadActiva(u)}
                  style={{padding:"8px 14px",borderRadius:10,border:`2px solid ${unidadActiva===u?COL_U[i]:"#e8e5de"}`,background:unidadActiva===u?COL_U[i]:"#fff",color:unidadActiva===u?"#fff":"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:700,fontSize:11,display:"flex",flexDirection:"column" as const,alignItems:"center",gap:1,minWidth:90}}>
                  <span>{u}</span>
                  <span style={{fontSize:8,fontWeight:400,opacity:.8}}>{cs.length} paradas</span>
                </button>
              )
            })}
          </div>
          {(()=>{
            const cs=porUnidad(unidadActiva)
            const cons=consolidado(unidadActiva)
            const ci=UNIDADES_CARGA.indexOf(unidadActiva)
            if(!cs.length) return(
              <div style={{textAlign:"center" as const,padding:48,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12}}>
                <div style={{fontSize:36,opacity:.15}}>🚚</div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:14,color:"#9a9590"}}>Sin contratos asignados a {unidadActiva}</div>
              </div>
            )
            return(
              <div style={{display:"grid",gridTemplateColumns:"1fr 280px",gap:14}}>
                <div>
                  {cs.map((x:Contrato,idx:number)=>{
                    const p=prog(x.id); const isExp=expandido===x.id
                    const modoX=asig[x.id]?.modo||"entrega"
                    const pct=p.total>0?Math.round(p.done/p.total*100):0
                    return(
                      <div key={x.id} style={{background:"#fff",border:`1.5px solid ${pct===100?"#2d6a4f40":COL_U[ci]+"30"}`,borderRadius:10,marginBottom:10,overflow:"hidden"}}>
                        <div style={{background:pct===100?"#edf7f2":COL_U[ci]+"10",padding:"10px 12px",cursor:"pointer",display:"flex",gap:8,alignItems:"flex-start"}} onClick={()=>setExpandido(isExp?null:x.id)}>
                          <div style={{width:26,height:26,borderRadius:"50%",background:pct===100?"#2d6a4f":COL_U[ci],color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,flexShrink:0}}>{idx+1}</div>
                          <div style={{flex:1}}>
                            <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                              <span style={{fontSize:9,padding:"1px 5px",borderRadius:4,background:modoX==="entrega"?"#edf3fa":"#f5f0fc",color:modoX==="entrega"?"#1a3a5c":"#4a2d6e",fontWeight:700}}>
                                {modoX==="entrega"?"🚚ENT":"📦DES"}
                              </span>
                              <span style={{fontFamily:"Playfair Display,serif",fontSize:13,fontWeight:700}}>{x.cliente}</span>
                            </div>
                            <div style={{fontSize:10,color:"#9a9590"}}>📍 {x.lugar?.slice(0,50)||"—"}</div>
                            <div style={{fontSize:10,color:"#4a4640"}}>📞 {x.tel||"—"}</div>
                          </div>
                          <div style={{textAlign:"center" as const,flexShrink:0}}>
                            <div style={{fontSize:18,fontWeight:800,color:pct===100?"#2d6a4f":COL_U[ci]}}>{pct}%</div>
                            <div style={{fontSize:8,color:"#9a9590"}}>{p.done}/{p.total}</div>
                          </div>
                        </div>
                        {isExp&&(
                          <div style={{borderTop:`1px solid ${COL_U[ci]}20`,padding:"10px 12px"}}>
                            <div style={{display:"flex",flexDirection:"column" as const,gap:4}}>
                              {(x.articulos||[]).map((a:Articulo,i:number)=>{
                                const done=check[x.id]?.[String(i)]||false
                                return(
                                  <label key={i} style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",padding:"5px 7px",borderRadius:6,background:done?"#edf7f2":"#fafaf8",border:`1px solid ${done?"#b7deca":"#e8e5de"}`}}>
                                    <input type="checkbox" checked={done} onChange={e=>setCheck(prev=>({...prev,[x.id]:{...(prev[x.id]||{}),[String(i)]:e.target.checked}}))}
                                      style={{width:15,height:15,cursor:"pointer",accentColor:modoX==="entrega"?"#1a3a5c":"#4a2d6e"}}/>
                                    <span style={{fontWeight:700,color:"#9a9590",minWidth:22,textAlign:"right" as const,fontSize:11}}>{a.cantidad}x</span>
                                    <span style={{flex:1,fontSize:11,color:done?"#9a9590":"#1a1814",textDecoration:done?"line-through":"none"}}>{a.nombre}</span>
                                    {done&&<span style={{color:"#2d6a4f",fontSize:13}}>✓</span>}
                                  </label>
                                )
                              })}
                            </div>
                          </div>
                        )}
                      </div>
                    )
                  })}
                </div>
                <div style={{position:"sticky" as const,top:80,alignSelf:"flex-start" as const}}>
                  <div style={{background:"#fff",border:`1.5px solid ${COL_U[ci]}40`,borderRadius:12,overflow:"hidden"}}>
                    <div style={{background:COL_U[ci],padding:"10px 14px"}}>
                      <div style={{fontWeight:800,fontSize:13,color:"#fff"}}>{unidadActiva}</div>
                      <div style={{fontSize:10,color:"rgba(255,255,255,.7)",marginTop:1}}>{cons.length} artículos · {cons.reduce((s,a)=>s+a.cantidad,0)} piezas</div>
                    </div>
                    <div style={{padding:"8px 12px",maxHeight:450,overflowY:"auto" as const}}>
                      {cons.map((a,i)=>(
                        <div key={i} style={{display:"flex",alignItems:"center",gap:7,padding:"5px 0",borderBottom:"1px solid #f5f4f0"}}>
                          <span style={{fontWeight:800,color:COL_U[ci],fontSize:12,minWidth:28,textAlign:"right" as const}}>{a.cantidad}x</span>
                          <div style={{flex:1}}>
                            <div style={{fontSize:11,fontWeight:600}}>{a.nombre}</div>
                            <div style={{fontSize:8,color:"#9a9590"}}>{a.clientes.join(", ").slice(0,35)}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )
          })()}
        </div>
      )}

      {/* ── MODAL REASIGNAR ── */}
      {reasigModal&&(()=>{
        const x=contratos.find(c=>c.id===reasigModal)
        if(!x) return null
        return(
          <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}}>
            <div style={{background:"#fff",borderRadius:14,padding:24,width:340,boxShadow:"0 20px 60px rgba(0,0,0,.2)"}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,marginBottom:4}}>↔ Reasignar unidad</div>
              <div style={{fontSize:12,color:"#9a9590",marginBottom:16}}>{x.cliente}</div>
              <div style={{display:"flex",flexDirection:"column" as const,gap:6,marginBottom:16}}>
                {UNIDADES_CARGA.map((u,i)=>(
                  <button key={u} onClick={()=>reasignar(reasigModal,u)}
                    style={{padding:"10px 14px",borderRadius:8,border:`1.5px solid ${asig[reasigModal]?.unidad===u?COL_U[i]:"#e8e5de"}`,background:asig[reasigModal]?.unidad===u?COL_U[i]+"15":"#fff",color:asig[reasigModal]?.unidad===u?COL_U[i]:"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:700,fontSize:12,textAlign:"left" as const,display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:COL_U[i]}}/>
                    {u}
                    {asig[reasigModal]?.unidad===u&&<span style={{marginLeft:"auto",fontSize:10}}>actual</span>}
                  </button>
                ))}
              </div>
              <button onClick={()=>setReasigModal(null)} style={{width:"100%",padding:"8px",borderRadius:8,background:"#f5f4f0",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12}}>Cancelar</button>
            </div>
          </div>
        )
      })()}

      {/* ── MODAL PREVIEW MENSAJE ── */}
      {msgModal&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000}}>
          <div style={{background:"#fff",borderRadius:16,width:560,maxHeight:"90vh",display:"flex",flexDirection:"column" as const,boxShadow:"0 24px 80px rgba(0,0,0,.25)"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #e8e5de",display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1}}>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800}}>📋 Ruta {msgModal.unidad}</div>
                <div style={{fontSize:11,color:"#9a9590"}}>Edita antes de enviar</div>
              </div>
              <button onClick={()=>setMsgModal(null)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#9a9590"}}>✕</button>
            </div>
            <div style={{flex:1,padding:"12px 16px",overflowY:"auto" as const}}>
              <textarea value={msgEdit} onChange={e=>setMsgEdit(e.target.value)}
                style={{width:"100%",height:380,padding:"10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"monospace",fontSize:12,lineHeight:1.65,resize:"vertical" as const,outline:"none",boxSizing:"border-box" as const,color:"#1a1814",background:"#fafaf8"}}/>
            </div>
            <div style={{padding:"12px 16px",borderTop:"1px solid #e8e5de",display:"flex",gap:8}}>
              <button onClick={()=>{navigator.clipboard.writeText(msgEdit).then(()=>{setCopiado(true);setTimeout(()=>setCopiado(false),2500)})}}
                style={{flex:1,padding:"9px",borderRadius:8,background:copiado?"#2d6a4f":"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700,transition:"background .3s"}}>
                {copiado?"✓ ¡Copiado!":"📋 Copiar texto"}
              </button>
              <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(msgEdit)}`,"_blank")}
                style={{flex:1,padding:"9px",borderRadius:8,background:"#25D366",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                WhatsApp →
              </button>
              <button onClick={()=>setMsgEdit(msgModal.txt)} title="Restaurar"
                style={{padding:"9px 12px",borderRadius:8,background:"#f5f4f0",border:"1px solid #e8e5de",cursor:"pointer",fontSize:13,color:"#9a9590"}}>↺</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── CATÁLOGO DE ARTÍCULOS ─────────────────────────────────────────
const CATS_CATALOGO = ["MOBILIARIO","FLORES","MANTELERIA","VAJILLA","SERVICIOS","CARPAS","ILUMINACION"]
const CAT_COLORS: Record<string,string> = {
  MOBILIARIO:"#1a3a5c",FLORES:"#2d6a4f",MANTELERIA:"#92580a",
  VAJILLA:"#4a2d6e",SERVICIOS:"#8b2e2e",CARPAS:"#1a5c4a",ILUMINACION:"#5c4a1a"
}
const CAT_ICONS: Record<string,string> = {
  MOBILIARIO:"🪑",FLORES:"🌸",MANTELERIA:"🪢",
  VAJILLA:"🍽️",SERVICIOS:"🚚",CARPAS:"⛺",ILUMINACION:"💡"
}

interface ArticuloCatalogo {
  id: string; codigo: string; nombre: string; categoria: string; subcategoria: string
  descripcion: string; precio_renta: number; precio_venta: number; unidad: string
  existencia_total: number; existencia_disponible: number; color: string
  material: string; medidas: string; notas: string; activo: boolean
}

function CatalogoSection({ token }: { token: string }) {
  const [articulos, setArticulos] = useState<ArticuloCatalogo[]>([])
  const [cargando, setCargando] = useState(true)
  const [catActiva, setCatActiva] = useState("todos")
  const [busq, setBusq] = useState("")
  const [soloActivos, setSoloActivos] = useState(true)
  const [modal, setModal] = useState<ArticuloCatalogo | null>(null)
  const [esNuevo, setEsNuevo] = useState(false)
  const [guardando, setGuardando] = useState(false)
  const [paginaActual, setPaginaActual] = useState(1)
  const POR_PAGINA = 50

  const cargarCatalogo = async (cat = catActiva, q = busq) => {
    setCargando(true)
    const params = new URLSearchParams()
    if (cat !== "todos") params.set("categoria", cat)
    if (q.trim()) params.set("busq", q.trim())
    params.set("activo", String(soloActivos))
    const res = await fetch(`/api/catalogo?${params}`, {
      headers: { Authorization: `Bearer ${token}` }
    })
    const data = await res.json()
    setArticulos(Array.isArray(data) ? data : [])
    setCargando(false)
    setPaginaActual(1)
  }

  useState(() => { cargarCatalogo() })

  const guardar = async () => {
    if (!modal) return
    setGuardando(true)
    const method = esNuevo ? "POST" : "PATCH"
    const url = esNuevo ? "/api/catalogo" : `/api/catalogo?id=${modal.id}`
    await fetch(url, {
      method, headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
      body: JSON.stringify(modal)
    })
    setGuardando(false)
    setModal(null)
    cargarCatalogo()
  }

  const toggleActivo = async (art: ArticuloCatalogo) => {
    await fetch(`/api/catalogo?id=${art.id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
      body: JSON.stringify({ activo: !art.activo })
    })
    setArticulos(prev => prev.map(a => a.id === art.id ? { ...a, activo: !a.activo } : a))
  }

  // Filtrar localmente
  const filtrados = articulos.filter(a => {
    if (soloActivos && !a.activo) return false
    if (catActiva !== "todos" && a.categoria !== catActiva) return false
    if (busq.trim()) return a.nombre.toLowerCase().includes(busq.toLowerCase()) ||
      a.codigo.toLowerCase().includes(busq.toLowerCase()) ||
      a.subcategoria.toLowerCase().includes(busq.toLowerCase())
    return true
  })

  // Agrupar por subcategoría
  const porSubcat: Record<string, ArticuloCatalogo[]> = {}
  filtrados.forEach(a => {
    const k = a.subcategoria || "Sin categoría"
    if (!porSubcat[k]) porSubcat[k] = []
    porSubcat[k].push(a)
  })

  // Paginación simple
  const inicio = (paginaActual - 1) * POR_PAGINA
  const pagina = filtrados.slice(inicio, inicio + POR_PAGINA)
  const totalPags = Math.ceil(filtrados.length / POR_PAGINA)

  // Conteos por categoría
  const conteos: Record<string, number> = { todos: articulos.length }
  articulos.forEach(a => { conteos[a.categoria] = (conteos[a.categoria] || 0) + 1 })

  const nuevaPlantilla: ArticuloCatalogo = {
    id: "", codigo: "", nombre: "", categoria: "MOBILIARIO", subcategoria: "",
    descripcion: "", precio_renta: 0, precio_venta: 0, unidad: "pza",
    existencia_total: 0, existencia_disponible: 0, color: "", material: "",
    medidas: "", notas: "", activo: true
  }

  return (
    <div>
      {/* ── HEADER ── */}
      <div style={{ background: "#fff", border: "1px solid #e8e5de", borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12, flexWrap: "wrap" as const }}>
          <div>
            <div style={{ fontFamily: "Playfair Display,serif", fontSize: 18, fontWeight: 800 }}>📦 Catálogo de Artículos</div>
            <div style={{ fontSize: 11, color: "#9a9590" }}>{articulos.length.toLocaleString()} artículos en total</div>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <div style={{ position: "relative" as const }}>
              <input value={busq} onChange={e => { setBusq(e.target.value); cargarCatalogo(catActiva, e.target.value) }}
                placeholder="Buscar artículo..." style={{ padding: "7px 32px 7px 12px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", width: 220 }} />
              {busq && <button onClick={() => { setBusq(""); cargarCatalogo(catActiva, "") }} style={{ position: "absolute" as const, right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "#9a9590" }}>✕</button>}
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, cursor: "pointer" }}>
              <input type="checkbox" checked={soloActivos} onChange={e => { setSoloActivos(e.target.checked); cargarCatalogo() }} />
              Solo activos
            </label>
            <button onClick={() => { setEsNuevo(true); setModal({ ...nuevaPlantilla }) }}
              style={{ padding: "7px 16px", borderRadius: 8, background: "#1a1814", color: "#fff", border: "none", cursor: "pointer", fontFamily: "Epilogue,sans-serif", fontSize: 12, fontWeight: 700 }}>
              + Nuevo artículo
            </button>
          </div>
        </div>
        {/* Filtros por categoría */}
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" as const }}>
          {[["todos", "Todos", articulos.length], ...CATS_CATALOGO.map(c => [c, `${CAT_ICONS[c]} ${c}`, conteos[c] || 0])].map(([k, l, n]) => (
            <button key={String(k)} onClick={() => { setCatActiva(String(k)); cargarCatalogo(String(k)) }}
              style={{ padding: "4px 12px", borderRadius: 16, border: `1.5px solid ${catActiva === k ? (CAT_COLORS[String(k)] || "#1a1814") : "#e8e5de"}`, background: catActiva === k ? (CAT_COLORS[String(k)] || "#1a1814") : "#fff", color: catActiva === k ? "#fff" : "#4a4640", fontSize: 11, fontWeight: catActiva === k ? 700 : 400, cursor: "pointer", fontFamily: "Epilogue,sans-serif" }}>
              {String(l)} ({Number(n).toLocaleString()})
            </button>
          ))}
        </div>
      </div>

      {/* ── KPIs ── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 14 }}>
        {[
          { l: "Total artículos", v: articulos.length, c: "#1a1814" },
          { l: "Activos", v: articulos.filter(a => a.activo).length, c: "#2d6a4f" },
          { l: "Con precio", v: articulos.filter(a => a.precio_renta > 0).length, c: "#1a3a5c" },
          { l: "Sin precio", v: articulos.filter(a => a.precio_renta === 0).length, c: "#92580a" },
        ].map((k, i) => (
          <div key={i} style={{ background: "#fff", border: "1px solid #e8e5de", borderRadius: 10, padding: "12px 14px" }}>
            <div style={{ fontSize: 9, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, letterSpacing: ".06em", marginBottom: 4 }}>{k.l}</div>
            <div style={{ fontFamily: "Playfair Display,serif", fontSize: 22, fontWeight: 800, color: k.c }}>{k.v.toLocaleString()}</div>
          </div>
        ))}
      </div>

      {/* ── TABLA ── */}
      <div style={{ background: "#fff", border: "1px solid #e8e5de", borderRadius: 12, overflow: "hidden" }}>
        {cargando ? (
          <div style={{ padding: 48, textAlign: "center" as const, color: "#9a9590" }}>Cargando...</div>
        ) : filtrados.length === 0 ? (
          <div style={{ padding: 48, textAlign: "center" as const, color: "#9a9590" }}>
            <div style={{ fontSize: 32, opacity: .2, marginBottom: 8 }}>📦</div>
            <div>Sin resultados</div>
          </div>
        ) : (
          <>
            <table style={{ width: "100%", borderCollapse: "collapse" as const, fontSize: 12 }}>
              <thead>
                <tr style={{ background: "#fafaf8" }}>
                  {["Código", "Nombre", "Categoría", "Subcategoría", "Precio Renta", "Existencia", "Estado", ""].map((h, i) => (
                    <th key={i} style={{ padding: "9px 12px", textAlign: "left" as const, fontSize: 10, fontWeight: 700, textTransform: "uppercase" as const, color: "#9a9590", letterSpacing: ".05em", borderBottom: "1px solid #e8e5de", whiteSpace: "nowrap" as const }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pagina.map((a, i) => (
                  <tr key={a.id} style={{ borderBottom: "1px solid #e8e5de", background: i % 2 === 0 ? "#fff" : "#fafaf8", opacity: a.activo ? 1 : .5 }}>
                    <td style={{ padding: "8px 12px", fontFamily: "monospace", fontSize: 11, color: "#9a9590" }}>{a.codigo}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 600, maxWidth: 280 }}>
                      <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" as const }}>{a.nombre}</div>
                      {a.color && <div style={{ fontSize: 9, color: "#9a9590" }}>{a.color}{a.material ? ` · ${a.material}` : ""}{a.medidas ? ` · ${a.medidas}` : ""}</div>}
                    </td>
                    <td style={{ padding: "8px 12px" }}>
                      <span style={{ fontSize: 10, padding: "2px 8px", borderRadius: 8, background: (CAT_COLORS[a.categoria] || "#9a9590") + "18", color: CAT_COLORS[a.categoria] || "#9a9590", fontWeight: 700 }}>
                        {CAT_ICONS[a.categoria] || "📦"} {a.categoria}
                      </span>
                    </td>
                    <td style={{ padding: "8px 12px", fontSize: 11, color: "#4a4640" }}>{a.subcategoria || "—"}</td>
                    <td style={{ padding: "8px 12px", fontFamily: "monospace", fontWeight: 700, color: a.precio_renta > 0 ? "#2d6a4f" : "#9a9590" }}>
                      {a.precio_renta > 0 ? `$${a.precio_renta.toLocaleString()}` : "—"}
                    </td>
                    <td style={{ padding: "8px 12px", textAlign: "center" as const }}>
                      {a.existencia_total > 0 ? (
                        <span style={{ fontSize: 11, color: "#1a3a5c", fontWeight: 700 }}>{a.existencia_disponible}/{a.existencia_total}</span>
                      ) : <span style={{ fontSize: 11, color: "#9a9590" }}>—</span>}
                    </td>
                    <td style={{ padding: "8px 12px" }}>
                      <button onClick={() => toggleActivo(a)}
                        style={{ fontSize: 10, padding: "2px 8px", borderRadius: 8, border: `1px solid ${a.activo ? "#b7deca" : "#e8b8b8"}`, background: a.activo ? "#edf7f2" : "#fdf0f0", color: a.activo ? "#2d6a4f" : "#8b2e2e", cursor: "pointer", fontWeight: 700 }}>
                        {a.activo ? "✓ Activo" : "Inactivo"}
                      </button>
                    </td>
                    <td style={{ padding: "8px 12px" }}>
                      <button onClick={() => { setEsNuevo(false); setModal({ ...a }) }}
                        style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #e8e5de", background: "#fff", cursor: "pointer", fontSize: 11, fontFamily: "Epilogue,sans-serif" }}>
                        ✏️ Editar
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {/* Paginación */}
            {totalPags > 1 && (
              <div style={{ padding: "10px 14px", borderTop: "1px solid #e8e5de", display: "flex", alignItems: "center", gap: 8, justifyContent: "center" }}>
                <button onClick={() => setPaginaActual(p => Math.max(1, p - 1))} disabled={paginaActual === 1}
                  style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid #e8e5de", background: "#fff", cursor: paginaActual === 1 ? "not-allowed" : "pointer", opacity: paginaActual === 1 ? .4 : 1, fontFamily: "Epilogue,sans-serif", fontSize: 12 }}>‹</button>
                <span style={{ fontSize: 11, color: "#9a9590" }}>{paginaActual} / {totalPags} · {filtrados.length} artículos</span>
                <button onClick={() => setPaginaActual(p => Math.min(totalPags, p + 1))} disabled={paginaActual === totalPags}
                  style={{ padding: "4px 12px", borderRadius: 6, border: "1px solid #e8e5de", background: "#fff", cursor: paginaActual === totalPags ? "not-allowed" : "pointer", opacity: paginaActual === totalPags ? .4 : 1, fontFamily: "Epilogue,sans-serif", fontSize: 12 }}>›</button>
              </div>
            )}
          </>
        )}
      </div>

      {/* ── MODAL EDITAR/CREAR ── */}
      {modal && (
        <div style={{ position: "fixed" as const, inset: 0, background: "rgba(0,0,0,.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000 }}>
          <div style={{ background: "#fff", borderRadius: 16, width: 600, maxHeight: "90vh", display: "flex", flexDirection: "column" as const, boxShadow: "0 24px 80px rgba(0,0,0,.25)" }}>
            <div style={{ padding: "16px 20px", borderBottom: "1px solid #e8e5de", display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontFamily: "Playfair Display,serif", fontSize: 16, fontWeight: 800 }}>{esNuevo ? "Nuevo artículo" : "Editar artículo"}</div>
                {!esNuevo && <div style={{ fontSize: 11, color: "#9a9590" }}>{modal.codigo}</div>}
              </div>
              <button onClick={() => setModal(null)} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#9a9590" }}>✕</button>
            </div>
            <div style={{ flex: 1, overflowY: "auto" as const, padding: "16px 20px" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                {/* Nombre */}
                <div style={{ gridColumn: "1/-1" }}>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Nombre *</label>
                  <input value={modal.nombre} onChange={e => setModal({ ...modal, nombre: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 13, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Categoría */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Categoría</label>
                  <select value={modal.categoria} onChange={e => setModal({ ...modal, categoria: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", boxSizing: "border-box" as const }}>
                    {CATS_CATALOGO.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                {/* Subcategoría */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Subcategoría</label>
                  <input value={modal.subcategoria} onChange={e => setModal({ ...modal, subcategoria: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Precio renta */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Precio renta</label>
                  <input type="number" value={modal.precio_renta} onChange={e => setModal({ ...modal, precio_renta: Number(e.target.value) })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "monospace", fontSize: 13, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Precio venta */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Precio venta</label>
                  <input type="number" value={modal.precio_venta} onChange={e => setModal({ ...modal, precio_venta: Number(e.target.value) })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "monospace", fontSize: 13, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Existencia total */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Existencia total</label>
                  <input type="number" value={modal.existencia_total} onChange={e => setModal({ ...modal, existencia_total: Number(e.target.value) })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "monospace", fontSize: 13, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Existencia disponible */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Disponible</label>
                  <input type="number" value={modal.existencia_disponible} onChange={e => setModal({ ...modal, existencia_disponible: Number(e.target.value) })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "monospace", fontSize: 13, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Color */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Color</label>
                  <input value={modal.color} onChange={e => setModal({ ...modal, color: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Material */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Material</label>
                  <input value={modal.material} onChange={e => setModal({ ...modal, material: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Medidas */}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Medidas</label>
                  <input value={modal.medidas} onChange={e => setModal({ ...modal, medidas: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", boxSizing: "border-box" as const }} />
                </div>
                {/* Descripción */}
                <div style={{ gridColumn: "1/-1" }}>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Descripción</label>
                  <textarea value={modal.descripcion} onChange={e => setModal({ ...modal, descripcion: e.target.value })} rows={2}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", resize: "none" as const, boxSizing: "border-box" as const }} />
                </div>
                {/* Notas */}
                <div style={{ gridColumn: "1/-1" }}>
                  <label style={{ fontSize: 10, fontWeight: 700, color: "#9a9590", textTransform: "uppercase" as const, display: "block", marginBottom: 4 }}>Notas</label>
                  <textarea value={modal.notas} onChange={e => setModal({ ...modal, notas: e.target.value })} rows={2}
                    style={{ width: "100%", padding: "8px 10px", border: "1.5px solid #e8e5de", borderRadius: 8, fontFamily: "Epilogue,sans-serif", fontSize: 12, outline: "none", resize: "none" as const, boxSizing: "border-box" as const }} />
                </div>
              </div>
            </div>
            <div style={{ padding: "12px 20px", borderTop: "1px solid #e8e5de", display: "flex", gap: 8 }}>
              <button onClick={() => setModal(null)}
                style={{ flex: 1, padding: "9px", borderRadius: 8, background: "#f5f4f0", border: "1px solid #e8e5de", cursor: "pointer", fontFamily: "Epilogue,sans-serif", fontSize: 12 }}>Cancelar</button>
              <button onClick={guardar} disabled={!modal.nombre.trim() || guardando}
                style={{ flex: 2, padding: "9px", borderRadius: 8, background: guardando ? "#9a9590" : "#1a1814", color: "#fff", border: "none", cursor: guardando ? "not-allowed" : "pointer", fontFamily: "Epilogue,sans-serif", fontSize: 12, fontWeight: 700 }}>
                {guardando ? "Guardando..." : esNuevo ? "Crear artículo" : "Guardar cambios"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}


// ─── COTIZACIONES ───
const VENDEDORES = [
  { nombre: "Karen",   prefijo: "K/H" },
  { nombre: "Cami",    prefijo: "C"   },
  { nombre: "Liliana", prefijo: "L"   },
  { nombre: "Alberto", prefijo: "A"   },
]

// Detectar vendedor desde folio (ej: "K/H-000123" -> "Karen", "L-000001" -> "Liliana")
function vendedorDesdeFolio(folio: string): string {
  if(!folio) return ""
  const f = folio.toUpperCase().trim()
  if(f.startsWith("K/H")||f.startsWith("KH")) return "Karen"
  if(f.startsWith("K-")) return "Karen"
  if(f.startsWith("C-")||f.startsWith("CAM")) return "Cami"
  if(f.startsWith("L-")||f.startsWith("LIL")||f.startsWith("AL-")||f.startsWith("A-L")) return "Liliana"
  if(f.startsWith("A-")||f.startsWith("ALB")) return "Alberto"
  // Detect from cell E6 value (L, K/H, AL, C)
  if(f==="L"||f==="L ") return "Liliana"
  if(f==="K/H"||f==="K") return "Karen"
  if(f==="C") return "Cami"
  if(f==="A"||f==="AL") return "Alberto"
  return ""
}
const ESTADOS_COT = ["borrador","enviada","vista","pendiente","por_vencer","expirada","rechazada"]
const ESTADOS_CONTRATO_NUEVO = ["confirmado","anticipo_recibido","en_produccion","en_montaje","evento_realizado","liquidado","cancelado"]
const ESTADO_LABEL_CONTRATO: Record<string,string> = {
  confirmado:"✅ Confirmado", anticipo_recibido:"💰 Anticipo recibido",
  en_produccion:"⚙️ En producción", en_montaje:"🔧 En montaje",
  evento_realizado:"🎉 Evento realizado", liquidado:"✓ Liquidado", cancelado:"🚫 Cancelado"
}
const ESTADO_COL: Record<string,string> = {
  borrador:"#9a9590", enviada:"#1a3a5c", vista:"#4a2d6e",
  pendiente:"#92580a", por_vencer:"#c47a1a", expirada:"#c4bfb8", rechazada:"#8b2e2e",
  // contrato
  confirmado:"#2d6a4f", anticipo_recibido:"#1a5c4a", en_produccion:"#1a3a5c",
  en_montaje:"#4a2d6e", evento_realizado:"#2d6a4f", liquidado:"#2d6a4f", cancelado:"#8b2e2e"
}
const ESTADO_BG: Record<string,string> = {
  borrador:"#f5f4f0", enviada:"#edf3fa", vista:"#f5f0fc",
  pendiente:"#fdf5e8", por_vencer:"#fef3e2", expirada:"#f5f4f0", rechazada:"#fdf0f0",
  // contrato
  confirmado:"#edf7f2", anticipo_recibido:"#edf7f2", en_produccion:"#edf3fa",
  en_montaje:"#f5f0fc", evento_realizado:"#edf7f2", liquidado:"#edf7f2", cancelado:"#fdf0f0"
}
const ESTADO_LABEL: Record<string,string> = {
  borrador:"📝 Borrador", enviada:"📤 Enviada", vista:"👁️ Vista",
  pendiente:"⏳ Pendiente", por_vencer:"⚠️ Por vencer",
  expirada:"⌛ Expirada", rechazada:"❌ Rechazada"
}

interface Partida {
  articulo_id: string; nombre: string; cantidad: number
  precio_unitario: number; subtotal: number; notas: string
}

interface Cotizacion {
  id: string; folio: string; cliente_nombre: string; cliente_tel: string
  cliente_email: string; lugar_evento: string; fecha_evento: string
  fecha_entrega: string; fecha_desmonte: string; fecha_vigencia: string
  estado: string; vendedor: string; subtotal: number; descuento_pct: number
  descuento_monto: number; aplica_iva: boolean; iva_monto: number; total: number
  notas_cliente: string; condiciones: string; partidas: Partida[]
  creado_en: string; actualizado_en: string
}

const COT_VACIA: Omit<Cotizacion,"id"|"folio"|"creado_en"|"actualizado_en"> = {
  cliente_nombre:"",cliente_tel:"",cliente_email:"",lugar_evento:"",
  fecha_evento:"",fecha_entrega:"",fecha_desmonte:"",
  fecha_vigencia: new Date(Date.now()+15*86400000).toISOString().slice(0,10),
  estado:"borrador",vendedor:"",subtotal:0,descuento_pct:0,descuento_monto:0,
  aplica_iva:false,iva_monto:0,total:0,notas_cliente:"",
  condiciones:"El mobiliario debe devolverse en las mismas condiciones en que fue entregado. El cliente es responsable de cualquier daño o pérdida durante el evento. La vajilla no se monta. El 50% del total se paga al confirmar y el resto el día del evento.",
  partidas:[]
}

function calcularTotales(
  partidas:Partida[],
  descPct:number,
  aplicaIva:boolean,
  descGlobalMonto:number=0  // monto fijo adicional
) {
  // 1. Subtotal bruto = suma pu * cant
  const subtotal = partidas.reduce((s,p)=>s+(p.precio_unitario||0)*(p.cantidad||0),0)
  // 2. Descuentos por artículo (porcentaje del precio del artículo)
  const descArticulos = partidas.reduce((s,p)=>{
    if(!(p as any).aplica_descuento) return s
    const base=(p.precio_unitario||0)*(p.cantidad||0)
    const pct=Number((p as any).descuento_pct_art)||0
    return s+Math.round(base*pct/100)
  },0)
  // 3. Descuento global: % sobre subtotal bruto, más monto fijo si hay
  const descPctMonto = Math.round(subtotal * (descPct||0) / 100)
  const descTotal = descArticulos + descPctMonto + (descGlobalMonto||0)
  // 4. Base después de descuentos
  const base = Math.max(0, subtotal - descTotal)
  // 5. IVA sobre base (después de descuentos)
  const ivaMonto = aplicaIva ? Math.round(base * 0.16) : 0
  // 6. Total
  const total = base + ivaMonto
  return { subtotal, descuento_monto: descTotal, iva_monto: ivaMonto, total }
}

const DELETE_PWD = "LITA2024"

function CotizacionesSection({ token, personal, logoUrl, vendedorActual, esAdmin, contratos: contratosAll = [] }: { token: string, personal: any, logoUrl?: string, vendedorActual?: string, esAdmin?: boolean, contratos?: any[] }) {
  const [cots, setCots] = useState<Cotizacion[]>([])
  const [cargando, setCargando] = useState(true)
  const [filtroEst, setFiltroEst] = useState("todos")
  const [filtroVendCot, setFiltroVendCot] = useState(vendedorActual||"todos")
  const [vista, setVista] = useState<"lista"|"form"|"detalle"|"calendario">("lista")
  const [vistaCalCot, setVistaCalCot] = useState<"semana"|"mes">("semana")
  const [semOffCot, setSemOffCot] = useState(0)
  const [cotActual, setCotActual] = useState<Partial<Cotizacion>>({...COT_VACIA})
  const [descTipo, setDescTipo] = useState<"pct"|"monto">("pct")
  const [descGlobalMonto, setDescGlobalMonto] = useState(0)
  const [esNueva, setEsNueva] = useState(true)
  const [guardando, setGuardando] = useState(false)
  const [busqArt, setBusqArt] = useState("")
  const [cantBusq, setCantBusq] = useState(1)
  const [catFiltroArt, setCatFiltroArt] = useState("TODOS")
  const [artsSugeridos, setArtsSugeridos] = useState<any[]>([])
  const [buscandoArt, setBuscandoArt] = useState(false)
  const [activeRow, setActiveRow] = useState<number|null>(null)
  const [rowSugeridos, setRowSugeridos] = useState<any[]>([])
  const [paso, setPaso] = useState(1) // 1=datos, 2=articulos, 3=resumen
  const [busqCliente, setBusqCliente] = useState("")
  const [clientesSug, setClientesSug] = useState<{nombre:string,tel:string,lugar:string,email:string}[]>([])
  const [mostrarSug, setMostrarSug] = useState(false)
  const [formGuardado, setFormGuardado] = useState(false)

  // Dirty = tiene cliente o artículos y no se ha guardado
  const formTieneDatos = !!(
    (cotActual.cliente_nombre && cotActual.cliente_nombre.trim()) ||
    ((cotActual.partidas||[]).length > 0)
  )
  const formDirty = vista === "form" && formTieneDatos && !formGuardado

  // Aviso al cerrar/recargar el navegador + flag global
  useEffect(()=>{
    (window as any).__cotFormDirty = formDirty
    const handler=(e:BeforeUnloadEvent)=>{
      if(formDirty){e.preventDefault();e.returnValue=""}
    }
    window.addEventListener("beforeunload",handler)
    return ()=>{
      window.removeEventListener("beforeunload",handler)
      ;(window as any).__cotFormDirty = false
    }
  },[formDirty])

  // Helper: navegar fuera del form con confirmación
  const salirDelForm = (action:()=>void) => {
    if(formDirty){
      const ok=window.confirm("¿Salir sin guardar?\n\nTienes una cotización en progreso. ¿Deseas descartarla o preferes guardarla como borrador?\n\nAcepta para descartar · Cancela para seguir editando")
      if(!ok) return
    }
    setFormGuardado(false)
    action()
  }

  // Password-protected delete
  const eliminarConContrasena = async (tipo: "cotizacion"|"contrato", id: string, onSuccess: ()=>void) => {
    const pwd = window.prompt(`Ingresa la contraseña para eliminar esta ${tipo}:`)
    if(pwd === null) return // canceló
    if(pwd !== DELETE_PWD){
      alert("❌ Contraseña incorrecta. No se eliminó el registro.")
      return
    }
    const endpoint = tipo === "cotizacion" ? `/api/cotizaciones?id=${id}` : `/api/contratos?id=${id}`
    const res = await fetch(endpoint, { method: "DELETE", headers: { Authorization: `Bearer ${token}` }})
    if(res.ok) onSuccess()
    else alert("Error al eliminar. Intenta de nuevo.")
  }

  const cargar = async () => {
    setCargando(true)
    // Load from cotizaciones table
    const url = filtroEst !== "todos" ? `/api/cotizaciones?estado=${filtroEst}` : "/api/cotizaciones"
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
    const fromDB = await res.json()
    // Also load contratos tipo=cotizacion from Excel system
    const res2 = await fetch("/api/contratos", { headers: { Authorization: `Bearer ${token}` } })
    const contratosData = await res2.json()
    const cotFromExcel: Cotizacion[] = (Array.isArray(contratosData) ? contratosData : [])
      .filter((x:any) => x.tipo === "cotizacion")
      .map((x:any) => ({
        id: "excel_"+x.id,
        folio: x.folio || x.archivo?.replace(".xlsx","") || "COT-EXCEL",
        cliente_nombre: x.cliente || x.archivo || "",
        cliente_tel: x.tel || x.telefono || "",
        cliente_email: "",
        lugar_evento: x.lugar || "",
        fecha_evento: x.fecha_evento || "",
        fecha_entrega: x.fecha_entrega || "",
        fecha_desmonte: x.fecha_desmonte || "",
        fecha_vigencia: "",
        estado: "enviada",
        vendedor: x.vendedor || "",
        subtotal: x.total || 0,
        descuento_pct: 0,
        descuento_monto: 0,
        aplica_iva: false,
        iva_monto: 0,
        total: x.total || 0,
        notas_cliente: x.notas || "",
        condiciones: "",
        partidas: (x.articulos||[]).map((a:any) => ({
          articulo_id: "", nombre: a.nombre, cantidad: a.cantidad,
          precio_unitario: a.pu||0, subtotal: a.importe||0, notas: ""
        })),
        creado_en: x.fecha_evento || "",
        actualizado_en: x.fecha_evento || "",
        _fromExcel: true
      } as any))
    // Cotizaciones: exclude converted ones (they become contratos)
    const dbCots = (Array.isArray(fromDB) ? fromDB : []).filter((x:any) => x.estado !== "convertida")
    const allCots = [...dbCots, ...cotFromExcel]
      .filter(c => filtroEst === "todos" || (c as any).estado === filtroEst)
      .sort((a,b) => (b.creado_en||"").localeCompare(a.creado_en||""))
    setCots(allCots)
    setCargando(false)
  }

  useState(() => { cargar() })

  // Buscar artículos del catálogo
  const buscarArticulos = async (q: string, cat?: string) => {
    if (!q || q.trim().length < 1) { setArtsSugeridos([]); return }
    setBuscandoArt(true)
    try {
      const params = new URLSearchParams({busq: q.trim(), activo: "true"})
      if(cat && cat !== "TODOS") params.append("categoria", cat)
      const res = await fetch(`/api/catalogo?${params}`, {
        headers: { Authorization: `Bearer ${token}` }
      })
      const data = await res.json()
      setArtsSugeridos(Array.isArray(data) ? data.slice(0, 40) : [])
    } catch { setArtsSugeridos([]) }
    setBuscandoArt(false)
  }
  const CATS_ART = ["TODOS","MOBILIARIO","FLORES","MANTELERÍA","VAJILLA","SERVICIOS","CARPAS"]
  // Buscar clientes desde contratos existentes
  const buscarClientes = (q: string) => {
    if (!q || q.trim().length < 1) { setClientesSug([]); setMostrarSug(false); return }
    const qLow = q.trim().toLowerCase()
    const vistos = new Set<string>()
    const unicos = contratosAll
      .filter((x:any) => {
        const nombre = (x.cliente||x.archivo||"").trim()
        const k = nombre.toLowerCase()
        if(!k || vistos.has(k)) return false
        if(!k.includes(qLow)) return false
        vistos.add(k); return true
      })
      .slice(0, 10)
      .map((x:any) => ({
        nombre: (x.cliente||x.archivo||"").trim(),
        tel: (x.tel||x.telefono||"").trim(),
        lugar: (x.lugar||"").trim(),
        email: ""
      }))
    setClientesSug(unicos)
    setMostrarSug(unicos.length > 0)
  }

  const seleccionarCliente = (cli: {nombre:string,tel:string,lugar:string,email:string}) => {
    setCotActual(prev => ({
      ...prev,
      cliente_nombre: cli.nombre,
      cliente_tel: cli.tel || prev.cliente_tel,
      cliente_email: cli.email || prev.cliente_email,
      lugar_evento: prev.lugar_evento || cli.lugar,
    }))
    setBusqCliente(cli.nombre)
    setClientesSug([])
    setMostrarSug(false)
  }

  // Generar folio con prefijo del vendedor
  const prefijoVendedor = (nombre: string) => {
    const v = VENDEDORES.find(v => v.nombre === nombre)
    return v ? v.prefijo : "COT"
  }

  const agregarPartida = (art: any) => {
    const nueva: Partida = {
      articulo_id: art.id, nombre: art.nombre,
      cantidad: 1, precio_unitario: art.precio_renta || 0,
      subtotal: art.precio_renta || 0, notas: ""
    }
    const partidas = [...(cotActual.partidas || []), nueva]
    const tots = calcularTotales(partidas, cotActual.descuento_pct || 0, cotActual.aplica_iva || false)
    setCotActual({ ...cotActual, partidas, ...tots })
    setBusqArt(""); setArtsSugeridos([])
  }

  const actualizarPartida = (i: number, campo: keyof Partida, valor: any) => {
    const partidas = [...(cotActual.partidas || [])]
    partidas[i] = { ...partidas[i], [campo]: valor }
    if (campo === "cantidad" || campo === "precio_unitario") {
      partidas[i].subtotal = partidas[i].cantidad * partidas[i].precio_unitario
    }
    const tots = calcularTotales(partidas, cotActual.descuento_pct || 0, cotActual.aplica_iva || false)
    setCotActual({ ...cotActual, partidas, ...tots })
  }

  const quitarPartida = (i: number) => {
    const partidas = (cotActual.partidas || []).filter((_,j) => j !== i)
    const tots = calcularTotales(partidas, cotActual.descuento_pct || 0, cotActual.aplica_iva || false)
    setCotActual({ ...cotActual, partidas, ...tots })
  }

  const actualizarDescuento = (pct: number) => {
    const tots = calcularTotales(cotActual.partidas || [], pct, cotActual.aplica_iva || false)
    setCotActual({ ...cotActual, descuento_pct: pct, ...tots })
  }

  const toggleIva = (v: boolean) => {
    const tots = calcularTotales(cotActual.partidas || [], cotActual.descuento_pct || 0, v)
    setCotActual({ ...cotActual, aplica_iva: v, ...tots })
  }

  const guardar = async (estadoNuevo?: string) => {
    setGuardando(true)
    setFormGuardado(true)
    const body = { ...cotActual }
    if (estadoNuevo) body.estado = estadoNuevo
    // Add vendedor prefix to body for folio generation
    if (esNueva && body.vendedor) {
      body._prefijo = prefijoVendedor(body.vendedor)
    }
    let res
    // Si viene de Excel (id con prefijo) o es nueva → POST
    const rawId = String(cotActual.id || "")
    const uuidClean = (rawId.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i)||[""])[0]
    const isExcel = rawId !== uuidClean || rawId.startsWith("_") || rawId.includes("excel")
    if (esNueva || isExcel) {
      // Crear nueva — quitar el id del excel para que Supabase genere uno nuevo
      const { id: _id, fromExcel: _fx, ...bodyClean } = body as any
      res = await fetch("/api/cotizaciones", {
        method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify(bodyClean)
      })
    } else {
      res = await fetch(`/api/cotizaciones?id=${uuidClean}`, {
        method: "PATCH", headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify(body)
      })
    }
    const data = await res.json()
    setGuardando(false)
    if (res.ok && (data.id || data.folio)) {
      setCotActual(data)
      setEsNueva(false)
      cargar()
      setVista("detalle")
    } else {
      alert("Error al guardar: " + (data?.error || data?.message || "Error desconocido"))
      console.error("Error guardando cotización:", data)
    }
  }

  const cambiarEstado = async (id: string, estado: string) => {
    await fetch(`/api/cotizaciones?id=${id}`, {
      method: "PATCH", headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
      body: JSON.stringify({ estado })
    })
    cargar()
  }

  const duplicar = (cot: Cotizacion) => {
    const { id, folio, creado_en, actualizado_en, ...resto } = cot
    setCotActual({ ...resto, estado: "borrador", fecha_vigencia: new Date(Date.now()+15*86400000).toISOString().slice(0,10) })
    setEsNueva(true); setPaso(1); setVista("form")
  }

  const convertirAContrato = async (cot: Cotizacion) => {
    if (!window.confirm(`¿Convertir cotización ${cot.folio} en contrato confirmado?\n\nSe creará un contrato con todos los artículos y datos del evento.`)) return

    const isExcel = (cot as any)._fromExcel

    if(isExcel){
      // Excel cotizacion: just update tipo in contratos table
      const realId = cot.id.replace("excel_","")
      const r = await fetch(`/api/contratos?id=${realId}`, {
        method:"PATCH",
        headers:{"Content-Type":"application/json", Authorization:`Bearer ${token}`},
        body: JSON.stringify({tipo:"contrato", estado_entrega:"pend"})
      })
      if(!r.ok){
        const err = await r.json().catch(()=>({}))
        alert("Error al convertir: " + (err.error || r.status))
        return
      }
    } else {
      // System cotizacion: create new row in contratos table
      const contratoBody = {
        archivo: cot.folio || "",
        cliente: cot.cliente_nombre || "",
        lugar: cot.lugar_evento || "",
        tel: cot.cliente_tel || "",
        telefono: cot.cliente_tel || "",
        fecha_evento: cot.fecha_evento || "",
        fecha_entrega: cot.fecha_entrega || cot.fecha_evento || "",
        fecha_desmonte: cot.fecha_desmonte || cot.fecha_evento || "",
        folio: cot.folio || "",
        vendedor: cot.vendedor || "",
        tipo: "contrato",
        total: cot.total || 0,
        a_cuenta: 0,
        cobrado: 0,
        pagos: [],
        estado_entrega: "pend",
        estado_desmonte: "pend",
        asig_entrega: [],
        asig_desmonte: [],
        notas: cot.notas_cliente || "",
        articulos: (cot.partidas || []).map((p:any) => ({
          nombre: p.nombre || "",
          cantidad: p.cantidad || 0,
          pu: p.precio_unitario || 0,
          importe: p.subtotal || 0,
          seccion: p.notas || ""
        }))
      }

      const res = await fetch("/api/contratos", {
        method:"POST",
        headers:{"Content-Type":"application/json", Authorization:`Bearer ${token}`},
        body: JSON.stringify(contratoBody)
      })

      if(!res.ok){
        const err = await res.json()
        alert("Error al crear el contrato: " + (err.error || res.status))
        return
      }

      // Mark cotizacion as converted
      await cambiarEstado(cot.id, "convertida")
    }

    await cargar(token)
    alert(`✓ Cotización ${cot.folio} convertida a contrato exitosamente.\nYa puedes verla en el módulo de Contratos.`)
  }

  // Generar mensaje WhatsApp
    const imprimirCotizacion = (cot: Cotizacion) => {
    const logoSrc = logoUrl || (typeof window!=="undefined" ? localStorage.getItem("pf_logo")||"/logo.png" : "/logo.png")
    const fechaEvento = cot.fecha_evento ? new Date(cot.fecha_evento+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"}) : "Por confirmar"
    const fechaVig = cot.fecha_vigencia ? new Date(cot.fecha_vigencia+"T12:00:00").toLocaleDateString("es-MX",{day:"numeric",month:"long",year:"numeric"}) : "15 días"
    const filas = (cot.partidas||[]).map((p,i) => {
      const bruto=(p.precio_unitario||0)*(p.cantidad||0)
      const descArt=(p as any).aplica_descuento?Math.round(bruto*((p as any).descuento_pct_art||0)/100):0
      const neto=bruto-descArt
      return `
      <tr style="border-bottom:1px solid #f0ece4;background:${i%2===0?"#fff":"#fafaf8"}">
        <td style="padding:8px 10px;text-align:center;font-size:13px">${p.cantidad}</td>
        <td style="padding:8px 10px;font-weight:600;font-size:13px">${p.nombre}</td>
        <td style="padding:8px 10px;text-align:right;font-family:monospace;font-size:13px">$${(p.precio_unitario||0).toLocaleString("es-MX")}</td>
        <td style="padding:8px 10px;text-align:right;font-family:monospace;font-weight:700;font-size:13px;color:#1a3a5c">
          $${neto.toLocaleString("es-MX")}
          ${descArt>0?`<div style="font-size:9px;color:#8b2e2e">-${(p as any).descuento_pct_art}% (-$${descArt.toLocaleString("es-MX")})</div>`:""}
        </td>
      </tr>`
    }).join("")

    const html = `<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<title>Cotización ${cot.folio}</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet"/>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family: Inter, Arial, sans-serif; font-size: 13px; color: #1a1814; background: #fff; }
  @media print {
    .no-print { display: none !important; }
    @page { margin: 12mm 15mm; size: A4 portrait; }
    body { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  }
  .page { max-width: 820px; margin: 0 auto; padding: 32px; }
  table { border-collapse: collapse; width: 100%; }
  th { background: #1a1814 !important; color: #fff !important; }
</style>
</head><body>
<div class="page">
  <!-- Botones -->
  <div class="no-print" style="text-align:right;margin-bottom:20px;display:flex;gap:10px;justify-content:flex-end">
    <button onclick="window.print()" style="padding:10px 24px;background:#1a1814;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">🖨️ Imprimir / PDF</button>
    <button onclick="window.close()" style="padding:10px 24px;background:#f5f4f0;border:1px solid #ccc;border-radius:8px;font-size:14px;cursor:pointer">✕ Cerrar</button>
  </div>

  <!-- Header con logo -->
  <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px;padding-bottom:20px;border-bottom:3px solid #1a1814">
    <div>
      <img src="${logoSrc}" alt="Poliflor" style="height:60px;width:auto;object-fit:contain;margin-bottom:6px" onerror="this.style.display='none'">
      <div style="font-size:11px;color:#9a9590;margin-top:4px">Renta de Mobiliario para Eventos</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:24px;font-weight:800;color:#1a1814;letter-spacing:-0.5px">COTIZACIÓN</div>
      <div style="font-size:20px;font-weight:700;color:#1a3a5c;margin-top:2px">${cot.folio}</div>
      <div style="font-size:11px;color:#9a9590;margin-top:4px">Vigencia: ${fechaVig}</div>
      ${cot.vendedor ? `<div style="font-size:11px;color:#9a9590">Vendedor: <strong>${cot.vendedor}</strong></div>` : ""}
    </div>
  </div>

  <!-- Datos del cliente y evento -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:28px">
    <div style="background:#f8f6f2;border-radius:10px;padding:16px">
      <div style="font-size:10px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Cliente</div>
      <div style="font-size:16px;font-weight:800;margin-bottom:4px">${cot.cliente_nombre||"—"}</div>
      ${cot.cliente_tel ? `<div style="font-size:12px;color:#4a4640;margin-top:3px">📞 ${cot.cliente_tel}</div>` : ""}
      ${cot.cliente_email ? `<div style="font-size:12px;color:#4a4640;margin-top:3px">✉️ ${cot.cliente_email}</div>` : ""}
    </div>
    <div style="background:#f8f6f2;border-radius:10px;padding:16px">
      <div style="font-size:10px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Evento</div>
      <div style="font-size:13px;font-weight:700;margin-bottom:4px">📅 ${fechaEvento}</div>
      ${cot.lugar_evento ? `<div style="font-size:12px;color:#4a4640;margin-top:4px">📍 ${cot.lugar_evento}</div>` : ""}
      ${cot.fecha_entrega ? `<div style="font-size:11px;color:#9a9590;margin-top:4px">Entrega: ${cot.fecha_entrega} · Desmonte: ${cot.fecha_desmonte||"—"}</div>` : ""}
    </div>
  </div>

  <!-- Tabla de artículos -->
  <table style="width:100%;border-collapse:collapse;margin-bottom:24px">
    <thead>
      <tr style="background:#1a1814;color:#fff">
        <th style="padding:10px;text-align:center;font-size:11px;width:70px">Cant.</th>
        <th style="padding:10px;text-align:left;font-size:11px">Artículo / Descripción</th>
        <th style="padding:10px;text-align:right;font-size:11px;width:100px">P. Unit.</th>
        <th style="padding:10px;text-align:right;font-size:11px;width:110px">Subtotal</th>
      </tr>
    </thead>
    <tbody>${filas}</tbody>
  </table>

  <!-- Totales -->
  <div style="display:flex;justify-content:flex-end;margin-bottom:28px">
    <div style="width:280px">
      <div style="display:flex;justify-content:space-between;padding:8px 0;font-size:13px;border-bottom:1px solid #f0ece4">
        <span style="color:#9a9590">Subtotal</span>
        <span style="font-family:monospace;font-weight:600">$${(cot.subtotal||0).toLocaleString("es-MX")}</span>
      </div>
      ${cot.descuento_pct > 0 ? `
      <div style="display:flex;justify-content:space-between;padding:8px 0;font-size:13px;border-bottom:1px solid #f0ece4;color:#2d6a4f">
        <span>Descuento (${cot.descuento_pct}%)</span>
        <span style="font-family:monospace;font-weight:600">-$${(cot.descuento_monto||0).toLocaleString("es-MX")}</span>
      </div>` : ""}
      ${cot.aplica_iva ? `
      <div style="display:flex;justify-content:space-between;padding:8px 0;font-size:13px;border-bottom:1px solid #f0ece4;color:#4a2d6e">
        <span>IVA (16%)</span>
        <span style="font-family:monospace;font-weight:600">$${(cot.iva_monto||0).toLocaleString("es-MX")}</span>
      </div>` : ""}
      <div style="display:flex;justify-content:space-between;padding:12px 0;background:#1a1814;color:#fff;border-radius:8px;margin-top:8px;padding:12px 14px">
        <span style="font-size:15px;font-weight:800">TOTAL</span>
        <span style="font-family:monospace;font-size:20px;font-weight:800">$${(cot.total||0).toLocaleString("es-MX")}</span>
      </div>
    </div>
  </div>

  <!-- Notas y condiciones -->
  ${cot.notas_cliente ? `
  <div style="background:#fdf5e8;border-left:4px solid #92580a;border-radius:0 8px 8px 0;padding:14px 16px;margin-bottom:16px">
    <div style="font-size:10px;font-weight:700;color:#92580a;text-transform:uppercase;margin-bottom:6px">Notas</div>
    <div style="font-size:12px;color:#4a4640;line-height:1.6">${cot.notas_cliente}</div>
  </div>` : ""}
  ${cot.condiciones ? `
  <div style="background:#f8f6f2;border-radius:8px;padding:14px 16px;margin-bottom:24px">
    <div style="font-size:10px;font-weight:700;color:#9a9590;text-transform:uppercase;margin-bottom:6px">Términos y condiciones</div>
    <div style="font-size:11px;color:#4a4640;line-height:1.7">${cot.condiciones}</div>
  </div>` : ""}

  <!-- Footer firma -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:40px;padding-top:20px;border-top:1px solid #e8e5de">
    <div style="text-align:center">
      <div style="border-top:1px solid #1a1814;padding-top:8px;margin-top:50px;font-size:11px;color:#9a9590">Firma del cliente</div>
      <div style="font-size:11px;color:#9a9590;margin-top:4px">${cot.cliente_nombre||""}</div>
    </div>
    <div style="text-align:center">
      <div style="border-top:1px solid #1a1814;padding-top:8px;margin-top:50px;font-size:11px;color:#9a9590">Autorizado por Poliflor</div>
      <div style="font-size:11px;color:#9a9590;margin-top:4px">${cot.vendedor||"Poliflor"}</div>
    </div>
  </div>
</div>
</body></html>`

    // Nombre: FOLIO_CLIENTE_FECHA.html
    // PDF: COTIZACION_FOLIO_CLIENTE_YYYY-MM-DD
    const slug=(s:string,n=25)=>s.slice(0,n).trim().replace(/\s+/g,"-").replace(/[^a-zA-Z0-9-]/g,"").toUpperCase()
    const nombrePDF=`COTIZACION_${slug(cot.folio||"COT",15)}_${slug(cot.cliente_nombre||"CLIENTE")}_${cot.fecha_evento||new Date().toISOString().slice(0,10)}`
    // Botones PDF: la clave es capturar #doc-main, NO body completo
    const btns=`
<div id="pflbtns" style="position:fixed;top:12px;right:12px;z-index:9999;background:#fff;padding:14px 16px;border-radius:12px;box-shadow:0 6px 28px rgba(0,0,0,.22);min-width:220px;font-family:Arial,sans-serif">
  <div style="font-size:10px;color:#9a9590;font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">📄 ${nombrePDF}</div>
  <div style="display:flex;flex-direction:column;gap:6px">
    <button id="btn-pdf" style="padding:10px 14px;background:#1a3a5c;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer;font-size:13px">⬇️ Descargar PDF</button>
    <button onclick="document.getElementById('pflbtns').style.display='none';window.print()" style="padding:10px 14px;background:#f5f4f0;color:#1a1814;border:1px solid #e8e5de;border-radius:8px;font-weight:600;cursor:pointer;font-size:12px">🖨️ Imprimir</button>
  </div>
  <div id="pdf-msg" style="font-size:10px;color:#9a9590;margin-top:6px;min-height:14px"></div>
</div>
<script>
document.getElementById("btn-pdf").onclick=function(){
  var btn=this;
  var msg=document.getElementById("pdf-msg");
  btn.textContent="Cargando...";
  btn.disabled=true;
  msg.textContent="Preparando PDF...";
  var s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js";
  s.onload=function(){
    var target=document.getElementById("doc-main");
    if(!target){msg.textContent="Error: contenido no encontrado";btn.textContent="⬇️ Descargar PDF";btn.disabled=false;return}
    btn.textContent="Generando...";
    msg.textContent="Esto toma unos segundos...";
    document.getElementById("pflbtns").style.opacity="0.5";
    html2pdf().set({
      margin:[8,8,8,8],filename:"${nombrePDF}.pdf",
      image:{type:"jpeg",quality:0.98},
      html2canvas:{scale:2,useCORS:true,logging:false,removeContainer:true},
      jsPDF:{unit:"mm",format:"a4",orientation:"portrait",compress:true}
    }).from(target).save().then(function(){
      btn.textContent="⬇️ Descargar PDF";btn.disabled=false;
      document.getElementById("pflbtns").style.opacity="1";
      msg.textContent="✓ PDF guardado";
    }).catch(function(e){
      btn.textContent="⬇️ Descargar PDF";btn.disabled=false;
      document.getElementById("pflbtns").style.opacity="1";
      msg.textContent="Error: "+e.message;
    });
  };
  s.onerror=function(){msg.textContent="Error cargando librería";btn.textContent="⬇️ Descargar PDF";btn.disabled=false;};
  document.head.appendChild(s);
};
</script>`
    const htmlFinal=html
      .replace("</head>",'</head><style>#pflbtns{font-family:Arial,sans-serif}@media print{#pflbtns{display:none!important}}</style>')
      .replace('<div class="page">','<div class="page" id="doc-main">')
      .replace("</body>",btns+"</body>")
      .replace('<button onclick="window.print()"','<button style="display:none"')
    const blob=new Blob([htmlFinal],{type:"text/html;charset=utf-8"})
    const url=URL.createObjectURL(blob)
    const a=document.createElement("a")
    a.href=url; a.target="_blank"; a.rel="noopener"
    document.body.appendChild(a); a.click()
    setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url)},1000)
  }


  const enviarWhatsApp = (cot: Cotizacion) => {
    const fecha = cot.fecha_evento ? new Date(cot.fecha_evento+"T12:00:00").toLocaleDateString("es-MX",{day:"numeric",month:"long",year:"numeric"}) : "por confirmar"
    const nl = "\n"
    const items = (cot.partidas||[]).slice(0,8).map((p,i)=>"  "+(i+1)+". "+p.cantidad+"x "+p.nombre+" — $"+p.subtotal.toLocaleString()).join(nl)
    const masItems = (cot.partidas||[]).length > 8 ? nl+"  ...y "+((cot.partidas||[]).length-8)+" artículos más" : ""
    let msg = "Hola "+cot.cliente_nombre+" 👋"+nl+nl
    msg += "Te compartimos la cotización *"+cot.folio+"* de *Poliflor* 🌸"+nl+nl
    msg += "📅 Fecha del evento: *"+fecha+"*"+nl
    if (cot.lugar_evento) msg += "📍 Lugar: "+cot.lugar_evento+nl
    msg += nl+"*📦 Artículos cotizados:*"+nl+items+masItems+nl+nl
    if (cot.descuento_pct > 0) msg += "🏷️ Descuento: "+cot.descuento_pct+"% (-$"+cot.descuento_monto.toLocaleString()+")"+nl
    if (cot.aplica_iva) msg += "📊 IVA (16%): $"+cot.iva_monto.toLocaleString()+nl
    msg += nl+"💰 *Total: $"+cot.total.toLocaleString()+"*"+nl
    msg += nl+"⏰ Vigencia: "+(cot.fecha_vigencia ? new Date(cot.fecha_vigencia+"T12:00:00").toLocaleDateString("es-MX",{day:"numeric",month:"long"}) : "15 días")+nl
    if (cot.condiciones) msg += nl+"📋 *Condiciones:*"+nl+cot.condiciones.slice(0,200)+"..."+nl
    msg += nl+"¿Te gustaría confirmar? Escríbenos para apartar la fecha 😊"+nl+nl+"— *Poliflor* 🌸"
    const tel = (cot.cliente_tel||"").replace(/[^0-9]/g,"")
    const url = tel ? "https://wa.me/52"+tel+"?text="+encodeURIComponent(msg) : "https://wa.me/?text="+encodeURIComponent(msg)
    window.open(url,"_blank")
    if (cot.estado === "borrador") cambiarEstado(cot.id, "enviada")
  }

  const fmt = (n:number) => "$"+Math.round(n||0).toLocaleString("es-MX")

  // ── Días hasta vencimiento
  const diasVigencia = (fecha: string) => {
    if (!fecha) return null
    const diff = Math.ceil((new Date(fecha+"T23:59:59").getTime() - Date.now()) / 86400000)
    return diff
  }

  // ── LISTA
  if (vista === "lista" || vista === "calendario") return (
    <div>
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 18px",marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap" as const}}>
          <div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>📋 Cotizaciones</div>
            <div style={{fontSize:11,color:"#9a9590"}}>{cots.length} cotizaciones</div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8}}>
            <button onClick={()=>{setCotActual({...COT_VACIA});setEsNueva(true);setPaso(1);setVista("form")}}
              style={{padding:"8px 18px",borderRadius:8,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
              + Nueva cotización
            </button>
            <button onClick={()=>setVista(v=>v==="calendario"?"lista":"calendario")}
              style={{padding:"8px 14px",borderRadius:8,background:vista==="calendario"?"#2563eb":"#f5f4f0",color:vista==="calendario"?"#fff":"#4a4640",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",gap:6}}>
              📅 {vista==="calendario"?"Lista":"Calendario"}
            </button>
          </div>
        </div>
        {/* Filtros estado */}
        <div style={{display:"flex",gap:6,marginTop:12,flexWrap:"wrap" as const}}>
          {["todos",...ESTADOS_COT].map(e=>{
            const n = e==="todos" ? cots.length : cots.filter(c=>c.estado===e).length
            return(
              <button key={e} onClick={()=>{setFiltroEst(e);setTimeout(cargar,0)}}
                style={{padding:"3px 10px",borderRadius:14,border:`1.5px solid ${filtroEst===e?(ESTADO_COL[e]||"#1a1814"):"#e8e5de"}`,background:filtroEst===e?(ESTADO_BG[e]||"#f5f4f0"):"#fff",color:filtroEst===e?(ESTADO_COL[e]||"#1a1814"):"#4a4640",fontSize:10,fontWeight:filtroEst===e?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>
                {e==="todos"?"Todos":ESTADO_LABEL[e]} ({n})
              </button>
            )
          })}
        </div>
        {/* Filtro por vendedor */}
        <div style={{display:"flex",gap:6,marginTop:8,flexWrap:"wrap" as const,alignItems:"center"}}>
          <span style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em"}}>Vendedor:</span>
          {/* Botón "Solo los míos" — acceso rápido */}
          {vendedorActual&&(
            <button onClick={()=>setFiltroVendCot(filtroVendCot===vendedorActual?"todos":vendedorActual)}
              style={{padding:"3px 12px",borderRadius:14,border:`1.5px solid ${filtroVendCot===vendedorActual?"#2563eb":"#e8e5de"}`,background:filtroVendCot===vendedorActual?"#2563eb":"#eff6ff",color:filtroVendCot===vendedorActual?"#fff":"#2563eb",fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"Epilogue,sans-serif",display:"flex",alignItems:"center",gap:4}}>
              {filtroVendCot===vendedorActual?"✓ Solo los míos":"👤 Solo los míos"}
            </button>
          )}
          {(esAdmin?["todos","Karen","Cami","Liliana","Alberto"]:[vendedorActual||""]).filter(Boolean).map(v=>{
            const n=v==="todos"?cots.length:cots.filter((x:any)=>(x.vendedor||vendedorDesdeFolio(x.folio||""))===v).length
            return(
              <button key={v} onClick={()=>setFiltroVendCot(v)}
                style={{padding:"3px 10px",borderRadius:14,border:`1.5px solid ${filtroVendCot===v?"#0f172a":"#e8e5de"}`,background:filtroVendCot===v?"#0f172a":"#fff",color:filtroVendCot===v?"#fff":"#4a4640",fontSize:10,fontWeight:filtroVendCot===v?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>
                {v==="todos"?"Todos":v} ({n})
              </button>
            )
          })}
        </div>
      </div>

      {/* KPIs operativos - sin montos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
        {[
          {l:"Total",v:cots.length,c:"#1a1814"},
          {l:"Enviadas",v:cots.filter(c=>c.estado==="enviada").length,c:"#1a3a5c"},
          {l:"Pendientes",v:cots.filter(c=>c.estado==="pendiente").length,c:"#92580a"},
          {l:"Por vencer",v:cots.filter(c=>{const d=diasVigencia(c.fecha_vigencia);return d!==null&&d>=0&&d<=7&&(c.estado==="enviada"||c.estado==="pendiente")}).length,c:"#8b2e2e"},
        ].map((k,i)=>(
          <div key={i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:9,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:4}}>{k.l}</div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:k.c}}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Lista */}
      {/* ── VISTA CALENDARIO ── */}
      {vista==="calendario"&&(()=>{
        const hoyC=new Date();hoyC.setHours(0,0,0,0)
        const dowC=hoyC.getDay()===0?6:hoyC.getDay()-1
        const lunesBase=new Date(hoyC);lunesBase.setDate(hoyC.getDate()-dowC+semOffCot*7)
        const diasSem=Array.from({length:7},(_,i)=>{const d=new Date(lunesBase);d.setDate(lunesBase.getDate()+i);return d})
        const DIAS=["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
        const MESES=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
        const isoD=(d:Date)=>d.toISOString().split("T")[0]

        // Group cots by fecha_evento
        const cotsByDate:Record<string,any[]>={}
        cots.forEach(cot=>{
          if(!cot.fecha_evento) return
          if(!cotsByDate[cot.fecha_evento]) cotsByDate[cot.fecha_evento]=[]
          cotsByDate[cot.fecha_evento].push(cot)
        })

        const ESTADO_COLOR:Record<string,string>={
          borrador:"#94a3b8",enviada:"#2563eb",pendiente:"#92580a",
          aceptada:"#2d6a4f",rechazada:"#8b2e2e",convertida:"#4a2d6e",expirada:"#9a9590"
        }

        return(
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
            {/* Toolbar calendario */}
            <div style={{padding:"12px 16px",borderBottom:"1px solid #f0ece4",display:"flex",alignItems:"center",gap:10}}>
              <button onClick={()=>setSemOffCot(v=>v-1)}
                style={{width:32,height:32,borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>‹</button>
              <button onClick={()=>setSemOffCot(0)}
                style={{padding:"4px 12px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,color:"#4a4640"}}>Hoy</button>
              <button onClick={()=>setSemOffCot(v=>v+1)}
                style={{width:32,height:32,borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>›</button>
              <span style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,flex:1}}>
                {MESES[lunesBase.getMonth()]} {lunesBase.getFullYear()}
                {lunesBase.getMonth()!==diasSem[6].getMonth()&&` — ${MESES[diasSem[6].getMonth()]}`}
              </span>
              <div style={{display:"flex",gap:4}}>
                {(["semana","mes"] as const).map(m=>(
                  <button key={m} onClick={()=>setVistaCalCot(m)}
                    style={{padding:"4px 10px",borderRadius:7,border:"none",background:vistaCalCot===m?"#0f172a":"#f5f4f0",color:vistaCalCot===m?"#fff":"#4a4640",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700,cursor:"pointer",textTransform:"capitalize" as const}}>
                    {m==="semana"?"Semana":"Mes"}
                  </button>
                ))}
              </div>
            </div>

            {/* Grid semana */}
            {vistaCalCot==="semana"&&(
              <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)"}}>
                {diasSem.map((dia,i)=>{
                  const key=isoD(dia)
                  const esHoy=key===isoD(hoyC)
                  const cotsDelDia=cotsByDate[key]||[]
                  return(
                    <div key={key} style={{borderRight:i<6?"1px solid #f0ece4":"none",minHeight:120,padding:6}}>
                      {/* Header día */}
                      <div style={{textAlign:"center" as const,marginBottom:6}}>
                        <div style={{fontSize:10,color:"#9a9590",fontWeight:600,textTransform:"uppercase" as const}}>{DIAS[i]}</div>
                        <div style={{width:28,height:28,borderRadius:"50%",background:esHoy?"#0f172a":"transparent",color:esHoy?"#fff":"#1a1814",fontWeight:700,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",margin:"2px auto"}}>
                          {dia.getDate()}
                        </div>
                        {cotsDelDia.length>0&&(
                          <div style={{fontSize:9,background:"#2563eb",color:"#fff",borderRadius:8,padding:"0 5px",display:"inline-block",fontWeight:700}}>{cotsDelDia.length}</div>
                        )}
                      </div>
                      {/* Cotizaciones del día */}
                      <div style={{display:"flex",flexDirection:"column" as const,gap:3}}>
                        {cotsDelDia.slice(0,4).map((cot:any,ci:number)=>(
                          <div key={ci} onClick={()=>{setVista("detalle");/* handled by click */}}
                            title={cot.cliente_nombre+" — "+cot.estado}
                            style={{fontSize:9,padding:"2px 5px",borderRadius:4,background:ESTADO_COLOR[cot.estado]||"#64748b",color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const,cursor:"pointer",fontWeight:600}}>
                            {cot.folio||cot.cliente_nombre}
                          </div>
                        ))}
                        {cotsDelDia.length>4&&(
                          <div style={{fontSize:9,color:"#9a9590",textAlign:"center" as const}}>+{cotsDelDia.length-4} más</div>
                        )}
                      </div>
                    </div>
                  )
                })}
              </div>
            )}

            {/* Vista mes */}
            {vistaCalCot==="mes"&&(()=>{
              const primerDia=new Date(lunesBase.getFullYear(),lunesBase.getMonth(),1)
              const dowPrimer=primerDia.getDay()===0?6:primerDia.getDay()-1
              const diasMes=new Date(lunesBase.getFullYear(),lunesBase.getMonth()+1,0).getDate()
              const celdas=Array.from({length:Math.ceil((dowPrimer+diasMes)/7)*7},(_,i)=>{
                const d=new Date(primerDia);d.setDate(1-dowPrimer+i);return d
              })
              return(
                <div>
                  {/* Headers días */}
                  <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",borderBottom:"1px solid #f0ece4"}}>
                    {DIAS.map(d=><div key={d} style={{padding:"6px",textAlign:"center" as const,fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const}}>{d}</div>)}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)"}}>
                    {celdas.map((dia,i)=>{
                      const key=isoD(dia)
                      const esHoy=key===isoD(hoyC)
                      const esMesActual=dia.getMonth()===lunesBase.getMonth()
                      const cotsDelDia=cotsByDate[key]||[]
                      return(
                        <div key={i} style={{borderRight:i%7<6?"1px solid #f0ece4":"none",borderBottom:"1px solid #f0ece4",minHeight:80,padding:4,opacity:esMesActual?1:.35}}>
                          <div style={{width:22,height:22,borderRadius:"50%",background:esHoy?"#0f172a":"transparent",color:esHoy?"#fff":"#4a4640",fontWeight:esHoy?700:400,fontSize:11,display:"flex",alignItems:"center",justifyContent:"center",marginBottom:2}}>
                            {dia.getDate()}
                          </div>
                          {cotsDelDia.slice(0,2).map((cot:any,ci:number)=>(
                            <div key={ci}
                              title={cot.cliente_nombre+" — "+cot.estado}
                              style={{fontSize:8,padding:"1px 4px",borderRadius:3,background:ESTADO_COLOR[cot.estado]||"#64748b",color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const,marginBottom:1,cursor:"pointer",fontWeight:600}}>
                              {cot.folio||cot.cliente_nombre}
                            </div>
                          ))}
                          {cotsDelDia.length>2&&<div style={{fontSize:8,color:"#9a9590"}}>+{cotsDelDia.length-2}</div>}
                        </div>
                      )
                    })}
                  </div>
                </div>
              )
            })()}

            {/* Leyenda */}
            <div style={{padding:"10px 16px",borderTop:"1px solid #f0ece4",display:"flex",gap:10,flexWrap:"wrap" as const}}>
              {Object.entries({borrador:"Borrador",enviada:"Enviada",pendiente:"Pendiente",aceptada:"Aceptada",rechazada:"Rechazada",convertida:"Convertida"}).map(([k,l])=>(
                <div key={k} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:"#4a4640"}}>
                  <div style={{width:10,height:10,borderRadius:2,background:ESTADO_COLOR[k]}}/>
                  {l}
                </div>
              ))}
            </div>
          </div>
        )
      })()}

            {vista!=="calendario"&&cargando ? <div style={{padding:48,textAlign:"center" as const,color:"#9a9590"}}>Cargando...</div> : vista!=="calendario"&&(
        <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
          {cots.length===0 && <div style={{padding:48,textAlign:"center" as const,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12,color:"#9a9590"}}>
            <div style={{fontSize:32,opacity:.2,marginBottom:8}}>📋</div>
            <div>No hay cotizaciones</div>
          </div>}
          {cots.filter((cot:any)=>filtroVendCot==="todos"||(cot.vendedor||vendedorDesdeFolio(cot.folio||""))===filtroVendCot).map(cot=>{
            const dias = diasVigencia(cot.fecha_vigencia)
            const porVencer = dias !== null && dias >= 0 && dias <= 7 && (cot.estado==="enviada"||cot.estado==="pendiente")
            const vencida = dias !== null && dias < 0 && (cot.estado==="enviada"||cot.estado==="pendiente")
            return(
              <div key={cot.id} style={{background:"#fff",border:`1.5px solid ${porVencer?"#e8d4b8":vencida?"#e8b8b8":"#e8e5de"}`,borderRadius:10,padding:"12px 16px"}}>
                <div style={{display:"flex",alignItems:"flex-start",gap:12,flexWrap:"wrap" as const}}>
                  <div style={{flex:1,minWidth:200}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap" as const}}>
                      <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1a1814"}}>{cot.folio}</span>
                      <span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:ESTADO_BG[cot.estado]||"#f5f4f0",color:ESTADO_COL[cot.estado]||"#9a9590",fontWeight:700}}>{ESTADO_LABEL[cot.estado]||cot.estado}</span>
                      {porVencer&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#fdf5e8",color:"#92580a",fontWeight:700}}>⚠️ Vence en {dias}d</span>}
                      {vencida&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:6,background:"#fdf0f0",color:"#8b2e2e",fontWeight:700}}>⌛ Vencida</span>}
                    </div>
                    <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700}}>{cot.cliente_nombre||"Sin nombre"}</div>
                    <div style={{fontSize:11,color:"#9a9590",marginTop:2}}>
                      {cot.lugar_evento&&<span>📍 {cot.lugar_evento.slice(0,40)} · </span>}
                      {cot.fecha_evento&&<span>📅 {cot.fecha_evento} · </span>}
                      <span>{(cot.partidas||[]).length} artículos</span>
                      {cot.vendedor&&<span> · 👤 {cot.vendedor}</span>}
                    </div>
                  </div>
                  <div style={{textAlign:"right" as const,flexShrink:0}}>
                    {cot.fecha_evento&&<div style={{fontSize:11,color:"#9a9590",fontFamily:"monospace"}}>{cot.fecha_evento}</div>}
                    {cot.fecha_vigencia&&(()=>{
                      const d=diasVigencia(cot.fecha_vigencia)
                      return d!==null&&d<=7&&d>=0?<div style={{fontSize:10,color:d<=3?"#8b2e2e":"#92580a",fontWeight:700}}>Vence {d}d</div>:null
                    })()}
                  </div>
                </div>
                {/* Acciones */}
                <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap" as const}}>
                  <button onClick={async()=>{
                    // Always fetch fresh full data to avoid state leak
                    setCotActual({...COT_VACIA})
                    setEsNueva(false)
                    setVista("detalle")
                    try{
                      const r=await fetch(`/api/cotizaciones?id=${cot.id}`,{headers:{Authorization:`Bearer ${token}`}})
                      const data=await r.json()
                      if(data&&data.id) setCotActual(data)
                      else setCotActual(cot) // fallback
                    }catch{setCotActual(cot)}
                  }}
                    style={{padding:"4px 10px",borderRadius:6,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:11,fontFamily:"Epilogue,sans-serif"}}>
                    👁️ Ver
                  </button>
                  <button onClick={()=>{setCotActual(cot);setEsNueva(false);setPaso(1);setVista("form")}}
                    style={{padding:"4px 10px",borderRadius:6,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:11,fontFamily:"Epilogue,sans-serif"}}>
                    ✏️ Editar
                  </button>
                  <button onClick={()=>imprimirCotizacion(cot)}
                    style={{padding:"4px 10px",borderRadius:6,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:11,fontFamily:"Epilogue,sans-serif"}}>
                    🖨️ PDF
                  </button>
                  <button onClick={()=>enviarWhatsApp(cot)}
                    style={{padding:"4px 10px",borderRadius:6,border:"1px solid #25D366",background:"#25D366",color:"#fff",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                    📱 WhatsApp
                  </button>
                  {!(cot as any)._fromExcel&&(
                    <button onClick={()=>{
                      const url=window.location.origin+"/cot/"+cot.id
                      navigator.clipboard.writeText(url)
                        .then(()=>alert("✓ Link copiado: "+url))
                        .catch(()=>window.open(url,"_blank"))
                    }} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #4a2d6e",background:"#f5f0fc",color:"#4a2d6e",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                      🔗 Link cliente
                    </button>
                  )}
                  <button onClick={()=>duplicar(cot)}
                    style={{padding:"4px 10px",borderRadius:6,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:11,fontFamily:"Epilogue,sans-serif"}}>
                    📋 Duplicar
                  </button>
                  {!(cot as any)._fromExcel&&(
                    <button onClick={()=>eliminarConContrasena("cotizacion",cot.id,()=>{setCots(prev=>prev.filter((x:any)=>x.id!==cot.id));cargar()})}
                      style={{padding:"4px 10px",borderRadius:6,border:"1px solid #fca5a5",background:"#fef2f2",color:"#8b2e2e",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                      🗑️ Eliminar
                    </button>
                  )}
                  {(cot.estado==="aceptada")&&(
                    <button onClick={()=>convertirAContrato(cot)}
                      style={{padding:"4px 10px",borderRadius:6,border:"1px solid #4a2d6e",background:"#f5f0fc",color:"#4a2d6e",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                      📄 → Contrato
                    </button>
                  )}
                  {/* Cambiar estado */}
                  <select value={cot.estado} onChange={e=>{
                    if(e.target.value==="convertida"){
                      convertirAContrato(cot)
                    } else {
                      cambiarEstado((cot as any)._fromExcel?cot.id.replace("excel_",""):cot.id, e.target.value)
                    }
                  }} style={{padding:"3px 8px",borderRadius:6,border:"1px solid #e8e5de",fontFamily:"Epilogue,sans-serif",fontSize:10,cursor:"pointer",outline:"none"}}>
                    {ESTADOS_COT.map(e=><option key={e} value={e}>{ESTADO_LABEL[e]||e}</option>)}
                    <option value="convertida">✅ Convertir a contrato</option>
                  </select>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )

  // ── DETALLE
  if (vista === "detalle" && cotActual) {
    const cot = cotActual as Cotizacion
    return (
      <div>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
          <button onClick={()=>salirDelForm(()=>setVista("lista"))} style={{padding:"6px 12px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12}}>← Volver</button>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,flex:1}}>{cot.folio}</div>
          <span style={{fontSize:11,padding:"3px 10px",borderRadius:8,background:ESTADO_BG[cot.estado]||"#f5f4f0",color:ESTADO_COL[cot.estado]||"#9a9590",fontWeight:700}}>{ESTADO_LABEL[cot.estado]||cot.estado}</span>
          <button onClick={()=>imprimirCotizacion(cot)} style={{padding:"7px 14px",borderRadius:8,background:"#f5f4f0",color:"#1a1814",border:"1.5px solid #1a1814",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>🖨️ Ver PDF</button>
          <button onClick={()=>enviarWhatsApp(cot)} style={{padding:"7px 14px",borderRadius:8,background:"#25D366",color:"#fff",border:"none",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>📱 Enviar WA</button>
          {!(cot as any)._fromExcel&&(
            <button onClick={()=>{
              const url=window.location.origin+"/cot/"+cot.id
              const msg="Hola "+cot.cliente_nombre+"! Te compartimos tu cotización de Poliflor. Puedes revisarla y aprobarla desde este link: "+url
              const tel=(cot.cliente_tel||"").replace(/[^0-9]/g,"")
              window.open("https://wa.me/52"+tel+"?text="+encodeURIComponent(msg),"_blank")
            }} style={{padding:"7px 14px",borderRadius:8,background:"#4a2d6e",color:"#fff",border:"none",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
              🔗 Enviar link
            </button>
          )}
          <button onClick={()=>{setEsNueva(false);setPaso(1);setVista("form")}} style={{padding:"7px 14px",borderRadius:8,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontSize:12,fontFamily:"Epilogue,sans-serif"}}>✏️ Editar</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 320px",gap:14}}>
          <div style={{display:"flex",flexDirection:"column" as const,gap:12}}>
            {/* Cliente */}
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:10}}>Cliente y evento</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,fontSize:12}}>
                <div><span style={{color:"#9a9590"}}>Cliente:</span> <strong>{cot.cliente_nombre}</strong></div>
                <div><span style={{color:"#9a9590"}}>Tel:</span> {cot.cliente_tel||"—"}</div>
                <div><span style={{color:"#9a9590"}}>Lugar:</span> {cot.lugar_evento||"—"}</div>
                <div><span style={{color:"#9a9590"}}>Evento:</span> {cot.fecha_evento||"—"}</div>
                <div><span style={{color:"#9a9590"}}>Entrega:</span> {cot.fecha_entrega||"—"}</div>
                <div><span style={{color:"#9a9590"}}>Desmonte:</span> {cot.fecha_desmonte||"—"}</div>
                {cot.vendedor&&<div><span style={{color:"#9a9590"}}>Vendedor:</span> {cot.vendedor}</div>}
                {cot.fecha_vigencia&&<div><span style={{color:"#9a9590"}}>Vigencia:</span> {cot.fecha_vigencia}</div>}
              </div>
            </div>
            {/* Partidas */}
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
              <div style={{padding:"12px 16px",borderBottom:"1px solid #e8e5de",fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>Artículos ({(cot.partidas||[]).length})</div>
              <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:12}}>
                <thead><tr style={{background:"#fafaf8"}}>
                  {["#","Artículo","Cant.","P. Unitario","Subtotal","Notas"].map((h,i)=>(
                    <th key={i} style={{padding:"8px 12px",textAlign:i>=2?"center" as const:"left" as const,fontSize:10,fontWeight:700,color:"#9a9590",borderBottom:"1px solid #e8e5de"}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {(cot.partidas||[]).map((p,i)=>(
                    <tr key={i} style={{borderBottom:"1px solid #e8e5de",background:i%2===0?"#fff":"#fafaf8"}}>
                      <td style={{padding:"8px 12px",color:"#9a9590",fontSize:11}}>{i+1}</td>
                      <td style={{padding:"8px 12px",fontWeight:600}}>{p.nombre}</td>
                      <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace"}}>{p.cantidad}</td>
                      <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace"}}>{fmt(p.precio_unitario)}</td>
                      <td style={{padding:"8px 12px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{fmt(p.subtotal)}</td>
                      <td style={{padding:"8px 12px",fontSize:10,color:"#9a9590"}}>{p.notas||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {/* Condiciones */}
            {cot.condiciones&&(
              <div style={{background:"#fafaf8",border:"1px solid #e8e5de",borderRadius:10,padding:14}}>
                <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:6}}>Condiciones</div>
                <div style={{fontSize:11,color:"#4a4640",lineHeight:1.6}}>{cot.condiciones}</div>
              </div>
            )}
            {cot.notas_cliente&&(
              <div style={{background:"#fdf5e8",border:"1px solid #e8d4b8",borderRadius:10,padding:14}}>
                <div style={{fontSize:10,fontWeight:700,color:"#92580a",textTransform:"uppercase" as const,marginBottom:6}}>Notas para el cliente</div>
                <div style={{fontSize:11,color:"#4a4640"}}>{cot.notas_cliente}</div>
              </div>
            )}
          </div>
          {/* Panel totales */}
          <div style={{position:"sticky" as const,top:80,alignSelf:"flex-start" as const}}>
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
              <div style={{background:"#1a1814",padding:"12px 16px"}}>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800,color:"#fff"}}>{cot.folio}</div>
                <div style={{fontSize:11,color:"rgba(255,255,255,.6)",marginTop:2}}>{cot.cliente_nombre}</div>
              </div>
              <div style={{padding:16}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:8}}>
                  <span style={{color:"#9a9590"}}>Subtotal</span>
                  <span style={{fontFamily:"monospace",fontWeight:700}}>{fmt(cot.subtotal)}</span>
                </div>
                {cot.descuento_pct>0&&(
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:8,color:"#2d6a4f"}}>
                    <span>Descuento ({cot.descuento_pct}%)</span>
                    <span style={{fontFamily:"monospace",fontWeight:700}}>-{fmt(cot.descuento_monto)}</span>
                  </div>
                )}
                {cot.aplica_iva&&(
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:8,color:"#4a2d6e"}}>
                    <span>IVA (16%)</span>
                    <span style={{fontFamily:"monospace",fontWeight:700}}>{fmt(cot.iva_monto)}</span>
                  </div>
                )}
                <div style={{borderTop:"2px solid #1a1814",paddingTop:10,marginTop:4,display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800}}>TOTAL</span>
                  <span style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,color:"#1a1814"}}>{fmt(cot.total)}</span>
                </div>
                <div style={{marginTop:16,display:"flex",flexDirection:"column" as const,gap:6}}>
                  <button onClick={()=>enviarWhatsApp(cot)}
                    style={{padding:"9px",borderRadius:8,background:"#25D366",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                    📱 Enviar por WhatsApp
                  </button>
                  {cot.estado==="aceptada"&&(
                    <button onClick={()=>convertirAContrato(cot)}
                      style={{padding:"9px",borderRadius:8,background:"#4a2d6e",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                      📄 Convertir a contrato
                    </button>
                  )}
                  <div style={{display:"flex",gap:6}}>
                    {["aceptada","rechazada","pendiente"].map(e=>(
                      <button key={e} onClick={()=>{cambiarEstado(cot.id,e);setCotActual({...cot,estado:e})}}
                        style={{flex:1,padding:"6px",borderRadius:6,border:`1px solid ${ESTADO_COL[e]}`,background:cot.estado===e?ESTADO_BG[e]:"#fff",color:ESTADO_COL[e],cursor:"pointer",fontSize:10,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                        {e==="aceptada"?"✅":"e"==="rechazada"?"❌":"⏳"} {e}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    )
  }

  // ── FORMULARIO
  return (
    <div>
      {/* Header pasos */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 18px",marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
          <button onClick={()=>salirDelForm(()=>setVista("lista"))} style={{padding:"6px 12px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12}}>← Volver</button>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,flex:1}}>{esNueva?"Nueva cotización":`Editar ${cotActual.folio||""}`}</div>
          <button onClick={()=>guardar()} disabled={guardando}
            style={{padding:"7px 16px",borderRadius:8,background:guardando?"#9a9590":"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
            {guardando?"Guardando...":"💾 Guardar"}
          </button>
        </div>
        {/* Pasos */}
        <div style={{display:"flex",gap:4}}>
          {[["1","Datos del cliente","👤"],["2","Artículos","📦"],["3","Resumen","💰"]].map(([n,l,ic])=>(
            <button key={n} onClick={()=>setPaso(Number(n))}
              style={{flex:1,padding:"8px",borderRadius:8,border:`1.5px solid ${paso===Number(n)?"#1a1814":"#e8e5de"}`,background:paso===Number(n)?"#1a1814":"#fff",color:paso===Number(n)?"#fff":"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:paso===Number(n)?700:400}}>
              {ic} {l}
            </button>
          ))}
        </div>
      </div>

      {/* ─ PASO 1: DATOS ─ */}
      {paso===1&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:20}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700,marginBottom:16}}>👤 Datos del cliente y evento</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            {/* AUTOCOMPLETE CLIENTE */}
            <div style={{gridColumn:"1/-1",position:"relative" as const}}>
              <label style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,display:"block",marginBottom:4}}>Nombre del cliente *</label>
              <input
                value={busqCliente||cotActual.cliente_nombre||""}
                onChange={e=>{
                  const v=e.target.value
                  setBusqCliente(v)
                  setCotActual(p=>({...p,cliente_nombre:v}))
                  if(v.trim().length>=1){buscarClientes(v)}
                  else{setClientesSug([]);setMostrarSug(false)}
                }}
                onFocus={()=>{
                  const v=(busqCliente||cotActual.cliente_nombre||"").trim()
                  if(v.length>=1) buscarClientes(v)
                  else { // Show all clients on empty focus
                    const vistos=new Set<string>()
                    const all=contratosAll.filter((x:any)=>{
                      const k=(x.cliente||x.archivo||"").trim().toLowerCase()
                      if(!k||vistos.has(k))return false
                      vistos.add(k);return true
                    }).slice(0,10).map((x:any)=>({nombre:(x.cliente||x.archivo||"").trim(),tel:(x.tel||x.telefono||"").trim(),lugar:(x.lugar||"").trim(),email:""}))
                    setClientesSug(all)
                    setMostrarSug(all.length>0)
                  }
                }}
                onBlur={()=>setTimeout(()=>setMostrarSug(false),300)}
                placeholder="Escribe el nombre del cliente..."
                autoComplete="off"
                style={{width:"100%",padding:"10px 12px",border:`2px solid ${mostrarSug&&clientesSug.length>0?"#2563eb":"#e8e5de"}`,borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,outline:"none",boxSizing:"border-box" as const}}
              />
              {mostrarSug&&clientesSug.length>0&&<div style={{fontSize:9,color:"#2563eb",marginTop:3}}>↓ Selecciona un cliente de la lista</div>}
              {mostrarSug&&clientesSug.length>0&&(
                <div style={{position:"absolute" as const,top:"calc(100% + 8px)",left:0,right:0,background:"#fff",border:"2px solid #2563eb",borderRadius:12,boxShadow:"0 16px 48px rgba(0,0,0,.2)",zIndex:9999,overflow:"hidden"}}>
                  <div style={{padding:"6px 12px",background:"#1a1814",fontSize:9,fontWeight:700,color:"rgba(255,255,255,.5)",textTransform:"uppercase" as const,letterSpacing:".08em"}}>
                    {clientesSug.length} clientes encontrados
                  </div>
                  {clientesSug.map((cli,i)=>(
                    <div key={i} onMouseDown={(e)=>{e.preventDefault();seleccionarCliente(cli)}}
                      style={{padding:"10px 14px",cursor:"pointer",borderBottom:"1px solid #f5f4f0",display:"flex",alignItems:"center",gap:10}}
                      onMouseEnter={e=>(e.currentTarget.style.background="#f5f4f0")}
                      onMouseLeave={e=>(e.currentTarget.style.background="#fff")}>
                      <div style={{width:36,height:36,borderRadius:"50%",background:"#0f172a",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:14,flexShrink:0}}>
                        {(cli.nombre||"?").charAt(0).toUpperCase()}
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:13,marginBottom:2}}>{cli.nombre}</div>
                        <div style={{display:"flex",gap:8,flexWrap:"wrap" as const}}>
                          {cli.tel&&<span style={{fontSize:10,color:"#4a4640"}}>📞 {cli.tel}</span>}
                          {cli.lugar&&<span style={{fontSize:10,color:"#9a9590"}}>📍 {cli.lugar.slice(0,40)}</span>}
                        </div>
                      </div>
                      <div style={{fontSize:11,color:"#2563eb",fontWeight:700,flexShrink:0,background:"#edf3fa",padding:"3px 8px",borderRadius:6}}>Seleccionar</div>
                    </div>
                  ))}
                  <div onMouseDown={(e)=>{e.preventDefault();setCotActual(p=>({...p,cliente_nombre:busqCliente||""}));setMostrarSug(false);setClientesSug([])}}
                    style={{padding:"10px 14px",cursor:"pointer",fontSize:12,color:"#9a9590",display:"flex",alignItems:"center",gap:8,background:"#fafaf8",borderTop:"1px solid #e8e5de"}}
                    onMouseEnter={e=>(e.currentTarget.style.background="#f0ece4")}
                    onMouseLeave={e=>(e.currentTarget.style.background="#fafaf8")}>
                    <span style={{fontSize:18,color:"#2d6a4f"}}>＋</span>
                    <span>Nuevo cliente: <strong>"{busqCliente}"</strong></span>
                  </div>
                </div>
              )}
            </div>
            {/* REST OF FIELDS */}
            {([
              {l:"Teléfono",k:"cliente_tel",type:"tel"},
              {l:"Email",k:"cliente_email",type:"email"},
              {l:"Lugar del evento",k:"lugar_evento",type:"text",col:"1/-1"},
              {l:"Fecha del evento",k:"fecha_evento",type:"date"},
              {l:"Fecha de entrega",k:"fecha_entrega",type:"date"},
              {l:"Fecha de desmonte",k:"fecha_desmonte",type:"date"},
              {l:"Vigencia cotización",k:"fecha_vigencia",type:"date"},
            ] as {l:string,k:string,type:string,col?:string}[]).map(f=>(
              <div key={f.k} style={{gridColumn:f.col||"auto"}}>
                <label style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,display:"block",marginBottom:4}}>{f.l}</label>
                <input type={f.type} value={(cotActual as any)[f.k]||""} onChange={e=>setCotActual(p=>({...p,[f.k]:e.target.value}))}
                  style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}/>
              </div>
            ))}
            <div>
              <label style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,display:"block",marginBottom:4}}>Vendedor</label>
              <select value={cotActual.vendedor||""} onChange={e=>setCotActual({...cotActual,vendedor:e.target.value})}
                style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none"}}>
                <option value="">Sin asignar</option>
                {VENDEDORES.map(v=>(
                  <option key={v.nombre} value={v.nombre}>{v.nombre} ({v.prefijo})</option>
                ))}
              </select>
              {cotActual.vendedor&&(
                <div style={{fontSize:10,color:"#9a9590",marginTop:3}}>
                  Folio: <strong>{prefijoVendedor(cotActual.vendedor)}-XXXXXX</strong>
                </div>
              )}
            </div>
            <div style={{gridColumn:"1/-1"}}>
              <label style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,display:"block",marginBottom:4}}>Notas para el cliente</label>
              <textarea value={cotActual.notas_cliente||""} onChange={e=>setCotActual({...cotActual,notas_cliente:e.target.value})} rows={2}
                style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}/>
            </div>
            <div style={{gridColumn:"1/-1"}}>
              <label style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,display:"block",marginBottom:4}}>Condiciones</label>
              <textarea value={cotActual.condiciones||""} onChange={e=>setCotActual({...cotActual,condiciones:e.target.value})} rows={3}
                style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}/>
            </div>
          </div>
          <div style={{marginTop:16,display:"flex",justifyContent:"flex-end"}}>
            <button onClick={()=>setPaso(2)} style={{padding:"9px 24px",borderRadius:8,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
              Siguiente → Artículos
            </button>
          </div>
        </div>
      )}

      {/* ─ PASO 2: ARTÍCULOS ─ */}
      {paso===2&&(
        <div>
          {/* ── TABLA ESTILO EXCEL ── */}
          <div style={{background:"#fff",border:"1px solid #c7c7c7",borderRadius:8,overflow:"visible",boxShadow:"0 2px 8px rgba(0,0,0,.06)",marginBottom:8}}>
            {/* Header */}
            <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:"#0f172a"}}>
              {["","Cant.","Artículo / Descripción","P. Unitario","Descuento","Subtotal","Nota",""].map((h,hi)=>(
                <div key={hi} style={{padding:"9px 8px",fontSize:10,fontWeight:700,color:"#94a3b8",textAlign:hi>=2&&hi<=4?"center" as const:"left" as const,letterSpacing:".06em",textTransform:"uppercase" as const,borderRight:hi<6?"1px solid rgba(255,255,255,.07)":"none"}}>
                  {h}
                </div>
              ))}
            </div>

            {/* Filas */}
            {(cotActual.partidas||[]).length===0?(
              <div style={{padding:"32px",textAlign:"center" as const,color:"#9a9590"}}>
                <div style={{fontSize:28,opacity:.2,marginBottom:6}}>📦</div>
                <div style={{fontSize:12}}>Agrega artículos con el botón de abajo</div>
              </div>
            ):(
              (cotActual.partidas||[]).map((p:Partida,i:number)=>(
                <div key={i} style={{position:"relative" as const}}>
                  <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:i%2===0?"#fff":"#fafafa",borderTop:"1px solid #ebebeb"}}>
                    {/* Mover ▲▼ */}
                    <div style={{display:"flex",flexDirection:"column" as const,alignItems:"center",justifyContent:"center",gap:1,borderRight:"1px solid #ebebeb",minHeight:40,padding:"2px 0",width:22}}>
                      <button onClick={()=>{
                        if(i===0)return
                        const p=[...(cotActual.partidas||[])]
                        ;[p[i-1],p[i]]=[p[i],p[i-1]]
                        const tots=calcularTotales(p,cotActual.descuento_pct||0,cotActual.aplica_iva||false)
                        setCotActual((prev:any)=>({...prev,partidas:p,...tots}))
                      }} disabled={i===0}
                        style={{width:18,height:14,background:"none",border:"none",cursor:i===0?"default":"pointer",fontSize:9,color:i===0?"#d0cdc8":"#4a4640",padding:0,lineHeight:1,display:"flex",alignItems:"center",justifyContent:"center"}}>▲</button>
                      <button onClick={()=>{
                        const partidas=cotActual.partidas||[]
                        if(i>=partidas.length-1)return
                        const p=[...partidas]
                        ;[p[i],p[i+1]]=[p[i+1],p[i]]
                        const tots=calcularTotales(p,cotActual.descuento_pct||0,cotActual.aplica_iva||false)
                        setCotActual((prev:any)=>({...prev,partidas:p,...tots}))
                      }} disabled={i>=(cotActual.partidas||[]).length-1}
                        style={{width:18,height:14,background:"none",border:"none",cursor:i>=(cotActual.partidas||[]).length-1?"default":"pointer",fontSize:9,color:i>=(cotActual.partidas||[]).length-1?"#d0cdc8":"#4a4640",padding:0,lineHeight:1,display:"flex",alignItems:"center",justifyContent:"center"}}>▼</button>
                    </div>

                    {/* Cantidad */}
                    <div style={{borderRight:"1px solid #ebebeb",display:"flex",alignItems:"center",justifyContent:"center"}}>
                      <input type="number" min="1" value={p.cantidad}
                        onChange={e=>actualizarPartida(i,"cantidad",Math.max(1,parseInt(e.target.value)||1))}
                        style={{width:60,padding:"9px 4px",border:"none",background:"transparent",textAlign:"center" as const,fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#0f172a",outline:"none"}}/>
                    </div>

                    {/* Nombre con autocomplete inline */}
                    <div style={{position:"relative" as const,borderRight:"1px solid #ebebeb"}}>
                      <input
                        value={p.nombre}
                        onChange={async(e)=>{
                          const v=e.target.value
                          actualizarPartida(i,"nombre",v)
                          setActiveRow(i)
                          if(v.trim().length>=1){
                            const params=new URLSearchParams({busq:v.trim(),activo:"true"})
                            if(catFiltroArt!=="TODOS") params.append("categoria",catFiltroArt)
                            const res=await fetch("/api/catalogo?"+params,{headers:{Authorization:"Bearer "+token}})
                            const data=await res.json()
                            setRowSugeridos(Array.isArray(data)?data.slice(0,10):[])
                          } else {
                            setRowSugeridos([])
                          }
                        }}
                        onFocus={()=>setActiveRow(i)}
                        onBlur={()=>setTimeout(()=>{setActiveRow(null);setRowSugeridos([])},200)}
                        placeholder="Escribe el artículo..."
                        style={{width:"100%",padding:"9px 10px",border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:500,color:"#1a1814",outline:"none"}}
                      />
                      {/* Dropdown autocomplete inline */}
                      {activeRow===i&&rowSugeridos.length>0&&(
                        <div style={{position:"absolute" as const,top:"100%",left:0,right:0,background:"#fff",border:"2px solid #2563eb",borderRadius:"0 0 8px 8px",boxShadow:"0 8px 24px rgba(37,99,235,.15)",zIndex:500,maxHeight:240,overflowY:"auto" as const}}>
                          {rowSugeridos.map((a:any,ai:number)=>(
                            <div key={ai}
                              onMouseDown={(e)=>{
                                e.preventDefault()
                                // Update nombre + precio in one shot to avoid stale closure bug
                                const newPartidas=[...(cotActual.partidas||[])]
                                const pu=a.precio_renta||0
                                const cant=newPartidas[i]?.cantidad||1
                                newPartidas[i]={...newPartidas[i],nombre:a.nombre,precio_unitario:pu,subtotal:pu*cant}
                                const tots=calcularTotales(newPartidas,cotActual.descuento_pct||0,cotActual.aplica_iva||false)
                                setCotActual((prev:any)=>({...prev,partidas:newPartidas,...tots}))
                                setActiveRow(null)
                                setRowSugeridos([])
                              }}
                              style={{padding:"8px 12px",cursor:"pointer",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,background:"#fff"}}
                              onMouseEnter={e=>(e.currentTarget as HTMLElement).style.background="#eff6ff"}
                              onMouseLeave={e=>(e.currentTarget as HTMLElement).style.background="#fff"}>
                              <span style={{fontSize:16,flexShrink:0}}>{a.categoria==="MOBILIARIO"?"🪑":a.categoria==="FLORES"?"🌸":a.categoria==="MANTELERÍA"?"🏮":a.categoria==="VAJILLA"?"🍽️":a.categoria==="CARPAS"?"⛺":"📦"}</span>
                              <div style={{flex:1,minWidth:0}}>
                                <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</div>
                                <div style={{fontSize:10,color:"#9a9590"}}>{a.categoria}{a.subcategoria?" · "+a.subcategoria:""}</div>
                              </div>
                              {(a.precio_renta||0)>0&&(
                                <div style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1a3a5c",flexShrink:0}}>${(a.precio_renta||0).toLocaleString("es-MX")}</div>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Precio unitario */}
                    <div style={{borderRight:"1px solid #ebebeb",display:"flex",alignItems:"center",padding:"0 8px",gap:3}}>
                      <span style={{fontSize:10,color:"#9a9590",flexShrink:0}}>$</span>
                      <input type="number" min="0" step="1" value={p.precio_unitario||""}
                        onChange={e=>actualizarPartida(i,"precio_unitario",parseFloat(e.target.value)||0)}
                        placeholder="0"
                        style={{width:"100%",padding:"9px 2px",border:"none",background:"transparent",textAlign:"right" as const,fontFamily:"monospace",fontSize:13,fontWeight:600,color:"#1a3a5c",outline:"none"}}/>
                    </div>

                    {/* Descuento % por artículo */}
                    <div style={{borderRight:"1px solid #ebebeb",padding:"5px 6px",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column" as const,gap:2,background:i%2===0?"#fff":"#fafafa"}}>
                      <input type="checkbox" checked={!!(p as any).aplica_descuento}
                        onChange={e=>{
                          const checked=e.target.checked
                          const newP={...p,aplica_descuento:checked,descuento_pct_art:checked?((p as any).descuento_pct_art||0):0} as any
                          const partidas=(cotActual.partidas||[]).map((x:any,j:number)=>j===i?newP:x)
                          const tots=calcularTotales(partidas,cotActual.descuento_pct||0,cotActual.aplica_iva||false,descGlobalMonto)
                          setCotActual((prev:any)=>({...prev,partidas,...tots}))
                        }}
                        style={{width:14,height:14,cursor:"pointer",accentColor:"#8b2e2e"}}
                        title="Aplicar descuento % a este artículo"/>
                      {(p as any).aplica_descuento&&(
                        <div style={{display:"flex",alignItems:"center",gap:1}}>
                          <input type="number" min="0" max="100"
                            value={(p as any).descuento_pct_art||""}
                            placeholder="0"
                            onChange={e=>{
                              const pct=Math.min(100,parseFloat(e.target.value)||0)
                              const newP={...p,descuento_pct_art:pct} as any
                              const partidas=(cotActual.partidas||[]).map((x:any,j:number)=>j===i?newP:x)
                              const tots=calcularTotales(partidas,cotActual.descuento_pct||0,cotActual.aplica_iva||false,descGlobalMonto)
                              setCotActual((prev:any)=>({...prev,partidas,...tots}))
                            }}
                            style={{width:38,padding:"2px 3px",border:"1.5px solid #8b2e2e",borderRadius:4,textAlign:"center" as const,fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#8b2e2e",outline:"none"}}/>
                          <span style={{fontSize:9,color:"#8b2e2e",fontWeight:700}}>%</span>
                        </div>
                      )}
                    </div>

                    {/* Subtotal */}
                    <div style={{borderRight:"1px solid #ebebeb",padding:"9px 10px",display:"flex",alignItems:"center",justifyContent:"flex-end",background:i%2===0?"#eff6ff":"#e8f1fc"}}>
                      <div>
                        <span style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#1a3a5c"}}>${((p.precio_unitario||0)*(p.cantidad||0)).toLocaleString("es-MX")}</span>
                        {(p as any).aplica_descuento&&(p as any).descuento_pct_art>0&&(
                          <div style={{fontSize:9,color:"#8b2e2e",fontFamily:"monospace",fontWeight:700}}>
                            -{(p as any).descuento_pct_art}% (-${Math.round((p.precio_unitario||0)*(p.cantidad||0)*(p as any).descuento_pct_art/100).toLocaleString("es-MX")})
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Nota */}
                    <div style={{borderRight:"1px solid #ebebeb",display:"flex",alignItems:"center"}}>
                      <input value={p.notas||""} onChange={e=>actualizarPartida(i,"notas",e.target.value)}
                        placeholder="Nota..."
                        style={{width:"100%",padding:"9px 8px",border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:11,color:"#64748b",outline:"none"}}/>
                    </div>

                    {/* Eliminar */}
                    <div style={{display:"flex",alignItems:"center",justifyContent:"center"}}>
                      <button onClick={()=>{
                        const partidas=(cotActual.partidas||[]).filter((_:any,j:number)=>j!==i)
                        const tots=calcularTotales(partidas,cotActual.descuento_pct||0,cotActual.aplica_iva||false)
                        setCotActual((prev:any)=>({...prev,partidas,...tots}))
                      }} style={{width:24,height:24,borderRadius:4,background:"#fdf0f0",border:"none",cursor:"pointer",color:"#8b2e2e",fontSize:16,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                    </div>
                  </div>
                </div>
              ))
            )}

            {/* Fila totales */}
            {(cotActual.partidas||[]).length>0&&(
              <>
                <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:"#f8fafc",borderTop:"2px solid #0f172a"}}>
                  <div style={{gridColumn:"1/5",padding:"10px 12px",fontSize:11,fontWeight:600,color:"#4a4640"}}>
                    {(cotActual.partidas||[]).reduce((s:number,p:Partida)=>s+(p.cantidad||0),0)} piezas
                  </div>
                  <div style={{padding:"10px 8px",fontSize:11,color:"#9a9590",borderLeft:"1px solid #ddd",display:"flex",alignItems:"center",justifyContent:"flex-end"}}>Subtotal</div>
                  <div style={{padding:"10px 10px",fontFamily:"monospace",fontSize:14,fontWeight:800,color:"#1a3a5c",borderLeft:"1px solid #ddd",background:"#dbeafe",display:"flex",alignItems:"center",justifyContent:"flex-end"}}>
                    ${(cotActual.subtotal||0).toLocaleString("es-MX")}
                  </div>
                  <div/>
                </div>
                {(cotActual.descuento_pct||0)>0&&(
                  <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:"#fffbeb",borderTop:"1px solid #fde68a"}}>
                    <div style={{gridColumn:"1/6",padding:"8px 12px",fontSize:11,color:"#92580a",fontWeight:600}}>Descuento {cotActual.descuento_pct}%</div>
                    <div style={{padding:"8px 10px",fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#92580a",borderLeft:"1px solid #fde68a",textAlign:"right" as const}}>
                      −${(cotActual.descuento_monto||0).toLocaleString("es-MX")}
                    </div>
                    <div/>
                  </div>
                )}
                {cotActual.aplica_iva&&(
                  <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:"#f0fdf4",borderTop:"1px solid #bbf7d0"}}>
                    <div style={{gridColumn:"1/6",padding:"8px 12px",fontSize:11,color:"#2d6a4f",fontWeight:600}}>IVA 16%</div>
                    <div style={{padding:"8px 10px",fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#2d6a4f",borderLeft:"1px solid #bbf7d0",textAlign:"right" as const}}>
                      +${(cotActual.iva_monto||0).toLocaleString("es-MX")}
                    </div>
                    <div/>
                  </div>
                )}
                <div style={{display:"grid",gridTemplateColumns:"22px 72px 1fr 110px 80px 110px 120px 32px",background:"#0f172a",borderTop:"2px solid #0f172a"}}>
                  <div style={{gridColumn:"1/6",padding:"12px",fontSize:13,fontWeight:800,color:"#fff"}}>TOTAL</div>
                  <div style={{padding:"12px 10px",fontFamily:"monospace",fontSize:16,fontWeight:800,color:"#60a5fa",borderLeft:"1px solid rgba(255,255,255,.15)",textAlign:"right" as const}}>
                    ${(cotActual.total||0).toLocaleString("es-MX")}
                  </div>
                  <div/>
                </div>
              </>
            )}
          </div>

          {/* Botón agregar fila */}
          <button onClick={()=>{
            const nueva:Partida={articulo_id:"",nombre:"",cantidad:1,precio_unitario:0,subtotal:0,notas:""}
            const partidas=[...(cotActual.partidas||[]),nueva]
            const tots=calcularTotales(partidas,cotActual.descuento_pct||0,cotActual.aplica_iva||false)
            setCotActual((prev:any)=>({...prev,partidas,...tots}))
          }} style={{width:"100%",padding:"12px",borderRadius:8,background:"#fff",border:"2px dashed #2563eb",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:13,color:"#2563eb",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",gap:8,marginBottom:12}}>
            <span style={{fontSize:20,lineHeight:1}}>+</span> Agregar artículo
          </button>

          {/* Opciones descuento + IVA */}
          <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap" as const}}>
            {/* ── DESCUENTO GLOBAL ── */}
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
              <div style={{fontSize:10,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".05em",marginBottom:10}}>🏷️ Descuento global</div>
              <div style={{display:"flex",gap:6,marginBottom:10}}>
                <button onClick={()=>{
                    setDescTipo("pct")
                    setDescGlobalMonto(0)
                    const tots=calcularTotales(cotActual.partidas||[],cotActual.descuento_pct||0,cotActual.aplica_iva||false,0)
                    setCotActual((prev:any)=>({...prev,...tots}))
                  }}
                  style={{flex:1,padding:"7px",borderRadius:7,border:`2px solid ${descTipo==="pct"?"#1a3a5c":"#e8e5de"}`,background:descTipo==="pct"?"#1a3a5c":"#fff",color:descTipo==="pct"?"#fff":"#4a4640",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  % Porcentaje
                </button>
                <button onClick={()=>{
                    setDescTipo("monto")
                    const tots=calcularTotales(cotActual.partidas||[],0,cotActual.aplica_iva||false,descGlobalMonto)
                    setCotActual((prev:any)=>({...prev,...tots,descuento_pct:0}))
                  }}
                  style={{flex:1,padding:"7px",borderRadius:7,border:`2px solid ${descTipo==="monto"?"#1a3a5c":"#e8e5de"}`,background:descTipo==="monto"?"#1a3a5c":"#fff",color:descTipo==="monto"?"#fff":"#4a4640",fontWeight:700,fontSize:12,cursor:"pointer"}}>
                  $ Monto fijo
                </button>
              </div>
              {descTipo==="pct"?(
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <input type="number" min="0" max="100"
                      value={cotActual.descuento_pct||""}
                      placeholder="0"
                      onChange={e=>{
                        const pct=Math.min(100,parseFloat(e.target.value)||0)
                        const tots=calcularTotales(cotActual.partidas||[],pct,cotActual.aplica_iva||false,0)
                        setCotActual((prev:any)=>({...prev,...tots,descuento_pct:pct}))
                      }}
                      style={{flex:1,padding:"10px 12px",border:"2px solid #e8e5de",borderRadius:8,fontFamily:"monospace",fontSize:18,fontWeight:800,outline:"none",textAlign:"center" as const,boxSizing:"border-box" as const}}
                      onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                      onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    <span style={{fontSize:18,fontWeight:800,color:"#4a4640"}}>%</span>
                  </div>
                  {(cotActual.descuento_pct||0)>0&&(
                    <div style={{marginTop:6,padding:"6px 10px",background:"#f0fdf4",borderRadius:7,fontSize:12,color:"#2d6a4f",fontWeight:700,display:"flex",justifyContent:"space-between" as const}}>
                      <span>= descuento</span>
                      <span style={{fontFamily:"monospace"}}>-${Math.round((cotActual.subtotal||0)*(cotActual.descuento_pct||0)/100).toLocaleString("es-MX")}</span>
                    </div>
                  )}
                </div>
              ):(
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:18,fontWeight:800,color:"#4a4640"}}>$</span>
                    <input type="number" min="0"
                      value={descGlobalMonto||""}
                      placeholder="0"
                      onChange={e=>{
                        const monto=parseFloat(e.target.value)||0
                        setDescGlobalMonto(monto)
                        const tots=calcularTotales(cotActual.partidas||[],0,cotActual.aplica_iva||false,monto)
                        setCotActual((prev:any)=>({...prev,...tots,descuento_pct:0}))
                      }}
                      style={{flex:1,padding:"10px 12px",border:"2px solid #e8e5de",borderRadius:8,fontFamily:"monospace",fontSize:18,fontWeight:800,outline:"none",textAlign:"center" as const,boxSizing:"border-box" as const}}
                      onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                      onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                  </div>
                  {descGlobalMonto>0&&(
                    <div style={{marginTop:6,padding:"6px 10px",background:"#f0fdf4",borderRadius:7,fontSize:12,color:"#2d6a4f",fontWeight:700,display:"flex",justifyContent:"space-between" as const}}>
                      <span>= descuento</span>
                      <span style={{fontFamily:"monospace"}}>-${descGlobalMonto.toLocaleString("es-MX")}</span>
                    </div>
                  )}
                </div>
              )}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8,background:"#fff",border:"1px solid #e8e5de",borderRadius:8,padding:"8px 12px"}}>
              <label style={{fontSize:12,color:"#4a4640",fontWeight:600}}>IVA 16%</label>
              <input type="checkbox" checked={cotActual.aplica_iva||false}
                onChange={e=>{
                  const iva=e.target.checked
                  const tots=calcularTotales(cotActual.partidas||[],cotActual.descuento_pct||0,iva)
                  setCotActual(prev=>({...prev,...tots,aplica_iva:iva}))
                }}
                style={{width:18,height:18,cursor:"pointer"}}/>
            </div>
          </div>

          {/* Navegación */}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setPaso(1)} style={{flex:1,padding:"10px",borderRadius:8,background:"#f5f4f0",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:600}}>← Datos</button>
            <button onClick={()=>setPaso(3)} style={{flex:2,padding:"10px",borderRadius:8,background:"#0f172a",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>Resumen →</button>
          </div>
        </div>
      )}

      {/* ─ PASO 3: RESUMEN ─ */}
      {paso===3&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 300px",gap:14}}>
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:20}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700,marginBottom:16}}>✅ Resumen final</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16,fontSize:12}}>
              <div style={{background:"#fafaf8",borderRadius:8,padding:"10px 12px"}}><div style={{color:"#9a9590",fontSize:10,marginBottom:2}}>CLIENTE</div><strong>{cotActual.cliente_nombre}</strong></div>
              <div style={{background:"#fafaf8",borderRadius:8,padding:"10px 12px"}}><div style={{color:"#9a9590",fontSize:10,marginBottom:2}}>TELÉFONO</div>{cotActual.cliente_tel||"—"}</div>
              <div style={{background:"#fafaf8",borderRadius:8,padding:"10px 12px"}}><div style={{color:"#9a9590",fontSize:10,marginBottom:2}}>EVENTO</div>{cotActual.fecha_evento||"—"}</div>
              <div style={{background:"#fafaf8",borderRadius:8,padding:"10px 12px"}}><div style={{color:"#9a9590",fontSize:10,marginBottom:2}}>LUGAR</div>{cotActual.lugar_evento||"—"}</div>
            </div>
            <div style={{marginBottom:16}}>
              <div style={{fontWeight:700,fontSize:12,marginBottom:8}}>Artículos ({(cotActual.partidas||[]).length})</div>
              {(cotActual.partidas||[]).map((p,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f5f4f0",fontSize:12}}>
                  <span>{p.cantidad}x {p.nombre}</span>
                  <span style={{fontFamily:"monospace",fontWeight:700}}>${p.subtotal.toLocaleString()}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#fafaf8",borderRadius:10,padding:14}}>
              <div style={{fontSize:10,fontWeight:700,color:"#9a9590",marginBottom:8}}>CONDICIONES</div>
              <div style={{fontSize:11,color:"#4a4640",lineHeight:1.6}}>{cotActual.condiciones}</div>
            </div>
          </div>
          <div style={{position:"sticky" as const,top:80,alignSelf:"flex-start" as const}}>
            <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:16}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:12}}>💰 Total</div>
              <div style={{fontSize:13,display:"flex",justifyContent:"space-between",marginBottom:6}}><span style={{color:"#9a9590"}}>Subtotal</span><span style={{fontFamily:"monospace",fontWeight:700}}>${(cotActual.subtotal||0).toLocaleString()}</span></div>
              {(cotActual.descuento_pct||0)>0&&<div style={{fontSize:13,display:"flex",justifyContent:"space-between",marginBottom:6,color:"#2d6a4f"}}><span>Descuento {cotActual.descuento_pct}%</span><span style={{fontFamily:"monospace",fontWeight:700}}>-${(cotActual.descuento_monto||0).toLocaleString()}</span></div>}
              {cotActual.aplica_iva&&<div style={{fontSize:13,display:"flex",justifyContent:"space-between",marginBottom:6,color:"#4a2d6e"}}><span>IVA 16%</span><span style={{fontFamily:"monospace",fontWeight:700}}>${(cotActual.iva_monto||0).toLocaleString()}</span></div>}
              <div style={{borderTop:"2px solid #1a1814",paddingTop:10,display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                <span style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800}}>TOTAL</span>
                <span style={{fontFamily:"Playfair Display,serif",fontSize:24,fontWeight:800}}>${(cotActual.total||0).toLocaleString()}</span>
              </div>
              <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
                <button onClick={()=>guardar("borrador")} disabled={guardando}
                  style={{padding:"9px",borderRadius:8,background:"#f5f4f0",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                  💾 Guardar borrador
                </button>
                <button onClick={()=>guardar("enviada")} disabled={guardando}
                  style={{padding:"9px",borderRadius:8,background:"#1a3a5c",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:700}}>
                  📤 Guardar y marcar enviada
                </button>
              </div>
              <button onClick={()=>setPaso(2)} style={{width:"100%",padding:"7px",borderRadius:8,background:"#fff",border:"1px solid #e8e5de",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,marginTop:8}}>← Editar artículos</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}


// ─── CONTRATOS CONFIRMADOS ─────────────────────────────────────────
function ContratosConfirmadosSection({token,contratos,onActualizar,isMobile,vendedorActual}:{token:string,contratos:Contrato[],onActualizar?:(id:string,updates:any)=>void,isMobile?:boolean,vendedorActual?:string,esAdmin?:boolean}){
  const [busq,setBusq]=useState("")
  const [filtroVendedor,setFiltroVendedor]=useState("todos")
  const [filtroMes,setFiltroMes]=useState("")
  const [soloMios,setSoloMios]=useState(!!vendedorActual)
  const [selContrato,setSelContrato]=useState<any|null>(null)
  const [modoEdicion,setModoEdicion]=useState(false)
  const [editForm,setEditForm]=useState<any>({})
  const [editArtsContrato,setEditArtsContrato]=useState<any[]>([])
  const [editDescTipo,setEditDescTipo]=useState<"pct"|"monto">("pct")
  const [editDescPct,setEditDescPct]=useState(0)
  const [editDescMonto,setEditDescMonto]=useState(0)
  const [editAplIva,setEditAplIva]=useState(false)
  const [busqArtCont,setBusqArtCont]=useState("")
  const [sugsArtCont,setSugsArtCont]=useState<any[]>([])
  const [guardandoEdit,setGuardandoEdit]=useState(false)
  const [showPagoForm,setShowPagoForm]=useState<string|null>(null)
  const [pagoMonto,setPagoMonto]=useState("")
  const [pagoMetodo,setPagoMetodo]=useState("efectivo")
  const [pagoNota,setPagoNota]=useState("")
  const [splitStatus,setSplitStatus]=useState<Record<string,boolean>>({})
  const [generandoSplit,setGenerandoSplit]=useState<string|null>(null)
  const [splitsDelContrato,setSplitsDelContrato]=useState<any[]>([])
  const [cargandoSplits,setCargandoSplits]=useState(false)
  const [mostrarSplits,setMostrarSplits]=useState(false)

  const eliminarContratoConPwd = async (id: string, onSuccess: ()=>void) => {
    const pwd = window.prompt("Ingresa la contraseña para eliminar este contrato:")
    if(pwd === null) return
    if(pwd !== DELETE_PWD){
      alert("❌ Contraseña incorrecta. No se eliminó el contrato.")
      return
    }
    const res = await fetch(`/api/contratos?id=${id}`, { method: "DELETE", headers: { Authorization: `Bearer ${token}` }})
    if(res.ok) onSuccess()
    else alert("Error al eliminar. Intenta de nuevo.")
  }

  const generarSplits=async(contrato:any)=>{
    setGenerandoSplit(contrato.id)
    // Eliminar splits anteriores EXCEPTO proveedores
    // Primero obtenemos los splits existentes
    const existingR=await fetch(`/api/splits?contrato_id=${contrato.id}`,{headers:{Authorization:`Bearer ${token}`}})
    const existing=await existingR.json()
    if(Array.isArray(existing)){
      // Solo borrar los que NO son proveedor
      const toDelete=existing.filter((s:any)=>s.tipo!=="proveedor")
      await Promise.all(toDelete.map((s:any)=>
        fetch(`/api/splits?id=${s.id}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}})
      ))
    }
    // Generar nuevos
    const nuevos=generarSplitsDesdeContrato(contrato)
    console.log("Splits a insertar:", JSON.stringify(nuevos.map(s=>({tipo:s.tipo,fecha_evento:s.fecha_evento,arts:s.articulos?.length}))))
    const res=await fetch("/api/splits",{
      method:"POST",
      headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
      body:JSON.stringify(nuevos)
    })
    const splitData=await res.json()
    if(res.ok){
      setSplitStatus(prev=>({...prev,[contrato.id]:true}))
      cargarSplitsDeContrato(contrato.id)
      setMostrarSplits(true)
    } else {
      const errMsg=splitData?.error||JSON.stringify(splitData)||"Error desconocido"
      alert("❌ Error al generar splits: "+errMsg)
      console.error("generarSplits error:", splitData)
    }
    setGenerandoSplit(null)
  }

  const cargarSplitsDeContrato=async(contratoId:string)=>{
    setCargandoSplits(true)
    const r=await fetch(`/api/splits?contrato_id=${contratoId}`,
      {headers:{Authorization:`Bearer ${token}`}})
    const data=await r.json()
    setSplitsDelContrato(Array.isArray(data)?data:[])
    setCargandoSplits(false)
  }

  // Check which contratos already have splits
  useEffect(()=>{
    if(typeof window==="undefined")return
    fetch("/api/splits",{headers:{Authorization:`Bearer ${token}`}})
      .then(r=>r.json())
      .then(data=>{
        if(!Array.isArray(data))return
        const map:Record<string,boolean>={}
        data.forEach((s:any)=>{if(s.contrato_id)map[s.contrato_id]=true})
        setSplitStatus(map)
      }).catch(()=>{})
  },[token])

  const abrirEdicionContrato=(ct:any)=>{
    setEditForm({
      cliente:ct.cliente||ct.archivo||"",
      folio:ct.folio||"",
      lugar:ct.lugar||"",
      tel:ct.tel||ct.telefono||"",
      fecha_evento:ct.fecha_evento||"",
      fecha_entrega:ct.fecha_entrega||"",
      fecha_desmonte:ct.fecha_desmonte||"",
      vendedor:ct.vendedor||"",
      total:ct.total||0,
      notas:ct.notas||"",
    })
    setEditArtsContrato(JSON.parse(JSON.stringify(ct.articulos||[])))
    setEditDescTipo("pct")
    setEditDescPct(ct.descuento_pct||0)
    setEditDescMonto(ct.descuento_monto_global||0)
    setEditAplIva(ct.aplica_iva||false)
    setBusqArtCont("")
    setSugsArtCont([])
    setModoEdicion(true)
  }

  // Recalcula el total del contrato mientras se edita
  const calcTotalContrato=(arts:any[],descPct:number,descMnt:number,aplIva:boolean)=>{
    const subtotal=arts.reduce((s:number,a:any)=>{
      const bruto=(a.pu||0)*(a.cantidad||0)
      const d=a.aplica_descuento?Math.round(bruto*(a.descuento_pct_art||0)/100):0
      return s+bruto-d
    },0)
    const descGlobal=editDescTipo==="pct"?Math.round(subtotal*descPct/100):descMnt
    const base=Math.max(0,subtotal-descGlobal)
    const iva=aplIva?Math.round(base*0.16):0
    return {subtotal,descGlobal,base,iva,total:base+iva}
  }

  const guardarEdicionContrato=async()=>{
    if(!selContrato)return
    setGuardandoEdit(true)
    const res=await fetch(`/api/contratos?id=${selContrato.id}`,{
      method:"PATCH",
      headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
      body:JSON.stringify({
        cliente:editForm.cliente,
        archivo:editForm.folio||editForm.cliente,
        folio:editForm.folio,
        lugar:editForm.lugar,
        telefono:editForm.tel,
        tel:editForm.tel,
        fecha_evento:editForm.fecha_evento,
        fecha_entrega:editForm.fecha_entrega,
        fecha_desmonte:editForm.fecha_desmonte,
        vendedor:editForm.vendedor,
        total:parseFloat(String(editForm.total))||calcTotalContrato(editArtsContrato,editDescPct,editDescMonto,editAplIva).total,
        descuento_pct:editDescPct,
        descuento_monto_global:editDescMonto,
        aplica_iva:editAplIva,
        notas:editForm.notas,
        articulos:editArtsContrato,
      })
    })
    if(res.ok){
      const tots=calcTotalContrato(editArtsContrato,editDescPct,editDescMonto,editAplIva)
      const updated={...selContrato,...editForm,tel:editForm.tel,telefono:editForm.tel,total:parseFloat(String(editForm.total))||tots.total,articulos:editArtsContrato,descuento_pct:editDescPct,descuento_monto_global:editDescMonto,aplica_iva:editAplIva}
      setSelContrato(updated)
      if(onActualizar) onActualizar(String(selContrato.id),editForm)
      setModoEdicion(false)
    }
    setGuardandoEdit(false)
  }

  const imprimirContrato=(x:any)=>{
    const logoSrc=typeof window!=="undefined"?(localStorage.getItem("pf_logo")||"/logo.png"):"/logo.png"
    const totalArts=(x.articulos||[]).reduce((s:number,a:any)=>s+(a.cantidad||0),0)
    const filas=(x.articulos||[]).map((a:any,i:number)=>{
      const bruto=(a.pu||a.precio_unitario||0)*(a.cantidad||0)
      const descArt=a.aplica_descuento?Math.round(bruto*(a.descuento_pct_art||0)/100):0
      const neto=bruto-descArt
      return `
      <tr style="border-bottom:1px solid #f0ece4;background:${i%2===0?"#fff":"#fafaf8"}">
        <td style="padding:8px 10px;text-align:center;font-weight:700;font-size:13px;color:#1a3a5c">${a.cantidad||0}</td>
        <td style="padding:8px 10px;font-weight:600;font-size:13px">${a.nombre||"—"}</td>
        <td style="padding:8px 10px;text-align:right;font-family:monospace;font-size:13px">$${(a.pu||a.precio_unitario||0).toLocaleString("es-MX")}</td>
        <td style="padding:8px 10px;text-align:right;font-family:monospace;font-weight:700;font-size:13px;color:#1a3a5c">
          $${neto.toLocaleString("es-MX")}
          ${descArt>0?`<div style="font-size:9px;color:#8b2e2e">-${a.descuento_pct_art}%</div>`:""}
        </td>
      </tr>`
    }).join("")
    const saldo=(x.total||0)-(x.cobrado||0)
    const html=`<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<title>Contrato ${x.folio||x.archivo||""}</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet"/>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:Inter,Arial,sans-serif;font-size:13px;color:#1a1814;background:#fff}
  @media print{.no-print{display:none!important}@page{margin:12mm 15mm;size:A4 portrait}body,body *{filter:grayscale(100%)!important;-webkit-filter:grayscale(100%)!important}img{filter:grayscale(100%) brightness(0.9)!important}}
  .page{max-width:820px;margin:0 auto;padding:32px}
  table{border-collapse:collapse;width:100%}
</style>
</head><body>
<div class="page">


  <!-- Header -->
  <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px;padding-bottom:20px;border-bottom:3px solid #1a1814">
    <div>
      <img src="${logoSrc}" alt="Poliflor" style="height:60px;width:auto;object-fit:contain;margin-bottom:6px" onerror="this.style.display='none'">
      <div style="font-size:11px;color:#9a9590;margin-top:4px">Renta de Mobiliario para Eventos</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:11px;color:#9a9590;text-transform:uppercase;letter-spacing:.08em">Contrato</div>
      <div style="font-size:22px;font-weight:800;color:#1a1814">${x.folio||x.archivo||"—"}</div>
      ${x.vendedor?`<div style="font-size:11px;color:#9a9590;margin-top:4px">Vendedor: <strong>${x.vendedor}</strong></div>`:""}
    </div>
  </div>

  <!-- Cliente y evento -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px">
    <div style="background:#f8f6f2;border-radius:10px;padding:16px">
      <div style="font-size:10px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Cliente</div>
      <div style="font-size:16px;font-weight:800;margin-bottom:6px">${x.cliente||x.archivo||"—"}</div>
      ${x.tel||x.telefono?`<div style="font-size:12px;color:#4a4640">📞 ${x.tel||x.telefono}</div>`:""}
    </div>
    <div style="background:#f8f6f2;border-radius:10px;padding:16px">
      <div style="font-size:10px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Evento</div>
      ${x.fecha_evento?`<div style="font-size:13px;font-weight:700;margin-bottom:4px">📅 ${x.fecha_evento}</div>`:""}
      ${x.lugar?`<div style="font-size:12px;color:#4a4640">📍 ${x.lugar}</div>`:""}
      ${x.fecha_entrega?`<div style="font-size:11px;color:#9a9590;margin-top:6px">Entrega: ${x.fecha_entrega?new Date(x.fecha_entrega+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"}):""}${x.fecha_desmonte?" · Desmonte: "+new Date(x.fecha_desmonte+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"}):""}</div>`:""}
    </div>
  </div>

  ${x.cuenta_bancaria?`<div style="margin:16px 0;padding:12px 16px;background:#f0fdf4;border:1px solid #b7deca;border-radius:8px;font-size:12px"><div style="font-weight:700;color:#2d6a4f;margin-bottom:6px">💳 Datos para transferencia</div><div><strong>${x.cuenta_bancaria.banco}</strong> · ${x.cuenta_bancaria.titular}</div>${x.cuenta_bancaria.cuenta?`<div>N° Cuenta: <strong>${x.cuenta_bancaria.cuenta}</strong></div>`:""}${x.cuenta_bancaria.clabe?`<div>CLABE: <strong>${x.cuenta_bancaria.clabe}</strong></div>`:""}</div>`:""}
  <!-- Artículos -->
  <table style="margin-bottom:24px">
    <thead>
      <tr style="background:#1a1814;color:#fff">
        <th style="padding:10px;text-align:center;font-size:11px;width:70px">Cant.</th>
        <th style="padding:10px;text-align:left;font-size:11px">Artículo / Descripción</th>
        <th style="padding:10px;text-align:right;font-size:11px;width:100px">P. Unit.</th>
        <th style="padding:10px;text-align:right;font-size:11px;width:110px">Importe</th>
      </tr>
    </thead>
    <tbody>${filas}</tbody>
    <tfoot>
      <tr style="background:#f8f6f2">
        <td colspan="2" style="padding:10px;font-weight:700;font-size:13px">Total artículos: ${totalArts} piezas</td>
        <td colspan="3" style="padding:10px;text-align:right;font-size:12px">
          ${(()=>{
            const arts=x.articulos||[]
            const subtotalBruto=arts.reduce((s,a)=>{
              const bruto=(a.pu||0)*(a.cantidad||0)
              const descArt=a.aplica_descuento?Math.round(bruto*(a.descuento_pct_art||0)/100):0
              return s+bruto-descArt
            },0)
            const descGlobal=x.descuento_pct>0?Math.round(subtotalBruto*x.descuento_pct/100):(x.descuento_monto_global||0)
            const base=Math.max(0,subtotalBruto-descGlobal)
            const iva=x.aplica_iva?Math.round(base*0.16):0
            const total=x.total||base+iva
            const rows=[]
            rows.push(`<div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:#9a9590">Subtotal</span><span style="font-family:monospace;font-weight:600">$${subtotalBruto.toLocaleString("es-MX")}</span></div>`)
            if(descGlobal>0) rows.push(`<div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:#2d6a4f">Descuento aplicado</span><span style="font-family:monospace;font-weight:600;color:#2d6a4f">-$${descGlobal.toLocaleString("es-MX")}</span></div>`)
            if(iva>0) rows.push(`<div style="display:flex;justify-content:space-between;margin-bottom:3px"><span style="color:#4a2d6e">IVA 16%</span><span style="font-family:monospace;font-weight:600;color:#4a2d6e">+$${iva.toLocaleString("es-MX")}</span></div>`)
            rows.push(`<div style="display:flex;justify-content:space-between;border-top:2px solid #1a1814;padding-top:6px;margin-top:4px"><span style="font-weight:800;font-size:14px">Total a pagar</span><span style="font-family:monospace;font-weight:800;font-size:16px;color:#1a1814">$${total.toLocaleString("es-MX")}</span></div>`)
            return rows.join("")
          })()}
        </td>
      </tr>
    </tfoot>
  </table>

  <!-- Pagos -->
  ${(x.pagos||[]).length>0?`
  <div style="margin-bottom:24px;background:#f0fdf4;border-radius:10px;padding:16px">
    <div style="font-size:10px;font-weight:700;color:#2d6a4f;text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px">Historial de pagos</div>
    ${(x.pagos||[]).map((p:any)=>`
      <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #dcfce7;font-size:12px">
        <span>${p.fecha} — ${p.metodo||"efectivo"}${p.nota?" ("+p.nota+")":""}</span>
        <span style="font-family:monospace;font-weight:700;color:#2d6a4f">$${(p.monto||0).toLocaleString("es-MX")}</span>
      </div>`).join("")}
    <div style="display:flex;justify-content:space-between;padding:8px 0;margin-top:4px;font-weight:700">
      <span>Cobrado</span>
      <span style="font-family:monospace;color:#2d6a4f">$${(x.cobrado||0).toLocaleString("es-MX")}</span>
    </div>
    ${saldo>0?`<div style="display:flex;justify-content:space-between;padding:8px 0;font-weight:700">
      <span style="color:#92580a">Saldo pendiente</span>
      <span style="font-family:monospace;color:#92580a">$${saldo.toLocaleString("es-MX")}</span>
    </div>`:""}
  </div>`:""}

  <!-- Firmas -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:40px;padding-top:20px;border-top:1px solid #e8e5de">
    <div style="text-align:center">
      <div style="border-top:1px solid #1a1814;padding-top:8px;margin-top:50px;font-size:11px;color:#9a9590">Firma del cliente</div>
      <div style="font-size:11px;color:#9a9590;margin-top:4px">${x.cliente||""}</div>
    </div>
    <div style="text-align:center">
      <div style="border-top:1px solid #1a1814;padding-top:8px;margin-top:50px;font-size:11px;color:#9a9590">Autorizado por Poliflor</div>
      <div style="font-size:11px;color:#9a9590;margin-top:4px">${x.vendedor||"Poliflor"}</div>
    </div>
  </div>
</div></body></html>`
    // PDF: CONTRATO_FOLIO_CLIENTE_YYYY-MM-DD
    const slugC=(s:string,n=25)=>s.slice(0,n).trim().replace(/\s+/g,"-").replace(/[^a-zA-Z0-9-]/g,"").toUpperCase()
    const nombreC=`CONTRATO_${slugC(x.folio||x.archivo||"CONT",15)}_${slugC(x.cliente||x.archivo||"CLIENTE")}_${x.fecha_evento||new Date().toISOString().slice(0,10)}`
    const btnsC=`
<div id="pflbtns" style="position:fixed;top:12px;right:12px;z-index:9999;background:#fff;padding:14px 16px;border-radius:12px;box-shadow:0 6px 28px rgba(0,0,0,.22);min-width:220px;font-family:Arial,sans-serif">
  <div style="font-size:10px;color:#9a9590;font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">📄 ${nombreC}</div>
  <div style="display:flex;flex-direction:column;gap:6px">
    <button id="btn-pdfc" style="padding:10px 14px;background:#1a3a5c;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer;font-size:13px">⬇️ Descargar PDF</button>
    <button onclick="document.getElementById('pflbtns').style.display='none';window.print()" style="padding:10px 14px;background:#f5f4f0;color:#1a1814;border:1px solid #e8e5de;border-radius:8px;font-weight:600;cursor:pointer;font-size:12px">🖨️ Imprimir</button>
  </div>
  <div id="pdf-msgc" style="font-size:10px;color:#9a9590;margin-top:6px;min-height:14px"></div>
</div>
<script>
// Load html2pdf lazily when user clicks Download
document.getElementById("btn-pdfc").onclick=function(){
  var btn=this;
  var msg=document.getElementById("pdf-msgc");
  btn.textContent="Cargando...";
  btn.disabled=true;
  msg.textContent="Preparando PDF...";
  var s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js";
  s.onload=function(){
    var target=document.getElementById("doc-mainc");
    if(!target){msg.textContent="Error: contenido no encontrado";btn.textContent="⬇️ Descargar PDF";btn.disabled=false;return}
    btn.textContent="Generando...";
    msg.textContent="Esto toma unos segundos...";
    document.getElementById("pflbtns").style.opacity="0.5";
    html2pdf().set({
      margin:[8,8,8,8],filename:"${nombreC}.pdf",
      image:{type:"jpeg",quality:0.98},
      html2canvas:{scale:2,useCORS:true,logging:false,removeContainer:true},
      jsPDF:{unit:"mm",format:"a4",orientation:"portrait",compress:true}
    }).from(target).save().then(function(){
      btn.textContent="⬇️ Descargar PDF";btn.disabled=false;
      document.getElementById("pflbtns").style.opacity="1";
      msg.textContent="✓ PDF guardado";
    }).catch(function(e){
      btn.textContent="⬇️ Descargar PDF";btn.disabled=false;
      document.getElementById("pflbtns").style.opacity="1";
      msg.textContent="Error: "+e.message;
    });
  };
  s.onerror=function(){msg.textContent="Error cargando librería";btn.textContent="⬇️ Descargar PDF";btn.disabled=false;};
  document.head.appendChild(s);
};
</script>`
    const htmlFinalC=html
      .replace("</head>",'</head><style>#pflbtns{font-family:Arial,sans-serif}@media print{#pflbtns{display:none!important}}</style>')
      .replace('<div class="page">','<div class="page" id="doc-mainc">')
      .replace("</body>",btnsC+"</body>")
    // Use Blob URL to avoid popup blocker
    const blob=new Blob([htmlFinalC],{type:"text/html;charset=utf-8"})
    const url=URL.createObjectURL(blob)
    const a=document.createElement("a")
    a.href=url; a.target="_blank"; a.rel="noopener"
    document.body.appendChild(a); a.click()
    setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url)},1000)
  }


  // Solo contratos tipo contrato (Excel importados)
  const base=contratos
    .filter(x=>(x.tipo||"contrato")==="contrato")
    .map(x=>({
      ...x,
      vendedor:x.vendedor||vendedorDesdeFolio(x.folio||""),
      pagos:x.pagos||[],
      cobrado:x.cobrado||0,
      total:x.total||0,
    }))

  // Meses disponibles
  const meses=[...new Set(base.map(x=>x.fecha_evento?.slice(0,7)).filter(Boolean))].sort().reverse()

  // Filtrar
  const filtrados=base.filter(x=>{
    if(filtroMes&&x.fecha_evento?.slice(0,7)!==filtroMes)return false
    if(filtroVendedor!=="todos"&&x.vendedor!==filtroVendedor)return false
    if(soloMios&&vendedorActual&&x.vendedor!==vendedorActual)return false
    if(busq){
      const q=busq.toLowerCase()
      return(x.cliente||"").toLowerCase().includes(q)||
        (x.folio||"").toLowerCase().includes(q)||
        (x.lugar||"").toLowerCase().includes(q)||
        (x.vendedor||"").toLowerCase().includes(q)
    }
    return true
  }).sort((a,b)=>(a.fecha_evento||"").localeCompare(b.fecha_evento||""))

  const vendedoresUnicos=[...new Set(base.map(x=>x.vendedor).filter(Boolean))]
  const totalArts=(x:Contrato)=>(x.articulos||[]).reduce((s:number,a:Articulo)=>s+(a.cantidad||0),0)

  return(
    <div style={{display:"grid",gridTemplateColumns:selContrato&&!isMobile?"1fr 540px":"1fr",gap:14}}>
      {/* ── LISTA ── */}
      <div>
        {/* Filtros */}
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"12px 14px",marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap" as const,marginBottom:10}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,flex:1}}>
              📄 Contratos
              <span style={{fontSize:12,fontWeight:400,color:"#9a9590",marginLeft:8}}>({filtrados.length})</span>
            </div>
            <input value={busq} onChange={e=>setBusq(e.target.value)} placeholder="🔍 Buscar..."
              style={{padding:"6px 12px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",width:180}}/>
            {vendedorActual&&(
              <button onClick={()=>setSoloMios(v=>!v)}
                style={{padding:"6px 14px",borderRadius:8,border:`1.5px solid ${soloMios?"#2563eb":"#e8e5de"}`,background:soloMios?"#2563eb":"#eff6ff",color:soloMios?"#fff":"#2563eb",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"Epilogue,sans-serif",whiteSpace:"nowrap" as const,flexShrink:0}}>
                {soloMios?"✓ Solo los míos":"👤 Solo los míos"}
              </button>
            )}
            <select value={filtroVendedor} onChange={e=>{setFiltroVendedor(e.target.value);setSoloMios(false)}}
              style={{padding:"6px 10px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none"}}>
              <option value="todos">Todos los vendedores</option>
              {vendedoresUnicos.map(v=><option key={v} value={v}>{v}</option>)}
            </select>
            <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)}
              style={{padding:"6px 10px",border:"1px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none"}}>
              <option value="">Todos los meses</option>
              {meses.map(m=><option key={m} value={m}>{m}</option>)}
            </select>
          </div>
        </div>

        {/* Cards */}
        {filtrados.length===0?(
          <div style={{padding:48,textAlign:"center" as const,background:"#fff",border:"1.5px dashed #e8e5de",borderRadius:12,color:"#9a9590"}}>
            <div style={{fontSize:32,opacity:.2,marginBottom:8}}>📄</div>
            <div>Sin contratos con los filtros actuales</div>
          </div>
        ):(
          <div style={{display:"flex",flexDirection:"column" as const,gap:6}}>
            {filtrados.map((x,i)=>{
              const isSelected=selContrato?.id===x.id
              const nArts=totalArts(x)
              const saldo=(x.total||0)-(x.cobrado||0)
              return(
                <div key={x.id} onClick={()=>{
                  setSelContrato(isSelected?null:x)
                  setMostrarSplits(false)
                  setSplitsDelContrato([])
                  if(!isSelected && splitStatus[x.id]) cargarSplitsDeContrato(x.id)
                }}
                  style={{background:"#fff",border:`1.5px solid ${isSelected?"#0f172a":"#e8e5de"}`,borderRadius:10,padding:"10px 14px",cursor:"pointer",transition:"all .15s",position:"relative" as const,
                    boxShadow:isSelected?"0 2px 12px rgba(15,23,42,.12)":"none"}}>
                  <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                    {/* Fecha box */}
                    <div style={{width:44,flexShrink:0,textAlign:"center" as const,background:isSelected?"#0f172a":"#f5f4f0",borderRadius:8,padding:"6px 4px"}}>
                      <div style={{fontSize:18,fontWeight:800,color:isSelected?"#fff":"#1a1814",lineHeight:1}}>{x.fecha_evento?.slice(8)||"—"}</div>
                      <div style={{fontSize:8,color:isSelected?"rgba(255,255,255,.6)":"#9a9590",textTransform:"uppercase" as const}}>
                        {["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][(parseInt(x.fecha_evento?.slice(5,7)||"1")-1)]||""}
                      </div>
                    </div>
                    {/* Info */}
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2,flexWrap:"wrap" as const}}>
                        <span style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente||x.archivo}</span>
                        {x.folio&&<span style={{fontSize:9,fontFamily:"monospace",background:"#f5f4f0",color:"#9a9590",padding:"1px 5px",borderRadius:4}}>{x.folio}</span>}
                        {x.vendedor&&<span style={{fontSize:9,background:"#edf3fa",color:"#1a3a5c",padding:"1px 5px",borderRadius:4,fontWeight:600}}>👤{x.vendedor}</span>}
                      </div>
                      <div style={{fontSize:11,color:"#9a9590",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>📍 {x.lugar?.slice(0,45)||"—"}</div>
                      {x.tel&&<div style={{fontSize:10,color:"#4a4640",marginTop:1}}>📞 {x.tel}</div>}
                      <div style={{display:"flex",gap:5,marginTop:5,flexWrap:"wrap" as const}}>
                        {nArts>0&&<span style={{fontSize:9,background:"#f5f4f0",color:"#4a4640",padding:"1px 6px",borderRadius:4}}>📦 {nArts} pzas</span>}
                        {(x.asig_entrega||[]).length>0&&<span style={{fontSize:9,background:"#edf3fa",color:"#1a3a5c",padding:"1px 6px",borderRadius:4}}>🚚 {(x.asig_entrega||[]).join(", ")}</span>}
                        {!(x.asig_entrega||[]).length&&x.fecha_entrega&&new Date(x.fecha_entrega)>=new Date()&&<span style={{fontSize:9,background:"#fdf0f0",color:"#8b2e2e",padding:"1px 6px",borderRadius:4,fontWeight:700}}>⚠️ Sin asignar</span>}
                      </div>
                    </div>
                    {/* Monto */}
                    <div style={{textAlign:"right" as const,flexShrink:0}}>
                      {(x.total||0)>0&&<div style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#1a1814"}}>${(x.total||0).toLocaleString("es-MX")}</div>}
                      {saldo>0&&<div style={{fontSize:10,color:"#92580a",fontWeight:600}}>Debe ${saldo.toLocaleString("es-MX")}</div>}
                      {saldo<=0&&(x.total||0)>0&&<div style={{fontSize:10,color:"#2d6a4f",fontWeight:600}}>✓ Liq.</div>}
                      <div style={{fontSize:10,color:"#9a9590",marginTop:2}}>{isSelected?"▴ Cerrar":"▾ Ver"}</div>
                    </div>
                  </div>
                  {/* Delete contrato */}
                  <button onClick={e=>{e.stopPropagation();eliminarContratoConPwd(String(x.id),()=>{
                      onActualizar&&onActualizar(String(x.id),null)
                      if(selContrato?.id===x.id)setSelContrato(null)
                    })}}
                    style={{position:"absolute" as const,bottom:8,right:8,padding:"2px 7px",borderRadius:5,border:"1px solid #fca5a5",background:"#fef2f2",color:"#8b2e2e",cursor:"pointer",fontSize:9,fontWeight:700,opacity:0.6}}
                    title="Eliminar contrato">🗑️</button>
                </div>
              )
            })}
          </div>
        )}
      </div>

      {/* ── DETALLE CONTRATO ── */}
      {selContrato&&(
        <div style={isMobile?{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.6)",zIndex:2000,display:"flex",flexDirection:"column" as const,justifyContent:"flex-end" as const}:{position:"sticky" as const,top:70,alignSelf:"flex-start" as const}}>
          <div style={{background:"#fff",borderRadius:isMobile?"16px 16px 0 0":"12px",overflow:"hidden",maxHeight:isMobile?"90vh":"85vh",overflowY:"auto" as const,border:isMobile?"none":"1px solid #e8e5de"}}>
            {/* Header rediseñado */}
            <div style={{background:"linear-gradient(135deg,#0f172a 0%,#1e3a5c 100%)",padding:"20px 20px 16px"}}>
              {/* Top row: cliente + acciones */}
              <div style={{display:"flex",alignItems:"flex-start",gap:10,marginBottom:14}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,color:"#fff",lineHeight:1.2,marginBottom:4}}>{selContrato.cliente||selContrato.archivo}</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap" as const,alignItems:"center"}}>
                    {selContrato.folio&&<span style={{fontFamily:"monospace",fontSize:12,color:"#60a5fa",background:"rgba(96,165,250,.15)",padding:"2px 8px",borderRadius:5,fontWeight:700}}>{selContrato.folio}</span>}
                    {selContrato.vendedor&&<span style={{fontSize:11,color:"rgba(255,255,255,.6)"}}>👤 {selContrato.vendedor}</span>}
                    {selContrato.total>0&&<span style={{fontSize:12,color:"#4ade80",fontFamily:"monospace",fontWeight:700}}>${(selContrato.total||0).toLocaleString("es-MX")}</span>}
                  </div>
                </div>
                <button onClick={()=>setSelContrato(null)} style={{background:"rgba(255,255,255,.1)",border:"none",color:"rgba(255,255,255,.7)",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>✕</button>
              </div>
              {/* Fechas rápidas */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
                {[{l:"Evento",v:selContrato.fecha_evento,ico:"🎉"},{l:"Entrega",v:selContrato.fecha_entrega,ico:"🚚"},{l:"Desmonte",v:selContrato.fecha_desmonte,ico:"📦"}].map((k,i)=>(
                  <div key={i} style={{background:"rgba(255,255,255,.08)",borderRadius:8,padding:"8px 10px",textAlign:"center" as const}}>
                    <div style={{fontSize:14,marginBottom:2}}>{k.ico}</div>
                    <div style={{fontSize:11,fontWeight:700,color:"#fff"}}>{k.v||"—"}</div>
                    <div style={{fontSize:9,color:"rgba(255,255,255,.4)",textTransform:"uppercase" as const}}>{k.l}</div>
                  </div>
                ))}
              </div>
              {/* Botones de acción */}
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>modoEdicion?setModoEdicion(false):abrirEdicionContrato(selContrato)}
                  style={{flex:1,padding:"9px 10px",borderRadius:8,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif",background:modoEdicion?"rgba(239,68,68,.3)":"rgba(255,255,255,.15)",color:"#fff"}}>
                  {modoEdicion?"✕ Cancelar edición":"✏️ Editar contrato"}
                </button>
                <button onClick={()=>generarSplits(selContrato)} disabled={generandoSplit===selContrato.id}
                  style={{flex:1,padding:"9px 10px",borderRadius:8,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif",background:splitStatus[selContrato.id]?"rgba(45,106,79,.8)":"rgba(255,255,255,.15)",color:"#fff"}}>
                  {generandoSplit===selContrato.id?"...":splitStatus[selContrato.id]?"✂️ Re-generar":"✂️ Splits"}
                </button>
                <button onClick={()=>imprimirContrato(selContrato)}
                  style={{padding:"9px 14px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,background:"rgba(255,255,255,.2)",color:"#fff",fontWeight:700}} title="Imprimir contrato">🖨️ PDF</button>
              </div>
              {/* Declinar contrato */}
              <button onClick={async()=>{
                const pwd=window.prompt("Contraseña para declinar contrato:")
                if(!pwd)return
                if(pwd!=="LITA2024"){window.alert("❌ Contraseña incorrecta");return}
                if(!window.confirm(`¿Declinar contrato de ${selContrato.cliente||selContrato.archivo}?\nEsta acción no se puede deshacer.`))return
                await fetch(`/api/contratos?id=${selContrato.id}`,{
                  method:"PATCH",
                  headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
                  body:JSON.stringify({tipo:"declinado"})
                })
                onActualizar&&onActualizar(String(selContrato.id),{tipo:"declinado"})
                setSelContrato(null)
              }}
                style={{width:"100%",marginTop:6,padding:"7px",borderRadius:8,border:"1px solid #fca5a5",background:"transparent",color:"#fca5a5",cursor:"pointer",fontSize:11,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
                ✗ Declinar contrato
              </button>
            </div>
                        {/* ── MODO EDICIÓN ── */}
            {modoEdicion?(
              <div style={{display:"flex",flexDirection:"column" as const,height:"100%"}}>

                {/* Sección: Datos del cliente */}
                <div style={{padding:"16px 20px",borderBottom:"1px solid #f0ece4",background:"#fafaf8"}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".08em",marginBottom:12}}>👤 Cliente</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div style={{gridColumn:"1/-1"}}>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>Nombre del cliente</label>
                      <input value={editForm.cliente||""} onChange={e=>setEditForm((p:any)=>({...p,cliente:e.target.value}))}
                        placeholder="Nombre completo..."
                        style={{width:"100%",padding:"7px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:600,outline:"none",boxSizing:"border-box" as const,transition:"border .15s"}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>Folio</label>
                      <input value={editForm.folio||""} onChange={e=>setEditForm((p:any)=>({...p,folio:e.target.value}))}
                        placeholder="K-001..."
                        style={{width:"100%",padding:"7px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"monospace",fontSize:12,fontWeight:700,outline:"none",boxSizing:"border-box" as const}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>Teléfono</label>
                      <input value={editForm.tel||""} onChange={e=>setEditForm((p:any)=>({...p,tel:e.target.value}))}
                        placeholder="55 1234 5678"
                        style={{width:"100%",padding:"7px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                    <div style={{gridColumn:"1/-1"}}>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>📍 Lugar del evento</label>
                      <input value={editForm.lugar||""} onChange={e=>setEditForm((p:any)=>({...p,lugar:e.target.value}))}
                        placeholder="Dirección o nombre del salón..."
                        style={{width:"100%",padding:"7px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                  </div>
                </div>

                {/* Sección: Fechas */}
                <div style={{padding:"16px 20px",borderBottom:"1px solid #f0ece4",background:"#fff"}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".08em",marginBottom:12}}>📅 Fechas</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                    {([
                      {l:"🎉 Evento",k:"fecha_evento"},
                      {l:"🚚 Entrega",k:"fecha_entrega"},
                      {l:"📦 Desmonte",k:"fecha_desmonte"},
                    ] as {l:string,k:string}[]).map(f=>(
                      <div key={f.k}>
                        <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>{f.l}</label>
                        <input type="date" value={(editForm as any)[f.k]||""}
                          onChange={e=>setEditForm((p:any)=>({...p,[f.k]:e.target.value}))}
                          style={{width:"100%",padding:"6px 8px",border:"1.5px solid #e8e5de",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:11,outline:"none",boxSizing:"border-box" as const}}
                          onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                          onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Sección: Vendedor + Notas */}
                <div style={{padding:"12px 20px",borderBottom:"1px solid #f0ece4",background:"#fafaf8"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>👤 Vendedor</label>
                      <select value={editForm.vendedor||""} onChange={e=>setEditForm((p:any)=>({...p,vendedor:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"2px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",background:"#fff"}}>
                        <option value="">Sin asignar</option>
                        {VENDEDORES.map((v:any)=><option key={v.nombre} value={v.nombre}>{v.nombre}</option>)}
                      </select>
                    </div>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:4}}>📝 Notas</label>
                      <textarea value={editForm.notas||""} onChange={e=>setEditForm((p:any)=>({...p,notas:e.target.value}))} rows={1}
                        style={{width:"100%",padding:"8px 10px",border:"2px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                  </div>
                </div>

                {/* Sección: Descuento + IVA */}
                <div style={{padding:"12px 20px",borderBottom:"1px solid #f0ece4",background:"#fff"}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>🏷️ Descuento global e IVA</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    {/* Tipo descuento */}
                    <div style={{display:"flex",gap:4}}>
                      <button onClick={()=>setEditDescTipo("pct")}
                        style={{flex:1,padding:"6px 4px",borderRadius:6,border:`2px solid ${editDescTipo==="pct"?"#1a3a5c":"#e8e5de"}`,background:editDescTipo==="pct"?"#1a3a5c":"#fff",color:editDescTipo==="pct"?"#fff":"#4a4640",fontWeight:700,fontSize:11,cursor:"pointer"}}>
                        % Porc.
                      </button>
                      <button onClick={()=>setEditDescTipo("monto")}
                        style={{flex:1,padding:"6px 4px",borderRadius:6,border:`2px solid ${editDescTipo==="monto"?"#1a3a5c":"#e8e5de"}`,background:editDescTipo==="monto"?"#1a3a5c":"#fff",color:editDescTipo==="monto"?"#fff":"#4a4640",fontWeight:700,fontSize:11,cursor:"pointer"}}>
                        $ Fijo
                      </button>
                    </div>
                    {/* Input descuento */}
                    <div style={{display:"flex",alignItems:"center",gap:4}}>
                      {editDescTipo==="pct"?<span style={{fontSize:13,fontWeight:800,color:"#4a4640"}}>%</span>:<span style={{fontSize:13,fontWeight:800,color:"#4a4640"}}>$</span>}
                      <input type="number" min="0" max={editDescTipo==="pct"?100:undefined}
                        value={editDescTipo==="pct"?(editDescPct||""):(editDescMonto||"")}
                        placeholder="0"
                        onChange={e=>{
                          const v=parseFloat(e.target.value)||0
                          if(editDescTipo==="pct") setEditDescPct(Math.min(100,v))
                          else setEditDescMonto(v)
                        }}
                        style={{flex:1,padding:"6px 8px",border:"2px solid #e8e5de",borderRadius:6,fontFamily:"monospace",fontSize:13,fontWeight:700,outline:"none",boxSizing:"border-box" as const,textAlign:"center" as const}}
                        onFocus={e=>e.target.style.border="2px solid #1a3a5c"}
                        onBlur={e=>e.target.style.border="2px solid #e8e5de"}/>
                    </div>
                    {/* IVA toggle */}
                    <div style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:"#f5f4f0",borderRadius:7}}>
                      <input type="checkbox" id="edit-iva" checked={editAplIva}
                        onChange={e=>setEditAplIva(e.target.checked)}
                        style={{width:14,height:14,cursor:"pointer",accentColor:"#4a2d6e"}}/>
                      <label htmlFor="edit-iva" style={{fontSize:12,fontWeight:600,color:"#4a2d6e",cursor:"pointer"}}>Aplicar IVA 16% (sobre base con descuentos)</label>
                    </div>
                  </div>
                </div>

                {/* Sección: Artículos — tabla inline estilo Excel */}
                <div style={{padding:"14px 20px 8px",borderBottom:"1px solid #f0ece4",background:"#fff",flex:1,overflowY:"auto" as const}}>
                  <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8,display:"flex",justifyContent:"space-between" as const,alignItems:"center"}}>
                    <span>📦 Artículos</span>
                    <span style={{fontFamily:"monospace",fontWeight:800,color:"#0f172a",fontSize:11}}>{editArtsContrato.length} artículos · {editArtsContrato.reduce((s:number,a:any)=>s+(a.cantidad||0),0)} pzas</span>
                  </div>

                  {/* Tabla inline */}
                  <div style={{border:"1.5px solid #e8e5de",borderRadius:10,overflow:"hidden",marginBottom:8}}>
                    {/* Header */}
                    <div style={{display:"grid",gridTemplateColumns:"58px 1fr 90px 58px 60px 24px",background:"#0f172a"}}>
                      {["Cant.","Artículo","P.U.","Desc%","Subtotal",""].map((h,hi)=>(
                        <div key={hi} style={{padding:"7px 6px",fontSize:9,fontWeight:700,color:"#94a3b8",textAlign:hi>=2?"center" as const:"left" as const,textTransform:"uppercase" as const,letterSpacing:".05em",borderRight:hi<6?"1px solid rgba(255,255,255,.07)":"none"}}>{h}</div>
                      ))}
                    </div>
                    {/* Filas */}
                    {editArtsContrato.length===0?(
                      <div style={{padding:16,textAlign:"center" as const,color:"#c4bfb8",fontSize:11,fontStyle:"italic"}}>Sin artículos — busca abajo para agregar</div>
                    ):editArtsContrato.map((a:any,i:number)=>{
                      const bruto=(a.pu||0)*(a.cantidad||0)
                      const descMonto=a.aplica_descuento?Math.round(bruto*(a.descuento_pct_art||0)/100):0
                      const neto=bruto-descMonto
                      return(
                        <div key={i} style={{display:"grid",gridTemplateColumns:"58px 1fr 90px 58px 60px 24px",borderBottom:i<editArtsContrato.length-1?"1px solid #f0ece4":"none",background:i%2===0?"#fff":"#fafaf8",alignItems:"center"}}>

                          {/* Cantidad */}
                          <div style={{borderRight:"1px solid #ebebeb",display:"flex",justifyContent:"center"}}>
                            <input type="number" min="1" value={a.cantidad||1}
                              onChange={e=>{const v=Math.max(1,parseInt(e.target.value)||1);setEditArtsContrato(prev=>prev.map((x:any,j:number)=>j===i?{...x,cantidad:v}:x))}}
                              style={{width:"100%",padding:"6px 4px",border:"none",textAlign:"center" as const,fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1a3a5c",outline:"none",background:"transparent"}}/>
                          </div>
                          {/* Nombre editable */}
                          <div style={{borderRight:"1px solid #ebebeb",position:"relative" as const}}>
                            <input value={a.nombre||""} onChange={e=>setEditArtsContrato(prev=>prev.map((x:any,j:number)=>j===i?{...x,nombre:e.target.value}:x))}
                              style={{width:"100%",padding:"6px 8px",border:"none",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:500,outline:"none",background:"transparent",boxSizing:"border-box" as const}}/>
                          </div>
                          {/* P.U. */}
                          <div style={{borderRight:"1px solid #ebebeb",display:"flex",justifyContent:"flex-end"}}>
                            <input type="number" min="0" value={a.pu||0}
                              onChange={e=>{const v=parseFloat(e.target.value)||0;setEditArtsContrato(prev=>prev.map((x:any,j:number)=>j===i?{...x,pu:v}:x))}}
                              style={{width:"100%",padding:"6px 6px",border:"none",textAlign:"right" as const,fontFamily:"monospace",fontSize:11,fontWeight:600,outline:"none",background:"transparent"}}/>
                          </div>
                          {/* Descuento % */}
                          <div style={{borderRight:"1px solid #ebebeb",display:"flex",flexDirection:"column" as const,alignItems:"center",justifyContent:"center",padding:"3px 2px",gap:1}}>
                            <input type="checkbox" checked={!!a.aplica_descuento}
                              onChange={e=>setEditArtsContrato(prev=>prev.map((x:any,j:number)=>j===i?{...x,aplica_descuento:e.target.checked,descuento_pct_art:e.target.checked?(x.descuento_pct_art||0):0}:x))}
                              style={{width:12,height:12,cursor:"pointer",accentColor:"#8b2e2e"}}/>
                            {a.aplica_descuento&&(
                              <div style={{display:"flex",alignItems:"center",gap:1}}>
                                <input type="number" min="0" max="100" value={a.descuento_pct_art||""}
                                  placeholder="0"
                                  onChange={e=>{const v=Math.min(100,parseFloat(e.target.value)||0);setEditArtsContrato(prev=>prev.map((x:any,j:number)=>j===i?{...x,descuento_pct_art:v}:x))}}
                                  style={{width:46,padding:"2px 4px",border:"1.5px solid #8b2e2e",borderRadius:4,textAlign:"center" as const,fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#8b2e2e",outline:"none"}}/>
                                <span style={{fontSize:9,color:"#8b2e2e",fontWeight:700}}>%</span>
                              </div>
                            )}
                          </div>
                          {/* Subtotal neto */}
                          <div style={{borderRight:"1px solid #ebebeb",padding:"4px 6px",textAlign:"right" as const}}>
                            <div style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#1a3a5c"}}>${neto.toLocaleString("es-MX")}</div>
                            {descMonto>0&&<div style={{fontSize:8,color:"#8b2e2e",fontFamily:"monospace"}}>-${descMonto.toLocaleString("es-MX")}</div>}
                          </div>
                          {/* Delete */}
                          <div style={{display:"flex",justifyContent:"center"}}>
                            <button onClick={()=>setEditArtsContrato(prev=>prev.filter((_:any,j:number)=>j!==i))}
                              style={{width:20,height:20,borderRadius:4,background:"#fdf0f0",border:"none",cursor:"pointer",color:"#8b2e2e",fontSize:13,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                          </div>
                        </div>
                      )
                    })}
                    {/* Fila subtotal */}
                    {editArtsContrato.length>0&&(()=>{
                      const _tots=calcTotalContrato(editArtsContrato,editDescPct,editDescMonto,editAplIva)
                      return(
                        <div style={{background:"#f8f6f2",borderTop:"2px solid #e8e5de",padding:"8px 10px"}}>
                          <div style={{display:"flex",justifyContent:"space-between" as const,fontSize:11,marginBottom:2}}>
                            <span style={{color:"#9a9590"}}>Subtotal artículos</span>
                            <span style={{fontFamily:"monospace",fontWeight:700}}>${_tots.subtotal.toLocaleString("es-MX")}</span>
                          </div>
                          {_tots.descGlobal>0&&(
                            <div style={{display:"flex",justifyContent:"space-between" as const,fontSize:11,marginBottom:2,color:"#2d6a4f"}}>
                              <span>Descuento global</span>
                              <span style={{fontFamily:"monospace",fontWeight:700}}>-${_tots.descGlobal.toLocaleString("es-MX")}</span>
                            </div>
                          )}
                          {editAplIva&&(
                            <div style={{display:"flex",justifyContent:"space-between" as const,fontSize:11,marginBottom:2,color:"#4a2d6e"}}>
                              <span>IVA 16%</span>
                              <span style={{fontFamily:"monospace",fontWeight:700}}>${_tots.iva.toLocaleString("es-MX")}</span>
                            </div>
                          )}
                          <div style={{display:"flex",justifyContent:"space-between" as const,fontSize:13,borderTop:"1.5px solid #1a1814",paddingTop:6,marginTop:4}}>
                            <span style={{fontWeight:800}}>TOTAL</span>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontFamily:"monospace",fontWeight:800,color:"#1a3a5c"}}>${_tots.total.toLocaleString("es-MX")}</span>
                              <button onClick={()=>{
                                  setEditForm((p:any)=>({...p,total:_tots.total}))
                                }}
                                style={{fontSize:9,padding:"2px 8px",borderRadius:4,background:"#1a3a5c",color:"#fff",border:"none",cursor:"pointer",fontWeight:700,whiteSpace:"nowrap" as const}}>← Aplicar</button>
                            </div>
                          </div>
                        </div>
                      )
                    })()}
                  </div>

                  {/* Buscador para agregar */}
                  <div style={{position:"relative" as const}}>
                    <input value={busqArtCont} onChange={async e=>{
                      const v=e.target.value; setBusqArtCont(v)
                      if(v.trim().length>=1){
                        const r=await fetch(`/api/catalogo?busq=${encodeURIComponent(v.trim())}&activo=true`,{headers:{Authorization:`Bearer ${token}`}})
                        const data=await r.json()
                        setSugsArtCont(Array.isArray(data)?data.slice(0,8):[])
                      }else{setSugsArtCont([])}
                    }} onBlur={()=>setTimeout(()=>setSugsArtCont([]),200)}
                    placeholder="+ Buscar artículo del catálogo..."
                    style={{width:"100%",padding:"6px 10px",border:"1.5px solid #1a3a5c",borderRadius:7,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",boxSizing:"border-box" as const,background:"#eff6ff"}}/>
                    {sugsArtCont.length>0&&(
                      <div style={{position:"absolute" as const,top:"100%",left:0,right:0,background:"#fff",border:"2px solid #1a3a5c",borderRadius:"0 0 9px 9px",boxShadow:"0 8px 24px rgba(26,58,92,.15)",zIndex:500,maxHeight:200,overflowY:"auto" as const}}>
                        {sugsArtCont.map((a:any,ai:number)=>(
                          <div key={ai} onMouseDown={e=>{
                            e.preventDefault()
                            const exists=editArtsContrato.some((x:any)=>x.nombre===a.nombre)
                            if(exists){setEditArtsContrato(prev=>prev.map((x:any)=>x.nombre===a.nombre?{...x,cantidad:(x.cantidad||0)+1}:x))}
                            else{setEditArtsContrato(prev=>[...prev,{nombre:a.nombre,cantidad:1,pu:a.precio_renta||0,importe:a.precio_renta||0,seccion:a.categoria||"",aplica_descuento:false,descuento_pct_art:0}])}
                            setBusqArtCont(""); setSugsArtCont([])
                          }} style={{padding:"8px 12px",cursor:"pointer",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,fontSize:12}}
                          onMouseEnter={e=>(e.currentTarget as HTMLElement).style.background="#eff6ff"}
                          onMouseLeave={e=>(e.currentTarget as HTMLElement).style.background="#fff"}>
                            <div style={{flex:1}}><div style={{fontWeight:600}}>{a.nombre}</div><div style={{fontSize:10,color:"#9a9590"}}>{a.categoria}</div></div>
                            {(a.precio_renta||0)>0&&<span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#1a3a5c"}}>${(a.precio_renta||0).toLocaleString("es-MX")}</span>}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>

                {/* Botón guardar sticky */}
                <div style={{padding:"14px 20px",background:"#fff",borderTop:"2px solid #e8e5de"}}>
                  <button onClick={guardarEdicionContrato} disabled={guardandoEdit}
                    style={{width:"100%",padding:"14px",borderRadius:10,background:guardandoEdit?"#9a9590":"#0f172a",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:15,fontWeight:800,letterSpacing:".02em"}}>
                    {guardandoEdit?"💾 Guardando...":"💾 Guardar cambios"}
                  </button>
                </div>
              </div>
            ):(
            <>
            {/* Artículos */}
            <div style={{padding:"12px 14px",borderBottom:"1px solid #e8e5de"}}>
              <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,marginBottom:8}}>
                Artículos — {totalArts(selContrato)} piezas
              </div>
              {(selContrato.articulos||[]).length===0
                ?<div style={{fontSize:11,color:"#c4bfb8",fontStyle:"italic"}}>Sin artículos</div>
                :(selContrato.articulos||[]).map((a:any,i:number)=>(
                  <div key={i} style={{display:"flex",gap:8,fontSize:11,padding:"5px 0",borderBottom:"1px solid #f5f4f0"}}>
                    <span style={{fontWeight:700,color:"#1a3a5c",minWidth:28,textAlign:"right" as const}}>{a.cantidad}x</span>
                    <span style={{flex:1}}>{a.nombre}</span>
                    {(a.importe||a.subtotal||0)>0&&<span style={{fontFamily:"monospace",color:"#2d6a4f",fontWeight:600}}>${(a.importe||a.subtotal||0).toLocaleString("es-MX")}</span>}
                  </div>
                ))
              }
            </div>
            {/* Totales + Pagos */}
            {(selContrato.total||0)>0&&(
              <div style={{padding:"12px 14px",borderBottom:"1px solid #e8e5de"}}>
                {/* Descuentos aplicados */}
                {((selContrato.descuento_pct||0)>0||(selContrato.descuento_monto_global||0)>0)&&(
                  <div style={{marginBottom:8,padding:"8px 10px",background:"#f0fdf4",borderRadius:8,border:"1px solid #b7deca"}}>
                    <div style={{fontSize:10,fontWeight:700,color:"#2d6a4f",marginBottom:4}}>🏷️ Descuentos aplicados</div>
                    {(selContrato.descuento_pct||0)>0&&(
                      <div style={{fontSize:11,color:"#2d6a4f",display:"flex",justifyContent:"space-between"}}>
                        <span>Descuento global ({selContrato.descuento_pct}%)</span>
                        <span style={{fontFamily:"monospace",fontWeight:700}}>-${Math.round((selContrato.subtotal_bruto||selContrato.total)*(selContrato.descuento_pct||0)/100).toLocaleString("es-MX")}</span>
                      </div>
                    )}
                    {(selContrato.descuento_monto_global||0)>0&&(
                      <div style={{fontSize:11,color:"#2d6a4f",display:"flex",justifyContent:"space-between"}}>
                        <span>Descuento fijo</span>
                        <span style={{fontFamily:"monospace",fontWeight:700}}>-${(selContrato.descuento_monto_global||0).toLocaleString("es-MX")}</span>
                      </div>
                    )}
                    {selContrato.aplica_iva&&(
                      <div style={{fontSize:11,color:"#4a2d6e",display:"flex",justifyContent:"space-between",marginTop:2}}>
                        <span>IVA incluido (16%)</span>
                        <span style={{fontFamily:"monospace",fontWeight:700}}>+${Math.round((selContrato.total||0)*0.16/1.16).toLocaleString("es-MX")}</span>
                      </div>
                    )}
                  </div>
                )}
                <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
                  <span style={{color:"#9a9590"}}>Total</span>
                  <span style={{fontFamily:"monospace",fontWeight:700}}>${(selContrato.total||0).toLocaleString("es-MX")}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
                  <span style={{color:"#2d6a4f"}}>Cobrado</span>
                  <span style={{fontFamily:"monospace",fontWeight:700,color:"#2d6a4f"}}>${(selContrato.cobrado||0).toLocaleString("es-MX")}</span>
                </div>
                {(selContrato.total||0)-(selContrato.cobrado||0)>0&&(
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"8px 0",borderTop:"2px solid #1a1814",marginTop:4}}>
                    <span style={{fontWeight:700,color:"#92580a"}}>Saldo pendiente</span>
                    <span style={{fontFamily:"monospace",fontWeight:800,color:"#92580a"}}>${((selContrato.total||0)-(selContrato.cobrado||0)).toLocaleString("es-MX")}</span>
                  </div>
                )}
                {(selContrato.total||0)<=(selContrato.cobrado||0)&&(
                  <div style={{textAlign:"center" as const,marginTop:6,fontSize:11,color:"#2d6a4f",fontWeight:700}}>✓ Liquidado</div>
                )}
              </div>
            )}
            {/* Historial de pagos — siempre visible */}
            <div style={{padding:"12px 14px",borderBottom:"1px solid #e8e5de"}}>
              <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>
                Pagos registrados ({(selContrato.pagos||[]).length})
              </div>
              {(selContrato.pagos||[]).length===0
                ?<div style={{fontSize:11,color:"#c4bfb8",fontStyle:"italic"}}>Sin pagos registrados</div>
                :(selContrato.pagos||[]).map((p:any,i:number)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #f5f4f0",fontSize:11}}>
                    <div>
                      <div style={{fontWeight:600}}>${(p.monto||0).toLocaleString("es-MX")}</div>
                      <div style={{fontSize:9,color:"#9a9590"}}>{p.fecha} · {p.metodo||"efectivo"}{p.nota?" · "+p.nota:""}</div>
                    </div>
                    <span style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:"#edf7f2",color:"#2d6a4f",fontWeight:700}}>{p.metodo||"efectivo"}</span>
                  </div>
                ))
              }
            </div>
            {/* Registrar nuevo pago — deshabilitado */}
          </>
          )}
          </div>

          {/* ── ✂️ SPLITS ── */}
          <div style={{padding:"12px 14px",borderTop:"2px solid #f0ece4"}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1a1814",flex:1}}>✂️ Hojas de trabajo</div>
              {splitStatus[selContrato.id]&&(
                <button onClick={()=>{
                  setMostrarSplits(v=>!v)
                  if(!mostrarSplits&&splitsDelContrato.length===0)
                    cargarSplitsDeContrato(selContrato.id)
                }}
                  style={{padding:"5px 12px",borderRadius:8,border:"none",background:mostrarSplits?"#f5f4f0":"#2563eb",color:mostrarSplits?"#4a4640":"#fff",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>
                  {mostrarSplits?"Ocultar":"👁️ Ver todos"}
                </button>
              )}
              <button onClick={()=>generarSplits(selContrato)}
                disabled={generandoSplit===selContrato.id}
                style={{padding:"5px 12px",borderRadius:8,border:"none",
                  background:generandoSplit===selContrato.id?"#9a9590":splitStatus[selContrato.id]?"#f5f4f0":"#0f172a",
                  color:splitStatus[selContrato.id]?"#4a4640":"#fff",
                  cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>
                {generandoSplit===selContrato.id?"Generando...":splitStatus[selContrato.id]?"🔄 Re-generar":"✂️ Generar splits"}
              </button>
            </div>

            {!splitStatus[selContrato.id]&&(
              <div style={{padding:"12px",background:"#f8f6f2",borderRadius:8,fontSize:11,color:"#9a9590",textAlign:"center" as const}}>
                Sin splits generados — haz clic en <strong>✂️ Generar splits</strong>
              </div>
            )}

            {mostrarSplits&&splitStatus[selContrato.id]&&(
              cargandoSplits?(
                <div style={{padding:16,textAlign:"center" as const,color:"#9a9590",fontSize:12}}>Cargando hojas...</div>
              ):(
                <div style={{display:"flex",flexDirection:"column" as const,gap:6}}>
                  {splitsDelContrato.map((s:any)=>{
                    const tipo=SPLIT_TIPOS.find((t:any)=>t.id===s.tipo)||SPLIT_TIPOS[0]
                    const est=SPLIT_ESTADOS.find((e:any)=>e.id===s.estado)||SPLIT_ESTADOS[0]
                    return(
                      <div key={s.id} style={{borderRadius:10,overflow:"hidden",border:`1px solid ${tipo.color}33`}}>
                        <div style={{background:tipo.color,padding:"8px 12px",display:"flex",alignItems:"center",gap:8}}>
                          <span style={{fontSize:16}}>{tipo.icono}</span>
                          <div style={{flex:1,fontSize:12,fontWeight:700,color:"#fff"}}>{s.nombre||tipo.nombre}</div>
                          <select value={s.estado}
                            onChange={async e=>{
                              await fetch(`/api/splits?id=${s.id}`,{
                                method:"PATCH",
                                headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
                                body:JSON.stringify({estado:e.target.value})
                              })
                              setSplitsDelContrato(prev=>prev.map((x:any)=>x.id===s.id?{...x,estado:e.target.value}:x))
                            }}
                            style={{padding:"2px 6px",borderRadius:6,border:"none",background:est.bg,color:est.color,fontFamily:"Epilogue,sans-serif",fontSize:10,fontWeight:700,cursor:"pointer",outline:"none"}}>
                            {SPLIT_ESTADOS.map((e:any)=><option key={e.id} value={e.id}>{e.label}</option>)}
                          </select>
                          <button onClick={()=>abrirHojaSplit(s,logoUrl)}
                            style={{padding:"3px 8px",borderRadius:6,border:"none",background:"rgba(255,255,255,.2)",color:"#fff",cursor:"pointer",fontSize:11,fontWeight:700,whiteSpace:"nowrap" as const}}>
                            🖨️
                          </button>
                          <button onClick={async()=>{
                            if(!window.confirm("¿Eliminar esta hoja de trabajo?"))return
                            await fetch(`/api/splits?id=${s.id}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}})
                            setSplitsDelContrato(prev=>prev.filter((x:any)=>x.id!==s.id))
                          }}
                            style={{padding:"3px 7px",borderRadius:6,border:"none",background:"rgba(255,255,255,.15)",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>
                            🗑️
                          </button>
                        </div>
                        <div style={{padding:"8px 12px",background:tipo.bg}}>
                          {(s.articulos||[]).length===0?(
                            <div style={{fontSize:11,color:"#9a9590",fontStyle:"italic"}}>Sin artículos clasificados automáticamente</div>
                          ):(
                            <div>
                              <div style={{fontSize:10,color:"#9a9590",marginBottom:3}}>
                                {(s.articulos||[]).length} artículos · {(s.articulos||[]).reduce((sum:number,a:any)=>sum+(a.cantidad||0),0)} piezas
                              </div>
                              {(s.articulos||[]).slice(0,5).map((a:any,ai:number)=>(
                                <div key={ai} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"1px 0"}}>
                                  <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const,flex:1,color:"#4a4640"}}>{a.nombre}</span>
                                  <span style={{fontWeight:700,color:tipo.color,marginLeft:8,flexShrink:0}}>{a.cantidad}</span>
                                </div>
                              ))}
                              {(s.articulos||[]).length>5&&<div style={{fontSize:10,color:"#9a9590"}}>+{(s.articulos||[]).length-5} más</div>}
                            </div>
                          )}
                        </div>
                      </div>
                    )
                  })}
                  <button onClick={async()=>{
                    const nombre=prompt("Nombre del proveedor:")
                    if(!nombre)return
                    const r=await fetch("/api/splits",{
                      method:"POST",
                      headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
                      body:JSON.stringify([{
                        contrato_id:selContrato.id,
                        contrato_folio:selContrato.folio||selContrato.archivo||"",
                        cliente:selContrato.cliente||selContrato.archivo||"",
                        lugar:selContrato.lugar||"",
                        fecha_evento:selContrato.fecha_evento||"",
                        tipo:"proveedor",nombre:`Proveedor: ${nombre}`,
                        proveedor_nombre:nombre,estado:"pendiente",
                        articulos:[],notas:"",observaciones:"",
                        mostrar_precios:false,mostrar_direccion:false,
                      }])
                    })
                    const data=await r.json()
                    if(Array.isArray(data))setSplitsDelContrato(prev=>[...prev,...data])
                  }}
                    style={{padding:"8px",borderRadius:8,background:"#fff",border:"1.5px dashed #8b2e2e",color:"#8b2e2e",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>
                    📄 + Agregar hoja de proveedor
                  </button>
                </div>
              )
            )}
          </div>
        </div>
      )}

    </div>
  )
}


// ─── INICIO DASHBOARD ─────────────────────────────────────────────

// ═══════════════════════════════════════════════════════
// RH SECTION — Recursos Humanos
// ═══════════════════════════════════════════════════════
const RH_PWD = "POLIFLOR_GERENCIA_2026"

function RHSection({token,isMobile}:{token:string,isMobile?:boolean}){
  const SB_URL="https://ohxehnsxfbvdflmqlzxq.supabase.co"
  const SB_KEY="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"

  const [auth,setAuth]=useState(false)
  const [pwdInput,setPwdInput]=useState("")
  const [pwdError,setPwdError]=useState("")
  const [tab,setTab]=useState("empleados")
  const [empleados,setEmpleados]=useState<any[]>([])
  const [vacaciones,setVacaciones]=useState<any[]>([])
  const [incidencias,setIncidencias]=useState<any[]>([])
  const [cargando,setCargando]=useState(false)
  const [busq,setBusq]=useState("")
  const [selEmp,setSelEmp]=useState<any>(null)
  const [modoForm,setModoForm]=useState<"emp"|"vac"|"inc"|null>(null)
  const [form,setForm]=useState<any>({})
  const [guardando,setGuardando]=useState(false)
  const [filtroVac,setFiltroVac]=useState("todos")

  const sbFetch=async(path:string,opts:any={})=>{
    const method=(opts.method||"GET").toUpperCase()
    const r=await fetch(SB_URL+path,{
      ...opts,
      headers:{
        "apikey":SB_KEY,
        "Authorization":`Bearer ${SB_KEY}`,
        "Content-Type":"application/json",
        ...(method==="PATCH"||method==="DELETE"?{"Prefer":"return=minimal"}:{}),
        ...(opts.headers||{})
      }
    })
    return r
  }

  const cargarEmpleados=async()=>{
    setCargando(true)
    const r=await sbFetch("/rest/v1/empleados?select=*&order=nombre.asc")
    const d=await r.json()
    setEmpleados(Array.isArray(d)?d:[])
    setCargando(false)
  }

  const cargarVacaciones=async()=>{
    const r=await sbFetch("/rest/v1/vacaciones?select=*,empleados(nombre,nss)&order=fecha_inicio.desc")
    const d=await r.json()
    setVacaciones(Array.isArray(d)?d:[])
  }

  const cargarIncidencias=async()=>{
    const r=await sbFetch("/rest/v1/incidencias?select=*,empleados(nombre)&order=fecha.desc")
    const d=await r.json()
    setIncidencias(Array.isArray(d)?d:[])
  }

  useEffect(()=>{
    if(auth){cargarEmpleados();cargarVacaciones();cargarIncidencias()}
  },[auth])

  const calcDias=(ini:string,fin:string)=>{
    if(!ini||!fin)return 0
    const d1=new Date(ini+"T12:00:00"),d2=new Date(fin+"T12:00:00")
    return Math.max(0,Math.round((d2.getTime()-d1.getTime())/86400000)+1)
  }

  const guardarEmpleado=async()=>{
    if(!form.nombre?.trim())return
    setGuardando(true)
    if(form.id){
      const {id,...rest}=form
      await sbFetch(`/rest/v1/empleados?id=eq.${id}`,{method:"PATCH",body:JSON.stringify(rest)})
    } else {
      await sbFetch("/rest/v1/empleados",{method:"POST",body:JSON.stringify({...form,activo:true}),headers:{"Prefer":"return=minimal"}})
    }
    setGuardando(false);setModoForm(null);setForm({});cargarEmpleados()
  }

  const guardarVacacion=async()=>{
    if(!form.empleado_id||!form.fecha_inicio||!form.fecha_fin)return
    setGuardando(true)
    const dias=calcDias(form.fecha_inicio,form.fecha_fin)
    if(form.id){
      // Edición: calcular diferencia de días para ajustar programados
      const vacAnterior=vacaciones.find((v:any)=>v.id===form.id)
      const diasAntes=vacAnterior?.dias||0
      const diff=dias-diasAntes
      const {id,...rest}=form
      await sbFetch(`/rest/v1/vacaciones?id=eq.${id}`,{method:"PATCH",body:JSON.stringify({...rest,dias})})
      if(diff!==0){
        const emp=empleados.find((e:any)=>e.id===form.empleado_id)
        if(emp){
          const nuevoProg=Math.max(0,(emp.dias_programados||0)+diff)
          await sbFetch(`/rest/v1/empleados?id=eq.${form.empleado_id}`,{method:"PATCH",body:JSON.stringify({dias_programados:nuevoProg})})
        }
      }
    } else {
      // Alta: sumar días a programados del empleado
      await sbFetch("/rest/v1/vacaciones",{method:"POST",body:JSON.stringify({...form,dias,estado:"pendiente"}),headers:{"Prefer":"return=minimal"}})
      const emp=empleados.find((e:any)=>e.id===form.empleado_id)
      if(emp){
        const nuevoProg=(emp.dias_programados||0)+dias
        await sbFetch(`/rest/v1/empleados?id=eq.${form.empleado_id}`,{method:"PATCH",body:JSON.stringify({dias_programados:nuevoProg})})
      }
    }
    setGuardando(false);setModoForm(null);setForm({});cargarVacaciones();cargarEmpleados()
  }

  const guardarIncidencia=async()=>{
    if(!form.empleado_id||!form.tipo)return
    setGuardando(true)
    await sbFetch("/rest/v1/incidencias",{method:"POST",body:JSON.stringify(form),headers:{"Prefer":"return=minimal"}})
    setGuardando(false);setModoForm(null);setForm({});cargarIncidencias()
  }

  const eliminar=async(tabla:string,id:string,reload:()=>void)=>{
    if(!window.confirm("¿Eliminar este registro?"))return
    // Si es vacacion, restaurar días al empleado
    if(tabla==="vacaciones"){
      const vac=vacaciones.find((v:any)=>v.id===id)
      if(vac){
        const emp=empleados.find((e:any)=>e.id===vac.empleado_id)
        if(emp){
          const nuevoProg=Math.max(0,(emp.dias_programados||0)-(vac.dias||0))
          await sbFetch(`/rest/v1/empleados?id=eq.${vac.empleado_id}`,{method:"PATCH",body:JSON.stringify({dias_programados:nuevoProg})})
        }
      }
    }
    await sbFetch(`/rest/v1/${tabla}?id=eq.${id}`,{method:"DELETE"})
    reload()
    if(tabla==="vacaciones") cargarEmpleados()
  }

  const actualizarEstadoVac=async(id:string,estado:string)=>{
    // Si se rechaza, restaurar días al empleado
    if(estado==="rechazado"){
      const vac=vacaciones.find((v:any)=>v.id===id)
      if(vac){
        const emp=empleados.find((e:any)=>e.id===vac.empleado_id)
        if(emp){
          const nuevoProg=Math.max(0,(emp.dias_programados||0)-(vac.dias||0))
          await sbFetch(`/rest/v1/empleados?id=eq.${vac.empleado_id}`,{method:"PATCH",body:JSON.stringify({dias_programados:nuevoProg})})
        }
      }
    }
    await sbFetch(`/rest/v1/vacaciones?id=eq.${id}`,{method:"PATCH",body:JSON.stringify({estado})})
    cargarVacaciones();cargarEmpleados()
  }

  const empFilt=empleados.filter(e=>!busq||(e.nombre||"").toLowerCase().includes(busq.toLowerCase())||(e.nss||"").includes(busq))

  const DIAS_ES=["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"]
  const fmtFecha=(f:string)=>{
    if(!f)return "—"
    const d=new Date(f+"T12:00:00")
    return `${DIAS_ES[d.getDay()]} ${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`
  }

  const vacPend=vacaciones.filter(v=>v.estado==="pendiente")
  const vacFilt=filtroVac==="todos"?vacaciones:vacaciones.filter(v=>v.estado===filtroVac)

  const TABS=[
    {id:"empleados",label:"👥 Empleados",badge:empleados.length},
    {id:"vacaciones",label:"🏖️ Vacaciones",badge:vacPend.length>0?vacPend.length:null,badgeColor:"#8b2e2e"},
    {id:"incidencias",label:"⚠️ Incidencias",badge:null},
    {id:"nomina",label:"💰 Nómina",badge:null},
  ]

  // ── LOGIN ──
  if(!auth) return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:400}}>
      <div style={{background:"#fff",borderRadius:16,padding:40,width:"min(360px,90vw)",boxShadow:"0 4px 24px rgba(0,0,0,.10)",textAlign:"center" as const}}>
        <div style={{fontSize:32,marginBottom:8}}>👥</div>
        <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:700,marginBottom:4}}>Recursos Humanos</div>
        <div style={{fontSize:12,color:"#9a9590",marginBottom:24}}>Módulo restringido — ingresa la contraseña de gerencia</div>
        <input type="password" value={pwdInput} onChange={e=>setPwdInput(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter"){if(pwdInput===RH_PWD){setAuth(true);setPwdError("")}else setPwdError("Contraseña incorrecta")}}}
          placeholder="Contraseña..."
          style={{width:"100%",padding:"10px 14px",border:`1.5px solid ${pwdError?"#fca5a5":"#e8e5de"}`,borderRadius:9,fontSize:14,outline:"none",boxSizing:"border-box" as const,marginBottom:8,textAlign:"center" as const}}/>
        {pwdError&&<div style={{fontSize:12,color:"#8b2e2e",marginBottom:8}}>❌ {pwdError}</div>}
        <button onClick={()=>{if(pwdInput===RH_PWD){setAuth(true);setPwdError("")}else setPwdError("Contraseña incorrecta")}}
          style={{width:"100%",padding:"10px",borderRadius:9,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontSize:14,fontWeight:700}}>
          Entrar
        </button>
      </div>
    </div>
  )

  return(
    <div style={{maxWidth:1100,margin:"0 auto",padding:isMobile?"10px":"0 0 40px"}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap" as const,gap:10}}>
        <div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:700}}>👥 Recursos Humanos</div>
          <div style={{fontSize:12,color:"#9a9590",marginTop:2}}>{empleados.length} empleados · {vacPend.length} vacaciones pendientes</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          {tab==="empleados"&&<button onClick={()=>{setModoForm("emp");setForm({})}}
            style={{padding:"8px 16px",borderRadius:8,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700}}>+ Empleado</button>}
          {tab==="vacaciones"&&<button onClick={()=>{setModoForm("vac");setForm({})}}
            style={{padding:"8px 16px",borderRadius:8,background:"#1a3a5c",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700}}>+ Vacaciones</button>}
          {tab==="incidencias"&&<button onClick={()=>{setModoForm("inc");setForm({})}}
            style={{padding:"8px 16px",borderRadius:8,background:"#92580a",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700}}>+ Incidencia</button>}
        </div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",gap:4,marginBottom:16,borderBottom:"2px solid #f0ece4",paddingBottom:0}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)}
            style={{padding:"8px 16px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontWeight:tab===t.id?700:400,
              color:tab===t.id?"#1a1814":"#9a9590",borderBottom:`2px solid ${tab===t.id?"#1a1814":"transparent"}`,
              marginBottom:-2,position:"relative" as const,display:"flex",alignItems:"center",gap:6}}>
            {t.label}
            {t.badge!=null&&<span style={{background:t.badgeColor||"#4a2d6e",color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:700}}>{t.badge}</span>}
          </button>
        ))}
      </div>

      {/* Modal Form */}
      {modoForm&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.5)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"#fff",borderRadius:16,padding:28,width:"min(520px,100%)",maxHeight:"90vh",overflowY:"auto" as const}}>

            {/* Empleado Form */}
            {modoForm==="emp"&&(
              <>
                <div style={{fontWeight:700,fontSize:16,marginBottom:16}}>{form.id?"✏️ Editar empleado":"👤 Nuevo empleado"}</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  {[
                    {k:"nombre",l:"Nombre completo",full:true},
                    {k:"nss",l:"NSS"},
                    {k:"curp",l:"CURP"},
                    {k:"puesto",l:"Puesto"},
                    {k:"departamento",l:"Departamento"},
                    {k:"fecha_ingreso",l:"Fecha ingreso",type:"date"},
                    {k:"dias_vacaciones",l:"Días vacaciones",type:"number"},
                    {k:"dias_programados",l:"Días programados",type:"number"},
                  ].map((f:any)=>(
                    <div key={f.k} style={{gridColumn:f.full?"1/-1":"auto"}}>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>{f.l}</label>
                      <input type={f.type||"text"} value={form[f.k]||""}
                        onChange={e=>setForm((p:any)=>({...p,[f.k]:f.type==="number"?parseInt(e.target.value)||0:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
                    </div>
                  ))}
                  <div style={{gridColumn:"1/-1"}}>
                    <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Notas</label>
                    <textarea value={form.notas||""} onChange={e=>setForm((p:any)=>({...p,notas:e.target.value}))} rows={2}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}/>
                  </div>
                </div>
              </>
            )}

            {/* Vacaciones Form */}
            {modoForm==="vac"&&(
              <>
                <div style={{fontWeight:700,fontSize:16,marginBottom:16}}>🏖️ Registrar vacaciones</div>
                <div style={{display:"flex",flexDirection:"column" as const,gap:10}}>
                  <div>
                    <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Empleado</label>
                    <select value={form.empleado_id||""} onChange={e=>setForm((p:any)=>({...p,empleado_id:e.target.value}))}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none"}}>
                      <option value="">Seleccionar empleado...</option>
                      {empleados.map((e:any)=><option key={e.id} value={e.id}>{e.nombre}</option>)}
                    </select>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Fecha inicio</label>
                      <input type="date" value={form.fecha_inicio||""} onChange={e=>setForm((p:any)=>({...p,fecha_inicio:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
                    </div>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Fecha fin</label>
                      <input type="date" value={form.fecha_fin||""} onChange={e=>setForm((p:any)=>({...p,fecha_fin:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
                    </div>
                  </div>
                  {form.fecha_inicio&&form.fecha_fin&&(
                    <div style={{padding:"8px 12px",background:"#f0fdf4",borderRadius:8,fontSize:13,color:"#2d6a4f",fontWeight:700}}>
                      📅 {calcDias(form.fecha_inicio,form.fecha_fin)} días hábiles · {fmtFecha(form.fecha_inicio)} → {fmtFecha(form.fecha_fin)}
                    </div>
                  )}
                  <div>
                    <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Notas</label>
                    <textarea value={form.notas||""} onChange={e=>setForm((p:any)=>({...p,notas:e.target.value}))} rows={2}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}/>
                  </div>
                </div>
              </>
            )}

            {/* Incidencia Form */}
            {modoForm==="inc"&&(
              <>
                <div style={{fontWeight:700,fontSize:16,marginBottom:16}}>⚠️ Registrar incidencia</div>
                <div style={{display:"flex",flexDirection:"column" as const,gap:10}}>
                  <div>
                    <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Empleado</label>
                    <select value={form.empleado_id||""} onChange={e=>setForm((p:any)=>({...p,empleado_id:e.target.value}))}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none"}}>
                      <option value="">Seleccionar empleado...</option>
                      {empleados.map((e:any)=><option key={e.id} value={e.id}>{e.nombre}</option>)}
                    </select>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Tipo</label>
                      <select value={form.tipo||""} onChange={e=>setForm((p:any)=>({...p,tipo:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none"}}>
                        <option value="">Seleccionar...</option>
                        {["Falta","Retardo","Permiso con goce","Permiso sin goce","Incapacidad","Sanción","Otro"].map(t=><option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>
                    <div>
                      <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Fecha</label>
                      <input type="date" value={form.fecha||""} onChange={e=>setForm((p:any)=>({...p,fecha:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
                    </div>
                  </div>
                  <div>
                    <label style={{fontSize:10,fontWeight:700,color:"#4a4640",display:"block",marginBottom:3}}>Descripción</label>
                    <textarea value={form.descripcion||""} onChange={e=>setForm((p:any)=>({...p,descripcion:e.target.value}))} rows={3}
                      style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e8e5de",borderRadius:7,fontSize:13,outline:"none",resize:"none" as const,boxSizing:"border-box" as const}}/>
                  </div>
                </div>
              </>
            )}

            <div style={{display:"flex",gap:8,marginTop:20,justifyContent:"flex-end" as const}}>
              <button onClick={()=>{setModoForm(null);setForm({})}}
                style={{padding:"8px 18px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:13}}>Cancelar</button>
              <button onClick={modoForm==="emp"?guardarEmpleado:modoForm==="vac"?guardarVacacion:guardarIncidencia} disabled={guardando}
                style={{padding:"8px 22px",borderRadius:8,background:"#1a1814",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700}}>
                {guardando?"Guardando...":"Guardar"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── EMPLEADOS TAB ── */}
      {tab==="empleados"&&(
        <div>
          <input value={busq} onChange={e=>setBusq(e.target.value)} placeholder="Buscar por nombre o NSS..."
            style={{width:"100%",padding:"9px 14px",border:"1.5px solid #e8e5de",borderRadius:9,fontSize:13,outline:"none",marginBottom:14,boxSizing:"border-box" as const}}/>
          {cargando?<div style={{textAlign:"center" as const,padding:40,color:"#9a9590"}}>Cargando...</div>:(
            <div style={{background:"#fff",borderRadius:12,border:"1px solid #e8e5de",overflow:"hidden"}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 130px 110px 70px 70px 70px 60px",background:"#f8f6f2",padding:"8px 14px",fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".04em"}}>
                <div>Nombre / CURP</div><div>NSS</div><div>Puesto</div><div>Vac.</div><div>Prog.</div><div>Pend.</div><div></div>
              </div>
              {empFilt.length===0&&<div style={{padding:30,textAlign:"center" as const,color:"#9a9590"}}>Sin empleados</div>}
              {empFilt.map((e:any)=>{
                const pend=(e.dias_vacaciones||0)-(e.dias_programados||0)
                return(
                  <div key={e.id} style={{display:"grid",gridTemplateColumns:"1fr 130px 110px 70px 70px 70px 60px",padding:"10px 14px",borderTop:"1px solid #f0ece4",alignItems:"center"}}>
                    <div>
                      <div style={{fontWeight:600,fontSize:13}}>{e.nombre}</div>
                      {e.curp&&<div style={{fontSize:10,color:"#9a9590",fontFamily:"monospace"}}>{e.curp}</div>}
                    </div>
                    <div style={{fontSize:11,fontFamily:"monospace",color:"#4a4640"}}>{e.nss||"—"}</div>
                    <div style={{fontSize:11,color:"#4a4640"}}>{e.puesto||"—"}</div>
                    <div style={{textAlign:"center" as const}}>
                      <span style={{background:"#f0fdf4",color:"#2d6a4f",padding:"2px 8px",borderRadius:10,fontWeight:700,fontSize:12}}>{e.dias_vacaciones||0}</span>
                    </div>
                    <div style={{textAlign:"center" as const}}>
                      <span style={{background:"#eff6ff",color:"#1a3a5c",padding:"2px 8px",borderRadius:10,fontWeight:700,fontSize:12}}>{e.dias_programados||0}</span>
                    </div>
                    <div style={{textAlign:"center" as const}}>
                      <span style={{background:pend<0?"#fef2f2":pend===0?"#f8f6f2":"#fffbeb",color:pend<0?"#8b2e2e":pend===0?"#9a9590":"#92580a",padding:"2px 8px",borderRadius:10,fontWeight:700,fontSize:12}}>{pend}</span>
                    </div>
                    <div style={{display:"flex",gap:4,justifyContent:"flex-end" as const}}>
                      <button onClick={()=>{setForm({...e});setModoForm("emp")}}
                        style={{padding:"3px 7px",borderRadius:5,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:11}}>✏️</button>
                      <button onClick={()=>eliminar("empleados",e.id,cargarEmpleados)}
                        style={{padding:"3px 7px",borderRadius:5,border:"1px solid #fca5a5",background:"#fef2f2",color:"#8b2e2e",cursor:"pointer",fontSize:11}}>🗑️</button>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </div>
      )}

      {/* ── VACACIONES TAB ── */}
      {tab==="vacaciones"&&(
        <div>
          {vacPend.length>0&&(
            <div style={{background:"#fffbeb",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 16px",marginBottom:14,fontSize:13,color:"#92580a",fontWeight:600}}>
              ⚠️ {vacPend.length} solicitud{vacPend.length>1?"es":""} pendiente{vacPend.length>1?"s":""} de aprobación
            </div>
          )}
          <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap" as const}}>
            {["todos","pendiente","aprobado","rechazado"].map(f=>(
              <button key={f} onClick={()=>setFiltroVac(f)}
                style={{padding:"5px 12px",borderRadius:16,border:`1.5px solid ${filtroVac===f?"#1a1814":"#e8e5de"}`,
                  background:filtroVac===f?"#1a1814":"#fff",color:filtroVac===f?"#fff":"#4a4640",
                  cursor:"pointer",fontSize:12,fontWeight:600,textTransform:"capitalize" as const}}>
                {f==="todos"?"Todos":f.charAt(0).toUpperCase()+f.slice(1)}
                {f!=="todos"&&<span style={{marginLeft:4,opacity:.7}}>({vacaciones.filter(v=>v.estado===f).length})</span>}
              </button>
            ))}
          </div>
          <div style={{background:"#fff",borderRadius:12,border:"1px solid #e8e5de",overflow:"hidden"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 120px 120px 50px 110px 100px",background:"#f8f6f2",padding:"8px 14px",fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const}}>
              <div>Empleado</div><div>Inicio</div><div>Fin</div><div>Días</div><div>Estado</div><div>Acciones</div>
            </div>
            {vacFilt.length===0&&<div style={{padding:30,textAlign:"center" as const,color:"#9a9590"}}>Sin registros</div>}
            {vacFilt.map((v:any)=>(
              <div key={v.id} style={{display:"grid",gridTemplateColumns:"1fr 120px 120px 50px 110px 100px",padding:"10px 14px",borderTop:"1px solid #f0ece4",alignItems:"center",fontSize:12}}>
                <div>
                  <div style={{fontWeight:600}}>{v.empleados?.nombre||"—"}</div>
                  {v.notas&&<div style={{fontSize:10,color:"#9a9590"}}>{v.notas}</div>}
                </div>
                <div style={{color:"#4a4640"}}>{fmtFecha(v.fecha_inicio)}</div>
                <div style={{color:"#4a4640"}}>{fmtFecha(v.fecha_fin)}</div>
                <div style={{textAlign:"center" as const,fontWeight:700,color:"#1a3a5c"}}>{v.dias||0}</div>
                <div>
                  <span style={{padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:700,
                    background:v.estado==="aprobado"?"#f0fdf4":v.estado==="rechazado"?"#fef2f2":"#fffbeb",
                    color:v.estado==="aprobado"?"#2d6a4f":v.estado==="rechazado"?"#8b2e2e":"#92580a"}}>
                    {v.estado||"pendiente"}
                  </span>
                </div>
                <div style={{display:"flex",gap:4}}>
                  {v.estado==="pendiente"&&<>
                    <button onClick={()=>actualizarEstadoVac(v.id,"aprobado")}
                      style={{padding:"2px 7px",borderRadius:5,border:"1px solid #b7deca",background:"#f0fdf4",color:"#2d6a4f",cursor:"pointer",fontSize:11,fontWeight:700}}>✓</button>
                    <button onClick={()=>actualizarEstadoVac(v.id,"rechazado")}
                      style={{padding:"2px 7px",borderRadius:5,border:"1px solid #fca5a5",background:"#fef2f2",color:"#8b2e2e",cursor:"pointer",fontSize:11,fontWeight:700}}>✗</button>
                  </>}
                  <button onClick={()=>eliminar("vacaciones",v.id,cargarVacaciones)}
                    style={{padding:"2px 7px",borderRadius:5,border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:10}}>🗑️</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── INCIDENCIAS TAB ── */}
      {tab==="incidencias"&&(
        <div>
          <div style={{background:"#fff",borderRadius:12,border:"1px solid #e8e5de",overflow:"hidden"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 120px 100px 1fr 60px",background:"#f8f6f2",padding:"8px 14px",fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const}}>
              <div>Empleado</div><div>Fecha</div><div>Tipo</div><div>Descripción</div><div></div>
            </div>
            {incidencias.length===0&&<div style={{padding:30,textAlign:"center" as const,color:"#9a9590"}}>Sin incidencias registradas</div>}
            {incidencias.map((inc:any)=>(
              <div key={inc.id} style={{display:"grid",gridTemplateColumns:"1fr 120px 100px 1fr 60px",padding:"10px 14px",borderTop:"1px solid #f0ece4",alignItems:"center",fontSize:12}}>
                <div style={{fontWeight:600}}>{inc.empleados?.nombre||"—"}</div>
                <div style={{color:"#4a4640"}}>{fmtFecha(inc.fecha)}</div>
                <div>
                  <span style={{padding:"2px 8px",borderRadius:10,background:"#fffbeb",color:"#92580a",fontWeight:700,fontSize:11}}>{inc.tipo||"—"}</span>
                </div>
                <div style={{color:"#4a4640",fontSize:11}}>{inc.descripcion||"—"}</div>
                <div>
                  <button onClick={()=>eliminar("incidencias",inc.id,cargarIncidencias)}
                    style={{padding:"2px 7px",borderRadius:5,border:"1px solid #fca5a5",background:"#fef2f2",color:"#8b2e2e",cursor:"pointer",fontSize:10}}>🗑️</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── NÓMINA TAB — PLACEHOLDER ── */}
      {tab==="nomina"&&(
        <div style={{textAlign:"center" as const,padding:"60px 20px"}}>
          <div style={{fontSize:48,marginBottom:16}}>💰</div>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:24,fontWeight:700,marginBottom:8}}>Módulo de Nómina</div>
          <div style={{fontSize:14,color:"#9a9590",maxWidth:400,margin:"0 auto",lineHeight:1.6}}>
            Este módulo está en construcción. Utilizará el catálogo de empleados de RH como base para el procesamiento de nómina.
          </div>
          <div style={{marginTop:24,display:"flex",gap:10,justifyContent:"center" as const,flexWrap:"wrap" as const}}>
            {["Cálculo de nómina","Percepciones y deducciones","IMSS / INFONAVIT","Recibos de pago","Timbrado CFDI","Reportes fiscales"].map(f=>(
              <span key={f} style={{padding:"6px 14px",borderRadius:20,background:"#f0ece4",color:"#9a9590",fontSize:12,fontWeight:600}}>{f}</span>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}


function InicioSection({contratos,esAdmin,vendedorActual,token}:{contratos:Contrato[],esAdmin:boolean,vendedorActual:string,token:string}){
  const [vistaMode,setVistaMode]=useState<"semana"|"mes">("semana")
  const [alertaVisible,setAlertaVisible]=useState(true)
  const [semOff,setSemOff]=useState(0)
  const [mesOff,setMesOff]=useState(0)
  const [diaFocus,setDiaFocus]=useState<string|null>(null)
  const [fichaContrato,setFichaContrato]=useState<Contrato|null>(null)
  const [listaModal,setListaModal]=useState<{titulo:string,items:Contrato[],tipo:string}|null>(null)

  const hoy=new Date();hoy.setHours(0,0,0,0)
  const hoyStr=isoDate(hoy)
  const DIAS_L=["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
  const MESES_N=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
  const MESES_C=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

  // Semana
  const dow=hoy.getDay()===0?6:hoy.getDay()-1
  const lunesBase=new Date(hoy);lunesBase.setDate(hoy.getDate()-dow)
  const lunesActivo=new Date(lunesBase);lunesActivo.setDate(lunesBase.getDate()+semOff*7)
  const dias7=Array.from({length:7},(_,i)=>{const d=new Date(lunesActivo);d.setDate(lunesActivo.getDate()+i);return d})
  const isoDomSem=isoDate(dias7[6])

  // Mes
  const mesRef=new Date(hoy.getFullYear(),hoy.getMonth()+mesOff,1)
  const diasMes=new Date(mesRef.getFullYear(),mesRef.getMonth()+1,0).getDate()
  const primerDiaMes=new Date(mesRef.getFullYear(),mesRef.getMonth(),1)
  const dowMes=primerDiaMes.getDay()===0?6:primerDiaMes.getDay()-1
  const diasGrid:Array<Date|null>=Array(dowMes).fill(null)
  for(let i=1;i<=diasMes;i++) diasGrid.push(new Date(mesRef.getFullYear(),mesRef.getMonth(),i))
  while(diasGrid.length%7!==0) diasGrid.push(null)

  const diaActivo=diaFocus||hoyStr
  const rangoInicio=vistaMode==="semana"?isoDate(lunesActivo):isoDate(primerDiaMes)
  const rangoFin=vistaMode==="semana"?isoDomSem:isoDate(new Date(mesRef.getFullYear(),mesRef.getMonth()+1,0))

  const eventosDia=(f:string)=>contratos.filter(x=>x.fecha_entrega===f||x.fecha_evento===f||x.fecha_desmonte===f)
  const entregasDia=contratos.filter(x=>x.fecha_entrega===diaActivo)
  const eventosDiaAct=contratos.filter(x=>x.fecha_evento===diaActivo)
  const desmontesDia=contratos.filter(x=>x.fecha_desmonte===diaActivo)

  const contratosRango=contratos.filter(x=>
    (x.fecha_evento>=rangoInicio&&x.fecha_evento<=rangoFin)||
    (x.fecha_entrega>=rangoInicio&&x.fecha_entrega<=rangoFin)||
    (x.fecha_desmonte>=rangoInicio&&x.fecha_desmonte<=rangoFin)
  )

  const [cotsPend,setCotsPend]=useState<any[]>([])
  useState(()=>{
    fetch("/api/cotizaciones",{headers:{Authorization:`Bearer ${token}`}})
      .then(r=>r.json()).then(data=>{
        const all=Array.isArray(data)?data:[]
        setCotsPend(all.filter((x:any)=>!["rechazada","expirada","convertida"].includes(x.estado)))
      })
  })
  const cotsPorVencer=cotsPend.filter((x:any)=>{
    if(!x.fecha_vigencia)return false
    const d=Math.ceil((new Date(x.fecha_vigencia+"T23:59:59").getTime()-Date.now())/86400000)
    return d>=0&&d<=7
  })

  // Cotizaciones aceptadas con evento la próxima semana que AÚN no son contrato
  const proxLunes=new Date(hoy);proxLunes.setDate(hoy.getDate()+(7-dow))
  const proxDomingo=new Date(proxLunes);proxDomingo.setDate(proxLunes.getDate()+6)
  const isoProxLunes=isoDate(proxLunes)
  const isoProxDomingo=isoDate(proxDomingo)

  // Folios que ya tienen contrato confirmado
  const foliosConContrato=new Set(
    contratos.filter(x=>(x.tipo||"contrato")==="contrato").map(x=>x.folio).filter(Boolean)
  )

  // Cotizaciones del SISTEMA (tabla cotizaciones) con evento próximo sin convertir
  const cotsSistemaPorConvertir=cotsPend.filter((x:any)=>{
    if(!x.fecha_evento) return false
    const diasEvento=Math.ceil((new Date(x.fecha_evento+"T12:00:00").getTime()-Date.now())/86400000)
    if(diasEvento<0||diasEvento>14) return false
    if(x.estado==="convertida") return false
    if(foliosConContrato.has(x.folio)) return false
    return true
  })

  // Cotizaciones Excel (tipo="cotizacion") con evento en próximos 14 días
  const cotsExcelPorConvertir=contratos
    .filter(x=>{
      // Solo tipo cotizacion o pendiente
      const tipo=(x.tipo||"").trim().toLowerCase()
      if(tipo!=="cotizacion"&&tipo!=="pendiente") return false
      if(!x.fecha_evento) return false
      const diasEvento=Math.ceil((new Date(x.fecha_evento+"T12:00:00").getTime()-Date.now())/86400000)
      if(diasEvento<0||diasEvento>14) return false
      return true
    })
    .map(x=>({
      id:x.id,
      folio:x.folio||x.archivo||"",
      cliente_nombre:x.cliente||x.archivo||"",
      lugar_evento:x.lugar||"",
      fecha_evento:x.fecha_evento||"",
      vendedor:x.vendedor||vendedorDesdeFolio(x.folio||""),
      total:x.total||0,
      estado:"cotizacion",
      _fromExcel:true,
    }))

  // Unir ambas y ordenar por fecha de evento
  const cotsPorConvertir=[...cotsSistemaPorConvertir,...cotsExcelPorConvertir]
    .sort((a:any,b:any)=>(a.fecha_evento||"").localeCompare(b.fecha_evento||""))

  const exportarExcel=()=>{
    const rows=[
      ["Fecha Evento","Fecha Entrega","Fecha Desmonte","Cliente","Lugar","Telefono","Vendedor","Articulos","Total","Cobrado","Saldo"],
      ...contratosRango.map(x=>[
        x.fecha_evento||"",x.fecha_entrega||"",x.fecha_desmonte||"",
        x.cliente||x.archivo||"",x.lugar||"",x.tel||"",x.vendedor||"",
        (x.articulos||[]).reduce((s:number,a:Articulo)=>s+(a.cantidad||0),0),
        x.total||0,x.cobrado||0,(x.total||0)-(x.cobrado||0)
      ])
    ]
    const esc=(v:any)=>{const s=String(v||"");return s.includes(",")?"'"+s+"'":s}
    const csv=rows.map(r=>r.map(esc).join(",")).join("\n")
    const blob=new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8"})
    const url=URL.createObjectURL(blob)
    const a=document.createElement("a");a.href=url
    a.download="Poliflor_"+rangoInicio+".csv"
    a.click();URL.revokeObjectURL(url)
  }

  // ── FICHA CONTRATO MODAL ──────────────────────────────────────────
  const FichaModal=({x,onClose}:{x:Contrato,onClose:()=>void})=>{
    const totalArts=(x.articulos||[]).reduce((s:number,a:Articulo)=>s+(a.cantidad||0),0)
    const saldo=(x.total||0)-(x.cobrado||0)
    return(
      <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.6)",zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
        onClick={e=>{if(e.target===e.currentTarget)onClose()}}>
        <div style={{background:"#fff",borderRadius:16,width:640,maxHeight:"88vh",display:"flex",flexDirection:"column" as const,boxShadow:"0 32px 80px rgba(0,0,0,.3)"}}>
          {/* Header */}
          <div style={{background:"#0f172a",borderRadius:"16px 16px 0 0",padding:"16px 20px",display:"flex",alignItems:"flex-start",gap:12}}>
            <div style={{flex:1}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,color:"#fff"}}>{x.cliente||x.archivo}</div>
              <div style={{fontSize:11,color:"rgba(255,255,255,.5)",marginTop:3}}>
                {x.folio&&<span style={{marginRight:10,fontFamily:"monospace"}}>{x.folio}</span>}
                {x.vendedor&&<span>👤 {x.vendedor}</span>}
              </div>
            </div>
            <button onClick={onClose} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
          </div>
          {/* Info grid */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:0,borderBottom:"1px solid #e8e5de"}}>
            {[
              {l:"Evento",v:x.fecha_evento||"—",ico:"🎉"},
              {l:"Entrega",v:x.fecha_entrega||"—",ico:"🚚"},
              {l:"Desmonte",v:x.fecha_desmonte||"—",ico:"📦"},
            ].map((k,i)=>(
              <div key={i} style={{padding:"12px 16px",borderRight:i<2?"1px solid #e8e5de":"none",textAlign:"center" as const}}>
                <div style={{fontSize:18,marginBottom:4}}>{k.ico}</div>
                <div style={{fontSize:12,fontWeight:700,color:"#1a1814"}}>{k.v}</div>
                <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".05em"}}>{k.l}</div>
              </div>
            ))}
          </div>
          {/* Body */}
          <div style={{flex:1,overflowY:"auto" as const,padding:"14px 18px"}}>
            {/* Contacto */}
            <div style={{display:"flex",gap:16,marginBottom:14,flexWrap:"wrap" as const}}>
              {x.lugar&&<div style={{fontSize:12,color:"#4a4640"}}>📍 {x.lugar}</div>}
              {x.tel&&<div style={{fontSize:12,color:"#4a4640"}}>📞 {x.tel}</div>}
            </div>
            {/* Piezas asignadas */}
            {((x.asig_entrega||[]).length>0||(x.asig_desmonte||[]).length>0)&&(
              <div style={{display:"flex",gap:8,marginBottom:14}}>
                {(x.asig_entrega||[]).length>0&&(
                  <div style={{background:"#edf3fa",borderRadius:8,padding:"6px 10px",fontSize:11}}>
                    <span style={{fontWeight:700,color:"#1a3a5c"}}>🚚 Entrega: </span>
                    <span style={{color:"#1a3a5c"}}>{(x.asig_entrega||[]).join(", ")}</span>
                  </div>
                )}
                {(x.asig_desmonte||[]).length>0&&(
                  <div style={{background:"#f5f0fc",borderRadius:8,padding:"6px 10px",fontSize:11}}>
                    <span style={{fontWeight:700,color:"#4a2d6e"}}>📦 Desmonte: </span>
                    <span style={{color:"#4a2d6e"}}>{(x.asig_desmonte||[]).join(", ")}</span>
                  </div>
                )}
              </div>
            )}
            {/* Artículos */}
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>
                Artículos — {totalArts} piezas en total
              </div>
              <div style={{border:"1px solid #e8e5de",borderRadius:10,overflow:"hidden"}}>
                {(x.articulos||[]).length===0?(
                  <div style={{padding:"16px",fontSize:11,color:"#c4bfb8",textAlign:"center" as const,fontStyle:"italic"}}>Sin artículos registrados</div>
                ):(
                  <table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:12}}>
                    <thead>
                      <tr style={{background:"#fafaf8"}}>
                        <th style={{padding:"7px 12px",textAlign:"left" as const,fontSize:10,fontWeight:700,color:"#9a9590",borderBottom:"1px solid #e8e5de"}}>Artículo</th>
                        <th style={{padding:"7px 12px",textAlign:"center" as const,fontSize:10,fontWeight:700,color:"#9a9590",borderBottom:"1px solid #e8e5de",width:60}}>Cant.</th>
                        <th style={{padding:"7px 12px",textAlign:"right" as const,fontSize:10,fontWeight:700,color:"#9a9590",borderBottom:"1px solid #e8e5de",width:90}}>Importe</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(x.articulos||[]).map((a:Articulo,i:number)=>(
                        <tr key={i} style={{borderBottom:"1px solid #f5f4f0",background:i%2===0?"#fff":"#fafaf8"}}>
                          <td style={{padding:"7px 12px",fontWeight:500}}>{a.nombre}</td>
                          <td style={{padding:"7px 12px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{a.cantidad}</td>
                          <td style={{padding:"7px 12px",textAlign:"right" as const,fontFamily:"monospace",color:(a.importe||0)>0?"#2d6a4f":"#c4bfb8"}}>{(a.importe||0)>0?"$"+(a.importe||0).toLocaleString("es-MX"):"—"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>
          {/* Footer con totales */}
          {(x.total||0)>0&&(
            <div style={{padding:"12px 18px",borderTop:"1px solid #e8e5de",display:"flex",gap:10,alignItems:"center"}}>
              <div style={{flex:1,display:"flex",gap:16}}>
                <div>
                  <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const}}>Total</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,color:"#1a1814"}}>${(x.total||0).toLocaleString("es-MX")}</div>
                </div>
                <div>
                  <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const}}>Cobrado</div>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,color:"#2d6a4f"}}>${(x.cobrado||0).toLocaleString("es-MX")}</div>
                </div>
                {saldo>0&&(
                  <div>
                    <div style={{fontSize:9,color:"#9a9590",textTransform:"uppercase" as const}}>Saldo</div>
                    <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,color:"#92580a"}}>${saldo.toLocaleString("es-MX")}</div>
                  </div>
                )}
              </div>
              {saldo<=0&&<div style={{padding:"6px 14px",borderRadius:8,background:"#edf7f2",color:"#2d6a4f",fontSize:11,fontWeight:700}}>✓ Liquidado</div>}
              {saldo>0&&<div style={{padding:"6px 14px",borderRadius:8,background:"#fdf5e8",color:"#92580a",fontSize:11,fontWeight:700}}>⚠️ Pendiente de pago</div>}
            </div>
          )}
        </div>
      </div>
    )
  }

  // ── TARJETA DE EVENTO ─────────────────────────────────────────────
  const TarjetaEvento=({x,tipo}:{x:Contrato,tipo:"entrega"|"evento"|"desmonte"})=>{
    const totalArts=(x.articulos||[]).reduce((s:number,a:Articulo)=>s+(a.cantidad||0),0)
    const saldo=(x.total||0)-(x.cobrado||0)
    const cols={entrega:{borde:"#1a3a5c",bg:"#edf3fa",ico:"🚚",tag:"ENTREGA"},evento:{borde:"#2d6a4f",bg:"#edf7f2",ico:"🎉",tag:"EVENTO"},desmonte:{borde:"#4a2d6e",bg:"#f5f0fc",ico:"📦",tag:"DESMONTE"}}
    const col=cols[tipo]
    const vendedor=x.vendedor||vendedorDesdeFolio(x.folio||"")
    return(
      <div onClick={()=>setFichaContrato(x)} style={{background:"#fff",border:`1px solid #e8e5de`,borderLeft:`4px solid ${col.borde}`,borderRadius:8,padding:"10px 12px",marginBottom:6,cursor:"pointer",transition:"box-shadow .15s"}}
        onMouseEnter={e=>(e.currentTarget.style.boxShadow="0 2px 12px rgba(0,0,0,.1)")}
        onMouseLeave={e=>(e.currentTarget.style.boxShadow="none")}>
        <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:col.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>{col.ico}</div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:3,flexWrap:"wrap" as const}}>
              <span style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente||x.archivo}</span>
              <span style={{fontSize:9,padding:"1px 6px",borderRadius:4,background:col.bg,color:col.borde,fontWeight:700,flexShrink:0}}>{col.tag}</span>
            </div>
            <div style={{fontSize:10,color:"#9a9590",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>📍 {x.lugar?.slice(0,45)||"Sin dirección"}</div>
            {x.tel&&<div style={{fontSize:10,color:"#4a4640",marginTop:1}}>📞 {x.tel}</div>}
            <div style={{display:"flex",gap:6,marginTop:5,flexWrap:"wrap" as const}}>
              {x.folio&&<span style={{fontSize:9,background:"#f5f4f0",color:"#9a9590",padding:"1px 5px",borderRadius:4,fontFamily:"monospace"}}>{x.folio}</span>}
              {vendedor&&<span style={{fontSize:9,background:"#fafaf8",color:"#4a4640",padding:"1px 5px",borderRadius:4}}>👤 {vendedor}</span>}
              {totalArts>0&&<span style={{fontSize:9,background:col.bg,color:col.borde,padding:"1px 5px",borderRadius:4,fontWeight:600}}>📦 {totalArts} pzas</span>}
              {(x.asig_entrega||[]).length>0&&tipo==="entrega"&&<span style={{fontSize:9,background:"#edf7f2",color:"#2d6a4f",padding:"1px 5px",borderRadius:4}}>🚛 {(x.asig_entrega||[]).join(", ")}</span>}
              {!(x.asig_entrega||[]).length&&tipo==="entrega"&&<span style={{fontSize:9,background:"#fdf0f0",color:"#8b2e2e",padding:"1px 5px",borderRadius:4,fontWeight:700}}>⚠️ Sin asignar</span>}
            </div>
          </div>
          <div style={{textAlign:"right" as const,flexShrink:0}}>
            {(x.total||0)>0&&<div style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1a1814"}}>${(x.total||0).toLocaleString("es-MX")}</div>}
            {saldo>0&&<div style={{fontSize:9,color:"#92580a",fontWeight:700,marginTop:1}}>Debe ${saldo.toLocaleString("es-MX")}</div>}
            {saldo<=0&&(x.total||0)>0&&<div style={{fontSize:9,color:"#2d6a4f",marginTop:1}}>✓ Pagado</div>}
          </div>
        </div>
      </div>
    )
  }

  // ── PANEL DÍA con agrupación por vendedor ────────────────────────
  const PanelDia=({titulo,items,tipo,color}:{titulo:string,items:Contrato[],tipo:"entrega"|"evento"|"desmonte",color:string})=>{
    // Agrupar por vendedor
    const grupos:Record<string,Contrato[]>={}
    const SIN_VEND="Sin vendedor"
    items.forEach(x=>{
      const v=x.vendedor||vendedorDesdeFolio(x.folio||"")||SIN_VEND
      if(!grupos[v])grupos[v]=[]
      grupos[v].push(x)
    })
    const vendedoresOrden=VENDEDORES.map(v=>v.nombre).filter(v=>grupos[v]).concat(grupos[SIN_VEND]?[SIN_VEND]:[])
    return(
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
        <div style={{background:color,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:15}}>{tipo==="entrega"?"🚚":tipo==="evento"?"🎉":"📦"}</span>
          <span style={{fontSize:13,fontWeight:700,color:"#fff",flex:1}}>{titulo}</span>
          <span style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:"#fff"}}>{items.length}</span>
        </div>
        <div style={{padding:"8px 10px",maxHeight:420,overflowY:"auto" as const}}>
          {items.length===0
            ?<div style={{fontSize:11,color:"#c4bfb8",padding:"16px",fontStyle:"italic",textAlign:"center" as const}}>Sin {titulo.toLowerCase()}</div>
            :vendedoresOrden.map(v=>(
              <div key={v}>
                <div style={{fontSize:9,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".07em",padding:"6px 4px 3px",display:"flex",alignItems:"center",gap:5}}>
                  <div style={{width:16,height:16,borderRadius:"50%",background:v===SIN_VEND?"#e8e5de":"#1a1814",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:800,flexShrink:0}}>
                    {v.charAt(0)}
                  </div>
                  {v} <span style={{background:"#f5f4f0",padding:"0 5px",borderRadius:4,color:"#4a4640"}}>{grupos[v].length}</span>
                </div>
                {grupos[v].map((x,i)=><TarjetaEvento key={i} x={x} tipo={tipo}/>)}
              </div>
            ))
          }
        </div>
      </div>
    )
  }

  const periodoLabel=vistaMode==="semana"
    ?`${isoDate(lunesActivo)} al ${isoDomSem}`
    :MESES_N[mesRef.getMonth()]+" "+mesRef.getFullYear()
  const semLabel=semOff===0?"Esta semana":semOff===-1?"Semana pasada":semOff===1?"Próxima semana":`${semOff>0?"+":""}${semOff} semanas`

  return(
    <div style={{display:"flex",flexDirection:"column" as const,gap:14}}>

      {/* ── MODAL FICHA ── */}
      {fichaContrato&&<FichaModal x={fichaContrato} onClose={()=>setFichaContrato(null)}/>}

      {/* ── HEADER NAV ── */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"12px 16px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap" as const}}>
        <div style={{display:"flex",gap:3,background:"#f5f4f0",borderRadius:8,padding:3}}>
          {([["semana","📅 Semana"],["mes","📆 Mes"]] as [string,string][]).map(([m,l])=>(
            <button key={m} onClick={()=>{setVistaMode(m as any);setDiaFocus(null)}}
              style={{padding:"5px 14px",borderRadius:6,border:"none",background:vistaMode===m?"#1a1814":"transparent",color:vistaMode===m?"#fff":"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:vistaMode===m?700:400}}>
              {l}
            </button>
          ))}
        </div>
        <button onClick={()=>{vistaMode==="semana"?setSemOff(semOff-1):setMesOff(mesOff-1);setDiaFocus(null)}}
          style={{width:32,height:32,borderRadius:"50%",border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>‹</button>
        <div style={{flex:1,textAlign:"center" as const}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:800}}>{periodoLabel}</div>
          <div style={{fontSize:10,color:"#9a9590",marginTop:1}}>{vistaMode==="semana"?semLabel:""} · {contratosRango.length} eventos</div>
        </div>
        <button onClick={()=>{vistaMode==="semana"?setSemOff(semOff+1):setMesOff(mesOff+1);setDiaFocus(null)}}
          style={{width:32,height:32,borderRadius:"50%",border:"1px solid #e8e5de",background:"#fff",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>›</button>
        <button onClick={()=>{vistaMode==="semana"?setSemOff(0):setMesOff(0);setDiaFocus(null)}}
          style={{padding:"5px 14px",borderRadius:8,border:"1px solid #e8e5de",background:"#fff",color:"#4a4640",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>Hoy</button>
        <button onClick={exportarExcel}
          style={{padding:"5px 12px",borderRadius:8,border:"1px solid #2d6a4f",background:"#edf7f2",color:"#2d6a4f",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontSize:11,fontWeight:700}}>
          📊 Exportar
        </button>
      </div>

      {/* ── CALENDARIO SEMANA ── */}
      {vistaMode==="semana"&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8}}>
          {dias7.map((dia,i)=>{
            const ds=isoDate(dia)
            const ents=contratos.filter(x=>x.fecha_entrega===ds)
            const evts=contratos.filter(x=>x.fecha_evento===ds)
            const dess=contratos.filter(x=>x.fecha_desmonte===ds)
            const total=ents.length+evts.length+dess.length
            const isHoy=ds===hoyStr
            const isSel=ds===diaActivo
            const isPast=ds<hoyStr&&!isHoy
            return(
              <div key={ds} onClick={()=>setDiaFocus(isSel&&!isHoy?null:ds)}
                style={{background:isSel?"#0f172a":isHoy?"#fff8f0":"#fff",border:`2px solid ${isSel?"#0f172a":isHoy?"#92580a":"#e8e5de"}`,borderRadius:12,padding:"10px 6px",cursor:"pointer",opacity:isPast?0.65:1,transition:"all .15s",boxShadow:isSel?"0 4px 16px rgba(15,23,42,.2)":"none"}}>
                <div style={{fontSize:9,fontWeight:700,letterSpacing:".07em",color:isSel?"rgba(255,255,255,.5)":isHoy?"#92580a":"#9a9590",textAlign:"center" as const}}>{DIAS_L[i]}</div>
                <div style={{fontFamily:"Playfair Display,serif",fontSize:24,fontWeight:800,color:isSel?"#fff":isHoy?"#92580a":"#1a1814",textAlign:"center" as const,lineHeight:1.1,margin:"4px 0"}}>{dia.getDate()}</div>
                <div style={{display:"flex",flexDirection:"column" as const,gap:2}}>
                  {ents.length>0&&<div style={{fontSize:8,padding:"2px 4px",borderRadius:4,background:isSel?"rgba(255,255,255,.15)":"#edf3fa",color:isSel?"#fff":"#1a3a5c",textAlign:"center" as const,fontWeight:700}}>🚚{ents.length}</div>}
                  {evts.length>0&&<div style={{fontSize:8,padding:"2px 4px",borderRadius:4,background:isSel?"rgba(255,255,255,.15)":"#edf7f2",color:isSel?"#fff":"#2d6a4f",textAlign:"center" as const,fontWeight:700}}>🎉{evts.length}</div>}
                  {dess.length>0&&<div style={{fontSize:8,padding:"2px 4px",borderRadius:4,background:isSel?"rgba(255,255,255,.15)":"#f5f0fc",color:isSel?"#fff":"#4a2d6e",textAlign:"center" as const,fontWeight:700}}>📦{dess.length}</div>}
                  {total===0&&<div style={{fontSize:9,color:isSel?"rgba(255,255,255,.2)":"#e8e5de",textAlign:"center" as const}}>—</div>}
                </div>
              </div>
            )
          })}
        </div>
      )}

      {/* ── CALENDARIO MES ── */}
      {vistaMode==="mes"&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",borderBottom:"1px solid #e8e5de"}}>
            {DIAS_L.map(d=><div key={d} style={{padding:"8px 4px",textAlign:"center" as const,fontSize:10,fontWeight:700,color:"#9a9590",letterSpacing:".06em"}}>{d}</div>)}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)"}}>
            {diasGrid.map((dia,i)=>{
              if(!dia) return <div key={"e"+i} style={{padding:8,minHeight:72,borderRight:"1px solid #f5f4f0",borderBottom:"1px solid #f5f4f0",background:"#fafaf8"}}/>
              const ds=isoDate(dia)
              const ents=contratos.filter(x=>x.fecha_entrega===ds)
              const evts=contratos.filter(x=>x.fecha_evento===ds)
              const dess=contratos.filter(x=>x.fecha_desmonte===ds)
              const total=ents.length+evts.length+dess.length
              const isHoy=ds===hoyStr
              const isSel=ds===diaActivo
              const isPast=ds<hoyStr&&!isHoy
              return(
                <div key={ds} onClick={()=>setDiaFocus(isSel&&!isHoy?null:ds)}
                  style={{padding:"6px 8px",minHeight:72,borderRight:"1px solid #f5f4f0",borderBottom:"1px solid #f5f4f0",cursor:total>0||isHoy?"pointer":"default",background:isSel?"#0f172a":isHoy?"#fff8f0":"#fff",opacity:isPast?0.7:1,transition:"background .1s"}}>
                  <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:800,color:isSel?"#fff":isHoy?"#92580a":isPast?"#c4bfb8":"#1a1814",marginBottom:3}}>{dia.getDate()}</div>
                  <div style={{display:"flex",flexDirection:"column" as const,gap:1}}>
                    {ents.length>0&&<div style={{fontSize:8,padding:"1px 4px",borderRadius:3,background:isSel?"rgba(255,255,255,.15)":"#edf3fa",color:isSel?"#fff":"#1a3a5c",fontWeight:700}}>🚚{ents.length}</div>}
                    {evts.length>0&&<div style={{fontSize:8,padding:"1px 4px",borderRadius:3,background:isSel?"rgba(255,255,255,.15)":"#edf7f2",color:isSel?"#fff":"#2d6a4f",fontWeight:700}}>🎉{evts.length}</div>}
                    {dess.length>0&&<div style={{fontSize:8,padding:"1px 4px",borderRadius:3,background:isSel?"rgba(255,255,255,.15)":"#f5f0fc",color:isSel?"#fff":"#4a2d6e",fontWeight:700}}>📦{dess.length}</div>}
                  </div>
                </div>
              )
            })}
          </div>
        </div>
      )}

      {/* ── DETALLE DÍA — AGRUPADO POR VENDEDOR ── */}
      {(entregasDia.length>0||eventosDiaAct.length>0||desmontesDia.length>0)&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:15,fontWeight:700,color:"#1a1814"}}>
              {new Date(diaActivo+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long"})}
            </div>
            {diaActivo===hoyStr&&<span style={{background:"#92580a",color:"#fff",fontSize:9,padding:"2px 8px",borderRadius:6,fontWeight:700}}>HOY</span>}
            <span style={{fontSize:11,color:"#9a9590"}}>{entregasDia.length+eventosDiaAct.length+desmontesDia.length} movimientos · Haz clic en cualquier tarjeta para ver el detalle</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
            <PanelDia titulo="Entregas" items={entregasDia} tipo="entrega" color="#1a3a5c"/>
            <PanelDia titulo="Eventos" items={eventosDiaAct} tipo="evento" color="#2d6a4f"/>
            <PanelDia titulo="Desmontes" items={desmontesDia} tipo="desmonte" color="#4a2d6e"/>
          </div>
        </div>
      )}
      {diaFocus&&entregasDia.length===0&&eventosDiaAct.length===0&&desmontesDia.length===0&&(
        <div style={{textAlign:"center" as const,padding:"24px",background:"#fff",borderRadius:12,border:"1px dashed #e8e5de",color:"#9a9590",fontSize:12}}>
          Sin movimientos para {diaActivo}
        </div>
      )}

      {/* ── KPIs PERÍODO ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
        {[
          {l:"Entregas",v:contratos.filter(x=>x.fecha_entrega>=rangoInicio&&x.fecha_entrega<=rangoFin).length,c:"#1a3a5c",ico:"🚚",tipo:"entrega"},
          {l:"Eventos",v:contratos.filter(x=>x.fecha_evento>=rangoInicio&&x.fecha_evento<=rangoFin).length,c:"#2d6a4f",ico:"🎉",tipo:"evento"},
          {l:"Desmontes",v:contratos.filter(x=>x.fecha_desmonte>=rangoInicio&&x.fecha_desmonte<=rangoFin).length,c:"#4a2d6e",ico:"📦",tipo:"desmonte"},
          {l:"Sin asignar",v:contratosRango.filter(x=>x.fecha_entrega>=hoyStr&&(!(x.asig_entrega||[]).length||!(x.asig_desmonte||[]).length)).length,c:"#92580a",ico:"⚠️",tipo:"sinasignar"},
        ].map((k,i)=>(
          <div key={i} onClick={()=>{
            if(k.tipo==="entrega"){const f=contratos.filter(x=>x.fecha_entrega>=rangoInicio&&x.fecha_entrega<=rangoFin);setListaModal({titulo:"Entregas del período",items:f,tipo:"entrega"})}
            else if(k.tipo==="evento"){const f=contratos.filter(x=>x.fecha_evento>=rangoInicio&&x.fecha_evento<=rangoFin);setListaModal({titulo:"Eventos del período",items:f,tipo:"evento"})}
            else if(k.tipo==="desmonte"){const f=contratos.filter(x=>x.fecha_desmonte>=rangoInicio&&x.fecha_desmonte<=rangoFin);setListaModal({titulo:"Desmontes del período",items:f,tipo:"desmonte"})}
            else{const f=contratosRango.filter(x=>x.fecha_entrega>=hoyStr&&(!(x.asig_entrega||[]).length||!(x.asig_desmonte||[]).length));setListaModal({titulo:"Contratos sin asignar",items:f,tipo:"sinasignar"})}
          }}
          style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",gap:10,cursor:"pointer",transition:"all .15s"}}
          onMouseEnter={e=>{e.currentTarget.style.border=`1px solid ${k.c}`;e.currentTarget.style.boxShadow=`0 4px 16px ${k.c}22`}}
          onMouseLeave={e=>{e.currentTarget.style.border="1px solid #e8e5de";e.currentTarget.style.boxShadow="none"}}>
            <span style={{fontSize:22}}>{k.ico}</span>
            <div style={{flex:1}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:k.c,lineHeight:1}}>{k.v}</div>
              <div style={{fontSize:10,color:"#9a9590",marginTop:2}}>{k.l}</div>
            </div>
            <div style={{fontSize:9,color:"#9a9590"}}>Ver →</div>
          </div>
        ))}
      </div>
      {/* Modal lista */}
      {listaModal&&(
        <div style={{position:"fixed" as const,inset:0,background:"rgba(0,0,0,.55)",zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
          onClick={e=>{if(e.target===e.currentTarget)setListaModal(null)}}>
          <div style={{background:"#fff",borderRadius:16,width:680,maxHeight:"85vh",display:"flex",flexDirection:"column" as const,boxShadow:"0 24px 80px rgba(0,0,0,.3)"}}>
            <div style={{background:"#0f172a",borderRadius:"16px 16px 0 0",padding:"16px 20px",display:"flex",alignItems:"center",gap:10}}>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:17,fontWeight:800,color:"#fff",flex:1}}>{listaModal.titulo}</div>
              <span style={{fontSize:13,color:"rgba(255,255,255,.5)"}}>{listaModal.items.length} registros</span>
              <button onClick={()=>setListaModal(null)} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",width:30,height:30,borderRadius:"50%",cursor:"pointer",fontSize:16}}>✕</button>
            </div>
            <div style={{overflowY:"auto" as const,flex:1}}>
              {listaModal.items.length===0
                ?<div style={{padding:32,textAlign:"center" as const,color:"#9a9590"}}>Sin registros</div>
                :listaModal.items.map((x:Contrato,i:number)=>{
                  const fecha=listaModal.tipo==="entrega"?x.fecha_entrega:listaModal.tipo==="desmonte"?x.fecha_desmonte:x.fecha_evento
                  const vendedor=x.vendedor||vendedorDesdeFolio(x.folio||"")
                  const totalArts=(x.articulos||[]).reduce((s:number,a:Articulo)=>s+(a.cantidad||0),0)
                  const saldo=(x.total||0)-(x.cobrado||0)
                  return(
                    <div key={i} onClick={()=>{setFichaContrato(x);setListaModal(null)}} style={{padding:"12px 18px",borderBottom:"1px solid #f5f4f0",display:"flex",gap:12,alignItems:"flex-start",cursor:"pointer",background:i%2===0?"#fff":"#fafaf8"}}
                      onMouseEnter={e=>(e.currentTarget.style.background="#f0ece4")}
                      onMouseLeave={e=>(e.currentTarget.style.background=i%2===0?"#fff":"#fafaf8")}>
                      <div style={{width:42,flexShrink:0,textAlign:"center" as const,background:"#0f172a",borderRadius:8,padding:"6px 4px"}}>
                        <div style={{fontSize:18,fontWeight:800,color:"#fff",lineHeight:1}}>{fecha?.slice(8)||"—"}</div>
                        <div style={{fontSize:8,color:"rgba(255,255,255,.5)",textTransform:"uppercase" as const}}>
                          {["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][(parseInt(fecha?.slice(5,7)||"1")-1)]||""}
                        </div>
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:13}}>{x.cliente||x.archivo}</div>
                        <div style={{fontSize:11,color:"#9a9590",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>📍 {x.lugar?.slice(0,45)||"—"}</div>
                        <div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap" as const}}>
                          {vendedor&&<span style={{fontSize:9,background:"#edf3fa",color:"#1a3a5c",padding:"1px 6px",borderRadius:4}}>👤 {vendedor}</span>}
                          {x.tel&&<span style={{fontSize:9,color:"#4a4640"}}>📞 {x.tel}</span>}
                          {totalArts>0&&<span style={{fontSize:9,background:"#f5f4f0",color:"#4a4640",padding:"1px 6px",borderRadius:4}}>📦 {totalArts} pzas</span>}
                          {listaModal.tipo==="sinasignar"&&<span style={{fontSize:9,background:"#fdf0f0",color:"#8b2e2e",padding:"1px 6px",borderRadius:4,fontWeight:700}}>⚠️ Sin asignar</span>}
                        </div>
                      </div>
                      <div style={{textAlign:"right" as const,flexShrink:0}}>
                        {(x.total||0)>0&&<div style={{fontFamily:"monospace",fontSize:12,fontWeight:700}}>${(x.total||0).toLocaleString("es-MX")}</div>}
                        {saldo>0&&<div style={{fontSize:10,color:"#92580a",fontWeight:600}}>Debe ${saldo.toLocaleString("es-MX")}</div>}
                        {saldo<=0&&(x.total||0)>0&&<div style={{fontSize:10,color:"#2d6a4f"}}>✓ Liq.</div>}
                      </div>
                    </div>
                  )
                })
              }
            </div>
          </div>
        </div>
      )}

      {/* ── ALERTAS ── */}
      {cotsPorVencer.length>0&&(
        <div style={{background:"#fdf5e8",border:"1px solid #e8d4b8",borderRadius:12,padding:"12px 14px"}}>
          <div style={{fontSize:11,fontWeight:700,color:"#92580a",marginBottom:8}}>⚠️ Cotizaciones por vencer esta semana ({cotsPorVencer.length})</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:6}}>
            {cotsPorVencer.map((x:any,i:number)=>{
              const dias=Math.ceil((new Date(x.fecha_vigencia+"T23:59:59").getTime()-Date.now())/86400000)
              return(
                <div key={i} style={{background:"#fff",borderRadius:8,padding:"8px 10px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{fontWeight:700,fontSize:12}}>{x.cliente_nombre||"—"}</div>
                    <div style={{fontSize:10,color:"#9a9590"}}>{x.folio} · {x.vendedor||"—"}</div>
                  </div>
                  <span style={{color:dias<=3?"#8b2e2e":"#92580a",fontWeight:700,fontSize:12,flexShrink:0}}>{dias}d</span>
                </div>
              )
            })}
          </div>
        </div>
      )}

      {/* ── COTIZACIONES POR CONVERTIR ── */}
      {cotsPorConvertir.length>0&&(
        <div style={{background:"#fff",border:"2px solid #8b2e2e",borderRadius:12,overflow:"hidden"}}>
          {/* Header */}
          <div style={{background:"#8b2e2e",padding:"12px 16px",display:"flex",alignItems:"center",gap:10,cursor:"pointer"}}
            onClick={()=>setAlertaVisible(v=>!v)}>
            <span style={{fontSize:18}}>🚨</span>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:800,color:"#fff",display:"flex",alignItems:"center",gap:8}}>
                {cotsPorConvertir.length} cotización{cotsPorConvertir.length!==1?"es":""} pendiente{cotsPorConvertir.length!==1?"s":""} de convertir a contrato
                <span style={{background:"rgba(255,255,255,.2)",borderRadius:10,padding:"1px 8px",fontSize:11}}>
                  {alertaVisible?"▲ Ocultar":"▼ Ver"}
                </span>
              </div>
              <div style={{fontSize:10,color:"rgba(255,255,255,.65)",marginTop:1}}>
                Tienen evento en los próximos 14 días — necesitan confirmarse ya
              </div>
            </div>
          </div>
          {/* Lista — colapsable */}
          {alertaVisible&&<div style={{padding:"8px 12px"}}>
            {cotsPorConvertir.map((x:any,i:number)=>{
              const diasEvento=Math.ceil((new Date(x.fecha_evento+"T12:00:00").getTime()-Date.now())/86400000)
              const urgente=diasEvento<=7
              return(
                <div key={i} style={{
                  display:"flex",alignItems:"center",gap:12,
                  padding:"10px 8px",
                  borderBottom:i<cotsPorConvertir.length-1?"1px solid #f5f4f0":"none",
                  background:urgente?"#fdf0f0":"#fff"
                }}>
                  {/* Días countdown */}
                  <div style={{
                    width:52,height:52,borderRadius:10,flexShrink:0,
                    background:urgente?"#8b2e2e":"#92580a",
                    display:"flex",flexDirection:"column" as const,
                    alignItems:"center",justifyContent:"center"
                  }}>
                    <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,color:"#fff",lineHeight:1}}>{diasEvento}</div>
                    <div style={{fontSize:8,color:"rgba(255,255,255,.7)",textTransform:"uppercase" as const,letterSpacing:".04em"}}>días</div>
                  </div>
                  {/* Info */}
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{x.cliente_nombre||"—"}</div>
                    <div style={{fontSize:11,color:"#9a9590",marginTop:1,display:"flex",gap:8,flexWrap:"wrap" as const}}>
                      {x.folio&&<span style={{fontFamily:"monospace",background:"#f5f4f0",padding:"0 5px",borderRadius:4}}>{x.folio}</span>}
                      {x.vendedor&&<span>👤 {x.vendedor}</span>}
                      {x.lugar_evento&&<span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>📍 {x.lugar_evento.slice(0,35)}</span>}
                    </div>
                    <div style={{fontSize:10,marginTop:4,display:"flex",gap:8}}>
                      <span style={{background:urgente?"#fdf0f0":"#fffbeb",color:urgente?"#8b2e2e":"#92580a",padding:"1px 6px",borderRadius:4,fontWeight:700,border:`1px solid ${urgente?"#e8b8b8":"#e8d4b8"}`}}>
                        📅 Evento: {x.fecha_evento}
                      </span>
                      <span style={{background:(x as any)._fromExcel?"#fff8f0":"#f5f4f0",color:(x as any)._fromExcel?"#92580a":"#4a4640",padding:"1px 6px",borderRadius:4,fontWeight:(x as any)._fromExcel?600:400}}>
                        {(x as any)._fromExcel?"📄 Cotización — confirmar":"Estado: "+x.estado}
                      </span>
                    </div>
                  </div>
                  {/* Total + Acción */}
                  <div style={{textAlign:"right" as const,flexShrink:0,display:"flex",flexDirection:"column" as const,gap:4,alignItems:"flex-end"}}>
                    {(x.total||0)>0&&(
                      <div style={{fontFamily:"monospace",fontSize:13,fontWeight:700,color:"#1a1814"}}>
                        ${(x.total||0).toLocaleString("es-MX")}
                      </div>
                    )}
                    <div style={{fontSize:10,color:urgente?"#8b2e2e":"#92580a",fontWeight:600}}>
                      {urgente?"⚠️ Urgente":"Próximo"}
                    </div>
                    {(x as any)._fromExcel&&(
                      <div style={{display:"flex",gap:4,flexDirection:"column" as const}}>
                        <button onClick={async(e)=>{
                          e.stopPropagation()
                          if(!window.confirm("¿Convertir a contrato?\n\n"+x.cliente_nombre+"\n"+x.fecha_evento)) return
                          await fetch("/api/contratos?id="+x.id,{
                            method:"PATCH",
                            headers:{"Content-Type":"application/json",Authorization:"Bearer "+token},
                            body:JSON.stringify({tipo:"contrato"})
                          })
                          window.location.reload()
                        }}
                          style={{fontSize:10,padding:"4px 8px",borderRadius:6,background:"#2d6a4f",color:"#fff",border:"none",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:700,whiteSpace:"nowrap" as const}}>
                          ✓ Convertir
                        </button>
                        <button onClick={async(e)=>{
                          e.stopPropagation()
                          const pwd=window.prompt("Contraseña para declinar:")
                          if(!pwd)return
                          if(pwd!=="LITA2024"){window.alert("❌ Contraseña incorrecta");return}
                          if(!window.confirm("¿Marcar como declinado?\n\n"+x.cliente_nombre+"\n"+x.fecha_evento)) return
                          await fetch("/api/contratos?id="+x.id,{
                            method:"PATCH",
                            headers:{"Content-Type":"application/json",Authorization:"Bearer "+token},
                            body:JSON.stringify({tipo:"declinado"})
                          })
                          window.location.reload()
                        }} style={{fontSize:10,padding:"4px 8px",borderRadius:6,background:"#fff",color:"#8b2e2e",border:"1px solid #8b2e2e",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:700,whiteSpace:"nowrap" as const}}>
                          ✗ Declinar
                        </button>
                      </div>
                    )}
                    {!(x as any)._fromExcel&&(
                      <button onClick={async(e)=>{
                        e.stopPropagation()
                        if(!window.confirm("¿Declinar cotización?\n\n"+x.cliente_nombre+"\n"+x.fecha_evento)) return
                        await fetch("/api/cotizaciones?id="+x.id,{
                          method:"PATCH",
                          headers:{"Content-Type":"application/json",Authorization:"Bearer "+token},
                          body:JSON.stringify({estado:"rechazada"})
                        })
                        window.location.reload()
                      }}
                        style={{fontSize:10,padding:"4px 8px",borderRadius:6,background:"#fff",color:"#8b2e2e",border:"1px solid #8b2e2e",cursor:"pointer",fontFamily:"Epilogue,sans-serif",fontWeight:700,whiteSpace:"nowrap" as const}}>
                        ✗ Declinar
                      </button>
                    )}
                  </div>
                </div>
              )
            })}
          </div>}
          {/* Footer — siempre visible */}
          <div style={{padding:"10px 16px",background:"#fdf0f0",borderTop:"1px solid #e8b8b8",fontSize:11,color:"#8b2e2e",display:"flex",alignItems:"center",gap:8}}>
            <span>💡</span>
            <span>
              Para cotizaciones <strong>del sistema</strong>: ve a <strong>Ventas → Cotizaciones</strong> y haz clic en <strong>"Convertir a contrato"</strong>.
              Para cotizaciones <strong>de Excel</strong> (📄): usa el botón <strong>✓ Convertir</strong> o <strong>✗ Declinar</strong> directamente.
            </span>
          </div>
        </div>
      )}

      {/* ── EQUIPO (solo admin) ── */}
      {esAdmin&&cotsPend.length>0&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 16px"}}>
          <div style={{fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700,marginBottom:12}}>👥 Equipo comercial</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:8}}>
            {VENDEDORES.map(v=>{
              const cotsV=cotsPend.filter((x:any)=>x.vendedor===v.nombre)
              const sinCerrar=cotsV.filter((x:any)=>x.estado==="enviada"||x.estado==="pendiente").length
              const eventosV=contratos.filter(x=>(x.vendedor||vendedorDesdeFolio(x.folio||""))===v.nombre&&x.fecha_evento>=hoyStr&&x.fecha_evento<=rangoFin)
              return(
                <div key={v.nombre} style={{background:"#fafaf8",borderRadius:10,padding:"12px",border:`1px solid ${sinCerrar>3?"#e8b8b8":"#e8e5de"}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <div style={{width:32,height:32,borderRadius:"50%",background:"#0f172a",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:14}}>{v.nombre.charAt(0)}</div>
                    <div>
                      <div style={{fontWeight:700,fontSize:13}}>{v.nombre}</div>
                      <div style={{fontSize:9,color:"#9a9590"}}>{v.prefijo}</div>
                    </div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
                    <div style={{background:"#fff",borderRadius:6,padding:"5px",textAlign:"center" as const}}>
                      <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,color:"#1a3a5c"}}>{cotsV.length}</div>
                      <div style={{fontSize:8,color:"#9a9590",textTransform:"uppercase" as const}}>cots.</div>
                    </div>
                    <div style={{background:"#fff",borderRadius:6,padding:"5px",textAlign:"center" as const}}>
                      <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:800,color:sinCerrar>3?"#8b2e2e":"#2d6a4f"}}>{sinCerrar}</div>
                      <div style={{fontSize:8,color:"#9a9590",textTransform:"uppercase" as const}}>sin cerrar</div>
                    </div>
                  </div>
                  {eventosV.length>0&&<div style={{fontSize:9,color:"#4a4640",marginTop:6,textAlign:"center" as const}}>{eventosV.length} evento{eventosV.length!==1?"s":""} próximos</div>}
                  {sinCerrar>3&&<div style={{fontSize:9,color:"#8b2e2e",marginTop:4,fontWeight:700,textAlign:"center" as const}}>⚠️ Necesita cerrar ventas</div>}
                </div>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}


// ─── INVENTARIO ────────────────────────────────────────────────────
// Cruza artículos del catálogo con contratos para detectar conflictos
function InventarioSection({contratos,token}:{contratos:Contrato[],token:string}){
  const [fechaDesde,setFechaDesde]=useState(isoDate(new Date()))
  const [fechaHasta,setFechaHasta]=useState(isoDate(new Date(Date.now()+30*86400000)))
  const [busqArt,setBusqArt]=useState("")
  const [vistaTab,setVistaTab]=useState<"conflictos"|"ocupacion"|"catalogo">("conflictos")
  const [catalogo,setCatalogo]=useState<any[]>([])
  const [cargandoCat,setCargandoCat]=useState(false)
  const fmt=(n:number)=>"$"+Math.round(n||0).toLocaleString("es-MX")

  useState(()=>{
    setCargandoCat(true)
    fetch("/api/catalogo?activo=true",{headers:{Authorization:`Bearer ${token}`}})
      .then(r=>r.json()).then(data=>{setCatalogo(Array.isArray(data)?data:[]);setCargandoCat(false)})
  })

  // Contratos en el rango
  const contratosRango=contratos.filter(x=>
    (x.tipo||"contrato")==="contrato"&&
    x.fecha_evento&&x.fecha_evento>=fechaDesde&&x.fecha_evento<=fechaHasta
  )

  // Calcular demanda por artículo en el período
  const demandaMap:Record<string,{nombre:string,demandaTotal:number,eventos:{fecha:string,cliente:string,cantidad:number}[]}>={}
  contratosRango.forEach(x=>{
    (x.articulos||[]).forEach((a:Articulo)=>{
      const k=(a.nombre||"").trim().toLowerCase()
      if(!k||k.length<3)return
      if(!demandaMap[k])demandaMap[k]={nombre:a.nombre.trim(),demandaTotal:0,eventos:[]}
      demandaMap[k].demandaTotal+=a.cantidad||0
      demandaMap[k].eventos.push({fecha:x.fecha_evento||"",cliente:x.cliente||x.archivo||"",cantidad:a.cantidad||0})
    })
  })

  // Detectar conflictos: misma fecha + mismo artículo + total > existencia
  const conflictos:{articulo:string,fecha:string,demanda:number,existencia:number,clientes:string[]}[]=[]
  Object.entries(demandaMap).forEach(([k,data])=>{
    // Agrupar por fecha
    const porFecha:Record<string,{cantidad:number,clientes:string[]}>={}
    data.eventos.forEach(e=>{
      if(!porFecha[e.fecha])porFecha[e.fecha]={cantidad:0,clientes:[]}
      porFecha[e.fecha].cantidad+=e.cantidad
      porFecha[e.fecha].clientes.push(e.cliente)
    })
    // Buscar artículo en catálogo
    const catArt=catalogo.find(c=>c.nombre.trim().toLowerCase()===k)
    const existencia=catArt?.existencia_total||0
    Object.entries(porFecha).forEach(([fecha,info])=>{
      if(existencia>0&&info.cantidad>existencia){
        conflictos.push({articulo:data.nombre,fecha,demanda:info.cantidad,existencia,clientes:info.clientes})
      }
    })
  })
  conflictos.sort((a,b)=>a.fecha.localeCompare(b.fecha))

  // Top artículos más demandados
  const topArts=Object.values(demandaMap)
    .sort((a,b)=>b.demandaTotal-a.demandaTotal)
    .slice(0,20)

  // Filtro catálogo
  const catFiltrado=catalogo
    .filter(a=>!busqArt||a.nombre.toLowerCase().includes(busqArt.toLowerCase()))
    .slice(0,50)

  return(
    <div>
      {/* Header */}
      <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,padding:"14px 16px",marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap" as const}}>
          <div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800}}>📦 Inventario</div>
            <div style={{fontSize:11,color:"#9a9590"}}>Disponibilidad y conflictos por período</div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap" as const}}>
            <div style={{display:"flex",alignItems:"center",gap:6,background:"#f5f4f0",borderRadius:8,padding:"5px 10px"}}>
              <span style={{fontSize:10,color:"#9a9590"}}>Desde</span>
              <input type="date" value={fechaDesde} onChange={e=>setFechaDesde(e.target.value)}
                style={{border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:600,outline:"none",cursor:"pointer"}}/>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:6,background:"#f5f4f0",borderRadius:8,padding:"5px 10px"}}>
              <span style={{fontSize:10,color:"#9a9590"}}>Hasta</span>
              <input type="date" value={fechaHasta} onChange={e=>setFechaHasta(e.target.value)}
                style={{border:"none",background:"transparent",fontFamily:"Epilogue,sans-serif",fontSize:12,fontWeight:600,outline:"none",cursor:"pointer"}}/>
            </div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:4}}>
            {([["conflictos","⚠️ Conflictos"],["ocupacion","📊 Ocupación"],["catalogo","📋 Catálogo"]] as [string,string][]).map(([v,l])=>(
              <button key={v} onClick={()=>setVistaTab(v as any)}
                style={{padding:"6px 12px",borderRadius:8,border:`1.5px solid ${vistaTab===v?"#1a1814":"#e8e5de"}`,background:vistaTab===v?"#1a1814":"#fff",color:vistaTab===v?"#fff":"#4a4640",fontSize:11,fontWeight:vistaTab===v?700:400,cursor:"pointer",fontFamily:"Epilogue,sans-serif"}}>
                {l}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
        {[
          {l:"Contratos en período",v:contratosRango.length,c:"#1a1814"},
          {l:"Artículos únicos demandados",v:Object.keys(demandaMap).length,c:"#1a3a5c"},
          {l:"Conflictos detectados",v:conflictos.length,c:conflictos.length>0?"#8b2e2e":"#2d6a4f"},
          {l:"Artículos en catálogo",v:catalogo.length,c:"#4a2d6e"},
        ].map((k,i)=>(
          <div key={i} style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:9,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:4}}>{k.l}</div>
            <div style={{fontFamily:"Playfair Display,serif",fontSize:22,fontWeight:800,color:k.c}}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* ── CONFLICTOS ── */}
      {vistaTab==="conflictos"&&(
        <div>
          {conflictos.length===0?(
            <div style={{background:"#edf7f2",border:"1px solid #b7deca",borderRadius:12,padding:"32px",textAlign:"center" as const}}>
              <div style={{fontSize:36,marginBottom:8}}>✅</div>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:16,fontWeight:700,color:"#2d6a4f"}}>Sin conflictos en el período</div>
              <div style={{fontSize:12,color:"#9a9590",marginTop:4}}>
                {catalogo.filter(x=>x.existencia_total>0).length===0
                  ?"Agrega existencias en el Catálogo para detectar conflictos automáticamente"
                  :"No hay eventos que superen el inventario disponible"}
              </div>
              {catalogo.filter(x=>x.existencia_total>0).length===0&&(
                <div style={{marginTop:12,padding:"10px 14px",background:"#fff8f0",border:"1px solid #e8d4b8",borderRadius:8,fontSize:11,color:"#92580a"}}>
                  💡 Ve a <strong>Catálogo → Artículos</strong> y edita las existencias de tus piezas para activar la detección de conflictos
                </div>
              )}
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column" as const,gap:8}}>
              {conflictos.map((cf,i)=>(
                <div key={i} style={{background:"#fff",border:"1.5px solid #e8b8b8",borderRadius:10,padding:"12px 16px",borderLeft:"4px solid #8b2e2e"}}>
                  <div style={{display:"flex",alignItems:"flex-start",gap:12}}>
                    <div style={{width:36,height:36,borderRadius:8,background:"#fdf0f0",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>⚠️</div>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:700,fontSize:14,color:"#1a1814",marginBottom:3}}>{cf.articulo}</div>
                      <div style={{fontSize:12,color:"#9a9590",marginBottom:6}}>📅 {cf.fecha}</div>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap" as const}}>
                        {cf.clientes.map((cl,j)=>(
                          <span key={j} style={{fontSize:10,background:"#fdf0f0",color:"#8b2e2e",padding:"2px 8px",borderRadius:6}}>{cl}</span>
                        ))}
                      </div>
                    </div>
                    <div style={{textAlign:"right" as const,flexShrink:0}}>
                      <div style={{fontSize:12,color:"#8b2e2e",fontWeight:700}}>Demanda: {cf.demanda}</div>
                      <div style={{fontSize:12,color:"#2d6a4f"}}>Existencia: {cf.existencia}</div>
                      <div style={{fontSize:11,color:"#8b2e2e",fontWeight:700,marginTop:4}}>Faltan: {cf.demanda-cf.existencia}</div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ── OCUPACIÓN ── */}
      {vistaTab==="ocupacion"&&(
        <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:12,overflow:"hidden"}}>
          <div style={{padding:"12px 14px",borderBottom:"1px solid #e8e5de",fontFamily:"Playfair Display,serif",fontSize:14,fontWeight:700}}>
            Top artículos más demandados en el período
          </div>
          {topArts.length===0
            ?<div style={{padding:32,textAlign:"center" as const,color:"#9a9590",fontSize:12}}>Sin contratos en el período seleccionado</div>
            :<table style={{width:"100%",borderCollapse:"collapse" as const,fontSize:12}}>
              <thead>
                <tr style={{background:"#fafaf8"}}>
                  {["Artículo","Demanda total","Eventos","Existencia","Disponibilidad"].map((h,i)=>(
                    <th key={i} style={{padding:"9px 12px",textAlign:i===0?"left" as const:"center" as const,fontSize:10,fontWeight:700,color:"#9a9590",borderBottom:"1px solid #e8e5de",textTransform:"uppercase" as const,letterSpacing:".04em"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {topArts.map((a,i)=>{
                  const catArt=catalogo.find(c=>c.nombre.trim().toLowerCase()===a.nombre.trim().toLowerCase())
                  const exist=catArt?.existencia_total||0
                  const pct=exist>0?Math.min(100,Math.round(a.demandaTotal/exist*100)):0
                  const ok=exist===0||a.demandaTotal<=exist
                  return(
                    <tr key={i} style={{borderBottom:"1px solid #f5f4f0",background:i%2===0?"#fff":"#fafaf8"}}>
                      <td style={{padding:"9px 12px",fontWeight:600,maxWidth:280}}>
                        <div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</div>
                        <div style={{height:3,background:"#f5f4f0",borderRadius:2,marginTop:4,overflow:"hidden"}}>
                          <div style={{height:"100%",background:ok?"#2d6a4f":"#8b2e2e",borderRadius:2,width:pct+"%"}}/>
                        </div>
                      </td>
                      <td style={{padding:"9px 12px",textAlign:"center" as const,fontFamily:"monospace",fontWeight:700,color:"#1a3a5c"}}>{a.demandaTotal}</td>
                      <td style={{padding:"9px 12px",textAlign:"center" as const,color:"#9a9590"}}>{a.eventos.length}</td>
                      <td style={{padding:"9px 12px",textAlign:"center" as const,fontFamily:"monospace",color:exist===0?"#c4bfb8":"#2d6a4f",fontWeight:exist>0?700:400}}>{exist===0?"—":exist}</td>
                      <td style={{padding:"9px 12px",textAlign:"center" as const}}>
                        {exist===0
                          ?<span style={{fontSize:10,color:"#c4bfb8"}}>Sin dato</span>
                          :<span style={{fontSize:10,padding:"2px 8px",borderRadius:8,background:ok?"#edf7f2":"#fdf0f0",color:ok?"#2d6a4f":"#8b2e2e",fontWeight:700}}>{ok?"✓ OK":"⚠️ Conflicto"}</span>
                        }
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          }
        </div>
      )}

      {/* ── CATÁLOGO ── */}
      {vistaTab==="catalogo"&&(
        <div>
          <div style={{marginBottom:12}}>
            <input value={busqArt} onChange={e=>setBusqArt(e.target.value)} placeholder="Buscar artículo..."
              style={{padding:"8px 14px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:12,outline:"none",width:280}}/>
            <span style={{fontSize:11,color:"#9a9590",marginLeft:10}}>{catFiltrado.length} artículos</span>
          </div>
          {cargandoCat
            ?<div style={{padding:32,textAlign:"center" as const,color:"#9a9590"}}>Cargando catálogo...</div>
            :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
              {catFiltrado.map((a,i)=>{
                const demanda=demandaMap[a.nombre.trim().toLowerCase()]?.demandaTotal||0
                const ok=a.existencia_total===0||demanda<=a.existencia_total
                return(
                  <div key={i} style={{background:"#fff",border:`1px solid ${!ok?"#e8b8b8":"#e8e5de"}`,borderRadius:10,padding:"12px 14px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                      <div style={{fontWeight:700,fontSize:12,flex:1,marginRight:8,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" as const}}>{a.nombre}</div>
                      {!ok&&<span style={{fontSize:9,background:"#fdf0f0",color:"#8b2e2e",padding:"1px 5px",borderRadius:4,fontWeight:700,flexShrink:0}}>⚠️</span>}
                    </div>
                    <div style={{fontSize:10,color:"#9a9590",marginBottom:8}}>{a.categoria}{a.subcategoria?` · ${a.subcategoria}`:""}</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:4,fontSize:10}}>
                      <div style={{background:"#f5f4f0",borderRadius:5,padding:"4px 6px",textAlign:"center" as const}}>
                        <div style={{fontWeight:700,color:"#1a1814"}}>{a.existencia_total||"—"}</div>
                        <div style={{color:"#9a9590",fontSize:8}}>TOTAL</div>
                      </div>
                      <div style={{background:"#edf7f2",borderRadius:5,padding:"4px 6px",textAlign:"center" as const}}>
                        <div style={{fontWeight:700,color:"#2d6a4f"}}>{Math.max(0,(a.existencia_total||0)-demanda)||"—"}</div>
                        <div style={{color:"#9a9590",fontSize:8}}>LIBRE</div>
                      </div>
                      <div style={{background:demanda>0?"#edf3fa":"#f5f4f0",borderRadius:5,padding:"4px 6px",textAlign:"center" as const}}>
                        <div style={{fontWeight:700,color:demanda>0?"#1a3a5c":"#c4bfb8"}}>{demanda||"—"}</div>
                        <div style={{color:"#9a9590",fontSize:8}}>OCUPADO</div>
                      </div>
                    </div>
                    {a.precio_renta>0&&<div style={{fontSize:10,color:"#9a9590",marginTop:6}}>Renta: <strong style={{color:"#2d6a4f"}}>{fmt(a.precio_renta)}</strong></div>}
                  </div>
                )
              })}
            </div>
          }
        </div>
      )}
    </div>
  )
}


// ─── LOGO CONFIG ──────────────────────────────────────────────────
function LogoSection({logoUrl,setLogoUrl}:{logoUrl:string,setLogoUrl:(u:string)=>void}){
  const [preview,setPreview]=useState<string|null>(null)
  const [guardado,setGuardado]=useState(false)
  const [drag,setDrag]=useState(false)

  const procesarArchivo=(file:File)=>{
    if(!file.type.startsWith("image/")){alert("Solo se aceptan imágenes (PNG, JPG, SVG)");return}
    const reader=new FileReader()
    reader.onload=(e)=>{
      const result=e.target?.result as string
      setPreview(result)
      setGuardado(false)
    }
    reader.readAsDataURL(file)
  }

  const guardarLogo=async()=>{
    if(!preview)return
    setGuardado(false)
    try {
      // Save to Supabase so ALL devices see the logo
      await fetch("/api/config",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({clave:"logo",valor:preview})
      })
      localStorage.setItem("pf_logo",preview)
      setLogoUrl(preview)
      setGuardado(true)
      setTimeout(()=>setGuardado(false),3000)
    } catch(e) {
      // Fallback to localStorage only
      localStorage.setItem("pf_logo",preview)
      setLogoUrl(preview)
      setGuardado(true)
      setTimeout(()=>setGuardado(false),3000)
    }
  }

  const quitarLogo=async()=>{
    await fetch("/api/config",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({clave:"logo",valor:""})
    }).catch(()=>{})
    localStorage.removeItem("pf_logo")
    setLogoUrl("/logo.png")
    setPreview(null)
    setGuardado(false)
  }

  const logoActual=preview||logoUrl

  return(
    <div style={{maxWidth:720}}>
      <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,marginBottom:6}}>🖼️ Logo de la empresa</div>
      <div style={{fontSize:13,color:"#9a9590",marginBottom:20}}>El logo aparecerá en el sidebar, en las cotizaciones (PDF e impresión) y en el portal del cliente.</div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:20}}>
        {/* Upload area */}
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>Cargar logo</div>
          <div
            onDragOver={e=>{e.preventDefault();setDrag(true)}}
            onDragLeave={()=>setDrag(false)}
            onDrop={e=>{
              e.preventDefault();setDrag(false)
              const file=e.dataTransfer.files[0]
              if(file) procesarArchivo(file)
            }}
            style={{
              border:`2px dashed ${drag?"#2563eb":"#e8e5de"}`,
              borderRadius:12,padding:"32px 20px",textAlign:"center" as const,
              background:drag?"#eff6ff":"#fafaf8",cursor:"pointer",transition:"all .2s"
            }}>
            <div style={{fontSize:36,marginBottom:8,opacity:.4}}>🖼️</div>
            <div style={{fontSize:13,fontWeight:600,color:"#4a4640",marginBottom:4}}>Arrastra tu logo aquí</div>
            <div style={{fontSize:11,color:"#9a9590",marginBottom:12}}>PNG, JPG o SVG — hasta 5MB</div>
            <label style={{padding:"8px 18px",background:"#0f172a",color:"#fff",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700,fontFamily:"Epilogue,sans-serif"}}>
              <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                const file=e.target.files?.[0]
                if(file) procesarArchivo(file)
                e.target.value=""
              }}/>
              Seleccionar archivo
            </label>
          </div>
          {preview&&(
            <div style={{marginTop:12,display:"flex",gap:8}}>
              <button onClick={guardarLogo}
                style={{flex:1,padding:"10px",borderRadius:8,background:guardado?"#2d6a4f":"#0f172a",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"Epilogue,sans-serif",transition:"background .2s"}}>
                {guardado?"✓ Logo guardado":"💾 Guardar logo"}
              </button>
              <button onClick={()=>setPreview(null)}
                style={{padding:"10px 14px",borderRadius:8,background:"#f5f4f0",color:"#4a4640",border:"none",cursor:"pointer",fontSize:13,fontFamily:"Epilogue,sans-serif"}}>
                Cancelar
              </button>
            </div>
          )}
          {!preview&&logoUrl!=="/logo.png"&&(
            <button onClick={quitarLogo}
              style={{marginTop:12,width:"100%",padding:"8px",borderRadius:8,background:"#fdf0f0",color:"#8b2e2e",border:"1px solid #e8b8b8",cursor:"pointer",fontSize:12,fontFamily:"Epilogue,sans-serif"}}>
              Quitar logo actual
            </button>
          )}
        </div>

        {/* Previews */}
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#9a9590",textTransform:"uppercase" as const,letterSpacing:".06em",marginBottom:8}}>Vista previa</div>

          {/* Sidebar preview */}
          <div style={{background:"#0f172a",borderRadius:10,padding:"12px 16px",marginBottom:8,display:"flex",alignItems:"center",gap:10}}>
            <img src={logoActual} alt="Logo" style={{height:36,width:"auto",objectFit:"contain" as const,maxWidth:140,filter:"brightness(0) invert(1)"}}
              onError={(e:any)=>{e.target.style.display="none"}}/>
            <div style={{fontSize:10,color:"rgba(255,255,255,.3)",fontFamily:"Epilogue,sans-serif"}}>Sidebar</div>
          </div>

          {/* Cotizacion preview */}
          <div style={{background:"#fff",border:"1px solid #e8e5de",borderRadius:10,padding:"12px 16px",marginBottom:8,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div>
              <img src={logoActual} alt="Logo" style={{height:48,width:"auto",objectFit:"contain" as const,maxWidth:180}}
                onError={(e:any)=>{e.target.style.display="none"}}/>
              <div style={{fontSize:9,color:"#9a9590",marginTop:4,fontFamily:"Epilogue,sans-serif"}}>Cotizaciones PDF</div>
            </div>
            <div style={{textAlign:"right" as const}}>
              <div style={{fontSize:9,color:"#9a9590",fontFamily:"Epilogue,sans-serif",letterSpacing:".06em",textTransform:"uppercase" as const}}>Cotización</div>
              <div style={{fontFamily:"Playfair Display,serif",fontSize:20,fontWeight:800,color:"#1a1814"}}>K/H-000001</div>
            </div>
          </div>

          {/* Portal preview */}
          <div style={{background:"#1a1814",borderRadius:10,padding:"10px 16px",display:"flex",alignItems:"center",gap:10}}>
            <img src={logoActual} alt="Logo" style={{height:28,width:"auto",objectFit:"contain" as const,maxWidth:120,filter:"brightness(0) invert(1)"}}
              onError={(e:any)=>{e.target.style.display="none"}}/>
            <div style={{fontSize:10,color:"rgba(255,255,255,.3)",fontFamily:"Epilogue,sans-serif"}}>Portal del cliente</div>
          </div>
        </div>
      </div>

      {/* Info */}
      <div style={{background:"#f0fdf4",border:"1px solid #b7deca",borderRadius:10,padding:"12px 14px",fontSize:12,color:"#2d6a4f"}}>
        <strong>💡 Tip:</strong> Para mejor resultado usa un PNG con fondo transparente. El logo se adapta automáticamente al sidebar (fondo oscuro) y a las cotizaciones (fondo blanco).
        El logo se guarda en tu navegador — si usas otro dispositivo necesitas cargarlo de nuevo.
      </div>
    </div>
  )
}


// ─── CAMBIAR CONTRASEÑA ───────────────────────────────────────────
function CambiarPasswordSection({token,user}:{token:string,user:any}){
  const [actual,setActual]=useState("")
  const [nueva,setNueva]=useState("")
  const [confirmar,setConfirmar]=useState("")
  const [loading,setLoading]=useState(false)
  const [msg,setMsg]=useState<{tipo:"ok"|"err",texto:string}|null>(null)
  const [showActual,setShowActual]=useState(false)
  const [showNueva,setShowNueva]=useState(false)

  const handleSubmit=async()=>{
    setMsg(null)
    if(!actual||!nueva||!confirmar){setMsg({tipo:"err",texto:"Completa todos los campos"});return}
    if(nueva!==confirmar){setMsg({tipo:"err",texto:"Las contraseñas nuevas no coinciden"});return}
    if(nueva.length<6){setMsg({tipo:"err",texto:"La contraseña debe tener al menos 6 caracteres"});return}
    setLoading(true)
    try{
      const res=await fetch("/api/auth/password",{
        method:"POST",
        headers:{"Content-Type":"application/json",Authorization:`Bearer ${token}`},
        body:JSON.stringify({password_actual:actual,password_nuevo:nueva})
      })
      const data=await res.json()
      if(!res.ok){setMsg({tipo:"err",texto:data.error||"Error al cambiar contraseña"})}
      else{
        setMsg({tipo:"ok",texto:"✓ Contraseña actualizada correctamente"})
        setActual("");setNueva("");setConfirmar("")
      }
    }catch{
      setMsg({tipo:"err",texto:"Error de conexión"})
    }
    setLoading(false)
  }

  return(
    <div style={{maxWidth:440}}>
      <div style={{fontFamily:"Playfair Display,serif",fontSize:18,fontWeight:800,marginBottom:4}}>🔑 Cambiar contraseña</div>
      <div style={{fontSize:13,color:"#9a9590",marginBottom:24}}>Hola <strong>{user?.nombre}</strong> — elige una contraseña segura que solo tú conozcas.</div>

      <div style={{display:"flex",flexDirection:"column" as const,gap:14}}>
        {/* Contraseña actual */}
        <div>
          <label style={{fontSize:11,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".06em",display:"block",marginBottom:6}}>Contraseña actual</label>
          <div style={{position:"relative" as const}}>
            <input type={showActual?"text":"password"} value={actual} onChange={e=>setActual(e.target.value)}
              placeholder="Tu contraseña actual"
              style={{width:"100%",padding:"10px 40px 10px 12px",border:"1.5px solid #e8e5de",borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
            <button onClick={()=>setShowActual(v=>!v)}
              style={{position:"absolute" as const,right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#9a9590"}}>
              {showActual?"🙈":"👁️"}
            </button>
          </div>
        </div>

        {/* Nueva contraseña */}
        <div>
          <label style={{fontSize:11,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".06em",display:"block",marginBottom:6}}>Nueva contraseña</label>
          <div style={{position:"relative" as const}}>
            <input type={showNueva?"text":"password"} value={nueva} onChange={e=>setNueva(e.target.value)}
              placeholder="Mínimo 6 caracteres"
              style={{width:"100%",padding:"10px 40px 10px 12px",border:`1.5px solid ${nueva.length>0&&nueva.length<6?"#8b2e2e":"#e8e5de"}`,borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
            <button onClick={()=>setShowNueva(v=>!v)}
              style={{position:"absolute" as const,right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#9a9590"}}>
              {showNueva?"🙈":"👁️"}
            </button>
          </div>
          {nueva.length>0&&nueva.length<6&&(
            <div style={{fontSize:11,color:"#8b2e2e",marginTop:4}}>Mínimo 6 caracteres</div>
          )}
        </div>

        {/* Confirmar */}
        <div>
          <label style={{fontSize:11,fontWeight:700,color:"#4a4640",textTransform:"uppercase" as const,letterSpacing:".06em",display:"block",marginBottom:6}}>Confirmar nueva contraseña</label>
          <input type="password" value={confirmar} onChange={e=>setConfirmar(e.target.value)}
            placeholder="Repite la nueva contraseña"
            style={{width:"100%",padding:"10px 12px",border:`1.5px solid ${confirmar.length>0&&confirmar!==nueva?"#8b2e2e":"#e8e5de"}`,borderRadius:8,fontFamily:"Epilogue,sans-serif",fontSize:13,outline:"none",boxSizing:"border-box" as const}}/>
          {confirmar.length>0&&confirmar!==nueva&&(
            <div style={{fontSize:11,color:"#8b2e2e",marginTop:4}}>Las contraseñas no coinciden</div>
          )}
        </div>

        {/* Mensaje */}
        {msg&&(
          <div style={{padding:"10px 14px",borderRadius:8,background:msg.tipo==="ok"?"#f0fdf4":"#fdf0f0",border:`1px solid ${msg.tipo==="ok"?"#b7deca":"#e8b8b8"}`,fontSize:13,color:msg.tipo==="ok"?"#2d6a4f":"#8b2e2e",fontWeight:600}}>
            {msg.texto}
          </div>
        )}

        {/* Botón */}
        <button onClick={handleSubmit} disabled={loading}
          style={{padding:"12px",borderRadius:8,background:loading?"#9a9590":"#0f172a",color:"#fff",border:"none",cursor:loading?"not-allowed":"pointer",fontFamily:"Epilogue,sans-serif",fontSize:13,fontWeight:700,transition:"background .15s"}}>
          {loading?"Guardando...":"Cambiar contraseña"}
        </button>
      </div>
    </div>
  )
}


// ═══════════════════════════════════════════════════════════════
// SPLITS — Configuración de tipos (escalable)
// ═══════════════════════════════════════════════════════════════
const SPLIT_TIPOS=[
  {id:"rutas",       nombre:"Entrega",            icono:"🚚",color:"#1a3a5c",bg:"#eff6ff",  cats:["__ALL__"], mostrarPrecio:false,mostrarDir:true},
  {id:"carpinteria", nombre:"Carpintería",      icono:"🪚",color:"#92580a",bg:"#fffbeb",  cats:["MOBILIARIO"],                                                mostrarPrecio:false,mostrarDir:false},
  {id:"vajilla",     nombre:"Vajilla",           icono:"🍽️",color:"#4a2d6e",bg:"#faf5ff",  cats:["VAJILLA"],                                                   mostrarPrecio:false,mostrarDir:false},
  {id:"flores",      nombre:"Flores",            icono:"🌸",color:"#2d6a4f",bg:"#f0fdf4",  cats:["FLORES"],                                                    mostrarPrecio:true, mostrarDir:false},
  {id:"bases",       nombre:"Bases",             icono:"🏗️",color:"#64748b",bg:"#f8fafc",  cats:["CARPAS","ESTRUCTURA"],                                       mostrarPrecio:false,mostrarDir:true},
  {id:"desmonte",    nombre:"Desmonte",          icono:"📦",color:"#334155",bg:"#f1f5f9",  cats:["__ALL__"],                                                                     mostrarPrecio:false,mostrarDir:true},
  {id:"proveedor",   nombre:"Hoja de Proveedor", icono:"📄",color:"#8b2e2e",bg:"#fdf0f0",  cats:[],                                                            mostrarPrecio:false,mostrarDir:false},
]
const SPLIT_ESTADOS=[
  {id:"pendiente",  label:"Pendiente",  color:"#92580a",bg:"#fffbeb"},
  {id:"en_proceso", label:"En proceso", color:"#1a3a5c",bg:"#eff6ff"},
  {id:"listo",      label:"Listo",      color:"#2d6a4f",bg:"#f0fdf4"},
  {id:"entregado",  label:"Entregado",  color:"#4a2d6e",bg:"#faf5ff"},
]

function generarSplitsDesdeContrato(contrato:any):any[]{
  const arts=contrato.articulos||[]
  // Fecha preparación = día anterior al evento
  const addDaysStr=(dateStr:string,days:number):string=>{
    if(!dateStr)return ""
    const d=new Date(dateStr+"T12:00:00")
    d.setDate(d.getDate()+days)
    return d.toISOString().slice(0,10)
  }
  const fechaEvento=contrato.fecha_evento||""
  const fechaPrep=addDaysStr(fechaEvento,-1)
  // Desmonte = fecha_desmonte si existe, si no fecha_evento+1
  const fechaDesmonte=contrato.fecha_desmonte||(fechaEvento?addDaysStr(fechaEvento,1):"")
  // If still no desmonte date, use fechaEvento as fallback
  const fechaDesmonteFinal=fechaDesmonte||fechaEvento

  const base={
    contrato_id:contrato.id,
    contrato_folio:contrato.folio||contrato.archivo||"",
    cliente:contrato.cliente||contrato.archivo||"",
    lugar:contrato.lugar||"",
    tel:contrato.tel||contrato.telefono||"",
    fecha_evento:fechaEvento,
    fecha_evento_original:fechaEvento,
    estado:"pendiente",notas:"",observaciones:"",responsable:"",
  }
  return SPLIT_TIPOS.filter(t=>t.id!=="proveedor").map(tipo=>({
    ...base,
    // prep areas: fecha_evento = día de prep (antes), rutas: evento, desmonte: fecha desmonte
    fecha_evento:tipo.id==="desmonte"?fechaDesmonteFinal:tipo.id==="rutas"?fechaEvento:fechaPrep,
    fecha_preparacion:tipo.id==="desmonte"?fechaEvento:tipo.id==="rutas"?fechaEvento:fechaEvento,
    // Fechas del contrato (inmodificables) para mostrar en hoja
    fecha_entrega_contrato:contrato.fecha_entrega||addDaysStr(fechaEvento,-1),
    fecha_desmonte_contrato:contrato.fecha_desmonte||addDaysStr(fechaEvento,1),
    tipo:tipo.id,nombre:tipo.nombre,
    mostrar_precios:tipo.mostrarPrecio,mostrar_direccion:tipo.mostrarDir,
    proveedor_nombre:"",
    articulos:arts.filter((a:any)=>{
      // Excluir TRASLADO, MONTAJE, ENVÍO de TODAS las hojas
      const nom=(a.nombre||"").toUpperCase()
      if(nom.includes("TRASLADO")||nom.includes("MONTAJE")||nom.includes("ENVIO")||nom.includes("ENVÍO"))return false
      // Rutas = todos los artículos
      if(tipo.cats.includes("__ALL__"))return true
      const cat=(a.categoria||a.seccion||"").toUpperCase().trim()
      return tipo.cats.some((cc:string)=>cat.includes(cc.toUpperCase()))
    })
  }))
}

function abrirHojaSplit(split:any,logoSrc:string){
  const tipo=SPLIT_TIPOS.find(t=>t.id===split.tipo)||SPLIT_TIPOS[0]
  const arts=split.articulos||[]
  // Usar la definición de SPLIT_TIPOS como fuente de verdad para dirección y precio
  const tipoConfig=SPLIT_TIPOS.find((t:any)=>t.id===split.tipo)
  const mostrarDir=tipoConfig ? tipoConfig.mostrarDir : split.mostrar_direccion!==false
  const mostrarPrecio=tipoConfig ? tipoConfig.mostrarPrecio : split.mostrar_precios===true
  const DIAS_ES=["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"]
  // Para desmonte: el evento es el día ANTES de la fecha de desmonte (split.fecha_evento)
  // Prioridad: fecha_preparacion (si existe) → fecha_evento - 1 día
  // fecha_preparacion guarda la fecha REAL del evento del contrato para rutas y desmonte
  // fechaEventoReal = fecha REAL del evento del contrato (para mostrar en "📅 Evento:")
  // Para prep areas: split.fecha_evento es la fecha de preparación (día antes), entonces evento = fecha_evento + 1
  // Para rutas: fecha_preparacion tiene el evento real, o entrega + 1
  // Para desmonte: fecha_preparacion tiene el evento real, o desmonte - 1
  // Para prep areas: si tiene fecha_entrega_contrato, el evento = entrega + 1 día
  // Si no, calcular desde fecha_evento del split
  const fechaEventoReal=(()=>{
    if(split.tipo==="rutas"||split.tipo==="desmonte"){
      if(split.fecha_preparacion) return split.fecha_preparacion
      if(!split.fecha_evento) return ""
      const d=new Date(split.fecha_evento+"T12:00:00")
      d.setDate(d.getDate()+(split.tipo==="rutas"?1:-1))
      return d.toISOString().slice(0,10)
    }
    // Prep areas: si tiene fecha_entrega_contrato, evento = entrega + 1 día (infalible)
    if(split.fecha_entrega_contrato){
      const d=new Date(split.fecha_entrega_contrato+"T12:00:00")
      d.setDate(d.getDate()+1)
      return d.toISOString().slice(0,10)
    }
    // Fallback: fecha_evento + 1
    if(!split.fecha_evento) return ""
    const d=new Date(split.fecha_evento+"T12:00:00")
    d.setDate(d.getDate()+1)
    return d.toISOString().slice(0,10)
  })()
  const fecha=fechaEventoReal
    ?new Date(fechaEventoReal+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})
    :"Por confirmar"
  // Fecha REAL de trabajo según el tipo:
  // - rutas: fecha del split (ya es la entrega)
  // - desmonte: fecha del split (ya es desmonte)
  // - resto: fechaPrep (día antes del evento)
  const esPrepArea=split.tipo!=="rutas"&&split.tipo!=="desmonte"&&split.tipo!=="proveedor"
  const fechaTrabajo=esPrepArea
    ?(split.fecha_entrega_contrato||split.fecha_preparacion||split.fecha_evento)
    :split.fecha_evento
  const fechaTrabajoFmt=fechaTrabajo
    ?new Date(fechaTrabajo+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})
    :"Por confirmar"
  const diaSemana=fechaTrabajo
    ?DIAS_ES[new Date(fechaTrabajo+"T12:00:00").getDay()]
    :"—"
  // Etiqueta del tipo de actividad
  const etiquetaTipo=split.tipo==="rutas"?"Entrega":split.tipo==="desmonte"?"Desmontaje":split.tipo==="carpinteria"?"Cargan":"Preparación"
  // Fecha de preparación = día antes del evento
  const fechaPrep=(()=>{
    const base=split.fecha_preparacion||(split.fecha_evento?new Date(new Date(split.fecha_evento+"T12:00:00").getTime()-86400000).toISOString().slice(0,10):"")
    return base?new Date(base+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long"}):"—"
  })()
  const filas=arts.map((a:any,i:number)=>`
    <tr style="border-bottom:1px solid #f0f0f0;background:${i%2===0?"#fff":"#fafafa"}">
      <td style="padding:8px 10px;font-size:14px;font-weight:800;text-align:center;color:${tipo.color}">${a.cantidad||0}</td>
      <td style="padding:8px 10px;font-size:13px;font-weight:600">${a.nombre||"—"}</td>
      ${mostrarPrecio?`<td style="padding:8px 10px;font-size:13px;font-weight:700;text-align:right;font-family:monospace">$${(a.pu||a.precio_unitario||0).toLocaleString("es-MX")}</td>`:""}
      ${a.notas?`<td style="padding:8px 10px;font-size:11px;color:#64748b;font-style:italic">${a.notas}</td>`:"<td></td>"}
    </tr>`).join("")
  const totalPiezas=arts.reduce((s:number,a:any)=>s+(a.cantidad||0),0)
  const totalImporte=arts.reduce((s:number,a:any)=>s+(a.importe||a.subtotal||0),0)
  const html=`<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<title>${tipo.nombre} — ${split.contrato_folio||split.cliente}</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet"/>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Inter,Arial,sans-serif;color:#1a1814;background:#fff;font-size:13px}
table{border-collapse:collapse;width:100%}
th,td{vertical-align:middle}
@media print{
  .no-print{display:none!important}
  @page{margin:12mm 14mm;size:A4}
  body{print-color-adjust:exact;-webkit-print-color-adjust:exact}
}
</style></head>
<body style="padding:24px 28px;max-width:820px;margin:0 auto">

<!-- Botones pantalla -->
<div class="no-print" style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:16px">
  <button onclick="window.print()" style="padding:8px 20px;background:#1a1814;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;font-family:Inter,sans-serif">🖨️ Imprimir</button>
  <button onclick="window.close()" style="padding:8px 14px;background:#f5f4f0;border:1px solid #ccc;border-radius:8px;cursor:pointer;font-family:Inter,sans-serif;font-size:13px">✕</button>
</div>

<!-- ══ FILA 1: Logo | Preparación/Entrega/Desmontaje + día ══ -->
<div style="display:flex;align-items:center;justify-content:space-between;padding-bottom:12px;margin-bottom:12px;border-bottom:2.5px solid ${tipo.color}">
  <!-- Logo -->
  <img src="${logoSrc}" alt="Poliflor" style="height:40px;width:auto;object-fit:contain" onerror="this.style.display='none'"/>
  <!-- Etiqueta + día + fecha (compacto, derecha) -->
  <div style="text-align:right">
    <div style="font-size:11px;font-weight:700;color:${tipo.color};text-transform:uppercase;letter-spacing:.1em">${etiquetaTipo}</div>
    <div style="font-size:22px;font-weight:800;color:${tipo.color};line-height:1.1">${diaSemana}</div>
    <div style="font-size:11px;color:#4a4640;font-weight:500">${fechaTrabajoFmt}</div>
  </div>
</div>

<!-- ══ FILA 2: Cliente | Hoja de Trabajo / Área ══ -->
<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #e8e5de">
  <!-- Cliente -->
  <div>
    <div style="font-size:9px;font-weight:700;color:#9a9590;text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px">Cliente</div>
    <div style="font-size:22px;font-weight:800;color:#1a1814">${split.cliente||"—"}</div>
    ${split.contrato_folio?`<div style="font-size:13px;color:#9a9590;font-family:monospace;font-weight:700">${split.contrato_folio}</div>`:""}
    <div style="font-size:13px;color:#4a4640;font-weight:600;margin-top:6px">📅 Evento: ${fecha}</div>
    ${mostrarDir&&split.lugar?`<div style="font-size:13px;color:#4a4640;font-weight:600;margin-top:4px">📍 ${split.lugar}</div>`:""}
    ${(split.tipo==="rutas"||split.tipo==="desmonte")&&split.tel?`<div style="font-size:14px;color:#1a1814;font-weight:700;margin-top:4px">📞 ${split.tel}</div>`:""}
    ${["vajilla","carpinteria","bases","flores"].includes(split.tipo)&&split.fecha_entrega_contrato?`<div style="font-size:11px;color:#9a9590;margin-top:3px">🚚 Entrega: ${new Date(split.fecha_entrega_contrato+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>`:""}\n    ${["vajilla","carpinteria","bases","flores"].includes(split.tipo)&&split.fecha_desmonte_contrato?`<div style="font-size:11px;color:#9a9590;margin-top:2px">📦 Desmonte: ${new Date(split.fecha_desmonte_contrato+"T12:00:00").toLocaleDateString("es-MX",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>`:""}
  </div>
  <!-- Hoja de Trabajo / Área --> -->
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
