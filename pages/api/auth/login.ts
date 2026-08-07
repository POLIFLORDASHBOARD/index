import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") return res.status(405).end()
  try {
    const { email, password } = req.body
    const url = "https://ohxehnsxfbvdflmqlzxq.supabase.co"
    const key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
    const supabase = createClient(url, key)
    const { data: user, error } = await supabase.from("usuarios").select("*").eq("email", email.toLowerCase()).single()
    if (error) return res.status(401).json({ error: "DB: " + error.message })
    if (!user) return res.status(401).json({ error: "No encontrado" })
    if (user.password_hash !== password) return res.status(401).json({ error: "Pass incorrecto" })
    const jwt = require("jsonwebtoken")
    const token = jwt.sign({ id: user.id, nombre: user.nombre, email: user.email, rol: user.rol }, "poliflor_super_secret_2026_dashboard_key", { expiresIn: "7d" })
    return res.status(200).json({ token, user: { id: user.id, nombre: user.nombre, email: user.email, rol: user.rol } })
  } catch(e: any) {
    return res.status(500).json({ error: e.message })
  }
}
