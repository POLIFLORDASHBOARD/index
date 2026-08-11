import { NextApiRequest, NextApiResponse } from "next"
import { createClient } from "@supabase/supabase-js"

const supabase = createClient(
  "https://ohxehnsxfbvdflmqlzxq.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
)

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  const { data, error } = await supabase
    .from("cuentas_bancarias")
    .select("*")
    .eq("activo", true)
    .order("banco")
  if (error) return res.status(500).json({ error: error.message })
  return res.status(200).json(data || [])
}
