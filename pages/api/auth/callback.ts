import { NextApiRequest, NextApiResponse } from "next"

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  const { code, error } = req.query
  if (error) return res.redirect("/?error=" + error)
  if (!code) return res.status(400).json({ error: "No code" })
  
  try {
    const params = new URLSearchParams({
      client_id: "d15ce363-0ef6-4c41-904b-f365a871cf7c",
      client_secret: "xMS8Q~Ayny.rs147vEghYxJkRcfyE85UC2euwaV3",
      code: code as string,
      redirect_uri: "https://index-eta-puce.vercel.app/api/auth/callback",
      grant_type: "authorization_code",
    })
    
    const tokenRes = await fetch("https://login.microsoftonline.com/b6497cb4-7fca-430c-96e1-dca4373ef61c/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString()
    })
    
    const tokens = await tokenRes.json()
    if (tokens.error) return res.redirect("/?error=" + tokens.error_description)
    
    const { createClient } = require("@supabase/supabase-js")
    const supabase = createClient(
      "https://ohxehnsxfbvdflmqlzxq.supabase.co",
      "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oeGVobnN4ZmJ2ZGZsbXFsenhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDk0MDYyMCwiZXhwIjoyMDk2NTE2NjIwfQ.v6Gh1ZmQSSPKc3ESTTsuoiUihZ1LrejFQbxpqDGpjoM"
    )
    
    const expires = new Date(Date.now() + tokens.expires_in * 1000).toISOString()
    
    await supabase.from("onedrive_config").upsert({
      carpeta_nombre: "Principal",
      carpeta_id: "root",
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      token_expires: expires,
      activa: true,
    }, { onConflict: "carpeta_id" })
    
    return res.redirect("/?onedrive=connected")
  } catch(e: any) {
    return res.redirect("/?error=" + e.message)
  }
}
