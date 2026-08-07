import { NextApiRequest, NextApiResponse } from 'next'
import bcrypt from 'bcryptjs'
import { supabaseAdmin } from '../../../lib/supabase'
import { verifyToken } from '../../../lib/auth'

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') return res.status(405).end()

  const auth = req.headers.authorization?.replace('Bearer ', '')
  if (!auth) return res.status(401).json({ error: 'No autorizado' })

  let userId: string
  try {
    const payload = verifyToken(auth) as any
    userId = payload.id
  } catch {
    return res.status(401).json({ error: 'Token inválido' })
  }

  const { password_actual, password_nuevo } = req.body
  if (!password_actual || !password_nuevo)
    return res.status(400).json({ error: 'Faltan campos' })
  if (password_nuevo.length < 6)
    return res.status(400).json({ error: 'La contraseña debe tener al menos 6 caracteres' })

  // Get current user
  const { data: user, error } = await supabaseAdmin
    .from('usuarios')
    .select('password_hash')
    .eq('id', userId)
    .single()

  if (error || !user) return res.status(404).json({ error: 'Usuario no encontrado' })

  // Verify current password
  let valid = false
  if (user.password_hash === '$2b$10$placeholder_will_be_replaced') {
    valid = password_actual === 'POLIFLOR2026'
  } else {
    valid = await bcrypt.compare(password_actual, user.password_hash)
  }

  if (!valid) return res.status(401).json({ error: 'Contraseña actual incorrecta' })

  // Hash and save new password
  const hash = await bcrypt.hash(password_nuevo, 10)
  await supabaseAdmin
    .from('usuarios')
    .update({ password_hash: hash })
    .eq('id', userId)

  return res.status(200).json({ ok: true })
}
