import jwt from 'jsonwebtoken'
import { NextApiRequest, NextApiResponse } from 'next'

const JWT_SECRET = process.env.JWT_SECRET || 'poliflor-secret-2024'

export interface TokenPayload {
  id: string
  email: string
  nombre: string
  rol: string
}

export function verifyToken(token: string): TokenPayload | null {
  try {
    return jwt.verify(token, JWT_SECRET) as TokenPayload
  } catch {
    return null
  }
}

export function createToken(payload: any): string {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' })
}

export function requireAuth(
  handler: (req: NextApiRequest, res: NextApiResponse, user: TokenPayload) => Promise<any>
) {
  return async (req: NextApiRequest, res: NextApiResponse) => {
    const authHeader = req.headers.authorization
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'No autorizado' })
    }
    const token = authHeader.slice(7)
    const user = verifyToken(token)
    if (!user) {
      return res.status(401).json({ error: 'Token inválido' })
    }
    return handler(req, res, user)
  }
}
