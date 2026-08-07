import jwt from 'jsonwebtoken'

const JWT_SECRET = process.env.JWT_SECRET || 'poliflor-secret-2024'

export function verifyToken(token: string): any {
  try {
    return jwt.verify(token, JWT_SECRET)
  } catch {
    return null
  }
}

export function createToken(payload: any): string {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' })
}
