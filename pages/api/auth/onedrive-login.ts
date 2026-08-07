import { NextApiRequest, NextApiResponse } from "next"
export default function handler(req: NextApiRequest, res: NextApiResponse) {
  const url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
    "?client_id=d15ce363-0ef6-4c41-904b-f365a871cf7c" +
    "&response_type=code" +
    "&redirect_uri=" + encodeURIComponent("https://index-eta-puce.vercel.app/api/auth/callback") +
    "&scope=" + encodeURIComponent("Files.Read.All User.Read offline_access") +
    "&response_mode=query" +
    "&prompt=consent"
  return res.redirect(url)
}
