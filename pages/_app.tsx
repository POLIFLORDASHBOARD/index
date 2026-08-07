import type { AppProps } from "next/app"
import Head from "next/head"
export default function App({ Component, pageProps }: AppProps) {
  return (
    <>
      <Head>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800&family=Epilogue:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet" />
        <style>{`* { box-sizing: border-box; margin: 0; padding: 0; } body { font-family: Epilogue, sans-serif; background: #f5f4f0; }`}</style>
      </Head>
      <Component {...pageProps} />
    </>
  )
}
