const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
const ISSUER = "https://auth.openai.com"

Office.onReady(async () => {
  try {
    const verifier = generateRandomString(43)
    const challenge = await createCodeChallenge(verifier)
    const state = generateState()
    const redirectUri = `${window.location.origin}/auth-callback.html`

    window.sessionStorage.setItem("open-excel-oauth-verifier", verifier)
    window.sessionStorage.setItem("open-excel-oauth-state", state)

    const params = new URLSearchParams({
      response_type: "code",
      client_id: CLIENT_ID,
      redirect_uri: redirectUri,
      scope: "openid profile email offline_access",
      code_challenge: challenge,
      code_challenge_method: "S256",
      id_token_add_organizations: "true",
      codex_cli_simplified_flow: "true",
      state,
      originator: "open-excel-add-in",
    })

    window.location.replace(`${ISSUER}/oauth/authorize?${params.toString()}`)
  } catch (error) {
    Office.context.ui.messageParent(
      JSON.stringify({
        type: "oauth-error",
        message: error instanceof Error ? error.message : String(error),
      }),
    )
  }
})

function generateRandomString(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
  const bytes = crypto.getRandomValues(new Uint8Array(length))
  return Array.from(bytes, (value) => chars[value % chars.length]).join("")
}

function generateState() {
  const bytes = crypto.getRandomValues(new Uint8Array(32))
  return toBase64Url(bytes.buffer)
}

async function createCodeChallenge(verifier) {
  const buffer = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))
  return toBase64Url(buffer)
}

function toBase64Url(buffer) {
  const bytes = new Uint8Array(buffer)
  let binary = ""
  for (const byte of bytes) {
    binary += String.fromCharCode(byte)
  }

  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "")
}
