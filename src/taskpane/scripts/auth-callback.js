const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann"
const ISSUER = "https://auth.openai.com"

Office.onReady(async () => {
  try {
    const callbackUrl = new URL(window.location.href)
    const error = callbackUrl.searchParams.get("error")
    const code = callbackUrl.searchParams.get("code")
    const state = callbackUrl.searchParams.get("state")
    const expectedState = window.sessionStorage.getItem("open-excel-oauth-state")
    const verifier = window.sessionStorage.getItem("open-excel-oauth-verifier")
    const redirectUri = `${window.location.origin}/auth-callback.html`

    if (error) {
      throw new Error(`OAuth failed: ${error}`)
    }

    if (!code || !state || state !== expectedState || !verifier) {
      throw new Error("Invalid OAuth callback state.")
    }

    const response = await fetch(`${ISSUER}/oauth/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri,
        client_id: CLIENT_ID,
        code_verifier: verifier,
      }).toString(),
    })

    if (!response.ok) {
      throw new Error(`Failed token exchange: ${response.status}`)
    }

    const tokens = await response.json()
    const auth = {
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt: Date.now() + (tokens.expires_in ?? 3600) * 1000,
      accountId: extractAccountId(tokens.id_token || tokens.access_token),
    }

    Office.context.ui.messageParent(JSON.stringify({ type: "oauth-success", auth }))
  } catch (error) {
    Office.context.ui.messageParent(
      JSON.stringify({
        type: "oauth-error",
        message: error instanceof Error ? error.message : String(error),
      }),
    )
  }
})

function extractAccountId(token) {
  if (!token || typeof token !== "string") {
    return undefined
  }

  const parts = token.split(".")
  if (parts.length !== 3) {
    return undefined
  }

  try {
    const claims = JSON.parse(fromBase64Url(parts[1]))
    const nested = claims["https://api.openai.com/auth"]
    if (nested && typeof nested === "object" && typeof nested.chatgpt_account_id === "string") {
      return nested.chatgpt_account_id
    }

    return claims.chatgpt_account_id || claims.organizations?.[0]?.id
  } catch {
    return undefined
  }
}

function fromBase64Url(value) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/")
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=")
  return atob(padded)
}
