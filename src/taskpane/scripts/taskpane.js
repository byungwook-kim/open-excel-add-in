const { formatValuesToText, getAnchorReference, parseTextToValues, splitAddress } = require("./rangeText")

const AUTH_STORAGE_KEY = "open-excel-auth"

let elements

Office.onReady((info) => {
  elements = {
    addressInput: document.getElementById("address-input"),
    authStatus: document.getElementById("auth-status"),
    loginButton: document.getElementById("login-button"),
    readButton: document.getElementById("read-button"),
    resetAuthButton: document.getElementById("reset-auth-button"),
    statusOutput: document.getElementById("status-output"),
    valuesInput: document.getElementById("values-input"),
    writeButton: document.getElementById("write-button"),
  }

  if (info.host !== Office.HostType.Excel) {
    setStatus("이 add-in은 Excel에서만 동작합니다.")
    disableActions(true)
    return
  }

  disableActions(false)
  hydrateAuthState()

  elements.loginButton.addEventListener("click", startExperimentalLogin)
  elements.readButton.addEventListener("click", readRange)
  elements.resetAuthButton.addEventListener("click", resetStoredAuth)
  elements.writeButton.addEventListener("click", writeRange)

  setStatus("Excel 준비 완료. Range를 입력하고 Read/Write를 사용할 수 있습니다.")
})

function disableActions(disabled) {
  if (!elements) return
  elements.loginButton.disabled = disabled
  elements.readButton.disabled = disabled
  elements.resetAuthButton.disabled = disabled
  elements.writeButton.disabled = disabled
  elements.addressInput.disabled = disabled
  elements.valuesInput.disabled = disabled
}

function hydrateAuthState() {
  const raw = window.localStorage.getItem(AUTH_STORAGE_KEY)
  if (!raw) {
    elements.authStatus.textContent = "로그인되지 않음"
    return
  }

  try {
    const auth = JSON.parse(raw)
    const expiresText = auth.expiresAt ? new Date(auth.expiresAt).toLocaleString() : "unknown"
    elements.authStatus.textContent = auth.accountId
      ? `실험적 로그인 저장됨: ${auth.accountId} (expires ${expiresText})`
      : `실험적 로그인 저장됨 (expires ${expiresText})`
  } catch {
    elements.authStatus.textContent = "저장된 로그인 상태를 읽지 못했습니다."
  }
}

function resetStoredAuth() {
  window.localStorage.removeItem(AUTH_STORAGE_KEY)
  hydrateAuthState()
  setStatus("저장된 실험용 OAuth 토큰을 삭제했습니다.")
}

function startExperimentalLogin() {
  setStatus("실험적 ChatGPT OAuth dialog를 여는 중...")
  Office.context.ui.displayDialogAsync(
    `${window.location.origin}/auth-start.html`,
    { height: 60, width: 30, displayInIframe: false },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus(`OAuth dialog를 열지 못했습니다: ${result.error.message}`)
        return
      }

      const dialog = result.value
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
        try {
          const payload = JSON.parse(event.message)
          if (payload.type === "oauth-success") {
            window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(payload.auth))
            hydrateAuthState()
            setStatus("실험적 ChatGPT OAuth 응답 저장 완료")
          } else {
            setStatus(payload.message || "OAuth가 실패했습니다.")
          }
        } catch {
          setStatus("OAuth 응답을 해석하지 못했습니다.")
        } finally {
          dialog.close()
        }
      })

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
        setStatus(`OAuth dialog 이벤트: ${event.error}`)
      })
    },
  )
}

async function readRange() {
  const address = elements.addressInput.value.trim()
  if (!address) {
    setStatus("Read할 range를 입력해 주세요.")
    return
  }

  setStatus(`Reading ${address}...`)

  try {
    const values = await Excel.run(async (context) => {
      const { reference, sheetName } = splitAddress(address)
      const worksheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const range = worksheet.getRange(reference)
      range.load("values")
      await context.sync()
      return range.values
    })

    elements.valuesInput.value = formatValuesToText(values)
    setStatus(`Read 완료: ${address}`)
  } catch (error) {
    setStatus(`Read 실패: ${toErrorMessage(error)}`)
  }
}

async function writeRange() {
  const address = elements.addressInput.value.trim()
  if (!address) {
    setStatus("Write할 anchor를 입력해 주세요.")
    return
  }

  const values = parseTextToValues(elements.valuesInput.value)
  setStatus(`Writing to ${address}...`)

  try {
    await Excel.run(async (context) => {
      const anchorAddress = getAnchorReference(address)
      const { reference, sheetName } = splitAddress(anchorAddress)
      const worksheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet()

      const anchor = worksheet.getRange(reference)
      const target = anchor.getResizedRange(values.length - 1, values[0].length - 1)
      target.values = values
      await context.sync()
    })

    setStatus(`Write 완료: ${address}`)
  } catch (error) {
    setStatus(`Write 실패: ${toErrorMessage(error)}`)
  }
}

function setStatus(message) {
  if (elements?.statusOutput) {
    elements.statusOutput.textContent = message
  }
}

function toErrorMessage(error) {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}
