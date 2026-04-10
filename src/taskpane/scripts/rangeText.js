function parseTextToValues(text) {
  const trimmed = text.replace(/\r/g, "")

  if (trimmed.length === 0) {
    return [[""]]
  }

  const rows = trimmed.split("\n").map((line) => line.split("\t"))
  const maxColumns = rows.reduce((current, row) => Math.max(current, row.length), 0)

  return rows.map((row) => {
    if (row.length === maxColumns) {
      return row
    }

    return row.concat(Array.from({ length: maxColumns - row.length }, () => ""))
  })
}

function formatValuesToText(values) {
  if (!Array.isArray(values) || values.length === 0) {
    return ""
  }

  return values
    .map((row) => (Array.isArray(row) ? row.map((cell) => (cell == null ? "" : String(cell))).join("\t") : ""))
    .join("\n")
}

function splitAddress(address) {
  const trimmed = address.trim()
  if (!trimmed) {
    throw new Error("Range address is required.")
  }

  const bangIndex = trimmed.lastIndexOf("!")
  if (bangIndex === -1) {
    return { sheetName: undefined, reference: trimmed }
  }

  const sheetName = trimmed.slice(0, bangIndex).replace(/^'+|'+$/g, "")
  return {
    sheetName: sheetName || undefined,
    reference: trimmed.slice(bangIndex + 1),
  }
}

function getAnchorReference(address) {
  const { sheetName, reference } = splitAddress(address)
  const anchor = reference.split(":")[0]
  return sheetName ? `${sheetName}!${anchor}` : anchor
}

module.exports = {
  formatValuesToText,
  getAnchorReference,
  parseTextToValues,
  splitAddress,
}
