const test = require("node:test")
const assert = require("node:assert/strict")

const { formatValuesToText, getAnchorReference, parseTextToValues, splitAddress } = require("./rangeText")

test("parseTextToValues handles single value", () => {
  assert.deepEqual(parseTextToValues("Hello"), [["Hello"]])
})

test("parseTextToValues handles TSV matrix", () => {
  assert.deepEqual(parseTextToValues("1\t2\n3\t4"), [["1", "2"], ["3", "4"]])
})

test("parseTextToValues normalizes jagged rows", () => {
  assert.deepEqual(parseTextToValues("1\t2\n3"), [["1", "2"], ["3", ""]])
})

test("formatValuesToText handles matrix", () => {
  assert.equal(formatValuesToText([["1", "2"], ["3", "4"]]), "1\t2\n3\t4")
})

test("splitAddress handles sheet prefix", () => {
  assert.deepEqual(splitAddress("Sheet1!A1:B2"), { sheetName: "Sheet1", reference: "A1:B2" })
})

test("getAnchorReference extracts top-left cell", () => {
  assert.equal(getAnchorReference("Sheet1!B2:D5"), "Sheet1!B2")
})
