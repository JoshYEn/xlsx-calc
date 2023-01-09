"use strict";

const XLSX_CALC = require("../src");
const assert = require("assert");

describe("formula that returns a matrix", () => {
  it("should set a matrix 3x3", () => {
    let workbook = {
      Sheets: {
        Sheet1: {
          E3: { f: "SET_MATRIX()" },
        },
      },
    };
    XLSX_CALC.set_fx("SET_MATRIX", () => {
      return [
        ["aa", "bb", "cc"],
        ["aaa", "bbb", "ccc"],
        ["aaaa", "bbbb", "cccc"],
      ];
    });
    XLSX_CALC(workbook);

    assert.strictEqual(workbook.Sheets.Sheet1.E3.v, "aa");
    assert.strictEqual(workbook.Sheets.Sheet1.E4.v, "aaa");
    assert.strictEqual(workbook.Sheets.Sheet1.E5.v, "aaaa");
    assert.strictEqual(workbook.Sheets.Sheet1.F3.v, "bb");
    assert.strictEqual(workbook.Sheets.Sheet1.F4.v, "bbb");
    assert.strictEqual(workbook.Sheets.Sheet1.F5.v, "bbbb");
    assert.strictEqual(workbook.Sheets.Sheet1.G3.v, "cc");
    assert.strictEqual(workbook.Sheets.Sheet1.G4.v, "ccc");
    assert.strictEqual(workbook.Sheets.Sheet1.G5.v, "cccc");
  });

  it("should replace empty blocks", () => {
    let workbook = {
      Sheets: {
        Sheet1: {
          E3: { f: "SET_MATRIX()" },
        },
      },
    };
    XLSX_CALC.set_fx("SET_MATRIX", () => {
      return [
        ["aa", "bb", "cc"],
        ["aaa", "bbb", "ccc"],
        ["aaaa", "bbbb", ,],
      ];
    });
    XLSX_CALC(workbook);

    assert.strictEqual(workbook.Sheets.Sheet1.E3.v, "aa");
    assert.strictEqual(workbook.Sheets.Sheet1.E4.v, "aaa");
    assert.strictEqual(workbook.Sheets.Sheet1.E5.v, "aaaa");
    assert.strictEqual(workbook.Sheets.Sheet1.F3.v, "bb");
    assert.strictEqual(workbook.Sheets.Sheet1.F4.v, "bbb");
    assert.strictEqual(workbook.Sheets.Sheet1.F5.v, "bbbb");
    assert.strictEqual(workbook.Sheets.Sheet1.G3.v, "cc");
    assert.strictEqual(workbook.Sheets.Sheet1.G4.v, "ccc");
    assert.strictEqual(workbook.Sheets.Sheet1.G5.v, undefined);
  });
});
