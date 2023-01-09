"use strict";

const XLSX_CALC = require("../src");
const assert = require("assert");
const errorValues = {
  "#NULL!": 0x00,
  "#DIV/0!": 0x07,
  "#VALUE!": 0x0f,
  "#REF!": 0x17,
  "#NAME?": 0x1d,
  "#NUM!": 0x24,
  "#N/A": 0x2a,
  "#GETTING_DATA": 0x2b,
};

describe("XLSX_CALC inconsistency test", function () {
  let workbook;
  function create_workbook() {
    return {
      Sheets: {
        Sheet1: {
          A1: {},
          A2: {
            v: 7,
          },
          C2: {
            v: "1",
          },
          C3: {
            v: "",
          },
          C4: {
            v: "",
          },
          C5: {
            v: "",
          },
        },
      },
    };
  }
  beforeEach(function () {
    workbook = create_workbook();
  });

  describe("MIN_new", function () {
    it("finds the min in range", function () {
      workbook.Sheets.Sheet1.A1.f = "MIN(C3:C5)";
      XLSX_CALC(workbook);
      assert.strictEqual(workbook.Sheets.Sheet1.A1.v, 0);
    });
    it("finds the min in range including some negative cell", function () {
      workbook.Sheets.Sheet1.A1.f = "MIN(C3:C5,-A2)";
      XLSX_CALC(workbook);
      assert.strictEqual(workbook.Sheets.Sheet1.A1.v, -7);
    });
    it("finds the min in 2 dimensionnal range", function () {
      workbook.Sheets.Sheet1.A1.f = "MIN(A2:C5)";
      XLSX_CALC(workbook);
      assert.strictEqual(workbook.Sheets.Sheet1.A1.v, 1);
    });
  });
});
