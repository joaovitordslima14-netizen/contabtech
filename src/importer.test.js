import test from "node:test";
import assert from "node:assert/strict";
import { parseExcelDate, parseLegacyWorkbook } from "./importer.js";

const cell = (value) => ({ v: value });

test("converte datas do Excel para a data local correta", () => {
  assert.equal(parseExcelDate(46298), "2026-10-03");
  assert.equal(parseExcelDate("03/10/2026"), "2026-10-03");
  assert.equal(parseExcelDate(""), "");
});

test("importa o mês e preserva a correspondência dos totais da planilha", () => {
  const workbook = {
    Sheets: {
      Out: {
        B15: cell(46298),
        C15: cell("Renda variável"),
        D15: cell("Recebimento"),
        E15: cell(40),
        R13: cell("Mercado"),
        S13: cell(46298),
        U13: cell("Padaria"),
        V13: cell("Pix"),
        W13: cell(11.7),
        H54: cell("Dívida"),
        I54: cell("Pix"),
        J54: cell("Credor A"),
        K54: cell(1),
        L54: cell(3),
        M54: cell(250),
        N54: cell(750),
        P54: cell(0),
        B10: cell(40),
        C10: cell(11.7),
        D10: cell(0),
        E10: cell(28.3),
        AC21: cell("Mercado"),
        AC42: cell("Pix"),
      },
    },
  };

  const result = parseLegacyWorkbook(workbook, { fallbackYear: 2025 });
  assert.equal(result.report.inferredYear, 2026);
  assert.equal(result.transactions.length, 2);
  assert.equal(result.invoices.length, 0);
  assert.equal(result.debts.length, 1);
  assert.equal(result.transactions[0].period, "2026-10");
  assert.equal(result.report.comparisons.find((item) => item.period === "2026-10").matches, true);
  assert.equal(result.report.mismatchedPeriods.length, 0);
});

test("não trata valores monetários nem linhas demonstrativas como datas reais", () => {
  const workbook = {
    Sheets: {
      Jan: {
        C15: cell("Renda fixa"),
        D15: cell("Exemplo 1"),
        E15: cell(5000),
        K13: cell("Beleza"),
        M13: cell("exemplo"),
        P13: cell(10),
      },
    },
  };

  const result = parseLegacyWorkbook(workbook, { fallbackYear: 2024 });
  assert.equal(result.report.inferredYear, 2024);
  assert.equal(result.transactions.length, 0);
  assert.ok(result.report.skippedExamples > 0);
});
