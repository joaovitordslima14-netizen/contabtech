import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import {
  createXlsxBackup,
  parseExcelDate,
  parseLegacyWorkbook,
  parseTransactionsWorksheet,
  parseXlsxBackup,
  readXlsxBackupFile,
  readXlsxImportFile,
} from "./importer.js";

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

test("importa movimentações de uma planilha XLSX do Contabtech", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Movimentações");
  worksheet.addRow(["Período", "Data", "Tipo", "Categoria/tipo", "Descrição", "Forma de pagamento", "Fixa", "Paga", "Valor (BRL)"]);
  worksheet.addRow(["2026-10", "03/10/2026", "Saída variável", "Alimentação", "Café; padaria", "Pix", "Não", "Sim", 1234.56]);
  worksheet.addRow(["2026-10", "2026-10-04", "Entrada", "Renda fixa", "Salário", "", "Não", "Sim", 5000]);

  const result = parseTransactionsWorksheet(worksheet, { year: 2025 });
  assert.equal(result.report.inferredYear, 2025);
  assert.equal(result.report.transactionCount, 2);
  assert.equal(result.transactions[0].date, "2025-10-03");
  assert.equal(result.transactions[0].description, "Café; padaria");
  assert.equal(result.transactions[0].amountCents, 123456);
  assert.equal(result.transactions[0].fixed, false);
  assert.equal(result.transactions[1].kind, "income");
  assert.equal(result.transactions[1].incomeType, "Renda fixa");
});

test("apresenta erro para planilha sem os cabeçalhos XLSX suportados", () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Movimentações");
  worksheet.addRow(["Data", "Valor"]);
  worksheet.addRow(["03/10/2026", 12]);

  assert.throws(
    () => parseTransactionsWorksheet(worksheet),
    /Planilha incompatível/,
  );
});

test("preserva todos os dados do backup no XLSX", async () => {
  const data = {
    schemaVersion: 1,
    transactions: [{ kind: "expense", amountCents: 12345, description: "Café ☕" }],
    invoices: [{ brand: "Visa", amountCents: 5000 }],
    debts: [{ creditor: "Banco", installmentAmountCents: 2000 }],
    goals: { "2026-10": 100000 },
    taxonomy: { categories: ["Alimentação"] },
  };
  const bytes = await createXlsxBackup(data);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  const backup = parseXlsxBackup(workbook);

  assert.equal(backup.application, "Contabtech Finanças");
  assert.deepEqual(backup.data, data);

  const file = {
    name: "contabtech-backup.xlsx",
    arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  };
  const loadedBackup = await readXlsxBackupFile(file);
  const importBatch = await readXlsxImportFile(file);
  assert.deepEqual(loadedBackup.data, data);
  assert.equal(importBatch.report.sourceFormat, "XLSX Backup");
  assert.equal(importBatch.transactions[0].description, "Café ☕");
});
