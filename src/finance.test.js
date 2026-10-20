import test from "node:test";
import assert from "node:assert/strict";
import {
  amountToCents,
  emptyState,
  formatMoney,
  getMonthSummary,
  getYearSummary,
  periodOf,
  syncPaymentTransaction,
} from "./finance.js";

test("marcar parcela como paga cria saída vinculada e desmarcar remove", () => {
  const state = emptyState();
  const debt = { id: "d1", creditor: "Loja", installmentAmountCents: 5000, currentInstallment: 2, totalInstallments: 10, paymentMethod: "Crédito", paid: true };
  syncPaymentTransaction(state, "debt", debt, "2025-03-10");
  syncPaymentTransaction(state, "debt", debt, "2025-03-11");
  assert.equal(state.transactions.length, 1);
  assert.equal(state.transactions[0].amountCents, 5000);
  assert.equal(state.transactions[0].description, "Loja (parcela 2/10)");
  debt.paid = false;
  syncPaymentTransaction(state, "debt", debt, "2025-03-10");
  assert.equal(state.transactions.length, 0);
});

test("converte valores de reais para centavos sem erro de ponto flutuante", () => {
  assert.equal(amountToCents("12,34"), 1234);
  assert.equal(amountToCents("1.234,56"), 123456);
  assert.equal(amountToCents(10.1), 1010);
  assert.equal(formatMoney(1010), "R$ 10,10");
});

test("reproduz entradas, saídas pagas, investimentos e saldo mensal", () => {
  const state = emptyState();
  state.transactions = [
    { id: "in", kind: "income", period: "2026-10", amountCents: 100000 },
    { id: "paid-fixed", kind: "expense", period: "2026-10", amountCents: 20000, fixed: true, paid: true, category: "Aluguel" },
    { id: "unpaid-fixed", kind: "expense", period: "2026-10", amountCents: 5000, fixed: true, paid: false, category: "Contas" },
    { id: "variable", kind: "expense", period: "2026-10", amountCents: 1250, fixed: false, category: "Mercado", paymentMethod: "Pix" },
    { id: "investment", kind: "investment", period: "2026-10", amountCents: 10000 },
  ];

  const summary = getMonthSummary(state, "2026-10");
  assert.equal(summary.incomeCents, 100000);
  assert.equal(summary.expenseCents, 21250);
  assert.equal(summary.investmentCents, 10000);
  assert.equal(summary.balanceCents, 68750);
  assert.equal(summary.allExpenseCents, 26250);
  assert.equal(summary.unpaidFixedCents, 5000);
  assert.deepEqual(
    Object.fromEntries(summary.categories.map(({ label, valueCents }) => [label, valueCents])),
    { Aluguel: 20000, Contas: 5000, Mercado: 1250 },
  );
});

test("inclui parcelas no gráfico de formas sem somá-las às saídas ou ao saldo", () => {
  const state = emptyState();
  state.transactions.push({
    id: "expense",
    kind: "expense",
    period: "2026-10",
    amountCents: 1170,
    category: "Mercado",
    paymentMethod: "Pix",
    fixed: false,
  });
  state.debts.push({
    id: "debt",
    period: "2026-10",
    paymentMethod: "Pix",
    installmentAmountCents: 25000,
  });
  state.invoices.push({ id: "invoice", period: "2026-10", brand: "Cartão", paymentMethod: "Crédito", amountCents: 9000 });

  const summary = getMonthSummary(state, "2026-10");
  assert.equal(summary.expenseCents, 1170);
  assert.equal(summary.balanceCents, -1170);
  assert.deepEqual(summary.paymentMethods, [
    { label: "Pix", valueCents: 26170 },
    { label: "Crédito", valueCents: 9000 },
  ]);
  assert.deepEqual(summary.cardBrands, [{ label: "Cartão", valueCents: 9000 }]);
});

test("separa períodos com o mesmo mês em anos diferentes", () => {
  const state = emptyState();
  state.transactions.push(
    { id: "old", kind: "income", date: "2025-10-01", amountCents: 100 },
    { id: "new", kind: "income", date: "2026-10-01", amountCents: 200 },
  );
  assert.equal(periodOf(state.transactions[0]), "2025-10");
  assert.equal(getMonthSummary(state, "2025-10").incomeCents, 100);
  assert.equal(getMonthSummary(state, "2026-10").incomeCents, 200);
  assert.equal(getYearSummary(state, 2026).incomeCents, 200);
});
