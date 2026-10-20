export const MONTHS = [
  "Janeiro",
  "Fevereiro",
  "Março",
  "Abril",
  "Maio",
  "Junho",
  "Julho",
  "Agosto",
  "Setembro",
  "Outubro",
  "Novembro",
  "Dezembro",
];

export const CATEGORIES = [
  "Bobeira",
  "Alimentação",
  "Saúde",
  "Beleza",
  "Roupa",
  "Lazer",
  "Uber/transporte",
  "Mercado",
  "Aluguel",
  "Assinaturas",
  "Essencial",
  "Outros",
  "Contas",
  "Casa",
  "Filhos",
  "Educação",
  "Pet",
  "Presentes",
  "Viagem",
];

export const PAYMENT_METHODS = ["Crédito", "Débito", "Dinheiro", "Pix", "Boleto"];
export const INCOME_TYPES = ["Renda fixa", "Renda variável"];
export const CARD_BRANDS = [
  "Nu Bank",
  "Santanter",
  "Pic Pay",
  "Inter",
  "C6 Bank",
  "Bradesco",
  "Caixa",
  "BB",
  "Itaú",
  "Mercado Pago",
  "XP",
  "Outros",
];
export const DEBT_TYPES = ["Dívida", "Compra parcelada", "Financiamento"];

export function emptyState() {
  return {
    schemaVersion: 1,
    transactions: [],
    invoices: [],
    debts: [],
    goals: {},
    taxonomy: {
      categories: [...CATEGORIES],
      paymentMethods: [...PAYMENT_METHODS],
      incomeTypes: [...INCOME_TYPES],
      cardBrands: [...CARD_BRANDS],
      debtTypes: [...DEBT_TYPES],
    },
  };
}

export function amountToCents(value) {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("O valor informado não é válido.");
    return Math.round(value * 100);
  }
  const text = String(value ?? "").trim();
  if (!text) throw new TypeError("Informe um valor.");
  const normalized = text.includes(",")
    ? text.replace(/\./g, "").replace(",", ".").replace(/[^\d.-]/g, "")
    : text.replace(/[^\d.-]/g, "");
  const amount = Number(normalized);
  if (!Number.isFinite(amount)) throw new TypeError("O valor informado não é válido.");
  return Math.round(amount * 100);
}

export function formatMoney(cents, options = {}) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    ...options,
  }).format((Number(cents) || 0) / 100);
}

export function periodOf(record) {
  if (record.period && /^\d{4}-\d{2}$/.test(record.period)) return record.period;
  return typeof record.date === "string" ? record.date.slice(0, 7) : "";
}

export function sumCents(items, select = (item) => item.amountCents) {
  return items.reduce((sum, item) => sum + (Number(select(item)) || 0), 0);
}

// Mantém uma saída vinculada a faturas/parcelas: criada ao marcar como paga e removida ao desmarcar.
export function syncPaymentTransaction(state, kind, record, date) {
  const existing = state.transactions.find((item) => item.linkedId === record.id);
  if (!record.paid) {
    if (existing) state.transactions = state.transactions.filter((item) => item !== existing);
    return null;
  }

  const invoice = kind === "invoice";
  const installment = record.totalInstallments
    ? ` (parcela ${record.currentInstallment || "?"}/${record.totalInstallments})`
    : "";
  const amountCents = invoice
    ? record.amountCents
    : record.installmentAmountCents || record.totalAmountCents;
  const fields = {
    kind: "expense",
    amountCents,
    category: "Contas",
    description: invoice ? `Fatura ${record.brand}` : `${record.creditor}${installment}`,
    paymentMethod: record.paymentMethod || "",
    fixed: false,
    paid: true,
    linkedId: record.id,
    linkedKind: kind,
  };

  if (existing) {
    Object.assign(existing, fields);
    return existing;
  }
  const transaction = {
    ...fields,
    id: globalThis.crypto.randomUUID(),
    date,
    period: date.slice(0, 7),
  };
  state.transactions.push(transaction);
  return transaction;
}

export function getMonthSummary(state, period) {
  const transactions = state.transactions.filter((record) => periodOf(record) === period);
  const incomes = transactions.filter((record) => record.kind === "income");
  const expenses = transactions.filter((record) => record.kind === "expense");
  const investments = transactions.filter((record) => record.kind === "investment");
  const fixedExpenses = expenses.filter((record) => record.fixed);
  const variableExpenses = expenses.filter((record) => !record.fixed);
  const paidExpenses = expenses.filter((record) => !record.fixed || record.paid !== false);
  const debts = state.debts.filter((record) => record.period === period);
  const invoices = state.invoices.filter((record) => record.period === period);

  const categories = new Map();
  const paymentMethods = new Map();
  const cardBrands = new Map();

  for (const expense of expenses) {
    const category = expense.category || "Sem categoria";
    categories.set(category, (categories.get(category) || 0) + expense.amountCents);
    const method = expense.paymentMethod || "Sem forma informada";
    paymentMethods.set(method, (paymentMethods.get(method) || 0) + expense.amountCents);
  }

  for (const debt of debts) {
    if (!debt.installmentAmountCents) continue;
    // Parcela paga já aparece como saída vinculada no mesmo mês; evita contá-la duas vezes.
    if (transactions.some((record) => record.linkedId === debt.id)) continue;
    const method = debt.paymentMethod || "Sem forma informada";
    paymentMethods.set(method, (paymentMethods.get(method) || 0) + debt.installmentAmountCents);
  }

  for (const invoice of invoices) {
    if (!invoice.brand) continue;
    cardBrands.set(invoice.brand, (cardBrands.get(invoice.brand) || 0) + invoice.amountCents);
    if (!invoice.amountCents || transactions.some((record) => record.linkedId === invoice.id)) continue;
    const method = invoice.paymentMethod || "Sem forma informada";
    paymentMethods.set(method, (paymentMethods.get(method) || 0) + invoice.amountCents);
  }

  const incomeCents = sumCents(incomes);
  const expenseCents = sumCents(paidExpenses);
  const investmentCents = sumCents(investments);

  return {
    incomeCents,
    expenseCents,
    investmentCents,
    balanceCents: incomeCents - expenseCents - investmentCents,
    allExpenseCents: sumCents(expenses),
    fixedExpenseCents: sumCents(fixedExpenses),
    variableExpenseCents: sumCents(variableExpenses),
    unpaidFixedCents: sumCents(fixedExpenses.filter((record) => record.paid === false)),
    categories: [...categories].map(([label, valueCents]) => ({ label, valueCents })),
    paymentMethods: [...paymentMethods].map(([label, valueCents]) => ({ label, valueCents })),
    cardBrands: [...cardBrands].map(([label, valueCents]) => ({ label, valueCents })),
    transactionCount: transactions.length,
    invoiceCount: invoices.length,
    debtCount: debts.length,
  };
}

export function getYearSummary(state, year) {
  const months = Array.from({ length: 12 }, (_, index) => {
    const period = `${year}-${String(index + 1).padStart(2, "0")}`;
    const summary = getMonthSummary(state, period);
    return {
      period,
      month: MONTHS[index],
      ...summary,
      goalCents: Number(state.goals[period]) || 0,
    };
  });

  return {
    year,
    months,
    incomeCents: sumCents(months, (month) => month.incomeCents),
    expenseCents: sumCents(months, (month) => month.expenseCents),
    investmentCents: sumCents(months, (month) => month.investmentCents),
    balanceCents: sumCents(months, (month) => month.balanceCents),
    goalCents: sumCents(months, (month) => month.goalCents),
  };
}

export function mergeTaxonomy(current, additions = {}) {
  const merged = { ...current };
  for (const [key, values] of Object.entries(additions)) {
    if (!Array.isArray(values)) continue;
    merged[key] = [...new Set([...(current[key] || []), ...values].filter(Boolean))];
  }
  return merged;
}
