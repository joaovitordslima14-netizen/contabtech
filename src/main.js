import "./style.css";
import ExcelJS from "exceljs";
import {
  CARD_BRANDS,
  CATEGORIES,
  DEBT_TYPES,
  INCOME_TYPES,
  MONTHS,
  PAYMENT_METHODS,
  amountToCents,
  createFutureDebtInstallments,
  emptyState,
  formatMoney,
  getMonthSummary,
  getYearSummary,
  mergeTaxonomy,
  periodOf,
  syncPaymentTransaction,
} from "./finance.js";
import { createXlsxBackup, readXlsxBackupFile, readXlsxImportFile } from "./importer.js";
import {
  getSession,
  isCloudConfigured,
  loadCloudState,
  saveCloudState,
  signIn,
  signOut,
  signUp,
} from "./cloud.js";
import { loadState, openDatabase, saveState } from "./storage.js";

const app = document.querySelector("#app");
const today = new Date();
const PAGE_LABELS = {
  dashboard: "Visão geral",
  transactions: "Movimentações",
  invoices: "Faturas",
  debts: "Dívidas e parcelas",
  investments: "Investimentos",
  data: "Dados e backup",
};

let database;
let state;
let lastPersistedState;
let session = null;
let cloudVersion = null;
let selectedYear = today.getFullYear();
let selectedMonth = today.getMonth();
let currentPage = "dashboard";
let stagedWorkbook = null;
let stagedBackup = null;
let toastTimer;

const esc = (value) =>
  String(value ?? "").replace(/[&<>"']/g, (character) => {
    const replacements = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    };
    return replacements[character];
  });

const period = () => `${selectedYear}-${String(selectedMonth + 1).padStart(2, "0")}`;
const periodLabel = () => `${MONTHS[selectedMonth]} de ${selectedYear}`;
const todayString = () => {
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

function selectedPeriodDate() {
  if (period() === todayString().slice(0, 7)) return todayString();
  return `${period()}-01`;
}

function transactionFingerprint(record) {
  return [
    record.kind,
    record.date,
    record.amountCents,
    record.category || record.incomeType || "",
    record.description || "",
    record.paymentMethod || "",
    record.fixed ? "fixa" : "variável",
  ]
    .join("|")
    .toLocaleLowerCase("pt-BR");
}

function getOptions(values, selected = "") {
  return values
    .map(
      (value) =>
        `<option value="${esc(value)}" ${value === selected ? "selected" : ""}>${esc(value)}</option>`,
    )
    .join("");
}

function showToast(message, kind = "success") {
  const toast = document.querySelector("#toast");
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast toast-${kind} is-visible`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 5000);
}

function requireState() {
  if (!state) throw new Error("Os dados ainda não foram carregados.");
}

async function persist(message = isCloudConfigured ? "Alterações salvas e sincronizadas." : "Alterações salvas neste dispositivo.") {
  requireState();
  try {
    if (isCloudConfigured) {
      cloudVersion = await saveCloudState(session.user.id, state, cloudVersion);
    } else {
      await saveState(database, state);
    }
    lastPersistedState = structuredClone(state);
  } catch (error) {
    if (lastPersistedState) {
      state = structuredClone(lastPersistedState);
      render();
    }
    throw error;
  }
  render();
  showToast(message);
}

function recordsInSelectedPeriod() {
  return state.transactions
    .filter((record) => periodOf(record) === period())
    .sort((left, right) => (right.date || right.period).localeCompare(left.date || left.period));
}

function currentTaxonomy() {
  return {
    categories: mergeTaxonomy({ values: CATEGORIES }, { values: state.taxonomy.categories }).values,
    paymentMethods: mergeTaxonomy({ values: PAYMENT_METHODS }, { values: state.taxonomy.paymentMethods }).values,
    incomeTypes: mergeTaxonomy({ values: INCOME_TYPES }, { values: state.taxonomy.incomeTypes }).values,
    cardBrands: mergeTaxonomy({ values: CARD_BRANDS }, { values: state.taxonomy.cardBrands }).values,
    debtTypes: mergeTaxonomy({ values: DEBT_TYPES }, { values: state.taxonomy.debtTypes }).values,
  };
}

function recordTitle(record) {
  if (record.kind === "income") return record.description || record.incomeType || "Entrada";
  if (record.kind === "investment") return record.description || "Investimento";
  return record.description || record.category || "Saída";
}

function recordKindLabel(record) {
  if (record.kind === "income") return "Entrada";
  if (record.kind === "investment") return "Investimento";
  return record.fixed ? "Saída fixa" : "Saída variável";
}

function dateLabel(date, record) {
  if (!date) return record?.period ? `${record.period.slice(5, 7)}/${record.period.slice(0, 4)} · sem dia` : "—";
  const [year, month, day] = date.split("-");
  return `${day}/${month}/${year}`;
}

function progressBars(items, emptyMessage = "Sem dados neste período.") {
  if (!items.length) return `<p class="empty-inline">${esc(emptyMessage)}</p>`;
  const max = Math.max(...items.map((item) => Math.abs(item.valueCents)), 1);
  return `<ol class="bar-list">${items
    .map((item) => {
      const width = Math.max(0, Math.round((Math.abs(item.valueCents) / max) * 100));
      return `<li class="bar-row">
        <div class="bar-copy"><span>${esc(item.label)}</span><strong>${formatMoney(item.valueCents)}</strong></div>
        <div class="bar-track" aria-hidden="true"><span class="bar-fill" style="width:${width}%"></span></div>
      </li>`;
    })
    .join("")}</ol>`;
}

function simpleTable(headers, rows, emptyText) {
  if (!rows.length) return `<div class="empty-state"><span class="empty-icon" aria-hidden="true">＋</span><p>${esc(emptyText)}</p></div>`;
  return `<div class="table-scroll"><table><thead><tr>${headers.map((header) => `<th>${esc(header)}</th>`).join("")}</tr></thead><tbody>${rows.join("")}</tbody></table></div>`;
}

function transactionTable(records, includeActions = true) {
  const rows = records.map(
    (record) => `<tr>
      <td>${esc(dateLabel(record.date, record))}</td>
      <td><span class="kind-pill kind-${esc(record.kind)}">${esc(recordKindLabel(record))}</span></td>
      <td><strong>${esc(recordTitle(record))}</strong><small>${esc(record.kind === "expense" ? record.category || "Sem categoria" : record.kind === "income" ? record.incomeType || "Tipo não informado" : "Aplicação")}</small></td>
      <td>${esc(record.paymentMethod || "—")}${record.kind === "expense" && record.fixed && record.paid === false ? '<small class="pending-label">Pendente</small>' : ""}</td>
      <td class="amount-cell">${formatMoney(record.amountCents)}</td>
      ${includeActions ? `<td class="row-actions"><button class="icon-button" data-action="edit-transaction" data-id="${esc(record.id)}" aria-label="Editar ${esc(recordTitle(record))}">Editar</button><button class="icon-button danger-text" data-action="delete-transaction" data-id="${esc(record.id)}" aria-label="Excluir ${esc(recordTitle(record))}">Excluir</button></td>` : ""}
    </tr>`,
  );
  return simpleTable(
    ["Data", "Tipo", "Descrição", "Forma", "Valor", ...(includeActions ? ["Ações"] : [])],
    rows,
    "Nenhuma movimentação registrada neste mês.",
  );
}

function metricCard(label, value, note, className = "") {
  return `<article class="metric-card ${className}">
    <span class="metric-label">${esc(label)}</span>
    <strong class="metric-value">${formatMoney(value)}</strong>
    <span class="metric-note">${esc(note)}</span>
  </article>`;
}

function monthComparisonChart(yearSummary) {
  const maxValue = Math.max(
    1,
    ...yearSummary.months.flatMap((month) => [month.expenseCents, Math.abs(month.balanceCents)]),
  );
  return `<div class="month-chart" role="group" aria-label="Despesas e saldo de cada mês de ${yearSummary.year}">
    ${yearSummary.months
      .map((month) => {
        const expenseHeight = month.expenseCents > 0 ? (month.expenseCents / maxValue) * 100 : 0;
        const balanceHeight = month.balanceCents !== 0 ? (Math.abs(month.balanceCents) / maxValue) * 100 : 0;
        return `<button class="month-column ${month.period === period() ? "is-current" : ""}" data-action="select-period" data-period="${month.period}" aria-label="${month.month}: despesas ${formatMoney(month.expenseCents)}, saldo ${formatMoney(month.balanceCents)}">
          <span class="month-bars">
            <i class="month-bar expense-bar" style="height:${expenseHeight}%"></i>
            <i class="month-bar ${month.balanceCents < 0 ? "negative-bar" : "balance-bar"}" style="height:${balanceHeight}%"></i>
          </span>
          <span class="month-tick">${esc(month.month.slice(0, 3))}</span>
        </button>`;
      })
      .join("")}
  </div>
  <div class="chart-legend"><span><i class="legend-dot expense-dot"></i>Saídas no saldo</span><span><i class="legend-dot balance-dot"></i>Saldo do mês (negativo em vermelho)</span></div>`;
}

function cardList(items, emptyText) {
  if (!items.length) return `<div class="empty-state"><span class="empty-icon" aria-hidden="true">＋</span><p>${esc(emptyText)}</p></div>`;
  return `<div class="record-grid">${items
    .map(
      (record) => `<article class="record-card">
        <div class="record-card-heading"><div><span class="record-eyebrow">${esc(record.kind || record.type || record.brand || "Registro")}</span><h3>${esc(record.creditor || record.brand || record.description || "Sem descrição")}</h3></div><strong>${formatMoney(record.amountCents ?? record.totalAmountCents ?? 0)}</strong></div>
        <p>${esc(record.paymentMethod || "Forma não informada")}${record.dueDate ? ` · vence ${dateLabel(record.dueDate)}` : record.date ? ` · ${dateLabel(record.date)}` : ""}</p>
        ${record.paid === false ? '<span class="status-chip">Em aberto</span>' : record.paid === true ? '<span class="status-chip status-paid">Pago</span>' : ""}
        <div class="record-actions"><button class="text-button" data-action="toggle-paid" data-kind="${esc(record._kind)}" data-id="${esc(record.id)}">${record.paid ? "Desfazer pagamento" : "Marcar como paga"}</button><button class="text-button" data-action="edit-record" data-kind="${esc(record._kind)}" data-id="${esc(record.id)}">Editar</button><button class="text-button danger-text" data-action="delete-record" data-kind="${esc(record._kind)}" data-id="${esc(record.id)}">Excluir</button></div>
      </article>`,
    )
    .join("")}</div>`;
}

function dashboardView() {
  const summary = getMonthSummary(state, period());
  const yearSummary = getYearSummary(state, selectedYear);
  const recent = recordsInSelectedPeriod().slice(0, 5);
  const categories = summary.categories.sort((left, right) => right.valueCents - left.valueCents);
  const payments = summary.paymentMethods.sort((left, right) => right.valueCents - left.valueCents);
  const invoices = state.invoices.filter((invoice) => invoice.period === period());
  const fixedNote = summary.unpaidFixedCents
    ? `${formatMoney(summary.unpaidFixedCents)} em despesas fixas pendentes não entra no saldo, conforme a planilha.`
    : "Saldo calculado pelas entradas, saídas pagas e investimentos.";

  return `<section class="page-heading">
    <div><p class="eyebrow">Seu dinheiro, sem complicação</p><h1>${esc(periodLabel())}</h1><p>Resumo fiel aos lançamentos guardados neste dispositivo.</p></div>
    <button class="button button-primary" data-action="new-transaction"><span aria-hidden="true">＋</span> Novo lançamento</button>
  </section>
  <section class="metric-grid" aria-label="Resumo do mês">
    ${metricCard("Entradas", summary.incomeCents, `${summary.transactionCount} movimentações registradas`, "metric-income")}
    ${metricCard("Saídas no saldo", summary.expenseCents, fixedNote, "metric-expense")}
    ${metricCard("Investido", summary.investmentCents, "Incluído no cálculo do saldo do mês.", "metric-investment")}
    ${metricCard("Saldo do mês", summary.balanceCents, "Entradas − saídas pagas − investimentos.", "metric-balance")}
  </section>
  <section class="panel annual-panel">
    <div class="panel-heading"><div><p class="eyebrow">Visão anual</p><h2>Despesas e saldo por mês</h2><p>Selecione uma coluna para abrir aquele mês.</p></div><div class="chart-legend"><span><i class="legend-dot expense-dot"></i>Saídas</span><span><i class="legend-dot balance-dot"></i>Saldo</span></div></div>
    ${monthComparisonChart(yearSummary)}
  </section>
  <section class="chart-grid">
    <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Categorias</p><h2>Gastos por categoria</h2><p>Inclui despesas fixas mesmo quando ainda não pagas, como no gráfico original.</p></div></div>${progressBars(categories)}</article>
    <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Pagamentos</p><h2>Gastos e parcelas por forma</h2><p>Faturas e parcelas informativas aparecem aqui pela forma de pagamento, mas não são somadas às saídas do mês.</p></div></div>${progressBars(payments)}</article>
  </section>
  <section class="chart-grid">
    <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Cartões</p><h2>Faturas por bandeira</h2><p>Informativas: não alteram o saldo nem as saídas.</p></div></div>${progressBars(summary.cardBrands)}</article>
    <article class="panel"><div class="panel-heading"><div><p class="eyebrow">Aplicações</p><h2>Meta e investido</h2><p>Meta do mês: ${formatMoney(Number(state.goals[period()]) || 0)} · total anual investido: ${formatMoney(yearSummary.investmentCents)}</p></div><button class="text-button" data-action="navigate" data-page="investments">Gerenciar investimentos</button></div>
      <div class="goal-progress" aria-label="Investido no mês em relação à meta">
        <div class="goal-progress-heading"><span>Investido</span><strong>${formatMoney(summary.investmentCents)}</strong></div>
        <div class="bar-track"><span class="bar-fill" style="width:${Math.min(100, Number(state.goals[period()]) > 0 ? Math.round((summary.investmentCents / Number(state.goals[period()])) * 100) : 0)}%"></span></div>
        <span class="goal-caption">${Number(state.goals[period()]) > 0 ? `${Math.round((summary.investmentCents / Number(state.goals[period()])) * 100)}% da meta mensal` : "Defina uma meta para este mês, se desejar."}</span>
      </div>
    </article>
  </section>
  <section class="panel">
    <div class="panel-heading"><div><p class="eyebrow">Últimos registros</p><h2>Movimentações do mês</h2></div><button class="text-button" data-action="navigate" data-page="transactions">Ver todas</button></div>
    ${transactionTable(recent, false)}
  </section>`;
}

function transactionsView() {
  const records = recordsInSelectedPeriod();
  return `<section class="page-heading"><div><p class="eyebrow">Histórico pessoal</p><h1>Movimentações</h1><p>Entradas, saídas e investimentos de ${esc(periodLabel())}.</p></div><button class="button button-primary" data-action="new-transaction">＋ Novo lançamento</button></section>
    <section class="panel"><div class="panel-heading"><div><h2>${records.length} ${records.length === 1 ? "registro" : "registros"}</h2><p>Edite ou exclua um registro. A confirmação é solicitada antes de excluir.</p></div></div>${transactionTable(records)}</section>`;
}

function invoicesView() {
  const items = state.invoices
    .filter((invoice) => invoice.period === period())
    .map((invoice) => ({ ...invoice, _kind: "invoice" }));
  return `<section class="page-heading"><div><p class="eyebrow">Apenas informativo</p><h1>Faturas de cartão</h1><p>Os valores cadastrados aqui não entram nas saídas nem no saldo, conforme a observação da planilha.</p></div><button class="button button-primary" data-action="new-record" data-kind="invoice">＋ Nova fatura</button></section>
    <section class="panel"><div class="panel-heading"><div><h2>${esc(periodLabel())}</h2><p>Distribuição das faturas por bandeira.</p></div><strong class="section-total">${formatMoney(items.reduce((sum, invoice) => sum + invoice.amountCents, 0))}</strong></div>${cardList(items, "Nenhuma fatura cadastrada neste mês.")}</section>`;
}

function debtsView() {
  const items = state.debts
    .filter((debt) => debt.period === period())
    .map((debt) => ({
      ...debt,
      amountCents: debt.installmentAmountCents,
      _kind: "debt",
    }));
  return `<section class="page-heading"><div><p class="eyebrow">Compras parceladas e financiamentos</p><h1>Dívidas e parcelas</h1><p>Registros informativos: não alteram o saldo. As parcelas aparecem no gráfico de forma de pagamento, como na planilha.</p></div><button class="button button-primary" data-action="new-record" data-kind="debt">＋ Nova dívida</button></section>
    <section class="panel"><div class="panel-heading"><div><h2>${esc(periodLabel())}</h2><p>${items.length} ${items.length === 1 ? "registro" : "registros"}</p></div></div>${cardList(items, "Nenhuma dívida ou parcela cadastrada neste mês.")}</section>`;
}

function investmentsView() {
  const summary = getMonthSummary(state, period());
  const investments = recordsInSelectedPeriod().filter((record) => record.kind === "investment");
  const goal = Number(state.goals[period()]) || 0;
  return `<section class="page-heading"><div><p class="eyebrow">Acompanhamento mensal</p><h1>Investimentos</h1><p>Aplicações reduzem o saldo do mês, de acordo com a planilha.</p></div><button class="button button-primary" data-action="new-transaction" data-kind="investment">＋ Registrar investimento</button></section>
    <section class="metric-grid metric-grid-two">${metricCard("Investido no mês", summary.investmentCents, periodLabel(), "metric-investment")}${metricCard("Meta mensal", goal, "A meta é opcional e não altera o saldo.", "metric-balance")}</section>
    <section class="panel"><form id="goal-form" class="goal-form"><label for="goalAmount">Meta de investimento para ${esc(periodLabel())}</label><div class="inline-form"><div class="currency-input"><span>R$</span><input id="goalAmount" name="goal" type="number" min="0" step="0.01" value="${(goal / 100).toFixed(2)}" /></div><button class="button button-secondary" type="submit">Salvar meta</button></div></form></section>
    <section class="panel"><div class="panel-heading"><div><h2>Aplicações do mês</h2><p>Local e valor podem ser descritos no lançamento.</p></div></div>${transactionTable(investments)}</section>`;
}

function dataView() {
  const imported = state.transactions.filter((record) => record.sourceFingerprint).length;
  return `<section class="page-heading"><div><p class="eyebrow">Privacidade e portabilidade</p><h1>Dados e backup</h1><p>${isCloudConfigured ? "Seus registros são sincronizados com a sua conta (Supabase) e protegidos por login." : "Seus registros ficam neste navegador. Nada é enviado para uma conta, servidor ou API."}</p></div></section>
    <section class="data-grid">
      <article class="panel data-panel"><div class="panel-heading"><div><p class="eyebrow">Histórico da planilha</p><h2>Importar arquivo Excel</h2><p>Importe a planilha contábil ou um XLSX de movimentações exportado pelo Contabtech.</p></div></div>
        <form id="workbook-form" class="stack-form"><label class="file-picker" for="workbookFile"><span class="file-picker-icon" aria-hidden="true">⇧</span><span><strong>Escolha um arquivo .xlsx</strong><small>O arquivo é lido localmente no navegador.</small></span><input id="workbookFile" name="workbook" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" /></label>
          <label>Ano da importação: todas as datas do arquivo serão ajustadas para este ano<input name="year" type="number" min="1900" max="2200" value="${new Date().getFullYear()}" /></label>
          <div id="import-preview" class="import-preview" aria-live="polite">${stagedWorkbook ? renderImportPreview(stagedWorkbook) : '<p>Após escolher o arquivo, você poderá revisar os registros antes de importar.</p>'}</div>
          <button class="button button-primary" type="submit" ${stagedWorkbook ? "" : "disabled"}>Importar sem substituir dados existentes</button>
        </form>
      </article>
      <article class="panel data-panel"><div class="panel-heading"><div><p class="eyebrow">Cópia de segurança</p><h2>Exportar e restaurar</h2><p>O backup XLSX inclui movimentações, faturas, dívidas, metas e categorias. ${imported ? `${imported} movimentações vieram de uma planilha.` : ""}</p></div></div>
        <div class="backup-actions"><button class="button button-secondary" data-action="export-backup">Baixar backup XLSX</button><button class="button button-secondary" data-action="export-xlsx">Exportar movimentações XLSX</button></div>
        <form id="backup-form" class="stack-form"><label class="file-picker" for="backupFile"><span class="file-picker-icon" aria-hidden="true">↻</span><span><strong>Selecionar backup XLSX</strong><small>Escolha entre mesclar ou substituir após validar o arquivo.</small></span><input id="backupFile" name="backup" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" /></label><div id="backup-preview" class="import-preview" aria-live="polite">${stagedBackup ? renderBackupPreview(stagedBackup) : "<p>Nenhum backup selecionado.</p>"}</div><div class="backup-actions"><button class="button button-secondary" type="submit" name="mode" value="merge" ${stagedBackup ? "" : "disabled"}>Mesclar sem substituir</button><button class="button button-warning" type="submit" name="mode" value="replace" ${stagedBackup ? "" : "disabled"}>Substituir após confirmação</button></div></form>
      </article>
    </section>`;
}

function renderImportPreview(result) {
  const { report } = result;
  const matches = report.comparisons.filter((item) => item.matches).length;
  const mismatches = report.mismatchedPeriods.length;
  return `<div class="preview-stats"><strong>${report.transactionCount} movimentações</strong><strong>${report.invoiceCount} faturas</strong><strong>${report.debtCount} dívidas</strong><strong>Ano identificado: ${report.inferredYear}</strong></div>
    <p>${report.sourceFormat === "XLSX" || report.sourceFormat === "XLSX Backup" ? "Conteúdo da planilha XLSX pronto para importar." : `${matches} meses conferidos com os totais da planilha.${mismatches ? ` ${mismatches} mês(es) apresentam diferença; revise antes de continuar.` : ""}`}</p>
    ${report.undatedRecords ? `<p class="warning-copy">${report.undatedRecords} registro(s) sem dia foram associados ao mês da aba. A data pode ser corrigida depois.</p>` : ""}
    ${report.skippedExamples ? `<p class="muted-copy">${report.skippedExamples} linha(s) de demonstração foram ignoradas.</p>` : ""}
    ${mismatches ? `<ul class="mismatch-list">${report.comparisons.filter((item) => !item.matches).map((item) => `<li>${esc(item.period)}: planilha ${item.spreadsheet.map((value) => formatMoney(value)).join(" / ")}; importação ${item.imported.map((value) => formatMoney(value)).join(" / ")}</li>`).join("")}</ul>` : ""}`;
}

function renderBackupPreview(backup) {
  const data = backup.data;
  return `<p>Backup válido · exportado em ${esc(new Date(backup.exportedAt).toLocaleString("pt-BR"))}.</p><div class="preview-stats"><strong>${data.transactions.length} movimentações</strong><strong>${data.invoices.length} faturas</strong><strong>${data.debts.length} dívidas</strong></div><p>Escolha mesclar ou substituir. A substituição exige uma confirmação adicional.</p>`;
}

function pageContent() {
  if (currentPage === "transactions") return transactionsView();
  if (currentPage === "invoices") return invoicesView();
  if (currentPage === "debts") return debtsView();
  if (currentPage === "investments") return investmentsView();
  if (currentPage === "data") return dataView();
  return dashboardView();
}

function render() {
  if (!state) return;
  const navigation = Object.entries(PAGE_LABELS)
    .map(
      ([page, label]) => `<button class="nav-link ${currentPage === page ? "is-active" : ""}" data-action="navigate" data-page="${page}" ${currentPage === page ? 'aria-current="page"' : ""}><span class="nav-icon" aria-hidden="true">${navIcon(page)}</span><span>${esc(label)}</span></button>`,
    )
    .join("");

  app.innerHTML = `<div class="app-shell">
    <aside class="sidebar">
      <a class="brand" href="#" data-action="navigate" data-page="dashboard"><span class="brand-mark" aria-hidden="true">C</span><span><strong>contabtech</strong><small>finanças pessoais</small></span></a>
      <nav class="primary-nav" aria-label="Navegação principal">${navigation}</nav>
      <div class="sidebar-bottom"><div class="offline-badge"><i></i><span>${isCloudConfigured ? `Sincronizado · ${esc(session?.user?.email || "")}` : "Dados somente neste dispositivo"}</span></div>${isCloudConfigured ? '<button class="button button-compact button-secondary" data-action="sign-out">Sair</button>' : ""}</div>
    </aside>
    <main class="main-area">
      <header class="topbar"><div class="period-control"><button class="icon-button period-arrow" data-action="previous-month" aria-label="Mês anterior">‹</button><label class="visually-hidden" for="monthSelect">Mês</label><select id="monthSelect">${MONTHS.map((month, index) => `<option value="${index}" ${index === selectedMonth ? "selected" : ""}>${month}</option>`).join("")}</select><label class="visually-hidden" for="yearSelect">Ano</label><input id="yearSelect" type="number" min="1900" max="2200" value="${selectedYear}" aria-label="Ano selecionado" /><button class="icon-button period-arrow" data-action="next-month" aria-label="Próximo mês">›</button></div><div class="topbar-right"><span class="privacy-label"><span aria-hidden="true">●</span> ${isCloudConfigured ? "Sincronizado entre dispositivos" : "Privado e offline"}</span><button class="button button-compact button-secondary" data-action="new-transaction">＋ Lançamento</button></div></header>
      <div class="content-area">${pageContent()}</div>
      <footer class="site-footer"><span>${isCloudConfigured ? "Contabtech · dados sincronizados na sua conta" : "Contabtech · seus dados não saem deste dispositivo"}</span><button class="text-button" data-action="navigate" data-page="data">Backup e privacidade</button></footer>
    </main>
  </div>
  <dialog id="formDialog" class="form-dialog"></dialog>
  <div id="toast" class="toast" role="status" aria-live="polite"></div>`;
  bindPageEvents();
}

function navIcon(page) {
  const icons = {
    dashboard: "⌂",
    transactions: "↕",
    invoices: "▤",
    debts: "◷",
    investments: "↗",
    data: "⚙",
  };
  return icons[page] || "·";
}

function bindPageEvents() {
  app.querySelectorAll("[data-action]").forEach((button) => {
    button.addEventListener("click", handleAction);
  });
  app.querySelector("#monthSelect")?.addEventListener("change", (event) => {
    selectedMonth = Number(event.target.value);
    render();
  });
  app.querySelector("#yearSelect")?.addEventListener("change", (event) => {
    const year = Number(event.target.value);
    if (!Number.isInteger(year) || year < 1900 || year > 2200) {
      showToast("Informe um ano entre 1900 e 2200.", "error");
      event.target.value = selectedYear;
      return;
    }
    selectedYear = year;
    render();
  });
  app.querySelector("#goal-form")?.addEventListener("submit", handleGoalSubmit);
  app.querySelector("#workbook-form")?.addEventListener("change", handleWorkbookSelect);
  app.querySelector("#workbook-form")?.addEventListener("submit", handleWorkbookImport);
  app.querySelector("#backup-form")?.addEventListener("change", handleBackupSelect);
  app.querySelector("#backup-form")?.addEventListener("submit", handleBackupRestore);
}

async function handleAction(event) {
  const button = event.currentTarget;
  const action = button.dataset.action;
  try {
    if (action === "navigate") {
      currentPage = button.dataset.page;
      render();
      return;
    }
    if (action === "select-period") {
      const [year, month] = button.dataset.period.split("-").map(Number);
      selectedYear = year;
      selectedMonth = month - 1;
      render();
      return;
    }
    if (action === "previous-month" || action === "next-month") {
      const delta = action === "previous-month" ? -1 : 1;
      const date = new Date(selectedYear, selectedMonth + delta, 1);
      selectedYear = date.getFullYear();
      selectedMonth = date.getMonth();
      render();
      return;
    }
    if (action === "new-transaction") {
      openTransactionDialog(button.dataset.kind || "expense");
      return;
    }
    if (action === "edit-transaction") {
      const record = state.transactions.find((item) => item.id === button.dataset.id);
      if (record) openTransactionDialog(record.kind, record);
      return;
    }
    if (action === "delete-transaction") {
      const record = state.transactions.find((item) => item.id === button.dataset.id);
      if (record && window.confirm(`Excluir "${recordTitle(record)}" no valor de ${formatMoney(record.amountCents)}?`)) {
        state.transactions = state.transactions.filter((item) => item.id !== record.id);
        await persist("Movimentação excluída.");
      }
      return;
    }
    if (action === "new-record") {
      openSupplementDialog(button.dataset.kind);
      return;
    }
    if (action === "edit-record" || action === "delete-record" || action === "toggle-paid") {
      const kind = button.dataset.kind;
      const collection = kind === "invoice" ? "invoices" : "debts";
      const record = state[collection].find((item) => item.id === button.dataset.id);
      if (!record) return;
      if (action === "edit-record") openSupplementDialog(kind, record);
      else if (action === "toggle-paid") {
        record.paid = !record.paid;
        const payment = syncPaymentTransaction(state, kind, record, todayString());
        await persist(payment ? "Marcado como pago; saída criada nas movimentações." : "Pagamento desfeito; saída removida das movimentações.");
      } else if (window.confirm(`Excluir "${record.creditor || record.brand}"?`)) {
        state[collection] = state[collection].filter((item) => item.id !== record.id);
        await persist("Registro excluído.");
      }
      return;
    }
    if (action === "sign-out") {
      await signOut();
      session = null;
      state = null;
      cloudVersion = null;
      renderLogin();
      return;
    }
    if (action === "export-backup") {
      await exportBackup();
      return;
    }
    if (action === "export-xlsx") {
      await exportTransactionsXlsx();
    }
  } catch (error) {
    showToast(error instanceof Error ? error.message : "A operação não pôde ser concluída.", "error");
  }
}

function transactionFormHtml(kind, record = {}) {
  const taxonomy = currentTaxonomy();
  const date = record.date || selectedPeriodDate();
  const isExpense = kind === "expense";
  const isIncome = kind === "income";
  const categories = taxonomy.categories;
  const previousExpense = [...state.transactions]
    .reverse()
    .find((item) => item.kind === "expense" && item.category);
  const previousIncome = [...state.transactions]
    .reverse()
    .find((item) => item.kind === "income" && item.incomeType);
  const previousPayment = [...state.transactions]
    .reverse()
    .find((item) => item.kind === "expense" && item.paymentMethod);
  const selectedCategory = record.category || previousExpense?.category || "";
  const selectedIncomeType = record.incomeType || previousIncome?.incomeType || "";
  const selectedPaymentMethod = record.paymentMethod || previousPayment?.paymentMethod || "";
  const fixed = record.fixed === true || record.fixed === "true";
  const amountValue =
    record.amount ?? (record.id ? (Number(record.amountCents) / 100).toFixed(2) : "");
  return `<dialog id="transactionModal" class="form-dialog">
    <form id="transactionForm" class="modal-card">
      <div class="modal-heading"><div><p class="eyebrow">${record.id ? "Editar registro" : "Lançamento rápido"}</p><h2>${record.id ? "Atualizar movimentação" : "Novo lançamento"}</h2></div><button type="button" class="icon-button" data-close-dialog aria-label="Fechar">×</button></div>
      <input type="hidden" name="id" value="${esc(record.id || "")}" />
      <div class="kind-switch" role="group" aria-label="Tipo do lançamento">${["expense", "income", "investment"].map((item) => `<button type="button" class="${item === kind ? "is-selected" : ""}" data-kind-select="${item}">${item === "expense" ? "Saída" : item === "income" ? "Entrada" : "Investimento"}</button>`).join("")}</div>
      <input type="hidden" name="kind" value="${kind}" />
      <div class="form-grid">
        <label>Data<input name="date" type="date" value="${esc(date)}" required /></label>
        <label>Valor<input name="amount" type="number" min="0.01" step="0.01" inputmode="decimal" placeholder="0,00" value="${esc(amountValue)}" required /></label>
        ${isExpense ? `        <label>Categoria<select name="category" required><option value="">Selecione uma categoria</option>${getOptions(categories, selectedCategory)}</select></label>
          <label>Forma de pagamento<select name="paymentMethod" required><option value="">Selecione</option>${getOptions(taxonomy.paymentMethods, selectedPaymentMethod)}</select></label>
          <label>Tipo de saída<select name="fixed"><option value="false" ${!fixed ? "selected" : ""}>Variável</option><option value="true" ${fixed ? "selected" : ""}>Fixa</option></select></label>
          ${fixed ? `<label class="checkbox-label"><input name="paid" type="checkbox" ${record.paid ? "checked" : ""} /><span>Já foi paga</span></label>` : ""}` : ""}
        ${isIncome ? `<label>Tipo de entrada<select name="incomeType" required><option value="">Selecione o tipo</option>${getOptions(taxonomy.incomeTypes, selectedIncomeType)}</select></label>` : ""}
        ${kind === "investment" ? `<label>Local / descrição<input name="description" type="text" maxlength="120" placeholder="Ex.: Tesouro Direto" value="${esc(record.description || "")}" /></label>` : ""}
        ${(isExpense || isIncome) ? `<label class="wide-field">Descrição<input name="description" type="text" maxlength="120" placeholder="${isIncome ? "Ex.: salário" : "Ex.: compra no mercado"}" value="${esc(record.description || "")}" /></label>` : ""}
      </div>
      <p class="form-error" id="transactionError" aria-live="polite"></p>
      <div class="modal-actions"><button type="button" class="button button-quiet" data-close-dialog>Cancelar</button><button class="button button-primary" type="submit">${record.id ? "Salvar alterações" : "Salvar lançamento"}</button></div>
    </form>
  </dialog>`;
}

function openTransactionDialog(kind = "expense", record = null) {
  const html = transactionFormHtml(kind, record || {});
  app.querySelector("#formDialog").outerHTML = html;
  bindTransactionDialog(app.querySelector("#transactionModal"), kind, record);
}

function bindTransactionDialog(dialog, kind, record) {
  dialog.showModal();
  dialog.querySelectorAll("[data-close-dialog]").forEach((button) => button.addEventListener("click", () => dialog.close()));
  dialog.querySelectorAll("[data-kind-select]").forEach((button) =>
    button.addEventListener("click", () => {
      const values = Object.fromEntries(new FormData(dialog.querySelector("form")));
      const nextKind = button.dataset.kindSelect;
      const nextRecord = {
        ...(record || {}),
        ...values,
        id: record?.id || "",
        fixed: values.fixed === "true",
        paid: values.paid === "on",
      };
      if (values.amount) {
        try {
          nextRecord.amountCents = amountToCents(values.amount);
        } catch {
          nextRecord.amountCents = 0;
        }
      }
      dialog.close();
      dialog.outerHTML = transactionFormHtml(nextKind, nextRecord);
      bindTransactionDialog(app.querySelector("#transactionModal"), nextKind, nextRecord);
    }),
  );
  dialog.querySelector("form").addEventListener("submit", handleTransactionSubmit);
}

async function handleTransactionSubmit(event) {
  event.preventDefault();
  try {
    const data = new FormData(event.currentTarget);
    const kind = String(data.get("kind"));
    const date = String(data.get("date") || "");
    const amountCents = amountToCents(data.get("amount"));
    if (amountCents <= 0) throw new Error("O valor deve ser maior que zero.");
    const id = String(data.get("id") || "");
    const existing = id ? state.transactions.find((record) => record.id === id) : null;
    const record = {
      ...(existing || {}),
      id: existing?.id || crypto.randomUUID(),
      kind,
      date,
      period: date.slice(0, 7),
      amountCents,
    };
    delete record.source;
    delete record.sourceFingerprint;

    if (kind === "expense") {
      record.category = String(data.get("category") || "");
      record.paymentMethod = String(data.get("paymentMethod") || "");
      record.fixed = data.get("fixed") === "true";
      record.paid = record.fixed ? data.get("paid") === "on" : true;
      record.description = String(data.get("description") || "").trim();
      if (!record.category || !record.paymentMethod) throw new Error("Escolha uma categoria e a forma de pagamento.");
    } else if (kind === "income") {
      record.incomeType = String(data.get("incomeType") || "");
      record.description = String(data.get("description") || "").trim();
    } else {
      record.description = String(data.get("description") || "").trim();
    }

    if (!existing) {
      const signature = transactionFingerprint(record);
      const duplicate = state.transactions.find((item) => transactionFingerprint(item) === signature);
      if (duplicate && !window.confirm("Já existe um lançamento igual nesta data. Deseja registrar outro mesmo assim?")) return;
    }
    if (existing) state.transactions = state.transactions.map((item) => (item.id === id ? record : item));
    else state.transactions.push(record);
    await persist(existing ? "Movimentação atualizada." : "Lançamento registrado.");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Não foi possível salvar a movimentação.";
    const errorNode = document.querySelector("#transactionError");
    if (errorNode) errorNode.textContent = message;
    else showToast(message, "error");
  }
}

function supplementFormHtml(kind, record = {}) {
  const taxonomy = currentTaxonomy();
  const invoice = kind === "invoice";
  const title = invoice ? "Fatura de cartão" : "Dívida ou parcela";
  return `<dialog id="supplementModal" class="form-dialog"><form id="supplementForm" class="modal-card">
    <div class="modal-heading"><div><p class="eyebrow">${record.id ? "Editar registro" : "Cadastro informativo"}</p><h2>${record.id ? "Atualizar" : "Nova"} ${title.toLocaleLowerCase("pt-BR")}</h2></div><button type="button" class="icon-button" data-close-dialog aria-label="Fechar">×</button></div>
    <input type="hidden" name="kind" value="${kind}" /><input type="hidden" name="id" value="${esc(record.id || "")}" />
    <div class="form-grid">
      ${invoice ? `<label>Bandeira / instituição<input name="brand" list="brandOptions" maxlength="80" value="${esc(record.brand || "")}" required /><datalist id="brandOptions">${taxonomy.cardBrands.map((brand) => `<option value="${esc(brand)}"></option>`).join("")}</datalist></label>
        <label>Valor da fatura<input name="amount" type="number" min="0" step="0.01" value="${(Number(record.amountCents) / 100 || 0).toFixed(2)}" required /></label>
        <label>Forma de pagamento<select name="paymentMethod"><option value="">Não informado</option>${getOptions(taxonomy.paymentMethods, record.paymentMethod || "")}</select></label>
        <label>Vencimento<input name="dueDate" type="date" value="${esc(record.dueDate || "")}" /></label>
        <label class="checkbox-label"><input name="paid" type="checkbox" ${record.paid ? "checked" : ""} /><span>Fatura paga</span></label>` : `<label>Tipo<select name="type">${getOptions(taxonomy.debtTypes, record.type || taxonomy.debtTypes[0])}</select></label>
        <label>Credor / compra<input name="creditor" maxlength="100" value="${esc(record.creditor || "")}" required /></label>
        <label>Forma de pagamento<select name="paymentMethod"><option value="">Não informado</option>${getOptions(taxonomy.paymentMethods, record.paymentMethod || "")}</select></label>
        <label>Valor da parcela<input name="installmentAmount" type="number" min="0" step="0.01" value="${(Number(record.installmentAmountCents) / 100 || 0).toFixed(2)}" /></label>
        <label>Valor total<input name="totalAmount" type="number" min="0" step="0.01" value="${(Number(record.totalAmountCents) / 100 || 0).toFixed(2)}" /></label>
        <label>Parcela atual<input name="currentInstallment" type="number" min="0" step="1" value="${Number(record.currentInstallment) || ""}" /></label>
        <label>Total de parcelas<input name="totalInstallments" type="number" min="0" step="1" value="${Number(record.totalInstallments) || ""}" /></label>
        <label>Data<input name="date" type="date" value="${esc(record.date || "")}" /></label>
        <label class="checkbox-label"><input name="paid" type="checkbox" ${record.paid ? "checked" : ""} /><span>Parcela paga</span></label>`}
    </div><p class="form-error" id="supplementError" aria-live="polite"></p>
    <div class="modal-actions"><button type="button" class="button button-quiet" data-close-dialog>Cancelar</button><button class="button button-primary" type="submit">Salvar</button></div>
  </form></dialog>`;
}

function openSupplementDialog(kind, record = null) {
  app.querySelector("#formDialog").outerHTML = supplementFormHtml(kind, record || {});
  const dialog = app.querySelector("#supplementModal");
  dialog.showModal();
  dialog.querySelectorAll("[data-close-dialog]").forEach((button) => button.addEventListener("click", () => dialog.close()));
  dialog.querySelector("form").addEventListener("submit", handleSupplementSubmit);
}

async function handleSupplementSubmit(event) {
  event.preventDefault();
  try {
    const data = new FormData(event.currentTarget);
    const kind = String(data.get("kind"));
    const id = String(data.get("id") || "");
    const collection = kind === "invoice" ? "invoices" : "debts";
    const existing = id ? state[collection].find((record) => record.id === id) : null;
    let record;

    if (kind === "invoice") {
      const amountCents = amountToCents(data.get("amount"));
      if (amountCents < 0) throw new Error("O valor da fatura não pode ser negativo.");
      record = {
        ...(existing || {}),
        id: existing?.id || crypto.randomUUID(),
        period: period(),
        brand: String(data.get("brand") || "").trim(),
        paymentMethod: String(data.get("paymentMethod") || ""),
        amountCents,
        dueDate: String(data.get("dueDate") || ""),
        paid: data.get("paid") === "on",
      };
      if (!record.brand) throw new Error("Informe a bandeira ou instituição.");
    } else {
      const installmentAmountCents = amountToCents(data.get("installmentAmount") || "0");
      const totalAmountCents = amountToCents(data.get("totalAmount") || "0");
      const currentInstallment = Number(data.get("currentInstallment")) || 0;
      const totalInstallments = Number(data.get("totalInstallments")) || 0;
      if (installmentAmountCents < 0 || totalAmountCents < 0) throw new Error("Os valores não podem ser negativos.");
      if (
        !Number.isInteger(currentInstallment) ||
        !Number.isInteger(totalInstallments) ||
        currentInstallment < 0 ||
        totalInstallments < 0 ||
        (currentInstallment > 0 && totalInstallments === 0) ||
        (totalInstallments > 0 && currentInstallment === 0) ||
        currentInstallment > totalInstallments
      ) {
        throw new Error("Informe uma parcela atual e um total de parcelas válidos.");
      }
      record = {
        ...(existing || {}),
        id: existing?.id || crypto.randomUUID(),
        period: period(),
        type: String(data.get("type") || ""),
        creditor: String(data.get("creditor") || "").trim(),
        paymentMethod: String(data.get("paymentMethod") || ""),
        installmentAmountCents,
        totalAmountCents,
        currentInstallment,
        totalInstallments,
        date: String(data.get("date") || ""),
        paid: data.get("paid") === "on",
      };
      if (!record.creditor) throw new Error("Informe o credor ou a compra.");
    }

    delete record.source;
    delete record.sourceFingerprint;
    if (kind === "debt") record.seriesId = record.seriesId || crypto.randomUUID();
    if (existing) state[collection] = state[collection].map((item) => (item.id === id ? record : item));
    else state[collection].push(record);
    if (kind === "debt") {
      for (const installment of createFutureDebtInstallments(record)) {
        const sibling = state.debts.find(
          (item) => item.seriesId === record.seriesId && item.period === installment.period,
        );
        if (!sibling) state.debts.push({ ...installment, id: crypto.randomUUID() });
        else if (!sibling.paid) Object.assign(sibling, { ...installment, id: sibling.id });
      }
    }
    const payment = syncPaymentTransaction(state, kind, record, todayString());
    await persist(
      payment && !existing?.paid
        ? "Registro salvo e saída criada nas movimentações."
        : existing ? "Registro atualizado." : "Registro salvo.",
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "Não foi possível salvar este registro.";
    const errorNode = document.querySelector("#supplementError");
    if (errorNode) errorNode.textContent = message;
    else showToast(message, "error");
  }
}

async function handleGoalSubmit(event) {
  event.preventDefault();
  try {
    const value = amountToCents(new FormData(event.currentTarget).get("goal") || "0");
    if (value < 0) throw new Error("A meta não pode ser negativa.");
    state.goals[period()] = value;
    await persist("Meta do mês salva.");
  } catch (error) {
    showToast(error.message, "error");
  }
}

async function handleWorkbookSelect(event) {
  if (event.target.name !== "workbook") return;
  const file = event.target.files?.[0];
  if (!file) return;
  const preview = document.querySelector("#import-preview");
  preview.innerHTML = "<p>Lendo o arquivo localmente…</p>";
  try {
    const year = Number(new FormData(event.currentTarget).get("year"));
    stagedWorkbook = await readXlsxImportFile(file, { year });
    render();
    showToast("Arquivo analisado. Revise o resumo antes de importar.");
  } catch (error) {
    stagedWorkbook = null;
    preview.textContent = error instanceof Error ? error.message : "Não foi possível ler este arquivo.";
    preview.className = "import-preview error-copy";
  }
}

async function handleWorkbookImport(event) {
  event.preventDefault();
  if (!stagedWorkbook) return;
  try {
    const batch = stagedWorkbook;
    const candidates = [...batch.transactions, ...batch.invoices, ...batch.debts];
    const existingSignatures = new Set(
      [
        ...state.transactions,
        ...state.invoices,
        ...state.debts,
      ].map((record) => record.sourceFingerprint || recordFingerprint(record)),
    );
    let skipped = 0;
    for (const record of candidates) {
      const fingerprint = record.sourceFingerprint || recordFingerprint(record);
      if (existingSignatures.has(fingerprint)) {
        skipped += 1;
        continue;
      }
      existingSignatures.add(fingerprint);
      const collection = record.kind === "expense" || record.kind === "income" || record.kind === "investment"
        ? state.transactions
        : record.brand
          ? state.invoices
          : state.debts;
      collection.push({ ...record, id: crypto.randomUUID() });
    }
    state.goals = { ...state.goals, ...batch.goals };
    state.taxonomy = mergeTaxonomy(state.taxonomy, batch.taxonomy);
    stagedWorkbook = null;
    await persist(`Importação concluída. ${skipped} registro(s) duplicado(s) ignorados.`);
  } catch (error) {
    showToast(error instanceof Error ? error.message : "Não foi possível concluir a importação.", "error");
  }
}

function recordFingerprint(record) {
  return [
    record.kind || (record.brand ? "invoice" : "debt"),
    record.period || "",
    record.date || record.dueDate || "",
    record.amountCents ?? record.installmentAmountCents ?? 0,
    record.category || record.brand || record.creditor || record.description || "",
    record.paymentMethod || "",
    record.description || "",
  ]
    .join("|")
    .toLocaleLowerCase("pt-BR");
}

async function handleBackupSelect(event) {
  if (event.target.name !== "backup") return;
  const file = event.target.files?.[0];
  if (!file) return;
  const preview = document.querySelector("#backup-preview");
  try {
    const parsed = await readXlsxBackupFile(file);
    stagedBackup = validateBackup(parsed);
    render();
    showToast("Backup XLSX válido. Escolha se deseja mesclar ou substituir.");
  } catch (error) {
    stagedBackup = null;
    preview.textContent = error instanceof Error ? error.message : "Não foi possível ler o backup XLSX.";
    preview.className = "import-preview error-copy";
  }
}

function validateBackup(backup) {
  if (
    !backup ||
    backup.application !== "Contabtech Finanças" ||
    backup.schemaVersion !== 1 ||
    !backup.data ||
    !Array.isArray(backup.data.transactions) ||
    !Array.isArray(backup.data.invoices) ||
    !Array.isArray(backup.data.debts) ||
    !backup.data.goals ||
    typeof backup.data.goals !== "object"
  ) {
    throw new Error("Arquivo inválido ou versão de backup não compatível.");
  }
  for (const record of [...backup.data.transactions, ...backup.data.invoices, ...backup.data.debts]) {
    if (!record || typeof record !== "object" || !Number.isFinite(Number(record.amountCents ?? record.installmentAmountCents ?? 0))) {
      throw new Error("O backup contém registros inválidos.");
    }
  }
  return backup;
}

async function handleBackupRestore(event) {
  event.preventDefault();
  if (!stagedBackup) return;
  const mode = event.submitter?.value;
  if (mode === "replace" && !window.confirm("Isso substituirá todos os dados locais atuais. Você já salvou um backup? Deseja continuar?")) return;
  try {
    const incoming = stagedBackup.data;
    if (mode === "replace") {
      state = {
        ...emptyState(),
        ...incoming,
        schemaVersion: 1,
        taxonomy: mergeTaxonomy(emptyState().taxonomy, incoming.taxonomy || {}),
      };
    } else {
      mergeBackupData(incoming);
    }
    stagedBackup = null;
    await persist(mode === "replace" ? "Backup restaurado; os dados locais anteriores foram substituídos." : "Backup mesclado; nenhum registro existente foi substituído.");
  } catch (error) {
    showToast(error instanceof Error ? error.message : "Não foi possível restaurar o backup.", "error");
  }
}

function mergeBackupData(incoming) {
  for (const key of ["transactions", "invoices", "debts"]) {
    const current = state[key];
    const fingerprints = new Set(current.map(recordFingerprint));
    for (const record of incoming[key]) {
      const signature = recordFingerprint(record);
      if (fingerprints.has(signature)) continue;
      fingerprints.add(signature);
      current.push({ ...record, id: crypto.randomUUID() });
    }
  }
  state.goals = { ...state.goals, ...incoming.goals };
  state.taxonomy = mergeTaxonomy(state.taxonomy, incoming.taxonomy || {});
}

async function exportBackup() {
  const backupData = {
    schemaVersion: 1,
    transactions: state.transactions,
    invoices: state.invoices,
    debts: state.debts,
    goals: state.goals,
    taxonomy: state.taxonomy,
  };
  const bytes = await createXlsxBackup(backupData);
  downloadFile(
    `contabtech-backup-${todayString()}.xlsx`,
    bytes,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
}

async function exportTransactionsXlsx() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Contabtech";
  const worksheet = workbook.addWorksheet("Movimentações");
  worksheet.columns = [
    { header: "Período", key: "period", width: 12 },
    { header: "Data", key: "date", width: 14 },
    { header: "Tipo", key: "type", width: 18 },
    { header: "Categoria/tipo", key: "category", width: 22 },
    { header: "Descrição", key: "description", width: 32 },
    { header: "Forma de pagamento", key: "paymentMethod", width: 22 },
    { header: "Fixa", key: "fixed", width: 10 },
    { header: "Paga", key: "paid", width: 10 },
    { header: "Valor (BRL)", key: "amount", width: 16 },
  ];
  for (const record of state.transactions) {
    worksheet.addRow({
      period: periodOf(record),
      date: record.date || "",
      type: recordKindLabel(record),
      category: record.category || record.incomeType || "",
      description: record.description || "",
      paymentMethod: record.paymentMethod || "",
      fixed: record.fixed ? "Sim" : "Não",
      paid: record.paid === false ? "Não" : "Sim",
      amount: record.amountCents / 100,
    });
  }
  worksheet.getRow(1).font = { bold: true };
  worksheet.autoFilter = { from: "A1", to: "I1" };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  const bytes = await workbook.xlsx.writeBuffer();
  downloadFile(
    `contabtech-movimentacoes-${todayString()}.xlsx`,
    bytes,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
}

function downloadFile(name, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function registerServiceWorker() {
  if (import.meta.env.PROD && "serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch((error) => {
      console.error("Falha ao habilitar o modo offline:", error);
      showToast(`Modo offline indisponível: ${error.message}`, "error");
    });
  }
}

function renderLogin(message = "") {
  app.innerHTML = `<main class="fatal-error"><span class="brand-mark">C</span><h1>Entrar no Contabtech</h1><p>Entre para sincronizar seus dados entre dispositivos.</p>
    <form id="login-form" class="stack-form">
      <label>E-mail<input name="email" type="email" autocomplete="email" required /></label>
      <label>Senha<input name="password" type="password" minlength="8" autocomplete="current-password" required /></label>
      <p id="login-message" class="error-copy" role="alert">${esc(message)}</p>
      <button class="button button-primary" type="submit" name="mode" value="signin">Entrar</button>
      <button class="button button-secondary" type="submit" name="mode" value="signup">Criar conta</button>
    </form></main><div id="toast" class="toast" role="status" aria-live="polite"></div>`;
  app.querySelector("#login-form").addEventListener("submit", handleLogin);
}

async function handleLogin(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  const email = String(data.get("email")).trim();
  const password = String(data.get("password"));
  const message = form.querySelector("#login-message");
  form.querySelectorAll("button").forEach((button) => (button.disabled = true));
  try {
    session = event.submitter?.value === "signup" ? await signUp(email, password) : await signIn(email, password);
    if (!session) {
      message.textContent = "Conta criada. Confirme o e-mail recebido e depois entre.";
      form.querySelectorAll("button").forEach((button) => (button.disabled = false));
      return;
    }
    await startCloudSession();
    render();
  } catch (error) {
    message.textContent = error instanceof Error ? error.message : "Não foi possível entrar.";
    form.querySelectorAll("button").forEach((button) => (button.disabled = false));
  }
}

async function startCloudSession() {
  const cloud = await loadCloudState();
  if (cloud) {
    state = normalizeState(cloud.value);
    cloudVersion = cloud.version;
  } else {
    // Primeiro acesso: leva para a nuvem os dados que já existiam neste aparelho.
    let local = null;
    try {
      database = await openDatabase();
      local = await loadState(database);
    } catch (error) {
      console.error("Dados locais indisponíveis para migração:", error);
    }
    state = normalizeState(local);
    cloudVersion = await saveCloudState(session.user.id, state, null);
  }
  lastPersistedState = structuredClone(state);
}

async function refreshFromCloud() {
  if (!isCloudConfigured || !session || !state || document.querySelector("dialog[open]")) return;
  try {
    const cloud = await loadCloudState();
    if (cloud && cloud.version !== cloudVersion) {
      state = normalizeState(cloud.value);
      cloudVersion = cloud.version;
      lastPersistedState = structuredClone(state);
      render();
      showToast("Dados atualizados com as alterações de outro dispositivo.");
    }
  } catch (error) {
    console.error(error);
  }
}

async function initialize() {
  try {
    if (isCloudConfigured) {
      session = await getSession();
      registerServiceWorker();
      if (!session) {
        renderLogin();
        return;
      }
      await startCloudSession();
      render();
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible") refreshFromCloud();
      });
      return;
    }
    database = await openDatabase();
    const saved = await loadState(database);
    state = normalizeState(saved);
    if (!saved) await saveState(database, state);
    lastPersistedState = structuredClone(state);
    render();
    registerServiceWorker();
  } catch (error) {
    app.innerHTML = `<main class="fatal-error"><span class="brand-mark">C</span><h1>${isCloudConfigured ? "Não foi possível carregar seus dados" : "Não foi possível abrir os dados locais"}</h1><p>${esc(error instanceof Error ? error.message : "Erro inesperado no armazenamento deste navegador.")}</p><p>${isCloudConfigured ? "Seus dados não foram apagados. Verifique sua conexão e recarregue o aplicativo." : "Seus dados não foram apagados. Verifique a disponibilidade de IndexedDB e recarregue o aplicativo."}</p></main>`;
    console.error(error);
  }
}

function normalizeState(saved) {
  const defaults = emptyState();
  if (!saved) return defaults;
  return {
    ...defaults,
    ...saved,
    schemaVersion: 1,
    transactions: Array.isArray(saved.transactions) ? saved.transactions : [],
    invoices: Array.isArray(saved.invoices) ? saved.invoices : [],
    debts: Array.isArray(saved.debts) ? saved.debts : [],
    goals: saved.goals && typeof saved.goals === "object" ? saved.goals : {},
    taxonomy: mergeTaxonomy(defaults.taxonomy, saved.taxonomy || {}),
  };
}

initialize();
