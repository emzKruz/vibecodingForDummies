const STORAGE_KEY = "ikt-lanelogg";

const loanDialog = document.querySelector("#loanDialog");
const loanForm = document.querySelector("#loanForm");
const loanList = document.querySelector("#loanList");
const emptyState = document.querySelector("#emptyState");
const searchInput = document.querySelector("#searchInput");
const statusFilter = document.querySelector("#statusFilter");

let loans = loadLoans();

function loadLoans() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) ?? [];
  } catch {
    return [];
  }
}

function saveLoans() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(loans));
}

function toDateInputValue(date) {
  const offset = date.getTimezoneOffset();
  return new Date(date.getTime() - offset * 60_000).toISOString().slice(0, 10);
}

function formatDate(dateString) {
  return new Intl.DateTimeFormat("nb-NO", {
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(new Date(`${dateString}T12:00:00`));
}

function isOverdue(loan) {
  return !loan.returned && loan.dueDate < toDateInputValue(new Date());
}

function getStatus(loan) {
  if (loan.returned) return { key: "returned", label: "Levert" };
  if (isOverdue(loan)) return { key: "overdue", label: "Forfalt" };
  return { key: "active", label: "Aktivt lån" };
}

function escapeHtml(value) {
  const element = document.createElement("div");
  element.textContent = value;
  return element.innerHTML;
}

function createId() {
  if (typeof crypto.randomUUID === "function") return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function render() {
  const query = searchInput.value.trim().toLowerCase();
  const filter = statusFilter.value;

  const visibleLoans = loans.filter((loan) => {
    const status = getStatus(loan).key;
    const matchesQuery = `${loan.borrower} ${loan.item} ${loan.serial ?? ""}`
      .toLowerCase()
      .includes(query);
    const matchesStatus = filter === "all" || status === filter;
    return matchesQuery && matchesStatus;
  });

  loanList.innerHTML = visibleLoans.map((loan) => {
    const status = getStatus(loan);
    return `
      <tr>
        <td>
          <span class="item-name">${escapeHtml(loan.item)}</span>
          <span class="item-meta">${escapeHtml(loan.serial || loan.note || "Ikke registrert")}</span>
        </td>
        <td>
          <strong>${escapeHtml(loan.borrower)}</strong>
        </td>
        <td>
          <span class="date-main">Til ${formatDate(loan.dueDate)}</span>
          <span class="date-meta">Lånt ${formatDate(loan.loanDate)}</span>
        </td>
        <td><span class="status status-${status.key}">${status.label}</span></td>
        <td class="actions">
          ${loan.returned
            ? ""
            : `<button class="text-button" type="button" data-action="return" data-id="${loan.id}">Marker levert</button>`}
          <button class="text-button danger" type="button" data-action="delete" data-id="${loan.id}">Slett</button>
        </td>
      </tr>
    `;
  }).join("");

  emptyState.hidden = visibleLoans.length > 0;
  document.querySelector("table").hidden = visibleLoans.length === 0;

  const active = loans.filter((loan) => !loan.returned).length;
  const overdue = loans.filter(isOverdue).length;
  const returned = loans.filter((loan) => loan.returned).length;
  document.querySelector("#activeCount").textContent = active;
  document.querySelector("#overdueCount").textContent = overdue;
  document.querySelector("#returnedCount").textContent = returned;
}

document.querySelector("#openFormButton").addEventListener("click", () => {
  const today = toDateInputValue(new Date());
  const dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + 14);

  loanForm.reset();
  loanForm.elements.loanDate.value = today;
  loanForm.elements.dueDate.value = toDateInputValue(dueDate);
  loanDialog.showModal();
  loanForm.elements.borrower.focus();
});

function closeDialog() {
  loanDialog.close();
}

document.querySelector("#closeFormButton").addEventListener("click", closeDialog);
document.querySelector("#cancelButton").addEventListener("click", closeDialog);

loanDialog.addEventListener("click", (event) => {
  if (event.target === loanDialog) closeDialog();
});

loanForm.addEventListener("submit", (event) => {
  event.preventDefault();
  if (!loanForm.reportValidity()) return;

  const data = new FormData(loanForm);
  const loan = {
    id: createId(),
    borrower: data.get("borrower").trim(),
    item: data.get("item").trim(),
    serial: data.get("serial").trim(),
    loanDate: data.get("loanDate"),
    dueDate: data.get("dueDate"),
    note: data.get("note").trim(),
    returned: false,
  };

  if (loan.dueDate < loan.loanDate) {
    loanForm.elements.dueDate.setCustomValidity("Forfallsdato må være etter utlånsdato.");
    loanForm.elements.dueDate.reportValidity();
    return;
  }

  loanForm.elements.dueDate.setCustomValidity("");
  loans.unshift(loan);
  saveLoans();
  closeDialog();
  render();
});

loanForm.elements.dueDate.addEventListener("input", (event) => {
  event.target.setCustomValidity("");
});

loanList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) return;

  const loan = loans.find((item) => item.id === button.dataset.id);
  if (!loan) return;

  if (button.dataset.action === "return") {
    loan.returned = true;
  }
  if (button.dataset.action === "delete") {
    if (!window.confirm(`Vil du slette lånet for ${loan.borrower}?`)) return;
    loans = loans.filter((item) => item.id !== loan.id);
  }

  saveLoans();
  render();
});

searchInput.addEventListener("input", render);
statusFilter.addEventListener("change", render);

render();
