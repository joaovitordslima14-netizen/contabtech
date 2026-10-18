const DATABASE_NAME = "contabtech-financas";
const STORE_NAME = "app-state";
const STATE_KEY = "main";

function requestResult(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Falha ao acessar o armazenamento local."));
  });
}

export async function openDatabase() {
  if (!("indexedDB" in globalThis)) {
    throw new Error("Este navegador não oferece armazenamento local IndexedDB.");
  }

  const request = indexedDB.open(DATABASE_NAME, 1);
  request.onupgradeneeded = () => {
    const database = request.result;
    if (!database.objectStoreNames.contains(STORE_NAME)) {
      database.createObjectStore(STORE_NAME, { keyPath: "id" });
    }
  };
  return requestResult(request);
}

export async function loadState(database) {
  const transaction = database.transaction(STORE_NAME, "readonly");
  const result = await requestResult(transaction.objectStore(STORE_NAME).get(STATE_KEY));
  return result?.value ?? null;
}

export async function saveState(database, value) {
  const transaction = database.transaction(STORE_NAME, "readwrite");
  transaction.objectStore(STORE_NAME).put({ id: STATE_KEY, value });
  await new Promise((resolve, reject) => {
    transaction.oncomplete = resolve;
    transaction.onerror = () =>
      reject(transaction.error || new Error("Não foi possível salvar os dados neste dispositivo."));
    transaction.onabort = () =>
      reject(transaction.error || new Error("O salvamento local foi cancelado."));
  });
}
