import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const isCloudConfigured = Boolean(url && key);

const client = isCloudConfigured ? createClient(url, key) : null;

function requireClient() {
  if (!client) throw new Error("A sincronização na nuvem não está configurada.");
  return client;
}

export async function getSession() {
  const { data, error } = await requireClient().auth.getSession();
  if (error) throw error;
  return data.session;
}

export async function signIn(email, password) {
  const { data, error } = await requireClient().auth.signInWithPassword({ email, password });
  if (error) throw new Error("E-mail ou senha incorretos.", { cause: error });
  return data.session;
}

export async function signUp(email, password) {
  const { data, error } = await requireClient().auth.signUp({ email, password });
  if (error) throw error;
  return data.session;
}

export async function signOut() {
  const { error } = await requireClient().auth.signOut();
  if (error) throw error;
}

export async function loadCloudState() {
  const { data, error } = await requireClient()
    .from("user_state")
    .select("data, version")
    .maybeSingle();
  if (error) throw new Error("Não foi possível carregar os dados da nuvem.", { cause: error });
  return data ? { value: data.data, version: data.version } : null;
}

// Grava só se a versão ainda for a que este dispositivo leu, evitando sobrescrever alterações de outro aparelho.
export async function saveCloudState(userId, value, expectedVersion) {
  const supabase = requireClient();
  if (expectedVersion === null) {
    const { data, error } = await supabase
      .from("user_state")
      .insert({ user_id: userId, data: value, version: 1 })
      .select("version")
      .single();
    if (error) throw new Error("Não foi possível salvar os dados na nuvem.", { cause: error });
    return data.version;
  }
  const { data, error } = await supabase
    .from("user_state")
    .update({ data: value, version: expectedVersion + 1, updated_at: new Date().toISOString() })
    .eq("user_id", userId)
    .eq("version", expectedVersion)
    .select("version");
  if (error) throw new Error("Não foi possível salvar os dados na nuvem.", { cause: error });
  if (!data?.length) {
    throw new Error("Os dados foram alterados em outro dispositivo. Recarregue a página para ver a versão mais recente.");
  }
  return data[0].version;
}
