import { QWEN_PROMPT, parseQwenResponse, type QwenPage } from "@/lib/qwen-act";

const DEFAULT_HOST = "http://127.0.0.1:11434";
const DEFAULT_MODEL = "qwen3-vl:2b-instruct";

type FetchLike = typeof fetch;

export function ollamaHost(): string {
  return (process.env.OLLAMA_HOST || DEFAULT_HOST).replace(/\/$/, "");
}

export function chooseQwenModel(names: string[], preferred = process.env.QWEN_MODEL || DEFAULT_MODEL): string | null {
  const exact = names.find((name) => name === preferred || name.startsWith(`${preferred}:`));
  if (exact) return exact;
  const instruct = names.find((name) => /qwen3-vl/i.test(name) && /instruct/i.test(name));
  if (instruct) return instruct;
  return names.find((name) => /qwen3-vl/i.test(name)) ?? null;
}

export async function qwenStatus(fetchImpl: FetchLike = fetch): Promise<{
  ok: boolean;
  model: string | null;
  message: string;
}> {
  try {
    const response = await fetchImpl(`${ollamaHost()}/api/tags`, { signal: AbortSignal.timeout(4000) });
    if (!response.ok) {
      return { ok: false, model: null, message: "Ollama не ответил. Запустите его на компьютере с программой." };
    }
    const body = (await response.json()) as { models?: { name?: string }[] };
    const names = (body.models ?? []).map((model) => model.name).filter((name): name is string => Boolean(name));
    const model = chooseQwenModel(names);
    if (!model) {
      return {
        ok: false,
        model: null,
        message: "Модель Qwen не скачана. На компьютере с программой выполните ollama pull qwen3-vl:2b-instruct.",
      };
    }
    return { ok: true, model, message: `Чтение через ${model}` };
  } catch {
    return {
      ok: false,
      model: null,
      message: "На этом компьютере не запущен Ollama с моделью Qwen.",
    };
  }
}

export async function readActPage(imageBase64: string, fetchImpl: FetchLike = fetch): Promise<QwenPage> {
  const image = imageBase64.replace(/^data:image\/[a-z0-9.+-]+;base64,/i, "").trim();
  if (!image) throw new Error("Пустая страница скана.");
  if (image.length > 12_000_000) throw new Error("Страница скана слишком большая для модели.");
  const status = await qwenStatus(fetchImpl);
  if (!status.ok || !status.model) throw new Error(status.message);
  const response = await fetchImpl(`${ollamaHost()}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    signal: AbortSignal.timeout(180_000),
    body: JSON.stringify({
      model: status.model,
      stream: false,
      format: "json",
      think: false,
      messages: [{ role: "user", content: QWEN_PROMPT, images: [image] }],
      options: { temperature: 0.1, num_predict: 4096 },
    }),
  });
  const body = (await response.json().catch(() => null)) as { error?: string; message?: { content?: string } } | null;
  if (!response.ok) {
    throw new Error(body?.error || "Модель Qwen не смогла прочитать страницу.");
  }
  const content = body?.message?.content ?? "";
  if (!content.trim()) throw new Error("Модель Qwen вернула пустой ответ.");
  return parseQwenResponse(content);
}
