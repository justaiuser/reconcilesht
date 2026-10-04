import { QWEN_PROMPT, QWEN_SCHEMA, looksLooping, parseQwenResponse, type QwenPage } from "@/lib/qwen-act";

const DEFAULT_HOST = "http://127.0.0.1:11434";
type FetchLike = typeof fetch;

export function ollamaHost(): string {
  return (process.env.OLLAMA_HOST || DEFAULT_HOST).replace(/\/$/, "");
}

export function chooseQwenModel(names: string[], preferred = process.env.QWEN_MODEL): string | null {
  if (preferred) {
    const exact = names.find((name) => name === preferred || name.startsWith(`${preferred}:`));
    if (exact) return exact;
  }
  const visual = names.filter((name) => /qwen3-vl/i.test(name));
  const ranked = [...visual].sort(
    (a, b) =>
      Number(/instruct/i.test(b)) - Number(/instruct/i.test(a)) || modelSize(b) - modelSize(a),
  );
  return ranked[0] ?? null;
}

function modelSize(name: string): number {
  const match = name.match(/:(\d+(?:\.\d+)?)b/i);
  return match ? Number(match[1]) : 0;
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
        message: "Модель Qwen не скачана. На компьютере с программой выполните ollama pull qwen3-vl:8b-instruct.",
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

const IDLE_MS = 90_000;
const LIMIT_MS = 600_000;

export async function readActPage(imageBase64: string, fetchImpl: FetchLike = fetch): Promise<QwenPage> {
  const image = imageBase64.replace(/^data:image\/[a-z0-9.+-]+;base64,/i, "").trim();
  if (!image) throw new Error("Пустая страница скана.");
  if (image.length > 12_000_000) throw new Error("Страница скана слишком большая для модели.");
  const status = await qwenStatus(fetchImpl);
  if (!status.ok || !status.model) throw new Error(status.message);
  return askModel(image, status.model, fetchImpl);
}

async function askModel(image: string, model: string, fetchImpl: FetchLike): Promise<QwenPage> {
  const controller = new AbortController();
  let stop: "idle" | "limit" | null = null;
  const limit = setTimeout(() => {
    stop = "limit";
    controller.abort();
  }, LIMIT_MS);
  let idle = setTimeout(() => {
    stop = "idle";
    controller.abort();
  }, 180_000);
  const bump = () => {
    clearTimeout(idle);
    idle = setTimeout(() => {
      stop = "idle";
      controller.abort();
    }, IDLE_MS);
  };
  try {
    const response = await fetchImpl(`${ollamaHost()}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        model,
        stream: true,
        format: QWEN_SCHEMA,
        think: false,
        messages: [{ role: "user", content: QWEN_PROMPT, images: [image] }],
        options: { temperature: 0, num_predict: 6144, num_ctx: 16384 },
      }),
    });
    bump();
    if (!response.ok || !response.body) {
      const body = (await response.json().catch(() => null)) as { error?: string } | null;
      throw new Error(body?.error || "Модель Qwen не смогла прочитать страницу.");
    }
    const content = await readStream(response.body, bump);
    if (!content.trim()) throw new Error(stoppedMessage(stop));
    const page = parseQwenResponse(content);
    if (stop) {
      page.note = page.note
        ? `${page.note} ${stoppedMessage(stop)}`
        : stoppedMessage(stop);
    }
    return page;
  } catch (error) {
    if (stop) throw new Error(stoppedMessage(stop));
    throw error;
  } finally {
    clearTimeout(limit);
    clearTimeout(idle);
  }
}

function stoppedMessage(stop: "idle" | "limit" | null): string {
  if (stop === "idle") {
    return "Модель замолчала на этой части страницы. Выполните ollama ps: в колонке PROCESSOR должно быть 100% GPU. Если там есть CPU, поставьте модель меньше, ollama pull qwen3-vl:4b-instruct.";
  }
  return "Модель не дописала эту часть страницы. Показаны только строки, которые она успела прочитать.";
}

// Ollama sends the answer as NDJSON chunks; streaming keeps headers flowing so
// Node's fetch does not hit its 300 s header timeout on slow CPUs.
export async function readStream(body: ReadableStream<Uint8Array>, bump?: () => void): Promise<string> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let content = "";
  const take = (line: string) => {
    if (!line.trim()) return;
    let chunk: { error?: string; message?: { content?: string } };
    try {
      chunk = JSON.parse(line);
    } catch {
      return;
    }
    if (chunk.error) throw new Error(chunk.error);
    content += chunk.message?.content ?? "";
  };
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      bump?.();
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";
      lines.forEach(take);
      if (looksLooping(content)) {
        await reader.cancel();
        return content;
      }
    }
  } catch (error) {
    if (content.trim()) return content;
    throw error;
  }
  take(buffer + decoder.decode());
  return content;
}
