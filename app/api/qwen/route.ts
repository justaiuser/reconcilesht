import { qwenStatus, readActPage } from "@/lib/qwen-ollama";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 800;

export async function GET() {
  return Response.json(await qwenStatus());
}

export async function POST(request: Request) {
  let image = "";
  try {
    const body = (await request.json()) as { image?: unknown };
    image = typeof body.image === "string" ? body.image : "";
  } catch {
    return Response.json({ message: "Не удалось прочитать страницу скана." }, { status: 400 });
  }
  if (!image) return Response.json({ message: "Нет изображения страницы." }, { status: 400 });
  try {
    const page = await readActPage(image);
    return Response.json({ page });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Модель Qwen не смогла прочитать страницу.";
    const status = /не запущен|не скачана|не ответил/i.test(message) ? 503 : 502;
    return Response.json({ message }, { status });
  }
}
