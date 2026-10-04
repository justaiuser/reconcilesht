import { readPageImage } from "@/lib/row-server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

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
    const page = await readPageImage(Buffer.from(image, "base64"));
    return Response.json({ page });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Не удалось прочитать строки страницы.";
    console.error(message);
    return Response.json({ message }, { status: 500 });
  }
}
