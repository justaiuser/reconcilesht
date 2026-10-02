export function stretchContrast(canvas: HTMLCanvasElement): void {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context || canvas.width === 0 || canvas.height === 0) return;
  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  const { data } = image;
  const histogram = new Uint32Array(256);
  const step = Math.max(4, Math.floor(data.length / 4 / 20000)) * 4;
  let samples = 0;
  for (let index = 0; index < data.length; index += step) {
    const luminance = Math.round(
      data[index] * 0.2126 + data[index + 1] * 0.7152 + data[index + 2] * 0.0722,
    );
    histogram[luminance] += 1;
    samples += 1;
  }
  const cutoff = samples * 0.015;
  let seen = 0;
  let low = 0;
  for (let tone = 0; tone < 256; tone += 1) {
    seen += histogram[tone];
    if (seen >= cutoff) {
      low = tone;
      break;
    }
  }
  seen = 0;
  let high = 255;
  for (let tone = 255; tone >= 0; tone -= 1) {
    seen += histogram[tone];
    if (seen >= cutoff) {
      high = tone;
      break;
    }
  }
  if (high - low < 30) return;
  const scale = 255 / (high - low);
  for (let index = 0; index < data.length; index += 4) {
    for (let channel = 0; channel < 3; channel += 1) {
      const next = (data[index + channel] - low) * scale;
      data[index + channel] = next < 0 ? 0 : next > 255 ? 255 : next;
    }
  }
  context.putImageData(image, 0, 0);
}
