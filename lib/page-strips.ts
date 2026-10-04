export type PageStrip = { top: number; bottom: number };

const INK_LUMINANCE = 150;
const ROWS_PER_STRIP = 5;

export function pageStrips(
  pixels: Uint8Array | Uint8ClampedArray,
  width: number,
  height: number,
  channels = 4,
  rowsPerStrip = ROWS_PER_STRIP,
): PageStrip[] {
  if (width < 2 || height < 2) return [{ top: 0, bottom: height }];
  const bands = textRows(pixels, width, height, channels);
  if (bands.length === 0) return [{ top: 0, bottom: Math.max(1, height) }];

  const heights = bands.map((band) => band.bottom - band.top).sort((a, b) => a - b);
  const median = heights[Math.floor(heights.length / 2)] ?? 12;
  const gaps: number[] = [];
  for (let index = 1; index < bands.length; index += 1) {
    gaps.push(bands[index].top - bands[index - 1].bottom);
  }
  gaps.sort((a, b) => a - b);
  const medianGap = gaps[Math.floor(gaps.length / 2)] ?? median;
  const dense = bands.length >= 18 && medianGap < median * 1.15;
  if (!dense) return [{ top: 0, bottom: height }];

  const rows = splitTallBands(bands, median);
  const strips: PageStrip[] = [];
  for (let index = 0; index < rows.length; index += rowsPerStrip) {
    const chunk = rows.slice(index, index + rowsPerStrip);
    const previous = rows[index - 1];
    const next = rows[index + rowsPerStrip];
    const padBefore = previous ? Math.min(8, Math.max(1, Math.floor((chunk[0].top - previous.bottom) / 2))) : 4;
    const last = chunk[chunk.length - 1];
    const padAfter = next ? Math.min(8, Math.max(1, Math.floor((next.top - last.bottom) / 2))) : 4;
    const top = Math.max(0, chunk[0].top - padBefore);
    const bottom = Math.min(height, last.bottom + padAfter);
    const prior = strips[strips.length - 1];
    if (prior && top < prior.bottom) prior.bottom = top;
    if (bottom - top < 24 && prior) {
      prior.bottom = bottom;
      continue;
    }
    strips.push({ top, bottom });
  }
  if (strips.length === 0) return [{ top: 0, bottom: height }];
  strips[0].top = 0;
  strips[strips.length - 1].bottom = height;
  return strips;
}

// Two-column acts put the companies' tables side by side; the darkest vertical rule near the
// middle is their shared border.
export function columnSplit(
  pixels: Uint8Array | Uint8ClampedArray,
  width: number,
  height: number,
  channels = 4,
): number {
  const from = Math.floor(width * 0.35);
  const to = Math.ceil(width * 0.65);
  const sample = height > 900 ? 2 : 1;
  let best = Math.round(width / 2);
  let bestCount = -1;
  for (let x = from; x < to; x += 1) {
    let count = 0;
    for (let y = 0; y < height; y += sample) {
      const index = channels === 1 ? y * width + x : (y * width + x) * channels;
      const luminance =
        channels === 1
          ? pixels[index]
          : pixels[index] * 0.2126 + pixels[index + 1] * 0.7152 + pixels[index + 2] * 0.0722;
      if (luminance < INK_LUMINANCE) count += 1;
    }
    if (count > bestCount) {
      best = x;
      bestCount = count;
    }
  }
  return bestCount >= (height / sample) * 0.4 ? best : Math.round(width / 2);
}

function splitTallBands(bands: PageStrip[], median: number): PageStrip[] {
  const rows: PageStrip[] = [];
  const limit = Math.max(median * 3.5, median + 8);
  const step = Math.max(8, Math.round(median * 1.2));
  for (const band of bands) {
    const height = band.bottom - band.top;
    if (height <= limit) {
      rows.push(band);
      continue;
    }
    for (let top = band.top; top < band.bottom; top += step) {
      rows.push({ top, bottom: Math.min(band.bottom, top + step) });
    }
  }
  return rows;
}

export function textRows(
  pixels: Uint8Array | Uint8ClampedArray,
  width: number,
  height: number,
  channels: number,
): PageStrip[] {
  const dark = new Uint32Array(height);
  const sample = width > 900 ? 2 : 1;
  for (let y = 0; y < height; y += 1) {
    let count = 0;
    const row = y * width * channels;
    for (let x = 0; x < width; x += sample) {
      const index = channels === 1 ? y * width + x : row + x * channels;
      const luminance =
        channels === 1
          ? pixels[index]
          : pixels[index] * 0.2126 + pixels[index + 1] * 0.7152 + pixels[index + 2] * 0.0722;
      if (luminance < INK_LUMINANCE) count += 1;
    }
    dark[y] = count;
  }
  const noise = Math.max(3, Math.floor((width / sample) * 0.004));
  const body = [...dark].filter((count) => count >= noise).sort((a, b) => a - b);
  const crowded = body.length > height * 0.45;
  const limit = crowded ? body[Math.floor(body.length * 0.22)] ?? noise : noise;
  const ink = Array.from(dark, (count) => count >= limit);
  for (let y = 1; y < height - 1; y += 1) {
    if (!ink[y] && ink[y - 1] && ink[y + 1]) ink[y] = true;
  }

  const bands: PageStrip[] = [];
  let top = -1;
  for (let y = 0; y < height; y += 1) {
    if (ink[y] && top < 0) top = y;
    if (!ink[y] && top >= 0) {
      bands.push({ top, bottom: y });
      top = -1;
    }
  }
  if (top >= 0) bands.push({ top, bottom: height });
  return joinCloseBands(bands.filter((band) => band.bottom - band.top >= 2));
}

function joinCloseBands(bands: PageStrip[]): PageStrip[] {
  if (bands.length === 0) return bands;
  const heights = bands.map((band) => band.bottom - band.top).sort((a, b) => a - b);
  const median = heights[Math.floor(heights.length / 2)] ?? 8;
  const joinGap = Math.max(2, Math.round(median * 0.22));
  const joined: PageStrip[] = [{ ...bands[0] }];
  for (let index = 1; index < bands.length; index += 1) {
    const previous = joined[joined.length - 1];
    if (bands[index].top - previous.bottom < joinGap) previous.bottom = bands[index].bottom;
    else joined.push({ ...bands[index] });
  }
  return joined;
}
