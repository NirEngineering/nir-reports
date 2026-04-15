/**
 * docWorker.js – Runs inside a Web Worker.
 *
 * Receives the document payload from the main thread (with photos
 * pre-processed: each photo already has { data, width, height, caption }).
 * Builds and packs the docx file, then posts the resulting Blob back.
 *
 * Running in a Worker keeps the main thread responsive during heavy
 * docx packing (Packer.toBlob can take 1-3 seconds for large documents).
 */
import '../polyfills.js';
import { buildDocument } from './docBuilder.js';

self.onmessage = async ({ data: payload }) => {
  try {
    const blob = await buildDocument(payload);
    // Blob is transferable – no copy overhead
    self.postMessage({ ok: true, blob });
  } catch (err) {
    self.postMessage({ ok: false, error: err?.message ?? String(err) });
  }
};
