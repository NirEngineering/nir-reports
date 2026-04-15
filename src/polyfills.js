/**
 * polyfills.js
 * Browser compatibility fixes for libraries that assume Node.js 22+ environment.
 *
 * Problem: docx v8.5.0 internally calls Buffer.prototype.toBase64()
 * which is a Node.js 22+ method. The npm 'buffer' polyfill (v6.x) does
 * not include this method, causing "Cd.toBase64 is not a function" at runtime.
 */
import { Buffer } from 'buffer';

// Make Buffer globally available (needed by docx / jszip internals)
if (typeof globalThis.Buffer === 'undefined') {
  globalThis.Buffer = Buffer;
}

// Polyfill Buffer.prototype.toBase64 (Node.js 22+ – missing from browser polyfill)
if (!Buffer.prototype.toBase64) {
  Buffer.prototype.toBase64 = function toBase64() {
    return this.toString('base64');
  };
}

// Polyfill Uint8Array.prototype.toBase64 (new Web API, not yet in all browsers)
if (!Uint8Array.prototype.toBase64) {
  Uint8Array.prototype.toBase64 = function toBase64() {
    let binary = '';
    for (let i = 0; i < this.length; i++) {
      binary += String.fromCharCode(this[i]);
    }
    return btoa(binary);
  };
}
