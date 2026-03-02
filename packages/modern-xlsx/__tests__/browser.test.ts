import { describe, expect, it } from 'vitest';

// Note: These tests verify that browser-relevant APIs exist at the JS runtime level.
// They run in Node.js (not a real browser), so we test runtime support for WebAssembly and Blob.

describe('Browser-compatible APIs', () => {
  it('has WebAssembly support', () => {
    expect(typeof WebAssembly).toBe('object');
    expect(typeof WebAssembly.instantiate).toBe('function');
  });

  it('has Blob support', () => {
    expect(typeof Blob).toBe('function');
  });

  it('can create a Blob', () => {
    const blob = new Blob(['hello'], { type: 'text/plain' });
    expect(blob.size).toBe(5);
    expect(blob.type).toBe('text/plain');
  });
});
