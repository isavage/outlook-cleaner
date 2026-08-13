import test from 'node:test';
import assert from 'node:assert/strict';
import { extractRelativeDateFromText } from './index.js';

test('converts older-than-day expressions into an exact receivedBefore date', () => {
  const fixedNow = new Date('2026-08-12T12:00:00Z');
  const result = extractRelativeDateFromText('Delete newsletters older than 360 days', fixedNow);
  assert.equal(result, '2025-08-17');
});

test('returns null when the prompt has no relative age cutoff', () => {
  const fixedNow = new Date('2026-08-12T12:00:00Z');
  const result = extractRelativeDateFromText('Delete newsletters from Acme', fixedNow);
  assert.equal(result, null);
});
