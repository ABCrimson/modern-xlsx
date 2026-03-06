import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('Threaded Comments', () => {
  it('empty sheet has no threaded comments', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.threadedComments).toHaveLength(0);
  });

  it('addThreadedComment creates comment', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'Hello', 'John');
    expect(id).toBeTruthy();
    expect(ws.threadedComments).toHaveLength(1);
    expect(ws.threadedComments[0]?.text).toBe('Hello');
    expect(ws.threadedComments[0]?.refCell).toBe('A1');
  });

  it('replyToComment creates reply chain', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id1 = ws.addThreadedComment('A1', 'Original', 'Alice');
    const id2 = ws.replyToComment(id1, 'Reply 1', 'Bob');
    const id3 = ws.replyToComment(id1, 'Reply 2', 'Alice');
    expect(ws.threadedComments).toHaveLength(3);
    expect(ws.threadedComments[1]?.parentId).toBe(id1);
    expect(ws.threadedComments[2]?.parentId).toBe(id1);
    // Replies should use the same cell ref as parent
    expect(ws.threadedComments[1]?.refCell).toBe('A1');
    expect(ws.threadedComments[2]?.refCell).toBe('A1');
    // IDs should be unique
    expect(id1).not.toBe(id2);
    expect(id2).not.toBe(id3);
  });

  it('shared person across comments', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'First', 'Alice');
    ws.addThreadedComment('B1', 'Second', 'Alice');
    // Same author should reuse person ID
    expect(ws.threadedComments[0]?.personId).toBe(ws.threadedComments[1]?.personId);
  });

  it('replyToComment throws for invalid commentId', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(() => ws.replyToComment('nonexistent', 'text', 'author')).toThrow(
      'Comment nonexistent not found',
    );
  });

  it('different authors create different persons', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'First', 'Alice');
    ws.addThreadedComment('B1', 'Second', 'Bob');
    expect(ws.threadedComments[0]?.personId).not.toBe(ws.threadedComments[1]?.personId);
  });

  it('persons are shared across sheets', () => {
    const wb = new Workbook();
    const ws1 = wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws1.addThreadedComment('A1', 'Hello from 1', 'Alice');
    ws2.addThreadedComment('A1', 'Hello from 2', 'Alice');
    // Same person should be used across sheets
    expect(ws1.threadedComments[0]?.personId).toBe(ws2.threadedComments[0]?.personId);
  });

  it('comment has timestamp', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'Hello', 'John');
    expect(ws.threadedComments[0]?.timestamp).toBeTruthy();
    // Should be a valid ISO date string
    expect(() => new Date(ws.threadedComments[0]?.timestamp)).not.toThrow();
  });
});
