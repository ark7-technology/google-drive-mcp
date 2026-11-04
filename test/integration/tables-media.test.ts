import assert from 'node:assert/strict';
import { describe, it, before, after, beforeEach } from 'node:test';
import { setupTestServer, callTool, type TestContext } from '../helpers/setup-server.js';

describe('Tables & Media tools', () => {
  let ctx: TestContext;

  before(async () => { ctx = await setupTestServer(); });
  after(async () => { await ctx.cleanup(); });
  beforeEach(() => {
    ctx.mocks.drive.tracker.reset();
    ctx.mocks.docs.tracker.reset();
  });

  // --- insertTable ---
  describe('insertTable', () => {
    it('happy path', async () => {
      const res = await callTool(ctx.client, 'insertTable', {
        documentId: 'doc-1', rows: 3, columns: 4, index: 1,
      });
      assert.equal(res.isError, false);
      assert.ok(res.content[0].text!.includes('3x4'));
    });

    it('validation error', async () => {
      const res = await callTool(ctx.client, 'insertTable', {});
      assert.equal(res.isError, true);
    });
  });

  // --- editTableCell ---
  describe('editTableCell', () => {
    it('happy path', async () => {
      ctx.mocks.docs.service.documents.get._setImpl(async () => ({
        data: {
          body: {
            content: [
              {
                table: {
                  tableRows: [
                    {
                      tableCells: [
                        { startIndex: 5, endIndex: 10, content: [{ paragraph: { elements: [{ textRun: { content: 'old\n' } }] } }] },
                      ],
                    },
                  ],
                },
                startIndex: 2,
              },
            ],
          },
        },
      }));
      const res = await callTool(ctx.client, 'editTableCell', {
        documentId: 'doc-1', tableStartIndex: 2, rowIndex: 0, columnIndex: 0,
        textContent: 'new value',
      });
      assert.equal(res.isError, false);
      assert.ok(res.content[0].text!.includes('edited cell'));
    });

    it('validation error', async () => {
      const res = await callTool(ctx.client, 'editTableCell', {});
      assert.equal(res.isError, true);
    });
  });

  // --- insertTableRow / deleteTableRow / insertTableColumn / deleteTableColumn ---
  describe('table rows and columns', () => {
    const cell = (start: number, text: string) => ({
      startIndex: start, endIndex: start + text.length + 2,
      content: [{ paragraph: { elements: [{ textRun: { content: `${text}\n` } }] } }],
    });
    // A 2x2 table at index 2: "If" / "End" style rows above and below content.
    const table2x2 = {
      data: {
        body: {
          content: [
            { startIndex: 1, endIndex: 2, paragraph: { elements: [] } },
            {
              startIndex: 2, endIndex: 30,
              table: { rows: 2, columns: 2, tableRows: [
                { tableCells: [cell(4, 'Part I'), cell(12, '3')] },
                { tableCells: [cell(18, 'Part II'), cell(27, '5')] },
              ] },
            },
          ],
        },
      },
    };
    const lastRequest = () => {
      const calls = ctx.mocks.docs.tracker.getCalls('documents.batchUpdate');
      assert.equal(calls.length, 1);
      const requests = calls[0]!.args[0].requestBody.requests;
      assert.equal(requests.length, 1);
      return requests[0];
    };

    beforeEach(() => {
      ctx.mocks.docs.service.documents.get._setImpl(async () => table2x2);
    });

    it('inserts a row above a reference row, at the table, and returns the cell ranges', async () => {
      const res = await callTool(ctx.client, 'insertTableRow', {
        documentId: 'doc-1', tableStartIndex: 2, rowIndex: 0, insertBelow: false, ifRevisionId: 'rev-1',
      });
      assert.equal(res.isError, false);
      assert.deepEqual(lastRequest(), {
        insertTableRow: { tableCellLocation: { tableStartLocation: { index: 2 }, rowIndex: 0, columnIndex: 0 }, insertBelow: false },
      });
      const calls = ctx.mocks.docs.tracker.getCalls('documents.batchUpdate');
      assert.equal(calls[0]!.args[0].requestBody.writeControl.requiredRevisionId, 'rev-1');
      const text = res.content[0].text!;
      assert.ok(text.includes('Inserted a row above row 0'));
      assert.ok(text.includes('r0c0 [4-12] "Part I"'));
      assert.ok(text.includes('r1c1 [27-30] "5"'));
    });

    it('inserts below by default and scopes to a tab', async () => {
      ctx.mocks.docs.service.documents.get._setImpl(async () => ({
        data: { tabs: [{ tabProperties: { tabId: 't.2' }, documentTab: table2x2.data }] },
      }));
      const res = await callTool(ctx.client, 'insertTableRow', { documentId: 'doc-1', tableStartIndex: 2, rowIndex: 1, tabId: 't.2' });
      assert.equal(res.isError, false);
      assert.deepEqual(lastRequest().insertTableRow, {
        tableCellLocation: { tableStartLocation: { index: 2, tabId: 't.2' }, rowIndex: 1, columnIndex: 0 }, insertBelow: true,
      });
    });

    it('deletes a row; refuses the only row', async () => {
      const res = await callTool(ctx.client, 'deleteTableRow', { documentId: 'doc-1', tableStartIndex: 2, rowIndex: 1 });
      assert.equal(res.isError, false);
      assert.deepEqual(lastRequest(), { deleteTableRow: { tableCellLocation: { tableStartLocation: { index: 2 }, rowIndex: 1, columnIndex: 0 } } });

      ctx.mocks.docs.tracker.reset();
      ctx.mocks.docs.service.documents.get._setImpl(async () => ({
        data: { body: { content: [{ startIndex: 2, table: { rows: 1, columns: 2, tableRows: [{ tableCells: [cell(4, 'a'), cell(8, 'b')] }] } }] } },
      }));
      const only = await callTool(ctx.client, 'deleteTableRow', { documentId: 'doc-1', tableStartIndex: 2, rowIndex: 0 });
      assert.equal(only.isError, true);
      assert.ok(only.content[0].text!.includes('only row'));
      assert.equal(ctx.mocks.docs.tracker.getCalls('documents.batchUpdate').length, 0);
    });

    it('inserts a column left or right, and deletes one', async () => {
      const left = await callTool(ctx.client, 'insertTableColumn', { documentId: 'doc-1', tableStartIndex: 2, columnIndex: 1, insertRight: false });
      assert.equal(left.isError, false);
      assert.deepEqual(lastRequest(), {
        insertTableColumn: { tableCellLocation: { tableStartLocation: { index: 2 }, rowIndex: 0, columnIndex: 1 }, insertRight: false },
      });
      ctx.mocks.docs.tracker.reset();
      const del = await callTool(ctx.client, 'deleteTableColumn', { documentId: 'doc-1', tableStartIndex: 2, columnIndex: 0 });
      assert.equal(del.isError, false);
      assert.deepEqual(lastRequest(), { deleteTableColumn: { tableCellLocation: { tableStartLocation: { index: 2 }, rowIndex: 0, columnIndex: 0 } } });
    });

    it('names the wrong table, row or column before writing anything', async () => {
      const noTable = await callTool(ctx.client, 'insertTableRow', { documentId: 'doc-1', tableStartIndex: 5, rowIndex: 0 });
      assert.ok(noTable.content[0].text!.includes('No table found at index 5'));
      const noRow = await callTool(ctx.client, 'deleteTableRow', { documentId: 'doc-1', tableStartIndex: 2, rowIndex: 7 });
      assert.ok(noRow.content[0].text!.includes('Row 7 not found'));
      const noColumn = await callTool(ctx.client, 'insertTableColumn', { documentId: 'doc-1', tableStartIndex: 2, columnIndex: 3 });
      assert.ok(noColumn.content[0].text!.includes('Column 3 not found'));
      assert.equal(ctx.mocks.docs.tracker.getCalls('documents.batchUpdate').length, 0);
      const invalid = await callTool(ctx.client, 'insertTableRow', { documentId: 'doc-1', tableStartIndex: 2 });
      assert.equal(invalid.isError, true);
    });
  });

  // --- insertImageFromUrl ---
  describe('insertImageFromUrl', () => {
    it('happy path', async () => {
      const res = await callTool(ctx.client, 'insertImageFromUrl', {
        documentId: 'doc-1', imageUrl: 'https://example.com/image.png', index: 1,
      });
      assert.equal(res.isError, false);
      assert.ok(res.content[0].text!.includes('inserted image'));
    });

    it('validation error', async () => {
      const res = await callTool(ctx.client, 'insertImageFromUrl', {});
      assert.equal(res.isError, true);
    });
  });

  // --- insertLocalImage ---
  describe('insertLocalImage', () => {
    it('validation error on missing args', async () => {
      const res = await callTool(ctx.client, 'insertLocalImage', {});
      assert.equal(res.isError, true);
    });
  });
});
