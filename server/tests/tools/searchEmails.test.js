import { describe, it, expect } from 'vitest';
import { mockGraphClient } from '../helpers/mockGraph.js';
import { searchEmailsTool } from '../../tools/email/searchEmails.js';
import { encodeODataQueryValue } from '../../graph/graphClient.js';

const parse = (res) => JSON.parse(res.content[0].text);

const message = (id, receivedDateTime, extra = {}) => ({
  id,
  subject: `Receipt ${id}`,
  from: { emailAddress: { address: 'billing@vendor.example.com', name: 'Example Billing' } },
  toRecipients: [],
  receivedDateTime,
  bodyPreview: '',
  ...extra
});

async function search(args, responder = () => ({ value: [] })) {
  const { authManager, makeRequest } = mockGraphClient(responder);
  const result = await searchEmailsTool(authManager, args);
  return {
    result,
    makeRequest,
    options: makeRequest.mock.calls[0]?.[1],
    payload: result.isError ? null : parse(result)
  };
}

const idsOf = (payload) => payload.emails.map((email) => email.id);

// The $search value Graph receives: the KQL expression wrapped in one pair of
// double quotes, with inner quotes backslash-escaped.
const sent = (kql) => `"${kql.replace(/"/g, '\\"')}"`;

describe('searchEmailsTool KQL query construction', () => {
  it('sends property restrictions unquoted with ISO dates', async () => {
    const { options, payload, makeRequest } = await search({
      query: 'receipt',
      startDate: '2026-10-01T00:00:00Z',
      endDate: '2026-10-31T23:59:59Z'
    });

    expect(makeRequest.mock.calls[0][0]).toBe('/me/messages');
    // widened by one day on each side; the client-side guard trims to the exact window
    expect(options.search).toBe(sent('receipt AND received:2026-09-30..2026-11-01'));
    expect(options.search.slice(1, -1)).not.toContain('"');
    // $filter / $orderby cannot be combined with $search on messages
    expect(options.filter).toBeUndefined();
    expect(options.orderby).toBeUndefined();
    expect(payload.searchSummary.searchApproach).toMatch(/KQL/);
    expect(payload.searchSummary.kqlQuery).toBe(options.search);
    expect(payload.searchSummary.parameters.dateRange).toBe('KQL: received:2026-09-30..2026-11-01');
    expect(payload.searchSummary.kqlWindowWidened).toBe(true);
  });

  it('accepts date-only bounds', async () => {
    const { options } = await search({ query: 'receipt', startDate: '2026-10-01', endDate: '2026-10-31' });
    expect(options.search).toBe(sent('receipt AND received:2026-09-30..2026-11-01'));
  });

  it('uses comparison operators for one-sided ranges', async () => {
    const startOnly = await search({ query: 'receipt', startDate: '2026-10-01T00:00:00Z' });
    expect(startOnly.options.search).toBe(sent('receipt AND received>=2026-09-30'));

    const endOnly = await search({ query: 'receipt', endDate: '2026-10-31' });
    expect(endOnly.options.search).toBe(sent('receipt AND received<=2026-11-01'));
  });

  it('takes the UTC date part of offset date-times', async () => {
    const { options } = await search({
      query: 'receipt',
      startDate: '2026-10-01T23:30:00-05:00',
      endDate: '2026-11-01T01:00:00+02:00'
    });
    // 2026-10-02T04:30Z .. 2026-10-31T23:00Z, widened by one day on each side
    expect(options.search).toBe(sent('receipt AND received:2026-10-01..2026-11-01'));
  });

  it('does not depend on the host timezone', async () => {
    const savedTz = process.env.TZ;
    try {
      const queries = [];
      for (const tz of ['Pacific/Pago_Pago', 'Pacific/Kiritimati', 'UTC']) {
        process.env.TZ = tz;
        const { options } = await search({
          query: 'receipt',
          startDate: '2026-10-01T00:00:00Z',
          endDate: '2026-10-31T23:59:59Z'
        });
        queries.push(options.search);
      }
      expect(new Set(queries)).toEqual(new Set([sent('receipt AND received:2026-09-30..2026-11-01')]));
    } finally {
      if (savedTz === undefined) delete process.env.TZ;
      else process.env.TZ = savedTz;
    }
  });

  it('quotes a free-text term only when it contains whitespace', async () => {
    const single = await search({ query: 'receipt' });
    expect(single.options.search).toBe(sent('receipt'));

    const multi = await search({ query: 'annual report' });
    expect(multi.options.search).toBe(sent('"annual report"'));

    const quoted = await search({ query: 'say "hello" now' });
    expect(quoted.options.search).toBe(sent('"say hello now"'));
  });

  it('quotes a term containing KQL syntax characters', async () => {
    const { options } = await search({ query: 'ref:1234' });
    expect(options.search).toBe(sent('"ref:1234"'));
  });

  it('emits from: and subject: restrictions, quoting only the value and only when needed', async () => {
    const plain = await search({
      query: 'receipt',
      from: 'billing@vendor.example.com',
      subject: 'statement'
    });
    expect(plain.options.search).toBe(sent('receipt AND from:billing@vendor.example.com AND subject:statement'));

    const spaced = await search({
      query: 'receipt',
      from: 'Example Billing',
      subject: 'monthly "pro" statement'
    });
    expect(spaced.options.search).toBe(sent('receipt AND from:"Example Billing" AND subject:"monthly pro statement"'));
    expect(spaced.options.search).not.toContain('\\"from:');
    expect(spaced.options.search).not.toContain('\\"subject:');
  });

  it('combines term, sender, subject and dates in a fixed order', async () => {
    const { options, payload } = await search({
      query: 'receipt',
      from: 'billing@vendor.example.com',
      subject: 'statement',
      startDate: '2026-10-01',
      endDate: '2026-10-31'
    });
    expect(options.search).toBe(sent(
      'receipt AND from:billing@vendor.example.com AND subject:statement AND received:2026-09-30..2026-11-01'
    ));
    expect(payload.searchSummary.parameters.sender).toBe('KQL: from:billing@vendor.example.com');
    expect(payload.searchSummary.parameters.subject).toBe('KQL: subject:statement');
  });

  it('survives query-string encoding unchanged apart from the plus sign', () => {
    const kql = 'receipt AND from:"Example Billing" AND received>=2026-10-01 AND received:2026-10-01..2026-10-31';
    expect(encodeODataQueryValue(kql)).toBe(kql);
    expect(encodeODataQueryValue('from:billing+statements@vendor.example.com'))
      .toBe('from:billing%2Bstatements@vendor.example.com');
  });

  it('rejects an unparseable date without calling Graph', async () => {
    const { result, makeRequest } = await search({ query: 'receipt', startDate: 'last tuesday' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/startDate/);
    expect(makeRequest).not.toHaveBeenCalled();
  });
});

describe('searchEmailsTool KQL date window widening', () => {
  it('reports exactly what was sent and flags the widening', async () => {
    const { options, payload } = await search({ query: 'receipt', startDate: '2026-10-01', endDate: '2026-10-31' });
    expect(payload.searchSummary.kqlQuery).toBe(options.search);
    expect(payload.searchSummary.kqlWindowWidened).toBe(true);
  });

  it('widens a one-sided range on its open side only', async () => {
    const startOnly = await search({ query: 'receipt', startDate: '2026-03-01' });
    expect(startOnly.options.search).toBe(sent('receipt AND received>=2026-02-28'));
    expect(startOnly.payload.searchSummary.kqlWindowWidened).toBe(true);

    const endOnly = await search({ query: 'receipt', endDate: '2026-12-31' });
    expect(endOnly.options.search).toBe(sent('receipt AND received<=2027-01-01'));
    expect(endOnly.payload.searchSummary.kqlWindowWidened).toBe(true);
  });

  it('is not flagged when there is no date restriction or no KQL search', async () => {
    const noDates = await search({ query: 'receipt' });
    expect(noDates.options.search).toBe(sent('receipt'));
    expect(noDates.payload.searchSummary.kqlWindowWidened).toBeUndefined();

    const odata = await search({ from: 'billing@vendor.example.com', startDate: '2026-10-01', endDate: '2026-10-31' });
    expect(odata.options.search).toBeUndefined();
    expect(odata.payload.searchSummary.kqlWindowWidened).toBeUndefined();
  });

  it('keeps a message on the requested end day and drops one a day later', async () => {
    // A service that honours the widened window returns the boundary days too
    const value = [
      message('day-before-start', '2026-09-30T20:00:00Z'),
      message('on-start-day', '2026-10-01T00:30:00Z'),
      message('on-end-day', '2026-10-31T22:00:00Z'),
      message('day-after-end', '2026-11-01T09:00:00Z')
    ];
    const { options, payload } = await search(
      { query: 'receipt', startDate: '2026-10-01', endDate: '2026-10-31' },
      () => ({ value })
    );

    expect(options.search).toBe(sent('receipt AND received:2026-09-30..2026-11-01'));
    expect(idsOf(payload)).toEqual(['on-start-day', 'on-end-day']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(2);
  });
});

describe('searchEmailsTool result guard (KQL mode)', () => {
  const window = { startDate: '2026-10-01T00:00:00Z', endDate: '2026-10-31T23:59:59Z' };

  it('drops messages outside the requested window and counts them', async () => {
    const value = [
      message('start-edge', '2026-10-01T00:00:00Z'),
      message('mid', '2026-10-15T12:00:00Z'),
      message('end-edge', '2026-10-31T23:59:59Z'),
      message('before', '2026-09-30T23:59:59Z'),
      message('after', '2026-11-01T00:00:00Z'),
      message('ancient', '2025-01-01T00:00:00Z'),
      message('undated', undefined)
    ];
    const { payload } = await search({ query: 'receipt', ...window }, () => ({ value }));

    expect(idsOf(payload)).toEqual(['start-edge', 'mid', 'end-edge']);
    expect(payload.searchSummary.totalResults).toBe(3);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(4);
    expect(payload.searchSummary.kqlQuery).toBe(sent('receipt AND received:2026-09-30..2026-11-01'));

    const from = Date.parse(window.startDate);
    const to = Date.parse(window.endDate);
    for (const email of payload.emails) {
      const received = Date.parse(email.receivedDateTime);
      expect(received).toBeGreaterThanOrEqual(from);
      expect(received).toBeLessThanOrEqual(to);
    }
  });

  it('treats a date-only end date as the whole end day', async () => {
    const value = [
      message('late-on-end-day', '2026-10-31T23:30:00Z'),
      message('next-midnight', '2026-11-01T00:00:00Z')
    ];
    const { payload } = await search(
      { query: 'receipt', startDate: '2026-10-01', endDate: '2026-10-31' },
      () => ({ value })
    );
    expect(idsOf(payload)).toEqual(['late-on-end-day']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(1);
  });

  it('treats an explicit date-time end as inclusive and nothing later', async () => {
    const value = [
      message('at-end', '2026-10-31T12:00:00Z'),
      message('after-end', '2026-10-31T12:00:01Z')
    ];
    const { payload } = await search(
      { query: 'receipt', endDate: '2026-10-31T12:00:00Z' },
      () => ({ value })
    );
    expect(idsOf(payload)).toEqual(['at-end']);
  });

  it('compares instants, so offset bounds are honoured', async () => {
    const value = [
      message('inside', '2026-09-30T22:30:00Z'),
      message('outside', '2026-09-30T21:30:00Z')
    ];
    const { payload } = await search(
      { query: 'receipt', startDate: '2026-10-01T00:00:00+02:00' },
      () => ({ value })
    );
    expect(idsOf(payload)).toEqual(['inside']);
  });

  it('applies one-sided bounds', async () => {
    const value = [
      message('old', '2026-01-01T00:00:00Z'),
      message('new', '2026-12-01T00:00:00Z')
    ];
    const startOnly = await search({ query: 'receipt', startDate: '2026-06-01' }, () => ({ value }));
    expect(idsOf(startOnly.payload)).toEqual(['new']);

    const endOnly = await search({ query: 'receipt', endDate: '2026-06-01' }, () => ({ value }));
    expect(idsOf(endOnly.payload)).toEqual(['old']);
  });

  it('drops messages that fail a requested sender restriction (address or name, any case)', async () => {
    const other = { emailAddress: { address: 'newsletter@other.example.net', name: 'Other Sender' } };
    const value = [
      message('by-address', '2026-10-05T00:00:00Z', {
        from: { emailAddress: { address: 'Billing@Vendor.Example.com', name: 'Someone' } }
      }),
      message('by-name', '2026-10-06T00:00:00Z', {
        from: { emailAddress: { address: 'x@y.example.org', name: 'Vendor Billing Team' } }
      }),
      message('other', '2026-10-07T00:00:00Z', { from: other }),
      message('no-sender', '2026-10-08T00:00:00Z', { from: undefined })
    ];
    const { payload } = await search({ query: 'receipt', from: 'vendor' }, () => ({ value }));

    expect(idsOf(payload)).toEqual(['by-address', 'by-name']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(2);
  });

  it('drops messages that fail a requested subject restriction (case-insensitive contains)', async () => {
    const value = [
      message('match', '2026-10-05T00:00:00Z', { subject: 'Your Monthly STATEMENT is ready' }),
      message('miss', '2026-10-06T00:00:00Z', { subject: 'Weekly digest' }),
      message('no-subject', '2026-10-07T00:00:00Z', { subject: undefined })
    ];
    const { payload } = await search({ query: 'receipt', subject: 'statement' }, () => ({ value }));

    expect(idsOf(payload)).toEqual(['match']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(2);
  });

  it('drops nothing when no restriction was requested, including undated messages', async () => {
    const value = [message('a', undefined), message('b', '1999-01-01T00:00:00Z')];
    const { payload } = await search({ query: 'receipt' }, () => ({ value }));

    expect(idsOf(payload)).toEqual(['a', 'b']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(0);
  });

  it('warns that results may be incomplete when messages had to be dropped', async () => {
    const value = [
      message('in', '2026-10-15T00:00:00Z'),
      message('out-1', '2026-08-15T00:00:00Z'),
      message('out-2', '2026-12-15T00:00:00Z')
    ];
    const { payload } = await search({ query: 'receipt', limit: 3, ...window }, () => ({ value }));

    const { warning, droppedOutsideFilters } = payload.searchSummary;
    expect(droppedOutsideFilters).toBe(2);
    expect(warning).toContain('returned 2 messages outside the requested filters');
    expect(warning).toContain('removed');
    expect(warning).toContain('smaller than the requested limit');
    expect(warning).toContain('possibly incomplete');
    expect(warning).toMatch(/narrow the date window or raise limit/);
  });

  it('uses the singular for one dropped message', async () => {
    const value = [message('in', '2026-10-15T00:00:00Z'), message('out', '2026-08-15T00:00:00Z')];
    const { payload } = await search({ query: 'receipt', ...window }, () => ({ value }));
    expect(payload.searchSummary.warning).toContain('returned 1 message outside the requested filters');
    expect(payload.searchSummary.warning).toContain('it was removed');
  });

  it('does not warn about margin-day messages the widened KQL window returns by design', async () => {
    const value = [
      message('in', '2026-10-15T00:00:00Z'),
      message('margin-before', '2026-09-30T12:00:00Z'),
      message('margin-after', '2026-11-01T12:00:00Z')
    ];
    const { payload } = await search({ query: 'receipt', ...window }, () => ({ value }));
    expect(payload.searchSummary.droppedOutsideFilters).toBe(2);
    expect(payload.searchSummary.warning).toBeUndefined();
    expect(payload.emails.map(e => e.id)).toEqual(['in']);
  });

  it('has no warning when nothing was dropped', async () => {
    const value = [message('in', '2026-10-15T00:00:00Z')];
    const { payload } = await search({ query: 'receipt', ...window }, () => ({ value }));
    expect(payload.searchSummary.droppedOutsideFilters).toBe(0);
    expect(payload.searchSummary.warning).toBeUndefined();
  });

  it('also warns for a folder ($filter) search', async () => {
    const value = [
      message('in', '2026-10-15T00:00:00Z'),
      message('out', '2026-08-15T00:00:00Z')
    ];
    const { payload } = await search({ folders: ['Receipts'], ...window }, () => ({ value }));
    expect(idsOf(payload)).toEqual(['in']);
    expect(payload.searchSummary.warning).toContain('returned 1 message outside the requested filters');
  });
});

describe('searchEmailsTool folder restrictions', () => {
  it('binds a single folder through its endpoint', async () => {
    const { makeRequest, options, payload } = await search({
      query: 'receipt',
      folders: ['Receipts'],
      startDate: '2026-10-01',
      endDate: '2026-10-31'
    });

    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest.mock.calls[0][0]).toBe('/me/mailFolders/Receipts/messages');
    expect(options.search).toBeUndefined();
    expect(options.filter).toContain('receivedDateTime ge 2026-10-01T00:00:00Z');
    expect(options.filter).toContain('receivedDateTime lt 2026-11-01T00:00:00Z');
    expect(payload.searchSummary.endpoint).toBe('/me/mailFolders/Receipts/messages');
    expect(payload.searchSummary.folders).toEqual(['Receipts']);
  });

  it('parenthesises the free-text alternatives so they cannot bypass other restrictions', async () => {
    const { options } = await search({
      query: 'receipt',
      from: 'billing@vendor.example.com',
      folders: ['Receipts'],
      startDate: '2026-10-01',
      endDate: '2026-10-31'
    });

    expect(options.filter).toBe([
      'receivedDateTime ge 2026-10-01T00:00:00Z',
      'receivedDateTime lt 2026-11-01T00:00:00Z',
      "from/emailAddress/address eq 'billing@vendor.example.com'",
      "(contains(subject,'receipt') or contains(body/content,'receipt'))"
    ].join(' and '));
  });

  it('still drops out-of-window, wrong-sender or wrong-subject messages from a folder search', async () => {
    const value = [
      message('ok', '2026-10-10T00:00:00Z'),
      message('too-old', '2026-08-10T00:00:00Z'),
      message('wrong-sender', '2026-10-11T00:00:00Z', {
        from: { emailAddress: { address: 'someone@other.example.net', name: 'Someone' } }
      }),
      message('wrong-subject', '2026-10-12T00:00:00Z', { subject: 'Unrelated' })
    ];
    const { payload } = await search({
      subject: 'receipt',
      from: 'billing@vendor.example.com',
      folders: ['Receipts'],
      startDate: '2026-10-01',
      endDate: '2026-10-31'
    }, () => ({ value }));

    expect(idsOf(payload)).toEqual(['ok']);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(3);
  });

  it('queries every folder of a multi-folder search instead of falling back to all folders', async () => {
    const byPath = {
      '/me/mailFolders/Receipts/messages': [
        message('r1', '2026-10-03T00:00:00Z'),
        message('r2', '2026-10-20T00:00:00Z')
      ],
      '/me/mailFolders/Invoices/messages': [
        message('i1', '2026-10-10T00:00:00Z'),
        message('i2', '2026-10-25T00:00:00Z')
      ]
    };
    const { makeRequest, payload } = await search(
      { query: 'receipt', folders: ['Receipts', 'Invoices'], limit: 3 },
      (path) => ({ value: byPath[path] || [] })
    );

    const paths = makeRequest.mock.calls.map((call) => call[0]);
    expect(paths).toEqual([
      '/me/mailFolders/Receipts/messages',
      '/me/mailFolders/Invoices/messages'
    ]);
    expect(paths).not.toContain('/me/messages');
    // merged in receivedDateTime desc order, capped at the limit
    expect(idsOf(payload)).toEqual(['i2', 'r2', 'i1']);
    expect(payload.searchSummary.endpoint).toEqual(paths);
  });

  it('queries a folder given twice only once', async () => {
    const { makeRequest } = await search({ query: 'receipt', folders: ['Receipts', 'Receipts'] });
    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest.mock.calls[0][0]).toBe('/me/mailFolders/Receipts/messages');
  });

  it('returns the first error from a folder request', async () => {
    const failure = { content: [{ type: 'text', text: 'boom' }], isError: true };
    const { result } = await search(
      { query: 'receipt', folders: ['Receipts', 'Invoices'] },
      (path) => (path.includes('Invoices') ? failure : { value: [] })
    );
    expect(result).toBe(failure);
  });
});

// A service that honours $top, like Graph does: it never returns more than the page size asked for.
const pagedService = (pool) => (path, options) => ({ value: pool.slice(0, options.top) });

describe('searchEmailsTool page size with a widened KQL window', () => {
  const window = { startDate: '2026-10-01', endDate: '2026-10-31' };
  const marginBefore = (n) => message(`margin-before-${n}`, '2026-09-30T12:00:00Z');
  const marginAfter = (n) => message(`margin-after-${n}`, '2026-11-01T12:00:00Z');
  const inWindow = (n) => message(`in-${n}`, `2026-10-${String(10 + n).padStart(2, '0')}T12:00:00Z`);

  it('does not let margin-day messages use up the requested limit', async () => {
    // The service hands back the margin days first
    const pool = [marginBefore(1), marginAfter(1), inWindow(1), inWindow(2), inWindow(3), inWindow(4)];
    const { payload } = await search({ query: 'receipt', limit: 3, ...window }, pagedService(pool));

    expect(idsOf(payload)).toEqual(['in-1', 'in-2', 'in-3']);
    expect(payload.searchSummary.totalResults).toBe(3);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(2);
    // the service returned fewer messages than it was allowed to, so nothing was cut off
    expect(payload.searchSummary.warning).toBeUndefined();
  });

  it.each([
    [1, 26],
    [3, 28],
    [25, 50],
    [100, 200],
    [600, 1000],
    [1000, 1000]
  ])('requests more than limit %i (top %i) when the window is widened', async (limit, top) => {
    const { options } = await search({ query: 'receipt', limit, ...window });
    expect(options.top).toBe(top);
  });

  it('requests exactly the limit when nothing is widened', async () => {
    const noDates = await search({ query: 'receipt', limit: 3 });
    expect(noDates.options.top).toBe(3);

    const odata = await search({ from: 'billing@vendor.example.com', limit: 3, ...window });
    expect(odata.options.top).toBe(3);

    const folder = await search({ query: 'receipt', folders: ['Receipts'], limit: 3, ...window });
    expect(folder.options.top).toBe(3);
  });

  it('warns that results may be incomplete when a full page came back and some of it was dropped', async () => {
    // Only margin-day messages fit in the page; in-window ones may sit behind them
    const pool = [...Array.from({ length: 28 }, (_, n) => marginBefore(n)), inWindow(1), inWindow(2)];
    const { options, payload } = await search({ query: 'receipt', limit: 3, ...window }, pagedService(pool));

    expect(options.top).toBe(28);
    expect(payload.emails).toEqual([]);
    expect(payload.searchSummary.droppedOutsideFilters).toBe(28);
    const { warning } = payload.searchSummary;
    expect(warning).toContain('full page');
    expect(warning).toContain('possibly incomplete');
    expect(warning).toMatch(/narrow the date window or raise limit/);
  });

  it('does not warn about a full page when nothing was dropped', async () => {
    const pool = Array.from({ length: 30 }, (_, n) => message(`in-${n}`, '2026-10-15T12:00:00Z'));
    const { payload } = await search({ query: 'receipt', limit: 3, ...window }, pagedService(pool));

    // 28 fetched (a full page), all inside the window, trimmed to the limit
    expect(payload.searchSummary.droppedOutsideFilters).toBe(0);
    expect(idsOf(payload)).toEqual(['in-0', 'in-1', 'in-2']);
    expect(payload.searchSummary.warning).toBeUndefined();
  });

  it('keeps the unexplained-drops warning and adds the page note when both apply', async () => {
    const pool = [
      ...Array.from({ length: 27 }, (_, n) => marginAfter(n)),
      message('way-out', '2026-01-15T12:00:00Z'),
      inWindow(1)
    ];
    const { payload } = await search({ query: 'receipt', limit: 3, ...window }, pagedService(pool));

    const { warning } = payload.searchSummary;
    expect(warning).toContain('returned 1 message outside the requested filters');
    expect(warning).toContain('full page');
  });
});

describe('searchEmailsTool trims to the limit', () => {
  it('trims a single-endpoint KQL result to the limit even if the service returns more', async () => {
    const value = Array.from({ length: 8 }, (_, n) => message(`m-${n}`, '2026-10-15T12:00:00Z'));
    const { payload } = await search({ query: 'receipt', limit: 3 }, () => ({ value }));
    expect(idsOf(payload)).toEqual(['m-0', 'm-1', 'm-2']);
    expect(payload.searchSummary.totalResults).toBe(3);
  });

  it('trims a single-endpoint $filter result to the limit even if the service returns more', async () => {
    const value = Array.from({ length: 8 }, (_, n) => message(`m-${n}`, '2026-10-15T12:00:00Z'));
    const { payload } = await search({ from: 'billing@vendor.example.com', limit: 3 }, () => ({ value }));
    expect(idsOf(payload)).toEqual(['m-0', 'm-1', 'm-2']);
  });

  it('keeps the first in-window messages when a widened window returns extra margin-day messages', async () => {
    const value = [
      message('a', '2026-10-05T12:00:00Z'),
      message('margin', '2026-11-01T12:00:00Z'),
      message('b', '2026-10-06T12:00:00Z'),
      message('c', '2026-10-07T12:00:00Z'),
      message('d', '2026-10-08T12:00:00Z')
    ];
    const { payload } = await search(
      { query: 'receipt', limit: 3, startDate: '2026-10-01', endDate: '2026-10-31' },
      () => ({ value })
    );
    expect(idsOf(payload)).toEqual(['a', 'b', 'c']);
  });

  it('still trims a multi-folder merge to the limit', async () => {
    const byPath = {
      '/me/mailFolders/Receipts/messages': [message('r1', '2026-10-03T00:00:00Z'), message('r2', '2026-10-20T00:00:00Z')],
      '/me/mailFolders/Invoices/messages': [message('i1', '2026-10-10T00:00:00Z'), message('i2', '2026-10-25T00:00:00Z')]
    };
    const { payload } = await search(
      { query: 'receipt', folders: ['Receipts', 'Invoices'], limit: 3 },
      (path) => ({ value: byPath[path] || [] })
    );
    expect(idsOf(payload)).toEqual(['i2', 'r2', 'i1']);
  });
});
