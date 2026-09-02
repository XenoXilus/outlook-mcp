import { describe, it, expect, vi, beforeEach } from 'vitest';
import { getEmailTool } from '../../tools/email/listEmails.js';
import { searchEmailsTool } from '../../tools/email/searchEmails.js';
import { stripHtml, truncateText } from '../../utils/textUtils.js';

describe('Email Tools Redesign', () => {

    describe('Text Utilities', () => {
        it('should strip HTML correctly', () => {
            const html = '<p>Hello <b>World</b></p><br><div>New Line</div>';
            const text = stripHtml(html);
            expect(text).toContain('Hello World');
            expect(text).toContain('New Line');
            expect(text).not.toContain('<p>');
            expect(text).not.toContain('<b>');
        });

        it('should truncate text correctly', () => {
            const text = 'This is a long text that needs truncation';
            const truncated = truncateText(text, 10);
            expect(truncated).toHaveLength(10 + '\n... [truncated, original length: 41 chars]'.length);
            expect(truncated).toContain('This is a ');
            expect(truncated).toContain('... [truncated');
        });
    });

    describe('getEmailTool', () => {
        let mockAuthManager;
        let mockGraphApiClient;

        beforeEach(() => {
            mockGraphApiClient = {
                makeRequest: vi.fn(),
                getFolderResolver: vi.fn()
            };
            mockAuthManager = {
                ensureAuthenticated: vi.fn().mockResolvedValue(true),
                getGraphApiClient: vi.fn().mockReturnValue(mockGraphApiClient)
            };
        });

        it('should return truncated text by default', async () => {
            const longContent = 'a'.repeat(2000);
            const mockEmail = {
                id: '123',
                subject: 'Test Email',
                body: {
                    contentType: 'html',
                    content: `<p>${longContent}</p>`
                }
            };

            mockGraphApiClient.makeRequest.mockResolvedValue(mockEmail);

            const result = await getEmailTool(mockAuthManager, { messageId: '123' });
            const content = JSON.parse(result.content[0].text);

            expect(content.body.contentType).toBe('text');
            expect(content.body.content.length).toBeLessThan(2000);
            expect(content.truncated).toBe(true);
        });

        it('should return full HTML when requested', async () => {
            const htmlContent = '<p>Test Content</p>';
            const mockEmail = {
                id: '123',
                subject: 'Test Email',
                body: {
                    contentType: 'html',
                    content: htmlContent
                }
            };

            mockGraphApiClient.makeRequest.mockResolvedValue(mockEmail);

            const result = await getEmailTool(mockAuthManager, {
                messageId: '123',
                truncate: false,
                format: 'html'
            });
            const content = JSON.parse(result.content[0].text);

            expect(content.body.contentType).toBe('html');
            expect(content.body.content).toBe(htmlContent);
            expect(content.truncated).toBeUndefined();
        });
    });

    describe('searchEmailsTool', () => {
        let mockAuthManager;
        let mockGraphApiClient;

        beforeEach(() => {
            mockGraphApiClient = {
                makeRequest: vi.fn(),
                getFolderResolver: vi.fn()
            };
            mockAuthManager = {
                ensureAuthenticated: vi.fn().mockResolvedValue(true),
                getGraphApiClient: vi.fn().mockReturnValue(mockGraphApiClient)
            };
        });

        it('should not include body by default', async () => {
            const mockResponse = {
                value: [{
                    id: '123',
                    subject: 'Test',
                    bodyPreview: 'Preview'
                }]
            };

            mockGraphApiClient.makeRequest.mockResolvedValue(mockResponse);

            const result = await searchEmailsTool(mockAuthManager, { query: 'test' });
            const content = JSON.parse(result.content[0].text);

            expect(content.emails[0].body).toBeUndefined();
            expect(content.emails[0].bodyPreview).toBe('Preview');
        });

        it('should include truncated body when includeBody is true', async () => {
            const longContent = 'a'.repeat(2000);
            const mockResponse = {
                value: [{
                    id: '123',
                    subject: 'Test',
                    body: {
                        contentType: 'html',
                        content: `<p>${longContent}</p>`
                    }
                }]
            };

            mockGraphApiClient.makeRequest.mockResolvedValue(mockResponse);

            const result = await searchEmailsTool(mockAuthManager, {
                query: 'test',
                includeBody: true
            });
            const content = JSON.parse(result.content[0].text);

            expect(content.emails[0].body).toBeDefined();
            expect(content.emails[0].body.contentType).toBe('text');
            expect(content.emails[0].body.content.length).toBeLessThan(2000);
            expect(content.emails[0].truncated).toBe(true);
        });

        // A date-only endDate used to be sent as `le 2026-08-31`, which Graph
        // reads as midnight — silently dropping the whole last day.
        it('expands a date-only endDate to a half-open bound covering the full day', async () => {
            mockGraphApiClient.makeRequest.mockResolvedValue({ value: [] });

            await searchEmailsTool(mockAuthManager, {
                subject: 'receipt',
                startDate: '2026-08-01',
                endDate: '2026-08-31'
            });

            const options = mockGraphApiClient.makeRequest.mock.calls[0][1];
            expect(options.filter).toContain('receivedDateTime ge 2026-08-01T00:00:00Z');
            expect(options.filter).toContain('receivedDateTime lt 2026-09-01T00:00:00Z');
            expect(options.filter).not.toContain('le 2026-08-31');
        });

        it('handles month-end rollover for date-only endDate in December', async () => {
            mockGraphApiClient.makeRequest.mockResolvedValue({ value: [] });

            await searchEmailsTool(mockAuthManager, {
                subject: 'receipt',
                endDate: '2026-12-31'
            });

            const options = mockGraphApiClient.makeRequest.mock.calls[0][1];
            expect(options.filter).toContain('receivedDateTime lt 2027-01-01T00:00:00Z');
        });

        it('leaves explicit date-time bounds untouched', async () => {
            mockGraphApiClient.makeRequest.mockResolvedValue({ value: [] });

            await searchEmailsTool(mockAuthManager, {
                subject: 'receipt',
                startDate: '2026-08-01T00:00:00Z',
                endDate: '2026-08-31T23:59:59Z'
            });

            const options = mockGraphApiClient.makeRequest.mock.calls[0][1];
            expect(options.filter).toContain('receivedDateTime ge 2026-08-01T00:00:00Z');
            expect(options.filter).toContain('receivedDateTime le 2026-08-31T23:59:59Z');
        });
    });
});
