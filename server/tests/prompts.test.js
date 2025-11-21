
import { getPrompt } from '../prompts/index.js';

describe('MCP Prompts', () => {
    describe('check_recent_unread', () => {
        it('should return the correct message', async () => {
            const result = await getPrompt('check_recent_unread');
            expect(result.messages).toHaveLength(1);
            expect(result.messages[0].role).toBe('user');
            expect(result.messages[0].content.text).toContain('unread emails received in the last 36 hours');
        });
    });

    describe('draft_reply', () => {
        it('should return the correct message with default tone', async () => {
            const result = await getPrompt('draft_reply', { email_id: '123' });
            expect(result.messages).toHaveLength(1);
            expect(result.messages[0].content.text).toContain('draft a reply to email with ID "123"');
            expect(result.messages[0].content.text).toContain('tone should be professional');
        });

        it('should return the correct message with custom tone', async () => {
            const result = await getPrompt('draft_reply', { email_id: '123', tone: 'casual' });
            expect(result.messages[0].content.text).toContain('tone should be casual');
        });

        it('should throw error if email_id is missing', async () => {
            await expect(getPrompt('draft_reply', {})).rejects.toThrow('Argument "email_id" is required');
        });
    });

    describe('summarize_schedule', () => {
        it('should return the correct message', async () => {
            const result = await getPrompt('summarize_schedule');
            expect(result.messages).toHaveLength(1);
            expect(result.messages[0].content.text).toContain('calendar events for today');
        });
    });

    describe('unknown prompt', () => {
        it('should throw error for unknown prompt', async () => {
            await expect(getPrompt('unknown_prompt')).rejects.toThrow('Prompt not found');
        });
    });
});
