
/**
 * MCP Prompts for Outlook
 * 
 * Defines standard prompts that users can use to quickly perform common tasks.
 */

export const PROMPTS = {
    'check_recent_unread': {
        name: 'check_recent_unread',
        description: 'Check unread emails from the last 36 hours, summarize them, and highlight action items',
        arguments: []
    },
    'draft_reply': {
        name: 'draft_reply',
        description: 'Draft a reply to a specific email with a specified tone',
        arguments: [
            {
                name: 'email_id',
                description: 'The ID of the email to reply to',
                required: true
            },
            {
                name: 'tone',
                description: 'The tone of the response (e.g., professional, casual, urgent)',
                required: false
            }
        ]
    },
    'summarize_schedule': {
        name: 'summarize_schedule',
        description: 'Summarize today\'s calendar events and schedule',
        arguments: []
    }
};

/**
 * Handler for getting a specific prompt
 * @param {string} name - The name of the prompt
 * @param {Object} args - Arguments for the prompt
 * @returns {Object} The prompt result containing messages
 */
export async function getPrompt(name, args = {}) {
    switch (name) {
        case 'check_recent_unread':
            return {
                messages: [
                    {
                        role: 'user',
                        content: {
                            type: 'text',
                            text: 'Please check for unread emails received in the last 36 hours. Provide a high-level summary of these emails and highlight any that require immediate action.'
                        }
                    }
                ]
            };

        case 'draft_reply':
            const { email_id, tone = 'professional' } = args;
            if (!email_id) {
                throw new Error('Argument "email_id" is required for prompt "draft_reply"');
            }
            return {
                messages: [
                    {
                        role: 'user',
                        content: {
                            type: 'text',
                            text: `Please draft a reply to email with ID "${email_id}". The tone should be ${tone}. First, fetch the email content to understand the context, then draft the response.`
                        }
                    }
                ]
            };

        case 'summarize_schedule':
            return {
                messages: [
                    {
                        role: 'user',
                        content: {
                            type: 'text',
                            text: 'Please list my calendar events for today and provide a summary of my schedule. Identify any conflicts or free blocks.'
                        }
                    }
                ]
            };

        default:
            throw new Error(`Prompt not found: ${name}`);
    }
}

export const promptList = Object.values(PROMPTS);
