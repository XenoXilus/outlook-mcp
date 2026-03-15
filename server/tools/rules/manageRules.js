import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// List all inbox message rules
export async function listRulesTool(authManager, args) {
  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const result = await graphApiClient.makeRequest('/me/mailFolders/inbox/messageRules', {}, 'GET');
    const rules = result.value || [];

    if (rules.length === 0) {
      return {
        content: [{ type: 'text', text: 'No inbox rules found.' }],
      };
    }

    const summary = rules.map(r => ({
      id: r.id,
      displayName: r.displayName,
      isEnabled: r.isEnabled,
      sequence: r.sequence,
      conditions: r.conditions,
      actions: r.actions,
    }));

    return {
      content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list inbox rules');
  }
}

// Create a new inbox message rule
export async function createRuleTool(authManager, args) {
  const { displayName, senderContains, moveToFolder, isEnabled = true, sequence = 1 } = args;

  if (!displayName) {
    return createValidationError('displayName', 'Parameter is required');
  }
  if (!senderContains || senderContains.length === 0) {
    return createValidationError('senderContains', 'At least one sender filter is required');
  }
  if (!moveToFolder) {
    return createValidationError('moveToFolder', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const ruleData = {
      displayName,
      sequence,
      isEnabled,
      conditions: {
        senderContains,
      },
      actions: {
        moveToFolder,
        stopProcessingRules: true,
      },
    };

    const result = await graphApiClient.postWithRetry(
      '/me/mailFolders/inbox/messageRules',
      ruleData
    );

    return {
      content: [
        {
          type: 'text',
          text: `Rule "${displayName}" created successfully. Rule ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create inbox rule');
  }
}

// Delete an inbox message rule
export async function deleteRuleTool(authManager, args) {
  const { ruleId } = args;

  if (!ruleId) {
    return createValidationError('ruleId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.deleteWithRetry(`/me/mailFolders/inbox/messageRules/${ruleId}`);

    return {
      content: [
        {
          type: 'text',
          text: `Rule ${ruleId} deleted successfully.`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to delete inbox rule');
  }
}
