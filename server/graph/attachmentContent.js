/**
 * Fetch a message attachment's raw bytes.
 *
 * Graph rejects a `$select` containing the `@odata.type` annotation (400
 * BadRequest), and returns that annotation on attachment items without being
 * asked, so no select is sent for the bytes. If the item endpoint cannot
 * deliver them, the raw `/$value` endpoint is tried before giving up.
 */

import { Buffer } from 'buffer';
import { formatGraphErrorText } from '../utils/mcpErrorResponse.js';

// Metadata-only select. Never add '@odata.type' here: it is not a selectable property.
export const ATTACHMENT_META_SELECT = 'id,name,contentType,size,isInline,lastModifiedDateTime';

const NOT_FILE_TYPES = new Set([
  '#microsoft.graph.itemAttachment',
  '#microsoft.graph.referenceAttachment'
]);

function isMcpError(result) {
  return Boolean(result) && result.isError !== undefined && Boolean(result.content);
}

function failureFromMcpError(result) {
  return {
    text: result.content?.[0]?.text || 'Unknown Graph error',
    graphError: result._errorDetails?.graphError || null
  };
}

// Run one Graph call. makeRequest/getBinary report failures by returning an MCP
// error object, but a client may also throw — normalise both to { ok, value | failure }.
async function attempt(call) {
  try {
    const result = await call();
    if (isMcpError(result)) return { ok: false, failure: failureFromMcpError(result) };
    return { ok: true, value: result };
  } catch (error) {
    if (isMcpError(error)) return { ok: false, failure: failureFromMcpError(error) };
    const graphError = error?.graphError || null;
    return {
      ok: false,
      failure: {
        text: graphError ? formatGraphErrorText(graphError) : (error?.message || String(error)),
        graphError
      }
    };
  }
}

function fetchError(message, details) {
  const error = new Error(message);
  error.fetchDetails = details;
  return error;
}

/**
 * @param {Object} graphApiClient - GraphApiClient (makeRequest + getBinary)
 * @param {string} base - mailbox base, e.g. '/me' or '/users/<mailbox>'
 * @param {string} messageId
 * @param {string} attachmentId - used in the request path exactly as given
 * @param {Object} [hints]
 * @param {string} [hints.attachmentName] - known name, used only in error details
 * @returns {Promise<{ buffer: Buffer, attachment: Object, route: 'json'|'value' }>}
 * @throws {Error} with `fetchDetails` ({ attemptedRoutes, messageId, attachmentId,
 *   attachmentName, graphError }) when the content cannot be fetched
 */
export async function fetchAttachmentContent(graphApiClient, base, messageId, attachmentId, hints = {}) {
  const itemPath = `${base}/messages/${messageId}/attachments/${attachmentId}`;
  let attachment = null;
  const nameOf = () => attachment?.name || hints.attachmentName || null;
  const details = (attemptedRoutes, graphError) => ({
    attemptedRoutes,
    messageId,
    attachmentId,
    attachmentName: nameOf(),
    graphError: graphError || null
  });

  // Route 1: the item itself; fileAttachment carries base64 contentBytes
  const json = await attempt(() => graphApiClient.makeRequest(itemPath));
  let jsonFailure = json.failure;
  if (json.ok) {
    const { contentBytes, ...meta } = json.value || {};
    attachment = meta;
    if (NOT_FILE_TYPES.has(meta['@odata.type'])) {
      throw fetchError(
        `Attachment '${nameOf() || attachmentId}' is not a file attachment (${meta['@odata.type']}); ` +
        'only file attachments have downloadable bytes',
        details(['json'], null)
      );
    }
    if (typeof contentBytes === 'string' && contentBytes.length > 0) {
      return { buffer: Buffer.from(contentBytes, 'base64'), attachment, route: 'json' };
    }
    jsonFailure = { text: 'Graph returned the attachment without contentBytes', graphError: null };
  }

  // Route 2: the raw bytes endpoint
  const valuePath = `${itemPath}/$value`;
  const value = await attempt(() => graphApiClient.getBinary(valuePath));
  let valueFailure = value.failure;
  if (value.ok) {
    const bytes = value.value instanceof ArrayBuffer ? Buffer.from(value.value) : value.value;
    if (bytes && bytes.length > 0) {
      if (!attachment) {
        // Metadata is best-effort here: the bytes are already in hand
        const meta = await attempt(() => graphApiClient.makeRequest(itemPath, { select: ATTACHMENT_META_SELECT }));
        attachment = meta.ok && meta.value ? meta.value : { id: attachmentId };
        if (NOT_FILE_TYPES.has(attachment['@odata.type'])) {
          throw fetchError(
            `Attachment '${nameOf() || attachmentId}' is not a file attachment (${attachment['@odata.type']}); ` +
            'only file attachments have downloadable bytes',
            details(['json', 'value'], null)
          );
        }
      }
      return { buffer: Buffer.from(bytes), attachment, route: 'value' };
    }
    valueFailure = { text: 'Graph returned 0 bytes', graphError: null };
  }

  const label = nameOf() ? `'${nameOf()}' ` : '';
  throw fetchError(
    `Could not fetch the content of attachment ${label}(${attachmentId}) on message ${messageId}. ` +
    `Route json failed: ${jsonFailure.text} | Route value failed: ${valueFailure.text}`,
    details(['json', 'value'], valueFailure.graphError || jsonFailure.graphError)
  );
}
