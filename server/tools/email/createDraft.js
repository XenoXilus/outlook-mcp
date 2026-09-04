import fs from 'fs';
import path from 'path';
import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

const MAX_DIRECT_ATTACHMENT = 3 * 1024 * 1024; // Graph's base64 direct-attach ceiling
const MAX_ATTACHMENT = 150 * 1024 * 1024; // Graph's per-attachment ceiling
// Upload-session chunks must be a multiple of 320 KiB (Graph requirement). 10 × 320 KiB.
const UPLOAD_CHUNK_SIZE = 10 * 320 * 1024; // 3,276,800 bytes

/**
 * Stream one large file to a draft via a Microsoft Graph upload session.
 *
 * The file is read chunk-by-chunk (never buffered whole) and PUT to the
 * pre-authenticated uploadUrl with plain fetch — no Authorization header.
 * Returns null on success, or an MCP tool-error response on failure.
 *
 * `mailboxBase` is required (from getMailboxBase) — no default, so an omitted
 * argument fails loudly rather than silently targeting the signed-in user.
 */
async function uploadLargeAttachment(graphApiClient, draftId, info, fetchImpl, mailboxBase) {
  const session = await graphApiClient.postWithRetry(
    `${mailboxBase}/messages/${draftId}/attachments/createUploadSession`,
    {
      AttachmentItem: {
        attachmentType: 'file',
        name: info.name,
        size: info.size,
      },
    }
  );
  if (session && session.isError !== undefined && session.content) return session;

  const uploadUrl = session && session.uploadUrl;
  if (!uploadUrl) {
    return createValidationError(
      'attachmentPaths',
      `Upload session for ${info.name} did not return an uploadUrl; the draft was created but this attachment could not be uploaded.`
    );
  }

  const total = info.size;
  const handle = await fs.promises.open(info.filePath, 'r');
  try {
    const buffer = Buffer.allocUnsafe(Math.min(UPLOAD_CHUNK_SIZE, total || UPLOAD_CHUNK_SIZE));
    let start = 0;
    while (start < total) {
      const length = Math.min(UPLOAD_CHUNK_SIZE, total - start);
      const { bytesRead } = await handle.read(buffer, 0, length, start);
      if (bytesRead <= 0) {
        return createValidationError(
          'attachmentPaths',
          `Unexpected end of file reading ${info.name} at byte ${start}; the draft was created but this attachment is incomplete.`
        );
      }
      const end = start + bytesRead - 1;
      const chunk = bytesRead === buffer.length ? buffer : buffer.subarray(0, bytesRead);

      let response;
      try {
        response = await fetchImpl(uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Length': String(bytesRead),
            'Content-Range': `bytes ${start}-${end}/${total}`,
          },
          body: chunk,
        });
      } catch (netError) {
        return convertErrorToToolError(
          new Error(
            `Upload of ${info.name} failed while sending bytes ${start}-${end} of ${total}: ${netError.message}. The draft was created — inspect it in Outlook.`
          ),
          'Failed to create draft'
        );
      }

      // Graph upload sessions return 200 (intermediate) / 201 (final) for sequential uploads;
      // ranges are never resent, so nextExpectedRanges is intentionally not parsed.
      const ok = response && (response.ok === true || (typeof response.status === 'number' && response.status >= 200 && response.status < 300));
      if (!ok) {
        const status = response && response.status !== undefined ? response.status : 'unknown';
        return createValidationError(
          'attachmentPaths',
          `Upload of ${info.name} failed at bytes ${start}-${end} of ${total} (HTTP ${status}). The draft was created but this attachment is incomplete; inspect the draft in Outlook.`
        );
      }

      start = end + 1;
    }
  } finally {
    await handle.close();
  }
  return null;
}

// Create draft email with user styling
export async function createDraftTool(authManager, args, deps = {}) {
  const { fetchImpl = fetch } = deps;
  const {
    to, subject, body, bodyHtml, bodyType = 'text', cc = [], bcc = [],
    importance = 'normal', preserveUserStyling = true, attachmentPaths = [], mailbox,
  } = args;

  if (!to || to.length === 0) {
    return createValidationError('to', 'At least one recipient is required');
  }

  if (!subject) {
    return createValidationError('subject', 'Subject is required');
  }

  if (bodyHtml !== undefined && body !== undefined) {
    return createValidationError('bodyHtml', 'Provide either bodyHtml or body, not both (ambiguous body).');
  }

  // Resolve the body: bodyHtml is an alias for body + bodyType: 'html'.
  let finalBody;
  let finalBodyType;
  if (bodyHtml !== undefined) {
    finalBody = bodyHtml || '';
    finalBodyType = 'html';
  } else {
    finalBody = body || '';
    finalBodyType = bodyType;
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    // Validate every attachment (existence + size) BEFORE creating the draft or
    // making any attachment call, so an unreadable/oversized file aborts cleanly.
    const fileInfos = [];
    for (const filePath of attachmentPaths) {
      let stat;
      try {
        stat = await fs.promises.stat(filePath);
      } catch (readError) {
        return createValidationError('attachmentPaths', `Cannot read ${filePath}: ${readError.message}`);
      }
      if (!stat.isFile()) {
        return createValidationError('attachmentPaths', `${filePath} is not a regular file`);
      }
      if (stat.size > MAX_ATTACHMENT) {
        return createValidationError(
          'attachmentPaths',
          `${filePath} is ${stat.size} bytes — exceeds the 150 MB per-attachment limit`
        );
      }
      fileInfos.push({ filePath, name: path.basename(filePath), size: stat.size });
    }

    // Apply user styling if enabled
    if (preserveUserStyling && finalBody) {
      const styledBody = await applyUserStyling(graphApiClient, finalBody, finalBodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const draft = {
      subject,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
      importance,
    };

    if (cc.length > 0) {
      draft.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      draft.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages`, draft);
    if (result && result.isError !== undefined && result.content) return result;

    const attached = [];
    for (const info of fileInfos) {
      if (info.size <= MAX_DIRECT_ATTACHMENT) {
        // Small file: base64 direct-attach.
        const buffer = await fs.promises.readFile(info.filePath);
        const attachResult = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${result.id}/attachments`, {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: info.name,
          contentBytes: buffer.toString('base64'),
        });
        if (attachResult && attachResult.isError !== undefined && attachResult.content) return attachResult;
      } else {
        // Large file: chunked upload session (never send).
        const uploadError = await uploadLargeAttachment(graphApiClient, result.id, info, fetchImpl, mailboxBase);
        if (uploadError) return uploadError;
      }
      attached.push({ name: info.name, size: info.size });
    }

    const draftInfo = await graphApiClient.makeRequest(`${mailboxBase}/messages/${result.id}`, { select: 'id,webLink' });

    return createSafeResponse({
      draftId: result.id,
      webLink: draftInfo?.webLink || null,
      attachments: attached,
      note: 'Draft created for review — this tool never sends.',
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create draft');
  }
}
