import { readFileSync } from 'fs';
import path from 'path';
import { describe, expect, it } from 'vitest';
import { buildAllowedScopeDiagnostics } from '../src/auth.js';
import { getCombinedPresetPattern, TOOL_CATEGORIES } from '../src/tool-categories.js';

// Contract tests for the mail capability presets (mail-read, mail-read-draft,
// mail-draft, mail-draft-send, mail-send). Each preset's exact tool list and
// computed token are pinned. The no-send boundaries are token-enforced (no
// Mail.Send in the token); the no-read boundaries are tool surface only -
// Graph has no write-without-read mail scope - which is exactly why the
// exclusion pins below matter. mail-send is the exception that has both: a
// Mail.Send-only token cannot read, draft, modify or delete mailbox content.

interface Endpoint {
  toolName: string;
  presets?: string[];
}

const endpoints = JSON.parse(
  readFileSync(path.join(__dirname, '../src/endpoints.json'), 'utf8')
) as Endpoint[];

const READ_TOOLS = [
  'get-mail-message',
  'get-mail-message-mime',
  'list-mail-attachments',
  'list-mail-child-folders',
  'list-mail-folder-messages',
  'list-mail-folder-messages-delta',
  'list-mail-folders',
  'list-mail-messages',
];

const DRAFT_TOOLS = [
  'add-mail-attachment',
  'create-draft-email',
  'create-mail-attachment-upload-session',
  'delete-mail-attachment',
  'update-mail-message',
];

// Reply/forward drafts reference existing message ids, so they only make sense
// where the preset can also read.
const REPLY_DRAFT_TOOLS = ['create-forward-draft', 'create-reply-all-draft', 'create-reply-draft'];

const SEND_TOOLS = ['send-draft-message', 'send-mail'];

// Mailbox management is deliberately in no capability preset: the token may
// permit it (Mail.ReadWrite), but moving, deleting and rule/folder management
// are not reading, drafting or sending.
const MANAGEMENT_TOOLS = [
  'copy-mail-message',
  'create-mail-folder',
  'create-mail-rule',
  'delete-mail-folder',
  'delete-mail-message',
  'delete-mail-rule',
  'move-mail-message',
  'update-mail-folder',
  'update-mailbox-settings',
];

const CONTRACTS: Record<
  string,
  { tools: string[]; token: string[]; includesDownloaders: boolean }
> = {
  'mail-read': {
    tools: READ_TOOLS,
    token: ['Mail.Read'],
    includesDownloaders: true,
  },
  'mail-read-draft': {
    tools: [...READ_TOOLS, ...DRAFT_TOOLS, ...REPLY_DRAFT_TOOLS],
    token: ['Mail.ReadWrite'],
    includesDownloaders: true,
  },
  'mail-draft': {
    tools: DRAFT_TOOLS,
    token: ['Mail.ReadWrite'],
    includesDownloaders: false,
  },
  'mail-draft-send': {
    tools: [...DRAFT_TOOLS, ...SEND_TOOLS],
    token: ['Mail.ReadWrite', 'Mail.Send'],
    includesDownloaders: false,
  },
  'mail-send': {
    tools: ['send-mail'],
    token: ['Mail.Send'],
    includesDownloaders: false,
  },
};

describe.each(Object.entries(CONTRACTS))('%s preset contract', (preset, contract) => {
  const presetEndpoints = endpoints.filter((e) => e.presets?.includes(preset));

  it('contains exactly the expected endpoints', () => {
    expect(presetEndpoints.map((e) => e.toolName).sort()).toEqual([...contract.tools].sort());
  });

  it('computes exactly the expected token', () => {
    const diagnostics = buildAllowedScopeDiagnostics({
      enabledTools: getCombinedPresetPattern([preset]),
    });
    expect(diagnostics.effectivePermissions).toEqual([...contract.token].sort());
    expect(diagnostics.disabledTools).toEqual([]);
  });

  it('handles byte downloaders per its read boundary', () => {
    const pattern = TOOL_CATEGORIES[preset].pattern;
    for (const tool of ['download-bytes', 'download-bytes-to-file']) {
      if (contract.includesDownloaders) {
        expect(tool).toMatch(pattern);
      } else {
        expect(tool).not.toMatch(pattern);
      }
    }
  });

  it('never exposes management tools or graph-batch', () => {
    const pattern = TOOL_CATEGORIES[preset].pattern;
    for (const tool of [...MANAGEMENT_TOOLS, 'graph-batch']) {
      expect(tool, `${tool} must not be in ${preset}`).not.toMatch(pattern);
    }
  });

  it('does not require org mode', () => {
    expect(TOOL_CATEGORIES[preset].requiresOrgMode).toBeUndefined();
  });
});

describe('mail capability preset boundaries', () => {
  it('no-send presets expose no send tools', () => {
    for (const preset of ['mail-read', 'mail-read-draft', 'mail-draft']) {
      const pattern = TOOL_CATEGORIES[preset].pattern;
      for (const tool of [
        ...SEND_TOOLS,
        'reply-mail-message',
        'reply-all-mail-message',
        'forward-mail-message',
      ]) {
        expect(tool, `${tool} must not be in ${preset}`).not.toMatch(pattern);
      }
    }
  });

  it('no-read presets expose no read tools', () => {
    for (const preset of ['mail-draft', 'mail-draft-send', 'mail-send']) {
      const pattern = TOOL_CATEGORIES[preset].pattern;
      for (const tool of [...READ_TOOLS, ...REPLY_DRAFT_TOOLS]) {
        expect(tool, `${tool} must not be in ${preset}`).not.toMatch(pattern);
      }
    }
  });

  it('mail-send exposes no draft tools', () => {
    const pattern = TOOL_CATEGORIES['mail-send'].pattern;
    for (const tool of [...DRAFT_TOOLS, 'send-draft-message']) {
      expect(tool, `${tool} must not be in mail-send`).not.toMatch(pattern);
    }
  });
});
