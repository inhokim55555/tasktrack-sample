export const MEETING_SECTION_KEYS = [
  'summary',
  'keyDiscussionPoints',
  'decisions',
  'openQuestions',
];

const MEETING_SOURCE_TYPES = new Set([
  'upload',
  'friday_mic',
  'browser_tab',
  'google_meet_transcript',
  'google_meet_media',
]);

function makeId(prefix) {
  const value = globalThis.crypto?.randomUUID?.()
    || `${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
  return `${prefix}_${value}`;
}

function cleanText(value) {
  return String(value || '').trim();
}

function cleanTextList(value) {
  const rows = Array.isArray(value) ? value : [];
  return rows.map(cleanText).filter(Boolean);
}

function cleanIsoDate(value) {
  const text = cleanText(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return '';
  const date = new Date(`${text}T00:00:00Z`);
  return Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== text ? '' : text;
}

function normalizeActionItem(item = {}) {
  return {
    id: cleanText(item.id) || makeId('meeting_action'),
    text: cleanText(item.text),
    suggestedOwnerName: cleanText(item.suggestedOwnerName),
    suggestedDueDate: cleanIsoDate(item.suggestedDueDate),
    confirmedOwnerId: '',
    confirmedDueDate: '',
    linkedTaskId: '',
    linkedBoardId: '',
    status: 'proposed',
  };
}

export function createMeetingDraft({
  id = '',
  workspaceId = '',
  uploaderId = '',
  sourceType = 'upload',
  externalSourceId = '',
  captureId = '',
  transcript = '',
  analysis = {},
  now = new Date().toISOString(),
} = {}) {
  return {
    id: cleanText(id) || makeId('meeting'),
    workspaceId: cleanText(workspaceId),
    uploaderId: cleanText(uploaderId),
    sourceType: MEETING_SOURCE_TYPES.has(cleanText(sourceType)) ? cleanText(sourceType) : 'upload',
    externalSourceId: cleanText(externalSourceId),
    captureId: cleanText(captureId),
    title: cleanText(analysis.title) || 'Untitled meeting',
    meetingDate: cleanIsoDate(analysis.meetingDate),
    attendees: cleanTextList(analysis.attendees),
    language: cleanText(analysis.language) || 'Unknown',
    visibility: 'workspace',
    allowedMemberIds: [],
    relatedBoardIds: [],
    status: 'ready',
    archivedAt: '',
    transcript: cleanText(transcript),
    sections: {
      summary: cleanText(analysis.summary),
      keyDiscussionPoints: cleanTextList(analysis.keyDiscussionPoints),
      decisions: cleanTextList(analysis.decisions),
      openQuestions: cleanTextList(analysis.openQuestions),
    },
    actionItems: (Array.isArray(analysis.actionItems) ? analysis.actionItems : [])
      .map(normalizeActionItem)
      .filter(item => item.text),
    createdAt: now,
    updatedAt: now,
  };
}

export function validateGeneratedMeetingAnalysis(analysis = {}) {
  const missing = [];
  if (!cleanText(analysis.transcript)) missing.push('transcript');
  if (!cleanText(analysis.summary)) missing.push('meeting summary');
  if (!cleanTextList(analysis.keyDiscussionPoints).length) missing.push('key discussion points');
  for (const [field, label] of [
    ['decisions', 'decisions'],
    ['actionItems', 'action items'],
    ['openQuestions', 'open questions'],
  ]) {
    if (!Array.isArray(analysis[field])) missing.push(label);
  }
  return {
    valid: missing.length === 0,
    missing,
  };
}

export function canViewMeeting(meeting, { memberId = '', isAdmin = false } = {}) {
  if (!meeting) return false;
  if (meeting.visibility !== 'restricted') return true;
  if (isAdmin) return true;
  const currentMemberId = cleanText(memberId);
  if (!currentMemberId) return false;
  if (currentMemberId === cleanText(meeting.uploaderId)) return true;
  return cleanTextList(meeting.allowedMemberIds).includes(currentMemberId);
}

export function canEditMeeting(meeting, actor = {}) {
  return canViewMeeting(meeting, actor);
}

export function createRegenerationPreview(analysis = {}) {
  return {
    sections: {
      summary: cleanText(analysis.summary),
      keyDiscussionPoints: cleanTextList(analysis.keyDiscussionPoints),
      decisions: cleanTextList(analysis.decisions),
      openQuestions: cleanTextList(analysis.openQuestions),
    },
    actionItems: (Array.isArray(analysis.actionItems) ? analysis.actionItems : [])
      .map(normalizeActionItem)
      .filter(item => item.text),
  };
}

export function applyRegenerationPreview(
  meeting,
  preview,
  acceptedSectionKeys = [],
  now = new Date().toISOString(),
) {
  const accepted = new Set(acceptedSectionKeys);
  const nextSections = { ...(meeting?.sections || {}) };
  for (const key of MEETING_SECTION_KEYS) {
    if (accepted.has(key)) nextSections[key] = preview?.sections?.[key];
  }
  return {
    ...meeting,
    sections: nextSections,
    actionItems: accepted.has('actionItems')
      ? (preview?.actionItems || []).map(item => ({ ...item }))
      : (meeting?.actionItems || []).map(item => ({ ...item })),
    updatedAt: now,
  };
}

export function suggestMeetingActionOwner(suggestedName, members = []) {
  const target = cleanText(suggestedName).toLocaleLowerCase();
  if (!target) return null;
  const matches = members.filter(member =>
    cleanText(member?.name).toLocaleLowerCase() === target
  );
  if (matches.length !== 1) return null;
  return { id: cleanText(matches[0].id), name: cleanText(matches[0].name) };
}

export function buildTaskDraftFromMeetingAction(action = {}, { meetingTitle = '' } = {}) {
  const suggestedOwnerName = cleanText(action.suggestedOwnerName);
  const ownerIds = cleanText(action.confirmedOwnerId) ? [cleanText(action.confirmedOwnerId)] : [];
  const notes = [
    cleanText(meetingTitle) ? `Created from meeting: ${cleanText(meetingTitle)}` : '',
    !ownerIds.length && suggestedOwnerName
      ? `Suggested owner from meeting: ${suggestedOwnerName}`
      : '',
  ].filter(Boolean).join('\n');
  return {
    name: cleanText(action.text) || 'Meeting action item',
    ownerIds,
    due: cleanIsoDate(action.confirmedDueDate) || cleanIsoDate(action.suggestedDueDate),
    notes,
  };
}

export function canonicalMeetingTaskResult(result = {}) {
  const normalized = {
    actionId:cleanText(result.actionId || result.action_id),
    taskId:cleanText(result.taskId || result.task_id),
    boardId:cleanText(result.boardId || result.board_id),
    groupId:cleanText(result.groupId || result.group_id),
    created:result.created === true,
  };
  if(!normalized.actionId || !normalized.taskId || !normalized.boardId || !normalized.groupId) return null;
  return {
    ...normalized,
    shouldEmitCreationSideEffects:normalized.created,
  };
}

function noteList(title, items = []) {
  const values = cleanTextList(items);
  return `${title}\n${values.length ? values.map(item => `• ${item}`).join('\n') : 'None recorded'}`;
}

export function formatMeetingNotes(meeting = {}) {
  const actionLines = (meeting.actionItems || []).map(action => {
    const details = [
      cleanText(action.suggestedOwnerName),
      cleanText(action.suggestedDueDate),
    ].filter(Boolean);
    return `${cleanText(action.text)}${details.length ? ` — ${details.join(' · ')}` : ''}`;
  }).filter(Boolean);
  return [
    cleanText(meeting.title) || 'Untitled meeting',
    cleanText(meeting.meetingDate) ? `Date: ${cleanText(meeting.meetingDate)}` : '',
    cleanTextList(meeting.attendees).length
      ? `Attendees: ${cleanTextList(meeting.attendees).join(', ')}`
      : '',
    '',
    `Meeting summary\n${cleanText(meeting.sections?.summary) || 'None recorded'}`,
    '',
    noteList('Key discussion points', meeting.sections?.keyDiscussionPoints),
    '',
    noteList('Decisions made', meeting.sections?.decisions),
    '',
    noteList('Action items', actionLines),
    '',
    noteList('Open questions and follow-ups', meeting.sections?.openQuestions),
  ].filter((line, index, rows) => line !== '' || rows[index - 1] !== '').join('\n').trim();
}
