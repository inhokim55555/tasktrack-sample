function esc(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function listText(value) {
  return (Array.isArray(value) ? value : []).join('\n');
}

function previewList(items, emptyLabel = 'None recorded') {
  const values = (Array.isArray(items) ? items : [])
    .map(value => String(value || '').trim())
    .filter(Boolean);
  return values.length
    ? `<ul>${values.map(value => `<li>${esc(value)}</li>`).join('')}</ul>`
    : `<span>${esc(emptyLabel)}</span>`;
}

export function renderMeetingAnalysisPreview(analysis = {}) {
  const actions = (Array.isArray(analysis.actionItems) ? analysis.actionItems : [])
    .map(action => {
      const text = String(action?.text || '').trim();
      if (!text) return '';
      const details = [action?.suggestedOwnerName, action?.suggestedDueDate]
        .map(value => String(value || '').trim())
        .filter(Boolean);
      return details.length ? `${text} — ${details.join(' · ')}` : text;
    })
    .filter(Boolean);
  return `<section class="meeting-ai-preview">
    <div class="meeting-ai-preview-head">
      <strong>AI meeting notes are ready</strong>
      <span>These notes were generated automatically from the transcript and will remain editable.</span>
    </div>
    <div class="meeting-ai-preview-section meeting-ai-preview-summary">
      <strong>Meeting summary</strong>
      <p>${esc(analysis.summary || 'No summary was generated.')}</p>
    </div>
    <div class="meeting-ai-preview-grid">
      <div class="meeting-ai-preview-section"><strong>Key discussion points</strong>${previewList(analysis.keyDiscussionPoints)}</div>
      <div class="meeting-ai-preview-section"><strong>Decisions made</strong>${previewList(analysis.decisions, 'No explicit decisions recorded')}</div>
      <div class="meeting-ai-preview-section"><strong>Proposed action items</strong>${previewList(actions, 'No action items recorded')}</div>
      <div class="meeting-ai-preview-section"><strong>Open questions and follow-ups</strong>${previewList(analysis.openQuestions)}</div>
    </div>
  </section>`;
}

function meetingMeta(meeting, boardNames) {
  return [
    meeting.meetingDate,
    (meeting.attendees || []).slice(0, 3).join(', '),
    (meeting.relatedBoardIds || []).map(id => boardNames[id]).filter(Boolean).join(', '),
  ].filter(Boolean).join(' · ');
}

function meetingListHtml({ meetings, selectedId, boardNames }) {
  if (!meetings.length) {
    return `<div class="meeting-empty-list">
      <strong>No meetings yet</strong>
      <span>Record, import, or upload a transcript to create the first meeting page.</span>
    </div>`;
  }
  return meetings.map(meeting => `
    <button class="meeting-list-row ${selectedId === meeting.id ? 'active' : ''}" data-open-meeting="${esc(meeting.id)}">
      <span class="meeting-list-icon">M</span>
      <span class="meeting-list-copy">
        <strong>${esc(meeting.title || 'Untitled meeting')}</strong>
        <small>${esc(meetingMeta(meeting, boardNames) || 'Meeting notes')}</small>
      </span>
      <span class="meeting-visibility ${meeting.visibility === 'restricted' ? 'restricted' : ''}">
        ${meeting.visibility === 'restricted' ? 'Restricted' : 'Workspace'}
      </span>
    </button>
  `).join('');
}

export function renderMeetingCaptureBanner(capture, { global = false } = {}) {
  if (!capture?.id) return '';
  const paused = capture.status === 'paused';
  const interrupted = capture.status === 'interrupted';
  const failed = capture.status === 'failed';
  const external = capture.status === 'external';
  const canRecover = interrupted && capture.canRecover === true;
  const canRetryFinalization = failed && capture.canRetryFinalization === true;
  const finalizing = ['uploading', 'transcribing', 'generating_notes', 'finalizing'].includes(capture.status);
  const statusLabel = external
    ? 'Recording active in another tab'
    : failed
    ? (canRetryFinalization ? 'Recording needs attention' : 'Recording incomplete')
    : interrupted
      ? (canRecover ? 'Recording interrupted' : 'Recording incomplete')
      : paused ? 'Recording paused' : finalizing ? 'Finalizing recording' : 'Recording in Friday';
  return `<section class="meeting-capture-banner ${global ? 'global-meeting-capture-banner ' : ''}${interrupted || failed ? 'interrupted' : paused ? 'paused' : finalizing ? 'finalizing' : 'recording'}">
    <div class="meeting-capture-indicator" aria-hidden="true"></div>
    <div class="meeting-capture-copy">
      <div><strong>${esc(statusLabel)}</strong><span>${esc(capture.elapsedLabel || '00:00')}</span></div>
      <p>${esc(capture.liveTranscript || (finalizing ? 'Preparing the final transcript and meeting notes…' : 'Listening for speech…'))}</p>
    </div>
    <div class="meeting-capture-actions">
      ${external ? '<button class="btn sm" data-action="meeting-capture-clear-external">Clear stopped recording</button>' : canRecover ? '<button class="btn sm" data-action="meeting-capture-recover">Resume recording</button>' : interrupted || failed || finalizing ? '' : `<button class="btn ghost sm" data-action="meeting-capture-${paused ? 'resume' : 'pause'}">${paused ? 'Resume' : 'Pause'}</button>`}
      ${external || interrupted || (failed && !canRetryFinalization) ? '' : `<button class="btn sm" data-action="meeting-capture-stop" ${finalizing ? 'disabled' : ''}>${finalizing ? 'Finalizing…' : failed ? 'Retry finalization' : 'Stop and prepare notes'}</button>`}
      ${external || finalizing ? '' : '<button class="btn ghost sm danger" data-action="meeting-capture-cancel">Cancel</button>'}
    </div>
  </section>`;
}

function pendingCapturesHtml(pendingCaptures = []) {
  if (!pendingCaptures.length) return '';
  return `<section class="meeting-imports-ready">
    <div class="meeting-imports-heading"><div><strong>Meeting drafts ready</strong><span>Review access and details before publishing these meeting pages.</span></div><span>${pendingCaptures.length}</span></div>
    <div class="meeting-imports-list">${pendingCaptures.map(capture => `
      <div class="meeting-import-row">
        <span class="meeting-list-icon">${capture.sourceType === 'google_meet_transcript' ? 'G' : 'F'}</span>
        <span><strong>${esc(capture.titleHint || capture.analysis?.title || 'Untitled meeting')}</strong><small>${capture.canRecoverFromAudio ? 'Saved audio needs recovery' : capture.sourceType === 'google_meet_transcript' ? 'Imported from Google Meet' : 'Recorded in Friday'}${capture.createdAt ? ` · ${esc(new Date(capture.createdAt).toLocaleDateString())}` : ''}</small></span>
        <span class="meeting-import-actions">${capture.meetingId
          ? `<button class="btn ghost sm" data-review-meeting-capture="${esc(capture.id)}">Resume save &amp; cleanup</button>`
          : `<button class="btn ghost sm" data-review-meeting-capture="${esc(capture.id)}">${capture.canRecoverFromAudio ? 'Recover & prepare notes' : 'Review'}</button><button class="btn ghost sm danger" data-discard-meeting-capture="${esc(capture.id)}">Discard</button>`}</span>
      </div>`).join('')}</div>
  </section>`;
}

function meetingOverviewHtml(meeting) {
  return `
    <div class="meeting-section-card meeting-summary-card">
      <label for="meeting-summary">Meeting summary</label>
      <textarea id="meeting-summary" data-meeting-section="summary" rows="6">${esc(meeting.sections?.summary || '')}</textarea>
    </div>
    <div class="meeting-section-grid">
      <div class="meeting-section-card">
        <label for="meeting-key-points">Key discussion points</label>
        <span>One point per line</span>
        <textarea id="meeting-key-points" data-meeting-section="keyDiscussionPoints" rows="9">${esc(listText(meeting.sections?.keyDiscussionPoints))}</textarea>
      </div>
      <div class="meeting-section-card">
        <label for="meeting-decisions">Decisions made</label>
        <span>One decision per line</span>
        <textarea id="meeting-decisions" data-meeting-section="decisions" rows="9">${esc(listText(meeting.sections?.decisions))}</textarea>
      </div>
      <div class="meeting-section-card">
        <label for="meeting-open-questions">Open questions and follow-ups</label>
        <span>One item per line</span>
        <textarea id="meeting-open-questions" data-meeting-section="openQuestions" rows="9">${esc(listText(meeting.sections?.openQuestions))}</textarea>
      </div>
    </div>`;
}

function meetingTranscriptHtml(meeting) {
  return `
    <div class="meeting-section-card meeting-transcript-card">
      <div class="meeting-card-heading">
        <div><label for="meeting-transcript">Editable transcript</label><span>This is Friday’s only retained transcript copy.</span></div>
        <span class="meeting-language">${esc(meeting.language || 'Unknown language')}</span>
      </div>
      <textarea id="meeting-transcript" data-meeting-transcript rows="28" spellcheck="true">${esc(meeting.transcript || '')}</textarea>
    </div>`;
}

function meetingActionsHtml(meeting, members) {
  const memberNames = Object.fromEntries((members || []).map(member => [member.id, member.name]));
  const rows = (meeting.actionItems || []).map(action => {
    const owner = action.confirmedOwnerId
      ? memberNames[action.confirmedOwnerId] || action.suggestedOwnerName
      : action.suggestedOwnerName;
    const stateLabel = action.status === 'created' ? 'Task created' : action.status === 'dismissed' ? 'Dismissed' : 'Proposed';
    const editable = action.status === 'proposed';
    return `<div class="meeting-action-row ${esc(action.status || 'proposed')}">
      <input type="checkbox" data-meeting-action-select="${esc(action.id)}" ${editable ? '' : 'disabled'}>
      <div class="meeting-action-main">
        <input class="meeting-action-text" data-meeting-action-field="${esc(action.id)}|text" value="${esc(action.text || '')}" placeholder="Action item" ${editable ? '' : 'disabled'}>
        <div class="meeting-action-suggestions">
          <label>
            <span>${editable ? 'Suggested owner' : 'Confirmed owner'}</span>
            <input data-meeting-action-field="${esc(action.id)}|suggestedOwnerName" value="${esc(owner || '')}" placeholder="Unassigned" ${editable ? '' : 'disabled'}>
          </label>
          <label>
            <span>${editable ? 'Suggested deadline' : 'Confirmed deadline'}</span>
            <input type="date" data-meeting-action-field="${esc(action.id)}|suggestedDueDate" value="${esc(editable ? action.suggestedDueDate || '' : action.confirmedDueDate || action.suggestedDueDate || '')}" ${editable ? '' : 'disabled'}>
          </label>
        </div>
      </div>
      <span class="meeting-action-status">${esc(stateLabel)}</span>
      ${action.linkedTaskId ? `<button class="btn ghost sm" data-open-meeting-task="${esc(action.linkedBoardId)}|${esc(action.linkedTaskId)}">Open task</button>` : ''}
    </div>`;
  }).join('');
  return `
    <div class="meeting-actions-head">
      <div><h3>Proposed action items</h3><p>Review selected items before they become Friday tasks.</p></div>
      <button class="btn" data-action="meeting-create-tasks">Review selected tasks</button>
    </div>
    <div class="meeting-action-list">${rows || '<div class="meeting-empty-list"><strong>No action items found</strong><span>You can add one manually below.</span></div>'}</div>
    <button class="btn ghost sm" data-action="meeting-add-action">Add action item</button>`;
}

function meetingDetailHtml({ meeting, activeTab, members, canPermanentlyDelete }) {
  const tab = ['overview', 'transcript', 'actions'].includes(activeTab) ? activeTab : 'overview';
  return `
    <section class="meeting-detail">
      <div class="meeting-detail-head">
        <div class="meeting-title-wrap">
          <input class="meeting-title-input" data-meeting-field="title" value="${esc(meeting.title || '')}" aria-label="Meeting title">
          <div class="meeting-meta-line">
            <input type="date" data-meeting-field="meetingDate" value="${esc(meeting.meetingDate || '')}" aria-label="Meeting date">
            <input data-meeting-attendees value="${esc((meeting.attendees || []).join(', '))}" placeholder="Attendees, separated by commas" aria-label="Meeting attendees">
            <span>${meeting.visibility === 'restricted' ? 'Restricted meeting' : 'Workspace meeting'}</span>
          </div>
        </div>
        <div class="meeting-detail-actions">
          <button class="btn ghost sm" data-action="meeting-access">Access</button>
          <button class="btn ghost sm" data-action="meeting-copy">Copy notes</button>
          <button class="btn ghost sm" data-action="meeting-pdf">Download PDF</button>
          <button class="btn ghost sm" data-action="meeting-regenerate">Regenerate</button>
          ${meeting.archivedAt
            ? '<button class="btn ghost sm" data-action="meeting-restore">Restore</button>'
            : '<button class="btn ghost sm" data-action="meeting-archive">Archive</button>'}
          ${canPermanentlyDelete ? '<button class="icon-btn meeting-delete" data-action="meeting-delete" title="Permanently delete">×</button>' : ''}
        </div>
      </div>
      <div class="meeting-tabs">
        <button class="${tab === 'overview' ? 'active' : ''}" data-meeting-tab="overview">Overview</button>
        <button class="${tab === 'transcript' ? 'active' : ''}" data-meeting-tab="transcript">Transcript</button>
        <button class="${tab === 'actions' ? 'active' : ''}" data-meeting-tab="actions">
          Action items <span>${(meeting.actionItems || []).filter(item => item.status === 'proposed').length}</span>
        </button>
      </div>
      <div class="meeting-tab-content">
        ${tab === 'overview' ? meetingOverviewHtml(meeting) : ''}
        ${tab === 'transcript' ? meetingTranscriptHtml(meeting) : ''}
        ${tab === 'actions' ? meetingActionsHtml(meeting, members) : ''}
      </div>
    </section>`;
}

export function renderMeetingsView({
  meetings = [],
  selectedId = '',
  detail = null,
  loading = false,
  search = '',
  showArchived = false,
  activeTab = 'overview',
  members = [],
  boards = [],
  canPermanentlyDelete = false,
  error = '',
  capture = null,
  pendingCaptures = [],
  googleMeetConnected = false,
  googleMeetImportEnabled = false,
} = {}) {
  const boardNames = Object.fromEntries(boards.map(board => [board.id, board.name]));
  const query = String(search || '').trim().toLowerCase();
  const filtered = meetings.filter(meeting => {
    if (Boolean(meeting.archivedAt) !== Boolean(showArchived)) return false;
    if (!query) return true;
    const haystack = [
      meeting.title,
      meeting.meetingDate,
      meeting.language,
      ...(meeting.attendees || []),
      ...(meeting.relatedBoardIds || []).map(id => boardNames[id]),
    ].filter(Boolean).join(' ').toLowerCase();
    return haystack.includes(query);
  });
  return `
    <div class="meetings-page">
      <div class="meetings-page-head">
        <div><h2>Meeting notes</h2><p>Record, import, or upload a meeting, then turn confirmed actions into Friday tasks.</p></div>
        <div class="meeting-primary-actions">
          <button class="btn" data-action="meeting-record" ${capture?.id ? 'disabled' : ''}>● Record in Friday</button>
          <button class="btn ghost" data-action="meeting-google-meet">${googleMeetImportEnabled ? 'Google Meet imports on' : googleMeetConnected ? 'Finish Google Meet setup' : 'Connect Google Meet'}</button>
          <button class="btn ghost" data-action="meeting-upload">Upload transcript</button>
        </div>
      </div>
      ${renderMeetingCaptureBanner(capture)}
      ${pendingCapturesHtml(pendingCaptures)}
      <div class="meetings-toolbar">
        <input id="meeting-search" value="${esc(search)}" placeholder="Search meetings, attendees, boards, or dates">
        <label class="meeting-archive-toggle">
          <input type="checkbox" id="meeting-show-archived" ${showArchived ? 'checked' : ''}>
          Archived
        </label>
      </div>
      <div class="meetings-layout">
        <aside class="meeting-list">
          ${error
            ? `<div class="meeting-empty-list meeting-error"><strong>Meetings unavailable</strong><span>${esc(error)}</span></div>`
            : loading && !meetings.length
              ? '<div class="meeting-loading">Loading meetings…</div>'
              : meetingListHtml({ meetings: filtered, selectedId, boardNames })}
        </aside>
        ${loading && selectedId && !detail
          ? '<section class="meeting-detail meeting-loading">Loading meeting…</section>'
          : detail
            ? meetingDetailHtml({ meeting: detail, activeTab, members, canPermanentlyDelete })
            : '<section class="meeting-detail meeting-empty-detail"><div><strong>Select a meeting</strong><span>Its notes and transcript will appear here.</span></div></section>'}
      </div>
    </div>`;
}
