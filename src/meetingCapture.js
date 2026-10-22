export const MEETING_CAPTURE_SOURCE_TYPES = Object.freeze({
  UPLOAD: 'upload',
  FRIDAY_MIC: 'friday_mic',
  BROWSER_TAB: 'browser_tab',
  GOOGLE_MEET_TRANSCRIPT: 'google_meet_transcript',
  GOOGLE_MEET_MEDIA: 'google_meet_media',
});

export const MEETING_CAPTURE_STATUSES = Object.freeze([
  'pending_consent',
  'recording',
  'paused',
  'uploading',
  'transcribing',
  'generating_notes',
  'finalizing',
  'awaiting_confirmation',
  'ready',
  'completed',
  'failed',
  'cancelled',
]);

const STATUS_TRANSITIONS = Object.freeze({
  pending_consent: new Set(['recording', 'cancelled', 'failed']),
  recording: new Set(['paused', 'uploading', 'transcribing', 'finalizing', 'cancelled', 'failed']),
  paused: new Set(['recording', 'uploading', 'transcribing', 'finalizing', 'cancelled', 'failed']),
  uploading: new Set(['transcribing', 'finalizing', 'cancelled', 'failed']),
  transcribing: new Set(['generating_notes', 'finalizing', 'cancelled', 'failed']),
  generating_notes: new Set(['awaiting_confirmation', 'failed']),
  finalizing: new Set(['transcribing', 'generating_notes', 'awaiting_confirmation', 'failed']),
  awaiting_confirmation: new Set(['ready', 'completed', 'cancelled', 'failed']),
  ready: new Set(['completed']),
  completed: new Set(),
  failed: new Set(['transcribing', 'finalizing', 'cancelled']),
  cancelled: new Set(),
});

function cleanText(value) {
  return String(value ?? '').trim();
}

function finiteNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

export function captureStatusCanTransition(currentStatus, nextStatus) {
  const current = cleanText(currentStatus);
  const next = cleanText(nextStatus);
  if (!current || !next) return false;
  if (current === next) return true;
  return Boolean(STATUS_TRANSITIONS[current]?.has(next));
}

export function normalizeCaptureSegment(segment = {}, fallbackSequence = 0) {
  const sequence = Math.max(0, Math.trunc(finiteNumber(segment.sequence, fallbackSequence)));
  const startMs = Math.max(0, Math.trunc(finiteNumber(segment.startMs ?? segment.start_ms, 0)));
  const endMs = Math.max(startMs, Math.trunc(finiteNumber(segment.endMs ?? segment.end_ms, startMs)));
  const text = cleanText(segment.text ?? segment.transcript);
  return {
    id: cleanText(segment.id ?? segment.itemId ?? segment.item_id) || `segment-${sequence}`,
    sequence,
    startMs,
    endMs,
    speakerKey: cleanText(segment.speakerKey ?? segment.speaker_key ?? segment.speaker),
    speakerName: cleanText(segment.speakerName ?? segment.speaker_name),
    language: cleanText(segment.language ?? segment.language_code),
    text,
    isFinal: segment.isFinal !== false && segment.is_final !== false,
  };
}

export function appendCaptureSegment(segments = [], segment = {}) {
  const normalized = normalizeCaptureSegment(segment, segments.length);
  if (!normalized.text) return [...segments];
  const next = (Array.isArray(segments) ? segments : []).map((item, index) =>
    normalizeCaptureSegment(item, index)
  );
  const existingIndex = next.findIndex(item => item.id === normalized.id);
  if (existingIndex >= 0) next[existingIndex] = normalized;
  else next.push(normalized);
  return next.sort((a, b) =>
    a.startMs - b.startMs || a.sequence - b.sequence || a.id.localeCompare(b.id)
  );
}

export function canonicalCaptureTranscript(segments = []) {
  const ordered = (Array.isArray(segments) ? segments : [])
    .map((segment, index) => normalizeCaptureSegment(segment, index))
    .filter(segment => segment.text && segment.isFinal)
    .sort((a, b) => a.startMs - b.startMs || a.sequence - b.sequence || a.id.localeCompare(b.id));
  return ordered.map(segment => {
    const speaker = segment.speakerName || segment.speakerKey;
    return speaker ? `${speaker}: ${segment.text}` : segment.text;
  }).join('\n').trim();
}

export function realtimeTranscriptEventToSegment(event = {}, fallbackSequence = 0) {
  const type = cleanText(event.type);
  const completed = type.endsWith('input_audio_transcription.completed')
    || type === 'transcript.text.segment'
    || type.endsWith('transcription.completed');
  if (!completed) return null;
  const text = cleanText(event.transcript ?? event.text ?? event.segment?.text);
  if (!text) return null;
  return normalizeCaptureSegment({
    id: event.item_id ?? event.id ?? event.segment?.id,
    sequence: event.sequence ?? fallbackSequence,
    startMs: event.start_ms ?? event.startMs ?? event.segment?.start_ms
      ?? (Number(event.segment?.start) * 1000),
    endMs: event.end_ms ?? event.endMs ?? event.segment?.end_ms
      ?? (Number(event.segment?.end) * 1000),
    speakerKey: event.speaker ?? event.segment?.speaker,
    speakerName: event.speaker_name ?? event.segment?.speaker_name,
    language: event.language ?? event.language_code,
    text,
    isFinal: true,
  }, fallbackSequence);
}

export function normalizeCaptureDraft(row = {}) {
  const metadata = row.source_metadata && typeof row.source_metadata === 'object'
    ? row.source_metadata
    : {};
  const analysis = metadata.analysis && typeof metadata.analysis === 'object'
    ? metadata.analysis
    : {};
  return {
    id: cleanText(row.id),
    sourceType: cleanText(row.source_type),
    status: cleanText(row.status),
    meetingId: cleanText(row.meeting_id),
    titleHint: cleanText(row.title_hint),
    meetingDateHint: cleanText(row.meeting_date_hint),
    languageHint: cleanText(row.language_hint),
    externalSourceId: cleanText(row.external_source_id),
    transcript: cleanText(metadata.transcript),
    analysis: { ...analysis },
    createdAt: cleanText(row.created_at),
    updatedAt: cleanText(row.updated_at),
    heartbeatAt: cleanText(row.heartbeat_at),
    clientSessionId: cleanText(metadata.clientSessionId),
    audioObjectPaths: Array.isArray(metadata.audioObjectPaths) ? [...metadata.audioObjectPaths] : [],
    error: cleanText(row.error),
  };
}

export function captureElapsedLabel(startedAt, now = Date.now()) {
  const start = new Date(startedAt).getTime();
  const end = Number(now);
  if (!Number.isFinite(start) || !Number.isFinite(end)) return '00:00';
  const totalSeconds = Math.max(0, Math.floor((end - start) / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return [hours ? String(hours).padStart(2, '0') : '', String(minutes).padStart(2, '0'), String(seconds).padStart(2, '0')]
    .filter(Boolean)
    .join(':');
}
