export const MEETING_TRANSCRIPT_CHUNK_CHARACTERS = 48_000;

export function splitMeetingTranscriptForAnalysis(transcript) {
  const text = String(transcript || '').trim();
  if (!text) return [];
  const chunks = [];
  let offset = 0;
  while (offset < text.length) {
    const remaining = text.length - offset;
    if (remaining <= MEETING_TRANSCRIPT_CHUNK_CHARACTERS) {
      chunks.push(text.slice(offset));
      break;
    }
    const ceiling = offset + MEETING_TRANSCRIPT_CHUNK_CHARACTERS;
    const paragraphBreak = text.lastIndexOf('\n\n', ceiling);
    const lineBreak = text.lastIndexOf('\n', ceiling);
    const breakAt = paragraphBreak > offset
      ? paragraphBreak + 2
      : lineBreak > offset
      ? lineBreak + 1
      : ceiling;
    chunks.push(text.slice(offset, breakAt));
    offset = breakAt;
  }
  return chunks;
}
