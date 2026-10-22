import { MEETING_SECTION_KEYS } from './meetingDomain.js';

function requireValue(value, label) {
  const clean = String(value || '').trim();
  if (!clean) throw new Error(`${label} is required.`);
  return clean;
}

function meetingRow(meeting, workspaceId, userId) {
  return {
    workspace_id: workspaceId,
    id: requireValue(meeting.id, 'Meeting id'),
    title: String(meeting.title || '').trim() || 'Untitled meeting',
    meeting_date: meeting.meetingDate || null,
    attendees: Array.isArray(meeting.attendees) ? meeting.attendees : [],
    language: String(meeting.language || 'Unknown'),
    visibility: meeting.visibility === 'restricted' ? 'restricted' : 'workspace',
    allowed_member_ids: Array.isArray(meeting.allowedMemberIds) ? meeting.allowedMemberIds : [],
    related_board_ids: Array.isArray(meeting.relatedBoardIds) ? meeting.relatedBoardIds : [],
    status: ['processing', 'ready', 'failed'].includes(meeting.status) ? meeting.status : 'ready',
    source_type: String(meeting.sourceType || 'upload'),
    external_source_id: String(meeting.externalSourceId || ''),
    capture_id: String(meeting.captureId || ''),
    archived_at: meeting.archivedAt || null,
    uploader_member_id: String(meeting.uploaderId || ''),
    uploader_user_id: requireValue(meeting.uploaderUserId || userId, 'Uploader user'),
    created_at: meeting.createdAt || new Date().toISOString(),
    updated_at: meeting.updatedAt || new Date().toISOString(),
  };
}

function rowToMeeting(row) {
  return {
    id: row.id,
    workspaceId: row.workspace_id,
    uploaderId: row.uploader_member_id || '',
    uploaderUserId: row.uploader_user_id || '',
    title: row.title || 'Untitled meeting',
    meetingDate: row.meeting_date || '',
    attendees: Array.isArray(row.attendees) ? row.attendees : [],
    language: row.language || 'Unknown',
    visibility: row.visibility === 'restricted' ? 'restricted' : 'workspace',
    allowedMemberIds: Array.isArray(row.allowed_member_ids) ? row.allowed_member_ids : [],
    relatedBoardIds: Array.isArray(row.related_board_ids) ? row.related_board_ids : [],
    status: row.status || 'ready',
    sourceType: row.source_type || 'upload',
    externalSourceId: row.external_source_id || '',
    captureId: row.capture_id || '',
    archivedAt: row.archived_at || '',
    contentSavedAt: row.content_saved_at || '',
    createdAt: row.created_at || '',
    updatedAt: row.updated_at || '',
  };
}

function actionRow(action, meeting, workspaceId, position) {
  return {
    workspace_id: workspaceId,
    meeting_id: meeting.id,
    id: requireValue(action.id, 'Action item id'),
    text: String(action.text || ''),
    suggested_owner_name: String(action.suggestedOwnerName || ''),
    suggested_due_date: action.suggestedDueDate || null,
    confirmed_owner_id: String(action.confirmedOwnerId || ''),
    confirmed_due_date: action.confirmedDueDate || null,
    linked_task_id: String(action.linkedTaskId || ''),
    linked_board_id: String(action.linkedBoardId || ''),
    linked_group_id: String(action.linkedGroupId || ''),
    status: ['proposed', 'created', 'dismissed'].includes(action.status) ? action.status : 'proposed',
    position,
    updated_at: meeting.updatedAt || new Date().toISOString(),
  };
}

export function createMeetingRepository({ supabase, workspaceId, userId }) {
  if (!supabase) throw new Error('Supabase is required for Meetings.');
  const currentWorkspaceId = requireValue(workspaceId, 'Workspace');
  const currentUserId = requireValue(userId, 'Signed-in user');

  async function list({ includeArchived = false } = {}) {
    let query = supabase
      .from('app_meetings')
      .select('*')
      .eq('workspace_id', currentWorkspaceId)
      .order('meeting_date', { ascending: false, nullsFirst: false })
      .order('updated_at', { ascending: false });
    if (!includeArchived) query = query.is('archived_at', null);
    const { data, error } = await query;
    if (error) throw error;
    return (data || []).map(rowToMeeting);
  }

  async function load(id) {
    const meetingId = requireValue(id, 'Meeting');
    const [meetingResult, transcriptResult, sectionsResult, actionsResult, linksResult] = await Promise.all([
      supabase.from('app_meetings').select('*')
        .eq('workspace_id', currentWorkspaceId).eq('id', meetingId).single(),
      supabase.from('app_meeting_transcripts').select('transcript')
        .eq('workspace_id', currentWorkspaceId).eq('meeting_id', meetingId).maybeSingle(),
      supabase.from('app_meeting_sections').select('section_key,content')
        .eq('workspace_id', currentWorkspaceId).eq('meeting_id', meetingId),
      supabase.from('app_meeting_action_items').select('*')
        .eq('workspace_id', currentWorkspaceId).eq('meeting_id', meetingId).order('position'),
      supabase.from('app_meeting_links').select('entity_type,entity_id')
        .eq('workspace_id', currentWorkspaceId).eq('meeting_id', meetingId),
    ]);
    const firstError = [
      meetingResult.error,
      transcriptResult.error,
      sectionsResult.error,
      actionsResult.error,
      linksResult.error,
    ].find(Boolean);
    if (firstError) throw firstError;

    const sections = Object.fromEntries(MEETING_SECTION_KEYS.map(key => [key, key === 'summary' ? '' : []]));
    for (const row of sectionsResult.data || []) sections[row.section_key] = row.content;
    const meeting = rowToMeeting(meetingResult.data);
    meeting.transcript = transcriptResult.data?.transcript || '';
    meeting.sections = sections;
    meeting.actionItems = (actionsResult.data || []).map(row => ({
      id: row.id,
      text: row.text || '',
      suggestedOwnerName: row.suggested_owner_name || '',
      suggestedDueDate: row.suggested_due_date || '',
      confirmedOwnerId: row.confirmed_owner_id || '',
      confirmedDueDate: row.confirmed_due_date || '',
      linkedTaskId: row.linked_task_id || '',
      linkedBoardId: row.linked_board_id || '',
      linkedGroupId: row.linked_group_id || '',
      status: row.status || 'proposed',
    }));
    meeting.links = (linksResult.data || []).map(row => ({
      entityType: row.entity_type,
      entityId: row.entity_id,
    }));
    return meeting;
  }

  async function save(meeting) {
    const next = {
      ...meeting,
      updatedAt: new Date().toISOString(),
      uploaderUserId: meeting.uploaderUserId || currentUserId,
    };
    const initialMeetingRow = {
      ...meetingRow(next, currentWorkspaceId, currentUserId),
      content_saved_at: null,
    };
    const { error: meetingError } = await supabase
      .from('app_meetings')
      .upsert(initialMeetingRow, { onConflict: 'workspace_id,id' });
    if (meetingError) throw meetingError;

    const transcriptRow = {
      workspace_id: currentWorkspaceId,
      meeting_id: next.id,
      transcript: String(next.transcript || ''),
      updated_at: next.updatedAt,
    };
    const { error: transcriptError } = await supabase
      .from('app_meeting_transcripts')
      .upsert(transcriptRow, { onConflict: 'workspace_id,meeting_id' });
    if (transcriptError) throw transcriptError;

    const sectionRows = MEETING_SECTION_KEYS.map(sectionKey => ({
      workspace_id: currentWorkspaceId,
      meeting_id: next.id,
      section_key: sectionKey,
      content: next.sections?.[sectionKey] ?? (sectionKey === 'summary' ? '' : []),
      updated_at: next.updatedAt,
    }));
    const { error: sectionError } = await supabase
      .from('app_meeting_sections')
      .upsert(sectionRows, { onConflict: 'workspace_id,meeting_id,section_key' });
    if (sectionError) throw sectionError;

    const actionRows = (next.actionItems || []).map((action, position) =>
      actionRow(action, next, currentWorkspaceId, position)
    );
    if (actionRows.length) {
      const { error: actionError } = await supabase
        .from('app_meeting_action_items')
        .upsert(actionRows, { onConflict: 'workspace_id,meeting_id,id' });
      if (actionError) throw actionError;
    }
    const { data: existingActions, error: existingActionsError } = await supabase
      .from('app_meeting_action_items')
      .select('id')
      .eq('workspace_id', currentWorkspaceId)
      .eq('meeting_id', next.id);
    if (existingActionsError) throw existingActionsError;
    const keepActionIds = new Set(actionRows.map(row => row.id));
    const staleActionIds = (existingActions || []).map(row => row.id).filter(id => !keepActionIds.has(id));
    if (staleActionIds.length) {
      const { error: staleActionError } = await supabase
        .from('app_meeting_action_items')
        .delete()
        .eq('workspace_id', currentWorkspaceId)
        .eq('meeting_id', next.id)
        .in('id', staleActionIds);
      if (staleActionError) throw staleActionError;
    }

    const links = [
      ...(next.relatedBoardIds || []).map(entityId => ({ entityType: 'board', entityId })),
      ...(next.actionItems || [])
        .filter(action => action.linkedTaskId)
        .map(action => ({ entityType: 'task', entityId: action.linkedTaskId })),
    ];
    const { error: deleteLinksError } = await supabase
      .from('app_meeting_links')
      .delete()
      .eq('workspace_id', currentWorkspaceId)
      .eq('meeting_id', next.id);
    if (deleteLinksError) throw deleteLinksError;
    if (links.length) {
      const { error: linkError } = await supabase.from('app_meeting_links').insert(
        links.map(link => ({
          workspace_id: currentWorkspaceId,
          meeting_id: next.id,
          entity_type: link.entityType,
          entity_id: link.entityId,
        })),
      );
      if (linkError) throw linkError;
    }
    const { error: completionMarkerError } = await supabase
      .from('app_meetings')
      .update({ content_saved_at: next.updatedAt, updated_at: next.updatedAt })
      .eq('workspace_id', currentWorkspaceId)
      .eq('id', next.id);
    if (completionMarkerError) throw completionMarkerError;
    return { ...next, contentSavedAt: next.updatedAt };
  }

  async function archive(id, archived = true) {
    const { error } = await supabase
      .from('app_meetings')
      .update({
        archived_at: archived ? new Date().toISOString() : null,
        updated_at: new Date().toISOString(),
      })
      .eq('workspace_id', currentWorkspaceId)
      .eq('id', requireValue(id, 'Meeting'));
    if (error) throw error;
  }

  async function remove(id) {
    const { error } = await supabase
      .from('app_meetings')
      .delete()
      .eq('workspace_id', currentWorkspaceId)
      .eq('id', requireValue(id, 'Meeting'));
    if (error) throw error;
  }

  return { list, load, save, archive, remove };
}
