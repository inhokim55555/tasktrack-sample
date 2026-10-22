export function sharedEventFeedItem(row, read, memberId) {
  const canceled = row.sync_status === 'canceled';
  const event = {
    id:row.id,
    title:row.title,
    location:row.location || '',
    description:row.description || '',
    all_day:row.all_day,
    start_at:row.start_at || null,
    end_at:row.end_at || null,
    start_date:row.start_date || null,
    end_date_exclusive:row.end_date_exclusive || null,
    sync_status:row.sync_status,
    sync_error:row.sync_error || '',
    last_synced_at:row.last_synced_at || null
  };
  return {
    id:`shared-calendar:${row.id}`,
    kind:'shared_calendar',
    sharedEventId:row.id,
    authorId:row.owner_member_id,
    title:row.title,
    event,
    canceled,
    body:canceled ? `${row.title} was canceled in Google Calendar.` : `${row.title} was shared to the workspace.`,
    boardId:'',groupId:'',taskId:'',
    mentions:['__everyone'],
    readBy:read && memberId ? [memberId] : [],
    createdAt:Date.parse(row.shared_at) || Date.now()
  };
}

export function mergeSharedCalendarFeed(updates, rows, readIds, memberId) {
  return [
    ...(updates || []),
    ...(rows || []).map(row=>sharedEventFeedItem(row,readIds?.has(row.id),memberId))
  ].sort((a,b)=>(Number(b.createdAt)||0)-(Number(a.createdAt)||0));
}

export function calendarFeedRelevantToMember(item) {
  return item?.kind === 'shared_calendar';
}
