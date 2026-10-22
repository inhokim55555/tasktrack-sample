function localDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
}

function localTime(date) {
  return `${String(date.getHours()).padStart(2,'0')}:${String(date.getMinutes()).padStart(2,'0')}`;
}

export function googleEventCalendarSpan(event) {
  const allDay = event.allDay ?? event.all_day;
  if (allDay) {
    const start = event.startDate ?? event.start_date;
    const exclusive = event.endDateExclusive ?? event.end_date_exclusive;
    const end = new Date(`${exclusive}T00:00:00Z`);
    end.setUTCDate(end.getUTCDate()-1);
    return {start,end:end.toISOString().slice(0,10),startTime:'',endTime:''};
  }
  const startDate = new Date(event.startAt ?? event.start_at);
  const endDate = new Date(event.endAt ?? event.end_at);
  return {start:localDate(startDate),end:localDate(new Date(endDate.getTime()-1)),
    startTime:localTime(startDate),endTime:localTime(endDate)};
}

export function googleEventDisplayEnd(event) {
  const span = googleEventCalendarSpan(event);
  if (event.allDay ?? event.all_day) return {date:span.end,time:''};
  const end = new Date(event.endAt ?? event.end_at);
  return {date:localDate(end),time:localTime(end)};
}

export function googleCalendarItems(privateEvents=[],sharedEvents=[],sharedGoogleIds=new Set()) {
  const privateItems = privateEvents.filter(event=>!sharedGoogleIds.has(event.id)).map(event=>({
    id:`private:${event.id}`,type:'google-private',event,
    ...googleEventCalendarSpan(event)
  }));
  const sharedItems = sharedEvents.filter(event=>event.sync_status !== 'canceled').map(event=>({
    id:`shared:${event.id}`,type:'google-shared',event,
    ...googleEventCalendarSpan(event)
  }));
  return [...privateItems,...sharedItems];
}
