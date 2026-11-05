/**
 * 나르카 회의실 예약 (Google Apps Script)
 *
 * 흐름: 웹 예약 → 회의실 캘린더 등록 + 참석자 초대 → 슬랙 채널 공유 → 시작 전 리마인드
 * 비밀값(슬랙 웹훅)은 코드에 넣지 않고 스크립트 속성 SLACK_WEBHOOK_URL 에 저장합니다.
 */
const CONFIG = {
  DOMAIN: 'uncmhm.com',
  TZ: 'Asia/Seoul',
  OPEN_FROM: '2026-11-26',   // 예약 가능 시작일 (사무실 이전일)
  HOUR_START: 8,
  HOUR_END: 21,
  REMIND_MINUTES: 10,        // 시작 몇 분 전에 슬랙 리마인드할지
  ROOMS: [
    { key: 'big', name: '대회의실', capacity: 12 },
    { key: 'a', name: '회의실 A', capacity: 6 },
    { key: 'b', name: '회의실 B', capacity: 6 },
  ],
};

/* ---------- 최초 1회 실행 ---------- */

/** 회의실 캘린더 3개와 리마인드 예약 작업을 만듭니다. 다시 실행해도 중복 생성되지 않습니다. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  CONFIG.ROOMS.forEach((room) => {
    const prop = 'CAL_' + room.key;
    if (props.getProperty(prop)) return;
    const cal = CalendarApp.createCalendar('나르카 회의실 · ' + room.name, {
      timeZone: CONFIG.TZ,
      summary: '회의실 예약 페이지에서 관리되는 캘린더입니다. 직접 수정하지 마세요.',
    });
    props.setProperty(prop, cal.getId());
  });
  const hasTrigger = ScriptApp.getProjectTriggers()
    .some((t) => t.getHandlerFunction() === 'sendReminders');
  if (!hasTrigger) ScriptApp.newTrigger('sendReminders').timeBased().everyMinutes(5).create();
}

/* ---------- 웹 페이지 ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('나르카 회의실 예약')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBootstrap() {
  const me = currentUser_();
  return {
    me,
    rooms: CONFIG.ROOMS,
    today: dateStr_(new Date()),
    openFrom: CONFIG.OPEN_FROM,
    hourStart: CONFIG.HOUR_START,
    hourEnd: CONFIG.HOUR_END,
    domain: CONFIG.DOMAIN,
    remindMinutes: CONFIG.REMIND_MINUTES,
    members: listMembers_(),
  };
}

function getBookings(date) {
  assertDate_(date);
  const me = currentUser_();
  const start = kst_(date, '00:00');
  const end = new Date(start.getTime() + 86400000);
  const out = [];
  CONFIG.ROOMS.forEach((room) => {
    roomCalendar_(room.key).getEvents(start, end).forEach((e) => {
      out.push(toBooking_(e, room, me.email));
    });
  });
  return out;
}

function getMyBookings() {
  const me = currentUser_();
  const start = new Date();
  const end = new Date(start.getTime() + 60 * 86400000);
  const out = [];
  CONFIG.ROOMS.forEach((room) => {
    roomCalendar_(room.key).getEvents(start, end).forEach((e) => {
      if (e.getTag('ownerEmail') === me.email) {
        out.push(Object.assign(toBooking_(e, room, me.email), {
          date: dateStr_(e.getStartTime()),
        }));
      }
    });
  });
  return out.sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
}

function createBooking(p) {
  const me = currentUser_();
  const room = CONFIG.ROOMS.find((r) => r.key === p.roomKey);
  if (!room) throw new Error('회의실을 선택해 주세요.');
  const title = String(p.title || '').trim();
  if (!title || title.length > 80) throw new Error('회의명은 1~80자로 입력해 주세요.');
  assertDate_(p.date);
  if (p.date < CONFIG.OPEN_FROM) throw new Error(CONFIG.OPEN_FROM + '부터 예약할 수 있습니다.');
  assertTime_(p.start);
  assertTime_(p.end);

  const start = kst_(p.date, p.start);
  const end = kst_(p.date, p.end);
  if (start >= end) throw new Error('종료 시각은 시작 시각보다 늦어야 합니다.');
  if (start < new Date()) throw new Error('이미 지난 시간은 예약할 수 없습니다.');

  // 회사 외부 주소로 초대가 나가지 않도록 도메인을 제한합니다.
  const guests = Array.from(new Set([me.email].concat(p.attendees || [])
    .map((x) => String(x).trim().toLowerCase())));
  guests.forEach((email) => {
    if (!isCompanyEmail_(email)) throw new Error('회사 계정만 초대할 수 있습니다: ' + email);
  });

  const isPrivate = !!p.isPrivate;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let event;
  try {
    const cal = roomCalendar_(room.key);
    const taken = cal.getEvents(start, end)
      .some((e) => e.getStartTime() < end && e.getEndTime() > start);
    if (taken) throw new Error('방금 다른 예약이 등록된 시간입니다. 다른 시간을 선택해 주세요.');

    event = cal.createEvent(title, start, end, {
      guests: guests.join(','),
      sendInvites: true,
      location: room.name + ' (서울 마포구 잔다리로 39, 201호)',
      description: '예약자: ' + me.name + '\n회의실 예약 페이지에서 등록되었습니다.\n'
        + '시간 변경·취소는 예약 페이지에서 진행해 주세요.',
    });
    event.setTag('ownerEmail', me.email);
    event.setTag('ownerName', me.name);
    event.setTag('private', isPrivate ? '1' : '');
    event.setTag('guestNames', guestNames_(guests).join(', ').slice(0, 900));
    // 곧 시작하는 회의는 예약 알림과 리마인드가 겹치지 않도록 처리합니다.
    if (start.getTime() - Date.now() <= CONFIG.REMIND_MINUTES * 60000) event.setTag('reminded', '1');
  } finally {
    lock.releaseLock();
  }

  const slackOk = postSlack_(':calendar: *회의실 예약*\n' + slackBody_(event, room));
  return { booking: toBooking_(event, room, me.email), slackOk };
}

function cancelBooking(roomKey, eventId) {
  const me = currentUser_();
  const room = CONFIG.ROOMS.find((r) => r.key === roomKey);
  if (!room) throw new Error('회의실 정보가 올바르지 않습니다.');
  const event = roomCalendar_(room.key).getEventById(eventId);
  if (!event) throw new Error('이미 취소된 예약입니다.');
  if (event.getTag('ownerEmail') !== me.email) throw new Error('본인 예약만 취소할 수 있습니다.');
  const body = slackBody_(event, room);
  event.deleteEvent();
  postSlack_(':x: *회의실 예약 취소*\n' + body);
  return true;
}

/* ---------- 리마인드 (5분마다 자동 실행) ---------- */

function sendReminders() {
  const now = new Date();
  const until = new Date(now.getTime() + CONFIG.REMIND_MINUTES * 60000);
  CONFIG.ROOMS.forEach((room) => {
    roomCalendar_(room.key).getEvents(now, until).forEach((e) => {
      if (e.getStartTime() <= now || e.getTag('reminded')) return;
      const ok = postSlack_(':bell: *' + CONFIG.REMIND_MINUTES + '분 내 시작*\n' + slackBody_(e, room));
      if (ok) e.setTag('reminded', '1');
    });
  });
}

/* ---------- 내부 함수 ---------- */

function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!isCompanyEmail_(email)) throw new Error('회사 구글 계정으로 접속해 주세요.');
  const found = listMembers_().find((m) => m.email === email);
  return { email, name: found ? found.name : email.split('@')[0] };
}

function isCompanyEmail_(email) {
  return /^[a-z0-9._%+-]+@/.test(email) && email.endsWith('@' + CONFIG.DOMAIN);
}

/** 사내 디렉터리에서 구성원 목록을 읽습니다. 실패하면 빈 목록을 돌려주고 페이지는 이메일 직접 입력으로 동작합니다. */
function listMembers_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('members');
  if (hit) return JSON.parse(hit);
  try {
    const users = [];
    let pageToken;
    do {
      const res = AdminDirectory.Users.list({
        domain: CONFIG.DOMAIN, viewType: 'domain_public', maxResults: 200, pageToken,
      });
      (res.users || []).forEach((u) => users.push({
        email: String(u.primaryEmail).toLowerCase(),
        name: u.name && u.name.fullName ? u.name.fullName : u.primaryEmail,
      }));
      pageToken = res.nextPageToken;
    } while (pageToken);
    users.sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    cache.put('members', JSON.stringify(users), 21600);
    return users;
  } catch (err) {
    console.warn('구성원 목록 조회 실패: ' + err);
    return [];
  }
}

function guestNames_(emails) {
  const members = listMembers_();
  return emails.map((email) => {
    const m = members.find((x) => x.email === email);
    return m ? m.name : email.split('@')[0];
  });
}

function roomCalendar_(key) {
  const id = PropertiesService.getScriptProperties().getProperty('CAL_' + key);
  if (!id) throw new Error('초기 설정(setup)이 실행되지 않았습니다.');
  return CalendarApp.getCalendarById(id);
}

function toBooking_(e, room, myEmail) {
  const mine = e.getTag('ownerEmail') === myEmail;
  const hidden = e.getTag('private') === '1' && !mine;
  return {
    id: e.getId(),
    roomKey: room.key,
    start: timeStr_(e.getStartTime()),
    end: timeStr_(e.getEndTime()),
    title: hidden ? '비공개 회의' : e.getTitle(),
    ownerName: hidden ? '' : (e.getTag('ownerName') || ''),
    guestNames: hidden ? '' : (e.getTag('guestNames') || ''),
    isPrivate: e.getTag('private') === '1',
    mine,
  };
}

/** 비공개 예약은 회의명·참석자를 슬랙에 노출하지 않습니다. */
function slackBody_(e, room) {
  const when = Utilities.formatDate(e.getStartTime(), CONFIG.TZ, 'M/d')
    + ' ' + timeStr_(e.getStartTime()) + '~' + timeStr_(e.getEndTime());
  if (e.getTag('private') === '1') return '> ' + room.name + ' · ' + when + '\n> 비공개 회의';
  const owner = e.getTag('ownerName') || '';
  const guests = e.getTag('guestNames') || '';
  return '> *' + slackEscape_(e.getTitle()) + '*\n> ' + room.name + ' · ' + when
    + '\n> 예약자 ' + slackEscape_(owner) + (guests ? ' · 참석 ' + slackEscape_(guests) : '');
}

function slackEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 슬랙 전송이 실패해도 예약은 유지합니다. */
function postSlack_(text) {
  const url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!url) return false;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text }),
      muteHttpExceptions: true,
    });
    return res.getResponseCode() === 200;
  } catch (err) {
    console.warn('슬랙 전송 실패: ' + err);
    return false;
  }
}

function kst_(date, time) { return new Date(date + 'T' + time + ':00+09:00'); }
function dateStr_(d) { return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd'); }
function timeStr_(d) { return Utilities.formatDate(d, CONFIG.TZ, 'HH:mm'); }

function assertDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) throw new Error('날짜 형식이 올바르지 않습니다.');
}

function assertTime_(s) {
  const m = /^(\d{2}):(00|30)$/.exec(String(s));
  const h = m ? Number(m[1]) : -1;
  const valid = m && h >= CONFIG.HOUR_START
    && (h < CONFIG.HOUR_END || (h === CONFIG.HOUR_END && m[2] === '00'));
  if (!valid) throw new Error('예약 시간은 30분 단위로 선택해 주세요.');
}
