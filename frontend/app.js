const main = document.getElementById('appMain');
const todayLabel = document.getElementById('todayLabel');
const navButtons = [...document.querySelectorAll('.nav-item')];

const today = new Date();
const STORAGE_KEY_EVENTS = 'allinone.calendar.events.v1';
let calendarCursor = new Date(today.getFullYear(), today.getMonth(), 1);
let calendarSelectedDate = null;

const weekdayFmt = new Intl.DateTimeFormat('it-IT', { weekday: 'short' });
const dayFmt = new Intl.DateTimeFormat('it-IT', { day: '2-digit' });
const fullFmt = new Intl.DateTimeFormat('it-IT', {
  weekday: 'long',
  day: '2-digit',
  month: 'long',
  year: 'numeric',
});
const monthFmt = new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' });

todayLabel.textContent = fullFmt.format(today);

function toIsoDate(year, monthIndex, day) {
  return `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function getSavedEvents() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_EVENTS);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function setSavedEvents(eventsMap) {
  localStorage.setItem(STORAGE_KEY_EVENTS, JSON.stringify(eventsMap));
}

function parseIsoDate(isoDate) {
  const [year, month, day] = isoDate.split('-').map(Number);
  return new Date(year, month - 1, day);
}

function escapeHtml(text) {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function getEasterDate(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

function getHolidayMap(year) {
  const fixed = {
    [toIsoDate(year, 0, 1)]: 'Capodanno',
    [toIsoDate(year, 0, 6)]: 'Epifania',
    [toIsoDate(year, 3, 25)]: 'Festa della Liberazione',
    [toIsoDate(year, 4, 1)]: 'Festa dei Lavoratori',
    [toIsoDate(year, 5, 2)]: 'Festa della Repubblica',
    [toIsoDate(year, 7, 15)]: 'Ferragosto',
    [toIsoDate(year, 10, 1)]: 'Ognissanti',
    [toIsoDate(year, 11, 8)]: 'Immacolata Concezione',
    [toIsoDate(year, 11, 25)]: 'Natale',
    [toIsoDate(year, 11, 26)]: 'Santo Stefano',
  };

  const easter = getEasterDate(year);
  const easterIso = toIsoDate(year, easter.getMonth(), easter.getDate());
  const easterMonday = new Date(easter);
  easterMonday.setDate(easterMonday.getDate() + 1);
  const easterMondayIso = toIsoDate(year, easterMonday.getMonth(), easterMonday.getDate());

  return {
    ...fixed,
    [easterIso]: 'Pasqua',
    [easterMondayIso]: "Lunedì dell'Angelo",
  };
}

function getStartOfWeek(date) {
  const d = new Date(date);
  const day = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function renderHome() {
  const start = getStartOfWeek(today);
  const weekDays = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const isToday = d.toDateString() === today.toDateString();
    return `
      <div class="day-chip ${isToday ? 'is-today' : ''}">
        <span class="name">${weekdayFmt.format(d).replace('.', '')}</span>
        <span class="num">${dayFmt.format(d)}</span>
      </div>
    `;
  }).join('');

  const actions = [
    ['Aggiungi allenamento', 'Routine / Cardio / Forza'],
    ['Registra pasto', 'Colazione, pranzo, cena, snack'],
    ['Sessione hobby', 'Musica, lettura, progetto'],
    ['Idratazione', 'Aggiungi bicchiere d\'acqua'],
    ['Check energia', 'Come ti senti oggi?'],
    ['Nota veloce', 'Promemoria rapido'],
  ];

  main.innerHTML = `
    <section class="panel">
      <h2>Settimana attuale</h2>
      <div class="week-strip">${weekDays}</div>
    </section>

    <section class="panel">
      <h2>Azioni rapide</h2>
      <div class="quick-actions">
        ${actions
          .map(
            ([title, hint]) => `
              <button class="quick-btn" type="button">
                ${title}
                <span>${hint}</span>
              </button>
            `,
          )
          .join('')}
      </div>
    </section>
  `;
}

function renderCalendar() {
  const nowYear = calendarCursor.getFullYear();
  const nowMonth = calendarCursor.getMonth();
  const eventMap = getSavedEvents();
  const holidayMap = getHolidayMap(nowYear);
  const firstDay = new Date(nowYear, nowMonth, 1);
  const monthName = monthFmt.format(firstDay);
  const daysInMonth = new Date(nowYear, nowMonth + 1, 0).getDate();
  const weekStartOffset = (firstDay.getDay() + 6) % 7;

  const cells = [];
  for (let i = 0; i < weekStartOffset; i += 1) {
    cells.push('<div class="cell is-empty"></div>');
  }

  for (let day = 1; day <= daysInMonth; day += 1) {
    const isoDate = toIsoDate(nowYear, nowMonth, day);
    const eventsCount = Array.isArray(eventMap[isoDate]) ? eventMap[isoDate].length : 0;
    const holidayName = holidayMap[isoDate];
    const isToday = isoDate === toIsoDate(today.getFullYear(), today.getMonth(), today.getDate());
    const isSelected = calendarSelectedDate === isoDate;

    cells.push(`
      <button
        type="button"
        class="cell ${isToday ? 'is-today' : ''} ${holidayName ? 'is-holiday' : ''} ${isSelected ? 'is-selected' : ''}"
        data-day="${day}"
        data-date="${isoDate}"
        title="${holidayName || ''}"
      >
        <span class="cell-num">${day}</span>
        ${holidayName ? '<span class="holiday-dot" aria-hidden="true"></span>' : ''}
        ${eventsCount > 0 ? `<span class="event-dot" aria-label="${eventsCount} eventi salvati"></span>` : ''}
      </button>
    `);
  }

  const selectedEvents = calendarSelectedDate ? eventMap[calendarSelectedDate] || [] : [];
  const selectedDateLabel = calendarSelectedDate
    ? new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: 'long', year: 'numeric' }).format(
        parseIsoDate(calendarSelectedDate),
      )
    : null;

  main.innerHTML = `
    <section class="panel">
      <div class="month-head">
        <div class="month-nav">
          <button type="button" class="month-nav-btn" id="prevMonthBtn" aria-label="Mese precedente">◀</button>
          <h2>${monthName.charAt(0).toUpperCase() + monthName.slice(1)}</h2>
          <button type="button" class="month-nav-btn" id="nextMonthBtn" aria-label="Mese successivo">▶</button>
        </div>
        <p>Vista mensile</p>
      </div>

      <div class="month-grid">
        <div class="wk">Lun</div>
        <div class="wk">Mar</div>
        <div class="wk">Mer</div>
        <div class="wk">Gio</div>
        <div class="wk">Ven</div>
        <div class="wk">Sab</div>
        <div class="wk">Dom</div>
        ${cells.join('')}
      </div>

      <div class="calendar-events">
        ${
          selectedDateLabel
            ? `
              <h3>Eventi del ${selectedDateLabel}</h3>
              ${
                selectedEvents.length
                  ? `<ul>${selectedEvents.map((event) => `<li>${escapeHtml(event)}</li>`).join('')}</ul>`
                  : '<p>Nessun evento salvato.</p>'
              }
            `
            : '<p>Seleziona un giorno per aggiungere un evento.</p>'
        }
      </div>
    </section>
  `;

  const prevBtn = document.getElementById('prevMonthBtn');
  const nextBtn = document.getElementById('nextMonthBtn');
  prevBtn?.addEventListener('click', () => {
    calendarCursor = new Date(nowYear, nowMonth - 1, 1);
    renderCalendar();
  });
  nextBtn?.addEventListener('click', () => {
    calendarCursor = new Date(nowYear, nowMonth + 1, 1);
    renderCalendar();
  });

  main.querySelectorAll('.month-grid .cell[data-date]').forEach((cellBtn) => {
    cellBtn.addEventListener('click', () => {
      const isoDate = cellBtn.dataset.date;
      if (!isoDate) return;

      calendarSelectedDate = isoDate;
      const eventTitle = window.prompt('Nome evento');
      if (eventTitle && eventTitle.trim()) {
        const updatedEvents = getSavedEvents();
        const dayEvents = Array.isArray(updatedEvents[isoDate]) ? updatedEvents[isoDate] : [];
        dayEvents.push(eventTitle.trim());
        updatedEvents[isoDate] = dayEvents;
        setSavedEvents(updatedEvents);
      }

      renderCalendar();
    });
  });
}

function renderWorkout() {
  main.innerHTML = `
    <section class="panel workout-empty">
      <div>
        <h2>Workout</h2>
        <p>Pagina pronta: qui arriveranno cronologia, schede e statistiche.</p>
      </div>
    </section>
  `;
}

function setView(view) {
  navButtons.forEach((btn) => {
    btn.classList.toggle('is-active', btn.dataset.view === view);
  });

  if (view === 'home') renderHome();
  else if (view === 'calendario') renderCalendar();
  else renderWorkout();
}

navButtons.forEach((btn) => {
  btn.addEventListener('click', () => setView(btn.dataset.view));
});

setView('home');
