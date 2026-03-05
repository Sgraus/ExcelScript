const main = document.getElementById('appMain');
const todayLabel = document.getElementById('todayLabel');
const navButtons = [...document.querySelectorAll('.nav-item')];

const today = new Date();
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
  const nowYear = today.getFullYear();
  const nowMonth = today.getMonth();

  const firstDay = new Date(nowYear, nowMonth, 1);
  const monthName = monthFmt.format(firstDay);
  const daysInMonth = new Date(nowYear, nowMonth + 1, 0).getDate();
  const weekStartOffset = (firstDay.getDay() + 6) % 7;

  const cells = [];
  for (let i = 0; i < weekStartOffset; i += 1) {
    cells.push('<div class="cell is-empty"></div>');
  }

  for (let day = 1; day <= daysInMonth; day += 1) {
    const isToday =
      day === today.getDate() && nowMonth === today.getMonth() && nowYear === today.getFullYear();

    cells.push(`<div class="cell ${isToday ? 'is-today' : ''}">${day}</div>`);
  }

  main.innerHTML = `
    <section class="panel">
      <div class="month-head">
        <h2>${monthName.charAt(0).toUpperCase() + monthName.slice(1)}</h2>
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
    </section>
  `;
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
