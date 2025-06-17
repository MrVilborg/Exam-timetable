document.addEventListener('DOMContentLoaded', () => {
  const fileInput        = document.getElementById('file-input');
  const boardSelect      = document.getElementById('board-select');
  const examSelectGroup  = document.getElementById('exam-select-group');
  const examContainer    = document.getElementById('exam-checkboxes');
  const generateBtn      = document.getElementById('generate-btn');
  const downloadBtn      = document.getElementById('download-btn');
  const timetableContainer = document.getElementById('timetable-container');

  let examsData    = [];
  let filteredData = [];
  let lastTimetable= [];

  // Helper: turn Excel date+time into a JS Date for sorting
  function parseDateTime(row) {
    // --- parse date ---
    let dt;
    const dv = row['Date'];
    if (typeof dv === 'number') {
      // Excel serial date → JS date
      dt = new Date(Math.round((dv - (25567 + 2)) * 86400 * 1000));
    } else {
      // Try splitting on common separators
      const s = dv.toString().trim();
      const sep = s.includes('/') ? '/' : s.includes('-') ? '-' : '.';
      const parts = s.split(sep);
      let Y, M, D;
      if (parts[0].length === 4) [Y,M,D] = parts;
      else [D,M,Y] = parts;
      if (Y.length === 2) Y = '20' + Y;
      dt = new Date(+Y, +M - 1, +D);
    }

    // --- parse time ---
    const tv = row['Start Time'];
    let hours = 0, mins = 0;
    if (typeof tv === 'number') {
      const totalMins = Math.round(tv * 24 * 60);
      hours = Math.floor(totalMins / 60);
      mins  = totalMins % 60;
    } else if (typeof tv === 'string' && tv.includes(':')) {
      const [h,m] = tv.split(':').map(x => parseInt(x,10));
      hours = h; mins = m;
    }
    dt.setHours(hours, mins, 0, 0);
    return dt;
  }

  // Helper: format time fraction or string → "HH:MM"
  function formatTime(v) {
    if (typeof v === 'number') {
      const totalMins = Math.round(v * 24 * 60);
      const h = Math.floor(totalMins / 60);
      const m = totalMins % 60;
      return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
    }
    return v || '';
  }

  // Load Excel
  fileInput.addEventListener('change', e => {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type:'binary' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      examsData = XLSX.utils.sheet_to_json(sheet, { defval:'' });
      // if board already chosen, refresh list
      if (boardSelect.value) filterAndRender(boardSelect.value);
    };
    reader.readAsBinaryString(f);
  });

  // When IGCSE/IB is chosen
  boardSelect.addEventListener('change', e => {
    const b = e.target.value;
    if (b && examsData.length) {
      filterAndRender(b);
    } else {
      examSelectGroup.style.display = 'none';
      examContainer.innerHTML = '';
    }
  });

  // Build checkbox list
  function filterAndRender(board) {
    filteredData = examsData.filter(r => {
      const eb = (r['Exam Board']||'').toString().toUpperCase();
      return board==='IGCSE' 
        ? eb.includes('IGCSE') 
        : eb.includes('IB');
    });
    examContainer.innerHTML = '';
    filteredData.forEach((r,i) => {
      const div = document.createElement('div');
      div.className = 'checkbox-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox'; 
      cb.id = `ex-${i}`; 
      cb.value = i;
      const lbl = document.createElement('label');
      lbl.htmlFor = cb.id;
      lbl.textContent = board==='IGCSE'
        ? `${r['Exam (Paper)']} (${r['Exam Code (IGCSE)']})`
        : r['Exam (Paper)'];
      div.append(cb, lbl);
      examContainer.append(div);
    });
    examSelectGroup.style.display = 'block';
  }

  // Generate HTML timetable
  generateBtn.addEventListener('click', () => {
    try {
      const name = document.getElementById('student-name').value.trim();
      if (!name) throw new Error('Please enter the student name.');
      const checked = examContainer.querySelectorAll('input:checked');
      if (!checked.length) throw new Error('Please select at least one exam.');
      // pick rows & sort
      const sel = Array.from(checked).map(cb => filteredData[+cb.value]);
      sel.sort((a,b) => parseDateTime(a) - parseDateTime(b));
      lastTimetable = sel;

      // render table
      let html = `<h2>${name}'s Exam Timetable</h2><table><thead><tr>`;
      ['Date','Exam (Paper)','Exam Code (IGCSE)','Length (m)',
       'Start Time','Room','Exam Board']
        .forEach(c => html += `<th>${c}</th>`);
      html += '</tr></thead><tbody>';
      sel.forEach(r => {
        html += '<tr>'+
          `<td>${r['Date']}</td>`+
          `<td>${r['Exam (Paper)']}</td>`+
          `<td>${r['Exam Code (IGCSE)']||''}</td>`+
          `<td>${r['Length (m)']}</td>`+
          `<td>${formatTime(r['Start Time'])}</td>`+
          `<td>${r['Room']}</td>`+
          `<td>${r['Exam Board']}</td>`+
        '</tr>';
      });
      html += '</tbody></table>';
      timetableContainer.innerHTML = html;
      downloadBtn.style.display = 'inline-block';
    } catch (err) {
      alert(err.message);
      console.error(err);
    }
  });

  // PDF download
  downloadBtn.addEventListener('click', () => {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p','pt','a4');
    const name = document.getElementById('student-name').value.trim();
    doc.setFontSize(16);
    doc.text(name + "'s Exam Timetable", 40, 40);
    const head = [[
      'Date','Exam (Paper)','Exam Code (IGCSE)','Length (m)',
      'Start Time','Room','Exam Board'
    ]];
    const body = lastTimetable.map(r => [
      r['Date'],
      r['Exam (Paper)'],
      r['Exam Code (IGCSE)']||'',
      r['Length (m)'],
      formatTime(r['Start Time']),
      r['Room'],
      r['Exam Board']
    ]);
    doc.autoTable({ startY:60, head, body, theme:'grid' });
    doc.save('exam_timetable.pdf');
  });
});
