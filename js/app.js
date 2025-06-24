document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('file-input');
  const boardSelect = document.getElementById('board-select');
  const examSelectGroup = document.getElementById('exam-select-group');
  const examAvailable = document.getElementById('exam-available');
  const examSelected = document.getElementById('exam-selected');
  const addBtn = document.getElementById('add-btn');
  const removeBtn = document.getElementById('remove-btn');
  const clearBtn = document.getElementById('clear-btn');
  const extra25Chk = document.getElementById('extra25');
  const extra50Chk = document.getElementById('extra50');
  const sepInvigChk = document.getElementById('sep-invig');
  const customRoomGroup = document.getElementById('custom-room-group');
  const customRoomInput = document.getElementById('custom-room');
  const generateBtn = document.getElementById('generate-btn');
  const downloadBtn = document.getElementById('download-btn');
  const timetableContainer = document.getElementById('timetable-container');

  let examsData = [], filteredData = [], lastTimetable = [];

  function updateUI() {
    if (boardSelect.value && examsData.length) {
      const board = boardSelect.value;
      filteredData = examsData.filter(r => (r['Exam Board']||'').toString().toUpperCase().includes(board));
      examAvailable.innerHTML = '';
      examSelected.innerHTML = '';
      filteredData.forEach((r,i) => {
        const opt = document.createElement('option');
        opt.value = i;
        opt.text = board==='IGCSE'
                  ? `${r['Exam (Paper)']} (${r['Exam Code (IGCSE)']})`
                  : r['Exam (Paper)'];
        examAvailable.append(opt);
      });
      examSelectGroup.style.display = 'block';
      extra25Chk.parentElement.style.display = 'inline-block';
      extra50Chk.parentElement.style.display = 'inline-block';
      sepInvigChk.parentElement.style.display = 'inline-block';
    } else {
      examSelectGroup.style.display = 'none';
      extra25Chk.parentElement.style.display = 'none';
      extra50Chk.parentElement.style.display = 'none';
      sepInvigChk.parentElement.style.display = 'none';
      customRoomGroup.style.display = 'none';
    }
  }

  fileInput.addEventListener('change', e => {
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, {type:'binary'});
      examsData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''});
      updateUI();
    };
    reader.readAsBinaryString(f);
  });

  boardSelect.addEventListener('change', updateUI);
  sepInvigChk.addEventListener('change', () => {
    customRoomGroup.style.display = sepInvigChk.checked ? 'block' : 'none';
  });

  addBtn.addEventListener('click', () => {
    Array.from(examAvailable.selectedOptions).forEach(o => examSelected.append(o));
  });
  removeBtn.addEventListener('click', () => {
    Array.from(examSelected.selectedOptions).forEach(o => examAvailable.append(o));
  });
  clearBtn.addEventListener('click', () => {
    Array.from(examSelected.options).forEach(o => examAvailable.append(o));
  });

  function parseDateTime(r) {
    let dt;
    const dv = r['Date'];
    if (typeof dv === 'number') {
      dt = new Date((dv - (25567 + 2)) * 86400 * 1000);
    } else {
      const parts = dv.split(/[-\/\.]/);
      const [Y,M,D] = parts[0].length===4 ? parts : [parts[2],parts[1],parts[0]];
      dt = new Date(+Y, +M-1, +D);
    }
    const tv = r['Start Time'];
    let h=0,m=0;
    if (typeof tv==='number') {
      const t = Math.round(tv*24*60);
      h = Math.floor(t/60);
      m = t%60;
    } else if (typeof tv==='string' && tv.includes(':')) {
      [h,m] = tv.split(':').map(n=>+n);
    }
    dt.setHours(h,m,0,0);
    return dt;
  }

  function formatTime(v) {
    if (typeof v==='number') {
      const t = Math.round(v*24*60);
      const h = Math.floor(t/60), m = t%60;
      return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
    }
    return v||'';
  }

  function loadLogo(cb) {
    const img = new Image();
    img.crossOrigin = '';
    img.onload = () => {
      const c = document.createElement('canvas');
      c.width = img.width; c.height = img.height;
      c.getContext('2d').drawImage(img,0,0);
      cb(c.toDataURL('PNG'));
    };
    img.onerror = () => cb(null);
    img.src = 'logo1.png';
  }

  document.getElementById('download-template-btn').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    const headers = [['Date','Exam (Paper)','Exam Code (IGCSE)','Length (m)','Start Time','Room','Exam Board']];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'exam_timetable_template.xlsx');
  });

  generateBtn.addEventListener('click', () => {
    const name = document.getElementById('student-name').value.trim();
    if (!name) { alert('Enter student name.'); return; }
    const opts = Array.from(examSelected.options);
    if (!opts.length) { alert('Select exams.'); return; }
    let rows = opts.map(o => filteredData[o.value]);
    rows = rows.map(r => {
      const base = parseFloat(r['Length (m)'])||0;
      const factor = extra50Chk.checked?1.5:extra25Chk.checked?1.25:1;
      const room = sepInvigChk.checked?(customRoomInput.value.trim()||r['Room']):r['Room'];
      return {...r,'_length':Math.round(base*factor),'_room':room};
    });
    rows.sort((a,b)=>parseDateTime(a)-parseDateTime(b));
    lastTimetable = rows;
    const isIB = boardSelect.value==='IB';
    const headers = isIB
      ? ['Date','Exam (Paper)','Length (m)','Start Time','Room','Exam Board']
      : ['Date','Exam (Paper)','Exam Code (IGCSE)','Length (m)','Start Time','Room','Exam Board'];
    let html = `<h2>${name}'s Exam Timetable</h2><table><thead><tr>` +
      headers.map(h=>`<th>${h}</th>`).join('') + `</tr></thead><tbody>`;
    rows.forEach(r => {
      html += '<tr>';
      html += `<td>${r['Date']}</td><td>${r['Exam (Paper)']}</td>`;
      if (!isIB) html += `<td>${r['Exam Code (IGCSE)']||''}</td>`;
      html += `<td>${r['_length']}</td><td>${formatTime(r['Start Time'])}</td>`; 
      html += `<td>${r['_room']}</td><td>${r['Exam Board']}</td>`;
      html += '</tr>';
    });
    html += '</tbody></table>';
    timetableContainer.innerHTML = html;
    downloadBtn.style.display = 'inline-block';
  });

  downloadBtn.addEventListener('click', () => {
    loadLogo(logoData => {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF('p','pt','a4');
      // Salutation
      const name = document.getElementById('student-name').value.trim();
      doc.setFontSize(12).setTextColor(0,0,0);
      doc.text(`Dear ${name},`, 40, 40);
      const intro = [
        "The following is your examination timetable for the Summer session.",
        "Please read carefully and ensure every examination is listed.",
        "Once done, please sign below to confirm."
      ];
      intro.forEach((l,i)=>doc.text(l,40,60 + i*14));
      // Table
      const isIB = boardSelect.value==='IB';
      const head = isIB
        ? [['Date','Exam (Paper)','Length (m)','Start Time','Room','Exam Board']]
        : [['Date','Exam (Paper)','Exam Code (IGCSE)','Length (m)','Start Time','Room','Exam Board']];
      const body = lastTimetable.map(r=>{
        const row=[r['Date'],r['Exam (Paper)']];
        if(!isIB) row.push(r['Exam Code (IGCSE)']||'');
        row.push(r['_length'],formatTime(r['Start Time']),r['_room'],r['Exam Board']);
        return row;
      });
      doc.autoTable({
        startY:100,head,body,theme:'grid',
        headStyles:{fillColor:[0,71,142],textColor:[255,255,255]},
        alternateRowStyles:{fillColor:[240,240,240]},
        tableLineColor:[14,32,74],tableLineWidth:0.5
      });
      // Confirmation
      const confirmText="I confirm that this timetable is accurate and reflects my official statement of entry.";
      const confirmY=doc.lastAutoTable.finalY+20;
      doc.setFontSize(12).setTextColor(0,0,0);
      doc.text(confirmText,40,confirmY);
      // Signatures
      const ph=doc.internal.pageSize.getHeight();
      doc.text("Candidate signature: ____________________",40,ph-60);
      doc.text("Date: ____________",40,ph-40);
      doc.text("Examinations Officer: ____________________",300,ph-60);
      doc.text("Date: ____________",300,ph-40);
      // Logos
      if(logoData) {
        doc.addImage(logoData,'PNG',450,15,100,50);
      }
      doc.save('exam_timetable.pdf');
    });
  });
});
