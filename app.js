// ... existing code ...
import Chart from "chart.js/auto";
import { jsPDF } from "jspdf";

const state = {
  processes: [] // {id,name,people: n,qualityTarget, tasks: [{id,name,diff,cost,qty,assignments:[]}]}
};

function $(id){return document.getElementById(id)}

// Utilities
const uid = (p='i') => (p+Math.random().toString(36).slice(2,9));

// DOM refs
const refs = {
  procName: $('proc-name'),
  procPeople: $('proc-people'),
  procQualityTarget: $('proc-quality-target'),
  addProcessBtn: $('add-process'),
  procSelect: $('proc-select'),
  updateProcessBtn: $('update-process'),

  taskProcess: $('task-process'),
  taskName: $('task-name'),
  taskDiff: $('task-diff'),
  taskCost: $('task-cost'),
  taskQty: $('task-qty'),
  addTaskBtn: $('add-task'),

  assignProcess: $('assign-process'),
  assignPersons: $('assign-persons'),
  saveAssign: $('save-assign'),

  runSim: $('run-sim'),
  resetBtn: $('reset'),

  // New refs for topbar buttons
  loadSampleBtn: $('load-sample'),
  resetTemplatesBtn: $('reset-templates'),

  records: $('records'),
  sumTasks: $('sum-tasks'),
  sumCost: $('sum-cost'),

  chartTasksCtx: $('chart-tasks').getContext('2d'),
  chartEffCtx: $('chart-eff').getContext('2d'),
  chartQualCtx: $('chart-qual').getContext('2d'),
  personDetails: $('person-details'),
  analysisBox: $('analysis-box'),

  simSpeed: $('sim-speed'),
  simNoise: $('sim-noise'),

  // New DOM refs for exports and help
  exportPdf: $('export-pdf'),
  exportWord: $('export-word'),
  loadBackupBtn: $('load-backup'),

  // NEW: task manager button and table modal
  manageTasksBtn: $('manage-tasks'),
  tasksModal: $('tasks-modal'),
  tasksTableBody: document.querySelector('#tasks-table tbody'),
  tasksClose: $('tasks-close'),
  tasksClose2: $('tasks-close-2')
};

// Message modal refs (custom notification + decision)
const msgModal = $('msg-modal');
const msgText = $('msg-text');
const msgYes = $('msg-yes');
const msgNo = $('msg-no');

// Help modal controls
const helpFab = document.getElementById('help-fab');
const helpModal = document.getElementById('help-modal');
const helpClose = document.getElementById('help-close');
const helpDownloadExample = document.getElementById('help-download-example');

helpFab.addEventListener('click', ()=>{ helpModal.style.display='flex'; helpModal.setAttribute('aria-hidden','false'); });
helpClose.addEventListener('click', ()=>{ helpModal.style.display='none'; helpModal.setAttribute('aria-hidden','true'); });
helpModal.addEventListener('click',(e)=>{ if(e.target===helpModal){ helpModal.style.display='none'; helpModal.setAttribute('aria-hidden','true'); } });

// Template menu modal wiring (open/close behavior)
const menuAddProcess = document.getElementById('menu-add-process');
const menuAddTask = document.getElementById('menu-add-task');
const menuAssign = document.getElementById('menu-assign');
const menuRecords = document.getElementById('menu-records');

const modalAddProcess = document.getElementById('modal-add-process');
const modalAddTask = document.getElementById('modal-add-task');
const modalAssign = document.getElementById('modal-assign');
const modalRecords = document.getElementById('modal-records');

function openTemplateModal(modal){
  if(!modal) return;
  modal.style.display = 'flex';
  modal.setAttribute('aria-hidden','false');
}
function closeTemplateModal(modal){
  if(!modal) return;
  modal.style.display = 'none';
  modal.setAttribute('aria-hidden','true');
}

[menuAddProcess, menuAddTask, menuAssign, menuRecords].forEach(btn=>{
  if(!btn) return;
  btn.addEventListener('click', async ()=>{
    // close all first
    [modalAddProcess, modalAddTask, modalAssign, modalRecords].forEach(m=> closeTemplateModal(m));

    // If the user clicked "Ajustar asignación", ensure we have data to work with
    if(btn.id === 'menu-assign'){
      if(!state.processes || state.processes.length === 0){
        await notify('No se han agregado datos al sistema. Añada procesos y tareas antes de ajustar asignaciones.');
        return;
      }
      const hasAnyTasks = state.processes.some(p=> Array.isArray(p.tasks) && p.tasks.length > 0);
      if(!hasAnyTasks){
        await notify('No hay tareas registradas. Añada tareas antes de ajustar asignaciones.');
        return;
      }
      openTemplateModal(modalAssign);
      return;
    }

    // open corresponding for other menu items
    if(btn.id === 'menu-add-process') openTemplateModal(modalAddProcess);
    if(btn.id === 'menu-add-task') openTemplateModal(modalAddTask);
    if(btn.id === 'menu-records') openTemplateModal(modalRecords);
  });
});

// close buttons inside modals (they carry data-target attribute)
document.querySelectorAll('.modal-close-template').forEach(btn=>{
  btn.addEventListener('click', (e)=>{
    const target = btn.dataset.target;
    const m = document.getElementById(target);
    closeTemplateModal(m);
  });
});
// close when clicking outside modal content
[modalAddProcess, modalAddTask, modalAssign, modalRecords].forEach(m=>{
  if(!m) return;
  m.addEventListener('click', (e)=>{ if(e.target === m) closeTemplateModal(m); });
});

// Show a notification box (OK only)
function notify(message){
  return new Promise((resolve)=>{
    msgText.textContent = message;
    msgNo.style.display = 'none';
    msgYes.textContent = 'OK';
    msgModal.style.display = 'flex';
    msgModal.setAttribute('aria-hidden','false');
    const onOk = ()=>{
      msgYes.removeEventListener('click',onOk);
      msgModal.style.display = 'none';
      msgModal.setAttribute('aria-hidden','true');
      resolve(true);
    };
    msgYes.addEventListener('click',onOk);
  });
}

// Show a decision box (Yes/No) returns Promise<boolean>
function decide(message){
  return new Promise((resolve)=>{
    msgText.textContent = message;
    msgNo.style.display = 'inline-block';
    msgYes.textContent = 'Sí';
    msgNo.textContent = 'No';
    msgModal.style.display = 'flex';
    msgModal.setAttribute('aria-hidden','false');
    const onYes = ()=>{
      cleanup();
      resolve(true);
    };
    const onNo = ()=>{
      cleanup();
      resolve(false);
    };
    function cleanup(){
      msgYes.removeEventListener('click',onYes);
      msgNo.removeEventListener('click',onNo);
      msgModal.style.display = 'none';
      msgModal.setAttribute('aria-hidden','true');
    }
    msgYes.addEventListener('click',onYes);
    msgNo.addEventListener('click',onNo);
  });
}

// Load the detailed sample from modal button
helpDownloadExample.addEventListener('click', async ()=>{
  const ok = await decide('Cargar ejemplo detallado reemplazará las plantillas actuales. Continuar?');
  if(ok) loadSampleData();
  helpModal.style.display='none';
});

// Export report generation helpers
function buildReportText(simResult){
  let txt = `Simulador de Procesos - Informe\n\nResumen:\n`;
  const totalTasks = state.processes.reduce((acc,p)=> acc + p.tasks.reduce((s,t)=>s+t.qty,0),0);
  const totalCost = state.processes.reduce((acc,p)=> acc + p.tasks.reduce((s,t)=>s + t.cost * t.qty,0),0);
  txt += `Total procesos: ${state.processes.length}\nTotal tareas: ${totalTasks}\nCosto total: $${totalCost.toFixed(2)}\n\n`;
  for(const p of state.processes){
    txt += `Proceso: ${p.name}\n Personas: ${p.people} • Meta calidad: ${p.qualityTarget}\n`;
    const procRes = simResult?.[p.id];
    if(procRes){
      txt += `Eficiencia promedio: ${procRes.overallEff.toFixed(1)}% • Calidad promedio: ${procRes.overallQuality.toFixed(1)}%\n`;
    }
    txt += `Tareas:\n`;
    for(const t of p.tasks){
      txt += ` - ${t.name}: qty ${t.qty}, diff ${t.diff}, cost $${t.cost} • asignaciones [${t.assignments.join(', ')}]\n`;
    }
    txt += `\n`;
  }
  txt += `Interpretación: Valores de eficiencia por debajo de 60% requieren revisión de carga/recursos; calidad por debajo de la meta indica necesidad de capacitación o ajuste en dificultad/asignaciones.\n`;
  txt += `Generado: ${new Date().toLocaleString()}\n`;
  return txt;
}

function generateReportPDF(){
  const simResult = simulate();
  // Build PDF with executive style, include header, charts as images and analysis box content
  const doc = new jsPDF({unit:'pt',format:'a4'});
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 40;
  const contentW = pageW - margin*2;
  // Header
  doc.setFillColor(255,122,0);
  doc.rect(0,0,pageW,64,'F');
  doc.setFontSize(18);
  doc.setTextColor(255,255,255);
  doc.text('Simulador de Procesos', margin, 42);
  doc.setFontSize(10);
  doc.setTextColor(255,255,255);
  doc.text(`Generado: ${new Date().toLocaleString()}`, pageW - margin, 42, {align:'right'});
  // Summary block
  doc.setFontSize(11);
  doc.setTextColor(34,34,34);
  const summaryY = 84;
  const totalTasks = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s+t.qty,0),0);
  const totalCost = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s + t.cost*t.qty,0),0);
  doc.text(`Resumen ejecutivo`, margin, summaryY);
  doc.setFontSize(10);
  doc.setTextColor(80);
  doc.text(`Procesos: ${state.processes.length}   •   Tareas totales: ${totalTasks}   •   Costo estimado: $${totalCost.toFixed(2)}`, margin, summaryY + 16);
  // Insert charts (convert Chart.js canvases to images)
  try{
    const imgTasks = chartTasks.toBase64Image();
    const imgEff = chartEff.toBase64Image();
    const imgQual = chartQual.toBase64Image();
    const imgH = 120;
    const gap = 10;
    // Place Tasks and Efficiency side by side
    const colW = (contentW - gap) / 2;
    let y = summaryY + 36;
    doc.addImage(imgTasks, 'PNG', margin, y, colW, imgH);
    doc.addImage(imgEff, 'PNG', margin + colW + gap, y, colW, imgH);
    y += imgH + 12;
    // Place Quality full width
    doc.addImage(imgQual, 'PNG', margin, y, contentW, imgH);
    y += imgH + 18;
    // Analysis box content (use the existing analysis box HTML as plain text)
    const analysisHtml = refs.analysisBox.innerText || refs.analysisBox.textContent || '';
    const analysisLines = doc.splitTextToSize(analysisHtml, contentW);
    doc.setFontSize(11);
    doc.setTextColor(34,34,34);
    doc.text('Análisis interpretativo', margin, y);
    doc.setFontSize(10);
    doc.setTextColor(60);
    doc.text(analysisLines, margin, y + 14);
    // If content exceeds page, jsPDF will automatically add pages if we compute height; do a simple check
    // (Calculate used height and add page if needed for detailed process sections)
    let usedY = y + 14 + (analysisLines.length * 12) + 12;
    if(usedY + 120 > doc.internal.pageSize.getHeight()){
      doc.addPage();
      usedY = margin;
    }
    // Detailed per-process section
    doc.setFontSize(12);
    doc.setTextColor(34,34,34);
    doc.text('Detalles por proceso', margin, usedY);
    let lineY = usedY + 16;
    doc.setFontSize(10);
    for(const p of state.processes){
      if(lineY + 60 > doc.internal.pageSize.getHeight() - margin){
        doc.addPage();
        lineY = margin;
      }
      const procRes = simResult[p.id];
      const effText = procRes ? `${procRes.overallEff.toFixed(1)}%` : 'N/A';
      const qualText = procRes ? `${procRes.overallQuality.toFixed(1)}%` : 'N/A';
      doc.text(`${p.name} — Personas: ${p.people} • Meta calidad: ${p.qualityTarget}`, margin, lineY);
      lineY += 12;
      doc.setFontSize(9);
      doc.setTextColor(90);
      // list tasks compactly
      const taskLines = p.tasks.map(t=>`• ${t.name} (qty ${t.qty}) — diff ${t.diff} — $${t.cost} — asign.: [${t.assignments.join(', ')}]`).join('\n');
      const wrapped = doc.splitTextToSize(taskLines, contentW);
      doc.text(wrapped, margin + 8, lineY);
      lineY += wrapped.length * 10 + 6;
      doc.text(`Eficiencia: ${effText}   •   Calidad: ${qualText}`, margin + 8, lineY);
      lineY += 14;
      doc.setFontSize(10);
      doc.setTextColor(34,34,34);
    }
    // Footer / final notes
    const finalY = doc.internal.pageSize.getHeight() - margin;
    doc.setFontSize(9);
    doc.setTextColor(120);
    doc.text('Informe generado por Simulador de Procesos — Recomendaciones: priorizar entrenamiento y balanceo antes de aumentar personal.', margin, finalY, {baseline:'bottom'});
  }catch(e){
    // Fallback: embed textual report if charts failed
    const content = buildReportText(simResult);
    const lines = doc.splitTextToSize(content, contentW);
    doc.text(lines, margin, 120);
  }

  doc.save(`informe_ejecutivo_${Date.now()}.pdf`);
}

function generateReportWord(){
  const simResult = simulate();
  // try to extract charts as images
  let imgTasks='', imgEff='', imgQual='';
  try{
    imgTasks = chartTasks.toBase64Image();
    imgEff = chartEff.toBase64Image();
    imgQual = chartQual.toBase64Image();
  }catch(e){
    // ignore - will still produce textual report
  }
  // Build executive HTML similar to PDF layout
  const now = new Date().toLocaleString();
  const totalTasks = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s+t.qty,0),0);
  const totalCost = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s + t.cost*t.qty,0),0);

  const analysisHtml = refs.analysisBox ? refs.analysisBox.innerHTML : '';

  let processesHtml = '';
  for(const p of state.processes){
    const procRes = simResult?.[p.id];
    const effText = procRes ? `${procRes.overallEff.toFixed(1)}%` : 'N/A';
    const qualText = procRes ? `${procRes.overallQuality.toFixed(1)}%` : 'N/A';
    const tasksLines = p.tasks.map(t=>`<li>${t.name} — qty ${t.qty} — diff ${t.diff} — $${t.cost} — asign.: [${t.assignments.join(', ')}]</li>`).join('');
    processesHtml += `
      <section style="margin-bottom:12px;">
        <h3 style="margin:6px 0 6px 0;font-size:14px;color:#222">${escapeHtml(p.name)} — Personas: ${p.people} • Meta calidad: ${p.qualityTarget}</h3>
        <div style="font-size:12px;color:#444;margin-bottom:6px">Eficiencia: ${effText} • Calidad: ${qualText}</div>
        <ul style="margin:0 0 6px 18px;padding:0;font-size:12px;color:#333">${tasksLines}</ul>
      </section>
    `;
  }

  const html = `<!doctype html>
  <html>
  <head>
    <meta charset="utf-8">
    <title>Informe - Simulador de Procesos</title>
    <style>
      body{font-family:Calibri,Arial,Helvetica,sans-serif;color:#222;margin:20px;}
      header{background:#ff7a00;color:#fff;padding:14px;border-radius:6px}
      .summary{margin:12px 0;font-size:13px;color:#333}
      .charts{display:flex;gap:10px;flex-wrap:wrap}
      .chart-col{flex:1 1 48%;min-width:220px}
      .full{flex-basis:100%}
      .analysis{background:#fff;padding:10px;border:1px solid #e6e6e6;border-radius:6px;margin-top:12px}
      h1,h2,h3{margin:6px 0}
      pre{font-family:Calibri,Arial}
    </style>
  </head>
  <body>
    <header>
      <h1 style="margin:0;font-size:18px">Simulador de Procesos</h1>
      <div style="font-size:11px;margin-top:6px">Generado: ${now}</div>
    </header>

    <div class="summary">
      <strong>Resumen ejecutivo</strong>
      <div style="margin-top:6px">Procesos: ${state.processes.length} • Tareas totales: ${totalTasks} • Costo estimado: $${totalCost.toFixed(2)}</div>
    </div>

    <div class="charts">
      <div class="chart-col">
        <h2 style="font-size:13px;margin-bottom:6px">Tareas por Proceso</h2>
        ${ imgTasks ? `<img src="${imgTasks}" style="width:100%;max-height:200px;object-fit:contain;border:1px solid #e6e6e6;border-radius:6px">` : '<div style="padding:18px;background:#f7f7f7;border:1px solid #eee;border-radius:6px;color:#666">Gráfica no disponible</div>' }
      </div>
      <div class="chart-col">
        <h2 style="font-size:13px;margin-bottom:6px">Eficiencia por Proceso</h2>
        ${ imgEff ? `<img src="${imgEff}" style="width:100%;max-height:200px;object-fit:contain;border:1px solid #e6e6e6;border-radius:6px">` : '<div style="padding:18px;background:#f7f7f7;border:1px solid #eee;border-radius:6px;color:#666">Gráfica no disponible</div>' }
      </div>
      <div class="chart-col full">
        <h2 style="font-size:13px;margin-bottom:6px">Calidad por Proceso</h2>
        ${ imgQual ? `<img src="${imgQual}" style="width:100%;max-height:200px;object-fit:contain;border:1px solid #e6e6e6;border-radius:6px">` : '<div style="padding:18px;background:#f7f7f7;border:1px solid #eee;border-radius:6px;color:#666">Gráfica no disponible</div>' }
      </div>
    </div>

    <div class="analysis">
      <h2 style="font-size:14px;margin:0 0 8px 0">Análisis interpretativo</h2>
      ${analysisHtml ? analysisHtml : '<div style="color:#666">Ejecute la simulación para obtener el análisis interpretativo aquí.</div>'}
    </div>

    <main style="margin-top:14px">
      <h2 style="font-size:14px;margin-bottom:8px">Detalles por proceso</h2>
      ${processesHtml}
    </main>

    <footer style="margin-top:18px;font-size:11px;color:#666">
      Informe generado por Simulador de Procesos — Recomendaciones: priorizar formación y balanceo antes de aumentar personal.
    </footer>
  </body>
  </html>`;

  // Create blob and trigger download as .doc for Word compatibility
  const blob = new Blob([html], {type: "application/msword"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `informe_simulador_${Date.now()}.doc`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(str){
  return str.replace(/[&<>"]/g, function(tag){ const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }; return map[tag] || tag; });
}

// Charts
let chartTasks, chartEff, chartQual;
function createCharts(){
  chartTasks = new Chart(refs.chartTasksCtx, {
    type: 'bar',
    data: {labels:[],datasets:[{label:'Tareas',data:[],backgroundColor:'#2b7a78'}]},
    options:{responsive:true,maintainAspectRatio:false}
  });
  chartEff = new Chart(refs.chartEffCtx, {
    type: 'bar',
    data: {labels:[],datasets:[{label:'Eficiencia (%)',data:[],backgroundColor:'#ffb703'}]},
    options:{responsive:true,maintainAspectRatio:false,scales:{y:{max:100,min:0}}}
  });
  chartQual = new Chart(refs.chartQualCtx, {
    type: 'bar',
    data: {labels:[],datasets:[{label:'Calidad (%)',data:[],backgroundColor:'#06d6a0'}]},
    options:{responsive:true,maintainAspectRatio:false,scales:{y:{max:100,min:0}}}
  });
}

// State management & rendering
function addProcess(name, people, qualityTarget){
  const proc = {id: uid('proc'), name, people: Number(people), qualityTarget: Number(qualityTarget), tasks: []};
  state.processes.push(proc);
  renderAll();
}

function addTaskToProcess(procId, name, diff, cost, qty){
  const proc = state.processes.find(p=>p.id===procId);
  if(!proc) return;
  const task = {id: uid('task'), name, diff:Number(diff), cost: Number(cost), qty: Number(qty), assignments: Array.from({length:proc.people},()=>0)};
  proc.tasks.push(task);
  renderAll();
}

function clearAll(){state.processes = []; renderAll();}

// Render helper functions
function populateSelects(){
  const fills = state.processes.map(p=>`<option value="${p.id}">${p.name}</option>`).join('');
  refs.taskProcess.innerHTML = fills;
  refs.assignProcess.innerHTML = fills;
  // process select for editing: prepend empty option
  refs.procSelect.innerHTML = `<option value="">-- Nuevo proceso --</option>${fills}`;
}

// When user chooses a process to edit, load values into form
refs.procSelect.addEventListener('change', ()=>{
  const pid = refs.procSelect.value;
  if(!pid){
    refs.procName.value = '';
    refs.procPeople.value = 1;
    refs.procQualityTarget.value = 90;
    refs.updateProcessBtn.disabled = true;
    return;
  }
  const p = state.processes.find(x=>x.id===pid);
  if(p){
    refs.procName.value = p.name;
    refs.procPeople.value = p.people;
    refs.procQualityTarget.value = p.qualityTarget;
    refs.updateProcessBtn.disabled = false;
  }
});

// Add process remains create-only
refs.addProcessBtn.addEventListener('click', ()=>{
  const name = refs.procName.value.trim();
  const people = refs.procPeople.value;
  const qt = refs.procQualityTarget.value;
  if(!name){ notify('Ingrese nombre de proceso'); return; }
  addProcess(name, people, qt);
  refs.procName.value = '';
  refs.procPeople.value = 1;
  refs.procQualityTarget.value = 90;
  // reset edit select
  refs.procSelect.value = '';
  refs.updateProcessBtn.disabled = true;
});

// Update existing process in-place
refs.updateProcessBtn.addEventListener('click', async ()=>{
  const pid = refs.procSelect.value;
  if(!pid){ await notify('Seleccione un proceso a modificar'); return; }
  const p = state.processes.find(x=>x.id===pid);
  if(!p){ await notify('Proceso no encontrado'); return; }
  const ok = await decide('Guardar cambios en el proceso seleccionado?');
  if(!ok) return;
  p.name = refs.procName.value.trim() || p.name;
  p.people = Number(refs.procPeople.value) || p.people;
  p.qualityTarget = Number(refs.procQualityTarget.value) || p.qualityTarget;
  // Ensure each task assignments length equals new people count
  for(const t of p.tasks){
    if(!Array.isArray(t.assignments)) t.assignments = [];
    if(t.assignments.length !== p.people){
      const newArr = Array.from({length:p.people},(v,i)=> Number(t.assignments[i]) || 0);
      t.assignments = newArr;
    }
  }
  renderAll();
  await notify('Proceso modificado correctamente');
  // Keep selection on updated process
  refs.procSelect.value = p.id;
});

// initialize update button disabled by default
refs.updateProcessBtn.disabled = true;

function renderAssignPersons(){
  const procId = refs.assignProcess.value;
  const proc = state.processes.find(p=>p.id===procId);
  refs.assignPersons.innerHTML = '';

  // Update label next to the select to include total tasks for the selected process
  try{
    const labelEl = refs.assignProcess.parentElement;
    if(proc){
      const totalTasks = proc.tasks.reduce((s,t)=>s + (Number(t.qty)||0), 0);
      // The label contains a text node before the select; update that node's text
      if(labelEl.childNodes && labelEl.childNodes[0] && labelEl.childNodes[0].nodeType === Node.TEXT_NODE){
        labelEl.childNodes[0].textContent = `Proceso — Total tareas: ${totalTasks} `;
      } else {
        // fallback: prepend a text node
        labelEl.insertBefore(document.createTextNode(`Proceso — Total tareas: ${totalTasks} `), refs.assignProcess);
      }
    } else {
      if(labelEl.childNodes && labelEl.childNodes[0] && labelEl.childNodes[0].nodeType === Node.TEXT_NODE){
        labelEl.childNodes[0].textContent = 'Proceso ';
      }
    }
  }catch(e){
    // ignore label update errors
    console.error(e);
  }

  if(!proc){
    refs.assignPersons.innerHTML = '<div class="muted">Seleccione un proceso con tareas</div>'; return;
  }
  if(proc.tasks.length===0){ refs.assignPersons.innerHTML = '<div class="muted">No hay tareas para este proceso</div>'; return; }
  // For each task in process show controls to allocate per person and a live total
  for(const t of proc.tasks){
    const div = document.createElement('div');
    div.className = 'task-assign';
    // add a span to show live total so it updates as user types
    div.innerHTML = `<strong>${t.name}</strong> <span class="muted" style="margin-left:8px">total: <span class="task-total" data-task="${t.id}">${t.qty}</span></span><div class="muted">Asigne tareas por persona</div>`;
    const personsCont = document.createElement('div');
    personsCont.style.marginTop='6px';
    personsCont.className = 'persons-grid';
    for(let i=0;i<proc.people;i++){
      const pdiv = document.createElement('div');
      pdiv.className = 'assign-person';
      pdiv.innerHTML = `<label>Persona ${i+1}<input type="number" min="0" value="${t.assignments[i]||0}" data-task="${t.id}" data-person="${i}"></label>`;
      personsCont.appendChild(pdiv);
    }
    div.appendChild(personsCont);
    refs.assignPersons.appendChild(div);
  }

  // delegate input events to update live totals as user changes values
  refs.assignPersons.querySelectorAll('input[type="number"]').forEach(inp=>{
    inp.addEventListener('input', (e)=>{
      const taskId = inp.dataset.task;
      const inputs = refs.assignPersons.querySelectorAll(`input[data-task="${taskId}"]`);
      let sum = 0;
      inputs.forEach(i=> sum += Number(i.value) || 0);
      const totalEl = refs.assignPersons.querySelector(`.task-total[data-task="${taskId}"]`);
      if(totalEl) totalEl.textContent = sum;
      // Also update overall total in the label when any assignment changes
      try{
        const labelEl = refs.assignProcess.parentElement;
        const procLocal = state.processes.find(p=>p.id===refs.assignProcess.value);
        if(procLocal){
          const overall = procLocal.tasks.reduce((s,t)=>{
            // compute current visible totals: if any task has assignments entered, prefer sum of inputs,
            // otherwise use t.qty
            const inputsAll = refs.assignPersons.querySelectorAll(`input[data-task="${t.id}"]`);
            if(inputsAll && inputsAll.length>0){
              const ssum = Array.from(inputsAll).reduce((a,ii)=>a + (Number(ii.value)||0),0);
              return s + ssum;
            }
            return s + (Number(t.qty)||0);
          }, 0);
          if(labelEl.childNodes && labelEl.childNodes[0] && labelEl.childNodes[0].nodeType === Node.TEXT_NODE){
            labelEl.childNodes[0].textContent = `Proceso — Total tareas: ${overall} `;
          }
        }
      }catch(e){
        // ignore
      }
    });
  });
}

async function saveAssignments(){
  const procId = refs.assignProcess.value;
  const proc = state.processes.find(p=>p.id===procId);
  if(!proc) return;
  const inputs = refs.assignPersons.querySelectorAll('input[type="number"]');
  // apply values and also update task.qty as the sum of assignments (so totals and records reflect changes)
  // We'll group inputs by task id
  const groups = {};
  for(const inp of inputs){
    const tid = inp.dataset.task;
    const personIndex = Number(inp.dataset.person);
    if(!groups[tid]) groups[tid] = [];
    groups[tid][personIndex] = Number(inp.value) || 0;
  }
  for(const tid of Object.keys(groups)){
    const task = proc.tasks.find(t=>t.id===tid);
    if(!task) continue;
    const assigns = groups[tid].map(n=>Number(n)||0);
    // store assignments normalized to process people length
    task.assignments = Array.from({length: proc.people}, (v,i)=> assigns[i] || 0);
    // IMPORTANT: update task.qty to be the sum of assignments so totals and simulation use new totals
    const sumAssigned = task.assignments.reduce((a,b)=>a+b,0);
    task.qty = sumAssigned;
  }
  renderAll();
  await notify('Asignaciones guardadas');
}

function renderRecords(){
  refs.records.innerHTML = '';
  for(const p of state.processes){
    const div = document.createElement('div');
    div.className = 'process-row';
    const totals = p.tasks.reduce((acc,t)=>{ acc.qty += t.qty; acc.cost += t.cost*t.qty; return acc },{qty:0,cost:0});
    div.innerHTML = `<div><strong>${p.name}</strong><div class="muted">Tareas:${totals.qty} • Personas:${p.people}</div></div><div style="text-align:right">$${totals.cost.toFixed(2)}</div>`;
    refs.records.appendChild(div);
  }
}

function renderPersonDetails(simResult){
  refs.personDetails.innerHTML = '';
  // Build compact table-like view with precise per-person task counts derived from assignments (or fair fallback)
  for(const p of state.processes){
    const heading = document.createElement('div');
    heading.innerHTML = `<strong>${p.name}</strong>`;
    refs.personDetails.appendChild(heading);

    const list = document.createElement('div');
    list.style.marginBottom = '8px';

    // initialize per-person counters
    const perPerson = Array.from({length: p.people}, () => ({tasks: 0, cost: 0}));

    // For each task, use assignments if present; otherwise distribute qty evenly
    for(const t of p.tasks){
      let assigns = Array.isArray(t.assignments) ? t.assignments.map(n => Number(n) || 0) : [];
      const sumAssigned = assigns.reduce((a,b) => a + b, 0);
      if(sumAssigned === 0){
        // even distribution fallback
        const base = Math.floor(t.qty / p.people);
        assigns = Array.from({length: p.people}, (_,i) => i < (t.qty % p.people) ? base + 1 : base);
      } else if(sumAssigned !== t.qty && t.qty > 0){
        // scale proportionally to match declared qty, then correct rounding
        const factor = t.qty / Math.max(1, sumAssigned);
        assigns = assigns.map(v => Math.round(v * factor));
        let diff = t.qty - assigns.reduce((a,b)=>a+b,0);
        for(let i=0; i<Math.abs(diff); i++){
          const idx = i % p.people;
          assigns[idx] += Math.sign(diff);
        }
      } else if(t.qty === 0 && sumAssigned > 0){
        // If qty is zero but assignments present, set qty = sumAssigned to reflect reality
        t.qty = sumAssigned;
      }

      for(let i=0;i<p.people;i++){
        const n = assigns[i] || 0;
        perPerson[i].tasks += n;
        perPerson[i].cost += n * t.cost;
      }
    }

    // Render per-person rows (more compact spacing)
    perPerson.forEach((pp, idx) => {
      const el = document.createElement('div');
      el.style.display = 'flex';
      el.style.justifyContent = 'space-between';
      el.style.padding = '4px 0';
      el.style.borderBottom = '1px dashed #eee';
      el.style.fontSize = '13px';
      const sim = simResult?.[p.id]?.people?.[idx];
      const right = sim ? `E:${sim.eff.toFixed(0)}% • Q:${sim.quality.toFixed(0)}%` : '';
      el.innerHTML = `<div>Persona ${idx+1} • T:${pp.tasks} • $${pp.cost.toFixed(2)}</div><div>${right}</div>`;
      list.appendChild(el);
    });

    refs.personDetails.appendChild(list);
  }
}

// New: render analysis box interpreting results
function renderAnalysis(simResult){
  const box = refs.analysisBox;
  if(!box) return;
  if(!simResult || state.processes.length===0){
    box.innerHTML = `<strong>Análisis</strong><div class="muted">Ejecute la simulación para ver el análisis interpretativo aquí.</div>`;
    return;
  }

  const lines = [];
  lines.push('<strong>Análisis interpretativo</strong>');
  for(const p of state.processes){
    const res = simResult[p.id];
    if(!res){ lines.push(`<div><strong>${p.name}:</strong> Sin datos de simulación.</div>`); continue; }
    const avgEff = res.overallEff;
    const avgQual = res.overallQuality;
    const totalTasks = res.people.reduce((s,pp)=>s+pp.tasks,0);
    // Determine suggestions
    let suggestion = 'OK';
    if(avgQual < p.qualityTarget - 5) suggestion = 'Calidad por debajo del objetivo: considere capacitación, revisar asignaciones o disminuir dificultad.';
    else if(avgEff < 60) suggestion = 'Baja eficiencia: redistribuir tareas o aumentar personal.';
    else if(totalTasks > p.people * 40) suggestion = 'Alta carga por persona: evaluar aumento de recursos o balanceo.';
    else suggestion = 'Indicadores dentro de rango esperado.';
    lines.push(`<div style="margin-top:8px"><strong>${p.name}:</strong> Tareas totales ${totalTasks} • Eficiencia ${avgEff.toFixed(1)}% • Calidad ${avgQual.toFixed(1)}%<br><em>Interpretación:</em> ${suggestion}</div>`);
  }

  // Overall summary
  const overallTasks = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s+t.qty,0),0);
  const overallCost = state.processes.reduce((a,p)=>a + p.tasks.reduce((s,t)=>s + t.cost*t.qty,0),0);
  lines.push(`<div style="margin-top:12px"><strong>Resumen global:</strong> Tareas ${overallTasks} • Costo $${overallCost.toFixed(2)}.</div>`);
  lines.push(`<div class="muted" style="margin-top:8px">Consejo: si varios procesos muestran eficiencia baja o calidad por debajo de la meta, priorice redistribución de tareas y formación antes de aumentar personal.</div>`);

  box.innerHTML = lines.join('');
}

// Aggregates for charts
function computeAggregates(simResult){
  const labels = state.processes.map(p=>p.name);
  const tasksData = state.processes.map(p=> p.tasks.reduce((s,t)=>s+t.qty,0) );
  const effData = state.processes.map(p=> {
    const sr = simResult?.[p.id];
    return sr ? sr.overallEff : 0;
  });
  const qualData = state.processes.map(p=> {
    const sr = simResult?.[p.id];
    return sr ? sr.overallQuality : 0;
  });

  chartTasks.data.labels = labels; chartTasks.data.datasets[0].data = tasksData; chartTasks.update();
  chartEff.data.labels = labels; chartEff.data.datasets[0].data = effData; chartEff.update();
  chartQual.data.labels = labels; chartQual.data.datasets[0].data = qualData; chartQual.update();

  const totalTasks = tasksData.reduce((a,b)=>a+b,0);
  const totalCost = state.processes.reduce((acc,p)=> acc + p.tasks.reduce((s,t)=>s + t.cost * t.qty,0),0);
  refs.sumTasks.textContent = totalTasks;
  refs.sumCost.textContent = totalCost.toFixed(2);
}

// Simulation engine (simple, transparent, fast)
function simulate(){
  // returns a result per process with per-person metrics
  const noise = Number(refs.simNoise.value);
  const speed = Number(refs.simSpeed.value);

  const result = {};
  for(const p of state.processes){
    const procRes = {people:[],overallEff:0,overallQuality:0};
    // initialize per person
    for(let i=0;i<p.people;i++) procRes.people[i] = {tasks:0,cost:0,eff:0,quality:0};
    // For each task, distribute assigned tasks; compute time and quality impact
    for(const t of p.tasks){
      // if assignments all zero: fallback equal distribution
      const sumAssigned = t.assignments.reduce((a,b)=>a+b,0);
      let assign = t.assignments.slice();
      if(sumAssigned === 0){
        const per = Math.floor(t.qty / p.people);
        assign = Array.from({length:p.people},(_,i)=> i < (t.qty % p.people) ? per+1 : per);
      } else {
        // if assigned sum differs from qty, scale proportionally
        const factor = t.qty / Math.max(1,sumAssigned);
        assign = assign.map(v=>Math.round(v*factor));
        // correct rounding to match total qty
        let diff = t.qty - assign.reduce((a,b)=>a+b,0);
        for(let i=0;i<Math.abs(diff);i++){
          const idx = i % p.people;
          assign[idx] += Math.sign(diff);
        }
      }

      for(let i=0;i<p.people;i++){
        const qty = assign[i] || 0;
        procRes.people[i].tasks += qty;
        procRes.people[i].cost += qty * t.cost;
        // Baseline time per task proportional to difficulty
        const baseTime = t.diff * 1.0;
        // Efficiency decreases with higher difficulty and random noise
        const speedFactor = 1 / (1 + t.diff*0.08);
        const randomFactor = 1 - (Math.random()*noise*0.5);
        const eff = Math.max(20, Math.min(110, 80 * speedFactor * randomFactor * (1 + (speed-1)*0.2)));
        // Quality influenced by difficulty, targets and noise
        const qualityDrop = t.diff * 2;
        const quality = Math.max(40, Math.min(100, p.qualityTarget - qualityDrop + (Math.random()*noise*30)));
        // aggregate (weighted by qty)
        procRes.people[i].eff += eff * qty;
        procRes.people[i].quality += quality * qty;
      }
    }

    // finalize per-person averages
    let overallSumEff=0, overallSumQual=0, totalTasks=0;
    for(let i=0;i<p.people;i++){
      const pp = procRes.people[i];
      if(pp.tasks>0){
        pp.eff = pp.eff / pp.tasks;
        pp.quality = pp.quality / pp.tasks;
      } else { pp.eff = 0; pp.quality = 0; }
      overallSumEff += pp.eff * pp.tasks;
      overallSumQual += pp.quality * pp.tasks;
      totalTasks += pp.tasks;
    }
    procRes.overallEff = totalTasks ? (overallSumEff / totalTasks) : 0;
    procRes.overallQuality = totalTasks ? (overallSumQual / totalTasks) : 0;
    result[p.id] = procRes;
  }

  return result;
}

// Main render — ensure a simulation result is computed when none is provided so analysis and charts match current data
function renderAll(simResult){
  const sim = simResult || (state.processes.length ? simulate() : {});
  populateSelects();
  renderRecords();
  renderAssignPersons();
  renderPersonDetails(sim);
  renderAnalysis(sim);
  computeAggregates(sim);
}

// Event wiring
refs.addTaskBtn.addEventListener('click', ()=>{
  const pid = refs.taskProcess.value;
  const name = refs.taskName.value.trim();
  const diff = refs.taskDiff.value;
  const cost = refs.taskCost.value;
  const qty = refs.taskQty.value;
  if(!pid){ notify('Seleccione proceso'); return; }
  if(!name){ notify('Ingrese nombre de tarea'); return; }
  addTaskToProcess(pid, name, diff, cost, qty);
  refs.taskName.value=''; refs.taskQty.value='1';
});

refs.assignProcess.addEventListener('change', renderAssignPersons);
refs.saveAssign.addEventListener('click', saveAssignments);

refs.runSim.addEventListener('click', async ()=>{
  // Run simulation, refresh all views and show short status message
  const simResult = simulate();
  renderAll(simResult);
  renderPersonDetails(simResult);

  // Insert a transient status line at top of records area
  try{
    const nowLabel = `Simulación ejecutada: ${new Date().toLocaleString()}`;
    // remove any previous status
    const prev = document.getElementById('sim-status');
    if(prev) prev.remove();
    const node = document.createElement('div');
    node.id = 'sim-status';
    node.style.fontSize = '13px';
    node.style.color = '#1f7a78';
    node.style.padding = '6px 8px';
    node.style.borderBottom = '1px dashed #e6e6e6';
    node.textContent = nowLabel;
    // place at the top of the records container
    refs.records.insertAdjacentElement('afterbegin', node);
    // auto-remove after 4 seconds
    setTimeout(()=>{ const el = document.getElementById('sim-status'); if(el) el.remove(); }, 4000);
  }catch(e){
    // fallback: no-op
    console.error(e);
  }

  // Also show a confirmation dialog-like toast to the user
  await notify('Simulación ejecutada');
});

refs.resetBtn.addEventListener('click', async ()=>{
  const ok = await decide('Restablecer todo?');
  if(ok) clearAll();
});

// Sample data loader
function loadSampleData(){
  // Build sample processes but normalize assignments to ensure they match people count
  state.processes = [];
  const makeProc = (name, people, qualityTarget, tasks) => {
    const p = {id:uid('proc'), name, people: Number(people), qualityTarget: Number(qualityTarget), tasks: []};
    for(const t of tasks){
      // ensure assignments length equals people and sum equals qty (distribute proportionally if needed)
      let assigns = Array.isArray(t.assignments) ? t.assignments.map(n=>Number(n)||0) : [];
      // if assignments length mismatch, create equal distribution
      if(assigns.length !== p.people){
        const base = Math.floor(t.qty / p.people);
        assigns = Array.from({length:p.people},(_,i)=> i < (t.qty % p.people) ? base+1 : base);
      } else {
        // scale to match qty if sums differ
        const sumAssigned = assigns.reduce((a,b)=>a+b,0);
        if(sumAssigned !== t.qty && sumAssigned > 0){
          const factor = t.qty / sumAssigned;
          assigns = assigns.map(v=>Math.round(v * factor));
          // correct any rounding diff
          let diff = t.qty - assigns.reduce((a,b)=>a+b,0);
          for(let i=0;i<Math.abs(diff);i++){
            const idx = i % p.people;
            assigns[idx] += Math.sign(diff);
          }
        } else if(sumAssigned === 0){
          const base = Math.floor(t.qty / p.people);
          assigns = Array.from({length:p.people},(_,i)=> i < (t.qty % p.people) ? base+1 : base);
        }
      }
      p.tasks.push({
        id: uid('task'),
        name: t.name,
        diff: Number(t.diff) || 1,
        cost: Number(t.cost) || 0,
        qty: Number(t.qty) || 0,
        assignments: assigns
      });
    }
    return p;
  };

  const p1 = makeProc('Recepción', 3, 92, [
    {name:'Verificar pedido', diff:3, cost:4.5, qty:30, assignments:[10,10,10]},
    {name:'Registrar entrada', diff:2, cost:2.0, qty:20, assignments:[7,7,6]},
    {name:'Inspección inicial', diff:4, cost:6.0, qty:15, assignments:[5,5,5]}
  ]);

  const p2 = makeProc('Procesamiento', 4, 88, [
    {name:'Revisión técnica', diff:6, cost:12.0, qty:40, assignments:[10,10,10,10]},
    {name:'Ajustes', diff:5, cost:8.5, qty:25, assignments:[7,6,6,6]},
    {name:'Pruebas funcionales', diff:7, cost:15.0, qty:18, assignments:[5,4,5,4]}
  ]);

  const p3 = makeProc('Despacho', 2, 95, [
    {name:'Empaquetado', diff:3, cost:3.5, qty:35, assignments:[18,17]},
    {name:'Generar guía', diff:2, cost:1.5, qty:35, assignments:[18,17]},
    {name:'Control de salida', diff:4, cost:2.5, qty:20, assignments:[10,10]}
  ]);

  state.processes.push(p1, p2, p3);
  renderAll();
}

// Reset templates to zero (clears processes without confirmation)
function resetTemplates(){
  state.processes = [];
  renderAll();
}

// Wire new buttons
refs.loadSampleBtn.addEventListener('click', async ()=>{
  const ok = await decide('Cargar datos de ejemplo completos reemplazará las plantillas actuales. Continuar?');
  if(ok) loadSampleData();
});
refs.resetTemplatesBtn.addEventListener('click', async ()=>{
  const ok = await decide('Reiniciar las plantillas a cero eliminará las plantillas actuales. Continuar?');
  if(ok) resetTemplates();
});

// Wire export buttons
refs.exportPdf.addEventListener('click', async ()=>{
  const ok = await decide('Generar informe en PDF?');
  if(ok) generateReportPDF();
});
refs.exportWord.addEventListener('click', async ()=>{
  const ok = await decide('Generar respaldo JSON?');
  if(ok) backupDataJSON();
});

// New: hidden file input for loading backups
const fileInput = document.createElement('input');
fileInput.type = 'file';
fileInput.accept = 'application/json';
fileInput.style.display = 'none';
document.body.appendChild(fileInput);

refs.loadBackupBtn.addEventListener('click', async ()=>{
  const ok = await decide('Cargar un archivo de respaldo reemplazará el estado actual. Continuar?');
  if(!ok) return;
  fileInput.value = '';
  fileInput.click();
});

fileInput.addEventListener('change', async (ev)=>{
  const f = ev.target.files && ev.target.files[0];
  if(!f) return;
  try{
    const text = await f.text();
    const data = JSON.parse(text);
    // Basic validation: expect object with state.processes
    if(!data || !data.state || !Array.isArray(data.state.processes)){
      await notify('Archivo no válido: no contiene datos de respaldo esperados.');
      return;
    }
    // Restore minimal state safely (keep nextId generation with new ids)
    // Convert/process to expected shapes: ensure assignments arrays exist for tasks
    const loaded = data.state;
    // normalize
    for(const p of loaded.processes){
      p.people = Number(p.people) || 1;
      p.qualityTarget = Number(p.qualityTarget) || 90;
      p.tasks = Array.isArray(p.tasks) ? p.tasks : [];
      for(const t of p.tasks){
        t.qty = Number(t.qty) || 0;
        t.diff = Number(t.diff) || 1;
        t.cost = Number(t.cost) || 0;
        if(!Array.isArray(t.assignments) || t.assignments.length !== p.people){
          t.assignments = Array.from({length:p.people},()=>0);
        } else {
          // coerce to numbers
          t.assignments = t.assignments.map(x=>Number(x)||0);
        }
      }
    }
    // replace state
    state.processes = loaded.processes;
    // ensure unique ids (if missing, add)
    for(const p of state.processes){
      if(!p.id) p.id = uid('proc');
      for(const t of p.tasks){
        if(!t.id) t.id = uid('task');
      }
    }
    renderAll();
    await notify('Respaldo cargado correctamente.');
  }catch(err){
    console.error(err);
    await notify('Error al leer el archivo de respaldo. Asegúrese de que sea JSON válido.');
  }
});

// Initial setup
createCharts();
renderAll();

 // New: backup current state as JSON and trigger download
function backupDataJSON(){
  const payload = {
    metadata: {
      generatedAt: new Date().toISOString(),
      processesCount: state.processes.length
    },
    state: state,
    analysis: refs.analysisBox ? refs.analysisBox.innerText : ''
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {type: 'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  // build timestamp safe for filenames: YYYYMMDD_HHMMSS
  const now = new Date();
  const pad = (n)=>String(n).padStart(2,'0');
  const ts = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  a.download = `SimuladorProcesos_${ts}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---- NEW: Task management modal behaviors ---- */

// build table rows for all tasks
function buildTasksTable(){
  const tbody = refs.tasksTableBody;
  tbody.innerHTML = '';
  for(const p of state.processes){
    for(const t of p.tasks){
      const tr = document.createElement('tr');
      tr.dataset.procId = p.id;
      tr.dataset.taskId = t.id;
      // assignments displayed as comma-separated
      const assigns = (Array.isArray(t.assignments) ? t.assignments.join(', ') : '');
      tr.innerHTML = `
        <td>${escapeHtml(p.name)}</td>
        <td><input type="text" data-field="name" value="${escapeHtml(t.name)}"></td>
        <td><input type="number" min="1" max="10" data-field="diff" value="${t.diff}"></td>
        <td><input type="number" min="0" step="0.01" data-field="cost" value="${t.cost}"></td>
        <td><input type="number" min="0" data-field="qty" value="${t.qty}"></td>
        <td><input type="text" data-field="assignments" value="${assigns}" class="assign-text"></td>
        <td>
          <button class="icon-btn save-row" title="Guardar">💾</button>
          <button class="icon-btn edit-row" title="Editar">✏️</button>
          <button class="icon-btn del-row" title="Eliminar">🗑️</button>
        </td>
      `;
      // wire per-row buttons
      tbody.appendChild(tr);
    }
  }
}

// open modal
function openTasksModal(){
  buildTasksTable();
  refs.tasksModal.style.display = 'flex';
  refs.tasksModal.setAttribute('aria-hidden','false');
}

// close modal
function closeTasksModal(){
  refs.tasksModal.style.display = 'none';
  refs.tasksModal.setAttribute('aria-hidden','true');
}

// parse assignments field "n1, n2, n3" into array with process people length
function parseAssignmentsField(procId, str){
  const proc = state.processes.find(p=>p.id===procId);
  const parts = (''+str).split(',').map(s=>Number(s.trim()||0));
  // normalize length to process.people
  if(!proc) return parts;
  const arr = Array.from({length:proc.people},(v,i)=> Number(parts[i]) || 0);
  return arr;
}

// save a single row changes back to state
async function saveRow(tr){
  const pid = tr.dataset.procId;
  const tid = tr.dataset.taskId;
  const proc = state.processes.find(p=>p.id===pid);
  if(!proc) return;
  const task = proc.tasks.find(x=>x.id===tid);
  if(!task) return;
  const name = tr.querySelector('[data-field="name"]').value.trim();
  const diff = Number(tr.querySelector('[data-field="diff"]').value) || 1;
  const cost = Number(tr.querySelector('[data-field="cost"]').value) || 0;
  const qty = Number(tr.querySelector('[data-field="qty"]').value) || 0;
  const assignsRaw = tr.querySelector('[data-field="assignments"]').value;
  const assignments = parseAssignmentsField(pid, assignsRaw);
  task.name = name || task.name;
  task.diff = diff;
  task.cost = cost;
  task.qty = qty;
  task.assignments = assignments;
  renderAll();
  await notify('Tarea actualizada');
  // refresh table to reflect normalized assignments
  buildTasksTable();
}

// delete a task row
async function deleteRow(tr){
  const pid = tr.dataset.procId;
  const tid = tr.dataset.taskId;
  const proc = state.processes.find(p=>p.id===pid);
  if(!proc) return;
  const ok = await decide('Eliminar esta tarea?');
  if(!ok) return;
  proc.tasks = proc.tasks.filter(t=>t.id !== tid);
  renderAll();
  buildTasksTable();
}

// wire manager button
refs.manageTasksBtn.addEventListener('click', async ()=>{
  // ensure there is at least one task
  const hasTasks = state.processes.some(p=>p.tasks && p.tasks.length>0);
  if(!hasTasks){ await notify('No hay tareas agregadas.'); return; }
  openTasksModal();
});

// close buttons
refs.tasksClose.addEventListener('click', closeTasksModal);
refs.tasksClose2.addEventListener('click', closeTasksModal);

// delegate table clicks
refs.tasksTableBody.addEventListener('click', async (e)=>{
  const btn = e.target.closest('button');
  if(!btn) return;
  const tr = btn.closest('tr');
  if(btn.classList.contains('save-row')){
    await saveRow(tr);
  } else if(btn.classList.contains('del-row')){
    await deleteRow(tr);
  } else if(btn.classList.contains('edit-row')){
    // focus first input in row
    const first = tr.querySelector('input');
    if(first) first.focus();
  }
});