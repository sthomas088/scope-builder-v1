const state = {
  disciplines: [],
  selectedTaskIds: new Set(),
  currentDocumentMode: 'preview',
  previousEditorContent: '',
};

function createDefaultEditorHtml() {
  return `<p class="document-placeholder">Select tasks and review the live preview. Then click "Send to Editor".</p>`;
}

const els = {
  discipline: document.getElementById('discipline'),
  coverLetterType: document.getElementById('coverLetterType'),
  previewModeBtn: document.getElementById('previewModeBtn'),
  editorModeBtn: document.getElementById('editorModeBtn'),
  sendPreviewToEditorBtn: document.getElementById('sendPreviewToEditorBtn'),
  restorePreviousEditorBtn: document.getElementById('restorePreviousEditorBtn'),
  documentModeMessage: document.getElementById('documentModeMessage'),
  previewDocumentPanel: document.getElementById('previewDocumentPanel'),
  editorDocumentPanel: document.getElementById('editorDocumentPanel'),
  previewDocument: document.getElementById('previewDocument'),
  taskCategoryLabel: document.getElementById('taskCategoryLabel'),
  taskList: document.getElementById('taskList'),
  generateBtn: document.getElementById('generateBtn'),
  exportBtn: document.getElementById('exportBtn'),
  scopeEditor: document.getElementById('scopeEditor'),
  projectName: document.getElementById('projectName'),
  projectAddress: document.getElementById('projectAddress'),
  projectCity: document.getElementById('projectCity'),
  projectCounty: document.getElementById('projectCounty'),
  customProjectCountyWrap: document.getElementById('customProjectCountyWrap'),
  customProjectCounty: document.getElementById('customProjectCounty'),
  projectState: document.getElementById('projectState'),
  regionalFramework: document.getElementById('regionalFramework'),
  customRegionalFrameworkWrap: document.getElementById('customRegionalFrameworkWrap'),
  customRegionalFramework: document.getElementById('customRegionalFramework'),
  acreageWrap: document.getElementById('acreageWrap'),
  acreage: document.getElementById('acreage'),
  apnWrap: document.getElementById('apnWrap'),
  apn: document.getElementById('apn'),
  generalProjectUnderstandingWrap: document.getElementById('generalProjectUnderstandingWrap'),
  generalProjectUnderstanding: document.getElementById('generalProjectUnderstanding'),
  client: document.getElementById('client'),
  contactName: document.getElementById('contactName'),
  contactEmail: document.getElementById('contactEmail'),
  clientAddress: document.getElementById('clientAddress'),
  subjectLine: document.getElementById('subjectLine'),
  scopeDate: document.getElementById('scopeDate'),
  signatory1Name: document.getElementById('signatory1Name'),
  signatory1Title: document.getElementById('signatory1Title'),
  signatory2Name: document.getElementById('signatory2Name'),
  signatory2Title: document.getElementById('signatory2Title'),
};

function todayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function sanitizeForFilename(input) {
  return input.trim().replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ');
}

function getEditorHtml() {
  return els.scopeEditor.innerHTML.trim();
}

function editorHasExportableContent() {
  const html = getEditorHtml();
  if (!html) return false;
  const text = els.scopeEditor.innerText.trim();
  return !!text;
}

function getCountyValue() {
  return els.projectCounty.value === 'Custom' ? els.customProjectCounty.value.trim() : els.projectCounty.value;
}

function getRegionalFrameworkValue() {
  return els.regionalFramework.value === 'Custom'
    ? els.customRegionalFramework.value.trim()
    : els.regionalFramework.value;
}

function updateCountyInputVisibility() {
  const isCustom = els.projectCounty.value === 'Custom';
  els.customProjectCountyWrap.classList.toggle('hidden', !isCustom);

  if (isCustom) {
    els.customProjectCounty.setAttribute('required', 'required');
  } else {
    els.customProjectCounty.removeAttribute('required');
    els.customProjectCounty.value = '';
  }
}

function updateRegionalFrameworkVisibility() {
  const isCustom = els.regionalFramework.value === 'Custom';
  els.customRegionalFrameworkWrap.classList.toggle('hidden', !isCustom);

  if (isCustom) {
    els.customRegionalFramework.setAttribute('required', 'required');
  } else {
    els.customRegionalFramework.removeAttribute('required');
    els.customRegionalFramework.value = '';
  }
}

function updateCoverLetterTypeVisibility() {
  const type = els.coverLetterType.value;
  const showAcreage = type === 'standard' || type === 'expanded';
  const showApn = type === 'standard' || type === 'expanded';
  const showGeneralUnderstanding = type === 'expanded';

  els.acreageWrap.classList.toggle('hidden', !showAcreage);
  els.apnWrap.classList.toggle('hidden', !showApn);
  els.generalProjectUnderstandingWrap.classList.toggle('hidden', !showGeneralUnderstanding);

  if (!showAcreage) els.acreage.value = '';
  if (!showApn) els.apn.value = '';
  if (!showGeneralUnderstanding) els.generalProjectUnderstanding.value = '';
}

async function loadTaskLibrary() {
  const response = await fetch('tasks.json');
  if (!response.ok) throw new Error('Unable to load tasks library.');
  const payload = await response.json();
  state.disciplines = payload.disciplines || [];
}

function getSelectedDiscipline() {
  return state.disciplines.find((discipline) => discipline.id === els.discipline.value) || null;
}

function getDisciplineCategoryMap(discipline) {
  const categories = Array.isArray(discipline.categories) ? discipline.categories : [];
  const tasks = Array.isArray(discipline.tasks) ? discipline.tasks : [];
  const grouped = new Map(categories.map((category) => [category, []]));

  tasks.forEach((task) => {
    const category = task.category || 'General / Miscellaneous';
    if (!grouped.has(category)) grouped.set(category, []);
    grouped.get(category).push(task);
  });

  return grouped;
}

function renderCategoryGroup(categoryName, tasks, index) {
  const groupEl = document.createElement('section');
  groupEl.className = 'task-group collapsed';

  const headerBtn = document.createElement('button');
  headerBtn.type = 'button';
  headerBtn.className = 'task-group-header';
  headerBtn.setAttribute('aria-expanded', 'false');
  headerBtn.setAttribute('aria-controls', `task-group-panel-${index}`);

  const arrow = document.createElement('span');
  arrow.className = 'task-group-arrow';
  arrow.textContent = '▸';

  const title = document.createElement('span');
  title.className = 'task-group-title';
  title.textContent = categoryName;

  headerBtn.append(arrow, title);

  const panel = document.createElement('div');
  panel.className = 'task-group-panel';
  panel.id = `task-group-panel-${index}`;

  if (tasks.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'task-group-empty';
    empty.textContent = 'No tasks defined yet in this category.';
    panel.appendChild(empty);
  } else {
    tasks.forEach((task) => {
      const wrapper = document.createElement('label');
      wrapper.className = 'task-item';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = task.id;
      checkbox.checked = state.selectedTaskIds.has(task.id);

      checkbox.addEventListener('change', (event) => {
        if (event.target.checked) {
          state.selectedTaskIds.add(task.id);
        } else {
          state.selectedTaskIds.delete(task.id);
        }
        refreshPreviewDocument();
      });

      const textWrap = document.createElement('span');
      textWrap.innerHTML = `<strong>${task.name}</strong><br>${task.description}`;
      wrapper.append(checkbox, textWrap);
      panel.appendChild(wrapper);
    });
  }

  headerBtn.addEventListener('click', () => {
    const expanded = headerBtn.getAttribute('aria-expanded') === 'true';
    headerBtn.setAttribute('aria-expanded', String(!expanded));
    groupEl.classList.toggle('collapsed', expanded);
    groupEl.classList.toggle('expanded', !expanded);
  });

  groupEl.append(headerBtn, panel);
  return groupEl;
}

function renderTaskSection() {
  const discipline = getSelectedDiscipline();
  state.selectedTaskIds.clear();

  if (!discipline) {
    els.taskCategoryLabel.textContent = 'Select a discipline to load tasks.';
    els.taskList.innerHTML = '';
    refreshPreviewDocument();
    return;
  }

  els.taskCategoryLabel.textContent = discipline.taskCategory;
  els.taskList.innerHTML = '';

  if (discipline.status === 'placeholder') {
    const note = document.createElement('p');
    note.className = 'placeholder-note';
    note.textContent = discipline.placeholderMessage || 'This discipline is a placeholder for future implementation.';
    els.taskList.appendChild(note);
    refreshPreviewDocument();
    return;
  }

  Array.from(getDisciplineCategoryMap(discipline).entries()).forEach(([categoryName, tasks], index) => {
    els.taskList.appendChild(renderCategoryGroup(categoryName, tasks, index));
  });

  refreshPreviewDocument();
}

function getSelectedTasks(discipline) {
  const allTasks = Array.isArray(discipline?.tasks) ? discipline.tasks : [];
  return allTasks.filter((task) => state.selectedTaskIds.has(task.id));
}

function getProjectInfo() {
  const coverLetterType = els.coverLetterType.value;

  return {
    projectName: els.projectName.value.trim(),
    address: els.projectAddress.value.trim(),
    city: els.projectCity.value.trim(),
    county: getCountyValue(),
    state: els.projectState.value.trim(),
    client: els.client.value.trim(),
    contactName: els.contactName.value.trim(),
    contactEmail: els.contactEmail.value.trim(),
    clientAddress: els.clientAddress.value.trim(),
    subjectLine: els.subjectLine.value.trim(),
    regionalFramework: getRegionalFrameworkValue(),
    acreage: els.acreage.value.trim(),
    apn: els.apn.value.trim(),
    generalProjectUnderstanding: els.generalProjectUnderstanding.value.trim(),
    coverLetterType,
    date: els.scopeDate.value || todayIsoDate(),
    signatories: [
      { name: els.signatory1Name.value.trim(), title: els.signatory1Title.value.trim() },
      { name: els.signatory2Name.value.trim(), title: els.signatory2Title.value.trim() },
    ],
  };
}

function escapeHtml(input) {
  return String(input ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formatDisplayDate(dateValue) {
  if (!dateValue) return '';
  const date = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(date.getTime())) return dateValue;

  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
}

function buildCoverLetterText(project) {
  if (project.coverLetterType === 'none') {
    return '';
  }

  const lines = [];
  lines.push(formatDisplayDate(project.date));

  if (project.client) lines.push(project.client);
  if (project.contactName) lines.push(project.contactName);
  if (project.clientAddress) lines.push(project.clientAddress);

  const subjectText = project.subjectLine || project.projectName || 'Scope of Work';
  lines.push(`Subject: ${subjectText}`);

  const greetingName = project.contactName || 'Client';
  lines.push(`Dear ${greetingName},`);

  const locationText =
    project.address ||
    [project.city, project.county, project.state].filter(Boolean).join(', ') ||
    'the project site';

  let body =
    `Please find Attachment A, Scope of Work for biological consulting services for the ${project.projectName || 'project'} ` +
    `located at ${locationText}.`;

  if (project.coverLetterType === 'standard' || project.coverLetterType === 'expanded') {
    if (project.acreage) {
      body += ` The site encompasses approximately ${project.acreage}.`;
    }
    if (project.apn) {
      body += ` Associated APN: ${project.apn}.`;
    }
  }

  if (project.coverLetterType === 'expanded' && project.generalProjectUnderstanding) {
    body += ` ${project.generalProjectUnderstanding}`;
  }

  lines.push(body);
  lines.push('We appreciate the opportunity to support your team.');

  return lines.join('\n\n');
}

function buildAttachmentAText(project, selectedTasks) {
  const header = [
    'ATTACHMENT A',
    'SCOPE OF WORK',
    'BIOLOGICAL CONSULTING SERVICES',
    project.projectName || 'Project',
  ].join('\n');

  const taskSection = selectedTasks.length
    ? selectedTasks
        .map((task, index) => {
          const titleLine = `${index + 1}. ${task.name}`;
          return `${titleLine}\n${task.description}`;
        })
        .join('\n\n')
    : 'No tasks selected.';

  const deliverables = selectedTasks
    .flatMap((task) => {
      if (!task.deliverables) return [];
      if (Array.isArray(task.deliverables)) return task.deliverables;
      return [task.deliverables];
    })
    .filter(Boolean);

  const deliverablesSection =
    deliverables.length > 0
      ? `\n\nDELIVERABLES\n${deliverables.map((item) => `- ${item}`).join('\n')}`
      : '';

  return `${header}\n\n${taskSection}${deliverablesSection}`;
}

function buildSignatureBlockHtml(project) {
  return `
    <div class="document-signature-grid">
      <div class="document-signature-block">
        <div class="document-signature-line"></div>
        <div class="document-signature-name">${escapeHtml(project.signatories[0]?.name || '')}</div>
        <div class="document-signature-title">${escapeHtml(project.signatories[0]?.title || '')}</div>
      </div>
      <div class="document-signature-block">
        <div class="document-signature-line"></div>
        <div class="document-signature-name">${escapeHtml(project.signatories[1]?.name || '')}</div>
        <div class="document-signature-title">${escapeHtml(project.signatories[1]?.title || '')}</div>
      </div>
    </div>
  `;
}

function buildFullDocumentText(project, selectedTasks) {
  const coverLetterText = buildCoverLetterText(project);
  const attachmentText = buildAttachmentAText(project, selectedTasks);

  if (!coverLetterText) return attachmentText;

  const signatureText = [
    '\n\n',
    '____________________________________',
    project.signatories[0]?.name || '',
    project.signatories[0]?.title || '',
    '\n\n',
    '____________________________________',
    project.signatories[1]?.name || '',
    project.signatories[1]?.title || '',
  ].join('\n');

  return `${coverLetterText}${signatureText}\n\n----------------------------------------\nATTACHMENT A BEGINS\n----------------------------------------\n\n${attachmentText}`;
}

function buildPreviewDocumentHtml(project, selectedTasks) {
  const coverLetterText = buildCoverLetterText(project);
  const coverLetterParagraphs = coverLetterText
    ? coverLetterText
        .split(/\n\s*\n/)
        .map((paragraph) => `<p class="document-paragraph">${escapeHtml(paragraph)}</p>`)
        .join('')
    : `<p class="document-paragraph">Cover letter disabled.</p>`;

  const attachmentTasksHtml = selectedTasks.length
    ? selectedTasks
        .map((task, index) => {
          return `
            <div class="document-task">
              <div class="document-task-title">${escapeHtml(`${index + 1}. ${task.name}`)}</div>
              <p class="document-task-text">${escapeHtml(task.description || '')}</p>
            </div>
          `;
        })
        .join('')
    : `<p class="document-paragraph">No tasks selected.</p>`;

  const deliverables = selectedTasks
    .flatMap((task) => {
      if (!task.deliverables) return [];
      if (Array.isArray(task.deliverables)) return task.deliverables;
      return [task.deliverables];
    })
    .filter(Boolean);

  const deliverablesHtml = deliverables.length
    ? `
      <div class="document-deliverables">
        <div class="document-deliverables-title">DELIVERABLES</div>
        <ul class="document-deliverables-list">
          ${deliverables.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}
        </ul>
      </div>
    `
    : '';

  const coverSection = `
    <div class="document-section">
      ${coverLetterParagraphs}
      ${project.coverLetterType === 'none' ? '' : buildSignatureBlockHtml(project)}
    </div>
  `;

  const attachmentSection = `
    <div class="document-section">
      <div class="document-section-title">ATTACHMENT A – SCOPE OF WORK</div>
      ${project.projectName ? `<p class="document-paragraph">${escapeHtml(project.projectName)}</p>` : ''}
      ${attachmentTasksHtml}
      ${deliverablesHtml}
    </div>
  `;

  if (project.coverLetterType === 'none') {
    return attachmentSection;
  }

  return `
    ${coverSection}
    <div class="document-page-break">
      <span class="document-page-break-label">Attachment A Begins</span>
    </div>
    ${attachmentSection}
  `;
}

function updateDocumentActionsVisibility() {
  const inPreview = state.currentDocumentMode === 'preview';
  els.sendPreviewToEditorBtn.classList.toggle('hidden', !inPreview);
  els.restorePreviousEditorBtn.classList.toggle(
    'hidden',
    inPreview || !state.previousEditorContent
  );
}

function setDocumentMode(mode) {
  state.currentDocumentMode = mode;

  const previewActive = mode === 'preview';
  els.previewModeBtn.classList.toggle('active', previewActive);
  els.previewModeBtn.setAttribute('aria-selected', String(previewActive));

  els.editorModeBtn.classList.toggle('active', !previewActive);
  els.editorModeBtn.setAttribute('aria-selected', String(!previewActive));

  els.previewDocumentPanel.classList.toggle('active', previewActive);
  els.editorDocumentPanel.classList.toggle('active', !previewActive);

  els.documentModeMessage.textContent = previewActive
    ? 'Live preview — updates automatically.'
    : 'Editing mode — changes are not auto-synced.';

  updateDocumentActionsVisibility();
}

function refreshPreviewDocument() {
  const discipline = getSelectedDiscipline();
  const project = getProjectInfo();
  const selectedTasks = discipline ? getSelectedTasks(discipline) : [];

  els.previewDocument.innerHTML = buildPreviewDocumentHtml(project, selectedTasks);
  els.exportBtn.disabled = !editorHasExportableContent();
}

function sendPreviewToEditor() {
  const discipline = getSelectedDiscipline();
  const project = getProjectInfo();
  const selectedTasks = discipline ? getSelectedTasks(discipline) : [];
  const newEditorHtml = buildPreviewDocumentHtml(project, selectedTasks);

  if (editorHasExportableContent()) {
    const confirmed = window.confirm(
      'This will replace your current editor content with the latest preview. Continue?'
    );
    if (!confirmed) return;
  }

  state.previousEditorContent = els.scopeEditor.innerHTML;
  els.scopeEditor.innerHTML = newEditorHtml;
  els.exportBtn.disabled = false;
  setDocumentMode('editor');
}

function restorePreviousEditorVersion() {
  const currentContent = els.scopeEditor.innerHTML;
  els.scopeEditor.innerHTML = state.previousEditorContent || createDefaultEditorHtml();
  state.previousEditorContent = currentContent;
  els.exportBtn.disabled = !editorHasExportableContent();
  updateDocumentActionsVisibility();
}

function validateForExport() {
  if (!editorHasExportableContent()) {
    alert('Please send preview content to the editor before exporting.');
    return false;
  }

  const discipline = getSelectedDiscipline();
  if (!discipline) {
    alert('Please select a discipline before exporting.');
    return false;
  }

  return true;
}

function buildFileName(projectName, discipline, date) {
  const cleanProjectName = sanitizeForFilename(projectName) || 'Project';
  return `${cleanProjectName}_${discipline.filenamePrefix}_${date}.docx`;
}

function convertHtmlToPlainText(html) {
  const temp = document.createElement('div');
  temp.innerHTML = html;

  temp.querySelectorAll('.document-page-break').forEach((el) => {
    const replacement = document.createTextNode(
      '\n----------------------------------------\nATTACHMENT A BEGINS\n----------------------------------------\n'
    );
    el.replaceWith(replacement);
  });

  temp.querySelectorAll('.document-signature-grid').forEach((grid) => {
    const blocks = Array.from(grid.querySelectorAll('.document-signature-block'));
    const lines = [];

    blocks.forEach((block, index) => {
      const name = block.querySelector('.document-signature-name')?.innerText?.trim() || '';
      const title = block.querySelector('.document-signature-title')?.innerText?.trim() || '';
      lines.push('____________________________________');
      if (name) lines.push(name);
      if (title) lines.push(title);
      if (index < blocks.length - 1) lines.push('');
    });

    grid.replaceWith(document.createTextNode(`\n${lines.join('\n')}\n`));
  });

  return temp.innerText.replace(/\n{3,}/g, '\n\n').trim();
}

function buildDocParagraphsFromText(text) {
  const { Paragraph } = window.docx;
  const lines = text.split('\n');

  return lines.map((line) => {
    const trimmed = line.trim();
    if (trimmed === '') return new Paragraph({ text: '' });
    if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      return new Paragraph({ text: trimmed.slice(2), bullet: { level: 0 } });
    }
    return new Paragraph({ text: line });
  });
}

async function exportToWord() {
  if (!validateForExport()) return;

  const discipline = getSelectedDiscipline();
  const project = getProjectInfo();
  const selectedTasks = getSelectedTasks(discipline);

  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    AlignmentType,
    TabStopType,
    TabStopPosition,
    PageBreak,
    convertInchesToTwip,
  } = window.docx;

  const rightTab = convertInchesToTwip(6.5);
  const subjectTab = convertInchesToTwip(1.2);

  const children = [];

  // ===== COVER LETTER =====

  if (project.coverLetterType !== 'none') {
    children.push(new Paragraph({
      text: formatDisplayDate(project.date),
      spacing: { after: 240 }
    }));

    children.push(new Paragraph({
      children: [
        new TextRun(project.client || ''),
        new TextRun('\t'),
        new TextRun({ text: 'VIA EMAIL', bold: true })
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: rightTab }]
    }));

    children.push(new Paragraph({
      children: [
        new TextRun(project.contactName || ''),
        new TextRun('\t'),
        new TextRun(project.contactEmail || '')
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: rightTab }]
    }));

    if (project.clientAddress) {
      project.clientAddress.split('\n').forEach(line => {
        children.push(new Paragraph({ text: line }));
      });
    }

    children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

    children.push(new Paragraph({
      children: [
        new TextRun({ text: 'Subject:', bold: true }),
        new TextRun('\t'),
        new TextRun(project.subjectLine || project.projectName)
      ],
      tabStops: [{ type: TabStopType.LEFT, position: subjectTab }],
      spacing: { after: 240 }
    }));

    children.push(new Paragraph({
      text: `Dear ${project.contactName || 'Client'}:`,
      spacing: { after: 240 }
    }));

    const location =
      project.address ||
      [project.city, project.county, project.state].filter(Boolean).join(', ');

    let body = `Please find attached Attachment A, Scope of Work for biological consulting services for the ${project.projectName} located at ${location}.`;

    if (project.acreage) body += ` The site encompasses approximately ${project.acreage}.`;
    if (project.apn) body += ` APN: ${project.apn}.`;

    children.push(new Paragraph({ text: body, spacing: { after: 240 } }));

    children.push(new Paragraph({
      text: 'We appreciate the opportunity to support your team.',
      spacing: { after: 240 }
    }));

    children.push(new Paragraph({
      text: 'Sincerely,',
      spacing: { after: 360 }
    }));

    children.push(new Paragraph({
      children: [
        new TextRun(project.signatories[0]?.name || ''),
        new TextRun('\t'),
        new TextRun(project.signatories[1]?.name || '')
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: rightTab }]
    }));

    children.push(new Paragraph({
      children: [
        new TextRun(project.signatories[0]?.title || ''),
        new TextRun('\t'),
        new TextRun(project.signatories[1]?.title || '')
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: rightTab }],
      spacing: { after: 240 }
    }));

    children.push(new Paragraph({
      text: 'Attachment A – Scope of Work and Fees'
    }));

    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // ===== ATTACHMENT A =====

  children.push(new Paragraph({
    text: 'ATTACHMENT A',
    alignment: AlignmentType.CENTER,
    bold: true
  }));

  children.push(new Paragraph({
    text: 'SCOPE OF WORK',
    alignment: AlignmentType.CENTER,
    bold: true
  }));

  children.push(new Paragraph({
    text: 'BIOLOGICAL CONSULTING SERVICES',
    alignment: AlignmentType.CENTER,
    bold: true
  }));

  children.push(new Paragraph({
    text: (project.projectName || '').toUpperCase(),
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 240 }
  }));

  children.push(new Paragraph({
    text: formatDisplayDate(project.date),
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 }
  }));

  // ===== TASKS =====

  selectedTasks.forEach((task, i) => {
    children.push(new Paragraph({
      children: [
        new TextRun({ text: `TASK ${i + 1} – ${task.name.toUpperCase()}`, bold: true }),
        new TextRun('\t'),
        new TextRun({ text: task.fee || '', bold: true })
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: rightTab }],
      spacing: { after: 160 }
    }));

    children.push(new Paragraph({
      text: task.description,
      spacing: { after: 200 }
    }));

    if (task.deliverables) {
      const text = Array.isArray(task.deliverables)
        ? task.deliverables.join('; ')
        : task.deliverables;

      children.push(new Paragraph({
        children: [
          new TextRun({ text: 'Deliverables: ', italics: true }),
          new TextRun({ text: text, italics: true })
        ],
        spacing: { after: 220 }
      }));
    }
  });

  const doc = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, buildFileName(project.projectName, discipline, project.date));
}

function handleRefreshPreview() {
  refreshPreviewDocument();
  setDocumentMode('preview');
}

function registerLivePreviewListeners() {
  const fields = [
    els.coverLetterType,
    els.projectName,
    els.projectAddress,
    els.projectCity,
    els.projectCounty,
    els.customProjectCounty,
    els.projectState,
    els.regionalFramework,
    els.customRegionalFramework,
    els.acreage,
    els.apn,
    els.generalProjectUnderstanding,
    els.client,
    els.contactName,
    els.contactEmail,
    els.clientAddress,
    els.subjectLine,
    els.scopeDate,
    els.signatory1Name,
    els.signatory1Title,
    els.signatory2Name,
    els.signatory2Title,
  ];

  fields.forEach((field) => {
    if (!field) return;
    field.addEventListener('input', refreshPreviewDocument);
    field.addEventListener('change', refreshPreviewDocument);
  });
}

function init() {
  els.scopeDate.value = todayIsoDate();
  els.scopeEditor.innerHTML = createDefaultEditorHtml();

  updateCountyInputVisibility();
  updateRegionalFrameworkVisibility();
  updateCoverLetterTypeVisibility();
  setDocumentMode('preview');

  els.discipline.addEventListener('change', renderTaskSection);

  els.projectCounty.addEventListener('change', () => {
    updateCountyInputVisibility();
    refreshPreviewDocument();
  });

  els.regionalFramework.addEventListener('change', () => {
    updateRegionalFrameworkVisibility();
    refreshPreviewDocument();
  });

  els.coverLetterType.addEventListener('change', () => {
    updateCoverLetterTypeVisibility();
    refreshPreviewDocument();
  });

  els.previewModeBtn.addEventListener('click', () => setDocumentMode('preview'));
  els.editorModeBtn.addEventListener('click', () => setDocumentMode('editor'));
  els.sendPreviewToEditorBtn.addEventListener('click', sendPreviewToEditor);
  els.restorePreviousEditorBtn.addEventListener('click', restorePreviousEditorVersion);

  els.generateBtn.addEventListener('click', handleRefreshPreview);

  els.exportBtn.addEventListener('click', () => {
    exportToWord().catch((error) => {
      console.error(error);
      alert('Unable to export document. Please try again.');
    });
  });

  els.scopeEditor.addEventListener('input', () => {
    els.exportBtn.disabled = !editorHasExportableContent();
  });

  registerLivePreviewListeners();

  loadTaskLibrary()
    .then(renderTaskSection)
    .catch((error) => {
      console.error(error);
      els.previewDocument.innerHTML = '<p class="document-placeholder">Error loading task library.</p>';
      els.scopeEditor.innerHTML = '<p class="document-placeholder">Error loading task library.</p>';
      els.exportBtn.disabled = true;
    });
}

init();