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
  saveProposalBtn: document.getElementById('saveProposalBtn'),
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
  clientAddress2: document.getElementById('clientAddress2'),
  subjectLine: document.getElementById('subjectLine'),
  scopeDate: document.getElementById('scopeDate'),
  signatory1Name: document.getElementById('signatory1Name'),
  signatory1Title: document.getElementById('signatory1Title'),
  signatory2Name: document.getElementById('signatory2Name'),
  signatory2Title: document.getElementById('signatory2Title'),
  loadProposalBtn: document.getElementById('loadProposalBtn'),
  loadProposalInput: document.getElementById('loadProposalInput'),
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
    clientAddress2: els.clientAddress2.value.trim(),
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
  if (project.contactEmail) lines.push(project.contactEmail);
  if (project.clientAddress) lines.push(project.clientAddress);
  if (project.clientAddress2) lines.push(project.clientAddress2);

  const subjectText = project.subjectLine || project.projectName || 'Scope of Work';
  lines.push(`Subject: ${subjectText}`);

  const greetingName = project.contactName || 'Client';
  lines.push(`Dear ${greetingName}:`);

  const locationText =
    project.address ||
    [project.city, project.county, project.state].filter(Boolean).join(', ') ||
    'the project site';

  let body =
    `Please find attached Attachment A, Scope of Work and Fee Estimate for biological consulting services for the ${project.projectName || 'project'} located at ${locationText}.`;

  if (project.coverLetterType === 'standard' || project.coverLetterType === 'expanded') {
    if (project.acreage) {
      body += ` The site encompasses approximately ${project.acreage}.`;
    }
    if (project.apn) {
      body += ` Associated APN${project.apn.includes(',') ? 's' : ''}: ${project.apn}.`;
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
    'SCOPE OF WORK AND FEE ESTIMATE',
    project.projectName || 'Project',
  ].join('\n');

  const taskSection = selectedTasks.length
    ? selectedTasks
        .map((task, index) => {
          const titleLine = `TASK ${index + 1} ${task.name}`;
          return `${titleLine}\n${task.description}`;
        })
        .join('\n\n')
    : 'No tasks selected.';

  return `${header}\n\n${taskSection}`;
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
          const feeDisplay = task.fee ? escapeHtml(formatCurrency(parseFeeValue(task.fee))) : '$';
          return `
            <div class="document-task">
              <div class="document-task-title">
                ${escapeHtml(`TASK ${index + 1}`)}
                <span style="display:inline-block; min-width: 1.5rem;"></span>
                ${escapeHtml(task.name || '')}
                <span style="float:right;">${feeDisplay}</span>
              </div>
              <p class="document-task-text">${escapeHtml(task.description || '')}</p>
            </div>
          `;
        })
        .join('')
    : `<p class="document-paragraph">No tasks selected.</p>`;

  const coverSection = `
    <div class="document-section">
      ${coverLetterParagraphs}
      ${project.coverLetterType === 'none' ? '' : buildSignatureBlockHtml(project)}
    </div>
  `;

  const attachmentSection = `
    <div class="document-section">
      <div class="document-section-title">ATTACHMENT A – SCOPE OF WORK AND FEE ESTIMATE</div>
      ${project.projectName ? `<p class="document-paragraph"><strong>${escapeHtml(project.projectName)}</strong></p>` : ''}
      ${project.date ? `<p class="document-paragraph"><strong>${escapeHtml(formatDisplayDate(project.date))}</strong></p>` : ''}
      ${attachmentTasksHtml}
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

function parseFeeValue(fee) {
  if (fee == null || fee === '') return 0;
  if (typeof fee === 'number' && Number.isFinite(fee)) return fee;

  const cleaned = String(fee).replace(/[^0-9.-]/g, '');
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrency(amount) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount || 0);
}

function formatFeeDisplay(fee) {
  const parsed = parseFeeValue(fee);
  if (!fee && fee !== 0) return '$';
  if (parsed === 0 && String(fee).trim() === '') return '$';
  return formatCurrency(parsed);
}

function isOptionalTask(task) {
  return Boolean(
    task?.optional === true ||
    task?.isOptional === true ||
    task?.taskType === 'optional' ||
    /optional/i.test(task?.category || '')
  );
}

function getPrimarySignatory(project) {
  const populated = (project.signatories || []).filter((s) => s?.name || s?.title);
  return populated[0] || { name: '', title: '' };
}

function getSecondarySignatory(project) {
  const populated = (project.signatories || []).filter((s) => s?.name || s?.title);
  return populated[1] || null;
}

function splitTextIntoParagraphs(text) {
  if (!text) return [];
  return String(text)
    .replace(/\r\n/g, '\n')
    .split(/\n\s*\n/)
    .map((part) => part.replace(/\n+/g, ' ').trim())
    .filter(Boolean);
}

function buildCoverLetterBodyParagraphs(project) {
  const locationText =
    project.address ||
    [project.city, project.county, project.state].filter(Boolean).join(', ') ||
    'the project site';

  let intro =
    `Please find attached Attachment A, Scope of Work and Fee Estimate for biological consulting services for the ${project.projectName || 'project'} located at ${locationText}.`;

  if (project.coverLetterType === 'standard' || project.coverLetterType === 'expanded') {
    if (project.acreage) {
      intro += ` The site encompasses approximately ${project.acreage}.`;
    }
    if (project.apn) {
      intro += ` Associated APN${project.apn.includes(',') ? 's' : ''}: ${project.apn}.`;
    }
  }

  const paragraphs = [intro];

  if (project.coverLetterType === 'expanded' && project.generalProjectUnderstanding) {
    paragraphs.push(project.generalProjectUnderstanding.trim());
  }

  paragraphs.push('We appreciate the opportunity to support your team.');

  return paragraphs;
}

function buildAttachmentTitleLines(project) {
  const projectName = (project.projectName || 'PROJECT SITE').toUpperCase();

  return [
    'ATTACHMENT A',
    'SCOPE OF WORK AND FEE ESTIMATE',
    'BIOLOGICAL CONSULTING SERVICES',
    projectName,
  ];
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
    BorderStyle,
    Header,
    Footer,
    PageNumber,
    NumberFormat,
    SectionType,
    convertInchesToTwip,
  } = window.docx;

  const pageMarginTwips = convertInchesToTwip(1.0);
  const coverRightTab = convertInchesToTwip(6.5);
  const signatureRightTab = convertInchesToTwip(3.5);
  const subjectTab = convertInchesToTwip(1.0);
  const taskTitleTab = convertInchesToTwip(1.0);
  const taskAmountTab = convertInchesToTwip(6.5);
  const footerCenterTab = convertInchesToTwip(3.25);
  const footerRightTab = convertInchesToTwip(6.5);

  function makeRun(text = '', options = {}) {
    return new TextRun({
      text,
      font: 'Arial',
      size: 22,
      ...options,
    });
  }

  function makeParagraph(text = '', options = {}) {
    const {
      bold = false,
      italics = false,
      alignment,
      spacing,
      tabStops,
      indent,
      border,
      keepLines,
      keepNext,
      children,
      pageBreakBefore,
    } = options;

    if (children) {
      return new Paragraph({
        children,
        alignment,
        spacing,
        tabStops,
        indent,
        border,
        keepLines,
        keepNext,
        pageBreakBefore,
      });
    }

    return new Paragraph({
      children: [makeRun(text, { bold, italics })],
      alignment,
      spacing,
      tabStops,
      indent,
      border,
      keepLines,
      keepNext,
      pageBreakBefore,
    });
  }

  function makeBlankParagraph(after = 0) {
    return new Paragraph({
      children: [makeRun('')],
      spacing: { after },
    });
  }

  function makeJustifiedParagraph(text, spacingAfter = 120, extraOptions = {}) {
    return makeParagraph(text, {
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: spacingAfter, line: 276 },
      ...extraOptions,
    });
  }

  function pushLinesAsParagraphs(target, blockText, options = {}) {
    if (!blockText) return;
    String(blockText)
      .replace(/\r\n/g, '\n')
      .split('\n')
      .map((line) => line.trim())
      .forEach((line) => {
        target.push(
          makeParagraph(line, {
            spacing: { after: options.after ?? 0 },
            keepLines: true,
          })
        );
      });
  }

  function buildCoverLetterChildren() {
    const children = [];
    const primarySignatory = getPrimarySignatory(project);
    const secondarySignatory = getSecondarySignatory(project);
    const subjectText = project.subjectLine || project.projectName || 'Scope of Work';
    const greetingName = project.contactName || 'Client';

    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());

    children.push(
      makeParagraph(formatDisplayDate(project.date), {
        spacing: { after: 0 },
      })
    );

    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());

    children.push(
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: coverRightTab }],
        spacing: { after: 0 },
        keepLines: true,
        children: [
          makeRun(project.client || ''),
          makeRun('\t'),
          makeRun('VIA EMAIL', { bold: true }),
        ],
      })
    );

    children.push(
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: coverRightTab }],
        spacing: { after: 0 },
        keepLines: true,
        children: [
          makeRun(project.contactName || ''),
          makeRun('\t'),
          makeRun(project.contactEmail || '', { bold: true }),
        ],
      })
    );

    if (project.clientAddress) {
  children.push(
    makeParagraph(project.clientAddress, {
      spacing: { after: 0 },
      keepLines: true,
    })
  );
}

if (project.clientAddress2) {
  children.push(
    makeParagraph(project.clientAddress2, {
      spacing: { after: 0 },
      keepLines: true,
    })
  );
}

    children.push(makeBlankParagraph());

    children.push(
      new Paragraph({
        tabStops: [{ type: TabStopType.LEFT, position: subjectTab }],
        spacing: { after: 220 },
        keepLines: true,
        children: [
          makeRun('Subject:', { bold: false }),
          makeRun('\t'),
          makeRun(subjectText),
        ],
      })
    );

    children.push(
      makeParagraph(`Dear ${greetingName}:`, {
        spacing: { after: 220 },
      })
    );

    buildCoverLetterBodyParagraphs(project).forEach((paragraphText, index, arr) => {
      children.push(
        makeJustifiedParagraph(
          paragraphText,
          index === arr.length - 1 ? 200 : 160
        )
      );
    });

    children.push(
      makeParagraph('Sincerely,', {
        spacing: { after: 0 },
      })
    );

    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());

    if (secondarySignatory && (secondarySignatory.name || secondarySignatory.title)) {
      children.push(
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: signatureRightTab }],
          spacing: { after: 0 },
          keepLines: true,
          children: [
            makeRun(primarySignatory.name || ''),
            makeRun('\t'),
            makeRun(secondarySignatory.name || ''),
          ],
        })
      );

      children.push(
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: signatureRightTab }],
          spacing: { after: 0 },
          keepLines: true,
          children: [
            makeRun(primarySignatory.title || ''),
            makeRun('\t'),
            makeRun(secondarySignatory.title || ''),
          ],
        })
      );
    } else {
      if (primarySignatory.name) {
        children.push(
          makeParagraph(primarySignatory.name, {
            spacing: { after: 0 },
            keepLines: true,
          })
        );
      }
      if (primarySignatory.title) {
        children.push(
          makeParagraph(primarySignatory.title, {
            spacing: { after: 0 },
            keepLines: true,
          })
        );
      }
    }

    children.push(makeBlankParagraph());
    children.push(makeBlankParagraph());

    children.push(
      makeParagraph('Attachment A – Scope of Work and Fee Estimate', {
        spacing: { after: 0 },
      })
    );

    return children;
  }

  function buildAttachmentHeader() {
    return new Header({
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              size: 4,
              color: 'A6A6A6',
              space: 1,
            },
          },
          spacing: { after: 0 },
          children: [
            makeRun(project.projectName || '', {
              italics: true,
              color: '7F7F7F',
              size: 16,
            }),
          ],
        }),
      ],
    });
  }

  function buildAttachmentFooter() {
    return new Footer({
      children: [
        new Paragraph({
          border: {
            top: {
              style: BorderStyle.SINGLE,
              size: 4,
              color: 'A6A6A6',
              space: 1,
            },
          },
          tabStops: [
            { type: TabStopType.CENTER, position: footerCenterTab },
            { type: TabStopType.RIGHT, position: footerRightTab },
          ],
          spacing: { before: 120, after: 0 },
          children: [
            makeRun('\t'),
            makeRun('A-'),
            PageNumber.CURRENT,
            makeRun('\t'),
            makeRun('Scope of Work and Fee Estimate', {
              italics: true,
              color: '7F7F7F',
              size: 16,
            }),
          ],
        }),
      ],
    });
  }

  function buildTaskHeadingParagraph(taskNumber, taskName, feeText, optional = false) {
    const labelPrefix = optional ? `OPTIONAL TASK ${taskNumber}` : `TASK ${taskNumber}`;
    const displayFee = feeText && String(feeText).trim() ? String(feeText).trim() : '$';

    return new Paragraph({
      tabStops: [
        { type: TabStopType.LEFT, position: taskTitleTab },
        { type: TabStopType.RIGHT, position: taskAmountTab },
      ],
      spacing: { after: 100 },
      keepLines: true,
      keepNext: true,
      children: [
        makeRun(labelPrefix, { bold: true }),
        makeRun('\t'),
        makeRun(String(taskName || '').toUpperCase(), { bold: true }),
        makeRun('\t'),
        makeRun(displayFee, { bold: true }),
      ],
    });
  }

  function buildTaskDescriptionParagraphs(text) {
    return splitTextIntoParagraphs(text).map((paragraphText, index, arr) =>
      makeJustifiedParagraph(paragraphText, index === arr.length - 1 ? 140 : 110, {
        keepLines: true,
      })
    );
  }

  function buildDeliverablesParagraph(deliverables) {
    if (!deliverables) return null;

    const content = Array.isArray(deliverables)
      ? deliverables.filter(Boolean).join('; ')
      : String(deliverables).trim();

    if (!content) return null;

    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 160, line: 276 },
      keepLines: true,
      children: [
        makeRun('Deliverables: ', { italics: true }),
        makeRun(content, { italics: true }),
      ],
    });
  }

  function buildTotalsParagraph(label, totalText) {
    return new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: taskAmountTab }],
      spacing: { before: 80, after: 110 },
      keepLines: true,
      children: [
        makeRun(label, { bold: true, italics: true }),
        makeRun('\t'),
        makeRun(totalText, { bold: true }),
      ],
    });
  }

  function buildAttachmentChildren() {
    const children = [];

    children.push(makeBlankParagraph());

children.push(
  makeParagraph('ATTACHMENT A', {
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 0 },
    keepLines: true,
    keepNext: true,
  })
);

children.push(makeBlankParagraph());

children.push(
  makeParagraph('SCOPE OF WORK AND FEE ESTIMATE', {
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 0 },
    keepLines: true,
    keepNext: true,
  })
);

children.push(
  makeParagraph('BIOLOGICAL CONSULTING SERVICES', {
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 0 },
    keepLines: true,
    keepNext: true,
  })
);

children.push(
  makeParagraph((project.projectName || 'PROJECT SITE').toUpperCase(), {
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 0 },
    keepLines: true,
    keepNext: true,
  })
);

children.push(makeBlankParagraph());

children.push(
  makeParagraph(formatDisplayDate(project.date), {
    alignment: AlignmentType.CENTER,
    bold: true,
    spacing: { after: 0 },
    keepLines: true,
    keepNext: true,
  })
);

children.push(makeBlankParagraph());

    const baseTasks = selectedTasks.filter((task) => !isOptionalTask(task));
    const optionalTasks = selectedTasks.filter((task) => isOptionalTask(task));
    const orderedTasks = [...baseTasks, ...optionalTasks];

    if (!orderedTasks.length) {
      children.push(
        makeParagraph('No tasks selected.', {
          spacing: { after: 120 },
        })
      );
      return children;
    }

    orderedTasks.forEach((task, index) => {
      const optional = isOptionalTask(task);
      const feeText = formatFeeDisplay(task.fee);

      children.push(
        buildTaskHeadingParagraph(index + 1, task.name, feeText, optional)
      );

      const taskParagraphs = buildTaskDescriptionParagraphs(task.description || '');
      if (taskParagraphs.length) {
        children.push(...taskParagraphs);
      } else {
        children.push(makeParagraph('', { spacing: { after: 120 } }));
      }

      const deliverablesParagraph = buildDeliverablesParagraph(task.deliverables);
      if (deliverablesParagraph) {
        children.push(deliverablesParagraph);
      }
    });

    const baseTotal = baseTasks.reduce((sum, task) => sum + parseFeeValue(task.fee), 0);
    const grandTotal = orderedTasks.reduce((sum, task) => sum + parseFeeValue(task.fee), 0);

    if (baseTasks.length > 0) {
      const firstTaskNumber = 1;
      const lastTaskNumber = baseTasks.length;
      const baseLabel =
        firstTaskNumber === lastTaskNumber
          ? `TOTAL TASK ${firstTaskNumber}`
          : `TOTAL TASKS ${firstTaskNumber}—${lastTaskNumber}`;
      children.push(buildTotalsParagraph(baseLabel, formatCurrency(baseTotal)));
    }

    if (optionalTasks.length > 0) {
      const firstTaskNumber = 1;
      const lastTaskNumber = orderedTasks.length;
      children.push(
        buildTotalsParagraph(
          `TOTAL TASKS ${firstTaskNumber}—${lastTaskNumber} INCLUDING OPTIONAL TASKS`,
          formatCurrency(grandTotal)
        )
      );
    }

    return children;
  }

  const coverLetterChildren = buildCoverLetterChildren();
  const attachmentChildren = buildAttachmentChildren();

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: 'Arial',
            size: 22,
          },
          paragraph: {
            spacing: {
              line: 276,
            },
          },
        },
      },
    },
    sections: [
      ...(project.coverLetterType !== 'none'
        ? [
            {
              properties: {
                page: {
                  margin: {
                    top: pageMarginTwips,
                    right: pageMarginTwips,
                    bottom: pageMarginTwips,
                    left: pageMarginTwips,
                  },
                },
              },
              children: coverLetterChildren,
            },
          ]
        : []),
      {
        properties: {
          type: project.coverLetterType !== 'none' ? SectionType.NEXT_PAGE : undefined,
          page: {
            margin: {
              top: convertInchesToTwip(0.7),
              right: pageMarginTwips,
              bottom: convertInchesToTwip(0.7),
              left: pageMarginTwips,
            },
            pageNumbers: {
              start: 1,
              formatType: NumberFormat.DECIMAL,
            },
          },
        },
        headers: {
          default: buildAttachmentHeader(),
        },
        footers: {
          default: buildAttachmentFooter(),
        },
        children: attachmentChildren,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, buildFileName(project.projectName, discipline, project.date));
}

function handleRefreshPreview() {
  refreshPreviewDocument();
  setDocumentMode('preview');
}

function handleSaveProposal() {
  const project = getProjectInfo();

  const payload = {
  schemaVersion: 1,
  savedAt: new Date().toISOString(),
  discipline: els.discipline.value,
  formData: project,
  selectedTaskIds: Array.from(state.selectedTaskIds),
  coverLetterType: project.coverLetterType,
  editorContent: els.scopeEditor.innerHTML || '',
};

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: 'application/json',
  });

  const fileNameBase = (project.projectName || 'proposal')
    .replace(/\s+/g, '_')
    .toLowerCase();

  const fileName = `${fileNameBase}_scopebuilder.json`;

  saveAs(blob, fileName);
}

function handleLoadProposalClick() {
  els.loadProposalInput.click();
}

function handleLoadProposalFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = JSON.parse(e.target.result);
      const project = data.formData || {};

      for (const key in project) {
        if (els[key] && typeof els[key].value !== 'undefined') {
          els[key].value = project[key];
        }
      }

      if (data.discipline) {
  els.discipline.value = data.discipline;
}

state.selectedTaskIds = new Set(data.selectedTaskIds || []);

renderTaskSection();
refreshPreviewDocument();

if (data.editorContent) {
  els.scopeEditor.innerHTML = data.editorContent;
  els.exportBtn.disabled = false;
} else {
  els.scopeEditor.innerHTML = createDefaultEditorHtml();
}

      if (data.editorContent) {
        els.exportBtn.disabled = false;
      }
    } catch (error) {
      console.error(error);
      alert('Unable to load proposal file.');
    } finally {
      els.loadProposalInput.value = '';
    }
  };

  reader.readAsText(file);
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
    els.clientAddress2,
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
els.saveProposalBtn.addEventListener('click', handleSaveProposal);
els.loadProposalBtn.addEventListener('click', handleLoadProposalClick);
els.loadProposalInput.addEventListener('change', handleLoadProposalFile);

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