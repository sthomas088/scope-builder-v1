const state = {
  disciplines: [],
  selectedTaskIds: new Set(),
  activePreviewTab: 'coverLetter',
};

const DEFAULT_EDITOR_TEXT = 'Select a discipline, complete project information, and click "Generate Scope".';

const els = {
  discipline: document.getElementById('discipline'),
  coverLetterType: document.getElementById('coverLetterType'),
  coverLetterTab: document.getElementById('coverLetterTab'),
  attachmentTab: document.getElementById('attachmentTab'),
  coverLetterPanel: document.getElementById('coverLetterPanel'),
  attachmentPanel: document.getElementById('attachmentPanel'),
  coverLetterDisabledMessage: document.getElementById('coverLetterDisabledMessage'),
  coverLetterEditor: document.getElementById('coverLetterEditor'),
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

function editorHasExportableContent() {
  return els.scopeEditor.value.trim() && els.scopeEditor.value.trim() !== DEFAULT_EDITOR_TEXT;
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

function setActivePreviewTab(tabName) {
  if (tabName === 'coverLetter' && els.coverLetterType.value === 'none') {
    state.activePreviewTab = 'attachment';
  } else {
    state.activePreviewTab = tabName;
  }

  const coverActive = state.activePreviewTab === 'coverLetter';
  els.coverLetterTab.classList.toggle('active', coverActive);
  els.coverLetterTab.setAttribute('aria-selected', String(coverActive));
  els.coverLetterPanel.classList.toggle('active', coverActive);

  const attachmentActive = state.activePreviewTab === 'attachment';
  els.attachmentTab.classList.toggle('active', attachmentActive);
  els.attachmentTab.setAttribute('aria-selected', String(attachmentActive));
  els.attachmentPanel.classList.toggle('active', attachmentActive);
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

  const coverLetterDisabled = type === 'none';
  els.coverLetterTab.disabled = coverLetterDisabled;
  els.coverLetterDisabledMessage.classList.toggle('hidden', !coverLetterDisabled);
  els.coverLetterEditor.classList.toggle('hidden', coverLetterDisabled);

  if (coverLetterDisabled && state.activePreviewTab === 'coverLetter') {
    setActivePreviewTab('attachment');
  } else {
    setActivePreviewTab(state.activePreviewTab);
  }
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

      checkbox.addEventListener('change', (event) => {
        if (event.target.checked) state.selectedTaskIds.add(task.id);
        else state.selectedTaskIds.delete(task.id);
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
  els.exportBtn.disabled = !editorHasExportableContent();

  if (!discipline) {
    els.taskCategoryLabel.textContent = 'Select a discipline to load tasks.';
    els.taskList.innerHTML = '';
    return;
  }

  els.taskCategoryLabel.textContent = discipline.taskCategory;
  els.taskList.innerHTML = '';

  if (discipline.status === 'placeholder') {
    const note = document.createElement('p');
    note.className = 'placeholder-note';
    note.textContent = discipline.placeholderMessage || 'This discipline is a placeholder for future implementation.';
    els.taskList.appendChild(note);
    return;
  }

  Array.from(getDisciplineCategoryMap(discipline).entries()).forEach(([categoryName, tasks], index) => {
    els.taskList.appendChild(renderCategoryGroup(categoryName, tasks, index));
  });
}

function getSelectedTasks(discipline) {
  const allTasks = Array.isArray(discipline.tasks) ? discipline.tasks : [];
  return allTasks.filter((task) => state.selectedTaskIds.has(task.id));
}

function buildAttachmentAText(project, selectedTasks) {
  const taskSection = selectedTasks
    .map((task, index) => `${index + 1}. ${task.name}\n${task.description}`)
    .join('\n\n');

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

  return `ATTACHMENT A\nSCOPE OF WORK\nBIOLOGICAL CONSULTING SERVICES\n${project.projectName}\n\n${taskSection}${deliverablesSection}`;
}

function buildCoverLetterText(project) {
  if (project.coverLetterType === 'none') {
    return '';
  }

  const locationText = project.address
    ? project.address
    : [project.city, project.county, project.state].filter(Boolean).join(', ');

  return `${project.date}\n\n${project.client}\n${project.contactName || ''}\n${project.clientAddress || ''}\n\nSubject: ${project.subjectLine || project.projectName}\n\nDear ${project.contactName || 'Client'},\n\nPlease find attached Attachment A, Scope of Work for biological consulting services for the ${project.projectName} project located at ${locationText}.\n\nWe appreciate the opportunity to support your team.\n`;
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
    coverLetterTypeLabel: coverLetterType.charAt(0).toUpperCase() + coverLetterType.slice(1),
    date: els.scopeDate.value || todayIsoDate(),
    signatories: [
      { name: els.signatory1Name.value.trim(), title: els.signatory1Title.value.trim() },
      { name: els.signatory2Name.value.trim(), title: els.signatory2Title.value.trim() },
    ],
  };
}

function validate(project, discipline, selectedTasks) {
  if (!discipline) {
    alert('Please select a discipline first.');
    return false;
  }

  const requiredFields = ['projectName', 'city', 'county', 'state', 'client'];
  const missing = requiredFields.filter((field) => !project[field]);
  if (missing.length > 0) {
    alert('Please complete all project information fields before generating scope.');
    return false;
  }

  if (discipline.status === 'active' && selectedTasks.length === 0) {
    alert('Please select at least one task.');
    return false;
  }

  return true;
}

function buildFileName(projectName, discipline, date) {
  const cleanProjectName = sanitizeForFilename(projectName) || 'Project';
  return `${cleanProjectName}_${discipline.filenamePrefix}_${date}.docx`;
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
  const currentScopeText = els.scopeEditor.value;
  if (!currentScopeText.trim() || currentScopeText.trim() === DEFAULT_EDITOR_TEXT) {
    alert('Please generate or enter scope content before exporting.');
    return;
  }

  const discipline = getSelectedDiscipline();
  if (!discipline) {
    alert('Please select a discipline before exporting.');
    return;
  }

  const { Document, Packer } = window.docx;
  const doc = new Document({
    sections: [{ properties: {}, children: buildDocParagraphsFromText(currentScopeText) }],
  });

  const blob = await Packer.toBlob(doc);
  const project = getProjectInfo();
  saveAs(blob, buildFileName(project.projectName, discipline, project.date));
}

function handleGenerateScope() {
  const discipline = getSelectedDiscipline();
  const project = getProjectInfo();
  const selectedTasks = discipline ? getSelectedTasks(discipline) : [];

  if (!validate(project, discipline, selectedTasks)) return;

  els.scopeEditor.value = buildAttachmentAText(project, selectedTasks);

  if (project.coverLetterType === 'none') {
    els.coverLetterEditor.value = '';
  } else {
    els.coverLetterEditor.value = buildCoverLetterText(project);
  }

  els.exportBtn.disabled = false;
}

function init() {
  els.scopeDate.value = todayIsoDate();
  els.scopeEditor.value = DEFAULT_EDITOR_TEXT;

  updateCountyInputVisibility();
  updateRegionalFrameworkVisibility();
  updateCoverLetterTypeVisibility();
  setActivePreviewTab('coverLetter');

  els.discipline.addEventListener('change', renderTaskSection);
  els.projectCounty.addEventListener('change', updateCountyInputVisibility);
  els.regionalFramework.addEventListener('change', updateRegionalFrameworkVisibility);
  els.coverLetterType.addEventListener('change', updateCoverLetterTypeVisibility);

  els.coverLetterTab.addEventListener('click', () => setActivePreviewTab('coverLetter'));
  els.attachmentTab.addEventListener('click', () => setActivePreviewTab('attachment'));

  els.generateBtn.addEventListener('click', handleGenerateScope);
  els.exportBtn.addEventListener('click', () => {
    exportToWord().catch((error) => {
      console.error(error);
      alert('Unable to export document. Please try again.');
    });
  });
  els.scopeEditor.addEventListener('input', () => {
    els.exportBtn.disabled = !editorHasExportableContent();
  });

  loadTaskLibrary()
    .then(renderTaskSection)
    .catch((error) => {
      console.error(error);
      els.scopeEditor.value = 'Error loading task library.';
      els.exportBtn.disabled = true;
    });
}

init();
