const state = {
  disciplines: [],
  selectedTaskIds: new Set(),
  generatedScope: '',
};

const els = {
  discipline: document.getElementById('discipline'),
  taskCategoryLabel: document.getElementById('taskCategoryLabel'),
  taskList: document.getElementById('taskList'),
  generateBtn: document.getElementById('generateBtn'),
  exportBtn: document.getElementById('exportBtn'),
  scopePreview: document.getElementById('scopePreview'),
  projectName: document.getElementById('projectName'),
  client: document.getElementById('client'),
  location: document.getElementById('location'),
  preparedBy: document.getElementById('preparedBy'),
  scopeDate: document.getElementById('scopeDate'),
};

function todayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function sanitizeForFilename(input) {
  return input.trim().replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ');
}

async function loadTaskLibrary() {
  const response = await fetch('tasks.json');
  if (!response.ok) {
    throw new Error('Unable to load tasks library.');
  }

  const payload = await response.json();
  state.disciplines = payload.disciplines || [];
}

function getSelectedDiscipline() {
  return state.disciplines.find((discipline) => discipline.id === els.discipline.value) || null;
}

function renderTaskSection() {
  const discipline = getSelectedDiscipline();
  state.selectedTaskIds.clear();
  state.generatedScope = '';
  els.exportBtn.disabled = true;

  if (!discipline) {
    els.taskCategoryLabel.textContent = 'Select a discipline to load tasks.';
    els.taskList.innerHTML = '';
    els.scopePreview.textContent = 'Select a discipline, complete project information, and click "Generate Scope".';
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

  discipline.tasks.forEach((task) => {
    const wrapper = document.createElement('label');
    wrapper.className = 'task-item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = task.id;

    checkbox.addEventListener('change', (event) => {
      if (event.target.checked) {
        state.selectedTaskIds.add(task.id);
      } else {
        state.selectedTaskIds.delete(task.id);
      }
    });

    const textWrap = document.createElement('span');
    textWrap.innerHTML = `<strong>${task.name}</strong><br>${task.description}`;

    wrapper.append(checkbox, textWrap);
    els.taskList.appendChild(wrapper);
  });
}

function getSelectedTasks(discipline) {
  return (discipline.tasks || []).filter((task) => state.selectedTaskIds.has(task.id));
}

function buildScopeText(project, discipline, selectedTasks) {
  if (discipline.status === 'placeholder') {
    return `${discipline.scopeTitle}

Project: ${project.projectName}
Client: ${project.client}
Location: ${project.location}
Prepared By: ${project.preparedBy}
Date: ${project.date}
Discipline: ${discipline.name}

Status
${discipline.placeholderMessage}

Next Step
Select Biology for fully implemented V1 scope generation and Word export.`;
  }

  const taskParagraphs = selectedTasks
    .map((task, index) => `${index + 1}. ${task.name}: ${task.description}`)
    .join('\n\n');

  return `${discipline.scopeTitle}

Project: ${project.projectName}
Client: ${project.client}
Location: ${project.location}
Prepared By: ${project.preparedBy}
Date: ${project.date}
Discipline: ${discipline.name}

Project Understanding
This scope of work describes biological services to support planning, environmental review, and regulatory strategy for the ${project.projectName} project in ${project.location}. Work products will be prepared to support agency coordination and to identify biological constraints that may affect project design, scheduling, and permitting.

Proposed Tasks
${taskParagraphs}

Deliverables
Consultant will provide a concise technical memorandum summarizing methods, observations, regulatory context, and project-specific recommendations. Mapping and supporting figures will be included where appropriate for clarity.

Assumptions
This scope assumes site access is available and that surveys can be scheduled during suitable seasonal and weather conditions. If agencies request additional studies beyond the tasks listed above, scope and budget updates may be provided for authorization.

Please contact ${project.preparedBy} with any questions regarding this scope.`;
}

function getProjectInfo() {
  return {
    projectName: els.projectName.value.trim(),
    client: els.client.value.trim(),
    location: els.location.value.trim(),
    preparedBy: els.preparedBy.value.trim(),
    date: els.scopeDate.value || todayIsoDate(),
  };
}

function validate(project, discipline, selectedTasks) {
  if (!discipline) {
    alert('Please select a discipline first.');
    return false;
  }

  const requiredFields = ['projectName', 'client', 'location', 'preparedBy'];
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

async function exportToWord() {
  if (!state.generatedScope) {
    return;
  }

  const discipline = getSelectedDiscipline();
  if (!discipline) {
    return;
  }

  const { Document, Packer, Paragraph } = window.docx;
  const paragraphs = state.generatedScope
    .split('\n\n')
    .flatMap((block) => block.split('\n'))
    .map((line) => new Paragraph({ text: line }));

  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });

  const blob = await Packer.toBlob(doc);
  const project = getProjectInfo();
  const fileName = buildFileName(project.projectName, discipline, project.date);
  saveAs(blob, fileName);
}

function handleGenerateScope() {
  const discipline = getSelectedDiscipline();
  const project = getProjectInfo();
  const selectedTasks = discipline ? getSelectedTasks(discipline) : [];

  if (!validate(project, discipline, selectedTasks)) {
    return;
  }

  state.generatedScope = buildScopeText(project, discipline, selectedTasks);
  els.scopePreview.textContent = state.generatedScope;
  els.exportBtn.disabled = false;
}

function init() {
  els.scopeDate.value = todayIsoDate();

  els.discipline.addEventListener('change', renderTaskSection);
  els.generateBtn.addEventListener('click', handleGenerateScope);
  els.exportBtn.addEventListener('click', () => {
    exportToWord().catch((error) => {
      console.error(error);
      alert('Unable to export document. Please try again.');
    });
  });

  loadTaskLibrary()
    .then(renderTaskSection)
    .catch((error) => {
      console.error(error);
      els.scopePreview.textContent = 'Error loading task library.';
    });
}

init();
