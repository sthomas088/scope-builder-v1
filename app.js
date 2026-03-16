const state = {
  tasks: [],
  selectedTaskIds: new Set(),
  generatedScope: '',
};

const els = {
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

async function loadTasks() {
  const response = await fetch('tasks.json');
  if (!response.ok) {
    throw new Error('Unable to load tasks library.');
  }

  state.tasks = await response.json();
  renderTaskList();
}

function renderTaskList() {
  els.taskList.innerHTML = '';

  state.tasks.forEach((task) => {
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

function getSelectedTasks() {
  return state.tasks.filter((task) => state.selectedTaskIds.has(task.id));
}

function buildScopeText(project, selectedTasks) {
  const taskParagraphs = selectedTasks
    .map((task, index) => `${index + 1}. ${task.name}: ${task.description}`)
    .join('\n\n');

  return `BIOLOGICAL CONSULTING SCOPE OF WORK

Project: ${project.projectName}
Client: ${project.client}
Location: ${project.location}
Prepared By: ${project.preparedBy}
Date: ${project.date}

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

function validate(project, selectedTasks) {
  const requiredFields = ['projectName', 'client', 'location', 'preparedBy'];
  const missing = requiredFields.filter((field) => !project[field]);

  if (missing.length > 0) {
    alert('Please complete all project information fields before generating scope.');
    return false;
  }

  if (selectedTasks.length === 0) {
    alert('Please select at least one biology task.');
    return false;
  }

  return true;
}

function buildFileName(projectName, date) {
  const cleanProjectName = sanitizeForFilename(projectName) || 'Project';
  return `${cleanProjectName}_BioScope_${date}.docx`;
}

async function exportToWord() {
  if (!state.generatedScope) {
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
  const fileName = buildFileName(project.projectName, project.date);
  saveAs(blob, fileName);
}

function handleGenerateScope() {
  const project = getProjectInfo();
  const selectedTasks = getSelectedTasks();

  if (!validate(project, selectedTasks)) {
    return;
  }

  state.generatedScope = buildScopeText(project, selectedTasks);
  els.scopePreview.textContent = state.generatedScope;
  els.exportBtn.disabled = false;
}

function init() {
  els.scopeDate.value = todayIsoDate();

  els.generateBtn.addEventListener('click', handleGenerateScope);
  els.exportBtn.addEventListener('click', () => {
    exportToWord().catch((error) => {
      console.error(error);
      alert('Unable to export document. Please try again.');
    });
  });

  loadTasks().catch((error) => {
    console.error(error);
    els.scopePreview.textContent = 'Error loading task library.';
  });
}

init();
