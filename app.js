// --- STATE ---
const state = {
  disciplines: [],
  selectedTaskIds: new Set(),
};

// --- HELPERS ---
function todayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function sanitizeForFilename(input) {
  return input.trim().replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ');
}

function getSelectedDiscipline() {
  return state.disciplines.find(d => d.id === els.discipline.value) || null;
}

function getSelectedTasks(discipline) {
  return discipline.tasks.filter(t => state.selectedTaskIds.has(t.id));
}

// --- WORD EXPORT ENGINE ---
async function exportToWord() {
  const discipline = getSelectedDiscipline();
  if (!discipline) return alert("Select discipline");

  const project = getProjectInfo();
  const selectedTasks = getSelectedTasks(discipline);

  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    AlignmentType,
    TabStopType,
  } = window.docx;

  const children = [];

  // --- COVER LETTER ---

  // Date
  children.push(new Paragraph(project.date));

  // spacing
  children.push(new Paragraph(""));

  // Address block with right tab
  children.push(new Paragraph({
    children: [
      new TextRun(project.client),
      new TextRun("\t"),
      new TextRun("VIA EMAIL")
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
  }));

  children.push(new Paragraph({
    children: [
      new TextRun(project.contactName),
      new TextRun("\t"),
      new TextRun(project.contactEmail)
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
  }));

  children.push(new Paragraph(project.clientAddress));

  children.push(new Paragraph(""));

  // Subject line
  children.push(new Paragraph({
    children: [
      new TextRun({ text: "Subject:", bold: true }),
      new TextRun("\t"),
      new TextRun(project.subjectLine || project.projectName)
    ],
    tabStops: [{ type: TabStopType.LEFT, position: 2000 }]
  }));

  children.push(new Paragraph(""));

  // Greeting
  children.push(new Paragraph(`Dear ${project.contactName || "Client"}:`));
  children.push(new Paragraph(""));

  // Body
  children.push(new Paragraph(
    `Please find attached Attachment A, Scope of Work for biological consulting services for the ${project.projectName} project located in ${project.city}, ${project.state}.`
  ));

  children.push(new Paragraph(""));

  children.push(new Paragraph("We appreciate the opportunity to support your team."));
  children.push(new Paragraph(""));

  // Sincerely
  children.push(new Paragraph("Sincerely,"));
  children.push(new Paragraph(""));

  // Signature block (tab aligned)
  const sig1 = project.signatories[0] || {};
  const sig2 = project.signatories[1] || {};

  children.push(new Paragraph({
    children: [
      new TextRun(sig1.name || ""),
      new TextRun("\t"),
      new TextRun(sig2.name || "")
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
  }));

  children.push(new Paragraph({
    children: [
      new TextRun(sig1.title || ""),
      new TextRun("\t"),
      new TextRun(sig2.title || "")
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
  }));

  children.push(new Paragraph(""));

  children.push(new Paragraph("Attachment A – Scope of Work and Fees"));

  // --- PAGE BREAK ---
  children.push(new Paragraph({ pageBreakBefore: true }));

  // --- ATTACHMENT A HEADER ---
  children.push(new Paragraph({
    text: "ATTACHMENT A",
    alignment: AlignmentType.CENTER
  }));

  children.push(new Paragraph({
    text: "SCOPE OF WORK",
    alignment: AlignmentType.CENTER
  }));

  children.push(new Paragraph({
    text: "BIOLOGICAL CONSULTING SERVICES",
    alignment: AlignmentType.CENTER
  }));

  children.push(new Paragraph({
    text: project.projectName,
    alignment: AlignmentType.CENTER
  }));

  children.push(new Paragraph(""));
  children.push(new Paragraph(project.date));
  children.push(new Paragraph(""));

  // --- TASKS ---
  selectedTasks.forEach((task, i) => {

    // Task header row (left title + right fee)
    children.push(new Paragraph({
      children: [
        new TextRun({ text: `TASK ${i + 1} – ${task.name.toUpperCase()}`, bold: true }),
        new TextRun("\t"),
        new TextRun(task.fee || "")
      ],
      tabStops: [{ type: TabStopType.RIGHT, position: 9000 }]
    }));

    // Description
    children.push(new Paragraph(task.description));

    // Deliverables
    if (task.deliverables) {
      children.push(new Paragraph(`Deliverables: ${task.deliverables}`));
    }

    children.push(new Paragraph(""));
  });

  // --- BUILD DOC ---
  const doc = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(doc);

  saveAs(
    blob,
    `${sanitizeForFilename(project.projectName)}.docx`
  );
}