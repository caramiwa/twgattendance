/**
 * Generates a project-level participation analysis from the currently active sheet.
 *
 * Source sheet structure:
 *   Column A: Project Code
 *   Column B: TWG Member
 *   Columns C onward: Activities
 *
 * The report puts activity attendance in the context of project exposure.
 * It does NOT calculate or label a "commitment score."
 *
 * Metrics per TWG member:
 *   Projects Invited       = distinct projects assigned to the member.
 *   Projects Participated  = assigned projects with at least one Y.
 *   Project Participation  = Projects Participated / Projects Invited.
 *   Activities             = applicable activities (Y + N); blanks excluded.
 *   Attended               = Y activity cells.
 *   Not Attended           = N activity cells.
 *   Attendance             = Attended / Activities.
 *
 * Blank cells are treated as not applicable.
 * Unexpected activity values are counted separately and excluded from Y/N metrics.
 *
 * The first three activity metrics (Activities, Attended, Not Attended) are
 * deliberately kept consistent with the activity-level attendance analysis
 * so this report can be checked against generateAttendanceSummary().
 */
function generateProjectParticipationAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const sourceName = sourceSheet.getName();
  const data = sourceSheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error('The active sheet does not contain any data rows.');
  }

  const byMember = {};

  // Row 1 is assumed to contain headers.
  for (let r = 1; r < data.length; r++) {
    const projectCode = String(data[r][0] ?? '').trim();
    const member = String(data[r][1] ?? '').trim();

    if (!projectCode || !member) continue;

    if (!byMember[member]) {
      byMember[member] = {
        projects: new Map(),
        activities: 0,
        attended: 0,
        notAttended: 0,
        unexpected: 0
      };
    }

    const stats = byMember[member];

    if (!stats.projects.has(projectCode)) {
      stats.projects.set(projectCode, {
        participated: false
      });
    }

    const project = stats.projects.get(projectCode);

    // Activities begin at column C (index 2).
    for (let c = 2; c < data[r].length; c++) {
      const value = String(data[r][c] ?? '').trim().toUpperCase();

      // Blank = not applicable.
      if (value === '') continue;

      if (value === 'Y') {
        stats.activities++;
        stats.attended++;
        project.participated = true;
      } else if (value === 'N') {
        stats.activities++;
        stats.notAttended++;
      } else {
        stats.unexpected++;
      }
    }
  }

  const baseName = `Project Participation - ${sourceName}`;
  let outputName = baseName;
  let counter = 2;

  while (ss.getSheetByName(outputName)) {
    outputName = `${baseName} (${counter})`;
    counter++;
  }

  const outputSheet = ss.insertSheet(outputName);

  const output = [[
    'TWG Member',
    'Projects Invited',
    'Projects Participated',
    'Project Participation %',
    'Activities',
    'Attended',
    'Not Attended',
    'Attendance %',
    'Unexpected Values'
  ]];

  const members = Object.keys(byMember).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  members.forEach(member => {
    const stats = byMember[member];
    const projectsInvited = stats.projects.size;
    let projectsParticipated = 0;

    stats.projects.forEach(project => {
      if (project.participated) projectsParticipated++;
    });

    const projectParticipation = projectsInvited > 0
      ? projectsParticipated / projectsInvited
      : 0;

    const attendance = stats.activities > 0
      ? stats.attended / stats.activities
      : 0;

    output.push([
      member,
      projectsInvited,
      projectsParticipated,
      projectParticipation,
      stats.activities,
      stats.attended,
      stats.notAttended,
      attendance,
      stats.unexpected
    ]);
  });

  outputSheet
    .getRange(1, 1, output.length, output[0].length)
    .setValues(output);

  outputSheet.getRange(1, 1, 1, output[0].length)
    .setFontWeight('bold');
  outputSheet.setFrozenRows(1);
  outputSheet.autoResizeColumns(1, output[0].length);

  if (output.length > 1) {
    outputSheet.getRange(2, 4, output.length - 1, 1)
      .setNumberFormat('0.00%');
    outputSheet.getRange(2, 8, output.length - 1, 1)
      .setNumberFormat('0.00%');
    outputSheet.getRange(2, 9, output.length - 1, 1)
      .setNumberFormat('0');
  }

  SpreadsheetApp.flush();
  return outputSheet.getName();
}
