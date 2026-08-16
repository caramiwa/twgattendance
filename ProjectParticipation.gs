/**
 * Generates a project-level participation analysis from the currently active sheet.
 *
 * Source sheet structure:
 *   Column A: Project Code
 *   Column B: TWG Member
 *   Columns C onward: Activities
 *
 * Metrics per TWG member:
 *   Projects Invited       = distinct project codes where the member appears.
 *   Projects Participated  = invited projects with at least one Y.
 *   Project Participation  = Projects Participated / Projects Invited.
 *   Activities             = non-blank Y/N activity cells.
 *   Attended               = Y activity cells.
 *   Activity Attendance    = Attended / Activities.
 *
 * Blank cells are treated as not applicable and excluded from activity totals.
 * Unexpected activity values are flagged and excluded from the Y/N calculations.
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

  // Start at row 2; row 1 is assumed to contain headers.
  for (let r = 1; r < data.length; r++) {
    const projectCode = String(data[r][0] ?? '').trim();
    const member = String(data[r][1] ?? '').trim();

    if (!projectCode || !member) continue;

    if (!byMember[member]) {
      byMember[member] = {
        projects: new Map(),
        activities: 0,
        attended: 0,
        unexpected: 0
      };
    }

    const stats = byMember[member];

    // Keep one project record per member/project combination.
    if (!stats.projects.has(projectCode)) {
      stats.projects.set(projectCode, {
        attended: false,
        hasApplicableActivity: false
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
        project.attended = true;
        project.hasApplicableActivity = true;
      } else if (value === 'N') {
        stats.activities++;
        project.hasApplicableActivity = true;
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
    'Applicable Activities',
    'Attended Activities',
    'Activity Attendance %',
    'Avg. Attended Activities / Project',
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
      if (project.attended) projectsParticipated++;
    });

    const projectParticipation = projectsInvited > 0
      ? projectsParticipated / projectsInvited
      : 0;

    const activityAttendance = stats.activities > 0
      ? stats.attended / stats.activities
      : 0;

    const averageAttendedPerProject = projectsInvited > 0
      ? stats.attended / projectsInvited
      : 0;

    output.push([
      member,
      projectsInvited,
      projectsParticipated,
      projectParticipation,
      stats.activities,
      stats.attended,
      activityAttendance,
      averageAttendedPerProject,
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
    outputSheet.getRange(2, 7, output.length - 1, 1)
      .setNumberFormat('0.00%');
    outputSheet.getRange(2, 8, output.length - 1, 1)
      .setNumberFormat('0.00');
    outputSheet.getRange(2, 9, output.length - 1, 1)
      .setNumberFormat('0');
  }

  SpreadsheetApp.flush();
  return outputSheet.getName();
}
