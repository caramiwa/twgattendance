/**
 * Generates a workload-contextualized TWG participation analysis.
 *
 * Source sheet structure:
 *   Column A: Project Code
 *   Column B: TWG Member
 *   Columns C onward: Activities
 *
 * This analysis does NOT produce a commitment score or a pass/fail recommendation.
 * It provides workload/exposure context alongside participation measures.
 *
 * Per member:
 *   Projects Invited       = distinct projects assigned.
 *   Projects Participated  = assigned projects with at least one Y.
 *   Project Participation  = Projects Participated / Projects Invited.
 *   Activities             = Y + N activity entries; blanks excluded.
 *   Attended               = Y entries.
 *   Not Attended           = N entries.
 *   Attendance             = Attended / Activities.
 *   Project Workload %ile  = percentile rank of project exposure among members.
 *   Activity Workload %ile = percentile rank of activity exposure among members.
 *
 * Population benchmarks are also written above the member table:
 *   Mean and median projects per member.
 *   Mean and median applicable activities per member.
 *
 * Percentile ranks are descriptive only. They do not imply that a higher
 * workload is better or worse. They simply show the member's relative exposure.
 *
 * Blank activity cells are treated as not applicable.
 */
function generateWorkloadAdjustedParticipationAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const sourceName = sourceSheet.getName();
  const data = sourceSheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error('The active sheet does not contain any data rows.');
  }

  const byMember = {};

  for (let r = 1; r < data.length; r++) {
    const projectCode = String(data[r][0] ?? '').trim();
    const member = String(data[r][1] ?? '').trim();

    if (!projectCode || !member) continue;

    if (!byMember[member]) {
      byMember[member] = {
        projects: new Map(),
        activities: 0,
        attended: 0,
        notAttended: 0
      };
    }

    const stats = byMember[member];

    if (!stats.projects.has(projectCode)) {
      stats.projects.set(projectCode, { participated: false });
    }

    const project = stats.projects.get(projectCode);

    for (let c = 2; c < data[r].length; c++) {
      const value = String(data[r][c] ?? '').trim().toUpperCase();

      if (value === '') continue;

      if (value === 'Y') {
        stats.activities++;
        stats.attended++;
        project.participated = true;
      } else if (value === 'N') {
        stats.activities++;
        stats.notAttended++;
      }
    }
  }

  const members = Object.keys(byMember).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  if (!members.length) {
    throw new Error('No valid TWG member/project records were found.');
  }

  const projectCounts = [];
  const activityCounts = [];

  members.forEach(member => {
    projectCounts.push(byMember[member].projects.size);
    activityCounts.push(byMember[member].activities);
  });

  const mean = values =>
    values.reduce((sum, value) => sum + value, 0) / values.length;

  const median = values => {
    const sorted = [...values].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2
      ? sorted[middle]
      : (sorted[middle - 1] + sorted[middle]) / 2;
  };

  // Percentile rank using the nearest-rank-style empirical proportion:
  // percentage of members with a workload less than or equal to this member.
  const percentileRank = (values, value) => {
    const count = values.filter(v => v <= value).length;
    return count / values.length;
  };

  const projectMean = mean(projectCounts);
  const projectMedian = median(projectCounts);
  const activityMean = mean(activityCounts);
  const activityMedian = median(activityCounts);

  const baseName = `Workload Participation - ${sourceName}`;
  let outputName = baseName;
  let counter = 2;

  while (ss.getSheetByName(outputName)) {
    outputName = `${baseName} (${counter})`;
    counter++;
  }

  const outputSheet = ss.insertSheet(outputName);

  // Population benchmark section.
  const benchmarkRows = [
    ['TWG WORKLOAD BENCHMARKS', ''],
    ['Members included', members.length],
    ['Mean projects per member', projectMean],
    ['Median projects per member', projectMedian],
    ['Mean applicable activities per member', activityMean],
    ['Median applicable activities per member', activityMedian],
    ['', ''],
    ['Interpretation note', 'Workload figures describe exposure only; they are not commitment scores.']
  ];

  outputSheet.getRange(1, 1, benchmarkRows.length, 2).setValues(benchmarkRows);
  outputSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  outputSheet.getRange(8, 1, 1, 2).setWrap(true);

  const headerRow = benchmarkRows.length + 2;
  const output = [[
    'TWG Member',
    'Projects Invited',
    'Projects Participated',
    'Project Participation %',
    'Activities',
    'Attended',
    'Not Attended',
    'Attendance %',
    'Project Workload Percentile',
    'Activity Workload Percentile'
  ]];

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
      percentileRank(projectCounts, projectsInvited),
      percentileRank(activityCounts, stats.activities)
    ]);
  });

  outputSheet.getRange(headerRow, 1, output.length, output[0].length).setValues(output);
  outputSheet.getRange(headerRow, 1, 1, output[0].length).setFontWeight('bold');
  outputSheet.setFrozenRows(headerRow);
  outputSheet.autoResizeColumns(1, output[0].length);

  if (output.length > 1) {
    outputSheet.getRange(headerRow + 1, 4, output.length - 1, 1).setNumberFormat('0.00%');
    outputSheet.getRange(headerRow + 1, 8, output.length - 1, 1).setNumberFormat('0.00%');
    outputSheet.getRange(headerRow + 1, 9, output.length - 1, 2).setNumberFormat('0.00%');
  }

  SpreadsheetApp.flush();
  return outputSheet.getName();
}
