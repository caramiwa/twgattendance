/**
 * Generates an attendance summary from the currently active sheet.
 *
 * Source sheet structure:
 *   Column A: Project Code
 *   Column B: TWG Member
 *   Columns C onward: Activities
 *
 * Activity values:
 *   Y = attended
 *   N = not attended
 *   blank = not applicable (excluded from the activity denominator)
 *
 * The invitation/project count is based on distinct project codes per TWG member.
 * The activity count is based only on non-blank activity cells for that member.
 */
function generateAttendanceSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const sourceName = sourceSheet.getName();
  const data = sourceSheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error('The active sheet does not contain any data rows.');
  }

  const summaryByMember = {};

  // Start at row 2, assuming row 1 contains the headers.
  for (let r = 1; r < data.length; r++) {
    const projectCode = String(data[r][0] ?? '').trim();
    const member = String(data[r][1] ?? '').trim();

    // Ignore completely unidentified rows.
    if (!projectCode || !member) continue;

    if (!summaryByMember[member]) {
      summaryByMember[member] = {
        projects: new Set(),
        activities: 0,
        attended: 0,
        notAttended: 0,
        unexpected: 0
      };
    }

    const stats = summaryByMember[member];
    stats.projects.add(projectCode);

    // Activities begin at column C (index 2).
    for (let c = 2; c < data[r].length; c++) {
      const value = String(data[r][c] ?? '').trim().toUpperCase();

      // Blank means not applicable; do not count it as an activity.
      if (value === '') continue;

      if (value === 'Y') {
        stats.activities++;
        stats.attended++;
      } else if (value === 'N') {
        stats.activities++;
        stats.notAttended++;
      } else {
        // Do not silently classify unexpected values as Y or N.
        stats.unexpected++;
      }
    }
  }

  // Create a unique summary sheet name so repeated runs do not overwrite results.
  const baseName = `Attendance Summary - ${sourceName}`;
  let summaryName = baseName;
  let counter = 2;
  while (ss.getSheetByName(summaryName)) {
    summaryName = `${baseName} (${counter})`;
    counter++;
  }

  const summarySheet = ss.insertSheet(summaryName);

  const output = [[
    'TWG Member',
    'Projects Invited',
    'Activities',
    'Attended',
    'Not Attended',
    'Attendance %',
    'Unexpected Values'
  ]];

  const members = Object.keys(summaryByMember).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  members.forEach(member => {
    const stats = summaryByMember[member];
    const attendanceRate = stats.activities > 0
      ? stats.attended / stats.activities
      : 0;

    output.push([
      member,
      stats.projects.size,
      stats.activities,
      stats.attended,
      stats.notAttended,
      attendanceRate,
      stats.unexpected
    ]);
  });

  summarySheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // Basic formatting.
  summarySheet.getRange(1, 1, 1, output[0].length)
    .setFontWeight('bold');
  summarySheet.setFrozenRows(1);
  summarySheet.getRange(2, 6, Math.max(output.length - 1, 1), 1)
    .setNumberFormat('0.00%');
  summarySheet.autoResizeColumns(1, output[0].length);

  // Make unexpected values visible without changing the calculations.
  if (output.length > 1) {
    const unexpectedRange = summarySheet.getRange(2, 7, output.length - 1, 1);
    unexpectedRange.setNumberFormat('0');
  }

  SpreadsheetApp.flush();
  return summarySheet.getName();
}
