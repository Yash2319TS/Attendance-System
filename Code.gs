/**
 * Attendance API — bind this script to your Google Sheet
 * (Extensions → Apps Script) and deploy it as a Web App.
 *
 * Sheet layout required:
 *
 *   Tab "Students"          Tab "Logs"
 *   ┌─────────────┬───────┐ ┌───────────┬───────┬───────┬────────┐
 *   │ RollNo      │ Name  │ │ Timestamp │ Roll  │ Name  │ Status │
 *   ├─────────────┼───────┤ ├───────────┼───────┼───────┼────────┤
 *   │ 230151520053│ Yash  │ │ (auto)    │(auto) │(auto) │(auto)  │
 *   └─────────────┴───────┘ └───────────┴───────┴───────┴────────┘
 *
 *   Fill "Students" once with your class roster.
 *   "Logs" fills itself — every fingerprint scan calls doPost() below,
 *   which appends one row per scan. Don't type into Logs by hand.
 */

const STUDENTS_SHEET = "Students";
const LOGS_SHEET = "Logs";
const CUTOFF_TIME = "09:15"; // scans after this time are marked "late"

// ── Called by the fingerprint device / gateway on every scan ───────────
// POST body: { "roll": "230151520053" }
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const roll = String(body.roll || "").trim();
  if (!roll) return json({ ok: false, error: "roll is required" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName(STUDENTS_SHEET);
  const logsSheet = ss.getSheetByName(LOGS_SHEET);

  const students = studentsSheet.getDataRange().getValues();
  const match = students.find((r) => String(r[0]).trim() === roll);
  const name = match ? match[1] : "Unknown";

  const now = new Date();
  const status = statusFor(now);
  logsSheet.appendRow([now, roll, name, status]);

  return json({ ok: true, roll, name, status, time: now.toISOString() });
}

// ── Called by the dashboard to read summarised attendance ──────────────
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = ss
    .getSheetByName(STUDENTS_SHEET)
    .getDataRange()
    .getValues()
    .slice(1);
  const logs = ss
    .getSheetByName(LOGS_SHEET)
    .getDataRange()
    .getValues()
    .slice(1);

  const today = startOfDay(new Date());
  const schoolDays = new Set(
    logs.map((l) => startOfDay(new Date(l[0])).getTime()),
  );

  const result = students
    .filter((r) => r[0] !== "")
    .map(([roll, name]) => {
      const mine = logs
        .filter((l) => String(l[1]).trim() === String(roll).trim())
        .sort((a, b) => new Date(b[0]) - new Date(a[0]));

      const todayLog = mine.find(
        (l) => startOfDay(new Date(l[0])).getTime() === today.getTime(),
      );

      const week = [];
      for (let i = 6; i >= 0; i--) {
        const day = new Date(today);
        day.setDate(day.getDate() - i);
        const log = mine.find(
          (l) => startOfDay(new Date(l[0])).getTime() === day.getTime(),
        );
        week.push(log ? log[3] : "absent");
      }

      const presentDays = new Set(
        mine
          .filter((l) => l[3] === "present" || l[3] === "late")
          .map((l) => startOfDay(new Date(l[0])).getTime()),
      );
      const attendance = schoolDays.size
        ? Math.round((presentDays.size / schoolDays.size) * 100)
        : 0;

      return {
        roll: String(roll),
        name,
        today: todayLog ? todayLog[3] : "absent",
        lastScan: mine[0]
          ? Utilities.formatDate(
              new Date(mine[0][0]),
              Session.getScriptTimeZone(),
              "hh:mm a",
            )
          : "—",
        week,
        attendance,
      };
    });

  return json(result);
}

function statusFor(date) {
  const [h, m] = CUTOFF_TIME.split(":").map(Number);
  const cutoff = new Date(date);
  cutoff.setHours(h, m, 0, 0);
  return date <= cutoff ? "present" : "late";
}

function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
