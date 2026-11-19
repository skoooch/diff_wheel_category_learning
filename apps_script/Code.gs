// Google Apps Script web app: receives the experiment data and appends it to the bound Google Sheet.
// Setup: Extensions > Apps Script in a Google Sheet, paste this, then Deploy > Web app (Execute as: me, Access: anyone).
function doPost(e) {
  // Serialize concurrent submissions from many participants
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return handlePost(e);
  } finally {
    lock.releaseLock();
  }
}

// Each request carries only the trials recorded since the participant's previous save
function handlePost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var trials = ss.getSheetByName('trials') || ss.insertSheet('trials');
  if (trials.getLastRow() === 0) {
    trials.appendRow(['Subject_id', 'condition', 'frame', 'distance_to_boundary', 'correct', 'key_press',
                      'round', 'category_label', 'rt_ms', 'trial_index', 'date']);
  }
  // Learning trials only (practice trials have disp_type 'practice_trial')
  var rows = d.data
    .filter(function (t) { return t.disp_type === 'learning_trial'; })
    .map(function (t) {
      return [d.participant_id, d.condition, t.frame, t.distance_to_boundary, t.correct, t.key,
              t.round, t.category_label, t.rt, t.trial_index, d.date];
    });
  if (rows.length) {
    trials.getRange(trials.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Memory (mouse-reconstruction) trials: target frame vs. the frame the participant selected
  var mem = ss.getSheetByName('memory') || ss.insertSheet('memory');
  if (mem.getLastRow() === 0) {
    mem.appendRow(['Subject_id', 'condition', 'block', 'target_frame', 'response_frame', 'error', 'abs_error',
                   'distance_to_boundary', 'category_label', 'rt_ms', 'trial_index', 'date']);
  }
  var memRows = d.data
    .filter(function (t) { return t.disp_type === 'main_memory_trial'; })
    .map(function (t) {
      return [d.participant_id, d.condition, t.block, t.frame, t.response_frame, t.error, t.abs_error,
              t.distance_to_boundary, t.category_label, t.rt, t.trial_index, d.date];
    });
  if (memRows.length) {
    mem.getRange(mem.getLastRow() + 1, 1, memRows.length, memRows[0].length).setValues(memRows);
  }

  // Per-participant boundary info (boundary frames differ across participants in condition 1)
  var subj = ss.getSheetByName('subjects') || ss.insertSheet('subjects');
  if (subj.getLastRow() === 0) {
    subj.appendRow(['Subject_id', 'prolific_pid', 'study_id', 'session_id', 'condition', 'wheel', 'boundary_frames', 'semantic_boundary_frames', 'category1_side', 'date']);
  }
  // Written once per participant (every chunk repeats this info)
  var known = subj.getLastRow() > 1 ? subj.getRange(2, 1, subj.getLastRow() - 1, 1).getValues() : [];
  var seen = known.some(function (r) { return r[0] === d.participant_id; });
  if (!seen) {
    subj.appendRow([d.participant_id, d.prolific_pid, d.study_id, d.session_id, d.condition, d.wheel, JSON.stringify(d.boundary_frames),
                    JSON.stringify(d.semantic_boundary_frames), d.category1_side, d.date]);
  }

  // Raw payload chunks (demographics, browser info, interaction data in the final one) as a backup
  var raw = ss.getSheetByName('raw') || ss.insertSheet('raw');
  var json = JSON.stringify(d);
  var chunks = [d.participant_id, d.date];
  for (var i = 0; i < json.length; i += 40000) chunks.push(json.substr(i, 40000));
  raw.appendRow(chunks);

  return ContentService.createTextOutput('ok');
}
