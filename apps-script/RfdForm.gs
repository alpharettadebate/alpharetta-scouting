/*****************************************************************
 * CHAIN SCOUT: RFD NOTES FORM
 * Add this as a second file in the same Apps Script project as the Chain Indexer
 * (Files "+" > Script, name it RfdForm), paste, Save, then run createRfdForm once.
 *
 * It makes a Google Form with four sections (Round / The round / Your thoughts /
 * Judge feedback). Anyone with the link can fill it in on their phone, no Google
 * or Claude account needed. Answers land in the "Chain Scout RFD Notes" sheet in
 * the "Chain Index 2026-27" folder, and Chain Scout's RFDs tab picks them up at
 * every Monday and Thursday update.
 *****************************************************************/

function createRfdForm() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty('RFD_FORM_ID');
  if (existing) {
    Logger.log('Form already exists. Link to share: ' + shortLink_(FormApp.openById(existing)));
    return;
  }

  const form = FormApp.create('Chain Scout: RFD Notes');
  form.setDescription('One response per round. Fill in what you can during the round and the RFD; ' +
                      'skip anything you missed. Shows up on Chain Scout after the next update (Mon + Thu).');
  try { form.setCollectEmail(false); } catch (e) {}
  try { form.setRequireLogin(false); } catch (e) {}       // only matters on school accounts
  form.setAllowResponseEdits(true);
  form.setShowLinkToRespondAgain(true);
  form.setConfirmationMessage('Saved. It will be on Chain Scout after the next update.');

  // Section 1: the round
  form.addTextItem().setTitle('Tournament').setHelpText('e.g. Westminster, Trevian, Practice');
  form.addListItem().setTitle('Round').setRequired(true)
      .setChoiceValues(['R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'R7', 'Doubles', 'Octas', 'Quarters', 'Semis', 'Finals', 'Practice']);
  form.addTextItem().setTitle('Aff team').setHelpText('Team code, e.g. Alpharetta JE or WaRu GP');
  form.addTextItem().setTitle('Neg team').setHelpText('Team code, e.g. Alpharetta AS or Bronx Science BH');
  form.addTextItem().setTitle('Judge');
  form.addMultipleChoiceItem().setTitle('Decision').setChoiceValues(['Aff', 'Neg', 'Not disclosed']);
  form.addTextItem().setTitle('Your name').setHelpText('Who took these notes (optional)');

  // Section 2: what each side read
  form.addPageBreakItem().setTitle('The round').setHelpText('What each speech read. Skip any you missed.');
  form.addParagraphTextItem().setTitle('1AC').setHelpText('Aff name + advantages');
  form.addParagraphTextItem().setTitle('1NC').setHelpText('Off-case in order, plus case args');
  form.addParagraphTextItem().setTitle('Block').setHelpText('2NC + 1NR: what they extended');
  form.addParagraphTextItem().setTitle('1AR');
  form.addParagraphTextItem().setTitle('2NR').setHelpText('What the neg went for');
  form.addParagraphTextItem().setTitle('2AR').setHelpText('What the aff went for');

  // Section 3: our thoughts
  form.addPageBreakItem().setTitle('Your thoughts');
  form.addParagraphTextItem().setTitle('Thoughts').setHelpText('What went well, what to fix, what you would do differently');

  // Section 4: judge feedback
  form.addPageBreakItem().setTitle('Judge feedback + RFD');
  form.addParagraphTextItem().setTitle('Judge feedback / RFD').setHelpText('Why they voted, what they wanted to hear, speech-by-speech advice');

  // Responses sheet, kept next to the Chain Index sheets
  const ss = SpreadsheetApp.create('Chain Scout RFD Notes');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  const sheets = SpreadsheetApp.openById(ss.getId()).getSheets();
  const resp = sheets.find(s => /^Form Responses/i.test(s.getName()));
  if (resp) {                                              // make the responses tab the first (and only) tab
    SpreadsheetApp.openById(ss.getId()).setActiveSheet(resp);
    SpreadsheetApp.openById(ss.getId()).moveActiveSheet(1);
    sheets.filter(s => s.getSheetId() !== resp.getSheetId()).forEach(s => { try { SpreadsheetApp.openById(ss.getId()).deleteSheet(s); } catch (e) {} });
  }

  const root = getRoot_();                                 // from the Chain Indexer file
  DriveApp.getFileById(form.getId()).moveTo(root);
  DriveApp.getFileById(ss.getId()).moveTo(root);

  props.setProperty('RFD_FORM_ID', form.getId());
  props.setProperty('RFD_SHEET_ID', ss.getId());
  Logger.log('Done. Link to share with the team (and paste to Claude): ' + shortLink_(form));
}

function shortLink_(form) {
  const url = form.getPublishedUrl();
  try { return form.shortenFormUrl(url); } catch (e) { return url; }
}
