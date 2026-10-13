// Site settings. This is the only file you should need to edit.
window.SCOUTING = {
  // Paste the Apps Script web app URL (ends in /exec) between the quotes.
  // Leave it empty and the site still works from the saved snapshot in /data.
  apiUrl: "",

  // Google Form the team uses to take notes during RFDs.
  formUrl: "https://forms.gle/ivYx9G15n152hwMGA",

  // Saved snapshot of everything indexed before the live pipeline (Oct 5, 2026).
  snapshotDocs: ["data/docs-2026-10-05.json"],
  snapshotOutlines: [
    "data/outlines-2026-10-05-1.json",
    "data/outlines-2026-10-05-2.json",
    "data/outlines-2026-10-05-3.json",
    "data/outlines-2026-10-05-4.json"
  ],

  // Team names that count as "us" on the Our Rounds page.
  ourPrefix: "Alpharetta"
};
