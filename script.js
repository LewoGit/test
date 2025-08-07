const GEMINI_API_KEY = "AIzaSyAbiZUIOUxqDdcPPvWs9UtPahq37h6HjqA";

function generateReportsFromSheet() {
  const sheet = SpreadsheetApp.openById("14OJQ0tORMrO1RKu_dZV16y2DYg8SQE9U3aZpuk4DBVY").getSheets()[0];
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const timestamp = row[0];
    const candidateName = row[1];
    const candidateEmail = row[2];
    const stage1Response = row[3];
    const stage2Response = row[4];
    const stage3Response = row[5];
    const stage4Response = row[6];
    const stage5Response = row[7];
    const stage6Response = row[8];

    const prompt = `
You are an expert AI trained in SACPCMP professional registration assessments. You have access to the Candidate Self-Assessment Framework (2025) and the full example report: https://github.com/216062607/gggggg/blob/main/Feedback%20Report%20for%20Selebaleng%20%20Ndwandwe%20Rev02.pdf.

Your task is to generate a professional feedback report identical in structure, tone, and table format to the one above, based on the candidate’s responses.

Candidate Name: ${candidateName}

Responses:
Stage 1 – Project Initiation and Briefing:
${stage1Response}

Stage 2 – Concept and Feasibility:
${stage2Response}

Stage 3 – Design Development:
${stage3Response}

Stage 4 – Tender Documentation and Procurement:
${stage4Response}

Stage 5 – Construction Documentation and Management:
${stage5Response}

Stage 6 – Project Close Out:
${stage6Response}

Output Requirements:
- Start with a “General Feedback” section.
- Then for each project stage, produce a table with:
  • Areas of Challenge  
  • Indicative Period to Address Gaps  
  • Recommended Training Resources  
  • Recommended Practical Experience  
  • Evaluation Method  
  • Additional Comments
- Follow with a paragraph titled "General Comments for Stage X"
- Repeat this table structure for each of the 6 stages
- Match the layout and format of the Word document template exactly
- Use structured plain text (no HTML)
`;

    const aiGeneratedResponse = callGeminiAPI(prompt);
    const pdfBlob = generatePDFUsingTemplate(candidateName, aiGeneratedResponse);

    MailApp.sendEmail({
      to: "lewokwetsi@gmail.com",
      subject: `SACPCMP Feedback Report – ${candidateName}`,
      body: `Dear Team,\n\nPlease find attached the SACPCMP Candidate Feedback Report for ${candidateName}.\n\nKind regards,\nAssessment System`,
      attachments: [pdfBlob]
    });
  }
}

function callGeminiAPI(userPrompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

  const payload = JSON.stringify({
    contents: [
      {
        parts: [
          { text: userPrompt }
        ]
      }
    ]
  });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.candidates && result.candidates.length > 0) {
      return result.candidates[0].content.parts[0].text;
    } else {
      return "AI response failed. No content returned.";
    }
  } catch (error) {
    Logger.log("Error calling Gemini API: " + error);
    return "An error occurred while contacting Gemini API.";
  }
}

function generatePDFUsingTemplate(candidateName, reportText) {
  const folderName = "Candidate Feedback Reports";
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  const templateId = "10nk0JsSCuBmgJZGqNF8hpXOwaqr858i0lDl6uY1AbsE"; // replace with actual ID from uploaded docx
  const template = DriveApp.getFileById(templateId).makeCopy(`Feedback Report for ${candidateName}`);
  const doc = DocumentApp.openById(template.getId());
  const body = doc.getBody();

  body.replaceText("{{CandidateName}}", candidateName);
  body.replaceText("{{ReportContent}}", reportText);

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs("application/pdf");
  folder.createFile(pdf);
  return pdf;
}