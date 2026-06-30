
// 1. SETTINGS
const SHEET_ID = '1-5Fc8FA_a4RLi9Z8H65-TwQ52Wsp2kqWpfLXeCmH5yg'; // Copy from the URL between /d/ and /edit
const SOURCE_TAB = 'PEVC V1';      // The tab where you send emails from
const TARGET_TAB = 'PEVC_Sent';  // The tab where "Sent" rows are moved to
const PROMPT_CHOICE = 1; // Change to 2 to use the second prompt version

const GEMINI_API_KEY = 'Gemini API key'; // Replace with your Gemini API key

const OUR_SERVICES = 'We provide plug-and-play dedicated Analyst bandwidth that works as an extension of your team without the cost, time, or commitment of hiring in-house. Our dedicated virtual analysts handle end-to-end execution across prospect research and identification, including ideal customer research, account and contact discovery, and building targeted outreach-ready lead lists. We also manage CRM operations such as cleaning messy data, contact enrichment, and maintaining accuracy and organization. In addition, we execute full outreach operations including email campaigns, LinkedIn outreach, list uploads, sequencing, tracking, and complete campaign management on your behalf. We do not just support your team, we execute. There is no recruitment, no training overhead, and no long-term lock-in. You simply plug us in and we start delivering. Our flexible engagement model allows you to choose 40, 80, or 160 hours and scale as needed, paying only for the hours worked. This is ideal for companies that want to run outbound without building a full team, manage or fix their CRM without hiring operations staff, scale outreach campaigns quickly, or offload time-consuming sales operations work.'; // full services text

// 2. MENU
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Sales AI')
    .addItem('Run Automation', 'startSalesAgent')
    .addItem('Move Sent Rows', 'moveAllExistingSentRows')
    .addToUi();
}

// 3. ENGINE WITH LOCK
function startSalesAgent() {

  console.log("=== Script Started ===");

  const lock = LockService.getScriptLock();

  try {

    // Prevent overlapping runs
    lock.tryLock(30000);

    console.log("Opening Spreadsheet...");

    const ss = SpreadsheetApp.openById(SHEET_ID);

    console.log("Spreadsheet Opened");

    const sheet = ss.getSheetByName(SOURCE_TAB);

    if (!sheet) {
      console.log("Sheet Not Found: " + SOURCE_TAB);
      return;
    }

    console.log("Sheet Loaded: " + SOURCE_TAB);

    const data = sheet.getDataRange().getValues();

    console.log("Total Rows Found: " + data.length);

    let batch = [];
    let rowIndexes = [];

    // LOOP THROUGH ROWS
    for (let i = 1; i < data.length; i++) {

      console.log("Checking Row: " + (i + 1));

      let [
        name,
        role,
        url,
        email,
        status,
        columnFValue // <-- Added to grab Column F (Index 5)
      ] = data[i];

      // SKIP CONDITIONS
      // NEW FILTER: Skip if Column F is NOT "No Response"
      if (!columnFValue || columnFValue.toString().trim() !== "No Response") {
        console.log("Skipping Row (Not 'No Response' in Col F): " + (i + 1));
        continue;
      }

      // SKIP CONDITIONS
      if (status === "Sent") {
        console.log("Already Sent Row: " + (i + 1));
        continue;
      }

      if (!url || !email) {
        console.log("Missing URL or Email Row: " + (i + 1));
        continue;
      }

      let cleanText = "";

      // FETCH LIGHTWEIGHT WEBSITE DATA
      try {

        console.log("Fetching Website: " + url);

        const response = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: false
        });

        const responseCode = response.getResponseCode();

        console.log("Response Code: " + responseCode);

        // BAD RESPONSE
        if (responseCode >= 400) {

          console.log("Bad Website Response Row: " + (i + 1));

          cleanText = `
            Company Website: ${url}
            Role: ${role}
          `;

        } else {

          // ONLY FETCH SMALL PART OF HTML
          const html = response
            .getContentText()
            .substring(0, 30000);

          // TITLE
          const titleMatch = html.match(/<title>(.*?)<\/title>/i);

          const title = titleMatch
            ? titleMatch[1]
            : "";

          // META DESCRIPTION
          const metaMatch = html.match(
            /<meta[^>]*name=["']description["'][^>]*content=["']([^"]*)["']/i
          );

          const metaDescription = metaMatch
            ? metaMatch[1]
            : "";

          cleanText = `
            Title: ${title}

            Description: ${metaDescription}

            Website: ${url}
          `;

          console.log("Website Fetch Success Row: " + (i + 1));
        }

      } catch (fetchError) {

        console.log("Website Fetch Failed Row: " + (i + 1));

        console.log(fetchError.toString());

        // FALLBACK DATA
        cleanText = `
          Company Website: ${url}
          Role: ${role}
        `;
      }

      // ADD TO BATCH
      batch.push({
        name: name,
        role: role,
        email: email,
        cleanText: cleanText
      });

      rowIndexes.push(i);

      console.log("Added To Batch Row: " + (i + 1));

      // PROCESS EVERY 10
      if (batch.length === 10) {

        try {

          console.log("Processing Batch...");

          processBatch(batch, rowIndexes, sheet);

          console.log("Batch Processed Successfully");

        } catch (batchError) {

          console.log("Batch Error: " + batchError.toString());
        }

        // RESET
        batch = [];
        rowIndexes = [];

        SpreadsheetApp.flush();

        Utilities.sleep(2000);
      }
    }

    // PROCESS REMAINING RECORDS
    if (batch.length > 0) {

      try {

        console.log("Processing Final Batch Size: " + batch.length);

        processBatch(batch, rowIndexes, sheet);

        console.log("Final Batch Completed");

      } catch (finalBatchError) {

        console.log("Final Batch Error: " + finalBatchError.toString());
      }
    }

    console.log("=== Script Finished Successfully ===");

  } catch (mainError) {

    console.log("MAIN SCRIPT ERROR: " + mainError.toString());

  } finally {

    lock.releaseLock();

    console.log("Lock Released");
  }
}


// 4. PROCESS BATCH
function processBatch(batch, rowIndexes, sheet) {
  try {
    // Logic to choose which prompt function to run
    let prompt;
    if (PROMPT_CHOICE === 1) {
      prompt = buildPromptV1(batch);
    } else {
      prompt = buildPromptV2(batch);
    }

    const aiResponse = callGemini(prompt);

    aiResponse.forEach((resp, idx) => {
      const { subject, body } = resp;
      const email = batch[idx].email;
      
    // Change: Wrap in try/catch to ensure Sheet only updates if Gmail succeeds
      try {
        // Step 1: Send Email FIRST
        GmailApp.sendEmail(email, subject, body);
        
        const rowNum = rowIndexes[idx] + 1;
        
        // Step 2: Update Sheet ONLY after successful send
        sheet.getRange(rowNum, 5).setValue("Sent");
        
        // Step 3: Format Date correctly (M/d/yyyy)
        const dateOnly = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
        sheet.getRange(rowNum, 7).setValue(dateOnly);
        
        // Force the sheet to save this row immediately in case of timeout later
        SpreadsheetApp.flush(); 

      } catch (e) {
        // If Gmail fails, mark the error so you know to retry
        const rowNum = rowIndexes[idx] + 1;
        sheet.getRange(rowNum, 5).setValue("Failed: " + e.message);
      }
      
    });

    SpreadsheetApp.flush();
  } catch (e) {
    console.error("Batch processing failed: " + e.toString());
  }
}

// 5. PROMPT BUILDER
function buildPromptV1(batch) {
  return `You are an expert B2B copywriter generating personalized cold outreach emails.

Your task is to write a personalized cold email for each prospect provided in the data.

---
STRICT OUTPUT FORMAT
You must return ONLY a raw JSON array. Do NOT wrap the response in markdown code blocks (e.g., do not use \`\`\`json). No conversational filler, no introductory text. The output must be immediately parsable by JSON.parse().

Each object in the JSON array must follow this exact schema:
{
  "subject": "Email subject line",
  "body": "Email body text. Use \\n\\n for all paragraph line breaks."
}

---
INPUT DATA
${JSON.stringify(batch)}

---
OUR SERVICES & CAPABILITIES (Use this to match our support to their needs)
We provide plug-and-play dedicated Analyst bandwidth that works as an extension of their team without the cost, time, or commitment of hiring in-house. 
- What we do: End-to-end execution across prospect research, account/contact discovery, targeted lead-list building, CRM operations (cleaning messy data, contact enrichment), and full outreach operations (email campaigns, LinkedIn outreach, sequencing, and tracking).
- How we support them: We don't just support, we execute. No recruitment or training overhead.

---
INSTRUCTIONS & MESSAGING FRAMEWORK

STEP 1: ANALYZE THE PROSPECT (Mandatory Internal Step)
Carefully read the website content ("cleanText") for each prospect to identify:
1. Exactly what their company does.
2. Who their target B2B audience/decision-makers are.

STEP 2: SUBJECT LINE RULES
- Must clearly mention their target audience.
- Must let them know we can identify and reach out to them on their behalf to free up their bandwidth for strategic work.
- Example: "Sourcing [Target Audience] + outreach & CRM support"

STEP 3: WRITE THE EMAIL BODY (WHAT WE DO & BENEFITS)
Draft a short, compelling email body that addresses exactly how we support them using these guidelines:
- Connection: Open by stating that we help companies like theirs build targeted lists and handle outreach to their specific target audience (identified in Step 1).
- What We Do & How We Support: Clearly mention that we provide plug-and-play dedicated Analyst bandwidth acting as an extension of their team to execute prospect research, list building, CRM cleanup, and email/LinkedIn outreach.
- The Core Benefits: Emphasize that they get full execution without the cost, time, or commitment of hiring in-house, and zero recruitment/training overhead.
- Pricing/Offer Lines: 
  * Line 1: Let them know they can start with a 40 hours pilot and scale as needed.
  * Line 2: Let them know they can pause anytime with just one week notice.
  * Line 3: Let them know our Analyst can work on any custom tasks.
- Keep these points crisp and punchy.

STEP 4: CALL TO ACTION (CTA)
- Use a low-friction CTA.
- Ask a simple yes/no question about a quick walkthrough in the coming days.

STEP 5: MANDATORY SIGNATURE
You must include this exact signature block at the end of every email body:
Regards,
Gayatri Lokhande
bizkonnect.com

Not interested? Reply 'Unsubscribe'.

---
TARGETING & STYLE RULES (STRICT)
- Only B2B (Target companies, teams, decision-makers).
- NEVER mention individuals or consumers.
- No intent-based phrases like "people looking for".
- If the company is B2C, target the business operators behind it.`;
}

// 5. PROMPT BUILDER
function buildPromptV(batch) {
  return `

Generate personalized cold emails for the following prospects.

Return ONLY a JSON array where each element has keys "subject" and "body".
Use "\\n\\n" for line breaks in the body.

Return ONLY raw JSON. Do not include any introductory text, markdown code blocks, or conversational filler. The output must be parsable by JSON.parse().

INPUT FORMAT

Each prospect includes:

Name
Role
Website content (cleanText)

STEP 1: ANALYZE (MANDATORY)

From the website content, internally identify:

What the company offers
Their target audience (STRICTLY B2B)
Industry / company type
Likely decision-makers they sell to
How they likely generate pipeline (sales-led, outbound, inbound, PLG, etc.)

Do NOT output this analysis, but use it to personalize the email.

STEP 2: WRITE EMAIL

Follow this structure:(MANDATORY and use proper spacing)

Greeting
Hi {{Name}},

Opener (1 line, question, max 20 words)
Must reflect a real observation from their business model, ICP, or offering.

Context(Max 20 words)
Show you understand how their team likely builds pipeline or runs outreach.
Keep it specific and relevant to their business model

Introduction (ONE-LINER((Max 20 words)) – MANDATORY)

After the context, add one single sentence that:

Introduces dedicated Plug and play analyst support (real human resource)
Positions them as working as vitual extended team member.
Implies no hiring/training, pause anytime.
Connects to execution work (lists, CRM, outreach)

Services (numbered, tailored)

Prospect research & ICP mapping: Tailor to their actual target companies and industries.
Account & contact discovery: Mention relevant job roles or departments.
CRM cleaning & enrichment: Focus on cleaning, enrichment, and maintaining accurate data.
Outreach execution: Mention email + LinkedIn workflows aligned to their motion.

Demonstration (VERY IMPORTANT)(Max 20 words)
Write 1 short paragraph (max 2 sentences) that:

Demonstrates how analysts would:
Identify and research target companies (based on their ICP)
Build and maintain outreach-ready lists
Keep CRM data clean and updated
Execute email + LinkedIn outreach
Mention:
Specific company types / industries
Relevant job roles or decision-makers
End by showing:
→ Their team can stay focused on closing deals and strategic work

USP (MANDATORY)

Write 3 short lines (numbered):

Line 1 → Let them know they can start with 40 hours pilot and scale as needed
Line 2 → Let them know they can pause anytime with just one week notice
Line 3 → Let them know Our Analyst can work on any custom tasks

Keep them crisp, phrase-style (not full heavy sentences).

CTA (low friction)
Ask a simple yes/no question about a quick walkthrough in the coming days.

Signature (MANDATORY)
Regards,
Gayatri Lokhande
bizkonnect.com
Not interested? Reply 'Unsubscribe'.

SUBJECT LINE

Must:

Clearly mention their target audience and let them know we can identify and reachout them on their behalf to free up your bandwidth for strategic work.

Example:
"Sourcing [Target Audience] + outreach & CRM support"



TARGETING RULES (STRICT)
Only B2B (companies, teams, decision-makers)
Never mention individuals or consumers
No intent-based phrases like “people looking for”
If B2C → target business operators behind it

STYLE RULES
Max 2 sentences per paragraph
Use line breaks for readability
No fluff or praise
No generic lines
No promises of deals or clients
Keep tone direct and peer-level
Write for fast scanning (F-pattern reading)

FINAL OUTPUT RULES (STRICT)
Use ONLY "\n\n" for spacing (NEVER use <br> or HTML)
Do NOT skip the signature — always end with:
Regards,
Gayatri Lokhande
bizkonnect.com
Not interested? Reply 'Unsubscribe'.

Keep all sections separate with proper spacing (no merging)
If any rule is violated, regenerate before returning output

Prospects:
${batch.map((p, idx) => 
  `${idx+1}. Name: ${p.name}, Role: ${p.role}, Website: ${p.cleanText}`
).join("\n")}
`;
}

// 5. PROMPT BUILDER
function buildPromptV2(batch) {
  return `

Generate personalized cold emails for the following prospects.

Return ONLY a JSON array where each element has keys "subject" and "body".
Use "\\n\\n" for line breaks in the body.

Return ONLY raw JSON. Do not include any introductory text, markdown code blocks, or conversational filler. The output must be parsable by JSON.parse().

INPUT FORMAT

Each prospect includes:

Name
Role
Website content (cleanText)

STEP 1: ANALYZE (MANDATORY)

From the website content, internally identify:

What the company offers
Their target audience (STRICTLY B2B)
Industry / company type
Likely decision-makers they sell to
How they likely generate pipeline (sales-led, outbound, inbound, PLG, etc.)

Do NOT output this analysis, but use it to personalize the email.

STEP 2: WRITE EMAIL

Follow this structure:(MANDATORY and use proper spacing)

Greeting
Hi {{Name}},

Opener (1 line, question, max 20 words)
Must reflect a real observation from their business model, ICP, or offering.

Context(Max 20 words)
Show you understand how their team likely builds pipeline or runs outreach.
Keep it specific and relevant to their business model

Introduction (ONE-LINER((Max 20 words)) – MANDATORY)

After the context, add one single sentence that:

Introduces dedicated Plug and play analyst support (real human resource)
Positions them as working as vitual extended team member.
Implies no hiring/training, pause anytime.
Connects to execution work (lists, CRM, outreach)

Services (numbered, tailored)

Prospect research & ICP mapping: Tailor to their actual target companies and industries.
Account & contact discovery: Mention relevant job roles or departments.
CRM cleaning & enrichment: Focus on cleaning, enrichment, and maintaining accurate data.
Outreach execution: Mention email + LinkedIn workflows aligned to their motion.



USP (MANDATORY)

Write 3 short lines (numbered):

1. Let them know they can start with 40 hours pilot and scale as needed
2. Let them know they can pause anytime with just one week notice
3. Let them know Our Analyst can work on any custom tasks

Keep them crisp, phrase-style (not full heavy sentences).

CTA (low friction)
Ask a simple yes/no question about a quick walkthrough in the coming days.

Signature (MANDATORY)
Regards,
Gayatri Lokhande
bizkonnect.com
Not interested? Reply 'Unsubscribe'.

SUBJECT LINE

Must:

Clearly mention their target audience and let them know we can identify and reachout them on their behalf to free up your bandwidth for strategic work.

Must have "Following Up:"

Example:
"Following up: Sourcing [Target Audience] + outreach & CRM support"


TARGETING RULES (STRICT)
Only B2B (companies, teams, decision-makers)
Never mention individuals or consumers
No intent-based phrases like “people looking for”
If B2C → target business operators behind it

STYLE RULES
Max 2 sentences per paragraph
Use line breaks for readability
No fluff or praise
No generic lines
No promises of deals or clients
Keep tone direct and peer-level
Write for fast scanning (F-pattern reading)

FINAL OUTPUT RULES (STRICT)
Use ONLY "\n\n" for spacing (NEVER use <br> or HTML)
Do NOT skip the signature — always end with:
Regards,
Gayatri Lokhande
bizkonnect.com
Not interested? Reply 'Unsubscribe'.

Keep all sections separate with proper spacing (no merging)
If any rule is violated, regenerate before returning output

Prospects:
${batch.map((p, idx) => 
  `${idx+1}. Name: ${p.name}, Role: ${p.role}, Website: ${p.cleanText}`
).join("\n")}
`;
}

// 6. GEMINI CALL
function callGemini(prompt) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=" + GEMINI_API_KEY;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
  // This line will show you the EXACT error in the Execution Log
  console.error("Gemini API Error Detail:", JSON.stringify(json.error, null, 2));

  if (json.error.code === 429) {
    let retryDelay = 60;
    // ... your existing retry logic ...
    console.log(`Rate limited. Retrying in ${retryDelay} seconds...`);
  } else {
    // This catches other errors like 400 (Bad Request) or 500 (Server Error)
    throw new Error(`Gemini API Failed: ${json.error.message}`);
  }
}

  if (!json.candidates || !json.candidates.length) {
    throw new Error("No response candidates returned from Gemini.");
  }

  const candidate = json.candidates[0];
  let text = candidate.content.parts.map(p => p.text || "").join(" ").trim();
  text = text.replace(/```json/i, "").replace(/```/g, "").trim();

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error("Gemini did not return valid JSON after cleanup: " + text);
  }

  if (!Array.isArray(parsed)) {
    throw new Error("Gemini did not return a JSON array: " + text);
  }

  return parsed;
}

// 8. CLEANUP
function allRowsProcessed(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== "Sent" && data[i][4] !== "Pending Retry" && !String(data[i][4]).startsWith("Error")) {
      return false;
    }
  }
  return true;
}

function moveAllExistingSentRows() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sourceSheet = ss.getSheetByName(SOURCE_TAB);
  const targetSheet = ss.getSheetByName(TARGET_TAB);
  
  // Safety check: ensure both tabs exist
  if (!sourceSheet || !targetSheet) {
    ss.toast("Error: Check your tab names!", "Failure");
    return;
  }

  const data = sourceSheet.getDataRange().getValues();

  // Loop backwards to avoid index shifting when rows are deleted
  for (let i = data.length - 1; i >= 1; i--) {
    // Column E is index 4 (Status)
    if (data[i][4] === "Sent") { 
      targetSheet.appendRow(data[i]);
      sourceSheet.deleteRow(i + 1);
    }
  }
  
  ss.toast("Cleanup complete! All 'Sent' rows moved to " + TARGET_TAB, 
  "Success");
}

function checkRepliesAndNotify() {
  // Store multiple email IDs in an array
  const notificationEmails = [
    'akshay.aher@bizkonnectsolutions.com',
    'shubhangi.mahajan@bizkonnectsolution.com',
    'vishwas.bhadange@bizkonnect.com'
  ]; 
  
  // Join the array elements with a comma to form a single recipient string
  const emailRecipientString = notificationEmails.join(',');
  
  // 1. Calculate Date (Last 2 days)
  const now = new Date();
  const twoDaysAgo = new Date(now.getTime() - (2 * 24 * 60 * 60 * 1000));
  const dateString = Utilities.formatDate(twoDaysAgo, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  // 2. Search for unread replies
  const searchQuery = `is:unread -from:me after:${dateString}`;
  const threads = GmailApp.search(searchQuery);
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const sender = lastMessage.getFrom().toLowerCase();
    const subject = lastMessage.getSubject().toLowerCase();
    
    // 3. BOUNCE FILTER LOGIC
    // Skip if sender is a system bot or subject contains failure keywords
    const isBounce = sender.includes('mailer-daemon') ||  
                     sender.includes('postmaster') ||  
                     subject.includes('delivery has failed') ||  
                     subject.includes('undeliverable');
    
    if (isBounce) {
      // Mark as read so it stops bothering you, but don't send notification
      thread.markRead();
      console.log('Skipped bounce/system message from: ' + sender);
      continue; 
    }
    
    // 4. SEND NOTIFICATION TO ALL RECIPIENTS
    // Construct the notification details
    const threadLink = thread.getPermalink();
    const emailSubject = `New Reply Received: ${lastMessage.getSubject()}`;
    const emailBody = `You have received a new reply from ${lastMessage.getFrom()}.\n\n` +
                      `Preview: ${lastMessage.getPlainBody().substring(0, 300)}...\n\n` +
                      `View the conversation here: ${threadLink}`;
    
    // Send email to the combined comma-separated string of users
    GmailApp.sendEmail(emailRecipientString, emailSubject, emailBody);
    
    // Optional: Mark thread as read after notifying so you aren't alerted twice next run
    thread.markRead(); 
  }
}

// 1. CONFIGURATION
const RESP_SHEET_ID = '1SMFmM6VZ9Fk_eiNGwDrO5oXMpPy_RQEDHqDdJON4YBA'; 
const TARGET_TAB_NAME = 'Sheet1'; 

function trackGmailResponses() {
  const ss = SpreadsheetApp.openById(RESP_SHEET_ID);
  const respSheet = ss.getSheetByName(TARGET_TAB_NAME);
  
  // Automatically detects your email to skip it
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  if (!respSheet) return;

  const labelMapping = {
    'DS Negative': 'Process Negative',
    'DS Positive': 'Process Positive'
  };

  const labelsToTrack = Object.keys(labelMapping);
  
  labelsToTrack.forEach(labelName => {
    const sourceLabel = GmailApp.getUserLabelByName(labelName);
    if (!sourceLabel) return;

    const targetLabelName = labelMapping[labelName];
    let targetLabel = GmailApp.getUserLabelByName(targetLabelName) || GmailApp.createLabel(targetLabelName);

    const threads = sourceLabel.getThreads();
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      let prospectEmail = "";
      let responseDate = "";

      // LOGIC: Start from the very first message (index 0) and move forward
      for (let i = 0; i < messages.length; i++) {
        let sender = messages[i].getFrom();
        // Extract clean email from "Name <email@domain.com>"
        let cleanEmail = (sender.match(/<([^>]+)>/) ? sender.match(/<([^>]+)>/)[1] : sender).toLowerCase().trim();

        // If the sender is NOT you, this is the prospect
        if (cleanEmail !== myEmail) {
          prospectEmail = cleanEmail;
          responseDate = messages[i].getDate();
          break; // Exit loop immediately after finding the first non-me sender
        }
      }

      if (prospectEmail) {
        // Append to DS_App Script_Negative/positive Sheet
        respSheet.appendRow([prospectEmail, targetLabelName, responseDate]);

        // Move Labels: This prevents the script from processing this thread again
        thread.addLabel(targetLabel);
        thread.removeLabel(sourceLabel);
        
        console.log("Success: Logged prospect " + prospectEmail + " to " + targetLabelName);
      }
    });
  });
}
