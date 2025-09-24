function convertMarkdownToGoogleDocs() {
  // Get the active document
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Clear existing content
  body.clear();

  // The markdown text to convert
  var markdownText = `## We Are Switching to Digital Time Tracking

**Effective Monday, we will be using a new mobile app for all time clock activities.**

### What You Need to Know:

📱 **New System:** All employees will use a smartphone app to clock in/out
🏢 **Location-Based:** You must be at your work location to clock in/out
📍 **GPS Required:** The app uses your phone's location services

### What You Need to Do:

#### ✅ **BY END OF DAY FRIDAY:**
**Send Cobey your Gmail address**
📧 Email: cobey@truerootsnm.com
📝 Subject: "My Gmail for Time Clock App"

*If you don't have a Gmail account, create one at gmail.com*

#### ✅ **THIS WEEKEND:**
**Scan the QR code below** to access complete setup instructions and app information

#### ✅ **STARTING MONDAY:**
**Begin using the new app** for all time clock activities
- No more paper timecards
- Clock in/out directly from your phone
- App must be used by ALL employees and contractors

---

## 📱 SCAN FOR COMPLETE INSTRUCTIONS

**[QR CODE GOES HERE]**

*The QR code links to: https://cobeycobb.github.io/employee-timeclock-info/*

This website contains:
- Step-by-step installation guide
- How to use the app daily
- Troubleshooting help
- Important policies you must know

---

## Questions?

Contact Cobey immediately if you have any questions or issues:
📧 **cobey@truerootsnm.com**

---

**⚠️ IMPORTANT:** Failure to provide your Gmail address by Friday or failure to use the app starting Monday may result in delays to your payroll. This system is mandatory for all employees and contractors.`;

  // Split text into lines for processing
  var lines = markdownText.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    // Skip empty lines but add spacing
    if (line.trim() === '') {
      body.appendParagraph('').setSpacingAfter(6);
      continue;
    }

    // Handle horizontal rules
    if (line.trim() === '---') {
      body.appendHorizontalRule();
      continue;
    }

    // Handle headers
    if (line.startsWith('## ')) {
      var headerText = line.substring(3);
      var header = body.appendParagraph(headerText);
      header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      header.editAsText().setFontSize(20).setBold(true);
      header.setSpacingBefore(12).setSpacingAfter(8);
      continue;
    }

    if (line.startsWith('### ')) {
      var headerText = line.substring(4);
      var header = body.appendParagraph(headerText);
      header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      header.editAsText().setFontSize(16).setBold(true);
      header.setSpacingBefore(10).setSpacingAfter(6);
      continue;
    }

    if (line.startsWith('#### ')) {
      var headerText = line.substring(5);
      var header = body.appendParagraph(headerText);
      header.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      header.editAsText().setFontSize(14).setBold(true);
      header.setSpacingBefore(8).setSpacingAfter(4);
      continue;
    }

    // Handle list items
    if (line.startsWith('- ')) {
      var listText = line.substring(2);
      var listItem = body.appendListItem(listText);
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      formatInlineText(listItem);
      continue;
    }

    // Handle regular paragraphs
    var paragraph = body.appendParagraph(line);
    formatInlineText(paragraph);

    // Add extra spacing for certain types of content
    if (line.includes('**') || line.includes('📧') || line.includes('⚠️')) {
      paragraph.setSpacingAfter(6);
    }
  }

  // Add title at the very top
  var title = body.insertParagraph(0, 'IMPORTANT NOTICE: NEW TIME CLOCK SYSTEM');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.editAsText().setFontSize(24).setBold(true);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setSpacingAfter(12);

  // Add horizontal rule after title
  body.insertHorizontalRule(1);
}

function formatInlineText(element) {
  var text = element.getText();
  var textStyle = element.editAsText();

  // Format bold text **text**
  var boldRegex = /\*\*(.*?)\*\*/g;
  var match;
  var offset = 0;

  while ((match = boldRegex.exec(text)) !== null) {
    var start = match.index - offset;
    var end = start + match[1].length;

    // Remove the markdown syntax
    textStyle.deleteText(start, start + 1); // Remove first *
    textStyle.deleteText(start, start + 1); // Remove second *
    textStyle.deleteText(end - 4, end - 3); // Remove first *
    textStyle.deleteText(end - 4, end - 3); // Remove second *

    // Apply bold formatting
    textStyle.setBold(start, end - 5, true);

    offset += 4; // Account for removed characters
  }

  // Format italic text *text*
  text = textStyle.getText(); // Get updated text
  var italicRegex = /\*(.*?)\*/g;
  offset = 0;

  while ((match = italicRegex.exec(text)) !== null) {
    var start = match.index - offset;
    var end = start + match[1].length;

    // Remove the markdown syntax
    textStyle.deleteText(start, start);
    textStyle.deleteText(end - 2, end - 2);

    // Apply italic formatting
    textStyle.setItalic(start, end - 3, true);

    offset += 2;
  }
}

// Alternative function to just insert and format specific sections
function insertFormattedNotice() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();

  if (cursor) {
    // Insert at cursor position
    var element = cursor.getElement();
    // Run the main conversion function
    convertMarkdownToGoogleDocs();
  } else {
    // If no cursor, just run the main function
    convertMarkdownToGoogleDocs();
  }
}