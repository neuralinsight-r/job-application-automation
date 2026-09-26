// ============================================================
//  DriveManager.gs — Google Drive folder + Doc creation
//
//  Creates [Company] - [Role] subfolders under the parent,
//  then writes Resume and Cover Letter as Google Docs.
// ============================================================


// ─────────────────────────────────────────────
//  FOLDER MANAGEMENT
// ─────────────────────────────────────────────

/**
 * Creates (or retrieves existing) subfolder under the parent folder.
 * Folder name format: "Company Name - Job Title"
 */
function createJobFolder(folderName) {
  const config       = getConfig();
  const parentFolder = DriveApp.getFolderById(config.PARENT_FOLDER_ID);

  // Check if folder already exists (avoids duplicates on re-runs)
  const existing = parentFolder.getFoldersByName(folderName);
  if (existing.hasNext()) {
    const folder = existing.next();
    Logger.log(`  Reusing existing folder: ${folderName}`);
    return folder;
  }

  const folder = parentFolder.createFolder(folderName);
  Logger.log(`  Created folder: ${folderName}`);
  return folder;
}


// ─────────────────────────────────────────────
//  GOOGLE DOC CREATION
// ─────────────────────────────────────────────

/**
 * Creates a Google Doc inside a given Drive folder.
 * Applies clean professional formatting.
 * Enforces no-em-dash rule throughout the content.
 * Returns the created Doc.
 */
function createGoogleDoc(folder, docTitle, content) {
  // Enforce no-em-dash rule: replace any em dashes Claude may have slipped in
  content = removeEmDashes(content);

  // Create the doc in the root (we'll move it to the folder)
  const doc  = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Apply document-level formatting
  const style = {};
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  style[DocumentApp.Attribute.FONT_SIZE]   = 11;
  style[DocumentApp.Attribute.LINE_SPACING] = 1.15;
  body.setAttributes(style);

  // Clear default content
  body.clear();

  // Parse and write the content with formatting
  writeFormattedContent(body, content);

  doc.saveAndClose();

  // Move from root to target folder
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log(`  Created doc: "${docTitle}" in folder "${folder.getName()}"`);
  return doc;
}

/**
 * Parses Claude's structured text output and writes it to a Doc body
 * with appropriate heading levels, bullet points, and paragraph formatting.
 *
 * Claude is prompted to use these markers:
 *   # Heading 1
 *   ## Heading 2
 *   ### Heading 3
 *   - Bullet item
 *   **Bold text**  (within paragraphs)
 *   [blank line] = paragraph break
 */
function writeFormattedContent(body, content) {
  const lines = content.split('\n');
  let inBulletList = false;

  for (let i = 0; i < lines.length; i++) {
    const raw  = lines[i];
    const line = raw.trimEnd();

    // Heading 1
    if (line.startsWith('# ') && !line.startsWith('## ')) {
      inBulletList = false;
      const text = line.replace(/^# /, '').trim();
      const p = body.appendParagraph(text);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      p.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]:   16,
        [DocumentApp.Attribute.BOLD]:        true,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#1a1a2e',
      });
      continue;
    }

    // Heading 2
    if (line.startsWith('## ') && !line.startsWith('### ')) {
      inBulletList = false;
      const text = line.replace(/^## /, '').trim();
      const p = body.appendParagraph(text);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      p.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]:   13,
        [DocumentApp.Attribute.BOLD]:        true,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#2c3e50',
      });
      continue;
    }

    // Heading 3
    if (line.startsWith('### ')) {
      inBulletList = false;
      const text = line.replace(/^### /, '').trim();
      const p = body.appendParagraph(text);
      p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      p.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]:   11,
        [DocumentApp.Attribute.BOLD]:        true,
        [DocumentApp.Attribute.ITALIC]:      true,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      });
      continue;
    }

    // Bullet item
    if (line.startsWith('- ') || line.startsWith('* ')) {
      inBulletList = true;
      const text = line.replace(/^[-*] /, '').trim();
      const item = body.appendListItem(text);
      item.setGlyphType(DocumentApp.GlyphType.BULLET);
      item.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]:   11,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      });
      applyInlineFormatting(item, text);
      continue;
    }

    // Empty line
    if (line.trim() === '') {
      inBulletList = false;
      // Only add spacing paragraph if not at start/end
      if (i > 0 && i < lines.length - 1) {
        const prev = lines[i - 1].trim();
        const next = lines[i + 1].trim();
        if (prev && next) {
          body.appendParagraph('').setAttributes({
            [DocumentApp.Attribute.FONT_SIZE]: 4,
          });
        }
      }
      continue;
    }

    // Horizontal rule (--- or ===)
    // Apps Script doesn't support true border lines — use a thin underscore line instead
    if (/^[-=]{3,}$/.test(line.trim())) {
      inBulletList = false;
      const p = body.appendParagraph('________________________________');
      p.setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]:       8,
        [DocumentApp.Attribute.FOREGROUND_COLOR]: '#cccccc',
      });
      continue;
    }

    // Normal paragraph
    inBulletList = false;
    const p = body.appendParagraph('');
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]:   11,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    });
    applyInlineFormatting(p, line.trim());
  }
}

/**
 * Applies inline bold formatting for **text** patterns.
 * Sets the full paragraph text, then bolds the marked segments.
 */
function applyInlineFormatting(paragraph, text) {
  // Split on **bold** markers
  const parts  = text.split(/\*\*([^*]+)\*\*/);
  // parts = [ normalText, boldText, normalText, boldText, ... ]

  // Clear the paragraph and rebuild with formatting
  paragraph.clear();
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part) continue;
    const textElement = paragraph.appendText(part);
    if (i % 2 === 1) {
      // Odd indices = bold segments
      textElement.setBold(true);
    } else {
      textElement.setBold(false);
    }
  }
}

/**
 * Removes all em dashes from content, replacing them with a space or comma
 * depending on context, then cleans up double spaces.
 */
function removeEmDashes(text) {
  if (!text) return text;
  // Replace em dash (—) with comma-space or just space
  return text
    .replace(/\s*—\s*/g, ', ')   // "word — word" → "word, word"
    .replace(/,,/g, ',')          // clean double commas
    .replace(/  +/g, ' ');        // clean double spaces
}
