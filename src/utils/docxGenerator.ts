import PizZip from 'pizzip';

export const ALLOWED_PLACEHOLDERS = [
  'MEMO_NO',
  'DATE',
  'GDE_NO',
  'GDE_DATE',
  'NCMEC_TIPLINE_NO',
  'PS_AUTHORITY',
  'ADDRESSEE_PS',
  'ADDRESSEE_NAME',
  'FULL_ADDRESS',
  'IO_NAME',
  'IO_MOBILE',
] as const;

export type PlaceholderKey = (typeof ALLOWED_PLACEHOLDERS)[number];

export type PlaceholderValues = Record<PlaceholderKey, string>;

const PLACEHOLDER_REGEX = new RegExp(
  `\\{\\{(${ALLOWED_PLACEHOLDERS.join('|')})\\}\\}`,
  'g'
);

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

function setXmlSpacePreserve(tNode: Element) {
  tNode.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  tNode.setAttribute('xml:space', 'preserve');
}

/**
 * Sets text content on a <w:t> node. If the replacement text contains newlines
 * (e.g., pasted multiline FULL_ADDRESS), inserts <w:br/> elements between lines
 * inside the same parent <w:r> run so each line stays in one clean <w:t> node.
 */
function setNodeTextWithLineBreaks(xmlDoc: Document, tNode: Element, text: string) {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  if (!normalized.includes('\n')) {
    tNode.textContent = normalized;
    setXmlSpacePreserve(tNode);
    return;
  }

  const lines = normalized.split('\n');
  tNode.textContent = lines[0];
  setXmlSpacePreserve(tNode);

  const parentRun = tNode.parentNode;
  if (!parentRun) return;

  const ns = tNode.namespaceURI || W_NS;
  let referenceNode: Node | null = tNode.nextSibling;

  for (let i = 1; i < lines.length; i++) {
    const brEl = xmlDoc.createElementNS(ns, 'w:br');
    parentRun.insertBefore(brEl, referenceNode);

    const newT = xmlDoc.createElementNS(ns, 'w:t');
    const lineText = lines[i];
    newT.textContent = lineText;
    setXmlSpacePreserve(newT);
    parentRun.insertBefore(newT, referenceNode);
  }
}

/**
 * Ensures the paragraph containing {{FULL_ADDRESS}} is strictly left-aligned
 * and does not inherit or apply any tab stops or <w:tab/> elements that could
 * push parts of an address line to the right side of the page.
 */
function enforceCleanLeftAlignedAddressParagraph(xmlDoc: Document, p: Element) {
  const ns = p.namespaceURI || W_NS;

  // Remove any <w:tab/> nodes inside the address paragraph
  const tabNodes = Array.from(p.getElementsByTagName('w:tab'));
  for (const tabNode of tabNodes) {
    tabNode.parentNode?.removeChild(tabNode);
  }

  // Ensure <w:pPr> exists on the paragraph
  let pPr = Array.from(p.childNodes).find(
    (node) => node.nodeType === 1 && (node as Element).tagName === 'w:pPr'
  ) as Element | undefined;

  if (!pPr) {
    pPr = xmlDoc.createElementNS(ns, 'w:pPr');
    p.insertBefore(pPr, p.firstChild);
  }

  // Remove any custom tab stops (<w:tabs>) from the paragraph properties
  const tabsEls = Array.from(pPr.getElementsByTagName('w:tabs'));
  for (const tabsEl of tabsEls) {
    tabsEl.parentNode?.removeChild(tabsEl);
  }

  // Force left alignment (<w:jc w:val="left"/>) so soft line breaks (<w:br/>)
  // never trigger justified word/character stretching across the page width
  let jcEl = Array.from(pPr.childNodes).find(
    (node) => node.nodeType === 1 && (node as Element).tagName === 'w:jc'
  ) as Element | undefined;

  if (!jcEl) {
    jcEl = xmlDoc.createElementNS(ns, 'w:jc');
    pPr.appendChild(jcEl);
  }
  jcEl.setAttributeNS(ns, 'w:val', 'left');
  jcEl.setAttribute('w:val', 'left');
}

function sanitizeFullAddress(rawAddress: string): string {
  return rawAddress
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\t+/g, ' ');
}

/**
 * Replaces ONLY exact {{PLACEHOLDER}} tokens inside a DOCX file without altering
 * paragraph structure, tables, runs, or formatting.
 * Handles cases where Word splits a {{PLACEHOLDER}} across multiple <w:t> runs
 * inside the same paragraph while preserving the original XML node tree.
 */
function replacePlaceholdersInXmlDom(xmlString: string, values: PlaceholderValues): string {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, 'application/xml');

  const parseError = xmlDoc.getElementsByTagName('parsererror');
  if (parseError.length > 0) {
    return replacePlaceholdersInRawXml(xmlString, values);
  }

  const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));

  for (const p of paragraphs) {
    const tNodes = Array.from(p.getElementsByTagName('w:t'));
    if (tNodes.length === 0) continue;

    // Ensure xml:space="preserve" is set on all <w:t> nodes so XMLSerializer
    // and Word never strip leading/trailing spaces between adjacent runs.
    for (const node of tNodes) {
      setXmlSpacePreserve(node);
    }

    let fullText = '';
    const nodeRanges: { node: Element; start: number; end: number; text: string }[] = [];

    for (const node of tNodes) {
      const text = node.textContent || '';
      const start = fullText.length;
      fullText += text;
      nodeRanges.push({
        node,
        start,
        end: fullText.length,
        text,
      });
    }

    // Fix word-spacing in the main body paragraph ("In exercise of the power conferred under section 179 BNSS...")
    // where adjacent Word runs in the master may lack separating spaces (e.g., "referredNotice", "purposeofinvestigationthiscase").
    const isMainBodyParagraph =
      /In\s*exercise\s*of\s*the\s*power\s*conferred/i.test(fullText) ||
      /section\s*179\s*BNSS/i.test(fullText) ||
      /purpose\s*of\s*investigation/i.test(fullText);

    if (isMainBodyParagraph) {
      const normalizedBody = fullText
        .replace(/\b(In)(?=exercise\b)/gi, '$1 ')
        .replace(/\b(exercise)(?=of\b)/gi, '$1 ')
        .replace(/\b(of)(?=the\b)/gi, '$1 ')
        .replace(/\b(the)(?=power\b)/gi, '$1 ')
        .replace(/\b(power)(?=conferred\b)/gi, '$1 ')
        .replace(/\b(conferred)(?=under\b)/gi, '$1 ')
        .replace(/\b(under)(?=section\b)/gi, '$1 ')
        .replace(/\b(section)(?=179\b)/gi, '$1 ')
        .replace(/\b(179)(?=BNSS\b)/gi, '$1 ')
        .replace(/(BNSS,)(?=\S)/gi, '$1 ')
        .replace(/\b(You)(?=are\b)/gi, '$1 ')
        .replace(/\b(are)(?=hereby\b)/gi, '$1 ')
        .replace(/\b(hereby)(?=directed\b)/gi, '$1 ')
        .replace(/\b(directed)(?=through\b)/gi, '$1 ')
        .replace(/\b(through)(?=this\b)/gi, '$1 ')
        .replace(/\b(this)(?=notice\b)/gi, '$1 ')
        .replace(/\b(notice)(?=to\b)/gi, '$1 ')
        .replace(/\b(to)(?=appear\b)/gi, '$1 ')
        .replace(/\b(appear)(?=before\b)/gi, '$1 ')
        .replace(/\b(before)(?=the\b)/gi, '$1 ')
        .replace(/\b(the)(?=undersigned\b)/gi, '$1 ')
        .replace(/\b(undersigned)(?=at\b)/gi, '$1 ')
        .replace(/\b(at)(?=Birbhum\b)/gi, '$1 ')
        .replace(/\b(Birbhum)(?=Cyber\b)/gi, '$1 ')
        .replace(/\b(Cyber)(?=Crime\b)/gi, '$1 ')
        .replace(/\b(Crime)(?=PS\b)/gi, '$1 ')
        .replace(/(PS,)(?=\S)/gi, '$1 ')
        .replace(/(Suri,)(?=\S)/gi, '$1 ')
        .replace(/\b(Birbhum)(?=\(West)/gi, '$1 ')
        .replace(/\b(West)(?=Bengal\))/gi, '$1 ')
        .replace(/(\(West\s+Bengal\))(?=\S)/gi, '$1 ')
        .replace(/\b(Opposite)(?=to\b)/gi, '$1 ')
        .replace(/\b(to)(?=M\b)/gi, '$1 ')
        .replace(/\b(M)(?=Bazar)/gi, '$1 ')
        .replace(/\b(Super)(?=Market\b)/gi, '$1 ')
        .replace(/\b(Market)(?=Suri)/gi, '$1 ')
        .replace(/\b(Near)(?=Suri\b)/gi, '$1 ')
        .replace(/\b(Suri)(?=Bus\b)/gi, '$1 ')
        .replace(/\b(Bus)(?=Stand\b)/gi, '$1 ')
        .replace(/\b(Stand)(?=within\b)/gi, '$1 ')
        .replace(/\b(within)(?=Three\b)/gi, '$1 ')
        .replace(/\b(Three)(?=\(03\))/gi, '$1 ')
        .replace(/(\(03\))(?=\S)/gi, '$1 ')
        .replace(/\b(days)(?=of\b)/gi, '$1 ')
        .replace(/\b(of)(?=receiving\b)/gi, '$1 ')
        .replace(/\b(receiving)(?=the\b)/gi, '$1 ')
        .replace(/\b(the)(?=referred\b)/gi, '$1 ')
        .replace(/\b(referred)(?=Notice\b)/gi, '$1 ')
        .replace(/\b(Notice)(?=for\b)/gi, '$1 ')
        .replace(/\b(for)(?=the\b)/gi, '$1 ')
        .replace(/\b(the)(?=purpose\b)/gi, '$1 ')
        .replace(/\b(purpose)(?=of\b)/gi, '$1 ')
        .replace(/\b(of)(?=investigation\b)/gi, '$1 ')
        .replace(/\b(investigation)(?=this\b)/gi, '$1 ')
        .replace(/\b(this)(?=case\b)/gi, '$1 ')
        .replace(/[ \t]{2,}/g, ' ');

      if (normalizedBody !== fullText) {
        // First insert missing spaces at run boundaries where words joined across runs
        for (let idx = 0; idx < nodeRanges.length - 1; idx++) {
          const curr = nodeRanges[idx];
          const next = nodeRanges[idx + 1];
          if (!curr.text || !next.text) continue;
          const lastChar = curr.text[curr.text.length - 1];
          const firstChar = next.text[0];
          if (/[A-Za-z0-9),.]/.test(lastChar) && /[A-Za-z0-9(]/.test(firstChar)) {
            if (!(lastChar === '/' || firstChar === '/')) {
              curr.text = curr.text + ' ';
            }
          }
        }

        // Also apply exact word separation inside each individual <w:t> node
        // so run-level bold/formatting stays on its original run while fixing joined words
        for (const range of nodeRanges) {
          range.text = range.text
            .replace(/\b(referred)(?=Notice\b)/gi, '$1 ')
            .replace(/\b(Notice)(?=for\b)/gi, '$1 ')
            .replace(/\b(for)(?=the\b)/gi, '$1 ')
            .replace(/\b(the)(?=purpose\b)/gi, '$1 ')
            .replace(/\b(purpose)(?=of\b)/gi, '$1 ')
            .replace(/\b(of)(?=investigation\b)/gi, '$1 ')
            .replace(/\b(investigation)(?=this\b)/gi, '$1 ')
            .replace(/\b(this)(?=case\b)/gi, '$1 ')
            .replace(/\b(within)(?=Three\b)/gi, '$1 ')
            .replace(/\b(Three)(?=\(03\))/gi, '$1 ')
            .replace(/(\(03\))(?=\S)/gi, '$1 ')
            .replace(/\b(days)(?=of\b)/gi, '$1 ')
            .replace(/\b(of)(?=receiving\b)/gi, '$1 ')
            .replace(/\b(receiving)(?=the\b)/gi, '$1 ')
            .replace(/\b(the)(?=referred\b)/gi, '$1 ')
            .replace(/[ \t]{2,}/g, ' ');
          range.node.textContent = range.text;
          setXmlSpacePreserve(range.node);
        }

        // Recompute fullText and nodeRanges offsets
        fullText = '';
        for (const range of nodeRanges) {
          range.start = fullText.length;
          fullText += range.text;
          range.end = fullText.length;
        }

        // If any joined words still remain across split runs, place the normalized body
        // into the primary non-empty <w:t> run while keeping paragraph and run formatting intact
        if (fullText !== normalizedBody) {
          const firstIdx = nodeRanges.findIndex((r) => r.text.length > 0);
          const targetIdx = firstIdx >= 0 ? firstIdx : 0;
          for (let idx = 0; idx < nodeRanges.length; idx++) {
            nodeRanges[idx].text = idx === targetIdx ? normalizedBody : '';
            nodeRanges[idx].node.textContent = nodeRanges[idx].text;
            setXmlSpacePreserve(nodeRanges[idx].node);
          }
          fullText = normalizedBody;
        }
      }
    }

    if (!fullText.includes('{{')) continue;

    const isReferenceParagraph =
      fullText.includes('{{GDE_NO}}') ||
      fullText.includes('{{GDE_DATE}}') ||
      fullText.includes('{{NCMEC_TIPLINE_NO}}') ||
      /\bRef\s*:/i.test(fullText);

    const matches: {
      key: PlaceholderKey;
      start: number;
      end: number;
      replacement: string;
    }[] = [];
    let match: RegExpExecArray | null;
    PLACEHOLDER_REGEX.lastIndex = 0;

    while ((match = PLACEHOLDER_REGEX.exec(fullText)) !== null) {
      const key = match[1] as PlaceholderKey;
      const rawValue = values[key] ?? '';
      let replacement =
        key === 'FULL_ADDRESS' ? sanitizeFullAddress(rawValue) : rawValue;

      // For Reference line placeholders, ensure a single space separates the inserted
      // value from adjacent words if the master runs had no space between them.
      if (
        isReferenceParagraph &&
        (key === 'GDE_NO' || key === 'GDE_DATE' || key === 'NCMEC_TIPLINE_NO')
      ) {
        const charBefore = match.index > 0 ? fullText[match.index - 1] : '';
        const charAfter =
          match.index + match[0].length < fullText.length
            ? fullText[match.index + match[0].length]
            : '';

        if (charBefore && !/\s/.test(charBefore) && !replacement.startsWith(' ')) {
          replacement = ' ' + replacement;
        }
        if (
          charAfter &&
          !/\s/.test(charAfter) &&
          !/^[,.;:!?)]/.test(charAfter) &&
          !replacement.endsWith(' ')
        ) {
          replacement = replacement + ' ';
        }
      }

      matches.push({
        key,
        start: match.index,
        end: match.index + match[0].length,
        replacement,
      });
    }

    if (matches.length === 0) continue;

    const hasFullAddress = matches.some((m) => m.key === 'FULL_ADDRESS');
    if (hasFullAddress) {
      enforceCleanLeftAlignedAddressParagraph(xmlDoc, p);
    }

    // Apply matches in reverse order so character offsets remain valid
    for (let i = matches.length - 1; i >= 0; i--) {
      const { start, end, replacement } = matches[i];

      for (let nIdx = 0; nIdx < nodeRanges.length; nIdx++) {
        const range = nodeRanges[nIdx];
        if (range.end <= start || range.start >= end) {
          continue;
        }

        const localStart = Math.max(0, start - range.start);
        const localEnd = Math.min(range.text.length, end - range.start);

        const isFirstOverlappingNode = range.start <= start && range.end > start;

        if (isFirstOverlappingNode) {
          range.text =
            range.text.slice(0, localStart) +
            replacement +
            range.text.slice(localEnd);
        } else {
          range.text =
            range.text.slice(0, localStart) +
            range.text.slice(localEnd);
        }
      }

      // Recalculate ranges in case multiple placeholders share the same paragraph
      let updatedFullText = '';
      for (const range of nodeRanges) {
        range.start = updatedFullText.length;
        updatedFullText += range.text;
        range.end = updatedFullText.length;
      }
    }

    // In the Reference paragraph, ensure words that were split across adjacent runs without
    // spaces in the master template (e.g. "No." + "140" + "dtd." + "05.10.2026" + "in" + "c/w")
    // have exactly one normal space between words while keeping all original run formatting intact.
    if (isReferenceParagraph) {
      const combinedRef = nodeRanges.map((r) => r.text).join('');
      const normalizedRef = combinedRef
        .replace(/\b(Ref:)(?=\S)/gi, '$1 ')
        .replace(/\b(Birbhum)(?=Cyber\b)/gi, '$1 ')
        .replace(/\b(Cyber)(?=Crime\b)/gi, '$1 ')
        .replace(/\b(Crime)(?=PS\b)/gi, '$1 ')
        .replace(/\b(PS)(?=GDE\b)/gi, '$1 ')
        .replace(/\b(GDE)(?=No\.)/gi, '$1 ')
        .replace(/(GDE\s+No\.)(?=\S)/gi, '$1 ')
        .replace(/(\S)(?=dtd\.)/gi, '$1 ')
        .replace(/(dtd\.)(?=\S)/gi, '$1 ')
        .replace(/(\S)(?=in\s*c\/w\b)/gi, '$1 ')
        .replace(/\b(in)(?=c\/w\b)/gi, '$1 ')
        .replace(/(c\/w)(?=\S)/gi, '$1 ')
        .replace(/\b(Cyber)(?=NCMEC\b)/gi, '$1 ')
        .replace(/\b(NCMEC)(?=Tipline\b)/gi, '$1 ')
        .replace(/\b(Tipline)(?=Report\b)/gi, '$1 ')
        .replace(/\b(Report)(?=No\.)/gi, '$1 ')
        .replace(/(Report\s+No\.)(?=\S)/gi, '$1 ')
        .replace(/[ \t]{2,}/g, ' ');

      if (normalizedRef !== combinedRef) {
        // Place the clean normalized Reference line into the first non-empty <w:t> run
        // (which already carries the paragraph's bold + underline run formatting)
        // and clear the remaining <w:t> nodes in this paragraph.
        const firstIdx = nodeRanges.findIndex((r) => r.text.length > 0);
        const targetIdx = firstIdx >= 0 ? firstIdx : 0;
        for (let idx = 0; idx < nodeRanges.length; idx++) {
          nodeRanges[idx].text = idx === targetIdx ? normalizedRef : '';
        }
      }
    }

    // Commit updated text (and line breaks if multiline) to the XML DOM nodes
    // and ensure xml:space="preserve" is set on every <w:t> in this paragraph
    // so Word never drops leading or trailing spaces between runs.
    for (const range of nodeRanges) {
      if ((range.node.textContent || '') !== range.text) {
        setNodeTextWithLineBreaks(xmlDoc, range.node, range.text);
      } else {
        setXmlSpacePreserve(range.node);
      }
    }
  }

  const serializer = new XMLSerializer();
  return serializer.serializeToString(xmlDoc);
}

function escapeXml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replacePlaceholdersInRawXml(xmlString: string, values: PlaceholderValues): string {
  return xmlString.replace(PLACEHOLDER_REGEX, (_, key: PlaceholderKey) => {
    const rawValue = values[key] ?? '';
    const cleaned =
      key === 'FULL_ADDRESS'
        ? sanitizeFullAddress(rawValue)
        : rawValue.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const lines = cleaned.split('\n').map(escapeXml);
    return lines.join('</w:t><w:br/><w:t xml:space="preserve">');
  });
}

function findUnresolvedPlaceholders(xmlString: string): string[] {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
  const parseError = xmlDoc.getElementsByTagName('parsererror');

  let combinedText = '';
  if (parseError.length > 0) {
    combinedText = xmlString;
  } else {
    const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
    for (const p of paragraphs) {
      const tNodes = Array.from(p.getElementsByTagName('w:t'));
      combinedText += tNodes.map((n) => n.textContent || '').join('') + '\n';
    }
  }

  const unresolved = new Set<string>();
  const regex = /\{\{[^{}]+\}\}/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(combinedText)) !== null) {
    unresolved.add(match[0]);
  }
  return Array.from(unresolved);
}

export function generateNoticeDocx(
  masterBuffer: ArrayBuffer,
  values: PlaceholderValues
): Blob {
  // Always use a fresh copy of the master buffer for every generated notice
  const freshBuffer = masterBuffer.slice(0);
  const zip = new PizZip(freshBuffer);

  const xmlFiles = Object.keys(zip.files).filter(
    (fileName) =>
      fileName === 'word/document.xml' ||
      /^word\/header\d*\.xml$/.test(fileName) ||
      /^word\/footer\d*\.xml$/.test(fileName)
  );

  const unresolvedSet = new Set<string>();

  for (const fileName of xmlFiles) {
    const file = zip.file(fileName);
    if (!file) continue;
    const xmlContent = file.asText();
    const updatedXml = replacePlaceholdersInXmlDom(xmlContent, values);

    for (const token of findUnresolvedPlaceholders(updatedXml)) {
      unresolvedSet.add(token);
    }

    zip.file(fileName, updatedXml);
  }

  if (
    unresolvedSet.has('{{PS_AUTHORITY}}') ||
    unresolvedSet.has('{{ADDRESSEE_PS}}')
  ) {
    throw new Error(
      'Failed to replace {{PS_AUTHORITY}} or {{ADDRESSEE_PS}} in the DOCX master.'
    );
  }

  if (unresolvedSet.size > 0) {
    throw new Error(
      `Unresolved placeholders found in document: ${Array.from(unresolvedSet).join(', ')}`
    );
  }

  return zip.generate({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}
