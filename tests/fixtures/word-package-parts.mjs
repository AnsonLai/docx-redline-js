const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const escapeText = value => String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const escapeAttribute = value => escapeText(value).replace(/"/g, '&quot;');

export function createCommentsPart(comments) {
    if (!Array.isArray(comments) || comments.length === 0) throw new Error('comments must be a non-empty array');
    const entries = comments.map(comment => {
        if (!Number.isInteger(comment.id) || comment.id < 0) throw new Error('comment IDs must be non-negative integers');
        return `  <w:comment w:id="${comment.id}" w:author="${escapeAttribute(comment.author || 'Reviewer')}"${comment.date ? ` w:date="${escapeAttribute(comment.date)}"` : ''}>
    <w:p><w:r><w:t>${escapeText(comment.text)}</w:t></w:r></w:p>
  </w:comment>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${NS_W}">
${entries.join('\n')}
</w:comments>`;
}

export function createNotesPart(kind, notes) {
    if (!['footnote', 'endnote'].includes(kind)) throw new Error('note kind must be footnote or endnote');
    if (!Array.isArray(notes) || notes.length === 0) throw new Error('notes must be a non-empty array');
    const plural = `${kind}s`;
    const reference = `${kind}Ref`;
    const entries = notes.map(note => {
        if (!Number.isInteger(note.id) || note.id < 1) throw new Error('note IDs must be positive integers');
        return `  <w:${kind} w:id="${note.id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${reference}/></w:r><w:r><w:t xml:space="preserve"> ${escapeText(note.text)}</w:t></w:r></w:p></w:${kind}>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${plural} xmlns:w="${NS_W}">
  <w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>
  <w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>
${entries.join('\n')}
</w:${plural}>`;
}

export function createHeaderFooterPart(kind, text) {
    if (!['header', 'footer'].includes(kind)) throw new Error('part kind must be header or footer');
    const root = kind === 'header' ? 'hdr' : 'ftr';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${root} xmlns:w="${NS_W}"><w:p><w:r><w:t>${escapeText(text)}</w:t></w:r></w:p></w:${root}>`;
}
