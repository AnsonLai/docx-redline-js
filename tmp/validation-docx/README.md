# Validation Fixtures

This folder contains generated OOXML parts for release-time validation.

The script writes document XML rather than complete .docx packages because this
package intentionally does not add a zip dependency.

To manually inspect these fixtures:

1. Copy a generated *.document.xml file into a minimal .docx package as word/document.xml.
2. Include any matching *.numbering.xml as word/numbering.xml.
3. Open with Word, LibreOffice, or another OOXML consumer.
