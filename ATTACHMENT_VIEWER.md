# Analytical attachment viewer

Images open in a near-fullscreen dialog, with zoom, fit, download and Escape
to close. The same viewer serves HPLC chromatograms and MS spectra.

PDF pages are rendered locally using PDF.js 6.5.438 (Apache-2.0). Zoom, fit and
previous/next page controls do not change the stored attachment. No analytical
attachment is sent to a remote conversion service. PDF JavaScript and XFA are
not executed by this viewer. Password-protected files are reported as such.

Publish `attachment-pdf.js`, updated `lab-ui.js`, `workspace.css`, `index.html`
and the complete `vendor/pdfjs/` directory, preserving paths and licenses.
PDF.js is loaded lazily only for PDF attachments.

HTTP(S), including the local preview server, is recommended. Opening the HTML
as `file://` can block the module loader; in that case the existing browser PDF
viewer is used as a fallback, with availability depending on the browser.

Primary API reference: https://mozilla.github.io/pdf.js/examples/

Verification: `node tests/attachment-viewer.test.cjs`. The test uses isolated
synthetic attachments, checks desktop/mobile dialog size, image zoom/fit and
download, and checks real PDF rendering using canvas pixel data.
