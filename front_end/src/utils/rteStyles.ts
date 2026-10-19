/**
 * RTE content styles, inlined so they travel with the email HTML
 * delivered to recipients and to our own Sent Mail view.
 *
 * The same styles are also defined in `src/index.css` for the
 * in-app editor and reader (where the CSS can live in a stylesheet).
 * For the recipient's mail client we cannot rely on our stylesheet,
 * so we embed the rules as a `<style>` block inside the HTML body
 * whenever we send or save a message.
 *
 * Source:
 * https://ej2.syncfusion.com/documentation/rich-text-editor/editor-value#styling-editor-content
 *
 * Tailwind preflight reset (also embedded so mail clients don't
 * strip list bullets / padding / margins):
 * https://ej2.syncfusion.com/angular/documentation/rich-text-editor/integrations/tailwind-preflight
 */

export const RTE_CONTENT_STYLES = `
.e-rte-content li { margin-bottom: 10px !important; padding: unset !important; }
.e-rte-content ul { list-style-type: disc !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol { list-style-type: decimal !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ul[style*="list-style-type: circle"] { list-style-type: circle !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ul[style*="list-style-type: square"] { list-style-type: square !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol[style*="list-style-type: upper-alpha"] { list-style-type: upper-alpha !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol[style*="list-style-type: lower-alpha"] { list-style-type: lower-alpha !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol[style*="list-style-type: upper-roman"] { list-style-type: upper-roman !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol[style*="list-style-type: lower-roman"] { list-style-type: lower-roman !important; padding: unset !important; margin-left: 40px; }
.e-rte-content ol[style*="list-style-type: lower-greek"] { list-style-type: lower-greek !important; padding: unset !important; margin-left: 40px; }
.e-rte-content p { margin: 0 0 10px; margin-bottom: 10px; }
.e-rte-content h1 { font-size: 2.857em; font-weight: 600; line-height: 1.2; margin: 10px 0; }
.e-rte-content h2 { font-size: 2.285em; font-weight: 600; line-height: 1.2;margin: 10px 0; }
.e-rte-content h3 { font-size: 2em; font-weight: 600; line-height: 1.2; margin: 10px 0; }
.e-rte-content h4 { font-size: 1.714em; font-weight: 600; line-height: 1.2; margin: 0; }
.e-rte-content h5 { font-size: 1.428em; font-weight: 600; line-height: 1.2; margin: 0; }
.e-rte-content h6 { font-size: 1.142em; line-height: 1.5; margin: 0; }
.e-rte-content blockquote { margin: 10px 0; margin-left: 0; padding-left: 5px; border-left: solid 3px #333; }
.e-rte-content pre { background-color: inherit; border: 0; border-radius: 0; color: #333; font-size: inherit; line-height: inherit; margin: 0 0 10px; overflow: visible; padding: 0; white-space: pre-wrap; word-break: inherit; word-wrap: break-word; }
.e-rte-content strong, .e-rte-content b { font-weight: 700; }
.e-rte-content a { color: #2e2ef1; text-decoration: none; -webkit-user-select: auto; -ms-user-select: auto; user-select: auto; }
.e-rte-content a:hover { text-decoration: underline; }
.e-rte-content h3 + h4, .e-rte-content h4 + h5, .e-rte-content h5 + h6 { margin-top: 0.6em; }
.e-rte-content .e-rte-table { border-collapse: collapse; empty-cells: show; }
.e-rte-content .e-rte-table td, .e-rte-content .e-rte-table th { border: 1px solid #bdbdbd; height: 20px; min-width: 20px; padding: 2px 5px; vertical-align: middle; }
.e-rte-content .e-rte-table.e-dashed-border td, .e-rte-content .e-rte-table.e-dashed-border th { border-style: dashed; }
.e-rte-content .e-rte-table th { background-color: #e0e0e0; }
.e-rte-image.e-img-inline, .e-img-caption-container.e-img-inline { margin-left: 5px; margin-right: 5px; display: inline-block; float: none; max-width: 100%; padding: 1px; vertical-align: bottom; }
.e-rte-image.e-img-center, .e-rte-image.e-img-break, .e-img-caption-container.e-img-center, .e-img-caption-container.e-img-break { cursor: pointer; display: block; float: none; margin: 5px auto; max-width: 100%; position: relative; }
.e-rte-content .e-rte-image.e-img-left-wrap, .e-rte-content .e-img-caption-container.e-img-left-wrap { float: left; margin: 5px 5px 5px 0; }
.e-rte-content .e-rte-image.e-img-right-wrap, .e-rte-content .e-img-caption-container.e-img-right-wrap { float: right; margin: 5px 0 5px 5px; }
.e-rte-content .e-rte-image.e-img-left, .e-rte-content .e-rte-image.e-img-right, .e-rte-content .e-img-caption-container.e-img-left, .e-rte-content .e-img-caption-container.e-img-right { margin-top: 5px; margin-bottom: 5px; }
.e-rte-content .e-rte-image.e-img-right, .e-rte-content .e-img-caption-container.e-img-right { margin-right: 0; margin-left: auto; }
.e-rte-content .e-rte-image.e-img-left, .e-rte-content .e-img-caption-container.e-img-left { margin-left: 0; margin-right: auto; }
.e-rte-content .e-img-caption-container { max-width: 100%; position: relative; }
.e-rte-content .e-img-caption-container .e-img-caption-text { box-sizing: border-box; display: block; font-size: 16px; font-weight: 400; margin: auto; opacity: 0.9; position: relative; text-align: center; width: 100%; }
.e-rte-content .e-img-caption-container.e-img-center, .e-rte-content .e-img-caption-container.e-img-break { margin-left: auto; margin-right: auto; }
`.trim();

/**
 * Wrap a raw RTE `value` (the inner HTML the editor exposes) in a
 * self-contained HTML envelope so that the styles travel with the
 * email body.
 *
 * The output is a single `<div class="e-rte-content">` containing a
 * scoped `<style>` block followed by the original HTML. This is the
 * same container class used by Syncfusion's RTE and by our in-app
 * `MailReader`, so:
 *   - In Outlook / Gmail / any HTML mail client, the embedded
 *     `<style>` block applies the same formatting the user saw
 *     inside the composer.
 *   - In our Sent Mail view, the styles are present in the body
 *     itself, so the message renders consistently even if the
 *     page-level CSS is unavailable (e.g. forwarded / printed).
 */
export function wrapRteBody(html: string): string {
  const body = html && html.trim().length > 0 ? html : '<p><br/></p>';
  return `<div class="e-rte-content"><style>${RTE_CONTENT_STYLES}</style>${body}</div>`;
}
