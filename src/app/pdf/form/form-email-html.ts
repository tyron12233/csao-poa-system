import type { ApprovalEntry, FormEmail, ResponseField } from "../../../lib/form-email";

// Renders any formapprovals.com request (FCA, Pull-out, Sustainability, ...) in the same
// layout as the original email, from the label/value rows parsed out of it.

function escapeHtml(str: string): string {
    return String(str ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

// Only allow link schemes that came from the original email; the data arrives via URL params.
function safeHref(url: string): string {
    return /^(https?:|mailto:)/i.test(url.trim()) ? escapeHtml(url.trim()) : "#";
}

const COL1_STYLE = "box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px;width:45%";
const COL2_STYLE = "box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px";

function renderApprovalRows(history: ApprovalEntry[]) {
    return history
        .map((h) => `<tr><td><span style="font-weight:500">${escapeHtml(h.action)}</span> by <a class="no-link">${escapeHtml(h.email)}</a></td></tr>`)
        .join("");
}

function renderFieldValue(field: ResponseField) {
    const unusedLinks = [...(field.links || [])];
    const lines = (field.lines || []).map((line) => {
        // Lines whose text is a link label (e.g. "File Upload 1", the requestor email) become anchors.
        const idx = unusedLinks.findIndex((l) => l.label === line);
        if (idx === -1) return escapeHtml(line);
        const [link] = unusedLinks.splice(idx, 1);
        return `<a href="${safeHref(link.url)}">${escapeHtml(line)}</a>`;
    });
    unusedLinks.forEach((l) => lines.push(`<a href="${safeHref(l.url)}">${escapeHtml(l.label || l.url)}</a>`));
    return lines.join("<br>");
}

function renderFieldRows(fields: ResponseField[]) {
    return fields
        .map((f) => `<tr><td class="response-items-col1" style="${COL1_STYLE}">${escapeHtml(f.label)}</td><td class="response-items-col2" style="${COL2_STYLE}">${renderFieldValue(f)}</td></tr>`)
        .join("");
}

export function buildFormEmailHtml(email: FormEmail) {
    const styles = `
    <style type="text/css">
      body{ font-family: -apple-system,system-ui,sans-serif; box-sizing: border-box; font-size: 14px; margin: 0; }
      td{ font-family: -apple-system,system-ui,sans-serif; }
      p{ margin: 10px 0; }
      .response-items tr { page-break-inside: avoid; }
      @media only screen and (max-width: 640px) {
        body { padding: 0 !important; }
        .container { padding: 0 !important; width: 100% !important; }
        .content-wrap { font-size:12px !important; padding: 10px !important; }
        #title{ font-size:26px !important; }
        #responses { padding:3px !important; font-size:13px !important; }
        #approval-history-label{ font-size:16px !important; }
      }
    </style>
  `;

    return `
  ${styles}
  <br>
  <table class="body-wrap" align="center" cellpadding="0" cellspacing="0" style="max-width:650px">
    <tbody><tr>
      <td class="container">
        <div class="content" style="padding:0px">
          <table class="main" width="100%" cellpadding="0" cellspacing="0" bgcolor="#fff" style="border-radius:5px;background-color:#fff;border:1px solid #e9e9e9">
            <tbody><tr style="box-sizing:border-box;font-size:14px;margin:0">
              <td class="content-wrap" style="font-size:15px;text-align:center;padding:20px 0 0 0">
                REQUEST
                <a href="${safeHref(email.requestUrl)}">#${escapeHtml(email.requestNumber)}</a>
                | ${escapeHtml(email.requestDate)}
                <br>
                <table width="100%" cellpadding="0" cellspacing="0">
                  <tbody><tr>
                    <td id="title" class="content-block" valign="top" align="center" style="padding:50px 10px 50px 10px;font-size:32px;color:#000;line-height:1.2em;font-weight:600">
                      ${escapeHtml(email.headerTitle)}
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <table id="responses" cellpadding="0" cellspacing="0" style="text-align:left;width:100%;padding:0 20px 0 20px">
                        <tbody><tr>
                          <td style="padding-bottom:10px">
                            <p>The request is now <strong>${escapeHtml(email.statusLabel.toLowerCase())}</strong>.</p>
                            <table class="approval-history" width="100%" cellpadding="0" cellspacing="0" bgcolor="#fff" style="border:1px dashed #e0e0e0;padding:10px;margin:10px 0">
                              <tbody><tr><td>
                                <table width="100%" style="padding-bottom:8px">
                                  <tbody><tr>
                                    <td id="approval-history-label" style="font-weight:500;font-size:18px">Approval history</td>
                                    <td style="padding:0 0 0 5px"><div style="text-align:right"><span style="color:#0e6245;background-color:#cbf4c9;padding:3px 7px 3px 7px;border-radius:7px">${escapeHtml(email.statusLabel)}</span></div></td>
                                  </tr></tbody>
                                </table>
                                <table width="100%"><tbody>${renderApprovalRows(email.approvalHistory || [])}</tbody></table>
                              </td></tr></tbody>
                            </table>
                            <table class="response-items" cellpadding="0" cellspacing="0" style="width:100%">
                              <tbody>${renderFieldRows(email.fields || [])}</tbody>
                            </table>
                            <br>
                          </td>
                        </tr></tbody>
                      </table>
                    </td>
                  </tr>
                </tbody></table>
                <br>
                <table id="footer" width="100%" cellpadding="0" cellspacing="0" style="box-sizing:border-box;font-size:13px;padding:0 5px 20px;text-align:center;color:#767676">
                  <tbody><tr>
                    <td>
                      This is an automated email sent by <a href="https://formapprovals.com">formapprovals.com</a>; do
                      not reply to or forward this email. You are receiving this email because you are a workflow participant of this
                      request. Your form administrator is <a href="mailto:${escapeHtml(email.adminEmail)}">${escapeHtml(email.adminEmail)}</a>
                    </td>
                  </tr></tbody>
                </table>
              </td>
            </tr></tbody>
          </table>
        </div>
      </td>
    </tr></tbody>
  </table>`;
}
