"use client";

import React from "react";

export type ApprovalEntry = {
    action: string; // e.g., "Approved", "Recommended", "Copy Sent"
    email: string; // e.g., "user@example.com"
};

export type ProgramFlowItem = {
    time: string; // e.g., "3:00 - 3:05"
    activity: string; // e.g., "Opening Prayer"
    inCharge: string; // e.g., "Mikayla Buno"
};

export type InvitationLink = {
    label: string; // e.g., "File Upload 1"
    url: string; // e.g., "https://..."
};

export interface POAEmailProps {
    // Header/request info
    requestNumber: string | number; // e.g., 3
    requestUrl: string; // link to the request tracking page
    requestDate: string; // formatted, e.g., "JUL 17, 2025"
    headerTitle: string; // e.g., "DEVELOPERS SOCIETY- POA"

    // Status / approval history
    statusLabel: string; // e.g., "Complete"
    approvalHistory: ApprovalEntry[];

    // Core details
    requestorEmail: string; // "developers.society@dlsl.edu.ph"
    emailAddress: string; // usually same as requestorEmail
    organizationName: string; // e.g., "DLSL Developers Society"
    activityTitle: string; // e.g., "Project Foundations: ..."
    startDate: string; // e.g., "July 22,2025"
    endDate: string; // e.g., "July 22,2025"
    implementationTime: string; // e.g., "3:00 PM - 5:00 PM"
    targetParticipants: string; // e.g., "DLSL Developers Society Officers"
    targetNumberOfParticipants: string | number; // e.g., 10
    estimatedActivityCost: string; // e.g., "₱ 0.00"
    unsdg: string; // e.g., "SDG 4: ... and SDG 9: ..."
    rationale: string; // paragraph
    objectives: string[]; // list of objectives
    mechanics: string[]; // list of mechanics/guidelines
    speakers: string | string[]; // names
    programFlow: ProgramFlowItem[]; // flow items
    budgetBreakdown: string; // e.g., "N/A"
    budgetCharging: string; // e.g., "N/A"
    participants: string[]; // facilitators and participants (line-separated)
    invitationLinks: InvitationLink[]; // list of links
    implementationType: string; // e.g., "Face to Face"
    venue: string; // e.g., "Mabini Building"
    preparedBy: string; // e.g., "Mikayla ... Buno"
    fbName: string; // e.g., "Mikayla Buno"
    position: string; // e.g., "Executive Secretary"

    // Footer/admin
    adminEmail: string; // e.g., "csao.poa@dlsl.edu.ph"
}

function escapeHtml(str: string): string {
    return str
        .replaceAll(/&/g, "&amp;")
        .replaceAll(/</g, "&lt;")
        .replaceAll(/>/g, "&gt;")
        .replaceAll(/\"/g, "&quot;")
        .replaceAll(/'/g, "&#39;");
}

function brJoin(items: string[]) {
    return items.map((s) => escapeHtml(s)).join("<br>");
}

function numberedBrJoin(items: string[]) {
    return items
        .map((s, i) => `${i + 1}. ${escapeHtml(s)}`)
        .join("<br>");
}

function renderApprovalRows(history: ApprovalEntry[]) {
    return history
        .map(
            (h) =>
                `<tr><td><span style="font-weight:500">${escapeHtml(
                    h.action
                )}</span> by <a class="no-link">${escapeHtml(h.email)}</a>  </td></tr>`
        )
        .join("");
}

function renderProgramFlow(flow: ProgramFlowItem[]) {
    return flow
        .map(
            (f) =>
                `${escapeHtml(f.time)}    ${escapeHtml(f.activity)}   ${escapeHtml(
                    f.inCharge
                )}`
        )
        .join("<br>");
}

function renderParticipantsList(list: string[]) {
    return list.map((p) => escapeHtml(p)).join("<br>");
}

function renderInvitationLinks(links: InvitationLink[]) {
    if (!links.length) return "";
    return links
        .map((l, idx) => `<a href="${l.url}">${escapeHtml(l.label || `File Upload ${idx + 1
            }`)}</a>`)
        .join("<br>");
}

export function buildPOAEmailHtml(props: POAEmailProps) {
    const styles = `
    <style type="text/css">
      body{
          font-family: -apple-system,system-ui,sans-serif; 
          box-sizing: border-box; 
          font-size: 14px; 
          margin: 0;
      }
      td{ font-family: -apple-system,system-ui,sans-serif; }
      p{ margin: 10px 0; }
      ul .ql-indent-1{ margin-left: 45px; }
      ul .ql-indent-2{ margin-left: 75px; }
      ul .ql-indent-3{ margin-left: 105px; }
      @media only screen and (max-width: 640px) {
        body { padding: 0 !important; }
        .container { padding: 0 !important; width: 100% !important; }
        .content-wrap { font-size:12px !important; padding: 10px !important; }
        #title{ font-size:26px !important; }
        #responses { padding:3px !important; font-size:13px !important; }
        #approval-history-label{ font-size:16px !important; }
        .button{ font-size:12px !important; padding: 7px 7px 3px 7px !important; border: 1px solid black !important; }
      }
    </style>
  `;

    const reqNumberDisplay = `#${props.requestNumber}`;
    const approvalRows = renderApprovalRows(props.approvalHistory || []);
    const objectivesHtml = numberedBrJoin(props.objectives || []);
    const mechanicsHtml = numberedBrJoin(props.mechanics || []);
    const speakersHtml = Array.isArray(props.speakers)
        ? brJoin(props.speakers)
        : escapeHtml(props.speakers);
    const programFlowHtml = renderProgramFlow(props.programFlow || []);
    const participantsHtml = renderParticipantsList(props.participants || []);
    const invitationLinksHtml = renderInvitationLinks(props.invitationLinks || []);

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
                <a href="${props.requestUrl}">${reqNumberDisplay}</a>
                | ${escapeHtml(props.requestDate)}
                <br>
                <table width="100%" cellpadding="0" cellspacing="0">
                  <tbody><tr>
                    <td id="title" class="content-block" valign="top" align="center" style="padding:50px 10px 50px 10px;font-size:32px;color:#000;line-height:1.2em;font-weight:600">
                      ${escapeHtml(props.headerTitle)}
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <table id="responses" cellpadding="0" cellspacing="0" style="text-align:left;width:100%;padding:0 20px 0 20px">
                        <tbody><tr>
                          <td style="padding-bottom:10px">
                            <p>The request is now <strong>${escapeHtml(
        props.statusLabel
    )}</strong>.</p>
                            <p>
                              <table class="approval-history" width="100%" cellpadding="0" cellspacing="0" bgcolor="#fff">
                                <tbody><tr><td>
                                  <table width="100%" style="padding-bottom:8px">
                                    <tbody><tr>
                                      <td id="approval-history-label" style="font-weight:500;font-size:18px">Approval history</td>
                                      <td style="padding:0 0 0 5px"><div style="text-align:right"><span style="color:#0e6245;background-color:#cbf4c9;padding:3px 7px 3px 7px;border-radius:7px">${escapeHtml(
        props.statusLabel
    )}</span></div></td>
                                    </tr></tbody>
                                  </table>
                                  <table width="100%"><tbody>${approvalRows}</tbody></table>
                                </td></tr></tbody>
                              </table>
                            </p>
                            <p>
                              <table class="response-items" cellpadding="0" cellspacing="0" style="width:100%">
                                <tbody>
                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Requestor:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px"><a href="mailto:${props.requestorEmail}">${escapeHtml(
        props.requestorEmail
    )}</a></td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">E-mail Address:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px"><a href="mailto:${props.emailAddress}">${escapeHtml(
        props.emailAddress
    )}</a></td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Name of Organization:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.organizationName
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Title of Activity:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.activityTitle
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Start Date of Implementation:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.startDate
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">End Date of Implementation:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.endDate
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Time of Implementation:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.implementationTime
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Target Participants:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.targetParticipants
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Target Number of Participants:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        String(props.targetNumberOfParticipants)
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Estimated Activity Cost:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.estimatedActivityCost
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">UNSDG:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.unsdg
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Rationale/Brief Description:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.rationale
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Objectives:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${objectivesHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Mechanics/Guidelines:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${mechanicsHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Name of Speakers::</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${speakersHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Program Flow::</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${programFlowHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Budget Breakdown:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.budgetBreakdown
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Budget Charging (CSAO Depository or Student Collection):</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.budgetCharging
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">List of Facilitators and Participants:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${participantsHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Letter of Invitation (If any) Sample Design/PubMaterial/Sample Video:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${invitationLinksHtml}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Type of Implementation (Online -Social Media Posting; Google Meet; Zoom etc)/Face to Face:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.implementationType
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Venue/Platform:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.venue
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Prepared by::</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.preparedBy
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">FBName:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.fbName
    )}</td></tr>

                                  <tr><td class="response-items-col1" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0;min-width:130px">Position/Designation:</td>
                                  <td class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px">${escapeHtml(
        props.position
    )}</td></tr>

                                  <tr><td colspan="2" class="response-items-col2" style="box-sizing:border-box;vertical-align:top;border-top-width:1px;border-top-color:#f4f4f4;border-top-style:solid;margin:0;padding:5px 0 0 10px;min-width:130px"></td></tr>
                                </tbody>
                              </table>
                            </p>
                            <br>
                          </td>
                        </tr></tbody>
                      </table>
                      <table width="100%">
                        <tbody><tr>
                          <td id="buttons" style="text-align:center;padding-bottom:20px">
                          </td>
                        </tr></tbody>
                      </table>
                    </td>
                  </tr>
                </tbody></table>
                <br>
                <table id="footer" width="100%" cellpadding="0" cellspacing="0" style="box-sizing:border-box;font-size:13px;padding:0 5px;text-align:center;color:#767676">
                  <tbody><tr>
                    <td>
                      This is an automated email sent by <a href="https://formapprovals.com">formapprovals.com</a>; do
                      not reply to or forward this email. You are receiving this email because you are a workflow participant of this
                      request. Your form administrator is <a href="mailto:${props.adminEmail}">${escapeHtml(
        props.adminEmail
    )}</a>
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

export default function POAEmail(props: POAEmailProps) {
    const html = buildPOAEmailHtml(props);
    return <div dangerouslySetInnerHTML={{ __html: html }} />;
}