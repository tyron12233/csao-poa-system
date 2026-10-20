"use client";

import { useEffect, useRef } from "react";
import { readPdfParams } from "../../lib/form-email";
import POAEmail, { buildPOAEmailHtml, POAEmailProps } from "./poa";

export default function PdfClient() {
    const searchParams = readPdfParams();
    const contentRef = useRef<HTMLIFrameElement>(null);

    const props: POAEmailProps = {
        requestNumber: searchParams.get("requestNumber") || "",
        requestUrl: searchParams.get("requestUrl") || "",
        requestDate: searchParams.get("requestDate") || "",
        headerTitle: searchParams.get("headerTitle") || "",
        statusLabel: searchParams.get("statusLabel") || "complete",
        approvalHistory: JSON.parse(searchParams.get("approvalHistory") || "[]"),
        requestorEmail: searchParams.get("requestorEmail") || "",
        emailAddress: searchParams.get("emailAddress") || "",
        organizationName: searchParams.get("organizationName") || "",
        activityTitle: searchParams.get("activityTitle") || "",
        startDate: searchParams.get("startDate") || "",
        endDate: searchParams.get("endDate") || "",
        implementationTime: searchParams.get("implementationTime") || "",
        targetParticipants: searchParams.get("targetParticipants") || "",
        targetNumberOfParticipants: searchParams.get("targetNumberOfParticipants") || "",
        estimatedActivityCost: searchParams.get("estimatedActivityCost") || "",
        unsdg: searchParams.get("unsdg") || "",
        rationale: searchParams.get("rationale") || "",
        objectives: JSON.parse(searchParams.get("objectives") || "[]"),
        mechanics: JSON.parse(searchParams.get("mechanics") || "[]"),
        speakers: JSON.parse(searchParams.get("speakers") || "[]"),
        programFlow: JSON.parse(searchParams.get("programFlow") || "[]"),
        budgetBreakdown: searchParams.get("budgetBreakdown") || "",
        budgetCharging: searchParams.get("budgetCharging") || "",
        participants: JSON.parse(searchParams.get("participants") || "[]"),
        invitationLinks: JSON.parse(searchParams.get("invitationLinks") || "[]"),
        implementationType: searchParams.get("implementationType") || "",
        venue: searchParams.get("venue") || "",
        preparedBy: searchParams.get("preparedBy") || "",
        fbName: searchParams.get("fbName") || "",
        position: searchParams.get("position") || "",
        adminEmail: searchParams.get("adminEmail") || "",
    };

    const html = buildPOAEmailHtml(props);

    useEffect(() => {
        // We need a slight delay to ensure all content is rendered before printing.
        const timer = setTimeout(() => {
            if (contentRef.current) {
                contentRef.current.contentWindow?.print();
            }
        }, 500);

        return () => clearTimeout(timer);
    }, []);

    return (
        <iframe
            ref={contentRef}
            style={{ width: "100%", height: "100vh", border: "none" }}
            // No allow-scripts: data comes from URL params, so the rendered email must never run code.
            // allow-same-origin lets us call print(); allow-modals lets the print dialog open.
            sandbox="allow-same-origin allow-modals allow-popups allow-popups-to-escape-sandbox"
            srcDoc={html}
        />
    );
}
