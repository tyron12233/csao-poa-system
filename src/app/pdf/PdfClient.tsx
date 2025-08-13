"use client";

import { useSearchParams } from "next/navigation";
import { useEffect, useRef } from "react";
import POAEmail, { buildPOAEmailHtml, POAEmailProps } from "./poa";

export default function PdfClient() {
    const searchParams = useSearchParams();
    const contentRef = useRef<HTMLDivElement>(null);

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
            window.print();
        }, 500);

        return () => clearTimeout(timer);
    }, []);

    return (
        <iframe
            style={{ width: "100%", height: "100vh", border: "none" }}
            srcDoc={html}
        />
    );
}
