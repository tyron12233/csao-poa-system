"use client";

import { useEffect, useRef } from "react";
import { formEmailFromSearchParams, readPdfParams } from "../../../lib/form-email";
import { buildFormEmailHtml } from "./form-email-html";

export default function FormPdfClient() {
    const contentRef = useRef<HTMLIFrameElement>(null);

    const html = buildFormEmailHtml(formEmailFromSearchParams(readPdfParams()));

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
