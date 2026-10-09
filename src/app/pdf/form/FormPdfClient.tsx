"use client";

import { useSearchParams } from "next/navigation";
import { useEffect, useRef } from "react";
import { formEmailFromSearchParams } from "../../../lib/form-email";
import { buildFormEmailHtml } from "./form-email-html";

export default function FormPdfClient() {
    const searchParams = useSearchParams();
    const contentRef = useRef<HTMLIFrameElement>(null);

    const html = buildFormEmailHtml(formEmailFromSearchParams(searchParams as unknown as URLSearchParams));

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
            srcDoc={html}
        />
    );
}
