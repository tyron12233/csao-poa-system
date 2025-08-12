"use client";

import { useSearchParams } from "next/navigation";
import { useEffect, useRef } from "react";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

export default function PdfClient() {
    const searchParams = useSearchParams();
    const htmlString = searchParams.get("html") || "";
    const contentRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        if (!htmlString || !contentRef.current) return;

        // Wait for the DOM to render the HTML
        const id = setTimeout(async () => {
            const input = contentRef.current!;
            const canvas = await html2canvas(input);
            const imgData = canvas.toDataURL("image/png");
            const pdf = new jsPDF();
            const imgProps = pdf.getImageProperties(imgData);
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
            pdf.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
            pdf.save("document.pdf");
        }, 100);

        return () => clearTimeout(id);
    }, [htmlString]);

    return (
        <div ref={contentRef} style={{ position: "absolute", left: "-9999px", top: 0 }}>
            <span dangerouslySetInnerHTML={{ __html: htmlString }} />
        </div>
    );
}
