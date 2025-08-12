"use client";

import { useSearchParams } from "next/navigation";
import { useEffect, useRef } from "react";

export default function PdfClient() {
    const searchParams = useSearchParams();
    const htmlString = searchParams.get("html") || "";
    const contentRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        if (!htmlString || !contentRef.current) return;

        window.print();
    }, [htmlString]);

    return (
        <div ref={contentRef} className="flex justify-center">
            <span dangerouslySetInnerHTML={{ __html: htmlString }} />
        </div>
    );
}
