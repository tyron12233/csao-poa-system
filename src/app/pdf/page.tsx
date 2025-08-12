"use client";

import dynamic from "next/dynamic";
import { Suspense } from "react";

const PdfClient = dynamic(() => import("./PdfClient"), { ssr: false });

export default function Page() {
    return (
        <Suspense fallback={null}>
            <PdfClient />
        </Suspense>
    );
}