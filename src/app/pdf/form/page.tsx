"use client";

import dynamic from "next/dynamic";
import { Suspense } from "react";

const FormPdfClient = dynamic(() => import("./FormPdfClient"), { ssr: false });

export default function Page() {
    return (
        <Suspense fallback={null}>
            <FormPdfClient />
        </Suspense>
    );
}
