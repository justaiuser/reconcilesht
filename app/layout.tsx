import type { Metadata } from "next";
import type { ReactNode } from "react";
import { Manrope } from "next/font/google";
import "./globals.css";

const sans = Manrope({
  subsets: ["latin", "cyrillic"],
  variable: "--font-sans",
});

export const metadata: Metadata = {
  title: "Сверка актов",
  description:
    "Сравнение скана акта сверки с файлом Excel из учётной программы. Документы остаются на вашем компьютере.",
};

const themeScript = `(function(){try{var t=localStorage.getItem("sverka.theme");var d=t==="dark"||(t!=="light"&&matchMedia("(prefers-color-scheme: dark)").matches);document.documentElement.classList.toggle("dark",d);}catch(e){}})();`;

// pdf.js 6 expects Map/WeakMap.getOrInsertComputed (Chrome 135+); polyfill for older office PCs.
const mapPolyfill = `(function(){function p(o){if(typeof o.getOrInsertComputed==="function")return;Object.defineProperty(o,"getOrInsertComputed",{configurable:true,writable:true,value:function(k,f){if(this.has(k))return this.get(k);var v=f(k);this.set(k,v);return v;}});}try{p(Map.prototype);p(WeakMap.prototype);}catch(e){}})();`;

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="ru" className={`${sans.variable} h-full antialiased`} suppressHydrationWarning>
      <head>
        <script dangerouslySetInnerHTML={{ __html: mapPolyfill }} />
        <script dangerouslySetInnerHTML={{ __html: themeScript }} />
      </head>
      <body className="min-h-full font-sans">{children}</body>
    </html>
  );
}
