import type { Metadata } from "next";

import "./globals.css";

export const metadata: Metadata = {
  title: "Forward Guard — Outlook add-in",
  description:
    "Intercepts Send on forwarded Outlook messages and submits extracted letter details.",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="h-full overflow-hidden">{children}</body>
    </html>
  );
}
