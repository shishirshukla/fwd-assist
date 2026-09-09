import type { Metadata } from "next";

import "./globals.css";

export const metadata: Metadata = {
  title: "Forward Guard — Outlook add-in",
  description:
    "Intercepts Send on forwarded Outlook messages and collects Priority, End Date, and Category.",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="h-full overflow-hidden">{children}</body>
    </html>
  );
}
