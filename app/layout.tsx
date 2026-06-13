import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "מערכת סיווג פעולות פיננסיות",
  description: "העלאת קבצי Excel וסיווג תנועות בנק וכרטיסי אשראי",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="he" dir="rtl" className="h-full antialiased" suppressHydrationWarning>
      <body className="min-h-full flex flex-col" suppressHydrationWarning>
        {children}
      </body>
    </html>
  );
}
