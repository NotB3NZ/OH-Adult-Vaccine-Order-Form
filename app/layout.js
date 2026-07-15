import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({
  variable: "--font-inter",
  subsets: ["latin"],
  display: "swap",
});

export const metadata = {
  title: "Adult Vaccines Order Form | OH Clinic",
  description:
    "Order adult vaccines quickly and easily. Choose from 13 available vaccines, manage deliveries, and submit your order in one simple form.",
  keywords: "vaccines, adult vaccines, order form, clinic, healthcare",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en" className={`${inter.variable} h-full antialiased`}>
      <body className="min-h-full flex flex-col bg-background text-foreground font-sans">
        {children}
      </body>
    </html>
  );
}
