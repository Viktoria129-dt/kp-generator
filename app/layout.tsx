import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Генератор коммерческих предложений",
  description: "Создание КП на фирменном бланке из таблицы Excel",
  robots: {
    index: false,
    follow: false,
    nocache: true,
    googleBot: {
      index: false,
      follow: false,
      noimageindex: true,
    },
  },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ru"><body>{children}</body></html>;
}
