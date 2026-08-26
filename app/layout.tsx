import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Генератор коммерческих предложений",
  description: "Создание КП на фирменном бланке из таблицы Excel",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ru"><body>{children}</body></html>;
}
