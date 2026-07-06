import Link from "next/link";

export default function NotFound() {
  return (
    <main className="flex min-h-screen w-full items-center justify-center px-4 py-10">
      <section className="w-full max-w-lg rounded-card bg-surface p-8 text-center shadow-card sm:p-12">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img
          src="https://res.cloudinary.com/dnpxugbk9/image/upload/v1783013365/logo_lrsktq.png"
          alt="Benny Vazana"
          className="mx-auto h-14 w-44 object-contain"
        />

        <p className="mt-8 text-7xl font-bold tracking-tight text-accent sm:text-8xl">404</p>

        <h1 className="mt-6 text-2xl font-semibold text-text-strong sm:text-3xl">
          הדף לא נמצא
        </h1>
        <p className="mx-auto mt-3 max-w-sm text-base leading-relaxed text-text">
          מצטערים, הדף שחיפשת לא קיים, הוסר או שהקישור שגוי.
        </p>

        <Link
          href="/"
          className="mt-8 inline-flex h-12 items-center justify-center rounded-pill bg-primary px-8 text-base font-semibold text-surface transition-colors hover:bg-primary-hover"
        >
          חזרה לדף הבית
        </Link>
      </section>
    </main>
  );
}
