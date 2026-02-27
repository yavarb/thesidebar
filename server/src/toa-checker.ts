import pdfParse from "pdf-parse";

interface ToaEntry {
  text: string;
  pages: string;
}

interface PageContent {
  pageNumber: number;
  text: string;
}

interface ToaCheckResult {
  entry: string;
  listedPages: string;
  actualPages: number[];
  status: "correct" | "incorrect" | "not_found" | "passim";
  details?: string;
}

export async function parsePageContent(pdfBase64: string): Promise<PageContent[]> {
  const buffer = Buffer.from(pdfBase64, "base64");
  const pages: PageContent[] = [];
  let currentPage = 0;

  await pdfParse(buffer, {
    pagerender: async function (pageData: any) {
      currentPage++;
      const textContent = await pageData.getTextContent();
      const text = textContent.items.map((item: any) => item.str).join(" ");
      pages.push({ pageNumber: currentPage, text });
      return text;
    },
  });

  return pages;
}

export function checkToaEntries(
  pages: PageContent[],
  entries: ToaEntry[]
): ToaCheckResult[] {
  const results: ToaCheckResult[] = [];

  for (const entry of entries) {
    if (entry.pages.toLowerCase().includes("passim")) {
      results.push({
        entry: entry.text,
        listedPages: entry.pages,
        actualPages: [],
        status: "passim",
        details: "Passim — appears throughout",
      });
      continue;
    }

    const listedNums = entry.pages
      .split(/[,\s]+/)
      .map((n) => parseInt(n))
      .filter((n) => !isNaN(n));

    const searchText = entry.text.replace(/[.,\s]+$/g, "").trim();
    const actualPages: number[] = [];

    for (const page of pages) {
      const normalizedPage = page.text.replace(/\s+/g, " ").toLowerCase();
      const normalizedSearch = searchText.replace(/\s+/g, " ").toLowerCase();

      if (
        normalizedPage.includes(normalizedSearch) ||
        normalizedPage.includes(normalizedSearch.substring(0, 50)) ||
        normalizedPage.includes(
          normalizedSearch.split(/\s+v\.\s+|,/)[0].trim()
        )
      ) {
        actualPages.push(page.pageNumber);
      }
    }

    if (actualPages.length === 0) {
      results.push({
        entry: entry.text,
        listedPages: entry.pages,
        actualPages,
        status: "not_found",
        details: "Citation not found in document body",
      });
    } else {
      const listedSet = new Set(listedNums);
      const actualSet = new Set(actualPages);
      const match =
        listedNums.length === actualPages.length &&
        listedNums.every((n) => actualSet.has(n));

      if (match) {
        results.push({
          entry: entry.text,
          listedPages: entry.pages,
          actualPages,
          status: "correct",
        });
      } else {
        const missing = actualPages.filter((p) => !listedSet.has(p));
        const extra = listedNums.filter((p) => !actualSet.has(p));
        let details = "";
        if (missing.length) details += `Missing pages: ${missing.join(", ")}. `;
        if (extra.length)
          details += `Listed but not found on: ${extra.join(", ")}. `;
        results.push({
          entry: entry.text,
          listedPages: entry.pages,
          actualPages,
          status: "incorrect",
          details: details.trim(),
        });
      }
    }
  }

  return results;
}
