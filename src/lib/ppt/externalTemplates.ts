import type { ExternalTemplateResult, ExternalTemplateSource } from "@/types/ppt";

const SEARCH_TIMEOUT_MS = 8500;
const USER_AGENT =
  "Mozilla/5.0 (compatible; AI-PPT-Studio/1.0; +https://localhost)";
const MAX_RESULTS_PER_SOURCE = 80;
const MAX_PAGES_PER_SOURCE = 6;
const DEFAULT_PAGE_SIZE = 12;
const MAX_PAGE_SIZE = 24;

interface SearchSourceResult {
  name: ExternalTemplateSource;
  status: "ok" | "error";
  message?: string;
  results: ExternalTemplateResult[];
}

interface SourceConfig {
  name: ExternalTemplateSource;
  searchUrl: (query: string) => string;
  pageUrl?: (query: string, page: number) => string;
  parse: (html: string, query: string) => ExternalTemplateResult[];
}

interface PreparedSearchQuery {
  originalQuery: string;
  searchQuery: string;
  relevanceTerms: string[];
}

interface QueryMapping {
  pattern: RegExp;
  search: string;
  terms: string[];
}

const queryMappings: QueryMapping[] = [
  { pattern: /圣诞|圣诞节|耶诞|christmas|xmas/i, search: "christmas", terms: ["christmas", "xmas", "holiday", "winter"] },
  { pattern: /新年|春节|过年|元旦|new year|lunar/i, search: "new year", terms: ["new", "year", "lunar", "spring", "festival"] },
  { pattern: /中秋|月饼|mid.?autumn|moon/i, search: "mid autumn", terms: ["mid", "autumn", "moon", "festival"] },
  { pattern: /万圣|halloween/i, search: "halloween", terms: ["halloween"] },
  { pattern: /情人|valentine/i, search: "valentine", terms: ["valentine", "love"] },
  {
    pattern: /商务|商业|公司|企业|汇报|报告|年终|总结|business|corporate|company|report|annual|summary/i,
    search: "business report",
    terms: ["business", "corporate", "company", "report", "annual", "summary", "finance", "strategy", "sales", "marketing"],
  },
  { pattern: /路演|融资|创业|pitch|startup|investor/i, search: "pitch deck", terms: ["pitch", "deck", "startup", "investor", "business"] },
  {
    pattern: /教育|教学|学校|课程|education|school|lesson|teach|student/i,
    search: "education",
    terms: ["education", "school", "lesson", "teach", "student", "academic"],
  },
  { pattern: /论文|答辩|毕业|学术|thesis|academic|research/i, search: "thesis", terms: ["thesis", "academic", "research", "university", "education"] },
  { pattern: /医学|医疗|健康|医院|medical|health|clinic|doctor|science/i, search: "medical", terms: ["medical", "health", "clinic", "doctor", "science"] },
  {
    pattern: /科技|技术|人工智能|智能|\bai\b|technology|tech|digital/i,
    search: "technology",
    terms: ["technology", "tech", "digital", "ai", "data", "innovation"],
  },
  { pattern: /数据|图表|信息图|infographic|chart|data/i, search: "data infographic", terms: ["data", "infographic", "chart", "dashboard"] },
  { pattern: /极简|简约|minimal|simple|clean/i, search: "minimal", terms: ["minimal", "minimalist", "simple", "clean"] },
  { pattern: /创意|艺术|设计|作品集|creative|art|design|portfolio/i, search: "creative design", terms: ["creative", "art", "design", "portfolio"] },
  { pattern: /营销|销售|市场|marketing|sales/i, search: "marketing", terms: ["marketing", "sales", "business"] },
  { pattern: /金融|财务|finance|financial/i, search: "finance", terms: ["finance", "financial", "business"] },
  { pattern: /产品|项目|方案|计划|product|project|proposal|plan/i, search: "project proposal", terms: ["project", "proposal", "plan", "product"] },
  { pattern: /旅游|旅行|景点|travel|tourism/i, search: "travel", terms: ["travel", "tourism", "trip"] },
  { pattern: /咖啡|coffee/i, search: "coffee", terms: ["coffee", "cafe"] },
  { pattern: /美食|餐饮|food|restaurant/i, search: "food", terms: ["food", "restaurant", "cooking"] },
  { pattern: /婚礼|婚庆|wedding/i, search: "wedding", terms: ["wedding", "love"] },
  { pattern: /体育|运动|健身|sport|fitness/i, search: "sports", terms: ["sport", "sports", "fitness"] },
  { pattern: /音乐|music/i, search: "music", terms: ["music"] },
  { pattern: /环保|可持续|environment|sustainability/i, search: "environment", terms: ["environment", "sustainability", "green"] },
];

function cleanText(value?: string) {
  return repairMojibake(decodeHtml(value ?? ""))
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function repairMojibake(value: string) {
  return value
    .replace(/\u00E2\u0080\u0099/g, "'")
    .replace(/\u00E2\u0080\u0098/g, "'")
    .replace(/\u00E2\u0080\u009C/g, '"')
    .replace(/\u00E2\u0080\u009D/g, '"')
    .replace(/\u00E2\u0080\u0093|\u00E2\u0080\u0094/g, "-")
    .replace(/\u00E2\u0080\u00A6/g, "...");
}

function decodeHtml(value: string) {
  return value
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCharCode(Number.parseInt(code, 16)))
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#039;|&apos;/g, "'")
    .replace(/&ndash;|&mdash;/g, "-")
    .replace(/&hellip;/g, "...")
    .replace(/&nbsp;/g, " ");
}

function absoluteUrl(url: string | undefined, base: string) {
  if (!url || url === "#") return undefined;
  try {
    return new URL(decodeHtml(url), base).toString();
  } catch {
    return undefined;
  }
}

function slugify(value: string) {
  return value
    .toLowerCase()
    .replace(/https?:\/\//g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 72);
}

function uniqueWords(words: string[]) {
  return Array.from(new Set(words.map((word) => word.toLowerCase()).filter(Boolean)));
}

function tokenizeSearchText(value: string) {
  return cleanText(value).toLowerCase().match(/[a-z0-9]+/g) ?? [];
}

function hasCjk(value: string) {
  return /[\u3400-\u9fff]/.test(value);
}

function prepareSearchQuery(query: string): PreparedSearchQuery {
  const originalQuery = query.trim() || "business";
  const asciiTerms = tokenizeSearchText(originalQuery);
  const matchedMappings = queryMappings.filter((mapping) => mapping.pattern.test(originalQuery));

  if (!matchedMappings.length) {
    return {
      originalQuery,
      searchQuery: hasCjk(originalQuery) && asciiTerms.length === 0 ? "" : originalQuery,
      relevanceTerms: uniqueWords(asciiTerms),
    };
  }

  const mappedSearchWords = matchedMappings.flatMap((mapping) => tokenizeSearchText(mapping.search));
  const searchWords = uniqueWords([...mappedSearchWords, ...asciiTerms]).slice(0, 5);

  return {
    originalQuery,
    searchQuery: searchWords.join(" "),
    relevanceTerms: uniqueWords([...matchedMappings.flatMap((mapping) => mapping.terms), ...searchWords]),
  };
}

function relevanceScore(result: ExternalTemplateResult, searchTerms: string[]) {
  if (!searchTerms.length) return 1;

  const title = cleanText(result.title).toLowerCase();
  const description = cleanText(result.description).toLowerCase();
  const tags = result.tags.map((tag) => cleanText(tag).toLowerCase()).join(" ");
  const url = decodeHtml(result.url).toLowerCase();

  return searchTerms.reduce((score, term) => {
    if (title.includes(term)) return score + 5;
    if (tags.includes(term)) return score + 3;
    if (description.includes(term)) return score + 2;
    if (url.includes(term)) return score + 1;
    return score;
  }, 0);
}

function filterAndRankResults(results: ExternalTemplateResult[], searchTerms: string[]) {
  if (!searchTerms.length) return results;

  return results
    .map((result, index) => ({ result, index, score: relevanceScore(result, searchTerms) }))
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score || a.index - b.index)
    .map(({ result }) => result);
}

function uniqueByUrl(results: ExternalTemplateResult[]) {
  const seen = new Set<string>();
  return results.filter((result) => {
    const key = result.url;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function mergeUniqueResults(existing: ExternalTemplateResult[], incoming: ExternalTemplateResult[]) {
  const seen = new Set(existing.map((result) => result.url));
  const merged = [...existing];
  for (const result of incoming) {
    if (seen.has(result.url)) continue;
    seen.add(result.url);
    merged.push(result);
  }
  return merged;
}

function splitArticles(html: string) {
  return html.match(/<article\b[\s\S]*?<\/article>/gi) ?? [];
}

function getAttr(html: string, attr: string) {
  const match = html.match(new RegExp(`${attr}=["']([^"']+)["']`, "i"));
  return match?.[1];
}

function extractFirstImage(article: string, base: string) {
  const img = article.match(/<img\b[\s\S]*?>/i)?.[0];
  return absoluteUrl(getAttr(img ?? "", "src"), base);
}

function extractTags(article: string) {
  return Array.from(article.matchAll(/<a\b[^>]*>([\s\S]*?)<\/a>/gi))
    .map((match) => cleanText(match[1]))
    .filter((tag) => tag.length > 1 && tag.length < 34)
    .slice(0, 5);
}

function parseSlidesCarnival(html: string): ExternalTemplateResult[] {
  return uniqueByUrl(
    splitArticles(html)
      .map((article) => {
        const titleAnchor = article.match(/<h3[^>]*class=["'][^"']*card-title[^"']*["'][\s\S]*?<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);
        const url = absoluteUrl(titleAnchor?.[1], "https://www.slidescarnival.com");
        const title = cleanText(titleAnchor?.[2]);
        if (!url || !title) return null;

        const pptxUrl = absoluteUrl(article.match(/href=["']([^"']+\/pptx)["']/i)?.[1], "https://www.slidescarnival.com");
        const googleSlidesUrl = absoluteUrl(
          article.match(/href=["']([^"']+\/google-slides)["']/i)?.[1],
          "https://www.slidescarnival.com",
        );
        const canvaUrl = absoluteUrl(article.match(/href=["']([^"']+\/canva)["']/i)?.[1], "https://www.slidescarnival.com");
        if (!pptxUrl && !googleSlidesUrl) return null;

        const slideCount = Number.parseInt(cleanText(article.match(/card-total-slides[\s\S]*?<\/img>\s*([^<]+)/i)?.[1]), 10);
        const aspectRatio = cleanText(article.match(/<span class=["']card-format["']>([\s\S]*?)<\/span>/i)?.[1]);
        const formats = [
          pptxUrl ? "PPTX" : "",
          googleSlidesUrl ? "Google Slides" : "",
          canvaUrl ? "Canva" : "",
        ].filter(Boolean);

        return {
          id: `slidescarnival-${slugify(url)}`,
          title,
          source: "SlidesCarnival" as const,
          url,
          thumbnail: extractFirstImage(article, "https://www.slidescarnival.com"),
          description: "SlidesCarnival free presentation template.",
          formats,
          aspectRatio: aspectRatio || undefined,
          slideCount: Number.isFinite(slideCount) ? slideCount : undefined,
          tags: extractTags(article),
          isFree: true,
          downloadUrl: pptxUrl ?? googleSlidesUrl,
          licenseNote: "Free template site; open the source page to confirm attribution and usage terms.",
        } satisfies ExternalTemplateResult;
      })
      .filter(Boolean) as ExternalTemplateResult[],
  ).slice(0, MAX_RESULTS_PER_SOURCE);
}

function parsePresentationGo(html: string): ExternalTemplateResult[] {
  return uniqueByUrl(
    splitArticles(html)
      .map((article) => {
        const titleAnchor = article.match(/<h2\b[\s\S]*?<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);
        const fallbackAnchor = article.match(/<a\b[^>]*href=["'](https:\/\/www\.presentationgo\.com\/presentation\/[^"']+)["']/i);
        const url = absoluteUrl(titleAnchor?.[1] ?? fallbackAnchor?.[1], "https://www.presentationgo.com");
        const title = cleanText(titleAnchor?.[2] ?? getAttr(article.match(/<img\b[\s\S]*?>/i)?.[0] ?? "", "alt"));
        if (!url || !title) return null;

        const description = cleanText(
          article.match(/gb-container-6990958b["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] ??
            getAttr(article.match(/<img\b[\s\S]*?>/i)?.[0] ?? "", "data-pin-description"),
        );
        const formatText = cleanText(article.match(/dynamic-term-class[\s\S]*?<\/p>/i)?.[0]);
        const aspectRatio = cleanText(article.match(/dl-aspect-ratio[^>]*>([\s\S]*?)<\/span>/i)?.[1]);

        return {
          id: `presentationgo-${slugify(url)}`,
          title,
          source: "PresentationGO" as const,
          url,
          thumbnail: extractFirstImage(article, "https://www.presentationgo.com"),
          description,
          formats: formatText.includes("PPTX") ? ["PPTX", "Google Slides"] : ["Google Slides"],
          aspectRatio: aspectRatio || undefined,
          tags: extractTags(article),
          isFree: true,
          licenseNote: "Free template site; open the source page to confirm attribution and usage terms.",
        } satisfies ExternalTemplateResult;
      })
      .filter(Boolean) as ExternalTemplateResult[],
  ).slice(0, MAX_RESULTS_PER_SOURCE);
}

function parsePptMon(html: string): ExternalTemplateResult[] {
  return uniqueByUrl(
    splitArticles(html)
      .map((article) => {
        const url = absoluteUrl(article.match(/location\.href=['"]([^'"]+)['"]/i)?.[1], "https://pptmon.com");
        const title = cleanText(article.match(/<h3>([\s\S]*?)<\/h3>/i)?.[1]);
        if (!url || !title) return null;
        const thumbnail = extractFirstImage(article, "https://pptmon.com");
        if (!thumbnail) return null;

        return {
          id: `pptmon-${slugify(url)}`,
          title,
          source: "PPTMON" as const,
          url,
          thumbnail,
          description: cleanText(article.match(/<p>([\s\S]*?)<\/p>/i)?.[1]),
          formats: ["PPTX", "Google Slides"],
          aspectRatio: "16:9",
          tags: extractTags(article),
          isFree: true,
          licenseNote: "Free template site; open the source page to confirm attribution and usage terms.",
        } satisfies ExternalTemplateResult;
      })
      .filter(Boolean) as ExternalTemplateResult[],
  ).slice(0, MAX_RESULTS_PER_SOURCE);
}

function parseSlidesgo(html: string): ExternalTemplateResult[] {
  const cards = html.match(/<div class=["']theme_post["'][\s\S]*?(?=<div class=["']theme_post["']|<\/main>|<\/body>)/gi) ?? [];

  return uniqueByUrl(
    cards
      .map((card) => {
        if (/data-premium=["']premium["']/i.test(card)) return null;
        const url = absoluteUrl(card.match(/<a\b[^>]*href=["'](https:\/\/slidesgo\.com\/theme\/[^"']+)["']/i)?.[1], "https://slidesgo.com");
        const title =
          cleanText(card.match(/handleEvents\([^,]+,\s*["']([^"']+)["']/i)?.[1]) ||
          cleanText(getAttr(card.match(/<img\b[\s\S]*?>/i)?.[0] ?? "", "alt")).replace(/\s*presentation template\s*$/i, "");
        if (!url || !title) return null;
        if (/rs=home-/i.test(decodeHtml(url))) return null;

        return {
          id: `slidesgo-${slugify(url)}`,
          title,
          source: "Slidesgo" as const,
          url,
          thumbnail: extractFirstImage(card, "https://slidesgo.com"),
          description: "Free Slidesgo presentation template.",
          formats: ["PPTX", "Google Slides"],
          aspectRatio: "16:9",
          tags: [],
          isFree: true,
          licenseNote: "Free result filtered from Slidesgo; open the source page to confirm download limits and attribution terms.",
        } satisfies ExternalTemplateResult;
      })
      .filter(Boolean) as ExternalTemplateResult[],
  ).slice(0, MAX_RESULTS_PER_SOURCE);
}

const sources: SourceConfig[] = [
  {
    name: "SlidesCarnival",
    searchUrl: (query) => {
      const category = categoryForSlidesCarnival(query);
      return category
        ? `https://www.slidescarnival.com/tag/${category}`
        : `https://www.slidescarnival.com/search/${encodeURIComponent(query)}`;
    },
    pageUrl: (query, page) => {
      const category = categoryForSlidesCarnival(query);
      return category
        ? `https://www.slidescarnival.com/tag/${category}/page/${page}`
        : `https://www.slidescarnival.com/search/${encodeURIComponent(query)}/page/${page}`;
    },
    parse: parseSlidesCarnival,
  },
  {
    name: "PresentationGO",
    searchUrl: (query) => `https://www.presentationgo.com/?s=${encodeURIComponent(query)}`,
    pageUrl: (query, page) => `https://www.presentationgo.com/page/${page}/?s=${encodeURIComponent(query)}`,
    parse: parsePresentationGo,
  },
  {
    name: "PPTMON",
    searchUrl: (query) => `https://pptmon.com/?s=${encodeURIComponent(query)}`,
    pageUrl: (query, page) => `https://pptmon.com/page/${page}/?s=${encodeURIComponent(query)}`,
    parse: parsePptMon,
  },
  {
    name: "Slidesgo",
    searchUrl: (query) => `https://slidesgo.com/search?search=${encodeURIComponent(query)}`,
    parse: parseSlidesgo,
  },
];

function categoryForSlidesCarnival(query: string) {
  const normalized = query.toLowerCase();
  if (/business|report|strategy|sales|finance|company|corporate|pitch|startup/.test(normalized)) return "business";
  if (/education|school|lesson|teach|student|academic|thesis/.test(normalized)) return "education";
  if (/medical|health|clinic|doctor|science/.test(normalized)) return "medical";
  if (/minimal|simple|clean/.test(normalized)) return "minimalist";
  if (/creative|design|portfolio|art/.test(normalized)) return "creative";
  return "";
}

async function fetchSource(source: SourceConfig, query: string): Promise<SearchSourceResult> {
  const loadPage = async (url: string) => {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), SEARCH_TIMEOUT_MS);
    try {
      return await fetch(url, {
      headers: {
        "user-agent": USER_AGENT,
        accept: "text/html,application/xhtml+xml",
      },
      signal: controller.signal,
      next: { revalidate: 1800 },
    });
    } finally {
      clearTimeout(timeout);
    }
  };

  try {
    let results: ExternalTemplateResult[] = [];
    let page = 1;
    let lastError = "";

    while (page <= MAX_PAGES_PER_SOURCE && results.length < MAX_RESULTS_PER_SOURCE) {
      const pageUrl = page === 1 ? source.searchUrl(query) : source.pageUrl?.(query, page);
      if (!pageUrl) break;

      const response = await loadPage(pageUrl);
      if (!response.ok) {
        lastError = `${response.status} ${response.statusText}`;
        break;
      }

      const html = await response.text();
      const pageResults = source.parse(html, query);
      const merged = mergeUniqueResults(results, pageResults);
      const added = merged.length - results.length;
      results = merged.slice(0, MAX_RESULTS_PER_SOURCE);
      if (added === 0 || pageResults.length === 0) break;
      page += 1;
    }

    return {
      name: source.name,
      status: results.length || !lastError ? "ok" : "error",
      message: results.length ? undefined : lastError || undefined,
      results,
    };
  } catch (error) {
    return {
      name: source.name,
      status: "error",
      message: error instanceof Error ? error.message : "Search failed.",
      results: [],
    };
  }
}

export async function searchExternalTemplates(query: string) {
  return searchExternalTemplatesPage(query, 1, DEFAULT_PAGE_SIZE);
}

export async function searchExternalTemplatesPage(query: string, page = 1, pageSize = DEFAULT_PAGE_SIZE) {
  const preparedQuery = prepareSearchQuery(query);
  const safePage = Math.max(1, Math.floor(page));
  const safePageSize = Math.min(MAX_PAGE_SIZE, Math.max(1, Math.floor(pageSize)));

  if (!preparedQuery.searchQuery) {
    return {
      query: preparedQuery.originalQuery,
      results: [],
      pagination: {
        page: safePage,
        pageSize: safePageSize,
        total: 0,
        totalPages: 0,
        hasNextPage: false,
        hasPreviousPage: false,
      },
      sources: sources.map(({ name }) => ({
        name,
        status: "ok" as const,
        message: "No English template-site keyword mapping for this query.",
      })),
    };
  }

  const settled = await Promise.all(sources.map((source) => fetchSource(source, preparedQuery.searchQuery)));
  const results = filterAndRankResults(
    settled.flatMap((source) => source.results),
    preparedQuery.relevanceTerms,
  );
  const total = results.length;
  const totalPages = Math.ceil(total / safePageSize);
  const currentPage = totalPages ? Math.min(safePage, totalPages) : 1;
  const start = (currentPage - 1) * safePageSize;
  const pagedResults = results.slice(start, start + safePageSize);

  return {
    query: preparedQuery.originalQuery,
    searchQuery: preparedQuery.searchQuery,
    results: pagedResults,
    pagination: {
      page: currentPage,
      pageSize: safePageSize,
      total,
      totalPages,
      hasNextPage: currentPage < totalPages,
      hasPreviousPage: currentPage > 1 && totalPages > 0,
    },
    sources: settled.map(({ name, status, message }) => ({ name, status, message })),
  };
}
