import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const site = "https://www.michealsheehy.com";
const personId = `${site}/about.html#person`;
const websiteId = `${site}/#website`;
const updated = "2026-07-28";

const articles = {
  "articles/chapter-1.html": {
    title: "The Industry Solved the Wrong Problem: Transaction Monitoring Effectiveness | Micheal Sheehy",
    headline: "The Industry Solved the Wrong Problem",
    description: "Why transaction monitoring must move beyond false-positive reduction toward risk coverage, detection effectiveness and adaptive governance.",
    published: "2026-07-24",
    image: `${site}/assets/chapter-1-graphic.png`,
    imageAlt: "Traditional compliance alerts evolving into an adaptive compliance system",
  },
  "articles/chapter-2.html": {
    title: "The Evidence Is Already Here: Regulatory Intelligence for Compliance | Micheal Sheehy",
    headline: "The Evidence Is Already Here",
    description: "How compliance teams can turn enforcement actions and regulatory findings into control changes, organizational learning and stronger governance.",
    published: "2026-07-24",
  },
  "articles/chapter-3.html": {
    title: "Every Model Begins Drifting on Day One: Model Drift in Compliance | Micheal Sheehy",
    headline: "Every Model Begins Drifting on Day One",
    description: "How financial institutions can detect data, concept and performance drift through continuous monitoring, clear thresholds and accountable governance.",
    published: "2026-07-24",
  },
  "articles/chapter-4.html": {
    title: "The Adaptive Compliance Maturity Model: Five Levels of Governance | Micheal Sheehy",
    headline: "The Adaptive Compliance Maturity Model",
    description: "A five-level maturity model for assessing whether compliance governance is reactive, periodic, continuously monitored, adaptive or intelligence-led.",
    published: "2026-07-24",
  },
  "articles/chapter-5.html": {
    title: "AI Won’t Replace Compliance: AI Governance and Human Oversight | Micheal Sheehy",
    headline: "AI Won’t Replace Compliance. It Will Replace Static Governance.",
    description: "A practical framework for AI governance in compliance, including human oversight, explainability, testing, incident response and accountability.",
    published: "2026-07-24",
  },
  "articles/chapter-6.html": {
    title: "Designing an Adaptive Compliance Organization: Global Operating Model | Micheal Sheehy",
    headline: "Designing an Adaptive Compliance Organization",
    description: "How to design a global compliance operating model that connects governance, regional intelligence, product compliance and specialist expertise.",
    published: "2026-07-24",
    image: `${site}/assets/chapters-6-8.png`,
    imageAlt: "Visual summary of an adaptive compliance operating model",
  },
  "articles/chapter-7.html": {
    title: "Measuring What Matters: Compliance Metrics for Adaptive Governance | Micheal Sheehy",
    headline: "Measuring What Matters",
    description: "Compliance metrics for measuring risk coverage, model health, customer friction, adaptability and governance response—not only operational activity.",
    published: "2026-07-24",
    image: `${site}/assets/chapters-6-8.png`,
    imageAlt: "Visual summary of adaptive compliance metrics and governance",
  },
  "articles/chapter-8.html": {
    title: "Adaptive Compliance as a Competitive Advantage: Product, Growth and Risk | Micheal Sheehy",
    headline: "Adaptive Compliance as a Competitive Advantage",
    description: "How adaptive compliance supports product development, market expansion, risk appetite, customer trust and better executive decisions.",
    published: "2026-07-24",
    image: `${site}/assets/chapters-6-8.png`,
    imageAlt: "Visual summary of compliance as a strategic capability",
  },
  "articles/compliance-advantage.html": {
    headline: "Compliance as a Competitive Advantage",
    description: "How mature compliance creates trust, protects market access, improves product design and allows payments and fintech companies to move with confidence.",
    published: "2026-07-26",
  },
  "articles/trust-at-the-speed-of-money.html": {
    headline: "Trust at the Speed of Money",
    published: "2026-07-27",
  },
  "articles/the-criminal-is-already-in-production.html": {
    headline: "The Criminal Is Already in Production",
    published: "2026-07-28",
  },
  "articles/global-compliance-without-the-global-bottleneck.html": {
    headline: "Global Compliance Without the Global Bottleneck",
    published: "2026-07-28",
  },
  "articles/your-new-platform-is-not-a-transformation.html": {
    headline: "Your New Platform Is Not a Transformation",
    published: "2026-07-28",
  },
  "articles/the-safest-customer-is-not-the-customer-you-refused-to-understand.html": {
    headline: "The Safest Customer Is Not the Customer You Refused to Understand",
    published: "2026-07-28",
  },
  "articles/global-standards-local-proof.html": {
    headline: "Global Standards, Local Proof",
    published: "2026-07-24",
  },
  "articles/global-kyc-not-translation.html": {
    headline: "Global KYC Is Not a Translation Exercise",
    published: "2026-07-24",
  },
  "articles/ai-agents-compliance.html": {
    headline: "What AI Agents Will Actually Change in Compliance",
    published: "2026-07-24",
  },
  "articles/c-suite-leadership.html": {
    headline: "What Nobody Tells You About Joining the C-Suite",
    published: "2026-07-24",
  },
  "articles/the-code-is-real-the-identity-is-fake.html": {
    headline: "The Code Is Real. The Identity Is Fake.",
    published: "2026-07-27",
  },
  "articles/personal-side-of-being-a-cco.html": {
    headline: "The Personal Side of Being a Chief Compliance Officer",
    published: "2026-07-27",
  },
};

const caseStudies = {
  "case-studies/global-kyc.html": {
    headline: "Building a Global KYC Operating Model",
    description: "How common infrastructure can create KYC consistency without erasing local regulatory context, documentary practices or customer evidence.",
  },
  "case-studies/global-operations.html": {
    headline: "Scaling Compliance Without Creating a Global Bottleneck",
    description: "How to design a distributed compliance organization that increases resilience while preserving one global standard and accountable local judgment.",
  },
  "case-studies/modernizing-transaction-monitoring.html": {
    headline: "Modernizing Transaction Monitoring and Model Governance",
    description: "How to connect detection technology, model management, investigations and governance into one adaptive transaction-monitoring capability.",
  },
  "case-studies/sanctions-screening-transformation.html": {
    headline: "Rebuilding Sanctions Screening Around Effectiveness",
    description: "How sanctions-screening transformation can improve control quality, customer experience and operational efficiency together.",
  },
};

const canonicalOnly = [
  "adaptive-compliance.html",
  "adaptive-compliance-manuscript.html",
  "case-studies.html",
];

function read(relative) {
  return fs.readFileSync(path.join(root, relative), "utf8");
}

function write(relative, html) {
  fs.writeFileSync(path.join(root, relative), html);
}

function pageUrl(relative) {
  return `${site}/${relative}`;
}

function textFromTag(html, tag) {
  const match = html.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`, "i"));
  return match ? match[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim() : "";
}

function attribute(html, pattern) {
  const match = html.match(pattern);
  return match ? match[1] : "";
}

function setTitle(html, title) {
  if (!title) return html;
  return html.replace(/<title>[\s\S]*?<\/title>/i, `<title>${title}</title>`);
}

function setDescription(html, description) {
  if (!description) return html;
  const tag = `<meta name="description" content="${description}"/>`;
  if (/<meta\s+name=["']description["'][^>]*>/i.test(html)) {
    return html.replace(/<meta\s+name=["']description["'][^>]*>/i, tag);
  }
  if (/<meta\s+name=["']viewport["'][^>]*>/i.test(html)) {
    return html.replace(/(<meta\s+name=["']viewport["'][^>]*>)/i, `$1\n${tag}`);
  }
  return html.replace(/<head>/i, `<head>\n${tag}`);
}

function setCanonical(html, canonical) {
  const tag = `<link rel="canonical" href="${canonical}"/>`;
  if (/<link\s+rel=["']canonical["'][^>]*>/i.test(html)) {
    return html.replace(/<link\s+rel=["']canonical["'][^>]*>/i, tag);
  }
  return html.replace(/(<title>[\s\S]*?<\/title>)/i, `$1\n${tag}`);
}

function ensureIcon(html, relativePrefix) {
  if (/<link\s+rel=["']icon["'][^>]*>/i.test(html)) return html;
  return html.replace(/(<link\s+rel=["']canonical["'][^>]*>)/i, `$1\n<link rel="icon" href="${relativePrefix}favicon.svg" type="image/svg+xml"/>`);
}

function removeExistingSocial(html) {
  return html
    .replace(/\s*<meta\s+property=["']og:[^"']+["'][^>]*>/gi, "")
    .replace(/\s*<meta\s+name=["']twitter:[^"']+["'][^>]*>/gi, "")
    .replace(/\s*<meta\s+property=["']article:(?:published_time|modified_time)["'][^>]*>/gi, "");
}

function setSocial(html, data, canonical, type = "article") {
  const title = data.title || textFromTag(html, "title");
  const description = data.description || attribute(html, /<meta\s+name=["']description["'][^>]*content=["']([^"']+)/i);
  const image = data.image || attribute(html, /<meta\s+property=["']og:image["'][^>]*content=["']([^"']+)/i) || `${site}/micheal-sheehy-headshot.jpg`;
  const imageAlt = data.imageAlt || data.headline || title.replace(/\s*\|\s*Micheal Sheehy\s*$/i, "");
  html = removeExistingSocial(html);
  const block = [
    `<meta property="og:type" content="${type}"/>`,
    `<meta property="og:site_name" content="Micheal Sheehy"/>`,
    `<meta property="og:title" content="${title}"/>`,
    `<meta property="og:description" content="${description}"/>`,
    `<meta property="og:url" content="${canonical}"/>`,
    `<meta property="og:image" content="${image}"/>`,
    `<meta property="og:image:alt" content="${imageAlt}"/>`,
    data.published ? `<meta property="article:published_time" content="${data.published}"/>` : "",
    data.published ? `<meta property="article:modified_time" content="${updated}"/>` : "",
    `<meta name="twitter:card" content="summary_large_image"/>`,
    `<meta name="twitter:title" content="${title}"/>`,
    `<meta name="twitter:description" content="${description}"/>`,
    `<meta name="twitter:image" content="${image}"/>`,
  ].filter(Boolean).join("\n");
  return html.replace(/(<link\s+rel=["']canonical["'][^>]*>)/i, `$1\n${block}`);
}

function setSchema(html, schema) {
  const block = `<script type="application/ld+json">${JSON.stringify(schema)}</script>`;
  if (/<script\s+type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/i.test(html)) {
    return html.replace(/<script\s+type=["']application\/ld\+json["'][^>]*>[\s\S]*?<\/script>/i, block);
  }
  return html.replace(/<\/head>/i, `${block}\n</head>`);
}

function articleSchema(data, canonical, html) {
  const description = data.description || attribute(html, /<meta\s+name=["']description["'][^>]*content=["']([^"']+)/i);
  const image = data.image || attribute(html, /<meta\s+property=["']og:image["'][^>]*content=["']([^"']+)/i) || `${site}/micheal-sheehy-headshot.jpg`;
  return {
    "@context": "https://schema.org",
    "@type": "Article",
    "@id": `${canonical}#article`,
    headline: data.headline,
    description,
    image,
    datePublished: data.published,
    dateModified: updated,
    author: {
      "@type": "Person",
      "@id": personId,
      name: "Micheal Sheehy",
      url: `${site}/about.html`,
    },
    publisher: {
      "@type": "Person",
      "@id": personId,
      name: "Micheal Sheehy",
      url: `${site}/about.html`,
    },
    isPartOf: { "@id": websiteId },
    mainEntityOfPage: { "@type": "WebPage", "@id": canonical },
  };
}

function addChapterByline(html, published) {
  if (/class=["']article-seo-meta["']/i.test(html)) return html;
  const formatted = new Intl.DateTimeFormat("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
    timeZone: "UTC",
  }).format(new Date(`${published}T00:00:00Z`));
  return html.replace(
    /(<section class="article-hero">[\s\S]*?<h1>[\s\S]*?<\/h1><p>[\s\S]*?<\/p>)(<\/div><\/section>)/i,
    `$1<p class="article-seo-meta">By <a href="../about.html" rel="author">Micheal Sheehy</a> · Published ${formatted} · Updated July 28, 2026</p>$2`,
  );
}

for (const [relative, data] of Object.entries(articles)) {
  let html = read(relative);
  const canonical = pageUrl(relative);
  html = setTitle(html, data.title);
  html = setDescription(html, data.description);
  html = setCanonical(html, canonical);
  html = ensureIcon(html, "../");
  html = setSocial(html, data, canonical);
  html = setSchema(html, articleSchema(data, canonical, html));
  if (/articles\/chapter-\d+\.html$/.test(relative)) {
    html = addChapterByline(html, data.published);
    html = html
      .replace('class="article-layout"', 'class="article-layout chapter-layout"')
      .replace('href="../index.html#about"', 'href="../about.html"')
      .replace('href="../index.html#topics"', 'href="../insights.html"')
      .replaceAll('href="../index.html#series"', 'href="../adaptive-compliance.html"');
  }
  write(relative, html);
}

for (const [relative, data] of Object.entries(caseStudies)) {
  let html = read(relative);
  const canonical = pageUrl(relative);
  const title = textFromTag(html, "title");
  html = setDescription(html, data.description);
  html = setCanonical(html, canonical);
  html = ensureIcon(html, "../");
  html = setSocial(html, { ...data, title }, canonical);
  html = setSchema(html, {
    "@context": "https://schema.org",
    "@type": "Article",
    "@id": `${canonical}#case-study`,
    headline: data.headline,
    description: data.description,
    image: `${site}/micheal-sheehy-headshot.jpg`,
    datePublished: "2026-07-24",
    dateModified: updated,
    author: { "@type": "Person", "@id": personId, name: "Micheal Sheehy", url: `${site}/about.html` },
    publisher: { "@type": "Person", "@id": personId, name: "Micheal Sheehy", url: `${site}/about.html` },
    isPartOf: { "@id": websiteId },
    mainEntityOfPage: { "@type": "WebPage", "@id": canonical },
  });
  write(relative, html);
}

for (const relative of canonicalOnly) {
  let html = read(relative);
  html = setCanonical(html, pageUrl(relative));
  write(relative, html);
}

const redirect = `<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<meta name="robots" content="noindex,follow"/>
<title>Compliance as a Competitive Advantage | Micheal Sheehy</title>
<link rel="canonical" href="${site}/articles/compliance-advantage.html"/>
<meta http-equiv="refresh" content="0; url=articles/compliance-advantage.html"/>
<script>location.replace("articles/compliance-advantage.html");</script>
<link rel="stylesheet" href="styles.css"/>
</head><body><main class="section"><div class="container"><h1>Compliance as a Competitive Advantage</h1><p>This page has moved. <a href="articles/compliance-advantage.html">Continue to the article →</a></p></div></main></body></html>
`;
write("compliance-advantage.html", redirect);

for (const relative of ["404.html", "insights-updated.html"]) {
  let html = read(relative);
  if (!/<meta\s+name=["']robots["']/i.test(html)) {
    html = html.replace(/<head>/i, `<head>\n<meta name="robots" content="noindex,nofollow"/>`);
  }
  write(relative, html);
}

const sitemapFiles = [
  "index.html",
  "about.html",
  "insights.html",
  "adaptive-compliance.html",
  "adaptive-compliance-manuscript.html",
  "case-studies.html",
  "speaking.html",
  "media.html",
  "now.html",
  "contact.html",
  ...Object.keys(articles),
  ...Object.keys(caseStudies),
];

const changedFiles = new Set([
  "index.html",
  "about.html",
  "adaptive-compliance.html",
  "adaptive-compliance-manuscript.html",
  "case-studies.html",
  ...Object.keys(articles),
  ...Object.keys(caseStudies),
]);

function lastModified(relative) {
  if (changedFiles.has(relative)) return updated;
  try {
    return execFileSync("git", ["log", "-1", "--format=%as", "--", relative], {
      cwd: root,
      encoding: "utf8",
    }).trim() || updated;
  } catch {
    return updated;
  }
}

const sitemapEntries = sitemapFiles.map((relative) => {
  const loc = relative === "index.html" ? `${site}/` : pageUrl(relative);
  return `  <url><loc>${loc}</loc><lastmod>${lastModified(relative)}</lastmod></url>`;
}).join("\n");

write("sitemap.xml", `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${sitemapEntries}
</urlset>
`);

console.log(`Updated ${Object.keys(articles).length} articles, ${Object.keys(caseStudies).length} case studies, canonical hub pages, redirects and sitemap.`);
