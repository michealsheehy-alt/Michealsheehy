import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const site = "https://www.michealsheehy.com";
const sitemap = fs.readFileSync(path.join(root, "sitemap.xml"), "utf8");
const urls = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map((match) => match[1]);
const failures = [];
const seenCanonicals = new Map();

function fileFor(url) {
  const pathname = new URL(url).pathname.replace(/^\/+/, "");
  return pathname ? pathname : "index.html";
}

for (const url of urls) {
  const relative = fileFor(url);
  const absolute = path.join(root, relative);
  if (!fs.existsSync(absolute)) {
    failures.push(`${relative}: sitemap target is missing`);
    continue;
  }

  const html = fs.readFileSync(absolute, "utf8");
  const title = html.match(/<title>([\s\S]*?)<\/title>/i)?.[1]?.trim();
  const description = html.match(/<meta\s+name=["']description["'][^>]*content=["']([^"']+)/i)?.[1];
  const canonical = html.match(/<link\s+rel=["']canonical["'][^>]*href=["']([^"']+)/i)?.[1];

  if (!title) failures.push(`${relative}: missing title`);
  if (!description) failures.push(`${relative}: missing meta description`);
  if (!canonical) failures.push(`${relative}: missing canonical`);
  if (canonical && canonical !== url) failures.push(`${relative}: canonical ${canonical} does not match sitemap URL ${url}`);

  if (canonical) {
    if (seenCanonicals.has(canonical)) {
      failures.push(`${relative}: duplicate canonical also used by ${seenCanonicals.get(canonical)}`);
    }
    seenCanonicals.set(canonical, relative);
  }

  const isArticle = relative.startsWith("articles/") || relative.startsWith("case-studies/");
  if (isArticle) {
    const blocks = [...html.matchAll(/<script\s+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi)];
    if (!blocks.length) {
      failures.push(`${relative}: missing JSON-LD`);
    } else {
      for (const block of blocks) {
        try {
          JSON.parse(block[1]);
        } catch (error) {
          failures.push(`${relative}: invalid JSON-LD (${error.message})`);
        }
      }
    }
  }
}

for (const relative of ["404.html", "insights-updated.html"]) {
  const html = fs.readFileSync(path.join(root, relative), "utf8");
  if (!/<meta\s+name=["']robots["'][^>]*noindex/i.test(html)) {
    failures.push(`${relative}: utility page is not marked noindex`);
  }
}

const htmlFiles = [];
function walk(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (entry.name === ".git") continue;
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) walk(absolute);
    if (entry.isFile() && entry.name.endsWith(".html")) htmlFiles.push(absolute);
  }
}
walk(root);

for (const absolute of htmlFiles) {
  const relative = path.relative(root, absolute);
  const html = fs.readFileSync(absolute, "utf8");
  for (const match of html.matchAll(/<a\b[^>]*href=["']([^"'#?]+)(?:[?#][^"']*)?["']/gi)) {
    const href = match[1];
    if (/^(?:https?:|mailto:|tel:|javascript:)/i.test(href)) continue;
    const target = path.resolve(path.dirname(absolute), href);
    let candidate = target;
    if (href.endsWith("/")) candidate = path.join(target, "index.html");
    if (!fs.existsSync(candidate)) failures.push(`${relative}: broken internal link ${href}`);
  }
}

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}

console.log(`SEO validation passed: ${urls.length} canonical sitemap pages, valid metadata/JSON-LD, no broken internal links.`);
