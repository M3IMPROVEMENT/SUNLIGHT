"""
Kompass Scraper - desktop app (Windows / Mac / Linux)

Install once:
    pip install playwright
    (Google Chrome must be installed)

Run:
    python kompass_app.py
"""
import csv, os, re, json, random, time, threading, queue, traceback
from urllib.parse import urljoin, urlparse, urlencode, parse_qsl, urlunparse, quote_plus
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

from playwright.sync_api import sync_playwright

FIELDS = ["name", "url", "vat", "enterprise_no", "phone", "gsm", "email", "website",
          "address", "postal_code", "city", "legal_form", "activity", "employees",
          "founded", "turnover", "other_info", "done"]

COMPANY_RE = re.compile(r"/c/[^/]+/[^/]+/?$")
PHONE_RE = re.compile(r"(?:\+|00)?\d[\d\s./()-]{6,16}\d")
LABEL_PHONE_RE = re.compile(r"(t[ée]l[ée]phone|\bt[ée]l\b|\btel\b|phone|mobile|gsm|portable)", re.I)
FAX_RE = re.compile(r"\bfax\b", re.I)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
VAT_RE = re.compile(r"\bBE[\s.]?[01]?\d{3}[\s.]?\d{3}[\s.]?\d{3}\b", re.I)
ENT_NO_RE = re.compile(r"\b[01]?\d{3}[\s.]?\d{3}[\s.]?\d{3}\b")
POSTAL_CITY_RE = re.compile(r"\b(\d{4})\s+([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'’ .-]{1,40})")

ENT_SCAN_RE = re.compile(r"(?<!\d)[01]\d{3}[\s.]?\d{3}[\s.]?\d{3}(?!\d)")
L_WEB = re.compile(r"(site web|site internet|website)", re.I)
L_VAT = re.compile(r"\b(tva|vat|btw)\b", re.I)
L_ENT = re.compile(r"(bce|kbo|num[ée]ro d.entreprise|n°\s*d.entreprise|company (number|no)|registration)", re.I)
L_ADDR = re.compile(r"^\s*(adresse|address|si[èe]ge)", re.I)
L_LEGAL = re.compile(r"(forme juridique|legal form|legal status)", re.I)
L_ACT = re.compile(r"(activit[ée]|activity|secteur|sector|nace)", re.I)
L_EMP = re.compile(r"(effectif|employ|nombre de salari)", re.I)
L_FOUND = re.compile(r"(ann[ée]e de cr[ée]ation|fond[ée]|founded|cr[ée]ation|established)", re.I)
L_TURN = re.compile(r"(chiffre d.affaires|turnover|revenue)", re.I)

REVEAL_RE = re.compile(
    r"(afficher|voir|montrer|show|display|reveal).{0,25}(t[ée]l|phone|num[ée]ro|number|contact|e-?mail)"
    r"|^\s*(t[ée]l[ée]phone|phone)\s*$", re.I)

SKIP_SITES = ("kompass.com", "b2bstars", "svaplus", "facebook.", "linkedin.", "twitter.",
              "instagram.", "youtube.", "google.", "cookiebot", "onetrust", "doubleclick",
              "gstatic", "w3.org", "schema.org", "apple.com", "play.google", "ksales", "ksales.ai")

COOKIE_SELECTORS = [
    "#onetrust-accept-btn-handler",
    "#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll",
    "button:has-text('Tout accepter')", "button:has-text('Accepter')",
    "button:has-text('Accept all')", "button:has-text('Accept')",
]
SEARCH_INPUTS = [
    "input[type='search']", "input[name*='search' i]", "input[id*='search' i]",
    "input[placeholder*='cherch' i]", "input[placeholder*='search' i]",
    "input[name='text']", "input[type='text']",
]
NEXT_SELECTORS = [
    "a[rel='next']", "li.next a", "a.next", "a.pagination-next", "button.next",
    "a[aria-label*='suivant' i]", "button[aria-label*='suivant' i]",
    "a[aria-label*='next' i]", "button[aria-label*='next' i]",
    "a:has-text('Suivant')", "button:has-text('Suivant')",
    "a:has-text('Next')", "button:has-text('Next')",
    "a:has-text('›')", "a:has-text('»')",
]


class Stopped(Exception):
    pass


class BrowserClosed(Stopped):
    pass


# ---------------------------------------------------------------- helpers

def normalize_be_phone(raw):
    """Valid Belgian number -> +32XXXXXXXX, else ''. Kompass often hides the leading 0."""
    digits = re.sub(r"\D", "", raw or "")
    if digits.startswith("0032"):
        national = digits[4:]
    elif digits.startswith("32") and len(digits) in (10, 11):
        national = digits[2:]
    elif digits.startswith("0"):
        national = digits[1:]
    else:
        national = digits
    if national.startswith("4") and len(national) == 9:          # mobile (GSM)
        return "+32" + national
    if national[:1] in list("123456789") and not national.startswith("4") and len(national) == 8:
        return "+32" + national                                    # landline
    return ""


def format_vat(raw):
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 9:
        digits = "0" + digits
    return "BE" + digits if len(digits) == 10 else ""


def format_ent_no(raw):
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 9:
        digits = "0" + digits
    if len(digits) == 10:
        return f"{digits[:4]}.{digits[4:7]}.{digits[7:]}"
    return ""


def valid_ent_no(raw):
    """Belgian enterprise number checksum: last 2 digits = 97 - (first 8 digits mod 97)."""
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 9:
        digits = "0" + digits
    return len(digits) == 10 and digits[0] in "01" and (97 - int(digits[:8]) % 97) == int(digits[8:])


def url_with_page(url, n, param="page"):
    u = urlparse(url)
    q = dict(parse_qsl(u.query))
    q[param] = str(n)
    return urlunparse(u._replace(query=urlencode(q)))


def slugify(text):
    return re.sub(r"[^\w]+", "_", (text or "").lower()).strip("_") or "results"


# ---------------------------------------------------------------- scraper

class Scraper:
    def __init__(self, cfg, log, stop):
        self.cfg, self.log, self.stop = cfg, log, stop
        self.base = cfg["base"].rstrip("/")
        self.out = cfg["out"]
        self.rows = {}
        self.slow = 1.0
        self._warned = False

    # ---- control
    def check(self):
        if self.stop.is_set():
            raise Stopped()

    def sleep(self, seconds):
        end = time.time() + seconds
        while time.time() < end:
            self.check()
            time.sleep(0.3)

    def pause(self, a=3, b=7):
        self.sleep(random.uniform(a, b) * self.slow)

    # ---- storage
    def load_existing(self):
        for path in (self.out, os.path.splitext(self.out)[0] + "_live.csv"):
            if os.path.exists(path):
                with open(path, encoding="utf-8-sig", newline="") as f:
                    for r in csv.DictReader(f):
                        if r.get("url"):
                            self.rows[r["url"]] = {k: r.get(k, "") for k in FIELDS}
            self.log(f"Resuming: {len(self.rows)} companies already in {os.path.basename(self.out)}")

    def save(self):
        live = os.path.splitext(self.out)[0] + "_live.csv"
        for path in (self.out, live):
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.DictWriter(f, fieldnames=FIELDS)
                    w.writeheader()
                    w.writerows(self.rows.values())
                if path != self.out and not self._warned:
                    self._warned = True
                    self.log(f"  ! The CSV is open in Excel - saving to {os.path.basename(path)} instead. "
                             "Close Excel to go back to the normal file.")
                return
            except PermissionError:
                continue
        self.log("  ! Cannot write any CSV - close it in Excel.")

    def save_xlsx(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.append(FIELDS)
            for r in self.rows.values():
                ws.append([r.get(k, "") for k in FIELDS])
            wb.save(os.path.splitext(self.out)[0] + ".xlsx")
        except Exception:
            pass

    # ---- captcha / cookies / navigation
    def blocked(self, page):
        try:
            if "captcha-delivery" in page.url:
                return True
            return page.query_selector("iframe[src*='captcha-delivery']") is not None
        except Exception:
            return False

    def wait_if_blocked(self, page):
        if not self.blocked(page):
            return
        self.log("Verification check on screen - solve it in the Chrome window. Waiting...")
        try:
            import winsound
            winsound.Beep(1000, 400)
            winsound.Beep(1400, 400)
        except Exception:
            pass
        while self.blocked(page):
            self.check()
            time.sleep(2)
        page.wait_for_timeout(2000)
        self.slow = min(self.slow * 1.5, 4.0)
        self.log(f"Check cleared. Slowing down (x{self.slow:.1f}) to avoid more checks...")
        self.sleep(random.uniform(20, 40))

    def accept_cookies(self, page):
        for sel in COOKIE_SELECTORS:
            try:
                el = page.query_selector(sel)
                if el and el.is_visible():
                    el.click(timeout=2000)
                    page.wait_for_timeout(1000)
                    return
            except Exception:
                continue

    def goto(self, page, url):
        self.check()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
        except Exception as e:
            if "has been closed" in str(e):
                raise BrowserClosed()
            self.log("  (navigation note) " + str(e).splitlines()[0][:120])
            try:
                page.wait_for_timeout(2500)
            except Exception:
                raise BrowserClosed()
        self.wait_if_blocked(page)
        self.accept_cookies(page)

    # ---- search + results pages
    def scroll_page(self, page):
        try:
            for _ in range(4):
                page.mouse.wheel(0, 1800)
                page.wait_for_timeout(500)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(800)
        except Exception as e:
            if "has been closed" in str(e):
                raise BrowserClosed()

    def links(self, page, tries=4):
        for _ in range(tries):
            self.check()
            self.wait_if_blocked(page)
            try:
                page.wait_for_selector("a[href*='/c/']", timeout=20000)
            except Exception:
                pass
            page.wait_for_timeout(1200)
            self.scroll_page(page)
            try:
                found = {}
                anchors = page.eval_on_selector_all(
                    "a[href*='/c/']",
                    "els => els.map(e => [e.getAttribute('href'), e.innerText])")
                for href, text in anchors:
                    full = urljoin(page.url, (href or "").split("?")[0])
                    if COMPANY_RE.search(full):
                        name = (text or "").strip()
                        if full not in found or (name and not found[full]):
                            found[full] = name
                return found
            except Exception:
                self.sleep(2)
        return {}

    def search(self, page, sector):
        q = sector.strip()
        self.log(f"Searching for: {q}")
        self.goto(page, self.base + "/fr")
        for sel in SEARCH_INPUTS:
            try:
                for el in page.query_selector_all(sel):
                    if el.is_visible():
                        el.click()
                        el.fill("")
                        el.type(q, delay=random.randint(60, 140))
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(3000)
                        self.wait_if_blocked(page)
                        if self.links(page, tries=2):
                            return True
                        break
            except Exception:
                continue
        for url in (f"{self.base}/fr/searchCompanies?text={quote_plus(q)}",
                    f"{self.base}/searchCompanies?text={quote_plus(q)}"):
            self.goto(page, url)
            if self.links(page, tries=2):
                return True
        return False

    @staticmethod
    def _disabled(el):
        try:
            cls = (el.get_attribute("class") or "").lower()
            pcls = (el.evaluate("e => e.parentElement ? e.parentElement.className : ''") or "").lower()
            aria = (el.get_attribute("aria-disabled") or "").lower()
            return "disabled" in cls or "disabled" in pcls or aria == "true" or el.is_disabled()
        except Exception:
            return False

    BAR_JS = r"""() => {
        const isNum = e => /^\d+$/.test((e.innerText || '').trim());
        const leaves = Array.from(document.querySelectorAll('a,button,li,span,div'))
            .filter(e => e.children.length === 0 && isNum(e));
        for (const l of leaves) {
            let p = l.parentElement;
            for (let i = 0; i < 4 && p; i++, p = p.parentElement) {
                if ((p.innerText || '').length > 200) continue;
                const nums = Array.from(p.querySelectorAll('a,button,li,span,div'))
                    .filter(e => e.children.length === 0 && isNum(e));
                if (nums.length >= 4) {
                    const items = Array.from(p.querySelectorAll('a,button,[role=button]'))
                        .filter(e => e.offsetParent !== null);
                    return {text: (p.innerText || '').replace(/\s+/g, ' ').trim(),
                            last: items.length ? items[items.length - 1] : null};
                }
            }
        }
        return null;
    }"""

    def bar_arrow(self, page):
        """Pagination bar like '< 1 ... 9 10 11 12 15 20 ... 40 >': return the last item (the arrow)."""
        try:
            h = page.evaluate_handle(self.BAR_JS)
            txt = h.get_property("text").json_value()
            last = h.get_property("last").as_element()
            self.log(f"  pagination bar: {str(txt)[:90]}")
            if last and not re.fullmatch(r"\d+", (last.inner_text() or "").strip()):
                return last
        except Exception as e:
            if "has been closed" in str(e):
                raise BrowserClosed()
        return None

        
    def find_next(self, page, current_no):
        """Find the next page reliably, including pagination with ellipses."""

        selectors = [
            "a[rel='next']",
            "button[rel='next']",
            "a[aria-label*='next' i]",
            "button[aria-label*='next' i]",
            "a[aria-label*='suivant' i]",
            "button[aria-label*='suivant' i]",
            "li.next a",
            "a.next",
            "button.next",
            "a.pagination-next",
            "button.pagination-next",
    ]

        for selector in selectors:
            try:
                elements = page.locator(selector)
                for i in range(min(elements.count(), 10)):
                    el = elements.nth(i)
                    if el.is_visible() and not self._disabled(el):
                         return el
            except Exception:
                continue

    # Search for a visible pagination control whose text is
    # a next arrow, rather than assuming the last item is the arrow.
        selectors = [
            "nav a", "nav button",
            ".pagination a", ".pagination button",
            "[class*='paginat'] a",
            "[class*='paginat'] button",
    ]

        for selector in selectors:
            try:
                elements = page.locator(selector)
                for i in range(min(elements.count(), 100)):
                    el = elements.nth(i)
                    if not el.is_visible() or self._disabled(el):
                        continue

                    text = (el.inner_text() or "").strip()
                    aria = (el.get_attribute("aria-label") or "").lower()
                    title = (el.get_attribute("title") or "").lower()

                    if (
                        text in ("›", "»", ">", "→")
                        or "next" in aria
                        or "suivant" in aria
                        or "next" in title
                        or "suivant" in title
                    ):
                        return el
            except Exception:
                continue

    # Last resort: locate the exact next page number.
        target = str(current_no + 1)

        for selector in [
            "nav a", "nav button",
            ".pagination a", ".pagination button",
            "[class*='paginat'] a",
            "[class*='paginat'] button",
        ]:
            try:
                elements = page.locator(selector)
                for i in range(min(elements.count(), 100)):
                    el = elements.nth(i)
                    if (
                        el.is_visible()
                        and (el.inner_text() or "").strip() == target
                        and not self._disabled(el)
                    ):
                        return el
            except Exception:
                continue

        return None  
 

     
    def next_page(self, page, current_no, current_links, results_url):
        """Advance only when the next page produces different results."""

        old_links = set(current_links)
        old_url = page.url

        def wait_for_new_results(timeout_seconds=15):
            deadline = time.time() + timeout_seconds

            while time.time() < deadline
                self.check()
                self.wait_if_blocked(page)

                try:
                    current = set(self.links(page, tries=1).keys())

                    if current and current != old_links:
                        self.log(
                            f"  Page navigation confirmed: "
                            f"{old_url} -> {page.url}"
                        )
                        return True

                except Stopped:
                    raise
                except Exception:
                    pass

                self.sleep(1)

            return False

    # First try the actual next-page button.
        btn = self.find_next(page, current_no)

        if btn:
            try:
                self.log(f"  Clicking next after page {current_no}...")

                btn.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
                btn.click(timeout=10000)

                if wait_for_new_results():
                    return True

                self.log("  Click did not change results; trying URL pagination.")

            except Stopped:
                raise
            except Exception as e:
                self.log(
                    "  Next-button error: "
                    + str(e).splitlines()[0][:120]
                )

    # Fallback: test common page-number query parameters.
        for param in ("page", "pageNumber", "pageNo", "p"):
            self.check()

            target_url = url_with_page(
                results_url, current_no + 1, param
            )

            try:
                self.log(f"  Trying fallback parameter: {param}")

                self.goto(page, target_url)

                if wait_for_new_results(timeout_seconds=8):
                    return True

            except Stopped:
                raise
            except Exception as e:
                self.log(
                    f"  Fallback {param} failed: "
                    + str(e).splitlines()[0][:100]
                )

        self.log(
            f"  Could not confirm page {current_no + 1}. "
            "Check the saved pagination debug files."
        )
        return False      

  
    def dump_pagination(self, page, page_no):
        try:
            html = page.evaluate("""() => {
                const el = document.querySelector("nav[aria-label*='pagination' i], ul.pagination, .pagination, [class*='paginat']");
                return el ? el.outerHTML.slice(0, 5000) : "(no pagination element found)";
            }""")
            with open("debug_pagination.html", "w", encoding="utf-8") as f:
                f.write(f"<!-- stopped after page {page_no} at {page.url} -->\n" + html)
            page.screenshot(path="debug_last_page.png")
            self.log(f"  Stopped after page {page_no}. URL: {page.url}")
            self.log("  Saved debug_pagination.html and debug_last_page.png (send me these if you expected more pages).")
        except Exception:
            pass

    def paginate(self, page):
        seen, page_no = set(), 1
        results_url = page.url
        while page_no <= self.cfg["max_pages"]:
            found = self.links(page)
            if page_no > 1 and len(found) < 5:
                for attempt in range(3):
                    self.log(f"  Page {page_no} shows only {len(found)} companies - waiting and retrying ({attempt + 1}/3)...")
                    self.sleep(8)
                    found = self.links(page)
                    if len(found) >= 5:
                        break
            new_this_run = set(found) - seen
            seen |= set(found)
            added = 0
            for u, n in found.items():
                if u not in self.rows:
                    self.rows[u] = {k: "" for k in FIELDS}
                    self.rows[u].update(name=n, url=u)
                    added += 1
            self.log(f"[Page {page_no}] {len(found)} companies, {added} new, total {len(self.rows)}")
            self.save()
            if not found:
                with open("debug_page.html", "w", encoding="utf-8") as f:
                    f.write(page.content())
                self.log("  No company links here. Saved debug_page.html. URL: " + page.url)
                break
            if page_no > 1 and len(found) < 5:
                self.log(f"  Page {page_no} shows only {len(found)} companies - the site may limit "
                         "results (login / result cap), or a check blocked the page.")
                self.dump_pagination(page, page_no)
                break
            
            if page_no >= self.cfg["max_pages"]:
                break
            if not self.next_page(page, page_no, found.keys(), results_url):
                self.log("  No further page found - end of results.")
                self.dump_pagination(page, page_no)
                break
            page_no += 1
            self.pause(3, 6)

    # ---- company page
    def click_reveal_buttons(self, page):
        clicked = 0
        for el in page.query_selector_all("button, a, span, div[role='button']"):
            if clicked >= 4:
                break
            try:
                txt = (el.inner_text() or "").strip()
                if txt and len(txt) < 40 and REVEAL_RE.search(txt) and el.is_visible():
                    el.click(timeout=2000)
                    page.wait_for_timeout(1200)
                    clicked += 1
            except Exception:
                continue

    @staticmethod
    def jsonld(page):
        out = {}
        try:
            blocks = page.eval_on_selector_all(
                "script[type='application/ld+json']", "els => els.map(e => e.textContent)")
        except Exception:
            return out

        def walk(o):
            if isinstance(o, list):
                for x in o:
                    walk(x)
            elif isinstance(o, dict):
                if o.get("telephone"):
                    out.setdefault("phone", str(o["telephone"]))
                if o.get("email"):
                    out.setdefault("email", str(o["email"]))
                if o.get("vatID"):
                    out.setdefault("vat", str(o["vatID"]))
                u = str(o.get("url") or "")
                if u.startswith("http") and not any(s in u for s in SKIP_SITES):
                    out.setdefault("website", u)
                a = o.get("address")
                if isinstance(a, dict):
                    out.setdefault("address", ", ".join(str(a.get(k, "")) for k in
                                   ("streetAddress", "postalCode", "addressLocality") if a.get(k)))
                    if a.get("postalCode"):
                        out.setdefault("postal_code", str(a["postalCode"]))
                    if a.get("addressLocality"):
                        out.setdefault("city", str(a["addressLocality"]))
                for v in o.values():
                    walk(v)

        for b in blocks:
            try:
                walk(json.loads(b))
            except Exception:
                pass
        return out

    @staticmethod
    def key_values(page):
        pairs = {}
        try:
            data = page.evaluate("""() => {
                const out = [];
                document.querySelectorAll('dt').forEach(dt => {
                    const dd = dt.nextElementSibling;
                    if (dd) out.push([dt.innerText, dd.innerText]);
                });
                document.querySelectorAll('tr').forEach(tr => {
                    const c = tr.querySelectorAll('th,td');
                    if (c.length === 2) out.push([c[0].innerText, c[1].innerText]);
                });
                return out;
            }""")
        except Exception:
            data = []
        for k, v in data:
            k = re.sub(r"\s+", " ", (k or "")).strip(" :")
            v = re.sub(r"\s+", " ", (v or "")).strip()
            if k and v and len(k) < 60 and len(v) < 400:
                pairs.setdefault(k, v)
        return pairs

    @staticmethod
    def pick(pairs, lines, label_re, used):
        for k, v in pairs.items():
            if label_re.search(k):
                used.add(k)
                return v
        for i, line in enumerate(lines):
            if len(line) < 80 and label_re.search(line):
                if ":" in line:
                    v = line.split(":", 1)[1].strip()
                    if v:
                        return v[:200]
                if i + 1 < len(lines):
                    return lines[i + 1][:200]
        return ""

    def parse_company(self, page, first=False):
        d = {k: "" for k in FIELDS if k not in ("url", "done")}
        try:
            body = page.inner_text("body")
        except Exception:
            body = ""
        lines = [re.sub(r"\s+", " ", l).strip() for l in body.splitlines() if l.strip()]
        if first:
            with open("debug_company.txt", "w", encoding="utf-8") as f:
                f.write(body)
            with open("debug_company.html", "w", encoding="utf-8") as f:
                f.write(page.content())

        ld = self.jsonld(page)
        pairs = self.key_values(page)
        used = set()

        h1 = page.query_selector("h1")
        if h1:
            d["name"] = " ".join((h1.inner_text() or "").split())

        # VAT + enterprise number
        vat_raw = ld.get("vat", "") or self.pick(pairs, lines, L_VAT, used)
        d["vat"] = format_vat(vat_raw)
        if not d["vat"]:
            m = VAT_RE.search(body)
            if m:
                d["vat"] = format_vat(m.group(0))
        ent_raw = self.pick(pairs, lines, L_ENT, used)
        m = ENT_NO_RE.search(ent_raw or "")
        d["enterprise_no"] = format_ent_no(m.group(0)) if m else ""
        if d["enterprise_no"] and not valid_ent_no(d["enterprise_no"]):
            d["enterprise_no"] = ""
        if not d["enterprise_no"] and d["vat"]:
            d["enterprise_no"] = format_ent_no(d["vat"][2:])
        if not d["enterprise_no"]:
            for m in ENT_SCAN_RE.finditer(body):      # scan the page for a checksum-valid number
                if valid_ent_no(m.group(0)):
                    d["enterprise_no"] = format_ent_no(m.group(0))
                    break
        if not d["vat"] and d["enterprise_no"]:       # in Belgium VAT = "BE" + enterprise number
            d["vat"] = "BE" + re.sub(r"\D", "", d["enterprise_no"])

        # phones: JSON-LD, tel: links, text next to Tel/GSM/Mobile labels
        cands = []
        if ld.get("phone"):
            cands.append(ld["phone"])
        for a in page.query_selector_all("a[href^='tel:']"):
            cands.append((a.get_attribute("href") or "").replace("tel:", ""))
        for i, line in enumerate(lines):
            if len(line) < 80 and LABEL_PHONE_RE.search(line) and not FAX_RE.search(line):
                chunk = " ".join(lines[i:i + 3])
                cands += [m.group(0) for m in PHONE_RE.finditer(chunk)]
        for c in cands:
            n = normalize_be_phone(c)
            if not n:
                continue
            if n.startswith("+324"):
                d["gsm"] = d["gsm"] or n
            else:
                d["phone"] = d["phone"] or n

        # email
        d["email"] = ld.get("email", "")
        if not d["email"]:
            mail = page.query_selector("a[href^='mailto:']")
            if mail:
                d["email"] = (mail.get_attribute("href") or "").replace("mailto:", "").split("?")[0].strip()
        if not d["email"]:
            for m in EMAIL_RE.finditer(body):
                if "kompass" not in m.group(0).lower():
                    d["email"] = m.group(0)
                    break

        # website: JSON-LD, a "Site web" label, or a link whose text looks like a domain
        site = ld.get("website", "")
        if not site:
            wv = self.pick(pairs, lines, L_WEB, used)
            m = re.search(r"(https?://\S+|www\.\S+|[\w-]+\.[a-z]{2,}(?:/\S*)?)", wv or "", re.I)
            if m:
                site = m.group(0)
        if not site:
            for a in page.query_selector_all("a[href^='http']"):
                h = a.get_attribute("href") or ""
                txt = (a.inner_text() or "").strip().lower()
                if any(x in h for x in SKIP_SITES):
                    continue
                if txt.startswith("www.") or re.fullmatch(r"(https?://)?[\w-]+(\.[\w-]+)+/?", txt):
                    site = h
                    break
        if any(x in site for x in SKIP_SITES):
            site = ""
        d["website"] = site

        # address, postal code, city
        addr = ld.get("address", "")
        if not addr:
            el = page.query_selector("[itemprop='address'], address")
            if el:
                addr = " ".join(el.inner_text().split())
        if not addr:
            addr = self.pick(pairs, lines, L_ADDR, used)
        d["address"] = addr
        d["postal_code"] = ld.get("postal_code", "")
        d["city"] = ld.get("city", "")
        if addr and not d["postal_code"]:
            m = POSTAL_CITY_RE.search(addr)
            if m:
                d["postal_code"], d["city"] = m.group(1), m.group(2).strip(" ,.-")

        # other info
        d["legal_form"] = self.pick(pairs, lines, L_LEGAL, used)
        d["activity"] = self.pick(pairs, lines, L_ACT, used)
        d["employees"] = self.pick(pairs, lines, L_EMP, used)
        d["founded"] = self.pick(pairs, lines, L_FOUND, used)
        d["turnover"] = self.pick(pairs, lines, L_TURN, used)
        rest = {k: v for k, v in pairs.items() if k not in used}
        d["other_info"] = json.dumps(dict(list(rest.items())[:40]), ensure_ascii=False) if rest else ""
        return d

    def open_company(self, page, url, first):
        self.goto(page, url)
        self.sleep(random.uniform(2, 3.5))
        self.click_reveal_buttons(page)
        return self.parse_company(page, first)

    # ---- main
    def run(self):
        try:
            self._run()
        except BrowserClosed:
            self.log("The Chrome window was closed, so the run stopped. Keep that window open "
                     "(minimize it instead). Press Start again to resume where it stopped.")
        except Stopped:
            self.log("Stopped by user.")
        except Exception:
            self.log("ERROR:\n" + traceback.format_exc())
        finally:
            self.save()
            self.save_xlsx()
            self.log(f"Saved {len(self.rows)} rows -> {self.out}")

    def _run(self):
        self.load_existing()
        with sync_playwright() as p:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=os.path.abspath("chrome_profile"), channel="chrome",
                headless=False, locale="fr-BE", viewport={"width": 1366, "height": 850},
                args=["--disable-blink-features=AutomationControlled"],
            )
            ctx.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()

                urls = [u for u in re.split(r"[\s;]+", self.cfg["start_url"]) if u.startswith("http")]
                if urls:
                    for n, u in enumerate(urls, 1):
                        self.check()
                        self.log(f"=== Results URL {n}/{len(urls)} ===")
                        self.goto(page, u)
                        self.paginate(page)
                else:
                    keywords = [k.strip() for k in re.split(r"[;,\n]", self.cfg["sector"]) if k.strip()]
                    for kw in keywords:
                        self.check()
                        if not self.search(page, kw):
                            self.log(f"No results found for '{kw}'.")
                            continue
                        self.log(f"=== Keyword: {kw} ===")
                        self.paginate(page)
                    if not self.rows:
                        self.log("Could not get results automatically.")
                        self.log("-> Search by hand on the site, copy the results URL into the 'Results URL' box, and start again.")
                        return

                if self.cfg["details"]:
                    todo = [u for u, r in self.rows.items() if r.get("done") != "1"]
                    if self.cfg["max_companies"]:
                        todo = todo[:self.cfg["max_companies"]]
                    self.log(f"\nOpening {len(todo)} company pages...")
                    dpage = ctx.new_page()
                    for i, url in enumerate(todo, 1):
                        self.check()
                        try:
                            d = self.open_company(dpage, url, first=(i == 1))
                            if not d["name"]:
                                d["name"] = self.rows[url]["name"]
                            self.rows[url].update(d)
                            self.rows[url]["done"] = "1"
                            self.log(f"[{i}/{len(todo)}] {d['name'][:35]} | VAT {d['vat'] or '-'} | "
                                     f"{d['gsm'] or d['phone'] or '-'} | {d['postal_code'] or '-'} {d['city']}")
                        except Stopped:
                            raise
                        except Exception as e:
                            if "has been closed" in str(e):
                                raise BrowserClosed()
                            self.log("  error: " + str(e).splitlines()[0][:120])
                        if i % 3 == 0:
                            self.save()
                        self.pause(3, 7)
                        if i % 25 == 0:
                            self.log("  short break...")
                            self.sleep(random.uniform(30, 60))
            finally:
                try:
                    ctx.close()
                except Exception:
                    pass


# ---------------------------------------------------------------- GUI

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kompass Scraper")
        self.geometry("860x660")
        self.q = queue.Queue()
        self.stop = threading.Event()
        self.thread = None
        self.last_out = ""

        f = ttk.Frame(self, padding=10)
        f.pack(fill="x")
        f.columnconfigure(1, weight=1)

        def row(r, label, widget):
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", pady=3)
            widget.grid(row=r, column=1, sticky="ew", pady=3, padx=6)

        self.sector = ttk.Entry(f)
        row(0, "Sector / keyword(s), comma-separated", self.sector)
        self.url = ttk.Entry(f)
        row(1, "Results URL(s) - optional, space-separated", self.url)
        self.site = ttk.Entry(f)
        self.site.insert(0, "https://be.kompass.com")
        row(2, "Site", self.site)
        self.pages = ttk.Spinbox(f, from_=1, to=1000, width=8)
        self.pages.set(50)
        row(3, "Max result pages", self.pages)
        self.maxc = ttk.Spinbox(f, from_=0, to=100000, width=8)
        self.maxc.set(0)
        row(4, "Max companies (0 = all)", self.maxc)
        self.details = tk.BooleanVar(value=True)
        row(5, "", ttk.Checkbutton(f, text="Open every company page (VAT, GSM, email, address...)",
                                   variable=self.details))
        self.outv = ttk.Entry(f)
        row(6, "Output CSV (blank = auto)", self.outv)

        b = ttk.Frame(self, padding=(10, 0))
        b.pack(fill="x")
        self.btn_start = ttk.Button(b, text="Start", command=self.start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(b, text="Stop", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        ttk.Button(b, text="Open CSV", command=self.open_csv).pack(side="left")

        self.logbox = scrolledtext.ScrolledText(self, height=24, font=("Consolas", 9))
        self.logbox.pack(fill="both", expand=True, padx=10, pady=10)
        self.log("Type a sector (e.g. 'construction') and press Start.\n"
                 "A Chrome window will open - leave it alone; if a check appears, solve it and the app continues.\n")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(200, self.poll)

    def log(self, msg):
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")

    def start(self):
        sector, url = self.sector.get().strip(), self.url.get().strip()
        if not sector and not url:
            messagebox.showwarning("Kompass Scraper", "Enter a sector or paste a results URL.")
            return
        try:
            pages, maxc = int(self.pages.get()), int(self.maxc.get())
        except ValueError:
            messagebox.showwarning("Kompass Scraper", "Pages / companies must be numbers.")
            return
        out = self.outv.get().strip() or f"kompass_{slugify(sector or 'results')}.csv"
        self.last_out = os.path.abspath(out)
        cfg = dict(sector=sector, start_url=url, base=self.site.get().strip() or "https://be.kompass.com",
                   max_pages=pages, max_companies=maxc, details=self.details.get(), out=self.last_out)
        self.stop.clear()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.thread = threading.Thread(target=self.work, args=(cfg,), daemon=True)
        self.thread.start()

    def work(self, cfg):
        try:
            Scraper(cfg, self.q.put, self.stop).run()
        finally:
            self.q.put("__DONE__")

    def on_stop(self):
        self.stop.set()
        self.log("Stopping after the current step...")

    def open_csv(self):
        if self.last_out and os.path.exists(self.last_out):
            try:
                os.startfile(self.last_out)
            except AttributeError:
                messagebox.showinfo("Kompass Scraper", self.last_out)
        else:
            messagebox.showinfo("Kompass Scraper", "No file yet.")

    def poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m == "__DONE__":
                    self.btn_start.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    self.log("--- finished ---")
                else:
                    self.log(m)
        except queue.Empty:
            pass
        self.after(200, self.poll)

    def on_close(self):
        self.stop.set()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()