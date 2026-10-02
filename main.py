
import re
import threading
import time
import urllib.parse
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


GOOGLE_TRANSLATE = "https://translate.google.com/"
SUPPORTED = {".docx", ".txt"}

# Practical sentence splitter. It keeps punctuation attached to the sentence.
SENTENCE_RE = re.compile(
    r".+?(?:[.!?]+(?:[\"”’'»)]*)?(?=\s+|$)|$)",
    re.DOTALL
)


def split_sentences(text):
    if not text or not text.strip():
        return []

    result = []
    for match in SENTENCE_RE.finditer(text):
        chunk = match.group(0)
        if chunk and chunk.strip():
            result.append(chunk)

    return result or [text]


class GoogleWebTranslator:
    """
    Automates the actual Google Translate website in a visible browser.

    It does NOT use googletrans and does NOT call a translation API.
    The program opens translate.google.com and enters the sentence into
    the same source-text box a user would use manually.
    """

    def __init__(self, log):
        self.log = log
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def start(self):
        self.playwright = sync_playwright().start()

        # Prefer installed Google Chrome because that is closest to
        # the user's normal manual workflow. Fall back to Playwright
        # Chromium if Chrome is unavailable.
        try:
            self.context = self.playwright.chromium.launch_persistent_context(
                user_data_dir=str(Path.home() / ".google_translate_rewriter"),
                channel="chrome",
                headless=False,
                viewport={"width": 1400, "height": 900},
                slow_mo=20,
            )
            self.log("Opened Google Chrome.")
        except Exception:
            self.context = self.playwright.chromium.launch_persistent_context(
                user_data_dir=str(Path.home() / ".google_translate_rewriter"),
                headless=False,
                viewport={"width": 1400, "height": 900},
                slow_mo=20,
            )
            self.log("Google Chrome was unavailable; opened Playwright Chromium.")

        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.set_default_timeout(15000)

        self.page.goto(
            GOOGLE_TRANSLATE + "?sl=en&tl=bn&op=translate",
            wait_until="domcontentloaded",
        )
        self.accept_consent_if_present()
        self.wait_for_source_box()

    def accept_consent_if_present(self):
        # Consent wording varies by country/language.
        candidates = [
            "Accept all",
            "I agree",
            "সব গ্রহণ করুন",
            "সবকটি গ্রহণ করুন",
            "Alle akzeptieren",
        ]
        for name in candidates:
            try:
                btn = self.page.get_by_role("button", name=name, exact=True)
                if btn.count() and btn.first.is_visible():
                    btn.first.click(timeout=1500)
                    time.sleep(1)
                    return
            except Exception:
                pass

    def wait_for_source_box(self):
        locators = [
            self.page.locator('textarea[aria-label="Source text"]'),
            self.page.locator("textarea").first,
            self.page.get_by_role("textbox").first,
        ]

        for loc in locators:
            try:
                loc.wait_for(state="visible", timeout=10000)
                return loc
            except Exception:
                pass

        raise RuntimeError(
            "Could not find the Google Translate source text box. "
            "Google may have changed its website layout."
        )

    def source_box(self):
        return self.wait_for_source_box()

    def clear_source(self):
        box = self.source_box()
        box.click()
        box.fill("")
        return box

    def get_result_text(self, target_lang):
        """
        Google Translate's DOM changes occasionally, so several selectors
        are tried. The known .ryNqvb result selector is included as a
        fallback used by several browser-automation implementations.
        """
        selectors = [
            ".ryNqvb",
            'span[lang="' + target_lang + '"]',
            '[aria-live="polite"]',
            '[data-result-index]',
        ]

        for selector in selectors:
            try:
                loc = self.page.locator(selector)
                count = loc.count()
                texts = []
                for i in range(min(count, 20)):
                    try:
                        t = loc.nth(i).inner_text(timeout=1000).strip()
                        if t:
                            texts.append(t)
                    except Exception:
                        pass

                if texts:
                    # Remove duplicates while preserving order.
                    unique = []
                    for t in texts:
                        if t not in unique:
                            unique.append(t)
                    joined = "\n".join(unique).strip()
                    if joined:
                        return joined
            except Exception:
                pass

        return ""

    def translate(self, sentence, src, dest):
        # Navigate directly to the same Google Translate webpage with
        # the sentence in the URL. This is equivalent to entering text
        # into the Translate page, and is more reliable than guessing
        # private JavaScript endpoints.
        encoded = urllib.parse.quote(sentence, safe="")
        url = (
            f"{GOOGLE_TRANSLATE}?sl={src}&tl={dest}"
            f"&text={encoded}&op=translate"
        )

        self.page.goto(url, wait_until="domcontentloaded")

        # Wait until the source text box is available.
        self.wait_for_source_box()

        # Wait for the result to become non-empty.
        deadline = time.time() + 15
        last = ""

        while time.time() < deadline:
            result = self.get_result_text(dest)
            if result and result != sentence:
                return result

            # Some sentences can legitimately translate identically.
            # Once the page has stopped changing, accept the result.
            if result:
                last = result

            time.sleep(0.25)

        if last:
            return last

        raise RuntimeError(
            "Google Translate did not return a result for this sentence."
        )

    def close(self):
        try:
            if self.context:
                self.context.close()
        finally:
            if self.playwright:
                self.playwright.stop()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("document humanizer - Google Translate Web Document Rewriter")
        self.root.geometry("900x650")
        self.root.minsize(760, 540)

        self.input_path = None
        self.running = False
        self.stop_requested = False
        self.delay = tk.DoubleVar(value=1.0)
        self.status = tk.StringVar(value="Choose a .docx or .txt document.")

        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self.root, padding=18)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="document humanizer - Google Translate Web Document Rewriter",
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            main,
            text=(
                "Uses the actual translate.google.com website in a visible browser.\n"
                "Every sentence: English → Bengali → English."
            ),
        ).pack(anchor="w", pady=(5, 14))

        filebox = ttk.LabelFrame(main, text="Document", padding=10)
        filebox.pack(fill="x")

        self.file_label = ttk.Label(filebox, text="No document selected")
        self.file_label.pack(side="left", fill="x", expand=True)

        self.choose = ttk.Button(
            filebox, text="Choose Document", command=self.choose_file
        )
        self.choose.pack(side="right")

        options = ttk.LabelFrame(main, text="Settings", padding=10)
        options.pack(fill="x", pady=10)

        ttk.Label(options, text="Delay after each Google translation (seconds):").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Spinbox(
            options,
            from_=0.3,
            to=10.0,
            increment=0.1,
            textvariable=self.delay,
            width=8,
        ).grid(row=0, column=1, padx=10, sticky="w")

        ttk.Label(
            options,
            text="The browser remains visible so you can watch the same website being used.",
            foreground="#555",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        prog = ttk.LabelFrame(main, text="Progress / verification", padding=10)
        prog.pack(fill="both", expand=True)

        self.progress = ttk.Progressbar(prog, maximum=1)
        self.progress.pack(fill="x")

        self.progress_text = ttk.Label(prog, text="0 / 0")
        self.progress_text.pack(anchor="w", pady=(5, 7))

        self.log = tk.Text(prog, wrap="word", height=18, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

        controls = ttk.Frame(main)
        controls.pack(fill="x", pady=(10, 0))

        self.start_button = ttk.Button(
            controls,
            text="▶ Start",
            command=self.start,
        )
        self.start_button.pack(side="left")

        self.stop_button = ttk.Button(
            controls,
            text="■ Stop",
            command=self.stop,
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=8)

        ttk.Label(main, textvariable=self.status).pack(anchor="w", pady=(8, 0))

    def choose_file(self):
        if self.running:
            return

        path = filedialog.askopenfilename(
            title="Choose English document",
            filetypes=[
                ("Word document", "*.docx"),
                ("Text file", "*.txt"),
                ("Supported files", "*.docx *.txt"),
            ],
        )
        if path:
            self.input_path = Path(path)
            self.file_label.config(text=str(self.input_path))
            self.status.set("Ready.")

    def write_log(self, text):
        def f():
            self.log.insert("end", text + "\n")
            self.log.see("end")
        self.root.after(0, f)

    def start(self):
        if not self.input_path:
            messagebox.showwarning("Choose a document", "Please choose a document first.")
            return

        if self.input_path.suffix.lower() not in SUPPORTED:
            messagebox.showerror("Unsupported", "Only .docx and .txt are supported.")
            return

        self.running = True
        self.stop_requested = False
        self.start_button.config(state="disabled")
        self.choose.config(state="disabled")
        self.stop_button.config(state="normal")
        self.log.delete("1.0", "end")
        self.status.set("Starting Google Translate...")
        threading.Thread(target=self.worker, daemon=True).start()

    def stop(self):
        self.stop_requested = True
        self.status.set("Stopping after the current sentence...")

    def load_text_document(self):
        if self.input_path.suffix.lower() == ".txt":
            text = self.input_path.read_text(encoding="utf-8-sig")
            return text, None

        doc = Document(str(self.input_path))
        return None, doc

    def all_doc_paragraphs(self, doc):
        paragraphs = list(doc.paragraphs)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    paragraphs.extend(cell.paragraphs)

        for section in doc.sections:
            paragraphs.extend(section.header.paragraphs)
            paragraphs.extend(section.footer.paragraphs)

        return paragraphs

    def replace_paragraph_text(self, paragraph, new_text):
        # Keep the formatting of the first run where possible.
        if paragraph.runs:
            first = paragraph.runs[0]
            first.text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(new_text)

    def translate_sentence_pair(self, translator, sentence):
        if self.stop_requested:
            raise InterruptedError()

        core = sentence.strip()
        if not core:
            return sentence, "", ""

        leading = sentence[:len(sentence) - len(sentence.lstrip())]
        trailing = sentence[len(sentence.rstrip()):]

        # EXACT requested workflow using Google Translate website:
        # 1. English -> Bengali
        bn = translator.translate(core, "en", "bn")

        time.sleep(max(0.3, float(self.delay.get())))

        if self.stop_requested:
            raise InterruptedError()

        # 2. Bengali -> English
        en = translator.translate(bn, "bn", "en")

        time.sleep(max(0.3, float(self.delay.get())))

        return leading + en + trailing, bn, en

    def process_text(self, translator, text):
        sentences = split_sentences(text)
        total = sum(1 for s in sentences if s.strip())
        done = 0
        output = []

        self.root.after(0, lambda: self.progress.config(maximum=max(total, 1), value=0))

        for sentence in sentences:
            if self.stop_requested:
                raise InterruptedError()

            if not sentence.strip():
                output.append(sentence)
                continue

            done += 1
            self.write_log(f"\n[{done}/{total}] ORIGINAL:\n{sentence.strip()}")

            new_sentence, bn, en = self.translate_sentence_pair(
                translator, sentence
            )

            self.write_log(f"    BENGALI:\n    {bn}")
            self.write_log(f"    BACK TO ENGLISH:\n    {en}")

            if en.strip() != sentence.strip():
                self.write_log("    ✓ RESULT DIFFERED FROM ORIGINAL")
            else:
                self.write_log("    = RESULT IDENTICAL TO ORIGINAL")

            output.append(new_sentence)

            self.root.after(
                0,
                lambda done=done, total=total: (
                    self.progress.config(value=done),
                    self.progress_text.config(text=f"{done} / {total} sentences"),
                    self.status.set(f"Processing sentence {done} of {total}..."),
                ),
            )

        return "".join(output)

    def process_docx(self, translator, doc):
        paragraphs = self.all_doc_paragraphs(doc)

        # First count sentences.
        data = []
        total = 0
        for p in paragraphs:
            sentences = split_sentences(p.text)
            data.append((p, sentences))
            total += sum(1 for s in sentences if s.strip())

        self.root.after(0, lambda: self.progress.config(maximum=max(total, 1), value=0))

        done = 0

        for p, sentences in data:
            if self.stop_requested:
                raise InterruptedError()

            if not p.text.strip():
                continue

            new_parts = []

            for sentence in sentences:
                if not sentence.strip():
                    new_parts.append(sentence)
                    continue

                done += 1
                self.write_log(
                    f"\n[{done}/{total}] ORIGINAL:\n{sentence.strip()}"
                )

                new_sentence, bn, en = self.translate_sentence_pair(
                    translator, sentence
                )

                self.write_log(f"    BENGALI:\n    {bn}")
                self.write_log(f"    BACK TO ENGLISH:\n    {en}")

                if en.strip() != sentence.strip():
                    self.write_log("    ✓ RESULT DIFFERED FROM ORIGINAL")
                else:
                    self.write_log("    = RESULT IDENTICAL TO ORIGINAL")

                new_parts.append(new_sentence)

                self.root.after(
                    0,
                    lambda done=done, total=total: (
                        self.progress.config(value=done),
                        self.progress_text.config(text=f"{done} / {total} sentences"),
                        self.status.set(f"Processing sentence {done} of {total}..."),
                    ),
                )

            self.replace_paragraph_text(p, "".join(new_parts))

        return doc

    def worker(self):
        translator = None

        try:
            translator = GoogleWebTranslator(self.write_log)
            translator.start()

            if self.input_path.suffix.lower() == ".txt":
                text = self.input_path.read_text(encoding="utf-8-sig")
                new_text = self.process_text(translator, text)

                output = self.input_path.with_name(
                    self.input_path.stem + "_google_web_bn_retranslated.txt"
                )
                output.write_text(new_text, encoding="utf-8")

            else:
                doc = Document(str(self.input_path))
                doc = self.process_docx(translator, doc)

                output = self.input_path.with_name(
                    self.input_path.stem + "_google_web_bn_retranslated.docx"
                )
                doc.save(str(output))

            self.root.after(
                0,
                lambda output=output: messagebox.showinfo(
                    "Finished",
                    "Finished.\n\nNew document saved as:\n" + str(output),
                ),
            )
            self.root.after(0, lambda: self.status.set("Finished."))

        except InterruptedError:
            self.write_log("\nSTOPPED BY USER.")
            self.root.after(0, lambda: self.status.set("Stopped."))

        except Exception as e:
            self.write_log(f"\nERROR: {type(e).__name__}: {e}")
            self.root.after(
                0,
                lambda e=e: messagebox.showerror(
                    "Error",
                    f"{type(e).__name__}: {e}\n\n"
                    "The Google Translate browser window was left visible so "
                    "you can see what happened.",
                ),
            )
            self.root.after(0, lambda: self.status.set("Error."))

        finally:
            if translator:
                try:
                    translator.close()
                except Exception:
                    pass

            self.root.after(0, self.finished_ui)

    def finished_ui(self):
        self.running = False
        self.start_button.config(state="normal")
        self.choose.config(state="normal")
        self.stop_button.config(state="disabled")


def main():
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.1)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
