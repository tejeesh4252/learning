# automation/downloader_updated.py

import os
from datetime import datetime
from playwright.async_api import Page, Frame, TimeoutError as PlaywrightTimeout
from PyQt6.QtCore import QObject, pyqtSignal


class FundDownloader(QObject):
    """
    Automates navigation and fund report downloading
    from the OlisNet Report Generation Tool.

    Signals:
        progress  – (current: int, total: int, message: str)
        status    – plain status string
        error     – error message string
        finished  – emitted when all downloads complete
    """

    progress = pyqtSignal(int, int, str)
    status   = pyqtSignal(str)
    error    = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(
        self,
        page:            Page,
        funds:           list[dict],  # [{"name": "AREEF GE 3.1", "code": "157257"}]
        start_date:      str,         # "dd/mm/yyyy"
        end_date:        str,         # "dd/mm/yyyy"
        download_folder: str = ".",
        parent=None
    ):
        super().__init__(parent)
        self.page            = page
        self.frame: Frame    = None
        self.funds           = funds
        self.start_date      = start_date
        self.end_date        = end_date
        self.download_folder = download_folder

        # ── Timeouts (ms) ─────────────────────────────────────────────
        self.NAV_TIMEOUT    = 20_000
        self.ACTION_TIMEOUT = 10_000
        self.POPUP_TIMEOUT  = 15_000
        self.GEN_TIMEOUT    = 120_000
        self.DL_TIMEOUT     = 60_000


    # ══════════════════════════════════════════════════════════════════
    #  PUBLIC ENTRY POINT
    # ══════════════════════════════════════════════════════════════════

    async def run(self):
        """Main entry point — navigates to the tool, then loops funds."""
        try:
            self.status.emit("🔍 Navigating to Report Generation Tool...")
            await self._navigate_to_report_tool()

            total = len(self.funds)
            for idx, fund in enumerate(self.funds, start=1):
                self.progress.emit(idx, total, f"Downloading: {fund['name']}")
                self.status.emit(
                    f"📥 [{idx}/{total}] Processing: {fund['name']} ({fund['code']})"
                )
                await self._download_fund(fund)

            self.status.emit("✅ All funds downloaded successfully.")
            self.finished.emit()

        except PlaywrightTimeout as e:
            self.error.emit(f"⏱ Timeout: {e}")
        except Exception as e:
            self.error.emit(f"❌ Error: {e}")


    # ══════════════════════════════════════════════════════════════════
    #  NAVIGATION  (3 steps — runs on main page, then locks onto iframe)
    # ══════════════════════════════════════════════════════════════════

    async def _navigate_to_report_tool(self):
        await self._click_hamburger_menu()
        await self._hover_report_factory()
        await self._click_report_generation_tool()


    async def _click_hamburger_menu(self):
        """Step 1 — Click the top-left MENU button."""
        MENU_SELECTOR = '[data-test="menu-entry-MENU"]'
        self.status.emit("  → Opening menu...")
        await self.page.wait_for_selector(MENU_SELECTOR, timeout=self.NAV_TIMEOUT)
        await self.page.click(MENU_SELECTOR, timeout=self.ACTION_TIMEOUT)
        await self.page.wait_for_timeout(600)


    async def _hover_report_factory(self):
        """Step 2 — Hover REPORT FACTORY to reveal sub-menu."""
        REPORT_FACTORY_SELECTOR = '[data-test="menu-item"]:has-text("REPORT FACTORY")'
        self.status.emit("  → Hovering Report Factory...")
        await self.page.wait_for_selector(
            REPORT_FACTORY_SELECTOR, timeout=self.NAV_TIMEOUT
        )
        await self.page.hover(REPORT_FACTORY_SELECTOR, timeout=self.ACTION_TIMEOUT)
        await self.page.wait_for_timeout(1_000)


    async def _click_report_generation_tool(self):
        """Step 3 — Click Report Generation Tool and lock onto iframe."""
        REPORT_GEN_TOOL_SELECTOR = '[data-test="menu-item"]:has-text("Report Generation Tool")'
        self.status.emit("  → Opening Report Generation Tool...")

        await self.page.wait_for_selector(
            REPORT_GEN_TOOL_SELECTOR, timeout=self.NAV_TIMEOUT, state="visible"
        )
        await self.page.hover(REPORT_GEN_TOOL_SELECTOR, timeout=self.ACTION_TIMEOUT)
        await self.page.wait_for_timeout(400)
        await self.page.click(REPORT_GEN_TOOL_SELECTOR, timeout=self.ACTION_TIMEOUT)

        self.status.emit("  → Waiting for tool iframe to load...")
        await self.page.wait_for_selector('iframe', timeout=self.NAV_TIMEOUT)
        await self.page.wait_for_timeout(1_000)

        self.frame = await self._find_tool_frame()
        if self.frame is None:
            raise RuntimeError(
                "Could not find the Report Generation Tool iframe. "
                "Check the page loaded correctly."
            )
        self.status.emit("  ✓ Report Generation Tool iframe found and ready.")


    async def _find_tool_frame(self) -> Frame | None:
        """
        Scans all frames for the one containing 'Catalogue'.
        Retries up to 10× with 500ms gaps (max 5 seconds).
        """
        for attempt in range(10):
            for frame in self.page.frames:
                try:
                    el = await frame.query_selector(
                        'span[unselectable="on"]:has-text("Catalogue")'
                    )
                    if el:
                        self.status.emit(f"  ✓ Tool frame found → URL: {frame.url}")
                        return frame
                except Exception:
                    continue
            await self.page.wait_for_timeout(500)
        return None


    # ══════════════════════════════════════════════════════════════════
    #  PER-FUND DOWNLOAD CYCLE  (Steps A → G)
    # ══════════════════════════════════════════════════════════════════

    async def _download_fund(self, fund: dict):
        """Full A→G cycle for one fund."""
        await self._select_report_type()
        await self._select_fund(fund)
        await self._check_all_accounts()
        await self._set_start_date(self.start_date)
        await self._set_end_date(self.end_date)
        await self._ensure_exclude_blank_yes()
        await self._click_generate()
        await self._wait_for_download_complete(fund)


    # ── Step A: Expand tree and click "Bank Statement - PDF" ──────────

    async def _select_report_type(self):
        """
        Expands: Catalogue → Custody → Statement of Movements → Cash
        Then clicks: Bank Statement - PDF
        Uses + expand icon anchored to node label.
        Skips already-expanded nodes.
        """
        self.status.emit("  → Expanding report tree...")

        def plus_icon(node_text: str) -> str:
            return (
                f'div.x-tree-node-el:has(span:has-text("{node_text}")) '
                f'img[class*="elbow"][class*="plus"]'
            )

        EXPAND_STEPS = [
            ("Custody",                "Statement of Movements"),
            ("Statement of Movements", "Cash"),
            ("Cash",                   "Bank Statement - PDF"),
        ]

        for node_label, child_label in EXPAND_STEPS:
            child_selector = f'span[unselectable="on"]:has-text("{child_label}")'
            already_open   = await self.frame.is_visible(child_selector)

            if already_open:
                self.status.emit(f"  ✓ '{node_label}' already expanded — skipping.")
            else:
                icon_selector = plus_icon(node_label)
                self.status.emit(f"  → Expanding '{node_label}'...")
                await self.frame.wait_for_selector(
                    icon_selector, timeout=self.NAV_TIMEOUT
                )
                await self.frame.click(icon_selector, timeout=self.ACTION_TIMEOUT)
                await self.frame.wait_for_selector(
                    child_selector, timeout=self.NAV_TIMEOUT, state="visible"
                )
                self.status.emit(f"  ✓ '{node_label}' expanded.")

            await self.frame.wait_for_timeout(400)

        REPORT_SELECTOR = 'span[unselectable="on"]:has-text("Bank Statement - PDF")'
        self.status.emit("  → Clicking 'Bank Statement - PDF'...")
        await self.frame.wait_for_selector(REPORT_SELECTOR, timeout=self.NAV_TIMEOUT)
        await self.frame.click(REPORT_SELECTOR, timeout=self.ACTION_TIMEOUT)
        await self.frame.wait_for_timeout(600)
        self.status.emit("  ✓ Bank Statement - PDF selected.")


    # ── Step B: Select fund via search modal ──────────────────────────

    async def _select_fund(self, fund: dict):
        """
        Clicks [Select] → modal → type code → Search → tick → Add → close.

        ⚠️  Modal is a div overlay inside the iframe — NOT a browser popup.
        ⚠️  Modal does NOT auto-close after Add — closed via Escape.
        ⚠️  ExtJS keeps input[name="key"] in DOM — never wait for it hidden.
        ✅  page.keyboard (NOT frame.keyboard — frames have no keyboard attr).
        """
        self.status.emit(
            f"  → Selecting fund: {fund['name']} ({fund['code']})..."
        )

        # ── B1. Click [Select] ────────────────────────────────────────
        SELECT_BTN = 'button:has-text("Select"), input[value="Select"]'
        self.status.emit("  → Clicking Select button...")
        await self.frame.wait_for_selector(SELECT_BTN, timeout=self.NAV_TIMEOUT)
        await self.frame.click(SELECT_BTN, timeout=self.ACTION_TIMEOUT)
        await self.frame.wait_for_timeout(600)

        # ── B2. Wait for Search modal ─────────────────────────────────
        CODE_INPUT = 'input[name="key"]'
        self.status.emit("  → Waiting for Search modal...")
        await self.frame.wait_for_selector(
            CODE_INPUT, timeout=self.NAV_TIMEOUT, state="visible"
        )

        # ── B3. Type fund code ────────────────────────────────────────
        self.status.emit(f"  → Typing fund code: {fund['code']}...")
        await self.frame.fill(CODE_INPUT, fund["code"])
        await self.frame.wait_for_timeout(300)

        # ── B4. Click [Search] ────────────────────────────────────────
        SEARCH_BTN = 'button.x-btn-text:has-text("Search")'
        self.status.emit("  → Clicking Search...")
        await self.frame.wait_for_selector(SEARCH_BTN, timeout=self.NAV_TIMEOUT)
        await self.frame.click(SEARCH_BTN, timeout=self.ACTION_TIMEOUT)

        # ── B5. Wait for fund checkbox in results grid ────────────────
        padded_code   = fund["code"].zfill(7)   # "157257" → "0157257"
        FUND_CHECKBOX = (
            f'input[name="prompt_p_FND_searchGrid_cb_selected"]'
            f'[value="{padded_code}"]'
        )
        self.status.emit(f"  → Waiting for result row ({padded_code})...")
        await self.frame.wait_for_selector(
            FUND_CHECKBOX, timeout=self.NAV_TIMEOUT, state="visible"
        )
        await self.frame.wait_for_timeout(400)

        # ── B5b. Tick the fund checkbox ───────────────────────────────
        self.status.emit(f"  → Ticking checkbox for {fund['name']}...")
        await self.frame.check(FUND_CHECKBOX)
        await self.frame.wait_for_timeout(400)

        # ── B6. Click [Add] — exact text, JS dispatch bypasses mask ──
        # :text-is() = exact — avoids matching "Add from list"
        # dispatch_event = bypasses x-dlg-mask overlay blocking
        ADD_BTN_EXACT = 'button.x-btn-text:text-is("Add")'
        self.status.emit("  → Clicking Add (exact + JS dispatch)...")
        await self.frame.wait_for_selector(ADD_BTN_EXACT, timeout=self.NAV_TIMEOUT)

        add_el = await self.frame.query_selector(ADD_BTN_EXACT)
        if add_el:
            await add_el.scroll_into_view_if_needed()
            await self.frame.wait_for_timeout(300)
            await add_el.dispatch_event("click")
        else:
            # Hard JS fallback — exact text match
            await self.frame.evaluate("""
                () => {
                    const btns = [...document.querySelectorAll('button.x-btn-text')];
                    const add  = btns.find(b => b.textContent.trim() === 'Add');
                    if (add) add.click();
                }
            """)
        await self.frame.wait_for_timeout(600)

        # ── B7. Confirm fund appeared in main form grid ───────────────
        ADDED_ROW = f'td:has-text("{fund["code"]}")'
        self.status.emit("  → Confirming fund in selection grid...")
        try:
            await self.frame.wait_for_selector(
                ADDED_ROW, timeout=self.NAV_TIMEOUT, state="visible"
            )
            self.status.emit(f"  ✓ Fund '{fund['name']}' confirmed in grid.")
        except Exception:
            self.status.emit("  ⚠ Fund row not confirmed — continuing anyway.")

        # ── B8. Close modal via × button or Escape ────────────────────
        # ✅ page.keyboard — NOT frame.keyboard (frames have no keyboard)
        self.status.emit("  → Closing Search modal...")
        closed = False
        CLOSE_BTN = 'div.x-tool-close'
        try:
            close_el = await self.frame.query_selector(CLOSE_BTN)
            if close_el and await close_el.is_visible():
                await close_el.click(timeout=self.ACTION_TIMEOUT)
                closed = True
                self.status.emit("  ✓ Modal closed via × button.")
        except Exception:
            pass

        if not closed:
            await self.page.keyboard.press("Escape")
            self.status.emit("  ✓ Modal closed via Escape.")

        await self.frame.wait_for_timeout(500)

        # ── B9. Confirm modal container gone ──────────────────────────
        # ✅ Check x-window container — NOT input[name="key"] (stays in DOM)
        MODAL_WINDOW = 'div.x-window:has(input[name="key"])'
        try:
            await self.frame.wait_for_selector(
                MODAL_WINDOW, timeout=5_000, state="hidden"
            )
            self.status.emit("  ✓ Search modal confirmed closed.")
        except Exception:
            self.status.emit("  ✓ Modal not detected — moving on.")

        self.status.emit(f"  ✓ Step B complete: '{fund['name']}' selected.")


    # ── Step C: Tick "All" checkbox under Account ─────────────────────

    async def _check_all_accounts(self):
        """
        Ticks the 'All' checkbox under Account.
        HTML: <input class="x-form-checkbox x-form-field" ...>
        ✅  Class-based — name/id are session-generated and change each login.
        """
        self.status.emit("  → Checking 'All' accounts checkbox...")

        ALL_CHECKBOX = 'input.x-form-checkbox.x-form-field'
        await self.frame.wait_for_selector(ALL_CHECKBOX, timeout=self.NAV_TIMEOUT)

        is_checked = await self.frame.is_checked(ALL_CHECKBOX)
        if not is_checked:
            await self.frame.check(ALL_CHECKBOX)
            await self.frame.wait_for_timeout(300)
            self.status.emit("  ✓ 'All' accounts checked.")
        else:
            self.status.emit("  ✓ 'All' accounts already checked.")

        await self.frame.wait_for_timeout(300)


    # ── Step D: Set Start and End dates ───────────────────────────────

    async def _set_start_date(self, date_str: str):
        """
        Clicks the FIRST x-form-date-trigger (= Start Date).
        Waits for calendar to appear, then picks the date.
        HTML: <img class="x-form-trigger x-form-date-trigger" id="ext-gen338">
        ✅  query_selector_all()[0] — position-based, id is session-generated.
        """
        self.status.emit(f"  → Setting start date: {date_str}...")

        DATE_TRIGGER = 'img.x-form-date-trigger'
        await self.frame.wait_for_selector(DATE_TRIGGER, timeout=self.NAV_TIMEOUT)

        triggers = await self.frame.query_selector_all(DATE_TRIGGER)
        if not triggers:
            raise RuntimeError("No date triggers found in frame.")

        await triggers[0].click()

        # ✅ Wait for calendar to fully render before picking
        await self.frame.wait_for_selector(
            'div.x-date-picker', timeout=self.ACTION_TIMEOUT
        )
        await self.frame.wait_for_timeout(500)

        await self._pick_date(date_str)
        self.status.emit(f"  ✓ Start date set: {date_str}")


    async def _set_end_date(self, date_str: str):
        """
        Clicks the SECOND x-form-date-trigger (= End Date).
        Waits for calendar to appear, then picks the date.
        HTML: <img class="x-form-trigger x-form-date-trigger" id="ext-gen348">
        ✅  query_selector_all()[1] — position-based, id is session-generated.
        """
        self.status.emit(f"  → Setting end date: {date_str}...")

        DATE_TRIGGER = 'img.x-form-date-trigger'
        await self.frame.wait_for_selector(DATE_TRIGGER, timeout=self.NAV_TIMEOUT)

        triggers = await self.frame.query_selector_all(DATE_TRIGGER)
        if len(triggers) < 2:
            raise RuntimeError(f"Expected 2 date triggers, found {len(triggers)}.")

        await triggers[1].click()

        # ✅ Wait for calendar to fully render before picking
        await self.frame.wait_for_selector(
            'div.x-date-picker', timeout=self.ACTION_TIMEOUT
        )
        await self.frame.wait_for_timeout(500)

        debug_info = await self.frame.evaluate("""
            () => {
                const pickers = [...document.querySelectorAll('div.x-date-picker')];
                return pickers.map((p, i) => ({
                    index:   i,
                    id:      p.id,
                    visible: p.offsetParent !== null,
                    label:   (p.querySelector('td.x-date-middle button')?.innerText || '').trim()
                }));
            }
        """)
        self.status.emit(f"  DEBUG pickers: {debug_info}")

        await self._pick_date(date_str)
        self.status.emit(f"  ✓ End date set: {date_str}")

    async def _pick_date(self, date_str: str):
        """
        Navigates the ExtJS 3.x DatePicker to the correct month/year,
        then clicks the day.
        date_str format: 'dd/mm/yyyy'

        ✅ Scopes ALL interactions to the VISIBLE calendar instance only.
           ExtJS keeps multiple hidden pickers in DOM — must not hit them.
        """
        day, month, year = date_str.split("/")
        target_dt = datetime(int(year), int(month), 1)
        target_day = int(day)

        # ── Step 1: Find the index of the VISIBLE x-date-picker ───────────
        picker_index: int = await self.frame.evaluate("""
            () => {
                const pickers = [...document.querySelectorAll('div.x-date-picker')];
                for (let i = 0; i < pickers.length; i++) {
                    const p = pickers[i];
                    // offsetParent is null for display:none elements
                    if (p.offsetParent !== null) return i;
                }
                return -1;
            }
        """)

        if picker_index == -1:
            raise RuntimeError("No visible x-date-picker found in DOM.")

        self.status.emit(f"  → Using calendar instance [{picker_index}]")

        # ── Step 2: Navigate to the correct month ─────────────────────────
        label = ""
        for attempt in range(24):

            # Read month label from SCOPED visible picker only
            label = await self.frame.evaluate(f"""
                () => {{
                    const picker = document.querySelectorAll(
                        'div.x-date-picker'
                    )[{picker_index}];
                    if (!picker) return '';
                    const btn = picker.querySelector('td.x-date-middle button');
                    return btn ? (btn.innerText || btn.textContent || '').trim() : '';
                }}
            """)

            if not label:
                raise ValueError(
                    f"Cannot read calendar month label from picker [{picker_index}]."
                )

            current_dt = datetime.strptime(label.strip(), "%B %Y")
            self.status.emit(
                f"  → Calendar: {label.strip()} | Target: {target_dt.strftime('%B %Y')}"
            )

            if current_dt == target_dt:
                break

            # Direction
            nav_title = (
                "Next Month (Control+Right)"
                if current_dt < target_dt
                else "Previous Month (Control+Left)"
            )

            # ✅ Dispatch events scoped to the VISIBLE picker only
            fired = await self.frame.evaluate(f"""
                () => {{
                    const picker = document.querySelectorAll(
                        'div.x-date-picker'
                    )[{picker_index}];
                    if (!picker) return false;
                    const btn = picker.querySelector('a[title="{nav_title}"]');
                    if (!btn) return false;
                    ['mousedown', 'mouseup', 'click'].forEach(type => {{
                        btn.dispatchEvent(new MouseEvent(type, {{
                            bubbles:    true,
                            cancelable: true,
                            view:       window
                        }}));
                    }});
                    return true;
                }}
            """)

            if not fired:
                # Fallback: try the ExtJS component API directly
                self.status.emit(
                    f"  ⚠️  Nav button not found via title — trying ExtJS API..."
                )
                await self.frame.evaluate(f"""
                    () => {{
                        const picker = document.querySelectorAll(
                            'div.x-date-picker'
                        )[{picker_index}];
                        if (!picker || !picker.id) return;
                        const cmp = Ext.getCmp(picker.id);
                        if (!cmp) return;
                        {'cmp.showNextMonth()' if 'Next' in nav_title else 'cmp.showPrevMonth()'}
                    }}
                """)

            await self.frame.wait_for_timeout(500)  # slightly longer — ExtJS re-renders

        else:
            raise ValueError(
                f"Could not navigate to {target_dt.strftime('%B %Y')} "
                "after 24 attempts."
            )

        await self.frame.wait_for_timeout(400)

        # ── Step 3: Click the correct day — scoped to visible picker ──────
        self.status.emit(f"  → Clicking day {target_day}...")

        clicked = await self.frame.evaluate(f"""
            () => {{
                const picker = document.querySelectorAll(
                    'div.x-date-picker'
                )[{picker_index}];
                if (!picker) return false;
                const cells = [...picker.querySelectorAll('a.x-date-date')];
                const target = cells.find(
                    a => (a.innerText || a.textContent || '').trim() === '{target_day}'
                );
                if (!target) return false;
                ['mousedown', 'mouseup', 'click'].forEach(type => {{
                    target.dispatchEvent(new MouseEvent(type, {{
                        bubbles:    true,
                        cancelable: true,
                        view:       window
                    }}));
                }});
                return true;
            }}
        """)

        if clicked:
            self.status.emit(f"  ✓ Day {target_day} clicked.")
        else:
            raise ValueError(
                f"Could not find day {target_day} in calendar. "
                f"Month shown: {label.strip()}"
            )

        # ── Wait for THIS picker to close ─────────────────────────────────
        try:
            await self.frame.wait_for_function(
                f"""
                () => {{
                    const p = document.querySelectorAll(
                        'div.x-date-picker'
                    )[{picker_index}];
                    return !p || p.offsetParent === null;
                }}
                """,
                timeout=self.ACTION_TIMEOUT,
            )
        except Exception:
            pass  # Calendar may stay open — not fatal
        await self.frame.wait_for_timeout(300)


    # ── Step E: Ensure "Exclude blank pages" = Yes ────────────────────

    async def _ensure_exclude_blank_yes(self):
        """
        Sets 'Exclude blank pages' dropdown to Yes if not already set.
        HTML: <img class="x-form-trigger x-form-arrow-trigger" id="ext-gen358">
        ✅  Only one x-form-arrow-trigger on this form — safe to use class.
        """
        self.status.emit("  → Checking 'Exclude blank pages'...")

        ARROW_TRIGGER = 'img.x-form-arrow-trigger'
        COMBO_LIST    = 'div.x-combo-list'
        YES_OPTION    = 'div.x-combo-list-item:has-text("Yes")'

        await self.frame.wait_for_selector(ARROW_TRIGGER, timeout=self.NAV_TIMEOUT)

        # Try to read current value
        COMBO_INPUT = 'input.x-combo-noedit'
        try:
            current = await self.frame.input_value(COMBO_INPUT)
        except Exception:
            current = ""

        if current.strip().lower() == "yes":
            self.status.emit("  ✓ 'Exclude blank pages' already Yes.")
            return

        self.status.emit("  → Setting 'Exclude blank pages' to Yes...")
        await self.frame.click(ARROW_TRIGGER, timeout=self.ACTION_TIMEOUT)
        await self.frame.wait_for_selector(COMBO_LIST, timeout=self.ACTION_TIMEOUT)
        await self.frame.click(YES_OPTION, timeout=self.ACTION_TIMEOUT)
        await self.frame.wait_for_timeout(300)
        self.status.emit("  ✓ 'Exclude blank pages' set to Yes.")


    # ── Step F: Click Generate ────────────────────────────────────────

    async def _click_generate(self):
        """
        Clicks the Generate button.
        HTML: <button class="x-btn-text" type="button" id="ext-gen409">Generate</button>
        ✅  :text-is() exact match — avoids any other x-btn-text buttons.
        """
        GENERATE_BTN = 'button.x-btn-text:text-is("Generate")'

        self.status.emit("  → Clicking Generate...")
        await self.frame.wait_for_selector(GENERATE_BTN, timeout=self.NAV_TIMEOUT)
        await self.frame.click(GENERATE_BTN, timeout=self.ACTION_TIMEOUT)
        self.status.emit("  ✓ Generate clicked — report building...")


    # ── Step G: Wait for download link and save file ──────────────────

    async def _wait_for_download_complete(self, fund: dict):
        """
        Waits for the download link to appear in the iframe,
        clicks it, and saves the PDF to download_folder.
        ✅  Download captured on MAIN PAGE context even though
            click is fired on self.frame.
        """
        DOWNLOAD_LINK = 'a[onclick*="downloadNewReport.submit"]'

        self.status.emit(
            f"  ⏳ Waiting for report (up to {self.GEN_TIMEOUT // 1000}s)..."
        )

        await self.frame.wait_for_selector(
            DOWNLOAD_LINK, timeout=self.GEN_TIMEOUT, state="visible"
        )
        self.status.emit("  → Link ready. Downloading...")

        # Sanitise filename for Windows
        safe_name = (
            fund['name']
            .replace("/", "-")
            .replace("\\", "-")
            .replace(":", "-")
        )
        filename  = f"{safe_name}_{fund['code']}.pdf"
        save_path = os.path.join(self.download_folder, filename)

        async with self.page.expect_download(timeout=self.DL_TIMEOUT) as dl_info:
            await self.frame.click(DOWNLOAD_LINK, timeout=self.ACTION_TIMEOUT)

        download = await dl_info.value
        await download.save_as(save_path)
        self.status.emit(f"  ✅ Saved → {save_path}")
