"""Paint.NET 5.1.12 MCP server.

Automation notes that are NOT obvious from the code, and that cost real debugging:

  * The main window title gains a "*" prefix when the document is dirty, so it can
    never be matched exactly. Match auto_id="MainForm" plus a loose title regex.
  * The spinner sub-panes inside redUpDown/greenUpDown/... carry numeric auto_ids
    that are regenerated every launch. Never match those. auto_id="1001" (brush
    size) is a different animal: a fixed WinForms ComboBox child id, and is safe.
  * The Colors panel has an accessibility bug: saturationUpDown is titled "R:",
    valueUpDown "B:", alphaUpDown "G:". Matching those by title writes to the
    wrong channel silently. Always go parent auto_id -> child Edit.
  * The history, layer and document lists are owner-drawn: their ROWS have no
    UIA representation at all, in the control view or the raw view. Their
    scrollbars are real controls though, and the Layers scrollbar's RangeValue
    maximum is the list's content height -- 62px per layer -- so the layer COUNT
    is observable even though no individual layer is. Every layer operation is
    verified against that. By contrast the Colors panel's labels and spinners
    ARE real WinForms child controls, which is why they appear in the tree.
  * canvasView is the whole viewport and does not shrink to the image, so the
    canvas rectangle is not the image rectangle. Image position is recovered by
    calibration against the status bar cursor readout, never by arithmetic.
  * The selected tool renders as control_type "CheckBox" while every unselected
    tool is a "Button", so the active tool is found by scanning for the CheckBox.
    The index map below was confirmed a second way, by driving every tool from
    its keyboard shortcut and reading back the resulting index -- all 19 agree.
    Several shortcuts CYCLE within a family (S walks rectangle/lasso/ellipse/
    magic wand), so the number of presses depends on where the cycle sits.

Design rule used throughout: a read that fails must raise, not return a plausible
default. Silently returning "100%" or {"width": 0} is worse than crashing, because
the caller then believes it.
"""

import contextlib
import ctypes
import ctypes.wintypes
import hashlib
import io
import os
import pathlib
import re
import time

from fastmcp import FastMCP
from PIL import Image, ImageGrab
from pywinauto import Desktop
from pywinauto.application import Application

mcp = FastMCP("Paint.NET MCP Server")

EXE_PATH = "C:/Program Files/paint.net/paintdotnet.exe"

# toolStripEx children, 0-based, row-major over 2 columns. Index 18 spans both.
TOOL_MAP = {
    0: "rectangle-select", 1: "move-selected",
    2: "lasso-select", 3: "move-selection",
    4: "ellipse-select", 5: "zoom",
    6: "magic-wand", 7: "pan",
    8: "paint-bucket", 9: "gradient",
    10: "paintbrush", 11: "eraser",
    12: "pencil", 13: "color-picker",
    14: "clone-stamp", 15: "recolor",
    16: "text", 17: "line-curve",
    18: "shapes",
}
TOOL_MAP_INV = {name: idx for idx, name in TOOL_MAP.items()}

# PdnAuxMenu checkbox titles carry "\r\n" plus shortcut text, so match the prefix.
PANEL_PREFIX = {
    "Tools": "Tools (F5)",
    "History": "History (F6)",
    "Layers": "Layers (F7)",
    "Colors": "Colors (F8)",
}

SETTLE = 0.05          # poll interval while waiting for the UI to catch up
TIMEOUT = 2.0          # how long to wait for a value to prove itself fresh

# The cursor is driven with SetCursorPos rather than pywinauto.mouse so that
# every move can be verified to have landed before anything is read. Positions
# here are virtual-desktop coordinates and are routinely negative when the
# window sits on a monitor above or left of the primary one.
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004


class PdnError(RuntimeError):
    """Any Paint.NET automation failure. Carries context, never a default value."""


class Paint_Net_MCP:
    def __init__(self):
        self.app = None
        self.window = None
        self.calib_cache = None
        # Resolving a control by auto_id searches the whole UIA subtree, which
        # costs seconds on this window. Calibration polls the status bar in a
        # loop, so the resolved wrappers are held and revalidated rather than
        # looked up again per read.
        self._status_cache = None
        self._canvas_cache = None
        self._tool_strip_cache = None
        self._main_wrapper = None
        self._main_hwnd = None
        self._layers_scroll_cache = None
        self._layer_row_px = None
        self._tool_strip_cache = None

    def _invalidateElementCache(self):
        # The learned layer row height goes too: a DPI or theme change alters
        # it, and it is cheap to relearn on the next layer operation.
        self._layer_row_px = None
        self._status_cache = None
        self._canvas_cache = None
        self._tool_strip_cache = None
        self._main_wrapper = None
        self._main_hwnd = None
        self._layers_scroll_cache = None

    # ---------------------------------------------------------------- window

    def start(self):
        """Launch Paint.NET. Prefer attachOrLaunch; this always spawns a new one."""
        self.app = Application(backend="uia").start(EXE_PATH)
        self.window = self.app.window(auto_id="MainForm", title_re=r".*Paint\.NET.*")
        self.window.wait("ready", timeout=30)
        self._invalidateElementCache()
        return {"launched": True, "pid": self.app.process}

    def attachOrLaunch(self):
        """Connect to a running Paint.NET, launching one only if none is found."""
        if self.window is not None and self.window.exists():
            return {"attached": True, "launched": False, "pid": self.app.process}
        try:
            self.app = Application(backend="uia").connect(
                path=EXE_PATH, timeout=2)
            self.window = self.app.window(auto_id="MainForm",
                                          title_re=r".*Paint\.NET.*")
            self.window.wait("ready", timeout=10)
            self._invalidateElementCache()
            return {"attached": True, "launched": False, "pid": self.app.process}
        except Exception:
            result = self.start()
            result["attached"] = False
            return result

    def _mainWrapper(self):
        """Resolved main window, held rather than re-found.

        WindowSpecification re-runs a UIA search on every attribute access --
        exists(), window_text() and rectangle() each cost over a second, and
        they sit under every other call. Resolving once takes the cost from
        seconds to microseconds.
        """
        if self._main_wrapper is None:
            self._main_wrapper = self.window.wrapper_object()
            self._main_hwnd = self._main_wrapper.element_info.handle
        return self._main_wrapper

    def doesWindowExist(self):
        if self.window is None:
            return False
        try:
            hwnd = self._main_hwnd or self._mainWrapper().element_info.handle
        except Exception:
            return False
        user32 = ctypes.windll.user32
        if user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
            return True
        self._invalidateElementCache()          # window went away; drop stale wrappers
        return self.window.exists() and self.window.is_visible()

    def checkForApplicationWindow(self):
        if self.window is None:
            raise PdnError("Paint.NET is not attached. Call attach_or_launch first.")
        if not self.doesWindowExist():
            needed = 'open' if not self.window.exists() else 'visible'
            raise PdnError(f"Paint.NET is not {needed}.")

    def focus(self):
        self.checkForApplicationWindow()
        self._mainWrapper().set_focus()
        return {"focused": True}

    def _main(self):
        self.checkForApplicationWindow()
        return self.window

    def getControls(self):
        """Dump the UIA tree. This is the primitive every anchor here was derived
        from -- when something stops matching, run this and diff it."""
        self.checkForApplicationWindow()
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            self.window.print_control_identifiers()
        return buffer.getvalue()

    def dumpDialog(self, title_re=".*"):
        """Dump a modal/child dialog by title. Needed to add File > New / Open /
        Save As, whose dialogs have never been captured."""
        self.checkForApplicationWindow()
        dlg = self.app.window(title_re=title_re, top_level_only=False)
        dlg.wait("exists ready", timeout=10)
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            dlg.print_control_identifiers()
        return buffer.getvalue()

    # ----------------------------------------------------------- status bar

    def _resolveStatusBar(self):
        """Find the status bar once and keep its Text children and zoom button.

        Each of these lookups costs seconds; the calibration loop needs
        sub-second reads to sample the same value twice inside its timeout.
        """
        bar = self._main().child_window(auto_id="statusBar")
        try:
            children = bar.wrapper_object().children()
        except Exception as exc:
            raise PdnError(f"statusBar unreadable: {exc}") from exc

        texts, zoom = [], None
        for child in children:
            kind = child.element_info.control_type
            if kind == "Text":
                texts.append(child)
            elif kind == "Button" and child.window_text().strip().endswith("%"):
                zoom = child
        if len(texts) < 3:
            kinds = [c.element_info.control_type for c in children]
            raise PdnError(f"statusBar has {len(texts)} Text children, need 3. "
                           f"Children: {kinds}")
        return {"texts": texts, "zoom": zoom}

    def _statusCache(self):
        cache = self._status_cache
        if cache is not None:
            try:
                # Probe every cached element, not just the hint. Reading one
                # of four proves only that one is alive, and the others are
                # exactly the fields whose staleness would be invisible:
                # dimensions, cursor position and zoom.
                for text in cache["texts"]:
                    text.window_text()
                if cache["zoom"] is not None:
                    cache["zoom"].window_text()
                return cache
            except Exception:
                self._status_cache = None
        self._status_cache = self._resolveStatusBar()
        return self._status_cache

    def _statusTexts(self):
        """[0] hint, [1] image dims, [2] cursor position in image space.

        Index 2 can legitimately hold "" when the pointer is outside the window,
        so callers must distinguish "empty" from "missing".
        """
        return self._statusCache()["texts"]

    def _statusText(self, index):
        return self._statusTexts()[index].window_text()

    def _zoom(self):
        """Zoom button text, e.g. "100%".

        Found by scanning for a Button whose text ends in "%" rather than by
        regex: this value keys the calibration cache, so a matcher that silently
        matches nothing would freeze the cache and every later coordinate.
        """
        zoom = self._statusCache()["zoom"]
        if zoom is None:
            raise PdnError("No zoom button found in statusBar.")
        return zoom.window_text().strip()

    def _imageDims(self):
        """Parse "1024 x 768". The separator is U+00D7 MULTIPLICATION SIGN, not
        an ASCII x -- splitting on "x" matches nothing and yields zeroes."""
        raw = self._statusText(1)
        parts = re.split(r"[×x]", raw)
        if len(parts) != 2:
            raise PdnError(f"Cannot parse image dims from {raw!r}.")
        try:
            return {"width": int(parts[0].strip()), "height": int(parts[1].strip())}
        except ValueError as exc:
            raise PdnError(f"Non-numeric image dims in {raw!r}.") from exc

    def _cursorPos(self):
        """Cursor in image space, signed. "" when outside the window."""
        raw = self._statusText(2).strip()
        if not raw:
            return None
        parts = raw.split(",")
        if len(parts) != 2:
            raise PdnError(f"Cannot parse cursor position from {raw!r}.")
        try:
            return (int(parts[0]), int(parts[1]))
        except ValueError as exc:
            raise PdnError(f"Non-numeric cursor position in {raw!r}.") from exc

    # --------------------------------------------------------------- panels

    def _ensurePanel(self, name):
        """Panels vanish from the UIA tree when closed, so anything touching
        ColorsForm/ToolsForm/etc must open them first."""
        prefix = PANEL_PREFIX.get(name)
        if prefix is None:
            raise PdnError(f"Unknown panel {name!r}. Known: {sorted(PANEL_PREFIX)}")
        win = self._main()
        checkbox = win.child_window(auto_id="PdnAuxMenu").child_window(
            title_re="^" + re.escape(prefix), control_type="CheckBox")
        if not checkbox.exists():
            raise PdnError(f"Panel toggle for {name!r} not found in PdnAuxMenu.")
        if not checkbox.get_toggle_state():
            checkbox.click()
            # Toggling a panel relays the window, which rebuilds status bar
            # controls; cached wrappers would answer for dead elements.
            self._invalidateElementCache()
        form = win.child_window(auto_id=f"{name}Form")
        try:
            form.wait("exists visible", timeout=5)
        except Exception as exc:
            raise PdnError(f"Panel {name!r} did not open: {exc}") from exc
        return form

    # ---------------------------------------------------------------- tools

    def _toolStrip(self):
        # Cached: this is a three-deep auto_id chase, and getActiveToolIndex is
        # called in polling loops after every tool change.
        cached = self._tool_strip_cache
        if cached is not None:
            try:
                cached.rectangle()
                return cached
            except Exception:
                self._tool_strip_cache = None
        self._ensurePanel("Tools")
        self._tool_strip_cache = self._main().child_window(auto_id="ToolsForm") \
            .child_window(auto_id="toolsControl") \
            .child_window(auto_id="toolStripEx").wrapper_object()
        return self._tool_strip_cache

    def getActiveToolIndex(self):
        """The selected tool is the only child rendered as a CheckBox."""
        children = self._toolStrip().children()
        for i, child in enumerate(children):
            if child.element_info.control_type == "CheckBox":
                return i
        kinds = [c.element_info.control_type for c in children]
        raise PdnError(f"No active tool: no CheckBox among {len(children)} "
                       f"toolStripEx children. Types: {kinds}")

    def selectTool(self, tool):
        """Select by name ("paintbrush") or index (10)."""
        idx = tool if isinstance(tool, int) else TOOL_MAP_INV.get(tool)
        if idx is None or idx not in TOOL_MAP:
            raise PdnError(f"Unknown tool {tool!r}. Known: {sorted(TOOL_MAP_INV)}")

        before = self.getActiveToolIndex()
        if before == idx:
            return {"changed": False, "index": idx, "tool": TOOL_MAP[idx]}

        hint_before = self._statusText(0)
        children = self._toolStrip().children()
        if idx >= len(children):
            raise PdnError(f"Tool index {idx} out of range "
                           f"({len(children)} children).")
        try:
            children[idx].click()
        except Exception:
            children[idx].click_input()

        # Poll until the CheckBox has actually moved. A read taken immediately
        # after the click often still reports the previous tool.
        deadline = time.time() + TIMEOUT
        active = None
        while time.time() < deadline:
            active = self.getActiveToolIndex()
            if active == idx:
                break
            time.sleep(SETTLE)
        if active != idx:
            raise PdnError(f"Tool selection failed: asked for {idx} "
                           f"({TOOL_MAP[idx]}), active is {active}.")

        hint_after = self._statusText(0)
        return {
            "changed": True, "index": idx, "tool": TOOL_MAP[idx],
            "hint_changed": hint_before != hint_after,
        }

    # Tool shortcuts. Several keys CYCLE through a family on repeated presses,
    # so the number of presses matters and depends on where the cycle currently
    # sits -- hence press-and-check rather than press-a-fixed-number-of-times.
    TOOL_KEYS = {
        0: ("s", 4), 2: ("s", 4), 4: ("s", 4), 6: ("s", 4),      # select family
        1: ("m", 2), 3: ("m", 2),                                 # move family
        17: ("o", 2), 18: ("o", 2),                               # line / shapes
        5: ("z", 1), 7: ("h", 1), 8: ("f", 1), 9: ("g", 1),
        10: ("b", 1), 11: ("e", 1), 12: ("p", 1), 13: ("k", 1),
        14: ("l", 1), 15: ("r", 1), 16: ("t", 1),
    }

    def selectToolByKey(self, tool):
        """Select a tool with its keyboard shortcut.

        Faster than clicking, and it does not need the Tools panel open at all.

        Two hazards. Some keys cycle within a family (S walks rectangle, lasso,
        ellipse, magic wand), so the key is pressed until the active tool index
        actually reaches the target rather than a fixed number of times. And if
        the Text tool has an uncommitted entry, a bare letter would be TYPED
        INTO THE IMAGE -- so any pending text is committed with Escape first.
        """
        idx = tool if isinstance(tool, int) else TOOL_MAP_INV.get(tool)
        if idx is None or idx not in TOOL_MAP:
            raise PdnError(f"Unknown tool {tool!r}. Known: {sorted(TOOL_MAP_INV)}")
        if idx not in self.TOOL_KEYS:
            raise PdnError(f"No keyboard shortcut recorded for {TOOL_MAP[idx]}.")

        key, cycle = self.TOOL_KEYS[idx]
        before = self.getActiveToolIndex()
        if before == idx:
            return {"changed": False, "index": idx, "tool": TOOL_MAP[idx],
                    "presses": 0, "key": key}

        # Commit any in-progress text first; otherwise the keystroke is content.
        if before == TOOL_MAP_INV["text"]:
            self._hotkey("{ESC}")
            time.sleep(0.2)

        for press in range(1, cycle + 1):
            self._hotkey(key)
            deadline = time.time() + 1.0
            active = None
            while time.time() < deadline:
                active = self.getActiveToolIndex()
                if active != before or active == idx:
                    break
                time.sleep(SETTLE)
            if active == idx:
                return {"changed": True, "index": idx, "tool": TOOL_MAP[idx],
                        "presses": press, "key": key}
            before = active
        raise PdnError(f"{key!r} pressed {cycle} times did not reach "
                       f"{TOOL_MAP[idx]} (index {idx}); stopped at "
                       f"{self.getActiveToolIndex()}.")

    # --------------------------------------------------------------- colors

    def _colorsEdit(self, auto_id):
        """Reach a Colors spinner's Edit through its parent auto_id. The Edit
        titles are their own values and several parent titles name the wrong
        channel, so neither can be matched on."""
        form = self._ensurePanel("Colors")
        pane = form.child_window(auto_id=auto_id)
        edits = [c for c in pane.wrapper_object().children()
                 if c.element_info.control_type == "Edit"]
        if not edits:
            raise PdnError(f"No Edit under {auto_id}.")
        return edits[0]

    def getPrimaryColor(self):
        form = self._ensurePanel("Colors")
        return form.child_window(auto_id="hexBox").get_value()

    def setPrimaryColor(self, hex_color):
        value = hex_color.lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", value):
            raise PdnError(f"Expected RRGGBB hex, got {hex_color!r}.")
        form = self._ensurePanel("Colors")
        box = form.child_window(auto_id="hexBox").wrapper_object()
        box.set_edit_text(value)
        box.type_keys("{ENTER}")

        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if box.get_value().lstrip("#").upper() == value:
                return {"primary_color": value}
            time.sleep(SETTLE)
        raise PdnError(f"Colour did not take: wanted {value}, "
                       f"hexBox reads {box.get_value()!r}.")

    # ----------------------------------------------------------- brush conf

    def _brushSizeEdit(self):
        """auto_id="1001" is a fixed WinForms id, not one of the randomised ones.

        drawConfigStrip is rebuilt for each tool, so this control is absent
        entirely for tools that have no brush size -- pencil, zoom, pan, the
        selection tools. Absence is a fact about the active tool, not a fault.
        """
        spec = self._main().child_window(auto_id="drawConfigStrip") \
            .child_window(auto_id="1001")
        if not spec.exists(timeout=1):
            raise PdnError(
                f"No brush-size control for the active tool "
                f"({TOOL_MAP.get(self.getActiveToolIndex())}). drawConfigStrip "
                f"is rebuilt per tool and this one has no brush size.")
        return spec.wrapper_object()

    def getBrushSize(self):
        return int(self._brushSizeEdit().get_value())

    def setBrushSize(self, pixels):
        if not isinstance(pixels, int) or pixels < 1:
            raise PdnError(f"Brush size must be a positive int, got {pixels!r}.")
        edit = self._brushSizeEdit()
        edit.set_edit_text(str(pixels))
        edit.type_keys("{ENTER}")

        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if edit.get_value().strip() == str(pixels):
                return {"brush_size": pixels}
            time.sleep(SETTLE)
        raise PdnError(f"Brush size did not take: wanted {pixels}, "
                       f"reads {edit.get_value()!r}.")

    def _toleranceSliders(self):
        """drawConfigStrip holds two panes with auto_id="ToleranceSliderControl".
        They are told apart only by x order: left is Hardness, right is Spacing."""
        strip = self._main().child_window(auto_id="drawConfigStrip").wrapper_object()
        panes = [c for c in strip.children()
                 if c.element_info.automation_id == "ToleranceSliderControl"]
        if len(panes) != 2:
            raise PdnError(f"Expected 2 ToleranceSliderControl panes, "
                           f"found {len(panes)}.")
        return sorted(panes, key=lambda c: c.rectangle().left)

    # The slider paints its filled portion in the accent colour. Nothing about
    # its value is exposed through UIA -- no Value or RangeValue pattern, no
    # children, no name, no legacy value -- so the fill is measured from the
    # pixels instead. Verified against known settings: 25 -> 24.9, 50 -> 50.3,
    # 75 -> 75.1, 100 -> 100.0.
    SLIDER_FILL_RGB = (0, 120, 215)
    SLIDER_FILL_TOLERANCE = 60
    SLIDER_MIN_FILL_PCT = 4.6      # the bar keeps a rounded cap even at zero

    def _sliderFillPercent(self, pane):
        """Filled fraction of a slider, as a percentage.

        Measured as the rightmost column containing fill colour, scanning every
        row: the percentage text is drawn ON TOP of the bar, so a simple run of
        matching pixels stops early where the digits cross it.
        """
        rect = pane.rectangle()
        self._parkCursor()
        image = ImageGrab.grab(
            bbox=(rect.left + 1, rect.top + 1, rect.right - 1, rect.bottom - 1),
            all_screens=True).convert("RGB")
        width, height = image.size
        if width < 4:
            raise PdnError(f"Slider is only {width}px wide; cannot measure it.")
        last = -1
        for x in range(width):
            for y in range(height):
                pixel = image.getpixel((x, y))
                if max(abs(pixel[i] - self.SLIDER_FILL_RGB[i])
                       for i in range(3)) <= self.SLIDER_FILL_TOLERANCE:
                    last = x
                    break
        return (last + 1) / width * 100.0

    def _setSlider(self, pane, percent, name):
        if not 0 <= percent <= 100:
            raise PdnError(f"Percent must be 0-100, got {percent}.")
        rect = pane.rectangle()
        x = int(rect.left + (rect.width() - 1) * percent / 100.0)
        y = int((rect.top + rect.bottom) / 2)
        self._clickAt(x, y)
        time.sleep(0.4)

        measured = self._sliderFillPercent(pane)
        target = max(percent, self.SLIDER_MIN_FILL_PCT)
        verified = abs(measured - target) <= 5.0
        result = {"requested_percent": percent,
                  "measured_percent": round(measured, 1),
                  "verified": verified}
        if not verified:
            result["reason"] = (f"{name} slider reads {measured:.1f}% after "
                                f"asking for {percent}%")
        return result

    def getHardness(self):
        return {"hardness_percent": round(
            self._sliderFillPercent(self._toleranceSliders()[0]), 1)}

    def getSpacing(self):
        return {"spacing_percent": round(
            self._sliderFillPercent(self._toleranceSliders()[1]), 1)}

    def setHardness(self, percent):
        return self._setSlider(self._toleranceSliders()[0], percent, "Hardness")

    def setSpacing(self, percent):
        return self._setSlider(self._toleranceSliders()[1], percent, "Spacing")

    # ---------------------------------------------------------- calibration

    def _canvasWrapper(self):
        cached = self._canvas_cache
        if cached is not None:
            try:
                cached.rectangle()          # cheap liveness probe
                return cached
            except Exception:
                self._canvas_cache = None
        self._canvas_cache = self._main().child_window(
            auto_id="canvasView").wrapper_object()
        return self._canvas_cache

    def _canvasRect(self):
        # Resolving canvasView by auto_id is a full subtree search costing
        # seconds, and this is called several times per drawing operation.
        return self._canvasWrapper().rectangle()

    def _occluders(self):
        """Rectangles of the floating panels, which sit ON TOP of canvasView.

        canvasView's rectangle is the entire viewport, so its corners are
        normally covered by Tools (top left), History (top right), Colors
        (bottom left) and Layers (bottom right). Probing a covered point moves
        the pointer over a panel, Paint.NET never sees a canvas mouse-move, and
        the status bar keeps reporting wherever the pointer was last -- a stale
        reading that looks perfectly valid.
        """
        # Win32, not UIA: resolving the four panels by auto_id costs about
        # five seconds each, which dominated calibration entirely. They are
        # top-level windows, so their rectangles come back instantly by handle.
        rects = []
        user32 = ctypes.windll.user32
        for hwnd, _cls, title in self._enumWindows():
            if title not in self._PANEL_TITLES:
                continue
            rect = ctypes.wintypes.RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                rects.append(rect)
        return rects

    def _cursorTolerance(self):
        """Landing tolerance in pixels, widened under display scaling.

        Coordinates round-trip through the DPI virtualisation layer, so a
        perfectly good move can land a pixel or two out on a scaled display."""
        try:
            dpi = ctypes.windll.user32.GetDpiForWindow(
                self._mainWrapper().element_info.handle) or 96
        except Exception:
            dpi = 96
        return max(2, int(round(2 * dpi / 96.0)))

    def _moveTo(self, x, y, verify=True):
        """Place the cursor and confirm it landed. Never silently mis-positions."""
        if not ctypes.windll.user32.SetCursorPos(int(x), int(y)):
            raise PdnError(f"SetCursorPos({int(x)}, {int(y)}) failed.")
        if verify:
            actual = self._pointerAt()
            tolerance = self._cursorTolerance()
            if (abs(actual[0] - int(x)) > tolerance
                    or abs(actual[1] - int(y)) > tolerance):
                raise PdnError(
                    f"Cursor did not land: asked for ({int(x)}, {int(y)}), "
                    f"got {actual}, tolerance {tolerance}px. Coordinate may "
                    f"be off-screen.")
        return self._pointerAt()

    @staticmethod
    def _mouseDown():
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    @staticmethod
    def _mouseUp():
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def _clickAt(self, x, y):
        self._moveTo(x, y)
        time.sleep(SETTLE)
        self._mouseDown()
        time.sleep(SETTLE)
        self._mouseUp()

    @staticmethod
    def _pointerAt():
        """Physical cursor position. Signed, so it works on monitors laid out
        left of or above the primary display."""
        point = ctypes.wintypes.POINT()
        if not ctypes.windll.user32.GetCursorPos(ctypes.byref(point)):
            raise PdnError("GetCursorPos failed.")
        return (point.x, point.y)

    @staticmethod
    def _isCovered(x, y, rects, margin=6):
        return any(r.left - margin <= x <= r.right + margin and
                   r.top - margin <= y <= r.bottom + margin for r in rects)

    def _probePoints(self):
        """Three uncovered points inside canvasView: two to solve from, one to
        verify against. Spread out so their readings cannot collide."""
        rect = self._canvasRect()
        occluders = self._occluders()
        fractions = (0.12, 0.26, 0.4, 0.5, 0.62, 0.76, 0.88)
        free = [(int(rect.left + rect.width() * fx),
                 int(rect.top + rect.height() * fy))
                for fx in fractions for fy in fractions]
        free = [p for p in free if not self._isCovered(p[0], p[1], occluders)]
        if len(free) < 3:
            raise PdnError(
                f"Only {len(free)} uncovered points inside canvasView; floating "
                f"panels leave too little canvas to calibrate. Close a panel "
                f"(F5-F8) or enlarge the window.")

        best = None
        for i, a in enumerate(free):
            for b in free[i + 1:]:
                dx, dy = abs(a[0] - b[0]), abs(a[1] - b[1])
                if dx < 50 or dy < 50:
                    continue          # need span on BOTH axes to solve
                if best is None or dx * dy > best[0]:
                    best = (dx * dy, a, b)
        if best is None:
            raise PdnError("No two uncovered canvas points differ enough on both "
                           "axes to solve a calibration.")
        _, a, b = best
        check = next((p for p in free
                      if p[0] not in (a[0], b[0]) and p[1] not in (a[1], b[1])), None)
        return a, b, check

    def _calibKeys(self):
        """Everything that invalidates the screen<->image mapping."""
        win_rect = self._mainWrapper().rectangle()
        canvas = self._canvasRect()
        return {
            "window": (win_rect.left, win_rect.top, win_rect.right, win_rect.bottom),
            "canvas": (canvas.left, canvas.top, canvas.right, canvas.bottom),
            "zoom": self._zoom(),
            "dims": self._statusText(1),
            "document": self.getActiveDocName(),
        }

    def _probeCursor(self, x, y):
        """Move the pointer there and return the settled readout.

        Movement only -- clicking here would draw on the user's image with
        whatever tool is active.

        "Settled" means the same value three times running. Settling alone does
        NOT prove the value is current: a status bar that has not updated is
        also perfectly stable. Freshness is established by the caller, which
        probes a second point and returns, and checks the readings track.
        """
        self._moveTo(x, y)
        deadline = time.time() + TIMEOUT
        last, repeats = None, 0
        while time.time() < deadline:
            # If the physical pointer is no longer where we put it, someone is
            # using the mouse. Say so, rather than blaming the status bar: the
            # readout is then reporting their pointer, not our probe.
            actual = self._pointerAt()
            tolerance = self._cursorTolerance()
            if (abs(actual[0] - int(x)) > tolerance
                    or abs(actual[1] - int(y)) > tolerance):
                raise PdnError(
                    f"Pointer moved externally: put it at ({int(x)}, {int(y)}), "
                    f"found it at {actual} (tolerance {tolerance}px). "
                    f"Calibration drives the real cursor -- do not use the "
                    f"mouse while it runs.")
            pos = self._cursorPos()
            if pos is not None and pos == last:
                repeats += 1
                if repeats >= 2:
                    return pos
            else:
                last, repeats = pos, 0
            time.sleep(SETTLE)
        if last is None:
            raise PdnError(f"Status bar cursor readout stayed empty at "
                           f"({x}, {y}); the pointer may be over a floating panel.")
        raise PdnError(f"Status bar cursor readout never settled at ({x}, {y}); "
                       f"last saw {last}.")

    def _wakeCanvasTracking(self):
        """Nudge Paint.NET into updating the cursor readout again.

        Selecting a tool restarts tracking; nothing else tried here does --
        not clicking the canvas, not synthesised move events, not refocusing
        the window. The original tool is restored afterwards, and no tool
        selection touches the image.
        """
        original = self.getActiveToolIndex()
        spare = 7 if original != 7 else 5          # pan, or zoom; both harmless
        self.selectTool(spare)
        self.selectTool(original)
        return original

    def calibrate(self, force=False):
        """Solve the affine screen->image mapping by probing two points.

        Derived rather than computed, so it survives zoom, scroll, DPI scaling
        and panel layout changes. The readout reports negative coordinates
        outside the image, so probe points need not land on the image itself.
        """
        keys = self._calibKeys()
        if not force and self.calib_cache and self.calib_cache["keys"] == keys:
            return dict(self.calib_cache["params"], cached=True)

        stray = self.openDialogs()
        if stray:
            raise PdnError(f"A modal dialog is open ({stray}); the canvas gets no "
                           f"mouse input while it is, so calibration cannot run. "
                           f"Close it first.")
        self.focus()
        rect = self._canvasRect()
        if rect.width() < 60 or rect.height() < 60:
            raise PdnError(f"canvasView too small to calibrate: "
                           f"{rect.width()}x{rect.height()}.")

        dirty_before = self.isDirty()

        # Points must avoid the floating panels that cover canvasView's corners.
        (ax_s, ay_s), (bx_s, by_s), check = self._probePoints()

        # Probe A, B, then A again. The return trip is what proves the readout
        # actually tracks the pointer: a stale status bar yields the same value
        # at both points, and a value that was stale when A was first read will
        # not reproduce on the way back. Comparing against a "previous" reading
        # cannot work here, because that reading is itself unverified -- if the
        # pointer already sits on A, A's correct value never changes and looks
        # exactly like a frozen one.
        ax_i, ay_i = self._probeCursor(ax_s, ay_s)
        bx_i, by_i = self._probeCursor(bx_s, by_s)

        if (ax_i, ay_i) == (bx_i, by_i):
            # Paint.NET sometimes stops updating the position readout entirely;
            # it then reports one frozen value everywhere. Switching tools makes
            # it resume, so do that once and retry before giving up. Tool
            # selection does not touch the image.
            self._wakeCanvasTracking()
            ax_i, ay_i = self._probeCursor(ax_s, ay_s)
            bx_i, by_i = self._probeCursor(bx_s, by_s)
            if (ax_i, ay_i) == (bx_i, by_i):
                raise PdnError(
                    f"Status bar reported the same position {(ax_i, ay_i)} at "
                    f"two different screen points {(ax_s, ay_s)} and "
                    f"{(bx_s, by_s)}, even after a tool change; it is not "
                    f"tracking the pointer.")

        again = self._probeCursor(ax_s, ay_s)
        # A couple of pixels of drift between visits to the same point is normal
        # -- the view can settle slightly after layer or selection changes. The
        # point of the return trip is to catch a FROZEN readout, which repeats a
        # completely unrelated value; the fit against the third point below is
        # what actually validates the numbers.
        drift = max(abs(again[0] - ax_i), abs(again[1] - ay_i))
        if drift > 2:
            raise PdnError(f"Status bar did not reproduce {(ax_i, ay_i)} on "
                           f"returning to {(ax_s, ay_s)}; it read {again} "
                           f"({drift}px away). The first reading was stale.")
        if bx_i == ax_i or by_i == ay_i:
            raise PdnError(f"Calibration probes gave a degenerate image-space "
                           f"span: A=({ax_i},{ay_i}) B=({bx_i},{by_i}).")

        mx = (bx_i - ax_i) / (bx_s - ax_s)
        my = (by_i - ay_i) / (by_s - ay_s)
        params = {
            "mx": mx, "bx": ax_i - mx * ax_s,
            "my": my, "by": ay_i - my * ay_s,
        }

        # Prove the solved mapping against a third point it was not fitted to.
        # Two points always fit a line; only a third can show the fit is wrong,
        # which is what a stale or occluded reading produces.
        if check is not None:
            cx_i, cy_i = self._probeCursor(check[0], check[1])
            px = params["mx"] * check[0] + params["bx"]
            py = params["my"] * check[1] + params["by"]
            if abs(px - cx_i) > 1.5 or abs(py - cy_i) > 1.5:
                raise PdnError(
                    f"Calibration failed verification at screen {check}: "
                    f"predicted image ({px:.1f}, {py:.1f}), status bar reports "
                    f"({cx_i}, {cy_i}). The mapping is not affine over the probed "
                    f"points -- a reading was probably stale.")

        if self.isDirty() and not dirty_before:
            raise PdnError("Calibration mutated the document -- it must not.")

        self.calib_cache = {"keys": keys, "params": params}
        return dict(params, cached=False)

    def screenToImage(self, x, y):
        p = self.calibrate()
        return (int(round(p["mx"] * x + p["bx"])),
                int(round(p["my"] * y + p["by"])))

    def imageToScreen(self, x, y):
        p = self.calibrate()
        if p["mx"] == 0 or p["my"] == 0:
            raise PdnError("Degenerate calibration; cannot invert.")
        return (int(round((x - p["bx"]) / p["mx"])),
                int(round((y - p["by"]) / p["my"])))

    # ----------------------------------------------------------- doc state

    def isDirty(self):
        self.checkForApplicationWindow()
        return self._mainWrapper().window_text().startswith("*")

    def getActiveDocName(self):
        self.checkForApplicationWindow()
        title = self._mainWrapper().window_text().lstrip("*")
        return re.sub(r"\s*-\s*Paint\.NET.*$", "", title)

    def getCanvasDimensions(self):
        """Image dimensions in pixels.

        Reads the status bar text. Note this is NOT the canvasView rectangle:
        canvasView is the viewport and stays the same size when the image is
        resized. Taking its rectangle -- or the rectangle of the label holding
        this text -- gives a wrong answer that never raises.
        """
        dims = self._imageDims()
        return (dims["width"], dims["height"])

    def getState(self):
        win = self._main()
        win_rect = self._mainWrapper().rectangle()
        canvas = self._canvasRect()
        active = self.getActiveToolIndex()
        try:
            brush_size, brush_reason = self.getBrushSize(), None
        except PdnError as exc:
            brush_size, brush_reason = None, str(exc)
        return {
            "window_rect": self._rectToDict(win_rect),
            "canvas_rect": self._rectToDict(canvas),
            "image_dims": self._imageDims(),
            "image_dims_raw": self._statusText(1),
            "zoom": self._zoom(),
            "active_tool": {"index": active, "name": TOOL_MAP.get(active)},
            "primary_color": self.getPrimaryColor(),
            # Absent for tools without a brush size. Report why rather than
            # crashing the whole state read over one optional control.
            "brush_size": brush_size,
            "brush_size_reason": brush_reason,
            "dirty": self.isDirty(),
            "document": {
                "name": self.getActiveDocName(),
                "count": None,
                "count_reason": "not enumerable via UIA "
                                "(documentStrip exposes no children)",
            },
            "hint": self._statusText(0),
        }

    @staticmethod
    def _rectToDict(rect):
        # RECT is not JSON-serialisable, and every tool return crosses MCP.
        return {"left": rect.left, "top": rect.top,
                "right": rect.right, "bottom": rect.bottom,
                "width": rect.width(), "height": rect.height()}

    # -------------------------------------------------------- mutation check

    def _imageScreenBox(self):
        """Screen rectangle of the image, clipped to the visible canvas."""
        dims = self._imageDims()
        left, top = self.imageToScreen(0, 0)
        right, bottom = self.imageToScreen(dims["width"], dims["height"])
        canvas = self._canvasRect()
        return (max(min(left, right), canvas.left),
                max(min(top, bottom), canvas.top),
                min(max(left, right), canvas.right),
                min(max(top, bottom), canvas.bottom))

    def _parkCursor(self):
        """Move the pointer off the canvas before any pixel capture.

        Paint.NET draws a brush outline on the canvas itself, and that overlay
        is part of the window, so a screen grab includes it. Left in place it
        makes every hash differ purely because the pointer moved, which would
        let the mutation check pass on an operation that drew nothing.
        """
        win = self._mainWrapper().rectangle()
        # Verify the park landed. An unverified move that silently failed would
        # leave the pointer over the canvas, and the brush outline Paint.NET
        # draws there would end up in the very grab this is protecting.
        canvas = self._canvasRect()
        for x, y in ((win.left + 8, win.top + 8),
                     (win.right - 8, win.top + 8),
                     (win.left + 8, win.bottom - 8)):
            # Landing anywhere over canvasView defeats the purpose: Paint.NET
            # would still draw its brush outline into the grab this protects.
            if (canvas.left <= x <= canvas.right
                    and canvas.top <= y <= canvas.bottom):
                continue
            try:
                self._moveTo(x, y)
                break
            except PdnError:
                continue
        else:
            raise PdnError("Could not park the pointer off the canvas; every "
                           "candidate point was over canvasView or "
                           "unreachable.")
        time.sleep(0.25)

    def _grabHash(self, box):
        shot = ImageGrab.grab(bbox=box, all_screens=True)
        return hashlib.sha256(shot.tobytes()).hexdigest()

    def _canvasHash(self, park=True):
        """Hash of the image region, sampled until two grabs agree.

        A single grab can land mid-repaint -- switching tools, dismissing the
        brush outline, or finishing a stroke all leave the canvas briefly
        different from its settled state. An unsettled hash would read as a
        mutation that never happened.
        """
        if park:
            self._parkCursor()
        box = self._imageScreenBox()
        if box[2] <= box[0] or box[3] <= box[1]:
            raise PdnError(f"Image is not visible on screen: box={box}.")

        # Three consecutive agreeing grabs, not two: dismissing the brush
        # outline can take longer than one sample interval, so a pair of grabs
        # can agree while both still contain it.
        deadline = time.time() + TIMEOUT
        previous, agreements = self._grabHash(box), 0
        while time.time() < deadline:
            time.sleep(0.12)
            current = self._grabHash(box)
            if current == previous:
                agreements += 1
                if agreements >= 2:
                    return current
            else:
                agreements = 0
            previous = current
        # Something on the canvas animates continuously -- marching ants around
        # a selection, for one. Report the last value; callers comparing it will
        # see spurious differences, which is why mutation checks also look at
        # the dirty flag rather than pixels alone.
        return previous

    def _mutationToken(self):
        """Dirty flag alone is insufficient: it is already set after the first
        edit, so a second edit would not change it. Pair it with a pixel hash."""
        return (self.isDirty(), self._canvasHash())

    def _assertMutated(self, before, what):
        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if self._mutationToken() != before:
                return True
            time.sleep(SETTLE)
        raise PdnError(f"{what} produced no visible change and left the document "
                       f"clean; it probably did nothing.")

    # ---------------------------------------------------------------- draw

    def drawPolyline(self, points, tool="paintbrush"):
        """Drag through a list of [x, y] image-space points."""
        if not points or len(points) < 2:
            raise PdnError("Need at least 2 points.")
        self.selectTool(tool)
        self.focus()
        screen = [self.imageToScreen(int(p[0]), int(p[1])) for p in points]
        before = self._mutationToken()

        self._moveTo(*screen[0])
        time.sleep(SETTLE)
        self._mouseDown()
        try:
            for point in screen[1:]:
                self._moveTo(*point)
                time.sleep(SETTLE)
        finally:
            self._mouseUp()

        self._assertMutated(before, "drawPolyline")
        return {"points": len(points), "tool": tool, "dirty": self.isDirty()}

    def drawLine(self, x1, y1, x2, y2, tool="paintbrush"):
        return self.drawPolyline([[x1, y1], [x2, y2]], tool=tool)

    def drawRect(self, x1, y1, x2, y2, tool="paintbrush"):
        """Traced as four segments rather than using the shapes tool, whose
        shape-picker dropdown has never been captured in a UIA dump."""
        return self.drawPolyline(
            [[x1, y1], [x2, y1], [x2, y2], [x1, y2], [x1, y1]], tool=tool)

    def drawEllipse(self, cx, cy, rx, ry, segments=48, tool="paintbrush"):
        import math
        points = [[cx + rx * math.cos(2 * math.pi * i / segments),
                   cy + ry * math.sin(2 * math.pi * i / segments)]
                  for i in range(segments + 1)]
        return self.drawPolyline(points, tool=tool)

    def floodFill(self, x, y):
        self.selectTool("paint-bucket")
        self.focus()
        before = self._mutationToken()
        self._clickAt(*self.imageToScreen(int(x), int(y)))
        self._assertMutated(before, "floodFill")
        return {"filled_at": [x, y], "dirty": self.isDirty()}

    def addText(self, x, y, text):
        self.selectTool("text")
        self.focus()
        before = self._mutationToken()
        self._clickAt(*self.imageToScreen(int(x), int(y)))
        time.sleep(0.2)
        self._main().type_keys(text, with_spaces=True, set_foreground=False)
        # Commit the text layer; until this the text is still an editable overlay.
        self._main().type_keys("{ESC}")
        self._assertMutated(before, "addText")
        return {"text": text, "at": [x, y], "dirty": self.isDirty()}

    def drawShape(self, start, stop, type="rectangle"):
        """Kept from the original surface. start/stop are [x, y] image points."""
        if type == "rectangle":
            return self.drawRect(start[0], start[1], stop[0], stop[1])
        if type == "ellipse":
            cx, cy = (start[0] + stop[0]) / 2, (start[1] + stop[1]) / 2
            return self.drawEllipse(cx, cy, abs(stop[0] - start[0]) / 2,
                                    abs(stop[1] - start[1]) / 2)
        if type == "line":
            return self.drawLine(start[0], start[1], stop[0], stop[1])
        raise PdnError(f"Unknown shape {type!r}. Known: rectangle, ellipse, line.")

    # -------------------------------------------------------------- history

    def _historyButtons(self):
        form = self._ensurePanel("History")
        strip = form.child_window(auto_id="toolStrip").wrapper_object()
        buttons = strip.children()
        if len(buttons) < 2:
            raise PdnError(f"HistoryForm toolStrip has {len(buttons)} buttons, "
                           f"expected 2 (undo, redo).")
        return buttons

    def undo(self):
        before = self._canvasHash()
        self._historyButtons()[0].click()
        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if self._canvasHash() != before:
                return {"undone": True, "dirty": self.isDirty()}
            time.sleep(SETTLE)
        return {"undone": False, "dirty": self.isDirty(),
                "reason": "canvas unchanged; history may be at its oldest entry"}

    def redo(self):
        before = self._canvasHash()
        self._historyButtons()[1].click()
        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if self._canvasHash() != before:
                return {"redone": True, "dirty": self.isDirty()}
            time.sleep(SETTLE)
        return {"redone": False, "dirty": self.isDirty(),
                "reason": "canvas unchanged; history may be at its newest entry"}

    # --------------------------------------------------------------- layers

    def _openMenuItems(self, main_class):
        """Visible MenuItems that belong to an open popup, not to the menu bar.

        The popup is its own top-level-ish window with a different WinForms
        class, so it cannot be reached by descending from MainForm; it has to
        be found process-wide and then told apart from the menu bar's own items
        by its parent's class.
        """
        # Only accept items belonging to an actual menu popup. Testing merely
        # that the parent class differs from the main window would also sweep in
        # menu items from any other window this process happens to have open --
        # a dialog with its own menu bar, for instance.
        popups = {hwnd for hwnd, cls, title in self._enumWindows()
                  if cls != main_class and not title and cls.startswith("WindowsForms")}
        found = []
        for item in Desktop(backend="uia").windows(
                process=self.app.process, top_level_only=False,
                control_type="MenuItem", visible_only=True):
            try:
                # "System" is the window's own system menu. It is always present
                # whether or not a menu is open, so it must not count as one.
                if item.window_text() == "System":
                    continue
                parent = item.parent()
                if parent.element_info.class_name == main_class:
                    continue
                if popups and parent.element_info.handle not in popups:
                    continue
                found.append(item)
            except Exception:
                continue
        return found

    def _menuPopupOpen(self):
        """True if a menu popup window exists. Cheap Win32 check.

        Popup menus live in their own WinForms window class, distinct from the
        main window's, so they can be spotted without walking the UIA tree.
        """
        main_class = None
        popups = []
        for hwnd, cls, title in self._enumWindows():
            if "Paint.NET" in title:
                main_class = cls
            popups.append((cls, title))
        if main_class is None:
            return False
        return any(cls != main_class and not title and cls.startswith("WindowsForms")
                   for cls, title in popups)

    def _closeOpenMenus(self):
        """Menus are modal to the mouse: while one is open the canvas receives
        no input and the next menu click closes it instead of opening one."""
        win = self._main()
        for _ in range(3):
            if not self._menuPopupOpen():
                return
            win.type_keys("{ESC}")
            time.sleep(0.2)

    def listMenuItems(self, menu):
        """Open a top-level menu, list what is in it, close it again."""
        win = self._main()
        self._closeOpenMenus()
        win.child_window(auto_id="PdnMainMenu").child_window(
            title=menu, control_type="MenuItem").click_input()
        time.sleep(0.4)
        try:
            names = [i.window_text() for i in
                     self._openMenuItems(win.element_info.class_name)]
        finally:
            win.type_keys("{ESC}")
        return {"menu": menu, "items": names}

    def _menuPick(self, menu, item):
        """Drive PdnMainMenu. Layer operations have to go through here because
        layersStrip exposes no children to click."""
        win = self._main()
        main_class = win.element_info.class_name
        bar = win.child_window(auto_id="PdnMainMenu")

        # Straight after a paste or a merge, Paint.NET can swallow the click
        # that opens the menu, so an empty popup means "try again" rather than
        # "no such item". Only an open menu that lacks the item is a real miss.
        seen = []
        for attempt in range(3):
            self._closeOpenMenus()
            bar.child_window(title=menu, control_type="MenuItem").click_input()

            deadline = time.time() + TIMEOUT
            while time.time() < deadline:
                entries = self._openMenuItems(main_class)
                if entries:
                    seen = [e.window_text() for e in entries]
                    for entry in entries:
                        if entry.window_text() == item:
                            entry.click_input()
                            return True
                    break                      # menu is open, item genuinely absent
                time.sleep(SETTLE)
            if seen:
                break
            time.sleep(0.3)                    # menu never opened; let it settle

        self._closeOpenMenus()
        raise PdnError(f"Menu item {item!r} not found under {menu!r}. "
                       f"Items offered: {seen or '(menu never opened)'}")

    # Titles of the dockable panels, which are windows but not dialogs.
    _PANEL_TITLES = ("Tools", "History", "Colors", "Layers")

    def _enumWindows(self):
        """(hwnd, class, title) for every visible window of the process.

        Win32 enumeration, not UIA: a Desktop(...).windows() walk costs seconds
        and this runs before every hotkey. That difference turned a single
        exact-draw into a two-and-a-half minute operation.
        """
        results = []
        pid = ctypes.wintypes.DWORD()
        proc = ctypes.wintypes.DWORD(self.app.process)
        user32 = ctypes.windll.user32

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND,
                            ctypes.wintypes.LPARAM)
        def collect(hwnd, _):
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == proc.value and user32.IsWindowVisible(hwnd):
                title = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title, 256)
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, cls, 256)
                results.append((hwnd, cls.value, title.value))
            return True

        user32.EnumWindows(collect, 0)
        return results

    def openDialogs(self):
        """Modal dialogs currently open.

        A modal dialog stops the canvas receiving mouse input, so the status bar
        cursor readout freezes and calibration cannot work. A leaked dialog
        therefore breaks everything downstream in a way that looks like a
        Paint.NET bug rather than an automation one.
        """
        return [title for _, _, title in self._enumWindows()
                if title and title not in self._PANEL_TITLES
                and "Paint.NET" not in title]

    def recover(self):
        """Get Paint.NET back to a usable state.

        A script killed mid-operation can leave a modal dialog or an open menu
        behind, and either one silently disables all canvas input -- every
        subsequent read then looks like a Paint.NET fault rather than a leaked
        dialog. This closes them and drops cached wrappers.
        """
        closed, menus = [], False
        for title in self.openDialogs():
            dlg = self.app.window(title=title, top_level_only=False)
            for _ in range(4):
                if not dlg.exists(timeout=1):
                    break
                for auto_id in ("cancelButton", "okButton"):
                    try:
                        dlg.child_window(auto_id=auto_id,
                                         control_type="Button").click_input()
                        break
                    except Exception:
                        continue
                else:
                    try:
                        dlg.type_keys("{ESC}")
                    except Exception:
                        pass
                time.sleep(0.4)
            closed.append(title)
        if self._menuPopupOpen():
            self._closeOpenMenus()
            menus = True
        self._invalidateElementCache()
        remaining = self.openDialogs()
        return {"closed_dialogs": closed, "closed_menus": menus,
                "still_open": remaining, "recovered": not remaining}

    def _closeDialog(self, dlg, name):
        """Close a dialog and confirm it actually went away."""
        for _ in range(4):
            if not dlg.exists(timeout=1):
                return True
            try:
                dlg.child_window(auto_id="cancelButton",
                                 control_type="Button").click_input()
            except Exception:
                try:
                    dlg.type_keys("{ESC}")
                except Exception:
                    pass
            time.sleep(0.5)
        raise PdnError(f"Could not close the {name} dialog. While it is open the "
                       f"canvas receives no mouse input and calibration will fail.")

    def _layerPropsDialog(self):
        """Open Layer Properties (F4). This is the only way to read anything
        about a layer: layersStrip exposes no children at all."""
        self._closeOpenMenus()
        self.focus()
        self._main().type_keys("{F4}")
        # Paint.NET dialogs are not top-level windows, so the default
        # top_level_only=True never finds them.
        dlg = self.app.window(title="Layer Properties", top_level_only=False)
        try:
            dlg.wait("exists ready visible", timeout=5)
        except Exception as exc:
            raise PdnError(f"Layer Properties dialog did not open: {exc}") from exc
        return dlg

    @staticmethod
    def _layerFields(dlg):
        wrapper = dlg.wrapper_object()
        name, opacity, blend, visible = None, None, None, None
        for child in wrapper.descendants():
            info = child.element_info
            if info.control_type == "Edit":
                if info.automation_id == "textBox":
                    name = child
                elif opacity is None:
                    opacity = child          # the unnamed Edit is the opacity box
            elif info.automation_id == "comboBox":
                blend = child
            elif info.automation_id == "checkBox":
                visible = child
        return name, opacity, blend, visible

    def getLayerProperties(self):
        """Name, opacity, blend mode and visibility of the ACTIVE layer.

        Cancels out rather than confirming, so reading never modifies anything.
        Note this describes only the active layer -- the layer list itself is
        still not enumerable, so there is no way to count layers or inspect
        the others without selecting them first.
        """
        dlg = self._layerPropsDialog()
        try:
            name, opacity, blend, visible = self._layerFields(dlg)
            state = {
                "name": name.get_value() if name else None,
                "opacity": int(opacity.get_value()) if opacity else None,
                "blend_mode": blend.window_text() if blend else None,
            }
            if visible is not None:
                try:
                    state["visible"] = bool(visible.get_toggle_state())
                except Exception as exc:
                    state["visible"] = None
                    state["visible_reason"] = f"toggle state unreadable: {exc}"
            return state
        finally:
            self._closeDialog(dlg, "Layer Properties")

    def setLayerProperties(self, name=None, opacity=None, blend_mode=None):
        """Change properties of the active layer, then verify by reading back.

        Visibility is deliberately absent: the checkbox exposes neither a Toggle
        pattern nor a checked state, so its current value cannot be read and a
        declarative set(visible=True) could not be honoured or verified. Use
        toggleLayerVisibility instead, which is explicit about flipping it.
        """
        if all(v is None for v in (name, opacity, blend_mode)):
            raise PdnError("Nothing to set.")
        if opacity is not None and not 0 <= int(opacity) <= 255:
            raise PdnError(f"Opacity must be 0-255, got {opacity}.")

        dlg = self._layerPropsDialog()
        try:
            name_box, opacity_box, blend_box, visible_box = self._layerFields(dlg)
            # Type, then confirm the control actually holds the value
            # BEFORE committing. A field that silently refused the text
            # would otherwise be committed as an unchanged value and the
            # failure would surface later as a mystery.
            for attempt in range(3):
                if name is not None and name_box:
                    name_box.set_edit_text(str(name))
                if opacity is not None and opacity_box:
                    opacity_box.set_edit_text(str(int(opacity)))
                if blend_mode is not None and blend_box:
                    blend_box.select(blend_mode)
                time.sleep(0.2)
                pending = []
                if name is not None and name_box:
                    if name_box.get_value() != str(name):
                        pending.append("name")
                if opacity is not None and opacity_box:
                    if opacity_box.get_value().strip() != str(int(opacity)):
                        pending.append("opacity")
                if not pending:
                    break
            else:
                raise PdnError(f"Layer Properties would not accept "
                               f"{pending} after 3 attempts.")
            # A real click, not the Invoke pattern: invoking okButton
            # raises nothing and does nothing -- the dialog just stays
            # open and the edit is never committed.
            ok = dlg.child_window(auto_id="okButton", control_type="Button")
            ok.click_input()
            time.sleep(0.3)
            if dlg.exists(timeout=1):
                try:
                    ok.click()          # last resort
                except Exception:
                    pass
        except Exception:
            self._closeDialog(dlg, "Layer Properties")
            raise

        # Read back only once this dialog is gone. Opening a second Layer
        # Properties on top of a first that failed to close leaves the
        # original dangling, and a dangling modal blocks all canvas input.
        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            if not dlg.exists(timeout=1):
                break
            time.sleep(SETTLE)
        else:
            # Do NOT cancel here: OK has already been pressed, and
            # cancelling a dialog that is merely slow to close would
            # discard the change and then report it as not applied.
            raise PdnError("Layer Properties did not close after OK; "
                           "leaving it open rather than cancelling, "
                           "which would discard the edit. Call recover().")

        after = self.getLayerProperties()
        wanted = {"name": name, "opacity": opacity, "blend_mode": blend_mode}
        mismatched = {k: {"wanted": v, "got": after.get(k)}
                      for k, v in wanted.items()
                      if v is not None and after.get(k) != v}
        if mismatched:
            raise PdnError(f"Layer properties did not take: {mismatched}")
        return after

    def toggleLayerVisibility(self):
        """Flip the active layer's Visible checkbox.

        Reported unverified on purpose: the checkbox has no Toggle pattern and
        never sets STATE_SYSTEM_CHECKED, so there is no way to read back whether
        the layer is now shown or hidden.
        """
        dlg = self._layerPropsDialog()
        try:
            _, _, _, visible_box = self._layerFields(dlg)
            if visible_box is None:
                raise PdnError("No Visible checkbox in Layer Properties.")
            visible_box.click_input()
            dlg.child_window(auto_id="okButton", control_type="Button").click()
        except Exception:
            self._closeDialog(dlg, "Layer Properties")
            raise
        time.sleep(0.4)
        return {"toggled": True, "verified": False,
                "reason": "Visible checkbox exposes no toggle or checked state"}

    # Keyboard shortcuts beat the menus for layer work: no popup to open, no
    # timing window where Paint.NET swallows the click that opens it.
    LAYER_KEYS = {
        "add": "^+n",           # Ctrl+Shift+N  new blank layer
        "delete": "^+{DEL}",    # Ctrl+Shift+Del
        "duplicate": "^+d",     # Ctrl+Shift+D
        "merge_down": "^m",     # Ctrl+M -- NOT Ctrl+Shift+M, which does nothing.
                                # A wrong merge key fails silently: the drawing
                                # stays on its own layer and the composited
                                # pixels look identical, so pixel verification
                                # still passes while layers quietly pile up.
    }

    def _hotkey(self, keys):
        """Send a shortcut to the main window with nothing in the way."""
        stray = self.openDialogs()
        if stray:
            raise PdnError(f"A modal dialog is open ({stray}); it would swallow "
                           f"the keystroke. Close it first.")
        self._closeOpenMenus()
        self.focus()
        self._mainWrapper().type_keys(keys)
        return keys

    # The layer list is owner-drawn and exposes no UIA children, but its
    # scrollbar does expose a RangeValue pattern whose maximum is the total
    # content height. That grows by exactly one row per layer, so the layer
    # count is observable even though the layers themselves are not.
    UIA_RANGEVALUE_PATTERN = 10003
    # Starting estimate only. Every layer operation changes the count by a
    # known amount, so the true row height is measured from the resulting
    # extent delta and replaces this -- a theme, DPI or UI-scale change
    # would otherwise skew every count derived from it.
    LAYER_ROW_PX = 62.0

    def _layersScrollBar(self):
        cached = self._layers_scroll_cache
        if cached is not None:
            try:
                cached.rectangle()
                return cached
            except Exception:
                self._layers_scroll_cache = None
        form = self._ensurePanel("Layers")
        # Paint.NET hides the scrollbar when the layer list fits, and
        # pywinauto searches only visible elements by default -- so with
        # few layers the control is not found unless asked for explicitly.
        for visible_only in (True, False):
            spec = form.child_window(auto_id="layersScrollBar",
                                     visible_only=visible_only)
            if spec.exists(timeout=1):
                self._layers_scroll_cache = spec.wrapper_object()
                return self._layers_scroll_cache
        raise PdnError("layersScrollBar is not present; the layer count "
                       "cannot be read for this document.")

    def _layersExtent(self):
        """Total content height of the Layers list, in pixels."""
        element = self._layersScrollBar().element_info.element
        try:
            import comtypes.gen.UIAutomationClient as UIAClient
            pattern = element.GetCurrentPattern(self.UIA_RANGEVALUE_PATTERN)
            pattern = pattern.QueryInterface(
                UIAClient.IUIAutomationRangeValuePattern)
            return float(pattern.CurrentMaximum)
        except Exception as exc:
            raise PdnError(f"Layers scrollbar exposes no range: {exc}") from exc

    def _rowPx(self):
        return self._layer_row_px or self.LAYER_ROW_PX

    def _learnRowHeight(self, delta_extent, layers_changed):
        """Derive the row height from an operation whose effect is known."""
        if not layers_changed:
            return
        measured = abs(delta_extent) / abs(layers_changed)
        if 8.0 <= measured <= 400.0:
            self._layer_row_px = measured

    def layerCount(self):
        """Number of layers, inferred from the list's scroll extent.

        Indirect by necessity. The row height is measured in pixels, so a theme
        or DPI change would alter it; a non-integral result is reported rather
        than rounded away.
        """
        try:
            extent = self._layersExtent()
        except PdnError as exc:
            return {"count": None, "extent": None,
                    "row_px": self._rowPx(), "reason": str(exc)}
        rows = extent / self._rowPx()
        count = round(rows)
        if abs(rows - count) > 0.05:
            return {"count": None, "extent": extent, "row_px": self._rowPx(),
                    "row_px_measured": self._layer_row_px is not None,
                    "reason": f"extent {extent} is not a whole number of "
                              f"{self._rowPx()}px rows ({rows:.3f}); the row "
                              f"height differs from the one in use"}
        return {"count": count, "extent": extent, "row_px": self._rowPx(),
                "row_px_measured": self._layer_row_px is not None}

    def _layerOp(self, action, verify):
        """Run a layer shortcut and verify it took effect.

        Verified against the Layers list scroll extent, which tracks the
        layer count. The active layer NAME is not used: a duplicate inherits
        the original's name, and repeated duplicates leave several layers
        sharing one, so an unchanged name proves nothing either way. Reading
        it would also cost two modal F4 dialog cycles per operation.

        The scrollbar only exists while the list overflows its panel, so on
        a short list there is no oracle at all and the result says so rather
        than claiming success.
        """
        # Verified by the scroll extent, which counts layers directly. The
        # active layer NAME is unusable as an oracle: a duplicate inherits the
        # original's name, and repeated duplicates leave several layers sharing
        # one name, so a name that did not change proves nothing either way.
        expected = {"add": 1, "duplicate": 1, "delete": -1, "merge_down": -1}
        before_extent = None
        if verify:
            try:
                before_extent = self._layersExtent()
            except PdnError:
                verify = False          # short list: no scrollbar, no oracle
        self._hotkey(self.LAYER_KEYS[action])
        time.sleep(0.5)
        result = {"action": action, "dirty": self.isDirty()}
        if not verify:
            result.update(verified=False,
                          reason="not verified; the Layers list scroll extent "
                                 "was unavailable, which happens when the "
                                 "list is short enough to need no scrollbar")
            return result
        try:
            after_extent = self._layersExtent()
        except PdnError as exc:
            return dict(result, verified=False,
                        reason=f"layer count unreadable after the operation: "
                               f"{exc}")
        delta = after_extent - before_extent
        delta_rows = delta / self._rowPx()
        matched = round(delta_rows) == expected[action]
        if matched:
            self._learnRowHeight(delta, expected[action])
        result.update(layers_before=round(before_extent / self._rowPx()),
                      layers_after=round(after_extent / self._rowPx()),
                      row_px=self._rowPx(), verified=matched)
        if not result["verified"]:
            result["reason"] = (f"layer count changed by {delta_rows:.2f} rows, "
                                f"expected {expected[action]:+d}")
        return result

    def zoomActualSize(self):
        """Ctrl+0. Pixel verification only works at 100%, so this makes the
        precondition reachable instead of merely reported."""
        self._hotkey("^0")
        time.sleep(0.4)
        self.calib_cache = None          # zoom change invalidates the mapping
        zoom = self._zoom()
        if zoom != "100%":
            # Before believing that, rule out a stale cached wrapper: a
            # zoom that did change would otherwise be reported as a failure.
            self._invalidateElementCache()
            zoom = self._zoom()
        if zoom != "100%":
            raise PdnError(f"Zoom is {zoom} after Ctrl+0, expected 100%.")
        return {"zoom": zoom}

    def flatten(self):
        """Ctrl+Shift+F. Collapses every layer into one."""
        before = self.layerCount().get("count")
        self._hotkey("^+f")
        time.sleep(0.8)
        after = self.layerCount().get("count")
        result = {"action": "flatten", "layers_before": before,
                  "layers_after": after, "dirty": self.isDirty(),
                  "verified": after == 1 or after is None}
        if after is None:
            result["reason"] = ("layer count unreadable after flattening: "
                                "with one layer the list needs no scrollbar. "
                                "Consistent with success, not proof of it")
        elif after != 1:
            result["reason"] = (f"expected a single layer after flattening, "
                                f"the list extent reports {after}")
        return result

    def toggleLayerVisibilityFast(self):
        """Ctrl+, -- no dialog.

        Preferred over the Layer Properties route: that dialog is modal, and a
        leaked one silently disables all canvas input. Still unverifiable,
        because visibility cannot be read back either way.
        """
        before = self._canvasHash()
        self._hotkey("^,")
        time.sleep(0.5)
        changed = self._canvasHash() != before
        return {"action": "toggle_layer_visibility", "dirty": self.isDirty(),
                "canvas_changed": changed, "verified": False,
                "reason": "visibility is not readable; a canvas change is "
                          "evidence only when the layer had visible content"}

    def layerAdd(self, verify=True):
        return self._layerOp("add", verify)

    def layerDuplicate(self, verify=True):
        return self._layerOp("duplicate", verify)

    def clearLayerContents(self):
        """Ctrl+A then Delete: empty the active layer without removing it."""
        before = self._canvasHash()
        self._hotkey("^a")
        time.sleep(0.2)
        self._hotkey("{DEL}")
        time.sleep(0.4)
        self._hotkey("^d")          # drop the select-all
        time.sleep(0.3)
        changed = self._canvasHash() != before
        return {"action": "clear_layer_contents", "dirty": self.isDirty(),
                "verified": changed,
                **({} if changed else
                   {"reason": "canvas unchanged; the layer was probably "
                              "already empty"})}

    def layerAddViaMenu(self):
        """Menu route, kept because it is the only one that proves the command
        exists rather than assuming a shortcut is bound to it."""
        before = self.layerCount().get("count")
        self._menuPick("Layers", "Add New Layer")
        time.sleep(0.5)
        after = self.layerCount().get("count")
        result = {"action": "layer_add_via_menu", "layers_before": before,
                  "layers_after": after, "dirty": self.isDirty(),
                  "verified": before is not None and after == before + 1}
        if not result["verified"]:
            result["reason"] = (f"layer count went {before} -> {after}; a "
                                f"menu click can be swallowed, so the click "
                                f"alone is not evidence")
        return result

    def layerDelete(self, verify=True):
        return self._layerOp("delete", verify)

    def layerMergeDown(self, verify=True):
        return self._layerOp("merge_down", verify)

    # ------------------------------------------------------------ capture

    def screenshotCanvas(self, path=None):
        """PNG of the image region only, not the whole viewport."""
        import base64
        self._parkCursor()      # keep the brush outline overlay out of the shot
        box = self._imageScreenBox()
        if box[2] <= box[0] or box[3] <= box[1]:
            raise PdnError(f"Image is not visible on screen: box={box}.")
        shot = ImageGrab.grab(bbox=box, all_screens=True)
        if path:
            shot.save(path, "PNG")
            return {"path": path, "size": list(shot.size)}
        buffer = io.BytesIO()
        shot.save(buffer, "PNG")
        return {"png_base64": base64.b64encode(buffer.getvalue()).decode("ascii"),
                "size": list(shot.size)}

    # ------------------------------------------------------- exact drawing

    # Pasting is preferable to dragging a synthetic mouse: the geometry is
    # exact, nothing is smoothed by the brush engine, no calibration is
    # involved, and the pointer never moves. Paint.NET honours the alpha
    # channel of a CF_DIBV5 paste, so a mostly-transparent overlay composites
    # over the existing pixels instead of replacing the layer.

    @staticmethod
    def _toDIBV5(img):
        """RGBA image -> BITMAPV5HEADER + bottom-up BGRA rows."""
        import struct
        img = img.convert("RGBA")
        width, height = img.size
        r, g, b, a = img.split()
        rows = Image.merge("RGBA", (b, g, r, a)).transpose(
            Image.FLIP_TOP_BOTTOM).tobytes()
        header = struct.pack(
            "<IiiHHIIiiII" "IIII" "I" + "i" * 9 + "III" "IIII",
            124, width, height, 1, 32, 3, len(rows), 0, 0, 0, 0,
            0x00FF0000, 0x0000FF00, 0x000000FF, 0xFF000000,
            0x73524742,                      # 'sRGB'
            *([0] * 9), 0, 0, 0,
            4, 0, 0, 0)                      # LCS_GM_IMAGES
        return header + rows

    def setClipboardImage(self, img):
        try:
            import win32clipboard, win32con
        except ImportError as exc:
            raise PdnError("pywin32 is required for clipboard drawing "
                           "(pip install pywin32).") from exc
        data = self._toDIBV5(img)
        # The clipboard is a single global lock. Any other process holding it --
        # a clipboard manager, a browser, another automation run -- makes
        # OpenClipboard fail, and it must be retried rather than waited on.
        deadline = time.time() + TIMEOUT
        while True:
            try:
                win32clipboard.OpenClipboard()
                break
            except Exception as exc:
                if time.time() >= deadline:
                    raise PdnError(f"Could not open the clipboard within "
                                   f"{TIMEOUT}s; another process is holding it: "
                                   f"{exc}") from exc
                time.sleep(SETTLE)
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_DIBV5, data)
        finally:
            win32clipboard.CloseClipboard()
        return len(data)

    def _blankOverlay(self):
        dims = self._imageDims()
        return Image.new("RGBA", (dims["width"], dims["height"]), (0, 0, 0, 0))

    def pasteOverlay(self, img, merge=True):
        """Composite a full-size RGBA overlay onto the image.

        The overlay must match the image dimensions exactly: a same-size paste
        lands at (0, 0) unambiguously, and Paint.NET never asks about resizing
        the canvas.

        Paint.NET's paste REPLACES the target layer's pixels rather than
        blending into them -- transparent parts of the overlay would erase what
        is underneath. So the paste goes onto a new layer, and merging that
        down is what actually composites. Set merge=False to keep the drawing
        on its own layer.
        """
        dims = self._imageDims()
        if img.size != (dims["width"], dims["height"]):
            raise PdnError(f"Overlay is {img.size}, image is "
                           f"{(dims['width'], dims['height'])}; sizes must match.")
        if not img.getbbox() or img.getchannel("A").getextrema()[1] == 0:
            raise PdnError("Overlay is entirely transparent; it would change "
                           "nothing. Refusing rather than reporting success.")

        self.focus()
        self._closeOpenMenus()
        win = self._main()

        # The pixel-hash oracle is useless here: Ctrl+V/Ctrl+D flash a selection
        # outline, so the hash differs even for a no-op paste. But the exact
        # result is known in advance, so compare against it instead.
        # "100% zoom" does not by itself mean one screen pixel per image pixel:
        # display scaling puts a factor in between. Calibration already measures
        # the real mapping, so check that instead of trusting the zoom label.
        mapping = self.calibrate()
        scale_off = max(abs(abs(mapping["mx"]) - 1.0), abs(abs(mapping["my"]) - 1.0))
        verifiable = self._zoom() == "100%" and scale_off <= 0.001
        baseline = self._grabCanvas() if verifiable else None

        self.setClipboardImage(img)
        self._hotkey("^d")           # drop any selection so paste lands at 0,0
        time.sleep(0.2)
        # Ctrl+Shift+V pastes straight into a NEW layer, replacing an
        # add-then-paste pair and the window between them where the paste could
        # land on the wrong layer.
        self._hotkey("^+v")
        time.sleep(0.6)
        win.type_keys("^d")          # deselect commits the floating paste
        time.sleep(0.4)
        if merge:
            self._hotkey(self.LAYER_KEYS["merge_down"])
            time.sleep(0.5)

        result = {"pasted": list(img.size), "merged": merge,
                  "dirty": self.isDirty()}
        if not verifiable:
            result.update(
                verified=False,
                reason=f"screen pixels are not 1:1 with image pixels "
                       f"(zoom={self._zoom()}, measured scale "
                       f"mx={mapping['mx']:.4f} my={mapping['my']:.4f}); the "
                       f"grab cannot be compared pixel for pixel. Call "
                       f"zoom_actual_size, and check display scaling if the "
                       f"measured scale is still not 1.0")
            return result

        expected = Image.alpha_composite(baseline.convert("RGBA"), img).convert("RGB")
        result.update(self._compareCanvas(expected))
        if not result["verified"]:
            raise PdnError(
                f"Paste did not produce the expected pixels: "
                f"{result['mismatch_percent']}% of the image differs "
                f"(mean channel error {result['mean_error']}).")
        return result

    def _grabCanvas(self):
        """RGB image of the image region. Only 1:1 with image pixels at 100%."""
        self._parkCursor()
        box = self._imageScreenBox()
        if box[2] <= box[0] or box[3] <= box[1]:
            raise PdnError(f"Image is not visible on screen: box={box}.")
        return ImageGrab.grab(bbox=box, all_screens=True).convert("RGB")

    def _compareCanvas(self, expected, tolerance=16, allowed_percent=0.5):
        """Compare what is on screen against what was predicted.

        A small mismatch budget is allowed: the screen grab goes through the
        display path, so colour management or subpixel rendering can shift a
        channel slightly without anything being wrong.
        """
        import numpy
        actual = self._grabCanvas()
        if actual.size != expected.size:
            raise PdnError(f"Grabbed {actual.size}, expected {expected.size}.")
        a = numpy.asarray(actual, dtype=numpy.int16)
        e = numpy.asarray(expected, dtype=numpy.int16)
        delta = numpy.abs(a - e)
        differing = (delta.max(axis=2) > tolerance)
        percent = round(float(differing.mean()) * 100, 3)
        return {"verified": percent <= allowed_percent,
                "mismatch_percent": percent,
                "mean_error": round(float(delta.mean()), 3)}

    @staticmethod
    def _rgba(colour, alpha=255):
        value = str(colour).lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", value):
            raise PdnError(f"Expected RRGGBB hex, got {colour!r}.")
        return (int(value[0:2], 16), int(value[2:4], 16),
                int(value[4:6], 16), alpha)

    def _drawOverlay(self, render):
        from PIL import ImageDraw
        overlay = self._blankOverlay()
        render(ImageDraw.Draw(overlay))
        return self.pasteOverlay(overlay)

    def drawLineExact(self, x1, y1, x2, y2, color="000000", width=1):
        return self._drawOverlay(lambda d: d.line(
            [(x1, y1), (x2, y2)], fill=self._rgba(color), width=width))

    def drawRectExact(self, x1, y1, x2, y2, color="000000", width=1, fill=None):
        """Square corners, unlike the brush-dragged version, which the
        Catmull-Rom 'Smoothed path' setting rounds off."""
        return self._drawOverlay(lambda d: d.rectangle(
            [(x1, y1), (x2, y2)], outline=self._rgba(color), width=width,
            fill=self._rgba(fill) if fill else None))

    def drawEllipseExact(self, x1, y1, x2, y2, color="000000", width=1, fill=None):
        return self._drawOverlay(lambda d: d.ellipse(
            [(x1, y1), (x2, y2)], outline=self._rgba(color), width=width,
            fill=self._rgba(fill) if fill else None))

    def drawPolygonExact(self, points, color="000000", width=1, fill=None):
        pts = [(int(p[0]), int(p[1])) for p in points]
        if len(pts) < 3:
            raise PdnError("A polygon needs at least 3 points.")
        return self._drawOverlay(lambda d: d.polygon(
            pts, outline=self._rgba(color), fill=self._rgba(fill) if fill else None))

    def drawPolylineExact(self, points, color="000000", width=1):
        pts = [(int(p[0]), int(p[1])) for p in points]
        if len(pts) < 2:
            raise PdnError("A polyline needs at least 2 points.")
        return self._drawOverlay(lambda d: d.line(
            pts, fill=self._rgba(color), width=width, joint="curve"))

    def pasteImageFile(self, path, x=0, y=0):
        """Composite an image file onto the layer at an exact image-space point."""
        source = Image.open(path).convert("RGBA")
        overlay = self._blankOverlay()
        overlay.alpha_composite(source, (int(x), int(y)))
        result = self.pasteOverlay(overlay)
        result["source"] = {"path": str(path), "size": list(source.size),
                            "at": [int(x), int(y)]}
        return result

    # -------------------------------------------------------------- files

    def _dialogFieldFor(self, dlg, label_auto_id):
        """The Edit belonging to a label, matched by vertical alignment.

        Every Edit in the New Image dialog has an EMPTY auto_id, and they are
        not in visual order: taking the first two would land on the CENTIMETRE
        width and height rather than the pixel ones, and silently create an
        image of the wrong size. The labels do have ids, so each field is found
        by the label beside it.
        """
        wrapper = dlg.wrapper_object()
        label = None
        edits = []
        for child in wrapper.descendants():
            info = child.element_info
            if info.automation_id == label_auto_id:
                label = child
            elif info.control_type == "Edit":
                edits.append(child)
        if label is None:
            raise PdnError(f"Label {label_auto_id!r} not found in the dialog.")
        if not edits:
            raise PdnError("No Edit controls in the dialog.")

        label_rect = label.rectangle()
        label_mid = (label_rect.top + label_rect.bottom) / 2
        best, best_gap = None, None
        for edit in edits:
            rect = edit.rectangle()
            if rect.left < label_rect.left:
                continue                      # field sits to the right of its label
            gap = abs((rect.top + rect.bottom) / 2 - label_mid)
            if best_gap is None or gap < best_gap:
                best, best_gap = edit, gap
        if best is None or best_gap > 20:
            raise PdnError(f"No Edit aligned with {label_auto_id!r} "
                           f"(closest was {best_gap}px away).")
        return best

    def newImage(self, width, height):
        """Create a new image of an exact pixel size.

        The dialog also carries centimetre and resolution fields, and a
        "Maintain aspect ratio" toggle that rewrites one dimension when the
        other changes. Both values are therefore typed, read back, and retyped
        until they agree before OK is pressed.
        """
        width, height = int(width), int(height)
        if width < 1 or height < 1:
            raise PdnError(f"Size must be positive, got {width}x{height}.")

        stray = self.openDialogs()
        if stray:
            raise PdnError(f"A dialog is already open ({stray}); call recover().")

        self._hotkey("^n")
        dlg = self.app.window(title="New", top_level_only=False)
        try:
            dlg.wait("exists ready visible", timeout=10)
        except Exception as exc:
            raise PdnError(f"New Image dialog did not appear: {exc}") from exc

        try:
            w_box = self._dialogFieldFor(dlg, "newWidthLabel1")
            h_box = self._dialogFieldFor(dlg, "newHeightLabel1")

            for attempt in range(4):
                w_box.set_edit_text(str(width))
                time.sleep(0.15)
                h_box.set_edit_text(str(height))
                time.sleep(0.25)
                got_w, got_h = w_box.get_value().strip(), h_box.get_value().strip()
                if (got_w, got_h) == (str(width), str(height)):
                    break
                if attempt == 0:
                    # Most likely the aspect-ratio lock rewriting the other
                    # field. It is a Button with no readable checked state, so
                    # clicking it once and retrying is the only way to tell.
                    try:
                        dlg.child_window(auto_id="constrainCheckBox",
                                         control_type="Button").click_input()
                        time.sleep(0.2)
                    except Exception:
                        pass
            else:
                raise PdnError(f"New Image dialog would not hold "
                               f"{width}x{height}; it reads {got_w}x{got_h}. "
                               f"The aspect-ratio lock may be active.")

            ok = dlg.child_window(auto_id="okButton", control_type="Button")
            ok.click_input()
        except Exception:
            self._closeDialog(dlg, "New")
            raise

        deadline = time.time() + 10
        while time.time() < deadline:
            if not dlg.exists(timeout=1):
                break
            time.sleep(SETTLE)
        else:
            raise PdnError("New Image dialog did not close after OK.")

        # A new document means new controls and a new coordinate mapping.
        self._invalidateElementCache()
        self.calib_cache = None
        time.sleep(0.5)

        deadline = time.time() + TIMEOUT
        while time.time() < deadline:
            dims = self._imageDims()
            if dims == {"width": width, "height": height}:
                return {"image_dims": dims, "document": self.getActiveDocName(),
                        "dirty": self.isDirty()}
            time.sleep(SETTLE)
        raise PdnError(f"New image is {self._imageDims()}, asked for "
                       f"{width}x{height}.")

    # ------------------------------------------------------------ file I/O

    # Save As / Open are standard Win32 common dialogs (class #32770). Most of
    # their auto_ids are useless -- the list view is full of controls all named
    # "System.ItemNameDisplay" -- but the file name Edit is "1001" (inside the
    # FileNameControlHost combo; NOT the classic 1148 of the legacy dialog) and
    # the accept button is "1". Both are scoped to the dialog, so the "1001"
    # shared with the brush-size box cannot collide.
    # Save As and Open are NOT the same dialog implementation: Save As uses the
    # modern FileNameControlHost whose Edit is "1001", while Open uses the
    # classic common-dialog ids where it is "1148". Both are tried.
    FILE_DIALOG_NAME_IDS = ("1001", "1148")
    FILE_DIALOG_ACCEPT_ID = "1"

    def _fileDialog(self, title, timeout=15):
        dlg = self.app.window(title=title, top_level_only=False, class_name="#32770")
        try:
            dlg.wait("exists ready visible", timeout=timeout)
        except Exception as exc:
            raise PdnError(f"{title} dialog did not appear: {exc}") from exc
        return dlg

    def _submitFileDialog(self, dlg, path):
        """Put a path in the file name box and accept it.

        The box is typed into rather than keystroked blind: a path contains
        characters that type_keys treats as syntax, and a mistyped path in a
        Save dialog writes to the wrong file rather than failing.
        """
        box = None
        for auto_id in self.FILE_DIALOG_NAME_IDS:
            candidate = dlg.child_window(auto_id=auto_id, control_type="Edit")
            if candidate.exists(timeout=2):
                box = candidate.wrapper_object()
                break
        if box is None:
            raise PdnError(f"File name box not found; tried control ids "
                           f"{list(self.FILE_DIALOG_NAME_IDS)}.")
        box.set_edit_text(str(path))
        time.sleep(0.2)
        # Compare normalised paths: the dialog may quote the value, and case
        # and separator differences are not disagreements on Windows.
        typed = os.path.normcase(os.path.normpath(
            box.get_value().strip().strip('"')))
        wanted = os.path.normcase(os.path.normpath(str(path)))
        if typed != wanted:
            raise PdnError(f"File name box holds {box.get_value()!r}, "
                           f"not {str(path)!r}.")
        accept = dlg.child_window(auto_id=self.FILE_DIALOG_ACCEPT_ID,
                                  control_type="Button")
        if accept.exists(timeout=2):
            accept.wrapper_object().click_input()
        else:
            box.type_keys("{ENTER}")

    # Buttons that mean "proceed", most specific first. baseOkButton is the id
    # Paint.NET gives the OK button on its own dialogs; the rest are for the
    # prompts that come from elsewhere.
    PROCEED_BUTTONS = ("baseOkButton", "okButton")
    PROCEED_LABELS = ("Flatten", "OK", "Yes", "Save")

    def _clearFollowUpDialogs(self, budget=25.0, done=None):
        """Dismiss the dialogs a save can raise, accepting their defaults.

        The format settings dialog differs per file type -- PNG offers bit depth
        and dithering, JPEG a quality slider -- so nothing inside it is touched;
        only its accept button is pressed. A flatten prompt may follow when the
        format cannot hold layers.

        Buttons are collected in one pass per dialog. Probing candidate titles
        one at a time costs a timeout per miss, which previously consumed the
        whole budget and returned as though everything had been handled.
        """
        handled, deadline = [], time.time() + budget
        while time.time() < deadline:
            # Wait for the WORK to finish, not for a momentary absence of
            # dialogs: the settings dialog takes a beat to appear, and checking
            # once immediately after clicking Save finds nothing and returns as
            # though everything had been dealt with.
            if done is not None and done():
                break
            titles = [t for t in self.openDialogs() if t not in ("Save As", "Open")]
            if not titles:
                if done is None:
                    break
                time.sleep(0.3)
                continue
            for title in titles:
                dlg = self.app.window(title=title, top_level_only=False)
                if not dlg.exists(timeout=1):
                    continue
                buttons = [child for child in dlg.wrapper_object().descendants()
                           if child.element_info.control_type == "Button"]
                chosen = None
                for auto_id in self.PROCEED_BUTTONS:
                    chosen = next((b for b in buttons
                                   if b.element_info.automation_id == auto_id), None)
                    if chosen is not None:
                        break
                if chosen is None:
                    for label in self.PROCEED_LABELS:
                        chosen = next((b for b in buttons
                                       if b.window_text().strip() == label), None)
                        if chosen is not None:
                            break
                if chosen is None:
                    raise PdnError(
                        f"Dialog {title!r} appeared after saving and has no "
                        f"recognised accept button. Buttons: "
                        f"{[b.window_text() for b in buttons]}. Call recover().")
                chosen.click_input()
                handled.append(f"{title}:{chosen.window_text().strip() or 'OK'}")
                time.sleep(0.6)
        return handled

    def saveAs(self, path):
        """Save the document to an exact path, then prove the file is right.

        Verified by reading the file back off disk and comparing its dimensions
        to the document -- a save dialog that silently went to a different name
        or format would otherwise look like success.
        """
        target = pathlib.Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        if not target.suffix:
            raise PdnError(f"{target} has no extension; Paint.NET picks the "
                           f"format from it.")
        expected = self._imageDims()
        before_mtime = target.stat().st_mtime if target.exists() else None

        stray = self.openDialogs()
        if stray:
            raise PdnError(f"A dialog is already open ({stray}); call recover().")

        self._hotkey("^+s")
        dlg = self._fileDialog("Save As")
        try:
            self._submitFileDialog(dlg, target)
        except Exception:
            self._closeDialog(dlg, "Save As")
            raise
        def written():
            return target.exists() and target.stat().st_mtime != before_mtime

        followed = self._clearFollowUpDialogs(done=written)
        if not written():
            raise PdnError(f"{target} was not written. Dialogs handled: "
                           f"{followed}; still open: {self.openDialogs()}.")

        with Image.open(target) as saved:
            size = saved.size
        result = {"path": str(target), "size": list(size),
                  "dialogs_handled": followed, "dirty": self.isDirty(),
                  "verified": list(size) == [expected["width"], expected["height"]]}
        if not result["verified"]:
            result["reason"] = (f"file is {size} but the document is "
                                f"{expected['width']}x{expected['height']}")
        return result

    def openFile(self, path):
        """Open an image file and confirm the document matches it."""
        target = pathlib.Path(path)
        if not target.exists():
            raise PdnError(f"{target} does not exist.")
        with Image.open(target) as source:
            expected = list(source.size)

        stray = self.openDialogs()
        if stray:
            raise PdnError(f"A dialog is already open ({stray}); call recover().")

        self._hotkey("^o")
        dlg = self._fileDialog("Open")
        try:
            self._submitFileDialog(dlg, target)
        except Exception:
            self._closeDialog(dlg, "Open")
            raise
        def opened():
            try:
                dims = self._imageDims()
            except PdnError:
                return False
            return [dims["width"], dims["height"]] == expected

        self._clearFollowUpDialogs(done=opened)

        # A different document means different controls and a new mapping.
        self._invalidateElementCache()
        self.calib_cache = None

        deadline = time.time() + 20
        while time.time() < deadline:
            try:
                dims = self._imageDims()
            except PdnError:
                time.sleep(0.3)
                continue
            if [dims["width"], dims["height"]] == expected:
                return {"path": str(target), "image_dims": dims,
                        "document": self.getActiveDocName(), "verified": True}
            time.sleep(0.3)
        raise PdnError(f"Opened document is {self._imageDims()}, but "
                       f"{target.name} is {expected[0]}x{expected[1]}.")

    def readCanvasExact(self, path=None):
        """Full-resolution pixels, via a file rather than the screen.

        screenshotCanvas grabs the display, so it only tells the truth while the
        image is at 100% zoom, unscrolled and entirely on screen. Saving to PNG
        and reading that back is exact at any zoom, for any image size, and
        includes layers composited properly.
        """
        temporary = path is None
        target = pathlib.Path(path) if path else pathlib.Path(
            os.environ.get("TEMP", ".")) / f"pdn_read_{os.getpid()}.png"
        saved = self.saveAs(target)
        with Image.open(target) as image:
            data = image.convert("RGBA")
            size = data.size
        if temporary:
            try:
                target.unlink()
            except OSError:
                pass
        return {"size": list(size), "path": None if temporary else str(target),
                "verified": saved["verified"]}


pdnmcp = Paint_Net_MCP()

_TOOLS = [
    (pdnmcp.start, "start_paint", "Launch a new Paint.NET instance"),
    (pdnmcp.attachOrLaunch, "attach_or_launch",
     "Attach to a running Paint.NET, launching one only if none is found"),
    (pdnmcp.focus, "focus", "Bring the Paint.NET window to the foreground"),
    (pdnmcp.getControls, "get_controls",
     "Dump the full UIA control tree for debugging"),
    (pdnmcp.dumpDialog, "dump_dialog",
     "Dump the UIA tree of a dialog matched by title regex"),
    (pdnmcp.getState, "get_state",
     "Window/canvas rects, image dims, zoom, active tool, colour, brush, dirty flag"),
    (pdnmcp.getCanvasDimensions, "get_canvas_dimensions",
     "Image width and height in pixels"),
    (pdnmcp.calibrate, "calibrate",
     "Solve and cache the screen<->image coordinate mapping"),
    (pdnmcp.screenToImage, "screen_to_image", "Convert screen point to image point"),
    (pdnmcp.imageToScreen, "image_to_screen", "Convert image point to screen point"),
    (pdnmcp.selectTool, "select_tool", "Select a tool by name or index"),
    (pdnmcp.selectToolByKey, "select_tool_by_key",
     "Select a tool via its keyboard shortcut (no Tools panel needed)"),
    (pdnmcp.getActiveToolIndex, "get_active_tool_index",
     "Index of the currently selected tool"),
    (pdnmcp.getPrimaryColor, "get_primary_color", "Primary colour as RRGGBB hex"),
    (pdnmcp.setPrimaryColor, "set_primary_color", "Set primary colour from RRGGBB hex"),
    (pdnmcp.getBrushSize, "get_brush_size", "Current brush size in pixels"),
    (pdnmcp.setBrushSize, "set_brush_size", "Set brush size in pixels"),
    (pdnmcp.setHardness, "set_hardness",
     "Set brush hardness percent, verified from the slider fill"),
    (pdnmcp.setSpacing, "set_spacing",
     "Set brush spacing percent, verified from the slider fill"),
    (pdnmcp.getHardness, "get_hardness", "Brush hardness percent"),
    (pdnmcp.getSpacing, "get_spacing", "Brush spacing percent"),
    (pdnmcp.drawPolyline, "draw_polyline", "Drag through image-space points"),
    (pdnmcp.drawLine, "draw_line", "Draw a line between two image points"),
    (pdnmcp.drawRect, "draw_rect", "Draw a rectangle outline"),
    (pdnmcp.drawEllipse, "draw_ellipse", "Draw an ellipse outline"),
    (pdnmcp.drawShape, "draw_shape", "Draw rectangle, ellipse or line from start/stop"),
    (pdnmcp.floodFill, "flood_fill", "Paint-bucket fill at an image point"),
    (pdnmcp.drawLineExact, "draw_line_exact",
     "Exact line via clipboard paste (no smoothing, no calibration)"),
    (pdnmcp.drawRectExact, "draw_rect_exact",
     "Exact rectangle, square corners, optional fill"),
    (pdnmcp.drawEllipseExact, "draw_ellipse_exact",
     "Exact ellipse, optional fill"),
    (pdnmcp.drawPolygonExact, "draw_polygon_exact", "Exact closed polygon"),
    (pdnmcp.drawPolylineExact, "draw_polyline_exact", "Exact open polyline"),
    (pdnmcp.pasteImageFile, "paste_image_file",
     "Composite an image file onto the layer at an exact point"),
    (pdnmcp.addText, "add_text", "Type text at an image point"),
    (pdnmcp.undo, "undo", "Undo via the History panel"),
    (pdnmcp.redo, "redo", "Redo via the History panel"),
    (pdnmcp.listMenuItems, "list_menu_items",
     "Open a top-level menu and list its items"),
    (pdnmcp.recover, "recover",
     "Close leaked dialogs/menus that block canvas input, and drop caches"),
    (pdnmcp.openDialogs, "open_dialogs",
     "List modal dialogs currently open (they block canvas input)"),
    (pdnmcp.toggleLayerVisibility, "toggle_layer_visibility",
     "Flip the active layer's visibility (unverifiable)"),
    (pdnmcp.getLayerProperties, "get_layer_properties",
     "Active layer's name, opacity, blend mode and visibility (via F4)"),
    (pdnmcp.setLayerProperties, "set_layer_properties",
     "Set active layer's name/opacity/blend mode/visibility, verified"),
    (pdnmcp.layerAdd, "layer_add", "Add a new layer via the Layers menu"),
    (pdnmcp.layerDelete, "layer_delete", "Delete the current layer"),
    (pdnmcp.layerDuplicate, "layer_duplicate", "Duplicate the current layer"),
    (pdnmcp.layerCount, "layer_count",
     "Number of layers, inferred from the Layers list scroll extent"),
    (pdnmcp.zoomActualSize, "zoom_actual_size",
     "Ctrl+0: set zoom to 100% (required for pixel verification)"),
    (pdnmcp.flatten, "flatten", "Ctrl+Shift+F: flatten all layers"),
    (pdnmcp.toggleLayerVisibilityFast, "toggle_layer_visibility_fast",
     "Ctrl+, : toggle layer visibility without opening a dialog"),
    (pdnmcp.clearLayerContents, "clear_layer_contents",
     "Ctrl+A then Delete: empty the active layer, keeping it"),
    (pdnmcp.layerAddViaMenu, "layer_add_via_menu",
     "Add a layer through the Layers menu instead of the shortcut"),
    (pdnmcp.layerMergeDown, "layer_merge_down", "Merge the current layer down"),
    (pdnmcp.screenshotCanvas, "screenshot_canvas",
     "PNG of the image region, as base64 or saved to a path"),
    (pdnmcp.newImage, "new_image", "Create a new image of the given size"),
    (pdnmcp.saveAs, "save_as", "Save the document to a path, verified against disk"),
    (pdnmcp.openFile, "open_file", "Open an image file, verified against its size"),
    (pdnmcp.readCanvasExact, "read_canvas_exact",
     "Full-resolution pixels via a temp file, valid at any zoom"),
]

for _fn, _name, _description in _TOOLS:
    mcp.tool(_fn, name=_name, description=_description)


if __name__ == "__main__":
    mcp.run()
