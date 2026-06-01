from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def create_scrollable_controls(
    parent: tk.Widget,
    *,
    outer_padding: int | tuple[int, ...] = 10,
    inner_padding: int | tuple[int, ...] = (0, 0, 6, 0),
    pack: bool = True,
    fixed_width: int | None = None,
    reset_y_on_resize: bool = False,
) -> tuple[ttk.Frame, ttk.Frame]:
    outer_kwargs = {}
    if fixed_width is not None:
        outer_kwargs["width"] = fixed_width

    outer = ttk.Frame(parent, padding=outer_padding, **outer_kwargs)
    if fixed_width is not None:
        outer.pack_propagate(False)
        outer.grid_propagate(False)
    if pack:
        outer.pack(fill="both", expand=True)

    outer.columnconfigure(0, weight=1, minsize=1)
    outer.columnconfigure(1, weight=0)
    outer.rowconfigure(0, weight=1)

    style = ttk.Style(parent)
    canvas_bg = style.lookup("App.TFrame", "background") or style.lookup("TFrame", "background") or "#10161d"
    canvas = tk.Canvas(
        outer,
        width=1,
        highlightthickness=0,
        borderwidth=0,
        bg=canvas_bg,
        bd=0,
        relief="flat",
    )
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas, padding=inner_padding)
    window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    def _update_scrollregion(_event=None) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _sync_width(event) -> None:
        canvas.itemconfigure(window_id, width=max(1, event.width))

    inner.bind("<Configure>", _update_scrollregion)
    canvas.bind("<Configure>", _sync_width)

    def _on_mousewheel(event) -> str:
        if getattr(event, "delta", 0):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            step = -1 if getattr(event, "num", None) == 4 else 1
            canvas.yview_scroll(step, "units")
        return "break"

    def _bind_mousewheel_tree(widget: tk.Widget) -> None:
        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(sequence, _on_mousewheel)

        for child in widget.winfo_children():
            _bind_mousewheel_tree(child)

    def _refresh_mousewheel_bindings(_event=None) -> None:
        for widget in (outer, canvas, scrollbar, inner):
            _bind_mousewheel_tree(widget)

    inner.bind("<Configure>", _refresh_mousewheel_bindings, add="+")
    outer.after_idle(_refresh_mousewheel_bindings)

    if reset_y_on_resize:
        outer.after_idle(lambda: canvas.yview_moveto(0))
        outer.bind("<Configure>", lambda _event: canvas.yview_moveto(0), add="+")

    return outer, inner


def create_resizable_plot_layout(
    parent: tk.Widget,
    *,
    sidebar_width: int = 320,
    min_sidebar_width: int = 220,
    min_plot_width: int = 420,
    sidebar_side: str = "left",
    plot_padding: int | tuple[int, ...] = 10,
) -> tuple[ttk.Frame, ttk.Frame]:
    pane = ttk.Panedwindow(parent, orient="horizontal")
    pane.pack(fill="both", expand=True)

    sidebar_host = ttk.Frame(pane, width=sidebar_width)
    plot_host = ttk.Frame(pane, padding=plot_padding)

    if sidebar_side == "left":
        pane.add(sidebar_host, weight=0)
        pane.add(plot_host, weight=1)
    elif sidebar_side == "right":
        pane.add(plot_host, weight=1)
        pane.add(sidebar_host, weight=0)
    else:
        raise ValueError(f"Unsupported sidebar side: {sidebar_side}")

    def _sidebar_limits(total_width: int) -> tuple[int, int]:
        if sidebar_side == "left":
            min_pos = min_sidebar_width
            max_pos = max(min_sidebar_width, total_width - min_plot_width)
        else:
            min_pos = min_plot_width
            max_pos = max(min_plot_width, total_width - min_sidebar_width)
        return min_pos, max_pos

    def _clamp_sash() -> None:
        if not pane.winfo_exists():
            return

        total_width = pane.winfo_width()
        if total_width <= 1:
            return

        min_pos, max_pos = _sidebar_limits(total_width)
        target = min(max(pane.sashpos(0), min_pos), max_pos)

        try:
            pane.sashpos(0, target)
        except tk.TclError:
            return

    def _set_initial_sash() -> None:
        if not pane.winfo_exists():
            return

        total_width = pane.winfo_width()
        if total_width <= 1:
            pane.after(16, _set_initial_sash)
            return

        min_pos, max_pos = _sidebar_limits(total_width)

        if sidebar_side == "left":
            sash_pos = min(max(min_pos, sidebar_width), max_pos)
        else:
            sash_pos = min(max(min_pos, total_width - sidebar_width), max_pos)

        try:
            pane.sashpos(0, sash_pos)
        except tk.TclError:
            return

    def _schedule_clamp(_event=None) -> None:
        pane.after_idle(_clamp_sash)

    pane.bind("<Configure>", _schedule_clamp, add="+")
    pane.bind("<ButtonRelease-1>", _schedule_clamp, add="+")
    pane.after_idle(_set_initial_sash)
    return sidebar_host, plot_host
