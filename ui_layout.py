from __future__ import annotations

import tkinter as tk
from tkinter import ttk


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
