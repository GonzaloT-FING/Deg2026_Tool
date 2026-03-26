from __future__ import annotations

from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk


@dataclass(frozen=True)
class AppTheme:
    name: str
    window_bg: str
    surface: str
    surface_alt: str
    border: str
    text: str
    muted_text: str
    accent: str
    accent_active: str
    accent_text: str
    input_bg: str
    input_fg: str
    input_border: str
    status_bg: str
    status_fg: str
    plot_viewport_bg: str
    plot_viewport_fg: str
    plot_viewport_muted: str


THEMES: dict[str, AppTheme] = {
    "light": AppTheme(
        name="light",
        window_bg="#eef3f7",
        surface="#ffffff",
        surface_alt="#e4edf5",
        border="#cfd9e3",
        text="#16202a",
        muted_text="#5b6b79",
        accent="#0f6cbd",
        accent_active="#0b5a9e",
        accent_text="#ffffff",
        input_bg="#f8fbfd",
        input_fg="#16202a",
        input_border="#b8c7d6",
        status_bg="#dceffc",
        status_fg="#0f3b66",
        plot_viewport_bg="#ffffff",
        plot_viewport_fg="#18222c",
        plot_viewport_muted="#5f6f7d",
    ),
    "dark": AppTheme(
        name="dark",
        window_bg="#10161d",
        surface="#17212b",
        surface_alt="#22303d",
        border="#304050",
        text="#e8eef5",
        muted_text="#9fb1c3",
        accent="#57a7ff",
        accent_active="#3f8fe3",
        accent_text="#09121a",
        input_bg="#111b24",
        input_fg="#e8eef5",
        input_border="#3a5062",
        status_bg="#17314a",
        status_fg="#d6ebff",
        plot_viewport_bg="#ffffff",
        plot_viewport_fg="#18222c",
        plot_viewport_muted="#5f6f7d",
    ),
}

THEME_LABELS = {
    "light": "Light",
    "dark": "Dark",
}

THEME_NAME_BY_LABEL = {label: name for name, label in THEME_LABELS.items()}


def get_theme(theme_name: str) -> AppTheme:
    return THEMES.get(theme_name, THEMES["light"])


def apply_theme(window: tk.Misc, style: ttk.Style, theme_name: str) -> AppTheme:
    theme = get_theme(theme_name)
    style.theme_use("clam")
    window.configure(bg=theme.window_bg)

    base_font = ("Segoe UI", 10)
    small_font = ("Segoe UI", 9)
    section_font = ("Segoe UI Semibold", 10)
    card_title_font = ("Segoe UI Semibold", 13)
    value_font = ("Segoe UI Semibold", 18)
    hero_title_font = ("Segoe UI Semibold", 24)
    hero_subtitle_font = ("Segoe UI", 10)
    viewport_title_font = ("Segoe UI Semibold", 12)

    style.configure(".", background=theme.window_bg, foreground=theme.text, font=base_font)
    style.configure("TFrame", background=theme.window_bg)
    style.configure("App.TFrame", background=theme.window_bg)
    style.configure("Card.TFrame", background=theme.surface)
    style.configure("Panel.TFrame", background=theme.surface_alt)
    style.configure("Status.TFrame", background=theme.status_bg)
    style.configure("PlotViewport.TFrame", background=theme.plot_viewport_bg)

    style.configure("Window.TLabel", background=theme.window_bg, foreground=theme.text, font=base_font)
    style.configure("WindowMuted.TLabel", background=theme.window_bg, foreground=theme.muted_text, font=small_font)
    style.configure("HeroTitle.TLabel", background=theme.window_bg, foreground=theme.text, font=hero_title_font)
    style.configure(
        "HeroSubtitle.TLabel",
        background=theme.window_bg,
        foreground=theme.muted_text,
        font=hero_subtitle_font,
    )
    style.configure("CardTitle.TLabel", background=theme.surface, foreground=theme.text, font=card_title_font)
    style.configure("CardValue.TLabel", background=theme.surface, foreground=theme.text, font=value_font)
    style.configure("CardBody.TLabel", background=theme.surface, foreground=theme.text, font=base_font)
    style.configure("CardMuted.TLabel", background=theme.surface, foreground=theme.muted_text, font=small_font)
    style.configure("FieldLabel.TLabel", background=theme.surface, foreground=theme.muted_text, font=section_font)
    style.configure("PanelLabel.TLabel", background=theme.surface_alt, foreground=theme.text, font=section_font)
    style.configure("Status.TLabel", background=theme.status_bg, foreground=theme.status_fg, font=base_font)
    style.configure(
        "PlotViewportTitle.TLabel",
        background=theme.plot_viewport_bg,
        foreground=theme.plot_viewport_fg,
        font=viewport_title_font,
    )
    style.configure(
        "PlotViewportBody.TLabel",
        background=theme.plot_viewport_bg,
        foreground=theme.plot_viewport_muted,
        font=small_font,
    )

    style.configure(
        "Accent.TButton",
        background=theme.accent,
        foreground=theme.accent_text,
        borderwidth=0,
        padding=(16, 10),
        focusthickness=0,
    )
    style.map(
        "Accent.TButton",
        background=[("pressed", theme.accent_active), ("active", theme.accent_active)],
        foreground=[("disabled", theme.muted_text)],
    )

    style.configure(
        "Subtle.TButton",
        background=theme.surface_alt,
        foreground=theme.text,
        borderwidth=0,
        padding=(14, 9),
        focusthickness=0,
    )
    style.map(
        "Subtle.TButton",
        background=[("pressed", theme.surface), ("active", theme.surface)],
        foreground=[("disabled", theme.muted_text)],
    )

    style.configure(
        "App.TEntry",
        foreground=theme.input_fg,
        fieldbackground=theme.input_bg,
        bordercolor=theme.input_border,
        lightcolor=theme.input_border,
        darkcolor=theme.input_border,
        padding=(10, 8),
    )
    style.map(
        "App.TEntry",
        bordercolor=[("focus", theme.accent)],
        lightcolor=[("focus", theme.accent)],
        darkcolor=[("focus", theme.accent)],
    )

    style.configure(
        "App.TCombobox",
        foreground=theme.input_fg,
        fieldbackground=theme.input_bg,
        background=theme.surface_alt,
        bordercolor=theme.input_border,
        lightcolor=theme.input_border,
        darkcolor=theme.input_border,
        arrowcolor=theme.text,
        padding=(8, 7),
    )
    style.map(
        "App.TCombobox",
        fieldbackground=[("readonly", theme.input_bg)],
        foreground=[("readonly", theme.input_fg)],
        background=[("readonly", theme.surface_alt), ("active", theme.surface)],
        bordercolor=[("focus", theme.accent)],
        lightcolor=[("focus", theme.accent)],
        darkcolor=[("focus", theme.accent)],
        selectbackground=[("readonly", theme.accent)],
        selectforeground=[("readonly", theme.accent_text)],
    )

    style.configure("Card.TCheckbutton", background=theme.surface, foreground=theme.text, font=base_font)
    style.map(
        "Card.TCheckbutton",
        background=[("active", theme.surface)],
        foreground=[("disabled", theme.muted_text)],
    )

    style.configure("TSeparator", background=theme.border)
    return theme
