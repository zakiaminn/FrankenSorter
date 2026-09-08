"""
Brand identity for FrankenSorter: Bricolage Grotesque + Martian Mono on the
"Sulfur on Chalk" palette. See brandkit for the source spec — this module
adapts it to a Tk/CustomTkinter runtime (private, per-process font loading;
color tokens as plain dicts instead of CSS variables).
"""
import ctypes
import ctypes.util
import os
import platform

FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "fonts")

BRICOLAGE = "Bricolage Grotesque"
BRICOLAGE_SEMIBOLD = "Bricolage Grotesque SemiBold"
MARTIAN = "Martian Mono"

# Fallback stacks used only if the bundled fonts fail to register.
BRICOLAGE_FALLBACK = "Helvetica Neue"
MARTIAN_FALLBACK = "Menlo" if platform.system() == "Darwin" else "Consolas"

FONT_FILES = [
    "BricolageGrotesque-Regular.ttf",
    "BricolageGrotesque-Bold.ttf",
    "BricolageGrotesque-SemiBold.ttf",
    "MartianMono-Regular.ttf",
    "MartianMono-Bold.ttf",
]


def _register_macos(path):
    cf = ctypes.CDLL(ctypes.util.find_library("CoreFoundation"))
    ct = ctypes.CDLL(ctypes.util.find_library("CoreText"))
    cf.CFStringCreateWithCString.restype = ctypes.c_void_p
    cf.CFStringCreateWithCString.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int32]
    cf.CFURLCreateWithFileSystemPath.restype = ctypes.c_void_p
    cf.CFURLCreateWithFileSystemPath.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int32, ctypes.c_bool]
    ct.CTFontManagerRegisterFontsForURL.restype = ctypes.c_bool
    ct.CTFontManagerRegisterFontsForURL.argtypes = [ctypes.c_void_p, ctypes.c_int32, ctypes.c_void_p]

    kCFStringEncodingUTF8 = 0x08000100
    kCFURLPOSIXPathStyle = 0
    kCTFontManagerScopeProcess = 1  # registered for this process only; never touches Font Book

    cf_path = cf.CFStringCreateWithCString(None, path.encode("utf-8"), kCFStringEncodingUTF8)
    cf_url = cf.CFURLCreateWithFileSystemPath(None, cf_path, kCFURLPOSIXPathStyle, False)
    return bool(ct.CTFontManagerRegisterFontsForURL(cf_url, kCTFontManagerScopeProcess, None))


def _register_windows(path):
    FR_PRIVATE = 0x10
    return bool(ctypes.windll.gdi32.AddFontResourceExW(path, FR_PRIVATE, 0))


def _register_linux(path):
    fc = ctypes.CDLL(ctypes.util.find_library("fontconfig"))
    fc.FcConfigGetCurrent.restype = ctypes.c_void_p
    fc.FcConfigAppFontAddFile.restype = ctypes.c_int
    fc.FcConfigAppFontAddFile.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    config = fc.FcConfigGetCurrent()
    return bool(fc.FcConfigAppFontAddFile(config, path.encode("utf-8")))


def register_fonts():
    """
    Loads the bundled brand fonts privately for this process only — no admin
    rights, no writing into the OS font library. Returns True if every file
    registered; callers should fall back to system fonts on False.
    """
    system = platform.system()
    registrar = {"Darwin": _register_macos, "Windows": _register_windows, "Linux": _register_linux}.get(system)
    if registrar is None:
        return False
    ok = True
    for filename in FONT_FILES:
        path = os.path.join(FONTS_DIR, filename)
        try:
            if not os.path.exists(path) or not registrar(path):
                ok = False
        except Exception:
            ok = False
    return ok


# Populated by init_fonts(); use these constants everywhere else in the app
# so a failed registration transparently degrades to system fonts.
FONT_BODY = BRICOLAGE_FALLBACK
FONT_HEADLINE = BRICOLAGE_FALLBACK
FONT_SUBHEADING = BRICOLAGE_FALLBACK
FONT_MONO = MARTIAN_FALLBACK


def init_fonts():
    """Call once, before any CTkFont is constructed."""
    global FONT_BODY, FONT_HEADLINE, FONT_SUBHEADING, FONT_MONO
    if register_fonts():
        FONT_BODY = BRICOLAGE
        FONT_HEADLINE = BRICOLAGE
        FONT_SUBHEADING = BRICOLAGE_SEMIBOLD
        FONT_MONO = MARTIAN


# --- Color tokens ("Sulfur on Chalk") -----------------------------------
# Fills use `brand`; accent text/links/icons use `brand_ink`; text sitting on
# a brand fill uses `brand_fg`. Green/red (`pos`/`neg`) are reserved for
# actual status/data signals, never decorative UI chrome.

LIGHT = {
    "bg": "#FAFAF9", "surface": "#F1F1EE", "surface_2": "#E9E9E5",
    "ink": "#16160E", "ink_2": "#565448", "ink_3": "#78766A",
    "rule": "#E5E5E1", "rule_2": "#D4D4CF",
    "brand": "#DCEC3A", "brand_ink": "#6F7A00", "brand_fg": "#16160E", "brand_wash": "#EFF0CE",
    "pos": "#2E7D4F", "neg": "#B33A3A", "focus": "#8A9600",
}

DARK = {
    "bg": "#0E0E0A", "surface": "#17170F", "surface_2": "#212118",
    "ink": "#EDEDE0", "ink_2": "#9C9A86", "ink_3": "#86846F",
    "rule": "#29291C", "rule_2": "#3A3A29",
    "brand": "#DCEC3A", "brand_ink": "#DCEC3A", "brand_fg": "#14140A", "brand_wash": "#23260F",
    "pos": "#55B37E", "neg": "#E8635E", "focus": "#DCEC3A",
}

PALETTES = {"dark": DARK, "light": LIGHT}


def spaced(text, gap=" "):
    """Fakes CSS letter-spacing for short uppercase eyebrow labels only —
    Tk has no per-character tracking control."""
    return gap.join(text)
