"""Where LinBASIC Studio keeps its read-only data, the user's settings, and where the
compiler and the runtime are (based on paths.py of Linpad++)."""
import os
import shutil

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, "data")
RES_DIR = os.path.join(APP_DIR, "res")
ICON_DIR = os.path.join(RES_DIR, "icons")

# the source tree: ide/lbstudio -> ide -> <root> with compiler/, runtime/, data/vb6/
ROOT = os.path.normpath(os.path.join(APP_DIR, "..", ".."))
COMPILER = os.path.join(ROOT, "compiler")
RUNTIME = os.path.join(ROOT, "runtime")
VB6_DATA = os.path.join(ROOT, "data", "vb6")

_config_override = None
NO_WINDOW = {"creationflags": 0x08000000} if os.name == "nt" else {}   # CREATE_NO_WINDOW: the tools of the IDE


def host_target():
    """what Run and Make build by default: linux, or win32 / win64 on Windows (the bitness of the IDE)"""
    if os.name != "nt":
        return "linux"
    import struct
    return "win64" if struct.calcsize("P") == 8 else "win32"


def set_config_dir(path):
    global _config_override
    _config_override = os.path.abspath(os.path.expanduser(path))


def portable():
    """the portable package of Windows (packaging/make-win.py writes portable.ini): the settings stay
    in its folder, so the folder can be carried on a USB stick"""
    return os.path.exists(os.path.join(ROOT, "portable.ini"))


def default_config_dir():
    if portable():
        return os.path.join(ROOT, "settings")
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "linbasic-studio")


def cache_dir(*parts):
    if portable():
        base = os.path.join(ROOT, "settings", "cache")
    elif os.name == "nt":
        base = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "linbasic-studio", "cache")
    else:
        base = os.path.join(os.environ.get("XDG_CACHE_HOME") or os.path.join(os.path.expanduser("~"), ".cache"),
                            "linbasic-studio")
    return os.path.join(base, *parts)


def config_dir():
    if _config_override:
        d = _config_override
    else:
        d = default_config_dir()
    os.makedirs(d, exist_ok=True)
    return d


def cfg(*parts):
    return os.path.join(config_dir(), *parts)


def data(*parts):
    return os.path.join(DATA_DIR, *parts)


def icon(name):
    return os.path.join(ICON_DIR, name)


def ensure_user_files():
    d = config_dir()
    for src, dst in (("stylers.model.xml", "stylers.xml"), ("langs.model.xml", "langs.xml")):
        target = os.path.join(d, dst)
        if not os.path.exists(target):
            try:
                shutil.copyfile(data(src), target)
            except OSError:
                pass
    os.makedirs(os.path.join(d, "themes"), exist_ok=True)
    return d


def user_or_model(user_name, model_name):
    p = cfg(user_name)
    return p if os.path.exists(p) else data(model_name)


def theme_dirs():
    return [data("themes"), cfg("themes")]


def list_themes():
    out = [("Default (stylers.xml)", user_or_model("stylers.xml", "stylers.model.xml"))]
    seen = set()
    for d in theme_dirs():
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d), key=str.lower):
            if f.lower().endswith(".xml") and f not in seen:
                seen.add(f)
                out.append((f[:-4], os.path.join(d, f)))
    final = {}
    for name, path in out:
        final[name] = path
    return list(final.items())
