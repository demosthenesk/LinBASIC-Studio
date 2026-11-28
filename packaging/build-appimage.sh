#!/usr/bin/env bash
# build-appimage.sh - the Linux AppImage of LinBASIC Studio (x86_64). Run on Ubuntu 24.04.
#   packaging/build-appimage.sh            -> build/LinBASIC-Studio-<version>-x86_64.AppImage
# Build deps: sudo apt install python3.12-venv libxcb-cursor0 curl gcc g++ libc6-dev qt6-base-dev
#             (qt6-base-dev: runtime/forms/build.sh; the AppImage itself carries Qt 6 of PyQt6)
# Needs PyPI and GitHub (appimagetool, the type2 runtime, the excludelist, tinycc) unless
# APPIMAGETOOL / RUNTIME / EXCLUDELIST point at local copies and runtime/tcc-linux exists.
#
# The AppImage: Python 3.12 + PyQt6 + QScintilla (the IDE), the compiler lbvc, the runtime
# (libvbrt.a, libvbdb.a, libvbforms.so) and the Tiny C Compiler (runtime/tcc-linux), so that
# the programs are built without gcc on the computer. The licenses: licenses/README.md.
set -euo pipefail

SRC="$(cd "$(dirname "$0")/.." && pwd)"
BUILD="${BUILD:-$SRC/build}"
APPDIR="$BUILD/AppDir"
PY=python3.12
VERSION="$($PY -c "import re;print(re.search(r'VERSION = \"(.*?)\"', open('$SRC/ide/lbstudio/__init__.py').read()).group(1))")"
OUT="$BUILD/LinBASIC-Studio-$VERSION-x86_64.AppImage"
TOOLS="$BUILD/tools"
mkdir -p "$TOOLS"

fetch() { [ -s "$2" ] || curl -sSL -o "$2" "$1"; }
APPIMAGETOOL="${APPIMAGETOOL:-$TOOLS/appimagetool}"
RUNTIME="${RUNTIME:-$TOOLS/runtime-x86_64}"
EXCLUDELIST="${EXCLUDELIST:-$TOOLS/excludelist}"
fetch https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-x86_64.AppImage "$APPIMAGETOOL"
fetch https://github.com/AppImage/type2-runtime/releases/download/continuous/runtime-x86_64 "$RUNTIME"
fetch https://raw.githubusercontent.com/AppImageCommunity/pkg2appimage/master/excludelist "$EXCLUDELIST"
chmod +x "$APPIMAGETOOL"

echo "== LinBASIC Studio $VERSION"
# ------------------------------------------------------------- the runtime and its tcc
sh "$SRC/runtime/build.sh" >/dev/null
sh "$SRC/runtime/db/build.sh" >/dev/null
sh "$SRC/runtime/forms/build.sh" >/dev/null
[ -x "$SRC/runtime/tcc-linux/tcc" ] || sh "$SRC/packaging/build-tcc-linux.sh"

rm -rf "$APPDIR"
SHARE="$APPDIR/usr/share/linbasic-studio"
LIBDIR="$SHARE/lib"
mkdir -p "$APPDIR/usr/bin" "$APPDIR/usr/lib" "$LIBDIR" \
         "$APPDIR/usr/share/applications" "$APPDIR/usr/share/icons/hicolor/scalable/apps" \
         "$APPDIR/usr/share/icons/hicolor/256x256/apps"

# ---------------------------------------------------------------- Python 3.12
cp -L /usr/bin/python3.12 "$APPDIR/usr/bin/python3.12"
cp -a /usr/lib/python3.12 "$APPDIR/usr/lib/"
( cd "$APPDIR/usr/lib/python3.12" && rm -rf test idlelib tkinter turtledemo ensurepip lib2to3 pydoc_data \
    config-3.12-* dist-packages site-packages venv sitecustomize.py && \
  find . -name '__pycache__' -prune -exec rm -rf {} + && find . -type d -name tests -prune -exec rm -rf {} + )

# ---------------------------------------------------------- PyQt6 + QScintilla + pefile
VENV="$BUILD/venv"
[ -x "$VENV/bin/pip" ] || $PY -m venv "$VENV"
"$VENV/bin/pip" install -q --upgrade "PyQt6==6.12.*" "PyQt6-Qt6==6.12.*" "PyQt6-QScintilla==2.14.*" pefile
SP="$("$VENV/bin/python" -c 'import site;print(site.getsitepackages()[0])')"
cp -a "$SP/pefile.py" "$SP/ordlookup" "$LIBDIR/"
mkdir -p "$LIBDIR/PyQt6"
cp -a "$SP/PyQt6/__init__.py" "$LIBDIR/PyQt6/"
cp -a "$SP"/PyQt6/sip* "$LIBDIR/PyQt6/" 2>/dev/null || true
for m in QtCore QtGui QtWidgets QtNetwork QtPrintSupport QtSvg QtDBus Qsci; do
  cp -a "$SP/PyQt6/$m".abi3.so "$LIBDIR/PyQt6/"
done
mkdir -p "$LIBDIR/PyQt6/Qt6/lib" "$LIBDIR/PyQt6/Qt6/plugins"
QT="$SP/PyQt6/Qt6"
for l in Core Gui Widgets Network PrintSupport DBus Svg XcbQpa WaylandClient WaylandEglClientHwIntegration \
         OpenGL EglFSDeviceIntegration; do
  cp -a "$QT"/lib/libQt6$l.so* "$LIBDIR/PyQt6/Qt6/lib/" 2>/dev/null || true
done
cp -a "$QT"/lib/libicu*.so* "$LIBDIR/PyQt6/Qt6/lib/"
cp -a "$QT"/lib/libqscintilla2_qt6.so* "$LIBDIR/PyQt6/Qt6/lib/" 2>/dev/null || true
for p in platforms platformthemes platforminputcontexts imageformats iconengines printsupport \
         xcbglintegrations wayland-decoration-client wayland-graphics-integration-client \
         wayland-shell-integration; do
  [ -d "$QT/plugins/$p" ] && cp -a "$QT/plugins/$p" "$LIBDIR/PyQt6/Qt6/plugins/"
done
rm -f "$LIBDIR"/PyQt6/Qt6/plugins/platforms/libqeglfs.so "$LIBDIR"/PyQt6/Qt6/plugins/platforms/libqlinuxfb.so \
      "$LIBDIR"/PyQt6/Qt6/plugins/platforms/libqvnc.so "$LIBDIR"/PyQt6/Qt6/plugins/platforms/libqvkkhrdisplay.so \
      "$LIBDIR"/PyQt6/Qt6/plugins/imageformats/libqpdf.so 2>/dev/null || true
[ -d "$QT/translations" ] && { mkdir -p "$LIBDIR/PyQt6/Qt6/translations"
  cp -a "$QT"/translations/qtbase_el.qm "$QT"/translations/qtbase_en.qm "$LIBDIR/PyQt6/Qt6/translations/" 2>/dev/null || true; }
check_qt_deps() {
  local f="$1"
  for need in $(ldd "$f" 2>/dev/null | awk '/libQt6/ {print $1}'); do
    [ -e "$LIBDIR/PyQt6/Qt6/lib/$need" ] || return 1
  done
  return 0
}
find "$LIBDIR/PyQt6/Qt6/plugins" -name '*.so' | while read -r f; do
  check_qt_deps "$f" || { echo "   drop plugin $(basename "$f")"; rm -f "$f"; }
done

# ------------------------------------------------------------------ LinBASIC Studio
mkdir -p "$SHARE/ide" "$SHARE/compiler" "$SHARE/data" "$SHARE/runtime/db" "$SHARE/runtime/forms"
cp -a "$SRC/ide/lbstudio" "$SHARE/ide/"
cp -a "$SRC/compiler/lbvc" "$SRC/compiler/lbvc.py" "$SHARE/compiler/"
cp -a "$SRC/data/vb6" "$SHARE/data/"
cp -a "$SRC/runtime/vbrt.h" "$SRC/runtime/libvbrt.a" "$SRC/runtime/libvbdb.a" "$SHARE/runtime/"
cp -a "$SRC/runtime/db/lbdbinfo" "$SHARE/runtime/db/"
cp -a "$SRC/runtime/forms/libvbforms.so" "$SHARE/runtime/forms/"
cp -a "$SRC/runtime/tcc-linux" "$SHARE/runtime/"
# the programs for Windows (Make, the Package and Deployment Wizard): the cross Tiny C Compiler, the runtime DLLs
# and Qt 5 for Windows of runtime/win32, runtime/win64 (lbvc with LBVC_WIN_CC=tcc, set by AppRun; no mingw-w64)
for k in win32 win64; do
  kd="$SRC/runtime/$k"
  if [ -x "$kd/tcc/cross-tcc" ] && [ -f "$kd/dll/vbrt.dll" ] && [ -f "$kd/qt/bin/Qt5Core.dll" ]; then
    mkdir -p "$SHARE/runtime/$k/tcc" "$SHARE/runtime/$k/qt/bin" "$SHARE/runtime/$k/qt/plugins"
    cp -a "$kd/tcc/cross-tcc" "$kd/tcc/include" "$kd/tcc/lib" "$kd/tcc/COPYING" "$SHARE/runtime/$k/tcc/"
    cp -a "$kd/dll" "$SHARE/runtime/$k/"
    cp -a "$kd"/qt/bin/*.dll "$SHARE/runtime/$k/qt/bin/"
    for p in platforms imageformats styles; do
      if [ -d "$kd/qt/plugins/$p" ]; then cp -a "$kd/qt/plugins/$p" "$SHARE/runtime/$k/qt/plugins/"; fi
    done
    echo "   + runtime/$k (programs for Windows)"
  else
    echo "   (no runtime/$k: no programs for $k - packaging/build-tcc.sh, build-qt-windows.sh, runtime/build-dll-win.sh)"
  fi
done
# the AppImages of the users' programs (lbvc --package --format appimage): appimagetool and the type2 runtime
mkdir -p "$SHARE/runtime/appimage"
cp "$APPIMAGETOOL" "$SHARE/runtime/appimage/appimagetool"
cp "$RUNTIME" "$SHARE/runtime/appimage/runtime-x86_64"
for l in appimagetool runtime-x86_64; do
  repo=$([ $l = appimagetool ] && echo appimagetool || echo type2-runtime)
  [ -s "$TOOLS/$l.LICENSE" ] || curl -fsSL -o "$TOOLS/$l.LICENSE" "https://raw.githubusercontent.com/AppImage/$repo/main/LICENSE" || true
  if [ -s "$TOOLS/$l.LICENSE" ]; then cp "$TOOLS/$l.LICENSE" "$SHARE/runtime/appimage/"; fi
done
find "$SHARE" -name '__pycache__' -prune -exec rm -rf {} +
cp "$SRC/LICENSE" "$SRC/RUNTIME-EXCEPTION" "$SRC/README.md" "$SHARE/"
cp -a "$SRC/licenses" "$SHARE/"
cp "$SP"/pyqt6_qt6-*.dist-info/LICENSE "$SHARE/licenses/Qt6-of-PyQt6-LICENSE.txt" 2>/dev/null || true

# --------------------------------------------- host libraries not on the excludelist
EXCL="$(grep -v '^\s*#' "$EXCLUDELIST" | awk '{print $1}' | grep -v '^$' | sed 's/\.so.*//' | sort -u)"
is_excluded() { local b="${1%%.so*}"; grep -qx "$b" <<<"$EXCL"; }
collect() {
  # name <TAB> path (the path may have spaces: "LinBASIC Studio")
  ldd "$1" 2>/dev/null | awk -F' => ' '/=> \// {n=$1; sub(/^[ \t]+/, "", n); p=$2; sub(/ \(0x[0-9a-f]+\)$/, "", p); print n "\t" p}' |
  while IFS=$'\t' read -r name path; do
    case "$path" in "$APPDIR"/*|/usr/local/*) continue;; esac
    is_excluded "$name" && continue
    case "$name" in libz.so*|libpython*|libQt6*|libicu*|libstdc++*|libgcc_s*) continue;; esac
    # core system libraries every desktop already has; bundling them only risks clashes
    case "$name" in libglib-2.0*|libgthread*|libgio-2*|libgobject*|libpcre2*|libdbus-1*|libsystemd*|libcap.so*|\
      libgcrypt*|libgpg-error*|liblz4*|libzstd*|libgssapi*|libkrb5*|libk5crypto*|libkeyutils*|libcom_err*|\
      libgnutls*|libp11-kit*|libidn2*|libunistring*|libtasn1*|libnettle*|libhogweed*|libgmp*|libbsd*|libmd.so*|\
      libXau*|libXdmcp*|libcrypt.so*|libdb-5*|libreadline*|libpanelw*|libncursesw*|libtinfo*|libavahi*|\
      libssl.so*|libcrypto.so*|libsqlite3*) continue;; esac
    case "$name" in libgtk-*|libgdk*|libgmodule*|libmount*|libblkid*|libselinux*|libcairo*|libpixman*|libpango*|\
      libthai*|libdatrie*|libgraphite2*|libepoxy*|libX*|libatk*|libatspi*|libjpeg*|libtiff*|libwebp*|librsvg*) continue;; esac
    [ -e "$APPDIR/usr/lib/$name" ] && continue
    cp -L "$path" "$APPDIR/usr/lib/$name" && echo "   + $name" && collect "$path"
  done
}
export LD_LIBRARY_PATH="$LIBDIR/PyQt6/Qt6/lib:$APPDIR/usr/lib"
# (one name a line, NUL-separated: the source folder may have spaces in its path, "LinBASIC Studio")
while IFS= read -r -d '' f; do
  collect "$f" || true                  # (ldd of a file that is not a dynamic object fails: nothing to collect)
done < <(printf '%s\0' "$APPDIR/usr/bin/python3.12"
         find "$APPDIR/usr/lib/python3.12/lib-dynload" "$LIBDIR" -name '*.so*' -type f -not -path '*/platformthemes/*' -print0)
unset LD_LIBRARY_PATH
# which Ubuntu package (and so which license) each of them comes from
( cd "$APPDIR/usr/lib" && for f in *.so*; do
    [ -f "$f" ] || continue
    real="$(readlink -f "$(ldconfig -p | awk -v n="$f" '$1==n {print $NF; exit}')" 2>/dev/null || true)"
    pkg="$( { dpkg -S "$real" 2>/dev/null || dpkg -S "/usr$real" 2>/dev/null || true; } | cut -d: -f1 | head -1)"
    echo "$f  ${pkg:-?}  (/usr/share/doc/${pkg:-?}/copyright)"
  done ) > "$APPDIR/usr/lib/LIBRARIES.txt"

# --------------------------------------------------------- pre-compile .pyc
"$APPDIR/usr/bin/python3.12" -I -m compileall -q -j0 "$APPDIR/usr/lib/python3.12" "$SHARE/ide" "$SHARE/compiler" \
    "$LIBDIR/pefile.py" >/dev/null || true

# ------------------------------------------------- desktop integration + AppRun
cp "$SRC/packaging/linux/linbasic-studio.desktop" "$APPDIR/linbasic-studio.desktop"
cp "$SRC/packaging/linux/linbasic-studio.desktop" "$APPDIR/usr/share/applications/"
cp "$SRC/ide/lbstudio/res/linbasic-studio.svg" "$APPDIR/linbasic-studio.svg"
cp "$SRC/ide/lbstudio/res/linbasic-studio.svg" "$APPDIR/usr/share/icons/hicolor/scalable/apps/"
cp "$SRC/ide/lbstudio/res/linbasic-studio-256.png" "$APPDIR/.DirIcon"
cp "$SRC/ide/lbstudio/res/linbasic-studio-256.png" "$APPDIR/usr/share/icons/hicolor/256x256/apps/linbasic-studio.png"
install -m 755 "$SRC/packaging/linux/AppRun" "$APPDIR/AppRun"

# ------------------------------------------------------------------- squash
# readable by everyone (files copied by some tools are 0600: appimagetool refuses an AppDir like that)
chmod -R a+rX "$APPDIR"
rm -f "$OUT"
ARCH=x86_64 APPIMAGE_EXTRACT_AND_RUN=1 "$APPIMAGETOOL" --no-appstream --runtime-file "$RUNTIME" \
     --comp zstd "$APPDIR" "$OUT"
( cd "$(dirname "$OUT")" && sha256sum "$(basename "$OUT")" > "$(basename "$OUT").sha256" )
du -h "$OUT"
