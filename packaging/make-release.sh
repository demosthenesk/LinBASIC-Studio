#!/bin/bash
# make-release.sh — LinBASIC Studio: puts the three builds of one version into Releases/
#
#   packaging/make-release.sh <version> <AppImage> <win32 folder> <win64 folder> [Releases folder]
#
# Layout (one folder per version, one zip per operating system):
#   Releases/LinBASIC-Studio-<version>/
#       LinBASIC-Studio-<version>-linux-x86_64.zip   the AppImage + README + LICENSE + RUNTIME-EXCEPTION + licenses/
#       LinBASIC-Studio-<version>-win32.zip          portable app, 32-bit Windows
#       LinBASIC-Studio-<version>-win64.zip          portable app, 64-bit Windows
#       SHA256SUMS
# Each zip unpacks into a single folder named like the zip. An existing version
# folder is never overwritten: a new build gets a new version number.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
src="$(dirname "$here")"
[ $# -ge 4 ] || { sed -n '3,14p' "$0"; exit 2; }
ver="$1"; appimage="$2"; win32="$3"; win64="$4"
rel="${5:-$(dirname "$src")/Releases}"
[[ "$ver" =~ ^[0-9]+\.[0-9]+\.[0-9]+([.-][0-9A-Za-z]+)?$ ]] || { echo "bad version: $ver"; exit 2; }
[ -f "$appimage" ] || { echo "no AppImage: $appimage"; exit 1; }
[ -d "$win32" ] || { echo "no win32 folder: $win32"; exit 1; }
[ -d "$win64" ] || { echo "no win64 folder: $win64"; exit 1; }
command -v zip >/dev/null || { echo "zip is needed (sudo apt install zip)"; exit 1; }

name="LinBASIC-Studio-$ver"
out="$rel/$name"
[ -e "$out" ] && { echo "$out exists already — releases are never overwritten"; exit 1; }
tmp="$(mktemp -d)"; trap 'rm -rf "$tmp"' EXIT

pack() {   # pack <os tag> : zips $tmp/<name>-<tag>/ into $out/<name>-<tag>.zip
  local dir="$name-$1"
  cp "$src/README.md" "$src/LICENSE" "$src/RUNTIME-EXCEPTION" "$tmp/$dir/"
  rm -rf "$tmp/$dir/licenses" && cp -r "$src/licenses" "$tmp/$dir/licenses"
  cp "$src/runtime/db/mdb/COPYING.LIB" "$tmp/$dir/licenses/libmdb-COPYING.LIB"
  chmod -R a+rX "$tmp/$dir"                 # (files copied by some tools are 0600)
  (cd "$tmp" && zip -qr9 "$out/$dir.zip" "$dir")
  echo "  $dir.zip"
}

mkdir -p "$out"
echo "$out"
mkdir -p "$tmp/$name-linux-x86_64"
cp "$appimage" "$tmp/$name-linux-x86_64/$name-x86_64.AppImage"
chmod +x "$tmp/$name-linux-x86_64/$name-x86_64.AppImage"
pack linux-x86_64
for bits in 32 64; do
  d="win$bits"; srcdir="$win32"; [ $bits = 64 ] && srcdir="$win64"
  mkdir -p "$tmp/$name-$d"
  cp -r "$srcdir/." "$tmp/$name-$d/"
  pack "$d"
done
(cd "$out" && sha256sum ./*.zip | sed 's| \./| |' > SHA256SUMS)
echo "  SHA256SUMS"
