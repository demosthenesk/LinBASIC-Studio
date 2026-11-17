#!/bin/bash
# extract-vb6-msdn.sh — LinBASIC Studio, Phase 0 (REA): the inputs of tools/rea/ from your own
# Windows virtual machine with VB6 (and MSDN) installed.
#
# Copies, READ-ONLY, from the VirtualBox disk (.vdi) of that machine:
#   1. the VB6 installation folder (VB98)                 -> <out>/VB6-REA/VB98
#   2. the VB6 runtime, the .ocx and the type libraries   -> <out>/VB6-REA/System32, CommonFiles
#   3. the whole MSDN Library folder                      -> <out>/_MSDN-full
#      (the folder with its .CHM files: what Tools > Options > MSDN Library of the IDE asks for)
#
# None of this ever goes into LinBASIC Studio or its repository: they are Microsoft's files, for
# local analysis and for the help of those who own a legal copy of VB6 / MSDN.
#
# Usage:   sudo tools/rea/extract-vb6-msdn.sh <disk.vdi> <output folder> [VM name]
# Needs:   qemu-utils (qemu-nbd), ntfs-3g or the kernel's ntfs3 driver
#          sudo apt install qemu-utils ntfs-3g
# The virtual machine must be powered off (not in a saved state).
# Then:    python3 tools/rea/dump_typelibs.py  <output folder>/VB6-REA data/vb6/typelibs
#          python3 tools/rea/dump_resources.py <output folder>/VB6-REA data/vb6

set -uo pipefail

[ $# -ge 2 ] || { sed -n '2,20p' "$0"; exit 2; }
VDI="$1"
OUT="$2"
VM_NAME="${3:-$(basename "$VDI" .vdi)}"
NBD=""
MNT="$(mktemp -d /tmp/w7-ro.XXXXXX)"

die()  { echo "ERROR: $*" >&2; exit 1; }
info() { echo ">> $*"; }

cleanup() {
  mountpoint -q "$MNT" && umount "$MNT"
  [ -n "$NBD" ] && qemu-nbd -d "$NBD" >/dev/null 2>&1
  rmdir "$MNT" 2>/dev/null
}
trap cleanup EXIT

[ "$(id -u)" -eq 0 ] || die "run it with sudo"
[ -f "$VDI" ] || die "$VDI not found"
command -v qemu-nbd >/dev/null || die "qemu-nbd is missing (sudo apt install qemu-utils)"

# --- 1. the virtual machine must be off -------------------------------------------
REAL_USER="${SUDO_USER:-$USER}"
if command -v VBoxManage >/dev/null; then
  STATE="$(sudo -u "$REAL_USER" VBoxManage showvminfo "$VM_NAME" --machinereadable 2>/dev/null \
           | sed -n 's/^VMState="\(.*\)"/\1/p')"
  info "VM state: ${STATE:-unknown}"
  if [ -n "$STATE" ] && [ "$STATE" != "poweroff" ] && [ "${FORCE:-0}" != 1 ]; then
    die "the VM is '$STATE'. Shut Windows down and run this again."
  fi
fi

# --- 2. the .vdi, read-only ------------------------------------------------------
modprobe nbd max_part=16 || die "the nbd module could not be loaded"
for d in /dev/nbd{0..15}; do
  if [ "$(cat /sys/block/$(basename "$d")/size 2>/dev/null)" = "0" ]; then NBD="$d"; break; fi
done
[ -n "$NBD" ] || die "no free /dev/nbdX"
info "connecting $VDI -> $NBD (read-only)"
qemu-nbd --read-only --format=vdi --connect="$NBD" "$VDI" || die "qemu-nbd failed"
sleep 2; partprobe "$NBD" 2>/dev/null; sleep 1

# --- 3. the partition of C: --------------------------------------------------------
FOUND=""
for p in "${NBD}"p*; do
  [ -b "$p" ] || continue
  if mount -t ntfs3 -o ro "$p" "$MNT" 2>/dev/null || mount -t ntfs-3g -o ro "$p" "$MNT" 2>/dev/null; then
    if [ -d "$MNT/Windows/System32" ]; then FOUND="$p"; break; fi
    umount "$MNT"
  fi
done
[ -n "$FOUND" ] || die "no partition with Windows found"
info "Windows on $FOUND, mounted on $MNT"

# case-insensitive search
ifind() { find "$1" -maxdepth "${3:-1}" -iname "$2" 2>/dev/null | head -n1; }
PF="$(ifind "$MNT" 'Program Files')"
SYS32="$(ifind "$MNT/Windows" 'System32')"
COMMON="$(ifind "$PF" 'Common Files')"

mkdir -p "$OUT/VB6-REA/System32" "$OUT/VB6-REA/CommonFiles"
copy() { cp -r --preserve=timestamps "$@"; }

# --- 4. VB6 (VB98) -------------------------------------------------------------------
VB6EXE="$(find "$PF" -maxdepth 5 -iname 'VB6.EXE' 2>/dev/null | head -n1)"
[ -n "$VB6EXE" ] || VB6EXE="$(find "$MNT" -xdev -iname 'VB6.EXE' 2>/dev/null | head -n1)"
if [ -n "$VB6EXE" ]; then
  VB98="$(dirname "$VB6EXE")"
  info "VB6: $VB98"
  rm -rf "$OUT/VB6-REA/VB98"
  copy "$VB98" "$OUT/VB6-REA/VB98"
else
  echo "WARNING: VB6.EXE not found" >&2
fi

# --- 5. runtime, .ocx and type libraries -----------------------------------------------
info "runtime / OCX / TLB from $SYS32"
for n in MSVBVM60.DLL VB6STKIT.DLL STDOLE2.TLB STDOLE32.TLB MSSTDFMT.DLL MSBIND.DLL \
         VB6EXT.OLB VB6DEBUG.DLL VBAME.DLL MSDBRPTR.DLL MSDERUN.DLL ASYCFILT.DLL \
         SCRRUN.DLL MSCOMCT2.OCX MSCOMCTL.OCX COMDLG32.OCX; do
  f="$(ifind "$SYS32" "$n")"; [ -n "$f" ] && copy "$f" "$OUT/VB6-REA/System32/"
done
# every ActiveX control (*.OCX): small, and their type libraries are needed
find "$SYS32" -maxdepth 1 -iname '*.ocx' -exec cp --preserve=timestamps {} "$OUT/VB6-REA/System32/" \;
# ADO and DAO
if [ -n "$COMMON" ]; then
  for sub in 'System/ado' 'Microsoft Shared/DAO'; do
    d="$COMMON/$sub"; [ -d "$d" ] && copy "$d" "$OUT/VB6-REA/CommonFiles/"
  done
fi

# --- 6. MSDN -----------------------------------------------------------------------------
MSDNCOL="$(find "$PF" -maxdepth 7 \( -iname '*.col' -o -iname 'MSDN*.CHM' \) -ipath '*msdn*' 2>/dev/null | head -n1)"
if [ -n "$MSDNCOL" ]; then
  MSDNDIR="$(dirname "$MSDNCOL")"
  info "MSDN: $MSDNDIR"
  rm -rf "$OUT/_MSDN-full"
  copy "$MSDNDIR" "$OUT/_MSDN-full"
  echo "$MSDNDIR" | sed "s|^$MNT|C:|" > "$OUT/_MSDN-full/_SOURCE.txt"
else
  echo "WARNING: no MSDN installation under $PF" >&2
  find "$MNT" -xdev -iname '*.chm' -ipath '*msdn*' 2>/dev/null | head -n 20 >&2
fi

# --- 7. list and owner -------------------------------------------------------------------
( cd "$OUT" && find VB6-REA _MSDN-full -type f -printf '%10s  %p\n' 2>/dev/null | sort -k2 ) \
  > "$OUT/VB6-REA/MANIFEST.txt"
chown -R "$REAL_USER": "$OUT/VB6-REA" "$OUT/_MSDN-full" 2>/dev/null
info "done: $(wc -l < "$OUT/VB6-REA/MANIFEST.txt") files, listed in $OUT/VB6-REA/MANIFEST.txt"
