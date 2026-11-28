/* jetcp.c - the code pages of jetfile.h through the runtime's own tables (intl.c): the texts of a Jet 3 file.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetfile.h"
#include "../vbrt_int.h"

int jet_cp_to_utf16(int cp, const unsigned char *b, int n, uint16_t *out)
{
    return n > 0 ? vb_cp_decode(cp, b, n, (vb_wchar *)out) : 0;
}

int jet_utf16_to_cp(int cp, const uint16_t *w, int n, unsigned char *out)
{
    return n > 0 ? vb_cp_encode(cp, (const vb_wchar *)w, n, out) : 0;
}
