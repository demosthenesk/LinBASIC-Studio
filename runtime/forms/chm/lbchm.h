/* lbchm.h - CHMLib (chm_lib.h) with names of its own: lbchm_open ... (the system's libchm, or another
 * library with an LZX decoder, may be in the same program). Included first by chm_lib.c and lzx.c
 * (gcc -include) and by forms/help.inc. */
#ifndef LBCHM_H
#define LBCHM_H
#define chm_open lbchm_open
#define chm_close lbchm_close
#define chm_set_param lbchm_set_param
#define chm_resolve_object lbchm_resolve_object
#define chm_retrieve_object lbchm_retrieve_object
#define chm_enumerate lbchm_enumerate
#define chm_enumerate_dir lbchm_enumerate_dir
#define LZXinit lbchm_LZXinit
#define LZXteardown lbchm_LZXteardown
#define LZXreset lbchm_LZXreset
#define LZXdecompress lbchm_LZXdecompress
#ifdef __cplusplus
extern "C" {
#endif
#include "chm_lib.h"
#ifdef __cplusplus
}
#endif
#endif
