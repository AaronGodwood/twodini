#include <R.h>
#include <Rinternals.h>
#include <R_ext/Rdynload.h>

/* Declarations from rtf_fast.c */
extern SEXP C_find_matching_brace(SEXP text, SEXP start);
extern SEXP C_rtf_unescape(SEXP text);
extern SEXP C_rtf_cell_to_text(SEXP raw_text);

static const R_CallMethodDef CallEntries[] = {
    {"C_find_matching_brace", (DL_FUNC) &C_find_matching_brace, 2},
    {"C_rtf_unescape",       (DL_FUNC) &C_rtf_unescape,        1},
    {"C_rtf_cell_to_text",   (DL_FUNC) &C_rtf_cell_to_text,    1},
    {NULL, NULL, 0}
};

void R_init_houdini(DllInfo *dll) {
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
