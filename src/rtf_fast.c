#include <R.h>
#include <Rinternals.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>

/*
 * C_find_matching_brace(text, start)
 *
 * Scan from byte position `start` (1-based) for the matching closing brace.
 * Returns the 1-based position of the matching '}', or NA_INTEGER if not found.
 */
SEXP C_find_matching_brace(SEXP text, SEXP start) {
    const char *s = CHAR(STRING_ELT(text, 0));
    int pos = INTEGER(start)[0] - 1;  /* convert to 0-based */
    int n = (int)strlen(s);
    int depth = 0;

    while (pos < n) {
        char c = s[pos];
        if (c == '{') {
            depth++;
        } else if (c == '}') {
            depth--;
            if (depth == 0) {
                return ScalarInteger(pos + 1);  /* back to 1-based */
            }
        }
        pos++;
    }

    return ScalarInteger(NA_INTEGER);
}

/*
 * C_rtf_unescape(text)
 *
 * Single-pass RTF unescaping:
 *   \uN?.  -> UTF-8 encoded codepoint N (skip replacement char after ?)
 *   \'xx   -> latin1 byte xx converted to UTF-8
 *
 * Returns a single UTF-8 string.
 */

/* Encode a Unicode codepoint as UTF-8 into buf, return number of bytes written */
static int utf8_encode(unsigned int cp, char *buf) {
    if (cp < 0x80) {
        buf[0] = (char)cp;
        return 1;
    } else if (cp < 0x800) {
        buf[0] = (char)(0xC0 | (cp >> 6));
        buf[1] = (char)(0x80 | (cp & 0x3F));
        return 2;
    } else if (cp < 0x10000) {
        buf[0] = (char)(0xE0 | (cp >> 12));
        buf[1] = (char)(0x80 | ((cp >> 6) & 0x3F));
        buf[2] = (char)(0x80 | (cp & 0x3F));
        return 3;
    } else if (cp < 0x110000) {
        buf[0] = (char)(0xF0 | (cp >> 18));
        buf[1] = (char)(0x80 | ((cp >> 12) & 0x3F));
        buf[2] = (char)(0x80 | ((cp >> 6) & 0x3F));
        buf[3] = (char)(0x80 | (cp & 0x3F));
        return 4;
    }
    /* invalid codepoint -> replacement character U+FFFD */
    buf[0] = (char)0xEF;
    buf[1] = (char)0xBF;
    buf[2] = (char)0xBD;
    return 3;
}

/* Convert a hex digit character to its value, or -1 */
static int hex_val(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

/* Latin-1 (0x80-0xFF) to UTF-8: each byte becomes 2 UTF-8 bytes */
static int latin1_to_utf8(unsigned char byte, char *buf) {
    if (byte < 0x80) {
        buf[0] = (char)byte;
        return 1;
    }
    buf[0] = (char)(0xC0 | (byte >> 6));
    buf[1] = (char)(0x80 | (byte & 0x3F));
    return 2;
}

SEXP C_rtf_unescape(SEXP text) {
    const char *s = CHAR(STRING_ELT(text, 0));
    int n = (int)strlen(s);

    /* Worst case: every char becomes 4 UTF-8 bytes. In practice, output <= input. */
    char *out = R_alloc((size_t)n * 4 + 1, 1);
    int oi = 0;
    int i = 0;

    while (i < n) {
        if (s[i] == '\\' && i + 1 < n) {
            if (s[i + 1] == 'u' && i + 2 < n &&
                (isdigit((unsigned char)s[i + 2]) || s[i + 2] == '-')) {
                /* \uN escape: parse the integer */
                int j = i + 2;
                int neg = 0;
                if (s[j] == '-') { neg = 1; j++; }
                long val = 0;
                while (j < n && isdigit((unsigned char)s[j])) {
                    val = val * 10 + (s[j] - '0');
                    j++;
                }
                if (neg) val = -val;

                /* Convert negative RTF values */
                if (val < 0) val += 65536;

                /* Skip optional '?' after the number */
                if (j < n && s[j] == '?') j++;
                /* Skip one replacement character after \uN? */
                if (j < n) j++;

                unsigned int cp = (unsigned int)(val & 0xFFFF);
                oi += utf8_encode(cp, out + oi);
                i = j;
                continue;
            }

            if (s[i + 1] == '\'' && i + 3 < n) {
                /* \'xx hex escape */
                int h1 = hex_val(s[i + 2]);
                int h2 = hex_val(s[i + 3]);
                if (h1 >= 0 && h2 >= 0) {
                    unsigned char byte = (unsigned char)((h1 << 4) | h2);
                    oi += latin1_to_utf8(byte, out + oi);
                    i += 4;
                    continue;
                }
            }
        }

        /* Normal character: copy as-is */
        out[oi++] = s[i++];
    }

    out[oi] = '\0';
    return ScalarString(mkCharCE(out, CE_UTF8));
}

/*
 * C_rtf_cell_to_text(raw)
 *
 * Strip RTF markup from a cell chunk, then unescape. Steps:
 *   1. Remove {\*...} destination groups
 *   2. Iteratively flatten innermost brace groups (strip control words, keep text)
 *   3. Remove remaining \word[-]N and \symbol sequences
 *   4. Remove stray braces
 *   5. Unescape \uN and \'xx sequences
 *   6. Trim whitespace
 */

/* Check if a character is an RTF control word character */
static int is_ctrl_char(char c) {
    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}

SEXP C_rtf_cell_to_text(SEXP raw_text) {
    const char *input = CHAR(STRING_ELT(raw_text, 0));
    int n = (int)strlen(input);

    /* Working buffer: we'll do in-place transformations.
     * Allocate enough for all passes. */
    char *buf = R_alloc((size_t)n + 1, 1);
    char *tmp = R_alloc((size_t)n + 1, 1);
    memcpy(buf, input, n + 1);

    /* Pass 1: Remove {\*...} destination groups (non-nested) */
    {
        int ri = 0, wi = 0;
        int len = (int)strlen(buf);
        while (ri < len) {
            if (buf[ri] == '{' && ri + 1 < len && buf[ri + 1] == '\\' &&
                ri + 2 < len && buf[ri + 2] == '*') {
                /* Find matching closing brace (handling nesting) */
                int depth = 1;
                int j = ri + 3;
                while (j < len && depth > 0) {
                    if (buf[j] == '{') depth++;
                    else if (buf[j] == '}') depth--;
                    j++;
                }
                ri = j; /* skip the entire group */
            } else {
                tmp[wi++] = buf[ri++];
            }
        }
        tmp[wi] = '\0';
        memcpy(buf, tmp, wi + 1);
    }

    /* Pass 2: Iteratively flatten innermost brace groups.
     * For each innermost {...}, strip control words but keep plain text. */
    for (int iter = 0; iter < 20; iter++) {
        int changed = 0;
        int ri = 0, wi = 0;
        int len = (int)strlen(buf);

        while (ri < len) {
            if (buf[ri] == '{') {
                /* Check if this is an innermost group (no nested braces) */
                int j = ri + 1;
                int has_nested = 0;
                int close = -1;
                while (j < len) {
                    if (buf[j] == '{') { has_nested = 1; break; }
                    if (buf[j] == '}') { close = j; break; }
                    /* Skip escape sequences */
                    if (buf[j] == '\\' && j + 1 < len) j++;
                    j++;
                }

                if (!has_nested && close >= 0) {
                    /* Innermost group: copy content, stripping control words */
                    changed = 1;
                    j = ri + 1;
                    while (j < close) {
                        if (buf[j] == '\\' && j + 1 < close && is_ctrl_char(buf[j + 1])) {
                            /* Skip \word[-]N and optional trailing space */
                            j++;
                            while (j < close && is_ctrl_char(buf[j])) j++;
                            if (j < close && buf[j] == '-') j++;
                            while (j < close && isdigit((unsigned char)buf[j])) j++;
                            if (j < close && buf[j] == ' ') j++;
                        } else {
                            tmp[wi++] = buf[j++];
                        }
                    }
                    ri = close + 1;
                } else {
                    tmp[wi++] = buf[ri++];
                }
            } else {
                tmp[wi++] = buf[ri++];
            }
        }
        tmp[wi] = '\0';
        memcpy(buf, tmp, wi + 1);
        if (!changed) break;
    }

    /* Pass 3: Remove remaining control words \word[-]N and optional space */
    {
        int ri = 0, wi = 0;
        int len = (int)strlen(buf);
        while (ri < len) {
            if (buf[ri] == '\\' && ri + 1 < len && is_ctrl_char(buf[ri + 1])) {
                ri++;
                while (ri < len && is_ctrl_char(buf[ri])) ri++;
                if (ri < len && buf[ri] == '-') ri++;
                while (ri < len && isdigit((unsigned char)buf[ri])) ri++;
                if (ri < len && buf[ri] == ' ') ri++;
            } else if (buf[ri] == '\\' && ri + 1 < len) {
                /* Control symbol \<char> — skip both */
                ri += 2;
            } else {
                tmp[wi++] = buf[ri++];
            }
        }
        tmp[wi] = '\0';
        memcpy(buf, tmp, wi + 1);
    }

    /* Pass 4: Remove stray braces */
    {
        int ri = 0, wi = 0;
        int len = (int)strlen(buf);
        while (ri < len) {
            if (buf[ri] != '{' && buf[ri] != '}') {
                tmp[wi++] = buf[ri];
            }
            ri++;
        }
        tmp[wi] = '\0';
        memcpy(buf, tmp, wi + 1);
    }

    /* Pass 5: Unescape \uN and \'xx — reuse the C_rtf_unescape logic inline */
    {
        int len = (int)strlen(buf);
        /* UTF-8 output can be up to 4x input */
        char *out = R_alloc((size_t)len * 4 + 1, 1);
        int oi = 0;
        int ii = 0;

        while (ii < len) {
            if (buf[ii] == '\\' && ii + 1 < len) {
                if (buf[ii + 1] == 'u' && ii + 2 < len &&
                    (isdigit((unsigned char)buf[ii + 2]) || buf[ii + 2] == '-')) {
                    int j = ii + 2;
                    int neg = 0;
                    if (buf[j] == '-') { neg = 1; j++; }
                    long val = 0;
                    while (j < len && isdigit((unsigned char)buf[j])) {
                        val = val * 10 + (buf[j] - '0');
                        j++;
                    }
                    if (neg) val = -val;
                    if (val < 0) val += 65536;
                    if (j < len && buf[j] == '?') j++;
                    if (j < len) j++;
                    oi += utf8_encode((unsigned int)(val & 0xFFFF), out + oi);
                    ii = j;
                    continue;
                }
                if (buf[ii + 1] == '\'' && ii + 3 < len) {
                    int h1 = hex_val(buf[ii + 2]);
                    int h2 = hex_val(buf[ii + 3]);
                    if (h1 >= 0 && h2 >= 0) {
                        unsigned char byte = (unsigned char)((h1 << 4) | h2);
                        oi += latin1_to_utf8(byte, out + oi);
                        ii += 4;
                        continue;
                    }
                }
            }
            out[oi++] = buf[ii++];
        }
        out[oi] = '\0';
        buf = out;
    }

    /* Pass 6: Trim leading and trailing whitespace */
    {
        int len = (int)strlen(buf);
        int start = 0, end = len - 1;
        while (start < len && (buf[start] == ' ' || buf[start] == '\t' ||
               buf[start] == '\r' || buf[start] == '\n')) start++;
        while (end >= start && (buf[end] == ' ' || buf[end] == '\t' ||
               buf[end] == '\r' || buf[end] == '\n')) end--;
        int result_len = end - start + 1;
        if (result_len < 0) result_len = 0;

        char *result = R_alloc((size_t)result_len + 1, 1);
        if (result_len > 0) memcpy(result, buf + start, result_len);
        result[result_len] = '\0';
        return ScalarString(mkCharCE(result, CE_UTF8));
    }
}
