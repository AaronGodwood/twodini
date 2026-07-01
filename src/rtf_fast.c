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

/* Windows-1252 codepoints for bytes 0x80-0x9F (latin1 has C1 controls here;
 * real RTF from SAS/Word uses CP1252 smart quotes, dashes, ellipsis, ...) */
static const unsigned short cp1252_high[32] = {
    0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
    0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
    0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
    0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178
};

/* CP1252 byte to UTF-8: bytes 0x80-0x9F via table, rest as latin1 */
static int cp1252_to_utf8(unsigned char byte, char *buf) {
    unsigned int cp = byte;
    if (byte >= 0x80 && byte <= 0x9F) cp = cp1252_high[byte - 0x80];
    return utf8_encode(cp, buf);
}

/* Check if a character is an RTF control word character */
static int is_ctrl_char(char c) {
    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}

/* Is s[i] the start of a \uN or \ucN sequence (which the unescape pass
 * decodes, so the control-word stripping passes must leave them alone)? */
static int is_unicode_escape(const char *s, int i, int end) {
    if (i + 1 >= end || s[i] != '\\' || s[i + 1] != 'u') return 0;
    if (i + 2 < end && (isdigit((unsigned char)s[i + 2]) || s[i + 2] == '-'))
        return 1;
    if (i + 3 < end && s[i + 2] == 'c' && isdigit((unsigned char)s[i + 3]))
        return 1;
    return 0;
}

/*
 * Shared unescape core. Decodes, from s[0..len):
 *   \ucN  -> sets the fallback length for subsequent \uN (default 1); consumed
 *   \uN   -> UTF-8 codepoint N; then skips the current uc count of fallback
 *            characters (a \'xx escape counts as one fallback character)
 *   \'xx  -> CP1252 byte xx as UTF-8
 * Everything else is copied through. out must have room for len*4+1 bytes.
 */
static void rtf_unescape_core(const char *s, int len, char *out) {
    int oi = 0;
    int i = 0;
    int uc = 1;

    while (i < len) {
        if (s[i] == '\\' && i + 1 < len) {
            /* \ucN : update fallback count, emit nothing */
            if (s[i + 1] == 'u' && i + 3 < len && s[i + 2] == 'c' &&
                isdigit((unsigned char)s[i + 3])) {
                int j = i + 3;
                int val = 0;
                while (j < len && isdigit((unsigned char)s[j])) {
                    val = val * 10 + (s[j] - '0');
                    j++;
                }
                uc = val;
                if (j < len && s[j] == ' ') j++;  /* control word delimiter */
                i = j;
                continue;
            }

            /* \uN : decode codepoint, then skip uc fallback characters */
            if (s[i + 1] == 'u' && i + 2 < len &&
                (isdigit((unsigned char)s[i + 2]) || s[i + 2] == '-')) {
                int j = i + 2;
                int neg = 0;
                if (s[j] == '-') { neg = 1; j++; }
                long val = 0;
                while (j < len && isdigit((unsigned char)s[j])) {
                    val = val * 10 + (s[j] - '0');
                    j++;
                }
                if (neg) val = -val;
                if (val < 0) val += 65536;
                if (j < len && s[j] == ' ') j++;  /* control word delimiter */

                /* Skip fallback chars; stop early at structure we mustn't eat */
                int k = uc;
                while (k > 0 && j < len) {
                    if (s[j] == '\\') {
                        if (j + 1 < len && s[j + 1] == '\'') j += 4;
                        else break;
                    } else if (s[j] == '{' || s[j] == '}') {
                        break;
                    } else {
                        j++;
                    }
                    k--;
                }

                oi += utf8_encode((unsigned int)(val & 0xFFFF), out + oi);
                i = j;
                continue;
            }

            /* \'xx hex escape */
            if (s[i + 1] == '\'' && i + 3 < len) {
                int h1 = hex_val(s[i + 2]);
                int h2 = hex_val(s[i + 3]);
                if (h1 >= 0 && h2 >= 0) {
                    unsigned char byte = (unsigned char)((h1 << 4) | h2);
                    oi += cp1252_to_utf8(byte, out + oi);
                    i += 4;
                    continue;
                }
            }
        }

        /* Normal character: copy as-is */
        out[oi++] = s[i++];
    }

    out[oi] = '\0';
}

/*
 * C_rtf_unescape(text)
 *
 * Resolve \ucN / \uN / \'xx escape sequences. Returns a UTF-8 string.
 */
SEXP C_rtf_unescape(SEXP text) {
    const char *s = CHAR(STRING_ELT(text, 0));
    int n = (int)strlen(s);

    /* Worst case: every escape byte expands to 4 UTF-8 bytes */
    char *out = R_alloc((size_t)n * 4 + 1, 1);
    rtf_unescape_core(s, n, out);
    return ScalarString(mkCharCE(out, CE_UTF8));
}

/*
 * C_rtf_cell_to_text(raw)
 *
 * Strip RTF markup from a cell chunk, then unescape. Steps:
 *   1. Remove {\*...} destination groups
 *   2. Iteratively flatten innermost brace groups (strip control words,
 *      keep plain text and \uN / \ucN / \'xx escapes)
 *   3. Remove remaining \word[-]N and \symbol sequences (preserving escapes)
 *   4. Remove stray braces
 *   5. Unescape \ucN / \uN / \'xx sequences
 *   6. Trim whitespace
 */
SEXP C_rtf_cell_to_text(SEXP raw_text) {
    const char *input = CHAR(STRING_ELT(raw_text, 0));
    int n = (int)strlen(input);

    /* Working buffer: we'll do in-place transformations.
     * Allocate enough for all passes. */
    char *buf = R_alloc((size_t)n + 1, 1);
    char *tmp = R_alloc((size_t)n + 1, 1);
    memcpy(buf, input, n + 1);

    /* Pass 1: Remove {\*...} destination groups */
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
     * For each innermost {...}, strip control words but keep plain text
     * and escape sequences (decoded in pass 5). */
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
                        if (buf[j] == '\\' && j + 1 < close &&
                            is_ctrl_char(buf[j + 1]) &&
                            !is_unicode_escape(buf, j, close)) {
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

    /* Pass 3: Remove remaining control words \word[-]N and control symbols,
     * preserving \uN / \ucN / \'xx for the unescape pass */
    {
        int ri = 0, wi = 0;
        int len = (int)strlen(buf);
        while (ri < len) {
            if (buf[ri] == '\\' && is_unicode_escape(buf, ri, len)) {
                /* Copy the backslash; the rest copies as plain text */
                tmp[wi++] = buf[ri++];
            } else if (buf[ri] == '\\' && ri + 1 < len &&
                       is_ctrl_char(buf[ri + 1])) {
                ri++;
                while (ri < len && is_ctrl_char(buf[ri])) ri++;
                if (ri < len && buf[ri] == '-') ri++;
                while (ri < len && isdigit((unsigned char)buf[ri])) ri++;
                if (ri < len && buf[ri] == ' ') ri++;
            } else if (buf[ri] == '\\' && ri + 1 < len && buf[ri + 1] == '\'') {
                /* Preserve \'xx hex escape: copy the \' pair, hex copies as text */
                tmp[wi++] = buf[ri++];
                tmp[wi++] = buf[ri++];
            } else if (buf[ri] == '\\' && ri + 1 < len) {
                /* Control symbol \<char> - skip both */
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

    /* Pass 5: Unescape \ucN / \uN / \'xx */
    {
        int len = (int)strlen(buf);
        char *out = R_alloc((size_t)len * 4 + 1, 1);
        rtf_unescape_core(buf, len, out);
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
