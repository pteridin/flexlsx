#include <R.h>
#include <Rinternals.h>
#include <R_ext/Rdynload.h>

SEXP R_merge_resolve_type(SEXP df_to_merge) {
  SEXP span_rows = VECTOR_ELT(df_to_merge, 0);
  SEXP span_cols = VECTOR_ELT(df_to_merge, 1);
  SEXP row_id = VECTOR_ELT(df_to_merge, 2);
  SEXP row_end = VECTOR_ELT(df_to_merge, 3);
  SEXP col_id = VECTOR_ELT(df_to_merge, 4);
  SEXP col_end = VECTOR_ELT(df_to_merge, 5);
  SEXP dims = VECTOR_ELT(df_to_merge, 6);
  SEXP merge_type = VECTOR_ELT(df_to_merge, 7);

  R_xlen_t length = XLENGTH(row_id);

  SEXP is_encapsulated = PROTECT(allocVector(LGLSXP, length));
  SEXP is_need_resolve = PROTECT(allocVector(LGLSXP, length));

  // Initialize all values to FALSE (0)
  memset(LOGICAL(is_encapsulated), 0, sizeof(int) * (size_t)length);
  memset(LOGICAL(is_need_resolve), 0, sizeof(int) * (size_t)length);

  for (R_xlen_t i = 0; i < length; i++) {
    if (REAL(span_rows)[i] > 1 && REAL(span_cols)[i] > 1) {
      INTEGER(merge_type)[i] = 1;
    } else if (REAL(span_rows)[i] > 1 && REAL(span_cols)[i] == 0) {
      INTEGER(merge_type)[i] = 2;
    }

    if (i == 0) {
      continue;
    }

    int current_is_encapsulated = 0;
    int current_is_need_resolve = 0;

    for (R_xlen_t j = 0; j < i; j++) {
      if (REAL(row_id)[i] >= REAL(row_id)[j] &&
          REAL(row_end)[i] <= REAL(row_end)[j] &&
          REAL(col_id)[i] >= REAL(col_id)[j] &&
          REAL(col_end)[i] <= REAL(col_end)[j]) {
        current_is_encapsulated = 1;
        break;
      }

      double overlap_row_start = (REAL(row_id)[i] > REAL(row_id)[j]) ? REAL(row_id)[i] : REAL(row_id)[j];
      double overlap_row_end = (REAL(row_end)[i] < REAL(row_end)[j]) ? REAL(row_end)[i] : REAL(row_end)[j];
      double overlap_col_start = (REAL(col_id)[i] > REAL(col_id)[j]) ? REAL(col_id)[i] : REAL(col_id)[j];
      double overlap_col_end = (REAL(col_end)[i] < REAL(col_end)[j]) ? REAL(col_end)[i] : REAL(col_end)[j];

      if (overlap_row_start <= overlap_row_end && overlap_col_start <= overlap_col_end) {
        current_is_need_resolve = 1;
        break;
      }
    }

    LOGICAL(is_encapsulated)[i] = current_is_encapsulated;
    LOGICAL(is_need_resolve)[i] = current_is_need_resolve;
  }

  SEXP result = PROTECT(allocVector(VECSXP, 10));
  SET_VECTOR_ELT(result, 0, span_rows);
  SET_VECTOR_ELT(result, 1, span_cols);
  SET_VECTOR_ELT(result, 2, row_id);
  SET_VECTOR_ELT(result, 3, row_end);
  SET_VECTOR_ELT(result, 4, col_id);
  SET_VECTOR_ELT(result, 5, col_end);
  SET_VECTOR_ELT(result, 6, dims);
  SET_VECTOR_ELT(result, 7, merge_type);
  SET_VECTOR_ELT(result, 8, is_encapsulated);
  SET_VECTOR_ELT(result, 9, is_need_resolve);

  SEXP result_names = PROTECT(allocVector(STRSXP, 10));
  SET_STRING_ELT(result_names, 0, mkChar("span.rows"));
  SET_STRING_ELT(result_names, 1, mkChar("span.cols"));
  SET_STRING_ELT(result_names, 2, mkChar("row_id"));
  SET_STRING_ELT(result_names, 3, mkChar("row_end"));
  SET_STRING_ELT(result_names, 4, mkChar("col_id"));
  SET_STRING_ELT(result_names, 5, mkChar("col_end"));
  SET_STRING_ELT(result_names, 6, mkChar("dims"));
  SET_STRING_ELT(result_names, 7, mkChar("merge_type"));
  SET_STRING_ELT(result_names, 8, mkChar("is_encapsulated"));
  SET_STRING_ELT(result_names, 9, mkChar("is_need_resolve"));
  setAttrib(result, R_NamesSymbol, result_names);

  SEXP row_names = PROTECT(allocVector(REALSXP, 2));
  REAL(row_names)[0] = NA_REAL;
  REAL(row_names)[1] = (double)-length;
  setAttrib(result, R_RowNamesSymbol, row_names);

  setAttrib(result, R_ClassSymbol, mkString("data.frame"));

  UNPROTECT(5);
  return result;
}

static const R_CallMethodDef CallEntries[] = {
  {"R_merge_resolve_type", (DL_FUNC) &R_merge_resolve_type, 1},
  {NULL, NULL, 0}
};

void R_init_flexlsx(DllInfo *dll) {
  R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
  R_useDynamicSymbols(dll, FALSE);
}
