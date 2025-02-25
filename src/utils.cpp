#include <Rcpp.h>
using namespace Rcpp;

// [[Rcpp::export]]
DataFrame cpp_merge_resolve_type(DataFrame df_to_merge) {
  NumericVector span_rows = df_to_merge["span.rows"];
  NumericVector span_cols = df_to_merge["span.cols"];
  NumericVector row_id = df_to_merge["row_id"];
  NumericVector row_end = df_to_merge["row_end"];
  NumericVector col_id = df_to_merge["col_id"];
  NumericVector col_end = df_to_merge["col_end"];
  CharacterVector dims = df_to_merge["dims"];
  IntegerVector merge_type = df_to_merge["merge_type"];

  LogicalVector is_encapsulated(row_id.length(), false);
  LogicalVector is_need_resolve(row_id.length(), false);

  for (int i = 0; i < row_id.length(); i++) {
    if(span_rows[i] > 1 && span_cols[i] > 1) {
      merge_type[i] = 1;
    } else if(span_rows[i] > 1 && span_cols[i] == 0) {
      merge_type[i] = 2;
    }

    if (i == 0) {
      continue;
    }

    bool current_is_encapsulated = false;
    bool current_is_need_resolve = false;

    for (int j = 0; j < i; j++) {
      // Is encapsulated?
      if (row_id[i] >= row_id[j] &&
          row_end[i] <= row_end[j] &&
          col_id[i] >= col_id[j] &&
          col_end[i] <= col_end[j]) {
        current_is_encapsulated = true;
        break;
      }

      // Is overlap?
      int overlap_row_start = std::max(row_id[i], row_id[j]);
      int overlap_row_end = std::min(row_end[i], row_end[j]);
      int overlap_col_start = std::max(col_id[i], col_id[j]);
      int overlap_col_end = std::min(col_end[i], col_end[j]);

      if (overlap_row_start <= overlap_row_end &&
          overlap_col_start <= overlap_col_end) {
        current_is_need_resolve = true;
        break;
      }
    }

    is_encapsulated[i] = current_is_encapsulated;
    is_need_resolve[i] = current_is_need_resolve;
  }


  return DataFrame::create(
    Named("span.rows") = span_rows,
    Named("span.cols") = span_cols,
    Named("row_id") = row_id,
    Named("row_end") = row_end,
    Named("col_id") = col_id,
    Named("col_end") = col_end,
    Named("dims") = dims,
    Named("merge_type") = merge_type,
    Named("is_encapsulated") = is_encapsulated,
    Named("is_need_resolve") = is_need_resolve
  );
}
