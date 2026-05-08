# Fetch and tidy data

Fetch and tidy data

## Usage

``` r
fetchAndTidy(tbl_name, connection, as.is)
```

## Arguments

- tbl_name:

  Name of data table

- connection:

  Database connection object

- as.is:

  which (if any) columns returned as character should be converted to
  another type? Allowed values are as for
  [`read.table`](https://rdrr.io/r/utils/read.table.html). See
  ‘Details’.

## Value

A tibble of tidy data
