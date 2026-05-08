# Generate column spec from data dictionary

Given a fields data dictionary, create a list of column specifications
that can be used in
[`readr::read_csv()`](https://readr.tidyverse.org/reference/read_delim.html)
or [`vroom::vroom()`](https://vroom.tidyverse.org/reference/vroom.html)

## Usage

``` r
makeColSpec(fields)
```

## Arguments

- fields:

  Fields data dictionary, as returned by
  [fetchFromAccess](https://wright13.github.io/imd-fetchaccess-package/reference/fetchFromAccess.md)

## Value

A list of lists
