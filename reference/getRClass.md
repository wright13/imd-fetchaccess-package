# Helper function to get primary R class of each data column

Helper function to get primary R class of each data column

## Usage

``` r
getRClass(fields, data)
```

## Arguments

- fields:

  Fields data dictionary, as returned by
  [fetchFromAccess](https://wright13.github.io/imd-fetchaccess-package/reference/fetchFromAccess.md)

- data:

  List of data tables, as returned by
  [fetchFromAccess](https://wright13.github.io/imd-fetchaccess-package/reference/fetchFromAccess.md)

## Value

`fields` with an additional rClass column
