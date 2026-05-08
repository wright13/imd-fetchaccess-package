# Write data and data dictionaries to files

You shouldn't need to call this function directly unless you are using
it to write a data export function for another R package. If you are
using this package on its own, you will usually want to call
`fetchFromAccess(save_to_files = TRUE)`.

## Usage

``` r
writeToFiles(
  all_tables,
  data_dir = here::here("data", "final"),
  dictionary_dir = here::here("data", "dictionary"),
  dictionary_filenames = c(tables = "data_dictionary_tables.txt", attributes =
    "data_dictionary_attributes.txt", categories = "data_dictionary_categories.txt"),
  lookup_dir = NA,
  verbose = FALSE
)
```

## Arguments

- all_tables:

  Output of
  [`fetchFromAccess()`](https://wright13.github.io/imd-fetchaccess-package/reference/fetchFromAccess.md)

- data_dir:

  Folder to store data csv's in

- dictionary_dir:

  Folder to store data dictionaries in

- dictionary_filenames:

  Named list with names `c("tables", "attributes", "categories")`
  indicating what to name the tables, attributes, and categories data
  dictionaries. You are encouraged to keep the default names unless you
  have a good reason to change them.

- lookup_dir:

  Optional folder to store lookup tables in. If left as `NA`, lookups
  won't be exported.

- verbose:

  Output feedback to console?
