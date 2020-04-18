## Libraries
library(openxlsx)
library(stringr)

## Read XLSX files
contacts <- vector(mode = "list", length = 5)
names(contacts) <- c("All", "Home", "School", "Work", "Other")

y <- vector(mode = "list", length = 152)
names(y) <- c(getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_all_locations_1.xlsx"),
              getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_all_locations_2.xlsx"))

contacts[["All"]] <- contacts[["Home"]] <- contacts[["School"]] <- contacts[["Work"]] <- contacts[["Other"]] <- y

## All locations
for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_all_locations_1.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_all_locations_1.xlsx", sheet = i)
  contacts[["All"]][[i]] <- x
}

for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_all_locations_2.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_all_locations_2.xlsx",
                 sheet = i,
                 colNames = FALSE)
  contacts[["All"]][[i]] <- x
}


## Home
for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_home_1.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_home_1.xlsx", sheet = i)
  contacts[["Home"]][[i]] <- x
}

for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_home_2.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_home_2.xlsx",
                 sheet = i,
                 colNames = FALSE)
  contacts[["Home"]][[i]] <- x
}


## School
for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_school_1.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_school_1.xlsx", sheet = i)
  contacts[["School"]][[i]] <- x
}

for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_school_2.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_school_2.xlsx",
                 sheet = i,
                 colNames = FALSE)
  contacts[["School"]][[i]] <- x
}


## Work
for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_work_1.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_work_1.xlsx", sheet = i)
  contacts[["Work"]][[i]] <- x
}

for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_work_2.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_work_2.xlsx",
                 sheet = i,
                 colNames = FALSE)
  contacts[["Work"]][[i]] <- x
}


## Other
for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_other_locations_1.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_other_locations_1.xlsx", sheet = i)
  contacts[["Other"]][[i]] <- x
}

for(i in getSheetNames("data-raw/contact_matrices_152_countries/MUestimates_other_locations_2.xlsx")) {
  x <- read.xlsx(xlsxFile = "data-raw/contact_matrices_152_countries/MUestimates_other_locations_2.xlsx",
                 sheet = i,
                 colNames = FALSE)
  contacts[["Other"]][[i]] <- x
}

socialcontacts <- contacts

usethis::use_data(socialcontacts, compress = "xz", overwrite = TRUE)
