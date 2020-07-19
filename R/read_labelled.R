#' Импортирует данные из Excel файла с метками переменных и значений
#' @param file Имя файла Excel с двумя листами: для данных и словаря
#' @param data.sheet Имя или номер листа книги, содержащего табличные данные
#' @param vars.sheet Имя или номер листа книги, содержащего словарь
#' @examples
#' # read_labelled("example.xlsx")
#' # read_labelled("example.xlsx", "Data View", "Variable View")
read_labelled <- function(file, data.sheet = 1, vars.sheet = 2) {
  data <- readxl::read_excel(file, sheet = data.sheet)
  vars <- readxl::read_excel(file, sheet = vars.sheet)

  if (!"variable" %in% names(vars)) stop("There is no 'variable' column in dictionary", call. = FALSE)
  if (!"label" %in% names(vars)) stop("There is no 'label' column in dictionary", call. = FALSE)
  if (!"values" %in% names(vars)) stop("There is no 'values' column in dictionary", call. = FALSE)

  varnames <- names(data)

  for (var in vars$variable) {
    if (var %in% varnames) {
      varlab <- vars[vars$variable == var, "label", TRUE]
      if (is.na(varlab)) varlab <- NULL
      values <- vars[vars$variable == var, "values", TRUE]
      if (!is.na(values)) values <- unlist(jsonlite::fromJSON(values)) else values <- NULL
      data[[var]] <- haven::labelled(data[[var]], values, varlab)
      attr(data[[var]], "format.spss") <- "F8.0"
    } else {
      cat("Variable does not exist:", var, "\n")
    }
  }
  return(data)
}
