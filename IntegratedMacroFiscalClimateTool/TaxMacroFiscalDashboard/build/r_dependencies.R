pkgs <- c("data.table","readxl","writexl","haven")
missing <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing)) install.packages(missing) else cat("all R dependencies present\n")
