# Downloading archive ITCD_MUNICIPIO

endereco <- 
  paste0("https://www5.sefaz.mt.gov.br/documents/6071037/51381700/",
         "4-+Arrecada%C3%A7%C3%A3o+de+ITCD+Por+Munic%C3%ADpio.xlsx/",
         "db2308c7-23da-a20b-8a6c-4a4f309354c1?t=1695161963407")

arquivo_local <- paste0("Z:/rstudio/sefaz_mt/receita/",
                        "sefaz_mt_receita_itcd", ".xlsx")

curl::curl_download(endereco, arquivo_local)

# Transforming Microdata

arquivo_folhas <- readxl::excel_sheets(arquivo_local)

arquivo_vetor <- vector(mode = 'list', length = (length(arquivo_folhas)))

process_data <- function(entrada) {
  # read to define columns names
  arquivo_variaveis <-
    readxl::read_excel(arquivo_local, sheet = entrada, col_names = F,
                       col_types = c("text", "text", "text", "text", "date",
                                     "date", "date", "date", "date","date",
                                     "date", "date", "date", "date", "date",
                                     "date", "text"))
  # extracting column names

arquivo_variaveis <-
  if(!anyNA(arquivo_variaveis[5,])){arquivo_variaveis[5,]}else{
    if(!anyNA(arquivo_variaveis[6,])){arquivo_variaveis[6,]}else{
    if(!anyNA(arquivo_variaveis[7,])){arquivo_variaveis[7,]}else{
    }}}

#arquivo_variaveis_vetor  
  
  # creating list to store values of a vector
  arquivo_variaveis_vetor <- vector(length = ncol(arquivo_variaveis))
  # store properly the values
  for(i in seq_along(arquivo_variaveis)){
    arquivo_variaveis_vetor[i] <- as.character(arquivo_variaveis[[i]])
  }
  # reading data
  arquivo <- readxl::read_excel(arquivo_local, sheet = entrada,
                                col_names = F, col_types = "text")
  
  arquivo <- arquivo |> dplyr::rename_with(~arquivo_variaveis_vetor,
                                           .cols = 1:ncol(arquivo))
  arquivo <- arquivo |>
    dplyr::filter(!stringr::str_detect(`Código do Município`,
                                       "Código|Total|1.1.1.2|Fonte")) |>
    dplyr::select(!contains("Acumulado|TOTAL")) |>
    tidyr::pivot_longer(matches("\\d{4}-\\d{2}-\\d{2}"), names_to = "data_mes")
}

for(i in seq_along(arquivo_folhas)){
  
  tryCatch({
    arquivo_vetor[[i]] <- process_data(arquivo_folhas[i])
  }, error = function(err){warning("file not processed")})
  
}

arquivo_vetor <- arquivo_vetor |> dplyr::bind_rows()

sefaz_mt_receita_itcd <- arquivo_vetor |>
  dplyr::mutate(across(matches("value"), as.numeric))

# Writing novocaged

nome_arquivo_csv <- "sefaz_mt_receita_itcd"

caminho_arquivo <- paste0(getwd(), "/", nome_arquivo_csv, ".txt")

readr::write_csv2(sefaz_mt_receita_itcd,
                  caminho_arquivo)
