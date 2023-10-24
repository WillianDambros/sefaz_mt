# Downloading archive IPVA_MUNICIPIO

ipva_endereco <- 
  paste0("https://www5.sefaz.mt.gov.br/documents/6071037/51381700/",
         "3-+Arrecada%C3%A7%C3%A3o+de+IPVA+por+Munic%C3%ADpio.xlsx/",
         "b54d1947-4675-b8e3-1abb-b3565d143ff6?t=1695161959447")

arquivo_local <- paste0("Z:/rstudio/sefaz_mt/receita/",
                        "sefaz_mt_receita_ipva", ".xlsx")

curl::curl_download(ipva_endereco, arquivo_local)

# Transforming Microdata

#icms_cnae_arquivo <- nome_arquivo_destino

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
  
  # creating list to store values of a vector
  arquivo_variaveis_vetor <- vector(length = ncol(arquivo_variaveis))
  # store properly the values
  for(i in seq_along(arquivo_variaveis)){
    arquivo_variaveis_vetor[i] <- as.character(arquivo_variaveis[[i]])
  }
  # reading data
  arquivo <- readxl::read_excel(arquivo_local, sheet = arquivo_folhas[9],
                                col_names = F, col_types = "text")
  
  arquivo <- arquivo |> dplyr::rename_with(~arquivo_variaveis_vetor,
                                           .cols = 1:ncol(arquivo))
  arquivo <- arquivo |>
    dplyr::filter(!stringr::str_detect(`Código do Município`,
                                       "Código|Total|1.1.1.2|Fonte")) |>
    dplyr::select(!matches("Acumulado|TOTAL")) |>
    tidyr::pivot_longer(matches("\\d{4}-\\d{2}-\\d{2}"), names_to = "data_mes")
}

for(i in seq_along(arquivo_folhas)){
  
  tryCatch({
    arquivo_vetor[[i]] <- process_data(arquivo_folhas[i])
  }, error = function(err){warning("file not processed")})
  
}

arquivo_vetor

arquivo_vetor <- arquivo_vetor |> dplyr::bind_rows()

sefaz_mt_receita_ipva <- arquivo_vetor |>
  dplyr::mutate(across(matches("value"), as.numeric))

# Writing novocaged

nome_arquivo_csv <- "sefaz_mt_receita_ipva"

caminho_arquivo <- paste0(getwd(), "/", nome_arquivo_csv, ".txt")

readr::write_csv2(sefaz_mt_receita_ipva,
                  caminho_arquivo)
