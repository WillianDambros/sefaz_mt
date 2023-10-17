# Downloading archive NATUREZA_RECEITA

endereco <- 
  paste0("https://www5.sefaz.mt.gov.br/documents/6071037/28918547/",
  "5-+Arrecada%C3%A7%C3%A3o+por+Natureza+de+Receita.xlsx/dcb91ddd-ab08-7700",
  "-a78b-93f808b19c9e?t=1697227113306")

arquivo_local <- paste0("Z:/rstudio/sefaz_mt/receita/",
                        "sefaz_mt_receita_natureza", ".xlsx")

curl::curl_download(endereco, arquivo_local)

# Transforming Microdata

#icms_cnae_arquivo <- nome_arquivo_destino

arquivo_folhas <- readxl::excel_sheets(arquivo_local)

arquivo_folhas <- arquivo_folhas[arquivo_folhas != "TOTAIS-ANUAIS"]

arquivo_vetor <- vector(mode = 'list', length = (length(arquivo_folhas)))

process_data <- function(entrada) {
  # read to define columns names
  arquivo_variaveis <-
    readxl::read_excel(arquivo_local, sheet = entrada, col_names = F,
                       col_types = c("text", "text", "date",
                                     "date", "date", "date", "date","date",
                                     "date", "date", "date", "date", "date",
                                     "date", "text"))
  # extracting column names
  
  arquivo_variaveis <-  arquivo_variaveis[7,]
  
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
  arquivo
  arquivo <- arquivo |>
    dplyr::filter(!stringr::str_detect(
      `CÓDIGO_DO_ATUAL`,
      "FONTE|Não|Rubricas|Sujeitas|natureza|disposto|ATUAL")) |>
    dplyr::select(!`TOTAL ANO`) |>
    tidyr::pivot_longer(matches("\\d{4}-\\d{2}-\\d{2}"), names_to = "data_mes")
}

for(i in seq_along(arquivo_folhas)){
  
  tryCatch({
    arquivo_vetor[[i]] <- process_data(arquivo_folhas[i])
  }, error = function(err){warning("file not processed")})
  
}

arquivo_vetor

arquivo_vetor <- arquivo_vetor |> dplyr::bind_rows()

arquivo <- arquivo_vetor |>
  dplyr::mutate(across(matches("value"), as.numeric))

# Writing novocaged

nome_arquivo_csv <- "sefaz_mt_receita_natureza"

caminho_arquivo <- paste0(getwd(), "/", nome_arquivo_csv, ".txt")

readr::write_csv2(arquivo,
                  caminho_arquivo)
