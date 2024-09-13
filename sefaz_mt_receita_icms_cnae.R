# Downloading arquive ICMS_CNAE

#### encontrar o bug da coluna que não sai, e remover os NA no R "" ou 0

icms_cnae_endereco <- 
  paste0("https://www5.sefaz.mt.gov.br/documents/6071037/51381700/",
         "1-+ICMS+por+CNAE+e+Grupo+CNAE.xlsx/",
         "0061779d-ce51-8852-fa19-94a757a5fad7?t=1695161941480")

nome_destino <- 
  paste0(getwd(),
         "/icms_cnae", ".xlsx")

curl::curl_download(icms_cnae_endereco, nome_destino)

# Transforming Microdata

icms_cnae_arquivo <- paste0(getwd(), "/icms_cnae", ".xlsx")

icms_cnae_folhas <- readxl::excel_sheets(icms_cnae_arquivo)

icms_cnae_vetor <- vector(mode = 'list', length = (length(icms_cnae_folhas)))

process_icms_cnae_data <- function(entrada) {
  # read to define columns names
  icms_cnae_names <-
    readxl::read_excel(icms_cnae_arquivo, sheet = entrada, col_names = F,
                       col_types = c("text", "text", "text", "text","text",
                                     "text","date", "date", "date", "date",
                                     "date","date", "date", "date", "date",
                                     "date", "date", "date", "text"))
  # extracting column names
  icms_cnae_names <- icms_cnae_names[5,]
  # creating list to store values of a vector
  icms_cnae_names_vetor <- vector(length = ncol(icms_cnae_names))
  # store properly the values
  for(i in seq_along(icms_cnae_names)){
    icms_cnae_names_vetor[i] <- as.character(icms_cnae_names[[i]])
  }
  # reading data
  icms_cnae <- readxl::read_excel(icms_cnae_arquivo, sheet = entrada,
                                  col_names = F, col_types = "text")
  
  icms_cnae <- icms_cnae |> dplyr::rename_with(~icms_cnae_names_vetor,
                                               .cols = 1:ncol(icms_cnae))
  icms_cnae <- icms_cnae |>
    dplyr::filter(!stringr::str_detect(SEÇÃO, "SEÇÃO|TOTAL|FONTE|Obs|Soma|
                                       Circulação|Adicional|Impostos")) |>
    dplyr::select(!TOTAL) |>
    tidyr::pivot_longer(matches("\\d{4}-\\d{2}-\\d{2}"), names_to = "data_mes")
}

for(i in seq_along(icms_cnae_folhas)){
  
  tryCatch({
    icms_cnae_vetor[[i]] <- process_icms_cnae_data(icms_cnae_folhas[i])
  }, error = function(err){warning("file not processed")})
  
}

sefaz_icms_cnae <- icms_cnae_vetor |> dplyr::bind_rows()

sefaz_icms_cnae <- sefaz_icms_cnae |>
  dplyr::mutate(across(matches("value"), as.numeric))

compilado_decodificador_endereço <-
  paste0("https://github.com/WillianDambros/data_source/",
         "raw/main/compilado_decodificador.xlsx")

decodificador_endereco <- paste0(getwd(), "/compilado_decodificador.xlsx")

curl::curl_download(compilado_decodificador_endereço,
                    decodificador_endereco)

decodificador_cnae <- readxl::read_xlsx(decodificador_endereco, sheet = "cnae",
                                        col_types = "text")

sefaz_icms_cnae <- sefaz_icms_cnae |>
  dplyr::left_join(decodificador_cnae,
                   by = dplyr::join_by(
                     SUBCLASSE == cnae_subclasse_codigo_7d_sem0)) |>
  dplyr::select(!matches(
    paste0("cnae_secao_codigo_sigla1d|cnae_divisao_codigo",
           "_2d|cnae_grupo_codigo_3d|cnae_classe_codigo_5d")))

sefaz_icms_cnae <- sefaz_icms_cnae |> dplyr::mutate(
  SEÇÃO = dplyr::case_when(SEÇÃO == "OUTROS CNAE (*)" ~ "OUTROS CNAE",
                           TRUE ~ SEÇÃO))

# Writing novocaged

#nome_arquivo_csv <- "sefaz_mt_receita_icms_cnae"

#caminho_arquivo <- paste0(getwd(), "/", nome_arquivo_csv, ".csv")

#readr::write_csv2(sefaz_icms_cnae,
#                  caminho_arquivo)

# writing PostgreSQL

conexao <- RPostgres::dbConnect(RPostgres::Postgres(),
                                dbname = "observatorio_db",
                                host = "10.43.88.8",
                                port = "5502",
                                user = "admin",
                                password = "adminadmin")

RPostgres::dbListTables(conexao)

schema_name <- "sefaz_mt"

table_name <- "sefaz_mt_receita_icms_cnae"

DBI::dbSendQuery(conexao, paste0("CREATE SCHEMA IF NOT EXISTS ", schema_name))

RPostgres::dbWriteTable(conexao,
                        name = DBI::Id(schema = schema_name,table = table_name),
                        value = sefaz_icms_cnae,
                        row.names = FALSE, overwrite = TRUE)

RPostgres::dbDisconnect(conexao)
