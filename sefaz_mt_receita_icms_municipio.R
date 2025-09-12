# Downloading archive ICMS_MUNICIPIO

icms_municipio_endereco <- 
  paste0("https://docs.google.com/spreadsheets/",
         "d/1DbcU8-jdlG_DrQyVlhw7jnFXWug_Rx5l/",
         "export?format=xlsx")

arquivo_local <- paste0(getwd(),
                        "/sefaz_mt_receita_icms_municipio", ".xlsx")

curl::curl_download(icms_municipio_endereco, arquivo_local)

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
        }}} |>
    dplyr::mutate(...1 = stringr::str_replace_all(...1,"[íìîïi]",""))
  
  # corrigindo bug folhas 09:11
  
  arquivo_variaveis <- arquivo_variaveis  |>
    dplyr::mutate(...1 = stringr::str_replace_all(...1,"[íìîïi]",""))
  
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
    dplyr::filter(!stringr::str_detect(`COD Muncpo`,
                                       "COD|Total|Fonte|Impostos|ICMS")) |>
    dplyr::select(!matches("Acumulado|TOTAL")) |>
    tidyr::pivot_longer(matches("\\d{4}-\\d{2}-\\d{2}"), names_to = "data_mes")
}

for(i in seq_along(arquivo_folhas)){
  
  tryCatch({
    arquivo_vetor[[i]] <- process_data(arquivo_folhas[i])
  }, error = function(err){warning("file not processed")})
  
}

arquivo_vetor <- arquivo_vetor |> dplyr::bind_rows()

sefaz_mt_receita_icms_municipio <- arquivo_vetor |>
  dplyr::mutate(across(matches("value"), as.numeric))

sefaz_mt_receita_icms_municipio <- sefaz_mt_receita_icms_municipio |>
  dplyr::mutate(data_mes = lubridate::ymd(data_mes))

sefaz_mt_receita_icms_municipio <- sefaz_mt_receita_icms_municipio |> 
dplyr::filter(!stringr::str_detect(`COD Muncpo`, "TOTAL"))


sefaz_mt_receita_icms_municipio |> dplyr::glimpse()
sefaz_mt_receita_icms_municipio$`COD Muncpo` |> unique()

####################################### Compilado_Decodificador ################

compilado_decodificador_endereço <-
  paste0("https://github.com/WillianDambros/data_source/raw/",
         "refs/heads/main/compilado_decodificador.xlsx")

decodificador_endereco <- paste0(getwd(), "/compilado_decodificador.xlsx")

curl::curl_download(compilado_decodificador_endereço,
                    decodificador_endereco)

"compilado_decodificador.xlsx" |> readxl::excel_sheets()

territorialidade_sedec <- 
  readxl::read_excel("compilado_decodificador.xlsx",
                     sheet =  "territorialidade_municipios_mt",
                     col_types = "text") |>
  dplyr::select("territorio_geo_munícipios",
                "rpseplan10340_regiao_decodificado",
                "imeia_regiao",
                "territorio_latitude", "territorio_longitude")

territorialidade_sedec <- territorialidade_sedec |>
  dplyr::rename(municipios_decodificado = territorio_geo_munícipios)


territorialidade_sedec <- territorialidade_sedec |>
  dplyr::mutate(
    municipios_decodificado = toupper(municipios_decodificado),
    territorio_latitude = as.numeric(gsub(",", ".", territorio_latitude)),
    territorio_longitude = as.numeric(gsub(",", ".", territorio_longitude))
  )

territorialidade_sedec <- territorialidade_sedec |>
  dplyr::mutate(municipios_decodificado =
                  gsub("[^[:alnum:] ]", "", iconv(municipios_decodificado,
                                                  to = "ASCII//TRANSLIT")))
territorialidade_sedec |> dplyr::glimpse()

territorialidade_sedec <- territorialidade_sedec |>
  dplyr::mutate(
    municipios_decodificado = toupper(municipios_decodificado),
    territorio_latitude = as.numeric(gsub(",", ".", territorio_latitude)),
    territorio_longitude = as.numeric(gsub(",", ".", territorio_longitude))
  )                                              

territorialidade_sedec |> dplyr::glimpse()

territorialidade_sedec <- territorialidade_sedec |>
  dplyr::mutate(
    municipios_decodificado = dplyr::recode(
      municipios_decodificado,
      "NOSSA SENHORA DO LIVRAMENTO" = "NOSSA SRA DO LIVRAMENTO",
        "POXOREU" = "POXOREO"
    )
  )

sefaz_mt_receita_icms_municipio <- sefaz_mt_receita_icms_municipio |>
  fuzzyjoin::stringdist_left_join(
    territorialidade_sedec,
    by = c("Nome do Município" = "municipios_decodificado"),
    method = "jw",          # Jaro-Winkler
    max_dist = 0.05,        # tolerância de distância
    distance_col = "dist"   # coluna com a distância calculada
  )
sefaz_mt_receita_icms_municipio |> dplyr::glimpse()
################################################## Estabelecimentos ############


# Writing novocaged

#nome_arquivo_csv <- "sefaz_mt_receita_icms_municipio"

#caminho_arquivo <- paste0(getwd(), "/", nome_arquivo_csv, ".txt")

#readr::write_csv2(sefaz_mt_receita_icms_municipio, caminho_arquivo)


# writing PostgreSQL

source("X:/POWER BI/NOVOCAGED/conexao.R")

RPostgres::dbListTables(conexao)

schema_name <- "sefaz_mt"

table_name <- "sefaz_mt_receita_icms_municipio"

DBI::dbSendQuery(conexao, paste0("CREATE SCHEMA IF NOT EXISTS ", schema_name))

RPostgres::dbWriteTable(conexao,
                        name = DBI::Id(schema = schema_name,table = table_name),
                        value = sefaz_mt_receita_icms_municipio,
                        row.names = FALSE, overwrite = TRUE)

RPostgres::dbDisconnect(conexao)
