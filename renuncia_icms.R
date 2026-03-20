
######################################################### renuncia ICMS ########

#googlesheets4::sheet_names("https://docs.google.com/spreadsheets/d/16j0h8IPX8tBjIQMnkXn6zHJ33qBlkRBX0SerzTlofBI")

googlesheets4::gs4_deauth()

arquivo <- googlesheets4::read_sheet(
  "https://docs.google.com/spreadsheets/d/16j0h8IPX8tBjIQMnkXn6zHJ33qBlkRBX0SerzTlofBI",
  sheet = "Sheet1"
)

arquivo <- arquivo |>
  dplyr::mutate(
    data_ano = as.integer(EXERCÍCIO),                # garante que ano seja numérico
    data_ano = lubridate::ymd(paste0(EXERCÍCIO, "-01-01")),
    .keep = "unused"
  )

# Using the particular produce decoder to adding more information in the novocaged

compilado_decodificador_endereço <-
  paste0("https://github.com/WillianDambros/data_source/raw/",
         "refs/heads/main/compilado_decodificador.xlsx")

decodificador_endereco <- paste0(getwd(), "/compilado_decodificador.xlsx")

curl::curl_download(compilado_decodificador_endereço,
                    decodificador_endereco)

"compilado_decodificador.xlsx" |> readxl::excel_sheets()

# reading, selecting and merging variables about a teme

cnae_sedec <- 
  readxl::read_excel("compilado_decodificador.xlsx",
                     sheet =  "cnae", col_types = "text") |> #dplyr::glimpse()
  dplyr::select("cnae_subclasse_codigo_7d_sem0",
                "cnae_secao_decodificado",
                "cnae_atividades_caracteristicas_turismo",
                "cnae_grande_grupamento_novocaged",
                "cnae_subclasse_decodificado")

arquivo <- arquivo |> 
  dplyr::left_join(cnae_sedec,
                   by = dplyr::join_by(CODG_CNAE ==
                                         cnae_subclasse_codigo_7d_sem0))


territorialidade_sedec <- 
  readxl::read_excel("compilado_decodificador.xlsx",
                     sheet =  "territorialidade_municipios_mt",
                     col_types = "text") |>
  dplyr::select("territorio_municipio_codigo_7d",
                "territorio_municipionovocaged_codigo_6d",
                "rpseplan10340_munícipio_polo_decodificado",
                "rpseplan10340_regiao_decodificado",
                "imeia_regiao",
                "imeia_municipios_polo_economico",
                "territorio_latitude", "territorio_longitude") |>
  dplyr::mutate(
    territorio_latitude =
      readr::parse_number(territorio_latitude,
                          locale = readr::locale(decimal_mark = ",")),
    territorio_longitude =
      readr::parse_number(territorio_longitude,
                          locale = readr::locale(decimal_mark = ","))
  )

arquivo <- arquivo |> 
  dplyr::left_join(territorialidade_sedec,
                   by = dplyr::join_by(CODG_IBGE ==
                                         territorio_municipio_codigo_7d))

arquivo |> dplyr::glimpse()
  
source("X:/POWER BI/NOVOCAGED/conexao.R")

RPostgres::dbListTables(conexao)

schema_name <- "sefaz_mt"

table_name <- "renuncia_icms"

DBI::dbSendQuery(conexao, paste0("CREATE SCHEMA IF NOT EXISTS ", schema_name))

RPostgres::dbWriteTable(conexao,
                        name = DBI::Id(schema = schema_name,
                                       table = table_name),
                        value = arquivo,
                        row.names = FALSE, overwrite = TRUE)

RPostgres::dbDisconnect(conexao)
