# Instalação de pacotes (apenas necessário uma vez)
# Ferramentas para ler, manipular e salvar dados
install.packages("tidyverse")   # Para manipulação de dados
install.packages("writexl")     # Para escrever arquivos Excel
install.packages("data.table")  # Para trabalhar com tabelas grandes
install.packages("readxl")      # Para ler arquivos Excel
install.packages("httr")        # Para baixar arquivos da internet
install.packages("openxlsx")    # Para criar arquivos Excel

# Carregando os pacotes instalados
# Torna as ferramentas disponíveis para uso
library(tidyverse)
library(writexl)
library(data.table)
library(readxl)
library(httr)
library(openxlsx)


# ---- SUBSTITUINDO CÓDIGOS POR DESCRIÇÕES LEGÍVEIS ----

# Baixando dados do Novo CAGED de 2025 para a ocupação Arqueólogo
url <- "https://raw.githubusercontent.com/machadotm/Novos_Dados_MTE_Arqueo/refs/heads/main/2025/Novo_CAGED/Dados_Brutos_Novo%20CAGED_2025.csv"

Arqueo_CAGED_2025 <- read_csv(url)  # Lê o arquivo CSV da internet

# Separando a coluna de competência (formato AAAAMM) em ano e mês
Arqueo_CAGED_2025 <- Arqueo_CAGED_2025 %>%
  mutate(anocompetênciamov = str_sub(as.character(competênciamov), 1, 4)) %>%  # Pega os 4 primeiros dígitos (ano)
  mutate(mêscompetênciamov = str_sub(as.character(competênciamov), 5, 6)) %>%    # Pega os 2 últimos dígitos (mês)
  relocate(c(anocompetênciamov,mêscompetênciamov), .before = competênciamov) %>%   # Reorganiza as colunas
  dplyr::select(-competênciamov)  # Remove a coluna original de competência

# Criando uma tabela para traduzir números de mês para nomes dos meses
mes <- data.frame(mes=c(
  "Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"),
  mêscompetênciamov= c("01","02","03","04","05","06","07","08","09","10","11","12")
)

# Substituindo os números dos meses pelos nomes
Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, mes, by = 'mêscompetênciamov') %>% 
  mutate(mêscompetênciamov = mes) %>% 
  dplyr::select(-mes)

# Baixando o dicionário de dados do Novo CAGED (contém as traduções de códigos)
url2 <- "https://github.com/machadotm/Dados_MTE_Arqueo/raw/refs/heads/main/CAGED/3-Novo_Caged/Dicionario_Dados/Layout_Novo_Caged.xlsx"
tmp <- tempfile(fileext = ".xlsx")  # Cria um arquivo temporário
GET(url2, write_disk(tmp))  # Baixa o dicionário para o arquivo temporário

# Traduzindo a coluna de região (códigos para nomes das regiões brasileiras)
regiao <- read_excel(tmp, 
                     sheet = "região") %>% 
  rename(região= Código)  # Renomeia a coluna Código para região (para fazer o join)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, regiao, by = 'região') %>% 
  mutate(região = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de UF (Unidade Federativa)
uf <- read_excel(tmp, 
                 sheet = "uf") %>% 
  rename(uf= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, uf, by = 'uf') %>% 
  mutate(uf = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de município
municipios <- read_excel(tmp, 
                         sheet = "município") %>% 
  rename(município= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, municipios, by = 'município') %>% 
  mutate(município = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de seção CNAE (classificação econômica)
secao <- read_excel(tmp, 
                    sheet = "seção") %>% 
  rename(seção= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, secao, by = 'seção') %>% 
  mutate(seção = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de subclasse CNAE (classificação econômica mais detalhada)
subclasse <- read_excel(tmp, 
                        sheet = "subclasse") %>% 
  rename(subclasse= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, subclasse, by = 'subclasse') %>% 
  mutate(subclasse = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de saldo de movimentação (admissão/desligamento)
saldomov <- data.frame(situacao=c("Admitido","Desligado"),
                       saldomovimentação=c(1,-1)
)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025,saldomov,by='saldomovimentação') %>% 
  mutate(saldomovimentação= situacao) %>% 
  dplyr::select(-situacao)

# Traduzindo a coluna de CBO (Classificação Brasileira de Ocupações)
cbo <- read_excel(tmp, 
                  sheet = "cbo2002ocupação") %>% 
  rename(cbo2002ocupação= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, cbo, by = 'cbo2002ocupação') %>% 
  mutate(cbo2002ocupação = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de categoria (Categoria de trabalhador)
categoria <- read_excel(tmp, 
                        sheet = "categoria") %>% 
  rename(categoria= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, categoria, by = 'categoria') %>% 
  mutate(categoria = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de grau de instrução (escolaridade)
escolaridade <- read_excel(tmp, 
                           sheet = "graudeinstrução") %>% 
  rename(graudeinstrução= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, escolaridade, by = 'graudeinstrução') %>% 
  mutate(graudeinstrução = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de raça/cor
racacor <- read_excel(tmp, 
                      sheet = "raçacor") %>% 
  rename(raçacor= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, racacor, by = 'raçacor') %>% 
  mutate(raçacor = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de sexo
sexo <- read_excel(tmp, 
                   sheet = "sexo") %>% 
  rename(sexo= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, sexo, by = 'sexo') %>% 
  mutate(sexo = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de tipo de empregador
empregador <- read_excel(tmp, 
                         sheet = "tipoempregador") %>% 
  rename(tipoempregador= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, empregador, by = 'tipoempregador') %>% 
  mutate(tipoempregador = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de tipo de estabelecimento
tipoestab <- read_excel(tmp, 
                        sheet = "tipoestabelecimento") %>% 
  rename(tipoestabelecimento= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, tipoestab, by = 'tipoestabelecimento') %>% 
  mutate(tipoestabelecimento = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de tipo de movimentação (detalhes da admissão/desligamento)
tipomov <- read_excel(tmp, 
                      sheet = "tipomovimentação") %>% 
  rename(tipomovimentação= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, tipomov, by = 'tipomovimentação') %>% 
  mutate(tipomovimentação = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de tipo de deficiência (para pessoas com deficiência)
tipodef <- read_excel(tmp, 
                      sheet = "tipodedeficiência") %>% 
  rename(tipodedeficiência= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, tipodef, by = 'tipodedeficiência') %>% 
  mutate(tipodedeficiência = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de indicador de trabalho intermitente
indtrabinter <- read_excel(tmp, 
                           sheet = "indtrabintermitente") %>% 
  rename(indtrabintermitente= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, indtrabinter, by = 'indtrabintermitente') %>% 
  mutate(indtrabintermitente = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de indicador de trabalho parcial
indtrabparcial <- read_excel(tmp, 
                             sheet = "indtrabparcial") %>% 
  rename(indtrabparcial= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, indtrabparcial, by = 'indtrabparcial') %>% 
  mutate(indtrabparcial = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de tamanho do estabelecimento em janeiro
tambestab <- read_excel(tmp, 
                        sheet = "tamestabjan") %>% 
  rename(tamestabjan= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, tambestab, by = 'tamestabjan') %>% 
  mutate(tamestabjan = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de indicador de aprendiz (jovem aprendiz)
aprendiz <- read_excel(tmp, 
                       sheet = "indicadoraprendiz") %>% 
  rename(indicadoraprendiz= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, aprendiz, by = 'indicadoraprendiz') %>% 
  mutate(indicadoraprendiz = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de origem da informação (como o dado foi obtido)
origeminfo <- read_excel(tmp, 
                         sheet = "origemdainformação") %>% 
  rename(origemdainformação= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, origeminfo, by = 'origemdainformação') %>% 
  mutate(origemdainformação = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de indicador de exclusão (se o registro foi excluído ou não)
indexclui <- read_excel(tmp, 
                        sheet = "indicadordeexclusão") %>% 
  rename(indicadordeexclusão= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, indexclui, by = 'indicadordeexclusão') %>% 
  mutate(indicadordeexclusão = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de indicador de fora do prazo (se a declaração foi feita fora do prazo)
indforaprazo <- read_excel(tmp, 
                           sheet = "indicadordeforadoprazo") %>% 
  rename(indicadordeforadoprazo= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, indforaprazo, by = 'indicadordeforadoprazo') %>% 
  mutate(indicadordeforadoprazo = Descrição) %>% 
  dplyr::select(-Descrição)

# Traduzindo a coluna de unidade de salário (código da unidade salarial)
unidsalario <- read_excel(tmp, 
                          sheet = "unidadesaláriocódigo") %>% 
  rename(unidadesaláriocódigo= Código)

Arqueo_CAGED_2025 <- left_join(Arqueo_CAGED_2025, unidsalario, by = 'unidadesaláriocódigo') %>% 
  mutate(unidadesaláriocódigo = Descrição) %>% 
  dplyr::select(-Descrição)


# SALVANDO OS RESULTADOS FINAIS EM UM ARQUIVO EXCEL


# Cria um novo arquivo Excel
wb <- createWorkbook()
nome_da_aba <- "Resultado Novo CAGED 2025" # Nome da aba (planilha) que será criada
addWorksheet(wb, nome_da_aba) # Adiciona uma planilha ao arquivo
# Escreve os dados tratados na planilha, com filtros ativados
writeData(wb, sheet = nome_da_aba, x = Arqueo_CAGED_2025, withFilter = TRUE)
# Ajusta a largura das colunas automaticamente
setColWidths(wb, sheet = nome_da_aba, cols = 1:ncol(Arqueo_CAGED_2025), widths = "auto")
# Salva o arquivo Excel com o nome especificado
saveWorkbook(wb, "Novo_CAGED_2025_Arqueo_Dados_Tratados.xlsx", overwrite = TRUE)

