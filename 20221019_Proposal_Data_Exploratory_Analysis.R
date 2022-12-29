#Script para Análise Exploratória dos Dados de Projetos Vendidos COMAU

#O Objetivo desse script é fazer uma análise exploratórias dos dados dos
#principais projetos vendidos pelo departamento de Propostas de uma empresa
#de automação industrical com o intuito de identificar padrões e análisar 
#a possibilidade de desenvolver modelos de Machine Learning Preditivos

################
#Libraries
################
install.packages("tidyverse")
install.packages("cluster")
install.packages("dendextend")
install.packages("factoextra")
install.packages("fpc")
install.packages("gridExtra")
install.packages("readxl")
install.packeges("xlsx")


library(tidyverse) #Pacote utilziado para manipulação de dados
library(cluster) #Pacote utilizado para elaboração do algoritmo de cluster
library(dendextend) #Pacote utilizado para compar dendogramas
library(factoextra) #Pacote utilizado para elaboração do algoritmo de cluster e visualização
library(fpc) #Pacote utilizado para elaboração do algoritmo de cluster e visualização
library(gridExtra) #Pacote utilizado para usufluir da função grid arrange
library(readxl)# Pacote Utilizado para a importação de Banco de Dados
library("xlsx")#Pacote mara manipulação de tabelas no Excel

################
#Data Load
################

Sold_Projects_Comm_Distribution <- read.csv("Reference_Data/20221019_Sold_Projects_Comm_Distribution.csv", sep = ";", header = T, dec = ".") # Carga arquivo CSV
# O Arquivo CSV precisa estar com os percentuais na forma númerica e as casas decimais precisão ser "." ao invés de ","


novos_nomes <- c("Capform",
                 "Cliente",
                 "BU",
                 "Code",
                 'Name',
                 'Revisão',
                 'Status',
                 'Nome Maqueda',
                 'Considered #1 Results?',
                 'Tabela Maqueda',
                 'Proposal Type',
                 'Line Type',
                 'Month',
                 "Ano",
                 "DESIGN  &  GENERAL STUDIES",
                 "MATERIAL COMPONENTS",
                 "INSTALLATION & COMMISSIONING IN HOUSE",
                 "INSTALLATION & COMMISSIONING ON SITE",
                 "GENERAL ITEMS",
                 "TOTAL MANAGEMENT",
                 "GRAND  TOTAL (%)",
                 "GRAND  TOTAL",
                 "COST VALUE") # Definindo os Nomes das Variáveis

names(Sold_Projects_Comm_Distribution) <- novos_nomes # Renomeando as Variáveis


View(Sold_Projects_Comm_Distribution) #Visuaização Arquivo Carregado

#Para que possamos realizar a Clusterização é necessário manter apenas os dados numérico desta forma vamos eliminar os dados não númericos e que não são relevantes para a análise
rownames(Sold_Projects_Comm_Distribution) <- Sold_Projects_Comm_Distribution[,5]#Nomeando as Linhas
Sold_Projects_Comm_Distribution_Dist <- Sold_Projects_Comm_Distribution[,-1:-14]# Removendo as Colunas com a descrição das linhas
Sold_Projects_Comm_Distribution_Dist <- Sold_Projects_Comm_Distribution_Dist[,-7:-9]# Removendo Linhas com os totais


summary(Sold_Projects_Comm_Distribution)# Resumo da Tabela Adicionada

# # Precisamos mudar  as classes para números para começar a fazer a análise exploratória
# Sold_Projects_Comm_Distribution_Dist$`DESIGN  &  GENERAL STUDIES`<- as.numeric(Sold_Projects_Comm_Distribution_Dist$`DESIGN  &  GENERAL STUDIES`)
# Sold_Projects_Comm_Distribution_Dist$`MATERIAL COMPONENTS`<- as.numeric(Sold_Projects_Comm_Distribution_Dist$`MATERIAL COMPONENTS`)
# Sold_Projects_Comm_Distribution_Dist$`INSTALLATION & COMMISSIONING IN HOUSE`<- as.numeric(Sold_Projects_Comm_Distribution_Dist$`INSTALLATION & COMMISSIONING IN HOUSE`)
# Sold_Projects_Comm_Distribution_Dist$`INSTALLATION & COMMISSIONING ON SITE`<- as.numeric(Sold_Projects_Comm_Distribution_Dist$`INSTALLATION & COMMISSIONING ON SITE`)
# Sold_Projects_Comm_Distribution_Dist$`GENERAL ITEMS` <- as.numeric(Sold_Projects_Comm_Distribution_Dist$`GENERAL ITEMS`)
# Sold_Projects_Comm_Distribution_Dist$`TOTAL MANAGEMENT` <- as.numeric(Sold_Projects_Comm_Distribution_Dist$`TOTAL MANAGEMENT`)

summary(Sold_Projects_Comm_Distribution_Dist)# Resumo da Tabela Adicionada


write.xlsx(Sold_Projects_Comm_Distribution_Dist,file="20221019-Sold_Projects_Comm_Distribution_Cluster.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel
getwd()#Mostra onde o Arquivo Excel Gerado foi Salvo


#CALCULANDO MATRIZ DE DISTANCIAS
dist_matrix <- dist(Sold_Projects_Comm_Distribution_Dist, method = "euclidean")
dist_matrix


#APLICAÇÃO DE DIFERENTES MÉTODOS DE CLUSTERIZAÇÃO PARA A COMPARAÇÃO
#metodos disponiveis "average", "single", "complete" e "ward.D"
hc1 <- hclust(dist_matrix, method = "single" )# Vizinho mais próximo
hc2 <- hclust(dist_matrix, method = "complete" )# Vizinho mais distânte
hc3 <- hclust(dist_matrix, method = "average" )#Média
hc4 <- hclust(dist_matrix, method = "ward.D" )

#DESENHANDO O DENDOGRAMA
plot(hc1,
     main="Dendograma da Clusterização pelos Vizinhos mais Próximos", 
     xlab = "Projetos Vendidos",
     ylab = "Altura da Clusterização",
     cex = 0.6, hang = -1)

plot(hc2,
     main="Dendograma da Clusterização pelos Vizinhos mais Distantes", 
     xlab = "Projetos Vendidos",
     ylab = "Altura da Clusterização",
     cex = 0.6, hang = -1)

plot(hc3,
     main="Dendograma da Clusterização pelo Método das Médias", 
     xlab = "Projetos Vendidos",
     ylab = "Altura da Clusterização", cex = 0.6, hang = -1)


plot(hc4,
     main="Dendograma da Clusterização pelo Método de Ward.D", 
     xlab = "Projetos Vendidos",
     ylab = "Altura da Clusterização", cex = 0.6, hang = -1)


#VERIFICANDO ELBOW 
fviz_nbclust(Sold_Projects_Comm_Distribution_Dist, FUN = hcut, method = "wss")

#Gerando os plots com segmentação dos Clusters
#rect.hclust(hc4, k = 5)


#criando 6 grupos de ProjetosCada um dos Métodos
#cuttree corta o método de calculo em 6 grupos

grupo_Vz_Prox <- cutree(hc1, k = 5)
grupo_Vz_Dist <- cutree(hc2, k = 5)
grupo_Media <- cutree(hc3, k = 5)
grupo_Ward.D <- cutree(hc4, k = 5)


#transformando em data frame a saida do cluster
Proj_Grupos_DH <- data.frame(grupo_Vz_Prox,grupo_Vz_Dist,grupo_Media,grupo_Ward.D)

novos_nomes_cluster <- c("Vizinhos Mais Próximos",
                         "Vizinhos Mais Distante",
                         "Média",
                         "Ward.D") # Definindo os Nomes das Variáveis

names(Proj_Grupos_DH) <- novos_nomes_cluster  # Renomeando as Variáveis

View(Proj_Grupos_DH)
write.xlsx(Proj_Grupos_DH,file="20221019-Proj_Grupos_DH.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel

#Optamos pela metodolodia do Vizinho mais distante

#juntando com a base original
Sold_Projects_Class <- cbind(Sold_Projects_Comm_Distribution_Dist, grupo_Vz_Dist) #União entre bancos de dados

novos_nomes_class <- c("Engenharia",
                 "Materiais",
                 "Install_int",
                 "Install_Ext",
                 "Miscl",
                 "Gerenciamento",
                 "Cluster"
                 ) # Definindo os Nomes das Variáveis

names(Sold_Projects_Class) <- novos_nomes_class # Renomeando as Variáveis
Sold_Projects_Class


# entendendo os clusters
#FAZENDO ANALISE DESCRITIVA -> Utilizado biblioteca tidyverse
#MEDIAS das variaveis por grupo
Clusters_Analysis_Mean <- Sold_Projects_Class %>% 
  group_by(Cluster) %>% 
  summarise(n = n(),
            Engenharia = mean(Engenharia),
            Materiais = mean(Materiais),
            Install_int = mean(Install_int),
            Install_Ext = mean(Install_Ext),
            Miscl = mean(Miscl),
            Gerenciamento = mean(Gerenciamento),
            Engenharia = mean(Engenharia),
            )

Clusters_Analysis_Mean

novos_nomes_class_end <- c("Cluster",
                        "Qtd. Projetos",
                       "Média de Custo de Engenharia",
                       "Média de Custo de Materiais",
                       "Média de Custo Com Instalação Interna",
                       "Média de Custo Com Instalação Externa",
                       "Média de Custo Com Itens Gerais",
                       "Média de Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Mean) <- novos_nomes_class_end # Renomeando as Variáveis
Clusters_Analysis_Mean
Resultado_fim <- data.frame(Clusters_Analysis_Mean)
write.xlsx(Resultado_fim,file="20221019-Resultado_fim.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel

#Gerar Tabela dos valores máximos de cada Cluster
Clusters_Analysis_Max <- Sold_Projects_Class %>% 
  group_by(Cluster) %>% 
  summarise(n = n(),
            Engenharia = max(Engenharia),
            Materiais = max(Materiais),
            Install_int = max(Install_int),
            Install_Ext = max(Install_Ext),
            Miscl = max(Miscl),
            Gerenciamento = max(Gerenciamento),
            Engenharia = max(Engenharia),
  )

Clusters_Analysis_Max

novos_nomes_class_end_max <- c("Cluster",
                           "Qtd. Projetos",
                           "Máximo Custo de Engenharia",
                           "Máximo Custo de Materiais",
                           "Máximo Custo Com Instalação Interna",
                           "Máximo Custo Com Instalação Externa",
                           "Máximo Custo Com Itens Gerais",
                           "Máximo Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Max) <- novos_nomes_class_end_max # Renomeando as Variáveis
Clusters_Analysis_Max
Resultado_CH_Max<- data.frame(Clusters_Analysis_Max)
write.xlsx(Resultado_CH_Max,file="20221019-Resultado_CH_Max.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel


#Gerar Tabela dos valores minimo de cada Cluster
Clusters_Analysis_Min <- Sold_Projects_Class %>% 
  group_by(Cluster) %>% 
  summarise(n = n(),
            Engenharia = min(Engenharia),
            Materiais = min(Materiais),
            Install_int = min(Install_int),
            Install_Ext = min(Install_Ext),
            Miscl = min(Miscl),
            Gerenciamento = min(Gerenciamento),
            Engenharia = min(Engenharia),
  )

Clusters_Analysis_Min

novos_nomes_class_end_Min <- c("Cluster",
                               "Qtd. Projetos",
                               "Mínimo Custo de Engenharia",
                               "Mínimo Custo de Materiais",
                               "Mínimo Custo Com Instalação Interna",
                               "Mínimo Custo Com Instalação Externa",
                               "Mínimo Custo Com Itens Gerais",
                               "Mínimo Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Min) <- novos_nomes_class_end_Min # Renomeando as Variáveis
Clusters_Analysis_Min
Resultado_CH_Min<- data.frame(Clusters_Analysis_Min)
write.xlsx(Resultado_CH_Min,file="20221019-Resultado_CH_Min.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel



#### CLUSTERIZAÇÂO NÃO HIERARQUICA#####


#VERIFICANDO ELBOW 
fviz_nbclust(Sold_Projects_Comm_Distribution_Dist, kmeans, method = "wss")


#Rodar o modelo
Cluster_NH_k5 <- kmeans(Sold_Projects_Comm_Distribution_Dist, centers = 5)# Cosiderando 5 Clusters
Cluster_NH_k7 <- kmeans(Sold_Projects_Comm_Distribution_Dist, centers = 7)# Cosiderando 7 Clusters

#Visualizar os clusters
fviz_cluster(Cluster_NH_k5, data = Sold_Projects_Comm_Distribution_Dist, main = "Clusterização Não Hierárquica com 5 Clusters")
fviz_cluster(Cluster_NH_k7, data = Sold_Projects_Comm_Distribution_Dist, main = "Clusterização Não Hierárquica com 7 Clusters")

#transformando em data frame a saida do cluster NH
Proj_Grupos_CNH <- data.frame(Cluster_NH_k5["cluster"],Cluster_NH_k7["cluster"])

# Optou-se por seguir com a segmentação em 5 CLusters para CNH


#juntando com a base original
Sold_Projects_Class_NH <- cbind(Sold_Projects_Comm_Distribution_Dist, Cluster_NH_k5["cluster"]) #União entre bancos de dados

novos_nomes_class_NH <- c("Engenharia",
                       "Materiais",
                       "Install_int",
                       "Install_Ext",
                       "Miscl",
                       "Gerenciamento",
                       "Cluster_NH"
) # Definindo os Nomes das Variáveis

names(Sold_Projects_Class_NH) <- novos_nomes_class_NH # Renomeando as Variáveis
Sold_Projects_Class_NH


# entendendo os clusters
#FAZENDO ANALISE DESCRITIVA -> Utilizado biblioteca tidyverse
#MEDIAS das variaveis por grupo
Clusters_Analysis_Mean_NH <- Sold_Projects_Class_NH %>% 
  group_by(Cluster_NH) %>% 
  summarise(n = n(),
            Engenharia = mean(Engenharia),
            Materiais = mean(Materiais),
            Install_int = mean(Install_int),
            Install_Ext = mean(Install_Ext),
            Miscl = mean(Miscl),
            Gerenciamento = mean(Gerenciamento),
            Engenharia = mean(Engenharia),
  )

Clusters_Analysis_Mean_NH

novos_nomes_class_end_NH <- c("Cluster NH",
                           "Qtd. Projetos",
                           "Média de Custo de Engenharia",
                           "Média de Custo de Materiais",
                           "Média de Custo Com Instalação Interna",
                           "Média de Custo Com Instalação Externa",
                           "Média de Custo Com Itens Gerais",
                           "Média de Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Mean_NH) <- novos_nomes_class_end_NH # Renomeando as Variáveis
Clusters_Analysis_Mean_NH
Resultado_NH_mean <- data.frame(Clusters_Analysis_Mean_NH)
write.xlsx(Resultado_NH_mean,file="20221019-Resultado_NH_mean.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel

#Gerar Tabela dos valores máximos de cada Cluster NH
Clusters_Analysis_Max_NH <- Sold_Projects_Class_NH %>% 
  group_by(Cluster_NH) %>% 
  summarise(n = n(),
            Engenharia = max(Engenharia),
            Materiais = max(Materiais),
            Install_int = max(Install_int),
            Install_Ext = max(Install_Ext),
            Miscl = max(Miscl),
            Gerenciamento = max(Gerenciamento),
            Engenharia = max(Engenharia),
  )

Clusters_Analysis_Max_NH

novos_nomes_class_end_max_NH <- c("Cluster NH",
                               "Qtd. Projetos",
                               "Máximo Custo de Engenharia",
                               "Máximo Custo de Materiais",
                               "Máximo Custo Com Instalação Interna",
                               "Máximo Custo Com Instalação Externa",
                               "Máximo Custo Com Itens Gerais",
                               "Máximo Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Max_NH) <- novos_nomes_class_end_max_NH # Renomeando as Variáveis
Clusters_Analysis_Max_NH
Resultado_NH_Max<- data.frame(Clusters_Analysis_Max_NH)
write.xlsx(Resultado_NH_Max,file="20221019-Resultado_NH_Max.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel


#Gerar Tabela dos valores minimo de cada Cluster NH
Clusters_Analysis_Min_NH <- Sold_Projects_Class_NH %>% 
  group_by(Cluster_NH) %>% 
  summarise(n = n(),
            Engenharia = min(Engenharia),
            Materiais = min(Materiais),
            Install_int = min(Install_int),
            Install_Ext = min(Install_Ext),
            Miscl = min(Miscl),
            Gerenciamento = min(Gerenciamento),
            Engenharia = min(Engenharia),
  )

Clusters_Analysis_Min_NH

novos_nomes_class_end_Min_NH <- c("Cluster",
                               "Qtd. Projetos",
                               "Mínimo Custo de Engenharia",
                               "Mínimo Custo de Materiais",
                               "Mínimo Custo Com Instalação Interna",
                               "Mínimo Custo Com Instalação Externa",
                               "Mínimo Custo Com Itens Gerais",
                               "Mínimo Custo com Gerenciamento"
) # Definindo os Nomes das Variáveis

names(Clusters_Analysis_Min_NH) <- novos_nomes_class_end_Min_NH # Renomeando as Variáveis
Clusters_Analysis_Min_NH
Resultado_NH_Min<- data.frame(Clusters_Analysis_Min_NH)
write.xlsx(Resultado_NH_Min,file="20221019-Resultado_NH_Min.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel


# Criando Dataset com a comparação entre CH e CNH
Sold_Projects_Class_CH_CNH <- cbind(Sold_Projects_Comm_Distribution_Dist,grupo_Vz_Dist, Cluster_NH_k5["cluster"])
View(Sold_Projects_Class_CH_CNH)

novos_nomes_class_CH_CNH <- c("Engenharia",
                          "Materiais",
                          "Install_int",
                          "Install_Ext",
                          "Miscl",
                          "Gerenciamento",
                          "Cluster Hierárquica",
                          "Cluster Não Hierárquica"
) # Definindo os Nomes das Variáveis

names(Sold_Projects_Class_CH_CNH) <- novos_nomes_class_CH_CNH # Renomeando as Variáveis
write.xlsx(Sold_Projects_Class_CH_CNH,file="20221019-Sold_Projects_Class_CH_CNH.xlsx") #Exportar Dados para análise de Cluster em um Arquivo Excel
