###########################################################
## SCRIPT PARA CRIAÇÃO DE CONTAS GOOGLE PARA PROFESSORES ##
##             JAYCKSON AMORIM - Junho/2022              ##
###########################################################

################################
## PARÂMETROS DE CONFIGURAÇÃO ##

# Nome do arquivo CSV ou endereço da publicação em CSV da base de dados do Google
nome_google <- "https://docs.google.com/spreadsheets/d/e/2PACX-1vScx2jZ9h9PxLcFebtGsQplCaSLhK8l3fh4A9xgGErc4pDAnh63mIaHqnokok52Pe1nghpq5FxAnA-t/pub?gid=1542627827&single=true&output=csv"

#Unidade organizacional onde serão criados os usuários
uo <-  "/ESCOLAS/PROFESSORES"

# Nome do arquivo da planilha da solicitação com as colunas: Matricula|Nome|Sobrenome|E-mail pessoal|Celular|CPF|Escola|Distrito (Devem estar tratados e em formato Tidy Data)
# Será aberta uma caixa de diálogo ao rodar o script para escolher o arquivo

################################

## BIBLIOTECAS UTILIZADAS ##
library(stringr)
library(readxl)
library(dplyr)
library(writexl)
library(tcltk)
############################

## BASE DO GOOGLE ##
base_google <- read.csv2(nome_google, sep = ",", encoding = "UTF-8")
names(base_google)[1:4] <- c("nome","sobrenome","e-mail","ultimo_acesso")
###################

## PLANILHA DE SOLICITAÇÃO ##
solicitacao <- read_xlsx(file.choose("Abrir solicitação"))
names(solicitacao)[1:8] <- c("matricula","nome","sobrenome","e-mail_pessoal","celular","cpf","escola","distrito")
#############################

## Função para remover acentos e caracteres especiais e converter para minúsculas
remove_acento <- function(texto,padrao="todos") {
  texto <- tolower(texto)
  if(!is.character(texto))
    texto <- as.character(texto)
  
  padrao <- unique(padrao)
  
  if(any(padrao=="Ç"))
    padrao[padrao=="Ç"] <- "ç"
  
  carac_esp <- c(
    acentos = "áéíóúÁÉÍÓÚýÝ",
    crases = "àèìòùÀÈÌÒÙ",
    circunflexos = "âêîôûÂÊÎÔÛ",
    tils = "ãõÃÕñÑ",
    tremas = "äëïöüÄËÏÖÜÿ",
    cedilhas = "çÇ"
  )
  
  carac_norm <- c(
    acentos = "aeiouAEIOUyY",
    crases = "aeiouAEIOU",
    circunflexos = "aeiouAEIOU",
    tils = "aoAOnN",
    tremas = "aeiouAEIOUy",
    cedilhas = "cC"
  )
  
  tipos_acentos <- c("´","`","^","~","¨","ç")
  
  if(any(c("all","al","a","todos","t","to","tod","todo")%in%padrao)) # opcao retirar todos
    return(chartr(paste(carac_esp, collapse=""), paste(carac_norm, collapse=""), texto))
  
  for(i in which(tipos_acentos%in%padrao))
    texto <- chartr(carc_esp[i],carac_norm[i], texto)
  
  
  return(texto)
}
#################################################
=
## FUNÇÃO PARA REMOVER PREPOSIÇÕES
remove_preposicao <- function(texto){
  arraynomes <- str_split(texto," ", simplify = TRUE)
  ### Preposições para remover
  preposicoes <- c("de","da","do","das","dos","e","")
  
  
  for(i in 1:length(preposicoes)){
    da_vez <- preposicoes[i]
    posicao <- match(da_vez,arraynomes, nomatch = 0)
    while(posicao>0){
      arraynomes <- arraynomes[-posicao]
      posicao <- match(da_vez,arraynomes, nomatch = 0)
    }
    
  }
  return(arraynomes)
}
########################################

## FUNÇÃO PARA GERAR SENHA
gera_senha <- function(){
  caracteres <- c("0","1","2","3","4","5","6","7","8","9",
                  "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
                  "@","#","$","%","&","*",">","<")
  senha <- caracteres[round(runif(1, min = 1, max = length(caracteres)))]
  for(i in 1:7){
    senha <- paste(senha,caracteres[round(runif(1, min = 1, max = length(caracteres)))], sep = "")
  }
  return(senha)
}

##########################


## VALIDAÇÃO DE E-MAILS
pessoa <- paste(solicitacao$nome, solicitacao$sobrenome)
Pessoa_sem_acentos <- remove_acento(pessoa)
solicitacao$email_criado <- NA
solicitacao$senha <- NA
for (i in 1:length(solicitacao$nome)) {
  ## MOTANDO PREFIXO
  array_pessoa <-remove_preposicao(Pessoa_sem_acentos[i])
  prefixo <- paste(array_pessoa[1],".",array_pessoa[2], sep = "")
  email_criado <- paste(prefixo,"@educacao.fortaleza.ce.gov.br",sep = "")
  
  ## EVITANDO DUPLICIDADE DE E-MAIL
  sequencial=0
  while(email_criado %in% base_google$`e-mail`){
    sequencial <- sequencial+1
    email_criado <- paste(prefixo,sequencial,"@educacao.fortaleza.ce.gov.br", sep = "")
  }
  solicitacao$email_criado[i] <- email_criado
  solicitacao$senha[i]<- gera_senha()

}

#########################
## CRIANDO OS ARQUIVOS ##
#########################

## PLANILHA DE RESPOSTA
nome_arquivo <- paste("arquivos/",format.Date(Sys.Date(),"%Y-%m-%d")," arquivo de resposta.xlsx",sep = "")
arq_resposta <- select(solicitacao,"matricula","nome","sobrenome","e-mail_pessoal","celular","cpf","escola","distrito","email_criado","senha")
names(arq_resposta)[1:10] <- c("Matrícula","Nome","Sobrenome","E-mail pessoal","Celular","CPF","Unidade escolar","Distrito","E-mail criado","Senha provisória")
write_xlsx(arq_resposta,nome_arquivo)

## Arquivo CSV para upload e XLSX para conferência
nome_arquivo <- paste("arquivos/",format.Date(Sys.Date(),"%Y-%m-%d")," arquivo para upload.csv",sep = "")
arq_upload <- select(solicitacao,"nome","sobrenome","email_criado","senha","e-mail_pessoal","celular","cpf","distrito","escola")
arq_upload$unidade <- uo
arq_upload$solicita_senha <- "TRUE"
arq_upload <- select(arq_upload,"nome","sobrenome","email_criado","senha","unidade","e-mail_pessoal","celular","cpf","distrito","escola","solicita_senha")
names(arq_upload)[1:11] <- c("First Name [Required]","Last Name [Required]","Email Address [Required]","Password [Required]","Org Unit Path [Required]","Recovery Email","Recovery Phone [MUST BE IN THE E.164 FORMAT]","Employee ID","Department","Cost Center","Change Password at Next Sign-In")
write.csv(arq_upload,nome_arquivo, row.names = FALSE)
nome_arquivo <- paste("arquivos/",format.Date(Sys.Date(),"%Y-%m-%d")," arquivo para conferência.xlsx",sep = "")
write_xlsx(arq_upload,nome_arquivo)
#######################

## MENSAGEM FINAL
mensagem <- tkmessageBox(title="FINAL DO SCRIPT",
                        message="Arquivos criados na pasta de arquivos do projetos!",
                        icon="info",
                        type="ok")
