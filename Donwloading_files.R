
library("xlsx")

main <- function(){
  matrix <- get_info()
  
  iterator(matrix)

}

# Função para pegar as informações referentes aos números dos chamados e links para download a partir
# De uma planilha 
get_info <- function(){
  read.xlsx("C:/Users/Paulo Henrique/Downloads/Download.xlsm", sheetIndex = 1, header = FALSE)
}

#Função para iterar dentro da matriz  e ir fazendo o download do arquivo 
iterator <- function(matriz) { 
  for (i in 1:length(matriz[,1])){
    zipfile <- download(matriz[i,2])
    
    zipdir <- tempfile() 
    
    dir.create(zipdir)
    
    unzip(zipfile, exdir=zipdir)
    
    file <- list.files(zipdir)
    
    for (j in 1:length(file)){
      from <- paste(zipdir, "/", file[j], sep = "")
      to <- paste ("C:/Users/Paulo Henrique/Desktop/Doispontocinco", "/",matriz[i,1], "_", file[j], sep = "")
      file.copy(from = from , to  = to)  
    }
    unlink(zipfile)
  }
}

#Função para realizar o download do arquivo
download <- function (fileurl,destfile = "C:/Users/Paulo Henrique/Desktop/Doispontocinco"){
  destfile <- paste(destfile, "/arquivo.zip", sep = "")
  download.file(fileurl, destfile = destfile )

  destfile
}

main()



