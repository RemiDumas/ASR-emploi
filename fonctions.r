# Installation et chargement de packages
install_and_load <- function(P) {
  Pi <- P[!(P %in% installed.packages()[,"Package"])]
  
  if (length(Pi)>0) install.packages(Pi)
  for(i in P) library(i,character.only = TRUE)
}

packages <-
  c("downloader",
    "glue",
    "httr",
    "lubridate"  ,
    "sf",
    "readxl",
    "rio"   ,
    "rlang"  ,
    "showtext"  ,
    "tidyverse",
    "RColorBrewer",
    "ggrepel",
    "ggthemes",
    "scales",
    "xlsx",
    "scales",
    "classInt"
  )

install_and_load(packages)

rm(packages, install_and_load)

# Pour récupérer les fichiers sur insee.fr
if(Sys.getenv()["USERDNSDOMAIN"] == "AD.INSEE.INTRA"){
  set_config(use_proxy(url="proxy-rie.http.insee.fr", port=8080))
}

font_add_google(name = "Oswald", family = "oswald") 
showtext_auto()

REGmap <- st_read("cartes/reg_France_DOM.shp", quiet=TRUE, options = "ENCODING=UTF-8") %>% mutate_if(is.factor, ~ as.character(.x))
DEPmap <- st_read("cartes/dep_France_DOM_zoomPC.shp", quiet=TRUE, options = "ENCODING=UTF-8") %>% mutate_if(is.factor, ~ as.character(.x))

# fonctions qui récupèrent les nomsdes départements et régions
get_libdep <- function(codedep){
  depbase %>% filter(dep == codedep) %>% pull(libelle)
}

get_libreg <- function(codereg){
  regbase %>% filter(reg == codereg) %>% pull(libelle)
}

"%notin%" <- Negate("%in%")

liste_trimestres <- function(debut, n_trim){
  trim <- data %>% distinct(Periode) %>% rowid_to_column() 
  id_deb <- trim %>% filter(Periode == debut) %>% pull(rowid)
  trim %>% filter(rowid %in% c(id_deb, id_deb + n_trim)) %>% pull(Periode)
}

liste_dep <- function(region){
  depbase %>% filter(reg == region) %>% pull(dep)
}

#### Fonctions pour xlsx ####

#++++++++++++++++++++++++
# Fonction helper pour ajouter des titres
#++++++++++++++++++++++++
# - sheet : la feuille Excel pour contenir le titre
# - rowIndex : numéro de la ligne pour contenir le titre 
# - title : texte du titre
# - titleStyle : l'objet style pour le titre
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=1)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}

#### Theme cartes ####
theme_carte <- function (base_size = 11, base_family = "") {
  theme_bw() %+replace%
  theme(
    panel.ontop = F,   ## Note: this is to make the panel grid visible in this example
    panel.grid = element_blank(),
    line = element_blank(),
    rect = element_blank(),
    axis.text = element_blank(),
    axis.title = element_blank(),
    plot.background = element_rect(fill = "white"),
    plot.title = element_text(size = 60, family = "oswald", lineheight = 0.3),
    plot.subtitle = element_text(size = 40, family = "oswald"),
    plot.caption = element_text(size = 20),
    legend.position="bottom",
    legend.title = element_blank(),
    legend.text = element_text(colour="grey10", size=30, family = "sans", face="bold")
  )
}

theme_graph1 <- function(base_size = 11, base_family = ""){
  theme_gdocs() +
    theme(
      plot.title=element_text(family = "oswald", size = 24, lineheight = 1.5),
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "none",
      plot.margin = unit(c(1,1,1,2), "cm")
    ) 
}

theme_graph2 <- function(base_size = 11, base_family = ""){
  theme_gdocs() +
    theme(
      plot.title = element_text(family = "oswald", size = 50),
      plot.caption = element_text(size = 20),
      strip.text.x = element_text(size = 18),
      axis.text.x = element_text(angle = 45, size = 16, hjust = 1),
      axis.text.y = element_text(size = 20),
      axis.title.y = element_text(size = 30),
      legend.position = "bottom",
      legend.direction = "horizontal",
      legend.text = element_text(size = 20),
      plot.margin = unit(c(1,1,1,2), "cm")
    ) 
}

#### Répertoires ####
if(!dir.exists("sorties")){
  dir.create("sorties")
}

if(!dir.exists("sorties/Bases")){
  dir.create("sorties/Bases")
}

if(!dir.exists("sorties/Régions")){
  dir.create("sorties/Régions")
}
