#### Chargement des packages et des fonctions ####
source("fonctions.r", encoding = "UTF-8")

#### chargement des bases et tables de passage ####
depbase <- read_csv(file = "tables/departement2019.csv")
regbase <- read_csv(file = "tables/region2019.csv") %>% 
  add_row(reg = c("FM","FR"), libelle = c("France métropolitaine", "France hors Mayotte")) 
a17base <- read_csv2(file = "tables/table_a17.csv") %>% replace_na(list(LIBA17 = ""))
                     
fichiers <- dir("data") %>% 
  str_sub(13, str_length(.)-4) %>% as.tibble %>% filter(value != "") %>% 
  mutate(
    annee = as.numeric(str_sub(value, 1,4)),
    trimestre = str_to_upper(str_sub(value, 5,6)),
    nom = paste0("ete_dep_a17_",value,".xls")
  ) %>% select(-value) %>% 
  arrange(desc(annee), desc(trimestre))

data <- read_xls(paste0("data/",first(fichiers$nom))) %>% # on charge le dernier fichier
  select(-REG, REG=REG2016) %>% 
  gather(key = "Periode", value = "nb_emp", -REG, -DEP, -A5, -A17) %>% 
  mutate(
    Annee = as.numeric(str_sub(Periode, 4,7)),
    Trimestre = str_to_upper(str_sub(Periode, 8,9)),
    Periode = paste(Annee, Trimestre, sep="-")
  ) %>% 
  arrange(Annee, Trimestre)
rm(fichiers)
#### fin chargement données et tables

#### variables globales ####
DEBUT_ETUDE <- "2019-T4"  # trimestre de début d'étude - format texte "aaaa-Tn"
FENETRE_ETUDE <- 1        # nombre de trimestre - format numérique (4 pour une année, 1 pour un trimestre)
#TODO tester si les variables sont valides

#### Détermination des périodes de calculs ####
Periodes <- liste_trimestres(DEBUT_ETUDE, FENETRE_ETUDE)
if(FENETRE_ETUDE == 1){
  texte_periodes <- glue("trimestriel au {Periodes[2]}")
  dirname <- glue("{Periodes[2]}")
} else {
  texte_periodes <- glue("entre le {Periodes[1]} et le {Periodes[2]}")
  dirname <- glue("{Periodes[1]} - {Periodes[2]}")
}

#### Création des répertoires periodes ####
if(!dir.exists(glue("sorties/Bases/{dirname}"))){
  dir.create(glue("sorties/Bases/{dirname}"))
}
if(!dir.exists(glue("sorties/Régions/{dirname}"))){
  dir.create(glue("sorties/Régions/{dirname}"))
}

#### filtrage des périodes pour l'étude ####
data_ASR <- data %>% filter(Periode %in% Periodes)

#### ASR REGxA5 ####
data_ASR_reg_A5 <- data_ASR %>% 
  group_by(Periode, REG, A5) %>% 
  summarise(nb_emp = sum(nb_emp,na.rm = T))

data_ASR_reg_A5 <- data_ASR_reg_A5 %>%
  bind_rows(data_ASR_reg_A5 %>%
              group_by(Periode, REG) %>%
              summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
              mutate(A5 = "Tous secteurs")
  ) #on rajoute le total Tous secteurs

data_ASR_reg_A5 <- data_ASR_reg_A5 %>%
  bind_rows(data_ASR_reg_A5 %>%
              group_by(Periode, A5) %>%
              summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
              mutate(REG = "FR")
  ) %>% ungroup #on rajoute le total France

data_ASR_reg_A5 <- data_ASR_reg_A5 %>%
  arrange(Periode, REG, A5) %>%
  group_by(REG, A5) %>%
  mutate(
    emploi_N0 = lag(nb_emp),
    emploi_N = nb_emp,
    ecart = nb_emp - lag(nb_emp),
    tx_evol = nb_emp / lag(nb_emp) - 1
  ) %>%
  filter(Periode == Periodes[2]) %>% 
  select(
    REG, A5, emploi_N0, emploi_N, ecart, tx_evol
  ) %>% ungroup

G <- data_ASR_reg_A5 %>% filter(REG == "FR", A5 == "Tous secteurs") %>% select(G=tx_evol)
Gi <- data_ASR_reg_A5 %>% filter(REG == "FR", A5 != "Tous secteurs") %>% select(A5, Gi=tx_evol)
gi <- data_ASR_reg_A5 %>% filter(REG != "FR", A5 != "Tous secteurs") %>% select(REG, A5, gi=tx_evol) 

ASR <- data_ASR_reg_A5 %>% mutate(G = G$G) %>% 
  left_join(Gi) %>% 
  left_join(gi)

ASR_REGxA5 <- ASR %>% mutate(
  Nat = emploi_N0 * G, 
  Sec = emploi_N0 * (Gi - G), 
  Reg = emploi_N0 * (gi - Gi),
  NATfx = G,
  SECfx = Gi - G,
  REGfx = gi - Gi
) %>% 
  select(REG, A5, emploi_N0, emploi_N, ecart, Nat, Sec, Reg, tx_evol, NATfx, SECfx, REGfx) %>% 
  filter(REG != "FR", A5 != "Tous secteurs")

ASR_REGxA5 <- bind_rows(
  ASR_REGxA5,
  ASR_REGxA5 %>% group_by(REG) %>% 
  summarise_at(
    vars(emploi_N0, emploi_N, ecart, Nat, Sec, Reg),
    sum,
    na.rm= T) %>% ungroup %>% 
  mutate(
    A5 = "Tous secteurs",
    tx_evol = ecart / emploi_N0,
    NATfx = Nat / emploi_N0 ,
    SECfx = Sec / emploi_N0 ,
    REGfx = Reg / emploi_N0 
  )
)

# Ecriture du fichiers de l'ASR REG en A5
write_csv2(ASR_REGxA5, glue("sorties/Bases/{dirname}/ASR_REGxA5.csv"))
# Suppression de la table de calcul
rm(ASR)

# tableau de l'onglet 1
tab_REGFRxA5 <- ASR_REGxA5 %>% select(REG, A5, emploi_N ,tx_evol, NATfx, SECfx, REGfx) %>% 
  gather(-REG, -A5, key = indic, value = val) %>% 
  unite("A5ind", c(A5, indic), sep = " - ") %>% 
  spread(A5ind, val) %>% inner_join(regbase, by = c("REG" = "reg")) %>% 
  select(REG, libelle, ends_with("- tx_evol"), ends_with("- NATfx")
         , ends_with("- SECfx"), ends_with("- REGfx")
  ) %>% 
  mutate_if(is.numeric, ~ round(.x * 100, 2)) #on arrondit à deux chiffres pour éviter les divergences et pour de meilleurs graphiques

#### Fin ASR_REG_A5

#### ASR REGxA17 ####
data_ASR_reg_A17 <- data_ASR %>% 
  group_by(Periode, REG, A17) %>% 
  summarise(nb_emp = sum(nb_emp,na.rm = T))

data_ASR_reg_A17 <- data_ASR_reg_A17 %>%
  bind_rows(data_ASR_reg_A17 %>%
              group_by(Periode, REG) %>%
              summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
              mutate(A17 = "Tous secteurs")
  ) #on rajoute le total Tous secteurs

data_ASR_reg_A17 <- data_ASR_reg_A17 %>%
  bind_rows(data_ASR_reg_A17 %>%
              group_by(Periode, A17) %>%
              summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
              mutate(REG = "FR")
  ) %>% ungroup #on rajoute le total France

data_ASR_reg_A17 <- data_ASR_reg_A17 %>%
  arrange(Periode, REG, A17) %>%
  group_by(REG, A17) %>%
  mutate(
    emploi_N0 = lag(nb_emp),
    emploi_N = nb_emp,
    ecart = nb_emp - lag(nb_emp),
    tx_evol = nb_emp / lag(nb_emp) - 1
  ) %>%
  filter(Periode == Periodes[2]) %>% 
  select(
    REG, A17, emploi_N0, emploi_N, ecart, tx_evol
  ) %>% ungroup

G <- data_ASR_reg_A17 %>% filter(REG == "FR", A17 == "Tous secteurs") %>% select(G=tx_evol)
Gi <- data_ASR_reg_A17 %>% filter(REG == "FR", A17 != "Tous secteurs") %>% select(A17, Gi=tx_evol)
gi <- data_ASR_reg_A17 %>% filter(REG != "FR", A17 != "Tous secteurs") %>% select(REG, A17, gi=tx_evol) 

ASR <- data_ASR_reg_A17 %>% mutate(G = G$G) %>% 
  left_join(Gi) %>% 
  left_join(gi)

ASR_REGxA17 <- ASR %>% mutate(
  Nat = emploi_N0 * G, 
  Sec = emploi_N0 * (Gi - G), 
  Reg = emploi_N0 * (gi - Gi),
  NATfx = G,
  SECfx = Gi - G,
  REGfx = gi - Gi
) %>% 
  select(REG, A17, emploi_N0, emploi_N, ecart, Nat, Sec, Reg, tx_evol, NATfx, SECfx, REGfx) %>% 
  filter(REG != "FR", A17 != "Tous secteurs")


ASR_REGxA17 <- bind_rows(
  ASR_REGxA17,
  ASR_REGxA17 %>% group_by(REG) %>% 
    summarise_at(
      vars(emploi_N0, emploi_N, ecart, Nat, Sec, Reg),
      sum,
      na.rm= T) %>% ungroup %>% 
    mutate(
      A17 = "Tous secteurs",
      tx_evol = ecart / emploi_N0,
      NATfx = Nat / emploi_N0 ,
      SECfx = Sec / emploi_N0 ,
      REGfx = Reg / emploi_N0 
    )
)

# Ecriture du fichiers de l'ASR REG en A17
write_csv2(ASR_REGxA17, glue("sorties/Bases/{dirname}/ASR_REGxA17.csv"))
# Suppression de la table de calcul
rm(ASR)

# tableau de l'onglet 3
tab_REGFRxA17 <- ASR_REGxA17 %>% select(REG, A17, emploi_N ,tx_evol, NATfx, SECfx, REGfx) %>% 
gather(-REG, -A17, key = indic, value = val) %>% 
  unite("A17ind", c(A17, indic), sep = " - ") %>% 
  spread(A17ind, val) %>% inner_join(regbase, by = c("REG" = "reg")) %>% 
  select(REG, libelle, ends_with("- tx_evol"), ends_with("- NATfx")
         , ends_with("- SECfx"), ends_with("- REGfx")
  ) %>% 
  mutate_if(is.numeric, ~ round(.x * 100, 2)) #on arrondit à deux chiffres pour éviter les divergences et pour de meilleurs graphiques

#### Fin ASR_REG_A17 

#### Données pour cartes régions ####
data_carto_REG <- inner_join(
  REGmap,
  ASR_REGxA17 %>% group_by(REG) %>%
    summarise_at(vars(emploi_N0, ecart, Reg), sum, na.rm = F) %>%
    mutate(REGfx = round(100 * Reg / emploi_N0, 2),
           tx_evol = round(100 * ecart / emploi_N0, 2)
    ),
  by = c("code" = "REG")
)
tx_fr <- ASR_REGxA17 %>% 
  summarise_at(vars(emploi_N0, ecart, Reg), sum, na.rm = F) %>% 
  mutate(tx_evol = round(100 * ecart / emploi_N0, 2)) %>% pull(tx_evol)
#### Cartes évolutions région ####
# get quantile breaks. Add .00001 offset to catch the lowest value
breaks_qt <- classIntervals(c(min(data_carto_REG$tx_evol) - .00001, data_carto_REG$tx_evol), n = 5, style = "jenks", dataPrecision = 1, digits = 1)
breaks_qt$brks <- round(breaks_qt$brks,1)
breaks_qt$brks[1] <- breaks_qt$brks[1] -0.1
breaks_qt$brks[6] <- breaks_qt$brks[6] +0.1
data_carto_REG <- data_carto_REG %>% mutate(tx_evol_cat = cut(tx_evol, breaks_qt$brks))
carte_REG <- ggplot(data_carto_REG) + 
  geom_sf(aes(geometry = geometry, fill = tx_evol_cat), colour = "black")  +
  geom_sf_text(aes(label = glue("{round(tx_evol,1)} %")), colour = "black", size = 10, fontface = 2, family = "sans")+
  scale_fill_brewer( palette = "RdBu" )  +
  labs(
    title = glue("Évolution de l'emploi régional {texte_periodes}"),
    caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
    subtitle = glue("Évolution France entière (hors Mayotte): {tx_fr} %")
  ) + 
  theme_carte()
  
#### Cartes effets résiduels régions ####
# get quantile breaks. Add .00001 offset to catch the lowest value
breaks_qt <- classIntervals(c(min(data_carto_REG$REGfx) - .00001, data_carto_REG$REGfx), n = 5, style = "jenks", dataPrecision = 1, digits = 1)
breaks_qt$brks <- round(breaks_qt$brks,1)
breaks_qt$brks[1] <- breaks_qt$brks[1] -0.1
breaks_qt$brks[6] <- breaks_qt$brks[6] +0.1
data_carto_REG <- data_carto_REG %>% mutate(REGfx_cat = cut(REGfx, breaks_qt$brks))
carte_REGfx <- ggplot(data_carto_REG) + 
  geom_sf(aes(geometry = geometry, fill = REGfx_cat), colour = "black")  +
  geom_sf_text(aes(label = glue("{round(REGfx,1)} %")), colour = "black", size = 10, fontface = 2, family = "sans")+
  scale_fill_brewer( palette = "OrRd" )  +
  labs(
    title = glue("Effet résiduel de l'emploi régional {texte_periodes}"),
    caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}")
  ) + 
  theme_carte()

#### Début fonction analyse par région ####
ASR_fonction <- function(region){
  
  # tableau de l'onglet 2
  tab_REGaaxA5 <- ASR_REGxA5 %>% filter(REG == region) %>% 
    select(A5, everything(), -REG) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    mutate_at(vars(emploi_N0, emploi_N, ecart, Nat, Sec, Reg ), ~ round(.x,0))
  
  # tableau de l'onglet 4
  tab_REGaaxA17 <- ASR_REGxA17 %>% filter(REG == region) %>% 
    inner_join(a17base) %>% mutate(Secteur = paste(A17,LIBA17,sep = "-")) %>% 
    select(Secteur, A5, everything(), -REG, -A17, -LIBA17) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    mutate_at(vars(emploi_N0, emploi_N, ecart, Nat, Sec, Reg ), ~ round(.x,0))
    
  #### ASR DEPxA5 ####
  data_ASR_dep_A5 <- data_ASR %>% 
    filter(REG == region) %>% 
    group_by(Periode, DEP, A5) %>% 
    summarise(nb_emp = sum(nb_emp,na.rm = T))
  
  data_ASR_dep_A5 <- data_ASR_dep_A5 %>%
    bind_rows(data_ASR_dep_A5 %>%
                group_by(Periode, DEP) %>%
                summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
                mutate(A5 = "Tous secteurs")
    ) #on rajoute le total Tous secteurs
  
  data_ASR_dep_A5 <- data_ASR_dep_A5 %>%
    bind_rows(data_ASR_dep_A5 %>%
                group_by(Periode, A5) %>%
                summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
                mutate(DEP = get_libreg(region))
    ) %>% ungroup #on rajoute le total France
  
  data_ASR_dep_A5 <- data_ASR_dep_A5 %>%
    arrange(Periode, DEP, A5) %>%
    group_by(DEP, A5) %>%
    mutate(
      emploi_N0 = lag(nb_emp),
      emploi_N = nb_emp,
      ecart = nb_emp - lag(nb_emp),
      tx_evol = nb_emp / lag(nb_emp) - 1
    ) %>%
    filter(Periode == Periodes[2]) %>% 
    select(
      DEP, A5, emploi_N0, emploi_N, ecart, tx_evol
    ) %>% ungroup
  
  G <- data_ASR_dep_A5 %>% filter(DEP == get_libreg(region), A5 == "Tous secteurs") %>% select(G=tx_evol)
  Gi <- data_ASR_dep_A5 %>% filter(DEP == get_libreg(region), A5 != "Tous secteurs") %>% select(A5, Gi=tx_evol)
  gi <- data_ASR_dep_A5 %>% filter(DEP != get_libreg(region), A5 != "Tous secteurs") %>% select(DEP, A5, gi=tx_evol) 
  
  ASR <- data_ASR_dep_A5 %>% mutate(G = G$G) %>% 
    left_join(Gi) %>% 
    left_join(gi)
  
  ASR_DEPxA5 <- ASR %>% mutate(
    Reg = emploi_N0 * G, 
    Sec = emploi_N0 * (Gi - G), 
    Dep = emploi_N0 * (gi - Gi),
    REGfx = G,
    SECfx = Gi - G,
    DEPfx = gi - Gi
  ) %>% 
    select(DEP, A5, emploi_N0, emploi_N, ecart, Reg, Sec, Dep, tx_evol, REGfx, SECfx, DEPfx) %>% 
    filter(DEP != get_libreg(region), A5 != "Tous secteurs")
  
  
  ASR_DEPxA5 <- bind_rows(
    ASR_DEPxA5,
    ASR_DEPxA5 %>% group_by(DEP) %>% 
      summarise_at(
        vars(emploi_N0, emploi_N, ecart, Reg, Sec, Dep),
        sum,
        na.rm= T) %>% ungroup %>% 
      mutate(
        A5 = "Tous secteurs",
        tx_evol = ecart / emploi_N0,
        REGfx = Reg / emploi_N0 ,
        SECfx = Sec / emploi_N0 ,
        DEPfx = Dep / emploi_N0
      )
  )
  
  # Ecriture du fichiers de l'ASR DEP en A5
  write_csv2(ASR_DEPxA5, glue("sorties/Bases/{dirname}/ASR_DEP{region}xA5.csv"))
  # Suppression de la table de calcul
  rm(ASR)
  
  # tableau de l'onglet 5
  tab_DEPREGxA5 <- ASR_DEPxA5 %>% select(DEP, A5, emploi_N ,tx_evol, REGfx, SECfx, DEPfx) %>% 
    gather(-DEP, -A5, key = indic, value = val) %>% 
    unite("A5ind", c(A5, indic), sep = " - ") %>% 
    spread(A5ind, val) %>% inner_join(depbase, by = c("DEP" = "dep")) %>% 
    select(DEP, libelle, ends_with("- tx_evol"), ends_with("- REGfx")
           , ends_with("- SECfx"), ends_with("- DEPfx")
    ) %>% 
    mutate_if(is.numeric, ~ round(.x * 100, 2)) #on arrondit à deux chiffres pour éviter les divergences et pour de meilleurs graphiques
  
  # tableau de l'onglet 6
  tab_DEPaaxA5 <- ASR_DEPxA5 %>% arrange(DEP,A5) %>% 
    inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>% 
    mutate_at(vars(emploi_N0, emploi_N, ecart, Reg, Sec, Dep ), ~ round(.x,0)) %>% 
    select(DEP, libelle, everything())
  #### Fin ASR_DEP_A5
  
  #### ASR DEPxA17 ####
  data_ASR_dep_A17 <- data_ASR %>% 
    filter(REG == region) %>% 
    group_by(Periode, DEP, A17) %>% 
    summarise(nb_emp = sum(nb_emp,na.rm = T))
  
  data_ASR_dep_A17 <- data_ASR_dep_A17 %>%
    bind_rows(data_ASR_dep_A17 %>%
                group_by(Periode, DEP) %>%
                summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
                mutate(A17 = "Tous secteurs")
    ) #on rajoute le total Tous secteurs
  
  data_ASR_dep_A17 <- data_ASR_dep_A17 %>%
    bind_rows(data_ASR_dep_A17 %>%
                group_by(Periode, A17) %>%
                summarise(nb_emp = sum(nb_emp, na.rm = T)) %>%
                mutate(DEP = get_libreg(region))
    ) %>% ungroup #on rajoute le total France
  
  data_ASR_dep_A17 <- data_ASR_dep_A17 %>%
    arrange(Periode, DEP, A17) %>%
    group_by(DEP, A17) %>%
    mutate(
      emploi_N0 = lag(nb_emp),
      emploi_N = nb_emp,
      ecart = nb_emp - lag(nb_emp),
      tx_evol = nb_emp / lag(nb_emp) - 1
    ) %>%
    filter(Periode == Periodes[2]) %>% 
    select(
      DEP, A17, emploi_N0, emploi_N, ecart, tx_evol
    ) %>% ungroup
  
  G <- data_ASR_dep_A17 %>% filter(DEP == get_libreg(region), A17 == "Tous secteurs") %>% select(G=tx_evol)
  Gi <- data_ASR_dep_A17 %>% filter(DEP == get_libreg(region), A17 != "Tous secteurs") %>% select(A17, Gi=tx_evol)
  gi <- data_ASR_dep_A17 %>% filter(DEP != get_libreg(region), A17 != "Tous secteurs") %>% select(DEP, A17, gi=tx_evol) 
  
  ASR <- data_ASR_dep_A17 %>% mutate(G = G$G) %>% 
    left_join(Gi) %>% 
    left_join(gi)
  
  ASR_DEPxA17 <- ASR %>% mutate(
    Reg = emploi_N0 * G, 
    Sec = emploi_N0 * (Gi - G), 
    Dep = emploi_N0 * (gi - Gi),
    REGfx = G,
    SECfx = Gi - G,
    DEPfx = gi - Gi
  ) %>% 
    select(DEP, A17, emploi_N0, emploi_N, ecart, Reg, Sec, Dep, tx_evol, REGfx, SECfx, DEPfx) %>% 
    filter(DEP != get_libreg(region), A17 != "Tous secteurs")
  
  ASR_DEPxA17 <- bind_rows(
    ASR_DEPxA17,
    ASR_DEPxA17 %>% group_by(DEP) %>% 
      summarise_at(
        vars(emploi_N0, emploi_N, ecart, Reg, Sec, Dep),
        sum,
        na.rm= T) %>% ungroup %>% 
      mutate(
        A17 = "Tous secteurs",
        tx_evol = ecart / emploi_N0,
        REGfx = Reg / emploi_N0 ,
        SECfx = Sec / emploi_N0 ,
        DEPfx = Dep / emploi_N0
      )
  )
  
  # Ecriture du fichiers de l'ASR DEP en A17
  write_csv2(ASR_DEPxA17, glue("sorties/Bases/{dirname}/ASR_DEP{region}xA17.csv"))
  # Suppression de la table de calcul
  rm(ASR)
  
  # tableau de l'onglet 7
  tab_DEPREGxA17 <- ASR_DEPxA17 %>% select(DEP, A17, emploi_N ,tx_evol, REGfx, SECfx, DEPfx) %>% 
    gather(-DEP, -A17, key = indic, value = val) %>% 
    unite("A17ind", c(A17, indic), sep = " - ") %>% 
    spread(A17ind, val) %>% inner_join(depbase, by = c("DEP" = "dep")) %>% 
    select(DEP, libelle, ends_with("- tx_evol"), ends_with("- REGfx")
           , ends_with("- SECfx"), ends_with("- DEPfx")
    ) %>% 
    mutate_if(is.numeric, ~ round(.x * 100, 2)) #on arrondit à deux chiffres pour éviter les divergences et pour de meilleurs graphiques
  
  # tableau de l'onglet 8
  tab_DEPaaxA17 <- ASR_DEPxA17 %>% arrange(DEP,A17) %>% inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% 
    inner_join(a17base) %>% mutate(Secteur = paste(A17,LIBA17,sep = "-")) %>% 
    select(DEP, libelle, Secteur, A5, everything(), -A17, -LIBA17) %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>% 
    mutate_at(vars(emploi_N0, emploi_N, ecart, Reg, Sec, Dep ), ~ round(.x,0))
  #### Fin ASR_DEP_A7
  
  #### Création du classeur Excel ####
  classeur <- createWorkbook(type="xlsx")
  
  #### Styles pour xlsx ####
  
  # Définir quelques styles de cellules
  #++++++++++++++++++++
  # Titres et sous-titres
  TITLE_STYLE <- CellStyle(classeur) + Font(classeur,  heightInPoints=14, color="grey1", isBold=TRUE, underline=1)
  SUB_TITLE_STYLE <- CellStyle(classeur) + Font(classeur,  heightInPoints=10, isItalic=TRUE, isBold=FALSE, color = "grey25")
  # Styles pour les noms de lignes/colonnes
  TABLE_ROWNAMES_STYLE <- CellStyle(classeur) + Font(classeur, isBold=TRUE)
  TABLE_COLNAMES_STYLE <- CellStyle(classeur) + Font(classeur, isBold=TRUE) + Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
    Border(color="black", position=c("TOP", "BOTTOM"), pen=c("BORDER_THIN", "BORDER_THICK")) 
  
  #### Les feuilles du classeur ####
  # Feuille 1
  sheet1 <- createSheet(classeur, sheetName = "ASR_REGFR_A5")
  xlsx.addTitle(sheet1, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi par région en secteurs A5 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet1, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_REGFRxA5), sheet1, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet1, colIndex=1, colWidth=6)
  setColumnWidth(sheet1, colIndex=2, colWidth=30)
  setColumnWidth(sheet1, colIndex=c(3:ncol(tab_REGFRxA5)), colWidth=14)
  # Feuille 2
  sheet2 <- createSheet(classeur, sheetName = glue("ASR_REG{region}_A5"))
  xlsx.addTitle(sheet2, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi de la région {get_libreg(region)} en secteurs A5 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet2, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_REGaaxA5), sheet2, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet2, colIndex=1, colWidth=30)
  setColumnWidth(sheet2, colIndex=c(2:ncol(tab_REGaaxA5)), colWidth=14)
  # Feuille 3
  sheet3 <- createSheet(classeur, sheetName = "ASR_REGFR_A17")
  xlsx.addTitle(sheet3, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi par région en secteurs A17 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet3, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_REGFRxA17), sheet3, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet3, colIndex=1, colWidth=6)
  setColumnWidth(sheet3, colIndex=2, colWidth=30)
  setColumnWidth(sheet3, colIndex=c(3:ncol(tab_REGFRxA17)), colWidth=14)
  # Feuille 4
  sheet4 <- createSheet(classeur, sheetName = glue("ASR_REG{region}_A17"))
  xlsx.addTitle(sheet4, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi de la région {get_libreg(region)} en secteurs A17 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet4, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_REGaaxA17), sheet4, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet4, colIndex=1, colWidth=60)
  setColumnWidth(sheet4, colIndex=2, colWidth=30)
  setColumnWidth(sheet4, colIndex=c(3:ncol(tab_REGaaxA17)), colWidth=14)
  # Feuille 5
  sheet5 <- createSheet(classeur, sheetName = glue("ASR_DEP{region}_A5"))
  xlsx.addTitle(sheet5, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi des département de la région {get_libreg(region)} en secteurs A5 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet5, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_DEPREGxA5), sheet5, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet5, colIndex=1, colWidth=6)
  setColumnWidth(sheet5, colIndex=2, colWidth=30)
  setColumnWidth(sheet5, colIndex=c(3:ncol(tab_DEPREGxA5)), colWidth=14)
  # Feuille 6
  sheet6 <- createSheet(classeur, sheetName = glue("ASR_DEP{region}_A5_2"))
  xlsx.addTitle(sheet6, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi des département de la région {get_libreg(region)} en secteurs A5 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet6, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_DEPaaxA5), sheet6, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet6, colIndex=1, colWidth=6)
  setColumnWidth(sheet6, colIndex=2, colWidth=20)
  setColumnWidth(sheet6, colIndex=3, colWidth=30)
  setColumnWidth(sheet6, colIndex=c(4:ncol(tab_DEPaaxA5)), colWidth=14)
  # Feuille 7
  sheet7 <- createSheet(classeur, sheetName = glue("ASR_DEP{region}_A17"))
  xlsx.addTitle(sheet7, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi des département de la région {get_libreg(region)} en secteurs A17 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet7, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_DEPREGxA17), sheet7, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet7, colIndex=1, colWidth=6)
  setColumnWidth(sheet7, colIndex=2, colWidth=30)
  setColumnWidth(sheet7, colIndex=c(3:ncol(tab_DEPREGxA17)), colWidth=14)
  # Feuille 8
  sheet8 <- createSheet(classeur, sheetName = glue("ASR_DEP{region}_A17_2"))
  xlsx.addTitle(sheet8, rowIndex=1, 
                title = glue("Analyse structurelle résiduelle des évolutions de l'emploi des département de la région {get_libreg(region)} en secteurs A17 entre {Periodes[1]} et {Periodes[2]}"), 
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet8, rowIndex=2, 
                title = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
                titleStyle = SUB_TITLE_STYLE)
  addDataFrame(as.data.frame(tab_DEPaaxA17), sheet8, startRow=4, startColumn=1, row.names = F)
  setColumnWidth(sheet8, colIndex=1, colWidth=6)
  setColumnWidth(sheet8, colIndex=2, colWidth=20)
  setColumnWidth(sheet8, colIndex=3, colWidth=30)
  setColumnWidth(sheet8, colIndex=4, colWidth=20)
  setColumnWidth(sheet8, colIndex=c(5:ncol(tab_DEPaaxA17)), colWidth=14)
  ####
  
  #### Sauvegarde du classeur ####
  if(!dir.exists(glue("sorties/Régions/{dirname}/{get_libreg(region)}"))){
    dir.create(glue("sorties/Régions/{dirname}/{get_libreg(region)}"))
  }
  saveWorkbook(classeur, glue("Sorties/Régions/{dirname}/{get_libreg(region)}/ASR_emploi_{region}.xlsx"))
 
  #### Cartes des Effets résiduels et de l'évolution d'emploi dans les régions ####
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/carte_REG.jpg"), plot = carte_REG, dpi = 300, width = 8, height = 8)
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/carte_REGfx.jpg"), plot = carte_REGfx, dpi = 300, width = 8, height = 8)
  
  #### Données pour cartes départementales ####
  data_carto_DEP <- inner_join(
    DEPmap %>% mutate(libelle = map_chr(code, ~ get_libdep(.x))),
    ASR_DEPxA17 %>% group_by(DEP) %>%
      summarise_at(vars(emploi_N0, ecart, Dep), sum, na.rm = F) %>%
      mutate(DEPfx = round(100 * Dep / emploi_N0, 2),
             tx_evol = round(100 * ecart / emploi_N0, 2)
      ),
    by = c("code" = "DEP")
  )
  tx_reg <- ASR_DEPxA17 %>% 
    summarise_at(vars(emploi_N0, ecart, Dep), sum, na.rm = F) %>% 
    mutate(tx_evol = round(100 * ecart / emploi_N0, 1)) %>% 
    pull(tx_evol)
  #### Cartes évolutions départements ####
  # get quantile breaks. Add .00001 offset to catch the lowest value
  breaks_qt <- classIntervals(c(min(data_carto_DEP$tx_evol) - .00001, data_carto_DEP$tx_evol), n = 5, style = "jenks", dataPrecision = 1, digits = 1)
  breaks_qt$brks <- unique(round(breaks_qt$brks,1))
  breaks_qt$brks[1] <- breaks_qt$brks[1] -0.1
  breaks_qt$brks[6] <- breaks_qt$brks[6] +0.1
  data_carto_DEP <- data_carto_DEP %>% mutate(tx_evol_cat = cut(tx_evol, breaks_qt$brks))
  carte_DEP <- ggplot(data_carto_DEP) + 
    geom_sf(aes(geometry = geometry, fill = tx_evol_cat), colour = "black")  +
    geom_sf_text(aes(label = glue("{libelle}\n{round(tx_evol,1)} %")), colour = "black", size = 10, fontface = 2, family = "sans", lineheight = 0.4)+
    scale_fill_brewer( palette = "RdBu" )  +
    labs(
      title = glue("Évolution de l'emploi départemental\n{texte_periodes}"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      subtitle = glue("Évolution région {get_libreg(region)}: {tx_reg} %")
    ) + 
    theme_carte()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/carte_DEP{region}.jpg"), plot = carte_DEP, dpi = 300, width = 8, height = 8)
  
  #### Cartes effets résiduels départements ####
  # get quantile breaks. Add .00001 offset to catch the lowest value
  breaks_qt <- classIntervals(c(min(data_carto_DEP$DEPfx) - .00001, data_carto_DEP$DEPfx), n = 5, style = "jenks", dataPrecision = 1, digits = 1)
  breaks_qt$brks <- unique(round(breaks_qt$brks,1))
  breaks_qt$brks[1] <- breaks_qt$brks[1] -0.1
  breaks_qt$brks[6] <- breaks_qt$brks[6] +0.1
  data_carto_DEP <- data_carto_DEP %>% mutate(DEPfx_cat = cut(DEPfx, breaks_qt$brks))
  carte_DEPfx <- ggplot(data_carto_DEP) + 
    geom_sf(aes(geometry = geometry, fill = DEPfx_cat), colour = "black")  +
    geom_sf_text(aes(label = glue("{libelle}\n{round(DEPfx,1)} %")), colour = "black", size = 10, fontface = 2, family = "sans", lineheight = 0.4)+
    scale_fill_brewer( palette = "OrRd" )  +
    labs(
      title = glue("Effet résiduel de l'emploi départemental\n{texte_periodes}"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}")
    ) + 
    theme_carte()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/carte_DEPfx{region}.jpg"), plot = carte_DEPfx, dpi = 300, width = 8, height = 8)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les régions ####
  ASR_REGxA5 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    filter(A5 == "Tous secteurs", REG %notin% c("01","02","03","04","06")) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= libelle, y=REGfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi dans les régions"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    ) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph1()
  
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG.jpg"), dpi = 60, width = 12, height = 15)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les régions en A5 ####
  ASR_REGxA5 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    filter(REG %notin% c("01","02","03","04","06")) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= libelle, y=REGfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
    title = glue("Effet résiduel de l'évolution d'emploi dans les régions par grand secteur"),
    y = glue("Effet résiduel"),
    caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
    fill = "",
    x = ""
    ) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    facet_wrap(~ as_factor(A5), ncol = 7) +
    theme_graph2()
    
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG_A5.jpg"), dpi = 150, width = 20, height = 12)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans la région en A5 ####
  ASR_REGxA5 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    filter(REG == region) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= A5, y=REGfx, fill = A5)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi dans la région\n{get_libreg(region)} par grand secteur"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    ) + 
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph1()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG{region}_A5.jpg"), dpi = 60, width = 12, height = 12)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les régions en A17 ####
  ASR_REGxA17 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    inner_join(a17base) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    mutate(A17 = paste(A17, LIBA17,sep =" - ")) %>% 
    filter(REG %notin% c("01","02","03","04","06")) %>% 
    ggplot(aes(x= libelle, y=REGfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi dans les régions en A17"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x =""
    ) +
    facet_wrap(~ as_factor(A17), ncol = 6) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph2()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG_A17.jpg"), dpi = 150, width = 20, height = 12)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans la région en A17 ####
  ASR_REGxA17 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    inner_join(a17base) %>% 
    #mutate(A17 = paste(A17, LIBA17,sep =" - ")) %>% 
    filter(REG == region) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= A17, y=REGfx, fill = A17)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi\ndans la région {get_libreg(region)} en A17"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    )  +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph1()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG{region}_A17.jpg"), dpi = 60, width = 14, height = 18)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les départements de la région ####
  ASR_DEPxA5 %>% inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% 
    filter(A5 == "Tous secteurs") %>% mutate(libelle = fct_reorder(factor(libelle), DEP, min)) %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x = libelle, y=DEPfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi\ndans les départements de la région {get_libreg(region)}"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    ) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph1()+
    theme(plot.title = element_text(lineheight = 0.8))
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_DEP{region}.jpg"), dpi = 100, width = 8, height = 8)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les départements de la région en A5 ####
  ASR_DEPxA5 %>% inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% filter(A5 !="Tous secteurs") %>% 
    mutate(libelle = fct_reorder(factor(libelle), DEP, min)) %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= libelle, y=DEPfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi dans les départements\nde la région {get_libreg(region)} par grand secteur"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    ) +
    facet_wrap(~ as_factor(A5), ncol = 3) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph2()
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_DEP{region}_A5.jpg"), dpi = 100, width = 16, height = 16)
  
  #### Graphique des Effets résiduels de l'évolution d'emploi dans les départements de la région en A17 ####
  ASR_DEPxA17 %>% inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% 
    mutate(libelle = fct_reorder(factor(libelle), DEP, min)) %>% 
    inner_join(a17base) %>% 
    mutate(A17 = paste(A17, LIBA17,sep =" - ")) %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>% 
    ggplot(aes(x= libelle, y=DEPfx, fill = libelle)) +
    geom_bar(stat = 'identity' ,color = "white") +
    labs(
      title = glue("Effet résiduel de l'évolution d'emploi dans les départements\nde la région {get_libreg(region)} en A17"),
      y = glue("Effet résiduel"),
      caption = glue("Source: Estimations localisées d'emploi, MSA, ACOSS, INSEE, DARES}"),
      fill = "",
      x = ""
    )  +
    facet_wrap(~ as_factor(A17), ncol = 6) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 8)) +
    theme_graph2() 
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_DEP{region}_A17.jpg"), dpi = 120, width = 14, height = 18)
  
  #### Effet structurel-résiduel tous secteurs des régions de France ####
  ASR_REGxA5 %>% inner_join(regbase %>% select(reg, libelle), by = c("REG" = "reg")) %>% 
    filter(A5 == "Tous secteurs", REG %notin% c("01","02","03","04","06")) %>% 
    mutate_at(vars(tx_evol,NATfx, SECfx, REGfx), ~ round(100 * .x,2)) %>%
    ggplot() +
    geom_point(aes(x = REGfx,
                   y = SECfx,
                   size = emploi_N, 
                   color = ecart)
    ) +
    geom_hline(yintercept = 0) +
    geom_vline(xintercept = 0) +
    geom_abline(slope = -1,
                intercept = 0,
                size = 1.5) +
    geom_text_repel(aes(
      x = REGfx,
      y = SECfx,
      label = libelle
    ),
    box.padding = 1.5
    ) +
    scale_size(range = c(1, 20),
               labels = label_comma(big.mark = " ", decimal.mark = ",")) +
    scale_color_gradient_tableau(palette= "Classic Green",
                                 labels = label_comma(big.mark = " ", decimal.mark = ",")
    )  +
    labs(
      title = "Évolution de l'emploi selon la structure sectorielle",
      subtitle = glue("Régions de France métropolitaine entre {Periodes[1]} et {Periodes[2]}"),
      caption = "Source: Estimations d'emplois localisés; Insee, ACOSS, MSA, DARES",
      x = "Effet résiduel",
      y = "Effet structurel de la région",
      color = "Évolution de l'emploi",
      size = glue("Nombre d'emploi au {Periodes[2]}")
    ) +
    scale_x_continuous(breaks = scales::pretty_breaks(n = 10)) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 10)) +
    theme_gdocs() +
    theme(
      legend.position = "right",
      panel.background = element_rect(fill = "grey98",
                                      colour = "lightblue",
                                      size = 0.5, linetype = "solid"),
      panel.grid.major = element_line(size = 0.5, linetype = 'solid',
                                      colour = "grey60"), 
      panel.grid.minor = element_line(size = 0.25, linetype = 'dashed',
                                      colour = "grey75"),
      plot.title = element_text(color="black", size=20, face="bold.italic"),
      plot.subtitle = element_text(color="grey30", size=14, face="plain"),
      plot.caption = element_text(color="grey20", size=10, face="italic")
    )
  
  ggsave(glue("Sorties/Régions/{dirname}/{get_libreg(region)}/graph_REG_struct_res.jpg"), dpi = 100, width = 12, height = 12)
  
    #### Effet structurel-résiduel tous secteurs des régions de France ####
  ASR_DEPxA5 %>% inner_join(depbase %>% select(dep, libelle), by = c("DEP" = "dep")) %>% 
    filter(A5 == "Tous secteurs") %>% 
    mutate_at(vars(tx_evol,REGfx, SECfx, DEPfx), ~ round(100 * .x,2)) %>%
    ggplot() +
    geom_point(aes(x = DEPfx,
                   y = SECfx,
                   size = emploi_N, 
                   color = ecart)
    ) +
    geom_hline(yintercept = 0) +
    geom_vline(xintercept = 0) +
    geom_abline(slope = -1,
                intercept = 0,
                size = 1.5) +
    geom_text_repel(aes(
      x = DEPfx,
      y = SECfx,
      label = libelle
    ),
    box.padding = 1.5
    ) +
    scale_size(range = c(1, 20),
               labels = label_comma(big.mark = " ", decimal.mark = ",")) +
    scale_color_gradient_tableau(palette= "Classic Blue",
                                 labels = label_comma(big.mark = " ", decimal.mark = ",")
    )  +
    labs(
      title = "Évolution de l'emploi selon la structure sectorielle",
      subtitle = glue("Départements de la région {get_libreg(region)} entre le {Periodes[1]} et le {Periodes[2]}"),
      caption = "Source: Estimations d'emplois localisés; Insee, ACOSS, MSA, DARES",
      x = "Effet résiduel",
      y = "Effet structurel du département",
      color = "Évolution de l'emploi",
      size = glue("Nombre d'emploi au {Periodes[2]}")
    ) +
    scale_x_continuous(breaks = scales::pretty_breaks(n = 10)) +
    scale_y_continuous(breaks = scales::pretty_breaks(n = 10)) +
    theme_gdocs() +
    theme(
      legend.position = "right",
      panel.background = element_rect(fill = "grey98",
                                      colour = "lightblue",
                                      size = 0.5, linetype = "solid"),
      panel.grid.major = element_line(size = 0.5, linetype = 'solid',
                                      colour = "grey60"), 
      panel.grid.minor = element_line(size = 0.25, linetype = 'dashed',
                                      colour = "grey75"),
      plot.title = element_text(color="black", size=20, face="bold.italic"),
      plot.subtitle = element_text(color="grey30", size=14, face="plain"),
      plot.caption = element_text(color="grey20", size=10, face="italic")
    )
  
  ggsave(glue("sorties/Régions/{dirname}/{get_libreg(region)}/graph_DEP{region}_struct_res.jpg"), dpi = 100, width = 12, height = 12)
  
  print(glue("Les fichiers de la région {get_libreg(region)} pour la période {texte_periodes} sont créés"))
}

regbase %>% filter(reg %notin% c("01","02","03","04","06", "FM", "FR")) %>% pull(reg) %>% walk(~ ASR_fonction(.x))

