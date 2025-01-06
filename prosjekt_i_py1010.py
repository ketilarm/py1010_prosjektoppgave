"""Del a) Skriv et program som leser inn filen 'support_uke_24.xlsx' og lagrer data fra kolonne 1 
i en array med variablenavn 'u_dag', dataen i kolonne 2 lagres i arrayen 'kl_slett', data i 
kolonne 3 lagres i arrayen 'varighet' og dataen i kolonne 4 lagres i arrayen 'score'.  Merk: 
filen 'support_uke_24.xlsx' må ligge i samme mappe som Python-programmet ditt. """

# Importerer moduler som trengs til alle oppgavene
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os


# file_path = "USN/Prosjektoppgave/support_uke_24.xlsx"
file_path = "support_uke_24.xlsx"

# leser inn excel filen
df = pd.read_excel(file_path)

# Lagrer dataen i kolonnene i arrayer
kl_slett = df["Klokkeslett"]
u_dag = df["Ukedag"]
varighet = df["Varighet"]
score = df["Tilfredshet"]


# Brukes til debugging
# print(kl_slett)
# print(u_dag)
# print(varighet)
# print(score)


"""Del b) Skriv et program som finner antall henvendelser for hver de 5 ukedagene. Resultatet 
visualiseres ved bruk av et søylediagram (stolpediagram)"""

# Lager en dictionary for å lagre antall henvendelser for hver dag
henvendelser = {"Mandag":0, "Tirsdag":0, "Onsdag":0, "Torsdag":0, "Fredag":0}

# ittererer gjennom hver av dagene og legger til 1 i dictionaryen for hver gang en dag dukker opp
for dag in u_dag:
    if dag in henvendelser:
        henvendelser[dag] += 1


# print(henvendelser)


# Data for x og y aksen
x= henvendelser.keys()
y = henvendelser.values()

# formatering av stolpediagramet
plt.xlabel("Dager")
plt.ylabel("Antall henvendelser")
plt.title("Antall henvendelser per dag")

# oppretter stolpediagram
plt.bar(x, y)
# vis stolpediagram
plt.show()


"""Del c) Skriv et program som finner minste og lengste samtaletid som er loggført for uke 24. 
Svaret skrives til skjerm med informativ tekst"""
print()
print("Del c")
varihet_max = varighet.max()
varihet_min = min(varighet)

print(f"Minste samtaletid: {varihet_min} minutter")
print(f"Lengste samtaletid: {varihet_max} minutter")

"""Del d) KREVENDE: Skriv et program som regner ut gjennomsnittlig samtaletid basert på alle 
henvendelser i uke 24. """

print()
print("Del d")
# Konverterer varighet til datetime format
varighet_tid = pd.to_datetime(varighet, format='%H:%M:%S')

# Beregner gjennomsnittlig varighet
varighet_gjennomsnitt = varighet_tid.mean()

# Formaterer gjennomsnittlig varighet til string med formatet HH:MM:SS, hvis ikke blir resultatet slik 1900-01-01 00:06:40.009216256

varighet_gjennomsnitt_str = varighet_gjennomsnitt.strftime("%H:%M:%S")


print(f"Gjennomsnittlig samtaletid: {varighet_gjennomsnitt_str}")
print()

"""Del e) Supportvaktene i MORSE er delt inn i 2-timers bolker: kl 08-10, kl 10-12, kl 12-14 og kl 
14-16. Skriv et program som finner det totale antall henvendelser supportavdelingen mottok 
for hver av tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24. Resultatet visualiseres ved 
bruk av et sektordiagram (kakediagram). """


import matplotlib.pyplot as plt
import numpy as np

# lager en dict med tidsrommene
vakt_tidsrom = {"08-10":0, "10-12":0, "12-14":0, "14-16":0}
# print(kl_slett)
# ittererer igjennom kl_slett, men må splitte på :, hente [0] fra arrayen og gjøre om til integer. legger til 1 i dict for hver gang en tid dukker opp
for kl in kl_slett:
    kl_hele_time= int(kl.split(":")[0])
    
    if 8 <= kl_hele_time < 10:
        vakt_tidsrom["08-10"] += 1
    elif 10 <= kl_hele_time < 12:
        vakt_tidsrom["10-12"] += 1
    elif 12 <= kl_hele_time < 14:
        vakt_tidsrom["12-14"] += 1
    elif 14 <= kl_hele_time < 16:
        vakt_tidsrom["14-16"] += 1

# print(vakt_tidsrom)
# Lager en array med antall henvendelser for hvert tidsrom
y=np.array([vakt_tidsrom["08-10"],vakt_tidsrom["10-12"],vakt_tidsrom["12-14"],vakt_tidsrom["14-16"]])


# Legger til en streng før nøklene i vakt_tidsrom for å formatere teksten som benyttes til labels i kakediagram
labels = []
for key in vakt_tidsrom.keys():
    labels.append(f"Vakt tidsrom {key}")

# Lager kakediagram
plt.pie(y, labels=labels)
plt.show()


"""Del f) Kundens tilfredshet loggføres som tall fra 1-10 hvor 1 indikerer svært misfornøyd og 
10 indikerer svært fornøyd. Disse tilbakemeldingene skal så overføres til NPS-systemet (Net 
Promoter Score).  
NPS-systemet er konstruert på følgende måte: 
 
Score 1-6 oppfattes som at kunden er negativ (vil trolig ikke anbefale MORSE til andre). 
Score 7-8 oppfattes som et nøytralt svar. 
Score 9-10 oppfattes som at kunden er positiv (vil trolig anbefale MORSE til andre).  
Supportavdelingens NPS beregnes som et tall, prosentandelen positive kunder minus 
prosentandelen negative kunder. Ved en formel kan dette gis slik: 
NPS = % positive kunder - % negative kunder 
 
Et eksempel på utregning av NPS er gitt i figuren under.
 
 
 
 
 
Kilde: https://www.blueprnt.com/2018/09/17/net-promoter-score/ 
Lag et program som regner ut supportavdelings NPS og skriver svaret til skjerm. Merk: 
Kunder som ikke har gitt tilbakemelding på tilfredshet, skal utelates fra utregningene. """

negativt = 0
nøytralt = 0
positivt = 0

# ittererer igjennom score og legger til 1 i riktig kategori
for scores in score:
    if 1 <= scores <= 6:
        negativt += 1
    elif 7 <= scores <= 8:
        nøytralt += 1
    elif 9 <= scores <= 10:
        positivt += 1

print()
print("Del f")
print("Antall tilbakemeldinger i hver kategori:")
print(f"Antall negative: {negativt}")
print(f"Antall nøytrale: {nøytralt}")
print(f"Antall positive: {positivt}")
print()


# Beregner NPS score

nps = (positivt/len(score)*100) - (negativt/len(score)*100)


print(f"NPS score: {nps}")
