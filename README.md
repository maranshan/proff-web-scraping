# proff-web-scraping
Hente ønskede nøkkeltall fra ønsket bedrift fra proff.no, og skriv til excel fil. Kan hente data fra resultat- og balanseregnskapet.

Bruker selenium for å navigere nettsiden og henteut spesifikk data. 

Fremgangsmåte:
1. Søk etter selskapet på proff.no, klikk deg inn på riktig selskap. Klikk på regnskap.
2. Kopier linken
3. Lim inn linken i variabelen URL
4. Bruk hent_data_resultat for å hente data om nøkkeltall fra resultatet, og hent_data_balanse for å hente data om nøkkeltall fra balanse.
5. Bruk funksjonen skriv_til_excel om du ønsker å skrive til excel-fil. (Husk å endre navnet på "filnavn" helt øverst til din fil.

hentKunRelevant henter 'Sum driftsinntekter', 'Sum salgsinntekter', 'Ordinære avskrivninger', 'Nedskrivning', 'Driftsresultat', 'Sum investeringer' og 'Sum egenkapital' 
Videre legger den sammen 'Ordinære avskrivninger', 'Nedskrivning' og 'Driftsresultat' for å finne EBITDA
Deretter skriver den til fil årlig driftsinntekter, salgsinntekter og EBITDA til excel

NB! Lagring til fil vil kun funke dersom fila ikke er åpen

*KUN MENT SOM ET PROSJEKT, OG IKKE EGNET FOR PROFESJONELT BRUK*
