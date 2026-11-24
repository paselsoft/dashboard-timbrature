# Dashboard Timbrature — guida d'uso

## Fogli

| Foglio | A cosa serve |
|---|---|
| **Dashboard** | Riepilogo dell'anno: saldi mensili, residui di ferie e permessi, buoni pasto, grafico. L'anno è in **A1**. |
| **Festività** | Festività dell'anno, calcolate da sole (Pasqua compresa). Il patrono è il 1° dicembre. La riga 15 è libera per un'eventuale festività in più. |
| **Impostazioni** | Orario previsto per giorno della settimana, buono pasto, plafond annui, massimo L. 104 al mese, pausa minima. **Si modificano solo le celle gialle.** |
| **Gennaio … Dicembre** | Cartellino del mese: timbrature, assenze, saldo. |

## Inserimento dei dati (fogli mensili)

| Caso | Colonne da compilare |
|---|---|
| Giorno lavorato | Entr 1 / Usc 1 (ed eventualmente Entr 2 / Usc 2), oppure pulsante **Importa** dal portale |
| **Ferie, Malattia, L. 104** (giornata intera) | Solo **Tipo Assenza**. "Assenze" resta vuota: l'orario previsto è coperto in automatico |
| **Art. 23, Art. 25, Art. 34** (permessi a ore) | **Assenze** (ore di permesso) **e** Tipo Assenza. Per l'intera giornata, le ore previste di quel giorno |
| Sabato, domenica, festivi | Nulla (salvo timbrature straordinarie) |

Un giorno senza timbrature e senza assenze conta 0: i giorni futuri non falsano il saldo.

## Protezione dei fogli

Tutti i fogli sono protetti **senza password**. Si possono modificare solo le celle da compilare:

| Foglio | Celle modificabili |
|---|---|
| Mesi | Assenze, Entr/Usc, Tipo Assenza |
| Impostazioni | Le celle gialle |
| Dashboard | Residuo anno precedente (C4) e Ferie anno precedente (I7) |
| Festività | La riga libera 15 |

Per una modifica straordinaria (per esempio aggiungere una nota a una cella): **Revisione → Rimuovi protezione foglio**, oppure la macro **SbloccaFogli** per tutti i fogli insieme. Finito, esegui **ProteggiFogli**. Le macro del file gestiscono la protezione da sole.

## Su iPad

Excel per iPad **non esegue le macro**. Si possono inserire timbrature e assenze e vedere saldi, colori e Dashboard. Importazione, PDF, verifica e nuovo anno vanno fatti dal PC.

## Cosa significano i colori

| Dove | Colore | Significato |
|---|---|---|
| Assenze / Tipo Assenza | Rosso | Ore senza Tipo, oppure permesso a ore senza ore |
| Entr / Usc | Rosso | Orari non in sequenza (es. uscita prima dell'entrata) |
| Usc 1 / Entr 2 | Rosso | Pausa più breve della pausa minima (Impostazioni) |
| Giorno ed Entr/Usc | Giallo | Giorno lavorativo già passato senza timbrature né assenze (dimenticato?) |
| Tipo Assenza e riepilogo L. 104 | Rosso | Giorni di L. 104 oltre il massimo mensile |
| Riga del giorno | Evidenziata | Numero dispari di timbrature, oppure più di 10 ore lavorate |
| Dashboard, Residuo | Arancione | Sotto il 20% del plafond |
| Dashboard, Residuo | Rosso | Esaurito (zero o negativo) |

## Macro da usare

| Macro | Dove si avvia | Quando |
|---|---|---|
| **ImportaDaPortaleWeb** | Pulsante nei fogli mensili | Per caricare le timbrature da un file .txt copiato dal portale |
| **EsportaMeseInPDF** | Pulsante nei fogli mensili | PDF del singolo mese (`Timbrature_AAAA_MM_Mese.pdf`) |
| **VerificaAssenze** | Alt+F8 (o un pulsante sulla Dashboard) | Prima di un report: elenca le anomalie di tutto l'anno |
| **EsportaReportPDF** | Pulsante sulla Dashboard | PDF annuale (Dashboard + 12 mesi, una pagina ciascuno). Prima controlla i dati e, se trova anomalie, chiede se esportare comunque |
| **NuovoAnno** | Alt+F8 | **Una volta a inizio gennaio** (vedi sotto) |
| **ProteggiFogli** / **SbloccaFogli** | Alt+F8 | Rimettere / togliere la protezione su tutti i fogli |

All'apertura il file va da solo sul foglio del mese corrente.

## Inizio anno (gennaio)

1. Completa le timbrature di dicembre e **salva**.
2. Esegui **NuovoAnno** (Alt+F8) e controlla il riepilogo:
   - saldo orario riportato (tutto, anche negativo);
   - ferie residue riportate (tutte).
3. Conferma. Il file dell'anno concluso **resta invariato come archivio**. Il nuovo anno viene salvato come file nuovo (es. `..._2027.xlsm`) ed è quello aperto.
4. Nel nuovo file apri **Impostazioni** e aggiorna, se cambiano: orario previsto, valore del buono pasto, plafond, massimo L. 104 al mese.
5. Se cambia una festività locale, aggiornala nel foglio **Festività**.

Calendario, Pasqua e intestazioni si aggiornano da soli.

## Moduli VBA

| Modulo | Stato |
|---|---|
| `Modulo1` (ImportaDaPortaleWeb) | Da tenere |
| `Modulo2` (EsportaMeseInPDF) | Da tenere |
| `modExportPDF` (EsportaReportPDF) | Da tenere |
| `modControlliAssenze` (VerificaAssenze) | Da tenere: lo usa anche il PDF annuale per il controllo dei dati |
| `modNuovoAnno` (NuovoAnno) | Da tenere |
| `modProtezione` (ProteggiFogli, SbloccaFogli) | Da tenere |
| `modBackup` (copia di sicurezza automatica) | Da tenere |
| `modAggiornaFormule` | **Da rimuovere — non rieseguire**: riscriverebbe l'orario fisso e scollegherebbe Impostazioni |
| `modImpostazioni`, `modAvvisi`, `modPausa`, `modGiorniMancanti` | Già eseguiti: si possono rimuovere |

## Copie di sicurezza automatiche

A ogni salvataggio il file viene copiato nella sottocartella **Backup**, con la data nel nome (`..._AAAA-MM-GG.xlsm`). C'è una copia per giorno, e si tengono le ultime 10 giornate. Per tornare a una versione precedente: apri la copia e usa **File → Salva con nome** al posto del file originale.

## Se qualcosa va storto

- **Backup:** le macro che modificano le formule salvano prima un backup `..._backup_AAAA-MM-GG_hhmmss.xlsm` nella stessa cartella. Se un messaggio dice di chiudere senza salvare, fallo e riapri il backup.
- **OneDrive/SharePoint:** le esportazioni PDF e le macro con backup richiedono che il file sia aperto dalla cartella sincronizzata sul PC, non da un indirizzo web.
- **Importare un modulo `.bas`:** se lo scarichi da GitHub, usa il pulsante *Download raw file*. Tutti i moduli contengono solo caratteri semplici, quindi non si rovinano.
