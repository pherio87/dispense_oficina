# Layout standard delle dispense Oficina

Specifica da cui nasce `scripts/build_dispensa.js`. La skill formatta solo il contenuto: niente copertina, logo, intestazione né piè di pagina (li aggiunge Oficina a mano in Word). Misure in DXA (567 = 1 cm) e punti.

## Pagina

A4 (11906 × 16838), margini 2 cm su tutti i lati, larghezza utile 17 cm, altezza utile 25,7 cm. Intestazione e piè di pagina a 1,25 cm dal bordo: il logo che Oficina aggiunge a mano ci sta senza spostare il testo. Una sola sezione, intestazione e piè di pagina vuoti.

## Caratteri

Solo **Arial**, tre misure:

| Uso | Misura |
|---|---|
| Titoli e unità (`#`) | 20 pt grassetto, nero, filetto rosso sotto |
| Intestazioni: capitoli (`##`) in blu, sezioni ed esercizi (`###`) in nero | 16 pt grassetto |
| Capacità e conoscenze nel riquadro Competenze | 10 pt, interlinea 1,15 |
| Tutto il resto: testo, elenchi, tabelle, didascalie, riquadri | 12 pt, interlinea 1,5, allineato a sinistra (mai giustificato: disturba chi ha dislessia) |

Spaziatura: dopo il paragrafo 10 pt; elenchi 5 pt; `###` con 20 pt sopra.

## Colori (`assets/brand.json`)

Nero (testo, titoli, elenchi, bordi), bianco, blu `0057B8`, rosso `E30613`. Nessun grigio.
- Rosso: **solo** filetto sotto le unità e riquadro Attenzione.
- Blu: capitoli (`##`), intestazioni di tabella (testo bianco), riquadro Da ricordare.
- Elenchi puntati, numerati e alfabetici: sempre neri.

## Pagine e interruzioni

- **`#` e `##`**: `pageBreakBefore` nello stile, quindi unità e capitoli iniziano sempre in pagina nuova.
- **Esercizi**: ogni paragrafo ha "mantieni con il successivo" tranne l'ultimo, così l'esercizio resta intero. Se è più lungo di una pagina la catena si interrompe tra una domanda e l'altra.
- **Tabelle** fino a 14 righe: righe che non si spezzano, intestazione ripetuta, tutte "con il successivo".
- **Elenchi** fino a 7 puntati / 10 numerati: restano insieme. Un paragrafo che finisce con ":" resta con ciò che introduce.
- **Immagine + didascalia**: sempre insieme.

## Immagini

Centrate; larghezza massima 17 cm; altezza massima 12 cm; testo alternativo obbligatorio. Didascalia sotto con campo `SEQ Figura`; quella delle tabelle sta sopra con `SEQ Tabella`. Due immagini affiancate stanno in un solo paragrafo con una didascalia.

## Riquadri

Paragrafi con bordo e sfondo, non tabelle. **Da ricordare**: bordo e barra blu, fondo bianco. **Attenzione**: fondo rosso pieno, testo bianco. **Competenze**: unica eccezione, è una tabella (codici e testi del repertorio sono dati). Bordo blu 1,5 pt; riga di intestazione blu piena con “COMPETENZE” (12 pt) e la qualifica (10 pt) in bianco; per ogni UC una riga con il titolo (12 pt, blu e nero) e una riga a due colonne, Capacità e Conoscenze, divise da un filetto blu sottile; voci in 10 pt con codice blu in grassetto e rientro sporgente di 1,06 cm. Uno per unità, dopo gli obiettivi, sempre intero (righe che non si spezzano, “mantieni con il successivo”). Stili `CompTitle` e `CompText`; testi da `references/qualifiche/`.

## Trappole note di docx-js (già risolte nello script)

- `w:lineRule` va scritto sempre, altrimenti LibreOffice taglia le immagini.
- Gli stili Heading 1-3 vanno ridefiniti in `styles.default`, non in `paragraphStyles`.
- I bordi dei paragrafi escono nell'ordine sbagliato per lo schema OOXML: si riordinano a file scritto.
- Gli stili predefiniti Heading 4-6 e Hyperlink sono blu: vanno ridefiniti.
