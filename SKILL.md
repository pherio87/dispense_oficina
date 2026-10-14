---
name: dispense-oficina
description: Formatta in Word (.docx) il contenuto delle dispense didattiche di Oficina con il layout standard: A4, Arial 12/16/20, unità e capitoli su pagina nuova, immagini con didascalia, esercizi interi in pagina, riquadri "Da ricordare" e "Attenzione", riquadro Competenze di ogni unità con UC, capacità e conoscenze della qualifica regionale. Solo il contenuto: niente copertina, intestazione o piè di pagina. Usala quando si chiede di scrivere, impaginare o convertire una dispensa, un manuale, degli appunti o delle schede per un corso, un modulo o un'unità formativa (termoidraulici, IeFP, 18-25, MSNA, SIA), anche se non si nomina Word né il layout, e quando si chiede di indicare UC, capacità o conoscenze in una dispensa o in una verifica.
---

# Dispense Oficina

Trasforma un testo Markdown in un `.docx` con il layout standard di Oficina. Formatta **solo il contenuto**: copertina, logo, intestazione e piè di pagina non ci sono (li aggiunge Oficina a mano). Lo script fa tutto in pochi secondi: il lavoro sta nel scrivere bene il testo, non nel ricontrollare l'impaginazione.

## Requisiti

Servono Node con il pacchetto `docx` (`npm i -g docx` se manca), `pandoc` e LibreOffice. Nel resto `<skill>` è la cartella di questa skill.

## Procedura

1. **Scrivi il Markdown** seguendo `references/sintassi.md` (leggilo la prima volta). Salva `dispensa.md` con le immagini in `img/`. Se parti da appunti `.docx`: `pandoc appunti.docx -t markdown -o dispensa.md --extract-media=img`. Non serve intestazione YAML.
2. **Indica le competenze di ogni unità.** Apri il file della qualifica in `references/qualifiche/` (se non sai quale, chiedilo). Per ogni unità scegli le capacità e le conoscenze che il testo tratta davvero e scrivi il riquadro `.competenze` con i soli codici, dopo l'elenco degli obiettivi. Non inventare codici e non ricopiare i testi: li mette lo script. Se nessun codice corrisponde a un argomento, dillo nella consegna.
3. **Costruisci:** `node <skill>/scripts/build_dispensa.js dispensa.md -o dispensa.docx` e correggi gli avvisi che stampa (immagini mancanti o senza testo alternativo, esercizi lunghi, codici inesistenti, unità senza competenze).
4. **Controlla una volta sola, alla fine:** `python <skill>/scripts/check_dispensa.py dispensa.docx` (con `--veloce` salta il rendering PDF). Se dice "TUTTO OK" hai finito: **non** convertire in PDF né guardare le pagine una a una, a meno che il controllo segnali problemi o la persona lo chieda.
5. **Consegna** il `.docx` con una riga su cosa contiene, sugli eventuali avvisi e su quali unità di competenza copre (lo script le elenca alla fine).

## Il layout in breve

| Cosa | Regola |
|---|---|
| Pagina | A4, margini 2 cm su tutti i lati. Nessuna copertina, intestazione o piè di pagina |
| Carattere | Solo Arial: **10 pt** solo per capacità e conoscenze del riquadro Competenze; **12 pt** testo, elenchi, tabelle, didascalie; **16 pt** intestazioni (`##` capitolo in blu, `###` sezione in nero); **20 pt** titoli e unità (`#`) con filetto rosso |
| Pagine | `#` e `##` sempre in cima a una pagina nuova |
| Colori | Nero, bianco, blu #0057B8, rosso #E30613. **Mai grigi.** Rosso solo per il filetto delle unità e il riquadro Attenzione; elenchi sempre neri; blu per capitoli, intestazioni di tabella e riquadro Da ricordare |
| Testo | Mai dentro tabelle: le tabelle servono per i dati. Riquadri e righe di scrittura sono paragrafi con bordo |
| Immagini | Centrate, entro i margini, al massimo 12 cm di altezza, con testo alternativo e didascalia numerata |
| Competenze | Un riquadro per unità dopo gli obiettivi: tabella con bordo e banda blu, titolo di ogni UC e due colonne (capacità, conoscenze) in 10 pt con codice blu. Resta intero; se non ci sta va alla pagina dopo |
| Esercizi | Restano interi sulla stessa pagina; se sono troppo lunghi si spezzano tra una domanda e l'altra |

## Scrivere per questi lettori

Ragazzi con DSA e BES, percorsi brevi (MSNA): conta la chiarezza. Pratica prima, teoria dopo (metodo induttivo). Le regole complete di Oficina sono in `references/guida-formattazione.md`: leggila prima di scrivere o riscrivere un testo, soprattutto la parte "Stile di scrittura". Il layout dello script la applica già; dove se ne discosta, vale `references/layout-standard.md`.
- Frasi brevi, un'idea per paragrafo; parole tecniche spiegate alla prima comparsa.
- Termini chiave in **grassetto**; niente corsivo, MAIUSCOLO o sottolineature.
- Procedure come elenchi numerati, un passaggio per voce.
- Ogni capitolo apre con ciò che si impara e chiude con un esercizio pratico.
- Tante immagini, ognuna con una didascalia che dice che cosa guardare.
- Versioni MSNA: stessa struttura, frasi più corte, vocabolario semplice, più immagini.

I dati tecnici e normativi (misure, temperature, D.Lgs. 81/2008) vanno controllati: se non sei certo, segnalalo nella consegna.

Le qualifiche RER stanno in `references/qualifiche/`, un file per corso, con i testi dei PDF regionali e i codici Oficina (la prima cifra è la UC, poi capacità e conoscenze numerate di seguito). Sono l'unica fonte: se un testo ha il segno […] è incompleto nel PDF regionale e lo script lo segnala. Per aggiungere una qualifica basta un nuovo file con la stessa struttura. Non rinumerare i codici esistenti: le dispense già fatte li usano.

## Cambiare colori o misure

- Colori: `assets/brand.json` (`primary`, `accent`, `secondary`). Misure: costante `T` in cima a `scripts/build_dispensa.js`.
