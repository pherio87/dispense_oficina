# Sintassi del testo (Markdown)

Il file `.md` è letto da pandoc. Una riga vuota prima e dopo ogni elenco, tabella e blocco `:::` evita quasi tutti i problemi.

## Inizio del file

Si parte direttamente dal contenuto: niente intestazione YAML, copertina, intestazione o piè di pagina. Se il file ha un blocco `---` iniziale, viene ignorato.

## Struttura

```markdown
# Unità 1 – Titolo dell’unità          ← pagina nuova
## Capitolo 1 – Titolo del capitolo    ← pagina nuova
### Sezione                            ← prosegue nella pagina
```
Il testo “Unità 1 –” e “Capitolo 1 –” si scrive a mano. Il numero degli esercizi, delle figure e delle tabelle invece è automatico (le figure e le tabelle usano i campi di Word, quindi si rinumerano da sole se se ne aggiunge una).
Dopo un `#` metti una breve introduzione e “Al termine di questa unità sarai in grado di:” seguito da un elenco: la pagina di apertura resta volutamente leggera.

## Testo

- `**grassetto**` per i termini chiave. Pedici e apici: `H~2~O`, `m^2^`.
- Elenco puntato: righe che iniziano con `- `. Elenco numerato (una procedura): `1. `, `2. `. Gli elenchi brevi non si spezzano tra due pagine.
- Un paragrafo che finisce con `:` resta attaccato all’elenco o alla tabella che segue.
- Glossario (definizioni):
  ```markdown
  Adduzione
  :   tratto che porta l’acqua dalla rete all’edificio.
  ```

## Tabelle

```markdown
| Materiale | Come si unisce |
|---|---|
| Rame | Saldatura capillare |

: Didascalia della tabella (va dopo la tabella)
```
Solo per dati. La didascalia è facoltativa; se c’è, diventa “Tabella N –”. Una tabella con celle vuote diventa una tabella da completare (righe alte).

## Immagini

```markdown
![Didascalia che dice cosa guardare](img/foto.png){alt="Descrizione per chi non vede l’immagine" fonte="Archivio Oficina" width=60%}
```
- Il testo tra `[ ]` è la didascalia (“Figura N –”). Con `![](...)` l’immagine non ha didascalia.
- `alt` è il testo alternativo: descrive l’immagine in una frase. Se manca si usa la didascalia e lo script avvisa.
- `width=60%` o `width=8cm` (facoltativo). Senza, l’immagine usa la sua dimensione naturale a 150 dpi, senza uscire dai margini. Non supera mai gli 11 cm di altezza.
- `fonte` (facoltativo) aggiunge “Fonte: …” sotto la didascalia.
- Formati: PNG e JPG. Per la stampa servono almeno 1200 px di larghezza per un’immagine a tutta pagina.

Due immagini affiancate, con una didascalia sola:
```markdown
::: {.coppia caption="Raccordo a pressare (a) e a saldare (b)"}
![](img/a.png){alt="Raccordo a pressare"} ![](img/b.png){alt="Raccordo a saldare"}
:::
```

## Riquadri

```markdown
::: {.ricorda title="DA RICORDARE"}
Testo breve. **Una sola idea.**
:::

::: {.attenzione title="ATTENZIONE"}
Rischio e comportamento corretto, in poche righe.
:::
```
`ricorda` è bianco con bordo e barra blu. `attenzione` è pieno rosso con testo bianco: usalo solo per sicurezza, altrimenti perde forza. Dentro ci vanno solo paragrafi ed elenchi.

## Competenze (UC, capacità e conoscenze)

Ogni unità indica le competenze della qualifica regionale su cui lavora. Il riquadro va nella pagina di apertura dell’unità, subito dopo l’elenco “Al termine di questa unità sarai in grado di:”.

```markdown
::: {.competenze qualifica="04" codici="1.2 1.4 1.6 1.9"}
:::
```
- `qualifica`: il numero del file in `references/qualifiche/` (`04` = Operatore impianti elettrici) oppure una parte del titolo (`"impianti elettrici"`). Basta scriverla nel primo riquadro: i successivi la ereditano. Se una dispensa usa due qualifiche, si riscrive nel riquadro in cui cambia.
- `codici`: i codici Oficina della qualifica, separati da spazi. La prima cifra è l’unità di competenza, quindi non serve indicarla a parte. Gli intervalli si scrivono `1.5-1.8`. `UC2` da solo mostra solo il titolo dell’unità di competenza.
- Si scrivono **solo i codici**: lo script prende il testo dal file della qualifica, scrive “Unità di competenza 1 (UC1)” per esteso e mette in ordine capacità e conoscenze. Un codice inesistente viene segnalato.

Nella verifica finale si può ripetere lo stesso riquadro dell’unità, senza riscrivere i codici:
```markdown
::: {.competenze riepilogo=true title="COMPETENZE VERIFICATE"}
:::
```

Il riquadro è una tabella con bordo blu: in alto una banda blu con il titolo e la qualifica, poi per ogni unità di competenza il titolo e, affiancate, la colonna delle capacità e quella delle conoscenze. Capacità e conoscenze sono in Arial 10 pt (l’unica eccezione ai 12 pt), con il codice in blu e il testo allineato. Il riquadro resta intero: se non c’è spazio nella pagina di apertura passa alla pagina dopo. Scegli i codici che l’unità tratta davvero, non tutta l’unità di competenza.

## Esercizi

Ogni esercizio è un blocco `esercizio`; il titolo “Esercizio N · tipo” e il numero (che riparte a ogni unità) sono automatici.

```markdown
:::: {.esercizio tipo="Vero o falso"}
Segna con una crocetta se l’affermazione è vera (V) o falsa (F).

::: {.vf}
- Il DN coincide sempre con il diametro reale del tubo.
- Lo scarico funziona in pressione.
:::
::::
```

| Tipo | Come si scrive |
|---|---|
| Risposta multipla | una domanda, riga vuota, poi `A) …` `B) …` `C) …` `D) …` su righe consecutive |
| Vero o falso | blocco `.vf` con un elenco puntato (vedi sopra) |
| Risposta aperta | consegna + `::: {.righe n=4}` e `:::` su due righe: 4 righe per scrivere |
| Tabella da completare | consegna + tabella con celle vuote |
| Osserva e rispondi | consegna + immagine + domanda + `.righe` |
| Quiz lungo | tante domande, ognuna con le sue risposte `A)…D)`; si apre il blocco con `:::::` (cinque) se contiene altri blocchi |

Se un blocco ne contiene un altro, quello esterno usa più due punti (`::::` o `:::::`) di quello interno.

**Lunghezza.** Un esercizio resta intero su una pagina. Se la stima supera circa 21 cm lo script lo segnala e lo lascia spezzare tra una domanda e l’altra: è un’eccezione. Per i quiz lunghi, scrivi le domande come paragrafi che iniziano con `**1.**`, `**2.**`… seguiti dalle risposte. Se vuoi forzare la divisione: `{.esercizio tipo="..." split=true}`.

## Pagina nuova a mano

```markdown
::: {.nuovapagina}
:::
```
Raramente serve: unità e capitoli vanno già a pagina nuova.

## Non supportato

Note a piè di pagina, titoli oltre il livello 3 (diventano livello 3), blocchi di codice, formule, elenchi dentro le tabelle. Lo script stampa un avviso quando li incontra.
