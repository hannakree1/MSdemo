# MSdemo – PRG ülevaade ETIS-ist

Skript laadib ETIS-i avalikust API-st kõik rühmagrandid (PRG) ja koostab neist Exceli ülevaate.
Käsitsi eksporti portaalist pole vaja.

## Käivitamine

```
pip3 install openpyxl
python3 prg_ulevaade.py
```

Tulemus: `valjund/PRG_ulevaade_<kuupäev>.xlsx`. Iga käivitus võtab ETIS-ist värsked andmed.

## Exceli lehed

| Leht | Sisu |
|---|---|
| Kokkuvõte | Kogusummad, metoodika, välja jäetud kirjed |
| Algusaasta | Alustatud PRG-de arv ja nende kogurahastus algusaasta järgi |
| Eelarveaasta | Aktiivsete PRG-de arv ja rahastus eelarveaasta järgi |
| Asutused | Projektide arv ja rahastus peaasutuse kaupa |
| Frascati põhivaldkonnad | 6 Frascati põhivaldkonda |
| Frascati alavaldkonnad | Frascati alavaldkonnad |
| Projektid | Kõik arvestatud projektid kontrollimiseks |

## Metoodika

- **Valdkonnad:** Frascati klassifikaator. Projekt ja selle rahastus jagatakse valdkondade vahel
  osakaalude järgi, nii et murdosad liituvad kokku täpselt projektide arvuks. ETIS-is pole Frascati
  valdkondadel osakaalu; need on salvestatud paarides ETIS-i valdkondadega, seega võetakse osakaal
  paariliselt ETIS-i valdkonnalt.
- **Asutused:** koondatud peaasutuse tasemele; mitme asutusega projekt jagatakse asutuste vahel võrdselt.
- **Eelarveaasta:** aastavahetust ületava rahastusperioodi summa jagatakse aastate vahel päevade järgi.
- **Välja jäetud kirjed:** `VALJA_JAETUD` skriptis (partnerikirje topeltloenduse vältimiseks ja ETIS-i testkirje).
- Jooksva ja tulevaste aastate summad võivad olla plaanitud, mitte välja makstud.
