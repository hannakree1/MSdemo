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
| Suurvaldkonnad | 4 ETIS-i suurvaldkonda |
| Alavaldkonnad | ETIS-i alavaldkonnad |
| Projektid | Kõik arvestatud projektid kontrollimiseks |

## Metoodika

- **Valdkonnad:** projekt ja selle rahastus jagatakse ETIS-i valdkondade vahel nende osakaalude järgi,
  nii et murdosad liituvad kokku täpselt projektide arvuks.
- **Asutused:** koondatud peaasutuse tasemele; mitme asutusega projekt jagatakse asutuste vahel võrdselt.
- **Eelarveaasta:** aastavahetust ületava rahastusperioodi summa jagatakse aastate vahel päevade järgi.
- **Välja jäetud kirjed:** `VALJA_JAETUD` skriptis (partnerikirje topeltloenduse vältimiseks ja ETIS-i testkirje).
- Jooksva ja tulevaste aastate summad võivad olla plaanitud, mitte välja makstud.
