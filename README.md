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
| Top vastutavad | 20 suurima PRG rahastusega vastutavat täitjat, nende EK projektid ja ERC grandid |
| EK projektid | Top vastutavate täitjate EK teadusprojektid |
| ERC grandid | Kõik Eesti osalusega ERC grandid CORDIS-est koos Eesti PI-ga |

## Metoodika

- **Valdkonnad:** Frascati klassifikaator. Projekt ja selle rahastus jagatakse valdkondade vahel
  osakaalude järgi, nii et murdosad liituvad kokku täpselt projektide arvuks. ETIS-is pole Frascati
  valdkondadel osakaalu; need on salvestatud paarides ETIS-i valdkondadega, seega võetakse osakaal
  paariliselt ETIS-i valdkonnalt.
- **Asutused:** koondatud peaasutuse tasemele; mitme asutusega projekt jagatakse asutuste vahel võrdselt.
- **Eelarveaasta:** aastavahetust ületava rahastusperioodi summa jagatakse aastate vahel päevade järgi.
- **Top vastutavad:** kui vastutav täitja vahetus, jagatakse PRG rahastus vastutuse perioodide järgi;
  20. kohaga viigis olevad võetakse kõik sisse.
- **EK projektid:** vastutava täitjana, rahastaja Euroopa Komisjon; Erasmus ja Interreg välja jäetud;
  katusprojekt jäetakse välja, kui selle alamprojekt on arvestuses; sama numbriga kirjed liidetakse.
- **ERC grandid:** CORDIS-e avaandmed (FP7, H2020, Euroopa Horisont), alla laaditud kausta
  `valjund/cordis` (umbes 125 MB, kord päevas). „Vastuvõtja Eestis“ = Eesti asutus on koordinaator
  või Synergy grandi osaleja. Kokkuvõttes on nii kõigi perioodide kui ka ainult H2020 ja
  Euroopa Horisondi arv, et tulemus oleks võrreldav analüüsidega, mis algavad 2014. aastast.
- **Välja jäetud kirjed:** `VALJA_JAETUD` skriptis (partnerikirje topeltloenduse vältimiseks ja ETIS-i testkirje).
- Jooksva ja tulevaste aastate summad võivad olla plaanitud, mitte välja makstud.
