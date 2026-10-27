# Töölogi

## 27.10.2026 – PRG ülevaade ETIS-ist (demo MS-iga)

Hanna lõi kausta ja privaatse GitHubi repositooriumi ning andis sessiooni üle kolleeg MS-ile.
Ülesanne: PRG-de ülevaade otse ETIS-ist, ilma portaalist Excelisse eksportimata.

### Andmeallikas
- ETIS-i avalik API `https://www.etis.ee:2346/api/project/getitems`, filter `ProgrammeCode=PRG`.
  Ametlikku dokumentatsiooni ei leitud; filter `PersonId=<guid>` annab inimese kõik projektid.
- Tulemus: Excel (sest MS kasutab Excelit) + skript GitHubis, et ülevaadet saaks uuesti teha.

### Otsused (MS)
| # | Küsimus | Otsus |
|---|---|---|
| 1 | Mida ülevaade näitab | Aastate, asutuste ja valdkondade kaupa; valdkonnas jagatakse projekt osakaalude järgi, nii et arv liitub õigeks |
| 2 | „Aastate kaupa“ | Mõlemad: algusaasta ja eelarveaasta |
| 3 | Valdkondade tase | Põhi- ja alavaldkonnad |
| 4 | Asutuste tase | Peaasutus; mitme asutusega projekt (PRG1667) jagatakse 50/50 |
| 5 | Valdkonna klassifikaator | Algul ETIS, hiljem **Frascati**. Frascati valdkondadel pole ETIS-is osakaalu, osakaal võetakse paariliselt ETIS-i valdkonnalt (kõik 47 ETIS-i koodi andsid sisuliselt sobiva paari) |
| 6 | Numbrita partnerikirje „Nutika ühiskonna…“ (TLÜ, 59 000 €) | Välja – sama projekt mis PRG1919, muidu topeltloendus |
| 7 | `HANGFIRETEST1` (1 päev, 100 000 €, KBFI) | Välja – ETIS-i süsteemi testkirje |
| 8 | Kahe vastutava täitjaga PRG-d (9 tk, täitja vahetus) | Rahastus jagatakse vastutuse perioodi järgi |
| 9 | Viik 20. kohal (10 inimest summaga 1 851 375 €) | Kõik viigis olevad sisse → 22 inimest |
| 10 | Mis on EK rahastus | Rahastaja „Euroopa Komisjon“ / „European Commission Directorate“; ainult teadusprogrammid – Erasmus ja Interreg välja (Erasmus ka siis, kui on programmi „Muu“ all) |
| 11 | Roll EK projektis | Ainult vastutav täitja |
| 12 | TeamPerMed katus- ja alamprojekt | Ainult EK alamprojekt (2,88 mln €); katus sisaldab riigi kaasrahastust |
| 13 | cGEM – kaks kirjet sama numbriga 810645 | Üks projekt, summad kokku (2,5 mln €) |

Otsused 12–13 on skriptis üldreeglitena: katusprojekt jäetakse välja, kui selle alamprojekt on
arvestuses; sama projektinumbriga kirjed liidetakse.

### Tulemused
- **405 PRG-d, 313,54 mln €** (2 kirjet välja jäetud). Kõikidel lehtedel samad summad;
  valdkondade murdosad liituvad 405-ks.
- Frascati: loodusteadused 47,5% rahast, tehnika 13,2%, arsti- ja terviseteadused 14,8%,
  põllumajandus 5,2%, sotsiaalteadused 8,2%, humanitaarteadused 11,1%.
- Top 22 vastutavat täitjat: 17-l on EK teadusprojekte vastutava täitjana.
- **ERC:** top 22 seas on ERC saajaid 3 – Tambet Teesalu (GLIOMADDS, GLIOGUIDE, MINerVA),
  Ülo Niinemets (SIP-VOL+, CRYPTO-AI-FUTURE), Ülo Mander (PeatlandN2O).
- **ERC grandid Eestis (CORDIS, 05.10.2026):** Eesti osalusega 36; neist vastuvõtja Eestis 29.
  Ainult H2020 + Euroopa Horisont: 30 (vastuvõtja Eestis 25).

### Kontrollid ja leitud vead
- Suurvaldkondade lehelt puudus arvu veerg (minu koodiviga) – parandatud.
- Ivo Leito „EACH Erasmus Mundus“ oli programmi „Muu“ all – välja jäetud (otsus 10).
- Mait Metspalu EK summa oli katusprojekti ja topeltkirjete tõttu 23,4 mln € – pärast otsuseid
  12–13 5,42 mln €. (Vahepeal ütlesin vestluses ekslikult 5,39 mln €.)
- Esialgu väitsin „5 ERC granti 4 inimesel“ – õige on 6 granti 3 inimesel.
- **ERC 36 vs ühe eelneva analüüsi 30:** sama CORDIS-e allikas, nimekirjad kattuvad täielikult.
  Vahe: eelnev analüüs võtab ainult H2020 ja Euroopa Horisondi (6 FP7 granti puuduvad);
  minu 29 jätab välja grandid, kus Eesti asutus on ainult partner.
- 3 ERC grandi PI-d (Loog, Sõukand, Kostakis) ei leia ETIS-ist automaatselt – võetud käsitsi
  kontrollitud vastetest. INVPROB, PSOPA, DEPRIVEDHOODS: Eesti PI-d ei leitud.
- Teadmiseks: PRG1296 ja PRG1473 on sama pealkirja ja vastutava täitjaga (tõenäoliselt jätk) –
  jäetud kaheks, nagu ETIS-is. PRG2043 rahastus on peatatud alates 04.2026.

### Avalikustamine
- MS otsustas repositooriumi avalikuks teha; viited eelnevale analüüsile muudeti enne seda
  neutraalseks ka git-ajaloos. Vana repositoorium kustutati ja loodi uus puhta ajalooga.
- Repositooriumi lisati ka tulemus (Excel); CORDIS-e toorandmed (~125 MB) jäetakse välja.

### Lahtised küsimused
- Jooksva ja tulevaste aastate summad võivad olla plaanitud, mitte välja makstud.
- CORDIS-es on ERC grandid alles pärast lepingu sõlmimist; värsked voorud võivad puududa.
