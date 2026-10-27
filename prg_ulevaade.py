"""PRG (rühmagrantide) ülevaade otse ETIS-i avalikust API-st.

Käivitamine:  python3 prg_ulevaade.py
Tulemus:      valjund/PRG_ulevaade_<kuupäev>.xlsx

Metoodika:
- Andmed: https://www.etis.ee:2346/api/project/getitems, filter ProgrammeCode=PRG.
- Algusaasta: projekt arvestatakse aastasse, mil see algas; summa = projekti kogurahastus.
- Eelarveaasta: iga rahastusperioodi summa jagatakse aastate vahel päevade järgi;
  aktiivne projekt = projekt, millel on sel aastal rahastust.
- Asutused: koondatud peaasutuse tasemele (nimi enne esimest koma). Mitme asutusega
  projekt jagatakse asutuste vahel võrdselt.
- Valdkonnad: Frascati klassifikaator. Projekt (ja selle rahastus) jagatakse valdkondade
  vahel osakaalude järgi, seega murdosad liituvad kokku täpselt projektide arvuks.
  Frascati valdkondadel ETIS-is osakaalu pole; ETIS hoiab need paarides ETIS-i
  valdkondadega, seega võetakse osakaal samal kohal olevalt ETIS-i valdkonnalt.
- Top vastutavad täitjad: PRG rahastus vastutava täitja kaupa; kui vastutav täitja
  vahetus, jagatakse rahastus vastutuse perioodide järgi. Võetakse 20 suurimat koos
  20. kohaga viigis olevatega.
- EK projektid: projektid, kus inimene on vastutav täitja ja rahastaja on Euroopa
  Komisjon; Erasmus ja Euroopa territoriaalne koostöö (Interreg) jäetakse välja.
- ERC grandid: CORDIS-e avaandmed (FP7, H2020, Euroopa Horisont), kõik ERC projektid,
  kus osaleb Eesti asutus. „Vastuvõtja Eestis“ = Eesti asutus on koordinaator või
  Synergy grandi osaleja. Eesti PI ETIS-ist grandinumbri või akronüümi järgi.
"""

import csv
import io
import json
import urllib.request
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

API = "https://www.etis.ee:2346/api/project/getitems?Format=json&{filter}&Take={take}&Skip={skip}"
LEHT = 200
TOP = 20

EK_RAHASTAJAD = {"Euroopa Komisjon", "European Commission Directorate"}
# EK rahastusega, aga mitte teadusprogrammid: Erasmus, Euroopa territoriaalne koostöö (Interreg)
EK_VALJA_PROGRAMMID = {"426", "414"}

# ETIS-i kirjed, mis ei ole iseseisvad PRG projektid: Guid -> põhjus
VALJA_JAETUD = {
    "8fef4bfa-2ba2-456e-ac5a-1ab12772354f": "Numbrita Tallinna Ülikooli partnerikirje projektile PRG1919 (topeltloendus)",
    "ebecf66e-1ec0-4923-a01b-cacf5c4abb92": "ETIS-i süsteemi testkirje",
}

CORDIS = "https://cordis.europa.eu/data/cordis-{fp}projects-csv.zip"
RAAMPROGRAMMID = {"fp7": "FP7", "h2020": "H2020", "HORIZON": "Euroopa Horisont"}
ERC_LIIGID = [("SYG", "Synergy"), ("POC", "Proof of Concept"), ("ADG", "Advanced"), ("COG", "Consolidator"),
              ("STG", "Starting"), ("ERC-AG", "Advanced"), ("ERC-CG", "Consolidator"), ("ERC-SG", "Starting")]
# ETIS-i programmid, kust ERC grandi vastutavat täitjat otsitakse
ETIS_EL_PROGRAMMID = ["439", "440", "442", "450", "403", "404", "137", "MUU"]
# ERC grandid, mille Eesti PI-d ETIS-ist grandinumbri ega akronüümi järgi ei leia;
# kontrollitud ühes eelnevas analüüsis
ERC_PI_KASITSI = {"649124": "Mart Loog", "714874": "Renata Sõukand", "802512": "Vasileios Kostakis"}

POHIVALDKONNAD = {
    "1": "Loodusteadused",
    "2": "Tehnika ja tehnoloogia",
    "3": "Arsti- ja terviseteadused",
    "4": "Põllumajandusteadused",
    "5": "Sotsiaalteadused",
    "6": "Humanitaarteadused",
}


def lae_projektid(filter="ProgrammeCode=PRG"):
    projektid, skip = [], 0
    while True:
        with urllib.request.urlopen(API.format(filter=filter, take=LEHT, skip=skip), timeout=120) as vastus:
            osa = json.load(vastus)
        projektid += osa
        if len(osa) < LEHT:
            return projektid
        skip += LEHT


def kp(tekst):
    p, k, a = tekst.split(".")
    return date(int(a), int(k), int(p))


def summad_aastate_kaupa(projekt):
    """Rahastus aastate kaupa; aastavahetust ületav periood jagatakse päevade järgi."""
    aastad = defaultdict(float)
    for periood in projekt["FinancingPeriods"]:
        algus, lopp, summa = kp(periood["PeriodStartDate"]), kp(periood["PeriodEndDate"]), periood["Amount"] or 0
        paevi = (lopp - algus).days + 1
        for aasta in range(algus.year, lopp.year + 1):
            a = max(algus, date(aasta, 1, 1))
            l = min(lopp, date(aasta, 12, 31))
            aastad[aasta] += summa * ((l - a).days + 1) / paevi
    return aastad


def valdkonnad(projekt):
    """Frascati valdkonnad koos osakaaluga, mis võetakse paarilise ETIS-i valdkonna juurest.
    Sama Frascati kood mitme ETIS-i valdkonna all liidetakse kokku."""
    osad, nimed = defaultdict(float), {}
    for e, f in zip(projekt["ResearchAreasEtis"], projekt["ResearchAreasFrascati"], strict=True):
        osad[f["Code"]] += (e["Percent"] or 0) / 100
        nimed[f["Code"]] = f["Name"]
    return [(kood, nimed[kood], osa) for kood, osa in osad.items()]


def asutused(projekt):
    return sorted({i["Name"].split(",")[0].strip() for i in projekt["Institutions"]})


def vastutavad(projekt):
    return ", ".join(x["Name"] for x in projekt["Persons"] if x["RoleName"] == "Vastutav täitja")


def top_vastutavad(projektid):
    """PRG rahastus vastutava täitja kaupa; TOP suurimat koos viigis olevatega.
    Tagastab [(guid, nimi, summa, prg_arv)]."""
    summa, arv, nimi = defaultdict(float), defaultdict(int), {}
    for p in projektid:
        vastutajad = [x for x in p["Persons"] if x["RoleName"] == "Vastutav täitja"]
        for x in vastutajad:
            nimi[x["Guid"]] = x["Name"]
            arv[x["Guid"]] += 1
            if len(vastutajad) == 1:
                summa[x["Guid"]] += p["FinancingInPeriodsTotal"] or 0
                continue
            a, l = kp(x["PeriodStartDate"]), kp(x["PeriodEndDate"])
            for periood in p["FinancingPeriods"]:
                pa, pl = kp(periood["PeriodStartDate"]), kp(periood["PeriodEndDate"])
                kattuvus = (min(l, pl) - max(a, pa)).days + 1
                if kattuvus > 0:
                    summa[x["Guid"]] += (periood["Amount"] or 0) * kattuvus / ((pl - pa).days + 1)
    jarjestus = sorted(summa, key=lambda g: -summa[g])
    piir = round(summa[jarjestus[TOP - 1]])
    return [(g, nimi[g], summa[g], arv[g]) for g in jarjestus if round(summa[g]) >= piir]


def ek_projektid(guid):
    """Inimese EK teadusprojektid, kus ta on (olnud) vastutav täitja.
    Katusprojekt jäetakse välja, kui selle alamprojekt on arvestuses (katus sisaldab
    ka kaasrahastust); sama projektinumbriga kirjed liidetakse üheks projektiks."""
    ek = [p for p in lae_projektid(f"PersonId={guid}")
            if any(x["Guid"] == guid and x["RoleName"] == "Vastutav täitja" for x in p["Persons"])
            and {i["Name"] for i in p["FinancingInstitutions"]} & EK_RAHASTAJAD
            and p["ProgrammeCode"] not in EK_VALJA_PROGRAMMID
            # Erasmus võib olla ka programmi "Muu" all
            and "erasmus" not in f'{p["Acronym"]} {p["Title"]} {p["TitleEng"]}'.lower()]

    katused = {r["Guid"] for p in ek for r in p["RoofProjects"] if r["Guid"] != p["Guid"]}
    uhendatud = {}
    for p in ek:
        if p["Guid"] in katused:
            continue
        voti = p["FinancierProjectNr"] or p["Guid"]
        if voti not in uhendatud:
            uhendatud[voti] = dict(p)
            continue
        k = uhendatud[voti]
        k["FinancingInPeriodsTotal"] = (k["FinancingInPeriodsTotal"] or 0) + (p["FinancingInPeriodsTotal"] or 0)
        k["Institutions"] = k["Institutions"] + p["Institutions"]
        k["Acronym"] = k["Acronym"] or p["Acronym"]
    return list(uhendatud.values())


def erc_liik(*tekstid):
    tekst = " ".join(tekstid).upper()
    return next((nimi for voti, nimi in ERC_LIIGID if voti in tekst), "")


def erc_grandid(kaust):
    """Eesti osalusega ERC grandid CORDIS-e avaandmetest; failid hoitakse päeva kaupa kaustas."""
    csv.field_size_limit(10**9)
    kaust.mkdir(parents=True, exist_ok=True)
    grandid = []
    for fp, raamprogramm in RAAMPROGRAMMID.items():
        tee = kaust / f"{date.today():%Y-%m-%d}_cordis_{fp}.zip"
        if not tee.exists():
            urllib.request.urlretrieve(CORDIS.format(fp=fp), tee)
        with zipfile.ZipFile(tee) as z:
            with z.open("project.csv") as f:
                erc = {r["id"]: r for r in csv.DictReader(io.TextIOWrapper(f, encoding="utf-8"), delimiter=";")
                       if r["fundingScheme"].startswith(("ERC", "HORIZON-ERC"))}
            with z.open("organization.csv") as f:
                for o in csv.DictReader(io.TextIOWrapper(f, encoding="utf-8"), delimiter=";"):
                    if o["projectID"] not in erc or o["country"] != "EE":
                        continue
                    p = erc[o["projectID"]]
                    liik = erc_liik(p["topics"], p["masterCall"], p["fundingScheme"])
                    grandid.append({
                        "nr": p["id"], "akronuum": p["acronym"], "pealkiri": p["title"], "raamprogramm": raamprogramm,
                        "liik": liik, "voor": p["masterCall"], "algus": p["startDate"], "asutus": o["name"],
                        "roll": "koordinaator" if o["role"] == "coordinator" else "partner",
                        "eestis": o["role"] == "coordinator" or liik == "Synergy",
                        "toetus": float(o["ecContribution"] or 0),
                    })
    return sorted(grandid, key=lambda g: (g["algus"], g["nr"]))


def lisa_erc_pi(grandid):
    """Eesti PI ETIS-ist: vastutav täitja projektil, mille number sisaldab grandinumbrit
    või mille akronüüm on sama."""
    etis = [p for kood in ETIS_EL_PROGRAMMID for p in lae_projektid(f"ProgrammeCode={kood}")]
    for g in grandid:
        vasted = [p for p in etis if g["nr"] in (p["FinancierProjectNr"] or "")
                  or (p["Acronym"] or "").strip().lower() == g["akronuum"].lower()]
        isikud = {x["Guid"]: x["Name"] for p in vasted for x in p["Persons"] if x["RoleName"] == "Vastutav täitja"}
        g["pi_guid"] = set(isikud)
        g["pi"] = ", ".join(sorted(set(isikud.values()))) or ERC_PI_KASITSI.get(g["nr"], "")


def arvuta(projektid):
    algusaasta = defaultdict(lambda: [0, 0.0])
    eelarveaasta = defaultdict(lambda: [0, 0.0])
    asutus = defaultdict(lambda: [0.0, 0.0])
    suur = defaultdict(lambda: [0.0, 0.0])
    ala = defaultdict(lambda: [0.0, 0.0])
    ala_nimi = {}

    for p in projektid:
        kokku = p["FinancingInPeriodsTotal"] or 0

        a = algusaasta[kp(p["ProjectStartDate"]).year]
        a[0] += 1
        a[1] += kokku

        for aasta, summa in summad_aastate_kaupa(p).items():
            if summa > 0:
                eelarveaasta[aasta][0] += 1
                eelarveaasta[aasta][1] += summa

        asut = asutused(p)
        for nimi in asut:
            asutus[nimi][0] += 1 / len(asut)
            asutus[nimi][1] += kokku / len(asut)

        for kood, nimi, osa in valdkonnad(p):
            ala_nimi[kood] = nimi
            for rida in (ala[kood], suur[kood.split(".")[0]]):
                rida[0] += osa
                rida[1] += kokku * osa

    return algusaasta, eelarveaasta, asutus, suur, ala, ala_nimi


PAIS = Font(bold=True, color="FFFFFF")
PAIS_TAUST = PatternFill("solid", fgColor="2F5597")
KOKKU = Font(bold=True)
EUR = '#,##0 "€"'
ARV = "0.00"
PROTS = "0.0%"


def leht(wb, nimi, pais, read, vormingud, laiused, kokku_rida=None):
    ws = wb.create_sheet(nimi)
    ws.append(pais)
    for lahter in ws[1]:
        lahter.font, lahter.fill = PAIS, PAIS_TAUST
        lahter.alignment = Alignment(wrap_text=True, vertical="center")
    for rida in read:
        ws.append(rida)
    if kokku_rida:
        ws.append(kokku_rida)
        for lahter in ws[ws.max_row]:
            lahter.font = KOKKU
    for veerg, vorming in enumerate(vormingud, start=1):
        if vorming:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, veerg).number_format = vorming
    for veerg, laius in enumerate(laiused, start=1):
        ws.column_dimensions[get_column_letter(veerg)].width = laius
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(pais))}{len(read) + 1}"
    return ws


def osakaaluga(andmed, koguarv, kogusumma):
    return [(v, arv, arv / koguarv, summa, summa / kogusumma) for v, (arv, summa) in andmed]


def kirjuta_excel(koik, tee):
    erc = erc_grandid(tee.parent / "cordis")
    lisa_erc_pi(erc)
    projektid = [p for p in koik if p["Guid"] not in VALJA_JAETUD]
    valja = [p for p in koik if p["Guid"] in VALJA_JAETUD]
    algusaasta, eelarveaasta, asutus, suur, ala, ala_nimi = arvuta(projektid)
    n = len(projektid)
    raha = sum(p["FinancingInPeriodsTotal"] or 0 for p in projektid)
    numbrita = [p for p in projektid if not p["FinancierProjectNr"]]

    wb = Workbook()
    ws = wb.active
    ws.title = "Kokkuvõte"
    ws.append(["PRG ülevaade (ETIS)"])
    ws["A1"].font = Font(bold=True, size=14)
    for rida in [
        [],
        ["Andmed laaditud", date.today().strftime("%d.%m.%Y")],
        ["Allikas", "ETIS avalik API, ProgrammeCode=PRG"],
        ["PRG projekte kokku", n],
        ["Rahastus kokku", raha],
        ["Projektinumbrita kirjeid", len(numbrita)],
        [],
        ["Metoodika"],
        ["Algusaasta", "Projekt loetakse aastasse, mil see algas; summa on projekti kogurahastus."],
        ["Eelarveaasta", "Rahastus jagatud aastate vahel päevade järgi; aktiivne = sel aastal rahastust."],
        ["Asutused", "Peaasutuse tasemel; mitme asutusega projekt jagatud asutuste vahel võrdselt."],
        ["Valdkonnad", "Frascati klassifikaator; projekt ja rahastus jagatud valdkondade vahel osakaalude järgi."],
        ["", "Osakaal võetud paariliselt ETIS-i valdkonnalt, sest Frascati valdkondadel ETIS-is osakaalu pole."],
        ["Top vastutavad", f"PRG rahastus vastutava täitja kaupa (vahetuse korral vastutuse perioodi järgi); {TOP} suurimat koos viigis olevatega."],
        ["EK projektid", "Vastutava täitjana; rahastaja Euroopa Komisjon; Erasmus ja Interreg välja jäetud. Summa on ETIS-is kajastatud rahastus."],
        ["ERC grandid", "CORDIS-e avaandmed (FP7, H2020, Euroopa Horisont); kõik ERC projektid, kus osaleb Eesti asutus."],
        ["", "Vastuvõtja Eestis = Eesti asutus on koordinaator või Synergy grandi osaleja. CORDIS-es on grandid alles pärast lepingu sõlmimist."],
        ["Märkus", "Jooksva ja tulevaste aastate summad võivad olla plaanitud, mitte välja makstud."],
        [],
        ["Välja jäetud kirjed"],
    ]:
        ws.append(rida)
    ws.cell(ws.max_row, 1).font = KOKKU
    for p in valja:
        summa = f'{p["FinancingInPeriodsTotal"]:,.0f}'.replace(",", " ")
        ws.append([p["FinancierProjectNr"] or "(number puudub)",
                   f'{p["Title"]} ({summa} €) – {VALJA_JAETUD[p["Guid"]]}'])
    uued = [g for g in erc if g["raamprogramm"] != "FP7"]
    ws.append([])
    ws.append(["ERC grandid (CORDIS)"])
    ws.cell(ws.max_row, 1).font = KOKKU
    for rida in [
        ["Eesti osalusega kokku", len(erc)],
        ["  neist vastuvõtja Eestis", sum(g["eestis"] for g in erc)],
        ["H2020 ja Euroopa Horisont", len(uued)],
        ["  neist vastuvõtja Eestis", sum(g["eestis"] for g in uued)],
    ]:
        ws.append(rida)
    ws["B6"].number_format = EUR
    ws["A9"].font = KOKKU
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 90

    read = [(a, arv, summa) for a, (arv, summa) in sorted(algusaasta.items())]
    leht(wb, "Algusaasta", ["Algusaasta", "Alustatud PRG-sid", "Nende kogurahastus"],
         read, [None, None, EUR], [12, 18, 20], ["Kokku", n, raha])

    read = [(a, arv, summa) for a, (arv, summa) in sorted(eelarveaasta.items())]
    leht(wb, "Eelarveaasta", ["Eelarveaasta", "Aktiivseid PRG-sid", "Rahastus aastal"],
         read, [None, None, EUR], [12, 18, 20],
         ["Kokku", None, sum(r[2] for r in read)])

    read = osakaaluga(sorted(asutus.items(), key=lambda x: -x[1][1]), n, raha)
    leht(wb, "Asutused", ["Asutus", "PRG-sid", "Osakaal arvust", "Rahastus", "Osakaal rahast"],
         read, [None, ARV, PROTS, EUR, PROTS], [60, 10, 14, 18, 14], ["Kokku", n, 1, raha, 1])

    read = [(f"{k}. {POHIVALDKONNAD[k]}", *r) for k, *r in
            osakaaluga(sorted(suur.items()), n, raha)]
    leht(wb, "Frascati põhivaldkonnad", ["Põhivaldkond", "PRG-sid (osakaalu järgi)", "Osakaal arvust", "Rahastus", "Osakaal rahast"],
         read, [None, ARV, PROTS, EUR, PROTS], [34, 14, 14, 18, 14], ["Kokku", n, 1, raha, 1])

    jarjestus = sorted(ala, key=lambda k: tuple(int(x) for x in k.split(".")))
    read = [(k, ala_nimi[k], POHIVALDKONNAD[k.split(".")[0]], arv, arv / n, summa, summa / raha)
            for k in jarjestus for arv, summa in [ala[k]]]
    leht(wb, "Frascati alavaldkonnad", ["Kood", "Alavaldkond", "Põhivaldkond", "PRG-sid (osakaalu järgi)", "Osakaal arvust", "Rahastus", "Osakaal rahast"],
         read, [None, None, None, ARV, PROTS, EUR, PROTS], [8, 60, 26, 14, 14, 18, 14],
         ["Kokku", None, None, n, 1, raha, 1])

    read = [(p["FinancierProjectNr"] or "(number puudub)", p["Title"], p["ProjectStartDate"], p["ProjectEndDate"],
             "; ".join(asutused(p)), vastutavad(p), p["FinancingInPeriodsTotal"] or 0,
             "; ".join(f"{kood} {nimi} ({osa:.0%})" for kood, nimi, osa in valdkonnad(p)))
            for p in sorted(projektid, key=lambda p: (kp(p["ProjectStartDate"]), p["FinancierProjectNr"]))]
    leht(wb, "Projektid", ["Nr", "Pealkiri", "Algus", "Lõpp", "Asutus", "Vastutav täitja", "Rahastus", "Frascati valdkonnad"],
         read, [None, None, None, None, None, None, EUR, None], [14, 60, 11, 11, 30, 28, 16, 70])

    top = top_vastutavad(projektid)
    ek = {g: ek_projektid(g) for g, *_ in top}
    summad = [s for _, _, s, _ in top]
    read = [(1 + sum(round(t) > round(s) for t in summad), nimi, s, arv, len(ek[g]),
             sum(p["FinancingInPeriodsTotal"] or 0 for p in ek[g]),
             ", ".join(f'{e["akronuum"]} ({e["liik"]})' for e in erc if g in e["pi_guid"]))
            for g, nimi, s, arv in top]
    leht(wb, "Top vastutavad", ["Koht", "Vastutav täitja", "PRG rahastus", "PRG-sid", "EK projekte", "EK projektide rahastus", "ERC grandid"],
         read, [None, None, EUR, None, None, EUR, None], [7, 28, 18, 10, 12, 22, 50])

    read = [(nimi, p["FinancierProjectNr"], p["Acronym"], p["Title"] or p["TitleEng"], p["ProgrammeName"],
             p["ProjectStartDate"], p["ProjectEndDate"], "; ".join(asutused(p)), p["FinancingInPeriodsTotal"] or 0)
            for g, nimi, *_ in top for p in sorted(ek[g], key=lambda p: kp(p["ProjectStartDate"]))]
    leht(wb, "EK projektid", ["Vastutav täitja", "Nr", "Akronüüm", "Pealkiri", "Programm", "Algus", "Lõpp", "Asutus", "Rahastus"],
         read, [None, None, None, None, None, None, None, None, EUR], [24, 14, 14, 60, 34, 11, 11, 30, 16])

    read = [(g["nr"], g["akronuum"], g["pealkiri"], g["raamprogramm"], g["liik"], g["voor"], g["algus"],
             g["asutus"], g["roll"], "jah" if g["eestis"] else "ei", g["pi"], g["toetus"]) for g in erc]
    leht(wb, "ERC grandid", ["Grandinumber", "Akronüüm", "Pealkiri", "Raamprogramm", "ERC liik", "Taotlusvoor", "Algus",
                             "Eesti asutus", "Eesti roll", "Vastuvõtja Eestis", "Eesti PI (ETIS)", "EK toetus Eesti asutusele"],
         read, [None] * 11 + [EUR], [12, 20, 60, 16, 16, 16, 11, 34, 13, 11, 26, 18],
         ["Kokku"] + [None] * 10 + [sum(g["toetus"] for g in erc)])

    tee.parent.mkdir(exist_ok=True)
    wb.save(tee)


if __name__ == "__main__":
    projektid = lae_projektid()
    tee = Path(__file__).parent / "valjund" / f"PRG_ulevaade_{date.today():%Y-%m-%d}.xlsx"
    kirjuta_excel(projektid, tee)
    valja = sum(p["Guid"] in VALJA_JAETUD for p in projektid)
    print(f"{len(projektid) - valja} PRG projekti ({valja} kirjet välja jäetud) -> {tee}")
