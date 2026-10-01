#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Miesięczna aktualizacja danych GUS DBW:
- import gazu wg kraju pochodzenia
- eksport gazu wg kraju
- CN 27111100 LNG
- CN 27112100 gaz ziemny w stanie gazowym

Skrypt:
1. czyta istniejące CSV z repo,
2. sprawdza, których miesięcy/lat brakuje,
3. dodatkowo odświeża ostatnie opublikowane okresy,
4. pobiera tylko potrzebne okresy,
5. zastępuje stare dane dla odświeżanych okresów,
6. sprawdza duplikaty,
7. zapisuje cztery CSV.

Nie sumuje miesięcy do danych rocznych.
Dane roczne są pobierane jako osobna oficjalna seria GUS.
"""

import calendar
import csv
import gzip
import json
import re
import sys
import time

from datetime import date
from decimal import Decimal
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


# ============================================================
# API
# ============================================================

API = "https://api-dbw.stat.gov.pl/api/1.2.0/"


# ============================================================
# KLUCZ API
# ============================================================
#
# WKLEJ TU SWÓJ KLUCZ GUS.
# MA BYĆ BEZPOŚREDNIO W KODZIE.
#
# Przykład:
# KLUCZ_API = "abc123xyz="
#
# Nie usuwaj końcowego "=" jeżeli występuje.
# ============================================================

KLUCZ_API = "WKLEJ_TUTAJ_SWOJ_KLUCZ_API"


# ============================================================
# USTAWIENIA
# ============================================================

ROK_OD = 2018

KODY_CN = {
    "27111100",
    "27112100",
}

ROZMIAR_STRONY = 5000

# Ile ostatnich miesięcy odświeżać nawet jeżeli już są w CSV.
# Chroni przed późniejszymi korektami GUS.
ODSWIEZ_OSTATNIE_MIESIACE = 2

# Ile ostatnich lat odświeżać.
ODSWIEZ_OSTATNIE_LATA = 1


# ============================================================
# IMPORT / EKSPORT
# ============================================================

KONFIGURACJE = {

    "import": {

        "id_zmiennej": 221,

        "oczekiwana_nazwa":
            "Import towarów wg kraju pochodzenia",

        "id_przekroju": 1434,

        "miesieczne":
            "gus_gaz_miesieczne_import.csv",

        "roczne":
            "gus_gaz_roczne_import.csv",
    },

    "export": {

        # Jeżeli GUS zmieni identyfikator,
        # kontrola nazwy poniżej przerwie program
        # zamiast pobrać złe dane.
        "id_zmiennej": 220,

        "oczekiwany_fragment":
            "eksport",

        "id_przekroju": 1434,

        "miesieczne":
            "gus_gaz_export_miesieczne.csv",

        "roczne":
            "gus_gaz_export_roczne.csv",
    },
}


# ============================================================
# KOLUMNY CSV
# ============================================================

KOLUMNY = [
    "okres",
    "rok",
    "miesiac",
    "kraj",
    "kod_kraju",
    "typ_kraju",
    "kod_cn",
    "towar",
    "wartosc",
    "jednostka",
    "status",
    "flaga_gus",
    "tajnosc_gus",
    "brak_wartosci_gus",
    "wartosc_opisowa_gus",
    "precyzja_gus",
    "id_kraju_gus",
    "id_cn_gus",
    "id_okresu_gus",
]


KATEGORIE_SPECJALNE = {
    "QP",
    "QQ",
    "QR",
    "QS",
    "QU",
    "QV",
    "QW",
    "QX",
    "QY",
    "QZ",
    "_no",
}


# ============================================================
# API GUS
# ============================================================

class KlientGUS:

    def __init__(self, klucz):

        self.klucz = klucz.strip()

        if (
            not self.klucz
            or self.klucz
            == "WKLEJ_TUTAJ_SWOJ_KLUCZ_API"
        ):
            raise RuntimeError(
                "Wpisz klucz GUS w zmiennej KLUCZ_API."
            )

        self.ostatnie_zadanie = 0.0


    def pobierz(
        self,
        endpoint,
        parametry=None,
        dopuszczaj_404=False
    ):

        url = (
            API
            + endpoint
            + "?"
            + urlencode(
                {
                    "lang": "pl",
                    **(parametry or {})
                }
            )
        )

        for proba in range(6):

            time.sleep(
                max(
                    0,
                    2.05
                    - (
                        time.monotonic()
                        - self.ostatnie_zadanie
                    )
                )
            )

            self.ostatnie_zadanie = (
                time.monotonic()
            )

            naglowki = {
                "Accept":
                    "application/json",

                "Accept-Encoding":
                    "gzip",

                "X-ClientId":
                    self.klucz,

                "User-Agent":
                    "GUS-gas-monthly-updater/1.0",
            }

            try:

                req = Request(
                    url,
                    headers=naglowki
                )

                with urlopen(
                    req,
                    timeout=90
                ) as response:

                    tresc = response.read()

                    if (
                        response.headers.get(
                            "Content-Encoding"
                        )
                        == "gzip"
                    ):
                        tresc = gzip.decompress(
                            tresc
                        )

                    return json.loads(
                        tresc.decode(
                            "utf-8-sig"
                        )
                    )


            except HTTPError as e:

                if (
                    e.code == 404
                    and dopuszczaj_404
                ):
                    return None

                if e.code in (401, 403):

                    raise RuntimeError(
                        f"HTTP {e.code}: "
                        "GUS odrzucił KLUCZ_API."
                    ) from None

                if e.code not in (
                    408,
                    429,
                    500,
                    502,
                    503,
                    504
                ):
                    raise RuntimeError(
                        f"HTTP {e.code}: {endpoint}"
                    ) from None

                powod = f"HTTP {e.code}"

            except (
                URLError,
                TimeoutError,
                OSError,
                ValueError
            ) as e:

                powod = type(e).__name__


            if proba == 5:

                raise RuntimeError(
                    f"Nie udało się pobrać "
                    f"{endpoint}: {powod}"
                )


            opoznienie = (
                120
                if powod == "HTTP 429"
                else min(
                    60,
                    5 * 2 ** proba
                )
            )

            print(
                f"  {powod}; "
                f"ponawiam za "
                f"{opoznienie} s.",
                flush=True
            )

            time.sleep(opoznienie)


    def slownik(
        self,
        nazwa,
        klucz
    ):

        wynik = {}

        strona = 1

        while True:

            dane = self.pobierz(
                "dictionaries/" + nazwa,
                {
                    "page":
                        strona,

                    "page-size":
                        ROZMIAR_STRONY,
                }
            )

            if (
                not isinstance(dane, dict)
                or
                not isinstance(
                    dane.get("data"),
                    list
                )
            ):
                raise RuntimeError(
                    f"Niepoprawny słownik: {nazwa}"
                )

            for x in dane["data"]:
                wynik[x[klucz]] = x

            if strona >= dane["page-count"]:
                return wynik

            strona += 1


# ============================================================
# CSV
# ============================================================

def wczytaj_csv(sciezka):

    if not sciezka.exists():
        return []

    with sciezka.open(
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        return list(
            csv.DictReader(
                f,
                delimiter=";"
            )
        )


def zapisz_csv(
    sciezka,
    wiersze
):

    tmp = sciezka.with_suffix(
        ".tmp"
    )

    with tmp.open(
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=KOLUMNY,
            delimiter=";"
        )

        writer.writeheader()
        writer.writerows(wiersze)

    tmp.replace(sciezka)


# ============================================================
# METADANE
# ============================================================

def pobierz_metadane(
    klient,
    konfiguracja
):

    id_zmiennej = (
        konfiguracja[
            "id_zmiennej"
        ]
    )

    id_przekroju = (
        konfiguracja[
            "id_przekroju"
        ]
    )


    meta = klient.pobierz(
        "variable/variable-meta",
        {
            "id-zmiennej":
                id_zmiennej
        }
    )


    nazwa = (
        meta.get(
            "nazwa",
            ""
        )
    )


    if "oczekiwana_nazwa" in konfiguracja:

        if (
            nazwa
            != konfiguracja[
                "oczekiwana_nazwa"
            ]
        ):
            raise RuntimeError(
                f"Zmienna {id_zmiennej}: "
                f"oczekiwano "
                f"{konfiguracja['oczekiwana_nazwa']!r}, "
                f"GUS zwrócił {nazwa!r}."
            )


    if "oczekiwany_fragment" in konfiguracja:

        if (
            konfiguracja[
                "oczekiwany_fragment"
            ].casefold()
            not in nazwa.casefold()
        ):
            raise RuntimeError(
                f"Zmienna {id_zmiennej} "
                f"nie wygląda na eksport. "
                f"GUS zwrócił {nazwa!r}."
            )


    print(
        f"  Zmienna GUS: {nazwa}"
    )


    pozycje = klient.pobierz(
        "variable/variable-section-position",
        {
            "id-przekroj":
                id_przekroju
        }
    )


    kraje = {

        x["id-pozycja"]: x

        for x in pozycje

        if (
            x.get(
                "nazwa-wymiar"
            )
            == "Kraje towary"

            and x.get(
                "symbol"
            )
            not in (
                "00",
                "EU"
            )

            and x.get(
                "nazwa-pozycja",
                ""
            )
            .strip()
            .casefold()
            != "ogółem"
        )
    }


    towary = {

        x["id-pozycja"]:
            dict(x)

        for x in pozycje

        if (
            "CN"
            in x.get(
                "nazwa-wymiar",
                ""
            )

            and str(
                x.get(
                    "symbol",
                    ""
                )
            )
            in KODY_CN
        )
    }


    znalezione = {

        str(
            x["symbol"]
        )

        for x in towary.values()
    }


    if znalezione != KODY_CN:

        raise RuntimeError(
            f"Nie znaleziono obu CN. "
            f"Znaleziono: {znalezione}"
        )


    if not kraje:

        raise RuntimeError(
            "Brak wymiaru krajów."
        )


    for x in towary.values():

        m = re.search(
            r"\[([^\[\]]+)\]\s*$",
            x.get(
                "nazwa-pozycja",
                ""
            )
        )

        x["jednostka"] = (
            m.group(1)
            if m
            else ""
        )


    okresy = klient.slownik(
        "periods-dictionary",
        "id-okres"
    )


    miesiace = {}

    for ident, x in okresy.items():

        m = re.fullmatch(
            r"M(\d{2})",
            x.get(
                "symbol",
                ""
            )
        )

        if (
            x.get(
                "id-czestotliwosc"
            )
            == 3

            and x.get(
                "id-typ"
            )
            == 1

            and m
        ):
            miesiace[
                int(
                    m.group(1)
                )
            ] = ident


    roczne = [

        ident

        for ident, x
        in okresy.items()

        if (
            x.get("opis")
            ==
            "rok - dane roczne - rok"
        )
    ]


    if len(miesiace) != 12:

        raise RuntimeError(
            "Brak wszystkich miesięcy "
            "w słowniku GUS."
        )


    if len(roczne) != 1:

        raise RuntimeError(
            "Nie rozpoznano okresu rocznego."
        )


    serie = {

        x[
            "id-czestotliwosc"
        ]: x

        for x
        in meta.get(
            "przekroje",
            []
        )

        if (
            x.get(
                "id-przekroj"
            )
            == id_przekroju
        )
    }


    if (
        1 not in serie
        or 3 not in serie
    ):
        raise RuntimeError(
            "Brak miesięcznej "
            "lub rocznej serii."
        )


    return {
        "id_zmiennej":
            id_zmiennej,

        "id_przekroju":
            id_przekroju,

        "kraje":
            kraje,

        "towary":
            towary,

        "miesiace":
            miesiace,

        "okres_roczny":
            roczne[0],

        "serie":
            serie,

        "flagi":
            klient.slownik(
                "flag-dictionary",
                "id-flaga"
            ),

        "braki":
            klient.slownik(
                "no-value-dictionary",
                "id-brak-wartosci"
            ),

        "tajnosc":
            klient.slownik(
                "confidentionality-dictionary",
                "id-tajnosci"
            ),
    }


# ============================================================
# WYMIAR
# ============================================================

def pozycja_wymiaru(
    rekord,
    wymiar
):

    for i in range(
        1,
        16
    ):

        if (
            rekord.get(
                f"id-wymiar-{i}"
            )
            == wymiar
        ):
            return rekord[
                f"id-pozycja-{i}"
            ]

    raise RuntimeError(
        f"Brakuje wymiaru {wymiar}."
    )


# ============================================================
# POBIERANIE JEDNEGO OKRESU
# ============================================================

def pobierz_okres(
    klient,
    meta,
    rok,
    miesiac
):

    if miesiac:

        id_okresu = (
            meta[
                "miesiace"
            ][
                miesiac
            ]
        )

    else:

        id_okresu = (
            meta[
                "okres_roczny"
            ]
        )


    etykieta = (
        f"{rok} M{miesiac}"
        if miesiac
        else str(rok)
    )


    wymiar_cn = next(
        iter(
            meta[
                "towary"
            ].values()
        )
    )[
        "id-wymiar"
    ]


    wymiar_kraju = next(
        iter(
            meta[
                "kraje"
            ].values()
        )
    )[
        "id-wymiar"
    ]


    wymiar_polski = 2


    wynik = []

    strona = 0


    while True:

        dane = klient.pobierz(
            "variable/variable-data-section",
            {
                "id-zmienna":
                    meta[
                        "id_zmiennej"
                    ],

                "id-przekroj":
                    meta[
                        "id_przekroju"
                    ],

                "id-rok":
                    rok,

                "id-okres":
                    id_okresu,

                "ile-na-stronie":
                    ROZMIAR_STRONY,

                "numer-strony":
                    strona,
            },
            dopuszczaj_404=(
                strona == 0
            )
        )


        if dane is None:

            print(
                f"    {etykieta}: "
                "jeszcze brak w API."
            )

            return None


        rekordy = dane.get(
            "data",
            []
        )


        if (
            not rekordy
            and strona == 0
        ):
            return None


        ostatnia = (
            dane[
                "page-count"
            ]
        )


        for rekord in rekordy:

            if (
                rekord.get(
                    "id-zmienna"
                )
                != meta[
                    "id_zmiennej"
                ]
            ):
                raise RuntimeError(
                    "API zwróciło inną zmienną."
                )


            cn = pozycja_wymiaru(
                rekord,
                wymiar_cn
            )


            if (
                cn
                not in
                meta["towary"]
            ):
                continue


            kraj = pozycja_wymiaru(
                rekord,
                wymiar_kraju
            )


            if (
                kraj
                not in
                meta["kraje"]
            ):
                continue


            if (
                pozycja_wymiaru(
                    rekord,
                    wymiar_polski
                )
                != 33617
            ):
                raise RuntimeError(
                    "Dane nie dotyczą Polski."
                )


            wynik.append(
                {
                    **rekord,
                    "kraj_id":
                        kraj,
                    "cn_id":
                        cn,
                }
            )


        print(
            f"    {etykieta}: "
            f"strona "
            f"{strona + 1}/"
            f"{ostatnia + 1}",
            flush=True
        )


        if strona >= ostatnia:
            break


        strona += 1


    return {
        "rok":
            rok,

        "miesiac":
            miesiac,

        "okres":
            etykieta,

        "id_okres":
            id_okresu,

        "dane":
            wynik,
    }


# ============================================================
# OPIS SŁOWNIKA
# ============================================================

def opis(
    slownik,
    ident
):

    if ident is None:
        return ""

    return (
        slownik
        .get(
            ident,
            {}
        )
        .get(
            "nazwa",
            ""
        )
    )


# ============================================================
# AKTYWNOŚĆ POZYCJI
# ============================================================

def aktywna(
    pozycja,
    rok,
    miesiac
):

    poczatek = date(
        rok,
        miesiac or 1,
        1
    ).isoformat()


    koniec = date(
        rok,
        miesiac or 12,

        calendar.monthrange(
            rok,
            miesiac or 12
        )[1]
    ).isoformat()


    return (

        (
            pozycja.get(
                "data-poczatku"
            )
            or "0001-01-01"
        )[:10]
        <= koniec

        and

        (
            pozycja.get(
                "data-konca"
            )
            or "9999-12-31"
        )[:10]
        >= poczatek
    )


# ============================================================
# GENEROWANIE WIERSZY
# ============================================================

def generuj_wiersze(
    meta,
    stan
):

    rekordy = {}


    for rekord in stan["dane"]:

        klucz = (
            rekord[
                "kraj_id"
            ],
            rekord[
                "cn_id"
            ]
        )

        if klucz in rekordy:

            raise RuntimeError(
                f"Duplikat w API "
                f"dla {stan['okres']}: "
                f"{klucz}"
            )

        rekordy[
            klucz
        ] = rekord


    kraje = sorted(
        meta[
            "kraje"
        ].items(),

        key=lambda x:
            x[1][
                "nazwa-pozycja"
            ]
    )


    towary = sorted(
        meta[
            "towary"
        ].items(),

        key=lambda x:
            x[1][
                "symbol"
            ]
    )


    wynik = []


    for kraj_id, kraj in kraje:

        if not aktywna(
            kraj,
            stan["rok"],
            stan["miesiac"]
        ):
            continue


        for cn_id, cn in towary:

            if not aktywna(
                cn,
                stan["rok"],
                stan["miesiac"]
            ):
                continue


            rekord = rekordy.get(
                (
                    kraj_id,
                    cn_id
                ),
                {}
            )


            wartosc = rekord.get(
                "wartosc"
            )


            if wartosc is not None:

                status = (
                    "wartosc_opublikowana"
                )


            elif not rekord:

                status = (
                    "brak_rekordu_w_api"
                )


            elif (
                meta[
                    "tajnosc"
                ]
                .get(
                    rekord.get(
                        "id-tajnosci"
                    ),
                    {}
                )
                .get(
                    "oznaczenie"
                )
                == "(:)"
            ):

                status = (
                    "tajemnica_statystyczna"
                )

                wartosc = None


            elif (
                rekord.get(
                    "id-brak-wartosci"
                )
                == 42
            ):

                status = (
                    "zjawisko_nie_wystapilo"
                )


            else:

                status = (
                    "brak_wartosci_w_api"
                )


            wynik.append({

                "okres":
                    stan[
                        "okres"
                    ],

                "rok":
                    str(
                        stan[
                            "rok"
                        ]
                    ),

                "miesiac":
                    (
                        str(
                            stan[
                                "miesiac"
                            ]
                        )
                        if stan[
                            "miesiac"
                        ]
                        else ""
                    ),

                "kraj":
                    kraj[
                        "nazwa-pozycja"
                    ],

                "kod_kraju":
                    kraj.get(
                        "symbol",
                        ""
                    ),

                "typ_kraju":
                    (
                        "kategoria_specjalna"

                        if kraj.get(
                            "symbol"
                        )
                        in
                        KATEGORIE_SPECJALNE

                        else
                        "kraj_lub_terytorium"
                    ),

                "kod_cn":
                    cn[
                        "symbol"
                    ],

                "towar":
                    cn[
                        "nazwa-pozycja"
                    ],

                "wartosc":
                    (
                        ""

                        if wartosc is None

                        else

                        format(
                            Decimal(
                                str(
                                    wartosc
                                )
                            ),
                            "f"
                        ).replace(
                            ".",
                            ","
                        )
                    ),

                "jednostka":
                    cn.get(
                        "jednostka",
                        ""
                    ),

                "status":
                    status,

                "flaga_gus":
                    opis(
                        meta[
                            "flagi"
                        ],
                        rekord.get(
                            "id-flaga"
                        )
                    ),

                "tajnosc_gus":
                    opis(
                        meta[
                            "tajnosc"
                        ],
                        rekord.get(
                            "id-tajnosci"
                        )
                    ),

                "brak_wartosci_gus":
                    opis(
                        meta[
                            "braki"
                        ],
                        rekord.get(
                            "id-brak-wartosci"
                        )
                    ),

                "wartosc_opisowa_gus":
                    rekord.get(
                        "wartosc-opisowa",
                        ""
                    ),

                "precyzja_gus":
                    rekord.get(
                        "precyzja",
                        ""
                    ),

                "id_kraju_gus":
                    str(
                        kraj_id
                    ),

                "id_cn_gus":
                    str(
                        cn_id
                    ),

                "id_okresu_gus":
                    str(
                        stan[
                            "id_okres"
                        ]
                    ),
            })


    return wynik


# ============================================================
# DOSTĘPNE LATA
# ============================================================

def zakres_lat(
    seria
):

    lata = [
        int(x)
        for x
        in re.findall(
            r"\d{4}",
            str(
                seria.get(
                    "szereg-czasowy",
                    ""
                )
            )
        )
    ]

    if not lata:

        raise RuntimeError(
            "Nie rozpoznano zakresu lat GUS."
        )

    return (
        min(lata),
        max(lata)
    )


# ============================================================
# OCZEKIWANE MIESIĄCE
# ============================================================

def oczekiwane_miesiace(
    meta
):

    dzis = date.today()

    rok_min, rok_max = (
        zakres_lat(
            meta[
                "serie"
            ][3]
        )
    )


    start = max(
        ROK_OD,
        rok_min
    )


    wynik = []


    for rok in range(
        start,
        min(
            rok_max,
            dzis.year
        )
        + 1
    ):

        for miesiac in range(
            1,
            13
        ):

            if (
                rok == dzis.year
                and
                miesiac >= dzis.month
            ):
                continue

            wynik.append(
                (
                    rok,
                    miesiac
                )
            )


    return wynik


# ============================================================
# OCZEKIWANE LATA
# ============================================================

def oczekiwane_lata(
    meta
):

    dzis = date.today()

    rok_min, rok_max = (
        zakres_lat(
            meta[
                "serie"
            ][1]
        )
    )


    pierwszy = max(
        ROK_OD,
        rok_min
    )


    ostatni = min(
        rok_max,
        dzis.year - 1
    )


    if ostatni < pierwszy:
        return []


    return list(
        range(
            pierwszy,
            ostatni + 1
        )
    )


# ============================================================
# OKRESY JUŻ W CSV
# ============================================================

def okresy_miesieczne_w_csv(
    wiersze
):

    wynik = set()


    for w in wiersze:

        try:

            rok = int(
                w["rok"]
            )

            miesiac = int(
                w["miesiac"]
            )

        except (
            ValueError,
            TypeError,
            KeyError
        ):
            continue


        wynik.add(
            (
                rok,
                miesiac
            )
        )


    return wynik


def lata_w_csv(
    wiersze
):

    wynik = set()


    for w in wiersze:

        try:
            wynik.add(
                int(
                    w["rok"]
                )
            )

        except (
            ValueError,
            TypeError,
            KeyError
        ):
            continue


    return wynik


# ============================================================
# WYMIANA OKRESU
# ============================================================

def wymien_miesiac(
    stare,
    rok,
    miesiac,
    nowe
):

    wynik = [

        w

        for w
        in stare

        if not (
            str(
                w.get(
                    "rok",
                    ""
                )
            )
            == str(rok)

            and

            str(
                w.get(
                    "miesiac",
                    ""
                )
            )
            == str(miesiac)
        )
    ]


    wynik.extend(
        nowe
    )

    return wynik


def wymien_rok(
    stare,
    rok,
    nowe
):

    wynik = [

        w

        for w
        in stare

        if (
            str(
                w.get(
                    "rok",
                    ""
                )
            )
            != str(rok)
        )
    ]


    wynik.extend(
        nowe
    )

    return wynik


# ============================================================
# DUPLIKATY
# ============================================================

def sprawdz_duplikaty(
    wiersze,
    miesieczne
):

    widziane = set()

    duplikaty = []


    for w in wiersze:

        if miesieczne:

            klucz = (
                w.get(
                    "rok",
                    ""
                ),

                w.get(
                    "miesiac",
                    ""
                ),

                w.get(
                    "id_kraju_gus",
                    ""
                ),

                w.get(
                    "id_cn_gus",
                    ""
                ),
            )

        else:

            klucz = (
                w.get(
                    "rok",
                    ""
                ),

                w.get(
                    "id_kraju_gus",
                    ""
                ),

                w.get(
                    "id_cn_gus",
                    ""
                ),
            )


        if klucz in widziane:

            duplikaty.append(
                klucz
            )

        else:

            widziane.add(
                klucz
            )


    if duplikaty:

        raise RuntimeError(
            f"Znaleziono "
            f"{len(duplikaty)} "
            f"duplikatów. "
            f"Przykład: "
            f"{duplikaty[0]}"
        )


# ============================================================
# SORTOWANIE
# ============================================================

def sortuj(
    wiersze
):

    def klucz(w):

        try:
            rok = int(
                w.get(
                    "rok",
                    0
                )
            )
        except ValueError:
            rok = 0


        try:
            miesiac = int(
                w.get(
                    "miesiac",
                    0
                )
                or 0
            )
        except ValueError:
            miesiac = 0


        return (
            rok,
            miesiac,
            w.get(
                "kraj",
                ""
            ),
            w.get(
                "kod_cn",
                ""
            ),
        )


    return sorted(
        wiersze,
        key=klucz
    )


# ============================================================
# AKTUALIZACJA JEDNEJ SERII
# ============================================================

def aktualizuj(
    klient,
    nazwa,
    konfiguracja
):

    print(
        "\n"
        "========================================"
    )

    print(
        f"AKTUALIZACJA: {nazwa.upper()}"
    )

    print(
        "========================================"
    )


    meta = pobierz_metadane(
        klient,
        konfiguracja
    )


    miesieczny_plik = Path(
        konfiguracja[
            "miesieczne"
        ]
    )


    roczny_plik = Path(
        konfiguracja[
            "roczne"
        ]
    )


    miesieczne = wczytaj_csv(
        miesieczny_plik
    )


    roczne = wczytaj_csv(
        roczny_plik
    )


    print(
        f"  Istniejące miesięczne: "
        f"{len(miesieczne)} wierszy"
    )

    print(
        f"  Istniejące roczne: "
        f"{len(roczne)} wierszy"
    )


    # ========================================================
    # MIESIĄCE
    # ========================================================

    wszystkie_miesiace = (
        oczekiwane_miesiace(
            meta
        )
    )


    obecne_miesiace = (
        okresy_miesieczne_w_csv(
            miesieczne
        )
    )


    brakujace_miesiace = [

        okres

        for okres
        in wszystkie_miesiace

        if okres
        not in obecne_miesiace
    ]


    ostatnie = (
        wszystkie_miesiace[
            -ODSWIEZ_OSTATNIE_MIESIACE:
        ]
    )


    do_pobrania_miesiace = sorted(
        set(
            brakujace_miesiace
            + ostatnie
        )
    )


    print(
        f"  Miesięcy do sprawdzenia: "
        f"{len(do_pobrania_miesiace)}"
    )


    for rok, miesiac in do_pobrania_miesiace:

        stan = pobierz_okres(
            klient,
            meta,
            rok,
            miesiac
        )


        if stan is None:

            continue


        nowe = generuj_wiersze(
            meta,
            stan
        )


        miesieczne = wymien_miesiac(
            miesieczne,
            rok,
            miesiac,
            nowe
        )


    # ========================================================
    # LATA
    # ========================================================

    wszystkie_lata = (
        oczekiwane_lata(
            meta
        )
    )


    obecne_lata = (
        lata_w_csv(
            roczne
        )
    )


    brakujace_lata = [

        rok

        for rok
        in wszystkie_lata

        if rok
        not in obecne_lata
    ]


    ostatnie_lata = (
        wszystkie_lata[
            -ODSWIEZ_OSTATNIE_LATA:
        ]
    )


    do_pobrania_lata = sorted(
        set(
            brakujace_lata
            + ostatnie_lata
        )
    )


    print(
        f"  Lat do sprawdzenia: "
        f"{len(do_pobrania_lata)}"
    )


    for rok in do_pobrania_lata:

        stan = pobierz_okres(
            klient,
            meta,
            rok,
            0
        )


        if stan is None:

            continue


        nowe = generuj_wiersze(
            meta,
            stan
        )


        roczne = wymien_rok(
            roczne,
            rok,
            nowe
        )


    # ========================================================
    # DUPLIKATY
    # ========================================================

    miesieczne = sortuj(
        miesieczne
    )

    roczne = sortuj(
        roczne
    )


    sprawdz_duplikaty(
        miesieczne,
        True
    )


    sprawdz_duplikaty(
        roczne,
        False
    )


    print(
        "  Duplikaty miesięczne: 0"
    )

    print(
        "  Duplikaty roczne: 0"
    )


    # ========================================================
    # ZAPIS
    # ========================================================

    zapisz_csv(
        miesieczny_plik,
        miesieczne
    )


    zapisz_csv(
        roczny_plik,
        roczne
    )


    print(
        f"  Zapisano "
        f"{miesieczny_plik}"
    )

    print(
        f"  Zapisano "
        f"{roczny_plik}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    klient = KlientGUS(
        KLUCZ_API
    )


    print(
        "GUS GAS MONTHLY UPDATE"
    )

    print(
        f"Data: {date.today()}"
    )


    aktualizuj(
        klient,
        "import",
        KONFIGURACJE[
            "import"
        ]
    )


    aktualizuj(
        klient,
        "export",
        KONFIGURACJE[
            "export"
        ]
    )


    print(
        "\n"
        "========================================"
    )

    print(
        "AKTUALIZACJA ZAKOŃCZONA"
    )

    print(
        "========================================"
    )


    print(
        "Pliki:"
    )

    print(
        "  gus_gaz_miesieczne_import.csv"
    )

    print(
        "  gus_gaz_roczne_import.csv"
    )

    print(
        "  gus_gaz_export_miesieczne.csv"
    )

    print(
        "  gus_gaz_export_roczne.csv"
    )


if __name__ == "__main__":

    try:

        main()

    except Exception as exc:

        print(
            f"\nBŁĄD: {exc}",
            file=sys.stderr
        )

        raise
