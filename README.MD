# etrade-parser

## Mikä?

Etrade-tietojen unperseraattori veroilmoitusta varten.

Veroilmoitukseen tarvittavien tietojen kaivaminen Etradesta on sen verran työlästä, että olen joka vuosi automatisoinut prosessia vähän enemmän ja nyt ajattelin jakaa skriptin muidenkin käyttöön. Tietäjät tietää, mihin tätä käytetään.

Scripti luo html-raportin, missä on eriteltynä kaikki myyntitapahtumat euromääräisinä summina, ja laskee yhteen luovutushinnat-, voitot ja tappiot. Raportin voi liittää veroilmoitukseen ja toivoa, että verottaja on tyytyväinen (käytännön kokemusta ei vielä ole).

## Asennus

1. Kloonaa repo tai lataa pelkkä skripti.
2. Asenna riippuvuudet, vaikka näin:
   ```
   pip install openpyxl pandas 
   ```

## Käyttö

1. Lataa Etradesta vuoden yhteenveto (.xlsx):
    - At work -> My Account -> Gains & Losses.
    - Valitse haluttu vuosi tai "custom interval" ja paina 'Apply'.
    - Paina 'Download' ja valitse 'Download Expanded'.
2. Lataa valuutanvaihtokurssit Suomen pankin sivuilta:

    https://www.suomenpankki.fi/fi/tilastot/taulukot-ja-kuviot/valuuttakurssit/taulukot/valuuttakurssit_taulukot_fi/valuuttakurssit_short_fi/

    Valitse seuraavat:
    - 'Kurssit': 'Yhdysvaltain dollari'. Varmista, että muita valuuttoja ei ole valittuna.
    - 'Alkupäivä': Valitse päivämäärä ennen ensimmäistä ostotapahtumaa.
    - 'Loppupäivä': Valitse joku päivämäärä viimeisimmän myynnin jälkeen.
    - 'Uusimmat': 'ylimpänä'
    - 'Aggregointi': 'Keskiarvo'
    - 'Näytä tiedot': 'päivätasolla'

    Paina "View report". Export -> CSV.

Suorita ohjelma seuraavalla komennolla (esimerkki):

```bash
python etrade-parser.py --input "G&L_Expanded_2025.xlsx" --valuuttakurssit "valuuttakurssit_short_fi.xlsx"
```

Jos myytyjen osakkeiden joukossa on ESPP-osakkeita, tarvitset lisäksi ESPP:n ostoihin käytetyn valuuttakurssin.
Tämä valuuttakurssi löytyy `ESPP purchase confirmation` -dokumenteista.

Dokumentit löytyy Etradesta:
    At work -> My Account -> Stock Plan Confirmation.

Scripti kysyy näitä tietoja tarvittaessa. Scripti tallentaa tiedot excel-tiedoston kopioon, jos jostain syystä scripti pitää ajaa uudelleen (harvoin mikään onnistuu ensimmäisellä kerralla). Huomaa, että tallennetun tiedoston nimi on *<alkuperäinen-nimi>_modified.xlsx*. 

Lisää ohjeita skriptin helpissä.

# Tiedot veroilmoitukseen

Huom: ota huomioon mahdolliset muut arvopaperiluovutukset kokonaissummia ilmoitettaessa.

Skripti luo html-raportin, missä *pitäisi* olla kaikki tarvittavat tiedot veroilmoitusta varten. Koska verottaja tykkää PDF-tiedostoista, html kannattaa muuttaa PDF-muotoon vaikka selaimen avulla.

Ilmoita veroilmoituksessa kohdassa *Muut tulot -> Pääomatulot -> Arvopaperien myyntivoitot*.

- Valitse *Ilmoittamistapa*: *Ilmoitan arvopaperien tiedot liitetiedostona*
- Lisää liitetiedostoksi PDF-tiedosto. (*Tyyppi*: Arvopaperikauppojen liite, *Kuvaus*: Arvopaperien luovutukset eriteltynä.)
- Ilmoita raportissa olevat Luovutushinnat, -voitot ja -tappiot *Luovutusten arvo*-kohtaan.
- Skripti ei ota huomioon valuutan vaihdosta aiheutuvia voittoja tai tappioita.

## Muuta

- Käyttö omalla vastuulla.
- 🐮
