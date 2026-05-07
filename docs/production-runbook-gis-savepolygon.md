# Production runbook: GIS saveParcelPolygon security/recovery

Status: **operativni runbook za incidente “parcela ima pogrešan polygon”, “neko je promenio geometriju”, “polygon save nije uspeo”, “Lat/Lng nije usklađen sa poligonom”, “public write endpoint je security rizik”.**

Aplikacija: **OtkupApp / AgriX PWA**
Domen: **Parcel GIS → parcel-draw.html → GAS getParcelGeo/saveParcelPolygon → tabela parcela/geometrije**
Glavni kod: `parcel-draw.html`, `gas/Code.gs`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Parcela je na pogrešnom mestu.”
* “Polygon parcele je obrisan.”
* “Koordinate Lat/Lng ne odgovaraju nacrtanom poligonu.”
* “Sačuvao sam polygon, ali se ne vidi u PWA.”
* “Mapa prikazuje staru parcelu.”
* “Neko je promenio polygon, a ne znamo ko.”
* “Kooperant vidi tuđu geometriju.”
* “saveParcelPolygon radi bez jasne autorizacije.”
* “Kliknuo sam Obriši polygon i mislio da je obrisano na serveru.”
* “Google Maps link vodi na pogrešnu lokaciju.”

Prvo pravilo:

> Geometrija parcele nije samo UI podatak. Ona utiče na meteo, geofencing, tretmane, karencu, berbu, izveštaje i poverenje u parcele. Ne menjaj polygon bez `ParcelaID`, stare vrednosti, nove vrednosti i odluke ko sme da menja.

Minimalni podaci koje operator mora da prikupi:

```text
ParcelaID:
KooperantID:
KatBroj:
KatOpstina:
Kultura:
Stari Lat:
Stari Lng:
Stari PolygonGeoJSON:
Novi Lat:
Novi Lng:
Novi PolygonGeoJSON:
Ko je menjao:
Vreme izmene:
Da li je korišćen parcel-draw.html:
Da li je postojao token u URL/localStorage:
Da li je incident security ili normalna korekcija:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: tabela parcela / GAS data source

Za svaku parcelu proveri:

```text
ParcelaID
KooperantID
KatBroj
KatOpstina
Kultura
PovrsinaHa
Lat
Lng
PolygonGeoJSON
GGAPStatus
```

`PolygonGeoJSON` je canonical geometrija. `Lat` i `Lng` su centroid ili marker za mapu.

### 2.2. Drugo mesto: `parcel-draw.html`

`parcel-draw.html` očekuje URL:

```text
parcel-draw.html?parcelaId=<ParcelaID>&token=<token>
```

Ako `token` nije u query string-u, stranica pokušava:

```text
localStorage.getItem('authToken')
```

Relevantne akcije:

* `getParcelGeo` učitava parcelu;
* `saveParcelPolygon` čuva polygon, centroid `lat/lng` i `polygonGeoJSON`.

### 2.3. Treće mesto: GAS endpoint

Proveriti `Code.gs`:

* da li je `getParcelGeo` public read;
* da li je `saveParcelPolygon` public write exception ili prolazi kroz token validation;
* da li `saveParcelPolygon` proverava role/entity ownership;
* da li je action logovan u ErrorLog ili poseban audit log.

### 2.4. Četvrto mesto: backup / Drive history

Ako je geometrija pogrešno promenjena, proveri:

```text
Google Sheet version history
Drive revision history
backup export ako postoji
ErrorLog oko vremena izmene
browser/user koji je radio izmenu ako je dostupno
```

Ako nema audit log-a, tretirati kao production gap.

---

## 3. Koji ID pratiš

Primarni ID:

```text
ParcelaID
```

Sekundarni ID-jevi:

```text
KooperantID
KatBroj
KatOpstina
PolygonGeoJSON hash
Lat/Lng centroid
Action request time
Token/session user ako postoji
```

Incident ticket minimum:

```text
ParcelaID:
KooperantID:
KatBroj:
Old Polygon hash:
New Polygon hash:
Old Lat/Lng:
New Lat/Lng:
Source: parcel-draw / manual sheet edit / unknown
Authorized: Da/Ne/Nepoznato
Security incident: Da/Ne/Nepoznato
Decision:
```

---

## 4. Normalan GIS tok

### 4.1. Učitavanje parcele

1. Operator otvara `parcel-draw.html?parcelaId=...`.
2. Stranica poziva:

```text
action = getParcelGeo
parcelaId = <ParcelaID>
```

3. GAS vraća parcelu.
4. Ako parcela ima `Lat`/`Lng`, mapa se centririra.
5. Ako parcela ima `PolygonGeoJSON`, polygon se renderuje na mapi.
6. Ako nema polygon, prikazuje se marker na `Lat`/`Lng` ako postoje.

### 4.2. Crtanje i čuvanje poligona

1. Operator iscrta polygon kroz Leaflet Draw.
2. UI izračuna centroid poligona.
3. UI popuni `Lat`/`Lng` iz centroida.
4. UI prikazuje JSON preview.
5. Klik na “Sačuvaj polygon” šalje:

```text
action = saveParcelPolygon
token = authToken
parcelaId = <ParcelaID>
polygonGeoJSON = JSON.stringify(geometry)
lat = centroid.lat
lng = centroid.lng
```

6. Ako backend vrati success, UI prikazuje “polygon sačuvan”.

### 4.3. Obriši polygon u UI-u

Dugme “Obriši polygon” u `parcel-draw.html` briše polygon **sa ekrana**.

Važno:

> Samo brisanje sa ekrana nije isto što i server-side delete. Ako korisnik klikne “Obriši polygon” i ne sačuva odgovarajuću promenu, server geometrija može ostati ista.

---

## 5. Security klasifikacija

### 5.1. Normalna korekcija

Signal:

```text
poznat operator
poznata parcela
postoji poslovni razlog
stara i nova geometrija su dokumentovane
operator ima pravo da menja parcelu
```

Akcija:

* dozvoljena izmena uz audit/ticket.

### 5.2. Neovlašćena ili nepoznata izmena

Signal:

```text
ne zna se ko je promenio
promena nije tražena
polygon pomeren daleko od originala
KooperantID ne odgovara korisniku
saveParcelPolygon nema auth check
nema audit log-a
```

Akcija:

* tretirati kao security/data integrity incident;
* zaustaviti dalje izmene parcele;
* vratiti validnu geometriju iz backup/history;
* proveriti endpoint autorizaciju.

### 5.3. Public write rizik

Ako `saveParcelPolygon` može da se pozove bez validnog tokena ili bez role/entity provere, to je production security gap.

Minimalno pravilo za produkciju:

```text
saveParcelPolygon mora tražiti validan token
Management sme da menja sve
Kooperant sme da menja samo svoje parcele ako je to poslovno dozvoljeno
svaka izmena mora imati audit log
```

---

## 6. Standardni incident flow

### Korak 1: Identifikuj parcelu

Zapiši:

```text
ParcelaID:
KooperantID:
KatBroj:
KatOpstina:
Kultura:
```

### Korak 2: Sačuvaj trenutno stanje

Pre bilo kakve izmene exportuj trenutno stanje:

```text
Lat:
Lng:
PolygonGeoJSON:
Polygon hash:
Timestamp:
```

Ako je moguće, kopiraj red parcele u incident ticket.

### Korak 3: Odredi da li je problem UI, backend ili data

| Signal                                        | Kategorija              | Sledeći korak                      |
| --------------------------------------------- | ----------------------- | ---------------------------------- |
| UI prikazuje staru geometriju, sheet ima novu | cache/UI problem        | reload/cache clear                 |
| UI prikazuje novu, sheet ima staru            | save nije uspeo         | proveri response/ErrorLog          |
| sheet ima pogrešnu geometriju                 | data incident           | restore/korekcija                  |
| Lat/Lng ne odgovara polygonu                  | centroid mismatch       | recalculacija centroida            |
| ne zna se ko je menjao                        | audit/security incident | backup/history + endpoint auth     |
| save radi bez tokena                          | security gap            | blokirati ili hardenovati endpoint |

### Korak 4: Proveri autorizaciju

Proveri:

```text
Da li request ima token?
Da li token validan?
Koja role?
Koji entityID?
Da li entityID sme da menja KooperantID parcele?
Da li action prolazi kroz requireRole/requireEntity?
```

### Korak 5: Doneti odluku

| Stanje                                   | Dozvoljena akcija                        |
| ---------------------------------------- | ---------------------------------------- |
| normalna korekcija, odobrena             | sačuvati novi polygon                    |
| pogrešan polygon, postoji validan backup | restore stare geometrije                 |
| nepoznat autor izmene                    | security incident, audit, restore        |
| public write potvrđen                    | hitan auth fix pre produkcije            |
| Lat/Lng pogrešan, polygon dobar          | recalculati centroid i sačuvati          |
| polygon pogrešan, Lat/Lng dobar          | iscrtati/restore polygon, zatim centroid |

---

## 7. Recovery scenariji

### 7.1. Polygon se ne vidi u PWA

Postupak:

1. Proveri da `ParcelaID` ima `PolygonGeoJSON` u backend tabeli.
2. Proveri da je `PolygonGeoJSON` validan JSON.
3. Proveri da li je geometry type `Polygon`.
4. Proveri da coordinates imaju format `[lng, lat]`, ne `[lat, lng]`.
5. Proveri browser console za JSON parse error.
6. Ako je sheet vrednost ispravna, uradi hard reload/cache clear.

### 7.2. Polygon je pogrešno pomeren

Postupak:

1. Ne crtati novi polygon pre čuvanja starog stanja.
2. Exportuj trenutni `PolygonGeoJSON`.
3. Proveri Google Sheet version history.
4. Nađi poslednju validnu geometriju.
5. Vrati staru geometriju ili iscrtaj novu uz odobrenje.
6. Sačuvaj novi centroid `Lat/Lng`.
7. Dokumentuj staru i novu vrednost.

### 7.3. Lat/Lng ne odgovara poligonu

Postupak:

1. Otvori parcelu u `parcel-draw.html`.
2. Učitaj polygon.
3. Edituj polygon minimalno ili ponovo sačuvaj da UI izračuna centroid.
4. Proveri da `Lat/Lng` pada unutar ili blizu poligona.
5. Ako je centroid nelogičan, proveri da polygon nije self-intersecting ili pogrešno formatiran.

### 7.4. Save polygon nije uspeo

Postupak:

1. Ne zatvarati stranicu dok se ne sačuva JSON preview.
2. Kopirati `PolygonGeoJSON` iz textarea preview-a u ticket.
3. Proveriti network response za `saveParcelPolygon`.
4. Ako je 401/403, proveriti token/session/role.
5. Ako je backend error, proveriti ErrorLog.
6. Nakon fix-a ponoviti save sa istim polygon JSON-om.

### 7.5. Korisnik kliknuo “Obriši polygon”

Postupak:

1. Utvrditi da li je kliknuo samo “Obriši polygon” ili i “Sačuvaj polygon”.
2. Ako nije sačuvao, server verovatno nije promenjen.
3. Reload parcele kroz “Učitaj parcelu”.
4. Ako je server ipak obrisan/promenjen, proveriti endpoint behavior i history.

### 7.6. Sumnja na neovlašćenu izmenu

Postupak:

1. Zaustaviti dalje GIS izmene.
2. Sačuvati trenutnu geometriju i history screenshot/export.
3. Proveriti da li je `saveParcelPolygon` bio public write.
4. Proveriti ErrorLog, Apps Script execution log i Drive history.
5. Vratiti poslednju validnu geometriju.
6. Tehnički owner zaključava endpoint tokenom/role check-om.
7. Security/poslovni owner odlučuje da li se incident prijavljuje kao security incident.

### 7.7. Kooperant promenio tuđu parcelu

Postupak:

1. Proveri token/session korisnika.
2. Proveri `KooperantID` parcele.
3. Ako se ne poklapaju, to je authorization bug.
4. Vratiti geometriju iz backup/history.
5. Blokirati endpoint ili dodati `requireEntity` proveru.
6. Pregledati sve parcele promenjene u istom periodu.

### 7.8. PolygonGeoJSON je nevalidan

Postupak:

1. Kopirati trenutnu vrednost u ticket.
2. Pokušati JSON parse u dev okruženju.
3. Ako JSON nije validan, vratiti staru validnu vrednost iz history-ja.
4. Ako je validan JSON ali nije validna GeoJSON geometry, ponovo nacrtati polygon.
5. Ne ostavljati nevalidan `PolygonGeoJSON` u produkciji jer može pokvariti mape i meteo/geofence tokove.

---

## 8. Kako sprečavaš pogrešne GIS izmene

Postojeće zaštite / očekivane zaštite:

1. `ParcelaID` je primarni ključ izmene.
2. UI prikazuje JSON preview pre save-a.
3. UI računa centroid iz poligona i šalje `Lat/Lng` zajedno sa polygon-om.
4. Leaflet Draw blokira intersection pri crtanju poligona.
5. Save endpoint treba da primi token.
6. Backend treba da proveri role/entity ownership.
7. Backend treba da čuva audit log pre/posle vrednosti.

Operativno pravilo:

> Ako nema audit traga ko je promenio polygon, svaka neočekivana promena se tretira kao data integrity incident.

---

## 9. Admin / DevTools provere

### 9.1. Provera URL-a

```text
parcel-draw.html?parcelaId=<ParcelaID>&token=<token>
```

### 9.2. Provera tokena u browser-u

```js
localStorage.getItem('authToken')
```

Ne kopirati pun token u ticket. Zabeležiti samo da li postoji.

### 9.3. Kopiranje trenutnog polygon preview-a

U `parcel-draw.html` kopirati sadržaj `PolygonGeoJSON preview` textarea.

### 9.4. Provera centroida

Ako je polygon učitan, UI automatski izračunava centroid pri draw/edit/save.

Proveriti:

```text
Lat field
Lng field
Google Maps open link
```

### 9.5. Test endpoint-a

Tehnički owner proverava:

```text
getParcelGeo bez tokena: da li je namerno public read?
saveParcelPolygon bez tokena: mora biti odbijen u produkciji
saveParcelPolygon sa Kooperant tokenom za tuđu parcelu: mora biti odbijen
saveParcelPolygon sa Management tokenom: dozvoljen
```

---

## 10. Ko donosi odluku

### Operator sme sam

* učitati parcelu;
* proveriti da li polygon postoji;
* kopirati JSON preview;
* proveriti Lat/Lng i Google Maps lokaciju;
* prijaviti mismatch;
* ne sme sam rešavati nepoznatu promenu.

### Management / GIS owner odlučuje

* da li je nova geometrija ispravna;
* da li se stara geometrija vraća;
* da li Kooperant sme sam menjati geometriju;
* da li `PRIV`/terenski nacrt ide u master geometriju;
* šta je canonical granica parcele.

### Tehnički owner odlučuje

* auth hardening `saveParcelPolygon` endpoint-a;
* restore iz version history-ja;
* ručnu izmenu `PolygonGeoJSON`;
* validaciju GeoJSON formata;
* audit log implementaciju;
* blokiranje public write endpoint-a.

### Security/poslovni owner odlučuje

* da li je nepoznata izmena security incident;
* da li se radi širi audit svih parcela;
* da li se incident prijavljuje interno/eksterno;
* da li se privremeno gasi GIS edit funkcija.

### Niko ne sme bez odobrenja

* menjati `PolygonGeoJSON` direktno u sheet-u;
* menjati `ParcelaID`;
* menjati `KooperantID` parcele radi “lakšeg save-a”;
* ostaviti public write u produkciji bez dokumentovane odluke;
* brisati history/error log tokom incidenta;
* precrtavati parcelu bez čuvanja starog stanja.

---

## 11. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan ParcelaID
[ ] Identifikovan KooperantID
[ ] Sačuvan stari PolygonGeoJSON
[ ] Sačuvan novi PolygonGeoJSON ako postoji
[ ] Proveren Lat/Lng
[ ] Proveren Google Maps link
[ ] Provereno da li je save uspeo ili pao
[ ] Proveren ErrorLog / Apps Script log ako je backend greška
[ ] Proverena autorizacija saveParcelPolygon endpoint-a
[ ] Ako je nepoznata izmena, klasifikovan security/data incident
[ ] Ako je restore rađen, dokumentovan izvor stare geometrije
[ ] Ako je korekcija rađena, odobrio GIS/Management owner
[ ] Korisnik obavešten
```

---

## 12. Primeri odluke

### Primer A: Polygon se ne vidi, ali `PolygonGeoJSON` postoji

Zaključak: verovatno UI parse/render problem ili nevalidan GeoJSON.
Akcija: proveriti JSON format i browser console; ne crtati novi polygon dok se ne sačuva stara vrednost.

### Primer B: Lat/Lng vodi u drugo selo, ali polygon je tačan

Zaključak: centroid/LatLng nije ažuriran.
Akcija: ponovo učitati/editovati/sačuvati polygon da se recalculiše centroid.

### Primer C: Polygon je promenjen, niko ne zna ko

Zaključak: data integrity/security incident.
Akcija: restore iz version history-ja, proveriti `saveParcelPolygon` auth, audit svih promena u periodu.

### Primer D: Kooperant može da sačuva polygon za tuđi `ParcelaID`

Zaključak: authorization bug.
Akcija: blokirati/hardenovati endpoint, vratiti pogođene geometrije, uraditi audit svih promena.

### Primer E: Kliknuto “Obriši polygon”, ali nije sačuvano

Zaključak: obrisano samo lokalno na ekranu.
Akcija: “Učitaj parcelu” vraća server stanje; ne radi restore.

---

## 13. Poznate production rupe koje treba zatvoriti

1. `saveParcelPolygon` mora biti eksplicitno iza token validation-a.
2. Dodati `requireRole` / `requireEntity` za GIS write.
3. Definisati politiku: Management-only GIS edit ili Kooperant-only own parcel edit.
4. Dodati `tblParcelGeoEventLog`: ko, kada, stari hash, novi hash, razlog.
5. Dodati pre-save confirmation sa starim i novim centroidom.
6. Dodati GeoJSON validator na backend-u.
7. Dodati size/area sanity check: polygon ne sme skočiti 100x bez override-a.
8. Dodati audit dashboard za parcele promenjene u poslednjih 24h/7d.
9. Dodati rollback funkciju iz poslednje validne geometrije.
10. Dodati warning ako `token` nedostaje u `parcel-draw.html`.
11. Dodati correlation/requestId u `saveParcelPolygon` response.
12. Dodati test: non-auth save mora pasti, Kooperant tuđa parcela mora pasti, Management save mora proći.

Do tada važi konzervativno pravilo:

> `PolygonGeoJSON` je production master podatak. Ako ne znaš ko ga je promenio i zašto, ne tretiraj to kao map bug nego kao data integrity incident.
