# Érettségi Vizsga Beosztás Optimalizáló

Ez a szoftver egy érettségi szóbeli vizsgabizottság és a vizsgázók optimális napi beosztásának elkészítésére szolgál. A program minimalizálja a vizsgabizottság üresjáratait és biztosítja a vizsgázók számára az igazságos és minimális teremben töltött várakozási időt.

A projekt egyszerre kínál **Python CLI (parancssoros)** változatot és egy könnyen kezelhető, interaktív **böngészős webes felületet**.

---

## Főbb funkciók

- **Intelligens optimalizáció**: A nyelvi tárgyak (0 perc felkészülési idő) közbeékelésével "lyuktöltőként" szolgálnak, így a bizottság nem áll üresen, amíg a nem-nyelvi vizsgázók felkészülnek.
- **Kettős felület**:
  - Parancssori Python változat ([vizsga_beosztas.py](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/vizsga_beosztas.py)) Excel exportálási lehetőséggel.
  - Webalkalmazás ([index.html](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/index.html)) kézi adatbevitellel, Excel-importtal és interaktív vizualizációkkal.
- **Automatikus többnapos felosztás**: A megadott napi időkeret alapján a szoftver automatikusan napokra osztja a vizsgázókat.
- **Belépési idő igazítása**: A diákok nem a nap elejétől várakoznak a teremben, hanem pontosan az első vizsgáját megelőző felkészülés kezdetekor lépnek be.
- **Excel kompatibilitás**: Sablon alapján beolvassa a vizsgázókat és exportálja a részletes napi beosztásokat (külön lapon a bizottság és a diákok számára).

---

## Vizsgáztatási szabályok és korlátok

- **Vizsga hossza**: 15 perc vizsgázónként.
- **Szünet**: 2 perc kötelező szünet a bizottságnak a vizsgák között.
- **Felkészülési idő**: 
  - Nem-nyelvi tárgyaknál: **minimum 20 perc**.
  - Nyelvi tárgyaknál: **0 perc** (azonnal megkezdhető a vizsga).
- **Teremkapacitás**: Egyidejűleg legfeljebb **5 tanuló** tartózkodhat a teremben.
- **Benttartózkodás**: A tanulónak a teremben kell maradnia az első vizsgájára való felkészülés megkezdésétől az utolsó vizsgája befejezéséig.

---

## Használati útmutató

### 1. Böngészős változat (`index.html`)
Nincs szükség különösebb telepítésre vagy szerverre.
1. Nyissa meg az [index.html](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/index.html) fájlt bármilyen modern böngészőben.
2. Adja meg a paramétereket (kezdési/befejezési idő, felkészülési idő stb.).
3. Töltsön be tesztadatokat, importáljon egy Excel fájlt, vagy vigyen fel sorokat manuálisan.
4. Kattintson a **"Beosztás készítése"** gombra.
5. Tekintse meg az eredményeket a *Statisztika*, *Bizottság beosztása* és *Diákok beosztása* füleken, majd töltse le a kész Excel táblázatot.

*(Megjegyzés: A webes felület automatikusan ellenőrzi a paramétereket; ha a megadott időkeret túl kicsi egyetlen vizsga lebonyolításához is, a számítás nem indul el, elkerülve a böngésző fagyását.)*

### 2. Python változat (`vizsga_beosztas.py`)
Futtatásához Python 3 és az `openpyxl` könyvtár szükséges (utóbbi az Excel kezeléséhez).

#### Telepítés:
```bash
pip install openpyxl
```

#### Futtatás:
```bash
python vizsga_beosztas.py
```
A program futás során bekéri:
- A vizsgák kezdetét és végét (óra szerint, pl. 8 és 16).
- A bemeneti Excel fájl nevét (alapértelmezett: `vizsga_bemenet_sablon.xlsx`).

A futás végeztével a konzolon megjelenik a részletes beosztás, és létrejön a `vizsga_beosztas.xlsx` fájl.

---

## Bemeneti Excel formátum

Az importáláshoz használt Excel fájlnak az alábbi szerkezetet kell követnie:
- Az **A oszlop** tartalmazza a vizsgázó nevét.
- A **B oszloptól kezdődően** (oszlopok száma tetszőleges) a vizsgázó tárgyai szerepelnek.
- **Nyelvi tárgyak jelölése**: A nyelvi tárgyak neve elé egy **csillag (`*`)** karaktert kell tenni (pl. `*Angol nyelv`, `*Német nyelv`).

A mintasablon megtalálható a projektben: [vizsga_bemenet_sablon.xlsx](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/vizsga_bemenet_sablon.xlsx).

---

## További dokumentáció

A szoftver belső működésével, a szimulációs modellel és az optimalizáló algoritmusokkal kapcsolatos részletes leírást az [ARCHITECTURE.md](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/ARCHITECTURE.md) fájlban találja.
An English description of the project is available in [README_EN.md](file:///f:/Antigravity%20projektek/Viszgabeoszt%C3%A1s/README_EN.md).
