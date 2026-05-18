Szerepkör: Te egy vezető műszaki oktató, prezentáció-tervező és Python fejlesztő vagy.

Feladat: A csatolt PDF dokumentumból készíts egy technikai fókuszú, felsővezetői szinten is érthető (Executive Summary) prezentációt, majd írj egy Python szkriptet (python-pptx és PyMuPDF/fitz használatával), ami ebből a tartalomból és a PDF-ből kinyert képekből egy valós, formázott .pptx fájlt generál.

Kritériumok a tartalomhoz:
1. A prezentáció pontosan 10 diából álljon.
2. A nyelvezet legyen szakmai, lényegretörő (Executive szintű).
3. Minden diának legyen: egy egyértelmű Címe, maximum 4-5 Vázlatpontja (bullet point), és egy profi Előadói jegyzete (amit egy vezető mondana el). A jegyzet NE laikusoknak szóljon, maradjon szakmai.

Kritériumok a Python szkripthez:
1. Adatszerkezet: A szkriptbe már fixen égesd bele a generált 10 dia tartalmát (Cím, vázlatpontok, jegyzetek) magyar nyelven.
2. Képkinyerés: Használj PyMuPDF (fitz) könyvtárat. A kódban hozz létre egy `slide_to_page_map` szótárat, ahol manuálisan hozzárendeled a 10 diát a PDF azon oldalaihoz, ahol a témához legjobban illeszkedő ábrák/képek találhatók. Ezeket a szkript futásidőben nyerje ki.
3. Abszolút pozicionálás (Layout): A rálógások (overlap) elkerülése végett a python-pptx-ben NE az alapértelmezett sablon helykitöltőit használd, hanem `Inches` segítségével fixáld a helyeket:
   - Cím: Fixen felül (pl. top 0.4 inch, left 0.5 inch). Betűtípus: Arial, Félkövér.
   - Vázlatpontok szövegdoboza: Bal oldalon (pl. top 1.6 inch, left 0.5 inch, width 4.8 inch). A szöveg automatikus méretezése legyen bekapcsolva (`MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`). Betűtípus: Calibri.
   - Képek: Jobb oldalon (pl. top 1.6 inch, left 5.5 inch). Úgy állítsd be a szélességet, hogy kényelmesen elférjen a szövegdoboz mellett.
4. Kimenet: Egyetlen, tiszta, azonnal futtatható Python kódblokkot adj vissza, ami elvégzi a képek kinyerését, a prezentáció összeállítását és elmentését!