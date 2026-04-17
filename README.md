# Druk rezerwacji · Gazetka FAMIX

Aplikacja PWA do obsługi potwierdzeń udziału producentów w gazetkach
promocyjnych. Stylistyka w duchu macOS, dane ładowane z `./data/SC.xlsx`,
eksport do XLSX (zgodny z wzorem DEKLARACJA) i wydruk do PDF z przeglądarki.

## Funkcje

- **Automatyczne ładowanie bazy** — aplikacja przy starcie pobiera
  `./data/SC.xlsx` (filtrowanie producenta po kolumnie `znacznik`,
  `cena_s` = Cena FAMIX).
- **Szybkie wyszukiwanie** producenta i filtr po nazwie/kodzie EAN/towar_id.
- **Wybór indeksów** — checkboxy, zaznacz widoczne, wyczyść.
- **Moduł** promocji (¼ / ½ / ¾ / 1 strona) i jedna opłata dla całego druku.
- **Osoba kontaktowa** — pole tekstowe.
- **Podgląd wydruku 1:1** z wzorem FAMIX.
- **Eksport XLSX** — pełne formatowanie (ramki, kolory, merge, numFmt).
- **Wydruk / PDF** — A4 poziomo.
- **Offline (PWA)** — service worker cache'uje UI; baza odświeżana on-line,
  ale dostępna też w trybie offline.
- **Zapamiętywanie rozpoczętej rezerwacji** — `localStorage`.

## Struktura plików

```
/
├── index.html
├── app.js
├── styles.css
├── manifest.webmanifest
├── sw.js
├── icons/
│   ├── icon-192.png
│   └── icon-512.png
├── data/
│   └── SC.xlsx          <- aktualna baza produktów
└── README.md
```

## Wdrożenie na GitHub Pages

1. Utwórz repozytorium (np. `famix-druk-rezerwacji`) i wrzuć wszystkie pliki.
2. W zakładce **Settings → Pages** ustaw:
   - *Source:* `Deploy from a branch`
   - *Branch:* `main` / folder: `/ (root)`
3. Po kilku sekundach aplikacja będzie dostępna pod:
   `https://<użytkownik>.github.io/<repo>/`.
4. Podmianę bazy robisz nadpisaniem `data/SC.xlsx` (commit + push).
5. Aby dodać aplikację do ekranu głównego / Docka — otwórz w Safari
   (iPad/Mac) lub Chrome, *Udostępnij → Dodaj do ekranu startowego*
   (lub instalator w pasku adresu w Chrome).

## Odświeżanie bazy

Przycisk **Odśwież** w prawym górnym rogu pobiera `./data/SC.xlsx` na
nowo (z pominięciem cache przeglądarki). Po zmianie pliku w repo wystarczy
odczekać chwilę na CDN GitHub Pages lub twardy reload (`Shift+R`).

## Rozwój lokalny

Nie wymaga bundlera. Najprostszy serwer:

```bash
python3 -m http.server 8080
```

i otwórz `http://localhost:8080`.

## Kolumny z SC.xlsx używane przez aplikację

| Kolumna | Rola |
|---|---|
| `towar_id` | Indeks |
| `nazwa` | Nazwa produktu |
| `znacznik` | **Producent** (po tym filtrujemy) |
| `kod` | Kod EAN |
| `jm` | Jednostka miary |
| `vat` | Stawka VAT |
| `cena_s` | **Cena FAMIX** |
| `stale` | Rabat stały |
| `refundacje` | Refundacja odsprzedaży (zł) |
| `grupa`, `podgrupa` | Kategoryzacja |

Pola *Rabat prom.*, *Rab. Z/O*, *Promocja cenowa Netto/Brutto*,
*Prom. rabat.*, *Promocja Pakietowa* i *Uwagi* są puste w wydruku —
uzupełnia je producent ręcznie przed odesłaniem formularza.

## Licencja

Projekt wewnętrzny — używaj wewnątrz firmy.
