# Domácí pomocník v3

Tablet aplikace pro motivaci dítěte (8-9 let) k denní rutině.  
Wall-mounted tablet 1280×800 + rodičovský overlay (PIN výchozí **1234**, lze změnit).

## Co je v3 nového oproti v1

- 🌤️ **Reálné počasí + návrh oblečení** (Open-Meteo, Praha, bez API klíče, cache 30 min)
- 🔥 **Streaks** — počítadlo dní v řadě, oslava na 7/14/30 dnů
- 🗣️ **Riki mluví česky** (Web Speech API) — pozdrav, povzbuzení, vzkaz nahlas
- 🌙 **Noční režim** — po večerce ztmaví, Riki spí pod hvězdami
- 💌 **Vzkazy od rodiče** — napíšeš v rodičovském overlay, objevují se po dokončení rutiny
- 💜 **Mood check-in** — denní emoji picker, týdenní historie v rodičovském overlay
- ⚙️ **Nastavení** — jméno dítěte, jméno maskota, PIN, večerka, zvuk on/off, TTS on/off
- 💾 **Záloha JSON** — Export / Import / Reset všech dat
- 🦊 **Lepší Riki animace** — tančí při dokončení rutiny, spí v noci, prošlapává

## Soubory

- **`index.html`** — produkční verze, kterou GitHub Pages servíruje
- **`domaci-pomocnik-20260620-1920-v3.html`** — verzovaná kopie (referenční, lze smazat)

## Deploy na GitHub Pages

1. Vytvoř nový repo, např. `domaci-pomocnik` (public)
2. Nahraj `index.html` + README
3. Settings → Pages → Source: `Deploy from a branch` → Branch: `main` / `/ (root)` → Save
4. Za ~30s běží na `https://<user>.github.io/domaci-pomocnik/`

## Instalace na tabletu (jako PWA)

### Android (Chrome)
1. Otevři URL v Chrome
2. Menu (⋮) → **„Přidat na plochu"** / „Install app"
3. Ikona se objeví jako native app, otevírá se fullscreen landscape

### iPad (Safari)
1. Otevři URL v Safari
2. Share → **„Přidat na plochu"**

### Kiosk mode (doporučeno pro dítě)
- **Android:** **Fully Kiosk Browser** (zdarma, kvalitní, podporuje PWA, autostart)
- **iPad:** Guided Access (trojklik Home → Guided Access)

## První spuštění — co nastavit

1. Otevři aplikaci
2. Klepni na malou tečku vpravo nahoře (rodičovský zámek) → zadej **1234**
3. V rodičovském overlay přepni na záložku **⚙️ Nastavení** a uprav:
   - Jméno dítěte
   - Jméno maskota (např. nech „Riki" nebo dej co dítě chce)
   - PIN (změň z výchozího 1234)
   - Večerku a probuzení (kdy se aktivuje noční režim)
   - Zvuk a TTS hlas
4. Přepni na **💌 Vzkazy** a napiš pár vzkazů od mámy/táty
5. Přepni na **🎁 Odměny** a uprav obchod (cena, dostupnost)

## TTS (Riki mluví) — poznámky

- Vyžaduje českou TTS hlas v systému
- **Android Chrome:** většinou má český hlas vestavěný
- **iPad Safari:** funguje, hlas Zuzana (česky)
- **Windows Chrome:** stáhni Microsoft Czech voice z Settings → Language & region
- Pokud nemá český hlas: fallback na default (může číst česky s přízvukem)
- V Nastavení je **🔊 Otestovat hlas** tlačítko

## Počasí

- Bere reálná data z **Open-Meteo** (`api.open-meteo.com`, free, no API key)
- Pevně nastavená lokace: **Praha** (50.0755, 14.4378)
- Pro jinou lokaci uprav konstantu `LOCATION` v kódu
- Cache 30 min v localStorage

## Datová bezpečnost

- Všechna data jsou jen v **localStorage** prohlížeče tabletu
- Nejde to nikam ven, žádný cloud, žádný backend
- Pro zálohu/migraci na jiný tablet použij **Export JSON** v Nastavení
- Pro reset: **Smazat všechna data** v Nastavení

## Co stále chybí (TODO v4)

- **Pravý Service Worker** (sw.js) pro 100% offline — momentálně se cachuje proaktivně přes Cache API, ale není to 100% spolehlivé bez plnohodnotného SW
- **Skutečný Google Calendar** přes OAuth (zatím hardcoded události)
- **Hue webhook endpoint** — admin UI hotové, chybí backend
- **Push notifikace** — vyžaduje SW + backend
- **Více dětí** — současná verze předpokládá jedno
- **Sběratelské nálepky** — nápad pro v4

## Reset PIN když zapomeneš

1. Otevři aplikaci v Chrome
2. F12 → Application → Local Storage → `https://...` → smaž `dompomocnik_v1`
3. Reload — všechno se resetuje včetně PINu

Nebo přes URL: otevři DevTools Console a zadej:
```js
localStorage.removeItem('dompomocnik_v1'); location.reload();
```

## Technické pozadí

- Single-file HTML (React 18 + Babel CDN, žádný build step)
- ~2400 řádků, ~155 kB
- React functional component + hooks
- localStorage pro persistenci (klíč `dompomocnik_v1`)
- Web Audio API (zvuky generované)
- Web Speech API (TTS)
- Cache API (proaktivní cache)
- Open-Meteo API (počasí)
- 3 témata přes CSS proměnné
