# PWMedia Word Add-in — Documentsjablonen

Een Word Add-in voor het invoegen van huisstijl-templates en bouwblokken.

---

## Projectstructuur

```
word-addin/
├── manifest.xml          ← Registratiebestand voor Office
├── src/
│   └── taskpane.html     ← De volledige add-in UI + logica
├── assets/
│   └── icons/            ← Iconen (16, 32, 64, 80, 128px PNG)
└── README.md
```

---

## Lokaal testen (stap voor stap)

### 1. Vereisten
- Node.js geïnstalleerd (https://nodejs.org)
- Microsoft Word (desktop)
- Windows of Mac

### 2. Installeer de Office generator (eenmalig)
```bash
npm install -g yo generator-office
npm install -g office-addin-dev-certs
```

### 3. Vertrouwd SSL-certificaat aanmaken (eenmalig)
```bash
npx office-addin-dev-certs install
```

### 4. Start een lokale server
De eenvoudigste manier:
```bash
npx serve . --ssl-cert ~/.office-addin-dev-certs/localhost.crt --ssl-key ~/.office-addin-dev-certs/localhost.key --listen 3000
```

Of gebruik de VS Code Live Server extensie met HTTPS ingeschakeld.

### 5. Laad de add-in in Word
1. Open Word
2. Ga naar **Invoegen → Add-ins → Mijn add-ins**
3. Klik op **Een aangepast manifest uploaden**
4. Kies het bestand `manifest.xml`
5. De add-in verschijnt in het lint onder **Start → Sjablonen**

---

## Aanpassen voor productie

### Huisstijlkleuren
Pas in `taskpane.html` de CSS variabelen aan:
```css
--brand:       #1A4F8A;   /* ← verander naar huisstijl primaire kleur */
--accent:      #D4821A;   /* ← verander naar accentkleur */
```

### Eigen sjabloonteksten
Zoek in `taskpane.html` op `buildDocumentTemplate` etc. en pas de standaardteksten aan.

### Beeldbibliotheek koppelen
In de functie `insertSelectedImage()`:
```javascript
// Laad afbeelding van SharePoint of Azure Blob:
const response = await fetch('https://uwsharepoint.sharepoint.com/...');
const blob = await response.blob();
// Converteer naar base64 en voeg in via:
// context.document.body.insertInlinePictureFromBase64(base64Data, 'End');
```

### Entiteiten configureren
Voeg in de instellingen-tab extra entiteiten toe aan de `<select>`:
```html
<option>Afdeling HR</option>
<option>Directie</option>
```

---

## Deployment naar klant (Microsoft 365 Admin)

1. Ga naar https://admin.microsoft.com
2. **Instellingen → Geïntegreerde apps → Apps uploaden**
3. Upload `manifest.xml`
4. Wijs toe aan gebruikers of groepen

De add-in verschijnt dan automatisch in Word van die gebruikers, zonder dat ze iets zelf hoeven te installeren.

---

## Volgende stappen voor uitbreiding

- [ ] Echt HTTPS hosting (Azure Static Web Apps = gratis)
- [ ] Afbeeldingen ophalen vanuit SharePoint-bibliotheek
- [ ] Automatische inhoudsopgave via OOXML field insertion
- [ ] Header/footer per entiteit instellen
- [ ] Huisstijlkleuren toepassen op tabelopmaak via OOXML
- [ ] E-mailhandtekening generator (apart HTML-bestand)

---

## Nuttige links

- [Office.js API documentatie](https://learn.microsoft.com/en-us/javascript/api/word)
- [Script Lab (live testen in Word)](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/explore-with-script-lab)
- [Office Add-ins samples (GitHub)](https://github.com/OfficeDev/Office-Add-in-samples)
