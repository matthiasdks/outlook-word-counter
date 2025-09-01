# Outlook 365 Wortzähler Add-In

Eine Outlook 365 Erweiterung, die Wörter und andere Textstatistiken in der aktuell ausgewählten E-Mail zählt.

## Features

- **Wortzählung**: Zählt die Anzahl der Wörter in der E-Mail
- **Zeichenzählung**: Mit und ohne Leerzeichen
- **Absätze und Sätze**: Zählt Absätze und Sätze
- **Lesezeit**: Schätzt die Lesezeit basierend auf durchschnittlicher Lesegeschwindigkeit
- **Durchschnittswerte**: Berechnet durchschnittliche Wörter pro Satz

## Installation

1. **Abhängigkeiten installieren:**
   ```bash
   npm install
   ```

2. **Entwicklungsserver starten:**
   ```bash
   npm run dev-server
   ```

3. **Add-In in Outlook laden:**
   - Öffnen Sie Outlook 365 (Web oder Desktop)
   - Gehen Sie zu "Add-Ins verwalten"
   - Wählen Sie "Benutzerdefiniertes Add-In hinzufügen"
   - Laden Sie die `manifest.xml` Datei hoch

## Entwicklung

### Verfügbare Scripts

- `npm run build` - Produktions-Build erstellen
- `npm run build:dev` - Entwicklungs-Build erstellen
- `npm run dev-server` - Entwicklungsserver starten
- `npm run start` - Add-In für Debugging starten
- `npm run stop` - Debugging stoppen
- `npm run validate` - Manifest validieren

### Projektstruktur

```
outlook-word-counter/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Haupt-UI
│   │   └── taskpane.js      # Wortzähler-Logik
│   └── commands/
│       ├── commands.html    # Ribbon-Commands
│       └── commands.js      # Command-Handler
├── assets/                  # Icons und Bilder
├── manifest.xml            # Add-In Manifest
├── package.json           # NPM Konfiguration
└── webpack.config.js      # Build-Konfiguration
```

## Verwendung

1. Öffnen Sie eine E-Mail in Outlook
2. Klicken Sie auf den "Wörter zählen" Button im Ribbon
3. Das Taskpane öffnet sich mit detaillierten Statistiken
4. Klicken Sie "Wörter zählen" um die Analyse zu starten

## Technische Details

- **Office.js API**: Verwendet die offizielle Microsoft Office JavaScript API
- **Webpack**: Für Build-Prozess und Entwicklungsserver
- **Babel**: Für JavaScript-Transpilation
- **HTTPS**: Erforderlich für Office Add-Ins

## Unterstützte Plattformen

- Outlook 365 (Web)
- Outlook 365 (Desktop)
- Outlook 2019
- Outlook 2016

## Lizenz

MIT License