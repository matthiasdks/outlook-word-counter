# Assets Verzeichnis

Dieses Verzeichnis enthält die Icons für das Outlook Add-In.

## Benötigte Icons:

- `icon-16.png` - 16x16 Pixel Icon
- `icon-32.png` - 32x32 Pixel Icon  
- `icon-80.png` - 80x80 Pixel Icon

## Icon-Anforderungen:

- Format: PNG
- Transparenter Hintergrund empfohlen
- Einfaches, klares Design
- Gut erkennbar bei kleinen Größen

## Temporäre Lösung:

Da keine echten Icons vorhanden sind, können Sie:
1. Einfache PNG-Dateien mit den entsprechenden Größen erstellen
2. Online Icon-Generatoren verwenden
3. Die URLs in der manifest.xml auf externe Icons zeigen lassen

## Beispiel für externe Icons:

```xml
<IconUrl DefaultValue="https://via.placeholder.com/32x32/0078d4/ffffff?text=WC"/>
```