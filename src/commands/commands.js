/* global Office */

Office.onReady(() => {
    // Commands werden hier registriert
});

/**
 * Zeigt das Taskpane an
 */
function showTaskpane(event) {
    // Diese Funktion wird vom Ribbon-Button aufgerufen
    // Das Taskpane wird automatisch durch die Manifest-Konfiguration geöffnet
    event.completed();
}

// Globale Funktionen für Office registrieren
if (typeof window !== 'undefined') {
    window.showTaskpane = showTaskpane;
}