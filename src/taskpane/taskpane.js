/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('count-words').onclick = countWords;
        // Nutzerinformationen anzeigen
        tryShowUserInfo();
    }
});

/**
 * Zeigt Nutzerinformationen (Name/E-Mail) im Taskpane, falls verfügbar
 */
function tryShowUserInfo() {
    try {
        const profile = Office.context?.mailbox?.userProfile;
        if (!profile) return;

        const name = profile.displayName || '';
        const email = profile.emailAddress || '';

        if (name || email) {
            const userInfo = document.getElementById('user-info');
            const nameEl = document.getElementById('user-name');
            const emailEl = document.getElementById('user-email');
            if (nameEl) nameEl.textContent = name;
            if (emailEl) emailEl.textContent = email;
            if (userInfo) userInfo.style.display = 'block';
        }
    } catch (e) {
        // still proceed silently; user info is optional
        console.warn('Konnte Nutzerinfo nicht lesen:', e);
    }
}

/**
 * Zählt Wörter und andere Statistiken in der aktuellen E-Mail
 */
async function countWords() {
    const button = document.getElementById('count-words');
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');
    const results = document.getElementById('results');
    
    // UI zurücksetzen
    button.disabled = true;
    loading.style.display = 'block';
    error.style.display = 'none';
    results.style.display = 'none';
    
    try {
        // E-Mail-Inhalt abrufen
        const emailContent = await getEmailContent();
        
        if (!emailContent) {
            throw new Error('Kein E-Mail-Inhalt gefunden. Bitte wählen Sie eine E-Mail aus.');
        }
        
        // Statistiken berechnen
        const stats = calculateTextStatistics(emailContent);
        
        // Ergebnisse anzeigen
        displayResults(stats);
        
    } catch (err) {
        console.error('Fehler beim Zählen der Wörter:', err);
        error.textContent = err.message || 'Ein unbekannter Fehler ist aufgetreten.';
        error.style.display = 'block';
    } finally {
        button.disabled = false;
        loading.style.display = 'none';
    }
}

/**
 * Ruft den Inhalt der aktuellen E-Mail ab
 */
function getEmailContent() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Text,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error('Fehler beim Abrufen des E-Mail-Inhalts: ' + result.error.message));
                }
            }
        );
    });
}

/**
 * Berechnet verschiedene Textstatistiken
 */
function calculateTextStatistics(text) {
    // Text bereinigen
    const cleanText = text.trim();
    
    if (!cleanText) {
        return {
            words: 0,
            characters: 0,
            charactersNoSpaces: 0,
            paragraphs: 0,
            sentences: 0,
            avgWordsPerSentence: 0,
            readingTimeMinutes: 0
        };
    }
    
    // Wörter zählen (Split by whitespace und Filter für leere Strings)
    const words = cleanText.split(/\s+/).filter(word => word.length > 0);
    const wordCount = words.length;
    
    // Zeichen zählen
    const characterCount = cleanText.length;
    const charactersNoSpaces = cleanText.replace(/\s/g, '').length;
    
    // Absätze zählen (Split by double line breaks oder mehr)
    const paragraphs = cleanText.split(/\n\s*\n/).filter(p => p.trim().length > 0);
    const paragraphCount = paragraphs.length;
    
    // Sätze zählen (Split by sentence endings)
    const sentences = cleanText.split(/[.!?]+/).filter(s => s.trim().length > 0);
    const sentenceCount = sentences.length;
    
    // Durchschnittliche Wörter pro Satz
    const avgWordsPerSentence = sentenceCount > 0 ? Math.round((wordCount / sentenceCount) * 10) / 10 : 0;
    
    // Lesezeit berechnen (durchschnittlich 200 Wörter pro Minute)
    const readingTimeMinutes = Math.ceil(wordCount / 200);
    
    return {
        words: wordCount,
        characters: characterCount,
        charactersNoSpaces: charactersNoSpaces,
        paragraphs: paragraphCount,
        sentences: sentenceCount,
        avgWordsPerSentence: avgWordsPerSentence,
        readingTimeMinutes: readingTimeMinutes
    };
}

/**
 * Zeigt die berechneten Statistiken in der UI an
 */
function displayResults(stats) {
    document.getElementById('word-count').textContent = stats.words.toLocaleString('de-DE');
    document.getElementById('char-count').textContent = stats.characters.toLocaleString('de-DE');
    document.getElementById('char-no-spaces').textContent = stats.charactersNoSpaces.toLocaleString('de-DE');
    document.getElementById('paragraph-count').textContent = stats.paragraphs.toLocaleString('de-DE');
    document.getElementById('sentence-count').textContent = stats.sentences.toLocaleString('de-DE');
    document.getElementById('avg-words-sentence').textContent = stats.avgWordsPerSentence.toLocaleString('de-DE');
    document.getElementById('reading-time').textContent = `${stats.readingTimeMinutes} Min.`;
    
    document.getElementById('results').style.display = 'block';
}

/**
 * Hilfsfunktion für Fehlerbehandlung
 */
function handleError(error, userMessage) {
    console.error('Fehler:', error);
    const errorDiv = document.getElementById('error');
    errorDiv.textContent = userMessage || 'Ein unbekannter Fehler ist aufgetreten.';
    errorDiv.style.display = 'block';
}