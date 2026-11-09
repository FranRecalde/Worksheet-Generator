// Pictorio & Pictorio Worksheets deterministic deck engine.
//
// Builds Worksheet (A4) and Slides (16:9 / 4:3) pages from Box A (Vocabulary)
// and Box B (Sentences) using 100% deterministic code and a single seeded RNG
// (mulberry32). The exact same pages feed the live HTML preview, A4 print view,
// and PowerPoint (.pptx) export.
(function () {
    'use strict';

    // ---------- Page sizes (inches) ----------
    const SIZES = {
        '16:9': { w: 10, h: 5.625 },
        '4:3': { w: 10, h: 7.5 },
        A4: { w: 8.27, h: 11.69 },
    };

    const NAVY = '0C1029';
    const GREY = '5A5E73';
    const MUSTARD = 'F2B33D';
    const MUSTARD_LIGHT = 'FDE68A';
    const TEAL = '2A7F7A';
    const ORANGE = 'E94C2D';
    const CREAM = 'F6E8CE';
    const PAPER = 'FFFBF3';
    const BORDER_GREY = 'C9CBD4';

    const ARASAAC_LINE = 'Pictograms: Sergio Palao. Origin: ARASAAC (https://arasaac.org). Licence: CC BY-NC-SA. Owner: Government of Aragón.';

    // ---------- Deterministic Seeded RNG (mulberry32) ----------
    const mulberry32 = (a) => () => {
        a |= 0;
        a = (a + 0x6D2B79F5) | 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };

    const shuffle = (arr, rnd) => {
        const a = arr.slice();
        for (let i = a.length - 1; i > 0; i--) {
            const j = Math.floor(rnd() * (i + 1));
            [a[i], a[j]] = [a[j], a[i]];
        }
        return a;
    };

    // True derangement: guarantees no element stays opposite its own match (order[i] !== i)
    // and if values are provided, guarantees values[order[i]] !== values[i] whenever possible.
    const derange = (n, rnd, values) => {
        if (n <= 1) return [0];
        const indices = Array.from({ length: n }, (_, i) => i);
        const isValid = (perm) => {
            for (let i = 0; i < n; i++) {
                if (perm[i] === i) return false;
                if (values && values[perm[i]] !== undefined && values[perm[i]] === values[i]) {
                    // Only reject if there exists another distinct value to swap with
                    const distinctCount = new Set(values).size;
                    if (distinctCount >= 2) return false;
                }
            }
            return true;
        };

        for (let attempt = 0; attempt < 120; attempt++) {
            const cand = shuffle(indices, rnd);
            if (isValid(cand)) return cand;
        }

        // Deterministic fallback: start from a shuffled permutation and fix any fixed points
        const base = shuffle(indices, rnd);
        for (let i = 0; i < n; i++) {
            if (base[i] === i || (values && values[base[i]] === values[i])) {
                // Find a j !== i such that swapping i and j leaves neither as a fixed point
                let swapped = false;
                for (let step = 1; step < n; step++) {
                    const j = (i + step) % n;
                    const valI = base[i];
                    const valJ = base[j];
                    const okI = valJ !== i && (!values || values[valJ] !== values[i]);
                    const okJ = valI !== j && (!values || values[valI] !== values[j]);
                    if (okI && okJ) {
                        base[i] = valJ;
                        base[j] = valI;
                        swapped = true;
                        break;
                    }
                }
                if (!swapped) {
                    const j = (i + 1) % n;
                    [base[i], base[j]] = [base[j], base[i]];
                }
            }
        }
        // Final absolute safety check for index identity
        if (base.some((v, idx) => v === idx)) {
            return indices.map((_, idx) => (idx + 1) % n);
        }
        return base;
    };

    const chunk = (arr, n) => {
        const out = [];
        for (let i = 0; i < arr.length; i += Math.max(1, n)) out.push(arr.slice(i, i + n));
        return out;
    };

    // ---------- Text Handling Rules (Section 4) ----------
    // Keep accents and special letters: á é í ó ú ñ ü, à â ç è é ê ë î ï ô ù û œ, ä ö ü ß, ì ò.
    // Vowels: a e i o u plus their accented forms (and ä ö ü). Y counts as a consonant.
    const VOWEL_CHARS = 'aeiouáéíóúàâèêëîïôùûäöüìòīūăąēėęĩįōŏőœæAEIOUÁÉÍÓÚÀÂÈÊËÎÏÔÙÛÄÖÜÌÒĪŪĂĄĒĖĘĨĮŌŎŐŒÆ';
    const VOWEL_SET = new Set(Array.from(VOWEL_CHARS));
    const isLetter = (ch) => /\p{L}/u.test(ch);
    const isVowel = (ch) => VOWEL_SET.has(ch);
    const isConsonant = (ch) => isLetter(ch) && !isVowel(ch);

    // Strip leading articles:
    // el/la/los/las/un/una, le/la/les/l'/un/une, der/die/das/ein/eine, il/lo/la/i/gli/le/un/una
    const ARTICLE_APOS_RE = /^(?:l['’])\s*(.+)$/i;
    const ARTICLE_WORD_RE = /^(?:el|la|los|las|un|una|le|les|une|der|die|das|ein|eine|il|lo|i|gli)\s+(.+)$/i;

    const stripArticle = (text) => {
        const t = (text || '').trim();
        if (!t) return '';
        const mApos = t.match(ARTICLE_APOS_RE);
        if (mApos && mApos[1]) return mApos[1].trim();
        const mWord = t.match(ARTICLE_WORD_RE);
        if (mWord && mWord[1]) return mWord[1].trim();
        return t;
    };

    // Uppercase helper where ß ALWAYS stays ß (Acceptance Check #6: die Straße -> STRAßE)
    const toPuzzleUpper = (str) => {
        return Array.from(str || '').map(ch => (ch === 'ß' ? 'ß' : ch.toUpperCase())).join('');
    };

    const removeDiacriticsKeepEszett = (str) => {
        return Array.from(str || '').map(ch => {
            if (ch === 'ß') return 'ß';
            if (ch === 'œ') return 'oe';
            if (ch === 'Œ') return 'OE';
            if (ch === 'æ') return 'ae';
            if (ch === 'Æ') return 'AE';
            return ch.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        }).join('');
    };

    // Converts a vocabulary word for puzzle grids (Word search, Crossword):
    // - optionally strips leading article (on by default)
    // - optionally removes accents (off by default; ß stays ß)
    // - removes spaces and non-letter punctuation
    // - uppercases with ß preserved as ß
    const toGridWord = (rawWord, opts = {}) => {
        const removeArticles = opts.removeArticles !== false;
        const removeAccents = !!opts.removeAccents;
        let w = removeArticles ? stripArticle(rawWord) : (rawWord || '').trim();
        if (removeAccents) {
            w = removeDiacriticsKeepEszett(w);
        }
        // Keep only Unicode letters
        const lettersOnly = Array.from(w).filter(ch => isLetter(ch)).join('');
        return toPuzzleUpper(lettersOnly);
    };

    // Formats target word for non-grid word puzzles (A4, A5, A6, A7, A8, A9)
    const getPuzzleTargetWord = (rawWord, opts = {}) => {
        const removeArticles = opts.removeArticles !== false;
        return removeArticles ? stripArticle(rawWord) : (rawWord || '').trim();
    };

    const NBSP = '\u00A0';
    const WORD_GAP = NBSP + NBSP + ' ';

    // A4: Fill in the vowels (vowels -> '_', consonants/spaces/apostrophes stay)
    const maskVowels = (text) => {
        return (text || '').split(/\s+/).filter(Boolean).map(w => {
            return Array.from(w).map(ch => (isVowel(ch) ? '_' : ch)).join(NBSP);
        }).join(WORD_GAP);
    };

    // B7: Missing vowels in sentences (keeps normal sentence spacing so sentences wrap naturally)
    const maskSentenceVowels = (sentence) => {
        return Array.from(sentence || '').map(ch => (isVowel(ch) ? '_' : ch)).join('');
    };

    // A5: Fill in the consonants (consonants -> '_', vowels/spaces/apostrophes stay; Y is consonant)
    const maskConsonants = (text) => {
        return (text || '').split(/\s+/).filter(Boolean).map(w => {
            return Array.from(w).map(ch => (isConsonant(ch) ? '_' : ch)).join(NBSP);
        }).join(WORD_GAP);
    };

    // A6: First and last letter: "p _ _ _ o". Words of 3 letters or fewer show only the first letter.
    // Each word of a multi-word item is treated on its own.
    const maskFirstLast = (text) => {
        return (text || '').split(/\s+/).filter(Boolean).map(w => {
            const chars = Array.from(w);
            const letterIndices = [];
            chars.forEach((ch, idx) => {
                if (isLetter(ch)) letterIndices.push(idx);
            });
            const n = letterIndices.length;
            if (n === 0) return w;
            const keepSet = new Set([letterIndices[0]]);
            if (n > 3) keepSet.add(letterIndices[n - 1]);
            return chars.map((ch, idx) => {
                if (!isLetter(ch)) return ch;
                return keepSet.has(idx) ? ch : '_';
            }).join(NBSP);
        }).join(WORD_GAP);
    };

    // A7: Anagrams: letters shuffled with the seed. Always different from the real word.
    const anagramWord = (text, rnd) => {
        return (text || '').split(/\s+/).filter(Boolean).map(w => {
            const chars = Array.from(w.toLowerCase());
            const letterPositions = [];
            const letters = [];
            chars.forEach((ch, idx) => {
                if (isLetter(ch)) {
                    letterPositions.push(idx);
                    letters.push(ch);
                }
            });
            if (letters.length < 2 || new Set(letters).size < 2) {
                return chars.join(NBSP);
            }
            const originalStr = letters.join('');
            let shuffled = letters.slice();
            for (let t = 0; t < 25; t++) {
                shuffled = shuffle(letters, rnd);
                if (shuffled.join('') !== originalStr) break;
            }
            if (shuffled.join('') === originalStr) {
                // Deterministic swap of first two distinct letters so it NEVER matches the answer
                shuffled = letters.slice();
                for (let i = 1; i < shuffled.length; i++) {
                    if (shuffled[i] !== shuffled[0]) {
                        [shuffled[0], shuffled[i]] = [shuffled[i], shuffled[0]];
                        break;
                    }
                }
            }
            const out = chars.slice();
            letterPositions.forEach((pos, idx) => {
                out[pos] = shuffled[idx];
            });
            return out.join(NBSP);
        }).join(WORD_GAP);
    };

    // B3: Jumbled sentence helper
    const jumbleSentence = (cleanSentence, rnd, lowercaseFirst = true) => {
        const s = (cleanSentence || '').trim();
        if (!s) return '';
        const leadMatch = s.match(/^([¿¡]+)/);
        const leadPunct = leadMatch ? leadMatch[1] : '';
        const afterLead = leadMatch ? s.slice(leadPunct.length).trim() : s;
        const trailMatch = afterLead.match(/([.!?…]+)$/);
        const trailPunct = trailMatch ? trailMatch[1] : '';
        const core = (trailMatch ? afterLead.slice(0, afterLead.length - trailPunct.length) : afterLead).trim();

        const words = core.split(/\s+/).filter(Boolean);
        if (words.length === 0) return s;
        const prepared = words.slice();
        if (lowercaseFirst && prepared[0]) {
            prepared[0] = prepared[0].toLowerCase();
        }
        const origOrder = prepared.join(' ');
        let mixed = prepared.slice();
        if (prepared.length >= 2 && new Set(prepared.map(w => w.toLowerCase())).size >= 2) {
            for (let t = 0; t < 30; t++) {
                mixed = shuffle(prepared, rnd);
                if (mixed.join(' ') !== origOrder) break;
            }
            if (mixed.join(' ') === origOrder) {
                mixed = prepared.slice();
                for (let i = 1; i < mixed.length; i++) {
                    if (mixed[i].toLowerCase() !== mixed[0].toLowerCase()) {
                        [mixed[0], mixed[i]] = [mixed[i], mixed[0]];
                        break;
                    }
                }
            }
        }
        return `${leadPunct ? leadPunct + ' ' : ''}${mixed.join(' / ')}${trailPunct ? ' ' + trailPunct : ''}`;
    };

    // ---------- Box A & Box B Input Parsers (Section 2) ----------
    // Box A: Vocabulary (one item per line): `el perro = the dog` (separator `=` or Tab)
    const parseVocabulary = (rawText) => {
        const text = String(rawText || '').replace(/\r\n/g, '\n');
        if (!text.trim()) {
            return { items: [], skipped: [], missingEnglishCount: 0, report: '0 words read, 0 lines skipped' };
        }
        const lines = text.split('\n');
        // If the very last line is empty because of a trailing newline, ignore trailing empty line unless it's in the middle
        const effectiveLength = (lines.length > 1 && lines[lines.length - 1] === '') ? lines.length - 1 : lines.length;
        const items = [];
        const skipped = [];
        let missingEnglishCount = 0;

        for (let i = 0; i < effectiveLength; i++) {
            const lineNum = i + 1;
            const rawLine = lines[i];
            const trimmed = rawLine.trim();
            if (!trimmed) {
                skipped.push({ line: lineNum, reason: 'no text' });
                continue;
            }
            const sepMatch = rawLine.match(/[=\t]/);
            if (sepMatch) {
                const idx = sepMatch.index;
                const left = rawLine.slice(0, idx).trim();
                const right = rawLine.slice(idx + 1).trim();
                if (!left && !right) {
                    skipped.push({ line: lineNum, reason: 'no text' });
                    continue;
                }
                if (!left) {
                    skipped.push({ line: lineNum, reason: 'no target word' });
                    continue;
                }
                const missingEnglish = !right;
                if (missingEnglish) missingEnglishCount++;
                items.push({
                    id: `vocab-${items.length}`,
                    index: items.length,
                    lineNum,
                    target: left,
                    english: right,
                    missingEnglish,
                });
            } else {
                missingEnglishCount++;
                items.push({
                    id: `vocab-${items.length}`,
                    index: items.length,
                    lineNum,
                    target: trimmed,
                    english: '',
                    missingEnglish: true,
                });
            }
        }

        const wordLabel = items.length === 1 ? '1 word read' : `${items.length} words read`;
        const skipLabel = skipped.length === 0
            ? '0 lines skipped'
            : `${skipped.length} ${skipped.length === 1 ? 'line' : 'lines'} skipped (${skipped.map(s => `line ${s.line}: ${s.reason}`).join('; ')})`;
        return {
            items,
            skipped,
            missingEnglishCount,
            report: `${wordLabel}, ${skipLabel}`,
        };
    };

    // Box B: Sentences (one per line): `Mi hermano tiene un [perro]. = My brother has a dog.`
    // Optional markup: [word] for gap-fill, *word* for picture prompt.
    const parseSentences = (rawText) => {
        const text = String(rawText || '').replace(/\r\n/g, '\n');
        if (!text.trim()) {
            return { items: [], skipped: [], missingEnglishCount: 0, picturePromptCount: 0, bracketCount: 0, report: '0 sentences read, 0 lines skipped' };
        }
        const lines = text.split('\n');
        const effectiveLength = (lines.length > 1 && lines[lines.length - 1] === '') ? lines.length - 1 : lines.length;
        const items = [];
        const skipped = [];
        let missingEnglishCount = 0;
        let picturePromptCount = 0;
        let bracketCount = 0;

        for (let i = 0; i < effectiveLength; i++) {
            const lineNum = i + 1;
            const rawLine = lines[i];
            const trimmed = rawLine.trim();
            if (!trimmed) {
                skipped.push({ line: lineNum, reason: 'no text' });
                continue;
            }
            const sepMatch = rawLine.match(/[=\t]/);
            let rawTarget = '';
            let english = '';
            if (sepMatch) {
                const idx = sepMatch.index;
                rawTarget = rawLine.slice(0, idx).trim();
                english = rawLine.slice(idx + 1).trim();
            } else {
                rawTarget = trimmed;
                english = '';
            }
            const cleanTarget = rawTarget.replace(/[\[\]*]/g, '').replace(/\s+/g, ' ').trim();
            if (!cleanTarget) {
                skipped.push({ line: lineNum, reason: 'no text' });
                continue;
            }
            const bracketMatch = rawTarget.match(/\[([^\][*]+)\]/);
            const asteriskMatch = rawTarget.match(/\*([^*[\]]+)\*/);
            const bracketWord = bracketMatch ? bracketMatch[1].trim() : '';
            const pictureKeyword = asteriskMatch ? asteriskMatch[1].trim() : '';

            if (bracketWord) bracketCount++;
            if (pictureKeyword) picturePromptCount++;
            const missingEnglish = !english;
            if (missingEnglish) missingEnglishCount++;

            items.push({
                id: `sent-${items.length}`,
                index: items.length,
                lineNum,
                rawTarget,
                target: cleanTarget,
                english,
                missingEnglish,
                bracketWord,
                pictureKeyword,
            });
        }

        const sentLabel = items.length === 1 ? '1 sentence read' : `${items.length} sentences read`;
        const skipLabel = skipped.length === 0
            ? '0 lines skipped'
            : `${skipped.length} ${skipped.length === 1 ? 'line' : 'lines'} skipped (${skipped.map(s => `line ${s.line}: ${s.reason}`).join('; ')})`;
        return {
            items,
            skipped,
            missingEnglishCount,
            picturePromptCount,
            bracketCount,
            report: `${sentLabel}, ${skipLabel}`,
        };
    };

    // ---------- Language & Rubric Lookup Tables (Section 5) ----------
    const I18N = {
        es: {
            langName: 'Spanish', nativeName: 'Español', defaultTitle: 'Ficha de vocabulario y frases',
            vocabTitle: 'Vocabulario', summary: 'Todas las palabras', answers: 'Soluciones', tf: ['V', 'F'],
            rubrics: {
                wordsearch: { title: 'Sopa de letras', rubric: 'Busca las palabras.' },
                crossword: { title: 'Crucigrama', rubric: 'Completa el crucigrama.' },
                matching: { title: 'Empareja', rubric: 'Empareja las columnas.' },
                fill_vowels: { title: 'Completa las vocales', rubric: 'Escribe las vocales que faltan.' },
                fill_consonants: { title: 'Completa las consonantes', rubric: 'Escribe las consonantes que faltan.' },
                first_last: { title: 'Primera y última letra', rubric: 'Completa las palabras.' },
                anagrams: { title: 'Anagramas', rubric: 'Ordena las letras para escribir cada palabra.' },
                label_picture: { title: 'Etiqueta los dibujos', rubric: 'Escribe la palabra debajo de cada dibujo.' },
                letter_boxes: { title: 'Casillas de letras', rubric: 'Escribe una letra en cada casilla.' },
                translate_to_target: { title: 'Traduce al español', rubric: 'Traduce las frases al español.' },
                translate_to_english: { title: 'Traduce al inglés', rubric: 'Traduce las frases al inglés.' },
                jumbled_sentences: { title: 'Frases desordenadas', rubric: 'Ordena las palabras para formar frases.' },
                sentence_gapfill: { title: 'Completa las frases', rubric: 'Completa cada frase con la palabra que falta.' },
                match_halves: { title: 'Une las mitades', rubric: 'Une las dos mitades de cada frase.' },
                match_translation: { title: 'Empareja frase y traducción', rubric: 'Empareja cada frase con su traducción.' },
                sentence_vowels: { title: 'Vocales en las frases', rubric: 'Completa las vocales que faltan en cada frase.' },
                picture_prompts: { title: 'Dibujos y frases', rubric: 'Mira el dibujo y completa la frase.' },
            }
        },
        fr: {
            langName: 'French', nativeName: 'Français', defaultTitle: 'Fiche de vocabulaire et phrases',
            vocabTitle: 'Vocabulaire', summary: 'Tous les mots', answers: 'Corrigé', tf: ['V', 'F'],
            rubrics: {
                wordsearch: { title: 'Mots mêlés', rubric: 'Trouve les mots.' },
                crossword: { title: 'Mots croisés', rubric: 'Complète les mots croisés.' },
                matching: { title: 'Relie', rubric: 'Relie les deux colonnes.' },
                fill_vowels: { title: 'Complète les voyelles', rubric: 'Écris les voyelles manquantes.' },
                fill_consonants: { title: 'Complète les consonnes', rubric: 'Écris les consonnes manquantes.' },
                first_last: { title: 'Première et dernière lettre', rubric: 'Complète les mots.' },
                anagrams: { title: 'Anagrammes', rubric: 'Remets les lettres dans l’ordre.' },
                label_picture: { title: 'Écris le mot', rubric: 'Écris le mot sous chaque image.' },
                letter_boxes: { title: 'Cases de lettres', rubric: 'Écris une lettre par case.' },
                translate_to_target: { title: 'Traduis en français', rubric: 'Traduis les phrases en français.' },
                translate_to_english: { title: 'Traduis en anglais', rubric: 'Traduis les phrases en anglais.' },
                jumbled_sentences: { title: 'Phrases mélangées', rubric: 'Remets les mots dans l’ordre.' },
                sentence_gapfill: { title: 'Complète les phrases', rubric: 'Complète les phrases avec le bon mot.' },
                match_halves: { title: 'Relie les deux moitiés', rubric: 'Relie le début et la fin de chaque phrase.' },
                match_translation: { title: 'Relie phrase et traduction', rubric: 'Relie chaque phrase à sa traduction.' },
                sentence_vowels: { title: 'Voyelles dans les phrases', rubric: 'Complète les voyelles dans chaque phrase.' },
                picture_prompts: { title: 'Images et phrases', rubric: 'Regarde l’image et complète la phrase.' },
            }
        },
        de: {
            langName: 'German', nativeName: 'Deutsch', defaultTitle: 'Arbeitsblatt: Wortschatz und Sätze',
            vocabTitle: 'Wortschatz', summary: 'Alle Wörter', answers: 'Lösungen', tf: ['R', 'F'],
            rubrics: {
                wordsearch: { title: 'Wortsuchrätsel', rubric: 'Finde die Wörter.' },
                crossword: { title: 'Kreuzworträtsel', rubric: 'Löse das Kreuzworträtsel.' },
                matching: { title: 'Zuordnen', rubric: 'Ordne die passenden Paare zu.' },
                fill_vowels: { title: 'Vokale ergänzen', rubric: 'Ergänze die fehlenden Vokale.' },
                fill_consonants: { title: 'Konsonanten ergänzen', rubric: 'Ergänze die fehlenden Konsonanten.' },
                first_last: { title: 'Erster und letzter Buchstabe', rubric: 'Ergänze die Wörter.' },
                anagrams: { title: 'Buchstabensalat', rubric: 'Ordne die Buchstaben.' },
                label_picture: { title: 'Bilder beschriften', rubric: 'Schreib das Wort unter jedes Bild.' },
                letter_boxes: { title: 'Buchstabenkästchen', rubric: 'Schreib einen Buchstaben in jedes Kästchen.' },
                translate_to_target: { title: 'Ins Deutsche übersetzen', rubric: 'Übersetze die Sätze ins Deutsche.' },
                translate_to_english: { title: 'Ins Englische übersetzen', rubric: 'Übersetze die Sätze ins Englische.' },
                jumbled_sentences: { title: 'Schüttelsätze', rubric: 'Bringe die Wörter in die richtige Reihenfolge.' },
                sentence_gapfill: { title: 'Lückentext', rubric: 'Ergänze das fehlende Wort in jedem Satz.' },
                match_halves: { title: 'Satzhälften verbinden', rubric: 'Verbinde die beiden Satzhälften.' },
                match_translation: { title: 'Satz und Übersetzung', rubric: 'Ordne jedem Satz die Übersetzung zu.' },
                sentence_vowels: { title: 'Vokale in Sätzen', rubric: 'Ergänze die fehlenden Vokale in jedem Satz.' },
                picture_prompts: { title: 'Bildimpulse', rubric: 'Schau das Bild an und ergänze den Satz.' },
            }
        },
        it: {
            langName: 'Italian', nativeName: 'Italiano', defaultTitle: 'Scheda di vocabolario e frasi',
            vocabTitle: 'Vocabolario', summary: 'Tutte le parole', answers: 'Soluzioni', tf: ['V', 'F'],
            rubrics: {
                wordsearch: { title: 'Crucipuzzle', rubric: 'Trova le parole.' },
                crossword: { title: 'Cruciverba', rubric: 'Completa il cruciverba.' },
                matching: { title: 'Abbina', rubric: 'Abbina le due colonne.' },
                fill_vowels: { title: 'Inserisci le vocali', rubric: 'Scrivi le vocali mancanti.' },
                fill_consonants: { title: 'Inserisci le consonanti', rubric: 'Scrivi le consonanti mancanti.' },
                first_last: { title: 'Prima e ultima lettera', rubric: 'Completa le parole.' },
                anagrams: { title: 'Anagrammi', rubric: 'Riordina le lettere.' },
                label_picture: { title: 'Scrivi la parola', rubric: 'Scrivi la parola sotto ogni immagine.' },
                letter_boxes: { title: 'Caselle di lettere', rubric: 'Scrivi una lettera in ogni casella.' },
                translate_to_target: { title: 'Traduci in italiano', rubric: 'Traduci le frasi in italiano.' },
                translate_to_english: { title: 'Traduci in inglese', rubric: 'Traduci le frasi in inglese.' },
                jumbled_sentences: { title: 'Frasi mescolate', rubric: 'Rimetti le parole in ordine.' },
                sentence_gapfill: { title: 'Completa le frasi', rubric: 'Completa ogni frase con la parola mancante.' },
                match_halves: { title: 'Unisci le metà', rubric: 'Unisci le due metà di ogni frase.' },
                match_translation: { title: 'Abbina frase e traduzione', rubric: 'Abbina ogni frase alla sua traduzione.' },
                sentence_vowels: { title: 'Vocali nelle frasi', rubric: 'Inserisci le vocali mancanti in ogni frase.' },
                picture_prompts: { title: 'Immagini e frasi', rubric: 'Guarda l’immagine e completa la frase.' },
            }
        },
        other: {
            langName: 'Target language', nativeName: 'Target language', defaultTitle: 'Vocabulary & Sentences Worksheet',
            vocabTitle: 'Vocabulary', summary: 'All words', answers: 'Answer Key', tf: ['T', 'F'],
            rubrics: {}
        }
    };

    // Complete Task Catalogue (Section 5: Group A = 9 word tasks, Group B = 8 sentence tasks)
    const TASK_CATALOGUE = [
        // Group A: Word tasks (from Box A)
        {
            id: 'wordsearch',
            group: 'A',
            code: 'A1',
            name: 'Word search',
            enRubric: 'Find the words in the grid.',
        },
        {
            id: 'crossword',
            group: 'A',
            code: 'A2',
            name: 'Crossword',
            enRubric: 'Complete the crossword using the clues.',
        },
        {
            id: 'matching',
            group: 'A',
            code: 'A3',
            name: 'Matching',
            enRubric: 'Match each item on the left with the correct item on the right. Write the letter in the box.',
        },
        {
            id: 'fill_vowels',
            group: 'A',
            code: 'A4',
            name: 'Fill in the vowels',
            enRubric: 'Fill in the missing vowels to complete each word.',
        },
        {
            id: 'fill_consonants',
            group: 'A',
            code: 'A5',
            name: 'Fill in the consonants',
            enRubric: 'Fill in the missing consonants to complete each word.',
        },
        {
            id: 'first_last',
            group: 'A',
            code: 'A6',
            name: 'First and last letter',
            enRubric: 'Complete each word using the given letters as clues.',
        },
        {
            id: 'anagrams',
            group: 'A',
            code: 'A7',
            name: 'Anagrams',
            enRubric: 'Unscramble the letters to write each word.',
        },
        {
            id: 'label_picture',
            group: 'A',
            code: 'A8',
            name: 'Label the picture',
            enRubric: 'Write the word for each picture.',
        },
        {
            id: 'letter_boxes',
            group: 'A',
            code: 'A9',
            name: 'Letter boxes',
            enRubric: 'Write one letter in each box to spell the word.',
        },
        // Group B: Sentence tasks (from Box B)
        {
            id: 'translate_to_target',
            group: 'B',
            code: 'B1',
            name: 'Translate into the target language',
            enRubric: 'Translate each sentence into the target language.',
        },
        {
            id: 'translate_to_english',
            group: 'B',
            code: 'B2',
            name: 'Translate into English',
            enRubric: 'Translate each sentence into English.',
        },
        {
            id: 'jumbled_sentences',
            group: 'B',
            code: 'B3',
            name: 'Jumbled sentences',
            enRubric: 'Put the words in the correct order to make a sentence.',
        },
        {
            id: 'sentence_gapfill',
            group: 'B',
            code: 'B4',
            name: 'Gap fill',
            enRubric: 'Complete each sentence with the missing word.',
        },
        {
            id: 'match_halves',
            group: 'B',
            code: 'B5',
            name: 'Match the halves',
            enRubric: 'Match the beginning of each sentence with its ending. Write the letter in the box.',
        },
        {
            id: 'match_translation',
            group: 'B',
            code: 'B6',
            name: 'Match sentence to translation',
            enRubric: 'Match each sentence on the left with its English translation on the right.',
        },
        {
            id: 'sentence_vowels',
            group: 'B',
            code: 'B7',
            name: 'Missing vowels in sentences',
            enRubric: 'Fill in the missing vowels to complete each sentence.',
        },
        {
            id: 'picture_prompts',
            group: 'B',
            code: 'B8',
            name: 'Picture prompts',
            enRubric: 'Look at the picture and write the missing word in each sentence.',
        },
    ];

    const DEFAULT_TASK_SETTINGS = {
        wordsearch: { gridSize: 12, difficulty: 'easy' }, // 'easy' | 'medium' | 'hard'
        crossword: {},
        matching: { pairMode: 'auto' }, // 'auto' | 'picture-target' | 'english-target' | 'picture-english'
        fill_vowels: {},
        fill_consonants: {},
        first_last: {},
        anagrams: {},
        label_picture: { mode: 'blank' }, // 'blank' | 'scaffold'
        letter_boxes: {},
        translate_to_target: {},
        translate_to_english: {},
        jumbled_sentences: { lowercaseFirst: true, showEnglish: false },
        sentence_gapfill: { wordBank: true, showFirstLetter: false },
        match_halves: {},
        match_translation: {},
        sentence_vowels: { showEnglish: false },
        picture_prompts: { wordBank: false },
    };

    // Evaluates whether a task can run with the current inputs and settings, and returns reason if disabled
    const getTaskAvailability = (taskId, state) => {
        const vocab = state.vocab || [];
        const sentences = state.sentences || [];
        const withPictures = !!state.withPictures;
        const vocabWithEng = vocab.filter(v => v.english && v.english.trim().length > 0);
        const sentWithEng = sentences.filter(s => s.english && s.english.trim().length > 0);
        const sentWithPicKeyword = sentences.filter(s => s.pictureKeyword && s.pictureKeyword.trim().length > 0);

        const meta = TASK_CATALOGUE.find(t => t.id === taskId);
        if (!meta) return { available: false, reason: 'Unknown task' };

        if (meta.group === 'A') {
            if (vocab.length === 0) {
                return { available: false, reason: 'Needs vocabulary in Box A' };
            }
            if (taskId === 'label_picture') {
                if (!withPictures) return { available: false, reason: 'Needs pictures on' };
                return { available: true, reason: '' };
            }
            if (taskId === 'matching') {
                if (vocab.length < 2) return { available: false, reason: 'Needs at least 2 words in Box A' };
                if (!withPictures && vocabWithEng.length < 2) {
                    return { available: false, reason: 'Needs English in Box A or pictures on' };
                }
                return { available: true, reason: '' };
            }
            if (taskId === 'crossword') {
                if (vocab.length < 2) return { available: false, reason: 'Needs at least 2 words in Box A' };
                if (!withPictures && vocabWithEng.length === 0) {
                    return { available: false, reason: 'Needs English in Box A or pictures on' };
                }
                return { available: true, reason: '' };
            }
            // wordsearch, fill_vowels, fill_consonants, first_last, anagrams, letter_boxes
            if (!withPictures && vocabWithEng.length === 0 && taskId !== 'wordsearch') {
                return { available: false, reason: 'Needs English in Box A or pictures on' };
            }
            return { available: true, reason: '' };
        }

        // Group B: Sentence tasks
        if (sentences.length === 0) {
            return { available: false, reason: 'Needs sentences in Box B' };
        }
        if (taskId === 'translate_to_target') {
            if (sentWithEng.length === 0) return { available: false, reason: 'Needs English translations in Box B' };
            return { available: true, reason: '' };
        }
        if (taskId === 'translate_to_english') {
            return { available: true, reason: '' };
        }
        if (taskId === 'jumbled_sentences') {
            const multiWord = sentences.filter(s => s.target.split(/\s+/).filter(Boolean).length >= 2);
            if (multiWord.length === 0) return { available: false, reason: 'Needs sentences with at least 2 words' };
            return { available: true, reason: '' };
        }
        if (taskId === 'sentence_gapfill') {
            return { available: true, reason: '' };
        }
        if (taskId === 'match_halves') {
            const multiWord = sentences.filter(s => s.target.split(/\s+/).filter(Boolean).length >= 2);
            if (multiWord.length < 2) return { available: false, reason: 'Needs at least 2 sentences in Box B' };
            return { available: true, reason: '' };
        }
        if (taskId === 'match_translation') {
            if (sentWithEng.length < 2) return { available: false, reason: 'Needs at least 2 sentences with English' };
            return { available: true, reason: '' };
        }
        if (taskId === 'sentence_vowels') {
            return { available: true, reason: '' };
        }
        if (taskId === 'picture_prompts') {
            if (!withPictures) return { available: false, reason: 'Needs pictures on' };
            if (sentWithPicKeyword.length === 0) return { available: false, reason: 'Needs *keyword* markup in Box B' };
            return { available: true, reason: '' };
        }

        return { available: true, reason: '' };
    };

    // ---------- Puzzle Generators (Word Search & Crossword) ----------
    // A1. Word Search Generator
    const generateWordSearch = (vocabItems, opts, rnd) => {
        const gridSize = Math.max(8, Math.min(20, Number(opts.gridSize) || 12));
        const difficulty = opts.difficulty || 'easy';
        let dirs = [[0, 1], [1, 0]]; // Easy: right and down
        if (difficulty === 'medium') {
            dirs = [[0, 1], [1, 0], [1, 1], [-1, 1]]; // Medium: adds diagonals
        } else if (difficulty === 'hard') {
            dirs = [[0, 1], [1, 0], [1, 1], [-1, 1], [0, -1], [-1, 0], [-1, -1], [1, -1]]; // Hard: adds backwards
        }

        const prepared = vocabItems.map((item, idx) => ({
            item,
            idx,
            gridWord: toGridWord(item.target, opts),
        })).filter(x => x.gridWord.length > 0);

        // Placement: longest word first, break ties by original order
        const sorted = prepared.slice().sort((a, b) => (b.gridWord.length - a.gridWord.length) || (a.idx - b.idx));
        const grid = Array.from({ length: gridSize }, () => Array(gridSize).fill(''));
        const placed = [];
        const unplaced = [];

        for (const entry of sorted) {
            const w = entry.gridWord;
            const len = Array.from(w).length;
            const chars = Array.from(w);
            if (len > gridSize) {
                unplaced.push({
                    item: entry.item,
                    gridWord: w,
                    reason: `"${entry.item.target}" (${len} letters) is longer than the ${gridSize}×${gridSize} grid`,
                });
                continue;
            }

            let ok = false;
            // Up to 200 tries per word using seeded positions
            for (let t = 0; t < 200 && !ok; t++) {
                const [dr, dc] = dirs[Math.floor(rnd() * dirs.length)];
                const minR = dr < 0 ? len - 1 : 0;
                const maxR = dr > 0 ? gridSize - len : gridSize - 1;
                const minC = dc < 0 ? len - 1 : 0;
                const maxC = dc > 0 ? gridSize - len : gridSize - 1;
                if (minR > maxR || minC > maxC) continue;

                const r0 = minR + Math.floor(rnd() * (maxR - minR + 1));
                const c0 = minC + Math.floor(rnd() * (maxC - minC + 1));

                let fits = true;
                for (let k = 0; k < len; k++) {
                    const cell = grid[r0 + dr * k][c0 + dc * k];
                    if (cell !== '' && cell !== chars[k]) {
                        fits = false;
                        break;
                    }
                }
                if (!fits) continue;

                const cells = [];
                for (let k = 0; k < len; k++) {
                    const r = r0 + dr * k;
                    const c = c0 + dc * k;
                    grid[r][c] = chars[k];
                    cells.push([r, c]);
                }
                placed.push({ ...entry, chars, cells, r0, c0, dr, dc });
                ok = true;
            }

            if (!ok) {
                unplaced.push({
                    item: entry.item,
                    gridWord: w,
                    reason: `"${entry.item.target}" could not be placed in the ${gridSize}×${gridSize} grid after 200 tries`,
                });
            }
        }

        // Fill empty cells with seeded random letters taken ONLY from the letters in the word list
        const allChars = Array.from(new Set(prepared.flatMap(p => Array.from(p.gridWord))));
        const fillPool = allChars.length > 0 ? allChars : Array.from('ABCDEFGHIJLMNOPQRSTUVZ');

        for (let r = 0; r < gridSize; r++) {
            for (let c = 0; c < gridSize; c++) {
                if (grid[r][c] === '') {
                    grid[r][c] = fillPool[Math.floor(rnd() * fillPool.length)];
                }
            }
        }

        return {
            size: gridSize,
            grid,
            placed: placed.slice().sort((a, b) => a.idx - b.idx),
            unplaced,
        };
    };

    // A2. Crossword Generator
    // Order words by length (longest first), break ties alphabetically. Place first word across.
    // For each next word, try every letter crossing. Pick the crossing with the most crossings
    // and the smallest grid. Break remaining ties with the seed.
    // Two words must never touch side by side except where they cross.
    const generateCrossword = (vocabItems, opts, rnd) => {
        const prepared = vocabItems.map((item, idx) => {
            const gridWord = toGridWord(item.target, opts);
            return {
                item,
                idx,
                gridWord,
                chars: Array.from(gridWord),
            };
        }).filter(x => x.chars.length >= 2);

        if (prepared.length === 0) {
            return { rows: 0, cols: 0, cells: [], across: [], down: [], unplaced: [] };
        }

        const sorted = prepared.slice().sort((a, b) => {
            if (b.chars.length !== a.chars.length) return b.chars.length - a.chars.length;
            return a.gridWord.localeCompare(b.gridWord);
        });

        const placed = [];
        const unplaced = [];

        // Build occupancy lookup from current placed words
        const buildBoardMap = () => {
            const map = new Map(); // "r,c" -> { letter, across: boolean, down: boolean }
            let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
            for (const p of placed) {
                const dr = p.dir === 'down' ? 1 : 0;
                const dc = p.dir === 'across' ? 1 : 0;
                for (let k = 0; k < p.chars.length; k++) {
                    const r = p.r + dr * k;
                    const c = p.c + dc * k;
                    if (r < minR) minR = r;
                    if (r > maxR) maxR = r;
                    if (c < minC) minC = c;
                    if (c > maxC) maxC = c;
                    const key = `${r},${c}`;
                    const ex = map.get(key) || { letter: p.chars[k], across: false, down: false };
                    if (p.dir === 'across') ex.across = true;
                    else ex.down = true;
                    map.set(key, ex);
                }
            }
            return { map, minR, maxR, minC, maxC };
        };

        // Place first word across at (0, 0)
        placed.push({
            ...sorted[0],
            r: 0,
            c: 0,
            dir: 'across',
        });

        // Evaluate candidate placement of `entry` at (r0, c0, dir)
        const evaluatePlacement = (entry, r0, c0, dir, board) => {
            const dr = dir === 'down' ? 1 : 0;
            const dc = dir === 'across' ? 1 : 0;
            const L = entry.chars.length;
            const { map } = board;

            // Cell immediately before start and immediately after end must be empty
            if (map.has(`${r0 - dr},${c0 - dc}`)) return null;
            if (map.has(`${r0 + dr * L},${c0 + dc * L}`)) return null;

            let crossings = 0;
            for (let k = 0; k < L; k++) {
                const r = r0 + dr * k;
                const c = c0 + dc * k;
                const ch = entry.chars[k];
                const ex = map.get(`${r},${c}`);

                if (ex) {
                    if (ex.letter !== ch) return null;
                    // Cannot overlap in the same direction
                    if (dir === 'across' && ex.across) return null;
                    if (dir === 'down' && ex.down) return null;
                    crossings++;
                } else {
                    // Non-crossing cell: perpendicular neighbors must be empty so words never touch side by side
                    if (dir === 'across') {
                        if (map.has(`${r - 1},${c}`) || map.has(`${r + 1},${c}`)) return null;
                    } else {
                        if (map.has(`${r},${c - 1}`) || map.has(`${r},${c + 1}`)) return null;
                    }
                }
            }

            if (crossings === 0) return null;

            const newMinR = Math.min(board.minR, r0);
            const newMaxR = Math.max(board.maxR, r0 + dr * (L - 1));
            const newMinC = Math.min(board.minC, c0);
            const newMaxC = Math.max(board.maxC, c0 + dc * (L - 1));
            const h = newMaxR - newMinR + 1;
            const w = newMaxC - newMinC + 1;
            // Reject absurdly huge aspect ratios if avoidable
            if (Math.max(w, h) > 26) return null;

            return {
                r: r0,
                c: c0,
                dir,
                crossings,
                area: w * h,
                maxSide: Math.max(w, h),
                tie: rnd(),
            };
        };

        for (let i = 1; i < sorted.length; i++) {
            const entry = sorted[i];
            const board = buildBoardMap();
            const candidates = [];
            const seen = new Set();

            for (const p of placed) {
                const candDir = p.dir === 'across' ? 'down' : 'across';
                const pDr = p.dir === 'down' ? 1 : 0;
                const pDc = p.dir === 'across' ? 1 : 0;
                const cDr = candDir === 'down' ? 1 : 0;
                const cDc = candDir === 'across' ? 1 : 0;

                for (let kp = 0; kp < p.chars.length; kp++) {
                    const chP = p.chars[kp];
                    const interR = p.r + pDr * kp;
                    const interC = p.c + pDc * kp;

                    for (let kw = 0; kw < entry.chars.length; kw++) {
                        if (entry.chars[kw] !== chP) continue;
                        const r0 = interR - cDr * kw;
                        const c0 = interC - cDc * kw;
                        const posKey = `${r0},${c0},${candDir}`;
                        if (seen.has(posKey)) continue;
                        seen.add(posKey);

                        const res = evaluatePlacement(entry, r0, c0, candDir, board);
                        if (res) candidates.push(res);
                    }
                }
            }

            if (candidates.length > 0) {
                candidates.sort((a, b) => {
                    if (b.crossings !== a.crossings) return b.crossings - a.crossings;
                    if (a.area !== b.area) return a.area - b.area;
                    if (a.maxSide !== b.maxSide) return a.maxSide - b.maxSide;
                    return a.tie - b.tie;
                });
                const best = candidates[0];
                placed.push({
                    ...entry,
                    r: best.r,
                    c: best.c,
                    dir: best.dir,
                });
            } else {
                unplaced.push({
                    item: entry.item,
                    gridWord: entry.gridWord,
                    reason: `"${entry.item.target}" shares no valid crossing with placed words`,
                });
            }
        }

        // Normalize coordinates to (0, 0)
        const finalBoard = buildBoardMap();
        const rowOffset = -finalBoard.minR;
        const colOffset = -finalBoard.minC;
        const rows = finalBoard.maxR - finalBoard.minR + 1;
        const cols = finalBoard.maxC - finalBoard.minC + 1;

        placed.forEach(p => {
            p.r += rowOffset;
            p.c += colOffset;
        });

        // Number clues in standard reading order (top-to-bottom, left-to-right)
        const startLocations = [];
        placed.forEach(p => {
            if (!startLocations.some(loc => loc.r === p.r && loc.c === p.c)) {
                startLocations.push({ r: p.r, c: p.c });
            }
        });
        startLocations.sort((a, b) => (a.r - b.r) || (a.c - b.c));

        const numMap = new Map();
        startLocations.forEach((loc, idx) => {
            numMap.set(`${loc.r},${loc.c}`, idx + 1);
        });

        placed.forEach(p => {
            p.clueNum = numMap.get(`${p.r},${p.c}`);
        });

        // Build 2D cell matrix
        const cellGrid = Array.from({ length: rows }, () => Array.from({ length: cols }, () => null));
        for (const p of placed) {
            const dr = p.dir === 'down' ? 1 : 0;
            const dc = p.dir === 'across' ? 1 : 0;
            for (let k = 0; k < p.chars.length; k++) {
                const r = p.r + dr * k;
                const c = p.c + dc * k;
                if (!cellGrid[r][c]) {
                    cellGrid[r][c] = { r, c, letter: p.chars[k], num: numMap.get(`${r},${c}`) || null };
                }
            }
        }

        const across = placed.filter(p => p.dir === 'across').sort((a, b) => a.clueNum - b.clueNum);
        const down = placed.filter(p => p.dir === 'down').sort((a, b) => a.clueNum - b.clueNum);

        return {
            rows,
            cols,
            cellGrid,
            placed,
            across,
            down,
            unplaced,
        };
    };

    // ---------- Sentence Gap-Fill Target Word Resolver (Section 5.B.4) ----------
    // Order of priority:
    // 1. words in [brackets]
    // 2. if none, the first word that is also in Box A
    // 3. if none, the longest word (first one if there's a tie)
    const resolveSentenceGap = (sentItem, vocabItems) => {
        // 1. Check [brackets] in rawTarget
        if (sentItem.bracketWord) {
            const cleanWithPlaceholder = sentItem.rawTarget
                .replace(/\*/g, '')
                .replace(/\[([^\][*]+)\]/, '___GAP___')
                .replace(/[\[\]]/g, '');
            return {
                answer: sentItem.bracketWord,
                template: cleanWithPlaceholder,
            };
        }

        // Tokenize clean sentence keeping word positions
        const target = sentItem.target;
        const tokenMatches = Array.from(target.matchAll(/\p{L}+(?:['’]\p{L}+)*/gu));
        if (tokenMatches.length === 0) {
            return { answer: target, template: '___GAP___' };
        }

        // Build lookup set of Box A words (both full target and article-stripped target, lowercased)
        const boxASet = new Set();
        (vocabItems || []).forEach(v => {
            const full = (v.target || '').toLowerCase().trim();
            const stripped = stripArticle(v.target).toLowerCase().trim();
            if (full) boxASet.add(full);
            if (stripped) boxASet.add(stripped);
        });

        // 2. First word in sentence that is also in Box A
        let chosenMatch = null;
        for (const m of tokenMatches) {
            const wLower = m[0].toLowerCase();
            if (boxASet.has(wLower)) {
                chosenMatch = m;
                break;
            }
        }

        // 3. Otherwise, the longest word (first one if there's a tie)
        if (!chosenMatch) {
            chosenMatch = tokenMatches[0];
            for (const m of tokenMatches) {
                if (Array.from(m[0]).length > Array.from(chosenMatch[0]).length) {
                    chosenMatch = m;
                }
            }
        }

        const startIdx = chosenMatch.index;
        const wordStr = chosenMatch[0];
        const template = target.slice(0, startIdx) + '___GAP___' + target.slice(startIdx + wordStr.length);
        return {
            answer: wordStr,
            template,
        };
    };

    // Split sentence at the word nearest the middle (Section 5.B.5)
    const splitSentenceInHalf = (cleanSentence) => {
        const words = (cleanSentence || '').trim().split(/\s+/).filter(Boolean);
        if (words.length < 2) {
            return [cleanSentence, ''];
        }
        const totalLen = cleanSentence.length;
        const midChar = totalLen / 2;
        let bestK = 1;
        let bestDist = Infinity;
        for (let k = 1; k < words.length; k++) {
            const leftLen = words.slice(0, k).join(' ').length;
            const dist = Math.abs(leftLen - midChar);
            if (dist < bestDist) {
                bestDist = dist;
                bestK = k;
            }
        }
        return [words.slice(0, bestK).join(' '), words.slice(bestK).join(' ')];
    };

    // ---------- Page Primitive Helpers ----------
    const fitPt = (text, w, h, maxPt, minPt, mono, bold) => {
        const paras = String(text || '').split('\n');
        for (let pt = maxPt; pt >= minPt; pt -= 0.5) {
            const charW = (pt / 72) * (mono ? 0.6 : bold ? 0.55 : 0.51);
            const perLine = Math.max(1, Math.floor(w / charW));
            let lines = 0;
            let ok = true;
            for (const p of paras) {
                const words = p.split(' ');
                let cur = 0;
                let l = 1;
                for (const wd of words) {
                    const len = Array.from(wd).length;
                    if (len > perLine) { ok = false; break; }
                    if (cur === 0) cur = len;
                    else if (cur + 1 + len <= perLine) cur += 1 + len;
                    else { l++; cur = len; }
                }
                if (!ok) break;
                lines += l;
            }
            if (ok && lines * (pt / 72) * 1.22 <= h) return pt;
        }
        return minPt;
    };

    const T = (x, y, w, h, text, o = {}) => {
        const pt = o.fit ? fitPt(text, w, h, o.pt || 12, o.min || 6.5, o.mono, o.bold) : (o.pt || 11);
        return {
            t: 'text', x, y, w, h, text: String(text ?? ''), pt,
            bold: !!o.bold, italic: !!o.italic, mono: !!o.mono,
            align: o.align || 'left', valign: o.valign || 'middle',
            color: o.color || NAVY, fill: o.fill || null
        };
    };
    const R = (x, y, w, h, o = {}) => ({
        t: 'rect', x, y, w, h, fill: o.fill || null,
        line: o.line === undefined ? NAVY : o.line,
        lineW: o.lineW || 1, radius: o.radius || 0, dash: !!o.dash
    });
    const L = (x, y, w, o = {}) => ({
        t: 'line', x, y, w, color: o.color || NAVY, lineW: o.lineW || 1, dash: !!o.dash
    });
    const V = (x, y, h, o = {}) => ({
        t: 'vline', x, y, h, color: o.color || NAVY, lineW: o.lineW || 1, dash: !!o.dash
    });
    const I = (x, y, w, h, cardIdx, imageUrl, attribution) => ({
        t: 'img', x, y, w, h, card: cardIdx, imageUrl: imageUrl || null, attribution: attribution || ''
    });

    // Helper to draw a picture clue (with optional subtle licence/source footer) or fallback text clue
    const drawPictureOrTextClue = (els, x, y, w, h, picObj, fallbackText, withPictures, showSourceLine = false) => {
        if (withPictures && picObj && picObj.image) {
            const srcH = showSourceLine && (picObj.provider || picObj.license) ? Math.min(0.16, h * 0.16) : 0;
            els.push(I(x, y, w, h - srcH, picObj.cardIndex, picObj.image, picObj.attribution));
            if (srcH > 0) {
                const srcLabel = [picObj.provider, picObj.license].filter(Boolean).join(' · ');
                els.push(T(x, y + h - srcH, w, srcH, srcLabel, { pt: 6, align: 'center', color: GREY, fit: true, min: 4.5 }));
            }
        } else {
            els.push(T(x, y, w, h, fallbackText || '—', { pt: 10, bold: false, italic: true, align: 'center', color: '2B304D', fit: true, min: 7 }));
        }
    };

    // Draw individual letter boxes for a word (`[ ] [ ] [ ]`)
    const drawLetterBoxesRow = (els, x, y, maxW, rowH, targetWord, isAnswer) => {
        const chars = Array.from(targetWord || '');
        if (chars.length === 0) return;
        const gap = 0.035;
        const maxBoxS = Math.min(rowH * 0.78, 0.28);
        const boxS = Math.min(maxBoxS, (maxW - (chars.length - 1) * gap) / chars.length);
        let curX = x;
        const boxY = y + (rowH - boxS) / 2;

        chars.forEach(ch => {
            if (ch === ' ') {
                curX += boxS * 0.65;
            } else if (!isLetter(ch)) {
                els.push(T(curX, boxY, boxS * 0.6, boxS, ch, { pt: 10, bold: true, align: 'center' }));
                curX += boxS * 0.6 + gap;
            } else {
                els.push(R(curX, boxY, boxS, boxS, { fill: isAnswer ? 'EEF8F6' : 'FFFFFF', line: NAVY, lineW: 1, radius: 0.02 }));
                if (isAnswer) {
                    els.push(T(curX, boxY, boxS, boxS, ch, { pt: Math.min(11, boxS * 44), bold: true, align: 'center', color: TEAL }));
                }
                curX += boxS + gap;
            }
        });
    };

    // ---------- Task Block Builders for Worksheet & Slides ----------
    // Each task builder returns { blocks: [{ h, draw(page, x, y, availW) }], warnings: [] }
    // A block is never split across pages unless the task itself has multiple blocks.
    const buildTaskBlocks = (taskId, sectionNum, ctx, isAnswer) => {
        const warnings = [];
        const settings = (ctx.taskSettings && ctx.taskSettings[taskId]) || DEFAULT_TASK_SETTINGS[taskId] || {};
        const rnd = mulberry32(((ctx.seed || 1) * 10007) + sectionNum * 313 + taskId.charCodeAt(0) * 17 + (taskId.charCodeAt(taskId.length - 1) || 0));

        const meta = TASK_CATALOGUE.find(t => t.id === taskId);
        const langRubricObj = (I18N[ctx.lang] && I18N[ctx.lang].rubrics && I18N[ctx.lang].rubrics[taskId]) || null;
        const showTargetRubric = ctx.showTargetRubric !== false && langRubricObj;

        let enRubric = meta ? meta.enRubric : '';
        if (taskId === 'translate_to_target') {
            const targetLangLabel = ctx.lang === 'other' ? (ctx.customLangName || 'the target language') : (I18N[ctx.lang]?.langName || 'Spanish');
            enRubric = `Translate each sentence into ${targetLangLabel}.`;
        }

        const headingTitle = showTargetRubric
            ? `${sectionNum}. ${langRubricObj.title} — ${meta.name}`
            : `${sectionNum}. ${meta ? meta.name : taskId}`;
        const instructionLine = showTargetRubric
            ? `${langRubricObj.rubric} (${enRubric})`
            : enRubric;

        const headerH = 0.52;
        const drawSectionHeader = (els, x, y, w, partLabel = '') => {
            const fullHead = partLabel ? `${headingTitle} (${partLabel})` : headingTitle;
            els.push(R(x, y, w, 0.27, { fill: PAPER, line: NAVY, lineW: 1, radius: 0.04 }));
            els.push(T(x + 0.1, y, w - 0.2, 0.27, fullHead, { pt: 11.5, bold: true, fit: true, min: 8.5 }));
            els.push(T(x + 0.04, y + 0.28, w - 0.08, 0.2, instructionLine, { pt: 9, italic: true, color: GREY, fit: true, min: 7 }));
        };

        const textOpts = {
            removeArticles: ctx.removeArticles !== false,
            removeAccents: !!ctx.removeAccents,
            gridSize: settings.gridSize || 12,
            difficulty: settings.difficulty || 'easy',
        };

        // ---- A1: WORD SEARCH ----
        if (taskId === 'wordsearch') {
            const wsData = generateWordSearch(ctx.vocab, textOpts, rnd);
            wsData.unplaced.forEach(u => warnings.push(`Word search: ${u.reason}`));

            const clueItems = wsData.placed;
            const warnH = wsData.unplaced.length > 0 ? 0.26 : 0;
            const gridSide = Math.min(3.55, Math.max(2.7, wsData.size * 0.21));
            const bodyH = Math.max(gridSide + 0.1, Math.ceil(clueItems.length / 2) * 0.42 + 0.2);
            const totalH = headerH + bodyH + warnH + 0.1;

            return {
                warnings,
                blocks: [{
                    h: totalH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;
                        const side = Math.min(gridSide, w * 0.5);
                        const cellS = side / wsData.size;
                        const pt = Math.max(6.5, Math.min(12.5, cellS * 72 * 0.54));

                        // Highlight placed words on Answer Key
                        if (isAnswer) {
                            wsData.placed.forEach(p => {
                                p.cells.forEach(([r, c]) => {
                                    els.push(R(x + c * cellS, cy + r * cellS, cellS, cellS, { fill: MUSTARD_LIGHT, line: null }));
                                });
                            });
                        }

                        // Draw cells & letters
                        for (let r = 0; r < wsData.size; r++) {
                            for (let c = 0; c < wsData.size; c++) {
                                els.push(R(x + c * cellS, cy + r * cellS, cellS, cellS, { line: BORDER_GREY, lineW: 0.5 }));
                                els.push(T(x + c * cellS, cy + r * cellS, cellS, cellS, wsData.grid[r][c], {
                                    pt, mono: true, bold: true, align: 'center'
                                }));
                            }
                        }
                        els.push(R(x, cy, side, side, { line: NAVY, lineW: 1.5 }));

                        // Clues on the right
                        const clueX = x + side + 0.25;
                        const clueW = w - side - 0.25;
                        const cols = clueItems.length > 6 ? 2 : 1;
                        const rows = Math.max(1, Math.ceil(clueItems.length / cols));
                        const colW = (clueW - (cols - 1) * 0.12) / cols;
                        const rowH = Math.min(ctx.withPictures ? 0.62 : 0.34, (bodyH - 0.05) / rows);

                        clueItems.forEach((p, i) => {
                            const c = i % cols;
                            const r = Math.floor(i / cols);
                            const ix = clueX + c * (colW + 0.12);
                            const iy = cy + r * rowH;
                            const pic = ctx.getPictureForVocab(p.item);

                            els.push(R(ix, iy + (rowH - 0.16) / 2, 0.16, 0.16, { line: NAVY, lineW: 1, fill: isAnswer ? MUSTARD_LIGHT : 'FFFFFF' }));
                            if (isAnswer) {
                                els.push(T(ix, iy + (rowH - 0.16) / 2, 0.16, 0.16, '✓', { pt: 8, bold: true, align: 'center', color: TEAL }));
                            }

                            if (ctx.withPictures && pic && pic.image) {
                                const imgS = Math.min(rowH - 0.06, 0.52);
                                els.push(I(ix + 0.22, iy + 0.03, imgS, rowH - 0.06, pic.cardIndex, pic.image, pic.attribution));
                                if (isAnswer) {
                                    els.push(T(ix + 0.26 + imgS, iy, colW - 0.26 - imgS, rowH, p.gridWord, { pt: 8.5, bold: true, color: TEAL, fit: true, min: 6 }));
                                }
                            } else {
                                const clueLabel = p.item.english || p.item.target;
                                const textToShow = isAnswer ? `${clueLabel} → ${p.gridWord}` : clueLabel;
                                els.push(T(ix + 0.22, iy, colW - 0.24, rowH, textToShow, {
                                    pt: 9.5, bold: isAnswer, color: isAnswer ? TEAL : NAVY, fit: true, min: 6.5
                                }));
                            }
                        });

                        if (wsData.unplaced.length > 0) {
                            const warnText = `⚠ Unplaced (${wsData.size}×${wsData.size} grid): ${wsData.unplaced.map(u => u.item.target).join(', ')}`;
                            els.push(R(x, cy + bodyH, w, 0.24, { fill: MUSTARD_LIGHT, line: NAVY, lineW: 0.75, radius: 0.03 }));
                            els.push(T(x + 0.08, cy + bodyH, w - 0.16, 0.24, warnText, { pt: 8, bold: true, color: NAVY, fit: true, min: 6 }));
                        }
                    }
                }]
            };
        }

        // ---- A2: CROSSWORD ----
        if (taskId === 'crossword') {
            const cwData = generateCrossword(ctx.vocab, textOpts, rnd);
            cwData.unplaced.forEach(u => warnings.push(`Crossword: ${u.reason}`));

            const maxDim = Math.max(1, cwData.rows, cwData.cols);
            const maxGridW = 3.5;
            const maxGridH = 3.5;
            const cellS = Math.min(0.32, maxGridW / Math.max(1, cwData.cols), maxGridH / Math.max(1, cwData.rows));
            const gridW = cwData.cols * cellS;
            const gridH = cwData.rows * cellS;
            const clueLinesCount = cwData.across.length + cwData.down.length + 2;
            const warnH = cwData.unplaced.length > 0 ? 0.26 : 0;
            const bodyH = Math.max(gridH + 0.15, Math.ceil(clueLinesCount * (ctx.withPictures ? 0.34 : 0.22)), 2.2);
            const totalH = headerH + bodyH + warnH + 0.1;

            return {
                warnings,
                blocks: [{
                    h: totalH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;

                        // Draw crossword cells
                        for (let r = 0; r < cwData.rows; r++) {
                            for (let c = 0; c < cwData.cols; c++) {
                                const cell = cwData.cellGrid[r][c];
                                if (!cell) continue;
                                const cx = x + c * cellS;
                                const ry = cy + r * cellS;
                                els.push(R(cx, ry, cellS, cellS, { fill: isAnswer ? 'EEF8F6' : 'FFFFFF', line: NAVY, lineW: 1 }));
                                if (cell.num) {
                                    els.push(T(cx + 0.015, ry + 0.01, cellS * 0.55, cellS * 0.38, String(cell.num), {
                                        pt: Math.max(5, Math.min(7, cellS * 22)), bold: true, align: 'left', valign: 'top', color: GREY
                                    }));
                                }
                                if (isAnswer) {
                                    els.push(T(cx, ry + cellS * 0.1, cellS, cellS * 0.88, cell.letter, {
                                        pt: Math.max(7, Math.min(12, cellS * 38)), mono: true, bold: true, align: 'center', color: TEAL
                                    }));
                                }
                            }
                        }

                        // Draw Across & Down clues on the right
                        const clueX = x + Math.max(gridW, 2.2) + 0.28;
                        const clueW = w - (clueX - x);
                        const colW = (clueW - 0.16) / 2;

                        const drawClueColumn = (title, list, colX) => {
                            els.push(T(colX, cy, colW, 0.22, title, { pt: 10, bold: true, color: NAVY }));
                            let curY = cy + 0.24;
                            const itemH = list.length > 0 ? Math.min(ctx.withPictures ? 0.44 : 0.24, (bodyH - 0.28) / Math.max(1, list.length)) : 0.24;
                            list.forEach(p => {
                                const pic = ctx.getPictureForVocab(p.item);
                                if (ctx.withPictures && pic && pic.image) {
                                    els.push(T(colX, curY, 0.26, itemH, `${p.clueNum}.`, { pt: 9, bold: true }));
                                    const imgW = Math.min(0.42, itemH - 0.04);
                                    els.push(I(colX + 0.26, curY + 0.02, imgW, itemH - 0.04, pic.cardIndex, pic.image, pic.attribution));
                                    if (isAnswer) {
                                        els.push(T(colX + 0.3 + imgW, curY, colW - 0.3 - imgW, itemH, p.gridWord, { pt: 8, bold: true, color: TEAL, fit: true, min: 6 }));
                                    }
                                } else {
                                    const clueStr = p.item.english || p.item.target;
                                    const label = isAnswer ? `${p.clueNum}. ${clueStr} (${p.gridWord})` : `${p.clueNum}. ${clueStr}`;
                                    els.push(T(colX, curY, colW, itemH, label, { pt: 8.5, bold: isAnswer, color: isAnswer ? TEAL : NAVY, fit: true, min: 6 }));
                                }
                                curY += itemH;
                            });
                        };

                        drawClueColumn('Across →', cwData.across, clueX);
                        drawClueColumn('Down ↓', cwData.down, clueX + colW + 0.16);

                        if (cwData.unplaced.length > 0) {
                            const warnText = `⚠ Unplaced words (no crossing): ${cwData.unplaced.map(u => u.item.target).join(', ')}`;
                            els.push(R(x, cy + bodyH, w, 0.24, { fill: MUSTARD_LIGHT, line: NAVY, lineW: 0.75, radius: 0.03 }));
                            els.push(T(x + 0.08, cy + bodyH, w - 0.16, 0.24, warnText, { pt: 8, bold: true, color: NAVY, fit: true, min: 6 }));
                        }
                    }
                }]
            };
        }

        // ---- A3: MATCHING (Article is ALWAYS kept!) ----
        if (taskId === 'matching') {
            const items = ctx.vocab;
            const k = items.length;
            let pairMode = settings.pairMode || 'auto';
            if (pairMode === 'auto') {
                pairMode = ctx.withPictures ? 'picture-target' : 'english-target';
            }
            if (!ctx.withPictures && (pairMode === 'picture-target' || pairMode === 'picture-english')) {
                pairMode = 'english-target';
            }

            const rightValues = items.map(it => (pairMode === 'picture-english' ? (it.english || it.target) : it.target));
            const rightOrder = derange(k, rnd, rightValues); // rightOrder[pos] = index of item shown at letter pos
            const letters = 'abcdefghijklmnopqrstuvwxyz';

            const leftHasPic = ctx.withPictures && (pairMode === 'picture-target' || pairMode === 'picture-english');
            const rowH = leftHasPic ? 0.52 : 0.31;
            const bodyH = k * rowH + 0.06;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;
                        const leftW = w * 0.44;
                        const rightX = x + w * 0.52;
                        const rightW = w * 0.48;

                        items.forEach((item, i) => {
                            const ry = cy + i * rowH;
                            const ansPos = rightOrder.indexOf(i);
                            const ansLetter = letters[ansPos] || String(ansPos + 1);

                            // Left column: number + picture or English + answer box
                            els.push(T(x, ry, 0.32, rowH, `${i + 1}.`, { pt: 10, bold: true }));
                            if (leftHasPic) {
                                const pic = ctx.getPictureForVocab(item);
                                drawPictureOrTextClue(els, x + 0.34, ry + 0.03, 0.65, rowH - 0.06, pic, item.english || item.target, true, false);
                            } else {
                                els.push(T(x + 0.34, ry, leftW - 0.72, rowH, item.english || item.target, { pt: 10, fit: true, min: 7 }));
                            }

                            // Answer box [ ]
                            const boxS = 0.24;
                            const bx = x + leftW - boxS;
                            const by = ry + (rowH - boxS) / 2;
                            els.push(R(bx, by, boxS, boxS, { fill: isAnswer ? 'EEF8F6' : 'FFFFFF', line: NAVY, lineW: 1.2, radius: 0.03 }));
                            if (isAnswer) {
                                els.push(T(bx, by, boxS, boxS, ansLetter, { pt: 10, bold: true, align: 'center', color: TEAL }));
                            }

                            // Right column: letter + target word (article ALWAYS kept!)
                            const rightItem = items[rightOrder[i]];
                            const rightText = pairMode === 'picture-english' ? (rightItem.english || rightItem.target) : rightItem.target;
                            els.push(T(rightX, ry, 0.32, rowH, `${letters[i] || (i + 1)})`, { pt: 10, bold: true, color: GREY }));
                            els.push(T(rightX + 0.32, ry, rightW - 0.32, rowH, rightText, { pt: 10, bold: false, fit: true, min: 7 }));
                        });
                    }
                }]
            };
        }

        // ---- A4, A5, A6, A7, A9: WORD PUZZLE ROWS/GRIDS ----
        if (['fill_vowels', 'fill_consonants', 'first_last', 'anagrams', 'letter_boxes'].includes(taskId)) {
            const items = ctx.vocab;
            const cols = 2;
            const rows = Math.ceil(items.length / cols);
            const rowH = ctx.withPictures ? 0.56 : 0.36;
            const bodyH = rows * rowH + 0.06;

            // Precompute puzzles deterministically
            const puzzles = items.map(item => {
                const targetClean = getPuzzleTargetWord(item.target, textOpts);
                let promptStr = '';
                if (taskId === 'fill_vowels') promptStr = maskVowels(targetClean);
                else if (taskId === 'fill_consonants') promptStr = maskConsonants(targetClean);
                else if (taskId === 'first_last') promptStr = maskFirstLast(targetClean);
                else if (taskId === 'anagrams') promptStr = anagramWord(targetClean, rnd);
                return { item, targetClean, promptStr };
            });

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;
                        const colW = (w - 0.28) / cols;

                        puzzles.forEach((pz, i) => {
                            const c = i % cols;
                            const r = Math.floor(i / cols);
                            const ix = x + c * (colW + 0.28);
                            const iy = cy + r * rowH;
                            const pic = ctx.getPictureForVocab(pz.item);

                            els.push(T(ix, iy, 0.28, rowH, `${i + 1}.`, { pt: 9.5, bold: true }));
                            const clueW = ctx.withPictures ? 0.56 : colW * 0.38;
                            drawPictureOrTextClue(els, ix + 0.28, iy + 0.03, clueW, rowH - 0.06, pic, pz.item.english || '—', ctx.withPictures, false);

                            const workX = ix + 0.32 + clueW;
                            const workW = colW - 0.32 - clueW;

                            if (taskId === 'letter_boxes') {
                                drawLetterBoxesRow(els, workX, iy, workW, rowH, pz.targetClean, isAnswer);
                            } else if (taskId === 'anagrams') {
                                els.push(T(workX, iy + 0.02, workW, rowH * 0.48, pz.promptStr, {
                                    pt: 10, mono: true, bold: true, color: GREY, fit: true, min: 6.5
                                }));
                                if (isAnswer) {
                                    els.push(T(workX, iy + rowH * 0.48, workW, rowH * 0.48, pz.targetClean, {
                                        pt: 10.5, bold: true, color: TEAL, fit: true, min: 7
                                    }));
                                } else {
                                    els.push(L(workX, iy + rowH - 0.08, workW - 0.05, { lineW: 0.9, color: NAVY }));
                                }
                            } else {
                                // fill_vowels, fill_consonants, first_last
                                const shown = isAnswer ? `${pz.promptStr}  →  ${pz.targetClean}` : pz.promptStr;
                                els.push(T(workX, iy, workW, rowH, shown, {
                                    pt: isAnswer ? 9.5 : 11, mono: !isAnswer, bold: true, color: isAnswer ? TEAL : NAVY, fit: true, min: 6.5
                                }));
                            }
                        });
                    }
                }]
            };
        }

        // ---- A8: LABEL THE PICTURE (Pictorio cards with source/licence line + blank or scaffolded labels) ----
        if (taskId === 'label_picture') {
            const items = ctx.vocab;
            const mode = settings.mode || 'blank'; // 'blank' | 'scaffold'
            const cols = items.length <= 4 ? 4 : items.length <= 6 ? 3 : 4;
            const rows = Math.ceil(items.length / cols);
            const cardH = 1.32;
            const gap = 0.14;
            const bodyH = rows * cardH + (rows - 1) * gap + 0.06;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;
                        const cardW = (w - (cols - 1) * gap) / cols;

                        items.forEach((item, i) => {
                            const c = i % cols;
                            const r = Math.floor(i / cols);
                            const cx = x + c * (cardW + gap);
                            const ry = cy + r * (cardH + gap);
                            const pic = ctx.getPictureForVocab(item);
                            const targetClean = getPuzzleTargetWord(item.target, textOpts);

                            els.push(R(cx, ry, cardW, cardH, { fill: 'FFFFFF', line: NAVY, lineW: 1, radius: 0.06 }));
                            els.push(T(cx + 0.05, ry + 0.03, 0.3, 0.18, `${i + 1}.`, { pt: 8.5, bold: true, color: GREY }));

                            // Source / licence line at top-right of card
                            if (pic && (pic.provider || pic.license)) {
                                const srcText = [pic.provider, pic.license].filter(Boolean).join(' · ');
                                els.push(T(cx + 0.3, ry + 0.03, cardW - 0.35, 0.16, srcText, {
                                    pt: 5.5, align: 'right', color: GREY, fit: true, min: 4.5
                                }));
                            }

                            if (pic && pic.image) {
                                els.push(I(cx + 0.08, ry + 0.2, cardW - 0.16, cardH * 0.56, pic.cardIndex, pic.image, pic.attribution));
                            }

                            const labelY = ry + cardH * 0.78;
                            if (isAnswer) {
                                els.push(T(cx + 0.06, labelY, cardW - 0.12, cardH * 0.2, targetClean, {
                                    pt: 10.5, bold: true, align: 'center', color: TEAL, fit: true, min: 7
                                }));
                            } else if (mode === 'scaffold') {
                                const scaffoldStr = maskFirstLast(targetClean);
                                els.push(T(cx + 0.06, labelY, cardW - 0.12, cardH * 0.2, scaffoldStr, {
                                    pt: 10, mono: true, bold: true, align: 'center', fit: true, min: 6.5
                                }));
                            } else {
                                els.push(L(cx + 0.14, ry + cardH - 0.14, cardW - 0.28, { lineW: 1, color: NAVY }));
                            }
                        });
                    }
                }]
            };
        }

        // ---- B1 & B2: TRANSLATE SENTENCES ----
        if (taskId === 'translate_to_target' || taskId === 'translate_to_english') {
            const intoTarget = taskId === 'translate_to_target';
            const sents = ctx.sentences;
            const rowH = 0.52;
            const bodyH = sents.length * rowH + 0.05;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;

                        sents.forEach((s, i) => {
                            const ry = cy + i * rowH;
                            const promptText = intoTarget ? (s.english || s.target) : s.target;
                            const answerText = intoTarget ? s.target : (s.english || '—');

                            els.push(T(x, ry, w, 0.23, `${i + 1}.  ${promptText}`, { pt: 10, bold: false, fit: true, min: 7.5 }));
                            if (isAnswer) {
                                els.push(T(x + 0.25, ry + 0.23, w - 0.25, 0.24, answerText, { pt: 10, bold: true, color: TEAL, fit: true, min: 7.5 }));
                            } else {
                                els.push(L(x + 0.25, ry + 0.44, w - 0.25, { lineW: 0.85, color: GREY }));
                            }
                        });
                    }
                }]
            };
        }

        // ---- B3: JUMBLED SENTENCES ----
        if (taskId === 'jumbled_sentences') {
            const sents = ctx.sentences;
            const lowercaseFirst = settings.lowercaseFirst !== false;
            const showEng = !!settings.showEnglish;
            const rowH = showEng ? 0.64 : 0.52;
            const bodyH = sents.length * rowH + 0.05;

            const jumbledList = sents.map(s => ({
                sent: s,
                jumbled: jumbleSentence(s.target, rnd, lowercaseFirst),
            }));

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;

                        jumbledList.forEach((item, i) => {
                            const ry = cy + i * rowH;
                            els.push(T(x, ry, w, 0.22, `${i + 1}.  ${item.jumbled}`, { pt: 10, bold: true, fit: true, min: 7 }));
                            if (showEng && item.sent.english) {
                                els.push(T(x + 0.25, ry + 0.2, w - 0.25, 0.16, `(${item.sent.english})`, { pt: 8.5, italic: true, color: GREY, fit: true, min: 6.5 }));
                            }
                            const lineY = showEng ? ry + 0.36 : ry + 0.24;
                            if (isAnswer) {
                                els.push(T(x + 0.25, lineY, w - 0.25, 0.24, item.sent.target, { pt: 10, bold: true, color: TEAL, fit: true, min: 7.5 }));
                            } else {
                                els.push(L(x + 0.25, ry + rowH - 0.07, w - 0.25, { lineW: 0.85, color: GREY }));
                            }
                        });
                    }
                }]
            };
        }

        // ---- B4: SENTENCE GAP FILL ----
        if (taskId === 'sentence_gapfill') {
            const sents = ctx.sentences;
            const showBank = settings.wordBank !== false;
            const showFirst = !!settings.showFirstLetter;

            const resolved = sents.map(s => resolveSentenceGap(s, ctx.vocab));
            const bankWords = shuffle(resolved.map(r => r.answer), rnd);
            // Ensure bank order differs from 1..N if >= 2 distinct words
            if (bankWords.length >= 2 && bankWords.every((w, i) => w === resolved[i].answer)) {
                [bankWords[0], bankWords[1]] = [bankWords[1], bankWords[0]];
            }

            const bankH = showBank ? 0.42 : 0;
            const rowH = 0.34;
            const bodyH = bankH + sents.length * rowH + 0.06;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        let cy = y + headerH;

                        if (showBank) {
                            els.push(R(x, cy, w, 0.34, { fill: 'FBF7EE', line: NAVY, lineW: 0.85, radius: 0.05 }));
                            const bankStr = `Word bank:   ${bankWords.join('   ·   ')}`;
                            els.push(T(x + 0.12, cy, w - 0.24, 0.34, bankStr, { pt: 9.5, bold: true, align: 'center', fit: true, min: 7 }));
                            cy += 0.42;
                        }

                        resolved.forEach((r, i) => {
                            const ry = cy + i * rowH;
                            const ans = r.answer || '';
                            const gapStr = isAnswer
                                ? `[ ${ans} ]`
                                : (showFirst && ans.length > 0 ? `${ans[0]}${'_'.repeat(Math.max(6, ans.length + 1))}` : '____________');
                            const fullLine = `${i + 1}.  ${r.template.replace('___GAP___', gapStr)}`;
                            els.push(T(x, ry, w, rowH, fullLine, {
                                pt: 10, bold: isAnswer, color: isAnswer ? TEAL : NAVY, fit: true, min: 7.5
                            }));
                        });
                    }
                }]
            };
        }

        // ---- B5: MATCH THE HALVES & B6: MATCH SENTENCE TO TRANSLATION ----
        if (taskId === 'match_halves' || taskId === 'match_translation') {
            const sents = ctx.sentences;
            const k = sents.length;
            const pairs = sents.map(s => {
                if (taskId === 'match_halves') {
                    const [left, right] = splitSentenceInHalf(s.target);
                    return { left, right };
                }
                return { left: s.target, right: s.english || s.target };
            });

            const rightValues = pairs.map(p => p.right);
            const rightOrder = derange(k, rnd, rightValues);
            const letters = 'abcdefghijklmnopqrstuvwxyz';
            const rowH = 0.36;
            const bodyH = k * rowH + 0.06;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;
                        const leftW = w * 0.46;
                        const rightX = x + w * 0.51;
                        const rightW = w * 0.49;

                        pairs.forEach((pair, i) => {
                            const ry = cy + i * rowH;
                            const ansPos = rightOrder.indexOf(i);
                            const ansLetter = letters[ansPos] || String(ansPos + 1);

                            els.push(T(x, ry, 0.28, rowH, `${i + 1}.`, { pt: 9.5, bold: true }));
                            els.push(T(x + 0.28, ry, leftW - 0.62, rowH, pair.left, { pt: 9.5, fit: true, min: 6.5 }));

                            const boxS = 0.24;
                            const bx = x + leftW - boxS;
                            const by = ry + (rowH - boxS) / 2;
                            els.push(R(bx, by, boxS, boxS, { fill: isAnswer ? 'EEF8F6' : 'FFFFFF', line: NAVY, lineW: 1.2, radius: 0.03 }));
                            if (isAnswer) {
                                els.push(T(bx, by, boxS, boxS, ansLetter, { pt: 9.5, bold: true, align: 'center', color: TEAL }));
                            }

                            const shownRight = pairs[rightOrder[i]].right;
                            els.push(T(rightX, ry, 0.28, rowH, `${letters[i] || (i + 1)})`, { pt: 9.5, bold: true, color: GREY }));
                            els.push(T(rightX + 0.28, ry, rightW - 0.28, rowH, shownRight, { pt: 9.5, fit: true, min: 6.5 }));
                        });
                    }
                }]
            };
        }

        // ---- B7: MISSING VOWELS IN SENTENCES ----
        if (taskId === 'sentence_vowels') {
            const sents = ctx.sentences;
            const rowH = 0.52;
            const bodyH = sents.length * rowH + 0.05;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;

                        sents.forEach((s, i) => {
                            const ry = cy + i * rowH;
                            const masked = maskSentenceVowels(s.target);
                            els.push(T(x, ry, w, 0.24, `${i + 1}.  ${masked}`, { pt: 10, mono: true, bold: true, fit: true, min: 7 }));
                            if (isAnswer) {
                                els.push(T(x + 0.25, ry + 0.24, w - 0.25, 0.24, s.target, { pt: 10, bold: true, color: TEAL, fit: true, min: 7.5 }));
                            } else {
                                els.push(L(x + 0.25, ry + 0.45, w - 0.25, { lineW: 0.85, color: GREY }));
                            }
                        });
                    }
                }]
            };
        }

        // ---- B8: PICTURE PROMPTS (Needs pictures on and *keyword* markup) ----
        if (taskId === 'picture_prompts') {
            const sents = ctx.sentences.filter(s => s.pictureKeyword && s.pictureKeyword.trim().length > 0);
            const rowH = 0.58;
            const bodyH = sents.length * rowH + 0.06;

            return {
                warnings,
                blocks: [{
                    h: headerH + bodyH,
                    draw: (page, x, y, w) => {
                        const els = page.els;
                        drawSectionHeader(els, x, y, w);
                        const cy = y + headerH;

                        sents.forEach((s, i) => {
                            const ry = cy + i * rowH;
                            const kw = s.pictureKeyword;
                            const gappedTemplate = s.rawTarget
                                .replace(/[\[\]]/g, '')
                                .replace(/\*([^*[\]]+)\*/, isAnswer ? `[ ${kw} ]` : '____________')
                                .replace(/\*/g, '');
                            const pic = ctx.getPictureForKeyword(kw, s.english);

                            els.push(T(x, ry, 0.28, rowH, `${i + 1}.`, { pt: 10, bold: true }));
                            drawPictureOrTextClue(els, x + 0.3, ry + 0.04, 0.62, rowH - 0.08, pic, kw, true, false);
                            els.push(T(x + 1.0, ry, w - 1.0, rowH, gappedTemplate, {
                                pt: 10.5, bold: isAnswer, color: isAnswer ? TEAL : NAVY, fit: true, min: 7.5
                            }));
                        });
                    }
                }]
            };
        }

        return { warnings, blocks: [] };
    };

    // ---------- Build Complete Worksheet & Slides Deck ----------
    const buildWorksheetDeck = (state) => {
        const fmt = state.format === 'slides' ? 'slides' : 'worksheet';
        const ratio = state.slideRatio === '4:3' ? '4:3' : '16:9';
        const size = fmt === 'worksheet' ? SIZES.A4 : SIZES[ratio];
        const W = size.w;
        const H = size.h;

        const lang = I18N[state.lang] ? state.lang : 'other';
        const langDict = I18N[lang];
        const title = (state.title || '').trim() || langDict.defaultTitle;
        const classLine = (state.classLine || '').trim();
        const seed = Number(state.seed) || 1234;
        const withPictures = state.withPictures !== false;

        const vocab = state.vocab || [];
        const sentences = state.sentences || [];
        const cards = state.cards || [];
        const pictureMap = state.pictureMap || {};

        // Lookup picture for a Box A vocabulary item
        const getPictureForVocab = (vocabItem) => {
            if (!withPictures || !vocabItem) return null;
            const key = (vocabItem.target || '').toLowerCase().trim();
            if (pictureMap[key] && pictureMap[key].image) return pictureMap[key];
            const idx = cards.findIndex(c => (c.word || '').toLowerCase().trim() === key || (c.target || '').toLowerCase().trim() === key);
            if (idx >= 0 && cards[idx].image) {
                return { ...cards[idx], cardIndex: idx };
            }
            return null;
        };

        // Lookup picture for a Box B *keyword*
        const getPictureForKeyword = (keyword) => {
            if (!withPictures || !keyword) return null;
            const kwLower = keyword.toLowerCase().trim();
            const kwStripped = stripArticle(kwLower);
            if (pictureMap[kwLower] && pictureMap[kwLower].image) return pictureMap[kwLower];
            if (pictureMap[kwStripped] && pictureMap[kwStripped].image) return pictureMap[kwStripped];
            const idx = cards.findIndex(c => {
                const w = (c.word || c.target || '').toLowerCase().trim();
                return w === kwLower || stripArticle(w) === kwStripped;
            });
            if (idx >= 0 && cards[idx].image) {
                return { ...cards[idx], cardIndex: idx };
            }
            return null;
        };

        // Collect all active pictures for credits footer & credits slide
        const usedPictures = [];
        const seenUrls = new Set();
        const registerPic = (label, p) => {
            if (p && p.image && !seenUrls.has(p.image)) {
                seenUrls.add(p.image);
                usedPictures.push({ word: label, ...p });
            }
        };
        if (withPictures) {
            vocab.forEach(v => registerPic(v.target, getPictureForVocab(v)));
            sentences.forEach(s => {
                if (s.pictureKeyword) registerPic(s.pictureKeyword, getPictureForKeyword(s.pictureKeyword));
            });
            cards.forEach((c, idx) => registerPic(c.word || c.target, { ...c, cardIndex: idx }));
        }

        const hasArasaac = usedPictures.some(c => c.provider === 'ARASAAC' || c.providerKey === 'arasaac');
        const others = Array.from(new Set(usedPictures.map(c => c.provider).filter(p => p && p !== 'ARASAAC' && p !== 'None' && p !== 'Upload')));
        let footerCredit = '';
        if (withPictures) {
            if (hasArasaac) footerCredit = ARASAAC_LINE + (others.length ? ` Other images: ${others.join(', ')}.` : '');
            else if (others.length) footerCredit = `Images: ${others.join(', ')}.`;
        }

        const ctx = {
            fmt, ratio, W, H, lang, L: langDict, title, classLine, seed, withPictures,
            customLangName: state.customLangName || '',
            vocab, sentences, cards,
            removeArticles: state.removeArticles !== false,
            removeAccents: !!state.removeAccents,
            showTargetRubric: state.showTargetRubric !== false,
            taskSettings: state.taskSettings || DEFAULT_TASK_SETTINGS,
            getPictureForVocab,
            getPictureForKeyword,
            footerCredit,
            hasArasaac,
            usedPictures,
        };

        // Filter selected tasks to those that are available
        const requestedTasks = Array.isArray(state.selectedTasks) ? state.selectedTasks : ['wordsearch', 'matching', 'sentence_gapfill', 'jumbled_sentences'];
        const activeTasks = [];
        const skippedTasks = [];

        requestedTasks.forEach(tId => {
            const avail = getTaskAvailability(tId, { vocab, sentences, withPictures });
            if (avail.available) activeTasks.push(tId);
            else skippedTasks.push({ id: tId, reason: avail.reason });
        });

        const allWarnings = [];
        const studentTaskBlocks = [];
        const answerTaskBlocks = [];

        activeTasks.forEach((tId, idx) => {
            const sectionNum = idx + 1;
            const st = buildTaskBlocks(tId, sectionNum, ctx, false);
            const an = buildTaskBlocks(tId, sectionNum, ctx, true);
            st.warnings.forEach(w => {
                if (!allWarnings.includes(w)) allWarnings.push(w);
            });
            studentTaskBlocks.push({ taskId: tId, sectionNum, blocks: st.blocks });
            answerTaskBlocks.push({ taskId: tId, sectionNum, blocks: an.blocks });
        });

        // Paginate A4 Worksheet pages so a task is never split across pages unless taller than 1 page
        const paginateA4 = (taskGroups, isAnswerKey) => {
            const pages = [];
            const marginX = 0.42;
            const contentW = W - marginX * 2;
            const bottomLimit = H - 0.52;

            let curPage = null;
            let curY = 0;

            const startPage = (isFirst) => {
                const pageNum = pages.length + 1;
                const label = isAnswerKey
                    ? `Answer key (Page ${pageNum})`
                    : `Student worksheet (Page ${pageNum})`;
                curPage = { w: W, h: H, bg: 'FFFFFF', els: [], label, isAnswerKey };
                const e = curPage.els;

                if (isFirst) {
                    e.push(T(marginX, 0.28, contentW * 0.56, 0.34, isAnswerKey ? `${title} — ANSWER KEY` : title, {
                        pt: 15, bold: true, color: isAnswerKey ? TEAL : NAVY, fit: true, min: 10
                    }));
                    if (isAnswerKey) {
                        e.push(T(marginX + contentW * 0.56, 0.28, contentW * 0.44, 0.34, `Version #${seed}`, {
                            pt: 10, bold: true, align: 'right', color: TEAL
                        }));
                    } else {
                        e.push(T(marginX + contentW * 0.56, 0.28, contentW * 0.44, 0.34, 'Name: __________________   Date: __________', {
                            pt: 9.5, align: 'right', color: NAVY
                        }));
                    }
                    let lineY = 0.66;
                    if (classLine) {
                        e.push(T(marginX, 0.62, contentW * 0.7, 0.22, classLine, { pt: 9.5, italic: true, color: GREY }));
                        if (!isAnswerKey) {
                            e.push(T(marginX + contentW * 0.7, 0.62, contentW * 0.3, 0.22, `Version #${seed}`, { pt: 8.5, align: 'right', color: GREY }));
                        }
                        lineY = 0.86;
                    } else if (!isAnswerKey) {
                        e.push(T(marginX, 0.62, contentW, 0.18, `Version #${seed}`, { pt: 8, align: 'right', color: GREY }));
                        lineY = 0.82;
                    }
                    e.push(L(marginX, lineY, contentW, { lineW: 1.5, color: NAVY }));
                    curY = lineY + 0.14;
                } else {
                    e.push(T(marginX, 0.26, contentW * 0.7, 0.26, isAnswerKey ? `${title} — Answer key (p. ${pageNum})` : `${title} (p. ${pageNum})`, {
                        pt: 10.5, bold: true, color: GREY, fit: true, min: 8
                    }));
                    e.push(T(marginX + contentW * 0.7, 0.26, contentW * 0.3, 0.26, `Version #${seed}`, {
                        pt: 8.5, align: 'right', color: GREY
                    }));
                    e.push(L(marginX, 0.55, contentW, { lineW: 1, color: BORDER_GREY }));
                    curY = 0.68;
                }

                // Footer credit & version
                const footerStr = [footerCredit, `Pictorio Worksheets · Version #${seed}`].filter(Boolean).join('  |  ');
                e.push(T(marginX, H - 0.38, contentW, 0.26, footerStr, { pt: 6.5, align: 'center', color: GREY, fit: true, min: 5 }));
                pages.push(curPage);
            };

            startPage(true);

            if (taskGroups.length === 0) {
                curPage.els.push(T(marginX, curY + 0.5, contentW, 0.5, 'Tick at least one task on the left to build your worksheet.', {
                    pt: 12, italic: true, align: 'center', color: GREY
                }));
                return pages;
            }

            taskGroups.forEach(tg => {
                tg.blocks.forEach(block => {
                    const availOnCurrent = bottomLimit - curY;
                    if (block.h > availOnCurrent && curY > 1.05) {
                        startPage(false);
                    }
                    block.draw(curPage, marginX, curY, contentW);
                    curY += block.h + 0.22;
                });
            });

            return pages;
        };

        // Slides format: one slide per task, then the answer slides, then the credits slide (Section 8)
        const buildSlidesPages = () => {
            const slidePages = [];
            const marginX = 0.45;
            const contentW = W - marginX * 2;

            // 1. One slide per task (Student)
            studentTaskBlocks.forEach(tg => {
                const meta = TASK_CATALOGUE.find(t => t.id === tg.taskId);
                tg.blocks.forEach((block) => {
                    const p = { w: W, h: H, bg: 'FFFFFF', els: [], label: `${tg.sectionNum}. ${meta ? meta.name : tg.taskId}`, isAnswerKey: false };
                    p.els.push(T(marginX, 0.2, contentW * 0.7, 0.28, title, { pt: 11, bold: true, color: GREY }));
                    p.els.push(T(marginX + contentW * 0.7, 0.2, contentW * 0.3, 0.28, `Version #${seed}`, { pt: 10, align: 'right', color: GREY }));
                    block.draw(p, marginX, 0.52, contentW);
                    slidePages.push(p);
                });
            });

            // 2. Answer slides
            answerTaskBlocks.forEach(tg => {
                const meta = TASK_CATALOGUE.find(t => t.id === tg.taskId);
                tg.blocks.forEach((block) => {
                    const p = { w: W, h: H, bg: 'FFFFFF', els: [], label: `Answer: ${tg.sectionNum}. ${meta ? meta.name : tg.taskId}`, isAnswerKey: true };
                    p.els.push(T(marginX, 0.2, contentW * 0.7, 0.28, `${title} — ANSWER KEY`, { pt: 11, bold: true, color: TEAL }));
                    p.els.push(T(marginX + contentW * 0.7, 0.2, contentW * 0.3, 0.28, `Version #${seed}`, { pt: 10, align: 'right', color: TEAL }));
                    block.draw(p, marginX, 0.52, contentW);
                    slidePages.push(p);
                });
            });

            // 3. Credits slide
            const credited = usedPictures.filter(c => c.attribution || c.license || c.author || c.provider);
            if (credited.length > 0) {
                const lines = credited.map(c => `${c.word}: ${[c.provider, c.author, c.attribution || c.license].filter(Boolean).join(' · ')}`);
                chunk(lines, 12).forEach((group, g, all) => {
                    const p = { w: W, h: H, bg: PAPER, els: [], label: `Image credits${all.length > 1 ? ` (${g + 1}/${all.length})` : ''}`, isAnswerKey: false };
                    p.els.push(T(marginX, 0.3, contentW, 0.4, 'Image credits & licences', { pt: 20, bold: true, color: NAVY }));
                    if (hasArasaac) {
                        p.els.push(T(marginX, 0.72, contentW, 0.28, ARASAAC_LINE, { pt: 9.5, italic: true, color: GREY, fit: true, min: 7 }));
                    }
                    const startY = hasArasaac ? 1.08 : 0.8;
                    const lh = Math.min(0.32, (H - startY - 0.35) / Math.max(1, group.length));
                    group.forEach((line, i) => {
                        p.els.push(T(marginX, startY + i * lh, contentW, lh, line, { pt: 9.5, fit: true, min: 6.5 }));
                    });
                    slidePages.push(p);
                });
            }

            return slidePages;
        };

        if (fmt === 'slides') {
            const pages = buildSlidesPages();
            return {
                pages,
                studentPages: pages.filter(p => !p.isAnswerKey),
                answerPages: pages.filter(p => p.isAnswerKey),
                size,
                warnings: allWarnings,
                skippedTasks,
            };
        }

        const studentPages = paginateA4(studentTaskBlocks, false);
        const answerPages = paginateA4(answerTaskBlocks, true);

        // Credits page for PowerPoint export
        const creditsSlidePages = [];
        const credited = usedPictures.filter(c => c.attribution || c.license || c.author || c.provider);
        if (credited.length > 0) {
            const lines = credited.map(c => `${c.word}: ${[c.provider, c.author, c.attribution || c.license].filter(Boolean).join(' · ')}`);
            chunk(lines, 24).forEach((group, g, all) => {
                const p = { w: W, h: H, bg: PAPER, els: [], label: `Image credits${all.length > 1 ? ` (${g + 1}/${all.length})` : ''}`, isCredits: true };
                p.els.push(T(0.42, 0.35, W - 0.84, 0.38, 'Image credits & licences', { pt: 16, bold: true, color: NAVY }));
                if (hasArasaac) {
                    p.els.push(T(0.42, 0.76, W - 0.84, 0.3, ARASAAC_LINE, { pt: 8.5, italic: true, color: GREY, fit: true, min: 6.5 }));
                }
                const startY = hasArasaac ? 1.12 : 0.82;
                const lh = Math.min(0.32, (H - startY - 0.45) / Math.max(1, group.length));
                group.forEach((line, i) => {
                    p.els.push(T(0.42, startY + i * lh, W - 0.84, lh, line, { pt: 8.5, fit: true, min: 6 }));
                });
                creditsSlidePages.push(p);
            });
        }

        return {
            pages: [...studentPages, ...answerPages, ...creditsSlidePages],
            studentPages,
            answerPages,
            creditsPages: creditsSlidePages,
            size,
            warnings: allWarnings,
            skippedTasks,
        };
    };

    // ---------- PowerPoint (.pptx) Renderer ----------
    const containRect = (el, img) => {
        if (!img || !img.w || !img.h) return { x: el.x, y: el.y, w: el.w, h: el.h };
        const s = Math.min(el.w / img.w, el.h / img.h);
        const w = img.w * s;
        const h = img.h * s;
        return { x: el.x + (el.w - w) / 2, y: el.y + (el.h - h) / 2, w, h };
    };

    // imagesByUrl: { [imageUrl]: { data, w, h } } and imagesByIndex: { [cardIdx]: { data, w, h } }
    const toPptx = (deck, imagesLookup, PptxGenJS, layoutName) => {
        const pptx = new PptxGenJS();
        pptx.defineLayout({ name: layoutName, width: deck.size.w, height: deck.size.h });
        pptx.layout = layoutName;
        deck.pages.forEach(page => {
            const slide = pptx.addSlide();
            slide.background = { color: page.bg || 'FFFFFF' };
            page.els.forEach(el => {
                if (el.t === 'text') {
                    slide.addText(el.text, {
                        x: el.x, y: el.y, w: el.w, h: el.h, margin: 0,
                        fontFace: el.mono ? 'Courier New' : 'Arial', fontSize: el.pt, bold: el.bold, italic: el.italic,
                        color: el.color, align: el.align, valign: el.valign, wrap: true,
                        ...(el.fill ? { fill: { color: el.fill } } : {}),
                    });
                } else if (el.t === 'rect') {
                    const opt = { x: el.x, y: el.y, w: el.w, h: el.h };
                    opt.fill = el.fill ? { color: el.fill } : { type: 'none' };
                    opt.line = el.line ? { color: el.line, width: el.lineW, ...(el.dash ? { dashType: 'dash' } : {}) } : { type: 'none' };
                    if (el.radius) opt.rectRadius = Math.min(el.radius / Math.min(el.w, el.h), 0.5);
                    slide.addShape(el.radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, opt);
                } else if (el.t === 'line' || el.t === 'vline') {
                    slide.addShape(pptx.ShapeType.line, {
                        x: el.x, y: el.y, w: el.t === 'line' ? el.w : 0, h: el.t === 'vline' ? el.h : 0,
                        line: { color: el.color, width: el.lineW, ...(el.dash ? { dashType: 'dash' } : {}) },
                    });
                } else if (el.t === 'img') {
                    const img = (el.imageUrl && imagesLookup[el.imageUrl]) || (el.card !== undefined && imagesLookup[el.card]);
                    if (!img || !img.data) return;
                    slide.addImage({ data: img.data, ...containRect(el, img) });
                }
            });
        });
        return pptx;
    };

    const api = {
        SIZES,
        I18N,
        TASK_CATALOGUE,
        DEFAULT_TASK_SETTINGS,
        mulberry32,
        shuffle,
        derange,
        stripArticle,
        toPuzzleUpper,
        toGridWord,
        isVowel,
        isConsonant,
        maskVowels,
        maskConsonants,
        maskFirstLast,
        anagramWord,
        jumbleSentence,
        parseVocabulary,
        parseSentences,
        getTaskAvailability,
        generateWordSearch,
        generateCrossword,
        resolveSentenceGap,
        splitSentenceInHalf,
        buildWorksheetDeck,
        toPptx,
        containRect,
    };

    if (typeof window !== 'undefined') window.PictorioDeck = api;
    if (typeof module !== 'undefined') module.exports = api;
})();
