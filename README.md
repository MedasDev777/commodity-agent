# commodity-agent

Claude-Code-Agent, der für Firmen mit Status `Lead` im Dolomite CRM E-Mail-Entwürfe vorbereitet. Er sendet nie selbst, die Entwürfe landen im CRM bei der jeweiligen Person.

**Privates Repo.** Das Handbook enthält Namen von Kontakten und darf nicht öffentlich werden.

## Handbook

| Datei | Rolle |
|---|---|
| `handbook/email-handbook.docx` | **Original.** Nur diese Datei wird bearbeitet. |
| `handbook/email-handbook.md` | Wird aus dem Docx erzeugt und vom Agent gelesen. **Nie von Hand ändern.** |
| `scripts/handbook.py` | Wandelt um und prüft. Nur Python-Standardbibliothek. |

### Neue Handbook-Version

1. Das neue Docx unter `handbook/email-handbook.docx` ablegen (gleicher Name).
2. Umwandeln:

   ```bash
   python scripts/handbook.py convert
   ```

3. Beide Dateien zusammen committen.

`convert` prüft danach automatisch, ob jedes Wort in derselben Reihenfolge übernommen wurde und ob Überschriften, Aufzählungen und Tabellenzeilen übereinstimmen. Bei einer Abweichung bricht es ab und schreibt nichts.

Neue Formatierungen im Docx (Bilder, Links, verschachtelte Listen, verbundene Tabellenzellen usw.) lehnt das Skript bewusst ab, statt sie still zu verlieren. Dann muss der Umwandler erweitert werden.

### Schutz vor veraltetem Handbook

In der ersten Zeile der Markdown-Datei steht der SHA-256-Fingerabdruck des Docx. Der Agent führt vor jedem Lauf Folgendes aus:

```bash
python scripts/handbook.py check
```

Passt der Fingerabdruck nicht, endet das mit Exit-Code 1 und „Handbook nicht aktuell“. Der Agent schreibt dann keine Mails.
