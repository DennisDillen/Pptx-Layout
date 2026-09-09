# W&P PPTX Template Renderer

Dieser Dienst erzeugt Marktanalyse-Folien auf Basis der gelieferten `Vorlage.pptx`.
Er baut keine neue Präsentation auf einer leeren Zeichenfläche. Jede Ausgabe beginnt
mit einer Kopie der Vorlage und ersetzt nur definierte Text- und Diagrammfelder.

## Lokaler Test

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Danach sind `http://localhost:10000/health` und
`http://localhost:10000/openapi.json` erreichbar. Ein Beispielrequest liegt in
`sample_request.json`.

## Render Free

Im Render-Dashboard **New → Web Service** wählen, das GitHub-Repository verbinden
und den Branch `main` verwenden. Wenn das Repository nur diesen Dienst enthält,
bleibt das Root Directory leer. Die Einstellungen lauten:

- Runtime: Python
- Build Command: `pip install -r requirements.txt`
- Start Command: `python app.py`
- Plan: Free
- Health Check Path: `/health`
- Environment Variable: `RENDERER_API_KEY` mit **Generate** erzeugen

Nach dem Deploy importiert das CustomGPT die Action aus
`https://<dein-render-name>.onrender.com/openapi.json`. Als Authentifizierung
verwendet die Action **API Key → Bearer** mit dem Wert `RENDERER_API_KEY`.

## Layoutvertrag

`template_manifest.json` enthält die geprüften Shape-IDs, Diagramm-Slots und die
Foliengröße. Die vier Diagramm-Slots behalten jeweils den Diagrammtyp und die
Formatierung der Vorlage. Kategorien und Reihenwerte sind dynamisch; ein Slot kann
bewusst entfernt werden, wenn die Story dort kein Diagramm benötigt. Eine Änderung
der Foliengröße oder der Position vorhandener Shapes führt zu einer `422`-Antwort.

Damit die Vorlage wirklich 1:1 bleibt, darf `Vorlage.pptx` nicht durch eine neu
erstellte PPTX ersetzt werden. Bei einer neuen Vorlagenversion müssen die Shape-IDs
und der Layoutvertrag erneut geprüft werden.

## Sicherheit und Free-Einschränkungen

Der Sugra-Key gehört ausschließlich in den Sugra-Proxy. Der Renderer kennt keinen
Sugra-Key und speichert keine Analyseergebnisse dauerhaft. Render Free kann den
Dienst nach Inaktivität schlafen legen; der erste Aufruf kann deshalb verzögert sein.

