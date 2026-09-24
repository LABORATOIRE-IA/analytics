# Mockup — Nova · Suivi Performance Marketing

Dashboard Flask de suivi de performance marketing, en version **démo autonome** : il fonctionne **uniquement sur des données fictives locales** (marque imaginaire **« Nova »**, mai-juillet 2026).

Les volumes contractualisés sont affichés en **chiffre d'affaires (€)** : chaque dossier, prescription et utilisateur UTM porte un montant tiré entre 100 et 1000 €. Le nombre de dossiers reste accessible en info-bulle, et les coûts unitaires (CPL, CPA, CAC) continuent de se calculer sur les volumes.

## Lancer

```bash
cd "Mockup"
pip install flask openpyxl pandas
python3 app.py            # → http://localhost:5002 (s'ouvre automatiquement)
```

Port **5002**.

## Contenu du dossier

| Élément | Rôle |
|---|---|
| `app.py` | Serveur Flask (API + export Excel/CSV/HTML) |
| `templates/index.html` | Front (onglets, graphiques Chart.js) |
| `static/chart.umd.min.js`, `static/xlsx-js-style.bundle.js` | Librairies JS servies en local (aucun CDN) |
| `data/` | Jeu de données fictives mai-juillet 2026 (voir `SOURCES_DONNEES.md`) |
| `SOURCES_DONNEES.md` | Description de chaque fichier de données (format, colonnes, usage) |

Le dossier est autonome : aucune ressource externe (CDN, API, fichier hors dossier) n'est nécessaire.

## Mode démo

- **Chemins relatifs** : toutes les données sont lues dans `data/`.
- **Horloge figée** au **02/07/2026 22:00** (`DEMO_NOW` côté serveur, `Date` décalé côté navigateur) → S27 = semaine en cours.
- **Semaines affichées** : S19 (04/05) → S27 (29/06).
- **Aucun appel réseau** : les connecteurs d'origine (API d'audience web, Google Sheets, téléchargement du budget, export backoffice) sont désactivés ; les boutons correspondants répondent « désactivé en mode démo ».
