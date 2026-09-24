# Sources de données — Mockup « Nova · Suivi Performance Marketing »

> **Toutes les données de `data/` sont fictives.** Elles décrivent une banque en ligne pour professionnels imaginaire, **Nova**, sur la période du **03/05/2026 au 02/07/2026** (semaines ISO S19 → S27).
> Les volumes, répartitions et délais sont réalistes mais simulés (tirages aléatoires). Les identifiants (clients, sessions, campagnes, agences, agents) sont aléatoires ou pseudonymisés, et les domaines utilisent le suffixe réservé `.example`.

Taille totale : ~90 Mo, dont ~70 Mo pour le cache de trafic web.

### Montant des contrats

Chaque dossier, prescription et utilisateur UTM porte une colonne **`Valeur`** : un montant en euros entre **100,00 €** et **1000,00 €** (moyenne 550 €), utilisé par le dashboard à la place du nombre de contrats.

Le montant est **dérivé par hachage d'une clé stable**, jamais tiré au fil de la lecture :

| Jeu | Clé | Formule |
|---|---|---|
| Souscriptions | `Date de création` (à la seconde) + rang parmi les lignes de même horodatage | `100 + (SHA-256(clé) mod 90 001) / 100` |
| Prescriptions | `Code_Contact` | idem |
| ALL_UTM | `HASHED_USER_ID` | idem |

Conséquence : un dossier porte **le même montant dans les 10 snapshots**, donc le calcul hebdomadaire par différence de fichiers reste valide et l'historique ne bouge pas d'un rafraîchissement à l'autre. Le montant étant indépendant de la date, du canal et du statut, le CA hebdomadaire vaut ≈ 550 € × nombre de contrats : les courbes gardent exactement la même forme, le ratio €/contrat restant à 552 € ± 3 € (amplitude 2,2 % sur les 9 semaines).

---

## 1. Vue d'ensemble

| # | Fichier(s) dans `data/` | Rôle dans l'application d'origine | Onglets du dashboard |
|---|---|---|---|
| 1 | `exports/AAAA.MM.JJ/Souscriptions/souscriptions_AAAA.MM.JJ_total.csv` | Exports du backoffice d'ouverture de compte (scraping hebdomadaire) | Souscriptions, Cohortes, Attribution, Suivi Budget |
| 2 | `exports/AAAA.MM.JJ/Prescriptions/prescriptions_AAAA.MM.JJ_total.csv` | Exports du backoffice des prescriptions agences (EER) | Prescriptions, Attribution, Suivi Budget |
| 3 | `caches/cache_presc_activations.json`, `caches/presc_pivot_cache.json` | Caches calculés par l'app à partir des exports prescriptions | Prescriptions |
| 4 | `cohortes/Matrice_cohorte_conversion.xlsx` | Matrice de conversion par cohorte, recalculée depuis les exports | Cohortes |
| 5 | `utm_cache/ALL_UTM.csv` | Extraction entrepôt de données : parcours + attribution marketing | Attribution, Suivi Budget, Budget Média, Tracking Marketing, exports Excel |
| 6 | `piano_cache/s{site}_{AAAA-MM-JJ}.json` | Cache jour par jour de l'outil d'audience web (Piano Analytics) | Tracking Marketing, Budget Média, Portail Entrepreneur |
| 7 | `budget/Nova - 2026 - Suivi Budgets.xlsx` | Fichier de suivi budgétaire hebdomadaire | Suivi Budget |
| 8 | `budget/media_plan_cache.json` | Plan média par plateforme (à l'origine un Google Sheet) | Budget Média, Suivi Budget |
| 9 | `projections/Nova - distribution_volume_cible.xlsx` | Objectifs hebdomadaires de souscriptions | Souscriptions (objectif), Attribution |
| 10 | `aggregates/kpi_daily.csv`, `aggregates/kpi_weekly.csv` | KPI pré-agrégés (spécifiques au mockup, non utilisés par `app.py`) | — |

---

## 2. Détail

### 2.1 Exports Souscriptions (ouvertures de compte)

- **10 snapshots** : chaque dimanche du 03/05 au 28/06, plus une clôture au **02/07**. Chaque fichier est **cumulatif** : il contient tous les dossiers créés jusqu'à sa date, avec leur statut à ce moment-là.
- Le dashboard calcule les volumes hebdomadaires par **différence entre deux snapshots** ou par date de création.
- **17 136 dossiers**, chacun avec un cycle de vie cohérent d'un snapshot à l'autre : non démarré → en cours → en attente d'analyse → validé / refusé.
- **Format :** CSV `;`, UTF-8.

| Colonne | Exemple | Remarque |
|---|---|---|
| `Numéro` | `1` | Rang dans l'export (1 = plus récent), **pas** un identifiant stable |
| `Date de création` | `28/06/2026 22:02:17` | `dd/mm/YYYY HH:MM:SS` |
| `Responsable` | *(vide)* / `Camille MARTIN` | Agent backoffice fictif |
| `Statut` | `Dossier validé` | Voir ci-dessous |
| `Étape en cours` | `Parcours terminé` | Étape du parcours d'ouverture |
| `Valeur` | `819.26` | Montant du contrat en € (voir « Montant des contrats ») |

- **Statuts :** `Souscription non démarrée`, `Souscription en cours`, `En attente d'analyse Agent - Dossier avec erreurs`, `En attente d'analyse Agent - Dossier sans erreur`, `Dossier refusé`, `Dossier validé`.
  L'app reconnaît aussi un format « court » (`Initié`, `En cours`, `À traiter`, `Refusé`, `Validé`), non utilisé ici.
- **Répartition finale simulée :** validé ≈ 41 %, refusé ≈ 14,5 %, en cours (abandon) ≈ 35,5 %, non démarré ≈ 8,5 %.
- **Couleurs utilisées par l'app :** non démarrée `#1a3d6b`, en cours `#0d7a8a`, attente avec erreurs `#e07840`, attente sans erreur `#0dab82`, refusé `#c0392b`, validé `#2dd44a`.
- **Routes :** `/api/chart/souscriptions`, `/api/stats/souscriptions`, `/api/stats/souscriptions/weekly_opened`, `/api/table/souscriptions`, `/api/chart/prospects_weekly`, `/api/chart/souscriptions_target`, `/api/cohortes/update`.

### 2.2 Exports Prescriptions (entrées en relation via agence — EER)

- Mêmes 10 dates de snapshot, **13 706 prescriptions**.
- **Format :** CSV `;`. L'app détecte les colonnes via l'en-tête.

| Colonne | Exemple | Remarque |
|---|---|---|
| `Code agence` | `02636` | Agence fictive (pool de 900) |
| `Statut` | `EER validé` | |
| `Création` | `02/07/2026` | `dd/mm/YYYY` sans heure |
| `Onboarding` | `ebcb3fa2` / `—` | ID onboarding, `—` si non démarré |
| `Code_Contact` | `1000063740` | **Identifiant unique** (10 chiffres) → jointure entre snapshots |
| `Valeur` | `549.09` | Montant du contrat en € (voir « Montant des contrats ») |

- **Statuts :** `EER non démarré`, `EER démarré`, `EER à valider`, `EER à rejeter`, `EER rejeté`, `EER validé`.
- **Activations** (calculées par l'app) : passage de « non démarré » à un autre statut entre deux snapshots, par jointure sur `Code_Contact`.
- **Routes :** `/api/chart/prescriptions`, `/api/stats/prescriptions`, `/api/chart/prescriptions_target`, `/api/chart/attribution_daily`.

### 2.3 Caches prescriptions

| Fichier | Contenu | Format |
|---|---|---|
| `cache_presc_activations.json` | Nombre d'EER démarrés par semaine (clé = lundi) | `{"2026-05-04": 296324.94, …}` (montants en €) |
| `presc_pivot_cache.json` | Répartition des statuts par semaine de création | `{"2026-05-04\|2026-05-10": {"EER validé": 205103.7, …}}` |

### 2.4 Matrice de cohortes

- `Taux Lundi-Dimanche` : ligne 4 = en-têtes (`Cohorte\n(semaine du)`, `Extract\nJJ/MM`…) ; cellules `"40.9%\n(2170/5300)"`.
- `Volumes Lundi-Dimanche` : ligne 3 = en-têtes (`Total\nJJ/MM`, `Validés\nJJ/MM`…) ; valeurs entières ou `—`.
- Cohorte = semaine lundi→dimanche de création ; chaque colonne = un export du dimanche. Le bouton ↺ du dashboard la recalcule (`/api/cohortes/update`).
- La première ligne (03/05) ne contient qu'un dimanche, la dernière (29/06 → 02/07) est partielle.

### 2.5 Parcours & attribution — `utm_cache/ALL_UTM.csv`

- Une ligne par (utilisateur × session), **~17 300 utilisateurs / ~25 800 lignes**.
- **Format :** CSV `,`, 35 colonnes. **Attention :** `AF_PID` et `AT_CAMPAIGN_AFF` apparaissent deux fois dans l'en-tête (format d'origine conservé).

| Groupe | Colonnes |
|---|---|
| Utilisateur | `HASHED_USER_ID`, `USER_LIFECYCLE_STATUS` (`validated`, `in_progress`, `rejected`, `not_started`, `rejected_awaiting_agent_confirmation`…), `USER_LIFECYCLE_START_DATE` (`2026-06-11 03:01:01.115 +0200`), `COMPANY_CREATED_AT` |
| Parcours | `STEP_ORDER`, `STEP_TYPE` (`contractsigning`, `companyinfo`, `identity`…), `STEP_UPDATED_AT` |
| Session | `REFERRER_URL`, `SOURCE_APPLICATION`, `SESSION_CREATED_AT`, `SESSION_ID` |
| Tracking | `AT_MEDIUM`, `AT_CAMPAIGN`, `AT_CHANNEL`, `AT_VARIANT`, `AT_CREATION`, `AT_REFERER`, `AT_AFF_IDENTIFIER`, `AT_CAMPAIGN_AFF`, `AT_AFF_TYPE` |
| UTM | `UTM_SOURCE`, `UTM_MEDIUM`, `UTM_CAMPAIGN`, `UTM_ID`, `UTM_CONTENT`, `UTM_TERM` |
| App mobile (AppsFlyer) | `AF_PID`, `AF_CAMPAIGN_NAME`, `AF_XP`, `AF_C_ID`, `AF_SITEID`, `AF_CLICK_LOOKBACK`, `AF_DP` |
| Montant | `VALEUR` (€, voir « Montant des contrats ») |

- **Logique app :** déduplication par utilisateur (dernière session marketing, hors relances `email-abandon`).
  - **Leads** : datés sur `USER_LIFECYCLE_START_DATE`.
  - **Conversions** : `validated`, datés sur `COMPANY_CREATED_AT`.
  - Classement par canal : prescription, affiliation, partenariat, event, media, organique, direct.

### 2.6 Trafic web — `piano_cache/`

- **Sites :** `100001` = site Nova (landing `www.l.nova.example/pro`, parcours `pro.nova.example/inscription/…`) ; `100002` = Portail Entrepreneur (`www.entrepreneur.nova.example`).
- **Fichier :** `s{site}_{AAAA-MM-JJ}.json` = liste de lignes déjà agrégées pour un jour (61 jours × 2 sites). ~111 000 lignes / 3,7 M visites pour le site principal, ~23 000 lignes / 1,15 M visites pour le portail.
- **Champs :**
  - date et source : `date`, `src`, `src_detail`, `src_campaign`, `src_creation`, `src_variant` ;
  - contexte : `device_type`, `visitor_privacy_consent` ;
  - mesures : **`m_visits`**, **`m_unique_visitors`** ;
  - page et tracking : `event_url_path`, `at_medium`, `at_campaign`, `at_channel`, `at_variant`, `at_creation`, `at_aff_identifier`, `tracking_platform` ;
  - indicateurs : `has_gclid`, `has_fbclid`, `has_ttclid`, `mode`, `client_id`.
- **Logique app :** `_fetch_range_cached()` lit ces fichiers ; les filtres sont évalués localement (`_eval_piano_filter`), par exemple le filtre consentement `{'visitor_privacy_consent': {'$eq': True}}`.

### 2.7 Budget — `budget/Nova - 2026 - Suivi Budgets.xlsx`

- Feuille `Suivi hebdo FULL` (lue par `/api/budget` et `/api/budget/export`) :
  - ligne 2 : libellés semaines (`S18` vide de « pré-affichage », puis `S19…S27`, puis `TOTAL`) ;
  - ligne 3 : lundis ;
  - à partir de la ligne 5 : `Dépense` | `Type` (`Perf` / `Ressource`) | valeurs hebdo en €.
- Lignes : Créa, Articles, Média, Influence, Affiliation, CRM, Partenariat Perf, Partenariat Pilotage, Site, Pilotage, TOTAL, Cumul Dépense, puis `Media Nova` / `Media Portail`.
- La colonne S18 vide conserve le mapping de l'app : colonne Excel = index de semaine affichée + 3.

### 2.8 Plan média — `budget/media_plan_cache.json`

```json
{ "n": 9, "weeks": ["2026-05-04", …],
  "nova_lines":    [{"category": "SEA", "platform": "Google Search", "values": [9 valeurs]}, …],
  "portail_lines": [{"category": "Social Ads", "platform": "Meta", "values": […]}, …],
  "nova_total": […], "portail_total": […] }
```
- Plateformes Nova : Google Search, Performance Max, Meta, TikTok, Demand Gen, Amazon.
- Plateformes Portail : Meta, Demand Gen.
- `n` doit être égal au nombre de semaines affichées (9).

### 2.9 Objectifs — `projections/Nova - distribution_volume_cible.xlsx`

- Feuille `Distribution Volume` : `Semaine` | `Du (Mardi)` | `Au (Lundi)` | `Volume Client` | `Tx Transfo` | `Volume Prospect` | `Période` | `Increment`.
- Semaines 15 → 24 (28/04 → 06/07). Modifiable depuis le dashboard (`/api/projection/save`).
- Les objectifs sont stockés **en nombre de contrats** ; l'API les convertit en € en les multipliant par `VALEUR_MOYENNE` (550 €), et la sauvegarde applique la division inverse.

### 2.10 Agrégats — `aggregates/`

- `kpi_daily.csv` (`;`) — 1 ligne par jour, avec les colonnes :
  - souscriptions : `souscriptions_creees`, `souscriptions_validees_02juil`, `souscriptions_refusees_02juil`, `validations_du_jour` ;
  - prescriptions : `prescriptions_creees`, `prescriptions_validees_02juil`, `eer_demarres_du_jour` ;
  - trafic et leads : `leads_utm`, `visites_nova`, `visites_portail` ;
  - chiffre d'affaires : `ca_souscriptions_creees_eur`, `ca_prescriptions_creees_eur` (somme des colonnes `Valeur`).
- `kpi_weekly.csv` — mêmes indicateurs par semaine ISO + `depense_totale_eur`.
- Pratique pour prototyper un nouveau look & feel sans passer par la logique de différence entre snapshots.

---

## 3. Comportement du mode démo

- **Horloge :** « aujourd'hui » = **02/07/2026 22:00** (`DEMO_NOW` dans `app.py`). Côté navigateur, l'horloge démarre à cette date puis avance, ce qui est nécessaire pour les animations des graphiques.
- **Aucun appel réseau :** API d'audience, Google Sheets, téléchargement du budget et export backoffice sont désactivés. Les boutons concernés répondent « désactivé en mode démo ».
- **Fichiers modifiés par l'app :** `cohortes/Matrice_cohorte_conversion.xlsx` (bouton ↺) et `projections/…xlsx` (sauvegarde d'objectif).
