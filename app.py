#!/usr/bin/env python3
"""
app.py — MOCKUP du dashboard « Nova — Suivi Performance Marketing »

Version de démonstration 100 % locale :
  • données FACTICES (mai-juillet 2026) lues dans ./data (voir SOURCES_DONNEES.md)
  • aucun appel réseau (audience web, Google Sheets, stockage budget, entrepôt de données, backoffice désactivés)
  • horloge figée au 02/07/2026 22:00 (DEMO_NOW) pour que « aujourd'hui » tombe dans la période

Pré-requis : pip install flask openpyxl pandas
Lancement  : python3 app.py
Accès      : http://localhost:5002
"""

import os, csv, json, logging, socket, threading, subprocess, time, urllib.request, urllib.error
from urllib.parse import urlparse, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, date as _date
from collections import defaultdict
from pathlib import Path

try:
    from flask import Flask, render_template, jsonify, Response, stream_with_context, request
except ImportError:
    import sys
    sys.exit("❌  Flask non installé. Lance : pip install flask")

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════
BASE_DIR    = Path(__file__).resolve().parent
DATA_DIR    = BASE_DIR / 'data'
EXPORTS_DIR = DATA_DIR / 'exports'          # dossiers AAAA.MM.JJ/{Souscriptions,Prescriptions}
COMP_DIR    = DATA_DIR / 'cohortes'
PRESC_ACTIVATION_CACHE = DATA_DIR / 'caches' / 'cache_presc_activations.json'

# ── Mode démo : horloge figée ───────────────────────────────────────────────
DEMO_NOW = datetime(2026, 7, 2, 22, 0, 0)

def _now():
    return DEMO_NOW

def _today():
    return DEMO_NOW.date()

# API Piano désactivée (mockup hors-ligne : uniquement le cache local data/piano_cache)
AT_KEY     = ''
AT_URL     = ''
AT_SITE    = 100001
AT_SITE_PE = 100002
# Filtre Piano pour les visiteurs ayant donné leur consentement
CONSENT_FILTER = {'visitor_privacy_consent': {'$eq': True}}

# Référence officielle S1 : semaine du lundi 19 jan 2026 (jeu factice décalé : S19 = 4 mai)
# Les semaines avant cette date sont affichées S-1, S-2, ...
S1_MON = _date(2026, 5, 4)       # Mon mode : 1re semaine du jeu factice (S19)
S1_THU = _date(2026, 4, 30)      # Thu mode

# Début de l'affichage (peut inclure des semaines pré-S1 nommées S-n)
WEEK1_THU = _date(2026, 4, 30)   # Thu mode : display from 30 avr
WEEK1_MON = _date(2026, 5, 4)    # Mon mode : display from 4 mai

def _wk_label(week_start, s1_ref=None):
    """Retourne le numéro de semaine calendaire ISO (ex: 'S25' pour le 15/06/2026)."""
    return f'S{week_start.isocalendar()[1]:02d}'

BUDGET_EXCEL        = DATA_DIR / 'budget' / 'Nova - 2026 - Suivi Budgets.xlsx'

def _budget_n_disp():
    """Détecte dynamiquement le nombre de semaines disponibles dans le fichier budget.
    Lit la ligne d'en-tête 2 de 'Suivi hebdo FULL' et compte les colonnes S* avant TOTAL.
    Mapping : bci = di + 3  (col Excel idx 3 = display idx 0 = S-1).
    Retourne au minimum 1 (mockup : 9 semaines S19→S27)."""
    if openpyxl is None or not BUDGET_EXCEL.exists():
        return 9
    try:
        wb  = openpyxl.load_workbook(BUDGET_EXCEL, data_only=True)
        ws  = wb['Suivi hebdo FULL']
        row2 = next(ws.iter_rows(min_row=2, max_row=2, values_only=True))
        # Compter les colonnes de données à partir de idx 2 (première étiquette 'S1')
        n_data = 0
        for v in row2[2:]:
            s = str(v).strip() if v is not None else ''
            if s.upper() == 'TOTAL' or not s:
                break
            n_data += 1
        # di_max = n_data - 2  (col bci=2+n_data-1 → di = bci-3 = n_data-2)
        # N_DISP  = n_data - 1
        return max(1, n_data - 1)
    except Exception:
        return 9

app = Flask(__name__, template_folder=str(BASE_DIR / 'templates'))
app.config['TEMPLATES_AUTO_RELOAD'] = True

# ── État de l'export (thread-safe) ──────────────────────────────────────────
_lock   = threading.Lock()
_export = {'running': False, 'done': False, 'log': []}

# ── Configuration des sections (reprend CHART_CONFIGS de run_weekly_export.py) ──
CHART_CFG = {
    'souscriptions': {
        'subdir':     'Souscriptions',
        'prefix':     'souscriptions',
        'date_col':   1,
        'statut_col': 3,
        'status_order': [
            'Souscription non démarrée',
            'Souscription en cours',
            "En attente d'analyse Agent - Dossier avec erreurs",
            "En attente d'analyse Agent - Dossier sans erreur",
            'Dossier refusé',
            'Dossier validé',
            # Nouveaux statuts courts (backoffice post-juillet 2026)
            'Initié',
            'En cours',
            'À traiter',
            'Refusé',
            'Validé',
        ],
        'colors': {
            'Souscription non démarrée':                             '#1a3d6b',
            'Souscription en cours':                                 '#0d7a8a',
            "En attente d'analyse Agent - Dossier avec erreurs":     '#e07840',
            "En attente d'analyse Agent - Dossier sans erreur":      '#0dab82',
            'Dossier refusé':                                        '#c0392b',
            'Dossier validé':                                        '#2dd44a',
            # Nouveaux statuts courts (mêmes couleurs que leurs équivalents)
            'Initié':    '#1a3d6b',
            'En cours':  '#0d7a8a',
            'À traiter': '#e07840',
            'Refusé':    '#c0392b',
            'Validé':    '#2dd44a',
        },
        'valide_key':    'Dossier validé',
        'valeur_col':    5,   # col 5 = 'Valeur' (montant du contrat en €)
        'finalisee_col': None,  # 'Finalisée le' : détectée par nom si le format l'expose
    },
    'prescriptions': {
        'subdir':     'Prescriptions',
        'prefix':     'prescriptions',
        'date_col':   1,
        'statut_col': 0,
        'id_col':     3,   # Code_Contact — identifiant unique de la prescription
        'status_order': [
            'EER non démarré', 'EER démarré', 'EER à valider',
            'EER à rejeter',   'EER rejeté',  'EER validé',
        ],
        'colors': {
            'EER non démarré': '#1a3d6b', 'EER démarré':  '#0d7a8a',
            'EER à valider':   '#0dab82', 'EER à rejeter': '#e07840',
            'EER rejeté':      '#c0392b', 'EER validé':    '#2dd44a',
        },
        'valide_key': 'EER validé',
        'valeur_col': 5,   # col 5 = 'Valeur' (montant du contrat en €)
    },
}

# Valeur moyenne d'un contrat (tirage uniforme 100–1000 €) : sert à convertir
# les objectifs, exprimés en nombre de contrats, dans l'unité affichée (€).
VALEUR_MOYENNE = 550.0


# ── Arrondi global des montants renvoyés par l'API ────────────────────────────
# Les agrégations somment des montants en euros : on coupe le bruit flottant
# (1234.5600000000002) au centime, une seule fois, à la frontière JSON.
_flask_jsonify = jsonify

def _round_floats(obj, nd=2):
    if isinstance(obj, float):
        return round(obj, nd)
    if isinstance(obj, dict):
        return {k: _round_floats(v, nd) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_round_floats(v, nd) for v in obj]
    return obj

def jsonify(*args, **kwargs):
    if kwargs and not args:
        return _flask_jsonify(**_round_floats(kwargs))
    if len(args) == 1 and not kwargs:
        return _flask_jsonify(_round_floats(args[0]))
    return _flask_jsonify(*_round_floats(list(args)), **_round_floats(kwargs))

# ── Corrections manuelles du file-diff souscriptions ───────────────────────
# Mécanisme prévu pour rattraper un export backoffice incomplet (comptes manquants
# sur une semaine donnée, réapparus dans l'export suivant, ce qui fausse le file-diff).
# Le jeu de données factices n'en a pas besoin : le dictionnaire reste vide.
_SOUS_FILEDIFF_CORRECTIONS = {}   # corrections propres aux données réelles : sans objet

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════════════════════
def week_start_thu(d):
    """Retourne le jeudi de la semaine (Jeu→Mer) contenant d."""
    return d - timedelta(days=(d.weekday() - 3) % 7)

def week_start_mon(d):
    """Retourne le lundi de la semaine ISO (Lun→Dim) contenant d."""
    return d - timedelta(days=d.weekday())

def is_chrome_debug():
    """Mode démo : pas de backoffice ni de Chrome debug."""
    return False

def _is_total_csv(f, prefix):
    """Vérifie si un fichier est un CSV total pour le prefix donné (insensible à la casse)."""
    return (f.stem.lower().startswith(prefix.lower())
            and 'total' in f.stem.lower()
            and f.suffix.lower() == '.csv')

def _presc_col_indices(path, delim=None):
    """Lit l'en-tête d'un CSV prescriptions et retourne (statut_col, date_col, id_col).
    Gère les deux formats : avec ou sans colonne 'Code agence' en tête.
    Fallback sur les indices fixes (0, 1, 3) si l'en-tête est absent ou non reconnu."""
    if delim is None:
        delim = detect_delimiter(path)
    try:
        with open(path, encoding='utf-8') as _f:
            header = [h.strip() for h in next(csv.reader(_f, delimiter=delim), [])]
        def _idx(name):
            for i, h in enumerate(header):
                if h.lower() == name.lower():
                    return i
            return None
        sc = _idx('Statut')
        dc = _idx('Création')
        ic = _idx('Code_Contact')
        if sc is not None and dc is not None and ic is not None:
            return sc, dc, ic
    except Exception:
        pass
    return 0, 1, 3  # indices fixes (ancien format)

def _valeur_col_of(path):
    """Index de la colonne 'Valeur' (montant du contrat), ou None si absente."""
    try:
        with open(path, encoding='utf-8') as _f:
            hdr = next(csv.reader(_f, delimiter=detect_delimiter(path)), [])
        for i, h in enumerate(hdr):
            if h.strip().lstrip('\ufeff').lower() == 'valeur':
                return i
    except Exception:
        pass
    return None


def _row_value(row, cfg, mode='eur'):
    """Mesure portée par la ligne : montant € ('eur') ou unité ('count').
    Retourne 1.0 si la colonne 'Valeur' est absente."""
    if mode == 'count':
        return 1.0
    vc = cfg.get('valeur_col')
    if vc is None or vc >= len(row):
        return 1.0
    try:
        return float(str(row[vc]).strip().replace(',', '.') or 0)
    except ValueError:
        return 1.0


def _cfg_for_file(cfg, path):
    """Retourne une copie de cfg avec date_col/statut_col/id_col détectés depuis l'en-tête du fichier.
    Gère :
    - prescriptions : colonnes dynamiques (ajout 'Code agence')
    - souscriptions : nouveau format backoffice (statuts courts 'Validé' vs 'Dossier validé')
    """
    if path is None:
        return cfg
    prefix = cfg.get('prefix', '')
    if prefix == 'prescriptions':
        sc, dc, ic = _presc_col_indices(path)
        updated = dict(cfg)
        updated['statut_col'] = sc
        updated['date_col']   = dc
        updated['id_col']     = ic
        updated['valeur_col'] = _valeur_col_of(path)
        return updated
    if prefix == 'souscriptions':
        # Détecte si le fichier utilise les nouveaux statuts courts ('Validé', 'En cours'…)
        # vs l'ancien format ('Dossier validé', 'Souscription en cours'…).
        # On lit juste les statuts des premières lignes data.
        _NEW_VALIDE   = 'Validé'
        _OLD_VALIDE   = 'Dossier validé'
        _NEW_STATUS_ORDER = ['Initié', 'En cours', 'À traiter', 'Refusé', _NEW_VALIDE]
        _NEW_COLORS   = {
            'Initié':    '#1a3d6b',
            'En cours':  '#0d7a8a',
            'À traiter': '#e07840',
            'Refusé':    '#c0392b',
            _NEW_VALIDE: '#2dd44a',
        }
        try:
            delim = detect_delimiter(path)
            with open(path, encoding='utf-8') as _f:
                rdr = csv.reader(_f, delimiter=delim)
                next(rdr, None)  # skip header
                sample = [row for _, row in zip(range(10), rdr)]
            statut_col = cfg.get('statut_col', 3)
            statuts = {r[statut_col].strip() for r in sample if statut_col < len(r)}
            updated = dict(cfg)
            updated['valeur_col'] = _valeur_col_of(path)
            if _NEW_VALIDE in statuts or ('En cours' in statuts and _OLD_VALIDE not in statuts):
                updated['valide_key']    = _NEW_VALIDE
                updated['status_order']  = _NEW_STATUS_ORDER
                updated['colors']        = _NEW_COLORS
            return updated
        except Exception:
            pass
    return cfg

def detect_delimiter(filepath):
    """Détecte le délimiteur d'un CSV (';' ou ',')."""
    with open(filepath, encoding='utf-8') as f:
        first = f.readline()
    return ';' if first.count(';') >= first.count(',') else ','

def _clean_date_cell(v: str) -> str:
    """Extrait la date d'une cellule CSV, en retirant le préfixe 'Voir les détails' si présent."""
    v = v.strip()
    if v.lower().startswith('voir les détails'):
        v = v[len('voir les détails'):].strip()
    return v.split(' ')[0]

def find_latest_export(prefix, subdir):
    """Retourne (Path, date) du fichier *_total.csv le plus récent."""
    best_path, best_date = None, None
    for entry in EXPORTS_DIR.iterdir():
        try:
            d = datetime.strptime(entry.name, '%Y.%m.%d').date()
        except ValueError:
            continue
        folder = entry / subdir
        if not folder.is_dir():
            continue
        for f in folder.iterdir():
            if _is_total_csv(f, prefix):
                if best_date is None or d > best_date:
                    best_date, best_path = d, f
    return best_path, best_date

def find_export_for_date(prefix, subdir, date_str):
    """Retourne (Path, date) pour une date spécifique (format 'YYYY-MM-DD')."""
    try:
        target = datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return None, None
    folder = EXPORTS_DIR / target.strftime('%Y.%m.%d') / subdir
    if folder.is_dir():
        for f in folder.iterdir():
            if _is_total_csv(f, prefix):
                return f, target
    return None, target

def list_exports(prefix, subdir):
    """Retourne les dates d'export disponibles, triées du plus récent au plus ancien."""
    dates = []
    for entry in EXPORTS_DIR.iterdir():
        try:
            d = datetime.strptime(entry.name, '%Y.%m.%d').date()
        except ValueError:
            continue
        folder = entry / subdir
        if not folder.is_dir():
            continue
        for f in folder.iterdir():
            if _is_total_csv(f, prefix):
                dates.append(str(d))
                break
    return sorted(dates, reverse=True)

def _find_file_on_or_before(target_date, prefix, subdir):
    """Fichier *_total.csv dont le dossier date est ≤ target_date (le plus récent possible)."""
    best_path, best_date = None, None
    for entry in EXPORTS_DIR.iterdir():
        try:
            d = datetime.strptime(entry.name, '%Y.%m.%d').date()
        except ValueError:
            continue
        if d > target_date:
            continue
        folder = entry / subdir
        if not folder.is_dir():
            continue
        for f in folder.iterdir():
            if _is_total_csv(f, prefix):
                if best_date is None or d > best_date:
                    best_date, best_path = d, f
    return best_path, best_date

def _find_first_file_on_or_after(target_date, prefix, subdir):
    """Fichier *_total.csv dont le dossier date est ≥ target_date (le plus ancien possible)."""
    best_path, best_date = None, None
    for entry in EXPORTS_DIR.iterdir():
        try:
            d = datetime.strptime(entry.name, '%Y.%m.%d').date()
        except ValueError:
            continue
        if d < target_date:
            continue
        folder = entry / subdir
        if not folder.is_dir():
            continue
        for f in folder.iterdir():
            if _is_total_csv(f, prefix):
                if best_date is None or d < best_date:
                    best_date, best_path = d, f
    return best_path, best_date

def _load_presc_cache():
    try:
        with open(PRESC_ACTIVATION_CACHE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _save_presc_cache(cache):
    try:
        with open(PRESC_ACTIVATION_CACHE, 'w', encoding='utf-8') as f:
            json.dump(cache, f)
    except Exception:
        pass

def _load_presc_full(path):
    """Retourne un dict {code_contact: {'statut': str, 'creation': date|None}}
    en détectant les colonnes depuis le header (robuste aux formats anciens).
    Déduplique par Code_Contact (dernière occurrence retenue)."""
    result = {}
    if path is None or not path.exists():
        return result
    try:
        delim = detect_delimiter(path)
        with open(path, encoding='utf-8') as fh:
            reader = csv.reader(fh, delimiter=delim)
            headers = [h.strip().lstrip('\ufeff') for h in next(reader)]
            if 'Code_Contact' not in headers or 'Statut' not in headers:
                return result
            id_col     = headers.index('Code_Contact')
            statut_col = headers.index('Statut')
            date_col   = headers.index('Création') if 'Création' in headers else None
            valeur_col = headers.index('Valeur') if 'Valeur' in headers else None
            for row in reader:
                if max(statut_col, id_col) >= len(row):
                    continue
                code    = row[id_col].strip()
                statut  = row[statut_col].strip()
                creation = None
                if date_col is not None and date_col < len(row):
                    try:
                        creation = datetime.strptime(row[date_col].strip(), '%d/%m/%Y').date()
                    except ValueError:
                        pass
                valeur = 1.0
                if valeur_col is not None and valeur_col < len(row):
                    try:
                        valeur = float(row[valeur_col].strip().replace(',', '.') or 0)
                    except ValueError:
                        pass
                result[code] = {'statut': statut, 'creation': creation, 'valeur': valeur}
    except Exception:
        pass
    return result

def _count_gross_activations(end_path, prev_path, wstart=None, wend=None, mode='eur'):
    """Compte les prescriptions qui ont DÉMARRÉ leur EER entre prev et end.

    Pour chaque prescription active (≠ EER non démarré) dans end_path :
      - Si absente de prev_path OU 'EER non démarré' dans prev_path → nouvelle activation

    Déduplique par Code_Contact (cohérent avec _load_presc_full).
    Les colonnes sont auto-détectées depuis les headers.

    Si wstart/wend sont fournis, retourne un tuple (same_week, prior_weeks) ;
    sinon retourne juste le total (int).
    """
    if end_path is None or not end_path.exists():
        return (0, 0) if wstart is not None else 0

    prev = _load_presc_full(prev_path)
    end  = _load_presc_full(end_path)

    same_week   = 0.0
    prior_weeks = 0.0

    for code, info in end.items():
        if info['statut'] == 'EER non démarré':
            continue
        prev_info = prev.get(code)
        if prev_info is not None and prev_info['statut'] != 'EER non démarré':
            continue  # déjà active avant → pas une activation cette semaine
        # Nouvelle activation : classer par date de création
        _v = 1.0 if mode == 'count' else info.get('valeur', 1.0)
        if wstart is not None:
            creation = info['creation']
            if creation is not None and wstart <= creation <= wend:
                same_week += _v
            else:
                prior_weeks += _v
        else:
            same_week += _v  # total quand pas de split

    if wstart is not None:
        return same_week, prior_weeks
    return same_week  # total

def _count_active_presc_in_range(path, statut_col, date_col, wstart, wend, valeur_col=None, mode='eur'):
    """Somme les montants des prescriptions actives créées dans [wstart, wend]."""
    if path is None or not path.exists():
        return 0
    _vcfg = {'valeur_col': valeur_col}
    n = 0.0
    try:
        delim = detect_delimiter(path)
        with open(path, encoding='utf-8') as fh:
            first = True
            for row in csv.reader(fh, delimiter=delim):
                if first:
                    first = False
                    continue
                if max(statut_col, date_col) < len(row):
                    if row[statut_col].strip() == 'EER non démarré':
                        continue
                    try:
                        raw = _clean_date_cell(row[date_col])
                        d = datetime.strptime(raw, '%d/%m/%Y').date()
                        if wstart <= d <= wend:
                            n += _row_value(row, _vcfg, mode)
                    except ValueError:
                        pass
    except Exception:
        pass
    return n

def _presc_activations_by_week(week_starts, presc_cfg, curr_week=None, mode='eur'):
    """Prescriptions ayant DÉMARRÉ leur EER (≠ EER non démarré) par semaine.

    Pour chaque semaine W :
      - Trouve l'export le plus récent ≤ fin de W  (end_export)
      - Trouve l'export le plus récent ≤ fin de W-1 (prev_export)
      - Si deux fichiers distincts : join par Code_Contact →
          compte les prescriptions actives dans end ET (nouvelles OU non démarrées dans prev)
          = activations BRUTES (indépendant des clôtures/disparitions)
      - Si même fichier ou pas d'antérieur : fallback sur date de création dans end_export

    Les semaines historiques (≠ curr_week) sont mises en cache JSON.
    curr_week : semaine toujours recalculée (semaine courante de l'export sélectionné).
    """
    prefix     = presc_cfg['prefix']
    subdir     = presc_cfg['subdir']
    statut_col = presc_cfg['statut_col']
    date_col   = presc_cfg['date_col']
    id_col     = presc_cfg.get('id_col', 3)

    cache = _load_presc_cache() if mode == 'eur' else {}
    results = []
    cache_updated = False

    for wstart in week_starts:
        wkey = str(wstart)
        # Semaine historique déjà calculée → utiliser le cache
        if wkey in cache and wstart != curr_week:
            results.append(cache[wkey])
            continue

        wend     = wstart + timedelta(days=6)
        prev_end = wstart - timedelta(days=1)
        end_path,  _ = _find_file_on_or_before(wend,     prefix, subdir)
        prev_path, _ = _find_file_on_or_before(prev_end, prefix, subdir)

        two_distinct_files = (
            end_path is not None and prev_path is not None
            and str(end_path) != str(prev_path)
        )

        if two_distinct_files:
            # Cas normal : join ligne à ligne, dédupliqué par Code_Contact
            same_w, prior_w = _count_gross_activations(end_path, prev_path, wstart, wend, mode)
            val = same_w + prior_w
        elif end_path is not None:
            # Même snapshot ou pas d'antérieur → fallback date de création
            _ep_cfg = _cfg_for_file(presc_cfg, end_path)
            val = _count_active_presc_in_range(end_path, _ep_cfg['statut_col'], _ep_cfg['date_col'], wstart, wend, _ep_cfg.get('valeur_col'), mode)
        else:
            # Pas encore de snapshot : premier disponible après la semaine
            fallback, _ = _find_first_file_on_or_after(wend, prefix, subdir)
            if fallback:
                _fb_cfg = _cfg_for_file(presc_cfg, fallback)
                val = _count_active_presc_in_range(fallback, _fb_cfg['statut_col'], _fb_cfg['date_col'], wstart, wend, _fb_cfg.get('valeur_col'), mode)
            else:
                val = 0

        results.append(val)
        # Mettre en cache uniquement les semaines historiques avec une valeur fiable :
        # - Pas la semaine courante (curr_week)
        # - Uniquement si on avait deux fichiers distincts (résultat fiable)
        #   OU si aucun fichier n'existe encore (val=0 justifié)
        # → Jamais en fallback "même fichier" : la semaine est en cours / données incomplètes
        if wstart != curr_week and (two_distinct_files or end_path is None):
            cache[wkey] = val
            cache_updated = True

    if cache_updated and mode == 'eur':   # le cache disque ne stocke que les montants
        _save_presc_cache(cache)

    return results

def _acquisition_by_file_diff(week_starts, cfg, max_date=None, count_all=False, mode='eur'):
    """Comptes par semaine = delta entre le fichier de fin de semaine et le précédent.
    count_all=False : uniquement les lignes avec statut valide_key (ex: 'Dossier validé').
    count_all=True  : toutes les lignes (tous statuts).
    max_date : plafonne les lookups de fichiers à cette date (export sélectionné)."""
    prefix     = cfg['prefix']
    subdir     = cfg['subdir']
    valide_key = cfg['valide_key']

    # Collecte les bornes de chaque semaine (wend et prev_end)
    bounds = set()
    for wstart in week_starts:
        bounds.add(wstart + timedelta(days=6))   # fin de semaine
        bounds.add(wstart - timedelta(days=1))   # fin de semaine précédente

    # Résout chaque borne → (path, actual_date), plafonné à max_date si fourni
    def _capped(d):
        cap = min(d, max_date) if max_date else d
        return _find_file_on_or_before(cap, prefix, subdir)
    file_cache = {d: _capped(d) for d in bounds}

    # Comptage par path (lecture unique) — colonnes détectées par fichier
    count_cache = {}
    def _count(target_date):
        path, _ = file_cache[target_date]
        if path is None:
            return None
        cache_key = (str(path), count_all, mode)
        if cache_key not in count_cache:
            n = 0.0
            try:
                _fcfg      = _cfg_for_file(cfg, path)
                statut_col = _fcfg['statut_col']
                _vk        = _fcfg['valide_key']   # valide_key adapté au format du fichier
                delim = detect_delimiter(path)
                with open(path, encoding='utf-8') as fh:
                    first = True
                    for row in csv.reader(fh, delimiter=delim):
                        if first:
                            first = False
                            continue
                        if count_all:
                            n += _row_value(row, _fcfg, mode)
                        elif statut_col < len(row) and row[statut_col].strip() == _vk:
                            n += _row_value(row, _fcfg, mode)
            except Exception:
                n = None
            count_cache[cache_key] = n
        return count_cache[cache_key]

    # Fallback : comptage par date de création dans une fenêtre [wstart, wend]
    date_count_cache = {}
    def _count_created_in_range(path, wstart, wend):
        key = (str(path), wstart, wend, count_all, mode)
        if key in date_count_cache:
            return date_count_cache[key]
        n = 0
        try:
            _fcfg      = _cfg_for_file(cfg, path)
            date_col   = _fcfg['date_col']
            statut_col = _fcfg['statut_col']
            _vk        = _fcfg['valide_key']   # valide_key adapté au format du fichier
            delim = detect_delimiter(path)
            with open(path, encoding='utf-8') as fh:
                first = True
                for row in csv.reader(fh, delimiter=delim):
                    if first:
                        first = False
                        continue
                    if date_col < len(row):
                        if count_all or (statut_col < len(row) and row[statut_col].strip() == _vk):
                            try:
                                raw = _clean_date_cell(row[date_col])
                                d = datetime.strptime(raw, '%d/%m/%Y').date()
                                if wstart <= d <= wend:
                                    n += _row_value(row, _fcfg, mode)
                            except ValueError:
                                pass
        except Exception:
            n = 0
        date_count_cache[key] = n
        return n

    # Première passe : résout end_cnt / prev_cnt pour chaque semaine
    week_info = []
    for wstart in week_starts:
        wend     = wstart + timedelta(days=6)
        prev_end = wstart - timedelta(days=1)
        week_info.append((wstart, wend, _count(wend), _count(prev_end)))

    # Trouve la semaine « ancre » : première semaine avec end_cnt mais sans prev_cnt.
    # Elle absorbera tous les comptes validés avant elle et non attribués aux semaines
    # précédentes, garantissant que sum(actuals) == total du dernier fichier.
    anchor_idx = next(
        (i for i, (_, _, ec, pc) in enumerate(week_info) if ec is not None and pc is None),
        None
    )
    # Fichier ancre : celui utilisé par la semaine ancre (même base pour tous les fallbacks)
    anchor_path = None
    if anchor_idx is not None:
        anchor_path, _ = file_cache[week_info[anchor_idx][1]]  # wend de l'ancre

    # Deuxième passe : calcul des actuals
    pre_anchor_sum = 0
    results = []
    for i, (wstart, wend, end_cnt, prev_cnt) in enumerate(week_info):
        if end_cnt is not None and prev_cnt is not None:
            # Cas normal : diff entre deux fichiers
            results.append(max(0, end_cnt - prev_cnt))
        elif end_cnt is not None and prev_cnt is None:
            # Semaine ancre : tout ce que les fallbacks précédents n'ont pas capturé
            results.append(max(0, end_cnt - pre_anchor_sum))
        else:
            # Pas de fichier S : comptage par date de création dans le fichier ancre
            if anchor_path is not None:
                val = _count_created_in_range(anchor_path, wstart, wend)
            else:
                fallback_path, _ = _find_first_file_on_or_after(wend, prefix, subdir)
                val = _count_created_in_range(fallback_path, wstart, wend) if fallback_path else 0
            results.append(val)
            if anchor_idx is not None and i < anchor_idx:
                pre_anchor_sum += val

    # Applique les corrections manuelles (bug export May17, premier format BO).
    # S'applique uniquement aux souscriptions et aux semaines listées dans la table.
    if prefix == CHART_CFG['souscriptions']['prefix'] and not count_all:
        results = [
            _SOUS_FILEDIFF_CORRECTIONS.get(ws, v)
            for ws, v in zip(week_starts, results)
        ]
    return results

def _acquisition_by_finalisee_date(week_starts, cfg, max_date=None):
    """Compte les souscriptions validées par semaine via la colonne 'Finalisée le'.
    Avantage par rapport à file_diff : lit un seul fichier (le plus récent) sans
    avoir besoin des exports précédents.
    Retourne None si la colonne est absente du fichier (fallback sur file_diff attendu)."""
    prefix     = cfg['prefix']
    subdir     = cfg['subdir']
    valide_key = cfg.get('valide_key', 'Dossier validé')

    # Fichier de référence : le plus récent, ou plafonné à max_date
    if max_date:
        ref_path, _ = _find_file_on_or_before(max_date, prefix, subdir)
    else:
        ref_path, _ = find_latest_export(prefix, subdir)

    if ref_path is None or not Path(ref_path).exists():
        return None

    delim = detect_delimiter(ref_path)
    with open(ref_path, encoding='utf-8') as fh:
        file_headers = next(csv.reader(fh, delimiter=delim), [])

    # Colonne 'Finalisée le' détectée par nom
    finalisee_col = next((i for i, h in enumerate(file_headers)
                          if 'finalis' in h.strip().lower()), -1)
    if finalisee_col < 0:
        return None   # Colonne absente → l'appelant tombe sur file_diff

    _fcfg      = _cfg_for_file(cfg, ref_path)
    statut_col = _fcfg['statut_col']
    valide_key = _fcfg['valide_key']   # adapté au format du fichier (court ou long)

    # Précalcul des bornes de chaque semaine
    week_ends = {ws: ws + timedelta(days=6) for ws in week_starts}
    counts    = {ws: 0.0 for ws in week_starts}

    with open(ref_path, encoding='utf-8') as fh:
        reader = csv.reader(fh, delimiter=delim)
        next(reader, None)  # skip header
        for row in reader:
            if statut_col >= len(row) or row[statut_col].strip() != valide_key:
                continue
            if finalisee_col >= len(row):
                continue
            raw = row[finalisee_col].strip()
            if not raw:
                continue
            try:
                d = datetime.strptime(_clean_date_cell(raw), '%d/%m/%Y').date()
            except ValueError:
                try:
                    d = datetime.strptime(raw[:10], '%Y-%m-%d').date()
                except ValueError:
                    continue
            for ws in week_starts:
                if ws <= d <= week_ends[ws]:
                    counts[ws] += _row_value(row, _fcfg)

    return [counts[ws] for ws in week_starts]


def _acquisition_sous(week_starts, cfg=None, max_date=None):
    """Wrapper : utilise 'Finalisée le' si disponible, sinon file_diff."""
    if cfg is None:
        cfg = CHART_CFG['souscriptions']
    result = _acquisition_by_finalisee_date(week_starts, cfg, max_date=max_date)
    if result is None:
        result = _acquisition_by_file_diff(week_starts, cfg, max_date=max_date)
    return result


def compute_pivot(csv_path, cfg, n_weeks=5, week_mode='thu', offset=0, granularity='week'):
    """Calcule le pivot hebdomadaire ou quotidien.
    offset=0 → n dernières semaines/jours ; offset=k → fenêtre décalée de k semaines en arrière."""
    week_fn = week_start_mon if week_mode == 'mon' else week_start_thu
    # Détection dynamique des colonnes (gère les évolutions de format CSV)
    fcfg  = _cfg_for_file(cfg, csv_path)
    delim = detect_delimiter(csv_path)
    rows = []
    with open(csv_path, encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=delim)
        next(reader)
        for row in reader:
            max_col = max(fcfg['date_col'], fcfg['statut_col'])
            if len(row) > max_col and row[fcfg['date_col']]:
                try:
                    raw = _clean_date_cell(row[fcfg['date_col']])
                    d   = datetime.strptime(raw, '%d/%m/%Y').date()
                    rows.append({'date': d, 'statut': row[fcfg['statut_col']].strip(),
                                 'valeur': _row_value(row, fcfg)})
                except ValueError:
                    continue

    if granularity == 'day':
        daily = defaultdict(lambda: defaultdict(float))
        daily_n = defaultdict(lambda: defaultdict(int))
        for r in rows:
            daily[r['date']][r['statut']] += r['valeur']
            daily_n[r['date']][r['statut']] += 1
        all_days = sorted(daily.keys())
        n_days   = n_weeks * 7
        if offset > 0:
            end_idx = max(0, len(all_days) - offset * 7)
            days    = all_days[max(0, end_idx - n_days):end_idx]
        else:
            days = all_days[-n_days:]
        datasets = []
        for status in fcfg['status_order']:
            data = [daily[d].get(status, 0) for d in days]
            if any(v > 0 for v in data):
                datasets.append({
                    'label':           status,
                    'data':            data,
                    'backgroundColor': fcfg['colors'].get(status, '#888888'),
                })
        _DAY = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim']
        labels = [f"{_DAY[d.weekday()]}\n{d.strftime('%d/%m')}" for d in days]
        totals   = [round(sum(daily[d].values()), 2) for d in days]
        detail   = [{s: round(daily[d].get(s, 0), 2) for s in fcfg['status_order']} for d in days]
        totals_n = [sum(daily_n[d].values()) for d in days]
        detail_n = [{s: daily_n[d].get(s, 0) for s in fcfg['status_order']} for d in days]
        return {
            'labels':      labels,
            'datasets':    datasets,
            'totals':      totals,
            'detail':      detail,
            'totals_n':    totals_n,
            'detail_n':    detail_n,
            'unit':        'eur',
            'weeks':       [str(d) for d in days],
            'total_weeks': len(all_days),
            'granularity': 'day',
        }

    weekly   = defaultdict(lambda: defaultdict(float))
    weekly_n = defaultdict(lambda: defaultdict(int))
    for r in rows:
        weekly[week_fn(r['date'])][r['statut']] += r['valeur']
        weekly_n[week_fn(r['date'])][r['statut']] += 1

    week1     = WEEK1_MON if week_mode == 'mon' else WEEK1_THU
    all_weeks = [w for w in sorted(weekly.keys()) if w >= week1]
    if offset > 0:
        end_idx = max(0, len(all_weeks) - offset)
        weeks   = all_weeks[max(0, end_idx - n_weeks):end_idx]
    else:
        weeks   = all_weeks[-n_weeks:]
    datasets = []
    for status in fcfg['status_order']:
        data = [weekly[w].get(status, 0) for w in weeks]
        if any(v > 0 for v in data):
            datasets.append({
                'label':           status,
                'data':            data,
                'backgroundColor': fcfg['colors'].get(status, '#888888'),
            })

    s1_ref = S1_MON if week_mode == 'mon' else S1_THU
    labels = [
        f"{_wk_label(w, s1_ref)}\n{w.strftime('%d/%m')}–{(w + timedelta(days=6)).strftime('%d/%m')}"
        for w in weeks
    ]
    totals   = [round(sum(weekly[w].values()), 2) for w in weeks]
    detail   = [{s: round(weekly[w].get(s, 0), 2) for s in fcfg['status_order']} for w in weeks]
    totals_n = [sum(weekly_n[w].values()) for w in weeks]
    detail_n = [{s: weekly_n[w].get(s, 0) for s in fcfg['status_order']} for w in weeks]

    return {
        'labels':      labels,
        'datasets':    datasets,
        'totals':      totals,
        'detail':      detail,
        'totals_n':    totals_n,
        'detail_n':    detail_n,
        'unit':        'eur',
        'weeks':       [str(w) for w in weeks],
        'total_weeks': len(all_weeks),
        'granularity': 'week',
    }

_PRESC_PIVOT_CACHE_FILE = DATA_DIR / 'caches' / 'presc_pivot_cache.json'

def _load_presc_pivot_cache():
    try:
        with open(_PRESC_PIVOT_CACHE_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _save_presc_pivot_cache(cache):
    try:
        with open(_PRESC_PIVOT_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f)
    except Exception:
        pass


def compute_presc_pivot_stable(cfg, n_weeks=5, week_mode='mon', offset=0,
                                granularity='week', date_param=None):
    """Pivot prescriptions avec comptes STABLES par semaine.

    Pour chaque semaine W, utilise l'export le plus récent ≤ fin de W et compte
    les prescriptions dont la Date de création tombe dans [wstart, wend].
    Les semaines passées sont mises en cache → les comptes ne changent plus
    quand un nouvel export global arrive.

    date_param : borne max (export sélectionné), None = pas de borne.
    granularity == 'day' : délègue à compute_pivot classique (pas de stabilisation intraday).
    """
    prefix   = cfg['prefix']
    subdir   = cfg['subdir']
    date_col = cfg['date_col']
    week_fn  = week_start_mon if week_mode == 'mon' else week_start_thu
    week1    = WEEK1_MON if week_mode == 'mon' else WEEK1_THU

    # Borne max si export sélectionné
    max_date = None
    if date_param:
        try:
            max_date = datetime.strptime(date_param, '%Y-%m-%d').date()
        except ValueError:
            pass

    # Récupère la liste des semaines disponibles depuis l'export le plus récent
    if max_date:
        ref_path, ref_date = _find_file_on_or_before(max_date, prefix, subdir)
    else:
        ref_path, ref_date = find_latest_export(prefix, subdir)
    if ref_path is None:
        return None  # pas de données

    # Détecte les indices de colonnes depuis l'en-tête du fichier de référence
    ref_cfg   = _cfg_for_file(cfg, ref_path)
    date_col  = ref_cfg['date_col']

    # Construit la liste de toutes les semaines présentes dans le ref_path
    all_week_keys = set()
    delim = detect_delimiter(ref_path)
    with open(ref_path, encoding='utf-8') as _f:
        _rdr = csv.reader(_f, delimiter=delim)
        next(_rdr, None)
        for _row in _rdr:
            if len(_row) <= date_col:
                continue
            try:
                _raw = _clean_date_cell(_row[date_col])
                _d   = datetime.strptime(_raw, '%d/%m/%Y').date()
                if week_fn(_d) >= week1:
                    all_week_keys.add(week_fn(_d))
            except ValueError:
                continue

    if granularity == 'day':
        # Mode jour : pas de stabilisation par snapshot, délègue au pivot classique
        from collections import defaultdict as _dd
        pivot = compute_pivot(ref_path, ref_cfg, n_weeks=n_weeks, week_mode=week_mode,
                              offset=offset, granularity='day')
        return pivot

    all_weeks = sorted(all_week_keys)
    if offset > 0:
        end_idx = max(0, len(all_weeks) - offset)
        weeks   = all_weeks[max(0, end_idx - n_weeks):end_idx]
    else:
        weeks   = all_weeks[-n_weeks:]

    today       = _now().date()
    curr_week   = week_fn(today)
    pivot_cache = _load_presc_pivot_cache()
    cache_updated = False
    weekly_by_status = {}   # {week_key: {statut: count}}

    for wk in weeks:
        wk_str = str(wk)
        wend   = wk + timedelta(days=6)
        # Clé cache inclut la date du snapshot utilisé (pour invalider si un meilleur
        # snapshot arrive pour cette semaine)
        snap_path, snap_date = _find_file_on_or_before(
            min(wend, max_date) if max_date else wend, prefix, subdir
        )
        cache_key = f"{wk_str}|{snap_date}" if snap_date else None

        if (cache_key and cache_key in pivot_cache
                and wk != curr_week):
            weekly_by_status[wk_str] = pivot_cache[cache_key]
            continue

        # Lecture du snapshot de la semaine (détection colonnes par fichier)
        snap_cfg   = _cfg_for_file(cfg, snap_path)
        snap_dc    = snap_cfg['date_col']
        snap_sc    = snap_cfg['statut_col']
        counts = defaultdict(float)
        if snap_path and snap_path.exists():
            _delim = detect_delimiter(snap_path)
            with open(snap_path, encoding='utf-8') as _f:
                _rdr = csv.reader(_f, delimiter=_delim)
                next(_rdr, None)
                for _row in _rdr:
                    if len(_row) <= max(snap_dc, snap_sc):
                        continue
                    try:
                        _raw = _clean_date_cell(_row[snap_dc])
                        _d   = datetime.strptime(_raw, '%d/%m/%Y').date()
                        if wk <= _d <= wend:
                            _st = _row[snap_sc].strip()
                            if '\n' in _st:
                                _st = _st.split('\n')[-1].strip()
                            counts[_st] += _row_value(_row, snap_cfg)
                    except ValueError:
                        continue

        weekly_by_status[wk_str] = dict(counts)
        if cache_key and wk != curr_week:
            pivot_cache[cache_key] = dict(counts)
            cache_updated = True

    if cache_updated:
        _save_presc_pivot_cache(pivot_cache)

    # Assemble la réponse au même format que compute_pivot
    s1_ref = S1_MON if week_mode == 'mon' else S1_THU
    labels = [
        f"{_wk_label(w, s1_ref)}\n{w.strftime('%d/%m')}–{(w + timedelta(days=6)).strftime('%d/%m')}"
        for w in weeks
    ]
    datasets = []
    for status in cfg['status_order']:
        data = [weekly_by_status.get(str(w), {}).get(status, 0) for w in weeks]
        if any(v > 0 for v in data):
            datasets.append({
                'label':           status,
                'data':            data,
                'backgroundColor': cfg['colors'].get(status, '#888888'),
            })
    totals = [sum(weekly_by_status.get(str(w), {}).values()) for w in weeks]
    detail = [{s: weekly_by_status.get(str(w), {}).get(s, 0) for s in cfg['status_order']} for w in weeks]

    return {
        'labels':      labels,
        'datasets':    datasets,
        'totals':      totals,
        'detail':      detail,
        'weeks':       [str(w) for w in weeks],
        'total_weeks': len(all_weeks),
        'granularity': 'week',
        'date':        str(ref_date) if ref_date else None,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# BUDGET — fichier local uniquement (mode démo)
# ═══════════════════════════════════════════════════════════════════════════════
_budget_refresh_status: dict = {'ok': None, 'ts': None, 'msg': ''}

def _refresh_budget_excel():
    """Mode démo : pas de téléchargement distant, le fichier budget local (data/budget) fait foi."""
    _budget_refresh_status.update({'ok': True, 'ts': _now().isoformat(),
                                   'msg': 'Mode démo — fichier budget local'})

GSHEET_ID        = ''      # Google Sheets désactivé (mode démo)
MEDIA_CACHE_FILE = DATA_DIR / 'budget' / 'media_plan_cache.json'

_gsheet_cache: dict = {'data': None, 'ts': 0.0, 'n': 0}

def _parse_euro(v) -> float:
    """Convertit une valeur texte (€1 234,56 ou 1234.56 ou float) en float."""
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = (str(v)
         .replace('€', '')
         .replace('\u202f', '')   # espace fine insécable (séparateur milliers FR)
         .replace('\xa0', '')     # espace insécable
         .replace(' ', '')
         .replace(',', '.')
         .strip())
    try: return float(s)
    except ValueError: return 0.0

def _load_media_data(N_DISP):
    """Mode démo : données média lues uniquement depuis le JSON local."""
    return _load_media_json(N_DISP)

def _load_media_json(N_DISP):
    """Charge depuis le cache JSON local (mis à jour à chaque fetch GSheet réussi)."""
    import json as _json
    if not MEDIA_CACHE_FILE.exists():
        raise FileNotFoundError('Pas de plan média JSON')
    payload = _json.loads(MEDIA_CACHE_FILE.read_text(encoding='utf-8'))
    if payload.get('n') != N_DISP:
        raise ValueError(f"Cache JSON incompatible (n={payload.get('n')} vs {N_DISP})")
    nova_lines    = payload['nova_lines']
    portail_lines = payload['portail_lines']
    nova_total    = payload['nova_total']
    portail_total = payload['portail_total']
    logging.getLogger(__name__).info('Données média chargées depuis le cache JSON local')
    return nova_lines, portail_lines, nova_total, portail_total



# ═══════════════════════════════════════════════════════════════════════════════
# AT INTERNET
# ═══════════════════════════════════════════════════════════════════════════════
MAX_PER_PAGE = 1000

LEAD_URL   = 'https://pro.nova.example/inscription/id-connexion'
TRAFIC_URL = 'https://pro.nova.example/inscription/creation-compte/choix-situation-entreprise'
LP_URL     = 'l.nova.example/pro'
PMAX_URL      = 'pmax'
META_URL      = 'nova_q1_meta'
SEA_BRAND_URL     = ['lancement-nova-01_gg_search_M', 'at_variant=marque']
SEA_NON_BRAND_URL = ['lancement-nova-01_gg_search_M', 'at_variant=generique']
PRESC_URL         = 'prescription-nova-pro'
LINKEDIN_URL      = 'nova_q1_linkedin'
TIKTOK_URL        = 'nova_q1_tiktok'
AMAZON_URL        = 'nova_q1_amazon'
AFFILIATION_URL   = 'affiliation'
PARTENARIAT_URL   = 'partenariat'
EVENT_URL         = 'event'

# Filtres secondaires Piano pour les 4 sources Budget Média (partagés snapshot + API)
_BMEDIA_URL_FILTERS = {'lp': LP_URL, 'init': TRAFIC_URL, 'lead': LEAD_URL, 'conv': '__UTM_CONV__'}
_BMEDIA_SPF = {
    'meta': {'$or': [
        {'event_url_full': {'$lk': META_URL}},
        {'src_detail':     {'$eq': 'nova_q1_meta'}},
    ]},
    'pmax': {'$or': [
        {'event_url_full': {'$lk': PMAX_URL}},
        {'src_detail':     {'$eq': 'nova_q1_gg_pmax'}},
    ]},
    'sea_brand': {'$or': [
        {'$and': [{'event_url_full': {'$lk': 'lancement-nova-01_gg_search_M'}},
                  {'event_url_full': {'$lk': 'at_variant=marque'}}]},
        {'$and': [{'src_detail':     {'$eq': 'lancement-nova-01_gg_search_M'}},
                  {'src_variant':    {'$lk': 'marque'}}]},
    ]},
    'sea_nb': {'$or': [
        {'$and': [{'event_url_full': {'$lk': 'lancement-nova-01_gg_search_M'}},
                  {'event_url_full': {'$lk': 'at_variant=generique'}}]},
        {'$and': [{'src_detail':     {'$eq': 'lancement-nova-01_gg_search_M'}},
                  {'src_variant':    {'$lk': 'generique'}}]},
    ]},
    'linkedin': {'$or': [
        {'event_url_full': {'$lk': LINKEDIN_URL}},
        {'src_detail':     {'$eq': LINKEDIN_URL}},
    ]},
    'tiktok': {'$or': [
        {'event_url_full': {'$lk': TIKTOK_URL}},
        {'src_detail':     {'$eq': TIKTOK_URL}},
    ]},
    'amazon': {'$or': [
        {'event_url_full': {'$lk': AMAZON_URL}},
        {'src_detail':     {'$eq': AMAZON_URL}},
    ]},
    'affiliation': {'$or': [
        {'event_url_full': {'$lk': AFFILIATION_URL}},
        {'src':            {'$eq': AFFILIATION_URL}},
    ]},
    'partenariat': {'$or': [
        {'event_url_full': {'$lk': PARTENARIAT_URL}},
        {'src':            {'$eq': PARTENARIAT_URL}},
    ]},
    'event': {'$or': [
        {'event_url_full': {'$lk': EVENT_URL}},
        {'src':            {'$eq': EVENT_URL}},
    ]},
}

# Mapping des noms de médias (colonne B du sheet "Sanity Check per day") → clé _BMEDIA_SPF
_SANITY_MEDIA_MAP = {
    'meta':              'meta',
    'tiktok':            'tiktok',
    'linkedin':          'linkedin',
    'amazon display':    'amazon',
    'amazon':            'amazon',
    'performance max':   'pmax',
    'pmax':              'pmax',
    'sea - brand':       'sea_brand',
    'sea brand':         'sea_brand',
    'sea - non brand':   'sea_nb',
    'sea nb':            'sea_nb',
    'affiliation':       'affiliation',
    'partenariat':       'partenariat',
    'prescription':      'prescription',
    # Demand Gen, Google UAC, Apple Search Ads → pas de filtre Piano → None
}

# ── Filtres UTM pour les LEADS ────────────────────────────────────────────
_UTM_FILTERS = {
    'meta':        lambda u: u.get('at_channel') == 'meta',
    'pmax':        lambda u: u.get('at_channel') in ('pmax', 'pmaxmobile'),
    'tiktok':      lambda u: u.get('at_channel') == 'tiktok',
    'sea_brand':   lambda u: u.get('at_channel') == 'gg_search' and 'marque'    in (u.get('at_variant') or ''),
    'sea_nb':      lambda u: u.get('at_channel') == 'gg_search' and 'generique' in (u.get('at_variant') or ''),
    'linkedin':    lambda u: u.get('at_channel') == 'linkedin',
    'amazon':      lambda u: 'amazon' in (u.get('at_campaign') or '').lower() or u.get('at_channel') == 'amazon',
    'affiliation':   lambda u: u.get('at_medium') == 'affiliation',
    'partenariat':   lambda u: u.get('at_medium') == 'partenariat',
    'event':         lambda u: u.get('at_medium') == 'event',
    'prescription':  lambda u: u.get('at_campaign') == 'prescription-nova-pro',
    'organique':     lambda u: u.get('at_medium') == 'organic',
}

UTM_CACHE_DIR = DATA_DIR / 'utm_cache'

_MEDIA_CANAL_KEYS = ('meta', 'pmax', 'sea_brand', 'sea_nb', 'linkedin', 'tiktok', 'amazon')

def _classify_canal(u):
    """Classifie un dict UTM en canal : prescription/affiliation/partenariat/event/media/organique/direct."""
    if _UTM_FILTERS['prescription'](u):   return 'prescription'
    if _UTM_FILTERS['affiliation'](u):    return 'affiliation'
    if _UTM_FILTERS['partenariat'](u):    return 'partenariat'
    if _UTM_FILTERS['event'](u):          return 'event'
    for k in _MEDIA_CANAL_KEYS:
        if _UTM_FILTERS[k](u):            return 'media'
    if _UTM_FILTERS['organique'](u):      return 'organique'
    return 'direct'


def _utm_dedup_df():
    """Charge et déduplique ALL_UTM.csv par HASHED_USER_ID.

    Règles de déduplication :
      1. Pour chaque utilisateur, on prend la DERNIÈRE session marketing
         (AT_MEDIUM non null OU AF_PID != 'eer_redirect_mobile_app'),
         en excluant les sessions internes applicatives (organique, eer).
      2. AT_MEDIUM = 'organique' est traité comme 'eer_redirect_mobile_app' :
         ce sont des sessions de navigation interne à pro.nova.example sans UTM,
         qui ne doivent pas masquer une vraie attribution marketing préalable.
      3. Les sessions 'email-abandon' (AT_MEDIUM == 'email-abandon') sont
         EXCLUES du choix d'attribution — elles ne peuvent pas être la
         source attribuée. Si seules des sessions email-abandon (ou eer)
         existent, on conserve la dernière session non-abandon ; sinon
         la toute dernière.
      4. Si aucune session marketing non-abandon n'existe, on revient à
         la dernière session marketing (email-abandon ou autre) ; en
         dernier recours on prend la dernière session toutes catégories.

    Retourne un DataFrame dédupliqué (une ligne par utilisateur),
    avec les colonnes originales + _is_eer / _is_marketing / _is_abandon
    + session_dt."""
    import pandas as _pd

    csv_path = UTM_CACHE_DIR / 'ALL_UTM.csv'
    if not csv_path.exists():
        return _pd.DataFrame()

    df = _pd.read_csv(csv_path, low_memory=False)
    df['session_dt'] = _pd.to_datetime(df['SESSION_CREATED_AT'], utc=True, errors='coerce')

    af_pid = df['AF_PID'].fillna('')
    # 'organique' = navigation interne pro.nova.example, traité comme EER (non-acquisition)
    df['_is_eer']       = (af_pid == 'eer_redirect_mobile_app') | (df['AT_MEDIUM'] == 'organique')
    df['_is_marketing'] = (df['AT_MEDIUM'].notna() & (df['AT_MEDIUM'] != 'organique')) | (~df['_is_eer'] & (af_pid != ''))
    # Les relances sont AT_MEDIUM='emailing' avec AT_CAMPAIGN contenant 'email-abandon'
    df['_is_abandon']   = df['AT_CAMPAIGN'].fillna('').str.contains('email-abandon', case=False, na=False)

    # Sélection vectorisée (compatible pandas 1.x / 2.x / 3.x) — même logique que
    # l'ancien groupby().apply(_pick_best) :
    #   priorité 3 = dernière session marketing NON email-abandon
    #   priorité 2 = à défaut, dernière session marketing (email-abandon inclus)
    #   priorité 1 = à défaut, dernière session toutes catégories (EER / organique)
    df['_prio'] = 1
    df.loc[df['_is_marketing'], '_prio'] = 2
    df.loc[df['_is_marketing'] & ~df['_is_abandon'], '_prio'] = 3
    df_sorted = df.sort_values(['_prio', 'session_dt'], kind='mergesort')
    out = df_sorted.drop_duplicates('HASHED_USER_ID', keep='last')
    return out.drop(columns=['_prio']).reset_index(drop=True)


def _utm_valeur(row):
    """Montant € porté par l'utilisateur (colonne VALEUR d'ALL_UTM), 1.0 si absente."""
    try:
        v = row.get('VALEUR', None)
        if v is None:
            return 1.0
        s = str(v).strip().replace(',', '.')
        return float(s) if s and s not in ('nan', 'None', 'NaT') else 1.0
    except (TypeError, ValueError):
        return 1.0


def _load_utm_cache(conv=False):
    """Charge ALL_UTM.csv depuis utm_cache/ (format enrichi, colonnes structurées).

    Déduplication par HASHED_USER_ID via _utm_dedup_df() :
      - Sélectionne la DERNIÈRE entrée marketing, en excluant 'email-abandon'
        de l'attribution pour ne pas masquer les vraies sources préalables.

    Date de référence :
      - conv=False (LEAD)              : USER_LIFECYCLE_START_DATE
      - conv=True  (CONVERSION)        : COMPANY_CREATED_AT si USER_LIFECYCLE_STATUS='validated'
                                          STEP_UPDATED_AT    si USER_LIFECYCLE_STATUS='validated_awaiting_agent_confirmation'
                                          (seules ces deux valeurs de statut sont retournées)

    Retourne {date_iso: [utms_with_meta, ...]}."""
    import pandas as _pd
    import numpy as _np

    df = _utm_dedup_df()
    if df.empty:
        return {}

    # ── Date de référence selon le type d'indicateur ──────────────────────────────────
    def _parse_col(col):
        if col not in df.columns:
            return _pd.Series(['NaT'] * len(df), index=df.index)
        return _pd.to_datetime(df[col], utc=True, errors='coerce').dt.strftime('%Y-%m-%d')

    if conv:
        # Filtrer uniquement les lifecycle valides pour les conversions
        lc = df['USER_LIFECYCLE_STATUS'].fillna('') if 'USER_LIFECYCLE_STATUS' in df.columns else _pd.Series([''] * len(df), index=df.index)
        df = df[lc.isin(['validated', 'validated_awaiting_agent_confirmation'])].copy()
        if df.empty:
            return {}
        lc = df['USER_LIFECYCLE_STATUS'].fillna('')
        company_dt = _parse_col('COMPANY_CREATED_AT')
        step_dt    = _parse_col('STEP_UPDATED_AT')
        df['ref_date'] = _np.where(lc == 'validated', company_dt, step_dt)
    else:
        # LEAD : date = USER_LIFECYCLE_START_DATE
        df['ref_date'] = _parse_col('USER_LIFECYCLE_START_DATE')

    # ── Construction du dict par date de référence ────────────────────────────
    by_date = {}
    for _, row in df.iterrows():
        d = row['ref_date']
        if not d or d == 'NaT' or _pd.isna(d):
            continue
        def _str(v):
            s = str(v) if v is not None else ''
            return '' if s in ('nan', 'NaT', 'None') else s
        utms_raw = _str(row.get('UTMS_RAW', ''))
        u = {
            'at_channel':  _str(row.get('AT_CHANNEL',  '')),
            'at_medium':   _str(row.get('AT_MEDIUM',   '')),
            'at_campaign': _str(row.get('AT_CAMPAIGN', '')),
            'at_variant':  _str(row.get('AT_VARIANT',  '')),
            'at_creation': _str(row.get('AT_CREATION', '')),
            'pid':         _str(row.get('AF_PID',      '')),
            '_statut':     _str(row.get('USER_TYPE',   '')),
            '_lifecycle':  _str(row.get('USER_LIFECYCLE_STATUS', '')),
            'utms_raw':    utms_raw,
            '_valeur':     _utm_valeur(row),
        }
        by_date.setdefault(d, []).append(u)
    return by_date


# Condition CONVERSION : lifecycle validated OU validated_awaiting_agent_confirmation)
# Utilisé uniquement comme garde-fou ; le filtrage principal est fait dans _load_utm_cache(conv=True).
_CONV_CONDITION = lambda u: u.get('_lifecycle') in ('validated', 'validated_awaiting_agent_confirmation')


def _utm_to_piano_rows(start, end, conv_only=False):
    """Convertit les entrées UTM (après dédup) en lignes synthétiques au format Piano.
    Si conv_only=True, utilise le cache conversions (daté sur COMPANY_CREATED_AT / STEP_UPDATED_AT,
    filtré sur validated / validated_awaiting_agent_confirmation)."""
    utm_by_date = _load_utm_cache(conv=conv_only)
    rows = []
    for d, utms in utm_by_date.items():
        if not (start <= d <= end):
            continue
        for u in utms:
            ch       = u.get('at_channel', '') or ''
            variant  = u.get('at_variant', '') or ''
            campaign = u.get('at_campaign', '') or ''
            medium   = u.get('at_medium', '') or ''
            pid      = u.get('pid', '') or ''

            _v = u.get('_valeur', 1.0)
            row = {'date': d, 'm_visits': _v, 'm_unique_visitors': _v,
                   'src': '', 'src_detail': '', 'src_variant': '',
                   'src_campaign': campaign, 'src_creation': '', 'device_type': '',
                   '_lifecycle': u.get('_lifecycle', '')}

            if ch in ('pmax', 'pmaxmobile'):
                row['src_detail'] = 'nova_q1_gg_pmax'
            elif ch == 'meta':
                row['src_detail'] = 'nova_q1_meta'
            elif ch == 'gg_search':
                row['src_detail'] = 'lancement-nova-01_gg_search_M'
                row['src_variant'] = variant
            elif campaign == 'prescription-nova-pro':
                row['src_detail'] = 'prescription-nova-pro'
            elif medium == 'affiliation':
                row['src'] = AFFILIATION_URL
            elif medium == 'partenariat':
                row['src'] = PARTENARIAT_URL
            elif medium == 'event':
                row['src'] = EVENT_URL
            elif ch == 'tiktok':
                row['src_detail'] = TIKTOK_URL
            elif ch == 'linkedin':
                row['src_detail'] = LINKEDIN_URL
            elif 'amazon' in campaign.lower():
                row['src_detail'] = AMAZON_URL
            elif medium == 'organic':
                row['src'] = 'organique_web'             # at_medium=organic → organique web tracké
            elif medium == 'organique':
                row['src'] = 'organique_sans_tracking'   # navigation interne pro.nova.example, traité comme EER
            elif pid == 'eer_redirect_mobile_app':
                row['src'] = 'organique_sans_tracking'   # deep-link mobile sans UTM marketing
            elif not u.get('utms_raw'):
                row['src'] = 'organique_sans_tracking'   # aucun tracking UTM
            else:
                row['src'] = pid or 'autres'

            rows.append(row)
    return rows


def _col_idx_to_letter(col_0idx: int) -> str:
    """Convertit un index 0-based en lettre(s) de colonne Excel (A, B, ..., Z, AA, ...)."""
    result = ''
    n = col_0idx + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def _bmedia_combined_filter(src: str, ut: str) -> dict:
    """Retourne le filtre Piano combiné source+URL pour la fonctionnalité bmedia_sub."""
    url_cond = {'event_url_full': {'$lk': _BMEDIA_URL_FILTERS[ut]}}
    def expand(branch):
        if '$and' in branch:
            return {'$and': branch['$and'] + [url_cond]}
        return {'$and': [branch, url_cond]}
    return {'$or': [expand(b) for b in _BMEDIA_SPF[src]['$or']]}

_AT_URL_PARAMS = ('at_medium', 'at_campaign', 'at_channel', 'at_variant', 'at_creation')

def _parse_piano_url(url: str) -> dict:
    """Parse event_url_full en colonnes structurées.

    Retourne un dict avec :
      event_url_path      — URL sans query string (ex: https://pro.nova.example/inscription/...)
      at_medium / at_campaign / at_channel / at_variant / at_creation — params Piano AT
      tracking_platform   — 'google' | 'meta' | 'tiktok' | 'amazon' | None
      has_gclid / has_fbclid / has_ttclid — booléens (présence du click ID)
      mode                — 'signup' | 'login' | None
      client_id           — 'EPRO' | 'EPROPSD2FED' | None
    """
    if not url:
        return {}
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query, keep_blank_values=False)

        def _first(key):
            v = params.get(key)
            return v[0] if v else None

        pid = _first('pid') or ''
        if 'google' in pid:
            platform = 'google'
        elif 'facebook' in pid or 'fb' in pid:
            platform = 'meta'
        elif 'tiktok' in pid:
            platform = 'tiktok'
        elif 'amazon' in pid:
            platform = 'amazon'
        elif _first('gclid'):
            platform = 'google'
        elif _first('fbclid'):
            platform = 'meta'
        elif _first('ttclid'):
            platform = 'tiktok'
        else:
            platform = None

        path = (parsed.scheme + '://' + parsed.netloc + parsed.path) if parsed.netloc else parsed.path
        return {
            'event_url_path':    path or None,
            'at_medium':         _first('at_medium'),
            'at_campaign':       _first('at_campaign'),
            'at_channel':        _first('at_channel'),
            'at_variant':        _first('at_variant'),
            'at_creation':       _first('at_creation'),
            'at_aff_identifier': _first('at_aff_identifier'),
            'tracking_platform': platform,
            'has_gclid':         bool(_first('gclid')),
            'has_fbclid':        bool(_first('fbclid')),
            'has_ttclid':        bool(_first('ttclid')),
            'mode':              _first('mode'),
            'client_id':         _first('client_id'),
        }
    except Exception:
        return {}


def _row_url_str(row: dict) -> str:
    """Reconstruit une chaîne synthétique depuis les colonnes parsées.

    Utilisé par _eval_piano_filter pour maintenir la compatibilité des filtres
    {'event_url_full': {'$lk': ...}} sans stocker l'URL brute en cache.
    Format : '<path>?at_medium=<v>&at_campaign=<v>&at_channel=<v>&at_variant=<v>&at_creation=<v>'
    """
    parts = [row.get('event_url_path') or '']
    for k in _AT_URL_PARAMS:
        v = row.get(k)
        if v:
            parts.append(f'{k}={v}')
    return '&'.join(parts)


def _fetch_traffic_page(start: str, end: str, page: int, url_filter: str = None, secondary_property_filter: dict = None, include_url: bool = False, site_id=None) -> tuple:
    """Récupère une page de l'API AT Internet.
    Retourne (rows, total_count) où total_count peut être None si non fourni."""
    cols = ['src', 'src_detail', 'src_campaign', 'src_creation', 'src_variant',
            'device_type', 'date', 'visitor_privacy_consent', 'm_visits', 'm_unique_visitors']
    if include_url:
        cols = ['event_url_full'] + cols
    body = {
        'columns': cols,
        'sort':    ['date', 'src'],
        'space':   {'s': [site_id or AT_SITE]},
        'period':  {'p1': [{'type': 'D', 'start': start, 'end': end}]},
        'max-results': MAX_PER_PAGE,
        'page-num': page,
    }
    conditions = []
    if url_filter:
        conditions.append({'event_url_full': {'$lk': url_filter}})
    if secondary_property_filter:
        conditions.append(secondary_property_filter)
    if len(conditions) == 1:
        body['filter'] = {'property': conditions[0]}
    elif len(conditions) > 1:
        body['filter'] = {'property': {'$and': conditions}}
    payload = json.dumps(body).encode('utf-8')
    req = urllib.request.Request(
        AT_URL, data=payload,
        headers={'x-api-key': AT_KEY, 'Content-Type': 'application/json'},
        method='POST',
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"API HTTP {e.code}: {e.read().decode()}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Connexion: {e.reason}")

    feed = body.get('DataFeed', body)
    counts = feed.get('RowCounts', {})
    total  = counts.get('Total', None)

    rows = []
    for item in feed.get('Rows', []):
        inner = item.get('Rows', None)
        if isinstance(inner, list):
            rows.extend(inner)
        elif isinstance(item, dict) and 'src' in item:
            rows.append(item)
    return rows, total


def _fetch_traffic_segment(start: str, end: str, url_filter: str = None, secondary_property_filter: dict = None, include_url: bool = False, site_id=None) -> dict:
    """Fetch paginé pour UNE plage de dates. Retourne un dict de lignes dédupliquées (keyed).
    Mode démo : API Piano désactivée → aucune ligne (seul le cache local est utilisé)."""
    return {}
    all_rows = {}
    page     = 1
    MAX_PAGES = 500   # garde-fou : max 500 000 lignes

    while True:
        rows, total = _fetch_traffic_page(start, end, page, url_filter, secondary_property_filter, include_url, site_id)

        if page == 1 and total is not None:
            app.logger.info(
                f'[Piano] fetch start={start} end={end} '
                f'url_filter={url_filter!r} include_url={include_url} '
                f'→ total annoncé={total} lignes'
            )

        for r in rows:
            # Parsage de l'URL en colonnes structurées (effectué ici, avant mise en cache)
            parsed_url = _parse_piano_url(r.get('event_url_full', '') or '') if include_url else {}
            key = (
                r.get('date', ''),
                r.get('src', ''),
                r.get('src_detail', ''),
                r.get('src_campaign', ''),
                r.get('src_creation', ''),
                r.get('src_variant', ''),
                r.get('device_type', ''),
                r.get('visitor_privacy_consent', ''),
                parsed_url.get('event_url_path', '') if include_url else '',
                parsed_url.get('at_campaign', '') if include_url else '',
                parsed_url.get('at_variant', '') if include_url else '',
            )
            if key in all_rows:
                all_rows[key]['m_visits']          += r.get('m_visits', 0) or 0
                all_rows[key]['m_unique_visitors'] += r.get('m_unique_visitors', 0) or 0
            else:
                row_data = {
                    'date':                     r.get('date', ''),
                    'src':                      r.get('src', ''),
                    'src_detail':               r.get('src_detail', ''),
                    'src_campaign':             r.get('src_campaign', ''),
                    'src_creation':             r.get('src_creation', ''),
                    'src_variant':              r.get('src_variant', ''),
                    'device_type':              r.get('device_type', ''),
                    'visitor_privacy_consent':  r.get('visitor_privacy_consent', '') or '',
                    'm_visits':                 r.get('m_visits', 0) or 0,
                    'm_unique_visitors':        r.get('m_unique_visitors', 0) or 0,
                }
                if include_url:
                    # Stocker les colonnes parsées — jamais l'URL brute
                    row_data.update(parsed_url)
                all_rows[key] = row_data

        if len(rows) < MAX_PER_PAGE:
            break
        if total is not None and page * MAX_PER_PAGE >= total:
            break
        if page >= MAX_PAGES:
            app.logger.warning(
                f'[Piano] garde-fou atteint ({MAX_PAGES} pages = {MAX_PAGES * MAX_PER_PAGE} lignes max) '
                f'url_filter={url_filter!r} include_url={include_url} total_annoncé={total}'
            )
            break
        page += 1

    app.logger.info(
        f'[Piano] segment terminé : {len(all_rows)} lignes en {page} page(s) '
        f'(start={start} end={end} url_filter={url_filter!r})'
    )
    return all_rows


def _fetch_traffic(start: str, end: str, url_filter: str = None, secondary_property_filter: dict = None, include_url: bool = False, site_id=None) -> list:
    """Appelle l'API AT Internet et retourne toutes les lignes dédupliquées.

    Piano interdit les périodes > 48h incluant le jour en cours. Si end == aujourd'hui
    et start < aujourd'hui, on split en deux requêtes :
      1. start → hier  (données historiques, pas de limite de période)
      2. aujourd'hui   (données temps réel, 1 seul jour → respecte la limite 48h)
    Les résultats sont fusionnés avec déduplication.
    """
    today     = _now().date().isoformat()
    yesterday = (_now().date() - timedelta(days=1)).isoformat()

    if end == today and start < yesterday:
        # Période > 2 jours incluant aujourd'hui → Piano refuse (limite 48h realtime).
        # On split : historique (start → hier) + aujourd'hui séparé.
        # Si start == yesterday ou start == today → ≤ 48h → un seul appel suffit.
        app.logger.info(f'[Piano] split hist+today : {start}→{yesterday} + {today}')
        hist_rows  = _fetch_traffic_segment(start, yesterday, url_filter, secondary_property_filter, include_url, site_id)
        today_rows = _fetch_traffic_segment(today, today,     url_filter, secondary_property_filter, include_url, site_id)
        # Fusion : today_rows peut chevaucher hist_rows sur la même clé → on additionne
        for key, r in today_rows.items():
            if key in hist_rows:
                hist_rows[key]['m_visits']          += r.get('m_visits', 0) or 0
                hist_rows[key]['m_unique_visitors'] += r.get('m_unique_visitors', 0) or 0
            else:
                hist_rows[key] = r
        return list(hist_rows.values())
    else:
        return list(_fetch_traffic_segment(start, end, url_filter, secondary_property_filter, include_url, site_id).values())

# ═══════════════════════════════════════════════════════════════════════════════
# CACHE PIANO LOCAL
# ═══════════════════════════════════════════════════════════════════════════════
CACHE_DIR = DATA_DIR / 'piano_cache'

def _cache_path(site_id: int, date_str: str) -> Path:
    return CACHE_DIR / f's{site_id}_{date_str}.json'

def _load_day_cache(site_id: int, date_str: str):
    """Retourne la liste de lignes mise en cache pour ce jour, ou None.
    Invalide automatiquement le cache si la dimension visitor_privacy_consent
    est absente (anciens fichiers pre-consent) — le jour sera re-fetché."""
    p = _cache_path(site_id, date_str)
    if p.exists():
        try:
            with open(p, encoding='utf-8') as f:
                data = json.load(f)
            if data:
                first = data[0]
                # Invalider si champ consent absent (ancien format pre-consent)
                if 'visitor_privacy_consent' not in first:
                    app.logger.info(f'[Cache] invalidation consent site={site_id} {date_str}')
                    try: p.unlink()
                    except Exception: pass
                    return None
                # Invalider si event_url_full présent (ancien format pre-parse)
                if 'event_url_full' in first:
                    app.logger.info(f'[Cache] invalidation url-parse site={site_id} {date_str}')
                    try: p.unlink()
                    except Exception: pass
                    return None
            return data
        except Exception:
            try: p.unlink()
            except Exception: pass
    return None

def _save_day_cache(site_id: int, date_str: str, rows: list):
    """Écrit le cache atomiquement (via un .tmp pour éviter les fichiers corrompus)."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = _cache_path(site_id, date_str).with_suffix('.tmp')
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False)
    tmp.rename(_cache_path(site_id, date_str))

def _fetch_range_cached(start: str, end: str, site_id: int,
                        realtime_url_filter: str = None,
                        realtime_spf: dict = None) -> list:
    """Charge les données Piano jour par jour avec mise en cache locale.

    - Jours révolus (< aujourd'hui) : cache JSON sur disque si disponible, sinon appel API
      puis mise en cache pour les prochains appels.
    - Jour en cours : jamais mis en cache (données incomplètes).
      Si realtime_url_filter ou realtime_spf sont fournis, ils sont passés directement à
      l'API Piano pour réduire le volume téléchargé (optimisation perf).
    - Inclut toujours event_url_full (nécessaire pour la classification CORRIGÉ).
    """
    today    = _now().date().isoformat()
    start_dt = datetime.strptime(start, '%Y-%m-%d').date()
    end_dt   = datetime.strptime(end,   '%Y-%m-%d').date()
    all_rows = []
    current  = start_dt

    while current <= end_dt:
        # Mode démo : lecture exclusive du cache local (jour en cours inclus), jamais d'API
        cached = _load_day_cache(site_id, str(current))
        if cached:
            all_rows.extend(cached)
        current += timedelta(days=1)

    return all_rows

def _eval_piano_filter(row: dict, f: dict) -> bool:
    """Évalue localement un filtre Piano JSON ($or / $and / $lk / $eq).
    Permet de remplacer les filtres API par un filtrage Python sur les données cachées.

    Les filtres sur 'event_url_full' sont évalués via _row_url_str() pour maintenir
    la compatibilité avec les constantes existantes (META_URL, PMAX_URL…) maintenant
    que l'URL brute n'est plus stockée en cache.
    """
    if '$or' in f:
        return any(_eval_piano_filter(row, sub) for sub in f['$or'])
    if '$and' in f:
        return all(_eval_piano_filter(row, sub) for sub in f['$and'])
    for field, cond in f.items():
        if field == 'event_url_full':
            val = _row_url_str(row)
        else:
            val = str(row.get(field, '') or '')
        if '$lk' in cond and str(cond['$lk'] or '') not in val:
            return False
        if '$eq' in cond and val != str(cond['$eq'] or ''):
            return False
    return True

# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    from flask import make_response as _make_response
    resp = _make_response(render_template('index.html'))
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    return resp

@app.route('/api/status')
def api_status():
    with _lock:
        return jsonify({
            'running':        _export['running'],
            'done':           _export['done'],
            'chrome':         is_chrome_debug(),
            'log_len':        len(_export['log']),
            'budget_refresh': _budget_refresh_status,
        })

@app.route('/api/export/start', methods=['POST'])
def start_export():
    """Mode démo : l'export backoffice (scraping) est désactivé."""
    return jsonify({'error': 'Export backoffice désactivé en mode démo (données factices)'}), 400

@app.route('/api/export/stream')
def stream_log():
    def generate():
        sent = 0
        while True:
            with _lock:
                lines   = list(_export['log'])
                is_done = _export['done']
            while sent < len(lines):
                yield f"data: {json.dumps(lines[sent])}\n\n"
                sent += 1
            if is_done and sent >= len(lines):
                yield "data: __DONE__\n\n"
                break
            time.sleep(0.2)

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )

@app.route('/api/exports/<kind>')
def api_exports(kind):
    if kind not in CHART_CFG:
        return jsonify({'error': 'Type inconnu'}), 404
    cfg = CHART_CFG[kind]
    return jsonify({'dates': list_exports(cfg['prefix'], cfg['subdir'])})

@app.route('/api/chart/<kind>')
def api_chart(kind):
    if kind not in CHART_CFG:
        return jsonify({'error': 'Type inconnu'}), 404
    cfg         = CHART_CFG[kind]
    week_mode   = request.args.get('week_mode', 'mon')
    granularity = request.args.get('granularity', 'week')
    date_param  = request.args.get('date')
    week_offset = max(0, int(request.args.get('week_offset', 0) or 0))
    n_weeks     = max(1, min(52, int(request.args.get('n_weeks', 5) or 5)))

    # Prescriptions : pivot stable par snapshot de semaine (évite les changements rétroactifs)
    if kind == 'prescriptions' and granularity == 'week':
        try:
            pivot = compute_presc_pivot_stable(
                cfg, n_weeks=n_weeks, week_mode=week_mode,
                offset=week_offset, granularity='week', date_param=date_param,
            )
            if pivot is None:
                return jsonify({'error': 'Aucun fichier trouvé', 'date': None}), 404
            return jsonify(pivot)
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    if date_param:
        path, d = find_export_for_date(cfg['prefix'], cfg['subdir'], date_param)
    else:
        path, d = find_latest_export(cfg['prefix'], cfg['subdir'])
    if not path:
        return jsonify({'error': 'Aucun fichier trouvé', 'date': None}), 404
    try:
        pivot = compute_pivot(path, cfg, n_weeks=n_weeks, week_mode=week_mode,
                              offset=week_offset, granularity=granularity)
        pivot['date'] = str(d)
        return jsonify(pivot)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/stats/souscriptions/weekly_opened')
def api_stats_sous_weekly_opened():
    """Dossiers dont la date de création tombe dans la semaine sélectionnée (dernier fichier)."""
    week_mode   = request.args.get('week_mode', 'mon')
    week_fn     = week_start_mon if week_mode == 'mon' else week_start_thu
    week_offset = max(0, int(request.args.get('week_offset', 0) or 0))
    today       = _now().date()
    wk_start    = week_fn(today) - timedelta(weeks=week_offset)
    wk_end      = wk_start + timedelta(days=6)
    cfg      = CHART_CFG['souscriptions']
    prefix   = cfg['prefix']
    subdir   = cfg['subdir']
    # Vérifier qu'un snapshot existe AVANT le début de la semaine (sinon delta impossible)
    prev_end = wk_start - timedelta(days=1)
    _, prev_date = _find_file_on_or_before(prev_end, prefix, subdir)
    if prev_date is None:
        # Pas de snapshot antérieur → impossible de calculer un delta fiable
        return jsonify({
            'count':       None,
            'wk_start':    wk_start.strftime('%d/%m'),
            'wk_end':      wk_end.strftime('%d/%m'),
            'week_offset': week_offset,
        })
    results = _acquisition_by_file_diff([wk_start], cfg)
    count   = results[0] if results else 0
    results_n = _acquisition_by_file_diff([wk_start], cfg, mode='count')
    count_n   = results_n[0] if results_n else 0

    # Semaine précédente pour le delta %
    prev_wk_start = wk_start - timedelta(weeks=1)
    prev_wk_prev_end = prev_wk_start - timedelta(days=1)
    _, prev_prev_date = _find_file_on_or_before(prev_wk_prev_end, prefix, subdir)
    if prev_prev_date is not None:
        prev_results = _acquisition_by_file_diff([prev_wk_start], cfg)
        prev_count   = prev_results[0] if prev_results else None
        prev_count_n = (_acquisition_by_file_diff([prev_wk_start], cfg, mode='count') or [None])[0]
    else:
        prev_count = prev_count_n = None

    return jsonify({
        'count':       count,
        'count_n':     count_n,
        'prev_count':  prev_count,
        'prev_count_n': prev_count_n,
        'unit':        'eur',
        'week_label':  _wk_label(wk_start),
        'prev_label':  _wk_label(prev_wk_start),
        'wk_start':    wk_start.strftime('%d/%m'),
        'wk_end':      wk_end.strftime('%d/%m'),
        'week_offset': week_offset,
    })


@app.route('/api/stats/<kind>')
def api_stats(kind):
    """Retourne uniquement les comptages par statut (plus léger que /api/table)."""
    if kind not in CHART_CFG:
        return jsonify({'error': 'Type inconnu'}), 404
    cfg = CHART_CFG[kind]
    date_param = request.args.get('date')
    if date_param:
        path, d = find_export_for_date(cfg['prefix'], cfg['subdir'], date_param)
    else:
        path, d = find_latest_export(cfg['prefix'], cfg['subdir'])
    if not path:
        return jsonify({'error': 'Aucun fichier'}), 404
    counts   = defaultdict(float)
    counts_n = defaultdict(int)
    total    = 0.0
    total_n  = 0
    delim  = detect_delimiter(path)
    with open(path, encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=delim)
        for row in reader:
            statut = row.get('Statut', '').strip()
            if '\n' in statut:
                statut = statut.split('\n')[-1].strip()
            try:
                v = float(str(row.get('Valeur', '') or 1).replace(',', '.'))
            except ValueError:
                v = 1.0
            counts[statut] += v
            counts_n[statut] += 1
            total   += v
            total_n += 1
    return jsonify({'date': str(d), 'total': round(total, 2),
                    'counts': {k: round(v, 2) for k, v in counts.items()},
                    'total_n': total_n, 'counts_n': dict(counts_n), 'unit': 'eur'})

@app.route('/api/table/<kind>')
def api_table(kind):
    if kind not in CHART_CFG:
        return jsonify({'error': 'Type inconnu'}), 404
    cfg  = CHART_CFG[kind]
    path, d = find_latest_export(cfg['prefix'], cfg['subdir'])
    if not path:
        return jsonify({'error': 'Aucun fichier trouvé'}), 404

    rows = []
    with open(path, encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            r = dict(row)
            # Nettoyer "Voir les détails\n{id}" → garder seulement l'id
            for k, v in r.items():
                if isinstance(v, str) and '\n' in v:
                    r[k] = v.split('\n')[-1].strip()
            rows.append(r)

    return jsonify({'date': str(d), 'rows': rows, 'total': len(rows)})

@app.route('/api/chart/prospects_weekly')
def api_chart_prospects_weekly():
    """Volumes de prospects (hebdo ou quotidien) : prescriptions actives + marketing.
    marketing = souscriptions − prescriptions actives.
    metric=leads       : leads créés par semaine (défaut)
    metric=acquisitions: comptes ouverts (validés) par semaine via file_diff"""
    week_mode   = request.args.get('week_mode', 'mon')
    week_fn     = week_start_mon if week_mode == 'mon' else week_start_thu
    granularity = request.args.get('granularity', 'week')
    n_weeks     = max(1, min(52, int(request.args.get('n_weeks', 5) or 5)))
    week_offset = max(0, int(request.args.get('week_offset', 0) or 0))
    date_param  = request.args.get('date')
    metric      = request.args.get('metric', 'leads')   # 'leads' | 'acquisitions'

    sous_cfg  = CHART_CFG['souscriptions']
    presc_cfg = CHART_CFG['prescriptions']

    if date_param:
        sous_path,  d = find_export_for_date(sous_cfg['prefix'],  sous_cfg['subdir'],  date_param)
        presc_path, _ = find_export_for_date(presc_cfg['prefix'], presc_cfg['subdir'], date_param)
    else:
        sous_path,  d = find_latest_export(sous_cfg['prefix'],  sous_cfg['subdir'])
        presc_path, _ = find_latest_export(presc_cfg['prefix'], presc_cfg['subdir'])
    if not sous_path:
        return jsonify({'error': 'Aucun fichier souscriptions trouvé', 'date': None}), 404

    week1 = WEEK1_MON if week_mode == 'mon' else WEEK1_THU
    # Paramètre optionnel pour démarrer à une semaine spécifique (ex: S04)
    start_from = request.args.get('start_from')
    if start_from:
        try:
            week1 = max(week1, _date.fromisoformat(start_from))
        except ValueError:
            pass

    # ── Lecture des fichiers ──────────────────────────────────────────────────
    # Clé = week_start (mode semaine) ou date (mode jour)
    key_fn = (lambda d_: d_) if granularity == 'day' else week_fn

    _sous_fcfg = _cfg_for_file(sous_cfg, sous_path)

    keyed_sous        = defaultdict(float)
    keyed_presc_total = defaultdict(float)
    keyed_presc_skip  = defaultdict(float)

    if sous_path and sous_path.exists():
        delim = detect_delimiter(sous_path)
        with open(sous_path, encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=delim)
            next(reader, None)
            for row in reader:
                if len(row) <= sous_cfg['date_col']:
                    continue
                try:
                    raw = _clean_date_cell(row[sous_cfg['date_col']])
                    d_ = datetime.strptime(raw, '%d/%m/%Y').date()
                    keyed_sous[key_fn(d_)] += _row_value(row, _sous_fcfg)
                except ValueError:
                    continue

    if presc_path and presc_path.exists():
        _pp_cfg   = _cfg_for_file(presc_cfg, presc_path)
        delim = detect_delimiter(presc_path)
        with open(presc_path, encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=delim)
            next(reader, None)
            for row in reader:
                date_ci   = _pp_cfg['date_col']
                statut_ci = _pp_cfg['statut_col']
                if len(row) <= max(date_ci, statut_ci):
                    continue
                try:
                    raw = _clean_date_cell(row[date_ci])
                    d_ = datetime.strptime(raw, '%d/%m/%Y').date()
                    key = key_fn(d_)
                    keyed_presc_total[key] += _row_value(row, _pp_cfg)
                    statut = row[statut_ci].strip()
                    if '\n' in statut:
                        statut = statut.split('\n')[-1].strip()
                    if statut == 'EER non démarré':
                        keyed_presc_skip[key] += _row_value(row, _pp_cfg)
                except ValueError:
                    continue


    # ── Sélection des périodes à afficher ────────────────────────────────────
    if granularity == 'day':
        all_keys = sorted(set(list(keyed_sous.keys()) + list(keyed_presc_total.keys())))
        all_keys = [k for k in all_keys if week_fn(k) >= week1]
        if not all_keys:
            all_keys = [week1 + timedelta(days=i) for i in range(n_weeks * 7)]
        n_days = n_weeks * 7
        end_idx   = max(1, len(all_keys) - week_offset * 7)
        start_idx = max(0, end_idx - n_days)
        sel_keys  = all_keys[start_idx:end_idx]

        JOURS_FR = ['Dim', 'Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam']
        labels_out, week_labels, dates_out, week_ends_out = [], [], [], []
        presc_actives, marketing_data, totals_out = [], [], []
        for k in sel_keys:
            dow = JOURS_FR[k.weekday() + 1 if k.weekday() < 6 else 0]  # Mon=0→Lun
            dow = ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'][k.weekday()]
            labels_out.append(f'{dow}\n{k.strftime("%d/%m")}')
            week_labels.append(str(k))
            dates_out.append(str(k))
            week_ends_out.append(str(k))
            presc_act = max(0, keyed_presc_total.get(k, 0) - keyed_presc_skip.get(k, 0))
            mkt       = max(0, keyed_sous.get(k, 0) - presc_act)
            presc_actives.append(presc_act)
            marketing_data.append(mkt)
            totals_out.append(presc_act + mkt)
    else:
        all_keys = sorted(set(list(keyed_sous.keys()) + list(keyed_presc_total.keys())))
        all_keys = [w for w in all_keys if w >= week1]
        if not all_keys:
            all_keys = [week1 + timedelta(weeks=i) for i in range(n_weeks)]

        # Compléter avec toutes les semaines de week1 jusqu'à aujourd'hui
        # → les semaines sans données encore auront 0, et s'alimenteront automatiquement
        #   dès que le prochain export CSV couvrira ces dates
        curr_wk  = week_fn(_today())
        full_set = set(all_keys)
        wk = week1
        while wk <= curr_wk:
            full_set.add(wk)
            wk += timedelta(weeks=1)
        all_keys = sorted(full_set)

        end_idx   = max(1, len(all_keys) - week_offset)
        start_idx = max(0, end_idx - n_weeks)
        sel_keys  = all_keys[start_idx:end_idx]

        labels_out, week_labels, dates_out, week_ends_out = [], [], [], []
        presc_actives, marketing_data, totals_out = [], [], []
        s1_ref = S1_MON if week_mode == 'mon' else S1_THU

        if metric == 'acquisitions':
            # Ouvertures de compte : file_diff souscriptions et prescriptions
            total_acq = _acquisition_by_file_diff(sel_keys, sous_cfg)
            presc_acq = _acquisition_by_file_diff(sel_keys, presc_cfg)
            for i, wk in enumerate(sel_keys):
                lbl    = _wk_label(wk, s1_ref)
                wk_end = wk + timedelta(days=6)
                labels_out.append(f'{lbl}\n{wk.strftime("%d/%m")}→{wk_end.strftime("%d/%m")}')
                week_labels.append(lbl)
                dates_out.append(str(wk))
                week_ends_out.append(str(wk_end))
                t = total_acq[i] if i < len(total_acq) else 0
                p = presc_acq[i]  if i < len(presc_acq)  else 0
                mkt = max(0, t - p)
                presc_actives.append(p)
                marketing_data.append(mkt)
                totals_out.append(t)
        else:
            # Leads : prescriptions activées + marketing (souscriptions hors presc)
            curr_week  = week_fn(d) if d else None
            presc_acts = _presc_activations_by_week(sel_keys, presc_cfg, curr_week=curr_week)
            for i, wk in enumerate(sel_keys):
                lbl    = _wk_label(wk, s1_ref)
                wk_end = wk + timedelta(days=6)
                labels_out.append(f'{lbl}\n{wk.strftime("%d/%m")}→{wk_end.strftime("%d/%m")}')
                week_labels.append(lbl)
                dates_out.append(str(wk))
                week_ends_out.append(str(wk_end))
                presc_act = presc_acts[i] if i < len(presc_acts) else 0
                mkt       = max(0, keyed_sous.get(wk, 0) - presc_act)
                presc_actives.append(presc_act)
                marketing_data.append(mkt)
                totals_out.append(presc_act + mkt)

    ds_presc_label = 'Agences'
    ds_mkt_label   = 'Hors Agence' if metric == 'acquisitions' else 'Digital'

    return jsonify({
        'labels':      labels_out,
        'datasets': [
            {'label': ds_presc_label, 'data': presc_actives,  'backgroundColor': '#f39c12', 'stack': 'p'},
            {'label': ds_mkt_label,   'data': marketing_data, 'backgroundColor': '#3245c1', 'stack': 'p'},
        ],
        'totals':      totals_out,
        'weeks':       week_labels,
        'dates':       dates_out,
        'week_ends':   week_ends_out,
        'date':        str(d) if d else None,
        'granularity': granularity,
        'metric':      metric,
    })


@app.route('/api/chart/souscriptions_target')
def api_chart_souscriptions_target():
    """Compare dossiers validés (souscriptions) hebdomadaires vs objectifs (Thu→Wed)."""
    dist_file = DATA_DIR / 'projections' / 'Nova - distribution_volume_cible.xlsx'
    if not dist_file.exists() or openpyxl is None:
        return jsonify({'error': 'Fichier de distribution introuvable'}), 404

    # ── Lecture des objectifs (Volume Validé col3, Volume EER col5) ──────────
    wb  = openpyxl.load_workbook(dist_file, data_only=True)
    ws  = wb['Distribution Volume']
    target_weeks = []   # (date_start: date, label: str, vol_valide: int, vol_eer: int)
    for row in list(ws.iter_rows(values_only=True))[1:]:
        if not row[0] or str(row[0]).strip().upper() == 'TOTAL':
            break
        try:
            week_num = int(row[0])
            ds = row[1]
            if hasattr(ds, 'date'):
                ds = ds.date()
            elif isinstance(ds, str):
                ds = datetime.strptime(ds.split(' ')[0], '%Y-%m-%d').date()
            else:
                continue
            # Objectifs saisis en nombre de contrats → convertis dans l'unité affichée (€)
            vol_valide = int(round(float(row[3]) * VALEUR_MOYENNE)) if isinstance(row[3], (int, float)) else 0
            # col F (index 5) = Volume Prospect, souvent une formule non cachée → recalcul depuis D/E
            if isinstance(row[5], (int, float)) and row[5]:
                vol_eer = round(float(row[5]) * VALEUR_MOYENNE)
            elif isinstance(row[3], (int, float)) and isinstance(row[4], (int, float)) and row[4]:
                vol_eer = round(float(row[3]) / float(row[4]) * VALEUR_MOYENNE)
            else:
                vol_eer = 0
            target_weeks.append((ds, f'S{week_num}', vol_valide, vol_eer))
        except (ValueError, TypeError):
            continue
    if not target_weeks:
        return jsonify({'error': 'Aucun objectif trouvé dans le fichier'}), 404

    # ── Mode semaine : Jeu→Mer (thu) ou Lun→Dim (mon) ───────────────────────
    week_mode   = request.args.get('week_mode', 'mon')
    week_fn     = week_start_mon if week_mode == 'mon' else week_start_thu
    # Les semaines Excel démarrent mardi : Thu mode = ds+2j (jeudi), Mon mode = ds-1j (lundi)
    week_offset = timedelta(days=-1) if week_mode == 'mon' else timedelta(days=2)

    # ── Actuals : date 'Finalisée le' (ou file-diff si colonne absente) ─────────
    cfg      = CHART_CFG['souscriptions']
    date_param = request.args.get('date')
    max_date   = None
    if date_param:
        try:
            max_date = datetime.strptime(date_param, '%Y-%m-%d').date()
        except ValueError:
            pass
    week_keys    = [ds + week_offset for ds, _, _, _ in target_weeks]
    actuals_list = _acquisition_by_file_diff(week_keys, cfg, max_date=max_date)

    # ── Totals Prospect : comptage par date de création (plus précis que file-diff) ──
    # Utilise le fichier d'export sélectionné (ou le plus récent) et groupe par
    # date de création → évite les dossiers à saisie tardive mal assignés par le diff.
    ref_path = None
    if max_date:
        ref_path, _ = _find_file_on_or_before(max_date, cfg['prefix'], cfg['subdir'])
    if ref_path is None:
        ref_path, _ = find_latest_export(cfg['prefix'], cfg['subdir'])

    totals_by_week = defaultdict(float)
    _ref_cfg = _cfg_for_file(cfg, ref_path) if ref_path else cfg
    if ref_path and ref_path.exists():
        delim = detect_delimiter(ref_path)
        with open(ref_path, encoding='utf-8') as _fh:
            _rdr = csv.reader(_fh, delimiter=delim)
            next(_rdr, None)
            for _row in _rdr:
                if len(_row) <= cfg['date_col']:
                    continue
                try:
                    _raw = _clean_date_cell(_row[cfg['date_col']])
                    _d   = datetime.strptime(_raw, '%d/%m/%Y').date()
                    totals_by_week[week_fn(_d)] += _row_value(_row, _ref_cfg)
                except ValueError:
                    continue

    # ── Assemblage + calcul quadrimestres ────────────────────────────────────
    today      = _now().date()
    cur_week   = week_fn(today)
    week1      = WEEK1_MON if week_mode == 'mon' else WEEK1_THU
    labels, targets, eer_targets, actuals, totals, dates, week_ends = [], [], [], [], [], [], []
    current_idx = None

    for i, (ds, _label, vol_valide, vol_eer) in enumerate(target_weeks):
        week_key = ds + week_offset
        week_end = week_key + timedelta(days=6)
        # Numéro de semaine aligné sur le mode choisi (même référence que le reste du dashboard)
        s1_ref = S1_MON if week_mode == 'mon' else S1_THU
        labels.append(_wk_label(week_key, s1_ref))
        targets.append(vol_valide)
        eer_targets.append(vol_eer)
        actuals.append(actuals_list[i])
        totals.append(totals_by_week.get(week_key, 0))
        dates.append(str(week_key))       # début réel de la semaine (Thu ou Mon)
        week_ends.append(str(week_end))   # fin réelle de la semaine (Wed ou Sun)
        if week_key == cur_week:
            current_idx = i

    last_actual = max((i for i, v in enumerate(actuals) if v > 0), default=-1)

    # Indices par quadrimestre (basé sur le mois du début de semaine réel)
    quarters = {'Q1': [], 'Q2': [], 'Q3': []}
    for i, wk_start_str in enumerate(dates):
        m = datetime.strptime(wk_start_str, '%Y-%m-%d').month
        if m <= 4:
            quarters['Q1'].append(i)
        elif m <= 8:
            quarters['Q2'].append(i)
        else:
            quarters['Q3'].append(i)

    return jsonify({
        'labels':      labels,
        'dates':       dates,
        'week_ends':   week_ends,
        'targets':     targets,
        'eer_targets': eer_targets,
        'actuals':     actuals,
        'totals':      totals,
        'current_idx': current_idx,
        'last_actual': last_actual,
        'quarters':    quarters,
    })

@app.route('/api/chart/prescriptions_target')
def api_chart_prescriptions_target():
    """EER validés (comptes ouverts via prescriptions) par semaine — file diff method.
    Utilise la même structure de semaines que souscriptions_target."""
    dist_file = DATA_DIR / 'projections' / 'Nova - distribution_volume_cible.xlsx'
    if not dist_file.exists() or openpyxl is None:
        return jsonify({'error': 'Fichier de distribution introuvable'}), 404

    # ── Lecture de la structure des semaines depuis le fichier distribution ────
    wb = openpyxl.load_workbook(dist_file, data_only=True)
    ws = wb['Distribution Volume']
    target_weeks = []
    for row in list(ws.iter_rows(values_only=True))[1:]:
        if not row[0] or str(row[0]).strip().upper() == 'TOTAL':
            break
        try:
            week_num = int(row[0])
            ds = row[1]
            if hasattr(ds, 'date'):
                ds = ds.date()
            elif isinstance(ds, str):
                ds = datetime.strptime(ds.split(' ')[0], '%Y-%m-%d').date()
            else:
                continue
            target_weeks.append((ds, f'S{week_num}'))
        except (ValueError, TypeError):
            continue
    if not target_weeks:
        return jsonify({'error': 'Aucune semaine trouvée'}), 404

    # ── Mode semaine ──────────────────────────────────────────────────────────
    week_mode   = request.args.get('week_mode', 'mon')
    week_fn     = week_start_mon if week_mode == 'mon' else week_start_thu
    week_offset = timedelta(days=-1) if week_mode == 'mon' else timedelta(days=2)

    # ── Actuals EER validés : même file-diff que souscriptions ───────────────
    cfg        = CHART_CFG['prescriptions']
    date_param = request.args.get('date')
    max_date   = None
    if date_param:
        try:
            max_date = datetime.strptime(date_param, '%Y-%m-%d').date()
        except ValueError:
            pass

    week_keys    = [ds + week_offset for ds, _ in target_weeks]
    actuals_list = _acquisition_by_file_diff(week_keys, cfg, max_date=max_date)
    # Total souscriptions validées (Dossier validé) pour calcul non-prescriptions
    sous_list    = _acquisition_by_file_diff(week_keys, CHART_CFG['souscriptions'], max_date=max_date)

    # ── Marketing : média + affiliation + partenariat (UTM conversions) ─────────────
    _MARKETING_CANAUX = ('media', 'affiliation', 'partenariat', 'event')
    utm_conv = _load_utm_cache(conv=True)
    def _marketing_for_week(wstart, wend):
        count = 0
        wstart_str = str(wstart)
        wend_str   = str(wend)
        for date_str, utms in utm_conv.items():
            if not (wstart_str <= date_str <= wend_str):
                continue
            for u in utms:
                if _classify_canal(u) in _MARKETING_CANAUX:
                    count += u.get('_valeur', 1.0)
        return count

    # ── Assemblage ────────────────────────────────────────────────────────────
    today       = _now().date()
    cur_week    = week_fn(today)
    s1_ref      = S1_MON if week_mode == 'mon' else S1_THU
    labels, actuals, sous_actuals, marketing_actuals, dates, week_ends = [], [], [], [], [], []
    current_idx = None

    for i, (ds, _label) in enumerate(target_weeks):
        week_key = ds + week_offset
        week_end = week_key + timedelta(days=6)
        labels.append(_wk_label(week_key, s1_ref))
        actuals.append(actuals_list[i] if i < len(actuals_list) else 0)
        sous_actuals.append(sous_list[i] if i < len(sous_list) else 0)
        marketing_actuals.append(_marketing_for_week(week_key, week_end))
        dates.append(str(week_key))
        week_ends.append(str(week_end))
        if week_key == cur_week:
            current_idx = i

    last_actual = max((i for i, v in enumerate(actuals) if v > 0), default=-1)

    quarters = {'Q1': [], 'Q2': [], 'Q3': []}
    for i, wk_start_str in enumerate(dates):
        m = datetime.strptime(wk_start_str, '%Y-%m-%d').month
        if m <= 4:
            quarters['Q1'].append(i)
        elif m <= 8:
            quarters['Q2'].append(i)
        else:
            quarters['Q3'].append(i)

    return jsonify({
        'labels':       labels,
        'dates':        dates,
        'week_ends':    week_ends,
        'actuals':            actuals,            # EER validés (prescriptions converties)
        'sous_actuals':       sous_actuals,       # Total Dossier validé (toutes sources)
        'marketing_actuals':  marketing_actuals,  # Média + affiliation + partenariat (UTM conversions)
        'current_idx':  current_idx,
        'last_actual':  last_actual,
        'quarters':     quarters,
    })


@app.route('/api/chart/attribution_daily')
def api_chart_attribution_daily():
    """Attribution comptes ouverts — granularité journalière via UTM conversions.

    Retourne les N derniers jours (param ?days=, défaut 30) classifiés en :
      - prescription  : classify_canal == 'prescription'
      - marketing     : classify_canal in (media, affiliation, partenariat)
      - hors_presc    : total - prescription - marketing
    """
    n_days = int(request.args.get('days', 30))
    today  = _now().date()

    utm_conv = _load_utm_cache(conv=True)   # {date_iso: [utms...]}

    _MARKETING = ('media', 'affiliation', 'partenariat')

    dates_out, labels_out = [], []
    actuals_out, sous_out, mkt_out = [], [], []
    current_idx = None

    for i in range(n_days - 1, -1, -1):
        d = today - timedelta(days=i)
        d_str = str(d)
        utms = utm_conv.get(d_str, [])

        presc = sum(u.get('_valeur', 1.0) for u in utms if _classify_canal(u) == 'prescription')
        mkt   = sum(u.get('_valeur', 1.0) for u in utms if _classify_canal(u) in _MARKETING)
        total = sum(u.get('_valeur', 1.0) for u in utms)

        dates_out.append(d_str)
        labels_out.append(f"{d.day:02d}/{d.month:02d}")
        actuals_out.append(presc)
        sous_out.append(total)
        mkt_out.append(mkt)
        if d == today:
            current_idx = len(dates_out) - 1

    last_actual = max((i for i, v in enumerate(sous_out) if v > 0), default=-1)

    return jsonify({
        'labels':            labels_out,
        'dates':             dates_out,
        'actuals':           actuals_out,
        'sous_actuals':      sous_out,
        'marketing_actuals': mkt_out,
        'current_idx':       current_idx,
        'last_actual':       last_actual,
    })


@app.route('/api/budget/refresh', methods=['POST'])
def api_budget_refresh():
    """Mode démo : confirme l’usage du fichier budget local."""
    _refresh_budget_excel()
    return jsonify(_budget_refresh_status)

@app.route('/api/budget/export')
def api_budget_export():
    """Exporte tous les budgets marketing & média en Excel, par dépense / catégorie / type / semaine."""
    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé'}), 500
    if not BUDGET_EXCEL.exists():
        return jsonify({'error': 'Fichier budget introuvable'}), 404

    from io import BytesIO
    from flask import send_file
    from datetime import date as _date

    # Inclure S2 (col C, idx 2) = Jan 5, première semaine du fichier
    DISP_START = _date(2026, 4, 27)  # S18 = colonne « pré-affichage » du fichier budget factice
    N_DISP_MEDIA = _budget_n_disp()          # semaines S3..S25 (comme le dashboard)
    N_DISP       = N_DISP_MEDIA + 1          # +1 pour inclure S2 dans les non-média
    week_starts = [DISP_START + timedelta(weeks=i) for i in range(N_DISP)]
    week_labels = [_wk_label(ws) for ws in week_starts]

    # ── 1. Dépenses non-média depuis Suivi hebdo FULL ─────────────────────────
    wb_src  = openpyxl.load_workbook(BUDGET_EXCEL, data_only=True)
    ws_src  = wb_src['Suivi hebdo FULL']
    src_rows = list(ws_src.iter_rows(values_only=True))

    # Lignes de dépenses : rows index 4-13 (0-based), col 0=label, 1=type, 2=S2, 3=S3, ...
    # bci = di + 2  (col C idx 2 = S2 = di 0)
    _SKIP = {'TOTAL', 'Cumul Dépense', 'Media Nova', 'Media Portail'}
    expense_rows = []          # (dépense, catégorie, type, plateforme, [vals])
    for ri in range(4, len(src_rows)):
        row   = src_rows[ri]
        label = str(row[0]).strip() if row[0] else ''
        typ   = str(row[1]).strip() if row[1] else ''
        if not label or not typ or label in _SKIP:
            continue
        vals = []
        for di in range(N_DISP):
            ci = di + 2          # col C (idx 2) = S2 → di=0, col D (idx 3) = S3 → di=1, …
            v  = row[ci] if ci < len(row) else None
            vals.append(round(float(v), 2) if isinstance(v, (int, float)) else 0.0)
        if label == 'Média':
            continue              # Média traité séparément ci-dessous
        # Dépense = Catégorie, Service = "Nova" pour toutes les lignes non-média
        expense_rows.append((label, label, typ, 'Nova', vals))

    # ── 2. Média par plateforme (Nova + Portail Entrepreneur) ─────────
    # Dépense = Plateforme, Catégorie = Média, Type = Perf, Service = produit
    # _load_media_data(N_DISP_MEDIA) retourne des valeurs alignées S3..S25
    # On préfixe un 0 pour S2 afin d'aligner avec les labels S2..S25 de l'export
    try:
        nova_lines, portail_lines, _, _ = _load_media_data(N_DISP_MEDIA)
    except Exception:
        nova_lines, portail_lines = [], []

    for line in nova_lines:
        plat     = line.get('platform', '?')
        vals_raw = [round(float(v), 2) if v else 0.0 for v in line.get('values', [0]*N_DISP_MEDIA)]
        vals_raw = (vals_raw + [0.0]*N_DISP_MEDIA)[:N_DISP_MEDIA]
        vals     = [0.0] + vals_raw   # S2=0 (pas de média S2), puis S3..S25
        expense_rows.append((plat, 'Média', 'Perf', 'Nova', vals))

    for line in portail_lines:
        plat     = line.get('platform', '?')
        vals_raw = [round(float(v), 2) if v else 0.0 for v in line.get('values', [0]*N_DISP_MEDIA)]
        vals_raw = (vals_raw + [0.0]*N_DISP_MEDIA)[:N_DISP_MEDIA]
        vals     = [0.0] + vals_raw   # S2=0, puis S3..S25
        expense_rows.append((plat, 'Média', 'Perf', 'Portail Entrepreneur', vals))

    # ── 3. Construction du classeur Excel ────────────────────────────────────
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Budgets par semaine'

    # Styles communs aux exports du dashboard
    st       = _excel_export_styles()
    HDR_BG   = st['HDR_BG'];  HDR_FNT = st['HDR_FNT']
    TTL_BG   = st['TTL_BG']
    TOT_BG   = st['TOT_BG'];  TOT_FNT = st['TOT_FNT']
    ALT_BG   = st['ALT_BG'];  LED_BG  = st['LED_BG']
    THIN     = st['THIN']
    C_LEFT   = st['LFT']
    C_CENTER = st['CTR']
    from openpyxl.styles import Font as _Font, Alignment as _Align
    C_RIGHT  = _Align(horizontal='right', vertical='center')

    def _h(r, c, v):
        x = ws_out.cell(r, c, v)
        x.fill = HDR_BG; x.font = HDR_FNT; x.alignment = C_CENTER; x.border = THIN

    def _d(r, c, v, fill=None, bold=False, align=None):
        x = ws_out.cell(r, c, v)
        if fill: x.fill = fill
        x.font = _Font(bold=bold, size=10)
        x.alignment = align or C_LEFT; x.border = THIN

    def _tot(r, c, v):
        x = ws_out.cell(r, c, v)
        x.fill = TOT_BG; x.font = TOT_FNT; x.alignment = C_CENTER; x.border = THIN

    # ── Titre ───────────────────────────────────────────────────────────────
    n_cols = 4 + N_DISP + 1
    from openpyxl.utils import get_column_letter
    ws_out.merge_cells(f'A1:{get_column_letter(n_cols)}1')
    t = ws_out['A1']
    t.value = f'Budgets Digital & Média — par dépense / semaine  (extrait {_now().strftime("%d/%m/%Y")})'
    t.fill = TTL_BG; t.font = _Font(bold=True, color='00D4FF', size=13); t.alignment = C_LEFT
    ws_out.row_dimensions[1].height = 26

    # ── En-têtes ────────────────────────────────────────────────────────────
    FIXED_COLS = ['Dépense', 'Catégorie', 'Type', 'Service']
    col_headers = FIXED_COLS + week_labels + ['TOTAL']
    for ci, h in enumerate(col_headers, 1):
        _h(2, ci, h)
    ws_out.row_dimensions[2].height = 16

    # Largeurs colonnes
    ws_out.column_dimensions['A'].width = 28
    ws_out.column_dimensions['B'].width = 18
    ws_out.column_dimensions['C'].width = 14
    ws_out.column_dimensions['D'].width = 22
    for ci in range(5, 5 + N_DISP + 1):
        ws_out.column_dimensions[get_column_letter(ci)].width = 10

    # ── Données ─────────────────────────────────────────────────────────────
    row_idx        = 3
    totals_by_week = [0.0] * N_DISP

    for (depense, categorie, typ, service, vals) in expense_rows:
        is_media = (categorie == 'Média')
        fill = LED_BG if is_media else (ALT_BG if row_idx % 2 == 0 else None)

        _d(row_idx, 1, depense,   fill=fill, align=C_LEFT)
        _d(row_idx, 2, categorie, fill=fill, align=C_LEFT)
        _d(row_idx, 3, typ,       fill=fill, align=C_LEFT)
        _d(row_idx, 4, service,   fill=fill, align=C_LEFT)

        row_total = 0.0
        for di, v in enumerate(vals):
            _d(row_idx, 5 + di, round(v, 2) if v else None, fill=fill, align=C_RIGHT)
            row_total += v or 0.0
            totals_by_week[di] += v or 0.0

        # Colonne TOTAL de la ligne : en gras
        _d(row_idx, 5 + N_DISP, round(row_total, 2), fill=fill, bold=True, align=C_RIGHT)
        row_idx += 1

    # ── Ligne TOTAL ─────────────────────────────────────────────────────────
    _tot(row_idx, 1, 'TOTAL')
    _tot(row_idx, 2, '')
    _tot(row_idx, 3, '')
    _tot(row_idx, 4, '')
    grand_total = 0.0
    for di, v in enumerate(totals_by_week):
        _tot(row_idx, 5 + di, round(v, 2) if v else None)
        grand_total += v
    _tot(row_idx, 5 + N_DISP, round(grand_total, 2))
    ws_out.row_dimensions[row_idx].height = 18

    # Figer titre + en-têtes + 4 premières colonnes
    ws_out.freeze_panes = ws_out['E3']

    # ── Export ───────────────────────────────────────────────────────────────
    buf = BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    fname = f"budgets_marketing_{_now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True,
                     download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/budget')
def api_budget():
    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé — pip install openpyxl'}), 500

    budget_file = BUDGET_EXCEL
    if not budget_file.exists():
        return jsonify({'error': 'Fichier budget introuvable'}), 404

    # ── Semaines d'affichage : Lun→Dim, S-1=12jan, S1=19jan (référence officielle) ──
    # Mapping colonnes Excel : budget_file_col = display_idx + 3
    from datetime import date as _date
    DISP_START = _date(2026, 5, 4)   # lundi
    N_DISP     = _budget_n_disp()
    week_starts = [DISP_START + timedelta(weeks=i) for i in range(N_DISP)]
    week_labels = [_wk_label(ws) for ws in week_starts]

    # ── Lecture du fichier budget ────────────────────────────────────────────
    wb   = openpyxl.load_workbook(budget_file, data_only=True)
    ws   = wb['Suivi hebdo FULL']
    rows = list(ws.iter_rows(values_only=True))

    # Lignes de dépenses (idx 4-14). Budget col = display_idx + 3
    # La ligne "Média" sera remplacée par les totaux du media plan.
    expenses = []
    for ri in range(4, 15):
        if ri >= len(rows):
            break
        row   = rows[ri]
        label = row[0] if len(row) > 0 else None
        typ   = row[1] if len(row) > 1 else None
        if not label or not typ:
            continue
        vals = []
        for di in range(N_DISP):
            bci = di + 3                    # budget file column index
            v   = row[bci] if bci < len(row) else None
            vals.append(float(v) if isinstance(v, (int, float)) else 0.0)
        expenses.append({'label': str(label), 'type': str(typ), 'values': vals})

    # Acquisition : calculé depuis les CSV souscriptions (plus fiable que la saisie Excel)
    # acquisition sera affecté après le calcul de souscriptions_w ci-dessous

    # ── Total média global (Nova + Portail Entrepreneur) ─────────────
    _, _, nova_total, portail_total = _load_media_data(N_DISP)
    global_media_total = [nova_total[i] + portail_total[i] for i in range(N_DISP)]
    # Remplace les valeurs de la ligne "Média" par le total global
    for exp in expenses:
        if exp['label'] == 'Média':
            exp['values'] = global_media_total
            break

    # Totaux Perf / Full (recalculés après remplacement Média)
    totals_perf = [0.0] * N_DISP
    totals_full = [0.0] * N_DISP
    for exp in expenses:
        for i, v in enumerate(exp['values']):
            if exp['type'] == 'Perf':
                totals_perf[i] += v
            totals_full[i] += v

    # ── CSV : souscriptions et prescriptions par semaine Jeu→Mer ────────────
    def _count_weeks_and_status(csv_path, date_col_idx, statut_col_idx, skip_status, mode='eur'):
        """Retourne (montants_total, montants_skip_status) par semaine d'affichage."""
        total_c = [0.0] * N_DISP
        skip_c  = [0.0] * N_DISP
        if not csv_path or not csv_path.exists():
            return total_c, skip_c
        _bcfg_t = {'valeur_col': _valeur_col_of(csv_path)}
        try:
            with open(csv_path, encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                next(reader, None)
                for row in reader:
                    max_col = max(date_col_idx, statut_col_idx)
                    if len(row) <= max_col or not row[date_col_idx]:
                        continue
                    try:
                        d = datetime.strptime(
                            _clean_date_cell(row[date_col_idx]), '%d/%m/%Y'
                        ).date()
                        statut = row[statut_col_idx].strip() if statut_col_idx < len(row) else ''
                        if '\n' in statut:
                            statut = statut.split('\n')[-1].strip()
                        for wi, wstart in enumerate(week_starts):
                            if wstart <= d <= wstart + timedelta(days=6):
                                total_c[wi] += _row_value(row, _bcfg_t, mode)
                                if statut == skip_status:
                                    skip_c[wi] += _row_value(row, _bcfg_t, mode)
                                break
                    except ValueError:
                        pass
        except Exception:
            pass
        return total_c, skip_c

    def _count_weeks(csv_path, date_col_idx, mode='eur'):
        total_c, _ = _count_weeks_and_status(csv_path, date_col_idx, date_col_idx, '', mode)
        return total_c

    def _count_weeks_keep(csv_path, date_col_idx, statut_col_idx, keep_status):
        """Somme les montants des lignes dont le statut correspond à keep_status."""
        counts = [0.0] * N_DISP
        if not csv_path or not csv_path.exists():
            return counts
        _bcfg_k = {'valeur_col': _valeur_col_of(csv_path)}
        try:
            with open(csv_path, encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                next(reader, None)
                for row in reader:
                    max_col = max(date_col_idx, statut_col_idx)
                    if len(row) <= max_col or not row[date_col_idx]:
                        continue
                    try:
                        d = datetime.strptime(
                            _clean_date_cell(row[date_col_idx]), '%d/%m/%Y'
                        ).date()
                        statut = row[statut_col_idx].strip() if statut_col_idx < len(row) else ''
                        if '\n' in statut:
                            statut = statut.split('\n')[-1].strip()
                        if statut != keep_status:
                            continue
                        for wi, wstart in enumerate(week_starts):
                            if wstart <= d <= wstart + timedelta(days=6):
                                counts[wi] += _row_value(row, _bcfg_k)
                                break
                    except ValueError:
                        pass
        except Exception:
            pass
        return counts

    sous_path,  _ = find_latest_export('souscriptions', 'Souscriptions')
    presc_path, _ = find_latest_export('prescriptions', 'Prescriptions')

    souscriptions_w = _count_weeks(sous_path, CHART_CFG['souscriptions']['date_col'])
    # Volumes : le CPL / CPA reste un coût par dossier, pas par euro de CA
    souscriptions_w_n = _count_weeks(sous_path, CHART_CFG['souscriptions']['date_col'], mode='count')
    # Acquisition = dossiers passés à "Dossier validé" dans la semaine
    # Méthode : 'Finalisée le' (ou file-diff si colonne absente dans les anciens exports)
    acquisition = _acquisition_by_file_diff(week_starts, CHART_CFG['souscriptions'])

    # Lead Prescriptions = activations brutes par semaine (join ligne à ligne par Code_Contact)
    # Même logique que l'onglet Prospects hebdo
    _curr_week_budget = week_start_mon(_now().date())
    prescriptions_active_w = _presc_activations_by_week(
        week_starts, CHART_CFG['prescriptions'], curr_week=_curr_week_budget
    )
    prescriptions_active_w_n = _presc_activations_by_week(
        week_starts, CHART_CFG['prescriptions'], curr_week=_curr_week_budget, mode='count'
    )
    acquisition_n = _acquisition_by_file_diff(week_starts, CHART_CFG['souscriptions'], mode='count')

    # Acquisition Prescription = même méthode que l'onglet Attribution (file diff sur prescriptions)
    # = barre verte du graphique Attribution, cohérente avec acquisition_weekly (barre totale)
    presc_acquisition_w = _acquisition_by_file_diff(week_starts, CHART_CFG['prescriptions'])
    presc_acquisition_w_n = _acquisition_by_file_diff(week_starts, CHART_CFG['prescriptions'], mode='count')

    # ── Conversions par canal (Attribution) ────────────────────────────
    # conv_marketing_w = media + affiliation + partenariat + event (même logique que l'onglet Attribution)
    # conv_presc_w     = prescription UTM (conservé pour référence, non utilisé dans le tableau Budget)
    # conv_hors_presc_w conservé pour compatibilité mais non affiché
    _MKT_CANAUX_BUDGET = ('media', 'affiliation', 'partenariat', 'event')
    utm_conv_budget   = _load_utm_cache(conv=True)
    conv_presc_w      = [0] * N_DISP
    conv_hors_presc_w = [0] * N_DISP
    conv_marketing_w  = [0] * N_DISP
    for date_str, utms in utm_conv_budget.items():
        for wi, wstart in enumerate(week_starts):
            wend_str   = str(wstart + timedelta(days=6))
            wstart_str = str(wstart)
            if wstart_str <= date_str <= wend_str:
                for u in utms:
                    canal = _classify_canal(u)
                    _v = u.get('_valeur', 1.0)
                    if canal == 'prescription':
                        conv_presc_w[wi] += _v
                    else:
                        conv_hors_presc_w[wi] += _v
                    if canal in _MKT_CANAUX_BUDGET:
                        conv_marketing_w[wi] += _v
                break

    # Semaine courante (Lun→Dim)
    today   = _now().date()
    cur_idx = N_DISP - 1
    for i, wstart in enumerate(week_starts):
        if wstart <= today <= wstart + timedelta(days=6):
            cur_idx = i
            break
        if wstart > today:
            cur_idx = max(0, i - 1)
            break

    return jsonify({
        'weeks':                       week_labels,
        'dates':                       [str(d) for d in week_starts],
        'current_week_idx':            cur_idx,
        'expenses':                    expenses,
        'totals_perf':                 totals_perf,
        'totals_full':                 totals_full,
        'souscriptions_weekly':        souscriptions_w,
        'prescriptions_active_weekly': prescriptions_active_w,
        'acquisition_weekly':          acquisition,
        'souscriptions_weekly_n':        souscriptions_w_n,
        'prescriptions_active_weekly_n': prescriptions_active_w_n,
        'acquisition_weekly_n':          acquisition_n,
        'unit':                          'eur',
        'conv_hors_presc_weekly':      conv_hors_presc_w,
        'conv_presc_weekly':           conv_presc_w,
        'conv_marketing_weekly':       conv_marketing_w,
        'presc_acquisition_weekly':    presc_acquisition_w,
        'presc_acquisition_weekly_n':  presc_acquisition_w_n,
    })


# ── Compactage traffic (module-level, partagé par les routes snapshot) ───────
_SNAP_KEEP = {'date','src','src_detail','src_campaign','src_creation',
               'src_variant','device_type','m_unique_visitors','m_visits'}

def _compact_traffic(data):
    """Déduplique les rows Piano/UTM et supprime prev_rows pour alléger le snapshot."""
    if not data or 'rows' not in data:
        return data
    by_key = {}
    for r in data['rows']:
        k = '|'.join([
            str(r.get('date','') or ''),
            str(r.get('src','') or ''),
            str(r.get('src_detail','') or ''),
            str(r.get('src_campaign','') or ''),
            str(r.get('src_creation','') or ''),
            str(r.get('src_variant','') or ''),
            str(r.get('device_type','') or ''),
        ])
        if k not in by_key:
            by_key[k] = {f: r.get(f) for f in _SNAP_KEEP}
            by_key[k]['m_visits'] = 0
            by_key[k]['m_unique_visitors'] = 0
        by_key[k]['m_visits']          += r.get('m_visits', 0) or 0
        by_key[k]['m_unique_visitors'] += r.get('m_unique_visitors', 0) or 0
    result = {k: v for k, v in data.items() if k != 'prev_rows'}
    result['rows'] = list(by_key.values())
    return result


@app.route('/export/snapshot')
def export_snapshot():
    """Génère un fichier HTML autonome avec toutes les données en dur."""
    from flask import make_response
    import base64

    now       = _now()
    date_str  = now.strftime('%Y-%m-%d')
    date_label = now.strftime('%d/%m/%Y à %H:%M')

    # ── Collecte des données via le test client Flask ──────────────────────
    with app.test_client() as c:
        def _get(url):
            try:
                r = c.get(url)
                return r.get_json() if r.status_code == 200 else None
            except Exception:
                return None

        yesterday   = (now.date() - timedelta(days=1)).isoformat()
        three_wk    = (now.date() - timedelta(days=20)).isoformat()
        four_weeks  = (now.date() - timedelta(days=27)).isoformat()

        snap = {
            'date':          date_str,
            'date_label':    date_label,
            'sous_thu':      _get('/api/chart/souscriptions?week_mode=thu'),
            'sous_mon':      _get('/api/chart/souscriptions?week_mode=mon'),
            'presc_thu':     _get('/api/chart/prescriptions?week_mode=thu'),
            'presc_mon':     _get('/api/chart/prescriptions?week_mode=mon'),
            'stats_sous':    _get('/api/stats/souscriptions'),
            'stats_presc':   _get('/api/stats/prescriptions'),
            'sous_wk_opened_thu': _get('/api/stats/souscriptions/weekly_opened?week_mode=thu'),
            'sous_wk_opened_mon': _get('/api/stats/souscriptions/weekly_opened?week_mode=mon'),
            'sous_target_thu': _get('/api/chart/souscriptions_target?week_mode=thu'),
            'sous_target_mon': _get('/api/chart/souscriptions_target?week_mode=mon'),
            'prospects_weekly_mon': _get('/api/chart/prospects_weekly?week_mode=mon'),
            'prospects_weekly_thu': _get('/api/chart/prospects_weekly?week_mode=thu'),
            'prospects_daily_mon':  _get('/api/chart/prospects_weekly?week_mode=mon&granularity=day'),
            'prospects_daily_thu':  _get('/api/chart/prospects_weekly?week_mode=thu&granularity=day'),
            'sous_daily_mon':       _get('/api/chart/souscriptions?week_mode=mon&granularity=day'),
            'sous_daily_thu':       _get('/api/chart/souscriptions?week_mode=thu&granularity=day'),
            'presc_daily_mon':      _get('/api/chart/prescriptions?week_mode=mon&granularity=day'),
            'presc_daily_thu':      _get('/api/chart/prescriptions?week_mode=thu&granularity=day'),
            'cohortes':      _get('/api/cohortes'),
            'budget':        _get('/api/budget'),
            'budget_media':  _get('/api/budget_media'),
            'traffic_by_filter': {
                'init':         _get(f'/api/traffic?start={three_wk}&end={yesterday}&trafic=1'),
                'lp':           _get(f'/api/traffic?start={three_wk}&end={yesterday}&lp=1'),
                'lead':         _get(f'/api/traffic?start={three_wk}&end={yesterday}&lead=1'),
                'conv':         _get(f'/api/traffic?start={three_wk}&end={yesterday}&conv=1'),
                'init_consent': _get(f'/api/traffic?start={three_wk}&end={yesterday}&trafic=1&consent=1'),
                'lp_consent':   _get(f'/api/traffic?start={three_wk}&end={yesterday}&lp=1&consent=1'),
                'lead_consent': _get(f'/api/traffic?start={three_wk}&end={yesterday}&lead=1&consent=1'),
                'conv_consent': _get(f'/api/traffic?start={three_wk}&end={yesterday}&conv=1&consent=1'),
            },
            'traffic_pie_by_filter': {
                'init':         _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&trafic=1&metric=visits'),
                'lp':           _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&lp=1&metric=visits'),
                'lead':         _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&lead=1&metric=visits'),
                'conv':         _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&conv=1&metric=visits'),
                'init_consent': _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&trafic=1&metric=visits&consent=1'),
                'lp_consent':   _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&lp=1&metric=visits&consent=1'),
                'lead_consent': _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&lead=1&metric=visits&consent=1'),
                'conv_consent': _get(f'/api/traffic/pie?start={four_weeks}&end={yesterday}&conv=1&metric=visits&consent=1'),
            },
            'traffic_bar_by_filter': {
                'init':         _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&trafic=1&metric=visits'),
                'lp':           _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&lp=1&metric=visits'),
                'lead':         _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&lead=1&metric=visits'),
                'conv':         _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&conv=1&metric=visits'),
                'init_consent': _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&trafic=1&metric=visits&consent=1'),
                'lp_consent':   _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&lp=1&metric=visits&consent=1'),
                'lead_consent': _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&lead=1&metric=visits&consent=1'),
                'conv_consent': _get(f'/api/traffic/pie?start={three_wk}&end={yesterday}&conv=1&metric=visits&consent=1'),
            },
        }

    # ── Compactage traffic : dédupliquer rows, supprimer prev_rows ───────────
    for fk in ('init', 'lp', 'lead', 'conv', 'init_consent', 'lp_consent', 'lead_consent', 'conv_consent'):
        tf = snap.get('traffic_by_filter', {})
        if tf.get(fk):
            tf[fk] = _compact_traffic(tf[fk])

    # ── Données Piano Budget Média sub-rows (12 appels parallèles) ───────────
    bm = snap.get('budget_media')
    bmedia_sub = {}
    if bm:
        cur_idx = bm.get('current_week_idx', 0)
        dates   = bm.get('dates', [])
        past_weeks = []
        for i in range(min(cur_idx, len(dates))):
            wk_start = dates[i]
            if wk_start:
                wk_end = (datetime.strptime(wk_start, '%Y-%m-%d').date() + timedelta(days=6)).isoformat()
                past_weeks.append((wk_start, wk_end))

        if past_weeks:
            full_start = past_weeks[0][0]
            full_end   = past_weeks[-1][1]

            # Charger toutes les lignes Piano depuis le cache disque (une seule fois)
            try:
                all_snap_rows = _fetch_range_cached(full_start, full_end, AT_SITE)
                app.logger.info(f'snapshot: {len(all_snap_rows)} lignes Piano chargées depuis le cache ({full_start}→{full_end})')
            except Exception as e:
                app.logger.warning(f'snapshot: erreur chargement cache Piano: {e}')
                all_snap_rows = []

            # Cache UTM : LEAD et CONVERSION avec dates différentes
            utm_by_date_snap      = _load_utm_cache(conv=False)
            utm_by_date_snap_conv = _load_utm_cache(conv=True)

            def _snap_fetch(src, ut):
                try:
                    if ut in ('lead', 'conv'):
                        filter_fn = _UTM_FILTERS.get(src)
                        if not filter_fn:
                            return src, ut, {}
                        cache = utm_by_date_snap_conv if ut == 'conv' else utm_by_date_snap
                        by_week = {}
                        for wk_start, wk_end in past_weeks:
                            by_week[wk_start] = sum(
                                 u.get('_valeur', 1.0) for d, utms in cache.items()
                                if wk_start <= d <= wk_end
                                for u in utms
                                if filter_fn(u)
                            )
                        return src, ut, by_week
                    combined = _bmedia_combined_filter(src, ut)
                    rows = [r for r in all_snap_rows if _eval_piano_filter(r, combined)]
                    by_week = {}
                    for wk_start, wk_end in past_weeks:
                        by_week[wk_start] = sum(
                            r.get('m_visits', 0) or 0 for r in rows
                            if wk_start <= (r.get('date') or '') <= wk_end
                        )
                    return src, ut, by_week
                except Exception as e:
                    app.logger.warning(f'bmedia_sub snap failed {src}/{ut}: {e}')
                    return src, ut, {}

            tasks = [(src, ut) for src in _BMEDIA_SPF for ut in _BMEDIA_URL_FILTERS]
            with ThreadPoolExecutor(max_workers=4) as pool:
                futures = {pool.submit(_snap_fetch, src, ut): (src, ut) for src, ut in tasks}
                for fut in as_completed(futures):
                    try:
                        src, ut, by_week = fut.result()
                        if src not in bmedia_sub:
                            bmedia_sub[src] = {k: {} for k in _BMEDIA_URL_FILTERS}
                        bmedia_sub[src][ut] = by_week
                    except Exception:
                        pass

    snap['bmedia_sub'] = bmedia_sub
    snap_json = json.dumps(snap, ensure_ascii=False)

    # ── Lecture du template HTML ───────────────────────────────────────────
    with open(str(BASE_DIR / 'templates' / 'index.html'), encoding='utf-8') as f:
        html = f.read()

    # ── Script d'override à injecter avant loadAll() ───────────────────────
    # Utilise un template de chaîne (pas f-string) pour éviter d'échapper les {}
    snap_script = (
        '\n/* ══════════════════════ SNAPSHOT MODE ══════════════════════\n'
        '   Données figées au ' + date_label + '\n'
        '   ══════════════════════════════════════════════════════════ */\n'
        '(function(){\n'
        'const __D__ = ' + snap_json + ';\n'
        '\n'
        '/* Mise à jour du header */\n'
        'const _hr = document.querySelector(\'.hdr-right\');\n'
        'if(_hr) _hr.innerHTML = \'<span style="background:#eef1fd;border:1px solid rgba(50,69,193,.3);border-radius:999px;padding:6px 14px;color:var(--accent);font-weight:600;font-size:12px">📸 Snapshot du ' + date_label + '</span>\';\n'
        '\n'
        '/* No-ops */\n'
        'window.pollStatus   = function(){};\n'
        'window.reloadData   = async function(){};\n'
        'window.startExport  = function(){ alert("Export non disponible en mode snapshot."); };\n'
        'window.exportSnapshot = function(){ alert("Export non disponible en mode snapshot."); };\n'
        '\n'
        'window.loadSection = async function(apiKind, kind, wid, canId, lgId, valideLabel){\n'
        '  const isSous  = kind === "sous";\n'
        '  const state   = _tabState[kind];\n'
        '  state.exportsLoaded = true;\n'
        '  const stats   = isSous ? __D__.stats_sous : __D__.stats_presc;\n'
        '  const isDay   = state.volGran === "day";\n'
        '  const mode    = state.weekMode === "mon" ? "mon" : "thu";\n'
        '  let chart, vlLabel;\n'
        '  if(isSous && !isDay){\n'
        '    chart   = __D__["prospects_weekly_" + mode];\n'
        '    vlLabel = null;\n'
        '  } else if(isSous && isDay){\n'
        '    chart   = __D__["prospects_daily_" + mode];\n'
        '    vlLabel = null;\n'
        '  } else if(!isDay){\n'
        '    chart   = __D__["presc_" + mode];\n'
        '    vlLabel = valideLabel;\n'
        '  } else {\n'
        '    chart   = __D__["presc_daily_" + mode];\n'
        '    vlLabel = valideLabel;\n'
        '  }\n'
        '  if(chart) renderChart(wid, canId, chart, lgId, vlLabel);\n'
        '  else document.getElementById(wid).innerHTML = \'<div class="empty"><div class="empty-ico">📭</div><div class="empty-txt">Aucune donnée</div></div>\';\n'
        '  if(stats) updateKpi(stats, kind);\n'
        '  if(isSous) _loadSousWeeklyOpened();\n'
        '};\n'
        '\n'
        'window._loadSousWeeklyOpened = function(){\n'
        '  const key = (typeof _sousWkOpenedMode !== "undefined" && _sousWkOpenedMode === "mon") ? "sous_wk_opened_mon" : "sous_wk_opened_thu";\n'
        '  const d = __D__[key];\n'
        '  if(!d) return;\n'
        '  const el  = document.getElementById("s-wk-opened");\n'
        '  const re  = document.getElementById("s-wk-opened-range");\n'
        '  const lbl = document.getElementById("s-wk-opened-label");\n'
        '  const nxt = document.getElementById("sWkOpenedNext");\n'
        '  if(el)  el.textContent  = (d.count || 0).toLocaleString("fr");\n'
        '  if(re)  re.textContent  = d.wk_start + " – " + d.wk_end;\n'
        '  if(lbl) lbl.textContent = "Comptes ouverts cette semaine";\n'
        '  if(nxt) nxt.disabled    = true;\n'
        '};\n'
        '/* Navigation semaine désactivée en snapshot */\n'
        'window.shiftWkOpened = function(){};\n'
        '\n'
        'window.loadAll = async function(){\n'
        '  await Promise.all([\n'
        '    loadSection("souscriptions","sous","cwSous","chartSous","lgSous","Dossier validé"),\n'
        '    loadSection("prescriptions","presc","cwPresc","chartPresc","lgPresc","EER validé"),\n'
        '    loadSousTargetChart(),\n'
        '    loadCohortes(),\n'
        '  ]);\n'
        '};\n'
        '\n'
        'window.loadCohortes = async function(){\n'
        '  const data = __D__.cohortes;\n'
        '  if(!data){ document.getElementById("tblCoh").innerHTML=\'<div class="empty"><div class="empty-ico">📊</div><div class="empty-txt">Données non disponibles</div></div>\'; return; }\n'
        '  const sheets = Object.keys(data);\n'
        '  document.getElementById("cohTabs").innerHTML = sheets.map((s,i)=>{\n'
        '    const label = s.replace(/\\s*Lundi-Dimanche\\s*/i,"").trim();\n'
        '    return `<button class="coh-tab${i===0?\' active\':\'\'}" data-sheet="${s}" onclick="selSheet(this,\'${s}\')">${label}</button>`;\n'
        '  }).join("");\n'
        '  window._cohData = data;\n'
        '  if(sheets.length) renderCoh(sheets[0]);\n'
        '};\n'
        '\n'
        'window.loadSousTargetChart = async function(){\n'
        '  const key  = (typeof _sousTargetWkMode !== "undefined" && _sousTargetWkMode === "mon") ? "sous_target_mon" : "sous_target_thu";\n'
        '  const data = __D__[key];\n'
        '  const wrap = document.getElementById("cwSousTarget");\n'
        '  if(!data){ if(wrap) wrap.innerHTML=\'<div class="empty"><div class="empty-ico">⚠️</div><div class="empty-txt">Données non disponibles</div></div>\'; return; }\n'
        '  _sousTargetData = data;\n'
        '  _autoSousWin(data);\n'
        '  renderSousTargetChart(data);\n'
        '};\n'
        '\n'
        'window.loadBudget = async function(){\n'
        '  if(!__D__.budget){ document.getElementById("budgetTableWrap").innerHTML=\'<div class="empty"><div class="empty-ico">⚠️</div><div class="empty-txt">Budget non disponible</div></div>\'; return; }\n'
        '  _budgetData = __D__.budget;\n'
        '  const ci = _budgetData.current_week_idx;\n'
        '  _budgetWinStart = Math.max(0, Math.min(_budgetData.weeks.length-8, ci-4));\n'
        '  renderBudgetTable();\n'
        '};\n'
        '\n'
        'window.loadBudgetMedia = async function(){\n'
        '  if(!__D__.budget_media){ document.getElementById("budgetMediaTableWrap").innerHTML=\'<div class="empty"><div class="empty-ico">⚠️</div><div class="empty-txt">Budget Média non disponible</div></div>\'; return; }\n'
        '  _budgetMediaData = __D__.budget_media;\n'
        '  const ci = _budgetMediaData.current_week_idx;\n'
        '  _budgetMediaWinStart = Math.max(0, Math.min(_budgetMediaData.weeks.length-8, ci-4));\n'
        '  // Pré-remplir le cache Piano bmedia_sub depuis les données snapshot\n'
        '  if(__D__.bmedia_sub){\n'
        '    for(const [src, utMap] of Object.entries(__D__.bmedia_sub)){\n'
        '      _bmediaSubCache[src] = {lp:{}, init:{}, lead:{}, conv:{}};\n'
        '      for(const [ut, weekMap] of Object.entries(utMap)){\n'
        '        _bmediaSubCache[src][ut] = Object.assign({}, weekMap);\n'
        '      }\n'
        '    }\n'
        '  }\n'
        '  renderBudgetMediaTable();\n'
        '  // Afficher le tableau Comparaison Média (utilise _bmediaSubCache déjà rempli)\n'
        '  _comparaisonWeekIdx = Math.max(0, ci - 1);\n'
        '  _updateComparaisonLabel();\n'
        '  loadComparaison();\n'
        '};\n'
        '\n'
        '// En mode snapshot, toutes les données Piano sont déjà en cache — pas de fetch\n'
        'window._ensureBmediaSubLoaded = function(){};\n'
        '// loadComparaison en snapshot : uniquement depuis _bmediaSubCache, jamais d\'appel API\n'
        'window.loadComparaison = async function(){\n'
        '  if(!_budgetMediaData) return;\n'
        '  const ws = _budgetMediaData.dates[_comparaisonWeekIdx];\n'
        '  if(!ws) return;\n'
        '  _tryBuildComparaisonFromCache(_comparaisonWeekIdx, ws);\n'
        '};\n'
        '\n'
        '/* ── Sélection du dataset selon le filtre actif (priorité: conv > lead > init > lp) + consent ── */\n'
        'function _snapFilterKey(){\n'
        '  const sfx = _trafConsentFilter ? "_consent" : "";\n'
        '  if(_trafConvFilter)   return "conv"   + sfx;\n'
        '  if(_trafLeadFilter)   return "lead"   + sfx;\n'
        '  if(_trafTraficFilter) return "init"   + sfx;\n'
        '  if(_trafLpFilter)     return "lp"     + sfx;\n'
        '  return "init" + sfx;\n'
        '}\n'
        '\n'
        '/* Toggles exclusifs en snapshot : activer un filtre désactive les autres */\n'
        'function _snapToggle(key){\n'
        '  _trafConvFilter   = (key==="conv");\n'
        '  _trafLeadFilter   = (key==="lead");\n'
        '  _trafTraficFilter = (key==="init");\n'
        '  _trafLpFilter     = (key==="lp");\n'
        '  const map={conv:["btnTrafConv","#9b5de5"],lead:["btnTrafLead","#e8962a"],init:["btnTrafTrafic","#3245c1"],lp:["btnTrafLp","#1f9d7a"]};\n'
        '  for(const [k,[id,col]] of Object.entries(map)){\n'
        '    const b=document.getElementById(id); if(!b) continue;\n'
        '    b.style.background=(k===key)?col:"transparent";\n'
        '    b.style.color=(k===key)?"#ffffff":col;\n'
        '  }\n'
        '  loadTraffic();\n'
        '  loadTrafficPie();\n'
        '}\n'
        'window.toggleTrafConv   = function(){ _snapToggle("conv"); };\n'
        'window.toggleTrafLead   = function(){ _snapToggle("lead"); };\n'
        'window.toggleTrafTrafic = function(){ _snapToggle("init"); };\n'
        'window.toggleTrafLp     = function(){ _snapToggle("lp");   };\n'
        'window.toggleTrafConsent = function(){\n'
        '  _trafConsentFilter = !_trafConsentFilter;\n'
        '  const btn = document.getElementById("btnTrafConsent");\n'
        '  if(btn) btn.classList.toggle("active", _trafConsentFilter);\n'
        '  loadTraffic();\n'
        '  loadTrafficPie();\n'
        '};\n'
        '\n'
        'window.loadTraffic = async function(){\n'
        '  const key  = _snapFilterKey();\n'
        '  const data = __D__.traffic_by_filter && __D__.traffic_by_filter[key];\n'
        '  const ts = document.getElementById("trafStart");\n'
        '  const te = document.getElementById("trafEnd");\n'
        '  if(data && ts) ts.value = data.start;\n'
        '  if(data && te) te.value = data.end;\n'
        '  if(!data){ document.getElementById("cwTraffic").innerHTML=\'<div class="empty"><div class="empty-ico">⚠️</div><div class="empty-txt">Données trafic non disponibles</div></div>\'; return; }\n'
        '  _trafData = data;\n'
        '  _trafHiddenSrcs = new Set();\n'
        '  _renderTraffic(_trafData);\n'
        '};\n'
        '\n'
        'window.loadTrafficPie = async function(){\n'
        '  const key = _snapFilterKey();\n'
        '  const metLabel = "Visites";\n'
        '  const data = __D__.traffic_pie_by_filter && __D__.traffic_pie_by_filter[key];\n'
        '  if(!data){\n'
        '    document.getElementById("cwTrafficPie").innerHTML=\'<div class="empty"><div class="empty-ico">📭</div><div class="empty-txt">Données PIE non disponibles</div></div>\';\n'
        '    return;\n'
        '  }\n'
        '  // Titre + sous-titre figés au snapshot (J-7)\n'
        '  const titleEl = document.getElementById("trafPieTitle");\n'
        '  if(titleEl){\n'
        '    titleEl.childNodes[0].textContent = "Part des sources  - semaine passée J-7";\n'
        '  }\n'
        '  const filterLabel = {conv:"CONVERSION",lead:"Lead",init:"INIT SOUSCRIPTION",lp:"LP"}[key] || "Tous";\n'
        '  const fmtD = s => s ? s.split("-").reverse().join("/") : "—";\n'
        '  const subEl = document.getElementById("trafPieSub");\n'
        '  if(subEl) subEl.innerHTML = `<strong>${filterLabel}</strong>  ·  ${fmtD(data.start)} → ${fmtD(data.end)}  ·  ${metLabel}`;\n'
        '  // Pour le bar chart, utiliser by_date des 3 semaines (pas des 7 jours du pie)\n'
        '  const barData = __D__.traffic_bar_by_filter && __D__.traffic_bar_by_filter[key];\n'
        '  _pieRawData = Object.assign({}, data, { by_date: barData ? barData.by_date : data.by_date });\n'
        '  _pieAutresDetail = false;\n'
        '  _applyPieRender();\n'
        '};\n'
        '\n'
        '/* ── UI snapshot : dates figées + boutons secondaires masqués ─────────── */\n'
        '(function(){\n'
        '  // Dater les inputs en lecture seule\n'
        '  ["trafStart","trafEnd"].forEach(function(id){\n'
        '    const el=document.getElementById(id);\n'
        '    if(el){el.setAttribute("readonly","");el.style.pointerEvents="none";el.style.opacity="0.55";el.title="Période figée au snapshot";}\n'
        '  });\n'
        '  // Masquer filtres secondaires (PMAX, META, SEA BRAND, SEA NB, Prescription)\n'
        '  ["btnTrafPmax","btnTrafMeta","btnTrafSeaBrand","btnTrafSeaNb","btnTrafPresc",\n'
        '   "btnBmediaConsent","btnPeConsent","trafMetricV","trafMetricUv",\n'
        '   "btnSanityCheck","btnExport","btnExportAffiliation",\n'
        '   "sousVolNext","sousVolLabel","prescVolNext","prescVolLabel"].forEach(function(id){\n'
        '    const el=document.getElementById(id); if(el) el.style.display="none";\n'
        '  });\n'
        '  // Masquer aussi le bouton ◀ des navigateurs vol sous/presc (pas d\'id, chercher par onclick)\n'
        '  document.querySelectorAll("[onclick*=\\"shiftVolWeek\\"]").forEach(function(el){ el.style.display="none"; });\n'
        '  // Masquer onglet Portail Entrepreneur\n'
        '  document.querySelectorAll(".tab").forEach(function(el){\n'
        '    if(el.getAttribute("onclick") && el.getAttribute("onclick").indexOf("\'pe\'")!==-1) el.style.display="none";\n'
        '  });\n'
        '  // Masquer boutons Actualiser, Export trafic et Export Pie Data\n'
        '  document.querySelectorAll(".btn-refresh,[onclick*=\\"exportTrafic\\"],[onclick*=\\"exportPie\\"]").forEach(function(el){el.style.display="none";});\n'
        '  // No-ops pour éviter les erreurs si les boutons sont appelés autrement\n'
        '  window.exportTrafic = function(){};\n'
        '  window.exportPie    = function(){};\n'
        '})();\n'
        '\n'
        '})();\n'
    )

    # Injecter juste avant le bloc d'appel final
    marker = 'pollStatus();\nsetInterval(pollStatus, 5000);\nloadAll();'
    html   = html.replace(marker, snap_script + marker)

    filename = f"Nova_Marketing_Dashboard_{now.strftime('%Y%m%d_%H%M')}.html"
    resp = make_response(html)
    resp.headers['Content-Type'] = 'text/html; charset=utf-8'
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp


@app.route('/export/snapshot/tracking')
def export_snapshot_tracking():
    """Rapport Excel Media : volumes journaliers J-28→J-1 par filtre (PMAX, META, TIKTOK, AMAZON, LINKEDIN)
    4 onglets : Conversions, Leads, Init Souscription, LP."""
    from flask import make_response
    import openpyxl as _xl
    from openpyxl.styles import Font as _Font, PatternFill as _PF, Alignment as _Al, Border as _Bo, Side as _Si
    import io

    from datetime import date as _date
    now       = _now()
    yesterday = now.date() - timedelta(days=1)
    # Paramètres optionnels : ?start=2026-05-27&end=2026-06-14
    _start_p = request.args.get('start')
    _end_p   = request.args.get('end')
    app.logger.info(f'[tracking] args start={_start_p!r} end={_end_p!r}')
    if _start_p and _end_p:
        try:
            start_dt = _date.fromisoformat(_start_p)
            end_dt   = _date.fromisoformat(_end_p)
        except Exception as _e:
            app.logger.warning(f'[tracking] fromisoformat failed: {_e}')
            start_dt = yesterday - timedelta(days=27)
            end_dt   = yesterday
    else:
        start_dt = yesterday - timedelta(days=27)
        end_dt   = yesterday
    n_days = (end_dt - start_dt).days + 1
    app.logger.info(f'[tracking] start_dt={start_dt} end_dt={end_dt} n_days={n_days}')
    dates  = [(start_dt + timedelta(days=i)).isoformat() for i in range(n_days)]

    MEDIA_CFG = [
        ('PMAX',      'pmax'),
        ('META',      'meta'),
        ('SEA Brand', 'sea_brand'),
        ('SEA NB',    'sea_nb'),
        ('TIKTOK',    'tiktok'),
        ('AMAZON',    'amazon'),
        ('LINKEDIN',  'linkedin'),
    ]

    # ── Données Piano (LP + Init Souscription) ──────────────────────────────────────────
    try:
        piano_rows = _fetch_range_cached(start_dt.isoformat(), end_dt.isoformat(), AT_SITE)
    except Exception:
        piano_rows = []

    # ── Données UTM (Leads + Conversions) ─────────────────────────────────────────────
    utm_lead = _load_utm_cache(conv=False)
    utm_conv = _load_utm_cache(conv=True)

    def _piano_daily(src_key, url_type, consent=False):
        """Volumes Piano journaliers : filtre source + URL (+ consent optionnel)."""
        combined = _bmedia_combined_filter(src_key, url_type)
        if consent:
            combined = {'$and': [combined, CONSENT_FILTER]}
        by_day = {d: 0 for d in dates}
        for r in piano_rows:
            d = r.get('date', '')
            if d in by_day and _eval_piano_filter(r, combined):
                by_day[d] += r.get('m_visits', 0) or 0
        return by_day

    def _utm_daily(src_key, cache):
        """Volumes UTM journaliers : filtre source."""
        filter_fn = _UTM_FILTERS.get(src_key)
        by_day = {d: 0 for d in dates}
        if not filter_fn:
            return by_day
        for d, utms in cache.items():
            if d in by_day:
                by_day[d] += sum(u.get('_valeur', 1.0) for u in utms if filter_fn(u))
        return by_day

    # ── Styles ────────────────────────────────────────────────────────────────
    ACCENT   = '0096D6'
    DARK_BG  = '0D1520'
    HDR_BG   = '1A2A3A'
    ALT_BG   = '111C28'
    TOT_BG   = '0A3050'
    THIN     = _Bo(left=_Si('thin', color='1E3048'), right=_Si('thin', color='1E3048'),
                   top=_Si('thin', color='1E3048'), bottom=_Si('thin', color='1E3048'))

    def _fill(hex_):
        return _PF('solid', fgColor=hex_)

    def _hdr(ws, row, col, val):
        c = ws.cell(row, col, val)
        c.font = _Font(bold=True, color='FFFFFF', size=10)
        c.fill = _fill(HDR_BG); c.alignment = _Al(horizontal='center'); c.border = THIN

    def _cell(ws, row, col, val, bg=None, bold=False):
        c = ws.cell(row, col, val)
        c.font = _Font(bold=bold, color='FFFFFF', size=10)
        if bg: c.fill = _fill(bg)
        c.alignment = _Al(horizontal='right' if isinstance(val, (int, float)) else 'left')
        c.border = THIN

    WHITE    = 'FFFFFF'
    BLACK    = '000000'
    OPT_COL  = '217346'   # vert foncé Excel (lisible sur fond blanc)

    def _col_letter(ci):
        return chr(64 + ci) if ci <= 26 else ('A' + chr(64 + ci - 26))

    def _build_sheet(wb, title, data_fn_by_media, data_fn_consent=None):
        ws = wb.create_sheet(title)
        n_media = len(MEDIA_CFG)
        n_cols  = 1 + n_media   # Date | canal1..N  (pas de colonne TOTAL)

        # Titre
        ws.merge_cells(f'A1:{_col_letter(n_cols)}1')
        t = ws.cell(1, 1, f'{title} — Volumes journaliers par canal média  ({dates[0][8:]}/{dates[0][5:7]}/{dates[0][:4]} → {dates[-1][8:]}/{dates[-1][5:7]}/{dates[-1][:4]})')
        t.font = _Font(bold=True, color='00D4FF', size=13)
        t.fill = _fill(DARK_BG)
        t.alignment = _Al(horizontal='left')
        ws.row_dimensions[1].height = 24

        # En-têtes
        _hdr(ws, 2, 1, 'Date')
        for ci, (label, _) in enumerate(MEDIA_CFG, 2):
            _hdr(ws, 2, ci, label)

        ws.column_dimensions['A'].width = 18
        for ci in range(2, n_cols + 1):
            ws.column_dimensions[_col_letter(ci)].width = 11

        # Pré-calculer toutes les séries
        series         = {src_key: data_fn_by_media(src_key) for _, src_key in MEDIA_CFG}
        series_consent = {src_key: data_fn_consent(src_key)  for _, src_key in MEDIA_CFG} if data_fn_consent else None

        totals_by_canal         = {src_key: 0 for _, src_key in MEDIA_CFG}
        totals_by_canal_consent = {src_key: 0 for _, src_key in MEDIA_CFG}

        current_row = 3
        for d in dates:
            # ── Ligne principale : fond blanc, police noire ───────────────────
            c0 = ws.cell(current_row, 1, d[8:] + '/' + d[5:7])
            c0.font = _Font(bold=True, color=BLACK, size=10)
            c0.fill = _fill(WHITE); c0.border = THIN
            for ci, (_, src_key) in enumerate(MEDIA_CFG, 2):
                v = series[src_key].get(d, 0)
                c = ws.cell(current_row, ci, v if v else '')
                c.font = _Font(color=BLACK, size=10)
                c.fill = _fill(WHITE); c.border = THIN
                c.alignment = _Al(horizontal='right')
                totals_by_canal[src_key] += v
            current_row += 1

            # ── Sous-ligne opt-in média : fond blanc, police verte ────────────
            if series_consent:
                c0 = ws.cell(current_row, 1, '↳ opt-in média')
                c0.font = _Font(italic=True, color=OPT_COL, size=9)
                c0.fill = _fill(WHITE); c0.border = THIN
                for ci, (_, src_key) in enumerate(MEDIA_CFG, 2):
                    v = series_consent[src_key].get(d, 0)
                    c = ws.cell(current_row, ci, v if v else '')
                    c.font = _Font(italic=True, color=OPT_COL, size=9)
                    c.fill = _fill(WHITE); c.border = THIN
                    c.alignment = _Al(horizontal='right')
                    totals_by_canal_consent[src_key] += v
                current_row += 1

        # ── Ligne TOTAL : fond blanc, police noire grasse ─────────────────────
        def _tot_cell(row, col, val, color=BLACK, italic=False):
            c = ws.cell(row, col, val if val else '')
            c.font = _Font(bold=True, italic=italic, color=color, size=10)
            c.fill = _fill(WHITE); c.border = THIN
            c.alignment = _Al(horizontal='right' if col > 1 else 'left')

        _tot_cell(current_row, 1, 'TOTAL')
        for ci, (_, src_key) in enumerate(MEDIA_CFG, 2):
            _tot_cell(current_row, ci, totals_by_canal[src_key])
        current_row += 1

        # ── Ligne TOTAL opt-in média : fond blanc, police verte grasse ───────
        if series_consent:
            _tot_cell(current_row, 1, 'TOTAL opt-in média', color=OPT_COL, italic=True)
            for ci, (_, src_key) in enumerate(MEDIA_CFG, 2):
                _tot_cell(current_row, ci, totals_by_canal_consent[src_key], color=OPT_COL, italic=True)

        ws.freeze_panes = 'B3'
        return ws

    # ── Onglet par canal : dates en ligne, indicateurs en colonne ─────────────────────
    KPI_CFG = [
        ('LP',          lambda src: _piano_daily(src, 'lp'),   lambda src: _piano_daily(src, 'lp',   consent=True)),
        ('Init Souscription',    lambda src: _piano_daily(src, 'init'), lambda src: _piano_daily(src, 'init', consent=True)),
        ('Leads',       lambda src: _utm_daily(src, utm_lead), None),
        ('Conversions', lambda src: _utm_daily(src, utm_conv), None),
    ]

    def _build_canal_sheet(wb, canal_label, src_key):
        ws = wb.create_sheet(canal_label)
        # Colonnes : Date | LP | LP opt-in | Init | Init opt-in | Leads | Conversions | TOTAL
        headers = ['Date']
        col_fns = []   # (fn_data, fn_consent_or_None) par colonne d'indicateur
        for kpi_label, fn, fn_c in KPI_CFG:
            headers.append(kpi_label)
            col_fns.append((kpi_label, fn, None))
            if fn_c is not None:
                headers.append('opt-in média')
                col_fns.append((kpi_label + ' opt-in', fn_c, None))
        n_cols = len(headers)   # pas de colonne TOTAL

        # Titre
        ws.merge_cells(f'A1:{_col_letter(n_cols)}1')
        t = ws.cell(1, 1, f'{canal_label} — Volumes journaliers par événement  ({dates[0][8:]}/{dates[0][5:7]}/{dates[0][:4]} → {dates[-1][8:]}/{dates[-1][5:7]}/{dates[-1][:4]})')
        t.font = _Font(bold=True, color='00D4FF', size=13)
        t.fill = _fill(DARK_BG); t.alignment = _Al(horizontal='left')
        ws.row_dimensions[1].height = 24

        # En-têtes
        for ci, h in enumerate(headers, 1):
            _hdr(ws, 2, ci, h)

        ws.column_dimensions['A'].width = 13
        for ci in range(2, n_cols + 1):
            ws.column_dimensions[_col_letter(ci)].width = 16

        # Pré-calculer séries
        series_by_col = [(lbl, fn(src_key)) for lbl, fn, _ in col_fns]

        # Totaux
        col_totals = [0] * len(col_fns)

        current_row = 3
        for d in dates:
            # Ligne principale (date)
            c0 = ws.cell(current_row, 1, d[8:] + '/' + d[5:7])
            c0.font = _Font(bold=True, color=BLACK, size=10)
            c0.fill = _fill(WHITE); c0.border = THIN
            for ci, (lbl, vals) in enumerate(series_by_col):
                v    = vals.get(d, 0)
                is_o = lbl.endswith(' opt-in')
                c = ws.cell(current_row, ci + 2, v if v else '')
                c.font = _Font(italic=is_o, color=OPT_COL if is_o else BLACK, size=10)
                c.fill = _fill(WHITE); c.border = THIN
                c.alignment = _Al(horizontal='right')
                col_totals[ci] += v
            current_row += 1

        # ── Ligne TOTAL : fond blanc, police noire grasse ─────────────────────
        def _tc(row, col, val, color=BLACK, italic=False):
            c = ws.cell(row, col, val if val else '')
            c.font = _Font(bold=True, italic=italic, color=color, size=10)
            c.fill = _fill(WHITE); c.border = THIN
            c.alignment = _Al(horizontal='right' if col > 1 else 'left')

        _tc(current_row, 1, 'TOTAL')
        for ci, (lbl, _) in enumerate(series_by_col):
            is_o = lbl.endswith(' opt-in')
            _tc(current_row, ci + 2, col_totals[ci], color=OPT_COL if is_o else BLACK, italic=is_o)

        ws.freeze_panes = 'B3'
        return ws

    wb = _xl.Workbook()
    wb.remove(wb.active)   # supprimer la feuille vide par défaut

    _build_sheet(wb, 'Conversions',
                 lambda src: _utm_daily(src, utm_conv))
    _build_sheet(wb, 'Leads',
                 lambda src: _utm_daily(src, utm_lead))
    _build_sheet(wb, 'Init Souscription',
                 lambda src: _piano_daily(src, 'init'),
                 data_fn_consent=lambda src: _piano_daily(src, 'init', consent=True))
    _build_sheet(wb, 'LP',
                 lambda src: _piano_daily(src, 'lp'),
                 data_fn_consent=lambda src: _piano_daily(src, 'lp', consent=True))

    for label, src_key in MEDIA_CFG:
        _build_canal_sheet(wb, label, src_key)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname = f"Nova_Media_Daily_{now.strftime('%Y%m%d_%H%M')}.xlsx"
    resp = make_response(buf.read())
    resp.headers['Content-Type']        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp



@app.route('/api/budget_media')
def api_budget_media():
    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé — pip install openpyxl'}), 500

    from datetime import date as _date
    DISP_START = _date(2026, 5, 4)   # lundi, display start (S-1)
    N_DISP     = _budget_n_disp()
    week_starts = [DISP_START + timedelta(weeks=i) for i in range(N_DISP)]
    week_labels = [_wk_label(ws) for ws in week_starts]

    today   = _now().date()
    cur_idx = N_DISP - 1
    for i, wstart in enumerate(week_starts):
        if wstart <= today <= wstart + timedelta(days=6):
            cur_idx = i
            break
        if wstart > today:
            cur_idx = max(0, i - 1)
            break

    nova_lines, portail_lines, nova_total, portail_total = _load_media_data(N_DISP)
    global_total = [nova_total[i] + portail_total[i] for i in range(N_DISP)]

    return jsonify({
        'weeks':            week_labels,
        'dates':            [str(d) for d in week_starts],
        'current_week_idx': cur_idx,
        'nova_lines':       nova_lines,
        'nova_total':       nova_total,
        'portail_lines':    portail_lines,
        'portail_total':    portail_total,
        'global_total':     global_total,
    })


@app.route('/api/budget_media/comparaison')
def api_budget_media_comparaison():
    """Budget vs trafic Piano par levier pour une semaine donnée.
    Paramètre : week_start=YYYY-MM-DD (défaut : semaine précédente complète)."""
    from datetime import date as _date
    DISP_START = _date(2026, 5, 4)   # lundi
    N_DISP     = _budget_n_disp()
    week_starts = [(DISP_START + timedelta(weeks=i)).isoformat() for i in range(N_DISP)]

    week_start = request.args.get('week_start')
    if not week_start:
        today = _now().date()
        week_start = week_starts[0]
        for ws in reversed(week_starts):
            ws_dt = datetime.strptime(ws, '%Y-%m-%d').date()
            if ws_dt + timedelta(days=6) < today:
                week_start = ws
                break

    if week_start not in week_starts:
        return jsonify({'error': f'Semaine non trouvée: {week_start}'}), 400

    week_idx = week_starts.index(week_start)
    week_end = (datetime.strptime(week_start, '%Y-%m-%d').date() + timedelta(days=6)).isoformat()

    # Budget Nova depuis Excel
    try:
        nova_lines, _, _, _ = _load_media_data(N_DISP)
    except Exception:
        nova_lines = []

    consent_only = request.args.get('consent') == '1'

    # Trafic Piano (cache local)
    try:
        all_rows = _fetch_range_cached(week_start, week_end, AT_SITE)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    if consent_only:
        all_rows = [r for r in all_rows if _eval_piano_filter(r, CONSENT_FILTER)]

    # Cache UTM : LEAD et CONVERSION avec dates différentes
    utm_by_date_comp      = _load_utm_cache(conv=False)
    utm_by_date_comp_conv = _load_utm_cache(conv=True)

    MEDIA_CONFIG = [
        ('PMAX',      ['Performance MAX'],  'pmax'),
        ('META',      ['Meta'],             'meta'),
        ('SEA Brand', ['SEA - Brand'],      'sea_brand'),
        ('SEA NB',    ['SEA - Non Brand'],  'sea_nb'),
        ('Amazon',    ['Amazon Display'],   'amazon'),
        ('TikTok',    ['TikTok'],           'tiktok'),
        ('LinkedIn',  ['LinkedIn'],         'linkedin'),
    ]

    def _utm_count_week(spf_key, wk_start, wk_end, conv=False, mode='eur'):
        filter_fn = _UTM_FILTERS.get(spf_key)
        if not filter_fn:
            return 0
        cache = utm_by_date_comp_conv if conv else utm_by_date_comp
        return sum(
            (1.0 if mode == 'count' else u.get('_valeur', 1.0))
            for d, utms in cache.items()
            if wk_start <= d <= wk_end
            for u in utms
            if filter_fn(u)
        )

    def _build_media(rows, wi, lines, wk_start, wk_end):
        result = []
        for (media_name, budget_platforms, spf_key) in MEDIA_CONFIG:
            budget = sum(
                (line['values'][wi] or 0)
                for line in lines
                if line['platform'] in budget_platforms
            )
            spf = _BMEDIA_SPF.get(spf_key)
            media_rows = [r for r in rows if not spf or _eval_piano_filter(r, spf)]
            lp_v   = sum(r.get('m_visits', 0) or 0 for r in media_rows if LP_URL    in (r.get('event_url_path', '') or ''))
            init_v = sum(r.get('m_visits', 0) or 0 for r in media_rows if TRAFIC_URL in (r.get('event_url_path', '') or ''))
            lead_v = _utm_count_week(spf_key, wk_start, wk_end)
            conv_v = _utm_count_week(spf_key, wk_start, wk_end, conv=True)
            # Volumes : le CPL / CPA reste un coût par unité
            lead_n = _utm_count_week(spf_key, wk_start, wk_end, mode='count')
            conv_n = _utm_count_week(spf_key, wk_start, wk_end, conv=True, mode='count')
            result.append({'label': media_name, 'budget': budget, 'lp': lp_v, 'init': init_v,
                           'lead': lead_v, 'conv': conv_v, 'lead_n': lead_n, 'conv_n': conv_n})
        return result

    media_result = _build_media(all_rows, week_idx, nova_lines, week_start, week_end)

    # Semaine S-1
    media_prev = None
    if week_idx > 0:
        prev_start = week_starts[week_idx - 1]
        prev_end   = (datetime.strptime(prev_start, '%Y-%m-%d').date() + timedelta(days=6)).isoformat()
        try:
            prev_rows = _fetch_range_cached(prev_start, prev_end, AT_SITE)
            if consent_only:
                prev_rows = [r for r in prev_rows if _eval_piano_filter(r, CONSENT_FILTER)]
            media_prev = _build_media(prev_rows, week_idx - 1, nova_lines, prev_start, prev_end)
        except Exception:
            media_prev = None

    return jsonify({
        'week_start': week_start,
        'week_end':   week_end,
        'week_label': _wk_label(datetime.strptime(week_start, '%Y-%m-%d').date()),
        'media':      media_result,
        'media_prev': media_prev,
    })


@app.route('/api/canaux/comparaison')
def api_canaux_comparaison():
    """Leads et conversions agrégés par canal d'acquisition pour une semaine + S-1.
    Paramètre : week_start=YYYY-MM-DD (défaut : dernière semaine complète)."""
    from datetime import date as _date
    DISP_START = _date(2026, 5, 4)
    N_DISP     = _budget_n_disp()
    week_starts = [(DISP_START + timedelta(weeks=i)).isoformat() for i in range(N_DISP)]

    week_start = request.args.get('week_start')
    if not week_start:
        today = _now().date()
        week_start = week_starts[0]
        for ws in reversed(week_starts):
            ws_dt = datetime.strptime(ws, '%Y-%m-%d').date()
            if ws_dt + timedelta(days=6) < today:
                week_start = ws
                break

    if week_start not in week_starts:
        return jsonify({'error': f'Semaine non trouvée: {week_start}'}), 400

    week_idx = week_starts.index(week_start)
    week_end = (datetime.strptime(week_start, '%Y-%m-%d').date() + timedelta(days=6)).isoformat()

    utm_lead = _load_utm_cache(conv=False)
    utm_conv = _load_utm_cache(conv=True)

    _MKT_CANAUX_CC = ('media', 'affiliation', 'partenariat', 'event')

    def _count_by_canal(cache, wk_start, wk_end, mode='eur'):
        counts = {c: 0 for c in ('organique', 'direct', 'prescription', 'media', 'affiliation', 'partenariat', 'event')}
        for d, utms in cache.items():
            if wk_start <= d <= wk_end:
                for u in utms:
                    counts[_classify_canal(u)] += 1.0 if mode == 'count' else u.get('_valeur', 1.0)
        # Fusionner SEO (organique) dans Direct pour l'affichage
        counts['direct'] += counts.pop('organique', 0)
        return counts

    def _override_lead_presc(counts, presc_acts):
        """Leads : remplace prescription UTM par _presc_activations_by_week.
        Direct = total UTM hors presc − marketing UTM (pas de total CSV fiable pour les leads)."""
        mkt_utm = sum(counts.get(c, 0) for c in _MKT_CANAUX_CC)
        utm_direct_orig = counts.get('direct', 0)
        counts['prescription'] = presc_acts
        # Direct inchangé (on ne dispose pas d'un total lead CSV exhaustif)
        return counts

    def _override_conv_presc_direct(counts, presc_csv, total_csv):
        """Conversions : remplace prescription et direct par les valeurs CSV (file-diff)."""
        mkt_utm = sum(counts.get(c, 0) for c in _MKT_CANAUX_CC)
        counts['prescription'] = presc_csv
        counts['direct']       = max(0, total_csv - presc_csv - mkt_utm)
        return counts

    def _csv_counts_for_week(wk_dt, mode='eur'):
        """Retourne (presc_acq_conv, sous_acq) depuis les CSV pour la semaine donnée (conversions)."""
        p = _acquisition_by_file_diff([wk_dt], CHART_CFG['prescriptions'], mode=mode)
        s = _acquisition_by_file_diff([wk_dt], CHART_CFG['souscriptions'], mode=mode)
        return (p[0] if p else 0), (s[0] if s else 0)

    def _presc_acts_for_week(wk_dt, mode='eur'):
        """Prescriptions activées (EER démarrés) pour la semaine = Leads."""
        curr_wk = week_start_mon(_now().date())
        res = _presc_activations_by_week([wk_dt], CHART_CFG['prescriptions'],
                                         curr_week=curr_wk, mode=mode)
        return res[0] if res else 0

    week_dt      = datetime.strptime(week_start, '%Y-%m-%d').date()
    presc_acts   = _presc_acts_for_week(week_dt)
    presc_csv, sous_csv = _csv_counts_for_week(week_dt)

    canaux_lead = _override_lead_presc(_count_by_canal(utm_lead, week_start, week_end), presc_acts)
    canaux_conv = _override_conv_presc_direct(_count_by_canal(utm_conv, week_start, week_end), presc_csv, sous_csv)

    # Volumes (nombre de dossiers) : servent au CPL / CPA, qui restent des coûts par unité
    presc_acts_n           = _presc_acts_for_week(week_dt, mode='count')
    presc_csv_n, sous_csv_n = _csv_counts_for_week(week_dt, mode='count')
    canaux_lead_n = _override_lead_presc(
        _count_by_canal(utm_lead, week_start, week_end, mode='count'), presc_acts_n)
    canaux_conv_n = _override_conv_presc_direct(
        _count_by_canal(utm_conv, week_start, week_end, mode='count'), presc_csv_n, sous_csv_n)

    prev_lead = prev_conv = prev_label = None
    prev_lead_n = prev_conv_n = None
    if week_idx > 0:
        prev_start  = week_starts[week_idx - 1]
        prev_end    = (datetime.strptime(prev_start, '%Y-%m-%d').date() + timedelta(days=6)).isoformat()
        prev_dt     = datetime.strptime(prev_start, '%Y-%m-%d').date()
        prev_presc_acts    = _presc_acts_for_week(prev_dt)
        ppresc_csv, psous_csv = _csv_counts_for_week(prev_dt)
        prev_lead   = _override_lead_presc(_count_by_canal(utm_lead, prev_start, prev_end), prev_presc_acts)
        prev_conv   = _override_conv_presc_direct(_count_by_canal(utm_conv, prev_start, prev_end), ppresc_csv, psous_csv)
        prev_label  = _wk_label(prev_dt)
        prev_presc_acts_n        = _presc_acts_for_week(prev_dt, mode='count')
        ppresc_csv_n, psous_csv_n = _csv_counts_for_week(prev_dt, mode='count')
        prev_lead_n = _override_lead_presc(
            _count_by_canal(utm_lead, prev_start, prev_end, mode='count'), prev_presc_acts_n)
        prev_conv_n = _override_conv_presc_direct(
            _count_by_canal(utm_conv, prev_start, prev_end, mode='count'), ppresc_csv_n, psous_csv_n)

    # ── Budgets par canal ────────────────────────────────────────────────────
    # Lire les lignes Excel (même logique que api_budget)
    def _budgets_for_idx(wi):
        bud = {'organique': None, 'prescription': None, 'media': None, 'affiliation': None, 'partenariat': None, 'event': None}
        if openpyxl is None or not BUDGET_EXCEL.exists():
            return bud
        try:
            wb   = openpyxl.load_workbook(BUDGET_EXCEL, data_only=True)
            ws_b = wb['Suivi hebdo FULL']
            rows_b = list(ws_b.iter_rows(values_only=True))

            def _val(lbl_match):
                total = 0.0
                for ri in range(4, 15):
                    if ri >= len(rows_b):
                        break
                    r = rows_b[ri]
                    lbl = str(r[0]) if r[0] else ''
                    if lbl_match(lbl):
                        bci = wi + 3
                        v   = r[bci] if bci < len(r) else None
                        total += float(v) if isinstance(v, (int, float)) else 0.0
                return total

            # Média : remplacé par total plan média (nova + portail)
            _, _, nova_total, portail_total = _load_media_data(N_DISP)
            bud['media']       = nova_total[wi] + portail_total[wi]
            bud['affiliation'] = _val(lambda l: l == 'Affiliation')
            bud['partenariat'] = _val(lambda l: l == 'Partenariat Perf')
            bud['event']       = _val(lambda l: l.lower().startswith('event')) or None
        except Exception:
            pass
        return bud

    budgets      = _budgets_for_idx(week_idx)
    prev_budgets = _budgets_for_idx(week_idx - 1) if week_idx > 0 else None

    return jsonify({
        'week_start':   week_start,
        'week_end':     week_end,
        'week_label':   _wk_label(datetime.strptime(week_start, '%Y-%m-%d').date()),
        'prev_label':   prev_label,
        'dates':        week_starts,
        'lead':         canaux_lead,
        'conv':         canaux_conv,
        'lead_n':       canaux_lead_n,
        'conv_n':       canaux_conv_n,
        'unit':         'eur',
        'prev_lead':    prev_lead,
        'prev_conv':    prev_conv,
        'prev_lead_n':  prev_lead_n,
        'prev_conv_n':  prev_conv_n,
        'budgets':      budgets,
        'prev_budgets': prev_budgets,
    })


@app.route('/api/sanity_check', methods=['POST'])
def api_sanity_check():
    """Exporte les données Piano consenties (par jour, par media) dans l'onglet
    'Sanity Check per day' du Google Sheet.
    Pour chaque colonne de date sans PIANO adjacent, insère une colonne PIANO à droite."""
    return jsonify({'ok': False, 'error': 'Sanity check Google Sheet désactivé en mode démo'}), 400


@app.route('/api/traffic/bmedia_sub')
def api_traffic_bmedia_sub():
    """Trafic Piano par source (pmax/meta/sea_brand/sea_nb) et type d'URL (lp/init/lead).
    Retourne les visites/visiteurs agrégés par date pour la plage demandée."""
    source   = request.args.get('source', '')
    url_type = request.args.get('url_type', '')
    start    = request.args.get('start')
    end      = request.args.get('end')

    if not start or not end:
        return jsonify({'error': 'start/end requis'}), 400

    if url_type not in _BMEDIA_URL_FILTERS:
        return jsonify({'error': f'url_type invalide: {url_type}'}), 400
    if source not in _BMEDIA_SPF:
        return jsonify({'error': f'source invalide: {source}'}), 400

    # LEAD et CONVERSION : données UTM au lieu de Piano
    if url_type in ('lead', 'conv'):
        filter_fn = _UTM_FILTERS.get(source)
        if not filter_fn:
            return jsonify({'visits': 0, 'visitors': 0})
        utm_by_date = _load_utm_cache(conv=(url_type == 'conv'))
        count = sum(
             u.get('_valeur', 1.0) for d, utms in utm_by_date.items()
            if start <= d <= end
            for u in utms
            if filter_fn(u)
        )
        return jsonify({'visits': count, 'visitors': count})

    combined_filter = _bmedia_combined_filter(source, url_type)
    if request.args.get('consent') == '1':
        combined_filter = {'$and': [combined_filter, CONSENT_FILTER]}

    try:
        rows = _fetch_traffic(start, end, url_filter=None, secondary_property_filter=combined_filter)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    total_visits   = sum(r.get('m_visits', 0) or 0 for r in rows)
    total_visitors = sum(r.get('m_unique_visitors', 0) or 0 for r in rows)
    return jsonify({'visits': total_visits, 'visitors': total_visitors})


@app.route('/api/traffic')
def api_traffic():
    try:
        start, end, start_dt, end_dt, url_filter, secondary_property_filter = _parse_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    n             = (end_dt - start_dt).days + 1
    prev_end_dt   = start_dt - timedelta(days=1)
    prev_start_dt = prev_end_dt - timedelta(days=n - 1)

    # LEAD et CONVERSION : données UTM
    if request.args.get('lead') == '1' or request.args.get('conv') == '1':
        conv = request.args.get('conv') == '1'
        rows      = _utm_to_piano_rows(start, end, conv_only=conv)
        prev_rows = _utm_to_piano_rows(str(prev_start_dt), str(prev_end_dt), conv_only=conv)
        return jsonify({
            'rows': rows, 'prev_rows': prev_rows,
            'start': start, 'end': end,
            'prev_start': str(prev_start_dt), 'prev_end': str(prev_end_dt),
        })

    def _apply(rows):
        out = []
        for r in rows:
            url = r.get('event_url_path', '') or ''
            if url_filter and url_filter not in url:
                continue
            if secondary_property_filter and not _eval_piano_filter(r, secondary_property_filter):
                continue
            out.append(r)
        return out

    try:
        rows      = _apply(_fetch_range_cached(start, end, AT_SITE,
                                               realtime_url_filter=url_filter,
                                               realtime_spf=secondary_property_filter))
        prev_rows = _apply(_fetch_range_cached(str(prev_start_dt), str(prev_end_dt), AT_SITE))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({
        'rows': rows, 'prev_rows': prev_rows,
        'start': start, 'end': end,
        'prev_start': str(prev_start_dt), 'prev_end': str(prev_end_dt),
    })

def _parse_traffic_params():
    """Extrait et valide les paramètres communs aux routes /api/traffic et /api/traffic/export."""
    start = request.args.get('start')
    end   = request.args.get('end')
    if not start or not end:
        end_dt   = _now().date()
        start_dt = end_dt - timedelta(days=6)
        start, end = str(start_dt), str(end_dt)
    start_dt = datetime.strptime(start, '%Y-%m-%d').date()
    end_dt   = datetime.strptime(end,   '%Y-%m-%d').date()

    lead_only      = request.args.get('lead')      == '1'
    trafic_only    = request.args.get('trafic')    == '1'
    lp_only        = request.args.get('lp')        == '1'
    pmax_only        = request.args.get('pmax')        == '1'
    meta_only        = request.args.get('meta')        == '1'
    sea_brand_only   = request.args.get('sea_brand')   == '1'
    sea_nb_only      = request.args.get('sea_nb')      == '1'
    presc_only       = request.args.get('presc')       == '1'
    linkedin_only    = request.args.get('linkedin')    == '1'
    tiktok_only      = request.args.get('tiktok')      == '1'
    amazon_only      = request.args.get('amazon')      == '1'
    affiliation_only  = request.args.get('affiliation')  == '1'
    partenariat_only  = request.args.get('partenariat')  == '1'

    consent_only = request.args.get('consent') == '1'
    url_filter   = LEAD_URL if lead_only else (TRAFIC_URL if trafic_only else (LP_URL if lp_only else None))

    if sea_nb_only:
        secondary_property_filter = {'$or': [
            {'$and': [{'event_url_full': {'$lk': 'lancement-nova-01_gg_search_M'}},
                      {'event_url_full': {'$lk': 'at_variant=generique'}}]},
            {'$and': [{'src_detail': {'$eq': 'lancement-nova-01_gg_search_M'}},
                      {'src_variant': {'$lk': 'generique'}}]},
        ]}
    elif sea_brand_only:
        secondary_property_filter = {'$or': [
            {'$and': [{'event_url_full': {'$lk': 'lancement-nova-01_gg_search_M'}},
                      {'event_url_full': {'$lk': 'at_variant=marque'}}]},
            {'$and': [{'src_detail': {'$eq': 'lancement-nova-01_gg_search_M'}},
                      {'src_variant': {'$lk': 'marque'}}]},
        ]}
    elif meta_only:
        secondary_property_filter = {'$or': [
            {'event_url_full': {'$lk': META_URL}},
            {'src_detail': {'$eq': 'nova_q1_meta'}},
        ]}
    elif pmax_only:
        secondary_property_filter = {'$or': [
            {'event_url_full': {'$lk': PMAX_URL}},
            {'src_detail': {'$eq': 'nova_q1_gg_pmax'}},
        ]}
    elif presc_only:
        secondary_property_filter = _BMEDIA_SPF.get('presc', {'$or': [
            {'event_url_full': {'$lk': PRESC_URL}},
            {'src_detail': {'$eq': 'prescription-nova-pro'}},
        ]})
    elif linkedin_only:
        secondary_property_filter = _BMEDIA_SPF['linkedin']
    elif tiktok_only:
        secondary_property_filter = _BMEDIA_SPF['tiktok']
    elif amazon_only:
        secondary_property_filter = _BMEDIA_SPF['amazon']
    elif affiliation_only:
        secondary_property_filter = _BMEDIA_SPF['affiliation']
    elif partenariat_only:
        secondary_property_filter = _BMEDIA_SPF['partenariat']
    else:
        secondary_property_filter = None

    # Filtre consentement : combiné avec tout filtre secondaire existant
    if consent_only:
        if secondary_property_filter:
            secondary_property_filter = {'$and': [secondary_property_filter, CONSENT_FILTER]}
        else:
            secondary_property_filter = CONSENT_FILTER

    return start, end, start_dt, end_dt, url_filter, secondary_property_filter


@app.route('/api/traffic/pie')
def api_traffic_pie():
    """Répartition par catégorie PIE — classification locale sur données mises en cache."""
    try:
        start, end, start_dt, end_dt, url_filter, _ = _parse_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    metric         = request.args.get('metric', 'visits')
    field          = 'm_unique_visitors' if metric == 'visitors' else 'm_visits'
    consent_filter = CONSENT_FILTER if request.args.get('consent') == '1' else None

    # LEAD et CONVERSION : données UTM
    if request.args.get('lead') == '1' or request.args.get('conv') == '1':
        rows = _utm_to_piano_rows(start, end, conv_only=(request.args.get('conv') == '1'))
    else:
        try:
            all_rows = _fetch_range_cached(start, end, AT_SITE,
                                           realtime_url_filter=url_filter,
                                           realtime_spf=consent_filter)
            rows = [r for r in all_rows if not url_filter or url_filter in (r.get('event_url_path', '') or '')]
            if consent_filter:
                rows = [r for r in rows if _eval_piano_filter(r, consent_filter)]
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    CLASSIFIERS = [
        ('PMAX_PIANO',        lambda r: r.get('src_detail','') == 'nova_q1_gg_pmax'),
        ('PMAX_CORRIGE',      lambda r: PMAX_URL in (r.get('at_channel','') or '').lower()
                                        or PMAX_URL in (r.get('at_campaign','') or '').lower()),
        ('META_PIANO',        lambda r: r.get('src_detail','') == 'nova_q1_meta'),
        ('META_CORRIGE',      lambda r: r.get('at_campaign','') == META_URL),
        ('SEA BRAND_PIANO',   lambda r: r.get('src_detail','') == 'lancement-nova-01_gg_search_M'
                                        and 'marque'    in (r.get('src_variant','') or '').lower()),
        ('SEA BRAND_CORRIGE', lambda r: r.get('at_campaign','') == 'lancement-nova-01_gg_search_M'
                                        and r.get('at_variant','') == 'marque'),
        ('SEA NB_PIANO',      lambda r: r.get('src_detail','') == 'lancement-nova-01_gg_search_M'
                                        and 'generique' in (r.get('src_variant','') or '').lower()),
        ('SEA NB_CORRIGE',    lambda r: r.get('at_campaign','') == 'lancement-nova-01_gg_search_M'
                                        and 'generique' in (r.get('at_variant','') or '')),
        ('AGENCE_PIANO',       lambda r: r.get('src_detail','') == 'prescription-nova-pro'),
        ('AGENCE_CORRIGE',     lambda r: r.get('at_campaign','') == PRESC_URL),
        ('LINKEDIN_PIANO',    lambda r: r.get('src_detail','') == LINKEDIN_URL),
        ('LINKEDIN_CORRIGE',  lambda r: r.get('at_campaign','') == LINKEDIN_URL),
        ('TIKTOK_PIANO',      lambda r: r.get('src_detail','') == TIKTOK_URL),
        ('TIKTOK_CORRIGE',    lambda r: r.get('at_campaign','') == TIKTOK_URL),
        ('AMAZON_PIANO',      lambda r: r.get('src_detail','') == AMAZON_URL),
        ('AMAZON_CORRIGE',    lambda r: r.get('at_campaign','') == AMAZON_URL),
        ('AFFILIATION_PIANO',   lambda r: r.get('src','')       == AFFILIATION_URL),
        ('AFFILIATION_CORRIGE', lambda r: r.get('at_medium','') == AFFILIATION_URL),
        ('PARTENARIAT_PIANO',   lambda r: r.get('src','')       == PARTENARIAT_URL),
        ('PARTENARIAT_CORRIGE', lambda r: r.get('at_medium','') == PARTENARIAT_URL),
    ]

    totals         = {label: 0 for label, _ in CLASSIFIERS}
    organique_total              = 0
    organique_web_total          = 0
    organique_sans_tracking_total = 0
    autres_src     = {}   # src → volume pour la décomposition AUTRES
    grand_total    = 0
    by_date        = {}   # date_str → {label: value}

    for row in rows:
        val      = row.get(field, 0) or 0
        date_key = row.get('date', '')
        grand_total += val
        if date_key not in by_date:
            by_date[date_key] = {}
        matched = False
        for label, test in CLASSIFIERS:
            if test(row):
                totals[label] += val
                by_date[date_key][label] = by_date[date_key].get(label, 0) + val
                matched = True
                break
        if not matched:
            src = row.get('src') or 'inconnu'
            if src == 'organique_web':
                organique_web_total += val
                by_date[date_key]['ORGANIQUE WEB'] = by_date[date_key].get('ORGANIQUE WEB', 0) + val
            elif src == 'organique_sans_tracking':
                organique_sans_tracking_total += val
                by_date[date_key]['ORGANIQUE SANS TRACKING'] = by_date[date_key].get('ORGANIQUE SANS TRACKING', 0) + val
            elif src == 'organique':
                # Trafic Piano organique (non-UTM) — conservé pour compatibilité
                organique_total += val
                by_date[date_key]['ORGANIQUE'] = by_date[date_key].get('ORGANIQUE', 0) + val
            else:
                autres_src[src] = autres_src.get(src, 0) + val
                by_date[date_key]['AUTRES'] = by_date[date_key].get('AUTRES', 0) + val

    # Parts principales
    slices = [{'label': k, 'value': v} for k, v in totals.items() if v > 0]
    if organique_web_total > 0:
        slices.append({'label': 'ORGANIQUE WEB', 'value': organique_web_total})
    if organique_sans_tracking_total > 0:
        slices.append({'label': 'ORGANIQUE SANS TRACKING', 'value': organique_sans_tracking_total})
    if organique_total > 0:
        slices.append({'label': 'ORGANIQUE', 'value': organique_total})
    autres_total = sum(autres_src.values())
    if autres_total > 0:
        slices.append({'label': 'AUTRES', 'value': autres_total})

    # Décomposition AUTRES par source (pour le mode détail côté frontend)
    autres_slices = [
        {'label': f'AUTRES · {src}', 'value': v}
        for src, v in sorted(autres_src.items(), key=lambda x: -x[1]) if v > 0
    ]

    return jsonify({'slices': slices, 'autres_slices': autres_slices, 'total': grand_total,
                    'by_date': by_date, 'start': start, 'end': end})


@app.route('/api/traffic/pie/export')
def api_traffic_pie_export():
    """Exporte un CSV classifié PIE (PIANO/CORRIGÉ/AUTRES) depuis le cache local."""
    try:
        start, end, start_dt, end_dt, url_filter, _ = _parse_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400
    consent_filter = CONSENT_FILTER if request.args.get('consent') == '1' else None

    export_cols = ['categorie', 'src', 'src_detail', 'src_campaign', 'src_creation',
                   'src_variant', 'device_type', 'event_url_path',
                   'at_medium', 'at_campaign', 'at_channel', 'at_variant',
                   'date', 'm_visits', 'm_unique_visitors']

    # Classifieurs TM (identiques à api_traffic_pie)
    CLASSIFIERS = [
        ('PMAX_PIANO',        lambda r: r.get('src_detail','') == 'nova_q1_gg_pmax'),
        ('PMAX_CORRIGE',      lambda r: PMAX_URL in (r.get('at_channel','') or '').lower()
                                        or PMAX_URL in (r.get('at_campaign','') or '').lower()),
        ('META_PIANO',        lambda r: r.get('src_detail','') == 'nova_q1_meta'),
        ('META_CORRIGE',      lambda r: r.get('at_campaign','') == META_URL),
        ('SEA BRAND_PIANO',   lambda r: r.get('src_detail','') == 'lancement-nova-01_gg_search_M'
                                        and 'marque'    in (r.get('src_variant','') or '').lower()),
        ('SEA BRAND_CORRIGE', lambda r: r.get('at_campaign','') == 'lancement-nova-01_gg_search_M'
                                        and r.get('at_variant','') == 'marque'),
        ('SEA NB_PIANO',      lambda r: r.get('src_detail','') == 'lancement-nova-01_gg_search_M'
                                        and 'generique' in (r.get('src_variant','') or '').lower()),
        ('SEA NB_CORRIGE',    lambda r: r.get('at_campaign','') == 'lancement-nova-01_gg_search_M'
                                        and 'generique' in (r.get('at_variant','') or '')),
        ('AGENCE_PIANO',       lambda r: r.get('src_detail','') == 'prescription-nova-pro'),
        ('AGENCE_CORRIGE',     lambda r: r.get('at_campaign','') == PRESC_URL),
        ('LINKEDIN_PIANO',    lambda r: r.get('src_detail','') == LINKEDIN_URL),
        ('LINKEDIN_CORRIGE',  lambda r: r.get('at_campaign','') == LINKEDIN_URL),
        ('TIKTOK_PIANO',      lambda r: r.get('src_detail','') == TIKTOK_URL),
        ('TIKTOK_CORRIGE',    lambda r: r.get('at_campaign','') == TIKTOK_URL),
        ('AMAZON_PIANO',      lambda r: r.get('src_detail','') == AMAZON_URL),
        ('AMAZON_CORRIGE',    lambda r: r.get('at_campaign','') == AMAZON_URL),
        ('AFFILIATION_PIANO',   lambda r: r.get('src','')       == AFFILIATION_URL),
        ('AFFILIATION_CORRIGE', lambda r: r.get('at_medium','') == AFFILIATION_URL),
        ('PARTENARIAT_PIANO',   lambda r: r.get('src','')       == PARTENARIAT_URL),
        ('PARTENARIAT_CORRIGE', lambda r: r.get('at_medium','') == PARTENARIAT_URL),
    ]

    def _classify(r):
        for label, test in CLASSIFIERS:
            if test(r):
                return label
        return 'AUTRES'

    try:
        all_rows = _fetch_range_cached(start, end, AT_SITE,
                                       realtime_url_filter=url_filter,
                                       realtime_spf=consent_filter)
        rows = [r for r in all_rows if not url_filter or url_filter in (r.get('event_url_path', '') or '')]
        if consent_filter:
            rows = [r for r in rows if _eval_piano_filter(r, consent_filter)]
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    import io
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=export_cols, delimiter=';', extrasaction='ignore')
    writer.writeheader()
    for row in sorted(rows, key=lambda r: (r.get('date', ''), r.get('src', ''))):
        row['categorie'] = _classify(row)
        writer.writerow(row)

    fname = f"trafic_pie_{start}_{end}.csv"
    from flask import make_response
    resp = make_response(out.getvalue())
    resp.headers['Content-Type']        = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@app.route('/api/traffic/export')
def api_traffic_export():
    """Retourne un CSV avec event_url_path + colonnes at_* + toutes les dimensions, depuis le cache local."""
    try:
        start, end, start_dt, end_dt, url_filter, secondary_property_filter = _parse_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    export_cols = ['src', 'src_detail', 'src_campaign', 'src_creation', 'src_variant',
                   'device_type', 'event_url_path',
                   'at_medium', 'at_campaign', 'at_channel', 'at_variant',
                   'date', 'm_visits', 'm_unique_visitors']

    try:
        all_rows = _fetch_range_cached(start, end, AT_SITE,
                                       realtime_url_filter=url_filter,
                                       realtime_spf=secondary_property_filter)
        rows = []
        for r in all_rows:
            if url_filter and url_filter not in (r.get('event_url_path', '') or ''):
                continue
            if secondary_property_filter and not _eval_piano_filter(r, secondary_property_filter):
                continue
            rows.append(r)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    import io
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=export_cols, delimiter=';', extrasaction='ignore')
    writer.writeheader()
    for row in sorted(rows, key=lambda r: (r.get('date', ''), r.get('src', ''))):
        writer.writerow(row)

    fname = f"trafic_{start}_{end}.csv"
    from flask import make_response
    resp = make_response(out.getvalue())
    resp.headers['Content-Type']        = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


# ═══════════════════════════════════════════════════════════════════════════════
# PORTAIL ENTREPRENEUR (site Piano 100002)
# ═══════════════════════════════════════════════════════════════════════════════
def _parse_pe_traffic_params():
    """Extrait et valide les paramètres de dates et filtres pour les routes /api/pe/."""
    start = request.args.get('start')
    end   = request.args.get('end')
    if not start or not end:
        end_dt   = _now().date()
        start_dt = end_dt - timedelta(days=6)
        start, end = str(start_dt), str(end_dt)
    start_dt     = datetime.strptime(start, '%Y-%m-%d').date()
    end_dt       = datetime.strptime(end,   '%Y-%m-%d').date()
    consent_only = request.args.get('consent') == '1'
    secondary_property_filter = CONSENT_FILTER if consent_only else None
    return start, end, start_dt, end_dt, secondary_property_filter


# Classifieurs PIE pour le Portail Entrepreneur (site 100002)
# Logique URL uniquement — pas de fallback src_detail sauf pour INFLUENCE
_PE_CLASSIFIERS = [
    ('META',              lambda r: 'meta'          in _row_url_str(r).lower()),
    ('GOOGLE',            lambda r: 'google_search' in _row_url_str(r)),
    ('LINKEDIN',          lambda r: 'linkedin'      in _row_url_str(r)),
    ('INFLUENCE_PIANO',   lambda r: r.get('src_detail','') == 'ifc-portail-Q1'),
    ('INFLUENCE_CORRIGE', lambda r: r.get('at_campaign','') == 'ifc-portail-Q1'),
]


@app.route('/api/pe/traffic')
def api_pe_traffic():
    try:
        start, end, start_dt, end_dt, secondary_property_filter = _parse_pe_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    n             = (end_dt - start_dt).days + 1
    prev_end_dt   = start_dt - timedelta(days=1)
    prev_start_dt = prev_end_dt - timedelta(days=n - 1)

    def _strip_apply(rows):
        out = []
        for r in rows:
            if secondary_property_filter and not _eval_piano_filter(r, secondary_property_filter):
                continue
            out.append(r)
        return out

    try:
        rows      = _strip_apply(_fetch_range_cached(start, end, AT_SITE_PE,
                                                     realtime_spf=secondary_property_filter))
        prev_rows = _strip_apply(_fetch_range_cached(str(prev_start_dt), str(prev_end_dt), AT_SITE_PE))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({
        'rows': rows, 'prev_rows': prev_rows,
        'start': start, 'end': end,
        'prev_start': str(prev_start_dt), 'prev_end': str(prev_end_dt),
    })


@app.route('/api/pe/traffic/pie')
def api_pe_traffic_pie():
    """Répartition PIE pour le Portail Entrepreneur (site Piano 100002)."""
    try:
        start, end, start_dt, end_dt, secondary_property_filter = _parse_pe_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    metric = request.args.get('metric', 'visits')
    field  = 'm_unique_visitors' if metric == 'visitors' else 'm_visits'

    try:
        all_rows = _fetch_range_cached(start, end, AT_SITE_PE, realtime_spf=secondary_property_filter)
        rows = [r for r in all_rows
                if not secondary_property_filter or _eval_piano_filter(r, secondary_property_filter)]
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    from urllib.parse import urlparse, parse_qs as _parse_qs

    totals        = {label: 0 for label, _ in _PE_CLASSIFIERS}
    autres_src    = {}
    grand_total   = 0
    by_date       = {}
    influence_aff = {}   # at_aff_identifier → volume (pour les lignes INFLUENCE)

    for row in rows:
        val      = row.get(field, 0) or 0
        date_key = row.get('date', '')
        grand_total += val
        if date_key not in by_date:
            by_date[date_key] = {}
        matched = False
        for label, test in _PE_CLASSIFIERS:
            if test(row):
                totals[label] += val
                by_date[date_key][label] = by_date[date_key].get(label, 0) + val
                if 'INFLUENCE' in label:
                    aff_id = row.get('at_aff_identifier') or '(non renseigné)'
                    influence_aff[aff_id] = influence_aff.get(aff_id, 0) + val
                matched = True
                break
        if not matched:
            src = row.get('src') or 'inconnu'
            autres_src[src] = autres_src.get(src, 0) + val
            by_date[date_key]['AUTRES'] = by_date[date_key].get('AUTRES', 0) + val

    slices = [{'label': k, 'value': v} for k, v in totals.items() if v > 0]
    autres_total = sum(autres_src.values())
    if autres_total > 0:
        slices.append({'label': 'AUTRES', 'value': autres_total})

    autres_slices = [
        {'label': f'AUTRES · {src}', 'value': v}
        for src, v in sorted(autres_src.items(), key=lambda x: -x[1]) if v > 0
    ]

    influence_aff_sorted = [
        {'id': k, 'value': v}
        for k, v in sorted(influence_aff.items(), key=lambda x: -x[1])
    ]

    return jsonify({'slices': slices, 'autres_slices': autres_slices, 'total': grand_total,
                    'by_date': by_date, 'start': start, 'end': end,
                    'influence_aff': influence_aff_sorted})


@app.route('/api/pe/traffic/pie/export')
def api_pe_traffic_pie_export():
    """Exporte un CSV classifié PIE pour le Portail Entrepreneur (site 100002)."""
    try:
        start, end, start_dt, end_dt, secondary_property_filter = _parse_pe_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    export_cols = ['categorie', 'src', 'src_detail', 'src_campaign', 'src_creation',
                   'src_variant', 'device_type', 'event_url_path',
                   'at_medium', 'at_campaign', 'at_channel', 'at_variant',
                   'date', 'm_visits', 'm_unique_visitors']

    def _classify_pe(r):
        for label, test in _PE_CLASSIFIERS:
            if test(r):
                return label
        return 'AUTRES'

    try:
        all_rows = _fetch_range_cached(start, end, AT_SITE_PE, realtime_spf=secondary_property_filter)
        rows = [r for r in all_rows
                if not secondary_property_filter or _eval_piano_filter(r, secondary_property_filter)]
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    import io
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=export_cols, delimiter=';', extrasaction='ignore')
    writer.writeheader()
    for row in sorted(rows, key=lambda r: (r.get('date', ''), r.get('src', ''))):
        row['categorie'] = _classify_pe(row)
        writer.writerow(row)

    fname = f"pe_trafic_pie_{start}_{end}.csv"
    from flask import make_response
    resp = make_response(out.getvalue())
    resp.headers['Content-Type']        = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@app.route('/api/pe/traffic/export')
def api_pe_traffic_export():
    """Exporte un CSV avec toutes les dimensions pour le Portail Entrepreneur (site 100002)."""
    try:
        start, end, start_dt, end_dt, secondary_property_filter = _parse_pe_traffic_params()
    except ValueError:
        return jsonify({'error': 'Dates invalides'}), 400

    export_cols = ['src', 'src_detail', 'src_campaign', 'src_creation', 'src_variant',
                   'device_type', 'event_url_path',
                   'at_medium', 'at_campaign', 'at_channel', 'at_variant',
                   'date', 'm_visits', 'm_unique_visitors']

    try:
        all_rows = _fetch_range_cached(start, end, AT_SITE_PE, realtime_spf=secondary_property_filter)
        rows = [r for r in all_rows
                if not secondary_property_filter or _eval_piano_filter(r, secondary_property_filter)]
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    import io
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=export_cols, delimiter=';', extrasaction='ignore')
    writer.writeheader()
    for row in sorted(rows, key=lambda r: (r.get('date', ''), r.get('src', ''))):
        writer.writerow(row)

    fname = f"pe_trafic_{start}_{end}.csv"
    from flask import make_response
    resp = make_response(out.getvalue())
    resp.headers['Content-Type']        = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@app.route('/api/projection/save', methods=['POST', 'OPTIONS'])
def api_projection_save():
    """Écrit les volumes 'nouveaux clients' du modèle de projection dans le fichier Excel
    Nova - distribution_volume_cible.xlsx (colonne 'Volume Client').
    Mapping : HTML S04 (index 3) → Excel S1 (ligne 2), décalage de 3 semaines."""
    # CORS pour permettre l'appel depuis file:// ou localhost
    if request.method == 'OPTIONS':
        resp = jsonify({'ok': True})
        resp.headers['Access-Control-Allow-Origin']  = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp

    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé'}), 500

    dist_file = DATA_DIR / 'projections' / 'Nova - distribution_volume_cible.xlsx'
    if not dist_file.exists():
        return jsonify({'error': f'Fichier introuvable : {dist_file}'}), 404

    data = request.get_json(silent=True) or {}
    week_clients = data.get('weekClients', [])   # 52 valeurs (S01–S52, index 0–51)
    if not week_clients or len(week_clients) != 52:
        return jsonify({'error': 'weekClients doit contenir exactement 52 valeurs'}), 400

    try:
        wb = openpyxl.load_workbook(dist_file)   # data_only=False → préserve les formules
        ws = wb['Distribution Volume']

        # Ligne 1 = en-tête, lignes 2–51 = semaines Excel 1–50
        # Excel S(n) ← HTML S(n+3) = weekClients[n+2] (0-indexed)
        # Ex : Excel S1 → weekClients[3] = HTML S04 ; Excel S49 → weekClients[51] = HTML S52
        # Excel S50 n'a pas de semaine HTML correspondante → on y écrit 0 (au lieu de skip)
        TARGET_TOTAL = 50000
        updated = 0
        written_rows = []   # (excel_row, valeur arrondie) pour correction finale

        for excel_row in range(2, ws.max_row + 1):
            week_num_cell = ws.cell(excel_row, 1).value   # colonne A = numéro de semaine
            try:
                week_num = int(str(week_num_cell))        # 1 à 50
            except (TypeError, ValueError):
                continue                                   # TOTAL ou ligne vide
            html_index = week_num + 2                     # S04=index3 → S1: 1+2=3 ✓
            if html_index >= len(week_clients):
                # Plus de semaines modèle disponibles → mettre explicitement 0
                ws.cell(excel_row, 4).value = 0
                updated += 1
                continue
            # Le dashboard envoie des montants € → le fichier stocke des contrats
            val = round(week_clients[html_index] / VALEUR_MOYENNE)
            ws.cell(excel_row, 4).value = val
            written_rows.append((excel_row, val))
            updated += 1

        # Correction d'arrondi : ajuster la valeur max pour que la somme = TARGET_TOTAL exactement
        if written_rows:
            current_sum = sum(v for _, v in written_rows)
            diff = TARGET_TOTAL - current_sum
            if diff != 0:
                max_row = max(written_rows, key=lambda x: x[1])[0]
                ws.cell(max_row, 4).value = (ws.cell(max_row, 4).value or 0) + diff

        wb.save(dist_file)
        resp = jsonify({'ok': True, 'updated': updated,
                        'file': dist_file.name})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

    except Exception as e:
        resp = jsonify({'error': str(e)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500


_VALIDE_STATUTS = {'Dossier validé', 'Validé'}   # ancien et nouveau format backoffice
_DATE_MIN_COHORTE = datetime(2026, 5, 3).date()

def _find_sous_exports(day_filter=None):
    """Retourne la liste triée (label, path) des CSV souscriptions _total.
    day_filter : ensemble de weekday ints (0=Lun … 6=Dim). None = tous les jours."""
    snapshots = []
    for entry in sorted(EXPORTS_DIR.iterdir()):
        try:
            d = datetime.strptime(entry.name, '%Y.%m.%d').date()
        except ValueError:
            continue
        if day_filter is not None and d.weekday() not in day_filter:
            continue
        sous_dir = entry / 'Souscriptions'
        if not sous_dir.is_dir():
            continue
        for f in sous_dir.iterdir():
            if 'total' in f.stem.lower() and f.suffix.lower() == '.csv':
                snapshots.append((d.strftime('%d/%m'), f))
                break
    return snapshots

def _build_cohorts(snapshots, week_fn):
    """Construit les données de cohorte depuis la liste (label, path)."""
    all_weeks, results = set(), []
    for label, fpath in snapshots:
        wt, wv = defaultdict(float), defaultdict(float)
        _ccfg = {'valeur_col': _valeur_col_of(fpath)}
        try:
            with open(fpath, encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                next(reader, None)
                for row in reader:
                    if len(row) > 3 and row[1]:
                        try:
                            dt = datetime.strptime(_clean_date_cell(row[1]), '%d/%m/%Y').date()
                            statut = row[3].strip()
                            if dt >= _DATE_MIN_COHORTE and statut:
                                w = week_fn(dt)
                                _v = _row_value(row, _ccfg)
                                wt[w] += _v
                                if statut in _VALIDE_STATUTS:
                                    wv[w] += _v
                        except ValueError:
                            continue
        except Exception:
            continue
        cohort = {w: {'total': round(wt[w]), 'valides': round(wv[w]),
                      'taux': wv[w] / wt[w] * 100 if wt[w] else 0} for w in wt}
        all_weeks.update(cohort.keys())
        results.append((label, cohort))
    return sorted(all_weeks), results

def _write_taux_sheet_app(wb, sheet_name, title_text, subtitle, all_weeks, snap_results, week_fn):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HEADER_FILL = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
    COHORT_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    COHORT_FONT = Font(color='FFFFFF', bold=True, size=11)
    LIGHT_BLUE  = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    WHITE_FILL  = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    DARK        = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
    GREEN_FONT  = Font(color='2E7D32', bold=True, size=11)
    RED_FONT    = Font(color='C62828', bold=True, size=11)
    BORDER      = Border(left=Side(style='thin',color='B0B0B0'), right=Side(style='thin',color='B0B0B0'),
                         top=Side(style='thin',color='B0B0B0'),  bottom=Side(style='thin',color='B0B0B0'))
    CENTER      = Alignment(horizontal='center', vertical='center', wrap_text=True)
    N  = len(snap_results)
    nc = 1 + N + (N - 1)
    ws = wb.create_sheet(sheet_name)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, title_text); c.font = Font(bold=True, size=14, color='1F3864'); c.alignment = Alignment(horizontal='left')
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nc)
    c = ws.cell(2, 1, subtitle); c.font = Font(size=11, color='666666', italic=True); c.alignment = Alignment(horizontal='left')
    R = 4
    headers_def = ([("Cohorte\n(semaine du)", HEADER_FILL)] +
                   [(f"Extract\n{lbl}", HEADER_FILL) for lbl,_ in snap_results] +
                   [(f"Évolution\n{snap_results[j][0]} → {snap_results[j+1][0]}",
                     PatternFill(start_color='2E4057', end_color='2E4057', fill_type='solid'))
                    for j in range(N-1)])
    for col, (hdr, fill) in enumerate(headers_def, 1):
        c = ws.cell(R, col, hdr); c.font = HEADER_FONT; c.fill = fill; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[R].height = 35
    for i, w in enumerate(all_weeks):
        r = R + 1 + i
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE_FILL
        ws.row_dimensions[r].height = 28
        w_end = w + timedelta(days=6)
        c = ws.cell(r, 1, f"{w.strftime('%d/%m')} → {w_end.strftime('%d/%m')}")
        c.font = COHORT_FONT; c.fill = COHORT_FILL; c.alignment = CENTER; c.border = BORDER
        tv = []
        for j, (_, cohort) in enumerate(snap_results):
            if w in cohort:
                d = cohort[w]
                c = ws.cell(r, 2+j, f"{d['taux']:.1f}%\n({d['valides']}/{d['total']})")
                c.font = Font(bold=True, size=10); tv.append(d['taux'])
            else:
                c = ws.cell(r, 2+j, '—'); c.font = Font(color='999999', size=10); tv.append(None)
            c.fill = bg; c.alignment = CENTER; c.border = BORDER
        for j in range(N-1):
            t1, t2 = tv[j], tv[j+1]
            col = 2 + N + j
            if t1 is not None and t2 is not None:
                delta = t2 - t1
                c = ws.cell(r, col, f"{delta:+.1f} pts")
                c.font = GREEN_FONT if delta > 0 else (RED_FONT if delta < 0 else Font(color='666666', size=11))
            else:
                c = ws.cell(r, col, '—'); c.font = Font(color='999999', size=10)
            c.fill = bg; c.alignment = CENTER; c.border = BORDER
    r_tot = R + 1 + len(all_weeks); ws.row_dimensions[r_tot].height = 28
    c = ws.cell(r_tot, 1, 'TOTAL'); c.font = Font(bold=True, size=11, color='FFFFFF')
    c.fill = DARK; c.alignment = CENTER; c.border = BORDER
    tt = []
    for j, (_, cohort) in enumerate(snap_results):
        tot = sum(cohort[w]['total']   for w in all_weeks if w in cohort)
        val = sum(cohort[w]['valides'] for w in all_weeks if w in cohort)
        tx  = val / tot * 100 if tot else 0; tt.append(tx)
        c = ws.cell(r_tot, 2+j, f"{tx:.1f}%\n({val}/{tot})")
        c.font = Font(bold=True, size=10, color='FFFFFF'); c.fill = DARK; c.alignment = CENTER; c.border = BORDER
    for j in range(N-1):
        delta = tt[j+1] - tt[j]
        c = ws.cell(r_tot, 2+N+j, f"{delta:+.1f} pts")
        c.font = Font(bold=True, size=11, color='70AD47' if delta > 0 else ('FF4B4B' if delta < 0 else 'AAAAAA'))
        c.fill = DARK; c.alignment = CENTER; c.border = BORDER
    ws.column_dimensions['A'].width = 18
    for j in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18

def _write_volumes_sheet_app(wb, sheet_name, title_text, all_weeks, snap_results):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HEADER_FILL = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
    COHORT_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    COHORT_FONT = Font(color='FFFFFF', bold=True, size=11)
    LIGHT_BLUE  = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    WHITE_FILL  = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    BORDER      = Border(left=Side(style='thin',color='B0B0B0'), right=Side(style='thin',color='B0B0B0'),
                         top=Side(style='thin',color='B0B0B0'),  bottom=Side(style='thin',color='B0B0B0'))
    CENTER      = Alignment(horizontal='center', vertical='center', wrap_text=True)
    nc = 1 + len(snap_results) * 2
    ws = wb.create_sheet(sheet_name)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, title_text); c.font = Font(bold=True, size=14, color='1F3864'); ws.row_dimensions[1].height = 30
    R = 3
    c = ws.cell(R, 1, "Cohorte\n(semaine du)"); c.font=HEADER_FONT; c.fill=HEADER_FILL; c.alignment=CENTER; c.border=BORDER
    for j, (lbl, _) in enumerate(snap_results):
        for k, (hdr, fill) in enumerate([(f"Total\n{lbl}", HEADER_FILL),
                                          (f"Validés\n{lbl}", PatternFill(start_color='2E7D32',end_color='2E7D32',fill_type='solid'))]):
            c = ws.cell(R, 2+j*2+k, hdr); c.font=HEADER_FONT; c.fill=fill; c.alignment=CENTER; c.border=BORDER
    ws.row_dimensions[R].height = 35
    for i, w in enumerate(all_weeks):
        r = R + 1 + i; bg = LIGHT_BLUE if i % 2 == 0 else WHITE_FILL
        ws.row_dimensions[r].height = 22
        w_end = w + timedelta(days=6)
        c = ws.cell(r, 1, f"{w.strftime('%d/%m')} → {w_end.strftime('%d/%m')}")
        c.font=COHORT_FONT; c.fill=COHORT_FILL; c.alignment=CENTER; c.border=BORDER
        for j, (_, cohort) in enumerate(snap_results):
            if w in cohort:
                ct = ws.cell(r, 2+j*2, cohort[w]['total']); ct.font=Font(size=11)
                cv = ws.cell(r, 3+j*2, cohort[w]['valides']); cv.font=Font(size=11, bold=True, color='2E7D32')
            else:
                ct = ws.cell(r, 2+j*2, '—'); ct.font=Font(color='999999')
                cv = ws.cell(r, 3+j*2, '—'); cv.font=Font(color='999999')
            for cell in (ct, cv):
                cell.fill=bg; cell.alignment=CENTER; cell.border=BORDER
    ws.column_dimensions['A'].width = 18
    for j in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14

def _run_matrice_update():
    """Recalcule les 4 onglets de la matrice de cohorte.
    Jeudi-Mercredi : extractions du mercredi (2) et jeudi (3).
    Lundi-Dimanche : extractions du dimanche (6)."""
    if openpyxl is None:
        raise RuntimeError('openpyxl non installé')
    matrice_path = COMP_DIR / 'Matrice_cohorte_conversion.xlsx'
    if not matrice_path.exists():
        raise FileNotFoundError('Fichier Matrice_cohorte_conversion.xlsx introuvable')

    # Uniquement les extractions du dimanche → Lundi-Dimanche
    snaps_mon = _find_sous_exports(day_filter={6})       # Dim → Lundi-Dimanche

    wb = openpyxl.load_workbook(matrice_path)
    for t in ['Taux Jeudi-Mercredi', 'Volumes Jeudi-Mercredi',
              'Taux Lundi-Dimanche', 'Volumes Lundi-Dimanche']:
        if t in wb.sheetnames:
            del wb[t]

    if snaps_mon:
        all_weeks_m, results_m = _build_cohorts(snaps_mon, week_start_mon)
        labels_m = ' & '.join(l for l, _ in results_m)
        _write_taux_sheet_app(wb, 'Taux Lundi-Dimanche',
            f"Taux de conversion par cohorte Lundi→Dimanche ({labels_m})",
            f"Cohortes Lundi→Dimanche — Extractions : {labels_m}",
            all_weeks_m, results_m, week_start_mon)
        _write_volumes_sheet_app(wb, 'Volumes Lundi-Dimanche',
            f"Volumes bruts par cohorte Lundi→Dimanche ({labels_m})",
            all_weeks_m, results_m)

    wb.save(matrice_path)
    return len(snaps_mon)

def _excel_export_styles():
    """Retourne les styles openpyxl communs aux exports Excel du dashboard."""
    from openpyxl.styles import Font as _Font, PatternFill as _Fill, Alignment as _Align, Border as _Border, Side as _Side
    HDR_BG  = _Fill('solid', fgColor='1A2942')
    HDR_FNT = _Font(bold=True, color='FFFFFF', size=11)
    TTL_BG  = _Fill('solid', fgColor='0D3B6E')
    TOT_BG  = _Fill('solid', fgColor='0A1628')
    TOT_FNT = _Font(bold=True, color='00D4FF', size=11)
    LED_BG  = _Fill('solid', fgColor='E8F4FD')
    CNV_BG  = _Fill('solid', fgColor='E8F8E8')
    ALT_BG  = _Fill('solid', fgColor='F5F8FB')
    THIN    = _Border(left=_Side(style='thin',color='D0D0D0'), right=_Side(style='thin',color='D0D0D0'),
                      top=_Side(style='thin',color='D0D0D0'), bottom=_Side(style='thin',color='D0D0D0'))
    CTR = _Align(horizontal='center', vertical='center')
    LFT = _Align(horizontal='left',   vertical='center')

    def _h(ws, r, c, v):
        from openpyxl.styles import Font as F
        x = ws.cell(r, c, v); x.fill = HDR_BG; x.font = HDR_FNT; x.alignment = CTR; x.border = THIN
    def _c(ws, r, c, v, fill=None, bold=False, align=None):
        from openpyxl.styles import Font as F
        x = ws.cell(r, c, v)
        if fill: x.fill = fill
        x.font = F(bold=bold, size=10)
        x.alignment = align or CTR; x.border = THIN
    def _t(ws, r, c, v):
        x = ws.cell(r, c, v); x.fill = TOT_BG; x.font = TOT_FNT; x.alignment = CTR; x.border = THIN

    return dict(HDR_BG=HDR_BG, HDR_FNT=HDR_FNT, TTL_BG=TTL_BG, TOT_BG=TOT_BG, TOT_FNT=TOT_FNT,
                LED_BG=LED_BG, CNV_BG=CNV_BG, ALT_BG=ALT_BG, THIN=THIN, CTR=CTR, LFT=LFT,
                h=_h, c=_c, t=_t)


def _build_medium_excel(medium_value, title_label, filename_prefix, exclude_aff_ids=None):
    """Génère le fichier Excel Leads+Conversions pour un AT_MEDIUM donné.
    Format 4 feuilles identique à l'export Affiliation.
    Retourne un objet Response Flask."""
    import pandas as _pd
    import numpy as _np
    import openpyxl as _xl
    from openpyxl.styles import Font as _Font
    import io
    from flask import make_response

    csv_path = UTM_CACHE_DIR / 'ALL_UTM.csv'
    if not csv_path.exists():
        return jsonify({'error': 'Fichier ALL_UTM.csv introuvable'}), 404

    df_dedup = _utm_dedup_df()
    if df_dedup.empty:
        return jsonify({'error': 'Aucune donnée UTM'}), 404

    # ── Dates leads / conversions ────────────────────────────────────────────────────
    # Lead : USER_LIFECYCLE_START_DATE (toujours renseigné, aligné avec le dashboard)
    # SESSION_CREATED_AT a des nulls qui faisaient disparaître des leads du groupby
    df_dedup['ref_date'] = _pd.to_datetime(
        df_dedup['USER_LIFECYCLE_START_DATE'], utc=True, errors='coerce'
    ).dt.strftime('%Y-%m-%d')

    lc = df_dedup['USER_LIFECYCLE_STATUS'].fillna('')
    company_dt = _pd.to_datetime(df_dedup['COMPANY_CREATED_AT'], utc=True, errors='coerce').dt.strftime('%Y-%m-%d')
    step_dt    = _pd.to_datetime(df_dedup['STEP_UPDATED_AT'],    utc=True, errors='coerce').dt.strftime('%Y-%m-%d')
    df_dedup['conv_date'] = _np.where(lc == 'validated', company_dt,
                            _np.where(lc == 'validated_awaiting_agent_confirmation', step_dt, _pd.NA))

    # ── Filtrer par medium ────────────────────────────────────────────────────
    med = df_dedup[df_dedup['AT_MEDIUM'] == medium_value].copy()
    if exclude_aff_ids and 'AT_AFF_IDENTIFIER' in med.columns:
        med = med[~med['AT_AFF_IDENTIFIER'].astype(str).isin([str(x) for x in exclude_aff_ids])]
    med['_lead'] = 1
    med['_conv'] = lc[med.index].isin(['validated', 'validated_awaiting_agent_confirmation']).astype(int)

    # ── Agrégations ───────────────────────────────────────────────────────────
    # AT_AFF_IDENTIFIER existe dans le CSV (colonnes affiliation) mais est vide pour partenariat
    # → vérifier qu'il y a au moins une valeur non-nulle, sinon fallback sur AT_CAMPAIGN
    has_aff_cols = (all(c in med.columns for c in ['AT_AFF_IDENTIFIER','AT_CAMPAIGN_AFF','AT_AFF_TYPE'])
                    and med['AT_AFF_IDENTIFIER'].notna().any())

    # Remplir les NaN dans les colonnes de regroupement (pandas dropna=True par défaut dans groupby)
    for _col in ['AT_AFF_IDENTIFIER','AT_CAMPAIGN_AFF','AT_AFF_TYPE','AT_CAMPAIGN']:
        if _col in med.columns:
            med[_col] = med[_col].fillna('(non renseigné)')

    med_conv = med[med['_conv'] == 1].copy()

    if has_aff_cols:
        id_raw  = ['AT_AFF_IDENTIFIER','AT_CAMPAIGN_AFF','AT_AFF_TYPE']
        id_disp = ['Identifiant','Nom campagne','Type']
        # Lead : daté sur USER_LIFECYCLE_START_DATE
        leads_grp = med.groupby(['ref_date'] + id_raw)[['_lead']].sum().reset_index()
        leads_grp.columns = ['Date'] + id_disp + ['Leads']
        # Conversion : datée sur COMPANY_CREATED_AT (validated) / STEP_UPDATED_AT (awaiting)
        conv_grp = med_conv.groupby(['conv_date'] + id_raw)[['_conv']].sum().reset_index()
        conv_grp.columns = ['Date'] + id_disp + ['Conversions']
        grp = leads_grp.merge(conv_grp, on=['Date'] + id_disp, how='outer').fillna(0)
        for _col in ['Leads','Conversions']:
            grp[_col] = grp[_col].astype(int)
        grp = grp[['Date'] + id_disp + ['Leads','Conversions']]
        by_detail = med.groupby(id_raw)[['_lead','_conv']].sum().reset_index()
        by_detail.columns = id_disp + ['Leads','Conversions']
        by_detail = by_detail.sort_values('Leads', ascending=False)
        detail_cols = id_disp[:1]
    else:
        id_raw  = ['AT_CAMPAIGN']
        id_disp = ['Campagne']
        # Lead : daté sur USER_LIFECYCLE_START_DATE
        leads_grp = med.groupby(['ref_date'] + id_raw)[['_lead']].sum().reset_index()
        leads_grp.columns = ['Date'] + id_disp + ['Leads']
        # Conversion : datée sur COMPANY_CREATED_AT (validated) / STEP_UPDATED_AT (awaiting)
        conv_grp = med_conv.groupby(['conv_date'] + id_raw)[['_conv']].sum().reset_index()
        conv_grp.columns = ['Date'] + id_disp + ['Conversions']
        grp = leads_grp.merge(conv_grp, on=['Date'] + id_disp, how='outer').fillna(0)
        for _col in ['Leads','Conversions']:
            grp[_col] = grp[_col].astype(int)
        grp = grp[['Date'] + id_disp + ['Leads','Conversions']]
        by_detail = med.groupby(id_raw)[['_lead','_conv']].sum().reset_index()
        by_detail.columns = id_disp + ['Leads','Conversions']
        by_detail = by_detail.sort_values('Leads', ascending=False)
        detail_cols = id_disp[:1]

    grp = grp.sort_values(['Date'] + detail_cols[:1])

    # Synthèse par date : leads par USER_LIFECYCLE_START_DATE, conversions par COMPANY_CREATED_AT/STEP_UPDATED_AT
    by_date_leads = med.groupby('ref_date')[['_lead']].sum().reset_index()
    by_date_leads.columns = ['Date','Leads']
    by_date_conv = med_conv.groupby('conv_date')[['_conv']].sum().reset_index()
    by_date_conv.columns = ['Date','Conversions']
    by_date = by_date_leads.merge(by_date_conv, on='Date', how='outer').fillna(0)
    by_date['Leads']       = by_date['Leads'].astype(int)
    by_date['Conversions'] = by_date['Conversions'].astype(int)
    by_date = by_date.sort_values('Date')

    total_leads = int(grp['Leads'].sum())
    total_conv  = int(grp['Conversions'].sum())
    export_date = _now().strftime('%d/%m/%Y')

    st = _excel_export_styles()
    _h = st['h']; _c = st['c']; _t = st['t']
    TTL_BG = st['TTL_BG']; TOT_BG = st['TOT_BG']; TOT_FNT = st['TOT_FNT']
    LED_BG = st['LED_BG']; CNV_BG = st['CNV_BG']; ALT_BG = st['ALT_BG']; THIN = st['THIN']
    LFT = st['LFT']

    wb = _xl.Workbook()

    # ── Feuille 1 : Détail par jour ───────────────────────────────────────────
    ws1 = wb.active; ws1.title = 'Détail par jour'
    n_cols = len(grp.columns)
    last_col = chr(ord('A') + n_cols - 1)
    ws1.merge_cells(f'A1:{last_col}1')
    t = ws1['A1']
    t.value = f'{title_label} — Leads & Conversions par jour  (extrait {export_date})'
    t.fill = TTL_BG; t.font = _Font(bold=True, color='00D4FF', size=13); t.alignment = LFT
    ws1.row_dimensions[1].height = 26
    ws1.merge_cells(f'A2:{last_col}2')
    s = ws1['A2']
    s.value = 'Déduplication par HASHED_USER_ID · Lead daté USER_LIFECYCLE_START_DATE · Conversion datée COMPANY_CREATED_AT (validated) / STEP_UPDATED_AT (awaiting) · une même ligne peut apparaître deux fois si lead et conversion tombent à des dates différentes'
    s.font = _Font(italic=True, color='888888', size=9); s.alignment = LFT

    for c, h in enumerate(grp.columns, 1):
        _h(ws1, 3, c, h)

    for i, (_, r) in enumerate(grp.iterrows()):
        row = 4 + i; fill = ALT_BG if i % 2 == 0 else None
        for ci, col in enumerate(grp.columns, 1):
            val = r[col]
            if col in ('Leads','Conversions'):
                _c(ws1, row, ci, int(val), LED_BG if col == 'Leads' else CNV_BG, bold=True)
            else:
                _c(ws1, row, ci, str(val) if str(val) not in ('nan','NaT') else '', fill, align=LFT)

    tr = 4 + len(grp)
    merge_end = len(grp.columns) - 2
    if merge_end > 1:
        ws1.merge_cells(f'A{tr}:{chr(ord("A")+merge_end-1)}{tr}')
    _t(ws1, tr, 1, 'TOTAL')
    for ci in range(2, merge_end + 1): ws1.cell(tr, ci).fill = TOT_BG; ws1.cell(tr, ci).border = THIN
    _t(ws1, tr, merge_end + 1, total_leads); _t(ws1, tr, merge_end + 2, total_conv)
    ws1.freeze_panes = 'A4'

    # ── Feuille 2 : Par détail ────────────────────────────────────────────────
    ws2 = wb.create_sheet('Synthèse par source')
    n2 = len(by_detail.columns)
    ws2.merge_cells(f'A1:{chr(ord("A")+n2-1)}1')
    t2 = ws2['A1']
    t2.value = f'Synthèse par source — toutes dates confondues'
    t2.fill = TTL_BG; t2.font = _Font(bold=True, color='00D4FF', size=13); t2.alignment = LFT
    ws2.row_dimensions[1].height = 26

    for c, h in enumerate(by_detail.columns, 1):
        _h(ws2, 2, c, h)

    for i, (_, r) in enumerate(by_detail.iterrows()):
        row = 3 + i; fill = ALT_BG if i % 2 == 0 else None
        for ci, col in enumerate(by_detail.columns, 1):
            val = r[col]
            if col in ('Leads','Conversions'):
                _c(ws2, row, ci, int(val), LED_BG if col == 'Leads' else CNV_BG, bold=True)
            else:
                _c(ws2, row, ci, str(val) if str(val) not in ('nan','NaT') else '', fill, align=LFT)

    tr2 = 3 + len(by_detail)
    merge2_end = len(by_detail.columns) - 2
    if merge2_end > 1:
        ws2.merge_cells(f'A{tr2}:{chr(ord("A")+merge2_end-1)}{tr2}')
    _t(ws2, tr2, 1, 'TOTAL')
    for ci in range(2, merge2_end + 1): ws2.cell(tr2, ci).fill = TOT_BG; ws2.cell(tr2, ci).border = THIN
    _t(ws2, tr2, merge2_end + 1, total_leads); _t(ws2, tr2, merge2_end + 2, total_conv)
    ws2.freeze_panes = 'A3'

    # ── Feuille 3 : Synthèse par date ────────────────────────────────────────
    ws3 = wb.create_sheet('Synthèse par date')
    ws3.merge_cells('A1:C1')
    t3 = ws3['A1']
    t3.value = 'Synthèse par date — toutes sources confondues'
    t3.fill = TTL_BG; t3.font = _Font(bold=True, color='00D4FF', size=13); t3.alignment = LFT
    ws3.row_dimensions[1].height = 26
    ws3.merge_cells('A2:C2')
    s3 = ws3['A2']
    s3.value = 'Lead daté USER_LIFECYCLE_START_DATE · Conversion datée COMPANY_CREATED_AT (validated) / STEP_UPDATED_AT (awaiting_agent_confirmation)'
    s3.font = _Font(italic=True, color='888888', size=9); s3.alignment = LFT

    for c, h in enumerate(['Date','Leads','Conversions'], 1):
        _h(ws3, 3, c, h)

    for i, (_, r) in enumerate(by_date.iterrows()):
        row = 4 + i; fill = ALT_BG if i % 2 == 0 else None
        _c(ws3, row, 1, r['Date'], fill)
        _c(ws3, row, 2, int(r['Leads']),       LED_BG, bold=True)
        _c(ws3, row, 3, int(r['Conversions']), CNV_BG, bold=True)

    tr3 = 4 + len(by_date)
    _t(ws3, tr3, 1, 'TOTAL')
    _t(ws3, tr3, 2, total_leads); _t(ws3, tr3, 3, total_conv)
    ws3.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname = f"{filename_prefix}_{_now().strftime('%Y%m%d_%H%M')}.xlsx"
    resp = make_response(buf.read())
    resp.headers['Content-Type']        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@app.route('/export/affiliation')
def export_affiliation():
    """Génère un fichier Excel Leads & Conversions Affiliation (délègue à _build_medium_excel)."""
    return _build_medium_excel('affiliation', 'Affiliation', 'Affiliation_Leads_Conversions',
                               exclude_aff_ids=[360])  # 360 = NetAffiliation (tests)


@app.route('/export/partenariat')
def export_partenariat():
    """Génère un fichier Excel Leads & Conversions Partenariat (même format que Affiliation)."""
    return _build_medium_excel('partenariat', 'Partenariat', 'Partenariat_Leads_Conversions')


@app.route('/export/relances')
def export_relances():
    """Génère un fichier Excel : contribution journalière leads & conversions des campagnes email-abandon.

    Logique :
      - Repart du CSV brut pour trouver TOUTES les sessions email-abandon (non filtrées par dédup).
      - Dédup par (HASHED_USER_ID, AT_CAMPAIGN) : conserve la session la plus récente.
      - Croise avec _utm_dedup_df() pour obtenir la source préalable attribuée à chaque user
        (qui exclut justement email-abandon de l'attribution).
      - Dates leads/conversions lues depuis les colonnes lifecycle de la session email-abandon elle-même.

    Feuilles :
      1. Détail par jour × campagne × source préalable
      2. Synthèse par campagne × source préalable
      3. Synthèse par date (leads)
      4. Synthèse par date (conversions)
    """
    import pandas as _pd
    import numpy as _np
    import openpyxl as _xl
    from openpyxl.styles import Font as _Font
    import io
    from flask import make_response

    csv_path = UTM_CACHE_DIR / 'ALL_UTM.csv'
    if not csv_path.exists():
        return jsonify({'error': 'Fichier ALL_UTM.csv introuvable'}), 404

    # ── 1. Sessions email-abandon depuis le CSV brut ──────────────────────────
    df_raw = _pd.read_csv(csv_path, low_memory=False)
    df_raw['session_dt'] = _pd.to_datetime(df_raw['SESSION_CREATED_AT'], utc=True, errors='coerce')

    # Les relances sont AT_MEDIUM='emailing' avec AT_CAMPAIGN contenant 'email-abandon'
    abn = df_raw[df_raw['AT_CAMPAIGN'].fillna('').str.contains('email-abandon', case=False, na=False)].copy()
    if abn.empty:
        return jsonify({'error': 'Aucune campagne email-abandon trouvée dans ALL_UTM.csv'}), 404

    # ── 2. Dédup par (HASHED_USER_ID, AT_CAMPAIGN) → dernière session ────────
    abn_dedup = (abn.sort_values('session_dt')
                    .groupby(['HASHED_USER_ID', 'AT_CAMPAIGN'], group_keys=False)
                    .last()
                    .reset_index())

    # ── 3. Source préalable depuis _utm_dedup_df() ───────────────────────────
    #    email-abandon exclu de l'attribution → on obtient la vraie source préalable.
    #    Si la seule source dédupliquée EST une campagne email-abandon (user sans
    #    autre UTM) → Organique, campagne vide.
    df_prior = _utm_dedup_df()

    def _classify_prior(row):
        """Retourne (label_source, campagne_préalable) pour une ligne dédupliquée."""
        if row is None:
            return 'Organique', ''
        ch       = str(row.get('AT_CHANNEL',  '') or '').lower()
        medium   = str(row.get('AT_MEDIUM',   '') or '').lower()
        pid      = str(row.get('AF_PID',      '') or '')
        campaign = str(row.get('AT_CAMPAIGN', '') or '')
        utms     = str(row.get('UTMS_RAW',    '') or '')
        camp_low = campaign.lower()

        # ── Vérifier d'abord si c'est une session email-abandon (pas de vraie source)
        if 'email-abandon' in camp_low:
            return 'Organique', ''
        # ── Puis classer normalement
        if ch in ('pmax', 'pmaxmobile'):
            label = 'PMAX'
        elif ch == 'meta':
            label = 'META'
        elif ch == 'gg_search':
            label = 'SEA'
        elif medium == 'affiliation':
            label = 'Affiliation'
        elif medium == 'partenariat':
            label = 'Partenariat'
        elif ch == 'tiktok':
            label = 'TikTok'
        elif ch == 'linkedin':
            label = 'LinkedIn'
        elif 'amazon' in camp_low:
            label = 'Amazon'
        elif medium == 'emailing':
            label = 'Emailing'
        elif pid == 'eer_redirect_mobile_app':
            return 'Organique', ''
        elif not utms or utms in ('nan', 'NaT', 'None'):
            return 'Organique', ''
        else:
            return (pid if pid else 'Organique'), ''

        camp_out = campaign if campaign not in ('nan', 'NaT', 'None', '') else ''
        return label, camp_out

    prior_map      = {}
    prior_camp_map = {}
    if not df_prior.empty:
        for _, prow in df_prior.iterrows():
            uid = prow['HASHED_USER_ID']
            label, camp = _classify_prior(prow)
            prior_map[uid]      = label
            prior_camp_map[uid] = camp

    abn_dedup['Source préalable']   = abn_dedup['HASHED_USER_ID'].map(prior_map).fillna('Organique')
    abn_dedup['Campagne préalable'] = abn_dedup['HASHED_USER_ID'].map(prior_camp_map).fillna('')

    # ── 4. Dates ──────────────────────────────────────────────────────────────
    # Feuille détail : SESSION_CREATED_AT de la session email-abandon
    abn_dedup['session_date'] = abn_dedup['session_dt'].dt.strftime('%Y-%m-%d')

    # Conversion : COMPANY_CREATED_AT (validated) ou STEP_UPDATED_AT (awaiting)
    lc         = abn_dedup['USER_LIFECYCLE_STATUS'].fillna('')
    company_dt = _pd.to_datetime(abn_dedup['COMPANY_CREATED_AT'], utc=True, errors='coerce').dt.strftime('%Y-%m-%d')
    step_dt    = _pd.to_datetime(abn_dedup['STEP_UPDATED_AT'],    utc=True, errors='coerce').dt.strftime('%Y-%m-%d')
    abn_dedup['conv_date'] = _np.where(lc == 'validated', company_dt,
                             _np.where(lc == 'validated_awaiting_agent_confirmation', step_dt, _pd.NA))

    abn_dedup['_conv'] = lc.isin(['validated', 'validated_awaiting_agent_confirmation']).astype(int)

    # ── 5. Filtrer sur les conversions uniquement ────────────────────────
    abn_conv = abn_dedup[abn_dedup['_conv'] == 1].copy()

    def _clean(v):
        s = str(v)
        return '' if s in ('nan', 'NaT', 'None') else s

    # ── 6. Agrégations (conversions uniquement) ───────────────────────────────
    # Feuille 1 : par SESSION_CREATED_AT × campagne × source préalable × campagne préalable
    grp = (abn_conv
           .groupby(['session_date', 'AT_CAMPAIGN', 'Source préalable', 'Campagne préalable'])[['_conv']]
           .sum().reset_index())
    grp.columns = ['Date session relance', 'Campagne relance', 'Source préalable', 'Campagne préalable', 'Conversions']
    grp = grp.sort_values(['Date session relance', 'Campagne relance', 'Source préalable'])

    # Feuille 2 : par campagne × source préalable
    by_camp = (abn_conv
               .groupby(['AT_CAMPAIGN', 'Source préalable'])[['_conv']]
               .sum().reset_index())
    by_camp.columns = ['Campagne relance', 'Source préalable', 'Conversions']
    by_camp = by_camp.sort_values(['Campagne relance', 'Conversions'], ascending=[True, False])

    # Feuille 3 : par date de conversion (conv_date)
    by_date7 = abn_conv.groupby('conv_date')[['_conv']].sum().reset_index()
    by_date7.columns = ['Date', 'Conversions']
    by_date7 = by_date7.sort_values('Date')

    total_conv  = int(abn_conv['_conv'].sum())
    export_date = _now().strftime('%d/%m/%Y')

    # ── 7. Excel ──────────────────────────────────────────────────────────────
    st     = _excel_export_styles()
    _h = st['h']; _c = st['c']; _t = st['t']
    TTL_BG = st['TTL_BG']; TOT_BG = st['TOT_BG']
    CNV_BG = st['CNV_BG']; ALT_BG = st['ALT_BG']; THIN = st['THIN']
    LFT    = st['LFT']

    from openpyxl.styles import PatternFill as _Fill
    SRC_BG = _Fill('solid', fgColor='FFF3E0')   # orange très pâle pour source/campagne préalable

    wb = _xl.Workbook()

    # ── Feuille 1 : Détail par date de session relance ────────────────────────
    ws1 = wb.active; ws1.title = 'Détail par jour'
    ws1.merge_cells('A1:E1')
    t = ws1['A1']
    t.value = f'Relances email-abandon — Conversions · détail par session  (extrait {export_date})'
    t.fill = TTL_BG; t.font = _Font(bold=True, color='00D4FF', size=13); t.alignment = LFT
    ws1.row_dimensions[1].height = 26
    ws1.merge_cells('A2:E2')
    s = ws1['A2']
    s.value = ('Uniquement les conversions (validated / validated_awaiting_agent_confirmation) '
               '· Date = SESSION_CREATED_AT de la relance · Source préalable = attribution hors email-abandon')
    s.font = _Font(italic=True, color='888888', size=9); s.alignment = LFT

    for c, h in enumerate(['Date session relance', 'Campagne relance', 'Source préalable', 'Campagne préalable', 'Conversions'], 1):
        _h(ws1, 3, c, h)

    for i, (_, r) in enumerate(grp.iterrows()):
        row = 4 + i; fill = ALT_BG if i % 2 == 0 else None
        _c(ws1, row, 1, _clean(r['Date session relance']),  fill)
        _c(ws1, row, 2, _clean(r['Campagne relance']),       fill, align=LFT)
        _c(ws1, row, 3, _clean(r['Source préalable']),       SRC_BG, align=LFT)
        _c(ws1, row, 4, _clean(r['Campagne préalable']),     SRC_BG, align=LFT)
        _c(ws1, row, 5, int(r['Conversions']),               CNV_BG, bold=True)

    tr = 4 + len(grp)
    ws1.merge_cells(f'A{tr}:D{tr}')
    _t(ws1, tr, 1, 'TOTAL')
    for ci in (2, 3, 4): ws1.cell(tr, ci).fill = TOT_BG; ws1.cell(tr, ci).border = THIN
    _t(ws1, tr, 5, total_conv)

    for w, col in zip([22, 22, 20, 28, 13], 'ABCDE'):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = 'A4'

    # ── Feuille 2 : Synthèse par campagne × source préalable ─────────────────
    ws2 = wb.create_sheet('Synthèse par campagne')
    ws2.merge_cells('A1:C1')
    t2 = ws2['A1']
    t2.value = 'Synthèse par campagne de relance et source préalable — toutes dates confondues'
    t2.fill = TTL_BG; t2.font = _Font(bold=True, color='00D4FF', size=13); t2.alignment = LFT
    ws2.row_dimensions[1].height = 26

    for c, h in enumerate(['Campagne relance', 'Source préalable', 'Conversions'], 1):
        _h(ws2, 2, c, h)

    for i, (_, r) in enumerate(by_camp.iterrows()):
        row = 3 + i; fill = ALT_BG if i % 2 == 0 else None
        _c(ws2, row, 1, _clean(r['Campagne relance']),  fill, align=LFT)
        _c(ws2, row, 2, _clean(r['Source préalable']),  SRC_BG, align=LFT)
        _c(ws2, row, 3, int(r['Conversions']),          CNV_BG, bold=True)

    tr2 = 3 + len(by_camp)
    ws2.merge_cells(f'A{tr2}:B{tr2}')
    _t(ws2, tr2, 1, 'TOTAL'); ws2.cell(tr2, 2).fill = TOT_BG; ws2.cell(tr2, 2).border = THIN
    _t(ws2, tr2, 3, total_conv)
    for w, col in zip([22, 20, 13], 'ABC'):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A3'

    # ── Feuille 3 : Synthèse par date de conversion ───────────────────────────────────
    ws3 = wb.create_sheet('Synthèse par date')
    ws3.merge_cells('A1:B1')
    t3 = ws3['A1']
    t3.value = 'Conversions par date — COMPANY_CREATED_AT (validated) · STEP_UPDATED_AT (awaiting)'
    t3.fill = TTL_BG; t3.font = _Font(bold=True, color='00D4FF', size=13); t3.alignment = LFT
    ws3.row_dimensions[1].height = 26
    ws3.merge_cells('A2:B2')
    s3 = ws3['A2']
    s3.value = 'Seuls validated et validated_awaiting_agent_confirmation sont comptabilisés'
    s3.font = _Font(italic=True, color='888888', size=9); s3.alignment = LFT

    for c, h in enumerate(['Date', 'Conversions'], 1):
        _h(ws3, 3, c, h)

    for i, (_, r) in enumerate(by_date7.iterrows()):
        row = 4 + i; fill = ALT_BG if i % 2 == 0 else None
        _c(ws3, row, 1, _clean(r['Date']), fill)
        _c(ws3, row, 2, int(r['Conversions']),    CNV_BG, bold=True)

    tr3 = 4 + len(by_date7)
    _t(ws3, tr3, 1, 'TOTAL'); _t(ws3, tr3, 2, total_conv)
    for w, col in zip([18, 13], 'AB'):
        ws3.column_dimensions[col].width = w
    ws3.freeze_panes = 'A4'

    # ── Réponse HTTP ──────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    fname = f"Relances_Conversions_{_now().strftime('%Y%m%d_%H%M')}.xlsx"
    resp = make_response(buf.read())
    resp.headers['Content-Type']        = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


@app.route('/api/cohortes/update', methods=['POST'])
def api_cohortes_update():
    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé'}), 500
    try:
        n = _run_matrice_update()
        return jsonify({'ok': True, 'snapshots': n})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/cohortes')
def api_cohortes():
    if openpyxl is None:
        return jsonify({'error': 'openpyxl non installé — pip install openpyxl'}), 500
    matrice = COMP_DIR / 'Matrice_cohorte_conversion.xlsx'
    if not matrice.exists():
        return jsonify({'error': 'Fichier matrice introuvable'}), 404

    wb     = openpyxl.load_workbook(matrice)
    result = {}
    for sheet_name in wb.sheetnames:
        if 'Lundi' not in sheet_name:
            continue
        ws   = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(min_row=3, values_only=True):
            cells = [str(c).replace('\n', ' ') if c is not None else '' for c in row]
            if any(c for c in cells):
                rows.append(cells)
        if rows:
            result[sheet_name] = rows
    return jsonify(result)

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    import webbrowser, logging
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s')
    app.logger.setLevel(logging.INFO)
    port = 5002
    (BASE_DIR / 'templates').mkdir(exist_ok=True)
    _refresh_budget_excel()
    print(f"\n🚀  Mockup (données factices, {DEMO_NOW:%d/%m/%Y}) → http://localhost:{port}\n")
    threading.Timer(1.5, lambda: webbrowser.open(f'http://localhost:{port}')).start()
    app.run(debug=False, port=port, threaded=True)
