# Certificat de Depot — Rendement Numerique Annuel

Outil d'apprentissage interactif et code MATLAB pour evaluer un certificat de depot structure (CD numerique annuel) par simulation Monte Carlo.

**Cours** : MATH40602 — Methodes Quantitatives 2, HEC Montreal

## Produit

Un CD a capital protege lie a un panier de 10 titres (AAPL, C, F, HPQ, JNJ, LLY, LOW, MO, MRK, WMT). Chaque annee, le coupon depend de la performance individuelle de chaque titre :

| Rendement du titre | Performance attribuee |
|---|---|
| > 0% | Coupon numerique (6,50%) |
| entre -30% et 0% | Rendement reel |
| <= -30% | Plancher (-30%) |

Le taux du coupon est `max(0, moyenne des 10 performances)`. Le capital de 1 000 $ est rembourse a echeance.

## Demo web

Ouvrir `demo_produit_structure.html` dans un navigateur. La page contient :

- Description du produit et parametres du term sheet
- Matrice de correlation et decomposition de Cholesky interactive
- Simulation Monte Carlo configurable (N, taux, seed)
- Resultats : prix estime, histogramme des coupons, convergence, distribution des S_T
- Detail de la simulation n°1 (matrice L, vecteurs Z, prix terminaux par regime)
- Code MATLAB complet avec guide d'execution

## Fichiers

| Fichier | Description |
|---|---|
| `demo_produit_structure.html` | Application web interactive (standalone, pas de serveur) |
| `data.json` | Donnees de marche : prix, volatilites, correlation, Cholesky |
| `simuler_gbm.m` | Fonction MATLAB : simulation GBM avec Cholesky |
| `calculer_payoff.m` | Fonction MATLAB : regles du produit et actualisation |
| `evaluer_cd.m` | Script principal MATLAB : calibrage, simulation, resultats |
| `extract_data.py` | Extraction des donnees depuis le fichier Excel vers `data.json` |
| `term_sheet.md` | Term sheet du produit (parametres, formules, panier) |
| `create_pptx.js` | Generateur de la presentation PowerPoint |
| `presentation_devoir2.pptx` | Presentation generee (16 slides) |

## Execution MATLAB

```matlab
% Dans le meme dossier :
evaluer_cd

% Sortie attendue :
% === MATRICE L (FACTEUR DE CHOLESKY) ===
% ...
% === RESULTAT ===
% Prix estime du CD : 1036.49 $
% Coupon moyen      : 48.44 $
% IC 95%            : [1035.62, 1037.36]
% Simulations       : 10000
```

## Methode

1. **Donnees** : 564 rendements log hebdomadaires (janv. 2006 — oct. 2016), prix ajustes (dividendes inclus)
2. **Correlation** : matrice 10x10 calculee sur les rendements log, decomposition de Cholesky C = LL^T
3. **Simulation** : GBM risque-neutre, S_T = S_0 * exp((r - sigma^2/2)*T + sigma*sqrt(T)*Z), Z correles via L
4. **Evaluation** : application des regles a 3 regimes, actualisation a r = 1,25%, intervalle de confiance a 95%
