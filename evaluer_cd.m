% =========================================================================
% ÉVALUATION DU CD NUMÉRIQUE ANNUEL - DEVOIR 2 (MATH40602)
% Script principal : calibrage, simulation et résultats
% Date d'évaluation : 31 Octobre 2016
% =========================================================================
clear; clc; close all;

%% 1. PARAMÈTRES
N             = 10000;      % Nombre de simulations
T             = 11/12;      % Temps restant (≈ 0.9167 an)
r             = 0.0125;     % Taux sans risque (US treasury 1 an, oct. 2016)
Nominal       = 1000;       % Dépôt ($)
DigitalCoupon = 0.065;      % Coupon numérique (6.50%)
Floor         = -0.30;      % Plancher (-30%)

%% 2. DONNÉES DE MARCHÉ
% Ordre : [AAPL, C, F, HPQ, JNJ, LLY, LOW, MO, MRK, WMT]
Tickers = {'AAPL','C','F','HPQ','JNJ','LLY','LOW','MO','MRK','WMT'};

% Prix initiaux ajustés (Trade Date : 26 sept. 2011)
S_init = [13.62; 29.51; 10.00; 10.72; 75.54; 47.72; 20.95; 37.23; 39.76; 19.44];

% Prix au 31 octobre 2016
S0 = [31.05; 57.79; 14.56; 16.87; 159.39; 114.51; 79.61; 115.52; 85.37; 29.80];

% Volatilités annualisées
sigma = [0.3392; 0.6552; 0.5156; 0.3334; 0.1514; 0.2158; 0.3070; 0.1810; 0.2425; 0.1876];

% Matrice de corrélation (rendements log hebdomadaires, 564 obs.)
CorrMat = [
  1.0000 0.3429 0.3367 0.3678 0.2397 0.2385 0.3543 0.2097 0.2720 0.2156;
  0.3429 1.0000 0.5853 0.2927 0.3807 0.4169 0.4834 0.2329 0.3883 0.2176;
  0.3367 0.5853 1.0000 0.3608 0.3791 0.4354 0.5305 0.3210 0.3008 0.3214;
  0.3678 0.2927 0.3608 1.0000 0.3737 0.2829 0.3733 0.1961 0.2919 0.2874;
  0.2397 0.3807 0.3791 0.3737 1.0000 0.5803 0.3846 0.4625 0.5404 0.4415;
  0.2385 0.4169 0.4354 0.2829 0.5803 1.0000 0.4118 0.4072 0.5361 0.4048;
  0.3543 0.4834 0.5305 0.3733 0.3846 0.4118 1.0000 0.3158 0.3358 0.4714;
  0.2097 0.2329 0.3210 0.1961 0.4625 0.4072 0.3158 1.0000 0.3931 0.3065;
  0.2720 0.3883 0.3008 0.2919 0.5404 0.5361 0.3358 0.3931 1.0000 0.3800;
  0.2156 0.2176 0.3214 0.2874 0.4415 0.4048 0.4714 0.3065 0.3800 1.0000];

%% 3. SIMULATION
% IMPORTANT :
% Cette version suppose que la fonction simuler_gbm a été modifiée pour retourner
% [S_final, L, Z_indep, Z_corr]
[S_final, L, Z_indep, Z_corr] = simuler_gbm(S0, sigma, CorrMat, r, T, N);

%% 4. TABLEAUX : MATRICE DE CHOLESKY ET Z
% --- Tableau de la matrice L ---
Table_L = array2table(L, 'VariableNames', Tickers, 'RowNames', Tickers);

% --- On affiche seulement la 1re simulation pour les Z ---
Table_Z_indep_1 = table(Z_indep(:,1), 'RowNames', Tickers, ...
    'VariableNames', {'Z_indep_sim1'});

Table_Z_corr_1 = table(Z_corr(:,1), 'RowNames', Tickers, ...
    'VariableNames', {'Z_corr_sim1'});

% --- Affichage dans la Command Window ---
disp(' ');
disp('=== MATRICE L (FACTEUR DE CHOLESKY) ===');
disp(Table_L);

disp(' ');
disp('=== Z INDÉPENDANT - SIMULATION 1 ===');
disp(Table_Z_indep_1);

disp(' ');
disp('=== Z CORRÉLÉ - SIMULATION 1 ===');
disp(Table_Z_corr_1);

%% 5. FENÊTRES GRAPHIQUES DE TABLEAUX
% --- Fenêtre pour la matrice L ---
f1 = figure('Name', 'Matrice L de Cholesky', 'Position', [100 100 1000 320]);
uitable(f1, ...
    'Data', round(L, 3), ...
    'ColumnName', Tickers, ...
    'RowName', Tickers, ...
    'Units', 'Normalized', ...
    'Position', [0 0 1 1]);

% --- Fenêtre pour Z indépendant et Z corrélé (simulation 1) ---
Data_Z = [Z_indep(:,1), Z_corr(:,1)];

f2 = figure('Name', 'Z indépendants et Z corrélés - Simulation 1', ...
    'Position', [150 150 500 320]);
uitable(f2, ...
    'Data', round(Data_Z, 3), ...
    'ColumnName', {'Z_indep', 'Z_corr'}, ...
    'RowName', Tickers, ...
    'Units', 'Normalized', ...
    'Position', [0 0 1 1]);

%% 6. ÉVALUATION DU PRODUIT
[PrixCD, CouponMoyen, IC95, Payoffs] = calculer_payoff( ...
    S_final, S_init, DigitalCoupon, Floor, Nominal, r, T);

%% 7. RÉSULTATS
fprintf('\n=== RÉSULTAT ===\n');
fprintf('Prix estimé du CD : %.2f $\n', PrixCD);
fprintf('Coupon moyen      : %.2f $\n', CouponMoyen);
fprintf('IC 95%%            : [%.2f, %.2f]\n', PrixCD - IC95, PrixCD + IC95);
fprintf('Simulations       : %d\n', N);

%% 8. EXPORT EXCEL (OPTIONNEL)
% Décommente ces lignes si tu veux sauvegarder les tableaux dans Excel

% writetable(Table_L, 'matrice_L.xlsx', 'WriteRowNames', true);
% writetable(Table_Z_indep_1, 'z_indep_sim1.xlsx', 'WriteRowNames', true);
% writetable(Table_Z_corr_1, 'z_corr_sim1.xlsx', 'WriteRowNames', true);

%% 9. GRAPHIQUES
CouponPaiements = Payoffs - Nominal;

% --- Histogramme des coupons ---
figure;
histogram(CouponPaiements, 20, 'FaceColor', [0 0.52 0.56]);
xlabel('Paiement du coupon ($)');
ylabel('Fréquence');
title('Distribution des paiements de coupon');
grid on;

% --- Convergence Monte Carlo ---
figure;
running_mean = cumsum(Payoffs) ./ (1:N);
running_std  = zeros(1, N);

for k = 1:N
    running_std(k) = std(Payoffs(1:k));
end

running_pv = exp(-r*T) * running_mean;
ci_up  = running_pv + 1.96 * exp(-r*T) * running_std ./ sqrt(1:N);
ci_low = running_pv - 1.96 * exp(-r*T) * running_std ./ sqrt(1:N);

plot(1:N, running_pv, 'k-', 'LineWidth', 1.5);
hold on;
fill([1:N, fliplr(1:N)], [ci_up, fliplr(ci_low)], ...
    [0.7 0.85 0.95], 'EdgeColor', 'none', 'FaceAlpha', 0.4);

xlabel('Nombre de simulations');
ylabel('Prix estimé ($)');
title('Convergence de l''estimateur Monte Carlo');
legend('Prix moyen', 'IC 95%', 'Location', 'best');
grid on;

%% 10. COMPARAISON : CD NUMÉRIQUE VS PORTEFEUILLE DIRECT
% Même investissement de 1000 $, réparti également entre les 10 titres
K = length(S0);
% S_final./S0 : rendement relatif [K x N], sum → [1 x N], * 100 → valeur par titre
PortfolioValue = sum(S_final ./ S0) * (Nominal / K);   % [1 x N]
CDValue = Payoffs;                                       % [1 x N] = Nominal + coupon

% Statistiques
fprintf('\n=== COMPARAISON CD vs PORTEFEUILLE DIRECT ===\n');
fprintf('CD      : E[valeur] = %.2f $, σ = %.2f $, P(perte) = %.1f%%\n', ...
    mean(CDValue), std(CDValue), 100*mean(CDValue < Nominal));
fprintf('Actions : E[valeur] = %.2f $, σ = %.2f $, P(perte) = %.1f%%\n', ...
    mean(PortfolioValue), std(PortfolioValue), 100*mean(PortfolioValue < Nominal));

% VaR 5%
CDSorted = sort(CDValue);
StSorted = sort(PortfolioValue);
idx5 = floor(0.05 * N);
fprintf('CD      : VaR 5%% = %.2f $\n', CDSorted(idx5));
fprintf('Actions : VaR 5%% = %.2f $\n', StSorted(idx5));

% --- Histogramme overlay ---
figure;
edges = linspace(min([CDValue, PortfolioValue]), ...
                 max([CDValue, PortfolioValue]), 50);
histogram(CDValue, edges, 'FaceColor', [0 0.52 0.56], ...
    'FaceAlpha', 0.6, 'DisplayName', 'CD numérique');
hold on;
histogram(PortfolioValue, edges, 'FaceColor', [0.18 0.20 0.21], ...
    'FaceAlpha', 0.4, 'DisplayName', 'Portefeuille direct');
xline(Nominal, '--', 'Seuil', 'Color', [0.75 0.34 0], 'LineWidth', 1.5);
xlabel('Valeur finale ($)');
ylabel('Fréquence');
title('CD numérique vs portefeuille direct (1 000 $)');
legend('Location', 'best');
grid on;

% --- CDF (fonction de répartition) ---
figure;
CDSorted = sort(CDValue);
StSorted = sort(PortfolioValue);
pct = (1:N) / N;
plot(CDSorted, pct, '-', 'Color', [0 0.52 0.56], ...
    'LineWidth', 2, 'DisplayName', 'CD numérique');
hold on;
plot(StSorted, pct, '--', 'Color', [0.18 0.20 0.21], ...
    'LineWidth', 2, 'DisplayName', 'Portefeuille direct');
xline(Nominal, ':', 'Color', [0.75 0.34 0]);
yline(0.05, ':', 'VaR 5%', 'Color', [0.72 0.11 0.11]);
xlabel('Valeur finale ($)');
ylabel('Probabilité cumulative');
title('Fonction de répartition : CD vs Actions');
legend('Location', 'southeast');
grid on;