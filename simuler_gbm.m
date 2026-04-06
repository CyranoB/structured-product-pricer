function [S_final, L, Z_indep, Z_corr] = simuler_gbm(S0, sigma, CorrMat, r, T, N)
% SIMULER_GBM  Simule N prix terminaux pour un panier d'actions corrélées
%   via le mouvement brownien géométrique sous la mesure risque-neutre.
%
%   S_final = simuler_gbm(S0, sigma, CorrMat, r, T, N)
%
%   Entrées :
%     S0      - vecteur [K x 1] des prix actuels (prix ajustés, date d'évaluation)
%     sigma   - vecteur [K x 1] des volatilités annualisées (calculées sur prix ajustés)
%     CorrMat - matrice [K x K] de corrélation des rendements
%     r       - taux sans risque (scalaire)
%     T       - horizon de simulation en années (scalaire)
%     N       - nombre de simulations Monte Carlo (scalaire)
%
%   Sortie :
%     S_final - matrice [K x N] des prix terminaux simulés
%
%   Note : Les prix ajustés intègrent déjà les dividendes, donc la
%   dérive risque-neutre est (r - sigma^2/2), sans terme q.
%
%   Formule GBM (un seul pas, mesure risque-neutre) :
%     S_T = S0 * exp((r - sigma^2/2)*T + sigma*sqrt(T)*Z)
%   où Z = L * X, avec L = chol(CorrMat) et X ~ N(0,I)

    K = length(S0);

    % --- Décomposition de Cholesky ---
    % CorrMat = L * L', L triangulaire inférieure
    % Transforme des variables indépendantes en variables corrélées
    L = chol(CorrMat, 'lower');

    % --- Dérive risque-neutre ---
    % Pas de q car les prix ajustés intègrent déjà les dividendes
    drift = (r - 0.5 * sigma.^2) * T;             % [K x 1]

    % --- Génération des chocs corrélés ---
    Z_indep = randn(K, N);                     % [K x N] normales indépendantes
    Z_corr  = L * Z_indep;                     % [K x N] normales corrélées

    % --- Prix terminaux ---
    choc    = sigma * sqrt(T) .* Z_corr;       % [K x N]
    S_final = S0 .* exp(drift + choc);         % [K x N]

end

%[appendix]{"version":"1.0"}
%---
