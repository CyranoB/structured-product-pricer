function [PrixCD, CouponMoyen, IC95, Payoffs] = calculer_payoff(S_final, S_init, ...
    DigitalCoupon, Floor, Nominal, r, T)
% CALCULER_PAYOFF  Applique les règles du CD numérique annuel aux prix simulés
%
%   [PrixCD, CouponMoyen, IC95, Payoffs] = calculer_payoff(S_final, S_init, ...
%       DigitalCoupon, Floor, Nominal, r, T)
%
%   Entrées :
%     S_final       - matrice [K x N] des prix terminaux simulés
%     S_init        - vecteur [K x 1] des prix initiaux (Trade Date, 2011)
%     DigitalCoupon - coupon numérique (ex: 0.065 pour 6.50%)
%     Floor         - plancher par titre (ex: -0.30 pour -30%)
%     Nominal       - dépôt initial (ex: 1000)
%     r             - taux sans risque
%     T             - horizon en années
%
%   Sorties :
%     PrixCD     - prix actualisé du CD ($)
%     CouponMoyen- coupon moyen ($)
%     IC95       - demi-largeur de l'intervalle de confiance à 95% ($)
%     Payoffs    - vecteur [1 x N] des paiements finaux bruts

    [K, N] = size(S_final);

    % --- Rendement par rapport au prix initial de 2011 ---
    StockReturns = (S_final - S_init) ./ S_init;   % [K x N]

    % --- Application des règles du produit ---
    Perf = zeros(K, N);

    % (1) Rendement > 0% => Performance = Digital Coupon (6.50%)
    Perf(StockReturns > 0) = DigitalCoupon;

    % (2) -30% < Rendement <= 0% => Performance = Rendement
    idx_mid = (StockReturns <= 0) & (StockReturns > Floor);
    Perf(idx_mid) = StockReturns(idx_mid);

    % (3) Rendement <= -30% => Performance = Floor (-30%)
    Perf(StockReturns <= Floor) = Floor;

    % --- Taux du coupon (moyenne pondérée, plancher à 0%) ---
    % Poids = 1/K pour chaque titre (pondération équiprobable)
    CouponRate = max(0, mean(Perf, 1));             % [1 x N]

    % --- Paiement final = Principal + Coupon ---
    Payoffs = Nominal + Nominal * CouponRate;        % [1 x N]

    % --- Actualisation et statistiques ---
    PrixCD     = exp(-r * T) * mean(Payoffs);
    CouponMoyen = mean(Payoffs) - Nominal;
    StdPayoffs = std(Payoffs);
    IC95       = 1.96 * StdPayoffs / sqrt(N) * exp(-r * T);

end