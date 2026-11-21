# Tom Bussard

**Quantitative Analyst Apprentice** at BNP Paribas CIB<br>
**Master's in Economics and Financial Engineering** (quantitative track) at Université Paris Dauphine PSL

Before joining BNP Paribas CIB, I worked in structuring at Amundi, following an internship at Société Générale CIB. I started my studies in economics and mathematics at Toulouse School of Economics.

I also build software outside finance. I developed **MyExperts** so people who refer renovation work to Expert Travaux can submit leads, track jobs and check their commissions. With **KeyDom**, I built a platform that helps students find medium-term rentals in Paris and its surrounding region.

[2026](#2026) · [2025](#2025) · [2024](#2024) · [2023](#2023) · [LinkedIn](https://www.linkedin.com/in/tom-bussard/)

## 2026

### [LatticeLab](https://github.com/TomBussard/LatticeLab)

**An object-oriented equity derivatives workbench in Python and Excel VBA.** I built the binomial and trinomial pricer around separate market, contract, node and tree objects. A contract supplies its payoff, while the tree builds the possible price paths and works backwards through them. That split let me add cash digitals, an up one-touch and custom expressions to the same engine that handles European and American vanilla options. A multi-leg position can combine different strikes, expiries and signed quantities.

One question I wanted to answer was how much a barrier crossing between two tree dates is worth. With the share at €100 and a €10 payment for touching €110, the 64-step tree gives **€6.3410** when I observe only the grid dates and **€6.7470** with Brownian-bridge monitoring. A digital that checks the same level only at expiry is worth **€3.9086**. I reproduced the two touch prices in desktop Excel.

![LatticeLab comparing a terminal digital with two ways of monitoring the same cash barrier](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/barrier-monitoring.png)

![LatticeLab showing a four-step American put with the share price and option value at each node](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/american-put-tree.jpg)

In this example the share starts at **€100**. After one quarter it can rise to **€111.93** or fall to **€91.60**. Two down moves bring the share to **€83.91**. Exercising the €100 put there pays **€16.09**, compared with **€15.39** for holding it one more step. The tree marks that decision with a dark outline.

![The Python pricer showing how a European call's delta changes as the share price moves from 80 to 120 euros](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/call-delta.jpg)

I then switched to a European call to study the hedge. At a share price of **€100**, the 128-step model gives a delta of **0.6333**. Near that point a €1 move in the share changes the price of one call by about **€0.63**. The curve shows how quickly that hedge changes as the share moves through the strike.

![LatticeLab displaying the European call price and its absolute error from Black-Scholes at six grid sizes](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/binomial-convergence.jpg)

The same call is worth **€10.4506** under Black-Scholes. The 96-step tree came closer than the 128-step tree. Checking several grids helped me distinguish a lucky match from a reliable price. I also ran the same contracts through desktop Excel to check the Python implementation.

The risk view covers **delta, gamma, vega, theta, rho, vanna, volga, charm and speed** for the signed position. I can change the spot to inspect each sensitivity, check the effect of pruning a trinomial tree, or time construction and backward induction separately. A saved two-call spread gives a second check on the portfolio layer and on the custom-payoff editor.

[Explore the workbench](https://github.com/TomBussard/LatticeLab#pricing-workbench) · [See the class design](https://github.com/TomBussard/LatticeLab#object-oriented-design-of-the-lattice) · [Reproduce the digital and one-touch comparison](https://github.com/TomBussard/LatticeLab#digital-cash-or-cash-at-first-touch)

### [VolDesk](https://github.com/TomBussard/VolDesk)

**A BTC and ETH options desk in C# and .NET.** I modelled a strategy as a collection of signed option legs. Listed and custom legs share a pricing interface but obtain their market inputs differently. The same domain objects take a trade from the option chain or an off-grid client quote through hedging and booking. A separate market-data layer lets me run the desk on public quotes, a synthetic market or a replay I can pause and step through.

![VolDesk market overview with the BTC volatility smile and term structure](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/market-overview.jpg)

The overview is running on synthetic quotes. Each point on the smile comes from an option mark. The table on the right puts that expiry in context by showing the forward and at-the-money volatility at the other maturities.

![VolDesk option chain with an expiry selector, implied-volatility curve and strike-level quotes](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/option-chain.jpg)

From the chain I can combine listed calls into a spread and compare its model value with the price available across both legs. In a controlled one-month example with strikes of USD 100,000 and 110,000, the model premium was **USD 3,584** and the net delta was **0.214 BTC**. That delta gives an initial hedge size. It changes as the underlying moves.

![A two-leg VolDesk quote with prices per leg, model premium and executable bid and ask](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/strategy-quote.jpg)

The pricer above uses the strikes and expiry from the captured session. Its two-leg quote can be checked against the bid and ask before adding the position to the simulated book.

For a custom trade, the desk interpolates the forward and total variance, then uses two listed options to hedge gamma and vega and a future to offset delta. I reran a client purchase of **25 off-grid BTC calls**. Before hedging, the desk's vega was **−USD 3,135.43 per volatility point**. After rounding the three hedge instruments to tradable quantities, the residual was **+USD 9.06 per point**.

![How the two listed calls and the future offset the Greeks of a custom VolDesk trade](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/otc-hedge.png)

The listed pricer also handles put spreads, butterflies, straddles, strangles and risk reversals. I can compare today's model value with expiry P&L in USD or crypto, read break-even prices, review realised and open P&L, and export the blotter, positions and expiry risk to Excel, CSV or JSON.

[Explore the desk](https://github.com/TomBussard/VolDesk#from-the-option-chain-to-the-book) · [See how the C# objects fit together](https://github.com/TomBussard/VolDesk#object-oriented-design-of-the-desk) · [Follow the off-grid quote and hedge](https://github.com/TomBussard/VolDesk#hedging-an-off-grid-call)

### [TransitionLens](https://github.com/TomBussard/TransitionLens)

**Carbon costs and systematic equity selection in Python.** I rank companies against their own industries, measure how a carbon price changes their relative cost base and take that signal all the way through stock selection, weights, turnover and net returns. The research modules also calculate value, quality, clean-technology and momentum scores from dated inputs.

![Gross and net portfolio paths alongside starting position sizes in the synthetic carbon-price experiment](https://raw.githubusercontent.com/TomBussard/TransitionLens/main/figures/portfolio-path.png)

At **EUR 65 per tonne**, the example selected **24 of 80** fictional companies. The largest opening position was 5%. Over 252 simulated trading days, gross return was **8.23%** and net return was **8.03%**. The gap comes from charging for turnover when the portfolio is formed and rebalanced.

I can switch between equal and capped weights, let positions drift between rebalances, and compare quarterly, annual and buy-and-hold policies. In the carbon example, raising the charge from **15 to 50 basis points per unit traded** lowers the net result from **8.03% to 7.58%** on the same synthetic returns. The project also shows when a small universe can satisfy the 5/10/40 concentration rules.

[Explore the research toolkit](https://github.com/TomBussard/TransitionLens#research-toolkit) · [Compare factor scores, caps and costs](https://github.com/TomBussard/TransitionLens#building-the-other-research-signals)

### [RatingSpread](https://github.com/TomBussard/RatingSpread)

**Comparing four ESG ratings at company, provider and portfolio level.** I built a Python comparison that aligns score direction and scale, computes all six provider correlations, ranks individual companies by disagreement and summarises scores by sector and region. I then take the top fifth from each method to see which holdings the four providers have in common.

![Pairwise score correlations for the reproducible four-method rating example](https://raw.githubusercontent.com/TomBussard/RatingSpread/main/figures/rating-agreement.png)

In the historical sample, **4 companies out of 103** selected by at least one provider appeared in all four lists. The mean pairwise correlation was **0.2824**. The figure uses a separate, reproducible sample of fictional companies so readers can inspect the ranking rules and rerun the comparison.

The public example lets me change the factor weights and see where agreement breaks down. In that run, **36 of the 72 selected companies** made exactly one provider's list and just **six** made all four. The project shows both the historical aggregate results and the company-level calculations on the fictional sample.

[Explore the four-provider comparison](https://github.com/TomBussard/RatingSpread#comparing-four-providers) · [See how many providers selected each firm](https://github.com/TomBussard/RatingSpread#account-for-every-selection)

## 2025

### [StratLedger](https://github.com/TomBussard/StratLedger)

**A portfolio simulator backed by a SQLite transaction ledger.** I can follow a sample client from an initial allocation through a sequence of strategy decisions, then ask which holdings were actually in the book on any given Monday. The dashboard brings together equities, bonds, ETFs, market-price comparisons, trade history and portfolio composition.

![Six weekly observations showing a position growing from ten to fifteen shares and falling to eleven](https://raw.githubusercontent.com/TomBussard/StratLedger/main/figures/ledger-example.png)

In the six-week example, buying 10 shares, adding 5 and selling 4 leaves **11 shares worth 1,122** at a final price of 102. The step chart makes it easy to check which changes came from trading. The dashboard then adds strategy rules and weekly risk measures.

The simulator includes a volatility-based low-risk rule, a lower-turnover rule with two eligible trading dates per month and a momentum-based equity rule. I replayed the two equity rules on five French stocks at the final 2023 close. Both picked L'Oréal and Schneider Electric. Only the higher-risk rule admitted Sanofi, whose **12.56%** volatility exceeded the low-risk entry threshold. The six-week walkthrough separately traces booked positions into weekly volatility, Sharpe and drawdown.

[Explore the client and strategy workflow](https://github.com/TomBussard/StratLedger#client-and-strategy-workflow) · [Trace weekly performance to its inputs](https://github.com/TomBussard/StratLedger#from-five-weekly-observations-to-the-dashboard-metrics)

### [GreenFolio](https://github.com/TomBussard/GreenFolio)

**A Streamlit builder for sustainability preferences and portfolio risk.** I can choose a risk profile, compare four ESG screens on the same universe, inspect company scores and set the weights on the surviving equities. The exposure view shows how those choices change environmental, social and governance measures and sector concentration.

![How many of eight fictional companies pass each of four sustainability screens](https://raw.githubusercontent.com/TomBussard/GreenFolio/main/figures/screen-comparison.png)

Of the eight example companies, **2** passed the environmental screen, **4** the social screen, **6** the balanced screen and **7** the general one. The repository lists the cutoffs so each inclusion can be checked against its four scores.

At equal weights the environmental selection has an average social risk score of **7.50**. Switching to the social selection takes it to **5.25** and raises the average environmental score from **2.50 to 5.00**. With price data, I can choose a benchmark and date range, plot a base-100 history and inspect return, volatility, drawdown, Sharpe, alpha and beta.

[Explore the portfolio builder](https://github.com/TomBussard/GreenFolio#portfolio-builder) · [See how the risk dashboard calculates its metrics](https://github.com/TomBussard/GreenFolio#what-the-risk-dashboard-calculates)

### [AlpineSwap](https://github.com/TomBussard/AlpineSwap)

**An Excel and Python model of the currency risk in a municipal interest-rate swap.** I can change the initial FX rate, volatility and loan assumptions, inspect each annual fixing and trace the barrier coupon through amortisation and discounting. Fixed shocks make a single contract path auditable. Monte Carlo shows how wide a range of borrower outcomes the same terms can produce.

![Distribution of the discounted borrower cash flows across ten thousand simulated EUR/CHF paths](https://raw.githubusercontent.com/TomBussard/AlpineSwap/main/figures/swap-distribution.png)

The mean discounted net cash flow was **EUR -35.67 million**. The fifth percentile was **EUR -65.14 million**. A single adverse fixing also changes later payments because the FX path carries the shock forward.

I can also isolate a single fixing, compare it with the zero-shock path, or take two payments from each simulated path. In the 10,000-path run, the 2014 and 2020 discounted payments had a **correlation of 0.699**. The VBA routine records both payments and the whole swap on the same row so the dependence survives the export.

[Explore the scenario tools](https://github.com/TomBussard/AlpineSwap#scenario-tools) · [Inspect amortisation and the distribution of losses](https://github.com/TomBussard/AlpineSwap#amortisation-and-the-distribution-of-losses)

### [IndexFoundry](https://github.com/TomBussard/IndexFoundry)

**Building an equity index from its constituents and rules.** The offline experiment isolates the effect of opening weights. The workbook view adds US and European price histories, price-coverage checks, FX conversion, sector and country filters, and comparison with the S&P 500 or STOXX Europe 600.

![Two base-100 index paths built from identical prices with different starting allocations](https://raw.githubusercontent.com/TomBussard/IndexFoundry/main/figures/weighting-rules.png)

With the same three fictional companies, the equal-weight index finished at **108.00**. A **60/25/15** starting allocation finished at **112.60**. In the offline dashboard, moving one slider changes the weights while keeping every component price fixed.

I also ran the full workbook on US and European equity histories. Across 2,168 dates, the Health Care basket grew from **100 to 315.47** and a ten-stock financial-strength basket finished at **259.08**. These are retrospective comparisons using the workbook's later fundamentals snapshot. I can change the fundamental weights, momentum lookback and selection size, link new baskets at rebalancing dates and export the results.

![Opening and year-end levels from two historical equity baskets built with IndexFoundry](https://raw.githubusercontent.com/TomBussard/IndexFoundry/main/figures/workbook-baskets.png)

[Explore the index workbench](https://github.com/TomBussard/IndexFoundry#index-workbench) · [See the historical workbook results](https://github.com/TomBussard/IndexFoundry#results-from-the-historical-workbook)

## 2024

### [TetherScope](https://github.com/TomBussard/TetherScope)

**An R time-series study of crypto volatility and Tether's market share.** I aligned 115 consecutive weekly observations, compared five transformations using ADF and KPSS diagnostics, and fitted both a univariate MA(1) forecast and a three-lag vector autoregression. The notebook checks residuals, tests predictive timing in each direction and traces a volatility shock through the fitted system.

![Weekly crypto volatility and Tether market share over the historical study window](https://raw.githubusercontent.com/TomBussard/TetherScope/main/figures/weekly-series.png)

In the three-lag model, past volatility changes helped predict changes in market-share growth (**p = 0.0157**). The reverse test gave **p = 0.5001**. This is evidence about timing within the sample. The report also shows the forecast intervals and how the estimated response changes after a shock.

I also reran the 12-week impulse response with **499 bootstrap draws**. One week after the shock the estimated change in log-share growth is **−0.02365**, with a 95% interval below zero. By the second week the interval spans zero. The project shows the response curve, the weekly forecast intervals and the effect of ordering volatility first.

[Explore the R study](https://github.com/TomBussard/TetherScope#inside-the-r-study) · [See the forecast and impulse response](https://github.com/TomBussard/TetherScope#forecast-errors-and-the-response-to-a-shock)

### [ParisTape](https://github.com/TomBussard/ParisTape)

**Five French equities and 10,000 ways to hold them.** I compared adjusted-price and daily-return correlations, examined return histograms and sampled long-only combinations of Air Liquide, L'Oréal, Schneider Electric, LVMH and Sanofi. The notebook maps each point on the risk-return chart back to the weights that produced it.

![Historical annualised volatility and return for ten thousand sampled five-stock allocations](https://raw.githubusercontent.com/TomBussard/ParisTape/main/figures/portfolio-cloud.png)

Across **1,283 common trading dates**, the lowest annualised volatility among the sampled portfolios was **17.64%**. The highest sampled Sharpe ratio was **1.08** at a zero risk-free rate. The workbook export preserves the weights behind each point.

The Excel export keeps all draws plus the four extremes in volatility and Sharpe. I worked back through the covariance matrix for the highest-Sharpe draw. Schneider Electric accounts for **37.64% of the capital** and **45.11% of estimated variance** in that portfolio. Its price-level correlation with Air Liquide was 0.936, compared with 0.658 for daily returns.

[Explore the notebook and its outputs](https://github.com/TomBussard/ParisTape#what-the-notebook-and-script-produce) · [See how the holdings contribute to risk](https://github.com/TomBussard/ParisTape#what-drives-the-risk-of-a-sampled-portfolio)

## 2023

### [EcoAtlas](https://github.com/TomBussard/EcoAtlas)

**A cross-country statistics project in R.** I assembled environmental performance, GDP per person, net plastic waste trade and political regime for 58 countries. The notebook lets me inspect the individual countries on a scatter plot, compare confidence levels for the sample mean and use Welch's test for groups of different sizes and variances.

![Each country’s GDP per person against its environmental performance score, with the sample mean marked](https://raw.githubusercontent.com/TomBussard/EcoAtlas/main/figures/income-and-environment.png)

Mean environmental performance was **52.74**, with a 95% confidence interval of **49.33 to 56.15**. In a separate continent-by-regime table, 22 of 24 expected counts were below five. I used 9,999 simulated tables with the same margins to assess that comparison.

The study also asks where the continent-by-regime association comes from. A heatmap reports observed and expected counts for each of the 24 cells. **Asia / authoritarian regime** and **South America / hybrid regime** together account for about **36%** of the chi-squared statistic. The report includes the underlying counts, the simulated test and the workbook used to reproduce both.

[Explore the country sample](https://github.com/TomBussard/EcoAtlas#explore-the-country-sample) · [Read the contingency table and its heatmap](https://github.com/TomBussard/EcoAtlas#where-the-continent-association-comes-from)
