# Tom Bussard

**Quantitative Analyst Apprentice** at BNP Paribas CIB<br>
**Master's in Economics and Financial Engineering** (quantitative track) at Université Paris Dauphine PSL

Before joining BNP Paribas CIB, I worked in structuring at Amundi, following an internship at Société Générale CIB. I started my studies in economics and mathematics at Toulouse School of Economics.

I also build software outside finance. I developed **MyExperts** so people who refer renovation work to Expert Travaux can submit leads, track jobs and check their commissions. With **KeyDom**, I built a platform that helps students find medium-term rentals in Paris and its surrounding region.

[2026](#2026) · [2025](#2025) · [2024](#2024) · [2023](#2023) · [LinkedIn](https://www.linkedin.com/in/tom-bussard/)

## 2026

### [LatticeLab](https://github.com/TomBussard/LatticeLab)

**A binomial option pricer in Python, with a second implementation in Excel VBA.** I built the tree so I could follow an option price all the way back from expiry. An American put was a good first test. At every node I have to decide whether the holder is better off exercising or waiting.

![LatticeLab showing a four-step American put with the share price and option value at each node](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/american-put-tree.jpg)

In this example the share starts at **€100**. After one quarter it can rise to **€111.93** or fall to **€91.60**. Two down moves bring the share to **€83.91**. Exercising the €100 put there pays **€16.09**, compared with **€15.39** for holding it one more step. The tree marks that decision with a dark outline.

![The Python pricer showing how a European call's delta changes as the share price moves from 80 to 120 euros](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/call-delta.jpg)

I then switched to a European call to study the hedge. At a share price of **€100**, the 128-step model gives a delta of **0.6333**. Near that point a €1 move in the share changes the price of one call by about **€0.63**. The curve shows how quickly that hedge changes as the share moves through the strike.

![LatticeLab displaying the European call price and its absolute error from Black-Scholes at six grid sizes](https://raw.githubusercontent.com/TomBussard/LatticeLab/main/figures/binomial-convergence.jpg)

The same call is worth **€10.4506** under Black-Scholes. The 96-step tree came closer than the 128-step tree. Checking several grids helped me distinguish a lucky match from a reliable price. I also ran the same contracts through desktop Excel to check the Python implementation.

[Explore the two contracts, the node calculations and the Excel workbook](https://github.com/TomBussard/LatticeLab#inside-the-pricer)

### [VolDesk](https://github.com/TomBussard/VolDesk)

**A BTC and ETH options desk in C# and .NET.** I wanted to see what happens between reading an option chain and carrying the resulting risk. I connected market data, Black-76 pricing and a simulated trading book so that the same contract can be followed across all three.

![VolDesk market overview with the BTC volatility smile and term structure](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/market-overview.jpg)

The overview is running on synthetic quotes. Each point on the smile comes from an option mark. The table on the right puts that expiry in context by showing the forward and at-the-money volatility at the other maturities.

![VolDesk option chain with an expiry selector, implied-volatility curve and strike-level quotes](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/option-chain.jpg)

From the chain I can combine listed calls into a spread and compare its model value with the price available across both legs. In a controlled one-month example with strikes of USD 100,000 and 110,000, the model premium was **USD 3,584** and the net delta was **0.214 BTC**. That delta gives an initial hedge size. It changes as the underlying moves.

![A two-leg VolDesk quote with prices per leg, model premium and executable bid and ask](https://raw.githubusercontent.com/TomBussard/VolDesk/main/figures/strategy-quote.jpg)

The pricer above uses the strikes and expiry from the captured session. Its two-leg quote can be checked against the bid and ask before adding the position to the simulated book.

[Walk through a quote, its Greeks and the simulated hedge](https://github.com/TomBussard/VolDesk#walk-through-a-trade)

### [TransitionLens](https://github.com/TomBussard/TransitionLens)

**Carbon costs and equity portfolio construction in Python.** I asked how a hypothetical carbon price would affect companies differently within the same industry. I compared Scope 1 and 2 emissions intensity with each industry's median, converted the gap into an annual cost advantage and scaled it by enterprise value before ranking stocks.

![Gross and net portfolio paths alongside starting position sizes in the synthetic carbon-price experiment](https://raw.githubusercontent.com/TomBussard/TransitionLens/main/figures/portfolio-path.png)

At **EUR 65 per tonne**, the example selected **24 of 80** fictional companies. The largest opening position was 5%. Over 252 simulated trading days, gross return was **8.23%** and net return was **8.03%**. The gap comes from charging for turnover when the portfolio is formed and rebalanced.

[Follow the signal into positions, turnover and returns](https://github.com/TomBussard/TransitionLens#the-experiment)

### [RatingSpread](https://github.com/TomBussard/RatingSpread)

**How much do ESG rating providers agree about individual companies?** I aligned the direction of four scores, compared each pair, then asked which firms survived a top-fifth screen under all four ratings. Correlation and overlap answer different questions. Only the second one tells me how similar the final holdings would be.

![Pairwise score correlations for the reproducible four-method rating example](https://raw.githubusercontent.com/TomBussard/RatingSpread/main/figures/rating-agreement.png)

In the historical sample, **4 companies out of 103** selected by at least one provider appeared in all four lists. The mean pairwise correlation was **0.2824**. The figure uses a separate, reproducible sample of fictional companies so readers can inspect the ranking rules and rerun the comparison.

[Compare the historical results with the reproducible example](https://github.com/TomBussard/RatingSpread#what-i-found)

## 2025

### [StratLedger](https://github.com/TomBussard/StratLedger)

**A portfolio simulator backed by a SQLite transaction ledger.** Working on investment strategies led me to a more basic question. Can I reconstruct what an account actually owned on any given date? I followed each buy and sale through the database, then valued the outstanding shares against the closing price.

![Six weekly observations showing a position growing from ten to fifteen shares and falling to eleven](https://raw.githubusercontent.com/TomBussard/StratLedger/main/figures/ledger-example.png)

In the six-week example, buying 10 shares, adding 5 and selling 4 leaves **11 shares worth 1,122** at a final price of 102. The step chart makes it easy to check which changes came from trading. The dashboard then adds strategy rules and weekly risk measures.

[Follow the three trades through the ledger](https://github.com/TomBussard/StratLedger#reconstruct-a-position)

### [GreenFolio](https://github.com/TomBussard/GreenFolio)

**Turning sustainability preferences into stock-selection rules.** I built a Streamlit portfolio builder so I could see which companies disappear when I switch from an environmental screen to a social one. Each screen uses the same environmental, social, governance and total risk scores. Lower values are better in this model.

![How many of eight fictional companies pass each of four sustainability screens](https://raw.githubusercontent.com/TomBussard/GreenFolio/main/figures/screen-comparison.png)

Of the eight example companies, **2** passed the environmental screen, **4** the social screen, **6** the balanced screen and **7** the general one. The repository lists the cutoffs so each inclusion can be checked against its four scores.

[Inspect the thresholds and build a screened portfolio](https://github.com/TomBussard/GreenFolio#one-universe-four-preferences)

### [AlpineSwap](https://github.com/TomBussard/AlpineSwap)

**The currency risk inside a municipal interest-rate swap.** A borrower receives a fixed coupon and pays one linked to EUR/CHF. I traced each annual fixing through the barrier, outstanding principal and discount factor. Then I priced the same contract over 10,000 currency paths, using Excel and Python to check the cash flows against each other.

![Distribution of the discounted borrower cash flows across ten thousand simulated EUR/CHF paths](https://raw.githubusercontent.com/TomBussard/AlpineSwap/main/figures/swap-distribution.png)

The mean discounted net cash flow was **EUR -35.67 million**. The fifth percentile was **EUR -65.14 million**. A single adverse fixing also changes later payments because the FX path carries the shock forward.

[Trace one coupon in Excel and repeat the FX stress in Python](https://github.com/TomBussard/AlpineSwap#follow-one-payment)

### [IndexFoundry](https://github.com/TomBussard/IndexFoundry)

**Constructing an equity index from its rules.** To understand why two indices can diverge even when they hold the same companies, I separated the price history from the weighting rule. The full workbench adds currency conversion, sector and country filters, fundamentals, momentum and benchmark comparisons.

![Two base-100 index paths built from identical prices with different starting allocations](https://raw.githubusercontent.com/TomBussard/IndexFoundry/main/figures/weighting-rules.png)

With the same three fictional companies, the equal-weight index finished at **108.00**. A **60/25/15** starting allocation finished at **112.60**. In the offline dashboard, moving one slider changes the weights while keeping every component price fixed.

[Try the weighting experiment and inspect the index calculations](https://github.com/TomBussard/IndexFoundry#run-locally)

## 2024

### [TetherScope](https://github.com/TomBussard/TetherScope)

**An R time-series study of volatility and Tether's share of the crypto market.** I matched **115 consecutive weekly observations** and tested both directions of a lead-lag relationship. To compare stationary series, the study models the change in log volatility and the change in the growth rate of Tether's market share.

![Weekly crypto volatility and Tether market share over the historical study window](https://raw.githubusercontent.com/TomBussard/TetherScope/main/figures/weekly-series.png)

In the three-lag model, past volatility changes helped predict changes in market-share growth (**p = 0.0157**). The reverse test gave **p = 0.5001**. This is evidence about timing within the sample. The report also shows the forecast intervals and how the estimated response changes after a shock.

[Read the econometric report and reproduce the tests](https://github.com/TomBussard/TetherScope#method-and-result)

### [ParisTape](https://github.com/TomBussard/ParisTape)

**Five French equities and 10,000 ways to hold them.** I started with adjusted prices for Air Liquide, L'Oréal, Schneider Electric, LVMH and Sanofi. I aligned their trading days, estimated how their returns moved together and sampled long-only allocations to see how much risk could change with the weights alone.

![Historical annualised volatility and return for ten thousand sampled five-stock allocations](https://raw.githubusercontent.com/TomBussard/ParisTape/main/figures/portfolio-cloud.png)

Across **1,283 common trading dates**, the lowest annualised volatility among the sampled portfolios was **17.64%**. The highest sampled Sharpe ratio was **1.08** at a zero risk-free rate. The workbook export preserves the weights behind each point.

[Open the notebook and inspect the sampled portfolios](https://github.com/TomBussard/ParisTape#explore-the-study)

## 2023

### [EcoAtlas](https://github.com/TomBussard/EcoAtlas)

**My first cross-country statistics project in R.** I brought together environmental performance, income per person, plastic waste trade and political regime for **58 countries**. Plotting individual countries first helped me see how much variation a single group average hides.

![Each country’s GDP per person against its environmental performance score, with the sample mean marked](https://raw.githubusercontent.com/TomBussard/EcoAtlas/main/figures/income-and-environment.png)

Mean environmental performance was **52.74**, with a 95% confidence interval of **49.33 to 56.15**. In a separate continent-by-regime table, 22 of 24 expected counts were below five. I used 9,999 simulated tables with the same margins to assess that comparison.

[Read the two-page study and inspect the country data](https://github.com/TomBussard/EcoAtlas#explore-the-study)
