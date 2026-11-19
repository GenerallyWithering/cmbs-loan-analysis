# CMBS loan analysis: 700 6th Street (Benchmark 2026-B44, Loan 3)

A loan-level credit analysis of a real, newly issued CMBS conduit loan, built only from the public SEC prospectus. The loan is a **$54.2M, 10-year, interest-only office loan** on a Class A building in Washington, DC.

| | |
|---|---|
| Deal | Benchmark 2026-B44 Mortgage Trust (10-year conduit), final prospectus, Form 424B2 filed Nov 10, 2026 |
| Loan | #3 of the pool (6.7%): 700 6th Street NW, Washington, DC, 304,947 SF office, built 2009 |
| Terms | $54,197,000, 7.38% fixed, Actual/360, interest-only for 120 months, matures Oct 6, 2036 |
| Use | Acquisition ($79.5M purchase price) by FitzWalter Capital Partners |

## What's in the repo
- `model/700_6th_Street_CMBS_Loan_Analysis.xlsx` is the Excel model:
  - **Inputs:** every input typed in blue, with the exact prospectus page.
  - **Metrics:** formulas in black, no typed numbers inside formulas.
  - **Historical:** the NOI trend.
  - **Stress:** NOI cuts, tenant events, rent rolling before maturity, a refinance test, and which lender limit binds.
- `writeup/credit_memo.md` is the 1-page credit memo: the loan, key metrics, stress results, risks, mitigants and verdict.
- `writeup/resume_and_pitch.md` has a resume bullet and a 60-second interview pitch.
- `STUDY_GUIDE.md` shows how to check every number yourself, plus the interview questions this project invites.
- `sources/SOURCES.md` lists the exact SEC EDGAR links and the page for every input.
- `scripts/build_model.py` is the Python script (openpyxl) that writes the workbook.

## Results
| Metric | My calculation | Prospectus |
|---|---|---|
| Annual debt service (IO, Actual/360) | $4,055,290.53 | $4,055,290.56 (cent rounding) |
| DSCR (UW NCF) | 2.43x | 2.43x |
| LTV at origination / maturity | 63.0% / 63.0% | 63.0% / 63.0% |
| Debt yield (UW NOI) | 19.3% | 19.3% |

**Stress tests:**
- NOI −10%: 2.17x DSCR, 17.3% debt yield.
- NOI −20%: 1.91x, 15.4%.
- The largest tenant renews at market rent: 1.69x, 13.7%.
- 99.7% of the rent expires before the loan matures.

**Refinance in 2036:**
- At underwritten NOI it refinances at 7–9% rates, about 1.25x the balloon.
- If the largest tenant renews at market rent, it is about $5.9M short.
- The binding limit is value, not the interest rate. The DSCR test only fails above about a 14% rate.

**Verdict:** comfortable at AAA–A, cautious at BBB- and below, because the risk is binary and concentrated in the 2033 lease rollover. See the memo.

## Method notes
- **Annual debt service for an interest-only loan on Actual/360** = loan × rate × 365/360. That matches the prospectus definition and Annex A-1.
- **DSCR** = underwritten NCF / annual debt service. **Debt yield** = underwritten NOI / loan. **LTV** = loan (or balloon) / as-is appraised value.
- **Stress cases** keep reserves (NOI − NCF) fixed. Tenant cases change base rent only, so lost expense reimbursements and new leasing costs are ignored.
- **The refinance test** uses my own lender assumptions (marked as assumptions in the model): 30-year amortization, at least 1.25x DSCR and 10% debt yield, at most 65% LTV, and a 10% exit cap rate.
- **Scope:** the deal's bond waterfall and the other loans in the pool were not modeled.

## How it was made
The inputs were read from the prospectus and every one is cited to its page. The model, memo and checks were built with help from an AI assistant (Claude). All figures can be traced and recalculated from the sources: see `STUDY_GUIDE.md`.

To rebuild the workbook: `python3 -m venv .venv && .venv/bin/pip install openpyxl && .venv/bin/python scripts/build_model.py`
