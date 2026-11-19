# Study guide: own every number (about 45 minutes)

Do these checks yourself with a calculator or a blank Excel sheet. After that, every number in the memo is one you can defend.

## 1. Find the inputs (10 min)
Open the prospectus (link in `sources/SOURCES.md`) and search for "Collateral Asset Summary – Loan No. 3".
- On **A-3-28**, find: loan $54,197,000, rate 7.38000%, interest-only for 120 months, appraised value $86,000,000, UW NOI $10,442,065, UW NCF $9,847,208.
- On **A-3-29**, find the two tenants and their lease end dates (Apr 2033 and Sep 2036, with a termination option in Sep 2031).

## 2. Debt service (5 min)
- Interest-only, on Actual/360, means: **54,197,000 × 7.38% × 365 / 360 = $4,055,290.53** a year.
- Why 365/360? The loan charges a 360-day year's rate for each actual day, so a full year costs 365/360 of the stated rate.
- Check: the prospectus's Annex A-1 (p. A-1-9) shows $337,940.88 a month and $4,055,290.56 a year. The 3-cent difference is because they round the monthly payment to the cent first.

## 3. The three ratios (5 min)
- **DSCR** = NCF / debt service = 9,847,208 / 4,055,290.53 = **2.43x**
- **LTV** = loan / value = 54,197,000 / 86,000,000 = **63.0%** (the same at maturity, because nothing is paid down)
- **Debt yield** = NOI / loan = 10,442,065 / 54,197,000 = **19.3%**

## 4. One stress test by hand (10 min)
- NOI −20% = 8,353,652.
- Reserves (NOI − NCF) stay at 594,857, so NCF = 7,758,795.
- DSCR = 7,758,795 / 4,055,290.53 = **1.91x**. Debt yield = 8,353,652 / 54,197,000 = **15.4%**.

## 5. The tenant story (10 min)
- Eversheds pays $81.58/SF on 143,160 SF of office. The market asking rent is $60.53 (A-3-32).
- Renewing at market: (60.53 − 81.58) × 143,160 = **−$3.0M** of NOI. NOI drops to $7.43M, a 1.69x DSCR and a 13.7% debt yield.
- At refinance (my assumptions): the largest new loan = 7.43M / 10% exit cap × 65% LTV = **$48.3M**, about $5.9M short of the $54.2M balloon.

## Interview questions to practice
1. **Why is the DSCR so high if the building is a quarter empty?**
   The loan is small against the building's cash flow. The buyer paid $79.5M, a 13.1% cap rate on UW NOI, and the 7.38% interest-only payment is about $4.1M against $9.8M of NCF.
2. **What is the single biggest risk?**
   Eversheds: 68.5% of rent, with its lease ending Apr 2033. Above-market rent plus a 22.7% submarket vacancy rate means a renewal could come at lower rent, or the tenant could leave.
3. **Why does interest-only matter here?**
   No principal is paid down, so the full $54.2M must be refinanced in 2036, right after the tenant events.
4. **What protects the lender?**
   - the DSCR sweep below 1.50x;
   - the Eversheds trigger, which traps cash up to $100/SF (about $14.6M) if they don't renew;
   - the leasing reserve and the hard lockbox;
   - 63% LTV with sponsor equity below it;
   - no mezzanine debt allowed.
5. **Which bonds would you buy?**
   The senior and upper-mezzanine classes (AAA to A), where this loan's risk is well covered. I'd be careful at BBB- and below, because a 2033 rollover problem would push losses to the bottom of the stack first.
6. **What didn't you model?**
   The deal's waterfall, the other loans, and lost expense reimbursements in the tenant cases. The refinance limits are my own assumptions.
