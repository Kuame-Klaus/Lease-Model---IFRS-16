# IFRS 16 Lease Portfolio Model

A single-workbook Excel model that computes IFRS 16 lease accounting (liability, right-of-use asset, interest, depreciation, and the profit-and-loss timing effect) across an 11,000-lease portfolio, split into an arrears-pay book and an advance-pay/ownership-transfer book, with a live calculation engine and a visual KPI/charts dashboard on top.

File: IFRS_16_Model_-Lease_Portfolio_11000.xlsx Scale: 11,000 leases · 6 sheets · ~135,000 formulas

# Architecture
┌───────────────────────────────────────────────────────────────────────┐
│                    IFRS 16 LEASE PORTFOLIO MODEL                      │
│                     11,000 leases across 2 books                      │
└───────────────────────────────────────────────────────────────────────┘

  RAW LEASE DATA  (source of truth — one row per lease)
  ┌───────────────────────────┐      ┌────────────────────────────────┐
  │ Lease Portfolio (Arrears) │      │ LP (Advance Pay & Ownership)    │
  │ 10,000 leases · cols A–X  │      │  1,000 leases · cols A–Y        │
  │────────────────────────────│      │──────────────────────────────────│
  │ Country · Asset Class      │      │ Country · Asset Class            │
  │ Commencement Date · Term   │      │ Commencement Date · Term         │
  │ IBR · Underlying Value     │      │ IBR · Payment Timing             │
  │ Exemption Applies?         │      │ Ownership Transfer?              │
  │ Lease Liability · ROU      │      │ Lease Liability · ROU            │
  │ Interest · Depreciation    │      │ Interest · Depreciation          │
  │ P&L Charge · Timing Effect │      │ P&L Charge · Timing Effect       │
  └──────────────┬─────────────┘      └────────────────┬─────────────---┘
                 │                                       │
                 └───────────────────┬───────────────────┘
                                      │  SUMIF / COUNTIF / SUMPRODUCT
                                      ▼
                     ┌────────────────────────────────────┐
                     │             Chart Data              │
                     │   calculation engine — formulas     │
                     │   only, no hardcoded values         │
                     │──────────────────────────────────────│
                     │  • By Asset Class   (Arrears / Adv / │
                     │    Consolidated)                     │
                     │  • By Country       (same 3 cuts)    │
                     │  • Lease Expiry Profile (2026–2046+) │
                     │  • IBR Distribution                  │
                     │  • On-Balance-Sheet vs Exempted      │
                     │  • Country × Asset Class heat-map    │
                     │  • Demographics: term bands, vintage,│
                     │    asset-value bands, book mix,      │
                     │    payment timing, ownership xfer    │
                     └───────────────────┬────────────────--┘
                                          │
                 ┌────────────────────────┼────────────────────────┐
                 ▼                                                  ▼
   ┌───────────────────────────┐                     ┌───────────────────────────┐
   │      Summary Dashboard     │                     │      Charts Dashboard      │
   │  tabular KPI totals and    │                     │  visual layer — first tab  │
   │  breakdowns (original)     │                     │──────────────────────────────│
   └───────────────────────────┘                     │  • KPI cards × 3 books      │
                                                       │  • Bar / donut / heat-map   │
                                                       │    charts × 3 books         │
                                                       │  • Portfolio Demographics   │
                                                       │    section                  │
                                                       └───────────────────────────┘

  SUPPORTING TABS
  ┌────────────────┐   ┌────────────────┐
  │  Assumptions     │   │  Methodology    │
  │  exemption        │   │  modelling notes │
  │  thresholds        │   │                 │
  └────────────────┘   └────────────────┘

Data flow, in one line: Raw lease rows (2 sheets) → Chart Data (aggregation engine) → Summary Dashboard + Charts Dashboard

Nothing calculates backwards — the two raw lease sheets are the only place with source data; every other tab is 100% formulas referencing either the raw sheets or Chart Data.

# Sheets
Sheet	Rows	Purpose
Charts Dashboard	~250	Visual KPI cards, bar/donut/heat-map charts, and a demographics section, split by Arrears book / Advance-Ownership book / Consolidated
Summary Dashboard	118	Original tabular breakdown of the portfolio by asset class and country, per book and consolidated
Assumptions	9	Exemption thresholds and other portfolio-level modelling assumptions
Methodology	26	Written notes on how the model calculates liability, ROU, interest, depreciation, and the timing effect
Lease Portfolio (Arrears)	10,000 leases	Lease-level data and IFRS 16 calculations for the arrears-pay book
LP (Advance Pay & Ownership)	1,000 leases	Lease-level data and IFRS 16 calculations for the advance-pay / ownership-transfer book
Chart Data	~232	Calculation engine — aggregates the two raw sheets by asset class, country, expiry year, IBR band, and demographic cut, feeding both dashboards
Key formulas

Lease liability (initial, arrears-pay):

= IF(Exemption Applies = "Yes", 0, Annual Payment × (1 − (1 + IBR)^−Term) ÷ IBR)

Years elapsed (drives the current-period figures):

= DATEDIF(Commencement Date, TODAY(), "y")

⚠️ This makes the model date-sensitive — see Known considerations below.

Aggregation pattern used throughout Chart Data:

= SUMIF('Lease Portfolio (Arrears)'!$D$2:$D$10001, $A6, 'Lease Portfolio (Arrears)'!$T$2:$T$10001)
+ SUMIF('LP (Advance Pay & Ownership)'!$D$2:$D$1001, $A6, 'LP (Advance Pay & Ownership)'!$U$2:$U$1001)

Every consolidated figure is the sum of the matching SUMIF/COUNTIF across both raw sheets — never a single-sheet lookup — so a lease booked in either book is always captured in the consolidated view.

Lease expiry year (used for the maturity profile chart):

= SUMPRODUCT((YEAR(Commencement Date range) + Term range) = target year)
Known considerations
TODAY()-driven recalculation. The "Years Elapsed" column (and everything downstream of it — opening/closing liability, interest, depreciation, P&L charge) recalculates relative to the current date every time the workbook opens. Figures will drift slightly day to day as leases cross an annual anniversary. This is expected behaviour, not a bug — but it means the numbers you see are always "as of today," not frozen at a point in time.
Manual vs. Automatic calculation. With ~135,000 formulas, some environments default this workbook to Manual calculation for performance. If numbers look stale or show 0, press Ctrl+Alt+F9 (full recalculation) and check Formulas → Calculation Options is set to Automatic.
Chart Data is a calculation layer, not an input sheet. Don't edit it directly — every cell is a formula derived from the two raw lease sheets. To change portfolio data, edit the raw sheets; to change exemption thresholds, edit Assumptions.

# Usage
1. Open the workbook and let it fully recalculate (Ctrl+Alt+F9 if needed).
2. For a quick portfolio read: start on Charts Dashboard.
3. For the underlying numbers behind any chart: the same figures live on Chart Data, laid out in the same order.
4. For lease-level detail or to add/edit leases: go to Lease Portfolio (Arrears) or LP (Advance Pay & Ownership) directly — every downstream sheet will recalculate automatically.

<img width="1519" height="761" alt="Screenshot 2026-08-17 at 9 44 36 PM" src="https://github.com/user-attachments/assets/3c9fd692-5884-469b-b0fc-28f18154d016" />
<img width="1519" height="720" alt="Screenshot 2026-08-17 at 9 44 51 PM" src="https://github.com/user-attachments/assets/c53f84a9-542e-4c9a-8564-56f02ddff1a4" />
<img width="1519" height="914" alt="Screenshot 2026-08-17 at 9 45 25 PM" src="https://github.com/user-attachments/assets/880b2851-826e-41be-bc6c-6a56ad498d4e" />
<img width="1519" height="914" alt="Screenshot 2026-08-17 at 9 45 40 PM" src="https://github.com/user-attachments/assets/e481f9b4-c452-41e3-b588-e307407b34a4" />
<img width="827" height="566" alt="Screenshot 2026-08-17 at 9 46 44 PM" src="https://github.com/user-attachments/assets/f27b26b3-bea8-4a2f-b9fb-571caad38db9" />
<img width="827" height="566" alt="Screenshot 2026-08-17 at 9 46 58 PM" src="https://github.com/user-attachments/assets/9be2c014-3c29-48ff-9c95-e85ef9163048" />
<img width="827" height="645" alt="Screenshot 2026-08-17 at 9 47 10 PM" src="https://github.com/user-attachments/assets/a99bf13f-75a6-4991-b6d6-dcbc22d967ec" />
<img width="734" height="291" alt="Screenshot 2026-08-17 at 9 47 35 PM" src="https://github.com/user-attachments/assets/5ab85150-7646-4e7c-b442-4eb318f8a325" />
<img width="798" height="857" alt="Screenshot 2026-08-17 at 9 48 33 PM" src="https://github.com/user-attachments/assets/aa0036f5-98fd-43b8-bbb6-6796c70d3720" />
<img width="558" height="726" alt="Screenshot 2026-08-17 at 9 48 45 PM" src="https://github.com/user-attachments/assets/198ee40f-db3d-4c4b-8dd9-a0c357533e05" />
<img width="1900" height="998" alt="Screenshot 2026-08-17 at 9 50 07 PM" src="https://github.com/user-attachments/assets/1aac2534-bfeb-4b5f-af0e-c1b1395052aa" />
<img width="1900" height="998" alt="Screenshot 2026-08-17 at 9 50 27 PM" src="https://github.com/user-attachments/assets/b306aabb-f72b-4b16-a594-14c60a273fd8" />
<img width="1900" height="998" alt="Screenshot 2026-08-17 at 9 50 37 PM" src="https://github.com/user-attachments/assets/dc06c80e-ec10-411f-8715-6bd2e0f544a6" />
<img width="1900" height="998" alt="Screenshot 2026-08-17 at 9 50 57 PM" src="https://github.com/user-attachments/assets/b62980e0-40c7-4851-a33c-f062a3398a7c" />
