# Maritime CLT & Inventory Cost Optimizer

Streamlit dashboard to compare maritime logistics scenarios using pipeline
inventory value, working capital, annual inventory financing and transportation
costs.

## Run locally

```powershell
cd Packaging\FCF_TDC
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Run the calculation tests from the same directory:

```powershell
python -m unittest discover -s tests -v
```

## Inputs and analysis basis

- Configure the financing rate using WACC or a separate inventory carrying
  cost rate. The two rates are alternatives and are never added together.
- Cargo value and average container consumption are common to both routes.
- Annual container volume drives annual transportation costs.
- Choose either **Transportation transit time only** or **Complete Customer
  Lead Time** for pipeline inventory and per-container financing. Complete CLT
  is treated as an end-to-end total; physical transit is not added to it.
- Complete CLT must be at least as long as physical transportation transit.
- If daily consumption and annual volume divided by days per year differ, the
  app shows a reconciliation warning. Pipeline financing uses daily consumption
  and analysis days; annual transportation cost uses annual volume.

## Calculation definitions

Let `V` be cargo value per container, `q` containers consumed per day, `r` the
selected annual financing rate, `d` the selected analysis days, `D` days per
year, and `N` annual container volume.

- Pipeline inventory value: `V × q × d`
- Working capital released: `Pipeline A − Pipeline B`
- Financing cost per container: `V × r × d ÷ D`
- Annual financing cost: `Pipeline inventory value × r`
- Transport cost per container: ocean freight + inland transport + additional
  logistics costs
- Annual transportation cost: transport cost per container × `N`
- Annual total economic logistics cost: annual transportation cost + annual
  financing cost
- Net annual economic benefit: Scenario A annual cost − Scenario B annual cost
- Financing-only maximum freight premium per container: Scenario A per-container
  financing cost − Scenario B per-container financing cost

The maximum freight premium is a financing-only break-even point. It excludes
service-level improvements, safety stock, obsolescence, warehousing, demurrage,
detention and stockout risks. Pipeline inventory value and working capital
released are capital measures, not P&L savings.

## Reference validation scenario

At 8% WACC, €100,000 cargo value, one container consumed daily, 365 annual
containers, and transit times of 40 versus 30 days:

- Pipeline inventory: €4,000,000 versus €3,000,000.
- Annual financing cost: €320,000 versus €240,000.
- Working capital released: €1,000,000.
- Annual financing savings: €80,000.
- Financing-only premium threshold: about €219.18 per container.

The net annual benefit from the full economic cost comparison is **€43,500**
using €1,900 and €2,000 transportation costs per container respectively:
€80,000 financing savings less €36,500 additional transportation cost. The
annual pipeline formula and per-container formula reconcile at the reference
assumptions (one container consumed per day and 365 containers per year).
