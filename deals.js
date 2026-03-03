// Simple in-memory deals configuration used by main.html and deal.html
// Edit this array to add or update deals.
const DEALS = [
  {
    id: "biolux-seed",
    name: "Biolux Seed Round",
    company: "Biolux",
    stage: "contacting investors", // prospect | onboarding | contacting investors
    targetAmount: "5 million £",
    raisedAmount: null,
    currency: "USD",
    owner: "Chris",
    fundraisingDashboardId: "biolux", // maps to ?dashboard=biolux in index.html
    CashCommission : "4%",
    EquityCommission : "4%",
    Retainer : "1800 £"
  },
  {
    id: "iq500-series-a",
    name: "IQ500 Series A",
    company: "IQ500",
    stage: "prospect",
    targetAmount: null,
    raisedAmount: null,
    currency: "USD",
    owner: "MJ",
    fundraisingDashboardId: "IQ500", // maps to ?dashboard=IQ500 in index.html
  },
];

