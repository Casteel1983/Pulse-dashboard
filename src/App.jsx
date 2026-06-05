import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, AreaChart, Area, PieChart, Pie, Legend
} from "recharts";
import * as Recharts from "recharts";

// API key — injected at build time by Vite from VITE_ANTHROPIC_KEY secret
// Falls back to empty string if not set (AI features will show error)
const ANTHROPIC_KEY = (typeof __VITE_ANTHROPIC_KEY__ !== "undefined" && __VITE_ANTHROPIC_KEY__)
  ? __VITE_ANTHROPIC_KEY__
  : "";

// ── Theme ─────────────────────────────────────────────────────────────────────
const AMBER  = "#1E5FCC";   // primary blue
const TEAL   = "#2980D9";   // mid blue
const GREEN  = "#1A9E5C";   // green (positive)
const RED    = "#D93025";   // red (negative)
const BG     = "#F4F7FB";   // light page bg
const BG2    = "#FFFFFF";   // card bg
const BORDER = "#D0DCF0";   // border
const MUTED  = "#6B7A99";   // muted text
const TEXT   = "#1A2340";   // primary text
const COLORS = ["#1E5FCC","#2980D9","#1A9E5C","#D93025","#7C3AED","#D97706","#0891B2","#DB2777"];
const REPS   = ["Tiffany", "Larry", "Austin"];
const REP_COLORS = { Tiffany: "#7C3AED", Larry: "#0891B2", Austin: "#D97706", House: "#059669", "Car Dealer": "#DC2626" };

const FILE_SLOTS = [
  { id: "master",     label: "⬡ Master Upload", desc: "Pulse_Master_Upload_Wxx.xlsx — uploads AR, CustomerComp, WTD, and all AD programs in one click", isMaster: true },
  { id: "customers",  label: "Customer List",   desc: "Account list with salesperson assignment" },
  { id: "ar",         label: "AR / Aging",      desc: "Balances · 30/60/90/120+ day aging (use Master Upload instead)" },
  { id: "weekComp",   label: "Week Comp",       desc: "Week-to-Week Customer Comp by Department (use Master Upload instead)" },
  { id: "sales",      label: "Sales Data",      desc: "Current & prior year sales — all years on one sheet" },
];


// Car Dealer accounts mapped to primary rep by city dominance
const CAR_DEALER_REP = {
  101256:"Larry",101298:"Larry",101300:"Larry",101374:"Larry",101568:"House",
  200232:"Larry",200295:"House",200336:"Tiffany",200348:"House",200349:"House",
  200360:"House",200365:"House",200445:"Tiffany",200455:"Larry",200456:"Larry",
  200461:"Tiffany",200484:"Tiffany",200524:"Larry",200527:"Larry",200533:"Tiffany",
  200539:"Tiffany",200552:"House",200554:"Larry",200580:"Larry",200587:"House",
  200588:"House",200626:"House",200682:"Tiffany",200688:"House",200734:"Larry",
  200765:"House",200767:"House",200769:"House",200772:"House",200773:"House",
  200774:"House",200777:"Larry",200778:"Larry",200783:"House",200784:"Tiffany",
  200790:"House",200801:"House",200831:"Larry",200832:"Larry",200833:"Larry",
  200835:"House",200838:"House",200841:"House",200850:"House",200851:"House",
  200854:"House",200859:"House",200861:"Tiffany",200864:"Tiffany",200865:"Larry",
  200882:"House",200908:"House",200930:"House",200933:"House",200943:"House",
  200948:"House",200951:"House",1999999:"Larry",2000000:"House",2000009:"House",
  2000011:"House",2000019:"House",2000024:"House",2000025:"House",2000026:"House",
  2000027:"Larry"
};

// Build shared car-dealer account entries from action plan + customers
function buildCarDealerAccounts(actionPlan, customers) {
  const apByNum = {};
  (actionPlan || []).forEach(a => { apByNum[a.custNum] = a; });
  const cdCustomers = (customers || SEED_CUSTOMERS).filter(c => c.salesman === "Car Dealer");
  return cdCustomers.map(c => {
    const existing = apByNum[c.num];
    const primaryRep = CAR_DEALER_REP[c.num] || "House";
    return existing
      ? { ...existing, salesman: primaryRep, _carDealer: true, _shared: true }
      : {
          custNum: c.num, customer: c.name, city: c.city || "", salesman: primaryRep,
          sales2025: 0, sales2026: c.ytdComp || 0, change: c.ytdComp || 0,
          gpPct: 0, action: "CAR DEALER", topDept: "Car Dealer Account",
          declinedDept: "", focus: "On track", _carDealer: true, _shared: true,
        };
  });
}

const SEED_BRANCH_DATA = {"weeklySales":[{"week":1,"Byron":71317.0,"Tifton":45986.0,"Statesboro":46756.0,"Athens":73570.0,"Byron25":163452.0,"Tifton25":124056.0,"Statesboro25":115721.0,"Athens25":115171.0},{"week":2,"Byron":570752.0,"Tifton":383464.0,"Statesboro":481873.0,"Athens":383818.0,"Byron25":377044.0,"Tifton25":303299.0,"Statesboro25":462387.0,"Athens25":273827.0},{"week":3,"Byron":426682.0,"Tifton":376015.0,"Statesboro":404746.0,"Athens":342498.0,"Byron25":425593.0,"Tifton25":246446.0,"Statesboro25":427419.0,"Athens25":349566.0},{"week":4,"Byron":522986.0,"Tifton":351882.0,"Statesboro":380511.0,"Athens":361139.0,"Byron25":247061.0,"Tifton25":162354.0,"Statesboro25":250715.0,"Athens25":268575.0},{"week":5,"Byron":475548.0,"Tifton":308672.0,"Statesboro":415940.0,"Athens":264278.0,"Byron25":320762.0,"Tifton25":262329.0,"Statesboro25":405301.0,"Athens25":254913.0},{"week":6,"Byron":476077.0,"Tifton":349456.0,"Statesboro":370412.0,"Athens":420047.0,"Byron25":474253.0,"Tifton25":350519.0,"Statesboro25":493114.0,"Athens25":354159.0},{"week":7,"Byron":465206.0,"Tifton":318158.0,"Statesboro":436169.0,"Athens":392013.0,"Byron25":420630.0,"Tifton25":308025.0,"Statesboro25":461629.0,"Athens25":386143.0},{"week":8,"Byron":509095.0,"Tifton":433433.0,"Statesboro":471785.0,"Athens":423549.0,"Byron25":421598.0,"Tifton25":305074.0,"Statesboro25":495207.0,"Athens25":352937.0},{"week":9,"Byron":545565.0,"Tifton":454495.0,"Statesboro":440677.0,"Athens":447750.0,"Byron25":388503.0,"Tifton25":356486.0,"Statesboro25":487012.0,"Athens25":332534.0},{"week":10,"Byron":578802.0,"Tifton":470119.0,"Statesboro":497236.0,"Athens":412614.0,"Byron25":634848.0,"Tifton25":414365.0,"Statesboro25":571197.0,"Athens25":375647.0},{"week":11,"Byron":640012.0,"Tifton":543489.0,"Statesboro":504057.0,"Athens":501260.0,"Byron25":475206.0,"Tifton25":301003.0,"Statesboro25":495204.0,"Athens25":403830.0},{"week":12,"Byron":601780.0,"Tifton":423219.0,"Statesboro":468571.0,"Athens":471418.0,"Byron25":453356.0,"Tifton25":300523.0,"Statesboro25":536190.0,"Athens25":364320.0},{"week":13,"Byron":644955.0,"Tifton":459113.0,"Statesboro":537096.0,"Athens":499530.0,"Byron25":397627.0,"Tifton25":315204.0,"Statesboro25":521326.0,"Athens25":359990.0},{"week":14,"Byron":695711.0,"Tifton":524967.0,"Statesboro":518940.0,"Athens":435410.0,"Byron25":531766.0,"Tifton25":342928.0,"Statesboro25":515104.0,"Athens25":404237.0},{"week":15,"Byron":722978.0,"Tifton":441601.0,"Statesboro":487888.0,"Athens":372553.0,"Byron25":591778.0,"Tifton25":382490.0,"Statesboro25":557293.0,"Athens25":549426.0},{"week":16,"Byron":581034.0,"Tifton":394978.0,"Statesboro":514215.0,"Athens":385858.0,"Byron25":596454.0,"Tifton25":334891.0,"Statesboro25":527256.0,"Athens25":454735.0},{"week":17,"Byron":631907.0,"Tifton":422233.0,"Statesboro":519390.0,"Athens":453314.0,"Byron25":484884.0,"Tifton25":357675.0,"Statesboro25":539418.0,"Athens25":324248.0},{"week":18,"Byron":550852.0,"Tifton":443646.0,"Statesboro":521227.0,"Athens":511785.0,"Byron25":519581.0,"Tifton25":396849.0,"Statesboro25":531112.0,"Athens25":347877.0},{"week":19,"Byron":533289.0,"Tifton":470135.0,"Statesboro":501654.0,"Athens":551444.0,"Byron25":595299.0,"Tifton25":430969.0,"Statesboro25":584156.0,"Athens25":381174.0},{"week":20,"Byron":537770.0,"Tifton":401595.0,"Statesboro":463117.0,"Athens":399092.0,"Byron25":488435.0,"Tifton25":327879.0,"Statesboro25":529185.0,"Athens25":381263.0},{"week":21,"Byron":581257.0,"Tifton":378909.0,"Statesboro":475853.0,"Athens":407850.0,"Byron25":550075.0,"Tifton25":395584.0,"Statesboro25":575972.0,"Athens25":405444.0},{"week":22,"Byron":385150.0,"Tifton":287880.0,"Statesboro":425248.0,"Athens":396320.0,"Byron25":444933.0,"Tifton25":347991.0,"Statesboro25":428728.0,"Athens25":387835.0}],"branches":{"Byron":{"q1_2025":5266440.63,"q1_2026":6822661.29,"q1_gp25":705844.68,"q1_gp26":861094.27,"q2_2025":4736699.45,"q2_2026":4926064.14,"q2_gp25":767300.85,"q2_gp26":658146.6},"Tifton":{"q1_2025":3825950.57,"q1_2026":5138331.47,"q1_gp25":566747.56,"q1_gp26":720485.34,"q2_2025":3240986.41,"q2_2026":3545113.76,"q2_gp25":570331.51,"q2_gp26":497861.47},"Statesboro":{"q1_2025":5835303.76,"q1_2026":5667773.51,"q1_gp25":745163.56,"q1_gp26":639736.99,"q2_2025":4675340.65,"q2_2026":4215588.46,"q2_gp25":724578.62,"q2_gp26":490241.22},"Athens":{"q1_2025":4287894.3,"q1_2026":5220784.25,"q1_gp25":577373.95,"q1_gp26":672226.68,"q2_2025":3539958.04,"q2_2026":3686327.89,"q2_gp25":567819.82,"q2_gp26":479579.67}},"tiftonQ1Depts":[{"dept":"1 - BYRON","sales2025":5266440.63,"sales2026":6822661.29,"gp2025":705844.68,"gp2026":861094.27},{"dept":"3 - STATESBORO","sales2025":5835303.76,"sales2026":5667773.51,"gp2025":745163.56,"gp2026":639736.99},{"dept":"5 - ATHENS","sales2025":4287894.3,"sales2026":5220784.25,"gp2025":577373.95,"gp2026":672226.68},{"dept":"2 - TIFTON","sales2025":3825950.57,"sales2026":5138331.47,"gp2025":566747.56,"gp2026":720485.34},{"dept":"RAD LT TRUCK","sales2025":1519974.15,"sales2026":2014306.08,"gp2025":222260.8,"gp2026":292854.91},{"dept":"TRUCK TIRES","sales2025":1039514.78,"sales2026":1133774.88,"gp2025":124413.63,"gp2026":111761.18},{"dept":"PASSENGER TIRES","sales2025":703217.54,"sales2026":1118847.12,"gp2025":138720.89,"gp2026":189943.94},{"dept":"FARM TIRES","sales2025":263687.6,"sales2026":278661.49,"gp2025":39236.24,"gp2026":33798.0},{"dept":"ST TRAILER","sales2025":111818.91,"sales2026":210417.35,"gp2025":21358.76,"gp2026":43126.35},{"dept":"OFF THE ROAD TIRES","sales2025":94922.14,"sales2026":167685.84,"gp2025":8105.2,"gp2026":14697.57},{"dept":"INDUSTRIAL TIRES","sales2025":26644.71,"sales2026":100424.94,"gp2025":3463.2,"gp2026":9692.53},{"dept":"TUBES","sales2025":41187.21,"sales2026":56254.35,"gp2025":8705.12,"gp2026":12141.58},{"dept":"GA EPD FEE","sales2025":20908.04,"sales2026":22857.42,"gp2025":0.04,"gp2026":0.42},{"dept":"VALVE STEMS","sales2025":527.18,"sales2026":16268.09,"gp2025":252.4,"gp2026":5252.01},{"dept":"WHEEL WEIGHTS","sales2025":206.21,"sales2026":9524.49,"gp2025":91.52,"gp2026":4431.67},{"dept":"TIRE TOOLS","sales2025":249.53,"sales2026":9513.72,"gp2025":67.84,"gp2026":2833.74},{"dept":"PATCHES AND REPAIR","sales2025":519.39,"sales2026":4511.29,"gp2025":161.42,"gp2026":1482.06},{"dept":"LAWN & GARDEN","sales2025":2881.42,"sales2026":4320.11,"gp2025":581.49,"gp2026":848.16},{"dept":"WHEELS","sales2025":1876.75,"sales2026":2216.06,"gp2025":285.71,"gp2026":635.91},{"dept":"MOUNTING LUBE","sales2025":260.79,"sales2026":1659.5,"gp2025":72.42,"gp2026":686.18},{"dept":"ALIGNMENT SHIMS","sales2025":0.0,"sales2026":719.96,"gp2025":0.0,"gp2026":243.44},{"dept":"FREIGHT CHARGES","sales2025":0.0,"sales2026":171.0,"gp2025":0.0,"gp2026":76.0},{"dept":"OUTSIDE PURCHASE","sales2025":354.31,"sales2026":96.0,"gp2025":66.38,"gp2026":0.0}]};


// Associate Dealer Programs — Toyo Tifton District 360
// PCR/LTR & TBR Tiers: Entry<60 | Tier 1: 60 | Tier 2: 125 | Tier 3: 250 | Tier 4: 400 | Tier 5: 750
const AD_PROGRAMS = {"200635":{"program":"Toyo Associate Dealer","toyoNum":"A222977","branch":"Tifton","pcr":{"primary":147,"secondary":12,"total":159,"pct":92.4},"tbr":{"primary":0,"secondary":0,"total":0}},"101326":{"program":"Toyo Associate Dealer","toyoNum":"A155757","branch":"Tifton","pcr":{"primary":132,"secondary":0,"total":132,"pct":100.0},"tbr":{"primary":0,"secondary":6,"total":5}},"200293":{"program":"Toyo Associate Dealer","toyoNum":"A004301","branch":"Tifton","pcr":{"primary":50,"secondary":6,"total":56,"pct":89.2},"tbr":{"primary":0,"secondary":20,"total":20}},"101080":{"program":"Toyo Associate Dealer","toyoNum":"A232704","branch":"Tifton","pcr":{"primary":53,"secondary":15,"total":68,"pct":77.9},"tbr":{"primary":0,"secondary":0,"total":0}},"100282":{"program":"Toyo Associate Dealer","toyoNum":"A004125","branch":"Tifton","pcr":{"primary":45,"secondary":0,"total":45,"pct":100.0},"tbr":{"primary":0,"secondary":0,"total":0}},"200827":{"program":"Toyo Associate Dealer","toyoNum":"A232705","branch":"Tifton","pcr":{"primary":20,"secondary":18,"total":38,"pct":52.6},"tbr":{"primary":0,"secondary":0,"total":0}},"101108":{"program":"Toyo Associate Dealer","toyoNum":"A232683","branch":"Tifton","pcr":{"primary":12,"secondary":21,"total":33,"pct":36.3},"tbr":{"primary":0,"secondary":4,"total":4}},"200885":{"program":"Toyo Associate Dealer","toyoNum":"A222590","branch":"Tifton","pcr":{"primary":8,"secondary":24,"total":32,"pct":25.0},"tbr":{"primary":0,"secondary":1,"total":1}}};
const ASCENSO_PROGRAMS = {"100043":{"qty":14,"amount":3994.49,"details":8,"invoices":8},"100076":{"qty":34,"amount":11158.75,"details":8,"invoices":8},"100252":{"qty":24,"amount":4332.88,"details":7,"invoices":6},"100350":{"qty":17,"amount":11371.91,"details":8,"invoices":8},"100366":{"qty":7,"amount":3251.89,"details":8,"invoices":8},"101161":{"qty":1,"amount":505.28,"details":1,"invoices":1},"101231":{"qty":12,"amount":6551.74,"details":9,"invoices":9},"101247":{"qty":9,"amount":2812.75,"details":7,"invoices":7},"101463":{"qty":105,"amount":12574.23,"details":8,"invoices":8},"101512":{"qty":16,"amount":3571.87,"details":9,"invoices":9},"101849":{"qty":12,"amount":2898.39,"details":13,"invoices":13},"101995":{"qty":4,"amount":706.4,"details":1,"invoices":1},"200220":{"qty":7,"amount":1053.22,"details":4,"invoices":4},"200266":{"qty":15,"amount":9718.38,"details":8,"invoices":8},"200595":{"qty":43,"amount":21225.16,"details":17,"invoices":14},"200635":{"qty":13,"amount":1877.72,"details":12,"invoices":12},"200690":{"qty":28,"amount":6245.49,"details":11,"invoices":11},"200891":{"qty":13,"amount":6027.63,"details":8,"invoices":7},"301222":{"qty":12,"amount":1802.62,"details":8,"invoices":7},"2000046":{"qty":34,"amount":5014.66,"details":1,"invoices":1},"2000049":{"qty":3,"amount":3698.35,"details":3,"invoices":3},"3000022":{"qty":5,"amount":1926.1,"details":7,"invoices":7},"3000268":{"qty":5,"amount":3971.66,"details":4,"invoices":4},"3000714":{"qty":26,"amount":6086.11,"details":7,"invoices":7}};
// Ascenso tiers — update when confirmed (placeholder: Entry<10, Bronze 10-24, Silver 25-49, Gold 50-99, Platinum 100+)
const ASCENSO_TIERS = [
  { label:"Tier 3", min:300000, payout:0.03, color:"#7C3AED" },
  { label:"Tier 2", min:100000, payout:0.02, color:"#0891B2" },
  { label:"Tier 1", min:35000,  payout:0.01, color:"#059669" },
  { label:"Entry",  min:0,      payout:0,    color:"#6B7A99" },
];
// Total Ascenso revenue across all accounts (sum of all customer amounts)
const ASCENSO_TOTAL = Object.values(ASCENSO_PROGRAMS).reduce((s,a)=>s+a.amount,0);

const AD_PCR_TIERS = [
  { label:"Tier 5", min:750, color:"#7C3AED" },
  { label:"Tier 4", min:400, color:"#DC2626" },
  { label:"Tier 3", min:250, color:"#D97706" },
  { label:"Tier 2", min:125, color:"#0891B2" },
  { label:"Tier 1", min:60,  color:"#059669" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];
const AD_TBR_TIERS = [
  { label:"Tier 5", min:750, color:"#7C3AED" },
  { label:"Tier 4", min:400, color:"#DC2626" },
  { label:"Tier 3", min:250, color:"#D97706" },
  { label:"Tier 2", min:125, color:"#0891B2" },
  { label:"Tier 1", min:60,  color:"#059669" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];
function getAdTier(units, tiers) {
  return tiers.find(t => units >= t.min) || tiers[tiers.length-1];
}
function getNextAdTier(units, tiers) {
  const tiersAsc = [...tiers].reverse();
  const nextIdx = tiersAsc.findIndex(t => units < t.min);
  return nextIdx >= 0 ? tiersAsc[nextIdx] : null;
}


// Americus Partners Program — as of 5/15/2026
// Tiers are unit-based: Entry<40 | Tier 1: 40 | Tier 2: 100 | Tier 3: 200+
const AMERICUS_PROGRAMS = {"200635":{"amerNum":10091,"enrollYear":2025,"primary":true,"units2025":465,"q1":106,"apr":36,"may":23,"jun":0,"q2":59,"ytd":165,"asOf":"5/29/2026"},"101161":{"amerNum":10083,"enrollYear":2024,"primary":true,"units2025":208,"q1":59,"apr":20,"may":5,"jun":0,"q2":25,"ytd":84,"asOf":"5/29/2026"},"101080":{"amerNum":10767,"enrollYear":2025,"primary":true,"units2025":136,"q1":8,"apr":0,"may":4,"jun":0,"q2":4,"ytd":12,"asOf":"5/29/2026"},"200628":{"amerNum":10168,"enrollYear":2025,"primary":true,"units2025":133,"q1":12,"apr":16,"may":4,"jun":0,"q2":20,"ytd":32,"asOf":"5/29/2026"},"200883":{"amerNum":10157,"enrollYear":2025,"primary":true,"units2025":329,"q1":46,"apr":13,"may":17,"jun":0,"q2":30,"ytd":76,"asOf":"5/29/2026"},"100551":{"amerNum":10175,"enrollYear":2025,"primary":true,"units2025":203,"q1":55,"apr":9,"may":9,"jun":0,"q2":18,"ytd":73,"asOf":"5/29/2026"},"101326":{"amerNum":10914,"enrollYear":2025,"primary":true,"units2025":143,"q1":26,"apr":9,"may":2,"jun":0,"q2":11,"ytd":37,"asOf":"5/29/2026"}};
const AMERICUS_TIERS = [
  { label:"Tier 3", min:200, color:"#7C3AED" },
  { label:"Tier 2", min:100, color:"#0891B2" },
  { label:"Tier 1", min:40,  color:"#059669" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];


// Bridgestone Firestone BARNN Program
// Tiers: placeholder — update when confirmed
const BARNN_PROGRAMS = {"200635":{"progNum":"969836","program":"Bridgestone Firestone BARNN","rep":"Larry Smith","role":"Primary","bs":3,"fs":4,"total":7,"priPct":100,"toNext":68,"asOf":"5/29/2026"},"101080":{"progNum":"900069","program":"Bridgestone Firestone BARNN","rep":"Larry Smith","role":"Secondary","bs":0,"fs":4,"total":4,"priPct":0,"toNext":71,"asOf":"5/29/2026"}};
const BARNN_TIERS = [
  { label:"Tier 10", min:750, color:"#7C3AED" },
  { label:"Tier 9",  min:675, color:"#6D28D9" },
  { label:"Tier 8",  min:600, color:"#DC2626" },
  { label:"Tier 7",  min:525, color:"#EA580C" },
  { label:"Tier 6",  min:450, color:"#D97706" },
  { label:"Tier 5",  min:375, color:"#0891B2" },
  { label:"Tier 4",  min:300, color:"#0369A1" },
  { label:"Tier 3",  min:225, color:"#059669" },
  { label:"Tier 2",  min:150, color:"#16A34A" },
  { label:"Tier 1",  min:75,  color:"#1E5FCC" },
  { label:"Entry",   min:0,   color:"#6B7A99" },
];


// Falken Fanatic PLT Program — as of 5/14/2026
// Tiers: Entry<30 | Tier 1:30 | Tier 2:60 | Tier 3:125 | Tier 4:250 | Tier 5:400 | Tier 6:600+
const FALKEN_PLT_PROGRAMS = {"100301":{"falkenId":30015,"q1":17,"q2":2,"ytd":2,"city":"Albany"},"101080":{"falkenId":19833,"q1":12,"q2":8,"ytd":8,"city":"Douglas"},"200807":{"falkenId":30946,"q1":8,"q2":4,"ytd":4,"city":"Cordele"},"101539":{"falkenId":39027,"q1":4,"q2":2,"ytd":2,"city":"Douglas"},"101103":{"falkenId":20852,"q1":150,"q2":132,"ytd":132,"city":"Tifton"},"101371":{"falkenId":30300,"q1":0,"q2":0,"ytd":0,"city":"Tifton"},"200327":{"falkenId":36783,"q1":6,"q2":0,"ytd":0,"city":"Valdosta"},"101436":{"falkenId":38153,"q1":0,"q2":8,"ytd":8,"city":"Sylvester"},"200891":{"falkenId":39881,"q1":10,"q2":8,"ytd":8,"city":"Sycamore"},"101512":{"falkenId":30314,"q1":0,"q2":9,"ytd":9,"city":"Ellenton"},"101323":{"falkenId":21528,"q1":42,"q2":24,"ytd":24,"city":"Tifton"},"200972":{"falkenId":30733,"q1":8,"q2":0,"ytd":0,"city":"Sylvester"},"200560":{"falkenId":38411,"q1":16,"q2":1,"ytd":1,"city":"Tifton"},"200266":{"falkenId":18255,"q1":27,"q2":16,"ytd":16,"city":"Valdosta"},"101513":{"falkenId":38475,"q1":4,"q2":0,"ytd":0,"city":"Moultrie"},"101322":{"falkenId":36793,"q1":6,"q2":11,"ytd":11,"city":"Tifton"},"101161":{"falkenId":28710,"q1":49,"q2":19,"ytd":19,"city":"Douglas"},"200631":{"falkenId":19578,"q1":-1,"q2":17,"ytd":17,"city":"Valdosta"},"100282":{"falkenId":38400,"q1":15,"q2":6,"ytd":6,"city":"Tifton"},"101463":{"falkenId":30957,"q1":9,"q2":11,"ytd":11,"city":"Fitzgerald"},"200220":{"falkenId":33722,"q1":21,"q2":16,"ytd":16,"city":"Sylvester"},"200628":{"falkenId":38456,"q1":6,"q2":12,"ytd":12,"city":"Moultrie"},"100551":{"falkenId":20125,"q1":22,"q2":4,"ytd":4,"city":"Tifton"},"101298":{"falkenId":31485,"q1":18,"q2":13,"ytd":13,"city":"Tifton"},"101415":{"falkenId":38439,"q1":23,"q2":8,"ytd":8,"city":"Nashville"},"200293":{"falkenId":29399,"q1":9,"q2":3,"ytd":3,"city":"Thomasville"},"101326":{"falkenId":30953,"q1":86,"q2":35,"ytd":35,"city":"Tifton"},"200294":{"falkenId":19568,"q1":9,"q2":2,"ytd":2,"city":"Valdosta"}};
const FALKEN_PLT_TIERS = [
  { label:"Tier 6", min:600, color:"#7C3AED" },
  { label:"Tier 5", min:400, color:"#DC2626" },
  { label:"Tier 4", min:250, color:"#D97706" },
  { label:"Tier 3", min:125, color:"#0891B2" },
  { label:"Tier 2", min:60,  color:"#059669" },
  { label:"Tier 1", min:30,  color:"#1E5FCC" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];


// Salesforce Call Log — uploaded 5/17/2026 (most recent entries)
const SEED_CALL_LOG = {"200166":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"spent some time with this strategic account to try to get the ship righted accessed inventory and there will be some returns but purchases as well to complete several sets","company":"DAVIS TIRE"}],"200895":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"visited with LB and Marcus was able to teach her about the TPMS scan tool finally, mission accomplished with a Davis Tire employee","company":"Tire Mart of Barnesville LLC"},{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dropped notepads attempted to get Mike trained on the scan tool but no luck Marcus was not in but I did let Angela know of looming price increase","company":"Tire Mart of Barnesville LLC"}],"2000043":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"new owner Cody reached out and wanted to pay his bill via ach Ireached out to Troy to have him send form","company":"CENTRAL AUTOMOTIVE LLC"}],"101691":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dropped notepads med truck dealer goal was to let Paula know of looming price increase","company":"JK MOBILE SERVICE LLC"}],"200884":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"BARNN prospect took Firestone guys by to go over the program with Jimmy Jimmy had to think about it a little more and get with me on the program","company":"TAYLOR CO. TIRE SHOP"}],"200897":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"med truck dealer dropped notepads reminded Danny of price increase coming on Friday afternoon","company":"187 TIRE & TRUCK REPAIR"}],"200482":[{"date":"5/17/2026","rep":"Nick Davis","type":"Phone Call","notes":"Leanne reached out for advice on which brand she should choose for a sale it was between Toyo and Falken I let her know she could not go wrong either way, but she is on the Falken program","company":"Watts Service Center Inc"}],"200446":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dropped handy dandy notepads reminded manager that to call me if pricing gets out of whack verse the market especially in regard to Toyo he stated pricing and inventory has been good","company":"MID GA SERVICE CENTER"}],"200362":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dealer still has not upgraded equipment and is still limited to what type of tires they can do but basically, they really are just not pushing tire sales at all","company":"BAR NONE AUTO"}],"200503":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dropped in with Firestone guys they reviewed all current promotions and dealers' numbers","company":"ALAN'S AUTOMOTIVE"}],"200583":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"dropped in with Firestone guys they attempted to log him into Treadnet but with no luck Mike and Bill will follow up with the assist","company":"YAUGHN TIRE"}],"101115":[{"date":"5/17/2026","rep":"Nick Davis","type":"Physical Visit","notes":"no Jay he was already gone for his Friday tee time did reach out to him with a friendly reminder of the looming price increase","company":"ECONOMY USED TIRE (WAREHOUSE)"}],"100165":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Physical Visit","notes":"Jason was not available to talk. They were very busy and said to come back.","company":"MOSLEY TIRE ALIGN.& BRAKE CTR."}],"200794":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Physical Visit","notes":"Visited Mr. Hayes with Jay Lee and Andrew from Ascenso. He is currently getting 2x delivery from gateway but he said he will start giving Ascenso a look. In the past he only wanted Galaxy Earth-Pro because they were tubeless. He was not aware that the Ascenso was available in tubeless.","company":"HAYES TIRE"},{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Joint call with Andrew - Ascenso and Austin Ballew. Went over new Ascenso program. Hayes wants to sign on TAG sooner than later","company":"HAYES TIRE"}],"200840":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Physical Visit","notes":"Visited Alex and Emma. My Gordy and Jonathen were out. Told them the history of the Ascenso brand and left the Associate dealer program for Mr. Gordy to turn in.","company":"Gordy Enterprises, Inc."},{"date":"5/15/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Met with the new owner Will, gave him the application, sent it in. Account is open and ready.","company":"STONE ENTERPRISES - FAUCETTS"}],"200406":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Physical Visit","notes":"Very good meeting with Chad and Jessica. They took Andrew on a tour of their facility and they also signed on the associate dealer program. We also sold them some irrigation tires.","company":"Appling Tire & Service Center"}],"200824":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Physical Visit","notes":"Visited with Jay and Andrew. We are converting him to the 12-16.5 L5 Ascenso from the Galaxy beefy baby. He also has the Associate dealer program info to return to Jay lee.","company":"Classic City Commercial"},{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Andrew and Austin went over Ascenso dealer program with Heath. Jerry got the call from Left Lane concerning PLT business that has left University.","company":"Classic City Commercial"},{"date":"5/15/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Inspected a Falken for possible adjustment. Heath and I agreed that it was an impact/foreign object. Scheduled Andrew/Ascenso for Wednesday.","company":"Classic City Commercial"}],"200132":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Phone Call","notes":"Sold Tim Wilks two Ascenso 600/65-20","company":"WILKS A-ONE TIRE SALES"}],"100301":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Phone Call","notes":"23.5R25 Titan STTR","company":"Albany General Tire Service Inc."}],"200266":[{"date":"5/17/2026","rep":"Austin Ballew","type":"Phone Call","notes":"Spoke to Byron and Bill a few times Friday. Coy had to have another operation. Surgery was successful and he is recovering.","company":"FUSSELL TIRE & SERVICE"}],"200885":[{"date":"5/17/2026","rep":"Tiffany Hilliard","type":"Physical Visit","notes":"Garret's last day was going to be Friday so I made sure to stop by and wish him farewell. He was the gatekeeper for Jay Lynn and made sure to make it damn near impossible to meet him years ago. So, I always made sure to say at ease soldier and salute him.","company":"Pierce Industrial Tire, LLC"}],"2000040":[{"date":"5/17/2026","rep":"Tiffany Hilliard","type":"Physical Visit","notes":"Got JD signed onto the Maxam Program. Even though D&K tried to get them to sign 1.5 months ago, they launched with us.","company":"ADVANCED TIRE SERVICE"}],"101530":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Check in with Bailey and Mr. Billy. Updated Toyo and Falken program numbers. OTR tread depth gauge, J Will is hitting them Tuesday. Americus ST road hazard. All is good. No complaints","company":"HARTLEY TIRE SERVICE"}],"200159":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Spoke to Michelle. Picked up a BH Ridgecrawler AT for road hazard. Updated Falken numbers. All good.","company":"Northside Tire Inc"}],"102049":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Spoke to Cj and Kyle (asst manager). They earned 4 Braves tickets as part of Toyo sellout for month of April. Tickets have been sent to Kyle. They love dealing with TDoG.","company":"RAFFIELD TIRE (RUSSELL)"},{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Let Leehim know that they had earned 4 tickets to Braves game for Toyo sellout for month of April. He was pumped. Tix have been delivered.","company":"RAFFIELD TIRE (MLK STORE)"}],"201003":[{"date":"5/16/2026","rep":"Curt Lane","type":"Phone Call","notes":"Todd reached out looking for 38x13.50r20 in a Toyo MT. None available, on back order with no ETA. Todd Doing well on Toyo and Falken programs.","company":"Thomas Tire Inc"}],"200603":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Updated Jeff and team on Toyo Driven and Falken Fanatic numbers. Looking good on both. Got a small PO for ST. Business has been solid. Multiple purchases a day.","company":"Perry Wholesale Tire Inc"}],"200798":[{"date":"5/16/2026","rep":"Curt Lane","type":"Phone Call","notes":"Spoke with Mr. Chris. Got a PO for a half a container load. Will be putting in a direct container order next week. Updated Toyo Driven. All is good.","company":"Raffield Tire Master Inc."}],"200940":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Not much positive here. Spoke to Andy and the inside sales team. Got an Airloc tube adjusted for Andy. Will keep popping in.","company":"PERRY BROS. OIL CO. INC."}],"200663":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Price level change completed. Showed Burger and Cody the tire return function. Updated Toyo Driven. Americus ST Road Hazard. We are go to distributor with NTW picking up scraps.","company":"BURGER'S TIRE CENTER & AUTO"}],"200386":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Justin called looking for trailer wheels. Told him how to look them up and he put order in for 4. Updated Toyo and Falken. Need to finish strong last month and a half.","company":"JUSTIN'S TIRE & AUTO - Eatonton"}],"102273":[{"date":"5/16/2026","rep":"Curt Lane","type":"Physical Visit","notes":"Spoke to Tim, updated Toyo Driven. All is good. We seem to be getting to his shop at a more consistent time, Tim likes.","company":"COMPLETE TIRE & SVC (#5)"}],"200495":[{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Lunch meet-up with David. He loves the TDOG family - went over Toyo, Yoko, and Falken numbers.","company":"Madison Car Care"},{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Firestone/Bridgestone call with Bill Munge. Went over all Q2 promos and Q1 earnings with David.","company":"Madison Car Care"}],"200736":[{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Daniel requested meeting about coming price increases. No deal killers but may need to shop some sensitive SKUs. Multiple competitors begging Daniel for more units.","company":"GLOBAL TIRES (SNELLVILLE)"}],"200297":[{"date":"5/16/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Went over price increase strategy with Jennifer. She will try to order 20k ahead of 5/15. Went over needed inventory at stores 2 & 3. All adjustments seem caught up.","company":"123 WHOLESALE TIRE"}],"101161":[{"date":"5/16/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Closed for a funeral. Will revisit.","company":"JMC Tire Inc."}],"101295":[{"date":"5/16/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Spoke with Jenny, tireco is trying to capture SOA in Douglas as well as Tifton and Moultrie.","company":"DAVID'S AUTO SALES / DOUGLAS"}],"101539":[{"date":"5/16/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Spoke with Austin, tireco trying to get all the business in Douglas.","company":"COURSON'S TIRE OF DOUGLAS"}],"100551":[{"date":"5/15/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Spoke with Randy, tire sales across the board are slow.","company":"Southside Tire & Auto Service"}],"200478":[{"date":"5/15/2026","rep":"Larry Smith","type":"Phone Call","notes":"Spoke with Roxanne, got a transfer misunderstanding solved.","company":"82 Tire & Lube"}],"2000052":[{"date":"5/15/2026","rep":"Jay Lee","type":"Physical Visit","notes":"Went over May promos with Josh. No issues. We are now their premier supplier.","company":"Moore Tire, LLC"}],"100282":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Spoke with Rudy, went over program goals and AR.","company":"RUDY'S TIRE SERVICE INC."},{"date":"5/11/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Checked in with Rudy on ST/TBR inventory and program status.","company":"RUDY'S TIRE SERVICE INC."}],"200891":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"With Jorge Herida, talked with Don about Falken TBR program. Jason wasn't available (off to the rodeo), left him the information and will revisit Tuesday.","company":"EJH Wrecker & Tire Service"}],"101103":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"With Jorge Herida, spoke with Amy and Juan about Falken TBR program. Left them the information.","company":"DEL TORO TIRE LLC"}],"100620":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"With Jorge Herida, spoke with Wil about the Falken TBR program. Left the information.","company":"ALLENS TIRE SERVICE"},{"date":"5/11/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Talked with Michael, checking inventory of ST, TBR. And program goals.","company":"ALLENS TIRE SERVICE"}],"101323":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"Spoke with Eric, cleared to sign up Allen's Tire on the Falken TBR program.","company":"ERIC'S TIRE SERVICE"}],"101326":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"With Jorge Herida, spoke with Steve and signed him up on the Falken TBR program.","company":"TIFTON GENERAL TIRE"}],"101463":[{"date":"5/14/2026","rep":"Larry Smith","type":"Physical Visit","notes":"With Jorge Herida from Falken TBR. Signed up Robbie on the Falken TBR program.","company":"SHELL RAPID LUBE (FITZGERALD)"}]};


// ── Authentication ────────────────────────────────────────────────────────────
const USERS = [
  { id:"tiffany", name:"Tiffany", color:"#7C3AED", hash:"ffc350e0cf3142060afedcbcc205ca749dcf6d78a82e29c36eecb8229352fe80", rep:true  },
  { id:"larry",   name:"Larry",   color:"#0891B2", hash:"03af02e864ff7b72aa5e1546d505089e06188ae3527c90af387c44aa748adec3", rep:true  },
  { id:"austin",  name:"Austin",  color:"#D97706", hash:"7c30f0708d4cd34a212c573bfafb8a41f56eb4ad9247ea50a638f9dff344183c", rep:true  },
  { id:"admin",   name:"Admin",   color:"#1E5FCC", hash:"428014c233df16655c4ce093a77fa4151788f71ab451eec009af91a9c1ba4001", rep:false },
];
async function hashPw(pw) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pw));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
}


// Falken Fanatic TBR Program — as of 5/22/2026
// Tiers: Entry<10 | Tier 1:10 | Tier 2:25 | Tier 3:50 | Tier 4:100 | Tier 5:200 | Tier 6:400+  (placeholder — update when confirmed)
const FALKEN_TBR_PROGRAMS = {"101161":{"falkenId":28710,"q1":0,"q2":30,"ytd":30,"city":"Douglas","segment":"TBR"},"101326":{"falkenId":30953,"q1":0,"q2":6,"ytd":6,"city":"Tifton","segment":"TBR"},"101463":{"falkenId":30957,"q1":0,"q2":4,"ytd":4,"city":"Fitzgerald","segment":"TBR"},"200266":{"falkenId":18255,"q1":0,"q2":0,"ytd":0,"city":"Valdosta","segment":"TBR"},"200595":{"falkenId":19331,"q1":0,"q2":10,"ytd":10,"city":"Mayo","segment":"TBR"}};
const FALKEN_TBR_TIERS = [
  { label:"Tier 6", min:400, color:"#7C3AED" },
  { label:"Tier 5", min:200, color:"#DC2626" },
  { label:"Tier 4", min:100, color:"#D97706" },
  { label:"Tier 3", min:50,  color:"#0891B2" },
  { label:"Tier 2", min:25,  color:"#059669" },
  { label:"Tier 1", min:10,  color:"#1E5FCC" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];


// Austin Statesboro territory — SOS all depts, all others Ag/OTR/Industrial only
// Source: W1-W21 2026 YTD as of 5/22/2026
const STATESBORO_DATA = {"3000374":{"name":"SNIDER FLEET SOLUTIONS","total":1781.26,"depts":{"INDUSTRIAL TIRES":{"qty":2.0,"amount":429.44,"gp":55.93},"OFF THE ROAD TIRES":{"qty":2.0,"amount":1351.82,"gp":455.76}}},"3000104":{"name":"SOS TIRE & AUTO","total":1104661.1,"depts":{"ST TRAILER":{"qty":353.0,"amount":28012.32,"gp":3291.89},"RAD LT TRUCK":{"qty":711.0,"amount":68963.25,"gp":9825.66},"TRUCK TIRES":{"qty":4058.0,"amount":957721.29,"gp":27249.69},"OFF THE ROAD TIRES":{"qty":39.0,"amount":40999.18,"gp":995.44},"PASSENGER TIRES":{"qty":95.0,"amount":19882.94,"gp":808.04},"FARM TIRES":{"qty":4.0,"amount":3328.85,"gp":327.17},"INDUSTRIAL TIRES":{"qty":13.0,"amount":7148.45,"gp":145.81},"LAWN & GARDEN":{"qty":8.0,"amount":442.2,"gp":72.32},"PATCHES AND REPAIR":{"qty":5.0,"amount":15.0,"gp":12.35},"VALVE STEMS":{"qty":165.0,"amount":2447.15,"gp":1037.42},"WHEEL WEIGHTS":{"qty":72.0,"amount":2902.0,"gp":1351.43}}},"101363":{"name":"TERWILLIGER TIRE & AG","total":39768.61,"depts":{"FARM TIRES":{"qty":77.0,"amount":21476.16,"gp":2503.44},"INDUSTRIAL TIRES":{"qty":24.0,"amount":5543.48,"gp":650.24},"OFF THE ROAD TIRES":{"qty":7.0,"amount":12748.97,"gp":1303.89}}},"101995":{"name":"APPLING TIRE & SERVICE CENTER","total":23834.3,"depts":{"FARM TIRES":{"qty":32.0,"amount":13442.9,"gp":1450.09},"OFF THE ROAD TIRES":{"qty":3.0,"amount":9685.0,"gp":745.53},"INDUSTRIAL TIRES":{"qty":4.0,"amount":706.4,"gp":116.36}}},"3000719":{"name":"BCS AUTO TRUCK & TIRE SERVICE","total":4393.5,"depts":{"FARM TIRES":{"qty":20.0,"amount":4393.5,"gp":598.12}}},"3000714":{"name":"BROWER TIRE & OIL","total":14769.42,"depts":{"FARM TIRES":{"qty":61.0,"amount":12470.01,"gp":1248.42},"OFF THE ROAD TIRES":{"qty":6.0,"amount":2299.41,"gp":194.89}}},"100332":{"name":"COLLINS AND SONS TIRE","total":38137.1,"depts":{"OFF THE ROAD TIRES":{"qty":7.0,"amount":5091.39,"gp":366.4},"INDUSTRIAL TIRES":{"qty":53.0,"amount":19441.34,"gp":2187.08},"FARM TIRES":{"qty":67.0,"amount":13604.37,"gp":1611.72}}},"3000274":{"name":"FORKLIFT TIRE SPECIALIST LLC","total":91710.86,"depts":{"INDUSTRIAL TIRES":{"qty":473.0,"amount":91710.86,"gp":12568.74}}},"101186":{"name":"GLENNVILLE AUTO PARTS","total":5805.12,"depts":{"FARM TIRES":{"qty":14.0,"amount":5805.12,"gp":1145.63}}},"101278":{"name":"GORDY ENTERPRISES INC","total":10757.23,"depts":{"FARM TIRES":{"qty":50.0,"amount":14624.24,"gp":1759.09},"INDUSTRIAL TIRES":{"qty":6.0,"amount":1085.87,"gp":185.09}}},"101267":{"name":"JESUP TIRE SHOP","total":13184.31,"depts":{"INDUSTRIAL TIRES":{"qty":3.0,"amount":407.73,"gp":62.73},"FARM TIRES":{"qty":17.0,"amount":3430.1,"gp":329.14},"OFF THE ROAD TIRES":{"qty":16.0,"amount":9346.48,"gp":1043.32}}},"100489":{"name":"LUMBER CITY ENT/BURKETT TIRE","total":68559.74,"depts":{"OFF THE ROAD TIRES":{"qty":27.0,"amount":65294.5,"gp":1829.73},"FARM TIRES":{"qty":6.0,"amount":3265.24,"gp":318.66}}},"100166":{"name":"MASSEY OIL","total":13622.48,"depts":{"FARM TIRES":{"qty":19.0,"amount":7985.36,"gp":1363.27},"INDUSTRIAL TIRES":{"qty":15.0,"amount":5637.12,"gp":790.93}}},"3000306":{"name":"MOODY TIRE SERVICE INC.","total":375.43,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":375.43,"gp":61.1}}},"101088":{"name":"MS GARAGE","total":1518.54,"depts":{"OFF THE ROAD TIRES":{"qty":2.0,"amount":834.54,"gp":59.32},"FARM TIRES":{"qty":4.0,"amount":503.16,"gp":85.32},"INDUSTRIAL TIRES":{"qty":1.0,"amount":180.84,"gp":39.69}}},"101231":{"name":"MS TIRE","total":24991.17,"depts":{"FARM TIRES":{"qty":38.0,"amount":24009.72,"gp":2330.08},"INDUSTRIAL TIRES":{"qty":3.0,"amount":981.45,"gp":180.24}}},"101212":{"name":"RANDYS WRECKER AND SERVICE","total":3549.32,"depts":{"FARM TIRES":{"qty":17.0,"amount":3265.58,"gp":541.08},"INDUSTRIAL TIRES":{"qty":2.0,"amount":283.74,"gp":55.0}}},"3000048":{"name":"SCREENS TIRE & MUFFLER","total":308.23,"depts":{"INDUSTRIAL TIRES":{"qty":3.0,"amount":233.68,"gp":35.38},"FARM TIRES":{"qty":1.0,"amount":74.55,"gp":8.8}}},"301222":{"name":"TATTNALL TIRE & BRAKE","total":3551.62,"depts":{"INDUSTRIAL TIRES":{"qty":2.0,"amount":489.17,"gp":53.34},"FARM TIRES":{"qty":22.0,"amount":3062.45,"gp":496.65}}},"100208":{"name":"WILLIAMS TIRE","total":22872.28,"depts":{"INDUSTRIAL TIRES":{"qty":7.0,"amount":1105.25,"gp":192.23},"OFF THE ROAD TIRES":{"qty":9.0,"amount":5124.91,"gp":846.5},"FARM TIRES":{"qty":78.0,"amount":16642.12,"gp":2346.78}}},"3000105":{"name":"ZEAGLER TIRE COMPANY","total":4519.63,"depts":{"OFF THE ROAD TIRES":{"qty":3.0,"amount":4519.63,"gp":889.25}}},"3000005":{"name":"CATES SALES & SERVICE INC.","total":3081.29,"depts":{"OFF THE ROAD TIRES":{"qty":1.0,"amount":1741.74,"gp":568.32},"FARM TIRES":{"qty":9.0,"amount":1339.55,"gp":240.33}}},"3000022":{"name":"DR. TIRE INC.","total":48085.1,"depts":{"INDUSTRIAL TIRES":{"qty":11.0,"amount":3931.16,"gp":476.34},"FARM TIRES":{"qty":50.0,"amount":23686.84,"gp":1796.59},"OFF THE ROAD TIRES":{"qty":10.0,"amount":20467.1,"gp":488.26}}},"101395":{"name":"K & L TIRE AND ALIGNMENT LLC","total":2207.22,"depts":{"FARM TIRES":{"qty":15.0,"amount":2207.22,"gp":402.72}}},"101242":{"name":"KREIDER REPAIR","total":632.26,"depts":{"FARM TIRES":{"qty":3.0,"amount":632.26,"gp":40.68}}},"100165":{"name":"MOSLEY TIRE ALIGN.& BRAKE CTR.","total":3037.01,"depts":{"INDUSTRIAL TIRES":{"qty":4.0,"amount":842.66,"gp":24.84},"OFF THE ROAD TIRES":{"qty":3.0,"amount":2613.0,"gp":378.33}}},"3000117":{"name":"SCREVEN TIRE","total":1277.82,"depts":{"FARM TIRES":{"qty":9.0,"amount":1132.49,"gp":197.65},"INDUSTRIAL TIRES":{"qty":1.0,"amount":145.33,"gp":7.42}}},"3000182":{"name":"SHEFFIELD OIL COMPANY INC","total":3730.47,"depts":{"FARM TIRES":{"qty":17.0,"amount":3330.37,"gp":461.34},"INDUSTRIAL TIRES":{"qty":2.0,"amount":400.1,"gp":51.82}}},"3000023":{"name":"SOUTHSIDE WRENS TIRE & BRAKE","total":1188.74,"depts":{"FARM TIRES":{"qty":4.0,"amount":1188.74,"gp":169.9}}},"101849":{"name":"TOTAL TIRE & AUTO CARE","total":15792.24,"depts":{"INDUSTRIAL TIRES":{"qty":14.0,"amount":3324.01,"gp":590.38},"OFF THE ROAD TIRES":{"qty":4.0,"amount":3498.66,"gp":470.94},"FARM TIRES":{"qty":38.0,"amount":8969.57,"gp":1321.81}}},"3000268":{"name":"BOOM BOOMS TIRE SHOP","total":15354.03,"depts":{"FARM TIRES":{"qty":18.0,"amount":9321.26,"gp":1263.11},"OFF THE ROAD TIRES":{"qty":2.0,"amount":2616.8,"gp":297.01},"INDUSTRIAL TIRES":{"qty":3.0,"amount":3415.97,"gp":382.86}}},"3000463":{"name":"FOUR HOLES SALES & SERVICE","total":3444.07,"depts":{"INDUSTRIAL TIRES":{"qty":13.0,"amount":2600.65,"gp":374.67},"FARM TIRES":{"qty":6.0,"amount":843.42,"gp":114.12}}},"201061":{"name":"HEAVY DUTY TIRE","total":3165.01,"depts":{"OFF THE ROAD TIRES":{"qty":3.0,"amount":1680.55,"gp":386.21},"INDUSTRIAL TIRES":{"qty":5.0,"amount":1484.46,"gp":101.92}}},"3000312":{"name":"SNIDER FLEET SOLUTIONS","total":5710.96,"depts":{"OFF THE ROAD TIRES":{"qty":3.0,"amount":6628.5,"gp":327.5}}},"3000121":{"name":"THOMSON TIRE CO.","total":562.74,"depts":{"OFF THE ROAD TIRES":{"qty":1.0,"amount":562.74,"gp":95.16}}},"3000562":{"name":"CHESTER SERVICES","total":272.5,"depts":{"INDUSTRIAL TIRES":{"qty":2.0,"amount":272.5,"gp":72.96}}},"3000732":{"name":"HEAVY DUTY TIRES","total":4303.7,"depts":{"INDUSTRIAL TIRES":{"qty":9.0,"amount":2176.3,"gp":208.48},"FARM TIRES":{"qty":11.0,"amount":2127.4,"gp":208.89}}},"3000025":{"name":"JASONS STATION","total":1339.62,"depts":{"FARM TIRES":{"qty":13.0,"amount":1339.62,"gp":168.48}}},"3000270":{"name":"JIM WHITEHEADS TIRE & SERVICE","total":7169.22,"depts":{"OFF THE ROAD TIRES":{"qty":3.0,"amount":2695.32,"gp":349.05},"INDUSTRIAL TIRES":{"qty":24.0,"amount":4473.9,"gp":777.44}}},"100262":{"name":"QUALITY TIRE CO.","total":3250.95,"depts":{"FARM TIRES":{"qty":10.0,"amount":3074.35,"gp":352.67},"INDUSTRIAL TIRES":{"qty":1.0,"amount":176.6,"gp":29.09}}},"101455":{"name":"AMERSON'S SALVAGE & USED CARS","total":1145.33,"depts":{"OFF THE ROAD TIRES":{"qty":1.0,"amount":1145.33,"gp":222.37}}},"101202":{"name":"BLOCKERS STATION","total":2011.25,"depts":{"FARM TIRES":{"qty":13.0,"amount":2011.25,"gp":306.02}}},"3000278":{"name":"KENS TIRE LLC","total":1405.97,"depts":{"FARM TIRES":{"qty":4.0,"amount":1405.97,"gp":175.36}}},"3000046":{"name":"R&R MOTOR COMPANY INC.","total":955.01,"depts":{"FARM TIRES":{"qty":5.0,"amount":559.78,"gp":121.07},"INDUSTRIAL TIRES":{"qty":1.0,"amount":395.23,"gp":61.0}}},"3000728":{"name":"RTS COMMERCIAL TIRES","total":1350.51,"depts":{"INDUSTRIAL TIRES":{"qty":2.0,"amount":318.8,"gp":33.38},"FARM TIRES":{"qty":2.0,"amount":628.04,"gp":108.52},"OFF THE ROAD TIRES":{"qty":1.0,"amount":403.67,"gp":43.72}}},"3000249":{"name":"STAR GARAGE INC.","total":229.19,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":229.19,"gp":30.41}}},"100224":{"name":"WARDS SERVICE CENTER","total":384.88,"depts":{"FARM TIRES":{"qty":4.0,"amount":384.88,"gp":131.0}}},"3000402":{"name":"GLENVILLE TIRE","total":1725.24,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":118.6,"gp":24.5},"FARM TIRES":{"qty":4.0,"amount":1606.64,"gp":190.54}}},"3000181":{"name":"GOODYEAR TIRE & SERVICE CENTER","total":229.34,"depts":{"FARM TIRES":{"qty":2.0,"amount":229.34,"gp":22.94}}},"101393":{"name":"HADDENS AUTO REPAIR","total":23318.28,"depts":{"INDUSTRIAL TIRES":{"qty":36.0,"amount":23318.28,"gp":1722.2}}},"3000026":{"name":"HWY 25 EXPRESS LLC","total":3070.04,"depts":{"FARM TIRES":{"qty":3.0,"amount":3070.04,"gp":211.21}}},"3000376":{"name":"PRIORITY TIRE (SBORO WHSE)","total":2716.99,"depts":{"INDUSTRIAL TIRES":{"qty":8.0,"amount":2638.72,"gp":291.68},"FARM TIRES":{"qty":1.0,"amount":78.27,"gp":13.24}}},"3000357":{"name":"SIMPLE TIRE - STATESBORO","total":525.66,"depts":{"INDUSTRIAL TIRES":{"qty":4.0,"amount":525.66,"gp":54.0}}},"100619":{"name":"WAYNE CO. TIRE & SERVICE INC","total":94.22,"depts":{"FARM TIRES":{"qty":1.0,"amount":94.22,"gp":19.21}}},"102367":{"name":"A.S.A.P MOBILE TIRE SALES","total":722.21,"depts":{"FARM TIRES":{"qty":4.0,"amount":722.21,"gp":91.37}}},"3000332":{"name":"BLACKS TIRE SERVICE INC.","total":930.0,"depts":{"FARM TIRES":{"qty":6.0,"amount":930.0,"gp":159.36}}},"101793":{"name":"COTTONS TIRE","total":83.76,"depts":{"FARM TIRES":{"qty":1.0,"amount":83.76,"gp":7.87}}},"3000453":{"name":"MORRELL TIRE SERVICE INC.","total":16679.47,"depts":{"FARM TIRES":{"qty":77.0,"amount":16679.47,"gp":1654.06}}},"3000706":{"name":"SOUTHERN TIRE MART #150","total":2095.99,"depts":{"FARM TIRES":{"qty":4.0,"amount":1612.87,"gp":205.98},"INDUSTRIAL TIRES":{"qty":3.0,"amount":483.12,"gp":39.41}}},"101344":{"name":"THE TIRE KING OF STATESBORO","total":120.97,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":120.97,"gp":26.73}}},"3000377":{"name":"TIRESEASY-LLC (SBORO WHSE)","total":3104.0,"depts":{"FARM TIRES":{"qty":8.0,"amount":3104.0,"gp":337.2}}},"3000109":{"name":"CLEMENS TRUCK REPAIR","total":290.82,"depts":{"FARM TIRES":{"qty":2.0,"amount":290.82,"gp":46.0}}},"101185":{"name":"LAKES ALIGNMENT SERVICE","total":1904.1,"depts":{"INDUSTRIAL TIRES":{"qty":8.0,"amount":1904.1,"gp":50.52}}},"3000718":{"name":"DAVIS TIRE (SPRINGFIELD)","total":60.42,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":60.42,"gp":8.39}}},"101378":{"name":"GRADYS TRUCK STOP","total":522.49,"depts":{"FARM TIRES":{"qty":3.0,"amount":522.49,"gp":66.62}}},"3000107":{"name":"JOHNSON TIRE & AUTO","total":93.58,"depts":{"FARM TIRES":{"qty":1.0,"amount":93.58,"gp":17.24}}},"100412":{"name":"PERFORMANCE TIRE & AUTO CARE","total":77.73,"depts":{"FARM TIRES":{"qty":1.0,"amount":77.73,"gp":6.66}}},"3000342":{"name":"STILSON TIRE CONNECTION","total":85.91,"depts":{"FARM TIRES":{"qty":1.0,"amount":85.91,"gp":9.99}}},"3000372":{"name":"CATES EXPRESS LLC","total":1357.96,"depts":{"FARM TIRES":{"qty":4.0,"amount":1357.96,"gp":144.4}}},"3000459":{"name":"SAVANNAH GLOBAL SOLUTIONS","total":4390.54,"depts":{"OFF THE ROAD TIRES":{"qty":2.0,"amount":4390.54,"gp":351.24}}},"3000198":{"name":"BROTHERS TIRE & SERVICE LLC","total":285.01,"depts":{"FARM TIRES":{"qty":2.0,"amount":142.81,"gp":12.01},"INDUSTRIAL TIRES":{"qty":1.0,"amount":142.2,"gp":4.29}}},"3000013":{"name":"CLARKS AUTO REPAIR","total":361.68,"depts":{"INDUSTRIAL TIRES":{"qty":2.0,"amount":361.68,"gp":79.38}}},"3000171":{"name":"MCCORKLE SALES INC.","total":203.74,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":203.74,"gp":30.72}}},"3000561":{"name":"ADAMS AUTO LLC","total":72.54,"depts":{"FARM TIRES":{"qty":1.0,"amount":72.54,"gp":13.15}}},"3000429":{"name":"HOLLAND TIRE","total":848.26,"depts":{"FARM TIRES":{"qty":6.0,"amount":848.26,"gp":117.8}}},"3000053":{"name":"JK&T TIRE & AUTO","total":60.42,"depts":{"INDUSTRIAL TIRES":{"qty":1.0,"amount":60.42,"gp":8.39}}},"3000021":{"name":"LANE TRACTOR & AUTO LLC","total":290.82,"depts":{"FARM TIRES":{"qty":2.0,"amount":290.82,"gp":46.0}}},"101237":{"name":"SOUTHEAST TIRE AND SERVICE","total":468.28,"depts":{"FARM TIRES":{"qty":5.0,"amount":468.28,"gp":62.67}}},"3000447":{"name":"MCCARTHY TIRE SERVICE COMPANY","total":2774.52,"depts":{"OFF THE ROAD TIRES":{"qty":3.0,"amount":2774.52,"gp":180.03}}},"101377":{"name":"BUTLER TIRE AND LUBE","total":83.76,"depts":{"FARM TIRES":{"qty":1.0,"amount":83.76,"gp":7.78}}},"301223":{"name":"JAMES TIRE 2","total":82.01,"depts":{"FARM TIRES":{"qty":1.0,"amount":82.01,"gp":4.08}}},"3000442":{"name":"SMITH BROS. TIRE SERVICE INC","total":1769.78,"depts":{"OFF THE ROAD TIRES":{"qty":2.0,"amount":1769.78,"gp":218.58}}},"101679":{"name":"JAM SALES INC","total":243.4,"depts":{"FARM TIRES":{"qty":2.0,"amount":243.4,"gp":63.18}}},"102273":{"name":"COMPLETE TIRE & SVC (CORDELE)","total":213.84,"depts":{"INDUSTRIAL TIRES":{"qty":4.0,"amount":213.84,"gp":19.64}}},"3000027":{"name":"SOUTH GA COMMERCIAL LUBE LLC","total":209.86,"depts":{"FARM TIRES":{"qty":1.0,"amount":209.86,"gp":35.8}}},"3000426":{"name":"GOODYEAR COMMERCIAL TIRE & SVC","total":239.5,"depts":{"FARM TIRES":{"qty":1.0,"amount":239.5,"gp":23.95}}},"3000281":{"name":"DORSEY AUTO CARE","total":178.62,"depts":{"INDUSTRIAL TIRES":{"qty":3.0,"amount":178.62,"gp":22.53}}},"101290":{"name":"PETES NEW AND USED TIRES","total":60.0,"depts":{"FARM TIRES":{"qty":1.0,"amount":60.0,"gp":6.73}}}};


// Yokohama Dealer Program — as of 5/31/2026
// Tiers: placeholder — update when official thresholds confirmed
const YOKOHAMA_PROGRAMS = {"100417":{"program":"Yokohama Dealer Program","progNum":"14112","rep":"Larry Smith","primary":32,"secondary":10,"ytd":42,"priPct":76,"toNext":18,"asOf":"5/31/2026"},"200266":{"program":"Yokohama Dealer Program","progNum":"22433","rep":"Austin Ballew","primary":0,"secondary":10,"ytd":10,"priPct":0,"toNext":50,"asOf":"5/31/2026"}};
const YOKOHAMA_TIERS = [
  { label:"Tier 5", min:200, color:"#7C3AED" },
  { label:"Tier 4", min:100, color:"#DC2626" },
  { label:"Tier 3", min:60,  color:"#D97706" },
  { label:"Tier 2", min:30,  color:"#0891B2" },
  { label:"Tier 1", min:10,  color:"#059669" },
  { label:"Entry",  min:0,   color:"#6B7A99" },
];


// ── Supabase Sync ─────────────────────────────────────────────────────────────
const SB_URL = "https://jcdajvjvengtbjdrbrsp.supabase.co";
const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpjZGFqdmp2ZW5ndGJqZHJicnNwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAyNjAzMzAsImV4cCI6MjA5NTgzNjMzMH0.OAzY4yuYzv-9D5UwwQW5Aqw6QpcVWHk6o_4_eecdx3k";

async function sbFetch(path, method="GET", body=null) {
  try {
    const headers = {
      "apikey": SB_KEY,
      "Authorization": `Bearer ${SB_KEY}`,
      "Content-Type": "application/json",
    };
    // Only add Prefer header for writes — never for GET (it suppresses response body)
    if (method === "POST")   headers["Prefer"] = "resolution=merge-duplicates,return=minimal";
    if (method === "PATCH")  headers["Prefer"] = "return=minimal";
    const opts = { method, headers };
    if (body) opts.body = JSON.stringify(body);
    const res = await fetch(`${SB_URL}/rest/v1/${path}`, opts);
    if (!res.ok) {
      const err = await res.text();
      console.warn(`[Supabase] ${method} ${path} → ${res.status}:`, err);
      return null;
    }
    const txt = await res.text();
    return txt ? JSON.parse(txt) : null;
  } catch(e) {
    console.warn("[Supabase] fetch error:", e.message);
    return null;
  }
}

// Notes sync
async function syncNotesUp(userId, custNum, notes) {
  // Use upsert with explicit on_conflict parameter
  await sbFetch("rep_notes?on_conflict=user_id,cust_num", "POST", {
    user_id: userId, cust_num: String(custNum), notes, updated_at: new Date().toISOString()
  });
}
async function syncNotesDown(userId, custNum) {
  const rows = await sbFetch(`rep_notes?user_id=eq.${encodeURIComponent(userId)}&cust_num=eq.${encodeURIComponent(custNum)}&select=notes,updated_at`);
  return rows?.[0] || null;
}

// Todos sync
async function syncTodosUp(userId, custNum, custName, city, salesman, todos) {
  // Delete all existing todos for this customer/user then re-insert
  await sbFetch(`rep_todos?user_id=eq.${encodeURIComponent(userId)}&cust_num=eq.${encodeURIComponent(custNum)}`, "DELETE");
  if (todos.length === 0) return;
  const rows = todos.map(t => ({
    user_id: userId, cust_num: String(custNum),
    cust_name: custName, city, salesman,
    text: t.text, done: t.done,
    created_date: t.date, created_by: t.by,
  }));
  await sbFetch("rep_todos", "POST", rows);
}
async function syncTodosDown(userId, custNum) {
  const rows = await sbFetch(`rep_todos?user_id=eq.${encodeURIComponent(userId)}&cust_num=eq.${encodeURIComponent(custNum)}&order=created_at.asc`);
  if (!rows) return null;
  return rows.map(r => ({ id: r.id, text: r.text, done: r.done, date: r.created_date, by: r.created_by }));
}
async function syncAllTodosDown(userId) {
  const rows = await sbFetch(`rep_todos?user_id=eq.${encodeURIComponent(userId)}&order=created_at.asc`);
  return rows || null;
}

// Activity log sync — write to Supabase so admin sees all reps
async function syncActivityUp(entry) {
  // Silently skip if table doesn't exist yet
  try {
    await sbFetch("rep_activity", "POST", {
      ts: entry.ts, user_name: entry.user, user_id: entry.userId,
      action: entry.action, detail: entry.detail,
      updated_at: new Date().toISOString(),
    });
  } catch {}
}
async function syncActivityDown() {
  const rows = await sbFetch("rep_activity?order=ts.desc&limit=500");
  return rows || [];
}

// Leads sync
async function syncLeadUp(lead) {
  // Only send columns that exist in rep_leads table
  const row = {
    id:               lead.id,
    name:             lead.name,
    city:             lead.city || null,
    phone:            lead.phone || null,
    business_type:    lead.businessType || lead.business_type || null,
    notes:            lead.notes || null,
    status:           lead.status || "open",
    assigned_to:      lead.assigned_to || null,
    assigned_to_name: lead.assigned_to_name || null,
    created_by:       lead.created_by || null,
    created_by_name:  lead.created_by_name || null,
    created_at:       lead.created_at || new Date().toISOString(),
    updated_at:       new Date().toISOString(),
  };
  await sbFetch("rep_leads?on_conflict=id", "POST", row);
}
async function syncLeadsDown(userId) {
  // Fetch leads assigned to this user (by name match — assigned_to stores lowercase name)
  const [assigned, created] = await Promise.all([
    sbFetch(`rep_leads?assigned_to=eq.${encodeURIComponent(userId)}&order=created_at.desc`),
    sbFetch(`rep_leads?created_by=eq.${encodeURIComponent(userId)}&order=created_at.desc`),
  ]);
  // Merge and dedupe
  const all = [...(assigned||[]), ...(created||[])];
  const seen = new Set();
  return all.filter(r => { if(seen.has(r.id)) return false; seen.add(r.id); return true; });
}
async function syncAllLeadsDown() {
  // Admin only — fetch ALL leads
  const rows = await sbFetch(`rep_leads?order=created_at.desc&limit=500`);
  return rows || [];
}
async function updateLeadUp(leadId, updates) {
  await sbFetch(`rep_leads?id=eq.${leadId}`, "PATCH", { ...updates, updated_at: new Date().toISOString() });
}
async function deleteLeadUp(leadId) {
  await sbFetch(`rep_leads?id=eq.${leadId}`, "DELETE");
}

// AI plan sync
async function syncAIPlanUp(userId, plan, generatedAt) {
  await sbFetch("rep_ai_plans?on_conflict=user_id", "POST", {
    user_id: userId, plan, generated_at: generatedAt, updated_at: new Date().toISOString()
  });
}
async function syncAIPlanDown(userId) {
  const rows = await sbFetch(`rep_ai_plans?user_id=eq.${encodeURIComponent(userId)}&select=plan,generated_at`);
  return rows?.[0] || null;
}


// ── Master Upload Workbook Parser ─────────────────────────────────────────────
function isMasterWorkbook(wb) {
  const sheets = wb.SheetNames;
  return sheets.includes("AR") && sheets.includes("CustomerComp") && sheets.includes("WTD");
}

function getSheetRows(wb, name, startRow = 1) {
  if (!wb.Sheets[name]) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null })
    .slice(startRow);
}

function isGAEPD(name) {
  return /GA.EPD|GAEPA/i.test(String(name || ""));
}

// ── AR sheet ─────────────────────────────────────────────────────────────────
function parseMasterAR(wb) {
  const rows = getSheetRows(wb, "AR", 2); // skip title + note rows
  const ar = [];
  for (const row of rows) {
    if (!row[1] || typeof row[1] !== "number") continue;
    let lastPaid = "";
    try {
      const d = row[12] instanceof Date ? row[12] : new Date(row[12]);
      if (!isNaN(d)) lastPaid = d.toISOString().slice(0, 10);
    } catch {}
    ar.push({
      salesman:  String(row[0] || "").trim(),
      custNum:   row[1],
      shortName: String(row[2] || "").trim(),
      name:      String(row[3] || "").trim(),
      phone:     String(row[4] || "").trim(),
      balance:   Number(row[5] || 0),
      futDue:    Number(row[6] || 0),
      curDue:    Number(row[7] || 0),
      due1_30:   Number(row[8] || 0),
      due31_60:  Number(row[9] || 0),
      due61_90:  Number(row[10] || 0),
      dueOver90: Number(row[11] || 0),
      lastPaid,
    });
  }
  return ar;
}

// ── CustomerComp sheet (HITS side-by-side 2025 A-I, 2026 K-S) ─────────────
function parseMasterCustomerComp(wb) {
  const rows = getSheetRows(wb, "CustomerComp", 4); // skip title/note/year-group/col-headers
  // Build per-customer YTD aggregation
  const by25 = {}, by26 = {};
  for (const row of rows) {
    const c25 = row[0], name25 = row[1], dept25 = row[2];
    const amt25 = Number(row[4] || 0), gp25 = Number(row[5] || 0);
    const c26 = row[10], name26 = row[11], dept26 = row[12];
    const amt26 = Number(row[14] || 0), gp26 = Number(row[15] || 0);

    if (c25 && typeof c25 === "number" && !isGAEPD(dept25)) {
      const k = String(c25);
      if (!by25[k]) by25[k] = { custNum: c25, customer: String(name25 || "").trim(), sales: 0, gp: 0, depts: {} };
      by25[k].sales += amt25;
      by25[k].gp += gp25;
      if (dept25) by25[k].depts[dept25] = (by25[k].depts[dept25] || 0) + amt25;
    }
    if (c26 && typeof c26 === "number" && !isGAEPD(dept26)) {
      const k = String(c26);
      if (!by26[k]) by26[k] = { custNum: c26, customer: String(name26 || "").trim(), sales: 0, gp: 0, depts: {} };
      by26[k].sales += amt26;
      by26[k].gp += gp26;
      if (dept26) by26[k].depts[dept26] = (by26[k].depts[dept26] || 0) + amt26;
    }
  }

  // Build action plan entries
  const allKeys = new Set([...Object.keys(by25), ...Object.keys(by26)]);
  const actionPlan = [];
  for (const k of allKeys) {
    const d25 = by25[k] || { sales: 0, gp: 0, depts: {} };
    const d26 = by26[k] || { sales: 0, gp: 0, depts: {} };
    const rec  = by26[k] || by25[k];
    const s25  = d25.sales, s26 = d26.sales;
    const change = s26 - s25;
    const gpPct  = s26 > 0 ? d26.gp / s26 : 0;

    // Determine top dept 2026
    const depts26 = d26.depts;
    const topDept = Object.keys(depts26).sort((a, b) => depts26[b] - depts26[a])[0] || "";

    // Most declined dept (in 2025 but less/missing in 2026)
    let declinedDept = "";
    let maxDecline = 0;
    for (const dept of Object.keys(d25.depts)) {
      const decline = (d25.depts[dept] || 0) - (d26.depts[dept] || 0);
      if (decline > maxDecline) { maxDecline = decline; declinedDept = dept; }
    }

    const action = change > 500 ? "GROW" : change < -500 ? "LOST" : Math.abs(change) < 100 ? "OK" : "WATCH";

    actionPlan.push({
      custNum:      Number(k),
      customer:     rec.customer,
      city:         "",       // not in comp report — will be filled from customer list
      salesman:     "House",  // default — will be overwritten by customer list match
      sales2025:    Math.round(s25 * 100) / 100,
      sales2026:    Math.round(s26 * 100) / 100,
      change:       Math.round(change * 100) / 100,
      gpPct:        Math.round(gpPct * 10000) / 10000,
      action,
      topDept,
      declinedDept,
      focus:        topDept,
    });
  }

  actionPlan.sort((a, b) => b.sales2026 - a.sales2026);
  return { actionPlan };
}

// ── Statesboro sheet (same format as CustomerComp) ──────────────────────────
function parseMasterStatesboro(wb) {
  if (!wb.Sheets["Statesboro"]) return { actionPlan: [] };
  const rows = getSheetRows(wb, "Statesboro", 4);
  const by26 = {}, by25 = {};
  for (const row of rows) {
    const c25 = row[0], dept25 = row[2], amt25 = Number(row[4] || 0);
    const c26 = row[10], dept26 = row[12], amt26 = Number(row[14] || 0), gp26 = Number(row[15] || 0);
    const name26 = row[11];
    if (c25 && typeof c25 === "number" && !isGAEPD(dept25)) {
      const k = String(c25);
      if (!by25[k]) by25[k] = { sales: 0, depts: {} };
      by25[k].sales += amt25;
    }
    if (c26 && typeof c26 === "number" && !isGAEPD(dept26)) {
      const k = String(c26);
      if (!by26[k]) by26[k] = { custNum: c26, customer: String(name26 || "").trim(), sales: 0, gp: 0, depts: {} };
      by26[k].sales += amt26; by26[k].gp += gp26;
      if (dept26) by26[k].depts[dept26] = (by26[k].depts[dept26] || 0) + amt26;
    }
  }
  const ap = [];
  for (const k of Object.keys(by26)) {
    const d = by26[k];
    const s25 = by25[k]?.sales || 0;
    const topDept = Object.keys(d.depts).sort((a, b) => d.depts[b] - d.depts[a])[0] || "";
    ap.push({
      custNum: Number(k), customer: d.customer, city: "Statesboro", salesman: "Austin",
      sales2025: Math.round(s25 * 100) / 100,
      sales2026: Math.round(d.sales * 100) / 100,
      change:    Math.round((d.sales - s25) * 100) / 100,
      gpPct:     d.sales > 0 ? Math.round((d.gp / d.sales) * 10000) / 10000 : 0,
      action: (d.sales - s25) > 500 ? "GROW" : (d.sales - s25) < -500 ? "LOST" : "OK",
      topDept, declinedDept: "", focus: topDept,
    });
  }
  return { actionPlan: ap };
}

// ── WTD sheet — parse branch comparison section (rows 10+) ──────────────────
function parseMasterWTD(wb) {
  if (!wb.Sheets["WTD"]) return null;
  const rows = getSheetRows(wb, "WTD", 0);
  const BRANCHES = {"1 - BYRON":"Byron","2 - TIFTON":"Tifton","3 - STATESBORO":"Statesboro","5 - ATHENS":"Athens"};
  const branchData = {};

  // Scan all rows for branch comparison data
  for (const row of rows) {
    const cell0 = String(row[0] || "").trim();
    const branch = BRANCHES[cell0];
    if (branch) {
      branchData[branch] = {
        sales2025: Number(row[1] || 0),
        sales2026: Number(row[2] || 0),
        gp2025:    Number(row[6] || row[5] || 0),
        gp2026:    Number(row[7] || row[6] || 0),
      };
    }
  }
  return Object.keys(branchData).length > 0 ? branchData : null;
}

// ── AD Programs ──────────────────────────────────────────────────────────────
function parseMasterToyo(wb) {
  if (!wb.Sheets["Toyo"]) return null;
  // rows 5+ (skip title, note, group-hdr, col-hdr)
  const rows = getSheetRows(wb, "Toyo", 4);
  const result = {};
  for (const row of rows) {
    if (!row[5] || !String(row[5]).includes("Tifton")) continue;
    const custName = String(row[3] || "").trim();
    const toyoNum  = String(row[2] || "").trim();
    result[toyoNum] = {
      toyoNum, dealerName: custName,
      pcr: {
        primary:   Number(row[6] || 0),
        secondary: Number(row[7] || 0),
        total:     Number(row[8] || 0),
        pct:       parseFloat(String(row[9] || "0").replace("%", "")) || 0,
      },
      tbr: {
        primary:   Number(row[10] || 0),
        secondary: Number(row[11] || 0),
        total:     Number(row[12] || 0),
      },
    };
  }
  return Object.keys(result).length > 0 ? result : null;
}

function parseMasterAmericus(wb) {
  if (!wb.Sheets["Americus"]) return null;
  const rows = getSheetRows(wb, "Americus", 4);
  const result = {};
  for (const row of rows) {
    const dealerNum = row[0];
    if (!dealerNum || typeof dealerNum !== "number") continue;
    const name = String(row[1] || "").trim();
    if (!name || name.toLowerCase().includes("total")) continue;
    result[String(dealerNum)] = {
      dealerNum,
      dealerName: name,
      units2025:  Number(row[6] || 0),
      jan:        Number(row[7] || 0),
      feb:        Number(row[8] || 0),
      mar:        Number(row[9] || 0),
      q1:         Number(row[10] || 0),
      apr:        Number(row[11] || 0),
      may:        Number(row[12] || 0),
      jun:        Number(row[13] || 0),
      q2:         Number(row[14] || 0),
      ytd:        Number(row[15] || 0),
      asOf:       new Date().toLocaleDateString("en-US", { month: "short", year: "numeric" }),
    };
  }
  return Object.keys(result).length > 0 ? result : null;
}

function parseMasterAscenso(wb) {
  if (!wb.Sheets["Ascenso"]) return null;
  const rows = getSheetRows(wb, "Ascenso", 4);
  const result = {};
  for (const row of rows) {
    const custNum = row[0];
    if (!custNum || typeof custNum !== "number") continue;
    result[String(custNum)] = {
      custNum,
      name:   String(row[1] || "").trim(),
      qty:    Number(row[3] || 0),
      amount: Number(row[4] || 0),
    };
  }
  return Object.keys(result).length > 0 ? result : null;
}

function parseMasterFalken(wb, sheetName) {
  if (!wb.Sheets[sheetName]) return null;
  const rows = getSheetRows(wb, sheetName, 4);
  const result = [];
  for (const row of rows) {
    const fanId = row[0];
    if (!fanId || typeof fanId !== "number") continue;
    result.push({
      fanId,
      segment:   String(row[1] || "").trim(),
      parentId:  row[2],
      dealer:    String(row[3] || "").trim(),
      city:      String(row[7] || "").trim(),
      territory: String(row[9] || "").trim(),
      q1:        Number(row[10] || 0),
      q2:        Number(row[11] || 0),
      q3:        Number(row[12] || 0),
      q4:        Number(row[13] || 0),
      ytd:       Number(row[14] || 0),
    });
  }
  return result.length > 0 ? result : null;
}

function parseMasterBARNN(wb) {
  if (!wb.Sheets["BARNN"]) return null;
  const rows = getSheetRows(wb, "BARNN", 4);
  const result = { primary: [], secondary: [] };
  for (const row of rows) {
    const parent  = row[0];
    const acctNum = row[1];
    const name    = String(row[2] || "").trim();
    if (!acctNum || typeof acctNum !== "number") continue;
    if (name.toLowerCase().includes("total")) continue;
    const entry = {
      parent: String(parent || ""),
      acctNum,
      name,
      bs:    Number(row[5] || 0),
      fs:    Number(row[8] || 0),
      total: Number(row[15] || 0),
      role:  String(parent) === "960788" ? "Primary" : "Secondary",
    };
    if (String(parent) === "960788") result.primary.push(entry);
    else result.secondary.push(entry);
  }
  return (result.primary.length + result.secondary.length) > 0 ? result : null;
}

function parseMasterYokohama(wb) {
  if (!wb.Sheets["Yokohama"]) return null;
  const rows = getSheetRows(wb, "Yokohama", 4);
  const result = [];
  for (const row of rows) {
    const dealerNum = row[0];
    if (!dealerNum || typeof dealerNum !== "number") continue;
    result.push({
      dealerNum,
      dealerName:  String(row[1] || "").trim(),
      salesRep:    String(row[2] || "").trim(),
      primary:     Number(row[4] || 0),
      priPct:      parseFloat(String(row[5] || "0").replace("%", "")) || 0,
      hasSecondary: String(row[6] || "").trim().toUpperCase() === "Y",
      secondary:   Number(row[8] || 0),
      qtd:         Number(row[10] || 0),
      toNext:      Number(row[11] || 0),
    });
  }
  return result.length > 0 ? result : null;
}

// ── Master entry point ────────────────────────────────────────────────────────
function parseMasterWorkbook(wb) {
  return {
    ar:         parseMasterAR(wb),
    customerComp: parseMasterCustomerComp(wb),
    statesboro: parseMasterStatesboro(wb),
    wtd:        parseMasterWTD(wb),
    toyo:       parseMasterToyo(wb),
    americus:   parseMasterAmericus(wb),
    ascenso:    parseMasterAscenso(wb),
    falkenPLT:  parseMasterFalken(wb, "Falken_PLT"),
    falkenTBR:  parseMasterFalken(wb, "Falken_TBR"),
    barnn:      parseMasterBARNN(wb),
    yokohama:   parseMasterYokohama(wb),
  };
}

const TABS = [
  { id: "setup",    label: "⬡ Files" },
  { id: "overview", label: "Overview" },
  { id: "tiffany",  label: "Tiffany", rep: true },
  { id: "larry",    label: "Larry",   rep: true },
  { id: "austin",   label: "Austin",  rep: true },
  { id: "house",    label: "House",   rep: true },
  { id: "cardealer",label: "Car Dealer", rep: true },
  { id: "ai",       label: "◈ AI" },
  { id: "map",      label: "⊙ Map" },
  { id: "ar",       label: "$ AR" },
];

// ── Styles ────────────────────────────────────────────────────────────────────
const S = {
  app:    { fontFamily: "'Inter','Segoe UI','Helvetica Neue',Arial,sans-serif", background: "#F4F7FB", minHeight: "100vh", color: TEXT, display: "flex", flexDirection: "column" },
  header: { padding: "0.85rem 2rem", borderBottom: `1px solid ${BORDER}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: "#1E5FCC" },
  logo:   { fontSize: "0.95rem", fontWeight: 700, letterSpacing: "0.18em", color: "#FFFFFF", textTransform: "uppercase" },
  sub:    { color: "rgba(255,255,255,0.7)", fontSize: "0.68rem", letterSpacing: "0.18em", marginTop: 2 },
  nav:    { display: "flex", borderBottom: `1px solid ${BORDER}`, background: "#FFFFFF", paddingLeft: "1.5rem", overflowX: "auto" },
  navBtn: (active, color) => ({
    padding: "0.75rem 1.2rem", background: "transparent", border: "none",
    borderBottom: `3px solid ${active ? (color || AMBER) : "transparent"}`,
    color: active ? (color || AMBER) : MUTED,
    cursor: "pointer", fontSize: "0.72rem", letterSpacing: "0.1em",
    textTransform: "uppercase", whiteSpace: "nowrap", transition: "color 0.2s",
    fontWeight: active ? 700 : 400,
  }),
  main:   { flex: 1, padding: "1.5rem 2rem", overflowY: "auto" },
  card:   { background: BG2, border: `1px solid ${BORDER}`, borderRadius: 6, padding: "1.25rem", marginBottom: "1rem" },
  kpiRow: { display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: "0.75rem", marginBottom: "1rem" },
  kpi:    (color) => ({ background: BG2, border: `1px solid ${BORDER}`, borderRadius: 6, padding: "1rem 1.25rem", borderLeft: `3px solid ${color || AMBER}` }),
  kpiVal: { fontSize: "1.4rem", fontWeight: 700, color: TEXT, lineHeight: 1.2 },
  kpiLbl: { fontSize: "0.68rem", color: MUTED, letterSpacing: "0.12em", textTransform: "uppercase", marginTop: 2 },
  table:  { width: "100%", borderCollapse: "collapse", fontSize: "0.72rem" },
  th:     { padding: "0.5rem 0.75rem", color: MUTED, borderBottom: `1px solid ${BORDER}`, textAlign: "left", fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", fontSize: "0.68rem" },
  td:     { padding: "0.45rem 0.75rem", borderBottom: `1px solid ${BORDER}`, color: TEXT },
  subNav: { display: "flex", gap: 4, marginBottom: "1.25rem" },
  subBtn: (a, c) => ({ padding: "0.4rem 0.9rem", background: a ? (c || AMBER) : "transparent", border: `1px solid ${a ? (c||AMBER) : BORDER}`, color: a ? BG2 : MUTED, borderRadius: 4, cursor: "pointer", fontSize: "0.68rem", letterSpacing: "0.08em" }),
  btn:    (c) => ({ padding: "0.35rem 0.85rem", background: "transparent", border: `1px solid ${c || BORDER}`, color: c || MUTED, borderRadius: 4, cursor: "pointer", fontSize: "0.68rem" }),
};

function fmt(n, pre="$") {
  if (n == null || isNaN(n)) return "—";
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  if (abs >= 1000000) return `${sign}${pre}${(abs/1000000).toFixed(1)}M`;
  if (abs >= 1000)    return `${sign}${pre}${(abs/1000).toFixed(1)}K`;
  return `${sign}${pre}${abs.toFixed(0)}`;
}
function pct(n) { return n == null || isNaN(n) ? "—" : `${(n*100).toFixed(1)}%`; }
function clr(n) { return n > 0 ? GREEN : n < 0 ? RED : TEXT; }

// ── Seed Data ─────────────────────────────────────────────────────────────────
const SEED_WEEK_COMP = {"weeks":[{"week":1,"sales2025":123427.64,"sales2026":45856.73,"change":-77570.91,"changePct":-0.628473,"gp2025":18015.72,"gp2026":7458.9,"gpChange":-10556.82,"locations":{"Byron":{"sales2025":163452.32,"sales2026":71316.71,"gp2025":21613.49,"gp2026":9927.56},"Tifton":{"sales2025":124055.64,"sales2026":45985.73,"gp2025":17983.88,"gp2026":7437.01},"Statesboro":{"sales2025":115720.85,"sales2026":46755.93,"gp2025":16574.42,"gp2026":5754.92},"Athens":{"sales2025":115170.69,"sales2026":73570.06,"gp2025":17369.69,"gp2026":11388.53}}},{"week":2,"sales2025":301822.72,"sales2026":382007.38,"change":80184.66,"changePct":0.265668,"gp2025":46329.9,"gp2026":54276.6,"gpChange":7946.7,"locations":{"Byron":{"sales2025":377044.43,"sales2026":570752.41,"gp2025":52249.02,"gp2026":76862.73},"Tifton":{"sales2025":303298.81,"sales2026":383463.71,"gp2025":46284.22,"gp2026":54183.13},"Statesboro":{"sales2025":462386.87,"sales2026":481873.11,"gp2025":60271.41,"gp2026":59412.46},"Athens":{"sales2025":273827.42,"sales2026":383818.48,"gp2025":39103.51,"gp2026":55728.0}}},{"week":3,"sales2025":245152.08,"sales2026":374483.37,"change":129331.29,"changePct":0.527555,"gp2025":38402.28,"gp2026":61237.1,"gpChange":22834.82,"locations":{"Byron":{"sales2025":425593.39,"sales2026":426682.0,"gp2025":58172.78,"gp2026":67610.08},"Tifton":{"sales2025":246446.08,"sales2026":376015.37,"gp2025":38342.58,"gp2026":61165.46},"Statesboro":{"sales2025":427419.41,"sales2026":404745.78,"gp2025":61195.15,"gp2026":48898.06},"Athens":{"sales2025":349566.16,"sales2026":342498.08,"gp2025":49966.03,"gp2026":47315.53}}},{"week":4,"sales2025":161565.13,"sales2026":350548.24,"change":188983.11,"changePct":1.169702,"gp2025":27003.95,"gp2026":50179.59,"gpChange":23175.64,"locations":{"Byron":{"sales2025":247061.2,"sales2026":522986.13,"gp2025":34166.35,"gp2026":70060.96},"Tifton":{"sales2025":162354.13,"sales2026":351882.3,"gp2025":26948.23,"gp2026":50151.79},"Statesboro":{"sales2025":250714.52,"sales2026":380511.43,"gp2025":34832.31,"gp2026":48248.13},"Athens":{"sales2025":268575.17,"sales2026":361138.66,"gp2025":37126.24,"gp2026":48640.81}}},{"week":5,"sales2025":261017.26,"sales2026":309344.09,"change":48326.83,"changePct":0.185148,"gp2025":39303.55,"gp2026":43783.31,"gpChange":4479.76,"locations":{"Byron":{"sales2025":320762.32,"sales2026":475548.29,"gp2025":41937.18,"gp2026":60561.96},"Tifton":{"sales2025":262329.26,"sales2026":308672.3,"gp2025":39271.71,"gp2026":43662.98},"Statesboro":{"sales2025":405300.67,"sales2026":415940.26,"gp2025":53454.33,"gp2026":49150.08},"Athens":{"sales2025":254913.44,"sales2026":264278.07,"gp2025":35273.09,"gp2026":33378.33}}},{"week":6,"sales2025":348604.74,"sales2026":348507.15,"change":-97.59,"changePct":-0.00028,"gp2025":50704.85,"gp2026":45700.15,"gpChange":-5004.7,"locations":{"Byron":{"sales2025":474252.62,"sales2026":476076.79,"gp2025":56572.2,"gp2026":61164.86},"Tifton":{"sales2025":350518.74,"sales2026":349455.92,"gp2025":50637.19,"gp2026":45500.63},"Statesboro":{"sales2025":493114.38,"sales2026":370412.24,"gp2025":61443.78,"gp2026":44444.13},"Athens":{"sales2025":354158.58,"sales2026":420047.08,"gp2025":45660.65,"gp2026":49245.91}}},{"week":7,"sales2025":306406.92,"sales2026":316725.6,"change":10318.68,"changePct":0.033676,"gp2025":43977.04,"gp2026":42577.08,"gpChange":-1399.96,"locations":{"Byron":{"sales2025":420629.79,"sales2026":465205.62,"gp2025":53125.83,"gp2026":57900.03},"Tifton":{"sales2025":308024.92,"sales2026":318158.47,"gp2025":43919.33,"gp2026":42513.41},"Statesboro":{"sales2025":461628.62,"sales2026":436169.39,"gp2025":42022.9,"gp2026":47512.59},"Athens":{"sales2025":386143.35,"sales2026":392013.04,"gp2025":49769.34,"gp2026":44201.35}}},{"week":8,"sales2025":303267.93,"sales2026":431601.88,"change":128333.95,"changePct":0.42317,"gp2025":36984.54,"gp2026":56432.01,"gpChange":19447.47,"locations":{"Byron":{"sales2025":421597.5,"sales2026":509094.67,"gp2025":54514.42,"gp2026":63853.94},"Tifton":{"sales2025":305073.85,"sales2026":433432.69,"gp2025":36920.78,"gp2026":56316.59},"Statesboro":{"sales2025":495206.73,"sales2026":471785.27,"gp2025":65989.66,"gp2026":51875.84},"Athens":{"sales2025":352936.66,"sales2026":423549.38,"gp2025":43955.8,"gp2026":51934.53}}},{"week":9,"sales2025":354550.14,"sales2026":452361.44,"change":97811.3,"changePct":0.275874,"gp2025":54039.46,"gp2026":65480.89,"gpChange":11441.43,"locations":{"Byron":{"sales2025":388503.35,"sales2026":545565.1,"gp2025":61624.26,"gp2026":66852.36},"Tifton":{"sales2025":356486.14,"sales2026":454495.49,"gp2025":53937.97,"gp2026":0},"Statesboro":{"sales2025":487011.7,"sales2026":440676.51,"gp2025":64815.44,"gp2026":36031.08},"Athens":{"sales2025":332534.01,"sales2026":447750.21,"gp2025":44801.32,"gp2026":54336.7}}},{"week":10,"sales2025":412345.78,"sales2026":468339.07,"change":55993.29,"changePct":0.135792,"gp2025":58080.16,"gp2026":65382.09,"gpChange":7301.93,"locations":{"Byron":{"sales2025":634848.31,"sales2026":578802.13,"gp2025":73784.03,"gp2026":77971.84},"Tifton":{"sales2025":414364.78,"sales2026":470118.83,"gp2025":57982.65,"gp2026":65184.07},"Statesboro":{"sales2025":571196.82,"sales2026":497235.86,"gp2025":61522.51,"gp2026":55352.06},"Athens":{"sales2025":375647.27,"sales2026":412614.08,"gp2025":47383.42,"gp2026":56149.13}}},{"week":11,"sales2025":299421.68,"sales2026":546099.8,"change":246678.12,"changePct":0.823849,"gp2025":47738.92,"gp2026":79416.32,"gpChange":31677.4,"locations":{"Byron":{"sales2025":475206.0,"sales2026":640012.25,"gp2025":66232.5,"gp2026":80560.68},"Tifton":{"sales2025":301002.68,"sales2026":543488.79,"gp2025":47663.3,"gp2026":78817.87},"Statesboro":{"sales2025":495203.93,"sales2026":504056.61,"gp2025":66603.02,"gp2026":55433.1},"Athens":{"sales2025":403830.44,"sales2026":501260.43,"gp2025":51959.3,"gp2026":70066.14}}},{"week":12,"sales2025":298950.36,"sales2026":424333.68,"change":125383.32,"changePct":0.419412,"gp2025":48950.19,"gp2026":60917.7,"gpChange":11967.51,"locations":{"Byron":{"sales2025":453355.91,"sales2026":601780.28,"gp2025":64771.83,"gp2026":67903.66},"Tifton":{"sales2025":300523.39,"sales2026":423218.69,"gp2025":48842.76,"gp2026":60222.15},"Statesboro":{"sales2025":536189.73,"sales2026":468570.64,"gp2025":71961.73,"gp2026":54358.2},"Athens":{"sales2025":364320.36,"sales2026":471417.78,"gp2025":51103.97,"gp2026":63914.76}}},{"week":13,"sales2025":313484.87,"sales2026":457493.26,"change":144008.39,"changePct":0.459379,"gp2025":46319.04,"gp2026":60789.56,"gpChange":14470.52,"locations":{"Byron":{"sales2025":397626.94,"sales2026":644955.08,"gp2025":57424.99,"gp2026":65278.52},"Tifton":{"sales2025":315203.87,"sales2026":459112.59,"gp2025":46243.42,"gp2026":60656.24},"Statesboro":{"sales2025":521325.91,"sales2026":537096.2,"gp2025":70354.11,"gp2026":58086.69},"Athens":{"sales2025":359990.21,"sales2026":499530.35,"gp2025":49554.54,"gp2026":60263.27}}},{"week":14,"sales2025":341112.78,"sales2026":522761.89,"change":181649.11,"changePct":0.532519,"gp2025":50418.76,"gp2026":66013.4,"gpChange":15594.64,"locations":{"Byron":{"sales2025":531766.41,"sales2026":695710.98,"gp2025":70423.4,"gp2026":88238.24},"Tifton":{"sales2025":342927.94,"sales2026":524966.89,"gp2025":50299.52,"gp2026":65927.83},"Statesboro":{"sales2025":515104.02,"sales2026":518940.04,"gp2025":70282.06,"gp2026":60444.4},"Athens":{"sales2025":404237.49,"sales2026":435410.47,"gp2025":55055.27,"gp2026":52300.73}}},{"week":15,"sales2025":380833.6,"sales2026":441159.9,"change":60326.3,"changePct":0.158406,"gp2025":54651.08,"gp2026":61749.54,"gpChange":7098.46,"locations":{"Byron":{"sales2025":591777.75,"sales2026":722977.83,"gp2025":87767.2,"gp2026":90146.06},"Tifton":{"sales2025":382489.6,"sales2026":441600.98,"gp2025":54545.61,"gp2026":61626.96},"Statesboro":{"sales2025":557293.32,"sales2026":487888.46,"gp2025":73158.27,"gp2026":54237.02},"Athens":{"sales2025":549425.64,"sales2026":372553.22,"gp2025":76248.34,"gp2026":50114.46}}},{"week":16,"sales2025":333294.04,"sales2026":393167.3,"change":59873.26,"changePct":0.179641,"gp2025":56625.53,"gp2026":53546.95,"gpChange":-3078.58,"locations":{"Byron":{"sales2025":596454.03,"sales2026":581034.42,"gp2025":96832.78,"gp2026":78029.98},"Tifton":{"sales2025":334890.98,"sales2026":394978.03,"gp2025":56567.76,"gp2026":53466.73},"Statesboro":{"sales2025":527256.33,"sales2026":514214.76,"gp2025":83874.62,"gp2026":56587.44},"Athens":{"sales2025":454735.18,"sales2026":385857.73,"gp2025":63725.16,"gp2026":49486.82}}},{"week":17,"sales2025":355870.61,"sales2026":420293.86,"change":64423.25,"changePct":0.18103,"gp2025":64902.25,"gp2026":58256.7,"gpChange":-6645.55,"locations":{"Byron":{"sales2025":484884.37,"sales2026":631906.61,"gp2025":83593.59,"gp2026":81159.27},"Tifton":{"sales2025":357674.61,"sales2026":422232.86,"gp2025":64782.85,"gp2026":58218.89},"Statesboro":{"sales2025":539417.6,"sales2026":519390.14,"gp2025":83810.45,"gp2026":61531.18},"Athens":{"sales2025":324248.22,"sales2026":453313.95,"gp2025":55300.1,"gp2026":56826.95}}},{"week":18,"sales2025":394844.47,"sales2026":441890.32,"change":47045.85,"changePct":0.11915,"gp2025":75014.24,"gp2026":60198.08,"gpChange":-14816.16,"locations":{"Byron":{"sales2025":519580.9,"sales2026":550851.99,"gp2025":83804.41,"gp2026":71336.72},"Tifton":{"sales2025":396848.53,"sales2026":443646.32,"gp2025":74924.75,"gp2026":60144.35},"Statesboro":{"sales2025":531111.92,"sales2026":521227.35,"gp2025":83353.85,"gp2026":60075.84},"Athens":{"sales2025":347877.2,"sales2026":511784.91,"gp2025":62474.46,"gp2026":61587.44}}},{"week":19,"sales2025":428994.06,"sales2026":468288.87,"change":39294.81,"changePct":0.091598,"gp2025":81852.17,"gp2026":66516.12,"gpChange":-15336.05,"locations":{"Byron":{"sales2025":595298.83,"sales2026":533288.93,"gp2025":100931.56,"gp2026":73801.54},"Tifton":{"sales2025":430969.01,"sales2026":470134.9,"gp2025":81700.88,"gp2026":66433.14},"Statesboro":{"sales2025":584155.89,"sales2026":501653.66,"gp2025":92927.32,"gp2026":59357.94},"Athens":{"sales2025":381173.58,"sales2026":551444.45,"gp2025":69387.3,"gp2026":72470.71}}},{"week":20,"sales2025":326178.17,"sales2026":399853.0,"change":73674.83,"changePct":0.225873,"gp2025":64917.52,"gp2026":57051.05,"gpChange":-7866.47,"locations":{"Byron":{"sales2025":488434.99,"sales2026":537770.24,"gp2025":82420.49,"gp2026":74998.44},"Tifton":{"sales2025":327879.41,"sales2026":401595.03,"gp2025":64832.19,"gp2026":57001.33},"Statesboro":{"sales2025":529184.8,"sales2026":463117.3,"gp2025":88137.64,"gp2026":54170.94},"Athens":{"sales2025":381262.85,"sales2026":399091.79,"gp2025":67066.79,"gp2026":51646.22}}},{"week":21,"sales2025":393701.7,"sales2026":377297.0,"change":-16404.7,"changePct":-0.041668,"gp2025":73092.23,"gp2026":58301.08,"gpChange":-14791.15,"locations":{"Byron":{"sales2025":550075.48,"sales2026":581257.34,"gp2025":97040.79,"gp2026":77922.52},"Tifton":{"sales2025":395583.7,"sales2026":378909.39,"gp2025":73030.54,"gp2026":58212.12},"Statesboro":{"sales2025":575972.34,"sales2026":475852.73,"gp2025":94703.25,"gp2026":58912.24},"Athens":{"sales2025":405443.52,"sales2026":407849.53,"gp2025":71959.77,"gp2026":56981.17}}},{"week":22,"sales2025":346389.91,"sales2026":286570.53,"change":-59819.38,"changePct":-0.172694,"gp2025":61512.47,"gp2026":46168.86,"gpChange":-15343.61,"locations":{"Byron":{"sales2025":444933.24,"sales2026":385149.63,"gp2025":74142.43,"gp2026":57098.92},"Tifton":{"sales2025":347990.91,"sales2026":287879.95,"gp2025":61416.95,"gp2026":46124.68},"Statesboro":{"sales2025":428728.05,"sales2026":425248.3,"gp2025":68453.95,"gp2026":50103.87},"Athens":{"sales2025":387834.9,"sales2026":396320.4,"gp2025":60949.68,"gp2026":53828.87}}}],"depts":[{"dept":"RAD LT TRUCK","sales":3336374.28,"gp":487161.06,"gpPct":0.146015,"lineItems":2344,"assessment":"TOP PERFORMER"},{"dept":"TRUCK TIRES","sales":1955142.51,"gp":188114.44,"gpPct":0.096215,"lineItems":1079,"assessment":"STRONG - watch margin"},{"dept":"PASSENGER TIRES","sales":1799123.91,"gp":308008.03,"gpPct":0.171199,"lineItems":2379,"assessment":"STRONG - good margins"},{"dept":"FARM TIRES","sales":530847.03,"gp":67175.63,"gpPct":0.126544,"lineItems":501,"assessment":"STRONG - watch margin"},{"dept":"ST TRAILER","sales":395387.08,"gp":82104.05,"gpPct":0.207655,"lineItems":929,"assessment":"STRONG - good margins"},{"dept":"OFF THE ROAD TIRES","sales":298114.81,"gp":25068.48,"gpPct":0.08409,"lineItems":103,"assessment":"STRONG - watch margin"},{"dept":"INDUSTRIAL TIRES","sales":149587.24,"gp":14347.6,"gpPct":0.095915,"lineItems":168,"assessment":"STRONG - watch margin"},{"dept":"TUBES","sales":99809.71,"gp":21415.89,"gpPct":0.214567,"lineItems":568,"assessment":"GOOD - Healthy margins"},{"dept":"VALVE STEMS","sales":28392.06,"gp":9262.57,"gpPct":0.326238,"lineItems":148,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"WHEEL WEIGHTS","sales":16773.06,"gp":7682.75,"gpPct":0.458041,"lineItems":105,"assessment":"BEST MARGIN - push as add-on"},{"dept":"TIRE TOOLS","sales":13680.43,"gp":3977.46,"gpPct":0.290741,"lineItems":60,"assessment":"GOOD - Healthy margins"},{"dept":"WHEELS","sales":11990.9,"gp":1750.6,"gpPct":0.145994,"lineItems":46,"assessment":"WATCH - GP needs improvement"},{"dept":"LAWN & GARDEN","sales":8019.26,"gp":1569.39,"gpPct":0.195703,"lineItems":106,"assessment":"STEADY - Monitor"},{"dept":"PATCHES AND REPAIR","sales":7675.58,"gp":2623.91,"gpPct":0.341852,"lineItems":124,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"OUTSIDE PURCHASE","sales":5481.43,"gp":585.01,"gpPct":0.106726,"lineItems":4,"assessment":"WATCH - GP needs improvement"},{"dept":"MOUNTING LUBE","sales":3474.85,"gp":1232.37,"gpPct":0.354654,"lineItems":43,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"ALIGNMENT SHIMS","sales":1767.14,"gp":429.79,"gpPct":0.243212,"lineItems":2,"assessment":"GOOD - Healthy margins"},{"dept":"FREIGHT CHARGES","sales":644.35,"gp":359.35,"gpPct":0.557694,"lineItems":3,"assessment":"BEST MARGIN - push as add-on"},{"dept":"ATV TIRES","sales":390.25,"gp":124.17,"gpPct":0.318181,"lineItems":2,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"DEPARTMENTS DOING WELL:","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 RAD LT TRUCK","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 TRUCK TIRES","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 PASSENGER TIRES","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 VALVE STEMS","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 WHEEL WEIGHTS","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 PATCHES AND REPAIR","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 MOUNTING LUBE","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 FREIGHT CHARGES","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 ATV TIRES","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"DEPARTMENTS TO FOCUS ON:","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 TRUCK TIRES (low GP 9.6%)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 OFF THE ROAD TIRES (low GP 8.4%)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 INDUSTRIAL TIRES (low GP 9.6%)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 LAWN & GARDEN (low volume $8019)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 PATCHES AND REPAIR (low volume $7676)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 OUTSIDE PURCHASE (low volume $5481)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 MOUNTING LUBE (low volume $3475)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 ALIGNMENT SHIMS (low volume $1767)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 FREIGHT CHARGES (low volume $644)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""},{"dept":"  \u2022 ATV TIRES (low volume $390)","sales":0.0,"gp":0.0,"gpPct":0.0,"lineItems":0,"assessment":""}],"actionPlan":[{"salesman":"Anthony","city":"ALBANY","custNum":200765,"customer":"ALBANY CHRYSLER DODGE JEEP RAM","sales2025":5834.95,"sales2026":4105.09,"change":-1729.86,"gpPct":0.258299,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$3085)"},{"salesman":"Anthony","city":"ALBANY","custNum":200838,"customer":"ALBANY CHRYSLER DODGE JEEP RAM","sales2025":1110.17,"sales2026":2583.76,"change":1473.59,"gpPct":0.112882,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (7.5%)"},{"salesman":"Anthony","city":"ALBANY","custNum":200360,"customer":"PRINCE CHEVY BUICK GMC","sales2025":281.88,"sales2026":0.0,"change":-281.88,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $282)"},{"salesman":"Anthony","city":"AUBURNDALE","custNum":2000024,"customer":"FORKLIFT TIRE OF CENTRAL FL","sales2025":0.0,"sales2026":18120.32,"change":18120.32,"gpPct":0.103881,"action":"GROWING - Expand","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"CAIRO","custNum":200626,"customer":"HOBSON CHEVROLET BUICK","sales2025":383.48,"sales2026":0.0,"change":-383.48,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $383)"},{"salesman":"Anthony","city":"CORDELE","custNum":2000025,"customer":"FORD CORDELE","sales2025":0.0,"sales2026":10936.79,"change":10936.79,"gpPct":0.17733,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (6.4%)"},{"salesman":"Anthony","city":"CORDELE","custNum":200882,"customer":"SUNBELT FORD CORDELE (AMI)","sales2025":690.84,"sales2026":0.0,"change":-690.84,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $691)"},{"salesman":"Anthony","city":"CORDELE","custNum":101568,"customer":"SUNBELT FORD OF CORDELE  INC","sales2025":7878.73,"sales2026":0.0,"change":-7878.73,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $5409) | LOST: PASSENGER TIRES (was $2469)"},{"salesman":"Anthony","city":"DOTHAN","custNum":200943,"customer":"BONDY'S NISSAN INC.","sales2025":482.63,"sales2026":0.0,"change":-482.63,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $483)"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200777,"customer":"ANDERSON BUICK GMC","sales2025":3385.62,"sales2026":6786.56,"change":3400.94,"gpPct":0.141415,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $1162)"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200734,"customer":"ANDERSON FORD","sales2025":0.0,"sales2026":11375.81,"change":11375.81,"gpPct":0.171655,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"DOUGLAS","custNum":101256,"customer":"ROBERT FENDER CHEVROLET","sales2025":0.0,"sales2026":60.13,"change":60.13,"gpPct":0.276401,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200778,"customer":"WOODY FOLSOM CDJR (DOUGLAS)","sales2025":0.0,"sales2026":716.64,"change":716.64,"gpPct":0.025508,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (2.6%)"},{"salesman":"Anthony","city":"FITZGERALD","custNum":200833,"customer":"FITZGERALD CHRYSLER DODGE RAM","sales2025":1925.62,"sales2026":8501.62,"change":6576.0,"gpPct":0.129414,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"FITZGERALD","custNum":101300,"customer":"FITZGERALD FORD AND LINCOLN","sales2025":1497.88,"sales2026":3057.49,"change":1559.61,"gpPct":0.081174,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.8%)"},{"salesman":"Anthony","city":"MOULTRIE","custNum":200365,"customer":"ROBERT HUTSON LINCOLN","sales2025":5103.78,"sales2026":10719.35,"change":5615.57,"gpPct":0.173107,"action":"GROWING - Expand","topDept":"OFF THE ROAD TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1982) | LOST: ST TRAILER (was $149)"},{"salesman":"Anthony","city":"NASHVILLE","custNum":2000009,"customer":"KING FORD OF NASHVILLE","sales2025":3915.44,"sales2026":8401.9,"change":4486.46,"gpPct":0.232158,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"NASHVILLE","custNum":200688,"customer":"NASHVILLE FORD","sales2025":7675.92,"sales2026":0.0,"change":-7675.92,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $3150) | LOST: RAD LT TRUCK (was $4526)"},{"salesman":"Anthony","city":"NASHVILLE","custNum":200908,"customer":"O'STEEN CHRYSLER DODGE JEEP","sales2025":223.16,"sales2026":845.43,"change":622.27,"gpPct":0.193795,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (8.2%)"},{"salesman":"Anthony","city":"QUITMAN","custNum":200783,"customer":"CASS BURCH CHEVROLET","sales2025":0.0,"sales2026":5365.84,"change":5365.84,"gpPct":0.095167,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (7.4%)"},{"salesman":"Anthony","city":"SYLVESTER","custNum":2000027,"customer":"FORD SYLVESTER","sales2025":0.0,"sales2026":2182.26,"change":2182.26,"gpPct":0.077099,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"SYLVESTER","custNum":1999999,"customer":"GRIFFIN CHEVROLET OF SYLVESTER","sales2025":1737.92,"sales2026":408.24,"change":-1329.68,"gpPct":0.246375,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1426)"},{"salesman":"Anthony","city":"SYLVESTER","custNum":200232,"customer":"SUNBELT FORD INC","sales2025":3522.1,"sales2026":0.0,"change":-3522.1,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $776) | LOST: RAD LT TRUCK (was $2692) | LOST: ST TRAILER (was $55)"},{"salesman":"Anthony","city":"TIFTON","custNum":200527,"customer":"GRIFFIN CHRY/DOD/JEEP/RAM","sales2025":8331.44,"sales2026":2038.78,"change":-6292.66,"gpPct":0.158281,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $3022) | DOWN: RAD LT TRUCK (-$3043) | LOST: ST TRAILER (was $228)"},{"salesman":"Anthony","city":"TIFTON","custNum":200831,"customer":"GRIFFIN CHRYSLER DODGE JEEP","sales2025":415.32,"sales2026":7858.3,"change":7442.98,"gpPct":0.142849,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"TIFTON","custNum":200832,"customer":"GRIFFIN FORD","sales2025":13146.27,"sales2026":11456.25,"change":-1690.02,"gpPct":0.167063,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$4375) | LOST: ST TRAILER (was $78)"},{"salesman":"Anthony","city":"TIFTON","custNum":200524,"customer":"GRIFFIN FORD LINCOLN  INC","sales2025":4936.42,"sales2026":0.0,"change":-4936.42,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $4079) | LOST: ST TRAILER (was $96) | LOST: PASSENGER TIRES (was $761)"},{"salesman":"Anthony","city":"TIFTON","custNum":200554,"customer":"HONDA OF SOUTH GEORGIA","sales2025":825.6,"sales2026":19.72,"change":-805.88,"gpPct":0.330629,"action":"DECLINING - Investigate","topDept":"MOUNTING LUBE","declinedDept":"RAD LT TRUCK","focus":"LOST: ST TRAILER (was $206) | LOST: RAD LT TRUCK (was $619)"},{"salesman":"Anthony","city":"TIFTON","custNum":200580,"customer":"JEFF FENDER BUICK  GMC, CAD.","sales2025":639.36,"sales2026":9360.58,"change":8721.22,"gpPct":0.08101,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (8.0%) | LOST: PASSENGER TIRES (was $639)"},{"salesman":"Anthony","city":"TIFTON","custNum":101374,"customer":"PRINCE CHEVY-OLDS  INC","sales2025":10604.56,"sales2026":10281.35,"change":-323.21,"gpPct":0.15976,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: FARM TIRES (was $183) | DOWN: TRUCK TIRES (-$550)"},{"salesman":"Anthony","city":"TIFTON","custNum":200455,"customer":"PRINCE HONDA","sales2025":1433.99,"sales2026":1152.12,"change":-281.87,"gpPct":0.240791,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$804)"},{"salesman":"Anthony","city":"TIFTON","custNum":200456,"customer":"PRINCE TOYOTA","sales2025":36256.29,"sales2026":18468.54,"change":-17787.75,"gpPct":0.042846,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: RAD LT TRUCK (-$14184) | LOW GP: PASSENGER TIRES (6.4%) | DOWN: PASSENGER TIRES (-$3604)"},{"salesman":"Anthony","city":"TIFTON","custNum":101298,"customer":"TENNESON NISSAN","sales2025":6079.94,"sales2026":13673.33,"change":7593.39,"gpPct":0.132687,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200933,"customer":"GRIFFIN CDJR VALDOSTA","sales2025":4956.24,"sales2026":5565.19,"change":608.95,"gpPct":0.212383,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200588,"customer":"LANGDALE HYUNDAI OF SOUTH GA","sales2025":2251.49,"sales2026":1855.24,"change":-396.25,"gpPct":0.211207,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $259)"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200348,"customer":"LANGDALE KIA (SERVICE)","sales2025":2756.25,"sales2026":845.43,"change":-1910.82,"gpPct":0.290574,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$307) | DOWN: RAD LT TRUCK (-$1604)"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200790,"customer":"PRINCE AUTO. VALDOSTA  BUICK","sales2025":0.0,"sales2026":5164.5,"change":5164.5,"gpPct":0.12206,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"WAYCROSS","custNum":200484,"customer":"ROBBIE ROBERSON FORD","sales2025":1293.68,"sales2026":184.22,"change":-1109.46,"gpPct":0.080067,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $164) | LOW GP: RAD LT TRUCK (8.0%) | DOWN: RAD LT TRUCK (-$945)"},{"salesman":"Anthony","city":"WAYCROSS","custNum":200461,"customer":"WALKER JONES CHEVY-BUICK","sales2025":2455.18,"sales2026":15608.85,"change":13153.67,"gpPct":0.057086,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.9%)"},{"salesman":"Austin","city":"MOULTRIE","custNum":2000030,"customer":"BEASON EQUIPMENT CO","sales2025":0.0,"sales2026":5135.11,"change":5135.11,"gpPct":0.104109,"action":"GROWING - Expand","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"On track"},{"salesman":"Austin","city":"VALDOSTA","custNum":200266,"customer":"FUSSELL TIRE & SERVICE","sales2025":248400.58,"sales2026":313280.74,"change":64880.16,"gpPct":0.075829,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: TRUCK TIRES (4.5%) | LOW GP: FARM TIRES (5.8%) | LOW GP: OFF THE ROAD TIRES (6.9%) | DOWN: OFF THE ROAD TIRES (-$23498) | LOST: OUTSIDE PURCHASE (was $1509) | LOST: FREIGHT CHARGES (was $275)"},{"salesman":"House","city":"ADEL","custNum":200401,"customer":"ADEL TIRE CO","sales2025":1322.48,"sales2026":3704.0,"change":2381.52,"gpPct":0.256061,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$159)"},{"salesman":"House","city":"ADEL","custNum":201037,"customer":"AFTER HOURS TIRE SERVICE","sales2025":8749.94,"sales2026":6006.96,"change":-2742.98,"gpPct":0.190521,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1919) | LOW GP: PASSENGER TIRES (7.3%) | DOWN: PASSENGER TIRES (-$724) | LOST: LAWN & GARDEN (was $19) | LOST: FARM TIRES (was $952)"},{"salesman":"House","city":"ADEL","custNum":200393,"customer":"BRUISER'S TIRE & TOWING","sales2025":0.0,"sales2026":721.58,"change":721.58,"gpPct":0.194074,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ADEL","custNum":200958,"customer":"BULLARD DIESEL & AUTO","sales2025":2708.98,"sales2026":1681.02,"change":-1027.96,"gpPct":0.157512,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $273) | LOST: INDUSTRIAL TIRES (was $145)"},{"salesman":"House","city":"ADEL","custNum":200416,"customer":"DENT'S SERVICE STATION","sales2025":13630.12,"sales2026":10221.92,"change":-3408.2,"gpPct":0.21874,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$883) | DOWN: PASSENGER TIRES (-$1939) | LOST: TUBES (was $4)"},{"salesman":"House","city":"ADEL","custNum":200647,"customer":"FAUSETTS TIRE CO.","sales2025":19193.02,"sales2026":-104.68,"change":-19297.7,"gpPct":0.0,"action":"NEGATIVE - Returns only","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$4336) | LOST: TUBES (was $391) | LOST: TRUCK TIRES (was $234) | LOST: LAWN & GARDEN (was $96) | LOST: FARM TIRES (was $43) | LOST: ST TRAILER (was $1813)"},{"salesman":"House","city":"ADEL","custNum":200980,"customer":"SLYDER'S GARAGE","sales2025":826.2,"sales2026":684.48,"change":-141.72,"gpPct":0.229298,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$287)"},{"salesman":"House","city":"ADEL","custNum":200822,"customer":"UNITED TIRE LLC","sales2025":140.52,"sales2026":0.0,"change":-140.52,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $130) | LOST: TUBES (was $11)"},{"salesman":"House","city":"ALAPAHA","custNum":101201,"customer":"TUCKERS SERVICE STATION","sales2025":65919.29,"sales2026":46305.24,"change":-19614.05,"gpPct":0.139573,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: RAD LT TRUCK (-$7568) | DOWN: FARM TIRES (-$10977) | DOWN: LAWN & GARDEN (-$252) | LOST: PATCHES AND REPAIR (was $203) | DOWN: TUBES (-$1664)"},{"salesman":"House","city":"ALBANY","custNum":201025,"customer":"A-1 WRECKER SERVICE","sales2025":10099.37,"sales2026":213.28,"change":-9886.09,"gpPct":0.278132,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$2244) | LOST: RAD LT TRUCK (was $6241) | LOST: ST TRAILER (was $769) | LOST: TRUCK TIRES (was $633)"},{"salesman":"House","city":"ALBANY","custNum":2000008,"customer":"ALBANY MOTORCARS","sales2025":-100.0,"sales2026":0.0,"change":100.0,"gpPct":0.0,"action":"MAINTAIN - Stable account","topDept":"None","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200675,"customer":"AUTO SOLUTIONS LLC","sales2025":234.24,"sales2026":0.0,"change":-234.24,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $234)"},{"salesman":"House","city":"ALBANY","custNum":200311,"customer":"AUTOMOTIVE NECESSITIES","sales2025":861.93,"sales2026":3490.16,"change":2628.23,"gpPct":0.014698,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (1.5%) | LOST: ST TRAILER (was $48)"},{"salesman":"House","city":"ALBANY","custNum":200600,"customer":"CHARLOT TRUCKING & TIRE SVC.","sales2025":0.0,"sales2026":925.74,"change":925.74,"gpPct":0.305561,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200870,"customer":"D&K USED TIRES","sales2025":484.48,"sales2026":3601.83,"change":3117.35,"gpPct":0.237149,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200411,"customer":"DJ'S CAR WASH","sales2025":3028.51,"sales2026":0.0,"change":-3028.51,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: ST TRAILER (was $114) | LOST: PASSENGER TIRES (was $1747) | LOST: RAD LT TRUCK (was $492) | LOST: FARM TIRES (was $675)"},{"salesman":"House","city":"ALBANY","custNum":201034,"customer":"ECONOMIC NICHOLAS TIRE","sales2025":874.62,"sales2026":407.92,"change":-466.7,"gpPct":0.206021,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $692)"},{"salesman":"House","city":"ALBANY","custNum":200655,"customer":"ECONOMY USED TIRE (ALBANY)","sales2025":4816.3,"sales2026":17345.24,"change":12528.94,"gpPct":0.106757,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: PASSENGER TIRES (7.4%) | LOST: PATCHES AND REPAIR (was $304) | LOST: TRUCK TIRES (was $457)"},{"salesman":"House","city":"ALBANY","custNum":200828,"customer":"HENRY'S ALIGNMNET","sales2025":2890.85,"sales2026":2566.58,"change":-324.27,"gpPct":0.158062,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1006)"},{"salesman":"House","city":"ALBANY","custNum":200750,"customer":"LAWN PERFORMANCE  LLC","sales2025":123.1,"sales2026":189.0,"change":65.9,"gpPct":0.279683,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":201056,"customer":"LEE'S AUTO SHOP","sales2025":4505.32,"sales2026":1428.8,"change":-3076.52,"gpPct":0.198929,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$1387) | LOW GP: PASSENGER TIRES (4.2%) | DOWN: PASSENGER TIRES (-$1689)"},{"salesman":"House","city":"ALBANY","custNum":200609,"customer":"LIBERTY AUTO CARE CENTER","sales2025":0.0,"sales2026":1491.38,"change":1491.38,"gpPct":0.305945,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200761,"customer":"MARIO NEW AND USED TIRE SHOP","sales2025":6685.85,"sales2026":4614.98,"change":-2070.87,"gpPct":0.211607,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$1468) | DOWN: RAD LT TRUCK (-$1176) | LOST: PATCHES AND REPAIR (was $18)"},{"salesman":"House","city":"ALBANY","custNum":101497,"customer":"PERFORMANCE MOTORSPORT","sales2025":317.7,"sales2026":285.98,"change":-31.72,"gpPct":0.30121,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $176)"},{"salesman":"House","city":"ALBANY","custNum":200357,"customer":"PONDER AUTO REPAIR","sales2025":504.64,"sales2026":2726.83,"change":2222.19,"gpPct":0.207769,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $421)"},{"salesman":"House","city":"ALBANY","custNum":200213,"customer":"PREMIER AUTOWORKS","sales2025":590.48,"sales2026":4237.99,"change":3647.51,"gpPct":0.090201,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$307) | LOW GP: RAD LT TRUCK (5.8%)"},{"salesman":"House","city":"ALBANY","custNum":200743,"customer":"SOUTHERN SALES & RENTALS  LLC","sales2025":1050.54,"sales2026":3035.81,"change":1985.27,"gpPct":0.255688,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200949,"customer":"SOUTHERN TIRE MART","sales2025":7067.4,"sales2026":7997.51,"change":930.11,"gpPct":0.160838,"action":"MAINTAIN - Stable account","topDept":"INDUSTRIAL TIRES","declinedDept":"FARM TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1532) | LOST: FARM TIRES (was $3868) | DOWN: TRUCK TIRES (-$1007)"},{"salesman":"House","city":"ALBANY","custNum":200546,"customer":"SOWEGA TIRE OF ALBANY","sales2025":5318.23,"sales2026":7235.16,"change":1916.93,"gpPct":0.137074,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.7%)"},{"salesman":"House","city":"ALBANY","custNum":200586,"customer":"STEEDLEY'S TRANSMISSION  INC","sales2025":1025.48,"sales2026":0.0,"change":-1025.48,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $1025)"},{"salesman":"House","city":"ALBANY","custNum":200136,"customer":"SUNBELT FORD ALBANY  INC (AMI)","sales2025":1481.0,"sales2026":0.0,"change":-1481.0,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1481)"},{"salesman":"House","city":"ALBANY","custNum":200756,"customer":"THE SHOP OF ALBANY  LLC","sales2025":4409.62,"sales2026":2945.24,"change":-1464.38,"gpPct":0.208353,"action":"DECLINING - Investigate","topDept":"ST TRAILER","declinedDept":"PASSENGER TIRES","focus":"DOWN: TRUCK TIRES (-$1271) | DOWN: PASSENGER TIRES (-$1603)"},{"salesman":"House","city":"ALBANY","custNum":200795,"customer":"WINCHESTER PAINT & BODY","sales2025":1798.9,"sales2026":198.0,"change":-1600.9,"gpPct":0.103333,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: ST TRAILER (was $114) | DOWN: PASSENGER TIRES (-$260) | LOST: RAD LT TRUCK (was $1226)"},{"salesman":"House","city":"ALLENTOWN","custNum":200974,"customer":"PRIORITY TIRE (TIFTON WHSE)","sales2025":28248.53,"sales2026":38655.64,"change":10407.11,"gpPct":0.127464,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (6.6%) | LOST: TRUCK TIRES (was $5027) | DOWN: ST TRAILER (-$284) | LOW GP: INDUSTRIAL TIRES (9.9%)"},{"salesman":"House","city":"ALMA","custNum":200897,"customer":"ALMA TIRE & AUTO REPAIR","sales2025":6574.8,"sales2026":6299.87,"change":-274.93,"gpPct":0.109896,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (8.4%) | DOWN: RAD LT TRUCK (-$3328) | LOW GP: FARM TIRES (1.9%) | LOST: LAWN & GARDEN (was $593)"},{"salesman":"House","city":"ALMA","custNum":101476,"customer":"HART'S SERVICE STATION","sales2025":1555.84,"sales2026":284.91,"change":-1270.93,"gpPct":0.271559,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1271)"},{"salesman":"House","city":"ALMA","custNum":201015,"customer":"LUBE KING & TIRES","sales2025":6844.85,"sales2026":10514.55,"change":3669.7,"gpPct":0.13906,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALMA","custNum":200985,"customer":"ROBERTS AUTO SERVICE LLC","sales2025":902.32,"sales2026":0.0,"change":-902.32,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $282) | LOST: PASSENGER TIRES (was $621)"},{"salesman":"House","city":"ARABI","custNum":200962,"customer":"GREENE'S TIRE SERVICE LLC","sales2025":8446.25,"sales2026":15056.15,"change":6609.9,"gpPct":0.166669,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$5552) | DOWN: TUBES (-$231)"},{"salesman":"House","city":"ASHBURN","custNum":201007,"customer":"SHORTY HUGHES TRUCKING  LLC","sales2025":33282.4,"sales2026":16949.89,"change":-16332.51,"gpPct":0.143903,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$16333) | LOST: ST TRAILER (was $35) | LOST: MOUNTING LUBE (was $39)"},{"salesman":"House","city":"ASHBURN","custNum":200803,"customer":"SOUTH MAIN GARAGE","sales2025":3062.17,"sales2026":2710.57,"change":-351.6,"gpPct":0.151027,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$575) | DOWN: RAD LT TRUCK (-$1202)"},{"salesman":"House","city":"BAINBRIDGE","custNum":200162,"customer":"DELTA TIRE CO","sales2025":67451.74,"sales2026":58318.3,"change":-9133.44,"gpPct":0.127503,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"FARM TIRES","focus":"LOST: TRUCK TIRES (was $6004) | LOST: TUBES (was $579) | LOST: ST TRAILER (was $1647)"},{"salesman":"House","city":"BAINBRIDGE","custNum":2000010,"customer":"SOUTHERN AUTOMOTIVE SVC & REP","sales2025":2938.36,"sales2026":365.0,"change":-2573.36,"gpPct":0.044164,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.4%) | DOWN: RAD LT TRUCK (-$2573)"},{"salesman":"House","city":"BAINBRIDGE","custNum":200230,"customer":"SOUTHERN TIRE & BATTERY","sales2025":1872.13,"sales2026":375.06,"change":-1497.07,"gpPct":0.323468,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1429) | LOST: PASSENGER TIRES (was $68)"},{"salesman":"House","city":"BAINBRIDGE","custNum":200297,"customer":"WHOLESALE BATTERY","sales2025":152.36,"sales2026":0.0,"change":-152.36,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $152)"},{"salesman":"House","city":"BENICIA","custNum":201038,"customer":"GIGA TIRES  LLC (TIFTON WHSE)","sales2025":47481.01,"sales2026":39512.57,"change":-7968.44,"gpPct":0.022541,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.9%) | LOW GP: TRUCK TIRES (8.4%) | DOWN: TRUCK TIRES (-$14513) | LOW GP: ST TRAILER (8.8%)"},{"salesman":"House","city":"BLACKSHEAR","custNum":200714,"customer":"C&S AUTO SERVICE INC.","sales2025":158.65,"sales2026":0.0,"change":-158.65,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TUBES","focus":"LOST: TUBES (was $129) | LOST: ST TRAILER (was $30)"},{"salesman":"House","city":"BLACKSHEAR","custNum":200465,"customer":"DIXON SERVICE CENTER","sales2025":0.0,"sales2026":414.49,"change":414.49,"gpPct":0.196362,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"BLACKSHEAR","custNum":201036,"customer":"GODWIN & SON REPAIR & SALES","sales2025":845.74,"sales2026":662.7,"change":-183.04,"gpPct":0.21005,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$241) | LOST: ST TRAILER (was $59)"},{"salesman":"House","city":"BROXTON","custNum":201028,"customer":"CLEMENT USED TIRES","sales2025":22854.52,"sales2026":20077.27,"change":-2777.25,"gpPct":0.165107,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $1478) | LOST: TRUCK TIRES (was $239)"},{"salesman":"House","city":"BROXTON","custNum":101364,"customer":"KENNY'S AUTO AND TRUCK SALVAGE","sales2025":100.91,"sales2026":810.12,"change":709.21,"gpPct":0.185787,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"CAIRO","custNum":200148,"customer":"AUTO & TRUCK CARE SPECIAL","sales2025":481.6,"sales2026":465.1,"change":-16.5,"gpPct":0.278779,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"House","city":"CAIRO","custNum":200146,"customer":"BRACEWELL AUTOMOTIVE SERVICE","sales2025":7653.78,"sales2026":10001.72,"change":2347.94,"gpPct":0.075717,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (7.9%) | LOW GP: PASSENGER TIRES (7.2%) | LOST: ST TRAILER (was $231) | LOST: TUBES (was $13)"},{"salesman":"House","city":"CAIRO","custNum":200658,"customer":"POWE AUTOMOTIVE","sales2025":3903.76,"sales2026":3289.68,"change":-614.08,"gpPct":0.046558,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (4.3%) | DOWN: PASSENGER TIRES (-$1909) | LOW GP: RAD LT TRUCK (4.7%)"},{"salesman":"House","city":"CAIRO","custNum":200287,"customer":"RIDLEY'S AUTOMOTIVE","sales2025":1349.64,"sales2026":399.16,"change":-950.48,"gpPct":0.20493,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$950)"},{"salesman":"House","city":"CAMILLA","custNum":200396,"customer":"PATE TIRE & SERVICE LLC","sales2025":8129.47,"sales2026":4799.24,"change":-3330.23,"gpPct":0.146461,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: TRUCK TIRES (-$1695) | LOST: RAD LT TRUCK (was $2041) | LOST: ST TRAILER (was $338) | LOST: TUBES (was $222) | LOW GP: OFF THE ROAD TIRES (8.4%) | LOST: FARM TIRES (was $833)"},{"salesman":"House","city":"CAMILLA","custNum":101545,"customer":"TOMMY'S TIRE","sales2025":1585.52,"sales2026":0.0,"change":-1585.52,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1586) | LOST: FARM TIRES (was $1047)"},{"salesman":"House","city":"CHICAGO","custNum":2000014,"customer":"UNITED TIRES ONLINE SALES -T","sales2025":0.0,"sales2026":9462.85,"change":9462.85,"gpPct":0.077144,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (9.8%) | LOW GP: PASSENGER TIRES (5.1%)"},{"salesman":"House","city":"CLIMAX","custNum":201002,"customer":"ZAPATA'S TIRE","sales2025":27270.87,"sales2026":19086.14,"change":-8184.73,"gpPct":0.097664,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $79) | LOW GP: TRUCK TIRES (8.5%) | DOWN: TRUCK TIRES (-$8206) | LOST: MOUNTING LUBE (was $323)"},{"salesman":"House","city":"COOLIDGE","custNum":201073,"customer":"GTO TIRE SERVICE & AUTO LLC","sales2025":1058.9,"sales2026":0.0,"change":-1058.9,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $663) | LOST: RAD LT TRUCK (was $396)"},{"salesman":"House","city":"CORDELE","custNum":200562,"customer":"BEST CARS OF CORDELE  LLC","sales2025":459.26,"sales2026":805.82,"change":346.56,"gpPct":0.307488,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"CORDELE","custNum":200912,"customer":"LANE'S TRK & TRL REPAIR & AUTO","sales2025":1258.7,"sales2026":715.68,"change":-543.02,"gpPct":0.202046,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $538)"},{"salesman":"House","city":"CORDELE","custNum":200866,"customer":"LEMUS TIRE SHOP","sales2025":8061.88,"sales2026":16629.32,"change":8567.44,"gpPct":0.21822,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"WHEEL WEIGHTS","focus":"DOWN: VALVE STEMS (-$12) | DOWN: WHEEL WEIGHTS (-$35)"},{"salesman":"House","city":"CORDELE","custNum":100741,"customer":"MASSEY'S MUFFLER","sales2025":4645.15,"sales2026":4263.76,"change":-381.39,"gpPct":0.214752,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: LAWN & GARDEN (was $197) | DOWN: RAD LT TRUCK (-$1339) | LOST: PATCHES AND REPAIR (was $4)"},{"salesman":"House","city":"CORDELE","custNum":101208,"customer":"MIKE FRASER AUTO REPAIR","sales2025":14791.65,"sales2026":9379.01,"change":-5412.64,"gpPct":0.148423,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5691) | LOW GP: TRUCK TIRES (5.9%)"},{"salesman":"House","city":"CORDELE","custNum":200940,"customer":"PERRY BROS. OIL (CORDELE)","sales2025":5271.96,"sales2026":3598.96,"change":-1673.0,"gpPct":0.116717,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $692) | LOW GP: TRUCK TIRES (8.9%)"},{"salesman":"House","city":"CORDELE","custNum":200890,"customer":"PMT TRK. TRAILER & TIRE REPAIR","sales2025":25786.49,"sales2026":21665.54,"change":-4120.95,"gpPct":0.104638,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.7%) | DOWN: TRUCK TIRES (-$7789) | LOW GP: RAD LT TRUCK (9.7%) | LOST: OFF THE ROAD TIRES (was $968)"},{"salesman":"House","city":"CUTHBERT","custNum":200506,"customer":"DEVANE TIRE & SERVICE LLC","sales2025":1664.23,"sales2026":1989.37,"change":325.14,"gpPct":0.111302,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (9.9%)"},{"salesman":"House","city":"DAWSON","custNum":200664,"customer":"ABR COMMERCIAL TRUCK & AUTO","sales2025":11800.58,"sales2026":3687.33,"change":-8113.25,"gpPct":0.128817,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7413) | LOST: RAD LT TRUCK (was $459) | LOST: ST TRAILER (was $249)"},{"salesman":"House","city":"DAWSON","custNum":200719,"customer":"FOSTER EASY PAY TIRE CO.  INC.","sales2025":38940.96,"sales2026":66703.46,"change":27762.5,"gpPct":0.150759,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"TUBES","focus":"LOW GP: TRUCK TIRES (9.5%) | LOST: TUBES (was $113) | LOW GP: OFF THE ROAD TIRES (9.0%)"},{"salesman":"House","city":"DOERUN","custNum":200193,"customer":"MCLEAN TIRES INC","sales2025":21691.31,"sales2026":36327.33,"change":14636.02,"gpPct":0.150914,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"DOWN: OFF THE ROAD TIRES (-$582)"},{"salesman":"House","city":"DOTHAN","custNum":200927,"customer":"MASTER TIRE","sales2025":809.44,"sales2026":322.39,"change":-487.05,"gpPct":0.073451,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (7.3%) | DOWN: RAD LT TRUCK (-$487)"},{"salesman":"House","city":"DOTHAN","custNum":200217,"customer":"SCOTT STEVENS TIRE & SERVICE","sales2025":158.78,"sales2026":0.0,"change":-158.78,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $159)"},{"salesman":"House","city":"DOTHAN","custNum":200241,"customer":"THE RIM SHOP INC","sales2025":642.55,"sales2026":179.71,"change":-462.84,"gpPct":0.110845,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $404)"},{"salesman":"House","city":"DOTHAN","custNum":200132,"customer":"WILKS A-ONE TIRE SALES","sales2025":8556.55,"sales2026":14155.12,"change":5598.57,"gpPct":0.108225,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $3924) | DOWN: OFF THE ROAD TIRES (-$1399) | LOST: RAD LT TRUCK (was $1064) | LOST: ST TRAILER (was $373)"},{"salesman":"House","city":"DOUGLAS","custNum":101366,"customer":"B & M AUTOMOTIVE SERVICE","sales2025":4573.22,"sales2026":457.85,"change":-4115.37,"gpPct":0.155684,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$1105) | LOST: RAD LT TRUCK (was $3010)"},{"salesman":"House","city":"DOUGLAS","custNum":101164,"customer":"DAVIS TIRE (DOUGLAS)","sales2025":7653.49,"sales2026":19645.52,"change":11992.03,"gpPct":0.165859,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"DOUGLAS","custNum":200648,"customer":"JOE'S AUTO REPAIR  LLC","sales2025":1295.76,"sales2026":2639.02,"change":1343.26,"gpPct":0.148961,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"House","city":"DOUGLAS","custNum":200720,"customer":"LUBE MASTERS","sales2025":0.0,"sales2026":968.36,"change":968.36,"gpPct":0.258251,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"EDISON","custNum":200166,"customer":"EDISON TIRE","sales2025":40661.64,"sales2026":85501.32,"change":44839.68,"gpPct":0.130421,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (9.7%) | DOWN: INDUSTRIAL TIRES (-$1177)"},{"salesman":"House","city":"EL SEGUNDO","custNum":201040,"customer":"TIRES EASY (NAP - TIFTON)","sales2025":30622.47,"sales2026":31847.11,"change":1224.64,"gpPct":0.125274,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$13316)"},{"salesman":"House","city":"EL SEGUNDO","custNum":200976,"customer":"TIRESEASY-LLC (TIFTON WHSE)","sales2025":237576.35,"sales2026":1023674.71,"change":786098.36,"gpPct":0.141912,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.1%) | LOW GP: INDUSTRIAL TIRES (9.2%) | LOW GP: FARM TIRES (0.9%)"},{"salesman":"House","city":"ELLENTON","custNum":101512,"customer":"ELLENTON TIRE AND AUTO","sales2025":95690.48,"sales2026":98730.03,"change":3039.55,"gpPct":0.124228,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (5.9%) | LOW GP: PASSENGER TIRES (3.5%) | DOWN: PASSENGER TIRES (-$5773) | DOWN: OFF THE ROAD TIRES (-$1411) | LOST: VALVE STEMS (was $14) | LOST: LAWN & GARDEN (was $16) | LOST: TIRE TOOLS (was $23)"},{"salesman":"House","city":"FLORHAM PARK","custNum":201041,"customer":"TIRETREADS LLC (TIFTON ACCT)","sales2025":535.53,"sales2026":704.53,"change":169.0,"gpPct":0.261834,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$338)"},{"salesman":"House","city":"FRIDLEY","custNum":201074,"customer":"TIRE DEPOT CO. - TAG (TIFTON)","sales2025":22516.59,"sales2026":25265.85,"change":2749.26,"gpPct":0.131769,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (7.0%)"},{"salesman":"House","city":"HAHIRA","custNum":201064,"customer":"CHAD'S AUTO REPAIR","sales2025":733.68,"sales2026":2375.68,"change":1642.0,"gpPct":0.111092,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $193) | LOW GP: RAD LT TRUCK (5.9%)"},{"salesman":"House","city":"HOMERVILLE","custNum":200391,"customer":"CLINCH BRAKE & ALIGNMENT","sales2025":0.0,"sales2026":1708.81,"change":1708.81,"gpPct":0.259175,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"HOMERVILLE","custNum":200806,"customer":"WALKERS AUTO & OUTDOOR  INC","sales2025":5368.2,"sales2026":1654.18,"change":-3714.02,"gpPct":0.236462,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$1835) | DOWN: PASSENGER TIRES (-$1879)"},{"salesman":"House","city":"JACKSONVILLE","custNum":200889,"customer":"CONLAN TIRE CO.","sales2025":1124.56,"sales2026":0.0,"change":-1124.56,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $1125)"},{"salesman":"House","city":"JACKSONVILLE","custNum":201070,"customer":"SNIDER INDUSTRIAL","sales2025":10603.01,"sales2026":3184.84,"change":-7418.17,"gpPct":0.147901,"action":"DECLINING - Investigate","topDept":"INDUSTRIAL TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $7401) | LOST: RAD LT TRUCK (was $908) | LOST: PASSENGER TIRES (was $412) | LOST: ST TRAILER (was $444) | LOST: FARM TIRES (was $1437)"},{"salesman":"House","city":"LAKE CITY","custNum":201032,"customer":"A-1 TIRE PLUS","sales2025":51915.97,"sales2026":32719.77,"change":-19196.2,"gpPct":0.139234,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: RAD LT TRUCK (-$4760) | DOWN: PASSENGER TIRES (-$6764) | DOWN: TRUCK TIRES (-$10107)"},{"salesman":"House","city":"LAKE CITY","custNum":201058,"customer":"AFTER 5 COMM. TIRE & OFF ROAD","sales2025":16909.88,"sales2026":1742.9,"change":-15166.98,"gpPct":0.229307,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7204) | LOST: TUBES (was $368) | LOST: ST TRAILER (was $2343) | DOWN: RAD LT TRUCK (-$2364) | LOST: INDUSTRIAL TIRES (was $488) | LOST: PASSENGER TIRES (was $2400)"},{"salesman":"House","city":"LAKE CITY","custNum":200916,"customer":"GATEWAY DIESEL  AUTO & MOBILE","sales2025":2187.3,"sales2026":328.7,"change":-1858.6,"gpPct":0.122117,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: ST TRAILER (was $124) | LOST: PASSENGER TIRES (was $637) | DOWN: TRUCK TIRES (-$1098)"},{"salesman":"House","city":"LAKE CITY","custNum":201048,"customer":"RRO 24 HR ROADSIDE ASSISTANCE","sales2025":4777.38,"sales2026":5948.25,"change":1170.87,"gpPct":0.163453,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: TRUCK TIRES (8.7%) | DOWN: PASSENGER TIRES (-$873) | LOST: INDUSTRIAL TIRES (was $491) | LOST: FARM TIRES (was $439)"},{"salesman":"House","city":"LAKE CITY","custNum":200895,"customer":"TIRE MART OF LAKE CITY","sales2025":1329.48,"sales2026":2057.28,"change":727.8,"gpPct":0.014971,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (1.5%)"},{"salesman":"House","city":"LAKE PARK","custNum":2000004,"customer":"ODELL AUTOMOTIVE","sales2025":0.0,"sales2026":185.0,"change":185.0,"gpPct":0.176865,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"LAKELAND","custNum":200315,"customer":"BOBBY'S CITGO","sales2025":8919.33,"sales2026":13995.61,"change":5076.28,"gpPct":0.116465,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: PASSENGER TIRES (8.3%)"},{"salesman":"House","city":"LEESBURG","custNum":101525,"customer":"BMS DISCOUNT TIRES","sales2025":11979.1,"sales2026":15245.75,"change":3266.65,"gpPct":0.166365,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1120)"},{"salesman":"House","city":"LEESBURG","custNum":200191,"customer":"LEE COUNTY AUTO SERVICE","sales2025":1727.2,"sales2026":6424.34,"change":4697.14,"gpPct":0.112178,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (10.0%)"},{"salesman":"House","city":"LEESBURG","custNum":101524,"customer":"MASTER BODY WORKS","sales2025":322.1,"sales2026":1939.64,"change":1617.54,"gpPct":0.144501,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"LEESBURG","custNum":200192,"customer":"MCGEHEE'S TIRE & AUTO","sales2025":0.0,"sales2026":295.64,"change":295.64,"gpPct":0.073332,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (0.8%)"},{"salesman":"House","city":"LENOX","custNum":200417,"customer":"HIGHWAY TIRE & DIESEL","sales2025":549.8,"sales2026":3343.13,"change":2793.33,"gpPct":0.185012,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$309)"},{"salesman":"House","city":"LENOX","custNum":200462,"customer":"QUALITY FEEDSTUFFS  INC","sales2025":3260.63,"sales2026":1219.31,"change":-2041.32,"gpPct":0.18075,"action":"DECLINING - Investigate","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2230) | LOST: WHEELS (was $175) | LOST: ST TRAILER (was $120)"},{"salesman":"House","city":"LIVE OAK","custNum":201018,"customer":"CRAWLEY'S AUTOMOTIVE & TIRE","sales2025":2396.54,"sales2026":5512.19,"change":3115.65,"gpPct":0.185153,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"House","city":"LIVE OAK","custNum":201008,"customer":"DBJ MOBILE TIRE SERVICE  INC.","sales2025":16806.88,"sales2026":22154.37,"change":5347.49,"gpPct":0.149361,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $739) | DOWN: RAD LT TRUCK (-$660)"},{"salesman":"House","city":"LIVE OAK","custNum":200900,"customer":"GILLETTES AUTO","sales2025":1872.92,"sales2026":782.06,"change":-1090.86,"gpPct":0.235289,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$384) | LOST: ST TRAILER (was $277) | DOWN: PASSENGER TIRES (-$429)"},{"salesman":"House","city":"LIVE OAK","custNum":200913,"customer":"LASHLEY'S HOMETOWN TIRE LLC","sales2025":11758.34,"sales2026":10021.31,"change":-1737.03,"gpPct":0.177108,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $105) | LOW GP: RAD LT TRUCK (0.9%) | DOWN: RAD LT TRUCK (-$1434) | LOST: LAWN & GARDEN (was $250) | LOST: ST TRAILER (was $298)"},{"salesman":"House","city":"LIVE OAK","custNum":200898,"customer":"PRECISION AUTO & MUFFLER LLC","sales2025":1917.68,"sales2026":2315.8,"change":398.12,"gpPct":0.255212,"action":"MAINTAIN - Stable account","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1918)"},{"salesman":"House","city":"MADISON","custNum":2000038,"customer":"FAST TIRE SERVICE","sales2025":0.0,"sales2026":4663.36,"change":4663.36,"gpPct":0.071108,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.1%)"},{"salesman":"House","city":"MADISON","custNum":200671,"customer":"STEWARTS AUTO SERVICE CENTER","sales2025":0.0,"sales2026":1130.95,"change":1130.95,"gpPct":0.248473,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MADISON","custNum":200624,"customer":"WALLACE MOTORS","sales2025":524.7,"sales2026":3721.7,"change":3197.0,"gpPct":0.122076,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (6.0%)"},{"salesman":"House","city":"MONTICELLO","custNum":200753,"customer":"AUTO TECH OF MIAMI INC.","sales2025":1043.28,"sales2026":2110.5,"change":1067.22,"gpPct":0.234996,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: WHEELS (was $44)"},{"salesman":"House","city":"MONTICELLO","custNum":2000007,"customer":"KEATON & SON TIRE LLC","sales2025":2159.61,"sales2026":3931.7,"change":1772.09,"gpPct":0.203947,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $60)"},{"salesman":"House","city":"MOULTRIE","custNum":200307,"customer":"ARREDONDO TIRE SERVICE","sales2025":3593.62,"sales2026":2635.99,"change":-957.63,"gpPct":0.219011,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$976) | DOWN: ST TRAILER (-$492)"},{"salesman":"House","city":"MOULTRIE","custNum":200690,"customer":"COLQUITT COUNTY TIRE LLC","sales2025":73373.11,"sales2026":96225.39,"change":22852.28,"gpPct":0.1467,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$2242)"},{"salesman":"House","city":"MOULTRIE","custNum":201003,"customer":"NICHOLAS TIRES INC.","sales2025":26313.34,"sales2026":20208.3,"change":-6105.04,"gpPct":0.227654,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$4433) | DOWN: ST TRAILER (-$1354) | LOST: TRUCK TIRES (was $217)"},{"salesman":"House","city":"MOULTRIE","custNum":200975,"customer":"SANTOS TIRE SHOP","sales2025":6083.42,"sales2026":11118.93,"change":5035.51,"gpPct":0.225541,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"On track"},{"salesman":"House","city":"MOULTRIE","custNum":200427,"customer":"SAUNDERS AUTO REPAIR","sales2025":0.0,"sales2026":655.6,"change":655.6,"gpPct":0.213362,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MOULTRIE","custNum":200896,"customer":"SOUTHERN AUTO SPECIALIST","sales2025":4221.81,"sales2026":5656.71,"change":1434.9,"gpPct":0.205301,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$353) | LOST: TUBES (was $20)"},{"salesman":"House","city":"MOULTRIE","custNum":200947,"customer":"SUNSET TIRE & AUTOMOTIVE","sales2025":9751.94,"sales2026":12516.71,"change":2764.77,"gpPct":0.238654,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MOULTRIE","custNum":200829,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2025":108954.72,"sales2026":31570.57,"change":-77384.15,"gpPct":0.111708,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.4%) | DOWN: TRUCK TIRES (-$59938) | LOST: FARM TIRES (was $8231) | LOST: OFF THE ROAD TIRES (was $5859) | DOWN: ST TRAILER (-$2049) | LOST: TUBES (was $131) | LOST: INDUSTRIAL TIRES (was $1713) | LOST: LAWN & GARDEN (was $32)"},{"salesman":"House","city":"MOULTRIE","custNum":200537,"customer":"TONY'S TIRE & ROAD SERVICE INC","sales2025":10689.57,"sales2026":16471.9,"change":5782.33,"gpPct":0.163593,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.0%) | DOWN: TRUCK TIRES (-$4816) | LOST: VALVE STEMS (was $38)"},{"salesman":"House","city":"NASHVILLE","custNum":200941,"customer":"D&S WHEELS & DEALS LLC","sales2025":11668.0,"sales2026":8503.69,"change":-3164.31,"gpPct":0.12359,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$4022)"},{"salesman":"House","city":"NASHVILLE","custNum":200451,"customer":"HARROD BROTHERS","sales2025":7588.56,"sales2026":6725.22,"change":-863.34,"gpPct":0.109879,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: ST TRAILER (-$426) | LOST: TRUCK TIRES (was $2315)"},{"salesman":"House","city":"NASHVILLE","custNum":200622,"customer":"MOORE'S ACCESSORIES & OFFROAD","sales2025":2802.91,"sales2026":20.86,"change":-2782.05,"gpPct":0.298178,"action":"DECLINING - Investigate","topDept":"TUBES","declinedDept":"RAD LT TRUCK","focus":"DOWN: TUBES (-$122) | LOST: RAD LT TRUCK (was $2660)"},{"salesman":"House","city":"NASHVILLE","custNum":101549,"customer":"NASHVILLE TIRE","sales2025":31130.87,"sales2026":32218.79,"change":1087.92,"gpPct":0.22583,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: TUBES (-$514) | LOST: TRUCK TIRES (was $630) | LOST: INDUSTRIAL TIRES (was $595)"},{"salesman":"House","city":"NEW YORK","custNum":201042,"customer":"TIRE AGENT CORP (TIFTON WHS)","sales2025":13547.23,"sales2026":13650.55,"change":103.32,"gpPct":0.089776,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.2%) | LOW GP: PASSENGER TIRES (8.4%)"},{"salesman":"House","city":"NORMAN PARK","custNum":200642,"customer":"E.G. AUTO SALES","sales2025":2432.66,"sales2026":1029.34,"change":-1403.32,"gpPct":0.245147,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$432) | DOWN: PASSENGER TIRES (-$905) | LOST: TUBES (was $10) | LOST: FARM TIRES (was $57)"},{"salesman":"House","city":"OMEGA","custNum":101439,"customer":"A.T. TIRE SERVICE","sales2025":21566.3,"sales2026":35950.45,"change":14384.15,"gpPct":0.154018,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $29)"},{"salesman":"House","city":"OMEGA","custNum":201039,"customer":"CERVANTES AUTO SALES","sales2025":1006.94,"sales2026":984.59,"change":-22.35,"gpPct":0.170203,"action":"MAINTAIN - Stable account","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $593)"},{"salesman":"House","city":"ORLANDO","custNum":2000005,"customer":"GOODYEAR COMMERCIAL TIRE & SVC","sales2025":1563.2,"sales2026":0.0,"change":-1563.2,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1563)"},{"salesman":"House","city":"PELHAM","custNum":200695,"customer":"PELHAM TIRE & EQUIPMENT COMP.","sales2025":883.24,"sales2026":0.0,"change":-883.24,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $883)"},{"salesman":"House","city":"PELHAM","custNum":2000013,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2025":0.0,"sales2026":15138.88,"change":15138.88,"gpPct":0.108533,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: FARM TIRES (9.8%) | LOW GP: INDUSTRIAL TIRES (8.0%)"},{"salesman":"House","city":"PLANT CITY","custNum":2000028,"customer":"SUNPOINT TIRES & ROAD SERVICE","sales2025":0.0,"sales2026":1957.92,"change":1957.92,"gpPct":0.100168,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.5%)"},{"salesman":"House","city":"POULAN","custNum":200504,"customer":"PLATINUM RECOVERY SERVICES LLC","sales2025":0.0,"sales2026":163.42,"change":163.42,"gpPct":0.212826,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"QUINCY","custNum":200585,"customer":"QUINCY TIRE AND RECAPPING","sales2025":54873.73,"sales2026":45997.69,"change":-8876.04,"gpPct":0.202198,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$15390)"},{"salesman":"House","city":"QUINCY","custNum":200599,"customer":"W&L TIRE & WHEEL CO. INC.","sales2025":11568.58,"sales2026":7789.26,"change":-3779.32,"gpPct":0.219145,"action":"DECLINING - Investigate","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5137) | LOST: PASSENGER TIRES (was $2248) | LOW GP: INDUSTRIAL TIRES (7.6%) | DOWN: INDUSTRIAL TIRES (-$1634) | LOST: OFF THE ROAD TIRES (was $132)"},{"salesman":"House","city":"QUITMAN","custNum":101566,"customer":"HARVEY'S GARAGE & MUFFLER","sales2025":200.0,"sales2026":0.0,"change":-200.0,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"DISCOUNTS/COUPONS","focus":"LOST: DISCOUNTS/COUPONS (was $200)"},{"salesman":"House","city":"QUITMAN","custNum":200356,"customer":"NEELY'S SERVICE CENTER","sales2025":3484.79,"sales2026":1634.86,"change":-1849.93,"gpPct":0.152588,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1083) | DOWN: RAD LT TRUCK (-$755) | LOST: ST TRAILER (was $114)"},{"salesman":"House","city":"QUITMAN","custNum":201023,"customer":"PEASE ON THE GO 24/7","sales2025":578.01,"sales2026":2306.76,"change":1728.75,"gpPct":0.232668,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ROCHELLE","custNum":101530,"customer":"MARTIN TIRE SERVICE","sales2025":27899.12,"sales2026":26123.01,"change":-1776.11,"gpPct":0.123028,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"FARM TIRES","focus":"DOWN: TUBES (-$449) | LOST: LAWN & GARDEN (was $167) | DOWN: RAD LT TRUCK (-$3718) | DOWN: FARM TIRES (-$7124) | LOW GP: WHEELS (5.3%)"},{"salesman":"House","city":"ROCHELLE","custNum":200880,"customer":"R&R TIRE CO.","sales2025":21415.62,"sales2026":12025.77,"change":-9389.85,"gpPct":0.112147,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$12875) | LOW GP: RAD LT TRUCK (9.0%) | DOWN: FARM TIRES (-$921)"},{"salesman":"House","city":"ROCHELLE","custNum":101025,"customer":"ROCHELLE TIRE","sales2025":28349.41,"sales2026":55859.07,"change":27509.66,"gpPct":0.142816,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"WHEELS","focus":"LOST: WHEELS (was $287)"},{"salesman":"House","city":"ROCHELLE","custNum":101588,"customer":"STEPHENS BROTHERS","sales2025":9897.93,"sales2026":13343.89,"change":3445.96,"gpPct":0.141229,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $811) | DOWN: ST TRAILER (-$323) | LOST: TUBES (was $39) | LOST: FARM TIRES (was $613)"},{"salesman":"House","city":"SYCAMORE","custNum":101305,"customer":"SYCAMORE SALES & SALVAGE LLC","sales2025":3229.76,"sales2026":2357.16,"change":-872.6,"gpPct":0.280931,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$177) | LOST: FARM TIRES (was $351)"},{"salesman":"House","city":"SYLVESTER","custNum":200681,"customer":"CITY OF SYLVESTER","sales2025":2801.44,"sales2026":3211.71,"change":410.27,"gpPct":0.103593,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"LOST: RAD LT TRUCK (was $427) | LOW GP: PASSENGER TIRES (9.6%) | LOW GP: TRUCK TIRES (6.2%) | LOST: INDUSTRIAL TIRES (was $431) | LOST: FARM TIRES (was $566)"},{"salesman":"House","city":"SYLVESTER","custNum":201052,"customer":"PRECISION DIESEL REPAIR LLC","sales2025":10037.09,"sales2026":14585.6,"change":4548.51,"gpPct":0.164974,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"DOWN: PASSENGER TIRES (-$408) | LOST: FARM TIRES (was $916)"},{"salesman":"House","city":"SYLVESTER","custNum":200715,"customer":"R&M AUTO TRUCKING  INC","sales2025":3349.52,"sales2026":3481.75,"change":132.23,"gpPct":0.207203,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$783)"},{"salesman":"House","city":"SYLVESTER","custNum":200860,"customer":"SUNBELT FORD INC. (AMI ACCT)","sales2025":17.54,"sales2026":0.0,"change":-17.54,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $169)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200979,"customer":"DISCOUNT TIRE & AUTO SHOP","sales2025":10139.57,"sales2026":2907.25,"change":-7232.32,"gpPct":0.250329,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$5029) | DOWN: RAD LT TRUCK (-$2083) | LOST: ST TRAILER (was $120)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200998,"customer":"FRIENDLY AUTO SALES","sales2025":541.86,"sales2026":0.0,"change":-541.86,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $542)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200990,"customer":"PATTON'S ALIGNMENT & BRAKE SVC","sales2025":1554.4,"sales2026":928.0,"change":-626.4,"gpPct":0.039526,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.0%) | DOWN: RAD LT TRUCK (-$626)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200991,"customer":"THE TIRE CENTRE OF FLORIDA LLC","sales2025":18073.68,"sales2026":29694.27,"change":11620.59,"gpPct":0.195044,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$1724) | LOST: INDUSTRIAL TIRES (was $57) | LOST: TUBES (was $10) | LOW GP: OFF THE ROAD TIRES (7.2%)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200796,"customer":"TRUCK N CAR CONCEPTS","sales2025":1915.96,"sales2026":0.0,"change":-1915.96,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1916)"},{"salesman":"House","city":"THOMASVILLE","custNum":201012,"customer":"AG PRO FUEL","sales2025":521.54,"sales2026":883.37,"change":361.83,"gpPct":0.075076,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.1%) | LOW GP: TRUCK TIRES (7.9%)"},{"salesman":"House","city":"THOMASVILLE","custNum":200641,"customer":"AUTO AIR OF THOMASVILLE","sales2025":54.85,"sales2026":0.0,"change":-54.85,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $55)"},{"salesman":"House","city":"THOMASVILLE","custNum":200607,"customer":"B AND B SERVICE CENTER  INC.","sales2025":4777.55,"sales2026":6553.8,"change":1776.25,"gpPct":0.240038,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1134) | LOW GP: INDUSTRIAL TIRES (9.0%)"},{"salesman":"House","city":"THOMASVILLE","custNum":200410,"customer":"EZDEALIN WHEELS AND TIRES","sales2025":121090.53,"sales2026":128634.19,"change":7543.66,"gpPct":0.20616,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (7.0%) | DOWN: PASSENGER TIRES (-$2189)"},{"salesman":"House","city":"THOMASVILLE","custNum":200526,"customer":"GERMAN IMPORT SERVICE","sales2025":70.04,"sales2026":0.0,"change":-70.04,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"ST TRAILER","focus":"LOST: TUBES (was $8) | LOST: ST TRAILER (was $62)"},{"salesman":"House","city":"THOMASVILLE","custNum":200277,"customer":"IMPORT SERVICE & SALES","sales2025":9234.05,"sales2026":8008.69,"change":-1225.36,"gpPct":0.204412,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$840)"},{"salesman":"House","city":"THOMASVILLE","custNum":200358,"customer":"PONDER'S AUTOMOTIVE INC","sales2025":9398.3,"sales2026":5792.14,"change":-3606.16,"gpPct":0.048901,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (3.9%) | DOWN: PASSENGER TIRES (-$3496) | LOW GP: RAD LT TRUCK (6.3%) | LOST: ST TRAILER (was $347)"},{"salesman":"House","city":"TIFTON","custNum":200759,"customer":"AADCO","sales2025":537.48,"sales2026":922.43,"change":384.95,"gpPct":0.276834,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$476)"},{"salesman":"House","city":"TIFTON","custNum":200709,"customer":"ASHLEY'S AUTOMOTIVE REPAIR","sales2025":12002.72,"sales2026":19467.09,"change":7464.37,"gpPct":0.118681,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: TUBES (-$25) | LOW GP: TRUCK TIRES (7.3%) | DOWN: ST TRAILER (-$303)"},{"salesman":"House","city":"TIFTON","custNum":101152,"customer":"BROOKS BODY SHOP","sales2025":284.0,"sales2026":1460.0,"change":1176.0,"gpPct":0.044164,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (4.4%) | LOST: PASSENGER TIRES (was $284)"},{"salesman":"House","city":"TIFTON","custNum":200331,"customer":"BUDGET CAR SALES","sales2025":38003.52,"sales2026":12565.57,"change":-25437.95,"gpPct":0.230756,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$7402) | DOWN: RAD LT TRUCK (-$18036)"},{"salesman":"House","city":"TIFTON","custNum":101252,"customer":"DNA DIESEL & AUTOMOTIVE REPAIR","sales2025":2356.9,"sales2026":1076.88,"change":-1280.02,"gpPct":0.111099,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: RAD LT TRUCK (was $200) | LOW GP: TRUCK TIRES (7.6%) | DOWN: TRUCK TIRES (-$1288)"},{"salesman":"House","city":"TIFTON","custNum":101322,"customer":"GRIMES AUTO SERVICE","sales2025":11941.67,"sales2026":10022.44,"change":-1919.23,"gpPct":0.146948,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2908) | DOWN: ST TRAILER (-$137)"},{"salesman":"House","city":"TIFTON","custNum":200407,"customer":"HOLLOWAY TRUCK & TRAILER REPAI","sales2025":3756.12,"sales2026":10440.72,"change":6684.6,"gpPct":0.179991,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $78) | LOW GP: RAD LT TRUCK (6.8%)"},{"salesman":"House","city":"TIFTON","custNum":200718,"customer":"JOBBER ACCT (TIFTON)","sales2025":4042.01,"sales2026":1977.54,"change":-2064.47,"gpPct":0.097652,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1470) | LOW GP: RAD LT TRUCK (9.1%) | DOWN: RAD LT TRUCK (-$370) | DOWN: ST TRAILER (-$358) | LOW GP: VALVE STEMS (9.1%)"},{"salesman":"House","city":"TIFTON","custNum":200805,"customer":"JOEY HALL AUTO SALES LLC","sales2025":573.83,"sales2026":988.0,"change":414.17,"gpPct":-0.060121,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $57)"},{"salesman":"House","city":"TIFTON","custNum":200901,"customer":"LENCHO'S & SON TIRE SHOP","sales2025":6820.85,"sales2026":10029.44,"change":3208.59,"gpPct":0.256205,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $295)"},{"salesman":"House","city":"TIFTON","custNum":200821,"customer":"MARK'S BODY SHOP-TBR ONLY","sales2025":0.0,"sales2026":235.93,"change":235.93,"gpPct":0.137371,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TIFTON","custNum":200918,"customer":"SOUTH GEORGIA TRUCKING SVC LLC","sales2025":5120.41,"sales2026":6379.37,"change":1258.96,"gpPct":0.159122,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1243) | DOWN: ST TRAILER (-$521) | LOST: INDUSTRIAL TIRES (was $333)"},{"salesman":"House","city":"TREVOSE","custNum":200953,"customer":"SIMPLE TIRE - TIFTON","sales2025":409775.18,"sales2026":600529.77,"change":190754.59,"gpPct":0.169633,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$57741) | LOW GP: INDUSTRIAL TIRES (5.3%) | LOW GP: FARM TIRES (9.4%) | DOWN: FARM TIRES (-$368)"},{"salesman":"House","city":"VALDOSTA","custNum":200668,"customer":"24/7 TIRE","sales2025":13376.28,"sales2026":10242.38,"change":-3133.9,"gpPct":0.253589,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$3209)"},{"salesman":"House","city":"VALDOSTA","custNum":200606,"customer":"AFFORDABLE TIRE SERVICE LLC","sales2025":4553.26,"sales2026":3233.73,"change":-1319.53,"gpPct":0.198263,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1301) | LOST: ST TRAILER (was $68)"},{"salesman":"House","city":"VALDOSTA","custNum":200475,"customer":"AZALEA CITY AUTO SALES/SERVICE","sales2025":3868.88,"sales2026":1656.47,"change":-2212.41,"gpPct":0.205998,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1811) | LOST: ST TRAILER (was $257)"},{"salesman":"House","city":"VALDOSTA","custNum":200914,"customer":"BESTDRIVE COMMERCIAL TIRE CTR","sales2025":4022.35,"sales2026":106.0,"change":-3916.35,"gpPct":0.362075,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: RAD LT TRUCK (was $1364) | LOST: INDUSTRIAL TIRES (was $230) | LOST: FARM TIRES (was $128) | LOST: TUBES (was $20) | LOST: TRUCK TIRES (was $2280)"},{"salesman":"House","city":"VALDOSTA","custNum":200674,"customer":"DASHER LLC","sales2025":30558.18,"sales2026":22949.6,"change":-7608.58,"gpPct":0.198658,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$7957)"},{"salesman":"House","city":"VALDOSTA","custNum":200327,"customer":"DRAPER TIRES & AUTOMOTIVE","sales2025":12823.79,"sales2026":10277.37,"change":-2546.42,"gpPct":0.099252,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (8.3%) | LOST: ST TRAILER (was $1311)"},{"salesman":"House","city":"VALDOSTA","custNum":200654,"customer":"ECONOMY USED TIRE (VALDOSTA)","sales2025":257.68,"sales2026":1720.09,"change":1462.41,"gpPct":0.052695,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $10) | LOW GP: RAD LT TRUCK (4.9%) | LOW GP: PASSENGER TIRES (3.9%)"},{"salesman":"House","city":"VALDOSTA","custNum":201055,"customer":"HERNANDEZ TIRES SHOP","sales2025":7044.34,"sales2026":3131.91,"change":-3912.43,"gpPct":0.243474,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2045) | DOWN: RAD LT TRUCK (-$1198) | DOWN: ST TRAILER (-$416) | LOST: TRUCK TIRES (was $252)"},{"salesman":"House","city":"VALDOSTA","custNum":200717,"customer":"JW AUTOMOTIVE","sales2025":4133.03,"sales2026":2578.4,"change":-1554.63,"gpPct":0.234258,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1848)"},{"salesman":"House","city":"VALDOSTA","custNum":200352,"customer":"MALUDA AUTO SALES","sales2025":6083.01,"sales2026":4475.56,"change":-1607.45,"gpPct":0.246146,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1765) | LOST: ST TRAILER (was $208)"},{"salesman":"House","city":"VALDOSTA","custNum":200631,"customer":"MARQUEZ TIRE SHOP","sales2025":35455.07,"sales2026":22760.01,"change":-12695.06,"gpPct":0.165687,"action":"DECLINING - Investigate","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$11679) | DOWN: PASSENGER TIRES (-$2185)"},{"salesman":"House","city":"VALDOSTA","custNum":200964,"customer":"MARTINEZ AUTO SERVICE","sales2025":0.0,"sales2026":551.4,"change":551.4,"gpPct":0.295285,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200699,"customer":"METTS PERFORMANCE & AUTOMOTIVE","sales2025":0.0,"sales2026":542.76,"change":542.76,"gpPct":0.238024,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200355,"customer":"MOTION WHEELS HUBCAPS,&TIRES","sales2025":3325.1,"sales2026":3334.53,"change":9.43,"gpPct":0.176469,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: PASSENGER TIRES (-$429) | DOWN: PATCHES AND REPAIR (-$33) | LOST: TRUCK TIRES (was $681) | LOST: LAWN & GARDEN (was $17)"},{"salesman":"House","city":"VALDOSTA","custNum":201031,"customer":"RENO'S QUALITY COLLISION","sales2025":5821.92,"sales2026":281.52,"change":-5540.4,"gpPct":0.147165,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$346) | LOST: RAD LT TRUCK (was $3270) | LOST: TRUCK TIRES (was $1924)"},{"salesman":"House","city":"VALDOSTA","custNum":200298,"customer":"WISENBAKER'S TIRE & BRAKE","sales2025":103.18,"sales2026":0.0,"change":-103.18,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $103)"},{"salesman":"House","city":"VALDOSTA","custNum":200300,"customer":"Z TIRE EXPRESS","sales2025":4664.9,"sales2026":0.0,"change":-4664.9,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $653) | LOST: PASSENGER TIRES (was $2100) | LOST: ST TRAILER (was $1912)"},{"salesman":"House","city":"WARWICK","custNum":101108,"customer":"PARKER TIRE DIRECT","sales2025":82394.55,"sales2026":85817.68,"change":3423.13,"gpPct":0.090175,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (8.4%) | LOW GP: TRUCK TIRES (9.0%) | DOWN: VALVE STEMS (-$124) | LOST: PATCHES AND REPAIR (was $5)"},{"salesman":"House","city":"WAYCROSS","custNum":200754,"customer":"JORGE USED TIRE SHOP","sales2025":21212.12,"sales2026":19303.45,"change":-1908.67,"gpPct":0.235525,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: RAD LT TRUCK (-$2080) | LOST: TRUCK TIRES (was $4080) | LOST: WHEELS (was $79) | LOST: LAWN & GARDEN (was $29) | LOST: FARM TIRES (was $549)"},{"salesman":"House","city":"WAYCROSS","custNum":201024,"customer":"LIBERTY CAR WASH & TIRE","sales2025":393.56,"sales2026":132.92,"change":-260.64,"gpPct":0.255191,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $494)"},{"salesman":"House","city":"WAYCROSS","custNum":200459,"customer":"LOW COUNTRY TIRE LLC","sales2025":2419.79,"sales2026":2460.17,"change":40.38,"gpPct":0.197084,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$847)"},{"salesman":"House","city":"WHIGHAM","custNum":200238,"customer":"CROSSROADS TIRE & ACC LLC","sales2025":4041.93,"sales2026":0.0,"change":-4041.93,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: FARM TIRES (was $495) | LOST: RAD LT TRUCK (was $1318) | LOST: TRUCK TIRES (was $2008) | LOST: PASSENGER TIRES (was $221)"},{"salesman":"House","city":"WHIGHAM","custNum":200268,"customer":"GIANT TIRE SALES/SERVICE","sales2025":16070.82,"sales2026":8983.83,"change":-7086.99,"gpPct":0.148164,"action":"DECLINING - Investigate","topDept":"OFF THE ROAD TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: TRUCK TIRES (was $2932) | DOWN: RAD LT TRUCK (-$3319) | DOWN: TUBES (-$131)"},{"salesman":"Larry","city":"ADEL","custNum":200804,"customer":"FIVE STAR TIRE ****ADEL****","sales2025":30178.2,"sales2026":20402.93,"change":-9775.27,"gpPct":0.178684,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$8822) | DOWN: ST TRAILER (-$2439) | LOST: LAWN & GARDEN (was $107)"},{"salesman":"Larry","city":"ADEL","custNum":200366,"customer":"ROUNTREE PERFORMANCE","sales2025":3435.36,"sales2026":1722.4,"change":-1712.96,"gpPct":0.051202,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1394) | LOW GP: RAD LT TRUCK (3.4%)"},{"salesman":"Larry","city":"ALBANY","custNum":100301,"customer":"ALBANY GENERAL TIRE SERVICE","sales2025":38536.07,"sales2026":22650.06,"change":-15886.01,"gpPct":0.108832,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: RAD LT TRUCK (4.1%) | DOWN: FARM TIRES (-$840) | LOW GP: PASSENGER TIRES (6.2%) | DOWN: TRUCK TIRES (-$10338) | LOW GP: INDUSTRIAL TIRES (8.5%) | LOST: OFF THE ROAD TIRES (was $12334)"},{"salesman":"Larry","city":"ALBANY","custNum":200592,"customer":"BERNEY'S TIRE SERVICE","sales2025":12435.51,"sales2026":21715.84,"change":9280.33,"gpPct":0.156623,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.8%)"},{"salesman":"Larry","city":"ALBANY","custNum":101479,"customer":"BILL THOMPSON TIRE SERVICES","sales2025":0.0,"sales2026":4627.0,"change":4627.0,"gpPct":0.065822,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (3.4%) | LOW GP: RAD LT TRUCK (3.2%) | LOW GP: TRUCK TIRES (8.0%)"},{"salesman":"Larry","city":"ALBANY","custNum":200214,"customer":"PETERSON TIRE & AUTO CENTER","sales2025":740.94,"sales2026":5158.84,"change":4417.9,"gpPct":0.234281,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALBANY","custNum":200752,"customer":"RAINEY USED CARS (ALBANY)","sales2025":5963.99,"sales2026":7402.88,"change":1438.89,"gpPct":0.142041,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (9.3%)"},{"salesman":"Larry","city":"ALBANY","custNum":200679,"customer":"RICHARD'S AUTO CARE & TIRE SVC","sales2025":6000.2,"sales2026":7273.04,"change":1272.84,"gpPct":0.235302,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: ST TRAILER (-$564)"},{"salesman":"Larry","city":"ALBANY","custNum":200288,"customer":"RIGHT CHOICE AUTO","sales2025":302.92,"sales2026":0.0,"change":-302.92,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $303)"},{"salesman":"Larry","city":"ALBANY","custNum":200939,"customer":"RNR TIRE EXPRESS","sales2025":55801.73,"sales2026":84193.09,"change":28391.36,"gpPct":0.205916,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"DISCOUNTS/COUPONS","focus":"LOST: ST TRAILER (was $694)"},{"salesman":"Larry","city":"ALBANY","custNum":200644,"customer":"SKIP'S AUTOMOTIVE","sales2025":0.0,"sales2026":362.32,"change":362.32,"gpPct":0.280691,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALBANY","custNum":100967,"customer":"SOUTHEASTERN COMMERCIAL TIRE","sales2025":60720.09,"sales2026":25139.31,"change":-35580.78,"gpPct":0.124212,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$37308) | LOW GP: RAD LT TRUCK (10.0%) | DOWN: TUBES (-$130) | DOWN: FARM TIRES (-$2851) | LOST: WHEEL WEIGHTS (was $12)"},{"salesman":"Larry","city":"ALBANY","custNum":200239,"customer":"T & S TIRE","sales2025":13265.53,"sales2026":10581.6,"change":-2683.93,"gpPct":0.250563,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2852)"},{"salesman":"Larry","city":"ALBANY","custNum":500373,"customer":"TOMAHAWK TIRE (ALBANY)","sales2025":3274.01,"sales2026":7779.13,"change":4505.12,"gpPct":0.213938,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$368) | LOW GP: TRUCK TIRES (8.9%)"},{"salesman":"Larry","city":"ALBANY","custNum":200601,"customer":"TUFF ENTERPRISES LLC","sales2025":12394.31,"sales2026":8792.16,"change":-3602.15,"gpPct":0.093201,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.6%) | DOWN: RAD LT TRUCK (-$4495) | LOW GP: PASSENGER TIRES (7.1%)"},{"salesman":"Larry","city":"ALMA","custNum":101507,"customer":"BURNETTE AUTOMOTIVE SERVICE","sales2025":9151.85,"sales2026":31661.41,"change":22509.56,"gpPct":0.151394,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALMA","custNum":200920,"customer":"PRECISION MAINTENANCE","sales2025":30668.91,"sales2026":22028.84,"change":-8640.07,"gpPct":0.147275,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$8286)"},{"salesman":"Larry","city":"ALMA","custNum":2000022,"customer":"R&R AUTO SERVICE & REPAIR","sales2025":0.0,"sales2026":4017.37,"change":4017.37,"gpPct":0.188957,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ASHBURN","custNum":101283,"customer":"CAMERON'S TOWING AND TIRE","sales2025":77330.1,"sales2026":104172.81,"change":26842.71,"gpPct":0.133667,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"BAINBRIDGE","custNum":200184,"customer":"JIMMY'S AUTO SALES","sales2025":689.08,"sales2026":0.0,"change":-689.08,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $439) | LOST: ST TRAILER (was $250)"},{"salesman":"Larry","city":"BROXTON","custNum":100107,"customer":"JOHNSON AUTO & TIRE","sales2025":6396.34,"sales2026":4532.8,"change":-1863.54,"gpPct":0.243119,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$507) | DOWN: RAD LT TRUCK (-$1190) | LOST: TRUCK TIRES (was $298)"},{"salesman":"Larry","city":"CAMILLA","custNum":200280,"customer":"JW PERFORMANCE & AUTO","sales2025":2819.76,"sales2026":0.0,"change":-2819.76,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2820)"},{"salesman":"Larry","city":"CAMILLA","custNum":200965,"customer":"MARQUEZ TIRE SHOP LLC","sales2025":7891.6,"sales2026":9556.19,"change":1664.59,"gpPct":0.155108,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.5%) | DOWN: RAD LT TRUCK (-$2348)"},{"salesman":"Larry","city":"CORDELE","custNum":2000043,"customer":"CENTRAL GA TIRE LLC","sales2025":0.0,"sales2026":1150.04,"change":1150.04,"gpPct":0.303642,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"CORDELE","custNum":200807,"customer":"CORDELE TIRE & WHEEL  LLC","sales2025":41255.53,"sales2026":55076.49,"change":13820.96,"gpPct":0.181404,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (9.1%)"},{"salesman":"Larry","city":"CORDELE","custNum":101241,"customer":"NEW PETTIS TIRE","sales2025":7257.96,"sales2026":6372.81,"change":-885.15,"gpPct":0.130782,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOST: LAWN & GARDEN (was $101) | LOW GP: PASSENGER TIRES (8.2%) | LOST: TUBES (was $310) | DOWN: TRUCK TIRES (-$672) | DOWN: FARM TIRES (-$1088)"},{"salesman":"Larry","city":"CORDELE","custNum":101491,"customer":"QUALITY AUTO & R.V. SERVICE","sales2025":4319.6,"sales2026":1388.24,"change":-2931.36,"gpPct":0.267619,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1628) | DOWN: RAD LT TRUCK (-$787) | LOST: ST TRAILER (was $517)"},{"salesman":"Larry","city":"CORDELE","custNum":100907,"customer":"SMITH'S DIESEL REPAIR","sales2025":8678.57,"sales2026":6428.18,"change":-2250.39,"gpPct":0.11887,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$3405) | LOST: ST TRAILER (was $231)"},{"salesman":"Larry","city":"DAWSON","custNum":200946,"customer":"ABBI'S 24 HOUR","sales2025":19013.41,"sales2026":36062.53,"change":17049.12,"gpPct":0.092876,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.4%) | LOW GP: RAD LT TRUCK (5.2%)"},{"salesman":"Larry","city":"DESOTO","custNum":100417,"customer":"CLARK BASS SERVICE","sales2025":62379.61,"sales2026":62450.89,"change":71.28,"gpPct":0.094554,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: RAD LT TRUCK (6.7%) | DOWN: OFF THE ROAD TIRES (-$2425) | LOW GP: PASSENGER TIRES (4.8%) | DOWN: FARM TIRES (-$11389)"},{"salesman":"Larry","city":"DOUGLAS","custNum":101080,"customer":"AMERSON TIRE INC.","sales2025":153513.73,"sales2026":98917.4,"change":-54596.33,"gpPct":0.080502,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.6%) | DOWN: PASSENGER TIRES (-$12730) | LOW GP: FARM TIRES (8.3%) | LOW GP: TRUCK TIRES (9.6%) | DOWN: TRUCK TIRES (-$13696) | LOW GP: OFF THE ROAD TIRES (7.1%) | DOWN: OFF THE ROAD TIRES (-$9853) | LOST: PATCHES AND REPAIR (was $46) | LOST: LAWN & GARDEN (was $502)"},{"salesman":"Larry","city":"DOUGLAS","custNum":101539,"customer":"COURSON'S TIRE OF DOUGLAS","sales2025":102164.83,"sales2026":64038.51,"change":-38126.32,"gpPct":0.136859,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$28687) | DOWN: PASSENGER TIRES (-$12657) | DOWN: TUBES (-$441) | LOW GP: TRUCK TIRES (7.2%) | DOWN: LAWN & GARDEN (-$73)"},{"salesman":"Larry","city":"DOUGLAS","custNum":101295,"customer":"DAVID'S AUTO SALES / DOUGLAS","sales2025":23360.62,"sales2026":23619.52,"change":258.9,"gpPct":0.23946,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"DOUGLAS","custNum":101161,"customer":"JMC TIRE CO.  INC.","sales2025":284415.85,"sales2026":211731.22,"change":-72684.63,"gpPct":0.114558,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: OFF THE ROAD TIRES (was $3263) | LOW GP: TRUCK TIRES (8.8%) | DOWN: RAD LT TRUCK (-$25926) | DOWN: FARM TIRES (-$11588) | LOST: INDUSTRIAL TIRES (was $2628) | LOST: WHEELS (was $500)"},{"salesman":"Larry","city":"DOUGLAS","custNum":200907,"customer":"KNOLLWOOD TIRE & WHEEL","sales2025":12111.89,"sales2026":3657.72,"change":-8454.17,"gpPct":0.165789,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$6579) | LOST: TRUCK TIRES (was $1742) | LOST: ST TRAILER (was $133)"},{"salesman":"Larry","city":"DOUGLAS","custNum":200510,"customer":"M & R TRUCK ACCESSORIES","sales2025":5075.05,"sales2026":3224.24,"change":-1850.81,"gpPct":0.085053,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (5.6%) | DOWN: RAD LT TRUCK (-$1284)"},{"salesman":"Larry","city":"DOUGLAS","custNum":2000042,"customer":"ROJAS AUTO REPAIR","sales2025":0.0,"sales2026":4307.51,"change":4307.51,"gpPct":0.262727,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"DOUGLAS","custNum":200959,"customer":"SOUTHERN GEORGIA TIRE LLC","sales2025":71785.32,"sales2026":49991.36,"change":-21793.96,"gpPct":0.178277,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$17786) | LOW GP: TRUCK TIRES (9.8%)"},{"salesman":"Larry","city":"FITZGERALD","custNum":100993,"customer":"D & G PERFORMANCE","sales2025":2712.92,"sales2026":0.0,"change":-2712.92,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2713)"},{"salesman":"Larry","city":"FITZGERALD","custNum":2000023,"customer":"FABOS AUTO SALES LLC","sales2025":0.0,"sales2026":104.08,"change":104.08,"gpPct":0.259416,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"FITZGERALD","custNum":101519,"customer":"MARK TAYLOR DBA/MTAA ENT.","sales2025":8873.23,"sales2026":7290.54,"change":-1582.69,"gpPct":0.133024,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$589) | DOWN: RAD LT TRUCK (-$1117)"},{"salesman":"Larry","city":"FITZGERALD","custNum":101463,"customer":"SHELL RAPID LUBE (FITZGERALD)","sales2025":108571.19,"sales2026":116343.49,"change":7772.3,"gpPct":0.147902,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: INDUSTRIAL TIRES (-$803) | DOWN: TRUCK TIRES (-$9050) | LOST: OUTSIDE PURCHASE (was $214) | LOST: WHEELS (was $4294) | LOST: TIRE TOOLS (was $300)"},{"salesman":"Larry","city":"FITZGERALD","custNum":200621,"customer":"SOUTH GA LUBE (FITZGERALD)","sales2025":5826.39,"sales2026":2148.68,"change":-3677.71,"gpPct":0.208854,"action":"DECLINING - Investigate","topDept":"INDUSTRIAL TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: WHEELS (was $1349) | DOWN: FARM TIRES (-$1473) | LOST: OFF THE ROAD TIRES (was $2365)"},{"salesman":"Larry","city":"HOBOKEN","custNum":200683,"customer":"HICKOX AUTO DEALERS","sales2025":1440.71,"sales2026":4254.88,"change":2814.17,"gpPct":0.119646,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (9.6%)"},{"salesman":"Larry","city":"HORTENSE","custNum":200502,"customer":"ATKINSON TIRE","sales2025":746.61,"sales2026":3116.12,"change":2369.51,"gpPct":0.044713,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (6.0%)"},{"salesman":"Larry","city":"LEESBURG","custNum":201053,"customer":"BERNEYS TIRE SERVICE","sales2025":11617.48,"sales2026":27020.89,"change":15403.41,"gpPct":0.12587,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":200317,"customer":"BROTHERS TIRES","sales2025":20578.09,"sales2026":15310.95,"change":-5267.14,"gpPct":0.202162,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: ST TRAILER (-$533) | LOST: TRUCK TIRES (was $620) | LOST: LAWN & GARDEN (was $122)"},{"salesman":"Larry","city":"MOULTRIE","custNum":201035,"customer":"DAVID'S AUTO SALES (MOULTRIE)","sales2025":10341.14,"sales2026":24587.36,"change":14246.22,"gpPct":0.2368,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":2000015,"customer":"EDWIN'S TIRES LLC","sales2025":0.0,"sales2026":493.48,"change":493.48,"gpPct":0.225825,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":101513,"customer":"GAY'S TIRE SERVICE","sales2025":133.86,"sales2026":568.77,"change":434.91,"gpPct":0.282821,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":200198,"customer":"MOULTRIE TIRE","sales2025":65856.64,"sales2026":99415.06,"change":33558.42,"gpPct":0.172093,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: WHEELS (was $67)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200755,"customer":"N-T TIRE SERVICE","sales2025":22035.18,"sales2026":20653.16,"change":-1382.02,"gpPct":0.16968,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$1058)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200628,"customer":"SOUTH GEORGIA TIRE","sales2025":33389.72,"sales2026":34480.43,"change":1090.71,"gpPct":0.208594,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (6.7%) | DOWN: TRUCK TIRES (-$2221) | DOWN: TUBES (-$152) | LOST: OFF THE ROAD TIRES (was $548)"},{"salesman":"Larry","city":"MOULTRIE","custNum":2000020,"customer":"T&D TIRE","sales2025":0.0,"sales2026":1063.08,"change":1063.08,"gpPct":0.208601,"action":"GROWING - Expand","topDept":"TUBES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":200242,"customer":"THOMAS TIRE RECAPPING INC.","sales2025":19466.54,"sales2026":404.85,"change":-19061.69,"gpPct":0.1711,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$3891) | LOST: RAD LT TRUCK (was $10705) | LOST: TRUCK TIRES (was $4326) | LOST: ST TRAILER (was $139)"},{"salesman":"Larry","city":"NAHUNTA","custNum":200478,"customer":"82 TIRE & LUBE","sales2025":14657.13,"sales2026":15794.59,"change":1137.46,"gpPct":0.228281,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"Larry","city":"NASHVILLE","custNum":200319,"customer":"BUCK'S AUTO REPAIR","sales2025":13446.03,"sales2026":10013.66,"change":-3432.37,"gpPct":0.124669,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$1683)"},{"salesman":"Larry","city":"NASHVILLE","custNum":101415,"customer":"THE TIRE STORE","sales2025":72970.86,"sales2026":111474.76,"change":38503.9,"gpPct":0.155769,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TIRE TOOLS (was $9) | LOST: MOUNTING LUBE (was $23)"},{"salesman":"Larry","city":"OCILLA","custNum":101181,"customer":"SOUTH GA LUBE CENTER","sales2025":2871.22,"sales2026":2686.78,"change":-184.44,"gpPct":0.184991,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $410)"},{"salesman":"Larry","city":"PEARSON","custNum":101297,"customer":"FOUR C'S LUBE","sales2025":42688.79,"sales2026":29676.31,"change":-13012.48,"gpPct":0.135408,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$1065) | DOWN: TRUCK TIRES (-$14420) | LOST: OFF THE ROAD TIRES (was $439) | LOST: INDUSTRIAL TIRES (was $710)"},{"salesman":"Larry","city":"PEARSON","custNum":200388,"customer":"PEARSON TIRE & LUBE","sales2025":4497.58,"sales2026":6687.42,"change":2189.84,"gpPct":0.186253,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $39)"},{"salesman":"Larry","city":"PEARSON","custNum":200762,"customer":"POWER MAN TIRE SHOP","sales2025":30414.43,"sales2026":10626.76,"change":-19787.67,"gpPct":0.22574,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$4175) | DOWN: RAD LT TRUCK (-$11825) | DOWN: VALVE STEMS (-$83) | DOWN: ST TRAILER (-$2096) | LOST: TUBES (was $215) | LOST: TRUCK TIRES (was $1453) | DOWN: LAWN & GARDEN (-$290)"},{"salesman":"Larry","city":"SYCAMORE","custNum":200971,"customer":"ALLEN'S TIRE","sales2025":73410.55,"sales2026":88373.18,"change":14962.63,"gpPct":0.13742,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: TRUCK TIRES (9.7%) | LOW GP: FARM TIRES (5.0%) | DOWN: PASSENGER TIRES (-$5185)"},{"salesman":"Larry","city":"SYCAMORE","custNum":200891,"customer":"EJH WRECKER & TIRE SERVICE","sales2025":69049.97,"sales2026":116075.04,"change":47025.07,"gpPct":0.141565,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $237)"},{"salesman":"Larry","city":"SYLVESTER","custNum":101436,"customer":"ED'S TIRE","sales2025":28691.82,"sales2026":43606.37,"change":14914.55,"gpPct":0.1813,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$498) | LOST: FARM TIRES (was $78)"},{"salesman":"Larry","city":"SYLVESTER","custNum":2000002,"customer":"EG AGRI PARTS LLC","sales2025":16535.6,"sales2026":0.0,"change":-16535.6,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $12513) | LOST: TRUCK TIRES (was $4421)"},{"salesman":"Larry","city":"SYLVESTER","custNum":200972,"customer":"ERIC'S TIRE OF SYLVESTER","sales2025":16106.12,"sales2026":10707.93,"change":-5398.19,"gpPct":0.190806,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: LAWN & GARDEN (was $39) | DOWN: RAD LT TRUCK (-$3705) | LOST: TRUCK TIRES (was $696)"},{"salesman":"Larry","city":"SYLVESTER","custNum":2000017,"customer":"JORDAN AUTOMOTIVE & TIRES","sales2025":0.0,"sales2026":3754.02,"change":3754.02,"gpPct":0.087464,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (4.3%)"},{"salesman":"Larry","city":"SYLVESTER","custNum":200500,"customer":"SHELL RAPID LUBE (SYLVESTER)","sales2025":465.28,"sales2026":78.96,"change":-386.32,"gpPct":0.215046,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$403)"},{"salesman":"Larry","city":"SYLVESTER","custNum":200220,"customer":"SINGLETARY & SON TIRE CO","sales2025":81555.68,"sales2026":146672.06,"change":65116.38,"gpPct":0.146837,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (9.7%)"},{"salesman":"Larry","city":"TIFTON","custNum":200869,"customer":"AFFORDABLE DIESEL REPAIR","sales2025":8138.56,"sales2026":0.0,"change":-8138.56,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $8021) | LOST: RAD LT TRUCK (was $118)"},{"salesman":"Larry","city":"TIFTON","custNum":200676,"customer":"ALL PURPOSE AUTO CENTER","sales2025":1774.36,"sales2026":635.62,"change":-1138.74,"gpPct":0.389667,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1009) | LOST: TRUCK TIRES (was $582)"},{"salesman":"Larry","city":"TIFTON","custNum":200503,"customer":"BB'S AUTOMOTIVE","sales2025":3065.43,"sales2026":3669.73,"change":604.3,"gpPct":0.176678,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.8%) | LOST: TRUCK TIRES (was $596)"},{"salesman":"Larry","city":"TIFTON","custNum":201043,"customer":"BILL'S TRAILER SERVICE","sales2025":2645.11,"sales2026":1821.59,"change":-823.52,"gpPct":0.085217,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"ST TRAILER","focus":"LOW GP: TRUCK TIRES (6.4%) | DOWN: ST TRAILER (-$905) | LOST: RAD LT TRUCK (was $183)"},{"salesman":"Larry","city":"TIFTON","custNum":200915,"customer":"DAVID'S AUTO SALES / TIFTON","sales2025":26810.34,"sales2026":18293.69,"change":-8516.65,"gpPct":0.210511,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5166) | LOST: ST TRAILER (was $92)"},{"salesman":"Larry","city":"TIFTON","custNum":200635,"customer":"DELTORO TIRE #2","sales2025":211854.01,"sales2026":271034.68,"change":59180.67,"gpPct":0.114846,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: TIRE TOOLS (-$442) | DOWN: OFF THE ROAD TIRES (-$1311)"},{"salesman":"Larry","city":"TIFTON","custNum":101371,"customer":"DIRTY SOUTH KUSTOMS","sales2025":12335.43,"sales2026":11788.11,"change":-547.32,"gpPct":0.047179,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (4.4%) | LOST: TRUCK TIRES (was $1761) | LOST: PASSENGER TIRES (was $112)"},{"salesman":"Larry","city":"TIFTON","custNum":200973,"customer":"ERIC'S TIRE (REBEL ROAD)","sales2025":24081.98,"sales2026":16464.1,"change":-7617.88,"gpPct":0.160554,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$3141) | DOWN: RAD LT TRUCK (-$4678) | DOWN: ST TRAILER (-$235) | LOST: TUBES (was $10)"},{"salesman":"Larry","city":"TIFTON","custNum":101323,"customer":"ERIC'S TIRE SERVICE","sales2025":73193.75,"sales2026":112850.09,"change":39656.34,"gpPct":0.162442,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: INDUSTRIAL TIRES (9.6%)"},{"salesman":"Larry","city":"TIFTON","custNum":201051,"customer":"EXPRESS OIL CHANGE #3168","sales2025":2913.64,"sales2026":0.0,"change":-2913.64,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1975) | LOST: PASSENGER TIRES (was $939)"},{"salesman":"Larry","city":"TIFTON","custNum":200560,"customer":"FIVE STAR TIRE SERVICE LLC","sales2025":65254.11,"sales2026":71878.62,"change":6624.51,"gpPct":0.191304,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $595)"},{"salesman":"Larry","city":"TIFTON","custNum":200868,"customer":"GOLDEN ENVIRONMENTAL","sales2025":23931.15,"sales2026":21018.79,"change":-2912.36,"gpPct":0.187208,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$520) | DOWN: RAD LT TRUCK (-$284)"},{"salesman":"Larry","city":"TIFTON","custNum":200501,"customer":"LARRY'S BODY SHOP","sales2025":896.92,"sales2026":404.16,"change":-492.76,"gpPct":0.265588,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$493)"},{"salesman":"Larry","city":"TIFTON","custNum":100842,"customer":"LOVE AVE. SERVICE CTR.","sales2025":3631.51,"sales2026":4958.14,"change":1326.63,"gpPct":0.121076,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$363)"},{"salesman":"Larry","city":"TIFTON","custNum":200409,"customer":"MASTER CRAFT IND.(NO PASS/LT)","sales2025":7898.08,"sales2026":8830.47,"change":932.39,"gpPct":0.198551,"action":"MAINTAIN - Stable account","topDept":"INDUSTRIAL TIRES","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$2413)"},{"salesman":"Larry","city":"TIFTON","custNum":200468,"customer":"MCKEE'S AUTO CENTER  INC","sales2025":1046.81,"sales2026":1426.91,"change":380.1,"gpPct":0.233182,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: PATCHES AND REPAIR (was $8) | LOW GP: TRUCK TIRES (7.3%) | DOWN: TRUCK TIRES (-$436)"},{"salesman":"Larry","city":"TIFTON","custNum":201016,"customer":"O&C AUTO REPAIR","sales2025":2381.49,"sales2026":2647.88,"change":266.39,"gpPct":0.242794,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$874)"},{"salesman":"Larry","city":"TIFTON","custNum":200608,"customer":"OFFROAD POWERSPORTS","sales2025":834.44,"sales2026":1899.83,"change":1065.39,"gpPct":0.092956,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.9%)"},{"salesman":"Larry","city":"TIFTON","custNum":200483,"customer":"PERRIN FARM EQUIPMENT","sales2025":7431.76,"sales2026":6280.62,"change":-1151.14,"gpPct":0.125394,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$1908) | DOWN: PASSENGER TIRES (-$1114) | LOST: ST TRAILER (was $506) | LOW GP: INDUSTRIAL TIRES (7.2%)"},{"salesman":"Larry","city":"TIFTON","custNum":201017,"customer":"PINEDA'S AUTOMOTIVE","sales2025":2523.12,"sales2026":2018.33,"change":-504.79,"gpPct":0.147305,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$600)"},{"salesman":"Larry","city":"TIFTON","custNum":100591,"customer":"RAINEY ALIGNMENT","sales2025":7084.44,"sales2026":14301.86,"change":7217.42,"gpPct":0.086422,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TUBES","focus":"LOW GP: RAD LT TRUCK (5.6%) | LOST: TUBES (was $8)"},{"salesman":"Larry","city":"TIFTON","custNum":100282,"customer":"RUDY'S TIRE SERVICE","sales2025":114154.42,"sales2026":97133.45,"change":-17020.97,"gpPct":0.173845,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.5%) | LOST: LAWN & GARDEN (was $141) | LOST: WHEEL WEIGHTS (was $48)"},{"salesman":"Larry","city":"TIFTON","custNum":100551,"customer":"SOUTHSIDE TIRE & AUTO SERVICE","sales2025":42375.11,"sales2026":41434.02,"change":-941.09,"gpPct":0.202269,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: INDUSTRIAL TIRES (-$224)"},{"salesman":"Larry","city":"TIFTON","custNum":200490,"customer":"T.C.A. IRRIGATION","sales2025":0.0,"sales2026":1021.42,"change":1021.42,"gpPct":0.181904,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":200956,"customer":"TENNESON COLLISION CENTER","sales2025":2277.72,"sales2026":4006.44,"change":1728.72,"gpPct":0.101731,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (7.0%)"},{"salesman":"Larry","city":"TIFTON","custNum":200744,"customer":"TIFTON COMMERCIAL","sales2025":1561.91,"sales2026":578.63,"change":-983.28,"gpPct":0.263225,"action":"DECLINING - Investigate","topDept":"TUBES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$719) | LOST: ST TRAILER (was $412) | LOST: RAD LT TRUCK (was $123)"},{"salesman":"Larry","city":"TIFTON","custNum":101326,"customer":"TIFTON GENERAL TIRE","sales2025":179625.69,"sales2026":175688.7,"change":-3936.99,"gpPct":0.094796,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.9%) | DOWN: TRUCK TIRES (-$19181) | LOW GP: PASSENGER TIRES (7.1%) | LOW GP: RAD LT TRUCK (8.8%) | DOWN: INDUSTRIAL TIRES (-$410) | DOWN: WHEELS (-$454) | LOST: OFF THE ROAD TIRES (was $139)"},{"salesman":"Larry","city":"TIFTON","custNum":201062,"customer":"TIRE MASTERS LLC","sales2025":7041.06,"sales2026":14439.05,"change":7397.99,"gpPct":0.124694,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"WARWICK","custNum":200760,"customer":"MALLARD'S SERVICE CENTER","sales2025":4760.32,"sales2026":4666.64,"change":-93.68,"gpPct":0.195346,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1792)"},{"salesman":"Larry","city":"WILLACOOCHEE","custNum":201014,"customer":"D&R AUTO SALES & SALVAGE PARTS","sales2025":3238.19,"sales2026":3758.46,"change":520.27,"gpPct":0.235027,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200885,"customer":"PIERCE INDUSTRIAL TIRE LLC","sales2025":21901.36,"sales2026":39147.4,"change":17246.04,"gpPct":0.024159,"action":"LOW MARGIN - Reprice","topDept":"OFF THE ROAD TIRES","declinedDept":"TUBES","focus":"LOST: TUBES (was $301) | LOW GP: OFF THE ROAD TIRES (2.0%) | LOW GP: RAD LT TRUCK (4.0%) | LOW GP: FARM TIRES (8.0%) | LOW GP: PASSENGER TIRES (1.0%)"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200883,"customer":"ROLLING BEAR TIRES LLC","sales2025":47164.0,"sales2026":45779.88,"change":-1384.12,"gpPct":0.192372,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6799) | DOWN: TRUCK TIRES (-$1153)"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200712,"customer":"TANNER AUTO REPAIR PLUS  LLC","sales2025":7193.61,"sales2026":186.14,"change":-7007.47,"gpPct":0.17836,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $5248) | DOWN: PASSENGER TIRES (-$1759)"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200474,"customer":"TIRE & WHEEL INC","sales2025":9254.93,"sales2026":7923.69,"change":-1331.24,"gpPct":0.038862,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (0.6%) | LOST: TRUCK TIRES (was $1480) | LOST: PASSENGER TIRES (was $693)"},{"salesman":"Tiffany","city":"BROOKSVILLE","custNum":2000040,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":3779.04,"change":3779.04,"gpPct":0.174912,"action":"GROWING - Expand","topDept":"OUTSIDE PURCHASE","declinedDept":"","focus":"LOW GP: OUTSIDE PURCHASE (5.0%)"},{"salesman":"Tiffany","city":"DONALSONVILLE","custNum":200246,"customer":"TRI COUNTY TIRE COMPANY","sales2025":24439.78,"sales2026":27762.36,"change":3322.58,"gpPct":0.127716,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (8.5%) | DOWN: TRUCK TIRES (-$5728) | DOWN: ST TRAILER (-$984) | DOWN: PASSENGER TIRES (-$316) | DOWN: INDUSTRIAL TIRES (-$823)"},{"salesman":"Tiffany","city":"DOTHAN","custNum":101878,"customer":"TRI STATE COMMERCIAL TIRE LLC","sales2025":1103.9,"sales2026":1452.12,"change":348.22,"gpPct":0.081557,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$1020) | LOW GP: RAD LT TRUCK (7.5%)"},{"salesman":"Tiffany","city":"GAINESVILLE","custNum":2000021,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":3533.93,"change":3533.93,"gpPct":0.251957,"action":"GROWING - Expand","topDept":"VALVE STEMS","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"GREENVILLE","custNum":200684,"customer":"OTR SERVICES  INC.","sales2025":1913.04,"sales2026":0.0,"change":-1913.04,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1748) | LOST: TUBES (was $165)"},{"salesman":"Tiffany","city":"HAHIRA","custNum":200270,"customer":"HAHIRA AUTOMOTIVE SERVICE","sales2025":21229.74,"sales2026":16861.38,"change":-4368.36,"gpPct":0.151728,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6140) | DOWN: TUBES (-$34) | LOST: FARM TIRES (was $78)"},{"salesman":"Tiffany","city":"JASPER","custNum":200876,"customer":"SUWANNEE VALLEY TIRE","sales2025":2500.0,"sales2026":322.44,"change":-2177.56,"gpPct":0.013925,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $2500) | LOW GP: RAD LT TRUCK (1.4%)"},{"salesman":"Tiffany","city":"LAKE BUTLER","custNum":201046,"customer":"BIELLING'S TIRE INC.","sales2025":4106.0,"sales2026":0.0,"change":-4106.0,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $4106)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":200923,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":2387.13,"change":2387.13,"gpPct":0.15576,"action":"GROWING - Expand","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"LOW GP: INDUSTRIAL TIRES (5.0%)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":2000018,"customer":"ALL PRO DIESEL  LLC","sales2025":0.0,"sales2026":6105.39,"change":6105.39,"gpPct":0.218115,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":2000044,"customer":"MURRAY'S TIRE & ROAD SERVICE","sales2025":0.0,"sales2026":7822.25,"change":7822.25,"gpPct":0.175141,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.2%)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":200878,"customer":"THOMAS TIRE REPAIR & ROAD SVC","sales2025":3377.8,"sales2026":4215.2,"change":837.4,"gpPct":0.174625,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"ST TRAILER","focus":"On track"},{"salesman":"Tiffany","city":"LAKELAND","custNum":200827,"customer":"LAKELAND TIRE DBA COOK & SONS","sales2025":97889.25,"sales2026":146856.72,"change":48967.47,"gpPct":0.157022,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: OFF THE ROAD TIRES (6.0%)"},{"salesman":"Tiffany","city":"LENOX","custNum":200406,"customer":"LENOX TIRE & SERVICE CENTER","sales2025":3732.01,"sales2026":10729.37,"change":6997.36,"gpPct":0.169936,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: WHEEL WEIGHTS (was $10)"},{"salesman":"Tiffany","city":"LENOX","custNum":101466,"customer":"WATTS REPAIR SERVICE","sales2025":13487.83,"sales2026":8685.34,"change":-4802.49,"gpPct":0.194827,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$4460) | LOST: ST TRAILER (was $1277)"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200961,"customer":"BABCOCK TIRE LLC","sales2025":25858.63,"sales2026":16109.96,"change":-9748.67,"gpPct":0.129545,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7789) | LOST: INDUSTRIAL TIRES (was $3182)"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200687,"customer":"LIVE OAK TIRE CENTER  LLC","sales2025":4691.95,"sales2026":20703.41,"change":16011.46,"gpPct":0.103023,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (5.8%) | LOW GP: FARM TIRES (10.0%)"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200659,"customer":"TOWN & COUNTRY TIRE","sales2025":66.99,"sales2026":1698.37,"change":1631.38,"gpPct":0.231375,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"TUBES","focus":"LOST: TUBES (was $67)"},{"salesman":"Tiffany","city":"MADISON","custNum":200691,"customer":"KENDA TRUCK CENTER","sales2025":24851.8,"sales2026":36329.05,"change":11477.25,"gpPct":0.08459,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: TRUCK TIRES (7.9%) | DOWN: PASSENGER TIRES (-$397) | LOST: OFF THE ROAD TIRES (was $1297) | LOST: ST TRAILER (was $272)"},{"salesman":"Tiffany","city":"MADISON","custNum":200636,"customer":"MTC SOUTH  INC.","sales2025":154268.09,"sales2026":96721.98,"change":-57546.11,"gpPct":0.103625,"action":"DECLINING - Investigate","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.4%) | DOWN: TRUCK TIRES (-$61348) | DOWN: FARM TIRES (-$1866) | LOST: VALVE STEMS (was $21) | DOWN: MOUNTING LUBE (-$59)"},{"salesman":"Tiffany","city":"MAYO","custNum":200595,"customer":"W.R. WILLIAMS","sales2025":99727.99,"sales2026":88761.77,"change":-10966.22,"gpPct":0.137617,"action":"MAINTAIN - Stable account","topDept":"FARM TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6844) | LOW GP: TRUCK TIRES (9.7%) | LOW GP: OFF THE ROAD TIRES (7.2%) | DOWN: OFF THE ROAD TIRES (-$2979)"},{"salesman":"Tiffany","city":"NORMAN PARK","custNum":101066,"customer":"WARRIOR CREEK TIRE  LLC","sales2025":41518.67,"sales2026":32833.95,"change":-8684.72,"gpPct":0.164558,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: TUBES (-$668) | DOWN: FARM TIRES (-$12237) | LOW GP: PASSENGER TIRES (6.5%) | LOST: LAWN & GARDEN (was $89)"},{"salesman":"Tiffany","city":"OCALA","custNum":200922,"customer":"ADVANCED TIRE SERVICE","sales2025":42083.6,"sales2026":296422.76,"change":254339.16,"gpPct":0.075772,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"DISCOUNTS/COUPONS","focus":"LOW GP: INDUSTRIAL TIRES (4.4%) | LOW GP: TRUCK TIRES (9.3%) | LOW GP: OFF THE ROAD TIRES (3.8%)"},{"salesman":"Tiffany","city":"OCALA","custNum":2000039,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":7685.51,"change":7685.51,"gpPct":0.144105,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"PELHAM","custNum":101544,"customer":"GODWIN TIRE & AUTO","sales2025":551.85,"sales2026":11320.38,"change":10768.53,"gpPct":0.198977,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"PERRY","custNum":200596,"customer":"CRIBBS TIRE","sales2025":593.18,"sales2026":0.0,"change":-593.18,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $593)"},{"salesman":"Tiffany","city":"PERRY","custNum":200673,"customer":"JB'S TIRE & REPAIR SVC.","sales2025":4820.23,"sales2026":15763.77,"change":10943.54,"gpPct":0.08378,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (4.9%) | LOW GP: PASSENGER TIRES (5.6%)"},{"salesman":"Tiffany","city":"PERRY","custNum":200698,"customer":"YARBROUGH TIRE CO.  INC.","sales2025":1755.68,"sales2026":0.0,"change":-1755.68,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $818) | LOST: PASSENGER TIRES (was $938)"},{"salesman":"Tiffany","city":"QUINCY","custNum":201005,"customer":"MIDWAY ENTERPRISE FL  LLC","sales2025":3847.59,"sales2026":1819.51,"change":-2028.08,"gpPct":0.229496,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1578) | DOWN: RAD LT TRUCK (-$384)"},{"salesman":"Tiffany","city":"QUITMAN","custNum":200383,"customer":"WILLIAMS ALIGNMENT & TIRE","sales2025":13044.88,"sales2026":21498.64,"change":8453.76,"gpPct":0.155497,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $803)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200921,"customer":"AUTO DOCTOR DIESEL & REPAIR","sales2025":5520.26,"sales2026":242.46,"change":-5277.8,"gpPct":0.355234,"action":"DECLINING - Investigate","topDept":"WHEEL WEIGHTS","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $550) | LOST: RAD LT TRUCK (was $4971)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200970,"customer":"BIG PINE REPAIR","sales2025":-100.0,"sales2026":0.0,"change":100.0,"gpPct":0.0,"action":"MAINTAIN - Stable account","topDept":"None","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200434,"customer":"CAPITAL AUTO PARTS  INC","sales2025":0.0,"sales2026":1025.06,"change":1025.06,"gpPct":0.140499,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200285,"customer":"PRECISION AUTOCRAFT INC","sales2025":577.84,"sales2026":0.0,"change":-577.84,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $578)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200291,"customer":"SINGLETARY TIRE PROS","sales2025":769.32,"sales2026":674.24,"change":-95.08,"gpPct":0.282881,"action":"MAINTAIN - Stable account","topDept":"PATCHES AND REPAIR","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $769)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200293,"customer":"THOMASVILLE TIRE DEPT.","sales2025":52912.63,"sales2026":58248.76,"change":5336.13,"gpPct":0.120936,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: FARM TIRES (7.3%) | DOWN: TRUCK TIRES (-$3836)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200385,"customer":"WILLIAMS AUTOMOTIVE","sales2025":1262.74,"sales2026":1392.73,"change":129.99,"gpPct":0.2737,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200867,"customer":"24/7 DIESEL AND TIRE REPAIR","sales2025":17556.05,"sales2026":17014.28,"change":-541.77,"gpPct":0.126465,"action":"MAINTAIN - Stable account","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200439,"customer":"BEAR TIRE SERVICE","sales2025":1511.22,"sales2026":4342.64,"change":2831.42,"gpPct":0.227247,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: FARM TIRES (7.0%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200663,"customer":"BEN'S TIRE & AUTO","sales2025":38949.51,"sales2026":35682.59,"change":-3266.92,"gpPct":0.203274,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: WHEEL WEIGHTS (was $120) | LOST: INDUSTRIAL TIRES (was $57)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200967,"customer":"EDDIES AUTOMOTIVE AND TIRE","sales2025":2904.13,"sales2026":13001.22,"change":10097.09,"gpPct":0.156491,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $29)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200935,"customer":"FROMETA USED CAR & TIRE CENTER","sales2025":1437.02,"sales2026":887.68,"change":-549.34,"gpPct":0.157061,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $800) | LOST: ST TRAILER (was $309)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":101565,"customer":"HARRY B ANDERSON","sales2025":0.0,"sales2026":5876.88,"change":5876.88,"gpPct":0.170665,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":101146,"customer":"NE-RO TIRE & BRAKE SERVICE INC","sales2025":34709.83,"sales2026":31276.62,"change":-3433.21,"gpPct":0.11775,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TUBES (-$1059) | DOWN: INDUSTRIAL TIRES (-$3795) | DOWN: TRUCK TIRES (-$13689) | LOW GP: OFF THE ROAD TIRES (7.7%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200362,"customer":"RAY NORTON TIRE & AUTO","sales2025":8933.73,"sales2026":22663.28,"change":13729.55,"gpPct":0.184536,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200932,"customer":"RNR TIRE EXPRESS","sales2025":16694.07,"sales2026":2735.1,"change":-13958.97,"gpPct":0.179299,"action":"DECLINING - Investigate","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$11935) | DOWN: PASSENGER TIRES (-$2024)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200370,"customer":"SMITH TIRE COMPANY","sales2025":0.0,"sales2026":1224.0,"change":1224.0,"gpPct":0.039706,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (4.0%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":2000029,"customer":"SOUTHERN TIRE MART @ PILOT","sales2025":0.0,"sales2026":4792.35,"change":4792.35,"gpPct":0.031855,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (6.9%) | LOW GP: TRUCK TIRES (4.4%) | LOW GP: RAD LT TRUCK (1.6%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200294,"customer":"TIRE KING OF VALDOSTA","sales2025":19232.35,"sales2026":36770.79,"change":17538.44,"gpPct":0.165877,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$827) | DOWN: TUBES (-$76) | DOWN: FARM TIRES (-$577)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200793,"customer":"ALL SEASON AUTO REPAIR","sales2025":1440.12,"sales2026":2007.2,"change":567.08,"gpPct":0.099935,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$939) | LOW GP: TRUCK TIRES (7.9%)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200810,"customer":"BOULEVARD TIRE CENTER","sales2025":27794.87,"sales2026":37840.11,"change":10045.24,"gpPct":0.12507,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"DOWN: TUBES (-$710) | DOWN: INDUSTRIAL TIRES (-$870) | LOST: ST TRAILER (was $458)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200886,"customer":"C&L PERFORMANCE INC","sales2025":1716.09,"sales2026":695.88,"change":-1020.21,"gpPct":0.243102,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$585) | LOST: PASSENGER TIRES (was $362) | LOST: TUBES (was $10) | LOST: FARM TIRES (was $64)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":101477,"customer":"DISCOUNT TIRE (ALMA)OSTEEN","sales2025":12960.79,"sales2026":15443.76,"change":2482.97,"gpPct":0.137627,"action":"MAINTAIN - Stable account","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: PASSENGER TIRES (8.3%) | LOST: OUTSIDE PURCHASE (was $140) | DOWN: FARM TIRES (-$3287) | LOW GP: OFF THE ROAD TIRES (10.0%) | LOST: INDUSTRIAL TIRES (was $101)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200543,"customer":"GATOR TIRE","sales2025":4630.86,"sales2026":4326.49,"change":-304.37,"gpPct":0.245099,"action":"MAINTAIN - Stable account","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$813) | DOWN: ST TRAILER (-$343) | LOST: TUBES (was $10)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":2000012,"customer":"GLOBAL TRUCK & EQUIPMENT SALES","sales2025":440.0,"sales2026":9858.88,"change":9418.88,"gpPct":0.184705,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200836,"customer":"KING MUFFLER","sales2025":297.01,"sales2026":948.56,"change":651.55,"gpPct":0.293076,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $297)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200496,"customer":"MILLER TIRE CO.","sales2025":0.0,"sales2026":130.8,"change":130.8,"gpPct":0.278287,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200579,"customer":"RODS CAR & TRUCK ACC.","sales2025":339.56,"sales2026":764.68,"change":425.12,"gpPct":0.078203,"action":"LOW MARGIN - Reprice","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (4.2%)"},{"salesman":"Tiffany","city":"WHITE SPRINGS","custNum":200906,"customer":"E&H TIRE","sales2025":21757.91,"sales2026":85525.66,"change":63767.75,"gpPct":0.130649,"action":"GROWING - Expand","topDept":"OFF THE ROAD TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$870) | DOWN: TIRE TOOLS (-$24)"},{"salesman":"Unknown","city":"Unknown","custNum":2000049,"customer":"BEALL TIRE WHOLESALE, LLC","sales2025":0.0,"sales2026":13867.95,"change":13867.95,"gpPct":0.094026,"action":"LOW MARGIN - Reprice","topDept":"FARM TIRES","declinedDept":"","focus":"LOW GP: FARM TIRES (7.9%) | LOW GP: RAD LT TRUCK (2.7%) | LOW GP: OFF THE ROAD TIRES (8.0%)"},{"salesman":"Unknown","city":"Unknown","custNum":101196,"customer":"BEASLEY AUTO & TRUCK REPAIR","sales2025":0.0,"sales2026":85.0,"change":85.0,"gpPct":0.297647,"action":"GROWING - Expand","topDept":"WHEEL WEIGHTS","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":500420,"customer":"BLUE LEVEL SERVICES LLC","sales2025":1683.64,"sales2026":0.0,"change":-1683.64,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1684)"},{"salesman":"Unknown","city":"Unknown","custNum":100395,"customer":"BUBBA'S MOBILE TIRE","sales2025":1688.26,"sales2026":0.0,"change":-1688.26,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1688)"},{"salesman":"Unknown","city":"Unknown","custNum":102361,"customer":"C. MCDOWELL TRUCK PARTS & SRV.","sales2025":9600.0,"sales2026":0.0,"change":-9600.0,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $9600)"},{"salesman":"Unknown","city":"Unknown","custNum":101776,"customer":"COLUMBUS TIRE CO","sales2025":0.0,"sales2026":203.3,"change":203.3,"gpPct":0.583965,"action":"GROWING - Expand","topDept":"WHEELS","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":102273,"customer":"COMPLETE TIRE & SVC (CORDELE)","sales2025":88595.06,"sales2026":53532.32,"change":-35062.74,"gpPct":0.0536,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (7.9%) | DOWN: TRUCK TIRES (-$23380) | LOW GP: PASSENGER TIRES (4.9%) | DOWN: PASSENGER TIRES (-$8536) | LOW GP: RAD LT TRUCK (1.8%) | DOWN: RAD LT TRUCK (-$8082) | LOW GP: OFF THE ROAD TIRES (2.0%) | LOW GP: INDUSTRIAL TIRES (1.9%)"},{"salesman":"Unknown","city":"Unknown","custNum":102396,"customer":"CONTENDER SERVICE LLC","sales2025":0.0,"sales2026":49.36,"change":49.36,"gpPct":0.260332,"action":"GROWING - Expand","topDept":"TUBES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000064,"customer":"COUNTY LINE MOTORS LLC","sales2025":0.0,"sales2026":655.92,"change":655.92,"gpPct":0.281071,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":102319,"customer":"D&W ELITE AUTO SALES","sales2025":1063.86,"sales2026":0.0,"change":-1063.86,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1064)"},{"salesman":"Unknown","city":"Unknown","custNum":101115,"customer":"ECONOMY USED TIRE (WAREHOUSE)","sales2025":81136.01,"sales2026":0.0,"change":-81136.01,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $81136)"},{"salesman":"Unknown","city":"Unknown","custNum":2000063,"customer":"GATEWAY TIRE","sales2025":0.0,"sales2026":579.63,"change":579.63,"gpPct":-0.092179,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000069,"customer":"IRWIN COUNTY CUSTOMS & REPAIR","sales2025":0.0,"sales2026":870.35,"change":870.35,"gpPct":0.384006,"action":"GROWING - Expand","topDept":"WHEEL WEIGHTS","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":101691,"customer":"JK MOBILE SERVICE LLC","sales2025":746.49,"sales2026":0.0,"change":-746.49,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $746)"},{"salesman":"Unknown","city":"Unknown","custNum":2000053,"customer":"JOINER CONTRACTING","sales2025":0.0,"sales2026":2462.48,"change":2462.48,"gpPct":0.218081,"action":"GROWING - Expand","topDept":"ST TRAILER","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (6.7%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000046,"customer":"KELLEY MANUFACTURING CO.","sales2025":0.0,"sales2026":7964.46,"change":7964.46,"gpPct":0.18627,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":101984,"customer":"KOUNTRY BOI TIRES, LLC","sales2025":589.23,"sales2026":2458.2,"change":1868.97,"gpPct":0.112155,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000052,"customer":"L&A TIRE, LLC","sales2025":0.0,"sales2026":8504.08,"change":8504.08,"gpPct":0.170263,"action":"GROWING - Expand","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000059,"customer":"LAKE CITY TIRE SHOP","sales2025":0.0,"sales2026":437.1,"change":437.1,"gpPct":0.320705,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":100489,"customer":"LUMBER CITY ENT/BURKETT TIRE","sales2025":11569.42,"sales2026":2194.89,"change":-9374.53,"gpPct":0.103326,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: RAD LT TRUCK (9.6%) | LOST: TRUCK TIRES (was $2964) | LOST: OFF THE ROAD TIRES (was $4350) | LOST: TUBES (was $1939)"},{"salesman":"Unknown","city":"Unknown","custNum":100125,"customer":"MACON COMMERCIAL TIRE, INC.","sales2025":0.0,"sales2026":1246.54,"change":1246.54,"gpPct":0.122996,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":102035,"customer":"MARSHALLVILLE TIRE & LUBE","sales2025":2350.06,"sales2026":0.0,"change":-2350.06,"gpPct":0.0,"action":"LOST - Win back immediately","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2179) | LOST: PASSENGER TIRES (was $171)"},{"salesman":"Unknown","city":"Unknown","custNum":100165,"customer":"MOSLEY TIRE ALIGN.& BRAKE CTR.","sales2025":0.0,"sales2026":381.2,"change":381.2,"gpPct":0.163169,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000065,"customer":"MOSS MOTORS LLC","sales2025":0.0,"sales2026":861.16,"change":861.16,"gpPct":0.270472,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000060,"customer":"NISSAN OF TIFTON","sales2025":0.0,"sales2026":5062.02,"change":5062.02,"gpPct":0.172526,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (5.3%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000055,"customer":"PARKER TIRE - TIFTON","sales2025":0.0,"sales2026":7543.51,"change":7543.51,"gpPct":0.016559,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.8%)"},{"salesman":"Unknown","city":"Unknown","custNum":102049,"customer":"RAFFIELD TIRE (TRUCK CENTER)","sales2025":0.0,"sales2026":957.96,"change":957.96,"gpPct":0.127521,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000048,"customer":"SAN LUIS CAR SALES","sales2025":0.0,"sales2026":2123.13,"change":2123.13,"gpPct":0.184562,"action":"GROWING - Expand","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":101999,"customer":"SIMPLE TIRE (NAT. ACCT. D.R.)","sales2025":13617.74,"sales2026":1633.73,"change":-11984.01,"gpPct":-0.005019,"action":"DECLINING - Investigate","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (9.4%) | DOWN: RAD LT TRUCK (-$9712)"},{"salesman":"Unknown","city":"Unknown","custNum":101128,"customer":"SIMPLE TIRE - *BILLING ACCT*","sales2025":0.0,"sales2026":2021.67,"change":2021.67,"gpPct":0.040061,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (4.0%)"},{"salesman":"Unknown","city":"Unknown","custNum":102274,"customer":"SNIDER FLEET SOLUTIONS","sales2025":0.0,"sales2026":2487.74,"change":2487.74,"gpPct":0.155487,"action":"GROWING - Expand","topDept":"OFF THE ROAD TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":3000389,"customer":"SOUTHERN TIRE MART, LLC (#134)","sales2025":29199.15,"sales2026":148805.75,"change":119606.6,"gpPct":0.136735,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: RAD LT TRUCK (6.0%) | LOW GP: PASSENGER TIRES (4.3%) | DOWN: FARM TIRES (-$5655) | LOW GP: OFF THE ROAD TIRES (8.8%) | LOST: INDUSTRIAL TIRES (was $1099)"},{"salesman":"Unknown","city":"Unknown","custNum":2000068,"customer":"STONE ENTERPRISES - FAUSSETTS LLC","sales2025":0.0,"sales2026":541.83,"change":541.83,"gpPct":0.242235,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000051,"customer":"STONE'S OUTDOOR POWER CENTER, LLC","sales2025":0.0,"sales2026":4776.72,"change":4776.72,"gpPct":0.088404,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (1.4%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000050,"customer":"SUN BELT FLEET SERVICES","sales2025":0.0,"sales2026":2623.94,"change":2623.94,"gpPct":0.042783,"action":"LOW MARGIN - Reprice","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (5.9%) | LOW GP: PASSENGER TIRES (4.8%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000047,"customer":"SUNPOINT TIRES","sales2025":0.0,"sales2026":6019.22,"change":6019.22,"gpPct":0.081614,"action":"LOW MARGIN - Reprice","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.3%)"},{"salesman":"Unknown","city":"Unknown","custNum":100224,"customer":"WARD'S SERVICE CENTER","sales2025":0.0,"sales2026":1305.35,"change":1305.35,"gpPct":0.132026,"action":"GROWING - Expand","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000057,"customer":"WOLFES WAY LLC","sales2025":0.0,"sales2026":3984.65,"change":3984.65,"gpPct":0.13419,"action":"GROWING - Expand","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"}]};

const SEED_CUSTOMERS = [{"num":100952,"salesman":"House","active":"Inactive","name":"OLD ACCT","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":102060,"salesman":"House","active":"Inactive","name":"TIFTON PROSPECT ACCOUNT","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200330,"salesman":"House","active":"Inactive","name":"EAGLE AUTO REPAIR","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200390,"salesman":"House","active":"Inactive","name":"CLOSED","address":"0","city":"0","state":"","zip":"","phone":"912-487-5159","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200722,"salesman":"House","active":"Inactive","name":"TIFTON>>>MACON","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200723,"salesman":"House","active":"Inactive","name":"TIFTON>>>STATES","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200925,"salesman":"House","active":"Active","name":"TIFTON TO STATESBORO TRANSFER","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200928,"salesman":"House","active":"Active","name":"TIFTON TO BYRON TRANSFER","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":200929,"salesman":"House","active":"Active","name":"TIFTON TO ATHENS TRANSFER","address":"0","city":"0","state":"","zip":"","phone":"","ytdComp":0.0,"lat":null,"lon":null,"accuracy":0.0,"accuracyType":""},{"num":100722,"salesman":"House","active":"Inactive","name":"MR.TIRE","address":"2002 SCREVEN CIRCLE","city":"ABBEVILLE","state":"GA","zip":"31001","phone":"","ytdComp":0.0,"lat":32.998608,"lon":-82.40699,"accuracy":0.3,"accuracyType":"street_center"},{"num":101125,"salesman":"House","active":"Inactive","name":"KNIGHT'S TIRE SERVICE","address":"239 WILSON RD.","city":"ABBEVILLE","state":"GA","zip":"31001","phone":"229-467-2072","ytdComp":0.0,"lat":31.982332,"lon":-83.315617,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200366,"salesman":"Larry","active":"Active","name":"ROUNTREE PERFORMANCE","address":"2953 EVERGREEN CHURCH RD","city":"ADEL","state":"GA","zip":"31620","phone":"2295498399","ytdComp":1576.52,"lat":31.189518,"lon":-83.516225,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200804,"salesman":"Larry","active":"Active","name":"FIVE STAR TIRE ****ADEL****","address":"1403 WEST 4TH ST","city":"ADEL","state":"GA","zip":"31620","phone":"2298965381","ytdComp":8937.13,"lat":31.137383,"lon":-83.442737,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200332,"salesman":"House","active":"Inactive","name":"FAUSETTS INC","address":"301 S HUTCHINSON AVE","city":"ADEL","state":"GA","zip":"31620","phone":"229-896-7487","ytdComp":0.0,"lat":31.133997,"lon":-83.422185,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200393,"salesman":"House","active":"Active","name":"BRUISER'S TIRE & TOWING","address":"904 WEST 4TH ST","city":"ADEL","state":"GA","zip":"31620","phone":"2295615095","ytdComp":418.7,"lat":31.134582,"lon":-83.435116,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200400,"salesman":"House","active":"Inactive","name":"0","address":"0","city":"ADEL","state":"GA","zip":"31620","phone":"","ytdComp":0.0,"lat":30.920883,"lon":-83.32225,"accuracy":0.92,"accuracyType":"street_center"},{"num":200401,"salesman":"House","active":"Active","name":"ADEL TIRE CO","address":"215 N HUTCHINSON AVE","city":"ADEL","state":"GA","zip":"31620","phone":"2298963086","ytdComp":1209.32,"lat":31.139413,"lon":-83.424345,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200416,"salesman":"House","active":"Active","name":"DENT'S SERVICE STATION","address":"314 W 4TH ST","city":"ADEL","state":"GA","zip":"31620","phone":"2298964160","ytdComp":6035.25,"lat":31.136204,"lon":-83.42707,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200498,"salesman":"House","active":"Inactive","name":"SOUTHERN AUTOMOTIVE WHOLESALE","address":"506 W. 4TH ST.","city":"ADEL","state":"GA","zip":"31620","phone":"229-560-5987","ytdComp":0.0,"lat":31.135474,"lon":-83.429803,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200552,"salesman":"Car Dealer","active":"Active","name":"COOK COUNTY FORD  INC.","address":"1000 S. HUTCHINSON AVE.","city":"ADEL","state":"GA","zip":"31620","phone":"2298967411","ytdComp":0.0,"lat":31.125564,"lon":-83.419557,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200620,"salesman":"House","active":"Inactive","name":"REAVES OFFROAD POWERSPORTS","address":"710 OLD QUITMAN RD","city":"ADEL","state":"GA","zip":"31620","phone":"2292325147","ytdComp":0.0,"lat":31.114371,"lon":-83.425638,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200647,"salesman":"House","active":"Inactive","name":"FAUSETTS TIRE CO.","address":"301 S HUTCHINSON AVE","city":"ADEL","state":"GA","zip":"31620","phone":"2298967481","ytdComp":-100.68,"lat":31.133997,"lon":-83.422185,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200764,"salesman":"House","active":"Inactive","name":"JERRY KELLEY GMC","address":"1100 S HUTCHINSON AVE.","city":"ADEL","state":"GA","zip":"31620","phone":"229-299-5042","ytdComp":0.0,"lat":31.122298,"lon":-83.418184,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200794,"salesman":"House","active":"Inactive","name":"JAY'S TIRE","address":"506 WEST 4TH ST.","city":"ADEL","state":"GA","zip":"31620","phone":"2298965469","ytdComp":0.0,"lat":31.135474,"lon":-83.429803,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200822,"salesman":"House","active":"Inactive","name":"UNITED TIRE LLC","address":"267 HARRELL RD","city":"ADEL","state":"GA","zip":"31620","phone":"2295070790","ytdComp":0.0,"lat":31.11043,"lon":-83.468331,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200845,"salesman":"House","active":"Inactive","name":"COOK CO. FORD INC. (AMI ACCT)","address":"1000 S HUTCHINSON AVE","city":"ADEL","state":"GA","zip":"31620","phone":"2298967411","ytdComp":0.0,"lat":31.125564,"lon":-83.419557,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200881,"salesman":"House","active":"Inactive","name":"EDGE REPAIR","address":"1100 S. ELM ST","city":"ADEL","state":"GA","zip":"31620","phone":"2292568073","ytdComp":0.0,"lat":31.148466,"lon":-83.433748,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200958,"salesman":"House","active":"Active","name":"BULLARD DIESEL & AUTO","address":"6980 VAL DEL RD","city":"ADEL","state":"GA","zip":"31620","phone":"2292567594","ytdComp":355.62,"lat":31.109739,"lon":-83.377613,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200980,"salesman":"House","active":"Active","name":"SLYDER'S GARAGE","address":"272 ELK LANE","city":"ADEL","state":"GA","zip":"31620","phone":"2295600388","ytdComp":283.84,"lat":31.109958,"lon":-83.471086,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201037,"salesman":"House","active":"Active","name":"AFTER HOURS TIRE SERVICE","address":"267 HARRELL RD","city":"ADEL","state":"GA","zip":"31620","phone":"2295896702","ytdComp":3008.15,"lat":31.11043,"lon":-83.468331,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000035,"salesman":"House","active":"Active","name":"CITY BAYS ALACHUA","address":"14570 NW US HWY 441","city":"ALACHUA","state":"FL","zip":"32615","phone":"3864623887","ytdComp":0.0,"lat":29.790822,"lon":-82.483247,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000036,"salesman":"House","active":"Active","name":"ROGERS TIRE & REPAIR","address":"15251 NW US 441","city":"ALACHUA","state":"FL","zip":"32615","phone":"3864623700","ytdComp":0.0,"lat":29.795236,"lon":-82.49616,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101201,"salesman":"House","active":"Active","name":"TUCKERS SERVICE STATION","address":"8928 HWY 82","city":"ALAPAHA","state":"GA","zip":"31622","phone":"2295326097","ytdComp":26961.33,"lat":31.390115,"lon":-83.229723,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":100301,"salesman":"Larry","active":"Active","name":"ALBANY GENERAL TIRE SERVICE","address":"1002 W. BROAD AVE.","city":"ALBANY","state":"GA","zip":"31702","phone":"2294362485","ytdComp":11852.47,"lat":31.577051,"lon":-84.176309,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":100967,"salesman":"Larry","active":"Active","name":"SOUTHEASTERN COMMERCIAL TIRE","address":"1116 SEMINOLE LN","city":"ALBANY","state":"GA","zip":"31701","phone":"2298883300","ytdComp":6087.99,"lat":31.552436,"lon":-84.17924,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101479,"salesman":"Larry","active":"Active","name":"BILL THOMPSON TIRE SERVICES","address":"200 E OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"2294357753","ytdComp":644.0,"lat":31.575037,"lon":-84.143258,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200214,"salesman":"Larry","active":"Active","name":"PETERSON TIRE & AUTO CENTER","address":"801 SOUTH SLAPPEY","city":"ALBANY","state":"GA","zip":"31701","phone":"2293525136","ytdComp":459.36,"lat":31.566524,"lon":-84.176704,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100753,"salesman":"House","active":"Inactive","name":"CAM COMPANY  INC.","address":"1007 E. OGLETHORPE","city":"ALBANY","state":"GA","zip":"31705","phone":"229-883-1575","ytdComp":0.0,"lat":31.573124,"lon":-84.127567,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200239,"salesman":"Larry","active":"Active","name":"T & S TIRE","address":"948 W BROAD AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"2298880696","ytdComp":2688.72,"lat":31.57658,"lon":-84.173347,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200288,"salesman":"Larry","active":"Active","name":"RIGHT CHOICE AUTO","address":"806 BYRON RD","city":"ALBANY","state":"GA","zip":"31701","phone":"","ytdComp":0.0,"lat":31.590592,"lon":-84.29093,"accuracy":0.69,"accuracyType":"street_center"},{"num":200592,"salesman":"Larry","active":"Active","name":"BERNEY'S TIRE SERVICE","address":"1105 MOULTRIE RD","city":"ALBANY","state":"GA","zip":"31721","phone":"2294350412","ytdComp":7461.9,"lat":31.552359,"lon":-84.124639,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200601,"salesman":"Larry","active":"Active","name":"TUFF ENTERPRISES LLC","address":"1918 LEDO RD","city":"ALBANY","state":"GA","zip":"31706","phone":"2298838700","ytdComp":520.48,"lat":31.622109,"lon":-84.189149,"accuracy":0.99,"accuracyType":"rooftop"},{"num":101020,"salesman":"House","active":"Inactive","name":"1ST CLASS TIRE & WHEEL LLC.","address":"819A 21ST AVE.","city":"ALBANY","state":"GA","zip":"31701","phone":"229-496-9422","ytdComp":0.0,"lat":31.616412,"lon":-84.173253,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200644,"salesman":"Larry","active":"Active","name":"SKIP'S AUTOMOTIVE","address":"525 W. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294851319","ytdComp":158.96,"lat":31.575136,"lon":-84.159698,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200679,"salesman":"Larry","active":"Active","name":"RICHARD'S AUTO CARE & TIRE SVC","address":"928 W. HIGHLAND AVE.","city":"ALBANY","state":"GA","zip":"31701","phone":"2294078600","ytdComp":3978.92,"lat":31.573459,"lon":-84.173209,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200752,"salesman":"Larry","active":"Active","name":"RAINEY USED CARS (ALBANY)","address":"706 E. OGLETHORPE BLVD.","city":"ALBANY","state":"GA","zip":"31705","phone":"2296339860","ytdComp":2117.9,"lat":31.57287,"lon":-84.134095,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200939,"salesman":"Larry","active":"Active","name":"RNR TIRE EXPRESS","address":"1407 DAWSON RD","city":"ALBANY","state":"GA","zip":"31707","phone":"2295181969","ytdComp":23219.67,"lat":31.587923,"lon":-84.185228,"accuracy":1.0,"accuracyType":"rooftop"},{"num":500373,"salesman":"Larry","active":"Active","name":"TOMAHAWK TIRE (ALBANY)","address":"1401 RADIUM SPRINGS RD","city":"ALBANY","state":"GA","zip":"31705","phone":"2294396594","ytdComp":4421.61,"lat":31.553881,"lon":-84.138035,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101369,"salesman":"House","active":"Inactive","name":"CLOSED","address":"270 GRAND ISLAND DR.","city":"ALBANY","state":"GA","zip":"31707","phone":"229-878-1800","ytdComp":0.0,"lat":31.626425,"lon":-84.205411,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101497,"salesman":"House","active":"Active","name":"PERFORMANCE MOTORSPORT","address":"213 RADIUM SPRING RD","city":"ALBANY","state":"GA","zip":"31705","phone":"2294385248","ytdComp":0.0,"lat":31.574524,"lon":-84.142473,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101521,"salesman":"House","active":"Inactive","name":"GOODLIFE TIRE","address":"501 W OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"229-435-1111","ytdComp":0.0,"lat":31.57567,"lon":-84.159336,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101522,"salesman":"House","active":"Inactive","name":"CLEVE WESTER'S TIRE MART","address":"525 W OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"229-435-7788","ytdComp":0.0,"lat":31.575136,"lon":-84.159698,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101523,"salesman":"House","active":"Inactive","name":"CLOSED ACCOUNT","address":"0","city":"ALBANY","state":"GA","zip":"31705","phone":"229-435-0412","ytdComp":0.0,"lat":31.52864,"lon":-83.84709,"accuracy":0.94,"accuracyType":"street_center"},{"num":200134,"salesman":"House","active":"Active","name":"ACREE AUTO SALES","address":"1634 LIBERTY EXPRESSWAY","city":"ALBANY","state":"GA","zip":"31705","phone":"2294365174","ytdComp":0.0,"lat":31.498501,"lon":-84.114826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200136,"salesman":"House","active":"Inactive","name":"SUNBELT FORD ALBANY  INC (AMI)","address":"2926 N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31702","phone":"2298833100","ytdComp":0.0,"lat":31.622577,"lon":-84.175633,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200137,"salesman":"House","active":"Inactive","name":"ALBANY HONDA","address":"903 E OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-432-9700","ytdComp":0.0,"lat":31.573736,"lon":-84.13036,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200142,"salesman":"House","active":"Inactive","name":"BERNEY'S TIRE SERVICE","address":"1105 MOULTRIE RD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-435-0412","ytdComp":0.0,"lat":31.552359,"lon":-84.124639,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200172,"salesman":"House","active":"Inactive","name":"ENGINE DIST INC","address":"408 S WESTOVER RD","city":"ALBANY","state":"GA","zip":"31707","phone":"912-432-7468","ytdComp":0.0,"lat":31.570455,"lon":-84.220635,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200197,"salesman":"House","active":"Inactive","name":"NEWTON CROUCH  INC","address":"1110 LIBERTY EXPRESSWAY","city":"ALBANY","state":"GA","zip":"31705","phone":"229-436-1100","ytdComp":0.0,"lat":31.525043,"lon":-84.114632,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200199,"salesman":"House","active":"Inactive","name":"EXPRESSWAY TIRE SERVICE","address":"1007 OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"229-432-5901","ytdComp":0.0,"lat":31.575615,"lon":-84.174299,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200213,"salesman":"House","active":"Active","name":"PREMIER AUTOWORKS","address":"2917 N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294352886","ytdComp":458.0,"lat":31.622024,"lon":-84.177455,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200234,"salesman":"House","active":"Inactive","name":"SWAIN HEAVY EQUIPMENT","address":"1718 W OAKRIDGE DR","city":"ALBANY","state":"GA","zip":"31706","phone":"800-554-1428","ytdComp":0.0,"lat":31.548214,"lon":-84.206656,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200240,"salesman":"House","active":"Inactive","name":"TANNER DEEN NISSAN","address":"1100 E OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-436-8806","ytdComp":0.0,"lat":31.571,"lon":-84.125481,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200274,"salesman":"House","active":"Inactive","name":"HOOVER DIESEL AUTO & TIRE","address":"1215 GILLIONVILLE RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-436-9883","ytdComp":0.0,"lat":31.58145,"lon":-84.185247,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200289,"salesman":"House","active":"Inactive","name":"S&S ALTERATIONS","address":"1822 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-883-5301","ytdComp":0.0,"lat":31.622452,"lon":-84.186767,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200295,"salesman":"Car Dealer","active":"Active","name":"TOYOTA OF ALBANY","address":"2865 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"2294367751","ytdComp":0.0,"lat":31.623356,"lon":-84.216653,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200304,"salesman":"House","active":"Inactive","name":"ALBANY BODYWORKS","address":"1219 GILLIONVILLE RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-888-1934","ytdComp":0.0,"lat":31.581234,"lon":-84.185774,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200305,"salesman":"House","active":"Inactive","name":"LINCOLN OF ALBANY","address":"632 W BROAD AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"229-432-7464","ytdComp":0.0,"lat":31.576448,"lon":-84.164424,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200311,"salesman":"House","active":"Active","name":"AUTOMOTIVE NECESSITIES","address":"2301 N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294341237","ytdComp":728.96,"lat":31.607349,"lon":-84.176238,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200313,"salesman":"House","active":"Inactive","name":"BEE LINE OF ALBANY","address":"1830","city":"ALBANY","state":"GA","zip":"31701","phone":"229-432-8988","ytdComp":0.0,"lat":31.620991,"lon":-84.178897,"accuracy":0.64,"accuracyType":"street_center"},{"num":200334,"salesman":"House","active":"Inactive","name":"FIVE STAR NISSAN ALBANY","address":"1100 E OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-436-8806","ytdComp":0.0,"lat":31.571,"lon":-84.125481,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200340,"salesman":"House","active":"Inactive","name":"HOLT AUTO SERVICE & TOWING","address":"1007 W OGLETHORPE","city":"ALBANY","state":"GA","zip":"31705","phone":"229-430-5648","ytdComp":0.0,"lat":31.575615,"lon":-84.174299,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200346,"salesman":"House","active":"Inactive","name":"KIA OF ALBANY","address":"700 E OGLETHORPE","city":"ALBANY","state":"GA","zip":"31705","phone":"229-446-4777","ytdComp":0.0,"lat":31.573045,"lon":-84.135007,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200357,"salesman":"House","active":"Active","name":"PONDER AUTO REPAIR","address":"838 PINE AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"2294851297","ytdComp":381.05,"lat":31.578446,"lon":-84.17045,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200360,"salesman":"Car Dealer","active":"Active","name":"PRINCE CHEVY BUICK GMC","address":"2701 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"2292994591","ytdComp":0.0,"lat":31.623393,"lon":-84.205144,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200364,"salesman":"House","active":"Inactive","name":"ROAD RUNNER TIRE & AUTO","address":"6402 NEWTON RD","city":"ALBANY","state":"GA","zip":"31721","phone":"229-485-1264","ytdComp":0.0,"lat":31.487058,"lon":-84.224383,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200372,"salesman":"House","active":"Inactive","name":"SPEED SHOP & TRUCK ACCESSORIES","address":"2755 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-888-7005","ytdComp":0.0,"lat":31.623199,"lon":-84.209691,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200376,"salesman":"House","active":"Inactive","name":"STREET'S TIRE & WRECKER","address":"1221 CLARK AVE","city":"ALBANY","state":"GA","zip":"31705","phone":"229-889-0487","ytdComp":0.0,"lat":31.580172,"lon":-84.123464,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200377,"salesman":"House","active":"Inactive","name":"SUPERIOR DISTRIBUTING","address":"(TUFF TRUCKIN)","city":"ALBANY","state":"GA","zip":"31707","phone":"229-883-0032","ytdComp":0.0,"lat":31.586919,"lon":-84.206431,"accuracy":1.0,"accuracyType":"place"},{"num":200379,"salesman":"House","active":"Inactive","name":"DEPOT QUICK LUBE","address":"122 PHILEMA RD","city":"ALBANY","state":"GA","zip":"31701","phone":"229-446-0604","ytdComp":0.0,"lat":31.609577,"lon":-84.152495,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200411,"salesman":"House","active":"Active","name":"DJ'S CAR WASH","address":"2535 STUART AVE","city":"ALBANY","state":"GA","zip":"31707","phone":"2298886262","ytdComp":0.0,"lat":31.613828,"lon":-84.21202,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200412,"salesman":"House","active":"Active","name":"GOODYEAR SERVICE CTR","address":"502 W BROAD AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"2294392285","ytdComp":0.0,"lat":31.577113,"lon":-84.158983,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200413,"salesman":"House","active":"Active","name":"GOODYEAR SERVICE CTR","address":"2537 STUART AVE","city":"ALBANY","state":"GA","zip":"31707","phone":"2298838800","ytdComp":0.0,"lat":31.61392,"lon":-84.212638,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200420,"salesman":"House","active":"Inactive","name":"TRANSPORT EQUIPMENT","address":"3416 SYLVESTER RD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-435-0254","ytdComp":0.0,"lat":31.572007,"lon":-84.049097,"accuracy":0.8,"accuracyType":"rooftop"},{"num":200423,"salesman":"House","active":"Active","name":"UNVERFERTH MANUFACTURING CO","address":"5430 NEWTON RD","city":"ALBANY","state":"GA","zip":"31702","phone":"2298888062","ytdComp":0.0,"lat":31.520696,"lon":-84.195068,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200431,"salesman":"House","active":"Inactive","name":"HONEST TIRE REPAIR","address":"502 POPULAR ST","city":"ALBANY","state":"GA","zip":"31707","phone":"229-255-5077","ytdComp":0.0,"lat":31.570314,"lon":-84.179202,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200458,"salesman":"House","active":"Inactive","name":"LMC-AG  LLC","address":"1715 S SLAPPEY","city":"ALBANY","state":"GA","zip":"31701","phone":"229-639-1775","ytdComp":0.0,"lat":31.551152,"lon":-84.177448,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200494,"salesman":"House","active":"Inactive","name":"POPS DETAIL SHOP INC","address":"501 S SLAPPEY","city":"ALBANY","state":"GA","zip":"31701","phone":"229-296-4806","ytdComp":0.0,"lat":31.570571,"lon":-84.176535,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200522,"salesman":"House","active":"Inactive","name":"BROWN MILLING AND PEANUT CO.","address":"4634 GA. HWY 112 SOUTH","city":"ALBANY","state":"GA","zip":"31705","phone":"229-776-3391","ytdComp":0.0,"lat":31.400218,"lon":-83.977213,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200544,"salesman":"House","active":"Inactive","name":"CLOSED","address":"706 E. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-435-8114","ytdComp":0.0,"lat":31.57287,"lon":-84.134095,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200546,"salesman":"House","active":"Active","name":"SOWEGA TIRE OF ALBANY","address":"1219 SOUTH MOCK ROAD","city":"ALBANY","state":"GA","zip":"31705","phone":"2298881881","ytdComp":718.16,"lat":31.557674,"lon":-84.095035,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200550,"salesman":"House","active":"Inactive","name":"LIBERTY MOTORS","address":"2211 LIBERTY EXPRESSWAY","city":"ALBANY","state":"GA","zip":"31705","phone":"229-496-1323","ytdComp":0.0,"lat":31.474765,"lon":-84.118081,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200555,"salesman":"House","active":"Inactive","name":"FIVE STAR NISSAN","address":"1100 OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-436-8806","ytdComp":0.0,"lat":31.571,"lon":-84.125481,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200586,"salesman":"House","active":"Inactive","name":"STEEDLEY'S TRANSMISSION  INC","address":"625 OLGLETHORPE BLVD.","city":"ALBANY","state":"GA","zip":"31701","phone":"2294398001","ytdComp":0.0,"lat":31.575666,"lon":-84.164073,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200600,"salesman":"House","active":"Active","name":"CHARLOT TRUCKING & TIRE SVC.","address":"100 S. MAGNOLIA ST.","city":"ALBANY","state":"GA","zip":"31706","phone":"2298814561","ytdComp":0.0,"lat":31.576575,"lon":-84.190005,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200609,"salesman":"House","active":"Active","name":"LIBERTY AUTO CARE CENTER","address":"1116 LIBERTY EXPRESSWAY S.E.","city":"ALBANY","state":"GA","zip":"31705","phone":"2298834505","ytdComp":0.0,"lat":31.52386,"lon":-84.11495,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200632,"salesman":"House","active":"Inactive","name":"GIERYIC'S AUTO","address":"2401 DAWSON RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-432-0351","ytdComp":0.0,"lat":31.608949,"lon":-84.207678,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200645,"salesman":"House","active":"Inactive","name":"CLOSED-AUTO  ENHANCERS","address":"810 MOULTRIE RD","city":"ALBANY","state":"GA","zip":"31705","phone":"229-733-6870","ytdComp":0.0,"lat":31.554099,"lon":-84.130028,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200651,"salesman":"House","active":"Inactive","name":"TRAILERLAND EQUIPMENT LLC","address":"1939 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"229-639-0262","ytdComp":0.0,"lat":31.623213,"lon":-84.193804,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200652,"salesman":"House","active":"Inactive","name":"EASTSIDE GARAGE LLC","address":"104 LEXINGTON DR","city":"ALBANY","state":"GA","zip":"31705","phone":"229-432-7020","ytdComp":0.0,"lat":31.574668,"lon":-84.082332,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200655,"salesman":"House","active":"Active","name":"ECONOMY USED TIRE (ALBANY)","address":"429 W. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294052512","ytdComp":1236.63,"lat":31.575758,"lon":-84.158571,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200666,"salesman":"House","active":"Inactive","name":"CLOSED-4 WHEEL MAINTENANCE","address":"501 WEST BROAD STREET","city":"ALBANY","state":"GA","zip":"31701","phone":"229-573-7000","ytdComp":0.0,"lat":31.577222,"lon":-84.159269,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200675,"salesman":"House","active":"Active","name":"AUTO SOLUTIONS LLC","address":"3009-A N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2297334830","ytdComp":0.0,"lat":31.580346,"lon":-84.175586,"accuracy":0.92,"accuracyType":"rooftop"},{"num":200725,"salesman":"House","active":"Inactive","name":"MASTER TECH AUTOMOTIVE LLC","address":"624 FLINT AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"229-573-5373","ytdComp":0.0,"lat":31.579575,"lon":-84.163868,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200732,"salesman":"House","active":"Inactive","name":"KAUFFMAN TIRE (ALBANY)","address":"1018 PINE AVENUE","city":"ALBANY","state":"GA","zip":"31701","phone":"229-299-8544","ytdComp":0.0,"lat":31.580227,"lon":-84.174893,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200743,"salesman":"House","active":"Active","name":"SOUTHERN SALES & RENTALS  LLC","address":"1939 LEDO RD","city":"ALBANY","state":"GA","zip":"31707","phone":"2294059991","ytdComp":1726.71,"lat":31.623213,"lon":-84.193804,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200750,"salesman":"House","active":"Active","name":"LAWN PERFORMANCE  LLC","address":"2814 WILMAR LANE","city":"ALBANY","state":"GA","zip":"31707","phone":"2294355007","ytdComp":0.0,"lat":31.621726,"lon":-84.19242,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200756,"salesman":"House","active":"Active","name":"THE SHOP OF ALBANY  LLC","address":"511 W. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2295737066","ytdComp":1281.06,"lat":31.575833,"lon":-84.159921,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200761,"salesman":"House","active":"Active","name":"MARIO NEW AND USED TIRE SHOP","address":"2003 OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"2295915893","ytdComp":1823.66,"lat":31.572102,"lon":-84.100861,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200765,"salesman":"Car Dealer","active":"Active","name":"ALBANY CHRYSLER DODGE JEEP RAM","address":"701 E. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"2292337769","ytdComp":1508.67,"lat":31.574874,"lon":-84.134503,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200766,"salesman":"House","active":"Inactive","name":"TRANS POWER  INC., FORD","address":"2702 PALMYRA RD","city":"ALBANY","state":"GA","zip":"31701","phone":"229-883-6550","ytdComp":0.0,"lat":31.614512,"lon":-84.179181,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200795,"salesman":"House","active":"Active","name":"WINCHESTER PAINT & BODY","address":"512A S. SLAPPEY","city":"ALBANY","state":"GA","zip":"31701","phone":"2294308078","ytdComp":200.0,"lat":31.569803,"lon":-84.176433,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200798,"salesman":"House","active":"Inactive","name":"REAL DEAL TIRE & SERVICE INC.","address":"1007 E. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"9042192850","ytdComp":0.0,"lat":31.573124,"lon":-84.127567,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200811,"salesman":"House","active":"Active","name":"DRAPERS AUTO BODY REPAIR","address":"629 PINE AVE.","city":"ALBANY","state":"GA","zip":"31701","phone":"2298882369","ytdComp":0.0,"lat":31.578941,"lon":-84.164412,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200824,"salesman":"House","active":"Active","name":"ALBANY COMMERCIAL","address":"402 CORDELE RD","city":"ALBANY","state":"GA","zip":"31705","phone":"2294962397","ytdComp":0.0,"lat":31.583011,"lon":-84.084093,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200828,"salesman":"House","active":"Active","name":"HENRY'S ALIGNMNET","address":"940 PINE AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"2294962019","ytdComp":1492.1,"lat":31.578811,"lon":-84.170186,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200835,"salesman":"Car Dealer","active":"Inactive","name":"SUNBELT FORD ALBANY","address":"2626 N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31702","phone":"2298833100","ytdComp":0.0,"lat":31.614843,"lon":-84.174572,"accuracy":0.88,"accuracyType":"nearest_rooftop_match"},{"num":200837,"salesman":"House","active":"Inactive","name":"ANYTIME TIRE SERVICE","address":"1701 EAST BROAD","city":"ALBANY","state":"GA","zip":"31705","phone":"2293499087","ytdComp":0.0,"lat":31.575954,"lon":-84.106823,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200838,"salesman":"Car Dealer","active":"Active","name":"ALBANY CHRYSLER DODGE JEEP RAM","address":"701 E. OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31706","phone":"2292337769","ytdComp":341.08,"lat":31.574874,"lon":-84.134503,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200863,"salesman":"House","active":"Inactive","name":"TRANS POWER  INC. FORD (AMI)","address":"2702 PALMYRA RD","city":"ALBANY","state":"GA","zip":"31701","phone":"2298836550","ytdComp":0.0,"lat":31.614512,"lon":-84.179181,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200870,"salesman":"House","active":"Active","name":"D&K USED TIRES","address":"911 N SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31707","phone":"2294961487","ytdComp":1095.4,"lat":31.589921,"lon":-84.176326,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200873,"salesman":"House","active":"Active","name":"ERICKSON AUTOMOTIVE","address":"1019 W. OGLETHORPE","city":"ALBANY","state":"GA","zip":"31701","phone":"2294325051","ytdComp":0.0,"lat":31.575754,"lon":-84.174967,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200893,"salesman":"House","active":"Inactive","name":"SOUTHERN TIRE MART @ PILOT(NA)","address":"310 CORDELE ROAD","city":"ALBANY","state":"GA","zip":"31705","phone":"2298781355","ytdComp":0.0,"lat":31.580644,"lon":-84.086102,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200949,"salesman":"House","active":"Active","name":"SOUTHERN TIRE MART","address":"1314 MOULTRIE RD","city":"ALBANY","state":"GA","zip":"31705","phone":"2299203030","ytdComp":0.0,"lat":31.542527,"lon":-84.111159,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201021,"salesman":"House","active":"Active","name":"ALBANY USED TIRES BRK & ALIGN.","address":"900 WEST OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294389391","ytdComp":0.0,"lat":31.574656,"lon":-84.170991,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201025,"salesman":"House","active":"Active","name":"A-1 WRECKER SERVICE","address":"213 7TH AVE","city":"ALBANY","state":"GA","zip":"31701","phone":"2294363990","ytdComp":0.0,"lat":31.596178,"lon":-84.153441,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201034,"salesman":"House","active":"Active","name":"ECONOMIC NICHOLAS TIRE","address":"833 W. OGLETHORPE BLVD.","city":"ALBANY","state":"GA","zip":"31701","phone":"2294545123","ytdComp":0.0,"lat":31.57575,"lon":-84.169456,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201049,"salesman":"House","active":"Inactive","name":"BROUSSARD ACCESSORIES LLC","address":"1402 SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2294968004","ytdComp":0.0,"lat":31.595787,"lon":-84.175824,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201056,"salesman":"House","active":"Active","name":"LEE'S AUTO SHOP","address":"2325 GILLONVILLE RD","city":"ALBANY","state":"GA","zip":"31707","phone":"8133684143","ytdComp":1149.48,"lat":31.586056,"lon":-84.207708,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201101,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 658","address":"1018 PINE AVENUE","city":"ALBANY","state":"GA","zip":"31701","phone":"2293528491","ytdComp":0.0,"lat":31.580227,"lon":-84.174893,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201200,"salesman":"House","active":"Active","name":"FIRESTONE STORE #020702","address":"2530 DAWSON RD","city":"ALBANY","state":"GA","zip":"31707","phone":"2294397707","ytdComp":0.0,"lat":31.611751,"lon":-84.214188,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000008,"salesman":"House","active":"Active","name":"ALBANY MOTORCARS","address":"805 EAST OGLETHORPE BLVD","city":"ALBANY","state":"GA","zip":"31705","phone":"2294362369","ytdComp":0.0,"lat":31.574209,"lon":-84.131408,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000026,"salesman":"Car Dealer","active":"Active","name":"FORD ALBANY","address":"2926 N/ SLAPPEY BLVD","city":"ALBANY","state":"GA","zip":"31701","phone":"2298833100","ytdComp":0.0,"lat":31.560674,"lon":-84.176751,"accuracy":0.5,"accuracyType":"place"},{"num":2000032,"salesman":"House","active":"Active","name":"J'S TIRE & MORE","address":"205 N. CARROLL ST.","city":"ALBANY","state":"GA","zip":"31705","phone":"2298699531","ytdComp":0.0,"lat":31.577975,"lon":-84.119654,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200974,"salesman":"House","active":"Active","name":"PRIORITY TIRE (TIFTON WHSE)","address":"1436 ECK RD.","city":"ALLENTOWN","state":"PA","zip":"18104","phone":"8664400177","ytdComp":22246.78,"lat":40.602039,"lon":-75.577191,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100438,"salesman":"House","active":"Inactive","name":"CLOSED","address":"127 DOGWOOD DR","city":"ALMA","state":"GA","zip":"31510","phone":"912-632-4359","ytdComp":0.0,"lat":31.507207,"lon":-82.461953,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101507,"salesman":"Larry","active":"Active","name":"BURNETTE AUTOMOTIVE SERVICE","address":"605 S PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"9126322713","ytdComp":9573.42,"lat":31.537473,"lon":-82.463803,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200920,"salesman":"Larry","active":"Active","name":"PRECISION MAINTENANCE","address":"140 GA-32 W BYPASS","city":"ALMA","state":"GA","zip":"31510","phone":"9122535237","ytdComp":7921.17,"lat":31.539256,"lon":-82.461836,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000022,"salesman":"Larry","active":"Active","name":"R&R AUTO SERVICE & REPAIR","address":"1976 HWY 32 WEST","city":"ALMA","state":"GA","zip":"31510","phone":"9123475301","ytdComp":1359.95,"lat":31.54435,"lon":-82.519198,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101435,"salesman":"House","active":"Inactive","name":"BELL'S  SERVICE STATION","address":"101 S PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"912-632-7133","ytdComp":0.0,"lat":31.541544,"lon":-82.461927,"accuracy":0.93,"accuracyType":"rooftop"},{"num":101476,"salesman":"House","active":"Active","name":"HART'S SERVICE STATION","address":"302 S PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"9126325180","ytdComp":0.0,"lat":31.539654,"lon":-82.462346,"accuracy":0.93,"accuracyType":"rooftop"},{"num":101510,"salesman":"House","active":"Inactive","name":"MIKE'S CORNER AUTO SALES  LLC","address":"410 W 16TH ST","city":"ALMA","state":"GA","zip":"31510","phone":"912-632-6448","ytdComp":0.0,"lat":31.53949,"lon":-82.468152,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101514,"salesman":"House","active":"Inactive","name":"YESTERDAYS TIRE TOWN","address":"934 S PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"9127229611","ytdComp":0.0,"lat":31.533118,"lon":-82.463506,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200711,"salesman":"House","active":"Active","name":"DIXON'S LUBE & TUBE","address":"1035 W. 12TH ST.","city":"ALMA","state":"GA","zip":"31510","phone":"9126322610","ytdComp":0.0,"lat":31.54355,"lon":-82.483241,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200825,"salesman":"House","active":"Active","name":"A&L AUTO SALES","address":"1038 W. 12TH ST.","city":"ALMA","state":"GA","zip":"31510","phone":"9122823157","ytdComp":0.0,"lat":31.542419,"lon":-82.483803,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200892,"salesman":"House","active":"Active","name":"STEVENS WHOLESALE TIRE DIST.","address":"934 S. PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"9125016573","ytdComp":0.0,"lat":31.533118,"lon":-82.463506,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200897,"salesman":"House","active":"Active","name":"ALMA TIRE & AUTO REPAIR","address":"937 S. PIERCE ST","city":"ALMA","state":"GA","zip":"31510","phone":"9122869140","ytdComp":5868.14,"lat":31.532615,"lon":-82.464317,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200985,"salesman":"House","active":"Active","name":"ROBERTS AUTO SERVICE LLC","address":"1976 HWY 32 W","city":"ALMA","state":"GA","zip":"31510","phone":"9126329942","ytdComp":0.0,"lat":31.54435,"lon":-82.519198,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201015,"salesman":"House","active":"Active","name":"LUBE KING & TIRES","address":"1035 W 12TH ST","city":"ALMA","state":"GA","zip":"31510","phone":"9126322610","ytdComp":1812.36,"lat":31.54355,"lon":-82.483241,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201071,"salesman":"House","active":"Active","name":"CAR MOD GUY LLC","address":"11175 CICERO DR","city":"ALPHARETTA","state":"GA","zip":"30027","phone":"4045902103","ytdComp":0.0,"lat":34.051596,"lon":-84.278464,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200598,"salesman":"House","active":"Inactive","name":"MARKHAM FARMS","address":"589 EBENEZER RD","city":"AMITY","state":"AR","zip":"71921","phone":"870-356-8434","ytdComp":0.0,"lat":34.274285,"lon":-93.599173,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":100622,"salesman":"House","active":"Inactive","name":"RANKIN PATE AUTO REPAIR SVC","address":"4496 2ND AVE","city":"ARABI","state":"GA","zip":"31712","phone":"229-273-6116","ytdComp":0.0,"lat":31.826858,"lon":-83.737953,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200962,"salesman":"House","active":"Active","name":"GREENE'S TIRE SERVICE LLC","address":"4530 INDUSTRIAL BLVD","city":"ARABI","state":"GA","zip":"31712","phone":"2296990748","ytdComp":5547.33,"lat":31.829278,"lon":-83.734012,"accuracy":0.72,"accuracyType":"nearest_rooftop_match"},{"num":200182,"salesman":"House","active":"Inactive","name":"HERRING TIRE SERVICE","address":"PO BOX 172","city":"ARITON","state":"AL","zip":"36311","phone":"334-762-2715","ytdComp":0.0,"lat":31.60017,"lon":-85.71883,"accuracy":1.0,"accuracyType":"place"},{"num":200322,"salesman":"House","active":"Inactive","name":"CHARLINE PARTS & REPAIR","address":"25703 BLAKELY RD","city":"ARLINGTON","state":"GA","zip":"39813","phone":"229-725-3343","ytdComp":0.0,"lat":31.427507,"lon":-84.727767,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200422,"salesman":"House","active":"Inactive","name":"WES GILLESPIE","address":"3988 NEW HOPE CHURCH RD","city":"ASHBORO","state":"NC","zip":"27205","phone":"336-953-6112","ytdComp":0.0,"lat":35.585263,"lon":-79.819527,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101283,"salesman":"Larry","active":"Active","name":"CAMERON'S TOWING AND TIRE","address":"220 N MAIN ST","city":"ASHBURN","state":"GA","zip":"31714","phone":"2295672437","ytdComp":32128.05,"lat":31.707754,"lon":-83.654007,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101416,"salesman":"House","active":"Inactive","name":"G & S 66","address":"341 E WASHINGTON","city":"ASHBURN","state":"GA","zip":"31714","phone":"229-567-3280","ytdComp":0.0,"lat":31.707742,"lon":-83.649116,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200329,"salesman":"House","active":"Inactive","name":"(BUSINESS SOLD) ASHBURN TIRE","address":"217 S MAIN ST","city":"ASHBURN","state":"GA","zip":"31714","phone":"2296130007","ytdComp":0.0,"lat":31.704365,"lon":-83.652267,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200623,"salesman":"House","active":"Inactive","name":"TURNER CO. BD. OF EDUCATION","address":"423 N. CLEVELAND ST.","city":"ASHBURN","state":"GA","zip":"31714","phone":"229-520-7564","ytdComp":0.0,"lat":31.709496,"lon":-83.648665,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200767,"salesman":"Car Dealer","active":"Active","name":"PARKER CHEVROLET BUICK GMC","address":"517 GORDAY DR","city":"ASHBURN","state":"GA","zip":"31714","phone":"2295131011","ytdComp":0.0,"lat":31.71146,"lon":-83.642373,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200768,"salesman":"House","active":"Inactive","name":"SOUTHLAND FORD INC","address":"801 NORTH ST","city":"ASHBURN","state":"GA","zip":"31714","phone":"229-567-3301","ytdComp":0.0,"lat":31.715311,"lon":-83.65462,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200803,"salesman":"House","active":"Active","name":"SOUTH MAIN GARAGE","address":"821 SOUTH MAIN","city":"ASHBURN","state":"GA","zip":"31714","phone":"2295663880","ytdComp":906.29,"lat":31.696362,"lon":-83.648276,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200812,"salesman":"House","active":"Inactive","name":"REAL DEAL ASHBURN","address":"217 S MAIN ST","city":"ASHBURN","state":"GA","zip":"31714","phone":"9042192850","ytdComp":0.0,"lat":31.704365,"lon":-83.652267,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200858,"salesman":"House","active":"Inactive","name":"SOUTHLAND FORD INC (AMI ACCT)","address":"801 NORTH ST","city":"ASHBURN","state":"GA","zip":"31714","phone":"2295673301","ytdComp":0.0,"lat":31.715311,"lon":-83.65462,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200945,"salesman":"House","active":"Active","name":"ASHBURN TIRE MART  LLC","address":"217 S. MAIN ST.","city":"ASHBURN","state":"GA","zip":"31714","phone":"2296130007","ytdComp":0.0,"lat":31.704365,"lon":-83.652267,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201007,"salesman":"House","active":"Inactive","name":"SHORTY HUGHES TRUCKING  LLC","address":"144 WEST END AVE","city":"ASHBURN","state":"GA","zip":"31714","phone":"2295672499","ytdComp":15569.81,"lat":31.709134,"lon":-83.657322,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200183,"salesman":"House","active":"Inactive","name":"KEY TIRE CENTER","address":"1916 OLD HWY 84","city":"ASHFORD","state":"AL","zip":"36312","phone":"3348998950","ytdComp":0.0,"lat":31.178056,"lon":-85.222076,"accuracy":0.7,"accuracyType":"street_center"},{"num":200742,"salesman":"House","active":"Active","name":"SOUTHERN SALES & RENTALS  LLC","address":"930 NEWTON BRIDGE RD","city":"ATHENS","state":"GA","zip":"30607","phone":"7065469760","ytdComp":0.0,"lat":34.006392,"lon":-83.401255,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200639,"salesman":"House","active":"Inactive","name":"GCR TIRE","address":"3803 JACK SPRINGS RD","city":"ATMORE","state":"AL","zip":"36502","phone":"2514468174","ytdComp":0.0,"lat":31.079547,"lon":-87.524548,"accuracy":0.85,"accuracyType":"nearest_rooftop_match"},{"num":2000024,"salesman":"Car Dealer","active":"Active","name":"FORKLIFT TIRE OF CENTRAL FL","address":"4863 JULIANA RESERVE DR","city":"AUBURNDALE","state":"FL","zip":"33823","phone":"8635599353","ytdComp":17505.64,"lat":28.135665,"lon":-81.81186,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201060,"salesman":"House","active":"Active","name":"CENTRAL FLORIDA TIRE TERMINAL","address":"2400 US HWY 27 S","city":"AVON PARK","state":"FL","zip":"33826","phone":"8634532030","ytdComp":0.0,"lat":27.564315,"lon":-81.514255,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200184,"salesman":"Larry","active":"Active","name":"JIMMY'S AUTO SALES","address":"1902 DOTHAN RD","city":"BAINBRIDGE","state":"GA","zip":"31817","phone":"2292431357","ytdComp":0.0,"lat":30.922405,"lon":-84.605289,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200157,"salesman":"House","active":"Inactive","name":"DARRINS AUTO & TIRE SERVICE","address":"200 W CALHOUN ST","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"229-246-0804","ytdComp":0.0,"lat":30.910189,"lon":-84.57701,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200162,"salesman":"House","active":"Active","name":"DELTA TIRE CO","address":"1616 DOTHAN RD","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"2292462750","ytdComp":15591.22,"lat":30.91873,"lon":-84.597752,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200179,"salesman":"House","active":"Inactive","name":"JE SHARBER OIL COMPANY","address":"250 WHIGHAM DAIRY RD","city":"BAINBRIDGE","state":"GA","zip":"39818","phone":"229-246-2183","ytdComp":0.0,"lat":30.913661,"lon":-84.53539,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200230,"salesman":"House","active":"Active","name":"SOUTHERN TIRE & BATTERY","address":"1129 E SHOTWELL","city":"BAINBRIDGE","state":"GA","zip":"39818","phone":"2292464925","ytdComp":0.0,"lat":30.904169,"lon":-84.559885,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200243,"salesman":"House","active":"Inactive","name":"TIM'S AUTOMOTIVE","address":"418 WIGHAM DAIRY RD","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"229-243-8867","ytdComp":0.0,"lat":30.926175,"lon":-84.535224,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200284,"salesman":"House","active":"Inactive","name":"PEACHSTATE DIESEL & ALIGNMENT","address":"0","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"229-246-0689","ytdComp":0.0,"lat":30.89087,"lon":-84.573277,"accuracy":0.94,"accuracyType":"street_center"},{"num":200297,"salesman":"House","active":"Active","name":"WHOLESALE BATTERY","address":"202 SPRING CREEK RD","city":"BAINBRIDGE","state":"GA","zip":"31717","phone":"2292468977","ytdComp":0.0,"lat":30.913966,"lon":-84.593262,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200361,"salesman":"House","active":"Inactive","name":"RATHEL'S AUTO PARTS","address":"917 COLQUITT HWY","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"229-243-8828","ytdComp":0.0,"lat":30.942211,"lon":-84.611728,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200384,"salesman":"House","active":"Inactive","name":"WILLIAMS AUTOMOTIVE","address":"1415 MLK JR DR","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"229-246-4075","ytdComp":0.0,"lat":30.908649,"lon":-84.55447,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200446,"salesman":"House","active":"Active","name":"BAINBRIDGE SERVICE CENTER","address":"2277 DOTHAN RD","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"2292468841","ytdComp":0.0,"lat":30.928281,"lon":-84.627147,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200481,"salesman":"House","active":"Inactive","name":"A-1 AUTO & EQUIPMENT","address":"1807 TALLAHASSEE HWY","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-246-9312","ytdComp":0.0,"lat":30.868274,"lon":-84.557819,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200493,"salesman":"House","active":"Inactive","name":"DAVIS & SONS","address":"710 FACEVILLE HWY","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-485-9650","ytdComp":0.0,"lat":30.891698,"lon":-84.584212,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200508,"salesman":"House","active":"Inactive","name":"WATSON TIRES","address":"0","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-416-4542","ytdComp":0.0,"lat":30.89087,"lon":-84.573277,"accuracy":1.0,"accuracyType":"street_center"},{"num":200640,"salesman":"House","active":"Inactive","name":"DW TRUCK AND TRAILER REPAIR","address":"619A FACEVILLE HWY","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-246-7513","ytdComp":0.0,"lat":30.888445,"lon":-84.581877,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200769,"salesman":"Car Dealer","active":"Active","name":"ACTION GM OF BAINBRIDGE","address":"2501 E STOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"2295154416","ytdComp":0.0,"lat":30.901863,"lon":-84.570049,"accuracy":0.5,"accuracyType":"place"},{"num":200770,"salesman":"House","active":"Inactive","name":"DEAN CHRYSLER JEEP DODGE","address":"1305 E. SHOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-515-4576","ytdComp":0.0,"lat":30.904306,"lon":-84.557172,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200771,"salesman":"House","active":"Inactive","name":"RIVER BEND FORD  INC","address":"1709 E. SHOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"229-246-0860","ytdComp":0.0,"lat":30.903704,"lon":-84.548029,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200846,"salesman":"House","active":"Inactive","name":"DEAN CDJR (AMI ACCOUNT)","address":"1305 E. SHOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"2295154576","ytdComp":0.0,"lat":30.904306,"lon":-84.557172,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200854,"salesman":"Car Dealer","active":"Active","name":"RIVER BEND FORD INC","address":"1709 E. SHOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"2292460860","ytdComp":0.0,"lat":30.903704,"lon":-84.548029,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000000,"salesman":"Car Dealer","active":"Inactive","name":"RIVER BEND CHRYSLER DODGE JEEP","address":"1305 E SHOTWELL ST","city":"BAINBRIDGE","state":"GA","zip":"39819","phone":"","ytdComp":0.0,"lat":30.904306,"lon":-84.557172,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000010,"salesman":"House","active":"Active","name":"SOUTHERN AUTOMOTIVE SVC & REP","address":"2277 DOTHAN RD","city":"BAINBRIDGE","state":"GA","zip":"39817","phone":"2292461333","ytdComp":366.0,"lat":30.928281,"lon":-84.627147,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101623,"salesman":"House","active":"Inactive","name":"LOTT'S E-Z OWN  INC","address":"627 S MAIN ST","city":"BAXLEY","state":"GA","zip":"31513","phone":"912-366-0500","ytdComp":0.0,"lat":31.769722,"lon":-82.351638,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201038,"salesman":"House","active":"Active","name":"GIGA TIRES  LLC (TIFTON WHSE)","address":"4562 E. 2ND ST.","city":"BENICIA","state":"CA","zip":"94510","phone":"8443470789","ytdComp":19701.68,"lat":38.081682,"lon":-122.134321,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200701,"salesman":"House","active":"Inactive","name":"TIRE ENGINEERS INC.","address":"3200 6TH AVE.","city":"BIRMINGHAM","state":"AL","zip":"35222","phone":"205-323-7282","ytdComp":0.0,"lat":33.515302,"lon":-86.784135,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200491,"salesman":"Larry","active":"Active","name":"JOT EM DOWN","address":"3425 US HWY 84 WEST","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124490095","ytdComp":0.0,"lat":31.301988,"lon":-82.245519,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200497,"salesman":"Larry","active":"Active","name":"MOBLEY AUTO CENTER","address":"418 GORDON STREET","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124494571","ytdComp":0.0,"lat":31.30742,"lon":-82.241854,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200465,"salesman":"House","active":"Active","name":"DIXON SERVICE CENTER","address":"527 GORDON ST.","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124496535","ytdComp":415.49,"lat":31.309419,"lon":-82.242471,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200474,"salesman":"Tiffany","active":"Active","name":"TIRE & WHEEL INC","address":"3599 E HWY 84","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124496164","ytdComp":2957.52,"lat":31.306276,"lon":-82.240466,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200513,"salesman":"House","active":"Inactive","name":"MUDSLINGERS 4WHEEL DRIVE","address":"845 SW CENTRAL AVE","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"912-449-5555","ytdComp":0.0,"lat":31.297467,"lon":-82.247193,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200539,"salesman":"Car Dealer","active":"Active","name":"MIKE BURCH FORD (BLACKSHEAR)","address":"304 WARE ST","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124494446","ytdComp":0.0,"lat":31.303883,"lon":-82.246498,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200712,"salesman":"Tiffany","active":"Active","name":"TANNER AUTO REPAIR PLUS  LLC","address":"845 S. W. CENTRAL AVE","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9128078277","ytdComp":0.0,"lat":31.297467,"lon":-82.247193,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200714,"salesman":"House","active":"Active","name":"C&S AUTO SERVICE INC.","address":"417 WARE ST.","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124493100","ytdComp":0.0,"lat":31.302466,"lon":-82.247527,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200883,"salesman":"Tiffany","active":"Active","name":"ROLLING BEAR TIRES LLC","address":"3154 US HWY 84","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9123876642","ytdComp":23383.3,"lat":31.294255,"lon":-82.255074,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200884,"salesman":"House","active":"Inactive","name":"PABLO'S TIRE SHOP","address":"803 GORDON ST","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9127227162","ytdComp":0.0,"lat":31.313207,"lon":-82.24232,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200885,"salesman":"Tiffany","active":"Active","name":"PIERCE INDUSTRIAL TIRE LLC","address":"PO BOX 312","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9128073685","ytdComp":0.0,"lat":31.326324,"lon":-82.271206,"accuracy":1.0,"accuracyType":"place"},{"num":200911,"salesman":"House","active":"Inactive","name":"CAR MART","address":"2710 WEST HWY 84","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9125843189","ytdComp":0.0,"lat":31.281084,"lon":-82.269578,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200982,"salesman":"House","active":"Active","name":"JJ & JAY'S TIRE & AUTO CARE","address":"3171 MIDWAY CHURCH RD","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"2529070899","ytdComp":0.0,"lat":31.247265,"lon":-82.267453,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201036,"salesman":"House","active":"Active","name":"GODWIN & SON REPAIR & SALES","address":"3091 HWY 84","city":"BLACKSHEAR","state":"GA","zip":"31516","phone":"9124527032","ytdComp":358.76,"lat":31.29174,"lon":-82.256651,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200144,"salesman":"House","active":"Active","name":"BROWNLEE SERVICE STATION INC.","address":"2150 S MAIN ST","city":"BLAKELY","state":"GA","zip":"39823","phone":"2297235855","ytdComp":0.0,"lat":31.374765,"lon":-84.933758,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200145,"salesman":"House","active":"Inactive","name":"BLAKELY TIRE INC","address":"PO BOX 545","city":"BLAKELY","state":"GA","zip":"39823","phone":"229-723-3160","ytdComp":0.0,"lat":31.346088,"lon":-84.901178,"accuracy":1.0,"accuracyType":"place"},{"num":200178,"salesman":"House","active":"Inactive","name":"HATTAWAY MOTOR CORP  FORD","address":"1503 S MAIN STREET","city":"BLAKELY","state":"GA","zip":"39823","phone":"2297233501","ytdComp":0.0,"lat":31.3648,"lon":-84.935177,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200205,"salesman":"House","active":"Inactive","name":"PERFORMANCE TIRE CENTER","address":"1493 N MAIN","city":"BLAKELY","state":"GA","zip":"39823","phone":"229-723-8473","ytdComp":0.0,"lat":31.40002,"lon":-84.932336,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200380,"salesman":"House","active":"Inactive","name":"THOMPSON MOTOR CO","address":"183 N MAIN","city":"BLAKELY","state":"GA","zip":"39823","phone":"229-723-3731","ytdComp":0.0,"lat":31.380806,"lon":-84.934587,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200633,"salesman":"House","active":"Inactive","name":"AUTO DOCTOR & EQUIPMENT","address":"965 N. CHURCH ST.","city":"BLAKELY","state":"GA","zip":"39823","phone":"229-723-7080","ytdComp":0.0,"lat":31.381267,"lon":-84.936254,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200772,"salesman":"Car Dealer","active":"Active","name":"BLAKELY CHEVROLET BUICK GMC","address":"183 N MAIN ST","city":"BLAKELY","state":"GA","zip":"39823","phone":"2297233731","ytdComp":0.0,"lat":31.380806,"lon":-84.934587,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200849,"salesman":"House","active":"Inactive","name":"HATTAWAY MOTOR CORP FORD (AMI)","address":"1503 S. MAIN ST.","city":"BLAKELY","state":"GA","zip":"39823","phone":"2297233501","ytdComp":0.0,"lat":31.3648,"lon":-84.935177,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200888,"salesman":"House","active":"Active","name":"EARLY AUTO PARTS","address":"87 MURDOCK AVE","city":"BLAKELY","state":"GA","zip":"39823","phone":"2297234121","ytdComp":0.0,"lat":31.380164,"lon":-84.934587,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200937,"salesman":"House","active":"Inactive","name":"TECHWAY AUTOMOTIVE 2","address":"1493 N. MAIN ST.","city":"BLAKELY","state":"GA","zip":"39823","phone":"3346993700","ytdComp":0.0,"lat":31.40002,"lon":-84.932336,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000016,"salesman":"Larry","active":"Active","name":"BEST VALUE COMMERCIAL TIRE","address":"19073 NE SR 64","city":"BLOUNTSTOWN","state":"FL","zip":"32424","phone":"8507929101","ytdComp":0.0,"lat":30.574608,"lon":-85.124075,"accuracy":0.46,"accuracyType":"street_center"},{"num":200559,"salesman":"House","active":"Inactive","name":"CITY TIRE","address":"19511 STATE ROAD 20 WEST","city":"BLOUNTSTOWN","state":"FL","zip":"32424","phone":"850-674-8784","ytdComp":0.0,"lat":30.44308,"lon":-85.064924,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200143,"salesman":"House","active":"Inactive","name":"ALL STAR TIRE","address":"304 W HWY 90","city":"BONIFAY","state":"FL","zip":"32425","phone":"850-547-2072","ytdComp":0.0,"lat":30.788272,"lon":-85.683465,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200222,"salesman":"House","active":"Inactive","name":"SON'S TIRE","address":"202 S WAUKESHA ST","city":"BONIFAY","state":"FL","zip":"32425","phone":"850-547-3388","ytdComp":0.0,"lat":30.79405,"lon":-85.6796,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200996,"salesman":"House","active":"Active","name":"UNEEK MOTORSPORTS","address":"3036 49TH AVE","city":"BRADENTON","state":"FL","zip":"34207","phone":"3136007100","ytdComp":0.0,"lat":27.455334,"lon":-82.592859,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201045,"salesman":"House","active":"Active","name":"AGTRAX LLC","address":"26474 CR49","city":"BRANFORD","state":"FL","zip":"32008","phone":"3369611171","ytdComp":0.0,"lat":29.95912,"lon":-82.92818,"accuracy":0.5,"accuracyType":"place"},{"num":200250,"salesman":"House","active":"Inactive","name":"WORTHINGTON FRONT END","address":"3234 BETHEL RD","city":"BRINSON","state":"GA","zip":"31725","phone":"229-246-7782","ytdComp":0.0,"lat":30.9953,"lon":-84.673829,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200726,"salesman":"House","active":"Active","name":"DK DIESEL & TIRE  LLC","address":"2020 DOTHAN RD","city":"BRINSON","state":"GA","zip":"39825","phone":"2292464847","ytdComp":0.0,"lat":30.92316,"lon":-84.60896,"accuracy":0.93,"accuracyType":"rooftop"},{"num":101372,"salesman":"Larry","active":"Active","name":"RAINEY USED CARS (BRONWOOD)","address":"100 MAIN ST E","city":"BRONWOOD","state":"GA","zip":"39826","phone":"2296959153","ytdComp":0.0,"lat":31.772481,"lon":-84.446805,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200839,"salesman":"House","active":"Inactive","name":"5 DOWN AUTO SALES","address":"100 MAIN ST.","city":"BRONWOOD","state":"GA","zip":"31794","phone":"2299953600","ytdComp":0.0,"lat":31.40108,"lon":-83.330141,"accuracy":0.94,"accuracyType":"range_interpolation"},{"num":200566,"salesman":"House","active":"Inactive","name":"AIMTRAC / BROOKLET","address":"9057 HWY 67","city":"BROOKLET","state":"GA","zip":"30415","phone":"912-839-2532","ytdComp":0.0,"lat":32.291877,"lon":-81.726818,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000040,"salesman":"Tiffany","active":"Active","name":"ADVANCED TIRE SERVICE","address":"1212 PONCE DE LEON BLVD","city":"BROOKSVILLE","state":"FL","zip":"34601","phone":"3526917771","ytdComp":582.87,"lat":28.569638,"lon":-82.39714,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101364,"salesman":"House","active":"Active","name":"KENNY'S AUTO AND TRUCK SALVAGE","address":"5898 DOUGLAS BROXTON HWY+","city":"BROXTON","state":"GA","zip":"31519","phone":"9123592222","ytdComp":167.0,"lat":31.592761,"lon":-82.854364,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100107,"salesman":"Larry","active":"Active","name":"JOHNSON AUTO & TIRE","address":"106 ALABAMA RD","city":"BROXTON","state":"GA","zip":"31519","phone":"9123592452","ytdComp":1649.87,"lat":31.62549,"lon":-82.885906,"accuracy":0.91,"accuracyType":"nearest_rooftop_match"},{"num":201028,"salesman":"House","active":"Active","name":"CLEMENT USED TIRES","address":"602 S. ALABAMA ST.","city":"BROXTON","state":"GA","zip":"31519","phone":"9126365505","ytdComp":10232.75,"lat":31.616582,"lon":-82.880077,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200511,"salesman":"House","active":"Inactive","name":"JEFF'S AUTOMOTIVE","address":"4765 NEW JESUP HWY","city":"BRUNSWICK","state":"GA","zip":"31520","phone":"912-342-7078","ytdComp":0.0,"lat":31.211947,"lon":-81.514132,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200516,"salesman":"House","active":"Inactive","name":"COOK'S AUTO REPAIR","address":"3704 COMMUNITY ROAD","city":"BRUNSWICK","state":"GA","zip":"31520","phone":"912-275-8247","ytdComp":0.0,"lat":31.198802,"lon":-81.491679,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200530,"salesman":"House","active":"Inactive","name":"KINGS COLONIAL FORD INC.","address":"3565 COMMUNITY RD","city":"BRUNSWICK","state":"GA","zip":"31520","phone":"912-264-6400","ytdComp":0.0,"lat":31.196156,"lon":-81.487897,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200536,"salesman":"House","active":"Inactive","name":"COMMUNITY TIRE","address":"3944 COMUNNITY RD","city":"BRUNSWICK","state":"GA","zip":"31523","phone":"912-223-0700","ytdComp":0.0,"lat":31.202276,"lon":-81.495559,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200538,"salesman":"House","active":"Inactive","name":"A+ TIRES (2244XW LTD)","address":"4504 ALTAMA AVE","city":"BRUNSWICK","state":"GA","zip":"31520","phone":"912-275-8730","ytdComp":0.0,"lat":31.193325,"lon":-81.484702,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200570,"salesman":"House","active":"Inactive","name":"AUTO SERVICE CENTER (BJS INC)","address":"116 CRISPEN BLVD","city":"BRUNSWICK","state":"GA","zip":"31525","phone":"912-267-9799","ytdComp":0.0,"lat":31.225602,"lon":-81.525063,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200741,"salesman":"House","active":"Inactive","name":"KAUFFMAN TIRE (BRUNSWICK)","address":"600 MALL BLVD.","city":"BRUNSWICK","state":"GA","zip":"31525","phone":"912-341-6677","ytdComp":0.0,"lat":31.214199,"lon":-81.490103,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101392,"salesman":"House","active":"Inactive","name":"JAY'S 24 HOUR TIRE REPAIR","address":"900 1ST AVE NE","city":"CAIRO","state":"GA","zip":"39828","phone":"229-397-0460","ytdComp":0.0,"lat":30.877256,"lon":-84.196087,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101543,"salesman":"House","active":"Inactive","name":"FLOWERS TIRE CO  LLC","address":"3294 HWY 84 E","city":"CAIRO","state":"GA","zip":"39828","phone":"229-377-6666","ytdComp":0.0,"lat":30.872634,"lon":-84.157409,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200146,"salesman":"House","active":"Active","name":"BRACEWELL AUTOMOTIVE SERVICE","address":"3434 US HWY 84 E","city":"CAIRO","state":"GA","zip":"39828","phone":"2293771771","ytdComp":2117.96,"lat":30.8687,"lon":-84.14993,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200148,"salesman":"House","active":"Active","name":"AUTO & TRUCK CARE SPECIAL","address":"238 ENOS LANE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293775342","ytdComp":471.1,"lat":30.777098,"lon":-84.219932,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200149,"salesman":"House","active":"Active","name":"CAIRO TIRE INC","address":"201 FIRST ST NE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293773622","ytdComp":0.0,"lat":30.879137,"lon":-84.206643,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200153,"salesman":"House","active":"Inactive","name":"CLEVELAND COMPLETE AUTO","address":"1951 1ST AVE SW","city":"CAIRO","state":"GA","zip":"39828","phone":"229-377-5632","ytdComp":0.0,"lat":30.876825,"lon":-84.232296,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200165,"salesman":"House","active":"Inactive","name":"FLOWERS AUTO CENTER","address":"3294 HWY 84 E","city":"CAIRO","state":"GA","zip":"31728","phone":"229-377-6666","ytdComp":0.0,"lat":30.872634,"lon":-84.157409,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200235,"salesman":"House","active":"Active","name":"SYRUPCITY TIRE & AUTO CENTER","address":"290 HWY 84 E","city":"CAIRO","state":"GA","zip":"39828","phone":"2293774200","ytdComp":0.0,"lat":30.884873,"lon":-84.204841,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200245,"salesman":"House","active":"Active","name":"TRAWICK SERVICE CENTER","address":"2060 HWY 111 N","city":"CAIRO","state":"GA","zip":"39828","phone":"2293774457","ytdComp":0.0,"lat":30.863297,"lon":-84.232577,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200287,"salesman":"House","active":"Active","name":"RIDLEY'S AUTOMOTIVE","address":"616 2ND AVE SE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293770170","ytdComp":0.0,"lat":30.875058,"lon":-84.199441,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200292,"salesman":"House","active":"Inactive","name":"STALLINGS MOTORS  CDJR","address":"1245 38TH BLVD NE","city":"CAIRO","state":"GA","zip":"39828","phone":"229-377-3333","ytdComp":0.0,"lat":30.878019,"lon":-84.194651,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200303,"salesman":"House","active":"Inactive","name":"AG PRO","address":"2025 HWY 84 E","city":"CAIRO","state":"GA","zip":"39828","phone":"229-377-3383","ytdComp":0.0,"lat":30.878235,"lon":-84.185404,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200339,"salesman":"House","active":"Inactive","name":"GRADY OIL CO","address":"325 5TH ST SE","city":"CAIRO","state":"GA","zip":"39828","phone":"229-377-4098","ytdComp":0.0,"lat":30.874921,"lon":-84.201576,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200345,"salesman":"House","active":"Inactive","name":"K&C PERFORMANCE AUTOMOTIVE","address":"135 2ND AVE","city":"CAIRO","state":"GA","zip":"39828","phone":"229-220-3225","ytdComp":0.0,"lat":30.725944,"lon":-84.345332,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200613,"salesman":"House","active":"Inactive","name":"HOPKINS MILLING CO.","address":"4392 GA HWY 111 SO.","city":"CAIRO","state":"GA","zip":"39828","phone":"229-872-3422","ytdComp":0.0,"lat":30.78228,"lon":-84.278974,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200614,"salesman":"House","active":"Inactive","name":"HOPKINS FARM","address":"272 OAK HILL","city":"CAIRO","state":"GA","zip":"39828","phone":"229-254-7230","ytdComp":0.0,"lat":30.773206,"lon":-84.282449,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200626,"salesman":"Car Dealer","active":"Active","name":"HOBSON CHEVROLET BUICK","address":"150 8TH AVE NE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293774162","ytdComp":0.0,"lat":30.885097,"lon":-84.206505,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200629,"salesman":"House","active":"Inactive","name":"AUTOMOTIVE SERVICE & PERF.","address":"674 3RD ST NE","city":"CAIRO","state":"GA","zip":"39828","phone":"229-421-9068","ytdComp":0.0,"lat":30.883826,"lon":-84.204162,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200653,"salesman":"House","active":"Active","name":"84 TIRE CENTER  LLC","address":"3294 HWY 84 EAST","city":"CAIRO","state":"GA","zip":"39828","phone":"2293778484","ytdComp":0.0,"lat":30.872634,"lon":-84.157409,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200658,"salesman":"House","active":"Active","name":"POWE AUTOMOTIVE","address":"139 2ND AVE SE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293970459","ytdComp":368.0,"lat":30.875186,"lon":-84.207349,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200859,"salesman":"Car Dealer","active":"Active","name":"STALLINGS MOTORS CHRYSLER","address":"1245 38TH BLVD NE","city":"CAIRO","state":"GA","zip":"39828","phone":"2293773333","ytdComp":0.0,"lat":30.878019,"lon":-84.194651,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200872,"salesman":"House","active":"Inactive","name":"QUICK PRO AUTO CARE","address":"805 1ST AVE N.E.","city":"CAIRO","state":"GA","zip":"39828","phone":"2293970375","ytdComp":0.0,"lat":30.877834,"lon":-84.197069,"accuracy":0.94,"accuracyType":"nearest_rooftop_match"},{"num":100926,"salesman":"House","active":"Active","name":"LIQUID TRANSPORT  INC.","address":"PO BOX 514","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293367103","ytdComp":0.0,"lat":31.249152,"lon":-84.242837,"accuracy":1.0,"accuracyType":"place"},{"num":200280,"salesman":"Larry","active":"Active","name":"JW PERFORMANCE & AUTO","address":"5745 RESEARCH RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"2298815927","ytdComp":0.0,"lat":31.179629,"lon":-84.228083,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200965,"salesman":"Larry","active":"Active","name":"MARQUEZ TIRE SHOP LLC","address":"557 NEWTON RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"2298868657","ytdComp":7099.95,"lat":31.237498,"lon":-84.226208,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101545,"salesman":"House","active":"Active","name":"TOMMY'S TIRE","address":"560 NEWTON RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293367696","ytdComp":0.0,"lat":31.237355,"lon":-84.226498,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101547,"salesman":"House","active":"Inactive","name":"CAMILLA TIRE & AUTO REPAIR  IN","address":"351 U.S. HWY 19 S","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-8473","ytdComp":0.0,"lat":31.222952,"lon":-84.196,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101548,"salesman":"House","active":"Inactive","name":"LARRY WOOD TRUCKING  INC","address":"4443 SKYLINE DR","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-5888","ytdComp":0.0,"lat":31.234748,"lon":-84.225779,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200150,"salesman":"House","active":"Active","name":"CAMILLA TIRE & AUTO CENTER","address":"351 US HWY 19 S","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293368473","ytdComp":0.0,"lat":31.222952,"lon":-84.196,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200159,"salesman":"House","active":"Inactive","name":"DAVID TIRE INC","address":"560 NEWTON RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-5852","ytdComp":0.0,"lat":31.237355,"lon":-84.226498,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200194,"salesman":"House","active":"Inactive","name":"MITCHELL COUNTY FORD","address":"231 HWY 19 N","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-9002","ytdComp":0.0,"lat":31.237102,"lon":-84.199312,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200244,"salesman":"House","active":"Inactive","name":"CLOSED ACCT","address":"0","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-7696","ytdComp":0.0,"lat":31.220318,"lon":-84.210211,"accuracy":0.9,"accuracyType":"street_center"},{"num":200282,"salesman":"House","active":"Inactive","name":"MITCHELL EMC","address":"475 CAIRO RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-903-3246","ytdComp":0.0,"lat":31.216025,"lon":-84.21625,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200316,"salesman":"House","active":"Inactive","name":"BROOKS SERVICE CENTER","address":"125 E BROAD ST","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-5555","ytdComp":0.0,"lat":31.230908,"lon":-84.204173,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200326,"salesman":"House","active":"Inactive","name":"C","address":"0","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-1227","ytdComp":0.0,"lat":31.220318,"lon":-84.210211,"accuracy":0.9,"accuracyType":"street_center"},{"num":200396,"salesman":"House","active":"Active","name":"PATE TIRE & SERVICE LLC","address":"100 A WEST OAKLAND AVE","city":"CAMILLA","state":"GA","zip":"31730","phone":"2298907389","ytdComp":3109.85,"lat":31.23161,"lon":-84.213258,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200487,"salesman":"House","active":"Active","name":"BROOKS SERVICE CENTER","address":"125 E BROAD ST","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293365555","ytdComp":0.0,"lat":31.230908,"lon":-84.204173,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200517,"salesman":"House","active":"Active","name":"CMW TIRE SERVICE  INC","address":"4368 GA. HWY 37","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293361227","ytdComp":0.0,"lat":31.236385,"lon":-84.16378,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200565,"salesman":"House","active":"Inactive","name":"AIMTRAC / CAMILLA","address":"1050 NEWTON RD","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-336-8780","ytdComp":0.0,"lat":31.245223,"lon":-84.241551,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200730,"salesman":"House","active":"Inactive","name":"RLK SERVICES LLC","address":"445 CHILI ST.","city":"CAMILLA","state":"GA","zip":"31730","phone":"229-344-1121","ytdComp":0.0,"lat":31.231748,"lon":-84.222387,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200773,"salesman":"Car Dealer","active":"Active","name":"FLINT RIVER FORD","address":"231 US 19 NORTH","city":"CAMILLA","state":"GA","zip":"31730","phone":"8552920747","ytdComp":0.0,"lat":31.237102,"lon":-84.199312,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200847,"salesman":"House","active":"Inactive","name":"FLINT RIVER FORD (AMI ACCT)","address":"231 US 19 NORTH","city":"CAMILLA","state":"GA","zip":"31730","phone":"8552920747","ytdComp":0.0,"lat":31.237102,"lon":-84.199312,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200874,"salesman":"House","active":"Inactive","name":"RAINEY USED CARS (CAMILLA)","address":"125 E. OAKLAND AVE.","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293360075","ytdComp":0.0,"lat":31.232442,"lon":-84.203516,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200879,"salesman":"House","active":"Inactive","name":"POWELL'S PAINT & BODY INC.","address":"10 EAST OAKLAND AVE","city":"CAMILLA","state":"GA","zip":"31730","phone":"2293368271","ytdComp":0.0,"lat":31.232867,"lon":-84.208019,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200909,"salesman":"House","active":"Inactive","name":"TOMAHAWK TIRE","address":"450 CHILI ST","city":"CAMILLA","state":"GA","zip":"31730","phone":"2294396594","ytdComp":0.0,"lat":31.231364,"lon":-84.223319,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101411,"salesman":"Larry","active":"Active","name":"LONE MOUNTAIN TRUCK LEASING","address":"200 OWEN PARKWAY CIRCLE","city":"CARTER LAKE","state":"IA","zip":"51510","phone":"8663400889","ytdComp":0.0,"lat":41.278223,"lon":-95.914365,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200840,"salesman":"House","active":"Active","name":"LAMEER ENTERPRISES  LLC","address":"1970 OLD COFFE RD.","city":"CECIL","state":"GA","zip":"31627","phone":"4704232136","ytdComp":0.0,"lat":31.04456,"lon":-83.395704,"accuracy":0.93,"accuracyType":"rooftop"},{"num":2000014,"salesman":"House","active":"Active","name":"UNITED TIRES ONLINE SALES -T","address":"3621 N HARLEM AVE","city":"CHICAGO","state":"IL","zip":"60634","phone":"3127147206","ytdComp":3258.71,"lat":41.945585,"lon":-87.807166,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200151,"salesman":"House","active":"Inactive","name":"AUTO GRAPHIX","address":"1417 JACKSON AVE","city":"CHIPLEY","state":"FL","zip":"32428","phone":"850-638-2999","ytdComp":0.0,"lat":30.777188,"lon":-85.550858,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200201,"salesman":"House","active":"Inactive","name":"MURRY TIRES INC","address":"623 N 6TH ST","city":"CHIPLEY","state":"FL","zip":"32428","phone":"850-638-1257","ytdComp":0.0,"lat":30.790102,"lon":-85.539261,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200206,"salesman":"House","active":"Inactive","name":"PANHANDLE 4X4","address":"1411 JACKSON AVE","city":"CHIPLEY","state":"FL","zip":"32428","phone":"850-638-9991","ytdComp":0.0,"lat":30.777354,"lon":-85.550411,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200219,"salesman":"House","active":"Inactive","name":"SEABORN'S AUTO REPAIR","address":"1141 MAIN ST","city":"CHIPLEY","state":"FL","zip":"32428","phone":"850-415-1776","ytdComp":0.0,"lat":30.768259,"lon":-85.537737,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200160,"salesman":"House","active":"Inactive","name":"DEEP SOUTH TIRE & BRAKE","address":"4061 THOMASVILLE RD","city":"CLIMAX","state":"GA","zip":"39834","phone":"229-243-0312","ytdComp":0.0,"lat":30.876279,"lon":-84.407635,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201002,"salesman":"House","active":"Active","name":"ZAPATA'S TIRE","address":"229 LAKE DOUGLAS EXT RD","city":"CLIMAX","state":"GA","zip":"39834","phone":"2293190228","ytdComp":5179.76,"lat":30.797066,"lon":-84.408375,"accuracy":0.95,"accuracyType":"rooftop"},{"num":200169,"salesman":"Larry","active":"Active","name":"HALL OIL & TIRE","address":"303 EAST CRAWFORD ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"2297583277","ytdComp":0.0,"lat":31.16871,"lon":-84.726754,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200175,"salesman":"House","active":"Inactive","name":"HWY 27 SERVICE CENTER","address":"111 W CRAWFORD ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"912-758-3281","ytdComp":0.0,"lat":31.169706,"lon":-84.73432,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200248,"salesman":"House","active":"Inactive","name":"CLOSED","address":"101 W COLLEGE ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"229-758-3929","ytdComp":0.0,"lat":31.170953,"lon":-84.733619,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200424,"salesman":"House","active":"Inactive","name":"CLOSED","address":"101 W COLLEGE ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"229-726-7870","ytdComp":0.0,"lat":31.170953,"lon":-84.733619,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200432,"salesman":"House","active":"Inactive","name":"CLOSED","address":"613 MLK DR","city":"COLQUITT","state":"GA","zip":"39837","phone":"229-220-4919","ytdComp":0.0,"lat":31.180811,"lon":-84.724212,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200472,"salesman":"House","active":"Inactive","name":"CLOSED","address":"201 W COLLEGE ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"229-758-6262","ytdComp":0.0,"lat":31.170992,"lon":-84.734952,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200774,"salesman":"Car Dealer","active":"Active","name":"FRANKLIN'S SPRING CREEK FORD","address":"220 WEST CRAWFORD ST","city":"COLQUITT","state":"GA","zip":"39837","phone":"2297583381","ytdComp":0.0,"lat":31.168923,"lon":-84.726228,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200848,"salesman":"House","active":"Inactive","name":"FRANKLIN'S SPRING CREEK FORD","address":"(AMI ACCOUNT)","city":"COLQUITT","state":"GA","zip":"39837","phone":"2297583381","ytdComp":0.0,"lat":31.17129,"lon":-84.73325,"accuracy":1.0,"accuracyType":"place"},{"num":200187,"salesman":"House","active":"Inactive","name":"K&C TIRE AUTOMOTIVE","address":"103 N MAIN ST","city":"COLUMBIA","state":"AL","zip":"36319","phone":"334-696-8473","ytdComp":0.0,"lat":31.292842,"lon":-85.111713,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200950,"salesman":"House","active":"Active","name":"SOUTHERN TIRE MART (MS)","address":"800 HWY 98","city":"COLUMBIA","state":"MS","zip":"39429","phone":"6014240161","ytdComp":0.0,"lat":31.239331,"lon":-89.818453,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201073,"salesman":"House","active":"Active","name":"GTO TIRE SERVICE & AUTO LLC","address":"2095 SOUTH PINE ST","city":"COOLIDGE","state":"GA","zip":"31738","phone":"2295290760","ytdComp":0.0,"lat":31.009175,"lon":-83.866918,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100084,"salesman":"House","active":"Inactive","name":"CITY OF CORDELE","address":"808 11TH AVE EAST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-276-2543","ytdComp":0.0,"lat":31.968496,"lon":-83.770772,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100371,"salesman":"House","active":"Inactive","name":"DISCOUNT TIRES & WHEELS","address":"1012 SOUTH 7TH STREET","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-9242","ytdComp":0.0,"lat":31.958181,"lon":-83.782763,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100621,"salesman":"House","active":"Inactive","name":"CORDELE TIRE CO. INC","address":"309 SOUTH SEVENTH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-5834","ytdComp":0.0,"lat":31.966219,"lon":-83.782188,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100666,"salesman":"House","active":"Inactive","name":"ABC WRECKER","address":"2044 HWY 300 SOUTH","city":"CORDELE","state":"GA","zip":"31015","phone":"229-535-3213","ytdComp":0.0,"lat":31.874868,"lon":-83.901289,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100741,"salesman":"House","active":"Active","name":"MASSEY'S MUFFLER","address":"109 DRAYTON LANE","city":"CORDELE","state":"GA","zip":"31015","phone":"2292734339","ytdComp":3168.18,"lat":31.992103,"lon":-83.798134,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100780,"salesman":"House","active":"Inactive","name":"A & D TIRES AND ACCESSORIES","address":"101 SOUTH PECAN ST.","city":"CORDELE","state":"GA","zip":"31015","phone":"2292710002","ytdComp":0.0,"lat":31.966664,"lon":-83.764141,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100837,"salesman":"House","active":"Inactive","name":"CLOSED","address":"12 DUVALL DR.","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-4420","ytdComp":0.0,"lat":31.952456,"lon":-83.787444,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100996,"salesman":"House","active":"Inactive","name":"T & C AUTOMOTIVE","address":"409 S. 7TH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-513-3240","ytdComp":0.0,"lat":31.964954,"lon":-83.782218,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101102,"salesman":"House","active":"Inactive","name":"ROCKHOUSE TRUCK AND TRAILER","address":"397 ROCKHOUSE RD","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-3359","ytdComp":0.0,"lat":31.913034,"lon":-83.744505,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101208,"salesman":"House","active":"Active","name":"MIKE FRASER AUTO REPAIR","address":"1204 S 7TH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"2292730652","ytdComp":5116.26,"lat":31.956705,"lon":-83.78288,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101490,"salesman":"House","active":"Inactive","name":"GENERAL REPAIR TRUCK & TRAILER","address":"401 S. HARRIS","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-6556","ytdComp":0.0,"lat":31.963886,"lon":-83.738411,"accuracy":0.88,"accuracyType":"rooftop"},{"num":100907,"salesman":"Larry","active":"Active","name":"SMITH'S DIESEL REPAIR","address":"157 FLOYD RD","city":"CORDELE","state":"GA","zip":"31015","phone":"2292731205","ytdComp":2751.62,"lat":31.901928,"lon":-83.734485,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101241,"salesman":"Larry","active":"Active","name":"NEW PETTIS TIRE","address":"410 8TH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"2292731153","ytdComp":5102.42,"lat":31.963609,"lon":-83.783988,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101568,"salesman":"Car Dealer","active":"Inactive","name":"SUNBELT FORD OF CORDELE  INC","address":"PO BOX 838","city":"CORDELE","state":"GA","zip":"31015","phone":"2292760607","ytdComp":0.0,"lat":31.917178,"lon":-83.785365,"accuracy":0.99,"accuracyType":"place"},{"num":101718,"salesman":"House","active":"Inactive","name":"DISCOUNT TIRES  ETC. LLC","address":"1012 S 7TH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-9242","ytdComp":0.0,"lat":31.958181,"lon":-83.782763,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101491,"salesman":"Larry","active":"Active","name":"QUALITY AUTO & R.V. SERVICE","address":"1002 E 11TH AVE","city":"CORDELE","state":"GA","zip":"31015","phone":"2292730720","ytdComp":0.0,"lat":31.968749,"lon":-83.795863,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200807,"salesman":"Larry","active":"Active","name":"CORDELE TIRE & WHEEL  LLC","address":"309 S. 7TH ST.","city":"CORDELE","state":"GA","zip":"31015","phone":"2294175099","ytdComp":20969.23,"lat":31.966219,"lon":-83.782188,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000043,"salesman":"Larry","active":"Active","name":"CENTRAL GA TIRE LLC","address":"260 BIG JIM RD","city":"CORDELE","state":"GA","zip":"31015","phone":"2299473023","ytdComp":0.0,"lat":31.877457,"lon":-83.781949,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200562,"salesman":"House","active":"Active","name":"BEST CARS OF CORDELE  LLC","address":"611 S. 7TH ST.","city":"CORDELE","state":"GA","zip":"31015","phone":"2292732378","ytdComp":224.42,"lat":31.962693,"lon":-83.781959,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200573,"salesman":"House","active":"Inactive","name":"PEI AUTOMOTIVE","address":"1218 HWY 280 W","city":"CORDELE","state":"GA","zip":"31015","phone":"229-271-9181","ytdComp":0.0,"lat":31.955723,"lon":-83.822979,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200729,"salesman":"House","active":"Inactive","name":"SHELL RAPID LUBE (CORDELE)","address":"306 16TH AVE E.","city":"CORDELE","state":"GA","zip":"31015","phone":"229-271-2776","ytdComp":0.0,"lat":31.963471,"lon":-83.779332,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200731,"salesman":"House","active":"Inactive","name":"KAUFFMAN TIRE (CORDELE)","address":"1409 16TH AVENUE EAST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-513-1847","ytdComp":0.0,"lat":31.961366,"lon":-83.760679,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200775,"salesman":"House","active":"Inactive","name":"SOUTHLAND CHRYSLER","address":"1096 HWY 280 WEST","city":"CORDELE","state":"GA","zip":"31015","phone":"229-273-3473","ytdComp":0.0,"lat":31.95564,"lon":-83.816454,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200823,"salesman":"House","active":"Inactive","name":"RAINEY'S USED CARS","address":"314 S 7TH ST","city":"CORDELE","state":"GA","zip":"31015","phone":"2292734600","ytdComp":0.0,"lat":31.966121,"lon":-83.782968,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200857,"salesman":"House","active":"Inactive","name":"SOUTHLAND CHRYSLER (AMI ACCT)","address":"1096 HWY 280 WEST","city":"CORDELE","state":"GA","zip":"31015","phone":"2292733473","ytdComp":0.0,"lat":31.95564,"lon":-83.816454,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200866,"salesman":"House","active":"Active","name":"LEMUS TIRE SHOP","address":"409 S. 7TH ST.","city":"CORDELE","state":"GA","zip":"31015","phone":"2294174737","ytdComp":6496.01,"lat":31.964954,"lon":-83.782218,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200882,"salesman":"Car Dealer","active":"Inactive","name":"SUNBELT FORD CORDELE (AMI)","address":"2511 E. 16TH AVE","city":"CORDELE","state":"GA","zip":"31015","phone":"2292760607","ytdComp":0.0,"lat":31.959041,"lon":-83.742584,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200890,"salesman":"House","active":"Active","name":"PMT TRK. TRAILER & TIRE REPAIR","address":"215 S MIDWAY RD","city":"CORDELE","state":"GA","zip":"31010","phone":"2294575167","ytdComp":4938.58,"lat":31.966994,"lon":-83.745157,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200912,"salesman":"House","active":"Active","name":"LANE'S TRK & TRL REPAIR & AUTO","address":"202 N. 7TH ST.","city":"CORDELE","state":"GA","zip":"31015","phone":"2293226338","ytdComp":0.0,"lat":31.969977,"lon":-83.782411,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200940,"salesman":"House","active":"Active","name":"PERRY BROS. OIL (CORDELE)","address":"302 NORTH MIDWAY RD","city":"CORDELE","state":"GA","zip":"31015","phone":"2292731412","ytdComp":455.42,"lat":31.971146,"lon":-83.745488,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201100,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 657","address":"1409 16TH AVENUE EAST","city":"CORDELE","state":"GA","zip":"31015","phone":"2293528490","ytdComp":0.0,"lat":31.961366,"lon":-83.760679,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000025,"salesman":"Car Dealer","active":"Active","name":"FORD CORDELE","address":"2511 E. 16TH AVE.","city":"CORDELE","state":"GA","zip":"31015","phone":"2292760607","ytdComp":4983.06,"lat":31.959041,"lon":-83.742584,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200141,"salesman":"House","active":"Inactive","name":"BARRON & BROTHERS","address":"470 S WAYSIDE ST","city":"CORNELIA","state":"GA","zip":"30531","phone":"706-778-2767","ytdComp":0.0,"lat":34.505197,"lon":-83.534315,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200492,"salesman":"House","active":"Inactive","name":"SALFORD BBI INC","address":"470 S WAYSIDE ST","city":"CORNELIA","state":"GA","zip":"30531","phone":"706-778-2767","ytdComp":0.0,"lat":34.505197,"lon":-83.534315,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200665,"salesman":"House","active":"Inactive","name":"SALFORD BBI  INC.","address":"470 S. WAYSIDE ST","city":"CORNELIA","state":"GA","zip":"30531","phone":"706-778-2767","ytdComp":0.0,"lat":34.505197,"lon":-83.534315,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200788,"salesman":"House","active":"Inactive","name":"BEST VALUE TIRES & ALIGNMENT","address":"2106 CRAWFORDVILLE HWY","city":"CRAWFORDVILLE","state":"FL","zip":"32327","phone":"8509262227","ytdComp":0.0,"lat":30.214724,"lon":-84.363838,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200506,"salesman":"House","active":"Active","name":"DEVANE TIRE & SERVICE LLC","address":"128 HOWELL MILL ROAD","city":"CUTHBERT","state":"GA","zip":"39840","phone":"2293109586","ytdComp":411.23,"lat":31.754989,"lon":-84.783804,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200531,"salesman":"House","active":"Inactive","name":"CUTHBERT TIRE INC.","address":"142 MAPLE ST","city":"CUTHBERT","state":"GA","zip":"39840","phone":"229-732-2600","ytdComp":0.0,"lat":31.770064,"lon":-84.788189,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200156,"salesman":"House","active":"Inactive","name":"B&R FIRESTONE","address":"609 HWY 134 E","city":"DALEVILLE","state":"AL","zip":"36322","phone":"334-598-4095","ytdComp":0.0,"lat":31.310421,"lon":-85.735625,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200204,"salesman":"House","active":"Inactive","name":"P & W GARAGE LLC","address":"204 OLD HWY 104 W","city":"DALEVILLE","state":"AL","zip":"36322","phone":"334-598-9149","ytdComp":0.0,"lat":31.301675,"lon":-85.772237,"accuracy":0.6,"accuracyType":"street_center"},{"num":200223,"salesman":"House","active":"Inactive","name":"SOUTHERN AUTOMOTIVE","address":"1856 JOE BRUER RD","city":"DALEVILLE","state":"AL","zip":"36322","phone":"334-347-0065","ytdComp":0.0,"lat":31.300925,"lon":-85.770693,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200154,"salesman":"House","active":"Active","name":"AUTO PRO & TIRE","address":"99 LAUREL AVE","city":"DAMASCUS","state":"GA","zip":"39841","phone":"2297256777","ytdComp":0.0,"lat":31.439828,"lon":-84.726518,"accuracy":0.93,"accuracyType":"rooftop"},{"num":101190,"salesman":"House","active":"Inactive","name":"DAWSON EASY PAY TIRE","address":"193 S MAIN ST","city":"DAWSON","state":"GA","zip":"39842","phone":"229-995-2167","ytdComp":0.0,"lat":31.773585,"lon":-84.446906,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101470,"salesman":"House","active":"Inactive","name":"TERRELL TRUCKS & TIRES","address":"1920 ALBANY HWY","city":"DAWSON","state":"GA","zip":"39842","phone":"229-995-2429","ytdComp":0.0,"lat":31.75568,"lon":-84.42819,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200946,"salesman":"Larry","active":"Active","name":"ABBI'S 24 HOUR","address":"880 INDUSTRIAL PARK BLVD","city":"DAWSON","state":"GA","zip":"39842","phone":"2294490762","ytdComp":10509.45,"lat":31.763031,"lon":-84.414445,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200602,"salesman":"House","active":"Inactive","name":"DAWSON TIRE AND AUTOMOTIVE CTR","address":"345 N. MAIN ST.","city":"DAWSON","state":"GA","zip":"39842","phone":"229-995-2429","ytdComp":0.0,"lat":31.776108,"lon":-84.44684,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200664,"salesman":"House","active":"Active","name":"ABR COMMERCIAL TRUCK & AUTO","address":"1126 FORRESTER DR SE","city":"DAWSON","state":"GA","zip":"39842","phone":"2299952169","ytdComp":1541.41,"lat":31.758895,"lon":-84.430907,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200719,"salesman":"House","active":"Active","name":"FOSTER EASY PAY TIRE CO.  INC.","address":"193 S. MAIN ST.","city":"DAWSON","state":"GA","zip":"39842","phone":"2299952167","ytdComp":20901.61,"lat":31.773585,"lon":-84.446906,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200749,"salesman":"House","active":"Inactive","name":"J & B TIRE REPAIR  INC.","address":"643 W. LEE ST.","city":"DAWSON","state":"GA","zip":"39842","phone":"2292887603","ytdComp":0.0,"lat":31.773892,"lon":-84.461226,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200776,"salesman":"House","active":"Inactive","name":"T CHILDRE BUICK","address":"1119 FORRESTER DR SE","city":"DAWSON","state":"GA","zip":"39842","phone":"229-995-2156","ytdComp":0.0,"lat":31.75899,"lon":-84.433091,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200800,"salesman":"House","active":"Inactive","name":"DAWSON AUTO SERVICE & REPAIR","address":"348 S. MAIN","city":"DAWSON","state":"GA","zip":"39842","phone":"2299955554","ytdComp":0.0,"lat":31.770916,"lon":-84.446552,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200131,"salesman":"House","active":"Inactive","name":"1 STOP AUTO SHOP","address":"1771 HWY 90 E","city":"DE FUNIAK SPRINGS","state":"FL","zip":"32433","phone":"850-951-1137","ytdComp":0.0,"lat":30.726552,"lon":-86.093482,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200809,"salesman":"House","active":"Active","name":"BOULEVARD TIRE CENTER","address":"816 S. WOODLAND BLVD","city":"DELAND","state":"FL","zip":"32720","phone":"3867346447","ytdComp":0.0,"lat":29.014709,"lon":-81.303663,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101386,"salesman":"House","active":"Inactive","name":"ATKINSON USED CARS","address":"2345 DOUGLAS HWY","city":"DENTON","state":"GA","zip":"31532","phone":"912-375-7128","ytdComp":0.0,"lat":31.725885,"lon":-82.694965,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100417,"salesman":"Larry","active":"Active","name":"CLARK BASS SERVICE","address":"2697 HWY 280","city":"DESOTO","state":"GA","zip":"31743","phone":"2298744685","ytdComp":22893.53,"lat":31.955582,"lon":-84.062858,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200969,"salesman":"House","active":"Active","name":"AG PRO- DIXIE","address":"4281 HWY 84 EAST","city":"DIXIE","state":"GA","zip":"31629","phone":"2292634133","ytdComp":0.0,"lat":30.792521,"lon":-83.667446,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200152,"salesman":"House","active":"Inactive","name":"CARTER'S AUTO SALES","address":"3988 GA HWY 33 S","city":"DOERUN","state":"GA","zip":"31744","phone":"229-776-9400","ytdComp":0.0,"lat":31.372644,"lon":-83.853964,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200193,"salesman":"House","active":"Active","name":"MCLEAN TIRES INC","address":"630 W BROAD AVE","city":"DOERUN","state":"GA","zip":"31744","phone":"2297827428","ytdComp":10373.82,"lat":31.320118,"lon":-83.924606,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200430,"salesman":"House","active":"Inactive","name":"CARVER BODY WORKS","address":"0","city":"DOERUN","state":"GA","zip":"31744","phone":"229-776-2421","ytdComp":0.0,"lat":31.33735,"lon":-83.91281,"accuracy":1.0,"accuracyType":"street_center"},{"num":200871,"salesman":"House","active":"Inactive","name":"COPELAND'S TIRE & DETAIL SHOP","address":"230 EAST BROAD AVE","city":"DOERUN","state":"GA","zip":"31744","phone":"2297827192","ytdComp":0.0,"lat":31.576758,"lon":-84.144891,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200173,"salesman":"House","active":"Active","name":"ESPY GRAY'S TIRE","address":"309 S DOWLING AVE","city":"DONALSONVILLE","state":"GA","zip":"31745","phone":"9125242700","ytdComp":0.0,"lat":31.039064,"lon":-84.87611,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":2000031,"salesman":"Larry","active":"Active","name":"W.C.'S TIRE TRAILER & TRUCK REPAIR LLC","address":"5206 FERTILIZER RD","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"2292020713","ytdComp":0.0,"lat":31.054189,"lon":-84.906123,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200246,"salesman":"Tiffany","active":"Active","name":"TRI COUNTY TIRE COMPANY","address":"805 E 5TH ST","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"2295242654","ytdComp":12460.76,"lat":31.036662,"lon":-84.869648,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200256,"salesman":"House","active":"Inactive","name":"AAA MOTORSPORTS","address":"604 W 3RD ST","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"2295248137","ytdComp":0.0,"lat":31.044264,"lon":-84.890037,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200306,"salesman":"House","active":"Inactive","name":"ALDAY COUNTRY STORE","address":"6834 HWY 374","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"229-861-2165","ytdComp":0.0,"lat":30.898038,"lon":-84.86442,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200314,"salesman":"House","active":"Inactive","name":"BILL CRAMER MOTORS","address":"731W 3RD ST","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"229-524-5180","ytdComp":0.0,"lat":31.047723,"lon":-84.894657,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200650,"salesman":"House","active":"Inactive","name":"UNIVERSAL IRRIGATION SUPPLY","address":"105 S. TENNILLE AVE","city":"DONALSONVILLE","state":"GA","zip":"39845","phone":"229-524-0115","ytdComp":0.0,"lat":31.041479,"lon":-84.878982,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101878,"salesman":"Tiffany","active":"Active","name":"TRI STATE COMMERCIAL TIRE LLC","address":"2324 ROSS CLARK CIRCLE","city":"DOTHAN","state":"AL","zip":"36301","phone":"3346788473","ytdComp":84.12,"lat":31.195083,"lon":-85.409376,"accuracy":1.0,"accuracyType":"rooftop"},{"num":102220,"salesman":"House","active":"Active","name":"JIM WHALEY TIRES (STORE 2)","address":"4143 W. MAIN ST.","city":"DOTHAN","state":"AL","zip":"36303","phone":"3347948633","ytdComp":0.0,"lat":31.233879,"lon":-85.454206,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200132,"salesman":"House","active":"Active","name":"WILKS A-ONE TIRE SALES","address":"4847 EAST HWY 84","city":"DOTHAN","state":"AL","zip":"36301","phone":"3347922225","ytdComp":1940.73,"lat":31.239222,"lon":-85.505367,"accuracy":0.69,"accuracyType":"street_center"},{"num":200181,"salesman":"Tiffany","active":"Active","name":"JIM WHALEY'S TIRES INC (W.H.)","address":"1940 S OATES ST SUITE 1","city":"DOTHAN","state":"AL","zip":"36301","phone":"3347943132","ytdComp":0.0,"lat":31.198645,"lon":-85.398321,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200202,"salesman":"House","active":"Inactive","name":"OLLIE HARRELL TIRE SERVICE","address":"2394 ROSS CLARK CIRCLE","city":"DOTHAN","state":"AL","zip":"36301","phone":"334-794-8604","ytdComp":0.0,"lat":31.197415,"lon":-85.411619,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200217,"salesman":"House","active":"Active","name":"SCOTT STEVENS TIRE & SERVICE","address":"2576 ROSS CLARK CIRCLE","city":"DOTHAN","state":"AL","zip":"36301","phone":"3347946969","ytdComp":0.0,"lat":31.204319,"lon":-85.419314,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200218,"salesman":"House","active":"Inactive","name":"SCOTTS GAS & TIRE","address":"3356 REEVES ST","city":"DOTHAN","state":"AL","zip":"36303","phone":"334-794-5574","ytdComp":0.0,"lat":31.268033,"lon":-85.378042,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200221,"salesman":"House","active":"Inactive","name":"SOLOMAN CHEVROLET","address":"4886 MONTGOMERY HWY","city":"DOTHAN","state":"AL","zip":"36303","phone":"334-793-3444","ytdComp":0.0,"lat":31.270884,"lon":-85.445771,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200237,"salesman":"House","active":"Inactive","name":"T & M TIRE AUTO SERVICE CENTER","address":"107 W WILSON ST","city":"DOTHAN","state":"AL","zip":"36303","phone":"334-673-1222","ytdComp":0.0,"lat":31.242398,"lon":-85.393069,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200241,"salesman":"House","active":"Active","name":"THE RIM SHOP INC","address":"766 S OATES ST","city":"DOTHAN","state":"AL","zip":"36301","phone":"3347939292","ytdComp":179.71,"lat":31.214666,"lon":-85.39345,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200247,"salesman":"House","active":"Inactive","name":"TRI STATE RUBBER CO","address":"1455 ROSS CLARK CIRCLE SE","city":"DOTHAN","state":"AL","zip":"36301","phone":"334-794-4739","ytdComp":0.0,"lat":31.199854,"lon":-85.368319,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200375,"salesman":"House","active":"Inactive","name":"ATLANTIC & SOUTHERN EQPT CO","address":"19223 KINSEY RD","city":"DOTHAN","state":"AL","zip":"36303","phone":"334-803-8141","ytdComp":0.0,"lat":31.270657,"lon":-85.347861,"accuracy":0.7,"accuracyType":"street_center"},{"num":200747,"salesman":"House","active":"Active","name":"JIM WHALEY'S TIRES (DOTHAN #4)","address":"1940 SOUTH OATES ST.","city":"DOTHAN","state":"AL","zip":"36301","phone":"3347943132","ytdComp":0.0,"lat":31.198645,"lon":-85.398321,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200926,"salesman":"House","active":"Active","name":"ACTION TRUCK CENTER","address":"211 ROSS CLARK CIRCLE","city":"DOTHAN","state":"AL","zip":"36303","phone":"3347948505","ytdComp":0.0,"lat":31.250805,"lon":-85.379555,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200927,"salesman":"House","active":"Active","name":"MASTER TIRE","address":"203 VULCAN WAY","city":"DOTHAN","state":"AL","zip":"36303","phone":"3347927518","ytdComp":322.39,"lat":31.243232,"lon":-85.373182,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200930,"salesman":"Car Dealer","active":"Active","name":"BONDY'S FORD INC.","address":"3615 ROSS CLARK CIR","city":"DOTHAN","state":"AL","zip":"36303","phone":"3347925171","ytdComp":0.0,"lat":31.252237,"lon":-85.42132,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200934,"salesman":"House","active":"Active","name":"JIM WHALEY TIRES (STORE 1)","address":"2237 MONTGOMERY HWY","city":"DOTHAN","state":"AL","zip":"36303","phone":"3347942030","ytdComp":0.0,"lat":31.240884,"lon":-85.41544,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200943,"salesman":"Car Dealer","active":"Active","name":"BONDY'S NISSAN INC.","address":"6393 W. MAIN ST.","city":"DOTHAN","state":"AL","zip":"36305","phone":"3347946936","ytdComp":0.0,"lat":31.239842,"lon":-85.492479,"accuracy":0.88,"accuracyType":"nearest_rooftop_match"},{"num":200951,"salesman":"Car Dealer","active":"Active","name":"DOTHAN KIA","address":"2985 ROSS CLARK CIR.","city":"DOTHAN","state":"AL","zip":"36301","phone":"3346738195","ytdComp":0.0,"lat":31.221687,"lon":-85.433101,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201201,"salesman":"House","active":"Active","name":"FIRESTONE STORE #005940","address":"127 JONES DR","city":"DOTHAN","state":"AL","zip":"36303","phone":"3346711082","ytdComp":0.0,"lat":31.250378,"lon":-85.423904,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100048,"salesman":"House","active":"Inactive","name":"DEEN TIRE","address":"2050 BROXTON HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-384-4652","ytdComp":0.0,"lat":31.530799,"lon":-82.85062,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101035,"salesman":"House","active":"Inactive","name":"K & R TIRES & MORE","address":"1472 IRON RD","city":"DOUGLAS","state":"GA","zip":"31535","phone":"912-393-7264","ytdComp":0.0,"lat":31.502168,"lon":-82.810127,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101111,"salesman":"House","active":"Inactive","name":"THUNDERBOLT ENTERPRISES","address":"1905 W. FOREST DR.","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-384-5977","ytdComp":0.0,"lat":31.508951,"lon":-82.822217,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101164,"salesman":"House","active":"Active","name":"DAVIS TIRE (DOUGLAS)","address":"1130 S E BOWENS MILL RD","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123837180","ytdComp":9927.92,"lat":31.48864,"lon":-82.836674,"accuracy":0.9,"accuracyType":"rooftop"},{"num":101256,"salesman":"Car Dealer","active":"Active","name":"ROBERT FENDER CHEVROLET","address":"3241 US 441","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9122929005","ytdComp":0.0,"lat":31.46309,"lon":-82.849992,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101296,"salesman":"House","active":"Inactive","name":"PHILLIP'S AUTO AND TIRES","address":"2440 HWY 221 N","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123816412","ytdComp":0.0,"lat":31.535768,"lon":-82.81166,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101365,"salesman":"House","active":"Inactive","name":"ALTON'S TIRE SERVICE","address":"105 W CHERRY ST","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-331-1044","ytdComp":0.0,"lat":31.5033,"lon":-82.851364,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101366,"salesman":"House","active":"Active","name":"B & M AUTOMOTIVE SERVICE","address":"1125 S.E. BOWENS MILL RD","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9123846115","ytdComp":464.85,"lat":31.527538,"lon":-82.877373,"accuracy":0.89,"accuracyType":"range_interpolation"},{"num":101367,"salesman":"House","active":"Inactive","name":"CROSBY OF COFFEE","address":"1135 S.W. BOWENS MILL RD","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-501-3111","ytdComp":0.0,"lat":31.489821,"lon":-82.869374,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101387,"salesman":"House","active":"Inactive","name":"SOUTH CENTRAL TRUCK SALES","address":"137 S.E. BOWENS MILL RD","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-383-9956","ytdComp":0.0,"lat":31.504231,"lon":-82.827975,"accuracy":0.9,"accuracyType":"rooftop"},{"num":101563,"salesman":"House","active":"Active","name":"BUSTED KNUCKLES AUTOMOTIVE CTR","address":"889 WALDRON RD.","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9123817962","ytdComp":0.0,"lat":31.409548,"lon":-82.817423,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101080,"salesman":"Larry","active":"Active","name":"AMERSON TIRE INC.","address":"609 BOWENS MILL RD SW","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123933674","ytdComp":37924.05,"lat":31.503625,"lon":-82.875643,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101161,"salesman":"Larry","active":"Active","name":"JMC TIRE CO.  INC.","address":"2680 GA-32 OCILLA HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123844940","ytdComp":95763.96,"lat":31.503644,"lon":-82.804317,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101295,"salesman":"Larry","active":"Active","name":"DAVID'S AUTO SALES / DOUGLAS","address":"1804 S PETERSON AVE.","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123848570","ytdComp":11146.81,"lat":31.478223,"lon":-82.850371,"accuracy":0.99,"accuracyType":"rooftop"},{"num":101539,"salesman":"Larry","active":"Active","name":"COURSON'S TIRE OF DOUGLAS","address":"601 E. BAKER HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123836188","ytdComp":37227.67,"lat":31.501665,"lon":-82.8434,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200510,"salesman":"Larry","active":"Active","name":"M & R TRUCK ACCESSORIES","address":"1714 PETERSON AVE","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9123842362","ytdComp":740.5,"lat":31.480922,"lon":-82.850565,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200907,"salesman":"Larry","active":"Active","name":"KNOLLWOOD TIRE & WHEEL","address":"1122 BOWENS MILL RD","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9122600005","ytdComp":2385.44,"lat":31.488173,"lon":-82.837459,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200959,"salesman":"Larry","active":"Active","name":"SOUTHERN GEORGIA TIRE LLC","address":"2639 US HWY 441 SOUTH","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9122920001","ytdComp":21271.43,"lat":31.465432,"lon":-82.850926,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000042,"salesman":"Larry","active":"Active","name":"ROJAS AUTO REPAIR","address":"1215 N. PETERSON AVE","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9124934004","ytdComp":2526.35,"lat":31.52274,"lon":-82.85093,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200540,"salesman":"House","active":"Inactive","name":"SKIP'S TIRE PLACE","address":"1472 IRON RD","city":"DOUGLAS","state":"GA","zip":"31533","phone":"912-393-7264","ytdComp":0.0,"lat":31.502168,"lon":-82.810127,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200648,"salesman":"House","active":"Active","name":"JOE'S AUTO REPAIR  LLC","address":"1301 N. PETERSON AVE.","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123842010","ytdComp":1708.17,"lat":31.523415,"lon":-82.850763,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200720,"salesman":"House","active":"Active","name":"LUBE MASTERS","address":"1120 BAKER HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123842248","ytdComp":831.92,"lat":31.500361,"lon":-82.864268,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200721,"salesman":"House","active":"Inactive","name":"SOUTH GEORGIA TIRE","address":"2639 S. HWY 441","city":"DOUGLAS","state":"GA","zip":"31535","phone":"912-292-0001","ytdComp":0.0,"lat":31.471153,"lon":-82.850422,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200734,"salesman":"Car Dealer","active":"Active","name":"ANDERSON FORD","address":"109 WESTGREEN HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123842600","ytdComp":5890.94,"lat":31.509409,"lon":-82.833057,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200777,"salesman":"Car Dealer","active":"Active","name":"ANDERSON BUICK GMC","address":"1001 BOWENS MILL RD SW","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123840811","ytdComp":4425.92,"lat":31.493633,"lon":-82.87209,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200778,"salesman":"Car Dealer","active":"Active","name":"WOODY FOLSOM CDJR (DOUGLAS)","address":"1280 BOWENS MILL RD SE","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9123842371","ytdComp":0.0,"lat":31.486877,"lon":-82.841938,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200844,"salesman":"House","active":"Inactive","name":"WOODY FOLSOM CDJR (AMI ACCT)","address":"1510 PETERSON AVE S.","city":"DOUGLAS","state":"GA","zip":"31535","phone":"9123842371","ytdComp":0.0,"lat":31.486065,"lon":-82.850721,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200865,"salesman":"Car Dealer","active":"Inactive","name":"ANDERSON FORD (AMI ACCOUNT)","address":"109 WESTGREEN HWY","city":"DOUGLAS","state":"GA","zip":"31533","phone":"9123842600","ytdComp":0.0,"lat":31.509409,"lon":-82.833057,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200166,"salesman":"House","active":"Active","name":"EDISON TIRE","address":"PO BOX 362","city":"EDISON","state":"GA","zip":"39846","phone":"2298352077","ytdComp":28680.1,"lat":31.506869,"lon":-84.721045,"accuracy":1.0,"accuracyType":"place"},{"num":200976,"salesman":"House","active":"Active","name":"TIRESEASY-LLC (TIFTON WHSE)","address":"101 CONTINENTAL BLVD","city":"EL SEGUNDO","state":"CA","zip":"90245","phone":"8443470789","ytdComp":517906.32,"lat":33.917015,"lon":-118.39262,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201040,"salesman":"House","active":"Active","name":"TIRES EASY (NAP - TIFTON)","address":"101 CONTINENTAL BLVD","city":"EL SEGUNDO","state":"CA","zip":"90245","phone":"8443470789","ytdComp":9475.91,"lat":33.917015,"lon":-118.39262,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101512,"salesman":"House","active":"Active","name":"ELLENTON TIRE AND AUTO","address":"128 BAKER ST","city":"ELLENTON","state":"GA","zip":"31747","phone":"2293242475","ytdComp":28323.53,"lat":31.176509,"lon":-83.587343,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200482,"salesman":"House","active":"Inactive","name":"520 SERVICE CENTER INC","address":"2452 HWY 82","city":"ENIGMA","state":"GA","zip":"31749","phone":"229-533-1002","ytdComp":0.0,"lat":31.41398,"lon":-83.332564,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200616,"salesman":"House","active":"Inactive","name":"SERVICEPRO TRUCK & TRAILER LLC","address":"2452 HWY 82","city":"ENIGMA","state":"GA","zip":"31749","phone":"2296463441","ytdComp":0.0,"lat":31.41398,"lon":-83.332564,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200161,"salesman":"House","active":"Inactive","name":"EASY PAY TIRE STORE","address":"510 E PARK AVE","city":"ENTERPRISE","state":"AL","zip":"36331","phone":"334-347-2291","ytdComp":0.0,"lat":31.308024,"lon":-85.844098,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200185,"salesman":"House","active":"Inactive","name":"L & M TIRE CO","address":"1312 E PARK AVE","city":"ENTERPRISE","state":"AL","zip":"36330","phone":"334-347-4818","ytdComp":0.0,"lat":31.307045,"lon":-85.830726,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101562,"salesman":"House","active":"Inactive","name":"BUSH MOTORS INC.","address":"2865 S. EUFAULA AVE","city":"EUFAULA","state":"AL","zip":"36027","phone":"334-687-2299","ytdComp":0.0,"lat":31.858485,"lon":-85.166236,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200507,"salesman":"House","active":"Inactive","name":"JAC'S TIRE & AUTO CARE INC.","address":"900 SOUTH EUFAULA AVE","city":"EUFAULA","state":"AL","zip":"36027","phone":"334-687-0101","ytdComp":0.0,"lat":31.884183,"lon":-85.145338,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200924,"salesman":"House","active":"Active","name":"EUFAULA TIRE","address":"3265 S. EUFAULA AVE.","city":"EUFAULA","state":"AL","zip":"36027","phone":"3342324433","ytdComp":0.0,"lat":31.835948,"lon":-85.165767,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200938,"salesman":"House","active":"Active","name":"JIM WHALEY TIRE (STORE 7)","address":"810 S. EUFAULA AVE","city":"EUFAULA","state":"AL","zip":"36027","phone":"3346873561","ytdComp":0.0,"lat":31.883583,"lon":-85.14545,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200227,"salesman":"House","active":"Inactive","name":"SOUTHERN STATES INC","address":"PO BOX 183","city":"FARMVILLE","state":"NC","zip":"27828","phone":"252-753-5371","ytdComp":0.0,"lat":35.59544,"lon":-77.58525,"accuracy":1.0,"accuracyType":"place"},{"num":100112,"salesman":"House","active":"Inactive","name":"PLANTERS TIRE & AUTO SERVICE","address":"130 HWY 129 SOUTH","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-424-0024","ytdComp":0.0,"lat":31.725063,"lon":-83.208652,"accuracy":0.9,"accuracyType":"range_interpolation"},{"num":100775,"salesman":"House","active":"Inactive","name":"0","address":"ACCT CLOSED  OUT OF BUSINESS","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-424-0524","ytdComp":0.0,"lat":31.748753,"lon":-83.235412,"accuracy":1.0,"accuracyType":"place"},{"num":101266,"salesman":"House","active":"Inactive","name":"NICK'S TIRE AND AUTO","address":"102 E CENTRAL AVE","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-424-9700","ytdComp":0.0,"lat":31.715017,"lon":-83.252957,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101300,"salesman":"Car Dealer","active":"Active","name":"FITZGERALD FORD AND LINCOLN","address":"161 BENJAMIN H HILL DR","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-423-8787","ytdComp":2114.63,"lat":31.693118,"lon":-83.250959,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101352,"salesman":"House","active":"Inactive","name":"BEN HILL COUNTY ROAD DEPT.","address":"147 COUNTY FARM RD.","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-426-5170","ytdComp":0.0,"lat":31.77438,"lon":-83.270273,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101551,"salesman":"House","active":"Inactive","name":"COOK'S PAINT AND BODY","address":"1403 HWY 280","city":"FITZGERALD","state":"GA","zip":"31079","phone":"229-365-7129","ytdComp":0.0,"lat":31.95556,"lon":-83.832728,"accuracy":0.92,"accuracyType":"range_interpolation"},{"num":101618,"salesman":"House","active":"Inactive","name":"DOSTER FARM SUPPLY","address":"8147 US HWY 280","city":"FITZGERALD","state":"GA","zip":"31079","phone":"229-365-2261","ytdComp":0.0,"lat":31.949703,"lon":-83.477021,"accuracy":0.82,"accuracyType":"nearest_rooftop_match"},{"num":100993,"salesman":"Larry","active":"Active","name":"D & G PERFORMANCE","address":"279 PERRY HOUSE RD","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294249566","ytdComp":0.0,"lat":31.691139,"lon":-83.261504,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101463,"salesman":"Larry","active":"Active","name":"SHELL RAPID LUBE (FITZGERALD)","address":"102 OCILLA HWY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294249348","ytdComp":38496.41,"lat":31.70571,"lon":-83.252548,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200441,"salesman":"House","active":"Inactive","name":"RICKS SERVICE CENTER","address":"815 N GRANT ST","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-423-7011","ytdComp":0.0,"lat":31.723582,"lon":-83.252826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101519,"salesman":"Larry","active":"Active","name":"MARK TAYLOR DBA/MTAA ENT.","address":"231 BOWENS MILL HWY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294254601","ytdComp":1832.51,"lat":31.732352,"lon":-83.245721,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200224,"salesman":"Larry","active":"Active","name":"SOUTHERN EXPRESS LUBE","address":"PO BOX 888","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2297770932","ytdComp":0.0,"lat":31.748753,"lon":-83.235412,"accuracy":1.0,"accuracyType":"place"},{"num":200621,"salesman":"Larry","active":"Active","name":"SOUTH GA LUBE (FITZGERALD)","address":"130 OCILLA HWY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2293451704","ytdComp":0.0,"lat":31.703389,"lon":-83.25219,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000023,"salesman":"Larry","active":"Active","name":"FABOS AUTO SALES LLC","address":"215 BOWENS MILL HWY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294230813","ytdComp":106.08,"lat":31.730912,"lon":-83.245576,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200597,"salesman":"House","active":"Inactive","name":"FITZGERALD TIRE & SERVICE","address":"205 N GRANT ST","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-622-0042","ytdComp":0.0,"lat":31.716452,"lon":-83.25258,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200689,"salesman":"House","active":"Inactive","name":"K&J HEAVY DUTY REPAIR","address":"582 JACKSONVILLE HWY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"229-423-5589","ytdComp":0.0,"lat":31.713521,"lon":-83.197549,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200705,"salesman":"House","active":"Inactive","name":"FITZGERALD CDJR","address":"187 BENJAMIN HILL DR.","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294235720","ytdComp":0.0,"lat":31.693168,"lon":-83.253283,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200826,"salesman":"House","active":"Active","name":"COVERED WAGON TRAILERS","address":"144 STUART WAY","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294234044","ytdComp":0.0,"lat":31.698197,"lon":-83.246519,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200833,"salesman":"Car Dealer","active":"Active","name":"FITZGERALD CHRYSLER DODGE RAM","address":"187 BENJAMIN HILL DR","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294235720","ytdComp":1600.0,"lat":31.693168,"lon":-83.253283,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200834,"salesman":"House","active":"Inactive","name":"FITZGERALD FORD (AMI ACCOUNT)","address":"161 BENJAMIN HILL DR","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294238787","ytdComp":0.0,"lat":31.693118,"lon":-83.250959,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200887,"salesman":"House","active":"Inactive","name":"SEVEN THREE ALLEY","address":"416 N GRANT ST","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2294572834","ytdComp":0.0,"lat":31.719235,"lon":-83.25329,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201065,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 2100","address":"131 BENJAMIN H HILL DR","city":"FITZGERALD","state":"GA","zip":"31750","phone":"2296354186","ytdComp":0.0,"lat":31.692571,"lon":-83.249335,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201041,"salesman":"House","active":"Active","name":"TIRETREADS LLC (TIFTON ACCT)","address":"258 HANOVER ROAD","city":"FLORHAM PARK","state":"NJ","zip":"7932","phone":"9738500121","ytdComp":373.4,"lat":40.809109,"lon":-74.465876,"accuracy":0.88,"accuracyType":"range_interpolation"},{"num":201057,"salesman":"House","active":"Active","name":"BOULEVARD TIRE CENTER #28","address":"4201 SOUTH STATE RD #7","city":"FORT LAUDERDALE","state":"FL","zip":"33314","phone":"9547927799","ytdComp":0.0,"lat":26.071652,"lon":-80.207621,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200694,"salesman":"House","active":"Inactive","name":"TIRE SUPER CENTER","address":"2800 FOWLER ST","city":"FORT MYERS","state":"FL","zip":"33901","phone":"239-288-6573","ytdComp":0.0,"lat":26.62592,"lon":-81.862415,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201074,"salesman":"House","active":"Active","name":"TIRE DEPOT CO. - TAG (TIFTON)","address":"8090 RANCHERS RD. NE","city":"FRIDLEY","state":"MN","zip":"55432","phone":"6122056109","ytdComp":14346.75,"lat":45.114704,"lon":-93.266632,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101495,"salesman":"House","active":"Inactive","name":"ST LUCIE BATTERY & TIRE","address":"5500 ORANGE AVE","city":"FT PIERCE","state":"FL","zip":"34947","phone":"772-672-8702","ytdComp":0.0,"lat":27.448272,"lon":-80.382128,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000021,"salesman":"Tiffany","active":"Active","name":"ADVANCED TIRE SERVICE","address":"912 N MAIN ST.","city":"GAINESVILLE","state":"FL","zip":"32601","phone":"3525590708","ytdComp":774.01,"lat":29.660434,"lon":-82.324773,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000037,"salesman":"House","active":"Active","name":"BOHANNON'S 1 STOP","address":"2215 N. MAIN ST.","city":"GAINESVILLE","state":"FL","zip":"32609","phone":"3523773170","ytdComp":0.0,"lat":29.673574,"lon":-82.321308,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200180,"salesman":"House","active":"Inactive","name":"HAYES TIRE & AUTO","address":"400 S COMMERCE ST","city":"GENEVA","state":"AL","zip":"36340","phone":"334-684-2765","ytdComp":0.0,"lat":31.034924,"lon":-85.864745,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200936,"salesman":"House","active":"Inactive","name":"INMAN STREET CUSTOMS","address":"845 US HWY 82","city":"GEORGETOWN","state":"GA","zip":"39854","phone":"2293766857","ytdComp":0.0,"lat":31.888351,"lon":-85.107471,"accuracy":0.85,"accuracyType":"nearest_rooftop_match"},{"num":201044,"salesman":"House","active":"Active","name":"JACK'S AUTO REPAIR","address":"845 US HWY 82","city":"GEORGETOWN","state":"GA","zip":"39854","phone":"2292345011","ytdComp":0.0,"lat":31.888351,"lon":-85.107471,"accuracy":0.85,"accuracyType":"nearest_rooftop_match"},{"num":200618,"salesman":"House","active":"Inactive","name":"ATLANTIC & SOUTHERN EQPT CO","address":"1504 HWY 117 SOUTH","city":"GOLDSBORO","state":"NC","zip":"27530","phone":"919-734-0781","ytdComp":0.0,"lat":35.35977,"lon":-78.010692,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200684,"salesman":"Tiffany","active":"Active","name":"OTR SERVICES  INC.","address":"11507 SW MOUNT GILEAD RD.","city":"GREENVILLE","state":"FL","zip":"32331","phone":"8503711129","ytdComp":0.0,"lat":30.303395,"lon":-83.760454,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200799,"salesman":"House","active":"Inactive","name":"W ROWELL FARMS","address":"11878 US 221","city":"GREENVILLE","state":"FL","zip":"32331","phone":"8505842310","ytdComp":0.0,"lat":30.287369,"lon":-83.632252,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":201063,"salesman":"House","active":"Active","name":"SOUTHERN TIRE EXPORTERS","address":"2820 STRAITS BLVD","city":"GULF SHORES","state":"AL","zip":"36547","phone":"4048195113","ytdComp":0.0,"lat":30.281835,"lon":-87.71141,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200270,"salesman":"Tiffany","active":"Active","name":"HAHIRA AUTOMOTIVE SERVICE","address":"407 W MAIN ST","city":"HAHIRA","state":"GA","zip":"31632","phone":"2297942429","ytdComp":8240.41,"lat":30.989953,"lon":-83.376883,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200448,"salesman":"House","active":"Inactive","name":"HAMILTON AUTOMOTIVE","address":"107 W COLEMAN","city":"HAHIRA","state":"GA","zip":"31632","phone":"229-794-1235","ytdComp":0.0,"lat":30.984139,"lon":-83.370506,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201064,"salesman":"House","active":"Active","name":"CHAD'S AUTO REPAIR","address":"107 W. COLEMAN DR","city":"HAHIRA","state":"GA","zip":"31632","phone":"2297941109","ytdComp":0.0,"lat":30.984139,"lon":-83.370506,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200177,"salesman":"House","active":"Inactive","name":"HARTFORD TIRE & AUTO CENTER","address":"401 W MAIN ST","city":"HARTFORD","state":"AL","zip":"36344","phone":"334-588-2122","ytdComp":0.0,"lat":31.099671,"lon":-85.701311,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200983,"salesman":"House","active":"Inactive","name":"ERRICK'S TIRE CENTER","address":"30 BOYD ST","city":"HAVANA","state":"FL","zip":"32333","phone":"8506161095","ytdComp":0.0,"lat":30.648112,"lon":-84.416479,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200986,"salesman":"House","active":"Active","name":"IRONHORSE EQUIPMENT SERVICES","address":"7092 FL GA HWY","city":"HAVANA","state":"FL","zip":"32333","phone":"8505391160","ytdComp":0.0,"lat":30.64208,"lon":-84.416276,"accuracy":0.9,"accuracyType":"rooftop"},{"num":200977,"salesman":"House","active":"Active","name":"TRI STATE COMMERCIAL TIRE LLC","address":"17794 AL-1 SOUTH","city":"HEADLAND","state":"AL","zip":"36345","phone":"3347855390","ytdComp":0.0,"lat":31.234295,"lon":-85.395863,"accuracy":0.64,"accuracyType":"street_center"},{"num":200683,"salesman":"Larry","active":"Active","name":"HICKOX AUTO DEALERS","address":"3114 HIGHWAY 82","city":"HOBOKEN","state":"GA","zip":"31542","phone":"9122813922","ytdComp":2639.86,"lat":31.178916,"lon":-82.149669,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000003,"salesman":"House","active":"Active","name":"RACETREADS LLC (TIFTON)","address":"80 RIVER STREET","city":"HOBOKEN","state":"NJ","zip":"7030","phone":"","ytdComp":0.0,"lat":40.736229,"lon":-74.029274,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200637,"salesman":"House","active":"Inactive","name":"CLOSED","address":"2309 DAMASCUS RD","city":"HOMER","state":"GA","zip":"30547","phone":"678-677-6770","ytdComp":0.0,"lat":34.421882,"lon":-83.441945,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200391,"salesman":"House","active":"Active","name":"CLINCH BRAKE & ALIGNMENT","address":"173 C CHURCH ST","city":"HOMERVILLE","state":"GA","zip":"31634","phone":"9124873746","ytdComp":875.36,"lat":31.038601,"lon":-82.748368,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200392,"salesman":"House","active":"Inactive","name":"ED'S SERVICE STATION","address":"173 S CHURCH ST","city":"HOMERVILLE","state":"GA","zip":"31634","phone":"912-487-2500","ytdComp":0.0,"lat":31.034612,"lon":-82.745677,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200470,"salesman":"House","active":"Inactive","name":"KIGHT'S HOME & AUTO  INC","address":"470 N CHURCH ST","city":"HOMERVILLE","state":"GA","zip":"31634","phone":"912-487-0084","ytdComp":0.0,"lat":31.042044,"lon":-82.749686,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200806,"salesman":"House","active":"Active","name":"WALKERS AUTO & OUTDOOR  INC","address":"470 N. CHURCH ST.","city":"HOMERVILLE","state":"GA","zip":"31634","phone":"9124870084","ytdComp":1103.35,"lat":31.042044,"lon":-82.749686,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200502,"salesman":"Larry","active":"Active","name":"ATKINSON TIRE","address":"21277 HWY 82(ATKINSON GA)","city":"HORTENSE","state":"GA","zip":"31543","phone":"9127785680","ytdComp":3030.75,"lat":31.26837,"lon":-81.810366,"accuracy":0.5,"accuracyType":"place"},{"num":200135,"salesman":"House","active":"Inactive","name":"AG-CHEM EQUIPMENT CO. INC.","address":"54927 796TH STREET","city":"JACKSON","state":"MN","zip":"56143","phone":"507-847-2690","ytdComp":0.0,"lat":43.635486,"lon":-94.871398,"accuracy":0.48,"accuracyType":"street_center"},{"num":101449,"salesman":"House","active":"Active","name":"MAGIC WHEEL & TIRE - JACKSONVILLE","address":"10448 ATLANTIC BLVD","city":"JACKSONVILLE","state":"FL","zip":"32225","phone":"904-998-9801","ytdComp":0.0,"lat":30.323583,"lon":-81.528293,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101997,"salesman":"House","active":"Active","name":"GOODYEAR COMMERCIAL TIRE & SER","address":"450 LANE AVE N.","city":"JACKSONVILLE","state":"FL","zip":"32254","phone":"9047832541","ytdComp":0.0,"lat":30.329127,"lon":-81.752801,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200678,"salesman":"House","active":"Inactive","name":"GERDAU","address":"16770 REBAR RD","city":"JACKSONVILLE","state":"FL","zip":"32234","phone":"904-266-1531","ytdComp":0.0,"lat":30.283895,"lon":-81.977552,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200745,"salesman":"House","active":"Inactive","name":"TIRE OUTLET DIRECT","address":"479 ZOO PKWY","city":"JACKSONVILLE","state":"FL","zip":"32226","phone":"904-696-0580","ytdComp":0.0,"lat":30.407795,"lon":-81.639826,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200813,"salesman":"House","active":"Inactive","name":"SUN TIRE OF 103RD","address":"7390 103RD ST.","city":"JACKSONVILLE","state":"FL","zip":"32210","phone":"","ytdComp":0.0,"lat":30.247827,"lon":-81.767851,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200814,"salesman":"House","active":"Inactive","name":"SUN TIRE OF SOUTHSIDE","address":"5942 UNIVERSITY BLVD. W.","city":"JACKSONVILLE","state":"FL","zip":"32216","phone":"","ytdComp":0.0,"lat":30.276368,"lon":-81.603251,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200815,"salesman":"House","active":"Inactive","name":"SUN TIRE OF REGENCY","address":"10101 ATLANTIC BLVD","city":"JACKSONVILLE","state":"FL","zip":"32225","phone":"","ytdComp":0.0,"lat":30.323554,"lon":-81.537937,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200816,"salesman":"House","active":"Inactive","name":"TINSELTOWN MAVIS","address":"9718 TOUCHTON RD","city":"JACKSONVILLE","state":"FL","zip":"32246","phone":"","ytdComp":0.0,"lat":30.268448,"lon":-81.556379,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200817,"salesman":"House","active":"Inactive","name":"SUN TIRE OF SAN PABLO","address":"14175 BEACJ BLVD","city":"JACKSONVILLE","state":"FL","zip":"32250","phone":"","ytdComp":0.0,"lat":30.288521,"lon":-81.443711,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200818,"salesman":"House","active":"Inactive","name":"SUN TIRE OF MANDARIN","address":"9950 SAN JOSE BLVD","city":"JACKSONVILLE","state":"FL","zip":"32257","phone":"","ytdComp":0.0,"lat":30.197227,"lon":-81.622261,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200819,"salesman":"House","active":"Inactive","name":"BARTRAM MAVIS","address":"13958 VILLAGE LAKE CIRCLE","city":"JACKSONVILLE","state":"FL","zip":"32258","phone":"","ytdComp":0.0,"lat":30.138724,"lon":-81.545612,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200820,"salesman":"House","active":"Inactive","name":"SUN TIRE OF RACE TRACK RD","address":"2660 RACE TRACK RD","city":"JACKSONVILLE","state":"FL","zip":"32259","phone":"","ytdComp":0.0,"lat":30.114225,"lon":-81.593212,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200889,"salesman":"House","active":"Active","name":"CONLAN TIRE CO.","address":"5233 COMMONWEALTH AVE","city":"JACKSONVILLE","state":"FL","zip":"32254","phone":"9047197460","ytdComp":0.0,"lat":30.337203,"lon":-81.734651,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201009,"salesman":"House","active":"Active","name":"BOULEVARD TIRE CENTER","address":"7500 COMMONWEALTH AVE","city":"JACKSONVILLE","state":"FL","zip":"32220","phone":"9047868550","ytdComp":0.0,"lat":30.34104,"lon":-81.775996,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201019,"salesman":"House","active":"Active","name":"MILLER BROS GIANT TIRE SVC","address":"11608 COLUMBIA PARK DR. WEST","city":"JACKSONVILLE","state":"FL","zip":"32258","phone":"9042688799","ytdComp":0.0,"lat":30.157399,"lon":-81.539564,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201069,"salesman":"House","active":"Active","name":"HEAVY DUTY TIRE","address":"694 EDGEWOOD AVE N","city":"JACKSONVILLE","state":"FL","zip":"32254","phone":"9048337750","ytdComp":0.0,"lat":30.329444,"lon":-81.732628,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201070,"salesman":"House","active":"Active","name":"SNIDER INDUSTRIAL","address":"5700 COMMONWEALTH AVE.","city":"JACKSONVILLE","state":"FL","zip":"32254","phone":"9043831143","ytdComp":3184.84,"lat":30.334508,"lon":-81.747045,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201202,"salesman":"House","active":"Active","name":"FIRESTONE STORE #018538","address":"4712 BLANDING BLVD","city":"JACKSONVILLE","state":"FL","zip":"32210","phone":"9047711381","ytdComp":0.0,"lat":30.25644,"lon":-81.742893,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201203,"salesman":"House","active":"Active","name":"FIRESTONE STORE #028444","address":"10553 SAN JOSE BLVD","city":"JACKSONVILLE","state":"FL","zip":"32257","phone":"9042682921","ytdComp":0.0,"lat":30.186177,"lon":-81.627505,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201204,"salesman":"House","active":"Active","name":"FIRESTONE STORE #014230","address":"9501 ARLINGTON EXPY","city":"JACKSONVILLE","state":"FL","zip":"32225","phone":"9047213049","ytdComp":0.0,"lat":30.323851,"lon":-81.549833,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201205,"salesman":"House","active":"Active","name":"TIRES PLUS #122262","address":"9350 SAN JOSE BLVD","city":"JACKSONVILLE","state":"FL","zip":"32257","phone":"9047317076","ytdComp":0.0,"lat":30.207176,"lon":-81.617674,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201206,"salesman":"House","active":"Active","name":"TIRES PLUS #629944","address":"4262 SOUTHSIDE BLVD","city":"JACKSONVILLE","state":"FL","zip":"32216","phone":"9046450733","ytdComp":0.0,"lat":30.270665,"lon":-81.559805,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201207,"salesman":"House","active":"Active","name":"TIRES PLUS #122068","address":"5880 UNIVERSITY BLVD W","city":"JACKSONVILLE","state":"FL","zip":"32216","phone":"9047332555","ytdComp":0.0,"lat":30.275715,"lon":-81.604641,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201208,"salesman":"House","active":"Active","name":"TIRES PLUS #676209","address":"9559 BAYMEADOWS RD","city":"JACKSONVILLE","state":"FL","zip":"32256","phone":"9047370401","ytdComp":0.0,"lat":30.221255,"lon":-81.560185,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201209,"salesman":"House","active":"Active","name":"TIRES PLUS #121738","address":"10361 ATLANTIC BLVD","city":"JACKSONVILLE","state":"FL","zip":"32225","phone":"9049980322","ytdComp":0.0,"lat":30.324039,"lon":-81.531332,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201210,"salesman":"House","active":"Active","name":"TIRES PLUS #122114","address":"6381 103RD ST","city":"JACKSONVILLE","state":"FL","zip":"32210","phone":"9047729494","ytdComp":0.0,"lat":30.248725,"lon":-81.746932,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201211,"salesman":"House","active":"Active","name":"TIRES PLUS #122084","address":"5403 NORMANDY BLVD","city":"JACKSONVILLE","state":"FL","zip":"32205","phone":"9047868560","ytdComp":0.0,"lat":30.30926,"lon":-81.739757,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201212,"salesman":"House","active":"Active","name":"TIRES PLUS #239627","address":"1332 DUNN AVE","city":"JACKSONVILLE","state":"FL","zip":"32218","phone":"9047143737","ytdComp":0.0,"lat":30.429445,"lon":-81.668974,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201213,"salesman":"House","active":"Active","name":"TIRES PLUS #354080","address":"13162 ATLANTIC BLVD","city":"JACKSONVILLE","state":"FL","zip":"32225","phone":"9042202141","ytdComp":0.0,"lat":30.319334,"lon":-81.46322,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201214,"salesman":"House","active":"Active","name":"TIRES PLUS #772996","address":"12575 BARTRAM PARK BLVD","city":"JACKSONVILLE","state":"FL","zip":"32258","phone":"9042929191","ytdComp":0.0,"lat":30.140076,"lon":-81.54673,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201215,"salesman":"House","active":"Active","name":"FIRESTONE STORE #786365","address":"4628 TROPEA WAY","city":"JACKSONVILLE","state":"FL","zip":"32246","phone":"9046429148","ytdComp":0.0,"lat":30.259807,"lon":-81.530667,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000006,"salesman":"House","active":"Active","name":"GOODYEAR COMMERCIL TIRE & SVC","address":"450 LANE AVE","city":"JACKSONVILLE","state":"FL","zip":"32254","phone":"9047832541","ytdComp":0.0,"lat":30.329127,"lon":-81.752801,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000045,"salesman":"Tiffany","active":"Active","name":"MANDARIN TIRES INCOPORATED","address":"3272 CORMORANT DR","city":"JACKSONVILLE","state":"FL","zip":"32223","phone":"9043189277","ytdComp":0.0,"lat":30.144394,"lon":-81.626062,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200660,"salesman":"House","active":"Active","name":"RO'S TIRES LLC","address":"4661 US HWY 41 SOUTH","city":"JASPER","state":"FL","zip":"32052","phone":"3867921059","ytdComp":0.0,"lat":30.502881,"lon":-82.941798,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200876,"salesman":"Tiffany","active":"Active","name":"SUWANNEE VALLEY TIRE","address":"1225 NW US HWY 129","city":"JASPER","state":"FL","zip":"32052","phone":"3867922420","ytdComp":0.0,"lat":30.524905,"lon":-82.961349,"accuracy":0.85,"accuracyType":"rooftop"},{"num":201046,"salesman":"Tiffany","active":"Active","name":"BIELLING'S TIRE INC.","address":"5559 W SR 238","city":"LAKE BUTLER","state":"FL","zip":"32054","phone":"3867526568","ytdComp":0.0,"lat":30.007626,"lon":-82.341422,"accuracy":0.5,"accuracyType":"place"},{"num":200408,"salesman":"House","active":"Inactive","name":"ATLANTIC & SOUTHERN EQUIPMENT","address":"1642 FOREST PKWY","city":"LAKE CITY","state":"GA","zip":"30260","phone":"404-487-0350","ytdComp":0.0,"lat":33.611184,"lon":-84.336988,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200877,"salesman":"Tiffany","active":"Inactive","name":"LEWIS MURRAY TIRE & WRECKER","address":"3827 N. US HWY 441","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3867525688","ytdComp":0.0,"lat":30.243573,"lon":-82.637868,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200878,"salesman":"Tiffany","active":"Inactive","name":"THOMAS TIRE REPAIR & ROAD SVC","address":"407 NW COUNTY RD 25A","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3863198148","ytdComp":2335.8,"lat":30.212913,"lon":-82.640809,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200894,"salesman":"House","active":"Inactive","name":"MURRAY'S TIRE SALES","address":"1100 N. MARION AVE","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3867524215","ytdComp":0.0,"lat":30.203874,"lon":-82.637486,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200895,"salesman":"House","active":"Active","name":"TIRE MART OF LAKE CITY","address":"2700 W. US HWY 90","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3867520054","ytdComp":2057.28,"lat":30.179276,"lon":-82.676542,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200899,"salesman":"House","active":"Inactive","name":"WISHES AUTO SHOP REPAIR","address":"1357 SE BAYA DR","city":"LAKE CITY","state":"FL","zip":"32025","phone":"3867540040","ytdComp":0.0,"lat":30.182711,"lon":-82.616223,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200916,"salesman":"House","active":"Active","name":"GATEWAY DIESEL  AUTO & MOBILE","address":"1020 NW MAIN BLVD","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3525580768","ytdComp":0.0,"lat":30.183246,"lon":-82.639186,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200923,"salesman":"Tiffany","active":"Active","name":"ADVANCED TIRE SERVICE","address":"837 EAST DUVAL ST","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3864066745","ytdComp":1549.01,"lat":30.189417,"lon":-82.624848,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201032,"salesman":"House","active":"Active","name":"A-1 TIRE PLUS","address":"1944 E. DUVAL ST","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3868675495","ytdComp":12980.17,"lat":30.186833,"lon":-82.606654,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201047,"salesman":"House","active":"Active","name":"BIELLING'S TIRE INC. NO. 2","address":"166 W DUVAL ST","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3867522120","ytdComp":0.0,"lat":30.189093,"lon":-82.638254,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201048,"salesman":"House","active":"Active","name":"RRO 24 HR ROADSIDE ASSISTANCE","address":"1620 NW OAKLAND AVE","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3869650117","ytdComp":3118.41,"lat":30.213223,"lon":-82.641446,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201058,"salesman":"House","active":"Inactive","name":"AFTER 5 COMM. TIRE & OFF ROAD","address":"1721 E DUVAL ST","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3524513863","ytdComp":1742.9,"lat":30.188292,"lon":-82.610204,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201066,"salesman":"House","active":"Active","name":"TIRE KINGDOM SVC CENTER - 4071","address":"2829 W US HWY 90","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3862597861","ytdComp":0.0,"lat":30.180212,"lon":-82.67864,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000018,"salesman":"Tiffany","active":"Active","name":"ALL PRO DIESEL  LLC","address":"378 NW WALDO ST","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3864389912","ytdComp":3546.37,"lat":30.210044,"lon":-82.643755,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000044,"salesman":"Tiffany","active":"Active","name":"MURRAY'S TIRE & ROAD SERVICE","address":"3827 N US HWY 441","city":"LAKE CITY","state":"FL","zip":"32055","phone":"3867525688","ytdComp":109.98,"lat":30.243573,"lon":-82.637868,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200347,"salesman":"House","active":"Inactive","name":"LAKE PARK TIRE & GARAGE","address":"6872 LAKE PARK BELLVILLE RD","city":"LAKE PARK","state":"GA","zip":"31636","phone":"229-559-1650","ytdComp":0.0,"lat":30.645686,"lon":-83.187378,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200630,"salesman":"House","active":"Active","name":"EXPLICIT OFFROAD","address":"1509 WEST MARION AVE.","city":"LAKE PARK","state":"GA","zip":"31636","phone":"2292515168","ytdComp":0.0,"lat":30.68616,"lon":-83.194858,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200875,"salesman":"House","active":"Inactive","name":"LAKE PARK TRUCK CENTER","address":"7061 LAKE PARK BELLVILLE RD","city":"LAKE PARK","state":"GA","zip":"31636","phone":"2295591913","ytdComp":0.0,"lat":30.640143,"lon":-83.192691,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201027,"salesman":"House","active":"Active","name":"ARTHUR BITZER TIRES","address":"239 LAKES BLVD","city":"LAKE PARK","state":"GA","zip":"31636","phone":"2294155720","ytdComp":0.0,"lat":30.686296,"lon":-83.201193,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201075,"salesman":"House","active":"Inactive","name":"SOUTHERN TIRE MART  LLC (NA)","address":"7061 BELLVILLE RD","city":"LAKE PARK","state":"GA","zip":"31636","phone":"2295591913","ytdComp":0.0,"lat":30.640252,"lon":-83.192339,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000004,"salesman":"House","active":"Active","name":"ODELL AUTOMOTIVE","address":"321 LAKES BLVD","city":"LAKE PARK","state":"GA","zip":"31636","phone":"2295068063","ytdComp":187.0,"lat":30.685693,"lon":-83.202837,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100786,"salesman":"House","active":"Inactive","name":"LAKELAND TIRE CO INC.","address":"937 W. THIGPEN AVE","city":"LAKELAND","state":"GA","zip":"31635","phone":"229-482-1000","ytdComp":0.0,"lat":31.042053,"lon":-83.08961,"accuracy":0.85,"accuracyType":"range_interpolation"},{"num":200315,"salesman":"House","active":"Active","name":"BOBBY'S CITGO","address":"46 N CARTER ST","city":"LAKELAND","state":"GA","zip":"31635","phone":"2294822724","ytdComp":5034.54,"lat":31.041327,"lon":-83.074953,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200827,"salesman":"Tiffany","active":"Active","name":"LAKELAND TIRE DBA COOK & SONS","address":"66 W. MAIN ST","city":"LAKELAND","state":"GA","zip":"31635","phone":"2294821000","ytdComp":62549.5,"lat":31.040747,"lon":-83.077522,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200591,"salesman":"House","active":"Inactive","name":"KEATON TIRE","address":"54 E. CAPPS ST","city":"LAMONT","state":"FL","zip":"32336","phone":"8505440694","ytdComp":0.0,"lat":30.4107,"lon":-83.91048,"accuracy":0.85,"accuracyType":"rooftop"},{"num":100635,"salesman":"House","active":"Inactive","name":"LEESBURG TIRE SHOP","address":"105 WALNUT AVE S","city":"LEESBURG","state":"GA","zip":"31763","phone":"229-759-8551","ytdComp":0.0,"lat":31.731691,"lon":-84.171893,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101524,"salesman":"House","active":"Active","name":"MASTER BODY WORKS","address":"1518 US HWY 19 S","city":"LEESBURG","state":"GA","zip":"31763","phone":"2294398833","ytdComp":933.97,"lat":31.634644,"lon":-84.177279,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101525,"salesman":"House","active":"Active","name":"BMS DISCOUNT TIRES","address":"127 WALNUT AVE","city":"LEESBURG","state":"GA","zip":"31763","phone":"2292340033","ytdComp":8544.65,"lat":31.730181,"lon":-84.171873,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101864,"salesman":"House","active":"Inactive","name":"SUPERIOR AUTO GROUP  INC -BILL","address":"1440 US 19 S","city":"LEESBURG","state":"GA","zip":"31763","phone":"229-349-6770","ytdComp":0.0,"lat":31.640131,"lon":-84.176128,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200138,"salesman":"House","active":"Inactive","name":"ALBANY TRACTOR COMPANY","address":"741 US HWY 82 WEST","city":"LEESBURG","state":"GA","zip":"31763","phone":"229-432-7468","ytdComp":0.0,"lat":31.661812,"lon":-84.275186,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200163,"salesman":"House","active":"Inactive","name":"FAIRCLOTH & SONS  INC","address":"1425 HWY 82 W","city":"LEESBURG","state":"GA","zip":"31763","phone":"912-435-2616","ytdComp":0.0,"lat":31.628842,"lon":-84.236839,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200167,"salesman":"House","active":"Inactive","name":"FLOYD ONE STOP TIRE","address":"105 WALNUT AVENUE","city":"LEESBURG","state":"GA","zip":"31763","phone":"229-759-2008","ytdComp":0.0,"lat":31.731691,"lon":-84.171893,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200191,"salesman":"House","active":"Active","name":"LEE COUNTY AUTO SERVICE","address":"264 WALNUT AVE","city":"LEESBURG","state":"GA","zip":"31763","phone":"2297592001","ytdComp":482.28,"lat":31.720766,"lon":-84.172766,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200192,"salesman":"House","active":"Active","name":"MCGEHEE'S TIRE & AUTO","address":"1359 US 82 W","city":"LEESBURG","state":"GA","zip":"31763","phone":"2294322846","ytdComp":72.64,"lat":31.63137,"lon":-84.240236,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200369,"salesman":"House","active":"Inactive","name":"SERVICE CEO'S","address":"107 STOCKS DAIRY RD","city":"LEESBURG","state":"GA","zip":"31763","phone":"229-434-0303","ytdComp":0.0,"lat":31.641628,"lon":-84.101668,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201053,"salesman":"Larry","active":"Active","name":"BERNEYS TIRE SERVICE","address":"1425 US HWY 82 WEST","city":"LEESBURG","state":"GA","zip":"31763","phone":"2294350413","ytdComp":8009.15,"lat":31.628842,"lon":-84.236839,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200590,"salesman":"House","active":"Inactive","name":"LEE COUNTY BOARD OF COMM.","address":"759 HWY 32 E","city":"LEESBURG","state":"GA","zip":"31763","phone":"2297596000","ytdComp":0.0,"lat":31.732461,"lon":-84.118474,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201067,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 2132","address":"1241 US HWY SOUTH","city":"LEESBURG","state":"GA","zip":"31763","phone":"2292314460","ytdComp":0.0,"lat":31.763715,"lon":-84.160359,"accuracy":0.5,"accuracyType":"place"},{"num":101466,"salesman":"Tiffany","active":"Active","name":"WATTS REPAIR SERVICE","address":"11 N OLD UNION RD","city":"LENOX","state":"GA","zip":"31637","phone":"2294024718","ytdComp":4174.75,"lat":31.094591,"lon":-83.478221,"accuracy":0.8,"accuracyType":"nearest_rooftop_match"},{"num":200212,"salesman":"House","active":"Inactive","name":"CLOSED","address":"101 OSBOURNE GRINER RD","city":"LENOX","state":"GA","zip":"31637","phone":"229-686-2047","ytdComp":0.0,"lat":31.324986,"lon":-83.38153,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200281,"salesman":"House","active":"Inactive","name":"LUKES SERVICE STATION","address":"271 KINARD BRIDGE RD","city":"LENOX","state":"GA","zip":"31637","phone":"2295464506","ytdComp":0.0,"lat":31.26798,"lon":-83.473078,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200406,"salesman":"Tiffany","active":"Active","name":"LENOX TIRE & SERVICE CENTER","address":"345 W CENTRAL AVE.","city":"LENOX","state":"GA","zip":"31637","phone":"2295464119","ytdComp":4772.14,"lat":31.270681,"lon":-83.469588,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200417,"salesman":"House","active":"Active","name":"HIGHWAY TIRE & DIESEL","address":"1070 HOMER JAMES RD","city":"LENOX","state":"GA","zip":"31637","phone":"2295464506","ytdComp":602.75,"lat":31.282699,"lon":-83.399424,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200462,"salesman":"House","active":"Active","name":"QUALITY FEEDSTUFFS  INC","address":"6122 ALAPAHA LENOX RD","city":"LENOX","state":"GA","zip":"31637","phone":"2296862770","ytdComp":659.22,"lat":31.318335,"lon":-83.311277,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200905,"salesman":"House","active":"Inactive","name":"DELTORO TIRE (LENOX)","address":"271 KINARD BRIDGE RD","city":"LENOX","state":"GA","zip":"31637","phone":"2295464438","ytdComp":0.0,"lat":31.26798,"lon":-83.473078,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200659,"salesman":"Tiffany","active":"Active","name":"TOWN & COUNTRY TIRE","address":"317 HOWARD ST E","city":"LIVE OAK","state":"FL","zip":"32064","phone":"3863624535","ytdComp":1096.96,"lat":30.294891,"lon":-82.981734,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200687,"salesman":"Tiffany","active":"Active","name":"LIVE OAK TIRE CENTER  LLC","address":"303 WEST HOWARD ST.","city":"LIVE OAK","state":"FL","zip":"32064","phone":"3863621972","ytdComp":8559.91,"lat":30.296928,"lon":-82.986543,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200898,"salesman":"House","active":"Active","name":"PRECISION AUTO & MUFFLER LLC","address":"500 WEST HOWARD ST.","city":"LIVE OAK","state":"FL","zip":"32064","phone":"3863641055","ytdComp":1881.28,"lat":30.29849,"lon":-82.98971,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200900,"salesman":"House","active":"Active","name":"GILLETTES AUTO","address":"4673 US HWY 129","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3863625171","ytdComp":71.64,"lat":30.364674,"lon":-82.941769,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200913,"salesman":"House","active":"Active","name":"LASHLEY'S HOMETOWN TIRE LLC","address":"9916 US HWY 90","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3862094919","ytdComp":7360.26,"lat":30.284124,"lon":-82.95516,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200931,"salesman":"House","active":"Active","name":"A&B CUSTOMS","address":"13732 US HWY 90","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3863621350","ytdComp":0.0,"lat":30.317121,"lon":-83.031841,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200961,"salesman":"Tiffany","active":"Active","name":"BABCOCK TIRE LLC","address":"315 MARTHA ST SE (SHIPPING)","city":"LIVE OAK","state":"FL","zip":"32064","phone":"3862056386","ytdComp":6824.86,"lat":30.285149,"lon":-82.985005,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201008,"salesman":"House","active":"Active","name":"DBJ MOBILE TIRE SERVICE  INC.","address":"14408 97 RD","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3862196036","ytdComp":12343.54,"lat":30.192251,"lon":-82.938432,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201018,"salesman":"House","active":"Active","name":"CRAWLEY'S AUTOMOTIVE & TIRE","address":"22922 COUNTY ROAD 250","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3866582007","ytdComp":1422.42,"lat":30.243443,"lon":-83.222562,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000011,"salesman":"Car Dealer","active":"Active","name":"MAVIS TIRES & BRAKES - 2180","address":"2245 OHIO AVENUE N","city":"LIVE OAK","state":"FL","zip":"32060","phone":"3868558004","ytdComp":0.0,"lat":30.32177,"lon":-82.96851,"accuracy":0.7,"accuracyType":"street_center"},{"num":200624,"salesman":"House","active":"Active","name":"WALLACE MOTORS","address":"1182 EAST US 90","city":"MADISON","state":"FL","zip":"32340","phone":"8509731230","ytdComp":2209.65,"lat":30.468946,"lon":-83.396827,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200636,"salesman":"Tiffany","active":"Active","name":"MTC SOUTH  INC.","address":"736 SW HARVEY GREEN DR.","city":"MADISON","state":"FL","zip":"32340","phone":"8502515393","ytdComp":42953.73,"lat":30.448958,"lon":-83.41431,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200670,"salesman":"House","active":"Active","name":"HALL'S TIRE & MUFFLER CENTER","address":"1064 EAST US 90","city":"MADISON","state":"FL","zip":"32340","phone":"8509733026","ytdComp":0.0,"lat":30.468765,"lon":-83.398892,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200671,"salesman":"House","active":"Active","name":"STEWARTS AUTO SERVICE CENTER","address":"115 S.W. BUNKER ST","city":"MADISON","state":"FL","zip":"32340","phone":"8509734088","ytdComp":432.4,"lat":30.465046,"lon":-83.410307,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200691,"salesman":"Tiffany","active":"Active","name":"KENDA TRUCK CENTER","address":"1087 NW COUNTY RD 150","city":"MADISON","state":"FL","zip":"32340","phone":"8509293700","ytdComp":2221.9,"lat":30.583567,"lon":-83.444093,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000038,"salesman":"House","active":"Active","name":"FAST TIRE SERVICE","address":"660 NE COLIN KELLY HWY","city":"MADISON","state":"FL","zip":"32340","phone":"3863241853","ytdComp":0.0,"lat":30.472578,"lon":-83.405593,"accuracy":0.87,"accuracyType":"nearest_rooftop_match"},{"num":200158,"salesman":"House","active":"Inactive","name":"BARNES TIRE & SUPPLY","address":"4458 JACKSON ST","city":"MARIANNA","state":"FL","zip":"32447","phone":"850-926-3813","ytdComp":0.0,"lat":30.773272,"lon":-85.224116,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200207,"salesman":"House","active":"Inactive","name":"PANHANDLE TRACTOR INC","address":"5003 HWY 90","city":"MARIANNA","state":"FL","zip":"32446","phone":"850-526-2257","ytdComp":0.0,"lat":30.746324,"lon":-85.178853,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200215,"salesman":"House","active":"Inactive","name":"RAHAL CHEVROLET","address":"4204 LAFAYETTE ST","city":"MARIANNA","state":"FL","zip":"32447","phone":"850-482-3051","ytdComp":0.0,"lat":30.782409,"lon":-85.245729,"accuracy":0.99,"accuracyType":"rooftop"},{"num":201216,"salesman":"House","active":"Active","name":"FIRESTONE STORE #003441","address":"4386 LAFAYETTE ST","city":"MARIANNA","state":"FL","zip":"32446","phone":"8504824991","ytdComp":0.0,"lat":30.775314,"lon":-85.230337,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100649,"salesman":"House","active":"Inactive","name":"CLOSED ACCT","address":"0","city":"MAYO","state":"FL","zip":"","phone":"","ytdComp":0.0,"lat":30.053,"lon":-83.17486,"accuracy":0.5,"accuracyType":"place"},{"num":100825,"salesman":"House","active":"Inactive","name":"CLOSED","address":"0","city":"MAYO","state":"FL","zip":"","phone":"","ytdComp":0.0,"lat":30.053,"lon":-83.17486,"accuracy":0.5,"accuracyType":"place"},{"num":200164,"salesman":"House","active":"Inactive","name":"CLOSED","address":"0","city":"MAYO","state":"FL","zip":"","phone":"","ytdComp":0.0,"lat":30.053,"lon":-83.17486,"accuracy":0.5,"accuracyType":"place"},{"num":200260,"salesman":"House","active":"Inactive","name":"A","address":"0","city":"MAYO","state":"FL","zip":"","phone":"","ytdComp":0.0,"lat":30.053,"lon":-83.17486,"accuracy":0.5,"accuracyType":"place"},{"num":200499,"salesman":"House","active":"Inactive","name":"0","address":"0","city":"MAYO","state":"FL","zip":"","phone":"","ytdComp":0.0,"lat":30.053,"lon":-83.17486,"accuracy":0.5,"accuracyType":"place"},{"num":200595,"salesman":"Tiffany","active":"Active","name":"W.R. WILLIAMS","address":"1404 EAST MAIN","city":"MAYO","state":"FL","zip":"32066","phone":"3862941888","ytdComp":27864.74,"lat":30.053083,"lon":-83.153478,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000019,"salesman":"Car Dealer","active":"Active","name":"CONLAN TIRE CO. LLC (MIAMI)","address":"2005 NW 110 TH AVE","city":"MIAMI","state":"FL","zip":"33172","phone":"3055256060","ytdComp":0.0,"lat":25.792477,"lon":-80.373546,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200910,"salesman":"House","active":"Inactive","name":"ANDY'S SMART TIRES LLC","address":"2901 CREEK ST","city":"MIDDLEBURG","state":"FL","zip":"32068","phone":"9045455328","ytdComp":0.0,"lat":30.04209,"lon":-81.872066,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201217,"salesman":"House","active":"Active","name":"TIRES PLUS #325121","address":"1781 BLANDING BLVD.","city":"MIDDLEBURG","state":"FL","zip":"32068","phone":"9045899993","ytdComp":0.0,"lat":30.104839,"lon":-81.826786,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200577,"salesman":"House","active":"Inactive","name":"MIDWAY TIRE","address":"1505 COMMERCE BLVD","city":"MIDWAY","state":"FL","zip":"32343","phone":"850-580-4010","ytdComp":0.0,"lat":30.499158,"lon":-84.412437,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200902,"salesman":"House","active":"Active","name":"CONLAN TIRE CO. LLC (#015)","address":"60 RAM BLVD","city":"MIDWAY","state":"FL","zip":"32304","phone":"8508488887","ytdComp":0.0,"lat":30.490915,"lon":-84.417903,"accuracy":0.85,"accuracyType":"nearest_rooftop_match"},{"num":200387,"salesman":"House","active":"Inactive","name":"TIRE SHOP","address":"7803 FULLER RD","city":"MILLWOOD","state":"GA","zip":"31552","phone":"912-283-6998","ytdComp":0.0,"lat":31.258189,"lon":-82.643888,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200575,"salesman":"House","active":"Active","name":"BIG BEND TIRE","address":"1300 NORTH JEFFERSON","city":"MONTICELLO","state":"FL","zip":"32344","phone":"8509974689","ytdComp":0.0,"lat":30.556587,"lon":-83.869872,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200576,"salesman":"House","active":"Inactive","name":"SORENSEN TIRE","address":"0","city":"MONTICELLO","state":"FL","zip":"32344","phone":"850-997-4689","ytdComp":0.0,"lat":30.524354,"lon":-83.927906,"accuracy":1.0,"accuracyType":"street_center"},{"num":200753,"salesman":"House","active":"Active","name":"AUTO TECH OF MIAMI INC.","address":"203 ASHVILLE HWY","city":"MONTICELLO","state":"FL","zip":"32344","phone":"8509970200","ytdComp":474.84,"lat":30.548862,"lon":-83.851346,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000007,"salesman":"House","active":"Active","name":"KEATON & SON TIRE LLC","address":"890 N. JEFFERSON ST","city":"MONTICELLO","state":"FL","zip":"32344","phone":"8502849344","ytdComp":2066.55,"lat":30.552324,"lon":-83.86975,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200672,"salesman":"House","active":"Active","name":"TIMMY'S AUTO REPAIR","address":"11774 ADEL RD","city":"MORVEN","state":"GA","zip":"31638","phone":"2295630470","ytdComp":0.0,"lat":30.948972,"lon":-83.499355,"accuracy":0.8,"accuracyType":"rooftop"},{"num":200978,"salesman":"House","active":"Active","name":"R&S LIFT TRUCKS CO. LLC","address":"11279 ADEL HWY","city":"MORVEN","state":"GA","zip":"31638","phone":"2297752221","ytdComp":0.0,"lat":30.941856,"lon":-83.49954,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200993,"salesman":"House","active":"Active","name":"STEVEN'S FORKLIFT REPAIR","address":"11311 ADEL HWY","city":"MORVEN","state":"GA","zip":"31638","phone":"2292516440","ytdComp":0.0,"lat":30.942192,"lon":-83.499484,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101361,"salesman":"House","active":"Active","name":"SOUTHERN TIRE CO","address":"176 W. BYPASS","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299855588","ytdComp":0.0,"lat":31.197149,"lon":-83.792317,"accuracy":0.94,"accuracyType":"nearest_rooftop_match"},{"num":200170,"salesman":"House","active":"Inactive","name":"EDWARDS MOTORS BUICK PONTIAC","address":"1825 NORTH VETERANS PARKWAY","city":"MOULTRIE","state":"GA","zip":"31788","phone":"229-985-3606","ytdComp":0.0,"lat":31.152706,"lon":-83.760863,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200190,"salesman":"House","active":"Inactive","name":"LASSITER TRACTOR CO","address":"1000 VETERANS PKWY","city":"MOULTRIE","state":"GA","zip":"31776","phone":"229-985-1027","ytdComp":0.0,"lat":31.201847,"lon":-83.771506,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200195,"salesman":"House","active":"Active","name":"MOULTRIE AUTOMOTIVE","address":"11 6TH ST SE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299851700","ytdComp":0.0,"lat":31.179458,"lon":-83.781416,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200225,"salesman":"House","active":"Inactive","name":"SOUTHERN STATES / EPM SHOP","address":"804 SECOND ST NE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-891-3931","ytdComp":0.0,"lat":31.187368,"lon":-83.786216,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200233,"salesman":"House","active":"Inactive","name":"SUNSET TIRE & AUTOMOTIVE","address":"2431 SOUTH MAIN","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-891-3366","ytdComp":0.0,"lat":31.137563,"lon":-83.783298,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200307,"salesman":"House","active":"Active","name":"ARREDONDO TIRE SERVICE","address":"362 HWY 319 N","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2295290882","ytdComp":1251.02,"lat":31.207326,"lon":-83.779667,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200343,"salesman":"House","active":"Inactive","name":"JIM 4 AUTO SERVICE","address":"301 2ND ST3","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-985-5464","ytdComp":0.0,"lat":31.18303,"lon":-83.786641,"accuracy":0.9,"accuracyType":"rooftop"},{"num":200353,"salesman":"House","active":"Inactive","name":"MERCER MOTORS","address":"2700 1ST AVE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"229-985-8782","ytdComp":0.0,"lat":31.172396,"lon":-83.751529,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200365,"salesman":"Car Dealer","active":"Active","name":"ROBERT HUTSON LINCOLN","address":"2316 1ST AVE SE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2299856603","ytdComp":627.96,"lat":31.174458,"lon":-83.755189,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200397,"salesman":"House","active":"Inactive","name":"LEWIS TIRE SERVICE","address":"13 5TH AVE NE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-890-1051","ytdComp":0.0,"lat":31.185121,"lon":-83.788248,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200398,"salesman":"House","active":"Inactive","name":"RICHARDSON TIRES","address":"425 W CENTRAL AVE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-891-3992","ytdComp":0.0,"lat":31.179808,"lon":-83.793662,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200399,"salesman":"House","active":"Inactive","name":"MCKEES AUTO CENTER","address":"10 5TH ST SE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-985-0137","ytdComp":0.0,"lat":31.179525,"lon":-83.783203,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200427,"salesman":"House","active":"Active","name":"SAUNDERS AUTO REPAIR","address":"1869 SYLVESTER HWY","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2296161041","ytdComp":352.64,"lat":31.210088,"lon":-83.79085,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101513,"salesman":"Larry","active":"Active","name":"GAY'S TIRE SERVICE","address":"400 2ND ST NE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299857656","ytdComp":0.0,"lat":31.184339,"lon":-83.785884,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200198,"salesman":"Larry","active":"Active","name":"MOULTRIE TIRE","address":"900 N MAIN ST","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299855619","ytdComp":28969.46,"lat":31.189244,"lon":-83.788314,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200242,"salesman":"Larry","active":"Active","name":"THOMAS TIRE RECAPPING INC.","address":"2700 1ST AVE SE","city":"MOULTRIE","state":"GA","zip":"31776","phone":"2299851839","ytdComp":0.0,"lat":31.172396,"lon":-83.751529,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200317,"salesman":"Larry","active":"Active","name":"BROTHERS TIRES","address":"1017 N MAIN ST","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2293192919","ytdComp":7559.19,"lat":31.19067,"lon":-83.789568,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200537,"salesman":"House","active":"Active","name":"TONY'S TIRE & ROAD SERVICE INC","address":"1320 N. MAIN ST","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2298909989","ytdComp":6780.03,"lat":31.19386,"lon":-83.789258,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200628,"salesman":"Larry","active":"Active","name":"SOUTH GEORGIA TIRE","address":"323 SOUTH MAIN ST","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2296688116","ytdComp":10156.52,"lat":31.176015,"lon":-83.788514,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200755,"salesman":"Larry","active":"Active","name":"N-T TIRE SERVICE","address":"233 PINEVIEW AVE.","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2298912234","ytdComp":10532.98,"lat":31.192402,"lon":-83.749331,"accuracy":0.99,"accuracyType":"rooftop"},{"num":201035,"salesman":"Larry","active":"Active","name":"DAVID'S AUTO SALES (MOULTRIE)","address":"828 VETERANS PKWY","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2292174959","ytdComp":10905.28,"lat":31.198643,"lon":-83.768172,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000015,"salesman":"Larry","active":"Active","name":"EDWIN'S TIRES LLC","address":"1528 SYLVESTER HWY","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2294296141","ytdComp":0.0,"lat":31.201424,"lon":-83.788414,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000020,"salesman":"Larry","active":"Active","name":"T&D TIRE","address":"120 CEPCOT MEADOWS","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2294563436","ytdComp":0.0,"lat":31.122842,"lon":-83.913376,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200634,"salesman":"House","active":"Inactive","name":"TEXAS TIRE","address":"1033 W. CENTRAL AVE.","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-798-5061","ytdComp":0.0,"lat":31.179529,"lon":-83.800439,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200643,"salesman":"House","active":"Inactive","name":"LANGDALE POWERSPORTS  LLC","address":"795 VETERANS PKWY","city":"MOULTRIE","state":"GA","zip":"31788","phone":"229-776-3473","ytdComp":0.0,"lat":31.194962,"lon":-83.767597,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200657,"salesman":"House","active":"Inactive","name":"KIMBRELL SERVICE CENTER","address":"324 S. MAIN ST.","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-985-1952","ytdComp":0.0,"lat":31.176148,"lon":-83.789264,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200690,"salesman":"House","active":"Active","name":"COLQUITT COUNTY TIRE LLC","address":"2356 HWY 133 S.","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2294541084","ytdComp":38904.19,"lat":31.12791,"lon":-83.712249,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200697,"salesman":"House","active":"Inactive","name":"SUNSET TIRE & AUTOMOTIVE SVC.","address":"2431 S. MAIN ST","city":"MOULTRIE","state":"GA","zip":"31768","phone":"229-891-3366","ytdComp":0.0,"lat":31.137563,"lon":-83.783298,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200780,"salesman":"House","active":"Inactive","name":"HUTSON CHRYSLER DODGE","address":"2500 1ST AVE SE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"229-329-1819","ytdComp":0.0,"lat":31.173369,"lon":-83.754282,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200802,"salesman":"House","active":"Inactive","name":"SOUTHERN AUTO. & AG REPAIR LLC","address":"2034 SYLVESTER HWY","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2292170558","ytdComp":0.0,"lat":31.214864,"lon":-83.789714,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200829,"salesman":"House","active":"Active","name":"TIRE SOLUTIONS & VEH. REPAIRS","address":"172 WEST BYPASS","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299858473","ytdComp":12884.35,"lat":31.199349,"lon":-83.871326,"accuracy":0.92,"accuracyType":"range_interpolation"},{"num":200850,"salesman":"Car Dealer","active":"Active","name":"HUTSON CHRYSLER DODGE JEEP RAM","address":"2500 1ST AVE SE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2293291819","ytdComp":0.0,"lat":31.173369,"lon":-83.754282,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200856,"salesman":"House","active":"Inactive","name":"ROBERT HUTSON LINC. (AMI ACCT)","address":"2316 1ST AVE SE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2299856603","ytdComp":0.0,"lat":31.174458,"lon":-83.755189,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200896,"salesman":"House","active":"Active","name":"SOUTHERN AUTO SPECIALIST","address":"2034 SYLVESTER HWY","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2292170558","ytdComp":3929.29,"lat":31.214864,"lon":-83.789714,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200947,"salesman":"House","active":"Active","name":"SUNSET TIRE & AUTOMOTIVE","address":"2431 S. MAIN ST.","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2298913366","ytdComp":6696.29,"lat":31.137563,"lon":-83.783298,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200948,"salesman":"Car Dealer","active":"Active","name":"LASSETER CHEVROLET","address":"1825 VETERANS PKWY","city":"MOULTRIE","state":"GA","zip":"31776","phone":"2299853606","ytdComp":0.0,"lat":31.152706,"lon":-83.760863,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200975,"salesman":"House","active":"Active","name":"SANTOS TIRE SHOP","address":"624 W CENTRAL AVE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299216310","ytdComp":4384.34,"lat":31.1802,"lon":-83.796401,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201003,"salesman":"House","active":"Active","name":"NICHOLAS TIRES INC.","address":"1014 W CENTRAL AVENUE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2298500289","ytdComp":8988.01,"lat":31.18014,"lon":-83.79989,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201022,"salesman":"House","active":"Inactive","name":"KUSTOM UPHOLSTERY LLC","address":"2509 1ST AVE","city":"MOULTRIE","state":"GA","zip":"31788","phone":"2294494546","ytdComp":0.0,"lat":31.174155,"lon":-83.753556,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000030,"salesman":"Austin","active":"Active","name":"BEASON EQUIPMENT CO","address":"720 1ST STREET NE","city":"MOULTRIE","state":"GA","zip":"31768","phone":"2299859785","ytdComp":405.56,"lat":31.186886,"lon":-83.78721,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200584,"salesman":"House","active":"Inactive","name":"WOODARD'S AUTO CENTER INC.","address":"10511 N. MAIN","city":"NAHUNTA","state":"GA","zip":"31553","phone":"912-462-6850","ytdComp":0.0,"lat":31.211642,"lon":-81.978826,"accuracy":0.85,"accuracyType":"nearest_rooftop_match"},{"num":200478,"salesman":"Larry","active":"Active","name":"82 TIRE & LUBE","address":"9324 MAIN ST SOUTH","city":"NAHUNTA","state":"GA","zip":"31553","phone":"9124627357","ytdComp":6270.54,"lat":31.200166,"lon":-81.983233,"accuracy":0.7,"accuracyType":"street_center"},{"num":101549,"salesman":"House","active":"Active","name":"NASHVILLE TIRE","address":"730 S DAVIS ST","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296861900","ytdComp":13329.71,"lat":31.194821,"lon":-83.251043,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200351,"salesman":"House","active":"Inactive","name":"LEONARD TIRE CO.","address":"7087 HWY 129","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-686-7931","ytdComp":0.0,"lat":31.171386,"lon":-83.231879,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200371,"salesman":"House","active":"Inactive","name":"SOUTH GEORGIA OFFROAD","address":"506 E MCPHERSON AVE","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-686-1850","ytdComp":0.0,"lat":31.205289,"lon":-83.243229,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200443,"salesman":"House","active":"Inactive","name":"MIKE BURCH FORD (NASHVILLE)","address":"723 S DAVIS ST","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-686-2058","ytdComp":0.0,"lat":31.195753,"lon":-83.252033,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200450,"salesman":"House","active":"Inactive","name":"FRANKLINS AUTOMOTIVE","address":"7705 HWY 129","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-686-3979","ytdComp":0.0,"lat":31.179665,"lon":-83.237826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200451,"salesman":"House","active":"Active","name":"HARROD BROTHERS","address":"807 TIFTON HWY","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296863959","ytdComp":1393.6,"lat":31.213533,"lon":-83.260501,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200453,"salesman":"House","active":"Inactive","name":"CJ MARTIN INC. CDJR","address":"(DBA MARTIN MOTORS)","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-686-2068","ytdComp":0.0,"lat":31.215906,"lon":-83.240207,"accuracy":0.78,"accuracyType":"street_center"},{"num":200569,"salesman":"House","active":"Inactive","name":"C&M TRANSMISSION & AUTO REPAIR","address":"208 EAST DENNIS AVE","city":"NASHVILLE","state":"GA","zip":"31639","phone":"229-326-3449","ytdComp":0.0,"lat":31.203624,"lon":-83.24916,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200622,"salesman":"House","active":"Active","name":"MOORE'S ACCESSORIES & OFFROAD","address":"7643 HWY 125 NORTH","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2292236389","ytdComp":20.86,"lat":31.279825,"lon":-83.332542,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101415,"salesman":"Larry","active":"Active","name":"THE TIRE STORE","address":"523 S JEFFERSON ST","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862073","ytdComp":35026.54,"lat":31.201194,"lon":-83.250115,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200319,"salesman":"Larry","active":"Active","name":"BUCK'S AUTO REPAIR","address":"307 HAZEL AVE","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862290","ytdComp":7271.12,"lat":31.198671,"lon":-83.246314,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200688,"salesman":"Car Dealer","active":"Active","name":"NASHVILLE FORD","address":"723 S. DAVIS ST","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862058","ytdComp":0.0,"lat":31.195753,"lon":-83.252033,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200843,"salesman":"House","active":"Inactive","name":"CJ MARTIN INC. CDJR (AMI ACCT)","address":"(DBA MARTIN MOTORS)","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862068","ytdComp":0.0,"lat":31.215906,"lon":-83.240207,"accuracy":0.78,"accuracyType":"street_center"},{"num":200853,"salesman":"House","active":"Inactive","name":"NASHVILLE FORD (AMI ACCT)","address":"723 S. DAVIS ST","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862058","ytdComp":0.0,"lat":31.195753,"lon":-83.252033,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200908,"salesman":"Car Dealer","active":"Active","name":"O'STEEN CHRYSLER DODGE JEEP","address":"706 S. DAVIS ST.","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862068","ytdComp":364.96,"lat":31.198461,"lon":-83.250952,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200941,"salesman":"House","active":"Active","name":"D&S WHEELS & DEALS LLC","address":"503 HAZEL AVE","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2292373167","ytdComp":4809.71,"lat":31.198612,"lon":-83.244222,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000009,"salesman":"Car Dealer","active":"Active","name":"KING FORD OF NASHVILLE","address":"723 S. DAVIS ST.","city":"NASHVILLE","state":"GA","zip":"31639","phone":"2296862058","ytdComp":3173.44,"lat":31.195753,"lon":-83.252033,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201042,"salesman":"House","active":"Active","name":"TIRE AGENT CORP (TIFTON WHS)","address":"101 WEST 23RD ST","city":"NEW YORK","state":"NY","zip":"10011","phone":"8338473463","ytdComp":6510.51,"lat":40.743264,"lon":-73.993244,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200174,"salesman":"House","active":"Inactive","name":"HUSCO AUTOMOTIVE","address":"41 SANDBED RD","city":"NEWTON","state":"AL","zip":"36352","phone":"334-692-5478","ytdComp":0.0,"lat":31.2317,"lon":-85.635431,"accuracy":0.7,"accuracyType":"street_center"},{"num":200489,"salesman":"House","active":"Inactive","name":"MCCONNELL'S","address":"191 GA. HWY 91 S.W.","city":"NEWTON","state":"GA","zip":"39870","phone":"229-734-5128","ytdComp":0.0,"lat":31.174059,"lon":-84.733495,"accuracy":0.89,"accuracyType":"range_interpolation"},{"num":101066,"salesman":"Tiffany","active":"Active","name":"WARRIOR CREEK TIRE  LLC","address":"2815 COOL SPRINGS RD","city":"NORMAN PARK","state":"GA","zip":"31771","phone":"2297980923","ytdComp":10477.74,"lat":31.220977,"lon":-83.681471,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200203,"salesman":"House","active":"Inactive","name":"OWENS AUTO SALES","address":"1428 HWY 319 N","city":"NORMAN PARK","state":"GA","zip":"31771","phone":"229-891-3791","ytdComp":0.0,"lat":31.221238,"lon":-83.754211,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200426,"salesman":"House","active":"Inactive","name":"WESLEY'S AUTO SALES","address":"4427 HWY 319 N","city":"NORMAN PARK","state":"GA","zip":"31771","phone":"229-769-3954","ytdComp":0.0,"lat":31.276091,"lon":-83.684296,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200642,"salesman":"House","active":"Active","name":"E.G. AUTO SALES","address":"4333 HWY 319","city":"NORMAN PARK","state":"GA","zip":"31771","phone":"2297695011","ytdComp":873.04,"lat":31.27292,"lon":-83.685638,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200696,"salesman":"House","active":"Active","name":"SNIDER FLEET SOLUTIONS","address":"443 SW 54TH CT","city":"OCALA","state":"FL","zip":"34474","phone":"2293168964","ytdComp":0.0,"lat":29.182946,"lon":-82.208993,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200922,"salesman":"Tiffany","active":"Active","name":"ADVANCED TIRE SERVICE","address":"2199 NW 10TH ST","city":"OCALA","state":"FL","zip":"34475","phone":"3522368825","ytdComp":137059.47,"lat":29.19712,"lon":-82.16037,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000033,"salesman":"House","active":"Active","name":"AFFORDABLE TIRE & AUTO CARE  LLC","address":"640 NW 27TH AVE","city":"OCALA","state":"FL","zip":"34475","phone":"3524215575","ytdComp":0.0,"lat":29.19278,"lon":-82.16859,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000034,"salesman":"House","active":"Active","name":"OCALA TRUCK & CAR CENTER","address":"2608 NW 6TH ST","city":"OCALA","state":"FL","zip":"34475","phone":"3523729033","ytdComp":0.0,"lat":29.19229,"lon":-82.1672,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000039,"salesman":"Tiffany","active":"Active","name":"ADVANCED TIRE SERVICE","address":"2418 EW SILVER SPRINGS BLVD","city":"OCALA","state":"FL","zip":"34470","phone":"3522368825","ytdComp":336.62,"lat":29.186417,"lon":-82.10481,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000041,"salesman":"Tiffany","active":"Active","name":"OCALA TIRE SERVICE","address":"424 S. MAGNOLIA AVE.","city":"OCALA","state":"FL","zip":"34471","phone":"3526227233","ytdComp":0.0,"lat":29.183143,"lon":-82.136982,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101163,"salesman":"House","active":"Inactive","name":"BARFIELD'S INC.","address":"410 S IRWIN AVE","city":"OCILLA","state":"GA","zip":"31774","phone":"2294685961","ytdComp":0.0,"lat":31.5906,"lon":-83.250633,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200542,"salesman":"House","active":"Inactive","name":"IRWIN COUNTY BD. OF EDUCATION","address":"255 CHIEFTON CIRCLE","city":"OCILLA","state":"GA","zip":"31774","phone":"2294687485","ytdComp":0.0,"lat":31.584483,"lon":-83.239633,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101181,"salesman":"Larry","active":"Active","name":"SOUTH GA LUBE CENTER","address":"133 FITZGERALD HWY","city":"OCILLA","state":"GA","zip":"31774","phone":"2294684435","ytdComp":2030.33,"lat":31.609883,"lon":-83.250083,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200955,"salesman":"House","active":"Inactive","name":"IRWIN CO. BD. OF COMMISSIONERS","address":"207 SOUTH IRWIN AVE","city":"OCILLA","state":"GA","zip":"31774","phone":"2294689441","ytdComp":0.0,"lat":31.592433,"lon":-83.250483,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101439,"salesman":"House","active":"Active","name":"A.T. TIRE SERVICE","address":"101 OMEGA ELLENTON RD","city":"OMEGA","state":"GA","zip":"31775","phone":"2298915428","ytdComp":14121.15,"lat":31.340815,"lon":-83.593635,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101511,"salesman":"House","active":"Inactive","name":"OMEGA FARM SUPPLY","address":"P.O. BOX 97","city":"OMEGA","state":"GA","zip":"31775","phone":"229-528-4227","ytdComp":0.0,"lat":31.34102,"lon":-83.5935,"accuracy":1.0,"accuracyType":"place"},{"num":200341,"salesman":"House","active":"Inactive","name":"JC TIRE SERVICE","address":"262 OAK ST","city":"OMEGA","state":"GA","zip":"31775","phone":"229-528-3019","ytdComp":0.0,"lat":31.338438,"lon":-83.594866,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201039,"salesman":"House","active":"Active","name":"CERVANTES AUTO SALES","address":"6516 US HWY 319 N","city":"OMEGA","state":"GA","zip":"31775","phone":"2297695201","ytdComp":0.0,"lat":31.312967,"lon":-83.63102,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201020,"salesman":"House","active":"Active","name":"MILLER BROS GIANT TIRE SVC","address":"925 W. LANCASTER RD","city":"ORLANDO","state":"FL","zip":"32809","phone":"4078559621","ytdComp":0.0,"lat":28.465619,"lon":-81.391971,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000005,"salesman":"House","active":"Active","name":"GOODYEAR COMMERCIAL TIRE & SVC","address":"971 TAFT VINELAND RD","city":"ORLANDO","state":"FL","zip":"32824","phone":"4078551182","ytdComp":0.0,"lat":28.42201,"lon":-81.39083,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200457,"salesman":"House","active":"Inactive","name":"ROADMART  INC","address":"1880 SOUTH UNION AVE","city":"OZARK","state":"AL","zip":"36360","phone":"3347749345","ytdComp":0.0,"lat":31.484427,"lon":-85.644796,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200954,"salesman":"House","active":"Active","name":"FIRST COAST TIRES INC.","address":"226 SAINT JOHNS DR","city":"PALATKA","state":"FL","zip":"32177","phone":"9042373356","ytdComp":0.0,"lat":29.749438,"lon":-81.568299,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200479,"salesman":"House","active":"Inactive","name":"FARMERS DEPOT","address":"6275 HWY 84","city":"PATTERSON","state":"GA","zip":"31557","phone":"912-647-0105","ytdComp":0.0,"lat":31.379942,"lon":-82.145635,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200389,"salesman":"House","active":"Inactive","name":"QUALITY LUBE","address":"26 MAIN ST","city":"PEARSON","state":"GA","zip":"31642","phone":"912-422-6729","ytdComp":0.0,"lat":31.297286,"lon":-82.852821,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101297,"salesman":"Larry","active":"Active","name":"FOUR C'S LUBE","address":"915 ALBANY AVENUE","city":"PEARSON","state":"GA","zip":"31642","phone":"9124226866","ytdComp":11989.49,"lat":31.300483,"lon":-82.867985,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200388,"salesman":"Larry","active":"Active","name":"PEARSON TIRE & LUBE","address":"229 MAIN ST","city":"PEARSON","state":"GA","zip":"31642","phone":"9124226820","ytdComp":4186.97,"lat":31.302665,"lon":-82.851539,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200762,"salesman":"Larry","active":"Active","name":"POWER MAN TIRE SHOP","address":"26 S MAIN ST.","city":"PEARSON","state":"GA","zip":"31642","phone":"9123814065","ytdComp":3476.8,"lat":31.297286,"lon":-82.852821,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101542,"salesman":"House","active":"Inactive","name":"WEST SERVICE STATION","address":"246 CHURCH ST","city":"PELHAM","state":"GA","zip":"31779","phone":"229-213-5025","ytdComp":0.0,"lat":31.123481,"lon":-84.154527,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101544,"salesman":"Tiffany","active":"Active","name":"GODWIN TIRE & AUTO","address":"7271 COUNTYLINE RD","city":"PELHAM","state":"GA","zip":"31779","phone":"2292949553","ytdComp":5831.5,"lat":31.079167,"lon":-84.288013,"accuracy":0.94,"accuracyType":"nearest_rooftop_match"},{"num":200176,"salesman":"House","active":"Inactive","name":"CLOSED","address":"7273 COUNTY LINE RD","city":"PELHAM","state":"GA","zip":"31779","phone":"2292942224","ytdComp":0.0,"lat":31.079167,"lon":-84.288013,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200210,"salesman":"House","active":"Inactive","name":"PELHAM TIRE CO","address":"254 GLAUSIER ST","city":"PELHAM","state":"GA","zip":"31779","phone":"229-294-0508","ytdComp":0.0,"lat":31.132001,"lon":-84.147943,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200328,"salesman":"House","active":"Inactive","name":"DREKO AUTOMOTIVE","address":"635 BARROW AVE","city":"PELHAM","state":"GA","zip":"31779","phone":"229-294-8571","ytdComp":0.0,"lat":31.114567,"lon":-84.164317,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200342,"salesman":"House","active":"Inactive","name":"JCB LOGISTICS","address":"259 WILLIAMS ST NE","city":"PELHAM","state":"GA","zip":"31779","phone":"229-294-8149","ytdComp":0.0,"lat":31.133427,"lon":-84.147849,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200367,"salesman":"House","active":"Inactive","name":"ROY'S PLACE ON THE CORNER","address":"145 MATTHEWSON AVE","city":"PELHAM","state":"GA","zip":"31779","phone":"229-294-5471","ytdComp":0.0,"lat":31.125672,"lon":-84.151158,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200518,"salesman":"House","active":"Inactive","name":"BURGESS AUTOMOTIVE","address":"0","city":"PELHAM","state":"GA","zip":"31779","phone":"","ytdComp":0.0,"lat":31.148991,"lon":-84.188934,"accuracy":0.9,"accuracyType":"street_center"},{"num":200695,"salesman":"House","active":"Active","name":"PELHAM TIRE & EQUIPMENT COMP.","address":"3027 HWY 19 N.E.","city":"PELHAM","state":"GA","zip":"31779","phone":"2292942801","ytdComp":0.0,"lat":31.15257,"lon":-84.144931,"accuracy":0.83,"accuracyType":"nearest_rooftop_match"},{"num":200779,"salesman":"House","active":"Inactive","name":"MITCHELL COUNTY CHRYSLER DODGE","address":"3604 NICKS RD.","city":"PELHAM","state":"GA","zip":"31779","phone":"229-329-1823","ytdComp":0.0,"lat":31.145192,"lon":-84.145929,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200852,"salesman":"House","active":"Inactive","name":"MITCHELL CO. CDJR (AMI ACCT)","address":"3604 MICKS RD.","city":"PELHAM","state":"GA","zip":"31779","phone":"2293291823","ytdComp":0.0,"lat":31.12768,"lon":-84.15185,"accuracy":0.5,"accuracyType":"place"},{"num":2000013,"salesman":"House","active":"Active","name":"TIRE SOLUTIONS & VEH. REPAIRS","address":"3495 MILL POND RD","city":"PELHAM","state":"GA","zip":"31779","phone":"2292942801","ytdComp":12103.26,"lat":31.155713,"lon":-84.142742,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200596,"salesman":"Tiffany","active":"Active","name":"CRIBBS TIRE","address":"1005 S JEFFERSON STREET","city":"PERRY","state":"FL","zip":"32348","phone":"8505843883","ytdComp":0.0,"lat":30.10916,"lon":-83.5821,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200627,"salesman":"House","active":"Inactive","name":"RYAN'S EVERYTHING AUTOMOTIVE","address":"4053 US HIGHWAY 19 S","city":"PERRY","state":"FL","zip":"32348","phone":"8505848900","ytdComp":0.0,"lat":30.068998,"lon":-83.560665,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200667,"salesman":"House","active":"Active","name":"CAMPBELL AUTO REPAIR","address":"1137 N. BYRON BUTLER","city":"PERRY","state":"FL","zip":"32347","phone":"8508382167","ytdComp":0.0,"lat":30.12965,"lon":-83.599246,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200673,"salesman":"Tiffany","active":"Active","name":"JB'S TIRE & REPAIR SVC.","address":"817 S. BYRON BUTLER","city":"PERRY","state":"FL","zip":"32347","phone":"8505842400","ytdComp":5175.95,"lat":30.111905,"lon":-83.589498,"accuracy":0.85,"accuracyType":"rooftop"},{"num":200698,"salesman":"Tiffany","active":"Inactive","name":"YARBROUGH TIRE CO.  INC.","address":"602 W. HAMPTON SPRINGS AVE.","city":"PERRY","state":"FL","zip":"32347","phone":"8505847554","ytdComp":0.0,"lat":30.11145,"lon":-83.587592,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200557,"salesman":"House","active":"Inactive","name":"AIMTRAC / CORPORATE","address":"321 FULLINGTON AVE","city":"PINEHURST","state":"GA","zip":"31070","phone":"888-861-0022","ytdComp":0.0,"lat":32.194954,"lon":-83.761617,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200563,"salesman":"House","active":"Inactive","name":"AIMTRAC / PINEHURST","address":"766 PINE AVE","city":"PINEHURST","state":"GA","zip":"31070","phone":"229-645-3331","ytdComp":0.0,"lat":32.198459,"lon":-83.761524,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200572,"salesman":"House","active":"Inactive","name":"PINEHURST PEANUT & GRAIN","address":"PO BOX 149","city":"PINEHURST","state":"GA","zip":"31070","phone":"229-645-3373","ytdComp":0.0,"lat":32.182652,"lon":-83.778422,"accuracy":1.0,"accuracyType":"place"},{"num":200574,"salesman":"House","active":"Inactive","name":"SOUTHERN STATES COOP  INC.","address":"265 HASLAM AVE","city":"PINEHURST","state":"GA","zip":"31070","phone":"229-645-3335","ytdComp":0.0,"lat":32.193393,"lon":-83.761357,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000028,"salesman":"Tiffany","active":"Active","name":"SUNPOINT TIRES & ROAD SERVICE","address":"3706 US-92","city":"PLANT CITY","state":"FL","zip":"33566","phone":"8632728823","ytdComp":0.0,"lat":28.024607,"lon":-82.070222,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200452,"salesman":"House","active":"Active","name":"TROY'S PAINT & BODY","address":"591 US HWY 82 NW","city":"POULAN","state":"GA","zip":"31781","phone":"2297760221","ytdComp":0.0,"lat":31.524769,"lon":-83.791666,"accuracy":0.86,"accuracyType":"nearest_rooftop_match"},{"num":200504,"salesman":"House","active":"Active","name":"PLATINUM RECOVERY SERVICES LLC","address":"407 N HUNTON ST","city":"POULAN","state":"GA","zip":"31781","phone":"2293952577","ytdComp":0.0,"lat":31.51681,"lon":-83.787359,"accuracy":0.76,"accuracyType":"nearest_rooftop_match"},{"num":200716,"salesman":"House","active":"Inactive","name":"POULAN DIESEL LLC","address":"407 HUNTON ST NW","city":"POULAN","state":"GA","zip":"31781","phone":"2298213586","ytdComp":0.0,"lat":31.51681,"lon":-83.787359,"accuracy":0.76,"accuracyType":"nearest_rooftop_match"},{"num":2,"salesman":"House","active":"Inactive","name":"quincy tire","address":"905 w. jefferson street","city":"quincy","state":"FL","zip":"32351","phone":"8506276050","ytdComp":0.0,"lat":30.587848,"lon":-84.586965,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200585,"salesman":"House","active":"Active","name":"QUINCY TIRE AND RECAPPING","address":"905 W. JEFFERSON","city":"QUINCY","state":"FL","zip":"32351","phone":"8506276050","ytdComp":23543.82,"lat":30.587672,"lon":-84.586963,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200599,"salesman":"House","active":"Active","name":"W&L TIRE & WHEEL CO. INC.","address":"820 W. JEFFERSON ST","city":"QUINCY","state":"FL","zip":"32351","phone":"8506278830","ytdComp":2077.7,"lat":30.588109,"lon":-84.585873,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200625,"salesman":"House","active":"Active","name":"DON SIRMONS ALIGNMENT & BRAKE","address":"396 E. JEFFERSON","city":"QUINCY","state":"FL","zip":"32351","phone":"8506278415","ytdComp":0.0,"lat":30.58752,"lon":-84.56919,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200589,"salesman":"Larry","active":"Active","name":"QUINCY ALIGNMENT & BRAKE","address":"804 W JEFFERSON ST","city":"QUINCY","state":"FL","zip":"32351","phone":"8508752444","ytdComp":0.0,"lat":30.588397,"lon":-84.585787,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201005,"salesman":"Tiffany","active":"Active","name":"MIDWAY ENTERPRISE FL  LLC","address":"804 W. JEFFERSON ST","city":"QUINCY","state":"FL","zip":"32351","phone":"8508752444","ytdComp":201.98,"lat":30.588397,"lon":-84.585787,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101566,"salesman":"House","active":"Active","name":"HARVEY'S GARAGE & MUFFLER","address":"422 N WASHINGTON","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292634272","ytdComp":0.0,"lat":30.788151,"lon":-83.557442,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200139,"salesman":"House","active":"Inactive","name":"WILLIE'S TIRE SHOP","address":"108 S WARREN ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292510310","ytdComp":0.0,"lat":30.784768,"lon":-83.562677,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200320,"salesman":"House","active":"Inactive","name":"CASS BURCH CHRYSLER","address":"801 E SCREVEN ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292632277","ytdComp":0.0,"lat":30.784406,"lon":-83.5532,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200356,"salesman":"House","active":"Active","name":"NEELY'S SERVICE CENTER","address":"302 S COURT ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292634454","ytdComp":1341.36,"lat":30.783282,"lon":-83.560036,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200383,"salesman":"Tiffany","active":"Active","name":"WILLIAMS ALIGNMENT & TIRE","address":"205 E JOHNSON ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292634797","ytdComp":9484.61,"lat":30.783301,"lon":-83.558917,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200437,"salesman":"House","active":"Inactive","name":"SHADES OF GREEN IRRIGATION","address":"111850 HWY 84 W","city":"QUITMAN","state":"GA","zip":"31643","phone":"229-370-0082","ytdComp":0.0,"lat":30.785659,"lon":-83.581076,"accuracy":0.7,"accuracyType":"street_center"},{"num":200463,"salesman":"House","active":"Inactive","name":"COASTAL PLAINS FARMERS CO-OP","address":"501 OGLESBY ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"229-263-7564","ytdComp":0.0,"lat":30.793902,"lon":-83.554995,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200783,"salesman":"Car Dealer","active":"Active","name":"CASS BURCH CHEVROLET","address":"12000 HWY 84 EAST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292637561","ytdComp":3198.0,"lat":30.785266,"lon":-83.538765,"accuracy":0.91,"accuracyType":"nearest_rooftop_match"},{"num":200841,"salesman":"Car Dealer","active":"Active","name":"CASS BURCH CHRYSLER DODGE JEEP","address":"801 E SCREVEN ST","city":"QUITMAN","state":"GA","zip":"31643","phone":"2292632277","ytdComp":0.0,"lat":30.784406,"lon":-83.5532,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201023,"salesman":"House","active":"Active","name":"PEASE ON THE GO 24/7","address":"403 WATERSIDE DR.","city":"QUITMAN","state":"GA","zip":"31643","phone":"2295396995","ytdComp":304.04,"lat":30.800969,"lon":-83.556099,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200262,"salesman":"House","active":"Inactive","name":"BERRIEN MOBILE TIRE","address":"365 STRAWDER RD","city":"RAY CITY","state":"GA","zip":"31645","phone":"229-686-1719","ytdComp":0.0,"lat":31.032141,"lon":-83.270473,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200229,"salesman":"House","active":"Inactive","name":"SOUTHERN STATES INC","address":"PO BOX 26234","city":"RICHMOND","state":"VA","zip":"23260","phone":"804-251-1696","ytdComp":0.0,"lat":37.524246,"lon":-77.493157,"accuracy":1.0,"accuracyType":"place"},{"num":101025,"salesman":"House","active":"Active","name":"ROCHELLE TIRE","address":"1266 1ST AVE","city":"ROCHELLE","state":"GA","zip":"31079","phone":"2293657943","ytdComp":16033.06,"lat":31.95043,"lon":-83.452368,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101223,"salesman":"House","active":"Inactive","name":"M & M REPAIR","address":"S 3RD AVE & MILLS ST.","city":"ROCHELLE","state":"GA","zip":"31079","phone":"229-365-7511","ytdComp":0.0,"lat":31.948551,"lon":-83.457902,"accuracy":1.0,"accuracyType":"intersection"},{"num":101530,"salesman":"House","active":"Active","name":"MARTIN TIRE SERVICE","address":"736 1ST AVE","city":"ROCHELLE","state":"GA","zip":"31079","phone":"2294254377","ytdComp":5264.27,"lat":31.950507,"lon":-83.461573,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101588,"salesman":"House","active":"Active","name":"STEPHENS BROTHERS","address":"11070 HIGHWAY 280","city":"ROCHELLE","state":"GA","zip":"31079","phone":"2294251055","ytdComp":4453.22,"lat":31.957885,"lon":-83.430832,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200880,"salesman":"House","active":"Active","name":"R&R TIRE CO.","address":"2275 GA HWY 233","city":"ROCHELLE","state":"GA","zip":"31079","phone":"2298054245","ytdComp":3314.61,"lat":31.879069,"lon":-83.431242,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201068,"salesman":"House","active":"Active","name":"HALO TIRES (TIFTON WHS)","address":"1953 N. WARREN RD","city":"SAN JACINTO","state":"CA","zip":"92582","phone":"9513941115","ytdComp":0.0,"lat":33.8201,"lon":-117.030384,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200808,"salesman":"House","active":"Inactive","name":"J&P SERVICE CENTER","address":"125 HWY 82","city":"SASSER","state":"GA","zip":"39885","phone":"229-698-2590","ytdComp":0.0,"lat":31.719268,"lon":-84.3487,"accuracy":0.93,"accuracyType":"range_interpolation"},{"num":200373,"salesman":"House","active":"Active","name":"PEERLESS MANUFACTURING CO","address":"2894 US 82 HWY E","city":"SHELLMAN","state":"GA","zip":"39886","phone":"2296795353","ytdComp":0.0,"lat":31.772821,"lon":-84.608831,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200272,"salesman":"House","active":"Inactive","name":"HAL ISRAEL FARM","address":"2285 GA HWY 308","city":"SMITHVILLE","state":"GA","zip":"31787","phone":"229-846-6477","ytdComp":0.0,"lat":31.925003,"lon":-84.279963,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200279,"salesman":"House","active":"Inactive","name":"ISRAEL FARM SUPPLY INC","address":"1979 GA HWY 308","city":"SMITHVILLE","state":"GA","zip":"31787","phone":"229-846-6655","ytdComp":0.0,"lat":31.929271,"lon":-84.299533,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200564,"salesman":"House","active":"Inactive","name":"AIMTRAC / SMITHVILLE","address":"2450 GA HWY 308","city":"SMITHVILLE","state":"GA","zip":"31787","phone":"229-924-3671","ytdComp":0.0,"lat":31.923444,"lon":-84.266329,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200685,"salesman":"House","active":"Inactive","name":"PAPA JOE'S TIRE","address":"104 STANTON DR.","city":"SMITHVILLE","state":"GA","zip":"31787","phone":"229-846-4402","ytdComp":0.0,"lat":31.903042,"lon":-84.253623,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200646,"salesman":"House","active":"Inactive","name":"BULLARD'S PERFORMANCE & AUTO.","address":"105 S GOODMAN ST","city":"SPARKS","state":"GA","zip":"31647","phone":"229-350-5040","ytdComp":0.0,"lat":31.166275,"lon":-83.437385,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101327,"salesman":"House","active":"Inactive","name":"NEW HOLLAND TIRE","address":"2982 US 82","city":"SUMNER","state":"GA","zip":"31789","phone":"229-776-4629","ytdComp":0.0,"lat":31.506842,"lon":-83.733008,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":100620,"salesman":"House","active":"Inactive","name":"ALLENS TIRE SERVICE","address":"1040 BUSSEY ROAD","city":"SYCAMORE","state":"GA","zip":"31790","phone":"229-567-3390","ytdComp":0.0,"lat":31.679679,"lon":-83.623642,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101200,"salesman":"House","active":"Inactive","name":"BUD'S TIRE SERVICE","address":"2779 HWY 41 S","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2295673138","ytdComp":0.0,"lat":31.669947,"lon":-83.634785,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101305,"salesman":"House","active":"Active","name":"SYCAMORE SALES & SALVAGE LLC","address":"3391 GA HWY 32 E","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2295672005","ytdComp":527.22,"lat":31.659565,"lon":-83.578823,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200891,"salesman":"Larry","active":"Active","name":"EJH WRECKER & TIRE SERVICE","address":"2779 US HWY 41","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2295663334","ytdComp":53788.54,"lat":31.669947,"lon":-83.634785,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200971,"salesman":"Larry","active":"Active","name":"ALLEN'S TIRE","address":"1040 BUSSEY RD.","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2295673390","ytdComp":37566.07,"lat":31.679679,"lon":-83.623642,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200903,"salesman":"House","active":"Inactive","name":"SMITH  LARRY (SALESMAN ACCT.)","address":"295 FOWLER SMITH LANE","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2293645833","ytdComp":0.0,"lat":31.610704,"lon":-83.55598,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200957,"salesman":"House","active":"Active","name":"DNR REPAIR SHOP","address":"1647 HWY 41","city":"SYCAMORE","state":"GA","zip":"31790","phone":"2292885336","ytdComp":0.0,"lat":31.685162,"lon":-83.642709,"accuracy":0.93,"accuracyType":"rooftop"},{"num":101537,"salesman":"House","active":"Active","name":"GILES TOWING & RECOVERY","address":"101 PENDLEY DR","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297763936","ytdComp":0.0,"lat":31.540444,"lon":-83.824099,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200140,"salesman":"House","active":"Inactive","name":"ALBANY TRACTOR","address":"HWY 33 S","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-5565","ytdComp":0.0,"lat":31.558023,"lon":-84.134173,"accuracy":0.54,"accuracyType":"street_center"},{"num":200186,"salesman":"House","active":"Inactive","name":"LANGDALE CHEVROLET","address":"1008 W FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-3473","ytdComp":0.0,"lat":31.533161,"lon":-83.852571,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200208,"salesman":"House","active":"Inactive","name":"PATTERSON TIRE CO","address":"200 N MAIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-3315","ytdComp":0.0,"lat":31.527791,"lon":-83.835192,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200216,"salesman":"House","active":"Active","name":"S & S TIRE SERVICE","address":"112 RICHIE DR","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2298836703","ytdComp":0.0,"lat":31.56381,"lon":-83.995143,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200232,"salesman":"Car Dealer","active":"Active","name":"SUNBELT FORD INC","address":"1002 W FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297767691","ytdComp":0.0,"lat":31.531934,"lon":-83.850424,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200299,"salesman":"House","active":"Inactive","name":"CLOSED","address":"401 W FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-0400","ytdComp":0.0,"lat":31.530457,"lon":-83.841879,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200338,"salesman":"House","active":"Active","name":"GIDDENS AUTO","address":"301 W FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297763373","ytdComp":0.0,"lat":31.530435,"lon":-83.840641,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200394,"salesman":"House","active":"Inactive","name":"POWELL ROYCE AUTO SERVICE","address":"401 E POPE ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-5218","ytdComp":0.0,"lat":31.529228,"lon":-83.831979,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200395,"salesman":"House","active":"Inactive","name":"SYLCREE INDUSTRIAL TIRE","address":"4605 SYLVESTER RD","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-438-1084","ytdComp":0.0,"lat":31.603336,"lon":-83.850054,"accuracy":0.5,"accuracyType":"place"},{"num":200681,"salesman":"House","active":"Active","name":"CITY OF SYLVESTER","address":"PO BOX 370","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297768504","ytdComp":1840.24,"lat":31.603336,"lon":-83.850054,"accuracy":1.0,"accuracyType":"place"},{"num":200692,"salesman":"House","active":"Inactive","name":"RAINEY USED CARS","address":"402 E. POPE ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"229-776-0400","ytdComp":0.0,"lat":31.528615,"lon":-83.832003,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200715,"salesman":"House","active":"Active","name":"R&M AUTO TRUCKING  INC","address":"638 REDROCK RD","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2292062102","ytdComp":1507.6,"lat":31.578654,"lon":-83.88765,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101436,"salesman":"Larry","active":"Active","name":"ED'S TIRE","address":"202 E FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297766952","ytdComp":17706.55,"lat":31.530234,"lon":-83.835118,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200220,"salesman":"Larry","active":"Active","name":"SINGLETARY & SON TIRE CO","address":"2528 US HWY 82 W","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297765535","ytdComp":44777.53,"lat":31.539741,"lon":-83.86788,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200500,"salesman":"Larry","active":"Active","name":"SHELL RAPID LUBE (SYLVESTER)","address":"1001 WEST FRANKLIN STREET","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297770932","ytdComp":16.64,"lat":31.530997,"lon":-83.852205,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200972,"salesman":"Larry","active":"Active","name":"ERIC'S TIRE OF SYLVESTER","address":"210 SOUTH MAIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2298212000","ytdComp":5427.25,"lat":31.524399,"lon":-83.836114,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200860,"salesman":"House","active":"Inactive","name":"SUNBELT FORD INC. (AMI ACCT)","address":"1002 W FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297767691","ytdComp":0.0,"lat":31.531934,"lon":-83.850424,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000002,"salesman":"Larry","active":"Active","name":"EG AGRI PARTS LLC","address":"206 E FRONT ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"7867827277","ytdComp":0.0,"lat":31.526321,"lon":-83.836701,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":2000017,"salesman":"Larry","active":"Active","name":"JORDAN AUTOMOTIVE & TIRES","address":"602 EAST FRANKLIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2294396881","ytdComp":352.0,"lat":31.529984,"lon":-83.829201,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201029,"salesman":"House","active":"Active","name":"UNIVERSAL AUTO SOLUTIONS LLC","address":"1121 E. FRANKLIN ST.","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297827131","ytdComp":0.0,"lat":31.528217,"lon":-83.811345,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201052,"salesman":"House","active":"Active","name":"PRECISION DIESEL REPAIR LLC","address":"807 SOUTH MAIN ST","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2294727268","ytdComp":9400.03,"lat":31.515209,"lon":-83.835776,"accuracy":1.0,"accuracyType":"rooftop"},{"num":1999999,"salesman":"Car Dealer","active":"Active","name":"GRIFFIN CHEVROLET OF SYLVESTER","address":"1006 W FRANKLIN","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297763473","ytdComp":0.0,"lat":31.53236,"lon":-83.851789,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000027,"salesman":"Car Dealer","active":"Active","name":"FORD SYLVESTER","address":"1002 W. FRANKLIN ST.","city":"SYLVESTER","state":"GA","zip":"31791","phone":"2297767691","ytdComp":0.0,"lat":31.531934,"lon":-83.850424,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200514,"salesman":"House","active":"Inactive","name":"GCR TIRES AND SERVICE","address":"4702 CAPITAL CIRCLE NW","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8502223130","ytdComp":0.0,"lat":30.516688,"lon":-84.366946,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200515,"salesman":"Tiffany","active":"Active","name":"NEECE TIRE & AUTO","address":"4792 BLOUNTSTOWN HWY","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"8505744100","ytdComp":0.0,"lat":30.436008,"lon":-84.358415,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200581,"salesman":"House","active":"Inactive","name":"AUTO ALLEY","address":"5019 W. THARPE STREET","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"850-576-9835","ytdComp":0.0,"lat":30.462832,"lon":-84.362686,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200583,"salesman":"House","active":"Inactive","name":"SUN TIRE","address":"2400 MILLCREEK LANE","city":"TALLAHASSEE","state":"FL","zip":"32308","phone":"850-553-9661","ytdComp":0.0,"lat":30.483888,"lon":-84.236384,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200737,"salesman":"House","active":"Inactive","name":"KAUFFMAN TIRE (BEECH RIDE CT.)","address":"3335 BEECH RIDE CT.","city":"TALLAHASSEE","state":"FL","zip":"32312","phone":"850-688-1196","ytdComp":0.0,"lat":30.56301,"lon":-84.218449,"accuracy":0.9,"accuracyType":"rooftop"},{"num":200738,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 817","address":"205 NORTH MAGNOLIA DR.","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8503295565","ytdComp":0.0,"lat":30.443378,"lon":-84.261039,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200739,"salesman":"House","active":"Inactive","name":"SUN TIRE OF NORTH MONROE","address":"2715 N. MONROE ST.","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"850-422-2024","ytdComp":0.0,"lat":30.479172,"lon":-84.300099,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200748,"salesman":"House","active":"Inactive","name":"ADVANTAGE COMMERCIAL TIRE","address":"4329 WEST PENSACOLA ST.","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"850-329-6238","ytdComp":0.0,"lat":30.435455,"lon":-84.35137,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200796,"salesman":"House","active":"Active","name":"TRUCK N CAR CONCEPTS","address":"3270 MAHAN DR","city":"TALLAHASSEE","state":"FL","zip":"32308","phone":"8506568800","ytdComp":0.0,"lat":30.462982,"lon":-84.221215,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200960,"salesman":"House","active":"Active","name":"BOULEVARD TIRE CENTER","address":"4702 CAPITAL CIRCLE","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8502223130","ytdComp":0.0,"lat":30.516688,"lon":-84.366946,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200979,"salesman":"House","active":"Active","name":"DISCOUNT TIRE & AUTO SHOP","address":"114 RIDGE RD","city":"TALLAHASSEE","state":"FL","zip":"32305","phone":"8505447234","ytdComp":1191.86,"lat":30.397894,"lon":-84.283229,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200981,"salesman":"Tiffany","active":"Active","name":"AUTO WORKS INTERNATIONAL","address":"2090 N. MONROE ST.","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"","ytdComp":0.0,"lat":30.467951,"lon":-84.287621,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200984,"salesman":"House","active":"Active","name":"SIRMONS BRAKE & ALIGNMENT","address":"5087 TENNESSEE CAPITAL BLVD","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8505743581","ytdComp":0.0,"lat":30.458894,"lon":-84.366962,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200987,"salesman":"House","active":"Inactive","name":"WESTSIDE TIRE & BRAKE (TEST)","address":"NO ORDERS UNTI APP REC'D","city":"TALLAHASSEE","state":"FL","zip":"32310","phone":"8055761151","ytdComp":0.0,"lat":30.441667,"lon":-84.407605,"accuracy":0.83,"accuracyType":"street_center"},{"num":200988,"salesman":"House","active":"Inactive","name":"REAL DEAL CUSTOM RIMS & TIRES","address":"1530 S MONROE ST","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8503296800","ytdComp":0.0,"lat":30.427631,"lon":-84.280998,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200989,"salesman":"House","active":"Active","name":"4S TIRE","address":"1471 CAPITAL CIRCLE N.W.","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8505977352","ytdComp":0.0,"lat":30.466457,"lon":-84.360664,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200990,"salesman":"House","active":"Active","name":"PATTON'S ALIGNMENT & BRAKE SVC","address":"2405 S. ADAMS ST","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8502220142","ytdComp":0.0,"lat":30.417865,"lon":-84.281938,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200991,"salesman":"House","active":"Active","name":"THE TIRE CENTRE OF FLORIDA LLC","address":"2620 S. MONROE","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8506714181","ytdComp":12243.98,"lat":30.413373,"lon":-84.28097,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200994,"salesman":"House","active":"Inactive","name":"MINCO AUTO TRUCK & ACCESSORIES","address":"200 N MAGNOLIA DR","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8506561919","ytdComp":0.0,"lat":30.443067,"lon":-84.262329,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200995,"salesman":"House","active":"Active","name":"BOB'S AUTO REPAIR & COLLISION","address":"2293 LAKE BRADFORD RD","city":"TALLAHASSEE","state":"FL","zip":"32310","phone":"8502249205","ytdComp":0.0,"lat":30.41559,"lon":-84.303631,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200997,"salesman":"House","active":"Inactive","name":"CAPITAL CITY IMPORTS","address":"4394 BLOUNTSTOWN HWY","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"8509421500","ytdComp":0.0,"lat":30.435815,"lon":-84.353509,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200998,"salesman":"House","active":"Inactive","name":"FRIENDLY AUTO SALES","address":"4900 BLOUNTSTOWN HWY","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"8505366080","ytdComp":0.0,"lat":30.435293,"lon":-84.359463,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200999,"salesman":"House","active":"Active","name":"USED CAR SUPERMARKET INC.","address":"3120 W. TENNESSEE ST.","city":"TALLAHASSEE","state":"FL","zip":"32304-1002","phone":"8505756702","ytdComp":0.0,"lat":30.458882,"lon":-84.349592,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201000,"salesman":"House","active":"Active","name":"JOSE TIRE & CAR SERVICE","address":"6320 BLOUNTSTOWN RD","city":"TALLAHASSEE","state":"FL","zip":"32310","phone":"8505677367","ytdComp":0.0,"lat":30.439444,"lon":-84.384932,"accuracy":0.85,"accuracyType":"rooftop"},{"num":201001,"salesman":"House","active":"Active","name":"CAPITAL HITCH SERVICE  INC","address":"7596 W. TENNESSEE ST.","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"8505758628","ytdComp":0.0,"lat":30.459983,"lon":-84.386407,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201004,"salesman":"House","active":"Inactive","name":"PORTERS AUTOMOTIVE  INC (TEST)","address":"415 CAPITAL CIRCLE SW","city":"TALLAHASSEE","state":"FL","zip":"32304","phone":"8502242886","ytdComp":0.0,"lat":30.451922,"lon":-84.358867,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201006,"salesman":"House","active":"Active","name":"FURRIN AUTO ALLEY","address":"5019 W. THARPE ST.","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8502226864","ytdComp":0.0,"lat":30.462832,"lon":-84.362686,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201010,"salesman":"House","active":"Active","name":"TIRE TOWN AUTOMOTIVE CENTER","address":"3206 APALACHEE PKWY.","city":"TALLAHASSEE","state":"FL","zip":"32311","phone":"8506568473","ytdComp":0.0,"lat":30.428149,"lon":-84.223201,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201011,"salesman":"House","active":"Active","name":"JIM WELLS TIRE CENTER  INC.","address":"1853 THOMASVILLE RD.","city":"TALLAHASSEE","state":"FL","zip":"32303","phone":"8502225305","ytdComp":0.0,"lat":30.464871,"lon":-84.269425,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201218,"salesman":"House","active":"Active","name":"FIRESTONE STORE #003557","address":"501 N MONROE ST","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8502220190","ytdComp":0.0,"lat":30.447271,"lon":-84.280207,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201219,"salesman":"House","active":"Active","name":"FIRESTONE STORE #024988","address":"2211 APALACHEE PKWY","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8508770184","ytdComp":0.0,"lat":30.429645,"lon":-84.244129,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201220,"salesman":"House","active":"Active","name":"TIRES PLUS #550655","address":"2800 APALACHEE PKWY","city":"TALLAHASSEE","state":"FL","zip":"32301","phone":"8508774091","ytdComp":0.0,"lat":30.427899,"lon":-84.23667,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201221,"salesman":"House","active":"Active","name":"TIRES PLUS #550728","address":"1883 CAPITAL CIR NE","city":"TALLAHASSEE","state":"FL","zip":"32308","phone":"8509427400","ytdComp":0.0,"lat":30.471082,"lon":-84.230879,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200942,"salesman":"House","active":"Active","name":"RNR TIRE EXPRESS","address":"13922 MONROES BUSINESS PARK","city":"TAMPA","state":"FL","zip":"33635","phone":"8139779800","ytdComp":0.0,"lat":28.037037,"lon":-82.64531,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201072,"salesman":"House","active":"Active","name":"FALCON TIRE CENTER","address":"7219 E BROADWAY AVE","city":"TAMPA","state":"FL","zip":"33619","phone":"8136359408","ytdComp":0.0,"lat":27.966343,"lon":-82.374063,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200251,"salesman":"House","active":"Inactive","name":"319 FLEET SERVICES LLC","address":"1309 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"229-225-9568","ytdComp":0.0,"lat":30.85219,"lon":-83.95988,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200258,"salesman":"House","active":"Active","name":"ALLIGOOD TIRE CO","address":"1305 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31799","phone":"2292260811","ytdComp":0.0,"lat":30.851678,"lon":-83.960549,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200264,"salesman":"House","active":"Inactive","name":"DAN BAIN AUTO SERVICE CENTER","address":"201 REMINGTON AVE","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-226-3603","ytdComp":0.0,"lat":30.837765,"lon":-83.977755,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200277,"salesman":"House","active":"Active","name":"IMPORT SERVICE & SALES","address":"4816 GA HWY 202","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292269844","ytdComp":4578.34,"lat":30.907063,"lon":-83.973884,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200285,"salesman":"Tiffany","active":"Inactive","name":"PRECISION AUTOCRAFT INC","address":"1017 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292280734","ytdComp":0.0,"lat":30.847477,"lon":-83.96693,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200291,"salesman":"Tiffany","active":"Active","name":"SINGLETARY TIRE PROS","address":"401 N MADISON","city":"THOMASVILLE","state":"GA","zip":"31799","phone":"2292262842","ytdComp":356.39,"lat":30.839825,"lon":-83.984374,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200293,"salesman":"Tiffany","active":"Active","name":"THOMASVILLE TIRE DEPT.","address":"1311 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292280260","ytdComp":20583.08,"lat":30.852467,"lon":-83.960239,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200301,"salesman":"House","active":"Inactive","name":"ADVANCED AUTOMOTIVE","address":"107 COMMERCE PARKDR","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"229-227-6010","ytdComp":0.0,"lat":30.863032,"lon":-83.941399,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200333,"salesman":"House","active":"Inactive","name":"FIRST VEHICLE SERVICES","address":"119 FAIRBANKS AVE","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-228-4545","ytdComp":0.0,"lat":30.831027,"lon":-83.961124,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200336,"salesman":"Car Dealer","active":"Active","name":"FLOWERS AUTOMOTIVE NISSAN","address":"1630 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292261106","ytdComp":0.0,"lat":30.85754,"lon":-83.950352,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200358,"salesman":"House","active":"Inactive","name":"PONDER'S AUTOMOTIVE INC","address":"1052 W JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292285779","ytdComp":4580.14,"lat":30.824611,"lon":-83.992073,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200359,"salesman":"House","active":"Inactive","name":"PRICE POINT MOTORS","address":"14638 HWY 19 S","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"229-551-0313","ytdComp":0.0,"lat":30.849516,"lon":-83.944871,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200368,"salesman":"House","active":"Inactive","name":"RUSSELL TIRE COMPANY","address":"1030 E JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-226-8475","ytdComp":0.0,"lat":30.84771,"lon":-83.964961,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200378,"salesman":"House","active":"Inactive","name":"T & T AUTO LUBE","address":"1441 REMINGTON AVE","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292276087","ytdComp":0.0,"lat":30.852635,"lon":-83.948923,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200385,"salesman":"Tiffany","active":"Active","name":"WILLIAMS AUTOMOTIVE","address":"1038 W JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292246328","ytdComp":0.0,"lat":30.824976,"lon":-83.991451,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200410,"salesman":"House","active":"Active","name":"EZDEALIN WHEELS AND TIRES","address":"114 COMMERCE PARK DR","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"8772472230","ytdComp":57409.96,"lat":30.862286,"lon":-83.941231,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200434,"salesman":"Tiffany","active":"Active","name":"CAPITAL AUTO PARTS  INC","address":"5042 GA HWY 202","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2292267878","ytdComp":0.0,"lat":30.909443,"lon":-83.972833,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200519,"salesman":"House","active":"Inactive","name":"RON GOODSON AUTOMOTIVE","address":"510 CAMPBELL ST","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-236-3190","ytdComp":0.0,"lat":30.832103,"lon":-83.992796,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200520,"salesman":"House","active":"Active","name":"OTTS AUTO DOCTOR","address":"10821 US HWY 84","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2292268100","ytdComp":0.0,"lat":30.83425,"lon":-83.933306,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200521,"salesman":"House","active":"Inactive","name":"EUROPEAN EDGE  LLC","address":"311 COMMERCIAL DRIVE","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292363343","ytdComp":0.0,"lat":30.851985,"lon":-83.941085,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200526,"salesman":"House","active":"Inactive","name":"GERMAN IMPORT SERVICE","address":"14262 US HWY 19 SOUTH","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2292262010","ytdComp":0.0,"lat":30.84452,"lon":-83.94411,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200528,"salesman":"House","active":"Inactive","name":"JACK'S SERVICE CENTER","address":"2980 E PINETREE","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292262150","ytdComp":0.0,"lat":30.855507,"lon":-83.952176,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200529,"salesman":"House","active":"Inactive","name":"STEPHENS MOTOR CARS INC","address":"826 E PINETREE BLVD","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-228-0322","ytdComp":0.0,"lat":30.832251,"lon":-83.947348,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200533,"salesman":"Car Dealer","active":"Active","name":"FLOWERS IMPORTS LLC (HONDA)","address":"1610 E. JACKSON ST","city":"THOMASVILLE","state":"GA","zip":"31758","phone":"2292251144","ytdComp":0.0,"lat":30.85649,"lon":-83.951152,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200561,"salesman":"House","active":"Inactive","name":"STALLINGS AUTOMOTIVE","address":"101 COMMERCE PARK DR","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"229-228-9100","ytdComp":0.0,"lat":30.863006,"lon":-83.942238,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200607,"salesman":"House","active":"Active","name":"B AND B SERVICE CENTER  INC.","address":"510 N. BLVD.","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292362886","ytdComp":1817.38,"lat":30.848773,"lon":-83.992189,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200615,"salesman":"House","active":"Inactive","name":"VICTORY TRANSPORT CO.","address":"2701 US HWY 84 BYPASS WEST","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2294132349","ytdComp":0.0,"lat":30.870203,"lon":-83.998971,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200617,"salesman":"House","active":"Inactive","name":"AUTO DETAIL CENTER","address":"366 COMMERCIAL DR","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292262850","ytdComp":0.0,"lat":30.851904,"lon":-83.94055,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200641,"salesman":"House","active":"Inactive","name":"AUTO AIR OF THOMASVILLE","address":"826 E. PINETREE BLVD","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292280322","ytdComp":0.0,"lat":30.832251,"lon":-83.947348,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200682,"salesman":"Car Dealer","active":"Active","name":"THOMASVILLE FORD LINCOLN","address":"1515 E. JACKSON ST.","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292265133","ytdComp":0.0,"lat":30.855796,"lon":-83.954998,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200736,"salesman":"House","active":"Inactive","name":"MAVIS TIRE (THOMASVILLE)","address":"14005 HIGHWAY 19 SOUTH","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"229-379-3041","ytdComp":0.0,"lat":30.840741,"lon":-83.944907,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200784,"salesman":"Car Dealer","active":"Active","name":"SPENCE CHEVROLET BUICK GMC","address":"11646 US 319 N","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2295164639","ytdComp":0.0,"lat":30.86352,"lon":-83.943055,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200786,"salesman":"House","active":"Inactive","name":"THOMASVILLE CHRYSLER DODGE","address":"101 COMMERCE PARK DR","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"877-360-6038","ytdComp":0.0,"lat":30.863006,"lon":-83.942238,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200861,"salesman":"Car Dealer","active":"Active","name":"THOMASVILLE CHRYSLER DODGE","address":"101 COMMERCE PARD DR","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"8773606038","ytdComp":0.0,"lat":30.863006,"lon":-83.942238,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200862,"salesman":"House","active":"Inactive","name":"THOMASVILLE FORD (AMI ACCT)","address":"1515 E. JACKSON ST.","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292265133","ytdComp":0.0,"lat":30.855796,"lon":-83.954998,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200302,"salesman":"Larry","active":"Active","name":"AG PRO","address":"12793 HWY 19 S","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2292264881","ytdComp":0.0,"lat":30.826388,"lon":-83.943028,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200921,"salesman":"Tiffany","active":"Active","name":"AUTO DOCTOR DIESEL & REPAIR","address":"10821 US HWY 84 E","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2292268100","ytdComp":242.46,"lat":30.83425,"lon":-83.933306,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200970,"salesman":"Tiffany","active":"Inactive","name":"BIG PINE REPAIR","address":"12390 US HWY 84 EAST","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2294032990","ytdComp":0.0,"lat":30.826597,"lon":-83.910699,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201012,"salesman":"House","active":"Active","name":"AG PRO FUEL","address":"201 REMMINGTON AVE.","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2299771313","ytdComp":0.0,"lat":30.837765,"lon":-83.977755,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201026,"salesman":"House","active":"Inactive","name":"THOMASVILLE TOYOTA (TEST ACCT)","address":"14724 US HIGHWAY 19 S","city":"THOMASVILLE","state":"GA","zip":"31757","phone":"2292280555","ytdComp":0.0,"lat":30.85021,"lon":-83.94482,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201103,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 663","address":"14005 HIGHWAY 19 SOUTH","city":"THOMASVILLE","state":"GA","zip":"31792","phone":"2293528493","ytdComp":0.0,"lat":30.840741,"lon":-83.944907,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201106,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 2078","address":"2226 HARRISON RD","city":"THOMSON","state":"GA","zip":"30824","phone":"7623097990","ytdComp":0.0,"lat":33.482871,"lon":-82.501891,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":100540,"salesman":"House","active":"Inactive","name":"TIRE CITY SALES & AUTO SERVICE","address":"1181 S. MAIN ST.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-386-1388","ytdComp":0.0,"lat":31.445096,"lon":-83.506751,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100546,"salesman":"House","active":"Inactive","name":"WHITES AUTO","address":"1395-A HWY 82 EAST","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-9967","ytdComp":0.0,"lat":31.444073,"lon":-83.478025,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100696,"salesman":"House","active":"Inactive","name":"DO-RITE TIRE AND AUTO","address":"2113 SOUTH CENTRAL AVE.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-386-8268","ytdComp":0.0,"lat":31.441189,"lon":-83.515513,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100715,"salesman":"House","active":"Inactive","name":"T&S AUTOMOTIVE GROUP","address":"101 RIVER BEND LANE.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-326-1316","ytdComp":0.0,"lat":31.478674,"lon":-83.580492,"accuracy":0.99,"accuracyType":"rooftop"},{"num":100954,"salesman":"House","active":"Inactive","name":"HALL TRUCKING GROUP  LLC","address":"3277 HALL WHITLEY RD","city":"TIFTON","state":"GA","zip":"31794","phone":"229-392-3975","ytdComp":0.0,"lat":31.367348,"lon":-83.385935,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101103,"salesman":"House","active":"Inactive","name":"DELTORO TIRE  LLC","address":"1221 MOORE HWY","city":"TIFTON","state":"GA","zip":"31794","phone":"229-445-3583","ytdComp":0.0,"lat":31.469291,"lon":-83.519295,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101152,"salesman":"House","active":"Active","name":"BROOKS BODY SHOP","address":"511 2ND ST E","city":"TIFTON","state":"GA","zip":"31794","phone":"2293861800","ytdComp":1464.0,"lat":31.452941,"lon":-83.503649,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101252,"salesman":"House","active":"Active","name":"DNA DIESEL & AUTOMOTIVE REPAIR","address":"605 OLD OMEGA RD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293820207","ytdComp":212.16,"lat":31.440833,"lon":-83.530334,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101298,"salesman":"Car Dealer","active":"Active","name":"TENNESON NISSAN","address":"535 OLD OMEGA RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293827777","ytdComp":5423.25,"lat":31.441115,"lon":-83.524944,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101322,"salesman":"House","active":"Active","name":"GRIMES AUTO SERVICE","address":"305 RIDGE AVE S","city":"TIFTON","state":"GA","zip":"31794","phone":"2293821501","ytdComp":2178.41,"lat":31.451744,"lon":-83.515632,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101325,"salesman":"House","active":"Inactive","name":"HOLT'S AUTO & TIRE SERVICES","address":"310 RIDGE AVE S","city":"TIFTON","state":"GA","zip":"31794","phone":"229-386-5500","ytdComp":0.0,"lat":31.451971,"lon":-83.516301,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101328,"salesman":"House","active":"Inactive","name":"TEN SPEED TRUCK SERVICES","address":"4445 UNION RD","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-5000","ytdComp":0.0,"lat":31.42765,"lon":-83.516859,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101330,"salesman":"House","active":"Inactive","name":"MICHAEL MOORE'S PAINT AND BODY","address":"2814 S CARPENTER RD","city":"TIFTON","state":"GA","zip":"31793","phone":"229-391-9318","ytdComp":0.0,"lat":31.438858,"lon":-83.547564,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101374,"salesman":"Car Dealer","active":"Active","name":"PRINCE CHEVY-OLDS  INC","address":"1410 US HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"2293864050","ytdComp":2160.83,"lat":31.447931,"lon":-83.536014,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200001,"salesman":"House","active":"Active","name":"CASH CUSTOMER","address":"DO NOT USE!","city":"TIFTON","state":"GA","zip":"31794","phone":"","ytdComp":0.0,"lat":31.423595,"lon":-83.545769,"accuracy":0.87,"accuracyType":"street_center"},{"num":200254,"salesman":"House","active":"Inactive","name":"LAIRSEY'S AUTO SERVICE CENTER","address":"3116 HWY 41 S","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-4919","ytdComp":0.0,"lat":31.426278,"lon":-83.49668,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200331,"salesman":"House","active":"Active","name":"BUDGET CAR SALES","address":"3302 HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"2293880020","ytdComp":6278.07,"lat":31.444871,"lon":-83.547336,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200374,"salesman":"House","active":"Inactive","name":"CLOSED-GO DADDY TIRE","address":"2205 LESLIE LOCKE RD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293965932","ytdComp":0.0,"lat":31.445513,"lon":-83.549511,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200407,"salesman":"House","active":"Active","name":"HOLLOWAY TRUCK & TRAILER REPAI","address":"80 TRI COUNTY RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293878795","ytdComp":6114.52,"lat":31.437901,"lon":-83.476747,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200435,"salesman":"House","active":"Inactive","name":"ATLANTIC & SOUTHERN EQPT CO","address":"4186 W. US HWY 82","city":"TIFTON","state":"GA","zip":"31793","phone":"2293398010","ytdComp":0.0,"lat":31.459957,"lon":-83.609165,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200440,"salesman":"House","active":"Inactive","name":"1ST PRODUCTS","address":"164 OAKRIDGE CHURCH RD","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-4769","ytdComp":0.0,"lat":31.401267,"lon":-83.507477,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200455,"salesman":"Car Dealer","active":"Active","name":"PRINCE HONDA","address":"1501 US HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"2293861400","ytdComp":338.36,"lat":31.448131,"lon":-83.538693,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200456,"salesman":"Car Dealer","active":"Active","name":"PRINCE TOYOTA","address":"2013 US HWY 82 WEST","city":"TIFTON","state":"GA","zip":"31793","phone":"2293864052","ytdComp":6391.76,"lat":31.44542,"lon":-83.545819,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200476,"salesman":"House","active":"Inactive","name":"BANNER GRAIN & PEANUT","address":"928 E GOLDEN RD","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-0415","ytdComp":0.0,"lat":31.429862,"lon":-83.499069,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200477,"salesman":"House","active":"Inactive","name":"BROWNLEE ENTERPRISES","address":"928 E GOLDEN RD","city":"TIFTON","state":"GA","zip":"31794","phone":"229-382-0415","ytdComp":0.0,"lat":31.429862,"lon":-83.499069,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200505,"salesman":"House","active":"Inactive","name":"PATRICK TRACTOR CO. INC.","address":"2718 MARSHALL DRIVE","city":"TIFTON","state":"GA","zip":"31793","phone":"229-382-3626","ytdComp":0.0,"lat":31.440496,"lon":-83.551858,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200524,"salesman":"Car Dealer","active":"Inactive","name":"GRIFFIN FORD LINCOLN  INC","address":"PO BOX 765","city":"TIFTON","state":"GA","zip":"31793","phone":"2293821300","ytdComp":0.0,"lat":31.462512,"lon":-83.593543,"accuracy":1.0,"accuracyType":"place"},{"num":200527,"salesman":"Car Dealer","active":"Inactive","name":"GRIFFIN CHRY/DOD/JEEP/RAM","address":"505 7TH ST. WEST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293820440","ytdComp":1881.08,"lat":31.451748,"lon":-83.520697,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200545,"salesman":"House","active":"Inactive","name":"CLOSED","address":"65 ETHRIDGE RD","city":"TIFTON","state":"GA","zip":"31793","phone":"229-402-9571","ytdComp":0.0,"lat":31.562862,"lon":-83.609868,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200548,"salesman":"House","active":"Inactive","name":"THE REPAIR SHOP OF TIFTON  LLC","address":"190 VICTORY DRIVE","city":"TIFTON","state":"GA","zip":"31794","phone":"229-387-1338","ytdComp":0.0,"lat":31.453965,"lon":-83.518983,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200554,"salesman":"Car Dealer","active":"Active","name":"HONDA OF SOUTH GEORGIA","address":"1025 CAMDEN WAY","city":"TIFTON","state":"GA","zip":"31794","phone":"2293964050","ytdComp":0.0,"lat":31.456651,"lon":-83.528678,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200556,"salesman":"House","active":"Inactive","name":"TIFTON AUTOMOTIVE & CAR AUDIO","address":"2302 S MAIN ST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293648209","ytdComp":0.0,"lat":31.441567,"lon":-83.506083,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200567,"salesman":"House","active":"Inactive","name":"AIMTRAC / TIFTON","address":"1825 US HWY 82 WEST","city":"TIFTON","state":"GA","zip":"31793","phone":"229-472-5732","ytdComp":0.0,"lat":31.446634,"lon":-83.54238,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200580,"salesman":"Car Dealer","active":"Active","name":"JEFF FENDER BUICK  GMC, CAD.","address":"736 2ND ST W.","city":"TIFTON","state":"GA","zip":"31793","phone":"2293861985","ytdComp":1980.46,"lat":31.457933,"lon":-83.522907,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200593,"salesman":"House","active":"Inactive","name":"CLOSED","address":"606 48TH ST","city":"TIFTON","state":"GA","zip":"31794","phone":"229-520-8546","ytdComp":0.0,"lat":31.490165,"lon":-83.497907,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200603,"salesman":"House","active":"Inactive","name":"PREMIER WHOLESALE TIRE COMPANY","address":"5705 N. PALAFAX ST","city":"TIFTON","state":"GA","zip":"32503","phone":"850-474-4999","ytdComp":0.0,"lat":30.475231,"lon":-87.245605,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200638,"salesman":"House","active":"Inactive","name":"HOLLOWAY TRUCK & OFFROAD ACCS.","address":"104 E. 3RD ST.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-472-1384","ytdComp":0.0,"lat":31.452132,"lon":-83.509568,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200669,"salesman":"House","active":"Inactive","name":"LENCHOS TIRE","address":"1021 FERRYLAKE DR.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-646-6518","ytdComp":0.0,"lat":31.455536,"lon":-83.494794,"accuracy":0.9,"accuracyType":"rooftop"},{"num":200686,"salesman":"House","active":"Inactive","name":"T&C AUTO REPAIR","address":"1508 REBEL ROAD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293862225","ytdComp":0.0,"lat":31.445513,"lon":-83.54098,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200693,"salesman":"House","active":"Inactive","name":"MOORE'S BODY SHOP","address":"756 TIFTON-ELDORADO RD.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-392-9115","ytdComp":0.0,"lat":31.406652,"lon":-83.474183,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200703,"salesman":"House","active":"Inactive","name":"STAR AUTO SALES & REPAIR","address":"510 RIDGE AVE. S.","city":"TIFTON","state":"GA","zip":"31794","phone":"229-396-5451","ytdComp":0.0,"lat":31.45122,"lon":-83.516569,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200709,"salesman":"House","active":"Active","name":"ASHLEY'S AUTOMOTIVE REPAIR","address":"310 S. RIDGE AVE.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293964640","ytdComp":3261.65,"lat":31.451971,"lon":-83.516301,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200713,"salesman":"House","active":"Inactive","name":"BALTES TIRES  LLC","address":"1021 FERRYLAKE RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2294299924","ytdComp":0.0,"lat":31.455536,"lon":-83.494794,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200718,"salesman":"House","active":"Active","name":"JOBBER ACCT (TIFTON)","address":"0","city":"TIFTON","state":"GA","zip":"31794","phone":"","ytdComp":1142.8,"lat":31.215302,"lon":-83.270126,"accuracy":0.9,"accuracyType":"street_center"},{"num":200728,"salesman":"House","active":"Inactive","name":"ADVANTAGE COMMERCIAL TIRE","address":"287 SOUTHWELL BLVD","city":"TIFTON","state":"GA","zip":"31794","phone":"2292072232","ytdComp":0.0,"lat":31.417354,"lon":-83.493885,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200751,"salesman":"House","active":"Inactive","name":"H&T TIRE & ALIGNMENT","address":"1508 REBEL RD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293964946","ytdComp":0.0,"lat":31.445513,"lon":-83.54098,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200759,"salesman":"House","active":"Active","name":"AADCO","address":"708 W. 3RD ST.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293828080","ytdComp":466.48,"lat":31.452431,"lon":-83.522133,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200763,"salesman":"House","active":"Inactive","name":"T.O.D. ONLINE (TIFTON)","address":"0","city":"TIFTON","state":"GA","zip":"31794","phone":"","ytdComp":0.0,"lat":31.215302,"lon":-83.270126,"accuracy":0.9,"accuracyType":"street_center"},{"num":200785,"salesman":"House","active":"Inactive","name":"PRINCE TRUCK CENTER","address":"1612 US HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"229-386-4040","ytdComp":0.0,"lat":31.446405,"lon":-83.539328,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200805,"salesman":"House","active":"Active","name":"JOEY HALL AUTO SALES LLC","address":"2548 HWY 82","city":"TIFTON","state":"GA","zip":"31794","phone":"2293826900","ytdComp":992.0,"lat":31.418678,"lon":-83.385569,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200821,"salesman":"House","active":"Active","name":"MARK'S BODY SHOP-TBR ONLY","address":"1541A HWY 82 EAST","city":"TIFTON","state":"GA","zip":"31794","phone":"2298483366","ytdComp":0.0,"lat":31.440904,"lon":-83.466185,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200831,"salesman":"Car Dealer","active":"Active","name":"GRIFFIN CHRYSLER DODGE JEEP","address":"505 7TH ST. WEST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293820440","ytdComp":2549.92,"lat":31.451748,"lon":-83.520697,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200832,"salesman":"Car Dealer","active":"Active","name":"GRIFFIN FORD","address":"PO BOX 765","city":"TIFTON","state":"GA","zip":"31793","phone":"2293821300","ytdComp":6298.74,"lat":31.462512,"lon":-83.593543,"accuracy":1.0,"accuracyType":"place"},{"num":200901,"salesman":"House","active":"Active","name":"LENCHO'S & SON TIRE SHOP","address":"1021 FERRY LAKE RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2296991259","ytdComp":3391.19,"lat":31.455536,"lon":-83.494794,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100282,"salesman":"Larry","active":"Active","name":"RUDY'S TIRE SERVICE","address":"222 E. 9TH. ST.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293825324","ytdComp":39838.71,"lat":31.448835,"lon":-83.507181,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100551,"salesman":"Larry","active":"Active","name":"SOUTHSIDE TIRE & AUTO SERVICE","address":"2700 S. MAIN ST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293876283","ytdComp":18491.16,"lat":31.439753,"lon":-83.504607,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200918,"salesman":"House","active":"Active","name":"SOUTH GEORGIA TRUCKING SVC LLC","address":"149 EARLY HUTCHINSON RD.","city":"TIFTON","state":"GA","zip":"31793","phone":"2294721056","ytdComp":4112.15,"lat":31.490881,"lon":-83.576122,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100591,"salesman":"Larry","active":"Active","name":"RAINEY ALIGNMENT","address":"111 MAGNOLIA DRIVE","city":"TIFTON","state":"GA","zip":"31794","phone":"2293827464","ytdComp":7278.62,"lat":31.454114,"lon":-83.524722,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100842,"salesman":"Larry","active":"Active","name":"LOVE AVE. SERVICE CTR.","address":"287 LOVE AVE.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293880430","ytdComp":2969.18,"lat":31.455819,"lon":-83.508372,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200944,"salesman":"House","active":"Inactive","name":"OILMASTERS","address":"116 7TH ST.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293825858","ytdComp":0.0,"lat":31.44969,"lon":-83.514252,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101323,"salesman":"Larry","active":"Active","name":"ERIC'S TIRE SERVICE","address":"1015 2ND ST W","city":"TIFTON","state":"GA","zip":"31794","phone":"2294721543","ytdComp":45533.29,"lat":31.457805,"lon":-83.528056,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":101326,"salesman":"Larry","active":"Active","name":"TIFTON GENERAL TIRE","address":"425 7TH ST W","city":"TIFTON","state":"GA","zip":"31794","phone":"2293826013","ytdComp":81623.3,"lat":31.451557,"lon":-83.519877,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101371,"salesman":"Larry","active":"Active","name":"DIRTY SOUTH KUSTOMS","address":"610 W SEVENTH ST","city":"TIFTON","state":"GA","zip":"31794","phone":"2292383992","ytdComp":2951.2,"lat":31.450666,"lon":-83.52229,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200409,"salesman":"Larry","active":"Active","name":"MASTER CRAFT IND.(NO PASS/LT)","address":"333 SOUTHWELL BLVD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293860610","ytdComp":5594.1,"lat":31.417415,"lon":-83.490027,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200468,"salesman":"Larry","active":"Active","name":"MCKEE'S AUTO CENTER  INC","address":"411 2ND ST E","city":"TIFTON","state":"GA","zip":"31794","phone":"2293827642","ytdComp":395.38,"lat":31.45337,"lon":-83.504987,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200483,"salesman":"Larry","active":"Active","name":"PERRIN FARM EQUIPMENT","address":"3664 HWY 82 W","city":"TIFTON","state":"GA","zip":"31794","phone":"2293829821","ytdComp":890.78,"lat":31.439768,"lon":-83.569354,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200490,"salesman":"Larry","active":"Active","name":"T.C.A. IRRIGATION","address":"2209 LESLIE LOCKE RD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293877097","ytdComp":243.06,"lat":31.445041,"lon":-83.550978,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200501,"salesman":"Larry","active":"Active","name":"LARRY'S BODY SHOP","address":"702 2ND STREET EAST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293864523","ytdComp":158.0,"lat":31.451819,"lon":-83.500555,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200503,"salesman":"Larry","active":"Active","name":"BB'S AUTOMOTIVE","address":"530 WEST 3RD STREET","city":"TIFTON","state":"GA","zip":"31794","phone":"2293824572","ytdComp":1710.89,"lat":31.452803,"lon":-83.519309,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200560,"salesman":"Larry","active":"Active","name":"FIVE STAR TIRE SERVICE LLC","address":"123 MAGNOLIA DRIVE","city":"TIFTON","state":"GA","zip":"31794","phone":"2293965412","ytdComp":22182.46,"lat":31.452723,"lon":-83.524982,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200608,"salesman":"Larry","active":"Active","name":"OFFROAD POWERSPORTS","address":"734 E 5TH ST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293877843","ytdComp":1905.83,"lat":31.449841,"lon":-83.502264,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200635,"salesman":"Larry","active":"Active","name":"DELTORO TIRE #2","address":"1678 SOUTH CARPENTER ROAD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293965510","ytdComp":115037.7,"lat":31.44805,"lon":-83.547412,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200676,"salesman":"Larry","active":"Active","name":"ALL PURPOSE AUTO CENTER","address":"2302 S. MAIN","city":"TIFTON","state":"GA","zip":"31794","phone":"2293260597","ytdComp":82.9,"lat":31.441567,"lon":-83.506083,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200744,"salesman":"Larry","active":"Active","name":"TIFTON COMMERCIAL","address":"181 EASTMAN DR","city":"TIFTON","state":"GA","zip":"31793","phone":"2294578600","ytdComp":10.73,"lat":31.444978,"lon":-83.536525,"accuracy":0.7,"accuracyType":"street_center"},{"num":200868,"salesman":"Larry","active":"Active","name":"GOLDEN ENVIRONMENTAL","address":"21 FARMERS MARKET RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293820309","ytdComp":6032.97,"lat":31.425697,"lon":-83.500492,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201105,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 2024","address":"1011 W 8TH ST","city":"TIFTON","state":"GA","zip":"31794","phone":"2292562228","ytdComp":0.0,"lat":31.463316,"lon":-83.527685,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200869,"salesman":"Larry","active":"Active","name":"AFFORDABLE DIESEL REPAIR","address":"4445 UNION RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293825000","ytdComp":0.0,"lat":31.42765,"lon":-83.516859,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200915,"salesman":"Larry","active":"Active","name":"DAVID'S AUTO SALES / TIFTON","address":"1826 US HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"2295205514","ytdComp":9378.88,"lat":31.445353,"lon":-83.541946,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200956,"salesman":"Larry","active":"Active","name":"TENNESON COLLISION CENTER","address":"2814 CARPENTER RD S.","city":"TIFTON","state":"GA","zip":"31794","phone":"2293919318","ytdComp":1234.21,"lat":31.438858,"lon":-83.547542,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200973,"salesman":"Larry","active":"Active","name":"ERIC'S TIRE (REBEL ROAD)","address":"1508 REBEL RD","city":"TIFTON","state":"GA","zip":"31793","phone":"2293964946","ytdComp":7145.0,"lat":31.445513,"lon":-83.54098,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201016,"salesman":"Larry","active":"Active","name":"O&C AUTO REPAIR","address":"503 SOUTH PARK AVE.","city":"TIFTON","state":"GA","zip":"31794","phone":"2298211221","ytdComp":291.04,"lat":31.450267,"lon":-83.51511,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201017,"salesman":"Larry","active":"Active","name":"PINEDA'S AUTOMOTIVE","address":"1271 US 82 EAST","city":"TIFTON","state":"GA","zip":"31794","phone":"2293821583","ytdComp":370.08,"lat":31.446529,"lon":-83.487976,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201043,"salesman":"Larry","active":"Active","name":"BILL'S TRAILER SERVICE","address":"1510 REBEL RD","city":"TIFTON","state":"GA","zip":"31794","phone":"2293964391","ytdComp":240.11,"lat":31.444925,"lon":-83.540723,"accuracy":0.99,"accuracyType":"rooftop"},{"num":201051,"salesman":"Larry","active":"Active","name":"EXPRESS OIL CHANGE #3168","address":"2006 US HWY 82 W","city":"TIFTON","state":"GA","zip":"31793","phone":"2295205980","ytdComp":0.0,"lat":31.445111,"lon":-83.544715,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201062,"salesman":"Larry","active":"Active","name":"TIRE MASTERS LLC","address":"101 MARTIN LUTHER KING JR DR","city":"TIFTON","state":"GA","zip":"31794","phone":"2294457500","ytdComp":5840.21,"lat":31.438634,"lon":-83.515764,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200495,"salesman":"House","active":"Inactive","name":"MADISON COUNTY CO-OP","address":"483 JACK THOMAS ROAD","city":"TONEY","state":"AL","zip":"35773","phone":"256-379-3040","ytdComp":0.0,"lat":34.892907,"lon":-86.601958,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200953,"salesman":"House","active":"Active","name":"SIMPLE TIRE - TIFTON","address":"5 NESHAMINY INTERPLEX DR.","city":"TREVOSE","state":"PA","zip":"19053","phone":"8884100604","ytdComp":155788.18,"lat":40.23538,"lon":-75.052822,"accuracy":0.86,"accuracyType":"range_interpolation"},{"num":201059,"salesman":"House","active":"Active","name":"HARRIS TIRE COMPANY","address":"1100 SOUTH BRUNDIDGE ST","city":"TROY","state":"AL","zip":"36081","phone":"8002391863","ytdComp":0.0,"lat":31.792676,"lon":-85.964922,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200402,"salesman":"House","active":"Inactive","name":"TIRE MASTERS","address":"4568 HWY 82 WEST","city":"TY TY","state":"GA","zip":"31795","phone":"229-232-7420","ytdComp":0.0,"lat":31.472433,"lon":-83.639417,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101329,"salesman":"House","active":"Inactive","name":"GRANDPA'S TIRE","address":"4614 US HWY 82","city":"TYTY","state":"GA","zip":"31795","phone":"229-472-1539","ytdComp":0.0,"lat":31.471442,"lon":-83.649752,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200231,"salesman":"House","active":"Inactive","name":"SOUTHERN TRACTOR SALES","address":"4021 US HWY 82 W","city":"TYTY","state":"GA","zip":"31795","phone":"229-777-3710","ytdComp":0.0,"lat":31.478354,"lon":-83.675888,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200418,"salesman":"House","active":"Inactive","name":"C","address":"4649 HWY 82","city":"TYTY","state":"GA","zip":"31795","phone":"2293828989","ytdComp":0.0,"lat":31.471653,"lon":-83.653742,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200419,"salesman":"House","active":"Inactive","name":"T","address":"4649 HWY 82","city":"TYTY","state":"GA","zip":"31795","phone":"2293821140","ytdComp":0.0,"lat":31.471653,"lon":-83.653742,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200467,"salesman":"House","active":"Inactive","name":"MCPHERSON SYSTEMS INC","address":"PO BOX 351","city":"TYTY","state":"GA","zip":"31795","phone":"229-386-2367","ytdComp":0.0,"lat":31.47185,"lon":-83.64684,"accuracy":1.0,"accuracyType":"place"},{"num":200612,"salesman":"House","active":"Inactive","name":"T","address":"142 INMAN ST","city":"TYTY","state":"GA","zip":"31795","phone":"229-382-1140","ytdComp":0.0,"lat":31.474386,"lon":-83.646967,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200757,"salesman":"House","active":"Inactive","name":"LINDSEY'S TIRE & TOWING","address":"4614 HWY 82 W","city":"TYTY","state":"GA","zip":"31795","phone":"229-206-3266","ytdComp":0.0,"lat":31.471442,"lon":-83.649752,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201030,"salesman":"House","active":"Inactive","name":"TY TY TIRE","address":"117 NORTH CHURCH ST.","city":"TYTY","state":"GA","zip":"31795","phone":"2293393745","ytdComp":0.0,"lat":31.472993,"lon":-83.649787,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100274,"salesman":"House","active":"Inactive","name":"HILL TIRE CO.","address":"2305 E.HILL AVE./ HWY 84","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-249-0013","ytdComp":0.0,"lat":30.846732,"lon":-83.237678,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":101146,"salesman":"Tiffany","active":"Active","name":"NE-RO TIRE & BRAKE SERVICE INC","address":"2311 S. PATTERSON ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292448353","ytdComp":12689.11,"lat":30.803136,"lon":-83.249894,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101553,"salesman":"House","active":"Inactive","name":"AUTO CONVENIENT CARE","address":"502 E CRANE AVE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-247-6369","ytdComp":0.0,"lat":30.831736,"lon":-83.272105,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101565,"salesman":"Tiffany","active":"Active","name":"HARRY B ANDERSON","address":"204 E CENTRAL AVE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292425945","ytdComp":1877.19,"lat":30.832951,"lon":-83.278251,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200147,"salesman":"House","active":"Inactive","name":"ALLSTATE USED TRUCKS OF VALDOS","address":"TA","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-242-2112","ytdComp":0.0,"lat":30.669063,"lon":-83.239765,"accuracy":0.79,"accuracyType":"street_center"},{"num":200266,"salesman":"Austin","active":"Active","name":"FUSSELL TIRE & SERVICE","address":"1775 WESTSIDE WAY","city":"VALDOSTA","state":"GA","zip":"31603","phone":"2292590034","ytdComp":122710.97,"lat":30.80999,"lon":-83.296291,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200286,"salesman":"House","active":"Inactive","name":"QUALITY TIRE CO","address":"2497 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292422338","ytdComp":0.0,"lat":30.78669,"lon":-83.270386,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200290,"salesman":"House","active":"Inactive","name":"S&S AUTOMOTIVE AND DIESEL","address":"3719 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292440530","ytdComp":0.0,"lat":30.744618,"lon":-83.276425,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200294,"salesman":"Tiffany","active":"Active","name":"TIRE KING OF VALDOSTA","address":"2608 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292471345","ytdComp":18321.92,"lat":30.874878,"lon":-83.277064,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200296,"salesman":"House","active":"Inactive","name":"VALDOSTA TOYOTA SCION","address":"2980 JAMES RD","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-247-1920","ytdComp":0.0,"lat":30.840324,"lon":-83.333471,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200298,"salesman":"House","active":"Active","name":"WISENBAKER'S TIRE & BRAKE","address":"819 S LEE ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292445229","ytdComp":0.0,"lat":30.81949,"lon":-83.269988,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200300,"salesman":"House","active":"Active","name":"Z TIRE EXPRESS","address":"4722 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31605","phone":"2292442084","ytdComp":0.0,"lat":30.932072,"lon":-83.242973,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200308,"salesman":"House","active":"Inactive","name":"AUDIO EXTREMES & TINT","address":"4309 N VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-242-5350","ytdComp":0.0,"lat":30.895714,"lon":-83.3376,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200321,"salesman":"House","active":"Inactive","name":"CASS BURCH CHRYSLER","address":"4164 N VALDOSTA","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-242-1540","ytdComp":0.0,"lat":30.897831,"lon":-83.330796,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200323,"salesman":"House","active":"Inactive","name":"CHOICE AUTOMOTIVE GROUP","address":"1606 W HILL ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292442777","ytdComp":0.0,"lat":30.823059,"lon":-83.307127,"accuracy":0.88,"accuracyType":"rooftop"},{"num":200324,"salesman":"House","active":"Inactive","name":"CITGO QUIK LUBE","address":"2183 NORTH ASHLEY","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-242-8777","ytdComp":0.0,"lat":30.860044,"lon":-83.281826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200325,"salesman":"House","active":"Active","name":"CMC TIRE & SERVICE","address":"1500 BAYTREE RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292478809","ytdComp":0.0,"lat":30.847351,"lon":-83.315084,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200327,"salesman":"House","active":"Active","name":"DRAPER TIRES & AUTOMOTIVE","address":"4207 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31605","phone":"2292474531","ytdComp":4730.71,"lat":30.916406,"lon":-83.257945,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200344,"salesman":"House","active":"Active","name":"JODY'S TIRES","address":"813 S ST AUGUSTINE RD","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292458880","ytdComp":0.0,"lat":30.811224,"lon":-83.295364,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200348,"salesman":"Car Dealer","active":"Active","name":"LANGDALE KIA (SERVICE)","address":"4021 N VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292423835","ytdComp":189.39,"lat":30.896289,"lon":-83.324064,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200349,"salesman":"Car Dealer","active":"Active","name":"HYUNDAI OF VALDOSTA","address":"4001 N VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292412880","ytdComp":0.0,"lat":30.89576,"lon":-83.322404,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200350,"salesman":"House","active":"Inactive","name":"LEE TIRE & AUTO","address":"300 W HILL","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-333-0011","ytdComp":0.0,"lat":30.832159,"lon":-83.276314,"accuracy":0.92,"accuracyType":"rooftop"},{"num":200352,"salesman":"House","active":"Active","name":"MALUDA AUTO SALES","address":"210 E GORDON ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292449163","ytdComp":2365.64,"lat":30.84074,"lon":-83.280544,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200354,"salesman":"House","active":"Inactive","name":"METTS AUTOMOTIVE","address":"1702 AL BROOKS DR","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-245-9909","ytdComp":0.0,"lat":30.843537,"lon":-83.311009,"accuracy":0.93,"accuracyType":"rooftop"},{"num":200355,"salesman":"House","active":"Active","name":"MOTION WHEELS HUBCAPS,&TIRES","address":"1302 S PATTERSON ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292530080","ytdComp":626.87,"lat":30.814195,"lon":-83.26831,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200362,"salesman":"Tiffany","active":"Active","name":"RAY NORTON TIRE & AUTO","address":"2606 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292471555","ytdComp":10759.84,"lat":30.874576,"lon":-83.27714,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200363,"salesman":"House","active":"Inactive","name":"REGISTER TIRE","address":"2183 N ASHLEY ST","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-247-0167","ytdComp":0.0,"lat":30.860044,"lon":-83.281826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200370,"salesman":"Tiffany","active":"Active","name":"SMITH TIRE COMPANY","address":"800 N PATTERSON","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292424830","ytdComp":0.0,"lat":30.838546,"lon":-83.283352,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200381,"salesman":"House","active":"Inactive","name":"VALDOSTA LINCOLN-MERCURY","address":"4534 N VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-242-7930","ytdComp":0.0,"lat":30.896721,"lon":-83.349784,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200382,"salesman":"House","active":"Inactive","name":"WALKER'S AUTO REPAIR","address":"2911 N PATTERSON ST","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-219-1114","ytdComp":0.0,"lat":30.871936,"lon":-83.290737,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200404,"salesman":"House","active":"Inactive","name":"SHOWTIME AUDIO & CUSTOMS","address":"2226 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-245-7455","ytdComp":0.0,"lat":30.86317,"lon":-83.280652,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200439,"salesman":"Tiffany","active":"Active","name":"BEAR TIRE SERVICE","address":"2600 HWY 41","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292421910","ytdComp":1367.91,"lat":30.791827,"lon":-83.238385,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200444,"salesman":"House","active":"Inactive","name":"TMR AUTO SALES","address":"2702 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-469-4733","ytdComp":0.0,"lat":30.876581,"lon":-83.27632,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200473,"salesman":"House","active":"Inactive","name":"SOUTHSIDE AUTOMOTIVE REPAIR","address":"1822 S PATTERSON ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-245-0660","ytdComp":0.0,"lat":30.808226,"lon":-83.25876,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200475,"salesman":"House","active":"Active","name":"AZALEA CITY AUTO SALES/SERVICE","address":"401 E GORDON ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292937337","ytdComp":800.74,"lat":30.840696,"lon":-83.277668,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200480,"salesman":"House","active":"Inactive","name":"ERIC'S TIRE (VALDOSTA)","address":"325 N ST AUGUSTINE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2294696353","ytdComp":0.0,"lat":30.831588,"lon":-83.30981,"accuracy":0.95,"accuracyType":"rooftop"},{"num":200486,"salesman":"House","active":"Inactive","name":"2-TIRE EXPRESS","address":"4722 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31605","phone":"229-244-2084","ytdComp":0.0,"lat":30.932072,"lon":-83.242973,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200488,"salesman":"House","active":"Inactive","name":"I-75 DIESEL","address":"3761 BOMAR LANE","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-740-1206","ytdComp":0.0,"lat":30.888172,"lon":-83.34985,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200551,"salesman":"House","active":"Inactive","name":"DAUGHARTY SERVICE STATION","address":"200 WEST HILL AVE.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292426028","ytdComp":0.0,"lat":30.830357,"lon":-83.280718,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200553,"salesman":"House","active":"Inactive","name":"LAMON HARVESTING  LLC.","address":"4421 HOLLY LANE","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-560-4591","ytdComp":0.0,"lat":30.897067,"lon":-83.362849,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200587,"salesman":"Car Dealer","active":"Active","name":"OSTEEN AUTOMOTIVE GROUP OF GA","address":"4140 N. VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292429920","ytdComp":0.0,"lat":30.898108,"lon":-83.328376,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200588,"salesman":"Car Dealer","active":"Active","name":"LANGDALE HYUNDAI OF SOUTH GA","address":"4001 NORTH VALDOSTA ROAD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292412880","ytdComp":1031.16,"lat":30.89576,"lon":-83.322404,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200606,"salesman":"House","active":"Active","name":"AFFORDABLE TIRE SERVICE LLC","address":"1508 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2293330690","ytdComp":1493.13,"lat":30.805484,"lon":-83.274979,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200631,"salesman":"House","active":"Active","name":"MARQUEZ TIRE SHOP","address":"608 W. HILL AVE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2299213897","ytdComp":5639.41,"lat":30.828255,"lon":-83.286516,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200654,"salesman":"House","active":"Active","name":"ECONOMY USED TIRE (VALDOSTA)","address":"511 W HILL AVE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292418988","ytdComp":37.0,"lat":30.828002,"lon":-83.284868,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200661,"salesman":"House","active":"Inactive","name":"OUT LAW KUSTOMS","address":"1305 BAYTREE RD","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-300-4834","ytdComp":0.0,"lat":30.846688,"lon":-83.307577,"accuracy":0.99,"accuracyType":"range_interpolation"},{"num":200662,"salesman":"House","active":"Inactive","name":"SOUTHERN AUTO REPAIR","address":"2148 N. ASHLEY ST.","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-375-3998","ytdComp":0.0,"lat":30.858919,"lon":-83.282102,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200663,"salesman":"Tiffany","active":"Active","name":"BEN'S TIRE & AUTO","address":"2183 N. ASHLEY ST.","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292428777","ytdComp":11875.81,"lat":30.860044,"lon":-83.281826,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200668,"salesman":"House","active":"Active","name":"24/7 TIRE","address":"810 N. ASHLEY ST.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292478473","ytdComp":3594.4,"lat":30.840178,"lon":-83.280278,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200674,"salesman":"House","active":"Active","name":"DASHER LLC","address":"805 N ST AUGUSTINE RD","city":"VALDOSTA","state":"GA","zip":"31621","phone":"2294696353","ytdComp":11192.32,"lat":30.837862,"lon":-83.315449,"accuracy":0.81,"accuracyType":"nearest_rooftop_match"},{"num":200677,"salesman":"House","active":"Inactive","name":"C&C AUTOMOTIVE & PERFORMANCE","address":"2620 BEMISS RD.","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2294748565","ytdComp":0.0,"lat":30.875853,"lon":-83.276591,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200680,"salesman":"House","active":"Inactive","name":"MIKE & SON SVC. CENTER","address":"1115 N. PATTERSON ST.","city":"VALDOSTA","state":"GA","zip":"31604","phone":"2292441567","ytdComp":0.0,"lat":30.843051,"lon":-83.284881,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200699,"salesman":"House","active":"Inactive","name":"METTS PERFORMANCE & AUTOMOTIVE","address":"4812 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31605","phone":"2292445511","ytdComp":480.76,"lat":30.935333,"lon":-83.24152,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200700,"salesman":"House","active":"Inactive","name":"J & W COMPLETE AUTO","address":"617 N. ASHLEY ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-506-7572","ytdComp":0.0,"lat":30.838328,"lon":-83.279513,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200708,"salesman":"House","active":"Inactive","name":"INJECTED DIESEL  LLC","address":"2497 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292422338","ytdComp":0.0,"lat":30.78669,"lon":-83.270386,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200717,"salesman":"House","active":"Active","name":"JW AUTOMOTIVE","address":"908 N LEE ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2297405299","ytdComp":1078.56,"lat":30.841787,"lon":-83.278235,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200733,"salesman":"House","active":"Inactive","name":"MAVIS TIRE (VALDOSTA)","address":"3121 INNER PERIMETER ROAD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-588-8030","ytdComp":0.0,"lat":30.879796,"lon":-83.297674,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200789,"salesman":"House","active":"Inactive","name":"LANGDALE FORD","address":"215 W MAGNOLIA ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"229-333-0762","ytdComp":0.0,"lat":30.833161,"lon":-83.28258,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200790,"salesman":"Car Dealer","active":"Active","name":"PRINCE AUTO. VALDOSTA  BUICK","address":"4550 N VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292423311","ytdComp":766.88,"lat":30.89639,"lon":-83.350935,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200791,"salesman":"House","active":"Inactive","name":"VALDOSTA CADILLAC","address":"3685 INNER PERIMETER RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"229-232-8925","ytdComp":0.0,"lat":30.883197,"lon":-83.270349,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200797,"salesman":"House","active":"Inactive","name":"DEATON TRUCK & TIRE SERVICE","address":"5849 PRODUCTION WAY","city":"VALDOSTA","state":"GA","zip":"31606","phone":"229-588-9009","ytdComp":0.0,"lat":30.831473,"lon":-83.225883,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200801,"salesman":"Car Dealer","active":"Active","name":"LANGDALE HONDA","address":"225 NORMAN DR.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292423835","ytdComp":0.0,"lat":30.823409,"lon":-83.314231,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200830,"salesman":"House","active":"Inactive","name":"REAMS AUTOMOTIVE","address":"302 E. GORDON ST.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292570494","ytdComp":0.0,"lat":30.841119,"lon":-83.279372,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200842,"salesman":"House","active":"Inactive","name":"CASS BURCH CHRYSLER (AMI ACCT)","address":"4164 N VALDOSTA","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292421540","ytdComp":0.0,"lat":30.897831,"lon":-83.330796,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200851,"salesman":"Car Dealer","active":"Active","name":"LANGDALE FORD","address":"215 W MAGNOLIA ST","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2293330762","ytdComp":0.0,"lat":30.833161,"lon":-83.28258,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200867,"salesman":"Tiffany","active":"Active","name":"24/7 DIESEL AND TIRE REPAIR","address":"3750 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2293758667","ytdComp":6021.71,"lat":30.74387,"lon":-83.278022,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200914,"salesman":"House","active":"Active","name":"BESTDRIVE COMMERCIAL TIRE CTR","address":"2299 E. HILL AVE","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292490013","ytdComp":107.0,"lat":30.846732,"lon":-83.237678,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200932,"salesman":"Tiffany","active":"Active","name":"RNR TIRE EXPRESS","address":"1500 W. HILL AVE.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2295888228","ytdComp":1174.44,"lat":30.823627,"lon":-83.305348,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200933,"salesman":"Car Dealer","active":"Active","name":"GRIFFIN CDJR VALDOSTA","address":"4164 NORTH VALDOSTA RD","city":"VALDOSTA","state":"GA","zip":"31604","phone":"2294714645","ytdComp":2665.08,"lat":30.897831,"lon":-83.330796,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200935,"salesman":"Tiffany","active":"Inactive","name":"FROMETA USED CAR & TIRE CENTER","address":"103 SOUTH SAINT AUGUSTINE RD","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2293338225","ytdComp":901.68,"lat":30.824913,"lon":-83.273457,"accuracy":0.8,"accuracyType":"nearest_rooftop_match"},{"num":200964,"salesman":"House","active":"Inactive","name":"MARTINEZ AUTO SERVICE","address":"500 N. PATTERSON ST.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"9728017027","ytdComp":402.98,"lat":30.835343,"lon":-83.281612,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200966,"salesman":"House","active":"Inactive","name":"SOUTH GEORGIA POWERSPORTS","address":"2713 BEMISS RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292426006","ytdComp":0.0,"lat":30.877359,"lon":-83.275085,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200967,"salesman":"Tiffany","active":"Active","name":"EDDIES AUTOMOTIVE AND TIRE","address":"210 EAST GORDON ST","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292324958","ytdComp":705.14,"lat":30.84074,"lon":-83.280544,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200968,"salesman":"House","active":"Inactive","name":"ROSS AUTOMOTIVE","address":"432 CONNELL RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292451253","ytdComp":0.0,"lat":30.875347,"lon":-83.27794,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201031,"salesman":"House","active":"Active","name":"RENO'S QUALITY COLLISION","address":"2421 HIGHWAY 41 SOUTH","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2295630752","ytdComp":286.52,"lat":30.798301,"lon":-83.243799,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201033,"salesman":"House","active":"Inactive","name":"BASSIC AUDIO & PERFROMANCE","address":"617 N. ASHLEY ST.","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2295067572","ytdComp":0.0,"lat":30.838328,"lon":-83.279513,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201050,"salesman":"House","active":"Active","name":"GOODYEAR COMMERCIAL TIRE & SVC","address":"3491 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292443179","ytdComp":0.0,"lat":30.752962,"lon":-83.270563,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201054,"salesman":"House","active":"Inactive","name":"EXPRESS OIL CHANGE #3197","address":"1072 N ST. AUGUSTINE RD","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2293755189","ytdComp":0.0,"lat":30.838831,"lon":-83.3192,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201055,"salesman":"House","active":"Active","name":"HERNANDEZ TIRES SHOP","address":"803 LANKFORD DR","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2296509357","ytdComp":1859.31,"lat":30.839196,"lon":-83.305539,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201102,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 662","address":"3121 INNER PERIMTER ROAD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2293528492","ytdComp":0.0,"lat":30.879796,"lon":-83.297674,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201222,"salesman":"House","active":"Active","name":"FIRESTONE STORE #029475","address":"3244 INNER PERIMETER RD","city":"VALDOSTA","state":"GA","zip":"31602","phone":"2292471889","ytdComp":0.0,"lat":30.881783,"lon":-83.291227,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000029,"salesman":"Tiffany","active":"Active","name":"SOUTHERN TIRE MART @ PILOT","address":"3491 MADISON HWY","city":"VALDOSTA","state":"GA","zip":"31601","phone":"2292443179","ytdComp":73.11,"lat":30.752962,"lon":-83.270563,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100566,"salesman":"House","active":"Inactive","name":"CLOSED ACCOUNT","address":"305 W. NORTH STREET","city":"VIDALIA","state":"GA","zip":"30474","phone":"912-537-3850","ytdComp":0.0,"lat":32.226004,"lon":-82.413021,"accuracy":0.89,"accuracyType":"nearest_rooftop_match"},{"num":200963,"salesman":"House","active":"Active","name":"CONLAN TIRE (BILL TO ACCT.)","address":"12225 STEPHENS RD","city":"WARREN","state":"MI","zip":"48089","phone":"5869397000","ytdComp":0.0,"lat":42.472148,"lon":-82.99914,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101108,"salesman":"House","active":"Active","name":"PARKER TIRE DIRECT","address":"123 PIKE RD","city":"WARWICK","state":"GA","zip":"31796","phone":"2294573818","ytdComp":41191.87,"lat":31.84531,"lon":-83.917851,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200523,"salesman":"House","active":"Inactive","name":"CLOSED","address":"4656 GA. HWY 313","city":"WARWICK","state":"GA","zip":"31796","phone":"229-535-6247","ytdComp":0.0,"lat":31.707882,"lon":-83.884277,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200760,"salesman":"Larry","active":"Active","name":"MALLARD'S SERVICE CENTER","address":"240 PEACHTREE ST","city":"WARWICK","state":"GA","zip":"31796","phone":"4789605246","ytdComp":816.49,"lat":31.833688,"lon":-83.922596,"accuracy":1.0,"accuracyType":"rooftop"},{"num":100796,"salesman":"House","active":"Inactive","name":"COMMERCIAL TIRE","address":"2623 KNIGHT AVE.","city":"WAYCROSS","state":"GA","zip":"31503","phone":"877-585-8473","ytdComp":0.0,"lat":31.202807,"lon":-82.316813,"accuracy":1.0,"accuracyType":"rooftop"},{"num":101477,"salesman":"Tiffany","active":"Active","name":"DISCOUNT TIRE (ALMA)OSTEEN","address":"7659 JAMESTOWN RD","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9123875359","ytdComp":4747.77,"lat":31.390665,"lon":-82.430768,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200428,"salesman":"House","active":"Inactive","name":"LUBE STOP","address":"700 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122852176","ytdComp":0.0,"lat":31.210904,"lon":-82.350129,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200445,"salesman":"Car Dealer","active":"Active","name":"CROSBY NISSAN","address":"2715 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122837783","ytdComp":0.0,"lat":31.191255,"lon":-82.317351,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":200459,"salesman":"House","active":"Active","name":"LOW COUNTRY TIRE LLC","address":"402 SYCAMORE ST","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9125906331","ytdComp":1565.8,"lat":31.207633,"lon":-82.340401,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200460,"salesman":"House","active":"Active","name":"COBB'S CORNER INC","address":"1730 ALBANY AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"912-285-7720","ytdComp":0.0,"lat":31.222301,"lon":-82.381982,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200461,"salesman":"Car Dealer","active":"Active","name":"WALKER JONES CHEVY-BUICK","address":"2700 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9124901314","ytdComp":2493.18,"lat":31.19175,"lon":-82.318972,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200484,"salesman":"Car Dealer","active":"Active","name":"ROBBIE ROBERSON FORD","address":"2825 MEMORIAL DRIVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122833131","ytdComp":0.0,"lat":31.190637,"lon":-82.315202,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200496,"salesman":"Tiffany","active":"Active","name":"MILLER TIRE CO.","address":"1915 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122875989","ytdComp":0.0,"lat":31.203075,"lon":-82.331841,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200534,"salesman":"House","active":"Inactive","name":"THE BODY SHOP OF WAYCROSS","address":"1970 BRUNSWICK HWY","city":"WAYCROSS","state":"GA","zip":"31502","phone":"912-285-3840","ytdComp":0.0,"lat":31.204783,"lon":-82.330348,"accuracy":0.99,"accuracyType":"rooftop"},{"num":200543,"salesman":"Tiffany","active":"Active","name":"GATOR TIRE","address":"1402 ALBANY AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9126709397","ytdComp":1680.36,"lat":31.220398,"lon":-82.375808,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200578,"salesman":"House","active":"Inactive","name":"WILBUR JAMES TIRE & BATTERY","address":"401 HICKS ST","city":"WAYCROSS","state":"GA","zip":"31501","phone":"912-283-6336","ytdComp":0.0,"lat":31.210894,"lon":-82.355625,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200579,"salesman":"Tiffany","active":"Active","name":"RODS CAR & TRUCK ACC.","address":"477 OSSIE DAVIS PKWY","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122830369","ytdComp":770.68,"lat":31.217922,"lon":-82.358322,"accuracy":0.87,"accuracyType":"nearest_rooftop_match"},{"num":200605,"salesman":"House","active":"Inactive","name":"GCR TIRES & SERVICE","address":"3710 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"912-285-9610","ytdComp":0.0,"lat":31.163855,"lon":-82.298313,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200611,"salesman":"House","active":"Inactive","name":"CLOSED","address":"2340 MINNESOTA AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"912-809-2261","ytdComp":0.0,"lat":31.188512,"lon":-82.391117,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200619,"salesman":"Tiffany","active":"Inactive","name":"DAVID'S AUTO SALES / WAYCROSS","address":"2251 KNIGHT AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122831818","ytdComp":0.0,"lat":31.204816,"lon":-82.323418,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200649,"salesman":"House","active":"Inactive","name":"CLOSED","address":"2103 KNIGHT AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"912-809-2632","ytdComp":0.0,"lat":31.204907,"lon":-82.324467,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200702,"salesman":"House","active":"Inactive","name":"TIRE ENGINEERS INC","address":"401 HICKS ST","city":"WAYCROSS","state":"GA","zip":"31501","phone":"912-283-6336","ytdComp":0.0,"lat":31.210894,"lon":-82.355625,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200706,"salesman":"House","active":"Inactive","name":"CARL'S SERVICE STATION LLC","address":"1006 PLANT AVE","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122838983","ytdComp":0.0,"lat":31.214451,"lon":-82.35346,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200707,"salesman":"House","active":"Active","name":"TIMBERLAND TIRE","address":"2498 C MEMORIAL DR.","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122831060","ytdComp":0.0,"lat":31.19364,"lon":-82.321092,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200710,"salesman":"House","active":"Inactive","name":"AARON'S AUTO REPAIR","address":"2623 KNIGHTS AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9123870503","ytdComp":0.0,"lat":31.202807,"lon":-82.316813,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200735,"salesman":"House","active":"Inactive","name":"TIRE ENGINEERS (WAYCROSS)","address":"401 HICKS ST.","city":"WAYCROSS","state":"GA","zip":"31501","phone":"912-283-6336","ytdComp":0.0,"lat":31.210894,"lon":-82.355625,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200754,"salesman":"House","active":"Active","name":"JORGE USED TIRE SHOP","address":"106 W. BLACKSHEAR AVE.","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9123095729","ytdComp":8152.04,"lat":31.236012,"lon":-82.371264,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200781,"salesman":"House","active":"Inactive","name":"WALKER-JONES CHRYSLER JEEP","address":"2730 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9125504113","ytdComp":0.0,"lat":31.191271,"lon":-82.317957,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200782,"salesman":"House","active":"Inactive","name":"WALKER-JONES GMC","address":"2731 MEMORIAL DR.","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9128164642","ytdComp":0.0,"lat":31.191199,"lon":-82.315971,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200793,"salesman":"Tiffany","active":"Active","name":"ALL SEASON AUTO REPAIR","address":"2623 KNIGHT AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122840006","ytdComp":718.6,"lat":31.202807,"lon":-82.316813,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200810,"salesman":"Tiffany","active":"Active","name":"BOULEVARD TIRE CENTER","address":"3710 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122859610","ytdComp":17366.63,"lat":31.163855,"lon":-82.298313,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200836,"salesman":"Tiffany","active":"Active","name":"KING MUFFLER","address":"1037 ALBANY AVE","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122854939","ytdComp":0.0,"lat":31.217608,"lon":-82.368287,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200855,"salesman":"House","active":"Inactive","name":"ROBBIE ROBERSON FORD (AMI ACT)","address":"2825 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122833131","ytdComp":0.0,"lat":31.190637,"lon":-82.315202,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200864,"salesman":"Car Dealer","active":"Active","name":"WALKER-JONES CHRYSLER DODGE","address":"2730 MEMORIAL DR","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9125504113","ytdComp":0.0,"lat":31.191271,"lon":-82.317957,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200886,"salesman":"Tiffany","active":"Active","name":"C&L PERFORMANCE INC","address":"1180 OSSIE DAVIS PKWY","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122830071","ytdComp":588.59,"lat":31.224612,"lon":-82.364172,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200904,"salesman":"House","active":"Inactive","name":"BOKOR  MICKY (SALESMAN ACCT.) CLOSED","address":"2469 CENTRAL AVE.","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9126141589","ytdComp":0.0,"lat":31.214402,"lon":-82.318722,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200917,"salesman":"House","active":"Active","name":"BENNETT'S TRACTOR SERVICE  LLC","address":"6698 ALMA HWY","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122841500","ytdComp":0.0,"lat":31.349149,"lon":-82.461911,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200919,"salesman":"House","active":"Inactive","name":"GUYS AUTOMOTIVE","address":"2630 BRUNSWICK HWY","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122854557","ytdComp":0.0,"lat":31.201008,"lon":-82.316867,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201013,"salesman":"House","active":"Inactive","name":"DIRTY SIDE DOWN OFFROAD","address":"1014 PLANT AVE","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9128164413","ytdComp":0.0,"lat":31.214899,"lon":-82.352854,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201024,"salesman":"House","active":"Active","name":"LIBERTY CAR WASH & TIRE","address":"817 AUGUSTA AVE","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9123870408","ytdComp":0.0,"lat":31.22503,"lon":-82.38742,"accuracy":1.0,"accuracyType":"range_interpolation"},{"num":201104,"salesman":"House","active":"Active","name":"MAVIS TIRES & BRAKES - 869","address":"401 HICKS STREET","city":"WAYCROSS","state":"GA","zip":"31501","phone":"9122540210","ytdComp":0.0,"lat":31.210894,"lon":-82.355625,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000001,"salesman":"House","active":"Active","name":"SHERROD CUSTOMS","address":"3151 INDUSTRIAL BLVD","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9124901210","ytdComp":0.0,"lat":31.253654,"lon":-82.419195,"accuracy":1.0,"accuracyType":"rooftop"},{"num":2000012,"salesman":"Tiffany","active":"Active","name":"GLOBAL TRUCK & EQUIPMENT SALES","address":"2632 BRUNSWICK HWY","city":"WAYCROSS","state":"GA","zip":"31503","phone":"9122854557","ytdComp":1531.78,"lat":31.20099,"lon":-82.316731,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200155,"salesman":"House","active":"Inactive","name":"COUNTY FOOD AND TIRE","address":"5324 GA HWY 111 S","city":"WHIGHAM","state":"GA","zip":"39897","phone":"229-872-3941","ytdComp":0.0,"lat":31.01003,"lon":-84.147336,"accuracy":0.92,"accuracyType":"range_interpolation"},{"num":200238,"salesman":"House","active":"Active","name":"CROSSROADS TIRE & ACC LLC","address":"PO BOX 31","city":"WHIGHAM","state":"GA","zip":"39839","phone":"","ytdComp":0.0,"lat":30.901959,"lon":-84.311493,"accuracy":1.0,"accuracyType":"place"},{"num":200268,"salesman":"House","active":"Active","name":"GIANT TIRE SALES/SERVICE","address":"3523 HWY 84 W","city":"WHIGHAM","state":"GA","zip":"39897","phone":"2297623230","ytdComp":465.12,"lat":30.881648,"lon":-84.335582,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200283,"salesman":"House","active":"Inactive","name":"MOBILE TIRE SERVICE","address":"3731 TIRED CREEK RD","city":"WHIGHAM","state":"GA","zip":"39897","phone":"229-872-3793","ytdComp":0.0,"lat":30.813541,"lon":-84.378625,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200906,"salesman":"Tiffany","active":"Active","name":"E&H TIRE","address":"16804 SPRING ST","city":"WHITE SPRINGS","state":"FL","zip":"32096","phone":"3863973333","ytdComp":32722.63,"lat":30.330855,"lon":-82.749876,"accuracy":1.0,"accuracyType":"rooftop"},{"num":200386,"salesman":"House","active":"Inactive","name":"SHANE'S TIRE & AUTO","address":"173 E MAIN ST","city":"WILLACOOCHEE","state":"GA","zip":"31650","phone":"9124225615","ytdComp":0.0,"lat":31.340074,"lon":-83.043729,"accuracy":1.0,"accuracyType":"rooftop"},{"num":201014,"salesman":"Larry","active":"Active","name":"D&R AUTO SALES & SALVAGE PARTS","address":"524 MAIN ST EAST","city":"WILLACOOCHEE","state":"GA","zip":"31650","phone":"9125346543","ytdComp":1681.99,"lat":31.337842,"lon":-83.03836,"accuracy":1.0,"accuracyType":"rooftop"}];

const SEED_AR = [{"salesman":"Larry","custNum":200198,"shortName":"MOULTRIE","name":"MOULTRIE TIRE","phone":"229-985-5619","balance":38152.35,"futDue":10140.19,"curDue":3814.34,"due1_30":24197.82,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"Tiffany","custNum":200636,"shortName":"MTC","name":"MTC SOUTH, INC.","phone":"850-251-5393","balance":49687.78,"futDue":11893.53,"curDue":21737.52,"due1_30":16056.73,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"Larry","custNum":200635,"shortName":"DEL","name":"DELTORO TIRE #2","phone":"229-396-5510","balance":81519.47,"futDue":9682.34,"curDue":68424.84,"due1_30":3412.29,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":101241,"shortName":"PETTIS","name":"NEW PETTIS TIRE","phone":"229-273-1153","balance":6659.99,"futDue":0.0,"curDue":174.48,"due1_30":1338.99,"due31_60":3573.36,"due61_90":1573.16,"dueOver90":0.0,"lastPaid":"2026-02-03"},{"salesman":"Tiffany","custNum":200293,"shortName":"THOM","name":"THOMASVILLE TIRE DEPT.","phone":"229-228-0260","balance":3664.84,"futDue":34.36,"curDue":2603.92,"due1_30":1026.56,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":201056,"shortName":"LEE","name":"LEE'S AUTO SHOP","phone":"813-368-4143","balance":841.23,"futDue":0.0,"curDue":12.43,"due1_30":828.8,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-20"},{"salesman":"House","custNum":200829,"shortName":"TIRE","name":"TIRE SOLUTIONS & VEH. REPAIRS","phone":"229-985-8473","balance":43129.6,"futDue":1469.93,"curDue":664.58,"due1_30":676.56,"due31_60":699.74,"due61_90":2309.43,"dueOver90":37309.36,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200979,"shortName":"DISCOUNT","name":"DISCOUNT TIRE & AUTO SHOP","phone":"850-544-7234","balance":844.5,"futDue":0.0,"curDue":318.92,"due1_30":525.58,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-17"},{"salesman":"House","custNum":2000013,"shortName":"TIRE","name":"TIRE SOLUTIONS & VEH. REPAIRS","phone":"229-294-2801","balance":32055.34,"futDue":1080.26,"curDue":467.24,"due1_30":489.36,"due31_60":958.54,"due61_90":5536.53,"dueOver90":23523.41,"lastPaid":"2026-05-20"},{"salesman":"House","custNum":200238,"shortName":"CROSSROADS","name":"CROSSROADS TIRE & ACC LLC","phone":"","balance":3776.48,"futDue":0.0,"curDue":109.22,"due1_30":53.42,"due31_60":52.64,"due61_90":51.86,"dueOver90":3509.34,"lastPaid":"2025-10-07"},{"salesman":"","custNum":2000053,"shortName":"JOINER","name":"JOINER CONTRACTING","phone":"229-269-3842","balance":7.08,"futDue":0.1,"curDue":0.0,"due1_30":6.98,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"","custNum":3000389,"shortName":"SOUTHERN","name":"SOUTHERN TIRE MART, LLC (#134)","phone":"229-559-1913","balance":32569.37,"futDue":21125.09,"curDue":11444.14,"due1_30":0.14,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"House","custNum":201040,"shortName":"TIRES","name":"TIRES EASY (NAP - TIFTON)","phone":"844-347-0789","balance":2691.13,"futDue":94.18,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":2596.95,"dueOver90":0.0,"lastPaid":""},{"salesman":"Tiffany","custNum":200867,"shortName":"24","name":"24/7 DIESEL AND TIRE REPAIR","phone":"229-375-8667","balance":2065.97,"futDue":2065.97,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"House","custNum":200668,"shortName":"247","name":"24/7 TIRE","phone":"229-247-8473","balance":3848.62,"futDue":3835.74,"curDue":12.88,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-27"},{"salesman":"Larry","custNum":200478,"shortName":"82","name":"82 TIRE & LUBE","phone":"912-462-7357","balance":-1803.54,"futDue":-1803.54,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":201032,"shortName":"A1","name":"A-1 TIRE PLUS","phone":"386-867-5495","balance":3310.16,"futDue":3310.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Larry","custNum":200946,"shortName":"ABBI","name":"ABBI'S 24 HOUR","phone":"229-449-0762","balance":5857.4,"futDue":5857.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":200664,"shortName":"ABR","name":"ABR COMMERCIAL TRUCK & AUTO","phone":"229-995-2169","balance":-0.1,"futDue":-0.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-01"},{"salesman":"House","custNum":200401,"shortName":"ADEL","name":"ADEL TIRE CO","phone":"229-896-3086","balance":-58.31,"futDue":-58.31,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Tiffany","custNum":200922,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-236-8825","balance":111189.45,"futDue":111189.45,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"Tiffany","custNum":200923,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"386-406-6745","balance":838.12,"futDue":838.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"Tiffany","custNum":2000021,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-559-0708","balance":2280.24,"futDue":2280.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"Tiffany","custNum":2000039,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-236-8825","balance":853.17,"futDue":853.17,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Tiffany","custNum":2000040,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-691-7771","balance":2890.26,"futDue":2890.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"House","custNum":201037,"shortName":"AFTER","name":"AFTER HOURS TIRE SERVICE","phone":"229-589-6702","balance":0.18,"futDue":0.18,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":201058,"shortName":"AFTER","name":"AFTER 5 COMM. TIRE & OFF ROAD","phone":"352-451-3863","balance":-67.1,"futDue":-67.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-12"},{"salesman":"Larry","custNum":100301,"shortName":"ALBANY","name":"ALBANY GENERAL TIRE SERVICE","phone":"229-436-2485","balance":3405.18,"futDue":3405.18,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Car Dealer","custNum":200765,"shortName":"ALBANY","name":"ALBANY CHRYSLER DODGE JEEP RAM","phone":"229-233-7769","balance":907.9,"futDue":907.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"Larry","custNum":200676,"shortName":"ALL","name":"ALL PURPOSE AUTO CENTER","phone":"229-326-0597","balance":-18.05,"futDue":-18.05,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-06"},{"salesman":"Tiffany","custNum":2000018,"shortName":"ALL","name":"ALL PRO DIESEL, LLC","phone":"386-438-9912","balance":-2.04,"futDue":-2.04,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200971,"shortName":"ALLEN","name":"ALLEN'S TIRE","phone":"229-567-3390","balance":10762.33,"futDue":9554.41,"curDue":1207.92,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"House","custNum":200897,"shortName":"ALMA","name":"ALMA TIRE & AUTO REPAIR","phone":"912-286-9140","balance":-671.28,"futDue":-671.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"Larry","custNum":101080,"shortName":"AMERSON","name":"AMERSON TIRE INC.","phone":"912-393-3674","balance":93.97,"futDue":93.97,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Car Dealer","custNum":200734,"shortName":"ANDERSON","name":"ANDERSON FORD","phone":"912-384-2600","balance":303.88,"futDue":303.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"House","custNum":200307,"shortName":"ARR","name":"ARREDONDO TIRE SERVICE","phone":"229-529-0882","balance":-0.94,"futDue":-0.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-13"},{"salesman":"House","custNum":200709,"shortName":"ASHLEY","name":"ASHLEY'S AUTOMOTIVE REPAIR","phone":"229-396-4640","balance":-174.99,"futDue":-174.99,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"House","custNum":101439,"shortName":"AT","name":"A.T. TIRE SERVICE","phone":"229-891-5428","balance":1280.15,"futDue":1280.15,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200311,"shortName":"AUTO","name":"AUTOMOTIVE NECESSITIES","phone":"229-434-1237","balance":-752.72,"futDue":-752.72,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":200753,"shortName":"AUTO","name":"AUTO TECH OF MIAMI INC.","phone":"850-997-0200","balance":-2.15,"futDue":-2.15,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"House","custNum":200475,"shortName":"AZALEA","name":"AZALEA CITY AUTO SALES/SERVICE","phone":"229-293-7337","balance":-0.14,"futDue":-0.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Larry","custNum":200503,"shortName":"BB","name":"BB'S AUTOMOTIVE","phone":"229-382-4572","balance":-0.57,"futDue":-0.57,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":200607,"shortName":"BB","name":"B AND B SERVICE CENTER, INC.","phone":"229-236-2886","balance":66.94,"futDue":66.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"","custNum":2000049,"shortName":"BEALL","name":"BEALL TIRE WHOLESALE, LLC","phone":"229-294-0600","balance":3208.04,"futDue":3208.04,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"Tiffany","custNum":200439,"shortName":"BEAR","name":"BEAR TIRE SERVICE","phone":"229-242-1910","balance":-125.21,"futDue":-125.21,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22"},{"salesman":"Austin","custNum":2000030,"shortName":"BEASON","name":"BEASON EQUIPMENT CO","phone":"229-985-9785","balance":2907.14,"futDue":2907.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07"},{"salesman":"Tiffany","custNum":200663,"shortName":"BEN","name":"BEN'S TIRE & AUTO","phone":"229-242-8777","balance":-341.84,"futDue":-341.84,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200592,"shortName":"BERNEY","name":"BERNEY'S TIRE SERVICE","phone":"229-435-0412","balance":3122.71,"futDue":3122.71,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"Larry","custNum":201053,"shortName":"BERNEY","name":"BERNEYS TIRE SERVICE","phone":"229-435-0413","balance":7935.43,"futDue":7935.43,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"Larry","custNum":101479,"shortName":"BILL","name":"BILL THOMPSON TIRE SERVICES","phone":"229-435-7753","balance":2253.01,"futDue":2253.01,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":201043,"shortName":"BILL","name":"BILL'S TRAILER SERVICE","phone":"229-396-4391","balance":184.17,"futDue":0.0,"curDue":184.17,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-20"},{"salesman":"House","custNum":101366,"shortName":"BM","name":"B & M AUTOMOTIVE SERVICE","phone":"912-384-6115","balance":-60.68,"futDue":-60.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-06"},{"salesman":"House","custNum":101525,"shortName":"BMS","name":"BMS DISCOUNT TIRES","phone":"229-234-0033","balance":-439.48,"futDue":-439.48,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":200315,"shortName":"BOB","name":"BOBBY'S CITGO","phone":"229-482-2724","balance":3679.99,"futDue":3679.99,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"House","custNum":200809,"shortName":"BOULEVARD","name":"BOULEVARD TIRE CENTER","phone":"386-734-6447","balance":7515.97,"futDue":7515.97,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":200146,"shortName":"BRACE","name":"BRACEWELL AUTOMOTIVE SERVICE","phone":"229-377-1771","balance":2728.52,"futDue":2728.52,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"Larry","custNum":200317,"shortName":"BRO","name":"BROTHERS TIRES","phone":"229-319-2919","balance":-16.8,"futDue":-16.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":101152,"shortName":"BROOKS","name":"BROOKS BODY SHOP","phone":"229-386-1800","balance":-8.11,"futDue":-8.11,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-27"},{"salesman":"House","custNum":201049,"shortName":"BROUSSARD","name":"BROUSSARD ACCESSORIES LLC","phone":"229-496-8004","balance":-63.31,"futDue":-63.31,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-06"},{"salesman":"Larry","custNum":200319,"shortName":"BUCK","name":"BUCK'S AUTO REPAIR","phone":"229-686-2290","balance":181.57,"futDue":181.57,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":200331,"shortName":"BUDGET","name":"BUDGET CAR SALES","phone":"229-388-0020","balance":1105.63,"futDue":1105.63,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-13"},{"salesman":"Larry","custNum":101507,"shortName":"BURNETTE","name":"BURNETTE AUTOMOTIVE SERVICE","phone":"912-632-2713","balance":8088.71,"futDue":8088.71,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-13"},{"salesman":"Larry","custNum":101283,"shortName":"CAMERON","name":"CAMERON'S TOWING AND TIRE","phone":"229-567-2437","balance":2065.56,"futDue":2065.56,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200600,"shortName":"CHARLOT","name":"CHARLOT TRUCKING & TIRE SVC.","phone":"229-881-4561","balance":-6.53,"futDue":-6.53,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200681,"shortName":"CITY","name":"CITY OF SYLVESTER","phone":"229-776-8504","balance":291.22,"futDue":291.22,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-23"},{"salesman":"Tiffany","custNum":200886,"shortName":"CL","name":"C&L PERFORMANCE INC","phone":"912-283-0071","balance":113.29,"futDue":113.29,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-10"},{"salesman":"Larry","custNum":100417,"shortName":"CLARK","name":"CLARK BASS SERVICE","phone":"229-874-4685","balance":11992.97,"futDue":11992.97,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200690,"shortName":"COLQUITT","name":"COLQUITT COUNTY TIRE LLC","phone":"229-454-1084","balance":1387.49,"futDue":1164.56,"curDue":222.93,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200807,"shortName":"CORDELE","name":"CORDELE TIRE & WHEEL, LLC","phone":"229-417-5099","balance":747.13,"futDue":747.13,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":101539,"shortName":"COURSON","name":"COURSON'S TIRE OF DOUGLAS","phone":"912-383-6188","balance":6516.22,"futDue":6516.22,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":201018,"shortName":"CRAWLEY","name":"CRAWLEY'S AUTOMOTIVE & TIRE","phone":"386-658-2007","balance":254.0,"futDue":254.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":200674,"shortName":"DASH","name":"DASHER LLC","phone":"229-469-6353","balance":2.0,"futDue":2.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Larry","custNum":101295,"shortName":"DAVID","name":"DAVID'S AUTO SALES / DOUGLAS","phone":"912-384-8570","balance":8538.93,"futDue":8538.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07"},{"salesman":"Larry","custNum":201035,"shortName":"DAVID","name":"DAVID'S AUTO SALES (MOULTRIE)","phone":"229-217-4959","balance":7694.63,"futDue":7694.63,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07"},{"salesman":"House","custNum":201008,"shortName":"DBJ","name":"DBJ MOBILE TIRE SERVICE, INC.","phone":"386-219-6036","balance":487.87,"futDue":487.87,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200162,"shortName":"DELTA","name":"DELTA TIRE CO","phone":"229-246-2750","balance":25880.34,"futDue":13312.71,"curDue":12567.63,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-24"},{"salesman":"House","custNum":200416,"shortName":"DENT","name":"DENT'S SERVICE STATION","phone":"229-896-4160","balance":204.91,"futDue":204.91,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200506,"shortName":"DEV","name":"DEVANE TIRE & SERVICE LLC","phone":"229-310-9586","balance":1000.44,"futDue":1000.44,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-24"},{"salesman":"Larry","custNum":101371,"shortName":"DIRTY","name":"DIRTY SOUTH KUSTOMS","phone":"229-238-3992","balance":0.6,"futDue":0.0,"curDue":0.6,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200870,"shortName":"DK","name":"D&K USED TIRES","phone":"229-496-1487","balance":0.88,"futDue":0.0,"curDue":0.88,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Larry","custNum":201014,"shortName":"DR","name":"D&R AUTO SALES & SALVAGE PARTS","phone":"912-534-6543","balance":-6.68,"futDue":-6.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-21"},{"salesman":"House","custNum":200327,"shortName":"DRAPER","name":"DRAPER TIRES & AUTOMOTIVE","phone":"229-247-4531","balance":-111.14,"futDue":-111.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":201034,"shortName":"ECONOMIC","name":"ECONOMIC NICHOLAS TIRE","phone":"229-454-5123","balance":-12.32,"futDue":-12.32,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-06"},{"salesman":"Larry","custNum":101436,"shortName":"ED","name":"ED'S TIRE","phone":"229-776-6952","balance":4742.69,"futDue":4742.69,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"Tiffany","custNum":200967,"shortName":"EDD","name":"EDDIES AUTOMOTIVE AND TIRE","phone":"229-232-4958","balance":940.4,"futDue":940.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200166,"shortName":"EDISON","name":"EDISON TIRE","phone":"229-835-2077","balance":5873.36,"futDue":5873.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":2000015,"shortName":"EDWIN","name":"EDWIN'S TIRES LLC","phone":"229-429-6141","balance":-12.6,"futDue":-12.6,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-05"},{"salesman":"House","custNum":200642,"shortName":"EG","name":"E.G. AUTO SALES","phone":"229-769-5011","balance":-4.0,"futDue":-4.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"Tiffany","custNum":200906,"shortName":"EH","name":"E&H TIRE","phone":"386-397-3333","balance":-192.9,"futDue":-192.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200891,"shortName":"EJH","name":"EJH WRECKER & TIRE SERVICE","phone":"229-566-3334","balance":-317.7,"futDue":-317.7,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":101512,"shortName":"ELLENTON","name":"ELLENTON TIRE AND AUTO","phone":"229-324-2475","balance":45588.21,"futDue":24263.71,"curDue":21324.5,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"Larry","custNum":101323,"shortName":"ERIC","name":"ERIC'S TIRE SERVICE","phone":"229-472-1543","balance":10840.67,"futDue":10840.67,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Larry","custNum":200972,"shortName":"ERIC","name":"ERIC'S TIRE OF SYLVESTER","phone":"229-821-2000","balance":2863.99,"futDue":1031.37,"curDue":1832.62,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-20"},{"salesman":"Larry","custNum":200973,"shortName":"ERIC","name":"ERIC'S TIRE (REBEL ROAD)","phone":"229-396-4946","balance":3058.49,"futDue":3058.49,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":200410,"shortName":"EZ","name":"EZDEALIN WHEELS AND TIRES","phone":"877-247-2230","balance":23539.99,"futDue":23539.99,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":2000038,"shortName":"FAST","name":"FAST TIRE SERVICE","phone":"386-324-1853","balance":1.12,"futDue":0.0,"curDue":1.12,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-13"},{"salesman":"House","custNum":200647,"shortName":"FAUSETTS","name":"FAUSETTS TIRE CO.","phone":"229-896-7481","balance":-22.2,"futDue":-22.2,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Car Dealer","custNum":101300,"shortName":"FITZGERALD","name":"FITZGERALD FORD AND LINCOLN","phone":"229-423-8787","balance":794.68,"futDue":794.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15"},{"salesman":"Larry","custNum":200560,"shortName":"FIVE","name":"FIVE STAR TIRE SERVICE,LLC","phone":"229-396-5412","balance":4123.75,"futDue":4123.75,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Car Dealer","custNum":2000025,"shortName":"FORD","name":"FORD CORDELE","phone":"229-276-0607","balance":4347.58,"futDue":1780.8,"curDue":2566.78,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-10"},{"salesman":"Car Dealer","custNum":2000024,"shortName":"FORKLIFT","name":"FORKLIFT TIRE OF CENTRAL FL","phone":"863-559-9353","balance":614.68,"futDue":0.0,"curDue":614.68,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14"},{"salesman":"House","custNum":200719,"shortName":"FOSTER","name":"FOSTER EASY PAY TIRE CO., INC.","phone":"229-995-2167","balance":15122.91,"futDue":15122.91,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":101297,"shortName":"FOUR","name":"FOUR C'S LUBE","phone":"912-422-6866","balance":-2.24,"futDue":-2.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Tiffany","custNum":200935,"shortName":"FROMETA","name":"FROMETA USED CAR & TIRE CENTER","phone":"229-333-8225","balance":-99.36,"futDue":-99.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-04"},{"salesman":"Austin","custNum":200266,"shortName":"FUSS","name":"FUSSELL TIRE & SERVICE","phone":"229-259-0034","balance":38311.36,"futDue":38311.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"","custNum":2000063,"shortName":"GATEWAY","name":"GATEWAY TIRE","phone":"662-449-3832","balance":579.63,"futDue":579.63,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":""},{"salesman":"Tiffany","custNum":200543,"shortName":"GATOR","name":"GATOR TIRE","phone":"912-670-9397","balance":5.9,"futDue":0.0,"curDue":5.9,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":200268,"shortName":"GIANT","name":"GIANT TIRE SALES/SERVICE","phone":"229-762-3230","balance":7522.69,"futDue":574.69,"curDue":6948.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14"},{"salesman":"House","custNum":200900,"shortName":"GILLETTES","name":"GILLETTES AUTO","phone":"386-362-5171","balance":-17.64,"futDue":-17.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-30"},{"salesman":"Larry","custNum":200868,"shortName":"GOLDEN","name":"GOLDEN ENVIRONMENTAL","phone":"229-382-0309","balance":7686.62,"futDue":3588.16,"curDue":4098.46,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":200962,"shortName":"GREENE","name":"GREENE'S TIRE SERVICE LLC","phone":"229-699-0748","balance":-19.88,"futDue":-19.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"Car Dealer","custNum":200831,"shortName":"GRIFFFIN","name":"GRIFFIN CHRYSLER DODGE JEEP","phone":"229-382-0440","balance":348.12,"futDue":348.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Car Dealer","custNum":200832,"shortName":"GRIFFIN","name":"GRIFFIN FORD","phone":"229-382-1300","balance":527.8,"futDue":527.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Car Dealer","custNum":200933,"shortName":"GRIFFIN","name":"GRIFFIN CDJR VALDOSTA","phone":"229-471-4645","balance":228.76,"futDue":228.76,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Car Dealer","custNum":1999999,"shortName":"GRIFFIN","name":"GRIFFIN CHEVROLET OF SYLVESTER","phone":"229-776-3473","balance":98.52,"futDue":98.52,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Tiffany","custNum":200270,"shortName":"HAHIRA","name":"HAHIRA AUTOMOTIVE SERVICE","phone":"229-794-2429","balance":3855.93,"futDue":3855.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"House","custNum":200451,"shortName":"HARROD","name":"HARROD BROTHERS","phone":"229-686-3959","balance":2813.38,"futDue":2813.38,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"Tiffany","custNum":101565,"shortName":"HARRY","name":"HARRY B ANDERSON","phone":"229-242-5945","balance":2587.47,"futDue":1379.04,"curDue":1208.43,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-16"},{"salesman":"House","custNum":201069,"shortName":"HEAVY","name":"HEAVY DUTY TIRE","phone":"904-833-7750","balance":-6.6,"futDue":-6.6,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-01-16"},{"salesman":"Larry","custNum":200683,"shortName":"HICKOX","name":"HICKOX AUTO DEALERS","phone":"912-281-3922","balance":-0.5,"futDue":-0.5,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"House","custNum":200417,"shortName":"HIGH","name":"HIGHWAY TIRE & DIESEL","phone":"229-546-4506","balance":245.04,"futDue":245.04,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-27"},{"salesman":"Car Dealer","custNum":200554,"shortName":"HONDA","name":"HONDA OF SOUTH GEORGIA","phone":"229-396-4050","balance":19.72,"futDue":19.72,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-20"},{"salesman":"House","custNum":200277,"shortName":"IMP","name":"IMPORT SERVICE & SALES","phone":"229-226-9844","balance":926.81,"futDue":926.81,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-08"},{"salesman":"","custNum":2000069,"shortName":"IRWIN","name":"IRWIN COUNTY CUSTOMS & REPAIR","phone":"229-326-5358","balance":0.9,"futDue":0.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Car Dealer","custNum":200580,"shortName":"JEFF","name":"JEFF FENDER BUICK, GMC, CAD.","phone":"229-386-1985","balance":4979.64,"futDue":4979.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Larry","custNum":101161,"shortName":"JMC","name":"JMC TIRE CO., INC.","phone":"912-384-4940","balance":-179.49,"futDue":-179.49,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200718,"shortName":"JOBBER #2","name":"JOBBER ACCT (TIFTON)","phone":"","balance":4.56,"futDue":0.0,"curDue":4.56,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22"},{"salesman":"House","custNum":200648,"shortName":"JOE","name":"JOE'S AUTO REPAIR, LLC","phone":"912-384-2010","balance":-78.86,"futDue":-78.86,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15"},{"salesman":"House","custNum":200805,"shortName":"JOEY","name":"JOEY HALL AUTO SALES LLC","phone":"229-382-6900","balance":-0.92,"futDue":-0.92,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-03"},{"salesman":"Larry","custNum":100107,"shortName":"JOHNSON","name":"JOHNSON AUTO & TIRE","phone":"912-359-2452","balance":969.24,"futDue":969.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Larry","custNum":2000017,"shortName":"JORDAN","name":"JORDAN AUTOMOTIVE & TIRES","phone":"229-439-6881","balance":-6.0,"futDue":-6.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"House","custNum":200754,"shortName":"JORGE","name":"JORGE USED TIRE SHOP","phone":"912-309-5729","balance":797.47,"futDue":797.47,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":2000007,"shortName":"KEATON","name":"KEATON & SON TIRE LLC","phone":"850-284-9344","balance":383.08,"futDue":383.08,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200691,"shortName":"KENDA","name":"KENDA TRUCK CENTER","phone":"850-929-3700","balance":30261.87,"futDue":30261.87,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14"},{"salesman":"Tiffany","custNum":200836,"shortName":"KING","name":"KING MUFFLER","phone":"912-285-4939","balance":956.26,"futDue":0.0,"curDue":956.26,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-11"},{"salesman":"Car Dealer","custNum":2000009,"shortName":"KING","name":"KING FORD OF NASHVILLE","phone":"229-686-2058","balance":-1625.88,"futDue":-1625.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"","custNum":2000052,"shortName":"LA","name":"L&A TIRE, LLC","phone":"386-855-3729","balance":-42.12,"futDue":-42.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"","custNum":2000059,"shortName":"LAKE","name":"LAKE CITY TIRE SHOP","phone":"386-344-4608","balance":441.1,"futDue":441.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":""},{"salesman":"Tiffany","custNum":200827,"shortName":"LAKELAND","name":"LAKELAND TIRE DBA COOK & SONS","phone":"229-482-1000","balance":1258.86,"futDue":-162.38,"curDue":1421.24,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200912,"shortName":"LANE","name":"LANE'S TRK & TRL REPAIR & AUTO","phone":"229-322-6338","balance":-0.1,"futDue":-0.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-06"},{"salesman":"Car Dealer","custNum":200348,"shortName":"LANG","name":"LANGDALE KIA (SERVICE)","phone":"229-242-3835","balance":75.02,"futDue":75.02,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14"},{"salesman":"Larry","custNum":200501,"shortName":"LAR","name":"LARRY'S BODY SHOP","phone":"229-386-4523","balance":3.78,"futDue":3.78,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":200913,"shortName":"LASHLEY","name":"LASHLEY'S HOMETOWN TIRE LLC","phone":"386-209-4919","balance":-2.93,"futDue":-2.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"House","custNum":200750,"shortName":"LAWN","name":"LAWN PERFORMANCE, LLC","phone":"229-435-5007","balance":191.0,"futDue":191.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-03-07"},{"salesman":"House","custNum":200191,"shortName":"LEE","name":"LEE COUNTY AUTO SERVICE","phone":"229-759-2001","balance":978.4,"futDue":978.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":200866,"shortName":"LEMUS","name":"LEMUS TIRE SHOP","phone":"229-417-4737","balance":252.3,"futDue":252.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200901,"shortName":"LENCHO","name":"LENCHO'S & SON TIRE SHOP","phone":"229-699-1259","balance":238.12,"futDue":238.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":201024,"shortName":"LIBERTY","name":"LIBERTY CAR WASH & TIRE","phone":"912-387-0408","balance":1.0,"futDue":1.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200687,"shortName":"LIVE","name":"LIVE OAK TIRE CENTER, LLC","phone":"386-362-1972","balance":338.24,"futDue":338.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200459,"shortName":"LOW","name":"LOW COUNTRY TIRE LLC","phone":"912-590-6331","balance":-0.02,"futDue":-0.02,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200352,"shortName":"MAL","name":"MALUDA AUTO SALES","phone":"229-244-9163","balance":-0.01,"futDue":-0.01,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Larry","custNum":200760,"shortName":"MALLARD","name":"MALLARD'S SERVICE CENTER","phone":"478-960-5246","balance":-82.09,"futDue":-82.09,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200761,"shortName":"MARIO","name":"MARIO NEW AND USED TIRE SHOP","phone":"229-591-5893","balance":-64.63,"futDue":-64.63,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":101519,"shortName":"MARK","name":"MARK TAYLOR DBA/MTAA ENT.","phone":"229-425-4601","balance":-0.16,"futDue":-0.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":200631,"shortName":"MARQ","name":"MARQUEZ TIRE SHOP","phone":"229-921-3897","balance":368.36,"futDue":368.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200965,"shortName":"MARQ","name":"MARQUEZ TIRE SHOP LLC","phone":"229-886-8657","balance":-505.8,"futDue":-505.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200964,"shortName":"MARTINEZ","name":"MARTINEZ AUTO SERVICE","phone":"972-801-7027","balance":-0.52,"futDue":-0.52,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-20"},{"salesman":"House","custNum":101524,"shortName":"MASTER","name":"MASTER BODY WORKS","phone":"229-439-8833","balance":88.33,"futDue":88.33,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"Larry","custNum":200409,"shortName":"MASTER","name":"MASTER CRAFT IND.(NO PASS/LT)","phone":"229-386-0610","balance":2499.95,"futDue":1326.21,"curDue":1173.74,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22"},{"salesman":"Larry","custNum":200468,"shortName":"MCKEE","name":"MCKEE'S AUTO CENTER, INC","phone":"229-382-7642","balance":742.42,"futDue":742.42,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"House","custNum":200193,"shortName":"MCLEAN","name":"MCLEAN TIRES INC","phone":"229-782-7428","balance":6045.22,"futDue":6045.22,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Tiffany","custNum":201005,"shortName":"MIDWAY","name":"MIDWAY ENTERPRISE FL, LLC","phone":"850-875-2444","balance":-159.42,"futDue":-159.42,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07"},{"salesman":"House","custNum":101208,"shortName":"MIKE","name":"MIKE FRASER AUTO REPAIR","phone":"229-273-0652","balance":-60.64,"futDue":-60.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"House","custNum":200622,"shortName":"MOORE","name":"MOORE'S ACCESSORIES & OFFROAD","phone":"229-223-6389","balance":-143.35,"futDue":-143.35,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-09"},{"salesman":"House","custNum":200355,"shortName":"MOTION","name":"MOTION WHEELS,HUBCAPS,&TIRES","phone":"229-253-0080","balance":-33.1,"futDue":-33.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-05"},{"salesman":"Larry","custNum":200510,"shortName":"MR","name":"M & R TRUCK ACCESSORIES","phone":"912-384-2362","balance":-0.3,"futDue":-0.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-28"},{"salesman":"Car Dealer","custNum":200688,"shortName":"NASH","name":"NASHVILLE FORD","phone":"229-686-2058","balance":-4.88,"futDue":-4.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-07-01"},{"salesman":"House","custNum":101549,"shortName":"NASHVILLE","name":"NASHVILLE TIRE","phone":"229-686-1900","balance":425.81,"futDue":425.81,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200356,"shortName":"NEEL","name":"NEELY'S SERVICE CENTER","phone":"229-263-4454","balance":308.5,"futDue":308.5,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-08"},{"salesman":"Tiffany","custNum":101146,"shortName":"NERO","name":"NE-RO TIRE & BRAKE SERVICE,INC","phone":"229-244-8353","balance":11001.01,"futDue":11001.01,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"House","custNum":201003,"shortName":"NICHOLAS","name":"NICHOLAS TIRES INC.","phone":"229-850-0289","balance":1277.67,"futDue":1277.67,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"","custNum":2000060,"shortName":"NISSAN","name":"NISSAN OF TIFTON","phone":"229-382-7777","balance":4201.94,"futDue":4201.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Larry","custNum":200755,"shortName":"NT","name":"N-T TIRE SERVICE","phone":"229-891-2234","balance":-47.56,"futDue":-47.56,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"Larry","custNum":201016,"shortName":"OC","name":"O&C AUTO REPAIR","phone":"229-821-1221","balance":-10.74,"futDue":-10.74,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"Car Dealer","custNum":200908,"shortName":"OSTEEN","name":"O'STEEN CHRYSLER DODGE JEEP","phone":"229-686-2068","balance":-44.89,"futDue":-44.89,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-17"},{"salesman":"House","custNum":101108,"shortName":"PARKER","name":"PARKER TIRE DIRECT","phone":"229-457-3818","balance":4198.67,"futDue":4198.67,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-21"},{"salesman":"Larry","custNum":200388,"shortName":"PEAR","name":"PEARSON TIRE & LUBE","phone":"912-422-6820","balance":-324.73,"futDue":-324.73,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":201023,"shortName":"PEASE","name":"PEASE ON THE GO 24/7","phone":"229-539-6995","balance":-55.66,"futDue":-55.66,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":101497,"shortName":"PERF","name":"PERFORMANCE MOTORSPORT","phone":"229-438-5248","balance":-3.06,"futDue":-3.06,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Larry","custNum":200214,"shortName":"PETERSON","name":"PETERSON TIRE & AUTO CENTER","phone":"229-352-5136","balance":330.61,"futDue":50.61,"curDue":280.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200885,"shortName":"PIERCE","name":"PIERCE INDUSTRIAL TIRE LLC","phone":"912-807-3685","balance":2367.16,"futDue":2367.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Larry","custNum":201017,"shortName":"PINEDA","name":"PINEDA'S AUTOMOTIVE","phone":"229-382-1583","balance":-267.52,"futDue":-267.52,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200890,"shortName":"PMT","name":"PMT TRK. TRAILER & TIRE REPAIR","phone":"229-457-5167","balance":-898.55,"futDue":-898.55,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200358,"shortName":"POND","name":"PONDER'S AUTOMOTIVE INC","phone":"229-228-5779","balance":556.0,"futDue":556.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"House","custNum":200658,"shortName":"POWE","name":"POWE AUTOMOTIVE","phone":"229-397-0459","balance":-223.82,"futDue":-223.82,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-01"},{"salesman":"Larry","custNum":200762,"shortName":"POWER","name":"POWER MAN TIRE SHOP","phone":"912-381-4065","balance":3470.61,"futDue":3269.41,"curDue":201.2,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200898,"shortName":"PRECISION","name":"PRECISION AUTO & MUFFLER LLC","phone":"386-364-1055","balance":144.84,"futDue":144.84,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-25"},{"salesman":"Larry","custNum":200920,"shortName":"PRECISION","name":"PRECISION MAINTENANCE","phone":"912-253-5237","balance":-395.26,"futDue":-395.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200213,"shortName":"PREMIER","name":"PREMIER AUTOWORKS","phone":"229-435-2886","balance":-406.0,"futDue":-406.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"Car Dealer","custNum":101374,"shortName":"PRINCE","name":"PRINCE CHEVY-OLDS, INC","phone":"229-386-4050","balance":4919.48,"futDue":4919.48,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"Car Dealer","custNum":200455,"shortName":"PRINCE","name":"PRINCE HONDA","phone":"229-386-1400","balance":161.98,"futDue":161.98,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15"},{"salesman":"Car Dealer","custNum":200456,"shortName":"PRINCE","name":"PRINCE TOYOTA","phone":"229-386-4052","balance":4890.48,"futDue":4890.48,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"Larry","custNum":101491,"shortName":"QUALITY","name":"QUALITY AUTO & R.V. SERVICE","phone":"229-273-0720","balance":411.2,"futDue":411.2,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22"},{"salesman":"House","custNum":200462,"shortName":"QUALITY","name":"QUALITY FEEDSTUFFS, INC","phone":"229-686-2770","balance":462.94,"futDue":462.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-23"},{"salesman":"House","custNum":200585,"shortName":"QUI","name":"QUINCY TIRE AND RECAPPING","phone":"850-627-6050","balance":7769.16,"futDue":7769.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Larry","custNum":101372,"shortName":"RAIN","name":"RAINEY USED CARS (BRONWOOD)","phone":"229-695-9153","balance":3549.26,"futDue":3549.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18"},{"salesman":"Tiffany","custNum":200362,"shortName":"RAY","name":"RAY NORTON TIRE & AUTO","phone":"229-247-1555","balance":3889.3,"futDue":3889.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200679,"shortName":"RICHARD","name":"RICHARD'S AUTO CARE & TIRE SVC","phone":"229-407-8600","balance":-47.98,"futDue":-47.98,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Larry","custNum":200288,"shortName":"RIGHT","name":"RIGHT CHOICE AUTO","phone":"","balance":-0.08,"futDue":-0.08,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-04-10"},{"salesman":"House","custNum":200241,"shortName":"RIM","name":"THE RIM SHOP INC","phone":"334-793-9292","balance":-13.64,"futDue":-13.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-26"},{"salesman":"House","custNum":200942,"shortName":"RNR","name":"RNR TIRE EXPRESS","phone":"813-977-9800","balance":13715.7,"futDue":13428.22,"curDue":287.48,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Car Dealer","custNum":101256,"shortName":"ROBERT","name":"ROBERT FENDER CHEVROLET","phone":"912-292-9005","balance":61.13,"futDue":61.13,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-08-27"},{"salesman":"Car Dealer","custNum":200365,"shortName":"ROBERT","name":"ROBERT HUTSON LINCOLN","phone":"229-985-6603","balance":500.68,"futDue":500.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"House","custNum":101025,"shortName":"ROCHELLE","name":"ROCHELLE TIRE","phone":"229-365-7943","balance":8541.6,"futDue":8541.6,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Tiffany","custNum":200883,"shortName":"ROLLING","name":"ROLLING BEAR TIRES LLC","phone":"912-387-6642","balance":246.66,"futDue":246.66,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200880,"shortName":"RR","name":"R&R TIRE CO.","phone":"229-805-4245","balance":-7.27,"futDue":-7.27,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":2000022,"shortName":"RR","name":"R&R AUTO SERVICE & REPAIR","phone":"912-347-5301","balance":-0.04,"futDue":-0.04,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22"},{"salesman":"House","custNum":201048,"shortName":"RRO","name":"RRO 24 HR ROADSIDE ASSISTANCE","phone":"386-965-0117","balance":-108.19,"futDue":-108.19,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"Larry","custNum":100282,"shortName":"RUDY","name":"RUDY'S TIRE SERVICE","phone":"229-382-5324","balance":26195.57,"futDue":18818.0,"curDue":7377.57,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200975,"shortName":"SANTOS","name":"SANTOS TIRE SHOP","phone":"229-921-6310","balance":-0.27,"futDue":-0.27,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200427,"shortName":"SAUNDERS","name":"SAUNDERS AUTO REPAIR","phone":"229-616-1041","balance":1372.76,"futDue":0.0,"curDue":1372.76,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-30"},{"salesman":"Larry","custNum":101463,"shortName":"SHELL","name":"SHELL RAPID LUBE (FITZGERALD)","phone":"229-424-9348","balance":20356.85,"futDue":20356.85,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200220,"shortName":"SING","name":"SINGLETARY & SON TIRE CO","phone":"229-776-5535","balance":7045.95,"futDue":7045.95,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Tiffany","custNum":200291,"shortName":"SING","name":"SINGLETARY TIRE PROS","phone":"229-226-2842","balance":317.85,"futDue":317.85,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-20"},{"salesman":"Larry","custNum":100907,"shortName":"SMITHS","name":"SMITH'S DIESEL REPAIR","phone":"229-273-1205","balance":3405.96,"futDue":0.0,"curDue":3405.96,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-27"},{"salesman":"Larry","custNum":100551,"shortName":"SOUTH","name":"SOUTHSIDE TIRE & AUTO SERVICE","phone":"229-387-6283","balance":40.79,"futDue":-107.95,"curDue":148.74,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":101181,"shortName":"SOUTH","name":"SOUTH GA LUBE CENTER","phone":"229-468-4435","balance":162.99,"futDue":162.99,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200621,"shortName":"SOUTH","name":"SOUTH GA LUBE (FITZGERALD)","phone":"229-345-1704","balance":831.89,"futDue":831.89,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200628,"shortName":"SOUTH","name":"SOUTH GEORGIA TIRE","phone":"229-668-8116","balance":144.0,"futDue":144.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200803,"shortName":"SOUTH","name":"SOUTH MAIN GARAGE","phone":"229-566-3880","balance":-0.3,"futDue":-0.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07"},{"salesman":"Larry","custNum":100967,"shortName":"SOUTHEAST","name":"SOUTHEASTERN COMMERCIAL TIRE","phone":"229-888-3300","balance":14737.98,"futDue":8619.56,"curDue":6118.42,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-10"},{"salesman":"Larry","custNum":200224,"shortName":"SOUTHERN","name":"SOUTHERN EXPRESS LUBE","phone":"229-777-0932","balance":702.94,"futDue":702.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14"},{"salesman":"House","custNum":200230,"shortName":"SOUTHERN","name":"SOUTHERN TIRE & BATTERY","phone":"229-246-4925","balance":377.06,"futDue":377.06,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-05-13"},{"salesman":"House","custNum":200742,"shortName":"SOUTHERN","name":"SOUTHERN SALES & RENTALS, LLC","phone":"706-546-9760","balance":613.9,"futDue":613.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11"},{"salesman":"House","custNum":200949,"shortName":"SOUTHERN","name":"SOUTHERN TIRE MART","phone":"229-920-3030","balance":6693.76,"futDue":5112.6,"curDue":1581.16,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-28"},{"salesman":"Larry","custNum":200959,"shortName":"SOUTHERN","name":"SOUTHERN GEORGIA TIRE LLC","phone":"912-292-0001","balance":-320.16,"futDue":-320.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":201063,"shortName":"SOUTHERN","name":"SOUTHERN TIRE EXPORTERS","phone":"404-819-5113","balance":-6.1,"futDue":-6.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-13"},{"salesman":"Tiffany","custNum":2000029,"shortName":"SOUTHERN","name":"SOUTHERN TIRE MART @ PILOT","phone":"229-244-3179","balance":2156.3,"futDue":2156.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20"},{"salesman":"House","custNum":200546,"shortName":"SOW","name":"SOWEGA TIRE OF ALBANY","phone":"229-888-1881","balance":304.59,"futDue":304.59,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":101588,"shortName":"STEPH","name":"STEPHENS BROTHERS","phone":"229-425-1055","balance":5427.22,"futDue":3410.98,"curDue":2016.24,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-01"},{"salesman":"Car Dealer","custNum":200232,"shortName":"SUNBELT","name":"SUNBELT FORD INC","phone":"229-776-7691","balance":-1014.5,"futDue":-1014.5,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-10-13"},{"salesman":"Tiffany","custNum":2000028,"shortName":"SUNPOINT","name":"SUNPOINT TIRES & ROAD SERVICE","phone":"863-272-8823","balance":478.56,"futDue":0.0,"curDue":478.56,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-06"},{"salesman":"","custNum":2000047,"shortName":"SUNPOINT","name":"SUNPOINT TIRES","phone":"863-272-8823","balance":166.32,"futDue":0.0,"curDue":166.32,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-26"},{"salesman":"Tiffany","custNum":200876,"shortName":"SUWANNEE","name":"SUWANNEE VALLEY TIRE","phone":"386-792-2420","balance":322.44,"futDue":322.44,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-12"},{"salesman":"House","custNum":101305,"shortName":"SYCAMORE","name":"SYCAMORE SALES & SALVAGE LLC","phone":"229-567-2005","balance":-14.36,"futDue":-14.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200712,"shortName":"TANNER","name":"TANNER AUTO REPAIR PLUS, LLC","phone":"912-807-8277","balance":-93.82,"futDue":-93.82,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-16"},{"salesman":"Larry","custNum":200490,"shortName":"TCA","name":"T.C.A. IRRIGATION","phone":"229-387-7097","balance":401.09,"futDue":401.09,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15"},{"salesman":"Larry","custNum":200956,"shortName":"TENNESON","name":"TENNESON COLLISION CENTER","phone":"229-391-9318","balance":403.19,"futDue":403.19,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"Larry","custNum":101415,"shortName":"THE","name":"THE TIRE STORE","phone":"229-686-2073","balance":22507.36,"futDue":22507.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"House","custNum":200756,"shortName":"THE","name":"THE SHOP OF ALBANY, LLC","phone":"229-573-7066","balance":-2.37,"futDue":-2.37,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"House","custNum":200991,"shortName":"THE","name":"THE TIRE CENTRE OF FLORIDA LLC","phone":"850-671-4181","balance":5967.12,"futDue":2967.12,"curDue":3000.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27"},{"salesman":"Larry","custNum":200242,"shortName":"THOMAS","name":"THOMAS TIRE COMPANY, LLC","phone":"229-985-1839","balance":407.85,"futDue":407.85,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-30"},{"salesman":"Larry","custNum":101326,"shortName":"TIFTON","name":"TIFTON GENERAL TIRE","phone":"229-382-6013","balance":31453.76,"futDue":31453.76,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200294,"shortName":"TIRE","name":"TIRE KING OF VALDOSTA","phone":"229-247-1345","balance":18672.34,"futDue":18672.34,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-24"},{"salesman":"Larry","custNum":201062,"shortName":"TIRE","name":"TIRE MASTERS LLC","phone":"229-445-7500","balance":2243.19,"futDue":572.29,"curDue":1670.9,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26"},{"salesman":"House","custNum":200537,"shortName":"TONY","name":"TONY'S TIRE & ROAD SERVICE,INC","phone":"229-890-9989","balance":0.8,"futDue":0.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Tiffany","custNum":200659,"shortName":"TOWN","name":"TOWN & COUNTRY TIRE","phone":"386-362-4535","balance":601.41,"futDue":601.41,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-14"},{"salesman":"Tiffany","custNum":200246,"shortName":"TRI","name":"TRI COUNTY TIRE COMPANY","phone":"229-524-2654","balance":9041.28,"futDue":3304.46,"curDue":5736.82,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21"},{"salesman":"Larry","custNum":200239,"shortName":"TS","name":"T & S TIRE","phone":"229-888-0696","balance":-4.0,"futDue":-4.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":101201,"shortName":"TUCKER","name":"TUCKERS SERVICE STATION","phone":"229-532-6097","balance":655.4,"futDue":655.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Larry","custNum":200601,"shortName":"TUFF","name":"TUFF ENTERPRISES LLC","phone":"229-883-8700","balance":2983.36,"futDue":2983.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"House","custNum":200624,"shortName":"WAL","name":"WALLACE MOTORS","phone":"850-973-1230","balance":565.97,"futDue":565.97,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Car Dealer","custNum":200461,"shortName":"WALKER","name":"WALKER JONES CHEVY-BUICK","phone":"912-490-1314","balance":1145.28,"futDue":1145.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200806,"shortName":"WALKERS","name":"WALKERS AUTO & OUTDOOR, INC","phone":"912-487-0084","balance":158.84,"futDue":158.84,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"Tiffany","custNum":101066,"shortName":"WARRIOR","name":"WARRIOR CREEK TIRE, LLC","phone":"229-798-0923","balance":3636.1,"futDue":3636.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"House","custNum":200132,"shortName":"WILKS","name":"WILKS A-ONE TIRE SALES","phone":"334-792-2225","balance":7193.07,"futDue":7193.07,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19"},{"salesman":"Tiffany","custNum":200383,"shortName":"WILL","name":"WILLIAMS ALIGNMENT & TIRE","phone":"229-263-4797","balance":-497.14,"futDue":-497.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29"},{"salesman":"Tiffany","custNum":200385,"shortName":"WILL","name":"WILLIAMS AUTOMOTIVE","phone":"229-224-6328","balance":92.0,"futDue":92.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28"},{"salesman":"House","custNum":200599,"shortName":"WL","name":"W&L TIRE & WHEEL CO. INC.","phone":"850-627-8830","balance":2347.15,"futDue":2347.15,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12"},{"salesman":"Tiffany","custNum":200595,"shortName":"WR","name":"W.R. WILLIAMS","phone":"386-294-1888","balance":30245.88,"futDue":12918.0,"curDue":17327.88,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15"},{"salesman":"House","custNum":200300,"shortName":"Z","name":"Z TIRE EXPRESS","phone":"229-244-2084","balance":-116.56,"futDue":-116.56,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-09-19"}];


// ── File Parsers ──────────────────────────────────────────────────────────────
function readSheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}

function parseWeekCompWorkbook(wb) {
  const sheetNames = wb.SheetNames;
  const weekSheets = sheetNames.filter(n => /^W\d+$/.test(n));

  // Parse each weekly sheet directly (works for WTD/QTD format)
  const BRANCHES_MAP = {"1 - BYRON":"Byron","2 - TIFTON":"Tifton","3 - STATESBORO":"Statesboro","5 - ATHENS":"Athens"};
  let weeks = [];
  weekSheets.forEach(sheetName => {
    const wnum = parseInt(sheetName.slice(1));
    if (!wnum || wnum > 52) return;
    const ws = wb.Sheets[sheetName];
    if (!ws) return;
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    const locData = {};
    rows.forEach(row => {
      const colB = row[1] ? String(row[1]).trim() : "";
      if (BRANCHES_MAP[colB]) {
        const branch = BRANCHES_MAP[colB];
        if (!locData[branch]) {
          locData[branch] = {
            sales2025: Number(row[2]||0), sales2026: Number(row[3]||0),
            gp2025: Number(row[6]||0),   gp2026:    Number(row[7]||0),
          };
        }
      }
    });
    const tif = locData["Tifton"] || {};
    const s25 = tif.sales2025 || 0;
    const s26 = tif.sales2026 || 0;
    if (s26 > 0) {
      weeks.push({
        week: wnum,
        sales2025: s25, sales2026: s26,
        change: s26-s25, changePct: s25 ? (s26-s25)/s25 : 0,
        gp2025: tif.gp2025||0, gp2026: tif.gp2026||0,
        gpChange: (tif.gp2026||0)-(tif.gp2025||0),
        locations: locData,
      });
    }
  });
  weeks.sort((a,b) => a.week - b.week);

  // Also try old "Week Analysis" sheet format
  const weeklyWs = wb.Sheets["Week Analysis"];
  if (weeklyWs && weeks.length === 0) {
    const rows = XLSX.utils.sheet_to_json(weeklyWs, { defval: null });
    weeks = rows.filter(r => r["Week"] && String(r["Week"]).startsWith("W")).map(r => ({
      week: parseInt(String(r["Week"]).slice(1)),
      sales2025: Number(r["2025 Sales"] || 0),
      sales2026: Number(r["2026 Sales"] || 0),
      change: Number(r["$ Change"] || 0),
      changePct: Number(r["% Change"] || 0),
      gp2025: Number(r["2025 GP"] || 0),
      gp2026: Number(r["2026 GP"] || 0),
    }));
  }

  // Dept analysis
  const deptWs = wb.Sheets["Dept Analysis"];
  let depts = [];
  if (deptWs) {
    const rows = XLSX.utils.sheet_to_json(deptWs, { defval: null });
    depts = rows.filter(r => r["Department"] && !String(r["Department"]).includes("EPD")).map(r => ({
      dept: String(r["Department"]),
      sales: Number(r["Total Sales"] || 0),
      gp: Number(r["Total GP $"] || 0),
      gpPct: Number(r["GP %"] || 0),
      lineItems: Number(r["Line Items"] || 0),
      assessment: String(r["Assessment"] || ""),
    }));
  }

  // Action plan
  const apWs = wb.Sheets["Action Plan"];
  let actionPlan = [];
  if (apWs) {
    const rows = XLSX.utils.sheet_to_json(apWs, { defval: null, range: 2 });
    actionPlan = rows.filter(r => r["Cust #"] || r["Cust#"] || r["Customer #"]).map(r => {
      const custNum   = r["Cust #"] || r["Cust#"] || r["Customer #"];
      const topRaw    = String(r["Top Dept (2026)"] || r["Top Dept"] || "");
      const decRaw    = String(r["Most Declined Dept"] || r["Most Declined"] || "");
      const city      = String(r["City"] || "");
      return {
        salesman:     String(r["Salesman"] || ""),
        city,
        custNum,
        customer:     String(r["Customer Name"] || r["Customer"] || ""),
        sales2025:    Number(r["2025 Sales"] || 0),
        sales2026:    Number(r["2026 Sales"] || 0),
        change:       Number(r["$ Change"] || 0),
        gpPct:        Number(r["GP %"] || 0),
        action:       String(r["Action Plan"] || ""),
        topDept:      topRaw.split("(")[0].trim(),
        declinedDept: decRaw.toLowerCase() === "none" ? "" : decRaw.split("(")[0].trim(),
        focus:        String(r["Dept Focus Areas"] || ""),
      };
    });
  }

  // Customers
  const custWs = wb.Sheets["Customer List"];
  let customers = [];
  if (custWs) {
    const rows = XLSX.utils.sheet_to_json(custWs, { defval: null });
    customers = rows.filter(r => r["CustomerNumber"]).map(r => ({
      num: r["CustomerNumber"],
      name: String(r["CompanyName"] || ""),
      salesman: String(r["Salesman"] || ""),
      active: String(r["Active"] || ""),
    }));
  }

  return { weeks, depts, actionPlan, customers };
}

function parseSalesSheet(rows) {
  const customers = {};
  for (const row of rows) {
    const name = row["Customer"] || row["Name"] || row["CustomerName"] || "";
    const rep  = row["Salesman"] || row["Rep"] || row["SalesRep"] || "";
    const dept = row["Department"] || row["Dept"] || "";
    const amt  = parseFloat(row["Amount"] || row["Sales"] || row["Total"] || 0);
    if (!name) continue;
    if (!customers[name]) customers[name] = { name, rep, depts: {}, total: 0 };
    customers[name].depts[dept] = (customers[name].depts[dept] || 0) + amt;
    customers[name].total += amt;
  }
  return Object.values(customers);
}

// ── Main App ──────────────────────────────────────────────────────────────────

// ── Login Screen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [selectedUser, setSelectedUser] = useState(null);
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);
  const inputRef = useRef(null);

  // Auto-login if remembered
  useEffect(() => {
    try {
      const saved = localStorage.getItem("pulse_remembered_user");
      if (saved) {
        const { userId } = JSON.parse(saved);
        const user = USERS.find(u => u.id === userId);
        if (user) onLogin(user);
      }
    } catch {}
  }, []);

  useEffect(() => {
    if (selectedUser && inputRef.current) inputRef.current.focus();
    // Pre-check remember me if this user was previously remembered
    try {
      const saved = localStorage.getItem("pulse_remembered_user");
      if (saved) {
        const { userId } = JSON.parse(saved);
        if (selectedUser?.id === userId) setRememberMe(true);
      }
    } catch {}
  }, [selectedUser]);

  async function handleLogin() {
    if (!selectedUser || !password) return;
    setLoading(true);
    setError("");
    try {
      const hashed = await hashPw(password);
      const user = USERS.find(u => u.id === selectedUser.id);
      if (hashed === user.hash) {
        if (rememberMe) {
          localStorage.setItem("pulse_remembered_user", JSON.stringify({ userId: user.id }));
        } else {
          localStorage.removeItem("pulse_remembered_user");
        }
        onLogin(user);
      } else {
        setError("Incorrect password. Please try again.");
        setPassword("");
      }
    } catch {
      setError("Login failed. Please try again.");
    }
    setLoading(false);
  }

  return (
    <div style={{ minHeight:"100vh", background:"#F4F7FB", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'Inter','Segoe UI',sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div style={{ width:"100%", maxWidth:420, padding:"0 1rem" }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:"2rem" }}>
          <div style={{ fontSize:"1.5rem", fontWeight:800, color:"#1E5FCC", letterSpacing:"0.15em", textTransform:"uppercase" }}>⬡ Pulse</div>
          <div style={{ fontSize:"0.72rem", color:"#6B7A99", letterSpacing:"0.15em", marginTop:4 }}>TIRES · AG · INDUSTRIAL</div>
        </div>

        <div style={{ background:"#FFFFFF", borderRadius:12, padding:"2rem", boxShadow:"0 4px 24px rgba(30,95,204,0.08)", border:"1px solid #D0DCF0" }}>
          {!selectedUser ? (
            <>
              <div style={{ fontSize:"0.78rem", fontWeight:700, color:"#1A2340", marginBottom:"1.25rem", textAlign:"center" }}>Who are you?</div>
              <div style={{ display:"flex", flexDirection:"column", gap:"0.6rem" }}>
                {USERS.map(u => (
                  <button key={u.id} onClick={() => { setSelectedUser(u); setError(""); }}
                    style={{ padding:"0.85rem 1rem", background:"#F4F7FB", border:`2px solid ${u.color}22`, borderRadius:8,
                      display:"flex", alignItems:"center", gap:12, cursor:"pointer", textAlign:"left", transition:"all 0.15s" }}
                    onMouseEnter={e => { e.currentTarget.style.background=u.color+"18"; e.currentTarget.style.borderColor=u.color; }}
                    onMouseLeave={e => { e.currentTarget.style.background="#F4F7FB"; e.currentTarget.style.borderColor=u.color+"22"; }}>
                    <div style={{ width:36, height:36, borderRadius:"50%", background:u.color, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:700, fontSize:"0.85rem", flexShrink:0 }}>
                      {u.name[0]}
                    </div>
                    <div>
                      <div style={{ fontSize:"0.85rem", fontWeight:700, color:"#1A2340" }}>{u.name}</div>
                      <div style={{ fontSize:"0.68rem", color:"#6B7A99" }}>{u.id==="admin" ? "Administrator" : "Sales Representative"}</div>
                    </div>
                  </button>
                ))}
              </div>
            </>
          ) : (
            <>
              <button onClick={() => { setSelectedUser(null); setPassword(""); setError(""); }}
                style={{ background:"none", border:"none", color:"#6B7A99", cursor:"pointer", fontSize:"0.72rem", marginBottom:"1rem", display:"flex", alignItems:"center", gap:4, padding:0 }}>
                ← Back
              </button>
              <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:"1.5rem" }}>
                <div style={{ width:44, height:44, borderRadius:"50%", background:selectedUser.color, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:700, fontSize:"1rem" }}>
                  {selectedUser.name[0]}
                </div>
                <div>
                  <div style={{ fontSize:"0.95rem", fontWeight:700, color:"#1A2340" }}>{selectedUser.name}</div>
                  <div style={{ fontSize:"0.7rem", color:"#6B7A99" }}>Enter your password</div>
                </div>
              </div>
              <form onSubmit={e => { e.preventDefault(); handleLogin(); }} method="post" action="#">
                <input
                  type="text"
                  name="username"
                  autoComplete="username"
                  value={selectedUser?.id || ""}
                  readOnly
                  style={{ position:"absolute", opacity:0, pointerEvents:"none", width:1, height:1 }}
                />
                <input
                  ref={inputRef}
                  type="password"
                  name="password"
                  autoComplete="current-password"
                  value={password}
                  onChange={e => { setPassword(e.target.value); setError(""); }}
                  placeholder="Password"
                  style={{ width:"100%", padding:"0.75rem 1rem", border:`2px solid ${error?"#DC2626":"#D0DCF0"}`, borderRadius:8, fontSize:"0.85rem", outline:"none", boxSizing:"border-box", marginBottom:"0.75rem", transition:"border 0.2s" }}
                  onFocus={e => e.target.style.borderColor=selectedUser.color}
                  onBlur={e => e.target.style.borderColor=error?"#DC2626":"#D0DCF0"}
                />
                {error && <div style={{ fontSize:"0.72rem", color:"#DC2626", marginBottom:"0.75rem" }}>⚠ {error}</div>}
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:"0.75rem" }}>
                  <input
                    type="checkbox"
                    id="rememberMe"
                    checked={rememberMe}
                    onChange={e => setRememberMe(e.target.checked)}
                    style={{ width:16, height:16, cursor:"pointer", accentColor:selectedUser.color }}
                  />
                  <label htmlFor="rememberMe" style={{ fontSize:"0.75rem", color:"#6B7A99", cursor:"pointer", userSelect:"none" }}>
                    Keep me signed in on this device
                  </label>
                </div>
                <button type="submit" disabled={loading || !password}
                  style={{ width:"100%", padding:"0.8rem", background:selectedUser.color, color:"#fff", border:"none", borderRadius:8,
                    fontSize:"0.85rem", fontWeight:700, cursor: loading||!password ? "not-allowed" : "pointer", opacity: loading||!password ? 0.7 : 1 }}>
                  {loading ? "Signing in…" : "Sign In"}
                </button>
              </form>
            </>
          )}
        </div>
        <div style={{ textAlign:"center", marginTop:"1.5rem", fontSize:"0.65rem", color:"#6B7A99" }}>
          Pulse Dashboard · Tire Distributors of GA
        </div>
      </div>
    </div>
  );
}

// ── Admin Tab ─────────────────────────────────────────────────────────────────

// ── Admin Notes & Todos View ──────────────────────────────────────────────────
function AdminNotesView() {
  const [view, setView]       = useState("todos");   // todos | notes
  const [repFilter, setRepFilter] = useState("all");
  const [loaded, setLoaded]   = useState(false);
  const [allTodos, setAllTodos] = useState([]);
  const [allNotes, setAllNotes] = useState([]);

  useEffect(() => {
    loadAll();
  }, []);

  async function loadAll() {
    // Load todos from Supabase — all reps
    try {
      const rows = await sbFetch("rep_todos?order=created_at.desc&limit=500");
      if (rows) setAllTodos(rows);
    } catch {}

    // Load notes from Supabase — all reps
    try {
      const noteRows = await sbFetch("rep_notes?order=updated_at.desc&limit=1000");
      if (noteRows && noteRows.length > 0) {
        setAllNotes(noteRows.map(r => ({
          custNum:  r.cust_num,
          userId:   r.user_id,
          text:     r.notes,
          updatedAt: r.updated_at,
        })));
      } else {
        // Fallback: scan admin's localStorage
        const notes = [];
        try {
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key && key.startsWith("notes_")) {
              const text = localStorage.getItem(key);
              if (text && text.trim()) notes.push({ custNum: key.replace("notes_",""), text });
            }
          }
        } catch {}
        setAllNotes(notes);
      }
    } catch {}
    setLoaded(true);
  }

  // Group todos by rep
  const reps = [...new Set(allTodos.map(t => t.salesman || t.user_id || "Unknown"))].filter(Boolean).sort();

  const filteredTodos = allTodos.filter(t =>
    repFilter === "all" || t.salesman === repFilter || t.user_id === repFilter
  );
  const filteredNotes = allNotes;

  const openCount  = allTodos.filter(t=>!t.done).length;
  const doneCount  = allTodos.filter(t=> t.done).length;
  const notesCount = allNotes.filter(n=>n.text).length;

  return (
    <div style={{ ...S.card, marginBottom:"1rem" }}>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.75rem", flexWrap:"wrap", gap:8 }}>
        <div style={{ fontSize:"0.72rem", fontWeight:700, color:"#1E5FCC", textTransform:"uppercase", letterSpacing:"0.1em" }}>
          📝 All Rep Notes & To-Dos
        </div>
        <div style={{ display:"flex", gap:5 }}>
          <button onClick={loadAll} style={{ fontSize:"0.65rem", color:MUTED, background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"0.25rem 0.6rem", cursor:"pointer" }}>↺ Refresh</button>
        </div>
      </div>

      {/* KPI strip */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.5rem", marginBottom:"0.85rem" }}>
        {[
          { label:"Open To-Dos",  val:openCount,  color:openCount>0?RED:GREEN },
          { label:"Completed",    val:doneCount,  color:GREEN },
          { label:"Total To-Dos", val:allTodos.length, color:MUTED },
          { label:"Cust. Notes",  val:notesCount, color:AMBER },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.5rem 0.6rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.9rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.58rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.05em", marginTop:1 }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* View + rep filter */}
      <div style={{ display:"flex", gap:6, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
        {[["todos","📋 To-Dos"],["notes","📝 Notes"]].map(([v,l])=>(
          <button key={v} onClick={()=>setView(v)}
            style={{ fontSize:"0.7rem", fontWeight:view===v?700:400, color:view===v?"#fff":MUTED,
              background:view===v?"#1E5FCC":"#F4F7FB", border:`1px solid ${view===v?"#1E5FCC":BORDER}`,
              borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer" }}>{l}</button>
        ))}
        {view==="todos" && reps.length > 1 && (
          <select value={repFilter} onChange={e=>setRepFilter(e.target.value)}
            style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
              padding:"0.3rem 0.5rem", borderRadius:4, fontSize:"0.7rem" }}>
            <option value="all">All Reps</option>
            {reps.map(r=><option key={r} value={r}>{r}</option>)}
          </select>
        )}
        <span style={{ marginLeft:"auto", fontSize:"0.65rem", color:MUTED }}>
          {view==="todos" ? `${filteredTodos.length} to-dos` : `${filteredNotes.length} customers with notes`}
        </span>
      </div>

      {!loaded && <div style={{ fontSize:"0.75rem", color:MUTED, textAlign:"center", padding:"1rem" }}>Loading…</div>}

      {/* To-Dos view */}
      {view === "todos" && loaded && (
        <div style={{ maxHeight:400, overflowY:"auto" }}>
          {filteredTodos.length === 0
            ? <div style={{ fontSize:"0.75rem", color:MUTED, textAlign:"center", padding:"1.5rem" }}>No to-dos found</div>
            : filteredTodos.map((t, i) => (
              <div key={i} style={{ display:"flex", gap:10, padding:"0.45rem 0",
                borderBottom: i<filteredTodos.length-1?`1px solid ${BORDER}`:"none",
                alignItems:"flex-start", opacity:t.done?0.6:1 }}>
                <span style={{ fontSize:"0.85rem", flexShrink:0, marginTop:1 }}>{t.done?"☑":"☐"}</span>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:"0.73rem", color:TEXT, textDecoration:t.done?"line-through":"none", wordBreak:"break-word" }}>{t.text}</div>
                  <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, display:"flex", gap:8, flexWrap:"wrap" }}>
                    <span style={{ fontWeight:600, color:AMBER }}>{t.cust_name || `#${t.cust_num}`}</span>
                    {t.salesman && <span>{t.salesman}</span>}
                    <span>{t.created_by} · {t.created_date || new Date(t.created_at||"").toLocaleDateString()}</span>
                    {t.done && <span style={{ color:GREEN, fontWeight:600 }}>✓ Done</span>}
                  </div>
                </div>
              </div>
            ))}
        </div>
      )}

      {/* Notes view */}
      {view === "notes" && loaded && (
        <div style={{ maxHeight:400, overflowY:"auto" }}>
          {filteredNotes.length === 0
            ? <div style={{ fontSize:"0.75rem", color:MUTED, textAlign:"center", padding:"1.5rem" }}>No notes found</div>
            : filteredNotes.map((n, i) => (
              <div key={i} style={{ padding:"0.55rem 0", borderBottom: i<filteredNotes.length-1?`1px solid ${BORDER}`:"none" }}>
                <div style={{ display:"flex", gap:8, alignItems:"center", marginBottom:3, flexWrap:"wrap" }}>
                  <span style={{ fontSize:"0.68rem", fontWeight:700, color:AMBER }}>Customer #{n.custNum}</span>
                  {n.userId && <span style={{ fontSize:"0.65rem", color:"#7C3AED", background:"#EDE9FE", borderRadius:8, padding:"1px 6px", fontWeight:600 }}>{n.userId}</span>}
                  {n.updatedAt && <span style={{ fontSize:"0.62rem", color:MUTED }}>{new Date(n.updatedAt).toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>}
                </div>
                <div style={{ fontSize:"0.74rem", color:TEXT, lineHeight:1.65, whiteSpace:"pre-wrap", wordBreak:"break-word" }}>{n.text}</div>
              </div>
            ))}
        </div>
      )}
    </div>
  );
}

function AdminTab({ currentUser, leads, onAddLead, onDeleteLead, convertLead }) {
  const [log, setLog] = useState([]);
  const [aiSuggestions, setAiSuggestions] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [filter, setFilter] = useState("all");
  const [loaded, setLoaded] = useState(false);
  const [inactiveList, setInactiveList] = useState({});
  const [suggestions, setSuggestions] = useState([]);

  function refreshLog() {
    try {
      const logRes = localStorage.getItem("shared_activity_log");
      setLog(logRes ? JSON.parse(logRes) : []);
      const inactRes = localStorage.getItem("inactive_customers");
      setInactiveList(inactRes ? JSON.parse(inactRes) : {});
      const sugRes = localStorage.getItem("pulse_suggestions");
      setSuggestions(sugRes ? JSON.parse(sugRes) : []);
    } catch {}
    setLoaded(true);
  }

  async function refreshFromSupabase() {
    // Pull ALL rep activity from Supabase — works across devices
    const rows = await syncActivityDown();
    if (rows.length > 0) {
      const mapped = rows.map(r => ({
        ts: r.ts, user: r.user_name, userId: r.user_id,
        action: r.action, detail: r.detail,
      }));
      // Merge with local, dedupe by ts+user+action
      const local = JSON.parse(localStorage.getItem("shared_activity_log") || "[]");
      const seen = new Set(local.map(e => e.ts + e.userId + e.action));
      const merged = [...local, ...mapped.filter(e => !seen.has(e.ts + e.userId + e.action))]
        .sort((a,b) => new Date(b.ts) - new Date(a.ts))
        .slice(0, 1000);
      setLog(merged);
      localStorage.setItem("shared_activity_log", JSON.stringify(merged));
    }
    setLoaded(true);
  }

  const [allLeads, setAllLeads] = useState(leads || []);

  useEffect(() => {
    refreshLog();
    refreshFromSupabase();
    // Pull all leads from Supabase
    syncAllLeadsDown().then(rows => {
      if (rows && rows.length > 0) setAllLeads(rows);
      else setAllLeads(leads || []);
    });
  }, []);

  // Keep in sync with prop
  useEffect(() => { if (!allLeads.length) setAllLeads(leads || []); }, [leads]);

  const users = [...new Set(log.map(e => e.user))].sort();
  const filtered = filter === "all" ? log : log.filter(e => e.user === filter);

  // Stats per user
  const stats = {};
  log.forEach(e => {
    if (!stats[e.user]) stats[e.user] = { logins:0, tabs:0, customers:0, lastSeen:"" };
    if (e.action === "login")         stats[e.user].logins++;
    if (e.action === "view_tab")      stats[e.user].tabs++;
    if (e.action === "view_customer") stats[e.user].customers++;
    if (!stats[e.user].lastSeen || e.ts > stats[e.user].lastSeen) stats[e.user].lastSeen = e.ts;
  });

  async function generateSuggestions() {
    setAiLoading(true);
    const logSummary = log.slice(0,100).map(e =>
      `${new Date(e.ts).toLocaleDateString()} ${e.user}: ${e.action} — ${e.detail}`
    ).join("\n");
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST", headers:{
          "Content-Type": "application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model:"claude-sonnet-4-6", max_tokens:1000,
          messages:[{ role:"user", content:`You are a sales manager analyzing rep activity in a tire & ag distribution company dashboard.

Activity log (most recent first):
${logSummary}

Based on this activity data:
1. Which reps are most/least active in the dashboard?
2. What customers are being researched most?
3. What coaching suggestions do you have for each rep?
4. Any patterns worth noting?

Keep it concise and actionable, written for a sales manager.` }]
        })
      });
      if (!res.ok) throw new Error(`API error ${res.status}: ${await res.text()}`);
      const data = await res.json();
      setAiSuggestions(data.content?.[0]?.text || "No suggestions generated.");
    } catch(e) { setAiSuggestions(`Error: ${e.message || 'Could not connect. Check API key.'}`); }
    setAiLoading(false);
  }

  async function clearLog() {
    if (!confirm("Clear all activity logs? This cannot be undone.")) return;
    localStorage.setItem("shared_activity_log", JSON.stringify([]));
    setLog([]);
  }

  function formatTime(ts) {
    const d = new Date(ts);
    return d.toLocaleDateString("en-US",{month:"short",day:"numeric"}) + " " + d.toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit"});
  }

  const actionIcon = a => a==="login"?"🔑":a==="logout"?"🚪":a==="view_customer"?"👤":a==="mark_inactive"?"⊘":a==="mark_active"?"✓":a==="bug_report"?"🐛":a==="suggestion"?"💡":a==="new_lead"?"🎯":a==="convert_lead"?"✓":a==="add_todo"?"☐":a==="update_notes"?"📝":a==="call_note"?"📞":"📋";
  const actionColor = a => a==="login"?GREEN:a==="logout"?MUTED:a==="view_customer"?AMBER:a==="mark_inactive"?RED:a==="mark_active"?GREEN:a==="bug_report"?RED:a==="suggestion"?"#7C3AED":a==="new_lead"?"#059669":a==="add_todo"?AMBER:a==="update_notes"?TEAL:a==="call_note"?"#0891B2":TEAL;

  return (
    <div>
      {/* Stats row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))", gap:"0.75rem", marginBottom:"1rem" }}>
        {Object.entries(stats).map(([user,s]) => {
          const u = USERS.find(x=>x.name===user)||{color:MUTED};
          return (
            <div key={user} style={{ ...S.card, borderTop:`4px solid ${u.color}`, marginBottom:0, padding:"0.75rem 1rem" }}>
              <div style={{ fontSize:"0.8rem", fontWeight:700, color:u.color, marginBottom:6 }}>{user}</div>
              <div style={{ fontSize:"0.68rem", color:MUTED, display:"flex", flexDirection:"column", gap:3 }}>
                <span>🔑 {s.logins} login{s.logins!==1?"s":""}</span>
                <span>📋 {s.tabs} tab view{s.tabs!==1?"s":""}</span>
                <span>👤 {s.customers} customer{s.customers!==1?"s":""} opened</span>
                {s.lastSeen && <span style={{ color:AMBER, fontWeight:600 }}>Last: {formatTime(s.lastSeen)}</span>}
              </div>
            </div>
          );
        })}
      </div>

      {/* AI Suggestions */}
      <div style={{ ...S.card, marginBottom:"1rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom: aiSuggestions?"0.75rem":"0" }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:AMBER, textTransform:"uppercase", letterSpacing:"0.1em" }}>◈ AI Manager Suggestions</div>
          <button onClick={generateSuggestions} disabled={aiLoading||log.length===0} style={{ ...S.btn(AMBER), fontSize:"0.7rem" }}>
            {aiLoading?"Analyzing…":"Generate Suggestions"}
          </button>
        </div>
        {aiSuggestions && (
          <div style={{ fontSize:"0.78rem", color:TEXT, lineHeight:1.8, whiteSpace:"pre-wrap" }}>{aiSuggestions}</div>
        )}
        {!aiSuggestions && !aiLoading && log.length === 0 && (
          <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:8 }}>No activity logged yet — suggestions will appear once reps start using the dashboard.</div>
        )}
      </div>

      {/* Inactive Accounts summary */}
      {Object.keys(inactiveList||{}).length > 0 && (
        <div style={{ ...S.card, marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:RED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>
            ⊘ Inactive Accounts <span style={{ color:MUTED, fontWeight:400 }}>({Object.keys(inactiveList).length})</span>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
            {Object.entries(inactiveList).map(([custNum, info]) => (
              <div key={custNum} style={{ display:"flex", gap:12, padding:"0.45rem 0.6rem", background:"#FEF2F2", borderRadius:6, alignItems:"flex-start" }}>
                <span style={{ fontSize:"0.9rem" }}>⊘</span>
                <span style={{ fontSize:"0.7rem", fontWeight:700, color:RED, minWidth:60 }}>#{custNum}</span>
                <span style={{ fontSize:"0.72rem", color:TEXT, flex:1 }}>{info.name}</span>
                <span style={{ fontSize:"0.68rem", color:MUTED, whiteSpace:"nowrap" }}>by {info.by} on {info.date}</span>
                <span style={{ fontSize:"0.68rem", color:"#991B1B", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{info.reason}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Leads Management */}
      <div style={{ ...S.card, marginBottom:"1rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.75rem" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div style={{ fontSize:"0.72rem", fontWeight:700, color:"#059669", textTransform:"uppercase", letterSpacing:"0.1em" }}>
              🎯 Leads <span style={{ color:MUTED, fontWeight:400 }}>({allLeads.filter(l=>l.status==="open").length} open · {allLeads.length} total)</span>
            </div>
            <button onClick={()=>syncAllLeadsDown().then(rows=>{if(rows&&rows.length>0)setAllLeads(rows);})}
              style={{ fontSize:"0.65rem", color:"#059669", background:"#F0FDF4", border:"1px solid #BBF7D0",
                borderRadius:4, padding:"0.2rem 0.6rem", cursor:"pointer" }}>
              ↺ Refresh
            </button>
          </div>
        </div>
        <LeadsTab leads={allLeads} repName="Admin" onAddLead={onAddLead} onDeleteLead={onDeleteLead} currentUser={currentUser} isAdmin={true} allReps={["Tiffany","Larry","Austin"]} convertLead={convertLead} />
      </div>

      {/* All Customer Notes & To-Dos */}
      <AdminNotesView />

      {/* Suggestions & Bug Reports */}
      {suggestions.length > 0 && (
        <div style={{ ...S.card, marginBottom:"1rem" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.75rem" }}>
            <div style={{ fontSize:"0.72rem", fontWeight:700, color:"#7C3AED", textTransform:"uppercase", letterSpacing:"0.1em" }}>
              💡 Suggestions & Bug Reports <span style={{ color:MUTED, fontWeight:400 }}>({suggestions.length})</span>
            </div>
            <button onClick={()=>{ localStorage.removeItem("pulse_suggestions"); setSuggestions([]); }}
              style={{ fontSize:"0.65rem", color:RED, background:"#FEF2F2", border:`1px solid #FECACA`, borderRadius:4, padding:"0.25rem 0.5rem", cursor:"pointer" }}>
              Clear All
            </button>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
            {suggestions.map((s,i) => (
              <div key={i} style={{ padding:"0.5rem 0.75rem", background: s.type==="bug"?"#FEF2F2":"#F5F3FF", border:`1px solid ${s.type==="bug"?"#FECACA":"#DDD6FE"}`, borderRadius:6 }}>
                <div style={{ display:"flex", gap:10, alignItems:"center", marginBottom:3 }}>
                  <span style={{ fontSize:"0.75rem" }}>{s.type==="bug"?"🐛":"💡"}</span>
                  <span style={{ fontSize:"0.7rem", fontWeight:700, color: s.type==="bug"?RED:"#7C3AED" }}>{s.type==="bug"?"BUG REPORT":"SUGGESTION"}</span>
                  <span style={{ fontSize:"0.65rem", color:MUTED }}>{s.user} · {new Date(s.ts).toLocaleDateString()}</span>
                </div>
                <div style={{ fontSize:"0.75rem", color:TEXT, paddingLeft:26 }}>{s.text}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Activity log */}
      <div style={S.card}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.75rem", flexWrap:"wrap", gap:8 }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:TEXT, textTransform:"uppercase", letterSpacing:"0.1em" }}>
            Activity Log <span style={{ color:MUTED, fontWeight:400 }}>({filtered.length} entries)</span>
          </div>
          <div style={{ display:"flex", gap:6, alignItems:"center" }}>
            <button onClick={()=>{ refreshLog(); refreshFromSupabase(); }} style={{ fontSize:"0.68rem", color:AMBER, background:"#EEF4FF", border:`1px solid ${BORDER}`, borderRadius:4, padding:"0.3rem 0.65rem", cursor:"pointer" }}>↺ Refresh</button>
            <select value={filter} onChange={e=>setFilter(e.target.value)}
              style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.3rem 0.5rem", borderRadius:4, fontSize:"0.7rem" }}>
              <option value="all">All Users</option>
              {users.map(u=><option key={u} value={u}>{u}</option>)}
            </select>
            <button onClick={clearLog} style={{ fontSize:"0.68rem", color:RED, background:"#FEF2F2", border:`1px solid #FECACA`, borderRadius:4, padding:"0.3rem 0.65rem", cursor:"pointer" }}>
              Clear Log
            </button>
          </div>
        </div>
        {!loaded && <div style={{ color:MUTED, fontSize:"0.75rem", textAlign:"center", padding:"1rem" }}>Loading…</div>}
        {loaded && filtered.length === 0 && (
          <div style={{ color:MUTED, fontSize:"0.75rem", textAlign:"center", padding:"1.5rem" }}>No activity recorded yet.</div>
        )}
        <div style={{ maxHeight:420, overflowY:"auto" }}>
          {filtered.map((e, i) => {
            const u = USERS.find(x=>x.name===e.user)||{color:MUTED};
            return (
              <div key={i} style={{ display:"flex", gap:10, padding:"0.45rem 0", borderBottom: i<filtered.length-1?`1px solid ${BORDER}`:"none", alignItems:"flex-start" }}>
                <span style={{ fontSize:"0.9rem", flexShrink:0 }}>{actionIcon(e.action)}</span>
                <span style={{ fontSize:"0.65rem", color:MUTED, whiteSpace:"nowrap", minWidth:110, marginTop:1 }}>{formatTime(e.ts)}</span>
                <span style={{ fontSize:"0.7rem", fontWeight:700, color:u.color, minWidth:60, marginTop:1 }}>{e.user}</span>
                <span style={{ fontSize:"0.72rem", color:actionColor(e.action), fontWeight:600, minWidth:55, marginTop:1 }}>{e.action.replace("_"," ")}</span>
                <span style={{ fontSize:"0.72rem", color:TEXT, flex:1, marginTop:1 }}>{e.detail}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [currentUser, setCurrentUser] = useState(null); // logged-in user object
  const [tab, setTab] = useState("overview");
  const [fileData, setFileData] = useState({ weekComp: SEED_WEEK_COMP, customers: SEED_CUSTOMERS, ar: SEED_AR });
  const [notice, setNotice] = useState("");
  const [aiPrompt, setAiPrompt] = useState("");
  const [custTabs, setCustTabs] = useState([]); // open customer detail tabs, max 10
  const [inactiveCustomers, setInactiveCustomers] = useState({}); // custNum → {reason, date, by}
  const [leads, setLeads] = useState([]); // array of lead objects

  // Load leads from storage + auto-sync
  useEffect(() => {
    // Load from localStorage first (instant)
    try {
      const saved = localStorage.getItem("pulse_leads");
      if (saved) setLeads(JSON.parse(saved));
    } catch {}
    // Always sync from Supabase after login
    if (currentUser) {
      refreshLeads(false);

      // Auto-refresh every 5 minutes
      const interval = setInterval(() => refreshLeads(false), 5 * 60 * 1000);

      // Also refresh when user switches back to this tab
      const handleVisibility = () => {
        if (document.visibilityState === "visible") refreshLeads(false);
      };
      document.addEventListener("visibilitychange", handleVisibility);

      return () => {
        clearInterval(interval);
        document.removeEventListener("visibilitychange", handleVisibility);
      };
    }
  }, [currentUser]);

  function saveLeads(updated) {
    setLeads(updated);
    localStorage.setItem("pulse_leads", JSON.stringify(updated));
  }

  async function refreshLeads(showAlert = false) {
    if (!currentUser) return;
    try {
      const rows = currentUser.id === "admin"
        ? await syncAllLeadsDown()
        : await syncLeadsDown(currentUser.id);
      if (rows && rows.length > 0) {
        setLeads(rows);
        localStorage.setItem("pulse_leads", JSON.stringify(rows));
        if (showAlert) alert(`✓ Synced ${rows.filter(l=>l.status==="open").length} open lead${rows.filter(l=>l.status==="open").length!==1?"s":""} from cloud`);
      } else {
        if (showAlert) alert("No leads found assigned to you yet — check with your manager.");
      }
    } catch(e) {
      if (showAlert) alert(`Sync failed: ${e.message}`);
    }
  }

  async function addLead(lead) {
    const assignedName = lead.assigned_to_name || lead.assigned_to || currentUser?.name || "";
    const assignedId   = assignedName.toLowerCase();
    const newLead = {
      id: `lead_${Date.now()}`,
      ...lead,
      assigned_to:      assignedId,
      assigned_to_name: assignedName,
      status: "open",
      created_by:       currentUser?.id || "",
      created_by_name:  currentUser?.name || "",
      created_at:       new Date().toISOString(),
    };
    const updated = [newLead, ...leads];
    saveLeads(updated);
    logActivity("new_lead", `${lead.name} — assigned to ${assignedName}`);
    // Push to Supabase — rep will pull this on next load/refresh
    await syncLeadUp(newLead);
    // Write a local notification flag so rep sees it immediately if on same device
    try {
      const notifyKey = `pulse_new_leads_${assignedId}`;
      const existing  = JSON.parse(localStorage.getItem(notifyKey) || "[]");
      localStorage.setItem(notifyKey, JSON.stringify([...existing, newLead.id]));
    } catch {}
    return newLead;
  }

  async function convertLead(leadId) {
    const updated = leads.map(l => l.id===leadId ? {...l, status:"converted"} : l);
    saveLeads(updated);
    await updateLeadUp(leadId, { status: "converted" });
    logActivity("convert_lead", leads.find(l=>l.id===leadId)?.name || leadId);
  }

  async function deleteLead(leadId) {
    const updated = leads.filter(l => l.id!==leadId);
    saveLeads(updated);
    await deleteLeadUp(leadId);
  }

  // Load inactive customers from storage on mount
  useEffect(() => {
    async function loadInactive() {
      try {
        const res = localStorage.getItem("inactive_customers");
        if (res) setInactiveCustomers(JSON.parse(res));
      } catch {}
    }
    loadInactive();
  }, []);

  async function markInactive(custNum, custName, reason, markedBy) {
    const entry = { reason, date: new Date().toISOString().slice(0,10), by: markedBy, name: custName };
    const updated = { ...inactiveCustomers, [String(custNum)]: entry };
    setInactiveCustomers(updated);
    localStorage.setItem("inactive_customers", JSON.stringify(updated));
    logActivity("mark_inactive", `${custName} — ${reason}`);
  }

  async function markActive(custNum, custName) {
    const updated = { ...inactiveCustomers };
    delete updated[String(custNum)];
    setInactiveCustomers(updated);
    localStorage.setItem("inactive_customers", JSON.stringify(updated));
    logActivity("mark_active", `${custName} reactivated`);
  }
  const [uploadDates, setUploadDates] = useState({}); // {slotId: ISO timestamp}
  const [dismissedWarning, setDismissedWarning] = useState(false);
  const [newWeeksDetected, setNewWeeksDetected] = useState([]);

  // Weekly upload slots that must be refreshed every week
  const WEEKLY_SLOTS = ["weekComp", "ar", "sales"];

  // Load upload dates from storage on mount — pre-seeded with 5/17/2026 upload
  const SEED_UPLOAD_DATES = {"weekComp":"2020-01-01T00:00:00.000Z","ar":"2020-01-01T00:00:00.000Z","sales":"2020-01-01T00:00:00.000Z"};
  useEffect(() => {
    async function loadDates() {
      try {
        const res=(()=>{ const v=localStorage.getItem("upload_dates"); return v?{value:v}:null; })();
        if (res?.value) {
          const stored = JSON.parse(res.value);
          const merged = {};
          ["weekComp","ar","sales"].forEach(slot => {
            const seeded = new Date(SEED_UPLOAD_DATES[slot]);
            const storedDate = stored[slot] ? new Date(stored[slot]) : new Date(0);
            merged[slot] = seeded > storedDate ? SEED_UPLOAD_DATES[slot] : stored[slot];
          });
          setUploadDates(merged);
        } else {
          setUploadDates(SEED_UPLOAD_DATES);
          localStorage.setItem("upload_dates", JSON.stringify(SEED_UPLOAD_DATES));
        }
      } catch {
        setUploadDates(SEED_UPLOAD_DATES);
      }
    }
    loadDates();
  }, []);

  // Compute warning state
  const warningInfo = (() => {
    const now = new Date();
    const dayOfWeek = now.getDay(); // 0=Sun
    // Find the most recent Sunday (start of current week)
    const lastSunday = new Date(now);
    lastSunday.setDate(now.getDate() - dayOfWeek);
    lastSunday.setHours(0, 0, 0, 0);

    const staleSlots = WEEKLY_SLOTS.filter(slotId => {
      const uploaded = uploadDates[slotId];
      if (!uploaded) return true;
      return new Date(uploaded) < lastSunday;
    });

    const isSundayOrLater = dayOfWeek === 0; // only warn on Sunday
    const isOverdue       = dayOfWeek === 0 && staleSlots.length > 0;
    const isPending       = dayOfWeek !== 0 && staleSlots.length > 0; // reminder any day

    return { staleSlots, isOverdue, isPending, isSundayOrLater, lastSunday };
  })();

  // Activity logger — writes to shared storage for admin view
  function logActivity(action, detail, userOverride) {
    const u = userOverride || currentUser;
    if (!u) return;
    try {
      const entry = { ts: new Date().toISOString(), user: u.name, userId: u.id, action, detail };
      // Write to localStorage (fast, local)
      let existing = [];
      try { const rv = localStorage.getItem("shared_activity_log"); if(rv) existing = JSON.parse(rv); } catch {}
      localStorage.setItem("shared_activity_log", JSON.stringify([entry, ...existing].slice(0, 1000)));
      // Push to Supabase (so admin sees ALL reps from any device)
      syncActivityUp(entry);
    } catch {}
  }

  function openCustomer(ap) {
    setCustTabs(prev => {
      const exists = prev.find(t => t.custNum === ap.custNum);
      if (exists) { setTab(`cust_${ap.custNum}`); return prev; }
      if (prev.length >= 10) { alert("Max 10 customer tabs open. Close one first."); return prev; }
      setTab(`cust_${ap.custNum}`);
      logActivity("view_customer", ap.customer);
      return [...prev, ap];
    });
  }
  function closeCustomer(custNum) {
    setCustTabs(prev => {
      const next = prev.filter(t => t.custNum !== custNum);
      if (tab === `cust_${custNum}`) setTab(next.length ? `cust_${next[next.length-1].custNum}` : "overview");
      return next;
    });
  }

  async function handleUpload(slotId, wb) {
    let parsed;

    // ── MASTER WORKBOOK DETECTION ─────────────────────────────────────────
    if (isMasterWorkbook(wb)) {
      const m = parseMasterWorkbook(wb);
      const summary = [];

      // AR
      if (m.ar && m.ar.length > 0) {
        setFileData(prev => ({ ...prev, ar: m.ar }));
        summary.push(`AR: ${m.ar.length} accounts`);
      }

      // CustomerComp → merge into actionPlan
      if (m.customerComp?.actionPlan?.length > 0) {
        setFileData(prev => {
          const existing = prev.weekComp || SEED_WEEK_COMP;
          let mergedAP = [...(existing.actionPlan || [])];
          m.customerComp.actionPlan.forEach(newAP => {
            const idx = mergedAP.findIndex(a => String(a.custNum) === String(newAP.custNum));
            // Preserve salesman/city from existing customer list
            const existingEntry = idx >= 0 ? mergedAP[idx] : null;
            const merged = {
              ...newAP,
              city:     existingEntry?.city     || newAP.city     || "",
              salesman: existingEntry?.salesman  || newAP.salesman || "House",
            };
            if (idx >= 0) mergedAP[idx] = merged;
            else mergedAP.push(merged);
          });
          return { ...prev, weekComp: { ...existing, actionPlan: mergedAP } };
        });
        summary.push(`CustomerComp: ${m.customerComp.actionPlan.length} accounts`);
      }

      // Statesboro
      if (m.statesboro?.actionPlan?.length > 0) {
        setFileData(prev => ({ ...prev, statesboroAP: m.statesboro.actionPlan }));
        summary.push(`Statesboro: ${m.statesboro.actionPlan.length} accounts`);
      }

      // AD Programs — merge into existing adPrograms state
      const adUpdates = {};
      if (m.toyo)     { adUpdates.toyoData     = m.toyo;     summary.push("Toyo ✓"); }
      if (m.americus) { adUpdates.americusData  = m.americus; summary.push("Americus ✓"); }
      if (m.ascenso)  { adUpdates.ascensoData   = m.ascenso;  summary.push("Ascenso ✓"); }
      if (m.falkenPLT){ adUpdates.falkenPLT     = m.falkenPLT; summary.push(`Falken PLT: ${m.falkenPLT.length}`); }
      if (m.falkenTBR){ adUpdates.falkenTBR     = m.falkenTBR; summary.push(`Falken TBR: ${m.falkenTBR.length}`); }
      if (m.barnn)    { adUpdates.barnnData      = m.barnn;    summary.push("BARNN ✓"); }
      if (m.yokohama) { adUpdates.yokohamaData   = m.yokohama; summary.push(`Yokohama: ${m.yokohama.length} dealers`); }
      if (Object.keys(adUpdates).length > 0) {
        setFileData(prev => ({ ...prev, ...adUpdates }));
      }

      // Record upload timestamp for all slots
      const now = new Date().toISOString();
      const newDates = { ...uploadDates, ar: now, weekComp: now, sales: now, master: now };
      setUploadDates(newDates);
      setDismissedWarning(false);
      try { localStorage.setItem("upload_dates", JSON.stringify(newDates)); } catch {}

      setNotice(`✓ Master Upload loaded — ${summary.join(" · ")}`);
      setTimeout(() => setNotice(""), 8000);
      return;
    }

    // ── LEGACY INDIVIDUAL FILE UPLOADS ────────────────────────────────────
    if (slotId === "weekComp") {
      parsed = parseWeekCompWorkbook(wb);
      setFileData(prev => {
        const existing = prev.weekComp || SEED_WEEK_COMP;

        // ── CUMULATIVE WEEK MERGE ──────────────────────────────────────────
        // Merge weeks: keep all existing weeks, add/update new ones from upload
        let mergedWeeks = [...(existing.weeks || [])];
        if (parsed.weeks && parsed.weeks.length > 0) {
          parsed.weeks.forEach(newWeek => {
            const idx = mergedWeeks.findIndex(w => w.week === newWeek.week);
            if (idx >= 0) {
              mergedWeeks[idx] = newWeek; // update existing week
            } else {
              mergedWeeks.push(newWeek);  // add new week
            }
          });
          mergedWeeks.sort((a,b) => a.week - b.week);
        }

        // ── CUMULATIVE ACTION PLAN MERGE ───────────────────────────────────
        // Update existing customers with new YTD figures, add new customers
        let mergedAP = [...(existing.actionPlan || [])];
        if (parsed.actionPlan && parsed.actionPlan.length > 0) {
          parsed.actionPlan.forEach(newAP => {
            const idx = mergedAP.findIndex(a => String(a.custNum) === String(newAP.custNum));
            if (idx >= 0) {
              mergedAP[idx] = newAP; // update with latest figures
            } else {
              mergedAP.push(newAP);  // add new customer
            }
          });
        }

        // ── DETECT NEW WEEKS ───────────────────────────────────────────────
        const existingWeekNums = new Set((existing.weeks||[]).map(w=>w.week));
        const newWeekNums = (parsed.weeks||[]).filter(w=>!existingWeekNums.has(w.week)).map(w=>w.week);
        if (newWeekNums.length > 0) {
          // Store new week numbers for AI analysis trigger
          localStorage.setItem("pulse_new_weeks", JSON.stringify(newWeekNums));
        }

        return {
          ...prev,
          weekComp: {
            weeks:      mergedWeeks,
            depts:      (parsed.depts && parsed.depts.length > 0) ? parsed.depts : existing.depts,
            actionPlan: mergedAP,
            customers:  (parsed.customers && parsed.customers.length > 0) ? parsed.customers : existing.customers,
          }
        };
      });
    } else if (slotId === "sales") {
      // Auto-detect year columns (e.g. "2025 Sales", "2026 Sales", "2027 Sales")
      const sheetName = wb.SheetNames[0];
      const rows = readSheet(wb, sheetName);
      if (rows.length === 0) { setNotice("⚠ Sales sheet appears empty"); return; }
      const cols = Object.keys(rows[0]);
      const yearCols = cols.filter(c => /^20\d{2}\s*Sales/i.test(c));
      // Build a parsed object keyed by year: { 2025: rows, 2026: rows, ... }
      const salesByYear = {};
      yearCols.forEach(col => {
        const yr = col.match(/20\d{2}/)[0];
        salesByYear[yr] = rows.map(r => ({ ...r, _salesAmt: Number(r[col] || 0) }));
      });
      // Also store as sales2025, sales2026 etc. for backward compat with week comp parser
      const updates = { sales: rows, _salesYears: yearCols.map(c => c.match(/20\d{2}/)[0]) };
      yearCols.forEach(col => {
        const yr = col.match(/20\d{2}/)[0];
        updates[`sales${yr}`] = salesByYear[yr];
      });
      const yearList = yearCols.map(c => c.match(/20\d{2}/)[0]).join(", ");
      setFileData(prev => ({ ...prev, ...updates }));
      setNotice(`✓ Sales data loaded — years detected: ${yearList || "none"}`);
    } else if (slotId === "ar") {
      // AR has a specific format — parse it like the seed data
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
      const arParsed = [];
      rows.slice(1).forEach(row => {
        if (!row[1] || typeof row[1] !== "number") return;
        let lastPaid = "";
        if (row[12]) {
          try {
            const d = row[12] instanceof Date ? row[12] : new Date(row[12]);
            if (!isNaN(d)) lastPaid = d.toISOString().slice(0,10);
          } catch {}
        }
        arParsed.push({
          salesman:  String(row[0]||"").trim(),
          custNum:   row[1],
          shortName: String(row[2]||"").trim(),
          name:      String(row[3]||"").trim(),
          phone:     String(row[4]||"").trim(),
          balance:   Number(row[5]||0),
          futDue:    Number(row[6]||0),
          curDue:    Number(row[7]||0),
          due1_30:   Number(row[8]||0),
          due31_60:  Number(row[9]||0),
          due61_90:  Number(row[10]||0),
          dueOver90: Number(row[11]||0),
          lastPaid,
        });
      });
      if (arParsed.length > 0) {
        setFileData(prev => ({ ...prev, ar: arParsed }));
        setNotice(`✓ AR loaded — ${arParsed.length} accounts, $${arParsed.reduce((s,a)=>s+a.balance,0).toLocaleString("en-US",{maximumFractionDigits:0})} total`);
      } else {
        setNotice("⚠ AR file could not be parsed — check format");
      }
    } else {
      const sheetName = wb.SheetNames[0];
      parsed = readSheet(wb, sheetName);
      setFileData(prev => ({ ...prev, [slotId]: parsed }));
    }
    if (slotId === "weekComp") {
      // Check if new weeks were detected
      const newWeeks = JSON.parse(localStorage.getItem("pulse_new_weeks") || "[]");
      if (newWeeks.length > 0) {
        setNotice(`✓ Week Comp loaded — ${newWeeks.length} new week${newWeeks.length>1?"s":""} added (W${newWeeks.join(", W")})`);
        localStorage.removeItem("pulse_new_weeks");
        setNewWeeksDetected(newWeeks);
      } else {
        setNotice("✓ Week Comp loaded — existing data updated");
      }
    } else if (slotId !== "sales") {
      setNotice(`✓ ${FILE_SLOTS.find(s=>s.id===slotId)?.label} loaded`);
    }
    setTimeout(() => setNotice(""), 5000);
    // Record upload timestamp
    const newDates = { ...uploadDates, [slotId]: new Date().toISOString() };
    setUploadDates(newDates);
    setDismissedWarning(false);
    try { localStorage.setItem("upload_dates", JSON.stringify(newDates)); } catch {}
  }

  function clearAll() {
    setFileData({ weekComp: SEED_WEEK_COMP, customers: SEED_CUSTOMERS, ar: SEED_AR, sales: null, sales2025: null, sales2026: null });
    setCustTabs([]);
    setNotice("Cleared");
    setTimeout(() => setNotice(""), 2000);
  }

  function goAI(prompt) { setAiPrompt(prompt); setTab("ai"); }

  const weekComp = fileData.weekComp;
  const fileDots = FILE_SLOTS.map(s => {
    if (s.id === "sales") return !!(fileData.sales || fileData.sales2025 || fileData.sales2026);
    return !!fileData[s.id];
  });

  if (!currentUser) {
    return <LoginScreen onLogin={user => { setCurrentUser(user); window.__pulseUser = user; logActivity("login","Logged in", user); }} />;
  }
  // Keep window ref in sync
  window.__pulseUser = currentUser;

  return (
    <div style={S.app}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div style={S.header}>
        <div>
          <div style={S.logo}>Pulse Dashboard</div>
          <div style={S.sub}>Tires · Ag · Industrial · Multi-line Distributor</div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {notice && <span style={{ fontSize: "0.67rem", color: notice.startsWith("✓") ? GREEN : RED }}>{notice}</span>}
          <div style={{ display: "flex", gap: 5, alignItems: "center" }}>
            {fileDots.map((loaded, i) => (
              <span key={i} title={FILE_SLOTS[i].label} style={{ width: 7, height: 7, borderRadius: "50%", background: loaded ? GREEN : BORDER, display: "inline-block" }} />
            ))}
            <span style={{ fontSize: "0.68rem", color: MUTED, marginLeft: 3 }}>{fileDots.filter(Boolean).length}/{FILE_SLOTS.length}</span>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginLeft:8, paddingLeft:8, borderLeft:`1px solid rgba(255,255,255,0.2)` }}>
            <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#FFFFFF" }}>
              <span style={{ opacity:0.7, fontWeight:400 }}>Logged in: </span>{currentUser.name}
            </span>
            <button
              onClick={async () => {
                const type = window.confirm("Is this a bug? Click OK for bug, Cancel for suggestion.") ? "bug" : "suggestion";
                const text = window.prompt(`${type === "bug" ? "🐛 Describe the bug:" : "💡 Enter your suggestion:"}`);
                if (text && text.trim()) {
                  try {
                    const entry = { ts: new Date().toISOString(), user: currentUser?.name || "Unknown", userId: currentUser?.id || "", type, text: text.trim() };
                    // Save locally
                    const existing = JSON.parse(localStorage.getItem("pulse_suggestions") || "[]");
                    localStorage.setItem("pulse_suggestions", JSON.stringify([entry, ...existing].slice(0, 100)));
                    // Write directly to activity log using logActivity
                    logActivity(type === "bug" ? "bug_report" : "suggestion", text.trim());
                    // Also sync to Supabase feedback table
                    await sbFetch("rep_feedback", "POST", {
      ts: entry.ts, user: entry.user, userId: entry.userId,
      type: entry.type, text: entry.text, updated_at: new Date().toISOString()
    }).catch(()=>{});
                    alert("✓ Submitted! Your feedback has been logged for Admin review.");
                  } catch {}
                }
              }}
              style={{ fontSize:"0.65rem", color:"rgba(255,255,255,0.8)", background:"rgba(255,255,255,0.1)", border:"1px solid rgba(255,255,255,0.25)", borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>
              💡 Feedback
            </button>
            <button
              onClick={() => { logActivity("logout","Logged out"); localStorage.removeItem("pulse_remembered_user"); setCurrentUser(null); }}
              style={{ fontSize:"0.65rem", color:"rgba(255,255,255,0.8)", background:"rgba(255,255,255,0.15)", border:"1px solid rgba(255,255,255,0.3)", borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>
              Sign Out
            </button>
          </div>
        </div>
      </div>

      {/* ── Weekly upload reminder banner ── */}
      {!dismissedWarning && (warningInfo.isOverdue || warningInfo.isPending) && (
        <div style={{
          padding: "0.6rem 1.5rem",
          background: warningInfo.isOverdue ? "#FEF2F2" : "#FFFBEB",
          borderBottom: `2px solid ${warningInfo.isOverdue ? "#FECACA" : "#FDE68A"}`,
          display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap",
        }}>
          <span style={{ fontSize: "1rem" }}>{warningInfo.isOverdue ? "🚨" : "📅"}</span>
          <div style={{ flex: 1 }}>
            <span style={{ fontSize: "0.75rem", fontWeight: 700, color: warningInfo.isOverdue ? "#DC2626" : "#D97706" }}>
              {warningInfo.isOverdue ? "OVERDUE — " : "REMINDER — "}
              Weekly uploads needed
            </span>
            <span style={{ fontSize: "0.72rem", color: warningInfo.isOverdue ? "#991B1B" : "#92400E", marginLeft: 8 }}>
              {warningInfo.staleSlots.map(s => {
                const slot = FILE_SLOTS.find(f => f.id === s);
                return slot?.label;
              }).filter(Boolean).join(", ")} {warningInfo.staleSlots.length === 1 ? "has" : "have"} not been uploaded this week
            </span>
          </div>
          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
            <button
              onClick={() => { setTab("setup"); setDismissedWarning(false); }}
              style={{ fontSize: "0.7rem", fontWeight: 700, color: "#FFFFFF",
                background: warningInfo.isOverdue ? "#DC2626" : "#D97706",
                border: "none", borderRadius: 6, padding: "0.3rem 0.85rem", cursor: "pointer" }}>
              Upload Now
            </button>
            <button
              onClick={() => setDismissedWarning(true)}
              style={{ fontSize: "0.68rem", color: warningInfo.isOverdue ? "#DC2626" : "#D97706",
                background: "none", border: `1px solid ${warningInfo.isOverdue ? "#FECACA" : "#FDE68A"}`,
                borderRadius: 6, padding: "0.3rem 0.65rem", cursor: "pointer" }}>
              Dismiss
            </button>
          </div>
        </div>
      )}
      {/* ── All clear ── */}
      {!dismissedWarning && warningInfo.staleSlots.length === 0 && (
        <div style={{ padding: "0.35rem 1.5rem", background: "#F0FDF4", borderBottom: `1px solid #BBF7D0`, display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: "0.68rem", color: "#059669", fontWeight: 600 }}>✓ All weekly uploads are current</span>
          <span style={{ fontSize: "0.65rem", color: MUTED }}>Week Comp · AR / Aging · Sales Data</span>
        </div>
      )}
      <div style={S.nav}>
        {TABS.map(t => {
          const color = t.rep ? REP_COLORS[t.label] : AMBER;
          return <button key={t.id} style={S.navBtn(tab === t.id, color)} onClick={() => { setTab(t.id); logActivity("view_tab", t.label); }}>{t.label}</button>;
        })}
        {currentUser?.id === "admin" && (
          <button style={S.navBtn(tab==="adminlog", "#DC2626")} onClick={() => setTab("adminlog")}>🔐 Admin</button>
        )}
        {custTabs.length > 0 && <div style={{ width: 1, background: BORDER, margin: "8px 4px" }} />}
        {custTabs.map(ct => (
          <div key={ct.custNum} style={{ display: "flex", alignItems: "center", borderBottom: `3px solid ${tab === `cust_${ct.custNum}` ? "#1E5FCC" : "transparent"}` }}>
            <button
              style={{ ...S.navBtn(tab === `cust_${ct.custNum}`, "#1E5FCC"), paddingRight: "0.4rem", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: "0.68rem" }}
              onClick={() => setTab(`cust_${ct.custNum}`)}
              title={ct.customer}
            >
              👤 {ct.customer.length > 14 ? ct.customer.slice(0,14)+"…" : ct.customer}
            </button>
            <button onClick={() => closeCustomer(ct.custNum)} style={{ background: "none", border: "none", color: MUTED, cursor: "pointer", fontSize: "0.75rem", padding: "0 6px 0 0", lineHeight: 1 }} title="Close">×</button>
          </div>
        ))}
      </div>

      <div style={S.main}>
        {tab === "setup"    && <FileSetup fileData={fileData} onUpload={handleUpload} onClear={clearAll} />}
        {tab === "overview" && newWeeksDetected.length > 0 && (
          <NewWeekBanner
            newWeeks={newWeeksDetected}
            weekComp={fileData.weekComp || SEED_WEEK_COMP}
            onDismiss={() => setNewWeeksDetected([])}
          />
        )}
        {tab === "overview" && <OverviewTab weekComp={weekComp} onAskAI={goAI} onCustomerClick={openCustomer} customers={fileData.customers || SEED_CUSTOMERS} />}
        {["tiffany","larry","austin","house"].includes(tab) && (
          <RepTab repName={tab.charAt(0).toUpperCase()+tab.slice(1)} weekComp={weekComp} onAskAI={goAI} onCustomerClick={openCustomer} customers={fileData.customers || SEED_CUSTOMERS} inactiveCustomers={inactiveCustomers} leads={leads} onAddLead={addLead} onDeleteLead={deleteLead} currentUser={currentUser} onLogActivity={logActivity} onRefreshLeads={refreshLeads} />
        )}
        {tab === "ai" && <AITab weekComp={weekComp} initialPrompt={aiPrompt} onClearPrompt={() => setAiPrompt("")} />}
        {tab === "map" && <MapTab customers={fileData.customers || SEED_CUSTOMERS} weekComp={weekComp} />}
        {tab === "ar" && <ARTab ar={fileData.ar || SEED_AR} />}
        {tab === "cardealer" && <CarDealerTab weekComp={weekComp} customers={fileData.customers || SEED_CUSTOMERS} onCustomerClick={openCustomer} />}
        {tab === "adminlog" && currentUser?.id === "admin" && <AdminTab currentUser={currentUser} leads={leads} onAddLead={addLead} onDeleteLead={deleteLead} convertLead={convertLead} />}
        {custTabs.map(ct => tab === `cust_${ct.custNum}` && (
          <CustomerDetailTab
            key={ct.custNum}
            ap={ct}
            customers={fileData.customers || SEED_CUSTOMERS}
            ar={fileData.ar || SEED_AR}
            weekComp={fileData.weekComp || SEED_WEEK_COMP}
            onClose={() => closeCustomer(ct.custNum)}
            inactiveRecord={inactiveCustomers[String(ct.custNum)] || null}
            onMarkInactive={(reason) => markInactive(ct.custNum, ct.customer, reason, currentUser?.name || "Admin")}
            onMarkActive={() => markActive(ct.custNum, ct.customer)}
            currentUser={currentUser}
            onLogActivity={logActivity}
          />
        ))}
      </div>
    </div>
  );
}

// ── FileSetup ─────────────────────────────────────────────────────────────────
function FileSetup({ fileData, onUpload, onClear }) {
  return (
    <div>
      <div style={{ ...S.card, marginBottom: "1.25rem" }}>
        <div style={{ fontSize: "0.7rem", color: MUTED, marginBottom: "0.5rem", letterSpacing: "0.1em" }}>FILE SLOTS</div>
        <div style={{ fontSize: "0.72rem", color: TEXT, lineHeight: 1.6 }}>
          Upload Excel files to each slot. <span style={{ color: AMBER }}>Sales Data</span> and <span style={{ color: AMBER }}>Week Comp</span> are pre-loaded — upload a new file any week to refresh. When 2027 data is ready, just upload the new sheet and it's detected automatically.
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: "0.75rem" }}>
        {FILE_SLOTS.map(slot => {
          const loaded = !!fileData[slot.id];
          return (
            <div key={slot.id} style={{ ...S.card, borderLeft: `3px solid ${loaded ? GREEN : BORDER}`, marginBottom: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <div style={{ fontSize: "0.78rem", fontWeight: 600, color: loaded ? GREEN : TEXT, marginBottom: 4 }}>{slot.label}</div>
                  <div style={{ fontSize: "0.68rem", color: MUTED }}>{slot.desc}</div>
                </div>
                <span style={{ fontSize: "0.65rem", color: loaded ? GREEN : MUTED }}>{loaded ? "✓ Loaded" : "Empty"}</span>
              </div>
              <label style={{ display: "block", marginTop: "0.75rem" }}>
                <span style={{ ...S.btn(AMBER), display: "inline-block", cursor: "pointer", fontSize: "0.65rem" }}>Upload .xlsx</span>
                <input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => {
                  const file = e.target.files[0];
                  if (!file) return;
                  const reader = new FileReader();
                  reader.onload = ev => {
                    const wb = XLSX.read(ev.target.result, { type: "array" });
                    onUpload(slot.id, wb);
                  };
                  reader.readAsArrayBuffer(file);
                  e.target.value = "";
                }} />
              </label>
            </div>
          );
        })}
      </div>

    </div>
  );
}

// ── OverviewTab ───────────────────────────────────────────────────────────────

// ── New Week Banner + AI Analysis ────────────────────────────────────────────
function NewWeekBanner({ newWeeks, weekComp, onDismiss }) {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading]   = useState(false);

  async function generateAnalysis() {
    setLoading(true);
    const weeks = weekComp.weeks || [];
    const newWkData = weeks.filter(w => newWeeks.includes(w.week));
    const prevWks   = weeks.filter(w => !newWeeks.includes(w.week)).slice(-4);

    // Build context
    const newWkLines = newWkData.map(w =>
      `W${w.week}: $${w.sales2026.toLocaleString()} 2026 vs $${w.sales2025.toLocaleString()} 2025 (${w.changePct>=0?"+":""}${(w.changePct*100).toFixed(1)}%)`
    ).join("\n");

    const prevWkLines = prevWks.map(w =>
      `W${w.week}: $${w.sales2026.toLocaleString()} (${w.changePct>=0?"+":""}${(w.changePct*100).toFixed(1)}%)`
    ).join("\n");

    const ytd26 = weeks.reduce((s,w)=>s+w.sales2026,0);
    const ytd25 = weeks.reduce((s,w)=>s+w.sales2025,0);

    // Top movers this week
    const ap = weekComp.actionPlan || [];
    const topGain = [...ap].sort((a,b)=>b.change-a.change).slice(0,5)
      .map(a=>`${a.customer} (${a.salesman}): +$${a.change.toLocaleString()}`).join("\n");
    const topDecline = [...ap].filter(a=>a.change<0).sort((a,b)=>a.change-b.change).slice(0,5)
      .map(a=>`${a.customer} (${a.salesman}): $${a.change.toLocaleString()}`).join("\n");

    // Dept analysis
    const depts = (weekComp.depts||[]).sort((a,b)=>b.sales-a.sales).slice(0,5)
      .map(d=>`${d.dept}: $${d.sales.toLocaleString()} (${d.assessment})`).join("\n");

    const prompt = `You are a sales manager at a tire and ag distributor analyzing a new week's performance.

NEW WEEK${newWeeks.length>1?"S":""}: W${newWeeks.join(", W")}
${newWkLines}

PRIOR 4 WEEKS TREND:
${prevWkLines}

YTD 2026: $${ytd26.toLocaleString()} vs 2025: $${ytd25.toLocaleString()} (${ytd25>0?(((ytd26-ytd25)/ytd25)*100).toFixed(1):"0"}% YTD)

TOP GROWING ACCOUNTS:
${topGain}

DECLINING ACCOUNTS:
${topDecline}

TOP DEPARTMENTS (YTD):
${depts}

Provide a concise weekly performance summary with:
1. Quick headline on this week's performance vs prior trend
2. Top 2-3 accounts to prioritize follow-up based on their trajectory
3. Any departments or patterns worth noting
4. One specific action recommendation for the team this week

Keep it punchy — 4 short paragraphs max. Write for a Monday morning sales meeting.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{
          "Content-Type":"application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version":"2023-06-01",
          "anthropic-dangerous-direct-browser-access":"true",
        },
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:800,
          messages:[{role:"user",content:prompt}] })
      });
      if (!res.ok) throw new Error(`API ${res.status}`);
      const data = await res.json();
      setAnalysis(data.content?.[0]?.text || "No response.");
    } catch(e) {
      setAnalysis(`Error: ${e.message}`);
    }
    setLoading(false);
  }

  return (
    <div style={{ margin:"0.75rem 1.5rem 0", padding:"0.85rem 1rem",
      background:"linear-gradient(135deg,#F0FDF4,#EFF6FF)",
      border:"2px solid #86EFAC", borderRadius:10 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
        <div>
          <span style={{ fontSize:"0.85rem", fontWeight:800, color:"#059669" }}>
            🆕 New week{newWeeks.length>1?"s":""} uploaded — W{newWeeks.join(", W")}
          </span>
          <span style={{ fontSize:"0.72rem", color:MUTED, marginLeft:10 }}>
            Data added to existing history — nothing replaced
          </span>
        </div>
        <div style={{ display:"flex", gap:6 }}>
          {!analysis && !loading && (
            <button onClick={generateAnalysis}
              style={{ fontSize:"0.72rem", fontWeight:700, color:"#fff", background:"#059669",
                border:"none", borderRadius:6, padding:"0.35rem 0.85rem", cursor:"pointer" }}>
              ◈ Analyze New Week
            </button>
          )}
          <button onClick={onDismiss}
            style={{ fontSize:"0.68rem", color:MUTED, background:"none",
              border:`1px solid ${BORDER}`, borderRadius:6, padding:"0.3rem 0.6rem", cursor:"pointer" }}>
            × Dismiss
          </button>
        </div>
      </div>
      {loading && (
        <div style={{ marginTop:"0.75rem", fontSize:"0.75rem", color:MUTED, textAlign:"center" }}>
          ◈ Analyzing W{newWeeks.join(", W")} performance…
        </div>
      )}
      {analysis && !loading && (
        <div style={{ marginTop:"0.85rem", paddingTop:"0.85rem", borderTop:`1px solid #BBF7D0` }}>
          <div style={{ fontSize:"0.7rem", fontWeight:700, color:"#059669", textTransform:"uppercase",
            letterSpacing:"0.1em", marginBottom:"0.5rem" }}>
            📊 Week Analysis — W{newWeeks.join(", W")}
          </div>
          <div style={{ fontSize:"0.78rem", color:TEXT, lineHeight:1.85, whiteSpace:"pre-wrap" }}>
            {analysis}
          </div>
          <button onClick={generateAnalysis}
            style={{ marginTop:"0.6rem", fontSize:"0.65rem", color:MUTED, background:"none",
              border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>
            ↺ Regenerate
          </button>
        </div>
      )}
    </div>
  );
}


// ── MSR — Monday Sales Review Tab ─────────────────────────────────────────────
const SEED_MSR = [
  {
    id: "msr_20260601",
    date: "2026-06-02",
    weekNum: 22,
    title: "W22 Review — June 2, 2026",
    headline: "W22 down 8.3% on units but GP% trending in the right direction after price increase. YTD still strong at +19.1% units, +11% sales.",
    sections: [
      {
        label: "📊 W22 Performance",
        color: "#DC2626",
        bullets: [
          "Units: 11,079 vs 12,076 PY — down 8.3% (997 units short)",
          "Revenue: $1.487M, down 7.2% ($114K)",
          "GP: $207K, down 21.8% ($58K) — BUT GP% at 13.9%, trending UP",
          "GP% progression post-price increase: 13.27 → 13.3 → 13.53 → 13.9% — target is 15-16%"
        ]
      },
      {
        label: "🏆 Departments",
        color: "#059669",
        bullets: [
          "Industrial tires: TOP performer for the week — first time ever",
          "Radio/light truck: BOTTOM — down 1,259 units (Venom, Predator, Nexon, Transamerica brands phased out)",
          "Truck tires (TBR): down 447 units — Byron down 254, Tifton and Statesboro also off",
          "Passenger: high volume but NOT profitable growth — focus elsewhere"
        ]
      },
      {
        label: "📅 YTD & Month-to-Month",
        color: "#0891B2",
        bullets: [
          "YTD through May: units +19.1%, sales +11%, +$3.9M, +50K units vs 2025",
          "May month-over-month: units -998 (-2%), sales -$400K (-5%)",
          "Still tracking well above 8% growth budget for 2026",
          "June focus: don't have a down month — containers in the pipe should help"
        ]
      },
      {
        label: "📦 Inventory",
        color: "#7C3AED",
        bullets: [
          "Current inventory: $21.1M (flat vs last week)",
          "3-4 direct containers expected to land this week — approx 1,000 TBR units",
          "TBR Town warehouse (Store 4): landlord wants to lease it — exit ASAP",
          "ACTION: Move TBR Town inventory to Statesboro and other branches immediately — creating out-of-stock illusion",
          "Toyo backorder issues — Scott working on ETAs, limited info from Toyo",
          "Blackhawk BFR-57: potential fill-in option being sourced, answer by tomorrow"
        ]
      },
      {
        label: "💰 AR Collections",
        color: "#D97706",
        bullets: [
          "New month — hammer AR hard, prime collection time",
          "Several accounts rolled past due from May",
          "3-4 accounts turned over to attorneys — complaints filed in court today",
          "Chris posting weekend payments this morning before running statements — updated AR coming"
        ]
      },
      {
        label: "🎯 Promotions",
        color: "#0891B2",
        bullets: [
          "Dunlop/Owala promo: 1 black laser-engraved Owala thermos per 4 Dunlops — first 2 weeks of June. Any dealer eligible. Flyer going on Tirelink today.",
          "Tag Team program: coming out shortly",
          "Americus May recon ended — Economy Tire big winner, Screens Tire & Muffler also in the money",
          "Falken Q2: only 1 DNR in the money — push hard through June",
          "Toyo Q2: 9 dealers in the money — one more month to close out"
        ]
      },
      {
        label: "🔑 Key Directives",
        color: "#1E5FCC",
        bullets: [
          "STOP THE MARGIN BLEED — volume does not fix margin. Chase GP, not units.",
          "TBR and light truck are the two big segment focuses",
          "Passenger segment growing but not profitable — don't chase that volume",
          "If cutting price aggressively, it must be for a new market or a down market — not existing well-serviced routes",
          "Our service is worth more than cutting 2 points off a tire to close a transaction",
          "Price levels review this month — adjusting across 3 accounts (Tifton)"
        ]
      }
    ],
    tiftonNotes: "JMZ and Del Toro called out as radio/LT accounts to watch. Eric's Tifton and Eric's Sylvester discussed for retention. Price level adjustments planned for 3 accounts — moving away from aggressive pricing on well-covered routes. Still no access to Dynamics CRM.",
    createdBy: "AI Summary"
  }
];

function MSRTab() {
  const MSR_KEY = "pulse_msr_meetings";
  const [meetings, setMeetings] = useState(() => {
    try {
      const saved = localStorage.getItem(MSR_KEY);
      return saved ? JSON.parse(saved) : SEED_MSR;
    } catch { return SEED_MSR; }
  });
  const [selected, setSelected] = useState(meetings[0]?.id || null);
  const [showUpload, setShowUpload] = useState(false);
  const [transcript, setTranscript] = useState("");
  const [meetingDate, setMeetingDate] = useState(new Date().toISOString().slice(0,10));
  const [generating, setGenerating] = useState(false);

  function saveMeetings(updated) {
    setMeetings(updated);
    try { localStorage.setItem(MSR_KEY, JSON.stringify(updated)); } catch {}
  }

  const current = meetings.find(m => m.id === selected);

  async function generateSummary() {
    if (!transcript.trim()) return;
    setGenerating(true);
    const prompt = `You are analyzing a tire distribution company's Monday Sales Review meeting transcript.

TRANSCRIPT:
${transcript.slice(0, 12000)}

Generate a structured meeting summary as JSON only (no markdown):
{
  "headline": "One sentence capturing the most important takeaway",
  "weekNum": <week number if mentioned, or null>,
  "sections": [
    {
      "label": "emoji + Section Title",
      "color": "#hexcolor",
      "bullets": ["specific bullet point", "another point"]
    }
  ],
  "tiftonNotes": "Any Tifton-specific items, action items, or callouts from the transcript",
  "createdBy": "AI Summary"
}

Use these section labels as relevant: 📊 Performance, 🏆 Departments, 📅 YTD, 📦 Inventory, 💰 AR Collections, 🎯 Promotions, 🔑 Key Directives, 👥 Rep Updates.
Colors: use #DC2626 red for down/bad, #059669 green for up/good, #0891B2 teal for info, #D97706 amber for caution, #7C3AED purple for inventory, #1E5FCC blue for strategy.
Be specific — use actual numbers, names, and account names from the transcript.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:2000,
          messages:[{role:"user",content:prompt}] })
      });
      const data = await res.json();
      const raw = data.content?.[0]?.text || "{}";
      const parsed = JSON.parse(raw.replace(/```json|```/g,"").trim());
      const newMeeting = {
        id:       `msr_${Date.now()}`,
        date:     meetingDate,
        weekNum:  parsed.weekNum || null,
        title:    `W${parsed.weekNum || "?"} Review — ${new Date(meetingDate).toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}`,
        headline: parsed.headline || "",
        sections: parsed.sections || [],
        tiftonNotes: parsed.tiftonNotes || "",
        createdBy: "AI Summary",
      };
      const updated = [newMeeting, ...meetings];
      saveMeetings(updated);
      setSelected(newMeeting.id);
      setShowUpload(false);
      setTranscript("");
    } catch(e) {
      alert(`Summary failed: ${e.message}`);
    }
    setGenerating(false);
  }

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"0.85rem", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontSize:"0.82rem", fontWeight:700, color:"#1E5FCC" }}>📋 Monday Sales Review</div>
          <div style={{ fontSize:"0.68rem", color:MUTED }}>{meetings.length} meeting{meetings.length!==1?"s":""} on record</div>
        </div>
        <button onClick={()=>setShowUpload(!showUpload)}
          style={{ ...S.btn("#1E5FCC"), fontSize:"0.7rem", fontWeight:700 }}>
          {showUpload ? "× Cancel" : "+ Add Meeting"}
        </button>
      </div>

      {/* Upload form */}
      {showUpload && (
        <div style={{ ...S.card, border:"2px solid #1E5FCC", marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:"#1E5FCC", marginBottom:"0.75rem" }}>◈ Add Meeting Summary</div>
          <div style={{ marginBottom:"0.6rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Meeting Date</div>
            <input type="date" value={meetingDate} onChange={e=>setMeetingDate(e.target.value)}
              style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem" }} />
          </div>
          <div style={{ marginBottom:"0.75rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Paste meeting transcript or notes</div>
            <textarea value={transcript} onChange={e=>setTranscript(e.target.value)}
              placeholder="Paste the full meeting transcript here. AI will extract key points, numbers, action items, and Tifton-specific notes..."
              rows={8}
              style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.75rem", resize:"vertical",
                outline:"none", boxSizing:"border-box", lineHeight:1.7 }}
              onFocus={e=>e.target.style.borderColor="#1E5FCC"}
              onBlur={e=>e.target.style.borderColor=BORDER}
            />
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:3 }}>{transcript.length.toLocaleString()} characters</div>
          </div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <button onClick={()=>{setShowUpload(false);setTranscript("");}} style={{ ...S.btn(MUTED) }}>Cancel</button>
            <button onClick={generateSummary} disabled={!transcript.trim()||generating}
              style={{ ...S.btn("#1E5FCC"), background:"#1E5FCC", color:"#fff",
                opacity:!transcript.trim()||generating?0.5:1 }}>
              {generating ? "◈ Analyzing…" : "◈ Generate Summary"}
            </button>
          </div>
        </div>
      )}

      {/* Meeting selector */}
      {meetings.length > 1 && (
        <div style={{ display:"flex", gap:5, marginBottom:"0.85rem", flexWrap:"wrap" }}>
          {meetings.map(m => (
            <button key={m.id} onClick={()=>setSelected(m.id)}
              style={{ fontSize:"0.68rem", fontWeight:selected===m.id?700:400,
                color:selected===m.id?"#fff":"#1E5FCC",
                background:selected===m.id?"#1E5FCC":"#EFF6FF",
                border:`1px solid ${selected===m.id?"#1E5FCC":"#BFDBFE"}`,
                borderRadius:6, padding:"0.25rem 0.7rem", cursor:"pointer", whiteSpace:"nowrap" }}>
              {m.title}
            </button>
          ))}
        </div>
      )}

      {/* Meeting content */}
      {current && (
        <div>
          {/* Headline */}
          <div style={{ padding:"0.75rem 1rem", background:"linear-gradient(135deg,#EFF6FF,#F0FDF4)",
            border:"1px solid #BFDBFE", borderRadius:8, marginBottom:"1rem" }}>
            <div style={{ fontSize:"0.65rem", color:"#1E5FCC", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>
              {current.title} · {new Date(current.date).toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"})}
            </div>
            <div style={{ fontSize:"0.82rem", fontWeight:700, color:TEXT, lineHeight:1.6 }}>{current.headline}</div>
          </div>

          {/* Sections */}
          <div style={{ display:"flex", flexDirection:"column", gap:"0.75rem", marginBottom:"1rem" }}>
            {(current.sections||[]).map((sec, i) => (
              <div key={i} style={{ ...S.card, borderLeft:`4px solid ${sec.color}`, padding:"0.75rem 1rem" }}>
                <div style={{ fontSize:"0.72rem", fontWeight:700, color:sec.color,
                  textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:"0.5rem" }}>
                  {sec.label}
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
                  {(sec.bullets||[]).map((b, j) => (
                    <div key={j} style={{ display:"flex", gap:8, fontSize:"0.76rem", color:TEXT, lineHeight:1.65 }}>
                      <span style={{ color:sec.color, flexShrink:0, marginTop:2 }}>▸</span>
                      <span>{b}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          {/* Tifton-specific notes */}
          {current.tiftonNotes && (
            <div style={{ ...S.card, background:"#FFFBEB", borderLeft:`4px solid ${AMBER}`, padding:"0.75rem 1rem" }}>
              <div style={{ fontSize:"0.7rem", fontWeight:700, color:AMBER,
                textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:"0.5rem" }}>
                📍 Tifton Notes
              </div>
              <div style={{ fontSize:"0.76rem", color:TEXT, lineHeight:1.7 }}>{current.tiftonNotes}</div>
            </div>
          )}

          {/* Footer */}
          <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:"0.75rem", textAlign:"right" }}>
            {current.createdBy} · {new Date(current.date).toLocaleDateString()}
          </div>
        </div>
      )}
    </div>
  );
}

function OverviewTab({ weekComp, onAskAI, onCustomerClick, customers }) {
  const [subTab, setSubTab] = useState("weekcomp");
  const branchData = SEED_BRANCH_DATA;

  const weeks = weekComp?.weeks || [];
  const totalYTD25 = weeks.reduce((s,w) => s+w.sales2025, 0);
  const totalYTD26 = weeks.reduce((s,w) => s+w.sales2026, 0);
  const totalGP26  = weeks.reduce((s,w) => s+w.gp2026, 0);
  const latest     = weeks[weeks.length-1];
  // Branch-level YTD (Q1 + Q2)
  const bSumQ1_26 = Object.values(branchData.branches).reduce((s,b)=>s+b.q1_2026,0);
  const bSumQ1_25 = Object.values(branchData.branches).reduce((s,b)=>s+b.q1_2025,0);
  const bSumQ2_26 = Object.values(branchData.branches).reduce((s,b)=>s+b.q2_2026,0);
  const bSumQ2_25 = Object.values(branchData.branches).reduce((s,b)=>s+b.q2_2025,0);
  const bYTD26 = bSumQ1_26 + bSumQ2_26;
  const bYTD25 = bSumQ1_25 + bSumQ2_25;
  const bGP26  = Object.values(branchData.branches).reduce((s,b)=>s+b.q1_gp26+b.q2_gp26,0);

  return (
    <div>
      <div style={S.kpiRow}>
        <div style={S.kpi(AMBER)}>
          <div style={S.kpiVal}>{fmt(totalYTD26)}</div>
          <div style={S.kpiLbl}>YTD 2026 Sales</div>
        </div>
        <div style={S.kpi(totalYTD26 >= totalYTD25 ? GREEN : RED)}>
          <div style={{ ...S.kpiVal, color: clr(totalYTD26-totalYTD25) }}>{fmt(totalYTD26-totalYTD25)}</div>
          <div style={S.kpiLbl}>vs 2025 ({pct(totalYTD25 ? (totalYTD26-totalYTD25)/totalYTD25 : 0)})</div>
        </div>
        <div style={S.kpi(TEAL)}>
          <div style={S.kpiVal}>{fmt(totalGP26)}</div>
          <div style={S.kpiLbl}>YTD GP ({pct(totalYTD26 ? totalGP26/totalYTD26 : 0)})</div>
        </div>
        <div style={S.kpi(MUTED)}>
          <div style={S.kpiVal}>{latest ? fmt(latest.sales2026) : "—"}</div>
          <div style={S.kpiLbl}>Latest {latest?.week || "—"} Sales</div>
        </div>
      </div>

      <div style={S.subNav}>
        <button style={S.subBtn(subTab==="weekcomp")} onClick={() => setSubTab("weekcomp")}>📅 Week by Week</button>
        <button style={S.subBtn(subTab==="depts")} onClick={() => setSubTab("depts")}>📦 Departments</button>
        <button style={S.subBtn(subTab==="action",   RED)}    onClick={() => setSubTab("action")}>🎯 Action Plan</button>
        <button style={S.subBtn(subTab==="msr",    "#1E5FCC")} onClick={() => setSubTab("msr")}>📋 MSR</button>
        <button style={S.subBtn(subTab==="trend",    TEAL)}   onClick={() => setSubTab("trend")}>📈 Trend Chart</button>
        <button style={S.subBtn(subTab==="branches", AMBER)}  onClick={() => setSubTab("branches")}>🏢 Branches</button>
        <button style={S.subBtn(subTab==="qtd",      "#0891B2")} onClick={() => setSubTab("qtd")}>📅 QTD</button>
      </div>

      {subTab === "weekcomp" && <WeekByWeekView weeks={weeks} />}
      {subTab === "depts"    && <DeptView depts={weekComp?.depts || []} />}
      {subTab === "action"   && <ActionPlanView actionPlan={weekComp?.actionPlan || []} onCustomerClick={onCustomerClick} customers={customers} />}
      {subTab === "branches"  && <BranchesTab branchData={branchData} />}
      {subTab === "qtd"       && <QTDTab branchData={branchData} />}
      {subTab === "msr" && <MSRTab />}
    {subTab === "trend"    && <TrendView weeks={weeks} />}
    </div>
  );
}

// ── Week by Week View ─────────────────────────────────────────────────────────
function WeekByWeekView({ weeks }) {
  if (!weeks.length) return <div style={{ color: MUTED, fontSize: "0.75rem" }}>No weekly data loaded.</div>;
  const latest = weeks[weeks.length - 1]?.week;
  const tot25 = weeks.reduce((s,w)=>s+w.sales2025,0);
  const tot26 = weeks.reduce((s,w)=>s+w.sales2026,0);
  const totGP25 = weeks.reduce((s,w)=>s+w.gp2025,0);
  const totGP26 = weeks.reduce((s,w)=>s+w.gp2026,0);

  return (
    <div style={S.card}>
      <div style={{ overflowX: "auto" }}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>Week</th>
              <th style={{ ...S.th, textAlign: "right" }}>2025 Sales</th>
              <th style={{ ...S.th, textAlign: "right" }}>2026 Sales</th>
              <th style={{ ...S.th, textAlign: "right" }}>$ Change</th>
              <th style={{ ...S.th, textAlign: "right" }}>% Change</th>
              <th style={{ ...S.th, textAlign: "right" }}>2025 GP</th>
              <th style={{ ...S.th, textAlign: "right" }}>2026 GP</th>
            </tr>
          </thead>
          <tbody>
            {weeks.map(w => (
              <tr key={w.week} style={{ background: w.week===latest ? "rgba(232,168,56,0.05)" : "transparent" }}>
                <td style={S.td}>
                  <span style={{ color: AMBER }}>{w.week}</span>
                  {w.week===latest && <span style={{ fontSize: "0.68rem", background: AMBER, color: BG2, padding: "1px 5px", borderRadius: 3, marginLeft: 6, fontWeight: 700 }}>LATEST</span>}
                </td>
                <td style={{ ...S.td, textAlign: "right", color: MUTED }}>{fmt(w.sales2025)}</td>
                <td style={{ ...S.td, textAlign: "right", fontWeight: 600 }}>{fmt(w.sales2026)}</td>
                <td style={{ ...S.td, textAlign: "right", color: clr(w.change) }}>{fmt(w.change)}</td>
                <td style={{ ...S.td, textAlign: "right", color: clr(w.changePct) }}>{pct(w.changePct)}</td>
                <td style={{ ...S.td, textAlign: "right", color: MUTED }}>{fmt(w.gp2025)}</td>
                <td style={{ ...S.td, textAlign: "right" }}>{fmt(w.gp2026)}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr style={{ borderTop: `1px solid ${AMBER}` }}>
              <td style={{ ...S.td, color: AMBER, fontWeight: 700 }}>YTD TOTAL</td>
              <td style={{ ...S.td, textAlign: "right", color: MUTED, fontWeight: 600 }}>{fmt(tot25)}</td>
              <td style={{ ...S.td, textAlign: "right", fontWeight: 700, color: AMBER }}>{fmt(tot26)}</td>
              <td style={{ ...S.td, textAlign: "right", fontWeight: 700, color: clr(tot26-tot25) }}>{fmt(tot26-tot25)}</td>
              <td style={{ ...S.td, textAlign: "right", fontWeight: 700, color: clr(tot26-tot25) }}>{pct(tot25 ? (tot26-tot25)/tot25 : 0)}</td>
              <td style={{ ...S.td, textAlign: "right", color: MUTED, fontWeight: 600 }}>{fmt(totGP25)}</td>
              <td style={{ ...S.td, textAlign: "right", fontWeight: 700 }}>{fmt(totGP26)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// ── Dept View ─────────────────────────────────────────────────────────────────
function DeptView({ depts }) {
  const sorted = [...depts].sort((a,b) => b.sales - a.sales);
  const maxSales = sorted[0]?.sales || 1;

  return (
    <div style={S.card}>
      <table style={S.table}>
        <thead>
          <tr>
            <th style={S.th}>Department</th>
            <th style={{ ...S.th, width: 160 }}>Sales Mix</th>
            <th style={{ ...S.th, textAlign: "right" }}>2026 Sales</th>
            <th style={{ ...S.th, textAlign: "right" }}>GP $</th>
            <th style={{ ...S.th, textAlign: "right" }}>GP %</th>
            <th style={{ ...S.th, textAlign: "right" }}>Lines</th>
            <th style={S.th}>Assessment</th>
          </tr>
        </thead>
        <tbody>
          {sorted.map((d, i) => {
            const barW = Math.round((d.sales / maxSales) * 140);
            const isUp = d.assessment.includes("⬆");
            const isDown = d.assessment.includes("⬇");
            const barColor = isUp ? GREEN : isDown ? RED : TEAL;
            return (
              <tr key={d.dept}>
                <td style={{ ...S.td, color: COLORS[i % COLORS.length], fontWeight: 600 }}>{d.dept}</td>
                <td style={S.td}>
                  <div style={{ width: barW, height: 6, background: barColor, borderRadius: 3, opacity: 0.8 }} />
                </td>
                <td style={{ ...S.td, textAlign: "right" }}>{fmt(d.sales)}</td>
                <td style={{ ...S.td, textAlign: "right" }}>{fmt(d.gp)}</td>
                <td style={{ ...S.td, textAlign: "right", color: d.gpPct < 0.08 ? RED : d.gpPct > 0.15 ? GREEN : TEXT }}>{pct(d.gpPct)}</td>
                <td style={{ ...S.td, textAlign: "right", color: MUTED }}>{d.lineItems.toLocaleString()}</td>
                <td style={{ ...S.td, fontSize: "0.65rem", color: isUp ? GREEN : isDown ? RED : MUTED, maxWidth: 200 }}>{d.assessment}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ── Action Plan View ──────────────────────────────────────────────────────────
function ActionPlanView({ actionPlan, onCustomerClick, customers, inactiveCustomers, leads, currentUser, repName }) {
  const [repFilter, setRepFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("change");
  const [userLoc, setUserLoc] = useState(null);
  const [locStatus, setLocStatus] = useState("idle"); // idle | loading | ready | denied

  const reps = ["All", ...Array.from(new Set(actionPlan.map(a=>a.salesman))).filter(Boolean).sort()];

  // Build custNum → coords map from SEED_CUSTOMERS
  const coordMap = {};
  (customers || SEED_CUSTOMERS).filter(c => c.lat && c.lon).forEach(c => { coordMap[c.num] = { lat: c.lat, lon: c.lon }; });

  function haversineAP(lat1, lon1, lat2, lon2) {
    const R = 3958.8;
    const dLat = (lat2-lat1)*Math.PI/180;
    const dLon = (lon2-lon1)*Math.PI/180;
    const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
    return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
  }

  function requestLocation() {
    if (!navigator.geolocation) { setLocStatus("denied"); return; }
    setLocStatus("loading");
    navigator.geolocation.getCurrentPosition(
      pos => { setUserLoc({ lat: pos.coords.latitude, lon: pos.coords.longitude }); setLocStatus("ready"); setSortBy("closest"); },
      () => setLocStatus("denied")
    );
  }

  // Annotate each account with distance if we have coords
  const withDist = actionPlan.map(a => {
    const coords = coordMap[a.custNum];
    const dist = (userLoc && coords) ? haversineAP(userLoc.lat, userLoc.lon, coords.lat, coords.lon) : null;
    return { ...a, _dist: dist };
  });

  const filtered = withDist
    .filter(a => repFilter === "All" || a.salesman === repFilter)
    .filter(a => !search || a.customer.toLowerCase().includes(search.toLowerCase()))
    .sort((a,b) => {
      if (sortBy === "closest") {
        if (a._dist === null && b._dist === null) return 0;
        if (a._dist === null) return 1;
        if (b._dist === null) return -1;
        return a._dist - b._dist;
      }
      if (sortBy === "change") return Math.abs(b.change)-Math.abs(a.change);
      return b.sales2026-a.sales2026;
    });

  function badge(action) {
    if (!action) return null;
    const a = action.toUpperCase();
    const color = a.includes("LOST") || a.includes("DECLINING") ? RED
      : a.includes("GROWING") || a.includes("EXPAND") ? GREEN
      : a.includes("MAINTAIN") || a.includes("STABLE") ? TEAL
      : MUTED;
    const label = a.includes("LOST") ? "LOST"
      : a.includes("DECLINING") ? "DECLINING"
      : a.includes("GROWING") ? "GROWING"
      : a.includes("MAINTAIN") ? "MAINTAIN"
      : "WATCH";
    return <span style={{ fontSize: "0.68rem", background: color, color: BG2, padding: "1px 5px", borderRadius: 2, fontWeight: 700, marginRight: 4 }}>{label}</span>;
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ display: "flex", gap: 4 }}>
          {reps.map(r => (
            <button key={r} style={S.subBtn(repFilter===r, REP_COLORS[r] || AMBER)} onClick={() => setRepFilter(r)}>{r}</button>
          ))}
        </div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search customer..." style={{ background: "#FFFFFF", border: `1px solid ${BORDER}`, color: TEXT, padding: "0.35rem 0.7rem", borderRadius: 4, fontSize: "0.7rem", width: 180 }} />
        <select
          value={sortBy}
          onChange={e => {
            if (e.target.value === "closest" && locStatus === "idle") { requestLocation(); }
            else setSortBy(e.target.value);
          }}
          style={{ background: "#FFFFFF", border: `2px solid ${sortBy === "closest" ? AMBER : BORDER}`, color: sortBy === "closest" ? AMBER : TEXT, fontWeight: sortBy === "closest" ? 700 : 400, padding: "0.35rem 0.5rem", borderRadius: 4, fontSize: "0.68rem" }}
        >
          <option value="change">Sort: $ Change</option>
          <option value="sales">Sort: 2026 Sales</option>
          <option value="closest">📍 Sort: Closest</option>
        </select>
        {locStatus === "loading" && <span style={{ fontSize: "0.68rem", color: AMBER }}>📡 Getting location…</span>}
        {locStatus === "ready"   && <span style={{ fontSize: "0.68rem", color: GREEN }}>✓ Sorted by distance</span>}
        {locStatus === "denied"  && <span style={{ fontSize: "0.68rem", color: RED }}>Location access denied</span>}
        {sortBy === "closest" && locStatus === "ready" && (
          <button onClick={() => { setUserLoc(null); setLocStatus("idle"); setSortBy("change"); }} style={{ background: "none", border: "none", color: MUTED, cursor: "pointer", fontSize: "0.7rem" }}>✕ Clear</button>
        )}
        <span style={{ color: MUTED, fontSize: "0.65rem", marginLeft: "auto" }}>{filtered.length} accounts</span>
      </div>
      <div style={S.card}>
        <div style={{ overflowX: "auto" }}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Customer</th>
                <th style={S.th}>Rep</th>
                {sortBy === "closest" && <th style={{ ...S.th, textAlign: "right", color: AMBER }}>Miles</th>}
                <th style={{ ...S.th, textAlign: "right" }}>2025</th>
                <th style={{ ...S.th, textAlign: "right" }}>2026</th>
                <th style={{ ...S.th, textAlign: "right" }}>$ Change</th>
                <th style={{ ...S.th, textAlign: "right" }}>GP %</th>
                <th style={S.th}>Status</th>
                <th style={S.th}>Top Dept</th>
                <th style={S.th}>Focus</th>
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0, 150).map((a, i) => (
                <tr key={i} onClick={() => onCustomerClick && onCustomerClick(a)}
                  style={{ cursor: onCustomerClick ? "pointer" : "default", opacity: (inactiveCustomers||{})[String(a.custNum)] ? 0.45 : 1 }}
                  onMouseEnter={e => { if(onCustomerClick) e.currentTarget.style.background="#EEF4FF"; }}
                  onMouseLeave={e => { e.currentTarget.style.background="transparent"; }}>
                  <td style={{ ...S.td, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: onCustomerClick ? AMBER : TEXT, fontWeight: onCustomerClick ? 600 : 400 }} title={a.customer}>
                    {onCustomerClick ? "↗ " : ""}{a.customer}
                    {a._carDealer && <span style={{ marginLeft: 5, fontSize: "0.6rem", background: "#FEF3C7", color: "#D97706", borderRadius: 8, padding: "1px 5px", fontWeight: 700 }}>CAR DEALER</span>}
                    {a._shared && !a._carDealer && <span style={{ marginLeft: 5, fontSize: "0.6rem", background: "#E0F7FA", color: "#0891B2", borderRadius: 8, padding: "1px 5px", fontWeight: 700 }}>SHARED</span>}
                  </td>
                  <td style={{ ...S.td, color: REP_COLORS[a.salesman] || MUTED }}>{a.salesman}</td>
                  {sortBy === "closest" && (
                    <td style={{ ...S.td, textAlign: "right", fontWeight: 700, color: a._dist !== null ? AMBER : MUTED }}>
                      {a._dist !== null ? `${a._dist.toFixed(1)} mi` : "—"}
                    </td>
                  )}
                  <td style={{ ...S.td, textAlign: "right", color: MUTED }}>{fmt(a.sales2025)}</td>
                  <td style={{ ...S.td, textAlign: "right" }}>{fmt(a.sales2026)}</td>
                  <td style={{ ...S.td, textAlign: "right", color: clr(a.change), fontWeight: 600 }}>{fmt(a.change)}</td>
                  <td style={{ ...S.td, textAlign: "right", color: a.gpPct < 0.08 ? RED : TEXT }}>{pct(a.gpPct)}</td>
                  <td style={S.td}>{badge(a.action)}</td>
                  <td style={{ ...S.td, color: MUTED, fontSize: "0.65rem", maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.topDept}</td>
                  <td style={{ ...S.td, color: MUTED, fontSize: "0.68rem", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.focus}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {filtered.length > 150 && <div style={{ color: MUTED, fontSize: "0.65rem", padding: "0.5rem 0.75rem" }}>Showing top 150 of {filtered.length} — use filter to narrow</div>}
        </div>
      </div>
    </div>
  );
}

// ── Trend Chart ───────────────────────────────────────────────────────────────
function TrendView({ weeks }) {
  return (
    <div style={S.card}>
      <div style={{ fontSize: "0.68rem", color: MUTED, marginBottom: "1rem", letterSpacing: "0.1em" }}>WEEKLY SALES TREND — 2025 vs 2026</div>
      <ResponsiveContainer width="100%" height={300}>
        <AreaChart data={weeks} margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
          <defs>
            <linearGradient id="g25" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={MUTED} stopOpacity={0.3} />
              <stop offset="95%" stopColor={MUTED} stopOpacity={0} />
            </linearGradient>
            <linearGradient id="g26" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={AMBER} stopOpacity={0.4} />
              <stop offset="95%" stopColor={AMBER} stopOpacity={0} />
            </linearGradient>
          </defs>
          <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
          <XAxis dataKey="week" tick={{ fill: MUTED, fontSize: 11 }} />
          <YAxis tickFormatter={v => fmt(v,"$")} tick={{ fill: MUTED, fontSize: 11 }} />
          <Tooltip contentStyle={{ background: BG2, border: `1px solid ${BORDER}`, fontSize: 12 }}
            formatter={(v, n) => [fmt(v), n === "sales2025" ? "2025" : "2026"]} />
          <Area type="monotone" dataKey="sales2025" stroke={MUTED} fill="url(#g25)" name="sales2025" strokeWidth={1.5} dot={false} />
          <Area type="monotone" dataKey="sales2026" stroke={AMBER} fill="url(#g26)" name="sales2026" strokeWidth={2} dot={{ fill: AMBER, r: 3 }} />
        </AreaChart>
      </ResponsiveContainer>
      <div style={{ display: "flex", gap: 16, marginTop: "0.5rem", fontSize: "0.65rem" }}>
        <span><span style={{ color: MUTED }}>▬</span> 2025</span>
        <span><span style={{ color: AMBER }}>▬</span> 2026</span>
      </div>
    </div>
  );
}

// ── Rep Tab ───────────────────────────────────────────────────────────────────
function RepTab({ repName, weekComp, onAskAI, onCustomerClick, customers, inactiveCustomers, leads, onAddLead, onDeleteLead, currentUser, onLogActivity, onRefreshLeads }) {
  const color = REP_COLORS[repName] || AMBER;
  const [cityFilter, setCityFilter] = useState("All");
  const [showShared, setShowShared] = useState(true);

  // Austin gets shared accounts from other reps where Industrial, OTR, or Farm is the top dept
  const AUSTIN_TARGET_DEPTS = ["INDUSTRIAL TIRES", "OFF THE ROAD TIRES", "FARM TIRES"];
  const isAustin = repName === "Austin";
  const repLeads = (leads||[]).filter(l => {
    if (currentUser?.id === "admin") return true;                         // admin sees all
    const assignedTo     = String(l.assigned_to      || "").toLowerCase();
    const assignedName   = String(l.assigned_to_name || "").toLowerCase();
    const userId         = String(currentUser?.id    || "").toLowerCase();
    const userName       = String(currentUser?.name  || "").toLowerCase();
    // Rep only sees leads explicitly assigned to them (by id or name)
    return assignedTo === userId || assignedTo === userName
        || assignedName === userName || assignedName === userId;
  });


  const ownAccounts = (weekComp?.actionPlan || []).filter(a => a.salesman.toLowerCase() === repName.toLowerCase());

  const austinShared = isAustin
    ? (weekComp?.actionPlan || []).filter(a =>
        a.salesman !== "Austin" &&
        AUSTIN_TARGET_DEPTS.some(d => (a.topDept || "").toUpperCase().includes(d))
      ).map(a => ({ ...a, _shared: true }))
    : [];

  const sharedAccounts = austinShared;

  const actionPlan = isAustin && showShared
    ? [...ownAccounts, ...sharedAccounts]
    : ownAccounts;
  const total25 = actionPlan.reduce((s,a)=>s+a.sales2025,0);
  const total26 = actionPlan.reduce((s,a)=>s+a.sales2026,0);
  const growing = actionPlan.filter(a=>a.action.toUpperCase().includes("GROW")).length;
  const declining = actionPlan.filter(a=>a.action.toUpperCase().includes("DECLIN") || a.action.toUpperCase().includes("LOST")).length;

  // Build sorted city list from this rep's accounts
  const cities = ["All", ...Array.from(new Set(actionPlan.map(a => a.city).filter(Boolean))).sort()];

  // City-filtered subset for KPIs and table
  const filtered = cityFilter === "All" ? actionPlan : actionPlan.filter(a => a.city === cityFilter);
  const filtTotal25 = filtered.reduce((s,a)=>s+a.sales2025,0);
  const filtTotal26 = filtered.reduce((s,a)=>s+a.sales2026,0);
  const filtGrowing = filtered.filter(a=>a.action.toUpperCase().includes("GROW")).length;
  const filtDeclining = filtered.filter(a=>a.action.toUpperCase().includes("DECLIN") || a.action.toUpperCase().includes("LOST")).length;

  const isFiltered = cityFilter !== "All";
  const [repSubTab, setRepSubTab] = useState("accounts");

  return (
    <div>
      <div style={S.kpiRow}>
        <div style={S.kpi(color)}>
          <div style={S.kpiVal}>{fmt(isFiltered ? filtTotal26 : total26)}</div>
          <div style={S.kpiLbl}>{isFiltered ? `${cityFilter} 2026 Sales` : "2026 YTD Sales"}</div>
        </div>
        <div style={S.kpi(clr((isFiltered ? filtTotal26 : total26)-(isFiltered ? filtTotal25 : total25)))}>
          <div style={{ ...S.kpiVal, color: clr((isFiltered?filtTotal26:total26)-(isFiltered?filtTotal25:total25)) }}>
            {fmt((isFiltered?filtTotal26:total26)-(isFiltered?filtTotal25:total25))}
          </div>
          <div style={S.kpiLbl}>vs 2025 ({pct((isFiltered?filtTotal25:total25) ? ((isFiltered?filtTotal26:total26)-(isFiltered?filtTotal25:total25))/(isFiltered?filtTotal25:total25) : 0)})</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>{isFiltered ? filtGrowing : growing}</div>
          <div style={S.kpiLbl}>Growing Accounts</div>
        </div>
        <div style={S.kpi(RED)}>
          <div style={S.kpiVal}>{isFiltered ? filtDeclining : declining}</div>
          <div style={S.kpiLbl}>Declining Accounts</div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.75rem", flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div style={{ fontSize: "0.75rem", fontWeight: 600, color: color }}>
            {repName.toUpperCase()} — {isFiltered ? cityFilter : "All Cities"}
            <span style={{ fontWeight: 400, color: MUTED, marginLeft: 6 }}>({filtered.length} accounts)</span>
          </div>
          {isAustin && sharedAccounts.length > 0 && (
            <button
              onClick={() => setShowShared(s => !s)}
              style={{
                padding: "0.3rem 0.75rem", borderRadius: 20,
                border: `2px solid ${showShared ? color : BORDER}`,
                background: showShared ? "#EEF4FF" : "#FFFFFF",
                color: showShared ? color : MUTED,
                fontWeight: 700, fontSize: "0.7rem", cursor: "pointer",
              }}
            >
              {showShared ? "✓ Shared Accts ON" : "Shared Accts OFF"}
              {showShared && <span style={{ marginLeft: 6, background: color, color: "#fff", borderRadius: 10, padding: "1px 6px", fontSize: "0.65rem" }}>{sharedAccounts.length}</span>}
            </button>
          )}
          {/* City dropdown */}
          <div style={{ position: "relative", display: "inline-block" }}>
            <select
              value={cityFilter}
              onChange={e => setCityFilter(e.target.value)}
              style={{
                appearance: "none",
                background: "#FFFFFF",
                border: `2px solid ${isFiltered ? color : BORDER}`,
                borderRadius: 6,
                color: isFiltered ? color : TEXT,
                fontWeight: isFiltered ? 700 : 400,
                padding: "0.4rem 2rem 0.4rem 0.75rem",
                fontSize: "0.75rem",
                cursor: "pointer",
                outline: "none",
                minWidth: 160,
              }}
            >
              {cities.map(c => <option key={c} value={c}>{c === "All" ? "📍 All Cities" : `📍 ${c}`}</option>)}
            </select>
            <span style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: isFiltered ? color : MUTED, fontSize: "0.7rem" }}>▼</span>
          </div>
          {isFiltered && (
            <button
              onClick={() => setCityFilter("All")}
              style={{ background: "none", border: "none", color: MUTED, cursor: "pointer", fontSize: "0.75rem", padding: "0.2rem 0.4rem" }}
            >✕ Clear</button>
          )}
        </div>
        <button style={S.btn(color)} onClick={() => onAskAI(`Analyze ${repName}'s performance${isFiltered ? ` in ${cityFilter}` : ""}. YTD 2026: ${fmt(isFiltered?filtTotal26:total26)}, vs 2025: ${pct((isFiltered?filtTotal25:total25) ? ((isFiltered?filtTotal26:total26)-(isFiltered?filtTotal25:total25))/(isFiltered?filtTotal25:total25) : 0)}. ${isFiltered?filtGrowing:growing} growing, ${isFiltered?filtDeclining:declining} declining accounts. Identify top opportunities and risks.`)}>
          ◈ AI Analysis
        </button>
      </div>

      {/* Rep sub-nav */}
      <div style={{ ...S.subNav, marginBottom:"0.75rem" }}>
        <button style={S.subBtn(repSubTab==="accounts", color)} onClick={()=>setRepSubTab("accounts")}>📋 Accounts</button>
        <button style={S.subBtn(repSubTab==="calllog",  "#0891B2")} onClick={()=>setRepSubTab("calllog")}>📞 Call Log</button>
        <button style={S.subBtn(repSubTab==="leads",    "#059669")} onClick={()=>setRepSubTab("leads")}>
          🎯 Leads
          {repLeads.filter(l=>l.status==="open").length > 0 && <span style={{ marginLeft:5, background:RED, color:"#fff", borderRadius:8, padding:"0px 6px", fontSize:"0.6rem", fontWeight:700 }}>{repLeads.filter(l=>l.status==="open").length}</span>}
        </button>
        <button style={S.subBtn(repSubTab==="ad",       "#7C3AED")} onClick={()=>setRepSubTab("ad")}>🏆 AD Programs</button>
        <button style={S.subBtn(repSubTab==="todo",     AMBER)}    onClick={()=>setRepSubTab("todo")}>📋 To Do</button>
        {isAustin && <button style={S.subBtn(repSubTab==="statesboro", "#0891B2")} onClick={()=>setRepSubTab("statesboro")}>📍 Statesboro</button>}
        <button style={S.subBtn(repSubTab==="specials", "#D97706")} onClick={()=>setRepSubTab("specials")}>🏷 Specials</button>
      </div>

      {repSubTab === "accounts" && (
        <ActionPlanView actionPlan={filtered} onCustomerClick={onCustomerClick} customers={customers} inactiveCustomers={inactiveCustomers} leads={leads} currentUser={currentUser} repName={repName} />
      )}
      {repSubTab === "ad" && (
        <RepADTab actionPlan={actionPlan} repName={repName} color={color} onCustomerClick={onCustomerClick} />
      )}
      {repSubTab === "todo" && (
        <RepTodoTab repName={repName} actionPlan={actionPlan} onCustomerClick={onCustomerClick} color={color} leads={leads} currentUser={currentUser} />
      )}
      {repSubTab === "calllog" && (
        <RepCallLog repName={repName} actionPlan={actionPlan} currentUser={currentUser} onLogActivity={onLogActivity} />
      )}
      {repSubTab === "leads" && (
        <LeadsTab leads={repLeads} repName={repName} onAddLead={onAddLead} onDeleteLead={onDeleteLead} currentUser={currentUser} isAdmin={false} allReps={["Tiffany","Larry","Austin"]} onRefreshLeads={onRefreshLeads} />
      )}
      {repSubTab === "statesboro" && isAustin && (
        <StatesboroTab />
      )}
      {repSubTab === "specials" && (
        <SpecialsTab repName={repName} color={color} />
      )}
    </div>
  );
}


// ── Rep AD Programs Tab ───────────────────────────────────────────────────────
function RepADTab({ actionPlan, repName, color, onCustomerClick }) {
  const ALL_PROGRAMS = [
    {
      key: "toyo", label: "Toyo AD", icon: "🏆", color: "#1E5FCC",
      data: AD_PROGRAMS, tiers: AD_PCR_TIERS,
      getUnits: d => d.pcr.total,
      getSub:   d => `PCR: ${d.pcr.total} units · ${d.pcr.pct}% primary${d.tbr.total>0?" · TBR: "+d.tbr.total:""}`,
    },
    {
      key: "ascenso", label: "Ascenso", icon: "🌀", color: "#059669",
      data: ASCENSO_PROGRAMS, tiers: ASCENSO_TIERS,
      getUnits: d => ASCENSO_TOTAL,  // company-wide dollar threshold
      getSub:   d => `${fmt(d.amount)} · ${d.qty} units · ${d.invoices} invoices`,
      isDollar: true,
      getDisplayVal: d => fmt(d.amount),
      getDisplayLabel: () => "This Acct Revenue",
    },
    {
      key: "americus", label: "Americus", icon: "🔵", color: "#0891B2",
      data: AMERICUS_PROGRAMS, tiers: AMERICUS_TIERS,
      getUnits: d => d.ytd,
      getSub:   d => `YTD: ${d.ytd} units · Q1: ${d.q1} · Q2: ${d.q2}`,
    },
    {
      key: "barnn", label: "BF BARNN", icon: "🔶", color: "#D97706",
      data: BARNN_PROGRAMS, tiers: BARNN_TIERS,
      getUnits: d => d.ytd,
      getSub:   d => `YTD: ${d.ytd} · PY: ${d.py} · ${d.ytd>d.py?"+":""}${d.ytd-d.py} units`,
    },
    {
      key: "falkenplt", label: "Falken PLT", icon: "🟢", color: "#059669",
      data: FALKEN_PLT_PROGRAMS, tiers: FALKEN_PLT_TIERS,
      getUnits: d => d.ytd,
      getSub:   d => `Q2: ${d.q2} units (Q1: ${d.q1})`,
    },
  ];

  // Build list of all AD-enrolled customers for this rep
  const repCustNums = new Set(actionPlan.map(a => String(a.custNum)));

  const enrolled = [];
  ALL_PROGRAMS.forEach(prog => {
    Object.entries(prog.data).forEach(([custNum, pData]) => {
      if (!repCustNums.has(custNum)) return;
      const ap = actionPlan.find(a => String(a.custNum) === custNum);
      if (!ap) return;
      const units   = prog.getUnits(pData);
      const tier    = getAdTier(units, prog.tiers);
      const nxtTier = getNextAdTier(units, prog.tiers);
      const toNext  = nxtTier ? nxtTier.min - units : 0;
      const tierMin = tier.min;
      const gap     = nxtTier ? nxtTier.min - tierMin : 0;
      const isClose = nxtTier && toNext <= Math.max(5, Math.round(gap * 0.15));
      const belowPrimary = prog.key === "toyo" && pData.pcr && pData.pcr.pct < 75;
      enrolled.push({
        custNum, custName: ap.customer, city: ap.city,
        prog, pData, units, tier, nxtTier, toNext, isClose, gap,
        ap, belowPrimary,
      });
    });
  });

  // Dedupe: group by custNum, collect all programs per customer
  const byCustomer = {};
  enrolled.forEach(e => {
    if (!byCustomer[e.custNum]) {
      byCustomer[e.custNum] = { custNum: e.custNum, custName: e.custName, city: e.city, ap: e.ap, programs: [] };
    }
    byCustomer[e.custNum].programs.push(e);
  });
  const customers = Object.values(byCustomer).sort((a,b) => b.programs.length - a.programs.length);

  // Close-to-tier alerts (any program)
  const closeAlerts = enrolled.filter(e => e.isClose || e.belowPrimary).sort((a,b) => {
    if (a.belowPrimary && !b.belowPrimary) return -1;
    if (!a.belowPrimary && b.belowPrimary) return 1;
    return a.toNext - b.toNext;
  });

  if (customers.length === 0) {
    return (
      <div style={{ ...S.card, textAlign:"center", padding:"2.5rem", color:MUTED }}>
        <div style={{ fontSize:"2rem", marginBottom:"0.5rem" }}>🏆</div>
        <div style={{ fontWeight:600, marginBottom:4 }}>No AD program accounts for {repName}</div>
        <div style={{ fontSize:"0.72rem" }}>Upload manufacturer program reports to see data here</div>
      </div>
    );
  }

  return (
    <div>
      {/* ── Close-to-tier alert box ── */}
      {closeAlerts.length > 0 && (
        <div style={{ ...S.card, background:"linear-gradient(135deg,#FEF3C7,#FFF7ED)", border:"2px solid #FCD34D", marginBottom:"1.25rem" }}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:"0.75rem" }}>
            <span style={{ fontSize:"1.1rem" }}>⚡</span>
            <span style={{ fontSize:"0.78rem", fontWeight:700, color:"#92400E" }}>
            ACTION NEEDED — {closeAlerts.length} alert{closeAlerts.length>1?"s":""}
            {enrolled.filter(e=>e.belowPrimary).length > 0 && <span style={{ marginLeft:8, color:"#DC2626" }}>({enrolled.filter(e=>e.belowPrimary).length} below Toyo 75% threshold)</span>}
          </span>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))", gap:"0.6rem" }}>
            {closeAlerts.map((e, i) => {
              const tColor = e.nxtTier?.color === "#6B7A99" ? MUTED : (e.nxtTier?.color || MUTED);
              return (
                <div key={i}
                  onClick={() => onCustomerClick && onCustomerClick(e.ap)}
                  style={{ background:"#FFFFFF", borderRadius:8, padding:"0.6rem 0.85rem", border:`2px solid ${tColor}`, cursor:"pointer", display:"flex", flexDirection:"column", gap:3 }}
                  onMouseEnter={e2=>e2.currentTarget.style.background="#FFFBEB"}
                  onMouseLeave={e2=>e2.currentTarget.style.background="#FFFFFF"}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <span style={{ fontSize:"0.72rem", fontWeight:700, color:TEXT }}>{e.custName}</span>
                    <span style={{ fontSize:"0.65rem", color:MUTED }}>{e.city}</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap" }}>
                    <span style={{ fontSize:"0.65rem", background:e.prog.color+"22", color:e.prog.color, borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{e.prog.icon} {e.prog.label}</span>
                    <span style={{ fontSize:"0.68rem", color:TEXT, fontWeight:600 }}>{e.tier.label}</span>
                    {e.belowPrimary
                      ? <span style={{ fontSize:"0.72rem", fontWeight:800, color:"#DC2626" }}>🚨 {e.pData.pcr.pct}% primary — below 75% rebate threshold</span>
                      : <><span style={{ fontSize:"0.68rem" }}>→</span>
                        <span style={{ fontSize:"0.72rem", fontWeight:800, color:tColor }}>⚡ {e.toNext}{e.prog.isDollar?"":" units"} to {e.nxtTier?.label}</span></>
                    }
                  </div>
                  <div style={{ marginTop:4, height:4, background:"#F3F4F6", borderRadius:2 }}>
                    <div style={{ height:4, background:tColor, borderRadius:2, width:`${Math.min(100,Math.round(((e.units-e.tier.min)/Math.max(1,e.gap))*100))}%` }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── Customer program cards ── */}
      <div style={{ fontSize:"0.7rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>
        {customers.length} enrolled customer{customers.length>1?"s":" "} · {enrolled.length} program enrollment{enrolled.length>1?"s":""}
      </div>

      <div style={{ display:"flex", flexDirection:"column", gap:"0.85rem" }}>
        {customers.map(cust => (
          <div key={cust.custNum} style={{ ...S.card, padding:"0.85rem 1rem" }}>
            {/* Customer header */}
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"0.75rem" }}>
              <div>
                <div
                  onClick={() => onCustomerClick && onCustomerClick(cust.ap)}
                  style={{ fontSize:"0.85rem", fontWeight:700, color:AMBER, cursor:"pointer", display:"inline-flex", alignItems:"center", gap:6 }}
                  onMouseEnter={e=>e.target.style.textDecoration="underline"}
                  onMouseLeave={e=>e.target.style.textDecoration="none"}>
                  ↗ {cust.custName}
                </div>
                <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:2 }}>📍 {cust.city}</div>
              </div>
              <div style={{ display:"flex", gap:4, flexWrap:"wrap", justifyContent:"flex-end" }}>
                {cust.programs.map((e,i) => (
                  <span key={i} style={{ fontSize:"0.65rem", background:e.prog.color+"18", color:e.prog.color, border:`1px solid ${e.prog.color}44`, borderRadius:8, padding:"2px 7px", fontWeight:700, whiteSpace:"nowrap" }}>
                    {e.prog.icon} {e.prog.label}
                  </span>
                ))}
              </div>
            </div>

            {/* Program rows */}
            <div style={{ display:"flex", flexDirection:"column", gap:"0.5rem" }}>
              {cust.programs.map((e, i) => {
                const pColor = e.prog.color;
                const tColor = e.tier.color === "#6B7A99" ? MUTED : e.tier.color;
                const nColor = e.nxtTier?.color === "#6B7A99" ? MUTED : (e.nxtTier?.color || MUTED);
                const barPct = e.nxtTier ? Math.min(100, Math.round(((e.units-e.tier.min)/Math.max(1,e.gap))*100)) : 100;
                return (
                  <div key={i} style={{ background:"#F8FAFF", borderRadius:6, padding:"0.55rem 0.75rem", borderLeft:`3px solid ${pColor}` }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <span style={{ fontSize:"0.7rem", fontWeight:700, color:pColor }}>{e.prog.icon} {e.prog.label}</span>
                        <span style={{ fontSize:"0.65rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"1px 6px", borderRadius:8 }}>{e.tier.label}</span>
                        {e.isClose && <span style={{ fontSize:"0.65rem", fontWeight:700, color:"#D97706", background:"#FEF3C7", padding:"1px 6px", borderRadius:8 }}>⚡ {e.toNext} to {e.nxtTier?.label}!</span>}
                      </div>
                      {!e.nxtTier && <span style={{ fontSize:"0.65rem", color:"#7C3AED", fontWeight:700 }}>★ Top Tier</span>}
                    </div>
                    <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:5 }}>{e.prog.getSub(e.pData)}</div>
                    {/* Mini progress bar */}
                    <div style={{ height:4, background:BORDER, borderRadius:2 }}>
                      <div style={{ height:4, background: e.isClose?"linear-gradient(90deg,#D97706,#FBBF24)":pColor, borderRadius:2, width:`${barPct}%`, transition:"width 0.4s ease" }} />
                    </div>
                    {e.nxtTier && (
                      <div style={{ display:"flex", justifyContent:"space-between", marginTop:3 }}>
                        <span style={{ fontSize:"0.58rem", color:tColor, fontWeight:600 }}>{e.tier.label} ({e.tier.min}{e.prog.isDollar?"":""}+)</span>
                        <span style={{ fontSize:"0.58rem", color:nColor, fontWeight:600 }}>{e.nxtTier.label} ({e.nxtTier.min}{e.prog.isDollar?"":""}+)</span>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── AI Tab ────────────────────────────────────────────────────────────────────
function AITab({ weekComp, initialPrompt, onClearPrompt }) {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const bottomRef = useRef(null);

  useEffect(() => {
    if (initialPrompt) {
      setInput(initialPrompt);
      onClearPrompt();
    }
  }, [initialPrompt]);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages]);

  // Auto-generate summary on load
  useEffect(() => {
    if (!weekComp?.weeks?.length || messages.length > 0) return;
    const weeks = weekComp.weeks;
    const tot25 = weeks.reduce((s,w)=>s+w.sales2025,0);
    const tot26 = weeks.reduce((s,w)=>s+w.sales2026,0);
    const topDepts = (weekComp.depts||[]).slice(0,5).map(d=>`${d.dept}: ${fmt(d.sales)} (GP ${pct(d.gpPct)})`).join(", ");
    const declining = (weekComp.actionPlan||[]).filter(a=>a.action.toUpperCase().includes("DECLIN")||a.action.toUpperCase().includes("LOST")).length;
    const prompt = `You are a sales analytics assistant for a multi-line tire and ag distributor. Provide a concise executive summary (5-7 bullets) of this week-by-week sales data:\n\n- YTD 2025: $${tot25.toFixed(0)}\n- YTD 2026: $${tot26.toFixed(0)}\n- Change: ${((tot26-tot25)/tot25*100).toFixed(1)}%\n- Weeks on file: ${weeks.length} (W1-${weeks[weeks.length-1]?.week})\n- Top departments: ${topDepts}\n- Declining accounts: ${declining}\n- Latest week: ${weeks[weeks.length-1]?.week} = $${weeks[weeks.length-1]?.sales2026.toFixed(0)}\n\nFocus on trend momentum, standout weeks (best/worst), department performance, and urgent actions needed.`;
    sendMessage(prompt, true);
  }, [weekComp]);

  async function sendMessage(text, isAuto = false) {
    const userMsg = isAuto ? null : { role: "user", content: text };
    if (!isAuto) setMessages(prev => [...prev, userMsg]);
    setLoading(true);
    setInput("");

    const systemCtx = weekComp ? `You have access to week-by-week sales comparison data (2025 vs 2026) for a multi-line tire and ag distributor. ${weekComp.weeks?.length || 0} weeks on file. YTD 2026: $${(weekComp.weeks||[]).reduce((s,w)=>s+w.sales2026,0).toFixed(0)}. Be concise and actionable.` : "You are a sales analytics assistant.";

    // Strip ALL extra fields — only role and content are allowed by the API
    const cleanMessages = messages
      .filter(m => m.role && m.content)
      .map(m => ({ role: m.role, content: String(m.content) }));
    const apiMessages = [
      ...(isAuto ? [] : cleanMessages),
      { role: "user", content: text }
    ];

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system: systemCtx,
          messages: apiMessages,
        }),
      });
      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`API ${res.status}: ${errText}`);
      }
      const data = await res.json();
      const reply = data.content?.[0]?.text || "No response.";
      setMessages(prev => [
        ...(isAuto ? [{ role: "assistant", content: reply, auto: true }] : [...prev, { role: "assistant", content: reply }])
      ]);
    } catch(e) {
      setMessages(prev => [...prev, { role: "assistant", content: `Error: ${e.message || "Could not connect to AI."}` }]);
    }
    setLoading(false);
  }

  const quickPrompts = [
    "Which weeks had the biggest swings? What might explain them?",
    "Which departments should we focus on to improve GP margin?",
    "Who are the top 5 accounts at risk of being lost?",
    "Summarize each rep's performance and top action item.",
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 160px)" }}>
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: "1rem" }}>
        {messages.length === 0 && !loading && (
          <div style={{ ...S.card, color: MUTED, fontSize: "0.72rem" }}>AI analysis will auto-generate from your week comp data. Ask any question about your sales.</div>
        )}
        {messages.map((m, i) => (
          <div key={i} style={{ marginBottom: "0.75rem", display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start" }}>
            <div style={{
              maxWidth: "80%", padding: "0.7rem 1rem", borderRadius: 6,
              background: m.role === "user" ? AMBER : BG2,
              color: m.role === "user" ? BG2 : TEXT,
              border: m.role === "assistant" ? `1px solid ${BORDER}` : "none",
              fontSize: "0.74rem", lineHeight: 1.6, whiteSpace: "pre-wrap",
              ...(m.auto ? { borderLeft: `3px solid ${TEAL}` } : {}),
            }}>
              {m.auto && <div style={{ fontSize: "0.68rem", color: TEAL, marginBottom: 4, letterSpacing: "0.1em" }}>◈ AUTO ANALYSIS</div>}
              {m.content}
            </div>
          </div>
        ))}
        {loading && (
          <div style={{ display: "flex", gap: 4, padding: "0.75rem 1rem" }}>
            {[0,1,2].map(i => <span key={i} style={{ width: 6, height: 6, background: AMBER, borderRadius: "50%", animation: "pulse 1s infinite", animationDelay: `${i*0.2}s` }} />)}
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      <div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: "0.5rem" }}>
          {quickPrompts.map((q, i) => (
            <button key={i} style={{ ...S.btn(MUTED), fontSize: "0.68rem" }} onClick={() => sendMessage(q)}>{q}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            value={input} onChange={e=>setInput(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&input.trim()&&sendMessage(input)}
            placeholder="Ask about your sales data..."
            style={{ flex: 1, background: "#FFFFFF", border: `1px solid ${BORDER}`, color: TEXT, padding: "0.65rem 1rem", borderRadius: 4, fontSize: "0.74rem", outline: "none" }}
          />
          <button style={{ ...S.btn(AMBER), padding: "0.65rem 1.25rem" }} onClick={() => input.trim() && sendMessage(input)} disabled={loading}>
            Send
          </button>
        </div>
      </div>
      <style>{`@keyframes pulse { 0%,100%{opacity:0.3} 50%{opacity:1} }`}</style>
    </div>
  );
}

// ── Map Tab ───────────────────────────────────────────────────────────────────
function MapTab({ customers, weekComp }) {
  const [userLoc, setUserLoc] = useState(null);
  const [nearest, setNearest] = useState([]);
  const [repFilter, setRepFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("name");
  const [locError, setLocError] = useState("");

  // Only active customers with real coordinates
  const activeCustomers = (customers || []).filter(c =>
    c.active === "Active" && c.lat && c.lon
  );

  // Sales data lookup by customer number
  const apByNum = {};
  (weekComp?.actionPlan || []).forEach(a => { apByNum[a.custNum] = a; });

  function haversine(lat1, lon1, lat2, lon2) {
    const R = 3958.8;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  }

  function getMyLocation() {
    setLocError("");
    if (!navigator.geolocation) { setLocError("Geolocation not supported by your browser."); return; }
    navigator.geolocation.getCurrentPosition(
      pos => {
        const loc = { lat: pos.coords.latitude, lon: pos.coords.longitude };
        setUserLoc(loc);
        const near = activeCustomers
          .map(c => ({ ...c, dist: haversine(loc.lat, loc.lon, c.lat, c.lon), ap: apByNum[c.num] }))
          .filter(c => repFilter === "All" || c.salesman === repFilter)
          .sort((a, b) => a.dist - b.dist)
          .slice(0, 25);
        setNearest(near);
      },
      err => setLocError("Location access denied. Please allow location in your browser.")
    );
  }

  const reps = ["All", ...Array.from(new Set(activeCustomers.map(c => c.salesman))).filter(Boolean).sort()];

  const filtered = activeCustomers
    .filter(c => repFilter === "All" || c.salesman === repFilter)
    .filter(c => !search || c.name.toLowerCase().includes(search.toLowerCase()) || c.city.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      if (sortBy === "name") return a.name.localeCompare(b.name);
      if (sortBy === "city") return a.city.localeCompare(b.city);
      const aAp = apByNum[a.num]; const bAp = apByNum[b.num];
      if (sortBy === "sales") return (bAp?.sales2026||0) - (aAp?.sales2026||0);
      return 0;
    });

  function actionBadge(ap) {
    if (!ap?.action) return null;
    const a = ap.action.toUpperCase();
    const color = a.includes("LOST")||a.includes("DECLIN") ? RED : a.includes("GROW") ? GREEN : a.includes("MAINTAIN") ? TEAL : MUTED;
    const label = a.includes("LOST") ? "LOST" : a.includes("DECLIN") ? "DECLINING" : a.includes("GROW") ? "GROWING" : "MAINTAIN";
    return <span style={{ fontSize:"0.55rem", background:color, color:BG2, padding:"1px 4px", borderRadius:2, fontWeight:700 }}>{label}</span>;
  }

  return (
    <div>
      {/* Stats bar */}
      <div style={S.kpiRow}>
        <div style={S.kpi(AMBER)}>
          <div style={S.kpiVal}>{activeCustomers.length}</div>
          <div style={S.kpiLbl}>Active Customers</div>
        </div>
        <div style={S.kpi(TEAL)}>
          <div style={S.kpiVal}>{[...new Set(activeCustomers.map(c=>c.city))].length}</div>
          <div style={S.kpiLbl}>Cities</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>{[...new Set(activeCustomers.map(c=>c.state))].length}</div>
          <div style={S.kpiLbl}>States</div>
        </div>
        <div style={S.kpi(MUTED)}>
          <div style={S.kpiVal}>{nearest.length > 0 ? `${nearest[0].dist.toFixed(1)} mi` : "—"}</div>
          <div style={S.kpiLbl}>Nearest Customer</div>
        </div>
      </div>

      {/* Controls */}
      <div style={{ ...S.card, marginBottom: "0.75rem" }}>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
          <div>
            <button style={{ ...S.btn(TEAL), padding:"0.5rem 1.1rem", fontSize:"0.74rem" }} onClick={getMyLocation}>
              ◉ Find Nearest Customers
            </button>
            {locError && <div style={{ fontSize:"0.65rem", color:RED, marginTop:4 }}>{locError}</div>}
            {userLoc && !locError && <div style={{ fontSize:"0.62rem", color:GREEN, marginTop:4 }}>✓ Location found — showing {nearest.length} nearest</div>}
          </div>
          <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
            {reps.map(r => (
              <button key={r} style={S.subBtn(repFilter===r, REP_COLORS[r]||AMBER)} onClick={()=>{ setRepFilter(r); if(userLoc){ /* re-sort nearest */ const loc=userLoc; const near=activeCustomers.filter(c=>r==="All"||c.salesman===r).map(c=>({...c,dist:haversine(loc.lat,loc.lon,c.lat,c.lon),ap:apByNum[c.num]})).sort((a,b)=>a.dist-b.dist).slice(0,25); setNearest(near); } }}>{r}</button>
            ))}
          </div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name or city..."
            style={{ background:BG, border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.7rem", borderRadius:4, fontSize:"0.7rem", width:190 }} />
          <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
            style={{ background:BG, border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.5rem", borderRadius:4, fontSize:"0.68rem" }}>
            <option value="name">Sort: Name</option>
            <option value="city">Sort: City</option>
            <option value="sales">Sort: 2026 Sales</option>
          </select>
          <span style={{ color:MUTED, fontSize:"0.65rem", marginLeft:"auto" }}>{filtered.length} customers</span>
        </div>
      </div>

      {/* Nearest panel */}
      {nearest.length > 0 && (
        <div style={{ ...S.card, marginBottom:"0.75rem" }}>
          <div style={{ fontSize:"0.68rem", color:TEAL, letterSpacing:"0.1em", marginBottom:"0.6rem" }}>
            ◉ NEAREST {nearest.length} CUSTOMERS TO YOUR LOCATION
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))", gap:"0.5rem" }}>
            {nearest.map((c,i) => (
              <div key={c.num} style={{ background:BG, border:`1px solid ${i===0?TEAL:BORDER}`, borderRadius:4, padding:"0.65rem 0.85rem", borderLeft:`3px solid ${REP_COLORS[c.salesman]||MUTED}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                  <span style={{ fontSize:"0.73rem", fontWeight:600, color:TEXT, lineHeight:1.3 }}>{c.name}</span>
                  <span style={{ fontSize:"0.7rem", color:i<3?TEAL:MUTED, fontWeight:700, marginLeft:8, whiteSpace:"nowrap" }}>{c.dist.toFixed(1)} mi</span>
                </div>
                <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:3 }}>{c.city}, {c.state} · <span style={{ color:REP_COLORS[c.salesman]||MUTED }}>{c.salesman}</span></div>
                {c.address && <div style={{ fontSize:"0.6rem", color:MUTED, marginTop:1 }}>{c.address}</div>}
                {c.phone && <div style={{ fontSize:"0.6rem", color:MUTED, marginTop:1 }}>📞 {c.phone}</div>}
                {c.ap && (
                  <div style={{ display:"flex", alignItems:"center", gap:6, marginTop:4 }}>
                    {actionBadge(c.ap)}
                    <span style={{ fontSize:"0.62rem", color:clr(c.ap.change) }}>{fmt(c.ap.sales2026)} ({c.ap.change>=0?"+":""}{fmt(c.ap.change)})</span>
                  </div>
                )}
                <div style={{ marginTop:5 }}>
                  <a href={`https://www.google.com/maps/dir/?api=1&destination=${c.lat},${c.lon}`} target="_blank" rel="noopener noreferrer"
                    style={{ fontSize:"0.6rem", color:AMBER, textDecoration:"none" }}>↗ Directions</a>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Full customer table */}
      <div style={S.card}>
        <div style={{ overflowX:"auto", maxHeight:520, overflowY:"auto" }}>
          <table style={S.table}>
            <thead style={{ position:"sticky", top:0, background:BG2 }}>
              <tr>
                <th style={S.th}>Name</th>
                <th style={S.th}>City</th>
                <th style={S.th}>ST</th>
                <th style={S.th}>Rep</th>
                <th style={{ ...S.th, textAlign:"right" }}>2026 Sales</th>
                <th style={{ ...S.th, textAlign:"right" }}>vs 2025</th>
                <th style={S.th}>Status</th>
                <th style={{ ...S.th, textAlign:"right" }}>Lat</th>
                <th style={{ ...S.th, textAlign:"right" }}>Lon</th>
              </tr>
            </thead>
            <tbody>
              {filtered.slice(0, 300).map(c => {
                const ap = apByNum[c.num];
                return (
                  <tr key={c.num}>
                    <td style={{ ...S.td, maxWidth:190, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={c.name}>{c.name}</td>
                    <td style={S.td}>{c.city}</td>
                    <td style={{ ...S.td, color:MUTED }}>{c.state}</td>
                    <td style={{ ...S.td, color:REP_COLORS[c.salesman]||MUTED }}>{c.salesman}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>{ap ? fmt(ap.sales2026) : "—"}</td>
                    <td style={{ ...S.td, textAlign:"right", color:ap?clr(ap.change):MUTED }}>{ap ? fmt(ap.change) : "—"}</td>
                    <td style={S.td}>{actionBadge(ap)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:TEAL, fontSize:"0.65rem" }}>{c.lat.toFixed(4)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:TEAL, fontSize:"0.65rem" }}>{c.lon.toFixed(4)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {filtered.length > 300 && <div style={{ color:MUTED, fontSize:"0.65rem", padding:"0.5rem 0.75rem" }}>Showing 300 of {filtered.length} — use filters to narrow</div>}
        </div>
      </div>
    </div>
  );
}






// ── Rep Call Log ──────────────────────────────────────────────────────────────
function RepCallLog({ repName, actionPlan, currentUser, onLogActivity }) {
  const CALL_LOG_KEY = `call_log_${currentUser?.id || repName}`;
  const [entries, setEntries]   = useState(() => {
    try { return JSON.parse(localStorage.getItem(CALL_LOG_KEY) || "[]"); } catch { return []; }
  });
  const [showAdd, setShowAdd]   = useState(false);
  const [newEntry, setNewEntry] = useState({ custNum:"", date: new Date().toISOString().slice(0,10), note:"" });
  const [expandedWeeks, setExpandedWeeks] = useState({});   // weekKey → bool
  const [copied, setCopied]     = useState(null);
  const [allCopied, setAllCopied] = useState(false);

  const custList = actionPlan.map(a => ({ custNum: String(a.custNum), name: a.customer, city: a.city }));

  function saveEntries(updated) {
    const sorted = [...updated].sort((a,b) => new Date(b.date)-new Date(a.date));
    setEntries(sorted);
    try { localStorage.setItem(CALL_LOG_KEY, JSON.stringify(sorted)); } catch {}
  }

  function addEntry() {
    if (!newEntry.note.trim() || !newEntry.custNum) return;
    const cust = custList.find(c => c.custNum === newEntry.custNum);
    const entry = {
      id:       `log_${Date.now()}`,
      custNum:  newEntry.custNum,
      custName: cust?.name || newEntry.custNum,
      city:     cust?.city || "",
      date:     newEntry.date,
      note:     newEntry.note.trim(),
      rep:      currentUser?.name || repName,
    };
    const updated = [entry, ...entries];
    saveEntries(updated);
    setNewEntry({ custNum:"", date: new Date().toISOString().slice(0,10), note:"" });
    setShowAdd(false);
    // Auto-expand the week this was added to
    setExpandedWeeks(prev => ({ ...prev, [getWeekKey(entry.date)]: true }));
    if (onLogActivity) onLogActivity("call_note", `${entry.custName} — ${entry.note.slice(0,60)}${entry.note.length>60?"…":""}`);
  }

  function deleteEntry(id) { saveEntries(entries.filter(e => e.id !== id)); }

  function getWeekKey(dateStr) {
    const d = new Date(dateStr); const mon = new Date(d);
    mon.setDate(d.getDate()-((d.getDay()+6)%7));
    return mon.toISOString().slice(0,10);
  }
  function getWeekLabel(dateStr) {
    const d = new Date(dateStr); const mon = new Date(d);
    mon.setDate(d.getDate()-((d.getDay()+6)%7));
    const fri = new Date(mon); fri.setDate(mon.getDate()+4);
    const fmt = dt => dt.toLocaleDateString("en-US",{month:"short",day:"numeric"});
    // Get ISO week number
    const jan1 = new Date(mon.getFullYear(),0,1);
    const wkNum = Math.ceil(((mon-jan1)/86400000+jan1.getDay()+1)/7);
    return `W${wkNum} · ${fmt(mon)} – ${fmt(fri)}`;
  }

  function copyWeek(weekEntries) {
    const text = weekEntries.map(e =>
      `${e.custName}${e.city?" ("+e.city+")":""} — ${e.date}\n${e.note}`
    ).join("\n\n");
    navigator.clipboard?.writeText(text).catch(()=>{});
    setCopied("copied"); setTimeout(()=>setCopied(null),2000);
  }

  function copyAll() {
    // Group and copy all entries formatted by week
    const text = sortedWeekKeys.map(wk => {
      const grp = weekGroups[wk];
      return `=== ${grp.label} ===\n\n` + grp.entries.map(e =>
        `${e.custName}${e.city?" ("+e.city+")":""} — ${e.date}\n${e.note}`
      ).join("\n\n");
    }).join("\n\n---\n\n");
    navigator.clipboard?.writeText(text).catch(()=>{});
    setAllCopied(true); setTimeout(()=>setAllCopied(false),2000);
  }

  // Build week groups — latest week first
  const weekGroups = {};
  entries.forEach(e => {
    const wk = getWeekKey(e.date);
    if (!weekGroups[wk]) weekGroups[wk] = { label: getWeekLabel(e.date), entries: [] };
    weekGroups[wk].entries.push(e);
  });
  const sortedWeekKeys = Object.keys(weekGroups).sort((a,b) => b.localeCompare(a));

  // Auto-expand most recent week on first load
  const latestWk = sortedWeekKeys[0];
  useState(() => {
    if (latestWk) setExpandedWeeks({ [latestWk]: true });
  });

  function toggleWeek(wk) {
    setExpandedWeeks(prev => ({ ...prev, [wk]: !prev[wk] }));
  }

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
        marginBottom:"0.85rem", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontSize:"0.82rem", fontWeight:700, color:TEXT }}>📞 Call Log History</div>
          <div style={{ fontSize:"0.68rem", color:MUTED }}>
            {entries.length} notes · {sortedWeekKeys.length} week{sortedWeekKeys.length!==1?"s":""} · latest first
          </div>
        </div>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {entries.length > 0 && (
            <button onClick={copyAll}
              style={{ fontSize:"0.68rem", fontWeight:700,
                color: allCopied?"#fff":"#0891B2",
                background: allCopied?"#0891B2":"#EFF6FF",
                border:`1px solid ${allCopied?"#0891B2":"#BFDBFE"}`,
                borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer", whiteSpace:"nowrap" }}>
              {allCopied ? "✓ Copied!" : "📋 Copy All for CRM"}
            </button>
          )}
          <button onClick={()=>setShowAdd(!showAdd)}
            style={{ ...S.btn("#0891B2"), fontSize:"0.7rem", fontWeight:700 }}>
            {showAdd ? "× Cancel" : "+ Add Note"}
          </button>
        </div>
      </div>

      {/* Add note form */}
      {showAdd && (
        <div style={{ ...S.card, border:"2px solid #0891B2", marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:"#0891B2", marginBottom:"0.75rem" }}>📝 Add Call Note</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0.6rem", marginBottom:"0.6rem" }}>
            <div>
              <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Customer *</div>
              <select value={newEntry.custNum}
                onChange={e=>setNewEntry(p=>({...p,custNum:e.target.value}))}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                  padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none" }}>
                <option value="">— Select Customer —</option>
                {custList.map(c=>(
                  <option key={c.custNum} value={c.custNum}>{c.name}{c.city?` (${c.city})`:""}</option>
                ))}
              </select>
            </div>
            <div>
              <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Date *</div>
              <input type="date" value={newEntry.date}
                onChange={e=>setNewEntry(p=>({...p,date:e.target.value}))}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                  padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none", boxSizing:"border-box" }} />
            </div>
          </div>
          <div style={{ marginBottom:"0.75rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Notes *</div>
            <textarea value={newEntry.note}
              onChange={e=>setNewEntry(p=>({...p,note:e.target.value}))}
              placeholder="What was discussed, outcomes, follow-up needed..."
              rows={4}
              style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.75rem", resize:"vertical",
                outline:"none", boxSizing:"border-box", lineHeight:1.65 }}
              onFocus={e=>e.target.style.borderColor="#0891B2"}
              onBlur={e=>e.target.style.borderColor=BORDER}
            />
          </div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <button onClick={()=>{setShowAdd(false);setNewEntry({custNum:"",date:new Date().toISOString().slice(0,10),note:""}); }}
              style={S.btn(MUTED)}>Cancel</button>
            <button onClick={addEntry} disabled={!newEntry.note.trim()||!newEntry.custNum}
              style={{ ...S.btn("#0891B2"), background:"#0891B2", color:"#fff",
                opacity:!newEntry.note.trim()||!newEntry.custNum?0.45:1 }}>
              Save Note
            </button>
          </div>
        </div>
      )}

      {/* Empty state */}
      {entries.length === 0 && !showAdd && (
        <div style={{ ...S.card, textAlign:"center", padding:"2rem 1rem", color:MUTED }}>
          <div style={{ fontSize:"2rem", marginBottom:"0.5rem" }}>📞</div>
          <div style={{ fontSize:"0.8rem", fontWeight:600, marginBottom:4 }}>No call notes yet</div>
          <div style={{ fontSize:"0.7rem" }}>Click + Add Note to log your first visit or call</div>
        </div>
      )}

      {/* Week-grouped history — latest week first */}
      {sortedWeekKeys.map(wk => {
        const grp      = weekGroups[wk];
        const isOpen   = !!expandedWeeks[wk];
        const isLatest = wk === latestWk;
        return (
          <div key={wk} style={{ marginBottom:"0.65rem" }}>
            {/* Week header — clickable accordion */}
            <div onClick={() => toggleWeek(wk)}
              style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
                padding:"0.55rem 0.85rem",
                background: isLatest ? "#EFF6FF" : "#F8FAFC",
                border: `1px solid ${isLatest?"#BFDBFE":BORDER}`,
                borderRadius: isOpen ? "8px 8px 0 0" : 8,
                cursor:"pointer", userSelect:"none" }}>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontSize:"0.82rem" }}>{isOpen?"▾":"▸"}</span>
                <span style={{ fontSize:"0.78rem", fontWeight:700,
                  color: isLatest?"#1E5FCC":TEXT }}>{grp.label}</span>
                {isLatest && <span style={{ fontSize:"0.6rem", fontWeight:700, color:"#fff",
                  background:"#1E5FCC", borderRadius:6, padding:"1px 7px" }}>Latest</span>}
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontSize:"0.68rem", color:MUTED }}>
                  {grp.entries.length} note{grp.entries.length!==1?"s":""}
                </span>
                <button onClick={e=>{e.stopPropagation();copyWeek(grp.entries);}}
                  style={{ fontSize:"0.62rem", color:"#0891B2", background:"#EFF6FF",
                    border:"1px solid #BFDBFE", borderRadius:5, padding:"2px 8px", cursor:"pointer" }}>
                  {copied==="copied"?"✓ Copied":"📋 CRM"}
                </button>
              </div>
            </div>

            {/* Week entries — shown when expanded */}
            {isOpen && (
              <div style={{ border:`1px solid ${isLatest?"#BFDBFE":BORDER}`, borderTop:"none",
                borderRadius:"0 0 8px 8px", overflow:"hidden" }}>
                {[...grp.entries].sort((a,b)=>new Date(b.date)-new Date(a.date)).map((e,i) => (
                  <div key={e.id} style={{ padding:"0.7rem 0.85rem",
                    background: i%2===0?"#FFFFFF":"#F8FAFC",
                    borderTop: i>0?`1px solid ${BORDER}`:undefined }}>
                    <div style={{ display:"flex", justifyContent:"space-between",
                      alignItems:"flex-start", marginBottom:"0.3rem" }}>
                      <div>
                        <span style={{ fontSize:"0.78rem", fontWeight:700, color:TEXT }}>
                          {e.custName}
                        </span>
                        {e.city && (
                          <span style={{ fontSize:"0.68rem", color:MUTED, marginLeft:6 }}>
                            {e.city}
                          </span>
                        )}
                      </div>
                      <div style={{ display:"flex", alignItems:"center", gap:6, flexShrink:0 }}>
                        <span style={{ fontSize:"0.65rem", color:MUTED }}>
                          {new Date(e.date).toLocaleDateString("en-US",{month:"short",day:"numeric"})}
                        </span>
                        <button onClick={()=>deleteEntry(e.id)}
                          style={{ fontSize:"0.6rem", color:RED, background:"none",
                            border:"none", cursor:"pointer", padding:"0 2px", lineHeight:1 }}>×</button>
                      </div>
                    </div>
                    <div style={{ fontSize:"0.75rem", color:TEXT, lineHeight:1.65,
                      whiteSpace:"pre-wrap" }}>{e.note}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function EntryCard({ entry: e, onDelete, formatForCRM }) {
  const [copied, setCopied] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const isLong = e.note.length > 180;

  return (
    <div style={{ ...S.card, padding:"0.75rem 1rem", borderLeft:"3px solid #0891B2" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"0.4rem", flexWrap:"wrap", gap:6 }}>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:"0.78rem", fontWeight:700, color:AMBER }}>↗ {e.custName}</span>
          {e.city && <span style={{ fontSize:"0.65rem", color:MUTED }}>📍 {e.city}</span>}
          <span style={{ fontSize:"0.65rem", color:MUTED, background:"#F4F7FB", borderRadius:6, padding:"1px 7px" }}>
            {new Date(e.date).toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"})}
          </span>
          <span style={{ fontSize:"0.65rem", color:"#0891B2", fontWeight:600 }}>{e.rep}</span>
        </div>
        <div style={{ display:"flex", gap:4 }}>
          <button onClick={()=>{ navigator.clipboard.writeText(formatForCRM([e])).then(()=>{ setCopied(true); setTimeout(()=>setCopied(false),2000); }); }}
            style={{ fontSize:"0.63rem", fontWeight:copied?700:400,
              color:copied?"#fff":"#0891B2", background:copied?"#0891B2":"transparent",
              border:`1px solid ${copied?"#0891B2":"#BFDBFE"}`, borderRadius:4, padding:"1px 7px", cursor:"pointer", transition:"all 0.2s" }}>
            {copied?"✓":"📋"}
          </button>
          <button onClick={()=>{ if(window.confirm("Delete this note?")) onDelete(e.id); }}
            style={{ fontSize:"0.7rem", color:MUTED, background:"none", border:"none", cursor:"pointer", padding:"0 2px" }}>×</button>
        </div>
      </div>
      <div style={{ fontSize:"0.77rem", color:TEXT, lineHeight:1.75,
        overflow: expanded||!isLong ? "visible" : "hidden",
        display: expanded||!isLong ? "block" : "-webkit-box",
        WebkitLineClamp: expanded ? "none" : 4,
        WebkitBoxOrient: "vertical",
        whiteSpace: "pre-wrap", wordBreak:"break-word" }}>
        {e.note}
      </div>
      {isLong && (
        <button onClick={()=>setExpanded(!expanded)}
          style={{ fontSize:"0.65rem", color:"#0891B2", background:"none", border:"none", cursor:"pointer", padding:"3px 0 0", display:"block" }}>
          {expanded ? "▲ Show less" : "▼ Show more"}
        </button>
      )}
    </div>
  );
}

// ── Leads Tab ─────────────────────────────────────────────────────────────────
function LeadsTab({ leads, repName, onAddLead, onDeleteLead, currentUser, isAdmin, allReps, convertLead, onRefreshLeads }) {
  const [showForm, setShowForm]   = useState(false);
  const [filter, setFilter]       = useState("open");
  const [form, setForm]           = useState({
    name:"", city:"", phone:"", businessType:"", notes:"",
    assigned_to: isAdmin ? "" : currentUser?.id || "",
    assigned_to_name: isAdmin ? "" : repName,
  });

  const openLeads      = leads.filter(l => l.status==="open");
  const convertedLeads = leads.filter(l => l.status==="converted");
  const displayed      = filter==="open" ? openLeads : filter==="converted" ? convertedLeads : leads;

  function resetForm() {
    setForm({ name:"", city:"", phone:"", businessType:"", notes:"",
      assigned_to: isAdmin?"":currentUser?.id||"",
      assigned_to_name: isAdmin?"":repName });
    setShowForm(false);
  }

  async function handleSubmit() {
    if (!form.name.trim()) return;
    const assigned_to_name = isAdmin ? form.assigned_to_name : repName;
    const assigned_to      = isAdmin ? form.assigned_to      : currentUser?.id || "";
    await onAddLead({ ...form, assigned_to, assigned_to_name });
    resetForm();
  }

  const priorityColors = { High:RED, Medium:AMBER, Normal:GREEN };

  return (
    <div>
      {/* Sync button for reps */}
      {!isAdmin && onRefreshLeads && (
        <div style={{ marginBottom:"0.6rem", display:"flex", justifyContent:"flex-end", alignItems:"center", gap:8 }}>
          <span style={{ fontSize:"0.65rem", color:MUTED }}>Don't see a lead?</span>
          <button onClick={async () => {
            const uid = window.__pulseUser?.id;
            // Quick diagnostic — test Supabase connection first
            const test = await sbFetch("rep_leads?limit=1");
            if (test === null) {
              alert("Cannot reach cloud database. Fix needed: In Supabase go to Authentication > URL Configuration and add https://casteel1983.github.io to Site URL and Redirect URLs.");
              return;
            }
            // Connection works — do the real sync
            onRefreshLeads(true);
          }}
            style={{ fontSize:"0.68rem", fontWeight:700, color:"#059669", background:"#F0FDF4",
              border:"1px solid #BBF7D0", borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer" }}>
            ↺ Sync from cloud
          </button>
        </div>
      )}

      {/* KPI strip */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"Open Leads",     val:openLeads.length,      color:openLeads.length>0?RED:MUTED },
          { label:"Converted",      val:convertedLeads.length, color:GREEN },
          { label:"Total",          val:leads.length,          color:MUTED },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.06em", marginTop:2 }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Actions row */}
      <div style={{ display:"flex", gap:6, marginBottom:"0.75rem", alignItems:"center" }}>
        {[["open","Open"],["converted","Converted"],["all","All"]].map(([v,l])=>(
          <button key={v} onClick={()=>setFilter(v)}
            style={{ fontSize:"0.7rem", fontWeight:filter===v?700:400,
              color:filter===v?"#fff":MUTED,
              background:filter===v?"#059669":"#F4F7FB",
              border:`1px solid ${filter===v?"#059669":BORDER}`,
              borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer" }}>{l}</button>
        ))}
        <button onClick={()=>setShowForm(!showForm)}
          style={{ ...S.btn("#059669"), fontSize:"0.72rem", marginLeft:"auto", fontWeight:700 }}>
          {showForm ? "× Cancel" : "+ New Lead"}
        </button>
      </div>

      {/* New Lead form */}
      {showForm && (
        <div style={{ ...S.card, border:`2px solid #059669`, marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:"#059669", marginBottom:"0.75rem" }}>🎯 New Lead</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0.6rem", marginBottom:"0.6rem" }}>
            {[
              { label:"Business Name *", key:"name",         placeholder:"e.g. Tifton Tire Co." },
              { label:"City",            key:"city",          placeholder:"e.g. Tifton" },
              { label:"Phone",           key:"phone",         placeholder:"229-555-0000" },
              { label:"Business Type",   key:"businessType",  placeholder:"e.g. Tire Shop, Auto Dealer..." },
            ].map(f=>(
              <div key={f.key}>
                <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>{f.label}</div>
                <input value={form[f.key]} onChange={e=>setForm(p=>({...p,[f.key]:e.target.value}))}
                  placeholder={f.placeholder}
                  style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                    padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none", boxSizing:"border-box" }}
                  onFocus={e=>e.target.style.borderColor="#059669"}
                  onBlur={e=>e.target.style.borderColor=BORDER}
                />
              </div>
            ))}
          </div>
          {isAdmin && (
            <div style={{ marginBottom:"0.6rem" }}>
              <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Assign to Rep *</div>
              <select value={form.assigned_to_name}
                onChange={e=>setForm(p=>({...p, assigned_to_name:e.target.value, assigned_to:e.target.value.toLowerCase()}))}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                  padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none" }}>
                <option value="">— Select Rep —</option>
                {(allReps||[]).map(r=><option key={r} value={r}>{r}</option>)}
              </select>
            </div>
          )}
          <div style={{ marginBottom:"0.75rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Notes</div>
            <textarea value={form.notes} onChange={e=>setForm(p=>({...p,notes:e.target.value}))}
              placeholder="What do you know about this prospect? How did you find them?"
              rows={3}
              style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem", resize:"vertical", outline:"none", boxSizing:"border-box" }}
              onFocus={e=>e.target.style.borderColor="#059669"}
              onBlur={e=>e.target.style.borderColor=BORDER}
            />
          </div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <button onClick={resetForm} style={{ ...S.btn(MUTED) }}>Cancel</button>
            <button onClick={handleSubmit} disabled={!form.name.trim()||(isAdmin&&!form.assigned_to_name)}
              style={{ ...S.btn("#059669"), background:"#059669", color:"#fff",
                opacity:!form.name.trim()||(isAdmin&&!form.assigned_to_name)?0.5:1 }}>
              Create Lead
            </button>
          </div>
        </div>
      )}

      {/* Lead list */}
      {displayed.length === 0 && (
        <div style={{ ...S.card, textAlign:"center", padding:"2rem", color:MUTED }}>
          <div style={{ fontSize:"1.5rem", marginBottom:8 }}>🎯</div>
          <div style={{ fontWeight:600, marginBottom:4 }}>
            {filter==="open" ? "No open leads" : filter==="converted" ? "No converted leads yet" : "No leads yet"}
          </div>
          <div style={{ fontSize:"0.72rem" }}>Click + New Lead to add a prospect</div>
        </div>
      )}

      <div style={{ display:"flex", flexDirection:"column", gap:"0.65rem" }}>
        {displayed.map(lead => (
          <div key={lead.id} style={{ ...S.card, padding:"0.85rem 1rem",
            borderLeft:`4px solid ${lead.status==="converted"?GREEN:RED}`,
            opacity: lead.status==="converted"?0.75:1 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
              <div style={{ flex:1 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap", marginBottom:4 }}>
                  <span style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>{lead.name}</span>
                  <span style={{ fontSize:"0.62rem", fontWeight:800, color:"#fff",
                    background: lead.status==="converted"?GREEN:RED,
                    borderRadius:8, padding:"1px 8px", letterSpacing:"0.05em" }}>
                    {lead.status==="converted" ? "✓ CONVERTED" : "LEAD"}
                  </span>
                  {lead.businessType && <span style={{ fontSize:"0.65rem", color:MUTED, background:"#F4F7FB", borderRadius:8, padding:"1px 7px" }}>{lead.businessType}</span>}
                </div>
                <div style={{ fontSize:"0.68rem", color:MUTED, display:"flex", gap:12, flexWrap:"wrap" }}>
                  {lead.city && <span>📍 {lead.city}</span>}
                  {lead.phone && <span>📞 {lead.phone}</span>}
                  <span>👤 {lead.assigned_to_name || lead.created_by_name}</span>
                  <span>📅 {new Date(lead.created_at).toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>
                  {isAdmin && lead.created_by_name && <span style={{ color:"#7C3AED" }}>Created by {lead.created_by_name}</span>}
                </div>
                {lead.notes && (
                  <div style={{ marginTop:6, fontSize:"0.73rem", color:TEXT, lineHeight:1.6,
                    padding:"0.4rem 0.6rem", background:"#F8FAFF", borderRadius:4, borderLeft:`2px solid ${BORDER}` }}>
                    {lead.notes}
                  </div>
                )}
              </div>
              <div style={{ display:"flex", gap:5, flexShrink:0, alignItems:"flex-start" }}>
                {lead.status==="open" && convertLead && (
                  <button onClick={()=>convertLead(lead.id)}
                    style={{ fontSize:"0.68rem", fontWeight:700, color:"#fff", background:GREEN,
                      border:"none", borderRadius:6, padding:"0.3rem 0.65rem", cursor:"pointer" }}
                    title="Mark as converted customer">
                    ✓ Convert
                  </button>
                )}
                <button onClick={()=>{ if(window.confirm(`Delete lead: ${lead.name}?`)) onDeleteLead(lead.id); }}
                  style={{ fontSize:"0.68rem", color:MUTED, background:"none",
                    border:`1px solid ${BORDER}`, borderRadius:6, padding:"0.3rem 0.5rem", cursor:"pointer" }}>
                  ×
                </button>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}


// ── AI Action Item (must be a proper component — hooks can't live in .map()) ──
function AIActionItem({ action: a, repName, actionPlan, color }) {
  const [added, setAdded] = useState(false);
  const pColor = a.priority==="HIGH" ? RED : a.priority==="MEDIUM" ? AMBER : GREEN;
  const catIcon = c => c==="Visit"?"🚗":c==="AD Program"?"🏆":c==="At Risk"?"⚠️":"📞";

  return (
    <div style={{ display:"flex", gap:10, padding:"0.55rem 0.75rem",
      background: added?"#F0FDF4":"#FFFFFF",
      border:`1px solid ${added?"#BBF7D0":pColor+"44"}`,
      borderLeft:`4px solid ${added?GREEN:pColor}`,
      borderRadius:6, alignItems:"flex-start", transition:"all 0.2s" }}>
      <div style={{ flex:1 }}>
        <div style={{ display:"flex", gap:6, alignItems:"center", marginBottom:3, flexWrap:"wrap" }}>
          <span style={{ fontSize:"0.68rem" }}>{catIcon(a.category)}</span>
          <span style={{ fontSize:"0.65rem", fontWeight:700, color:pColor, background:pColor+"18", borderRadius:8, padding:"1px 6px" }}>{a.priority}</span>
          <span style={{ fontSize:"0.7rem", fontWeight:700, color:AMBER }}>↗ {a.custName}</span>
          <span style={{ fontSize:"0.65rem", color:MUTED, background:"#F4F7FB", borderRadius:8, padding:"1px 6px" }}>{a.category}</span>
        </div>
        <div style={{ fontSize:"0.75rem", color:TEXT, lineHeight:1.6 }}>{a.action}</div>
      </div>
      <button
        onClick={() => {
          if (added || !a.custNum) return;
          const key = `todos_${a.custNum}`;
          const existing = JSON.parse(localStorage.getItem(key)||"[]");
          const newTodo = {
            id: Date.now(),
            text: `[AI] ${a.action}`,
            done: false,
            date: new Date().toISOString().slice(0,10),
            by: window.__pulseUser?.name || repName,
          };
          const updated = [newTodo, ...existing];
          localStorage.setItem(key, JSON.stringify(updated));
          const userId = window.__pulseUser?.id;
          if (userId) {
            const apRow = actionPlan.find(ap => String(ap.custNum)===String(a.custNum));
            syncTodosUp(userId, a.custNum, a.custName, apRow?.city||"", apRow?.salesman||repName, updated);
          }
          setAdded(true);
        }}
        disabled={added || !a.custNum}
        style={{ fontSize:"0.65rem", fontWeight:700, flexShrink:0, marginTop:2,
          color: added?"#059669":"#FFFFFF",
          background: added?"transparent":color,
          border:`1px solid ${added?GREEN:color}`,
          borderRadius:6, padding:"0.25rem 0.65rem",
          cursor: added||!a.custNum?"default":"pointer",
          transition:"all 0.2s" }}>
        {added ? "✓ Added" : "+ Add"}
      </button>
    </div>
  );
}

// ── Rep To Do Tab ─────────────────────────────────────────────────────────────
function RepTodoTab({ repName, actionPlan, onCustomerClick, color, leads, currentUser }) {
  const aiPlanKey = `ai_action_plan_${repName}`;
  const [aiPlan, setAiPlanState] = useState(() => {
    try { return localStorage.getItem(aiPlanKey) || null; } catch { return null; }
  });
  const [aiLoading, setAiLoading] = useState(false);
  const [aiPlanDate, setAiPlanDate] = useState(() => {
    try { return localStorage.getItem(aiPlanKey + "_date") || null; } catch { return null; }
  });

  function setAiPlan(text) {
    setAiPlanState(text);
    try {
      if (text) {
        localStorage.setItem(aiPlanKey, text);
        const now = new Date().toLocaleString("en-US",{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"});
        localStorage.setItem(aiPlanKey + "_date", now);
        setAiPlanDate(now);
        // Sync to Supabase
        const userId = window.__pulseUser?.id;
        if (userId) syncAIPlanUp(userId, text, now);
      } else {
        localStorage.removeItem(aiPlanKey);
        localStorage.removeItem(aiPlanKey + "_date");
        setAiPlanDate(null);
      }
    } catch {}
  }

  // Load AI plan from Supabase on mount
  useEffect(() => {
    const userId = window.__pulseUser?.id;
    if (!userId) return;
    syncAIPlanDown(userId).then(row => {
      if (row?.plan && !aiPlan) {
        setAiPlanState(row.plan);
        setAiPlanDate(row.generated_at);
        try {
          localStorage.setItem(aiPlanKey, row.plan);
          localStorage.setItem(aiPlanKey + "_date", row.generated_at || "");
        } catch {}
      }
    });
  }, []);

  async function generateAIPlan() {
    setAiLoading(true);
    setAiPlan(null);
    try {
      // Build rep context
      const repAP = actionPlan.filter(a => a.salesman === repName || a.salesman?.toLowerCase() === repName.toLowerCase());
      const sorted = [...repAP].sort((a,b) => b.change - a.change);
      const top5   = sorted.slice(0,5).map(a => `${a.customer} (${a.city}): $${a.sales2026.toFixed(0)} 2026, ${a.change>=0?"+":""}$${a.change.toFixed(0)} vs PY, top dept: ${a.topDept}`).join("\n");
      const bot5   = sorted.slice(-5).map(a => `${a.customer} (${a.city}): $${a.sales2026.toFixed(0)} 2026, ${a.change>=0?"+":""}$${a.change.toFixed(0)} vs PY`).join("\n");
      const declined = repAP.filter(a=>a.change<0).length;

      // Gather open todos
      const openTodos = [];
      repAP.forEach(ap => {
        try {
          const todos = JSON.parse(localStorage.getItem(`todos_${ap.custNum}`) || "[]");
          todos.filter(t=>!t.done).forEach(t => openTodos.push(`${ap.customer}: ${t.text}`));
        } catch {}
      });

      // Gather AD program enrollments for this rep
      const adEnrolled = [];
      const allPrograms = [
        { label:"Toyo AD",    data:AD_PROGRAMS,        getVal:d=>d.pcr?.total||0 },
        { label:"Ascenso",    data:ASCENSO_PROGRAMS,   getVal:d=>d.amount||0 },
        { label:"Americus",   data:AMERICUS_PROGRAMS,  getVal:d=>d.ytd||0 },
        { label:"BARNN",      data:BARNN_PROGRAMS,     getVal:d=>d.total||0 },
        { label:"Falken PLT", data:FALKEN_PLT_PROGRAMS,getVal:d=>d.q2||0 },
        { label:"Falken TBR", data:FALKEN_TBR_PROGRAMS,getVal:d=>d.ytd||0 },
        { label:"Yokohama",   data:YOKOHAMA_PROGRAMS,  getVal:d=>d.ytd||0 },
      ];
      repAP.forEach(ap => {
        allPrograms.forEach(prog => {
          const d = prog.data[String(ap.custNum)];
          if (d) adEnrolled.push(`${ap.customer} — ${prog.label}: ${prog.getVal(d)} units`);
        });
      });

      // Build customer lookup for AI (name → custNum)
      const custLookup = repAP.slice(0,30).map(a => `${a.customer}|${a.custNum}`).join(",");

      const prompt = `You are a sales manager at a tire and ag supply distributor. Generate a specific, actionable weekly action plan for sales rep ${repName}.

REP OVERVIEW:
- Total accounts: ${repAP.length}
- Accounts with positive growth: ${repAP.filter(a=>a.change>0).length}
- Accounts in decline: ${declined}
- Total 2026 sales: $${repAP.reduce((s,a)=>s+a.sales2026,0).toFixed(0)}

TOP GROWING ACCOUNTS:
${top5}

MOST DECLINED ACCOUNTS:
${bot5}

AD PROGRAM ENROLLMENTS:
${adEnrolled.length > 0 ? adEnrolled.join("\n") : "No program enrollments found"}

OPEN TO-DOS (${openTodos.length} total):
${openTodos.slice(0,10).join("\n") || "None"}

CUSTOMER LOOKUP (Name|CustNum):
${custLookup}

Return ONLY valid JSON (no markdown, no backticks) in this exact format:
{
  "summary": "2-3 sentence overall assessment of ${repName}'s territory",
  "coaching": "One specific coaching insight based on their performance pattern",
  "actions": [
    {
      "custNum": "the customer number from the lookup above",
      "custName": "exact customer name",
      "priority": "HIGH|MEDIUM|LOW",
      "category": "Visit|AD Program|At Risk|Follow Up",
      "action": "Specific actionable instruction for this customer (1-2 sentences)"
    }
  ]
}

Include 5-8 actions. Use real customer names and numbers from the lookup. Be specific — reference actual products, programs, dollar amounts, or trends.`;

      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000,
          messages:[{ role:"user", content:prompt }]
        })
      });
      if (!res.ok) throw new Error(`API ${res.status}`);
      const data = await res.json();
      const raw = data.content?.[0]?.text || "{}";
      try {
        const clean = raw.replace(/```json|```/g,"").trim();
        const parsed = JSON.parse(clean);
        setAiPlan(JSON.stringify(parsed)); // store as JSON string
      } catch {
        // Fallback: store raw text as legacy format
        setAiPlan(JSON.stringify({ summary: raw, actions: [], coaching: "" }));
      }
    } catch(e) {
      setAiPlan(`Error: ${e.message}`);
    }
    setAiLoading(false);
  }

  // Load all todos + notes from localStorage (fast) then refresh from Supabase
  const [data, setData] = useState(() => {
    const result = [];
    try {
      actionPlan.forEach(ap => {
        const todosRaw = localStorage.getItem(`todos_${ap.custNum}`);
        const notesRaw = localStorage.getItem(`notes_${ap.custNum}`);
        const todos    = todosRaw ? JSON.parse(todosRaw) : [];
        const notes    = notesRaw || "";
        if (todos.length > 0 || notes) result.push({ ap, todos, notes });
      });
    } catch {}
    return result;
  });
  const [sbLoaded, setSbLoaded] = useState(false);

  // Sync from Supabase on mount
  useEffect(() => {
    const userId = window.__pulseUser?.id;
    if (!userId) return;
    syncAllTodosDown(userId).then(rows => {
      if (!rows) return;
      // Group by custNum
      const byCust = {};
      rows.forEach(r => {
        if (!byCust[r.cust_num]) byCust[r.cust_num] = [];
        byCust[r.cust_num].push({ id:r.id, text:r.text, done:r.done, date:r.created_date, by:r.created_by });
      });
      // Update localStorage and state
      Object.entries(byCust).forEach(([cnum, todos]) => {
        try { localStorage.setItem(`todos_${cnum}`, JSON.stringify(todos)); } catch {}
      });
      // Rebuild data
      const result = [];
      actionPlan.forEach(ap => {
        const todos    = byCust[String(ap.custNum)] || JSON.parse(localStorage.getItem(`todos_${ap.custNum}`) || "[]");
        const notes    = localStorage.getItem(`notes_${ap.custNum}`) || "";
        if (todos.length > 0 || notes) result.push({ ap, todos, notes });
      });
      setData(result);
      setSbLoaded(true);
    });
  }, []);

  const [filter, setFilter] = useState("open"); // open | all | notes
  const [refreshKey, setRefreshKey] = useState(0);

  function refresh() {
    const result = [];
    try {
      actionPlan.forEach(ap => {
        const todosRaw = localStorage.getItem(`todos_${ap.custNum}`);
        const notesRaw = localStorage.getItem(`notes_${ap.custNum}`);
        const todos    = todosRaw ? JSON.parse(todosRaw) : [];
        const notes    = notesRaw || "";
        if (todos.length > 0 || notes) {
          result.push({ ap, todos, notes });
        }
      });
    } catch {}
    setData(result);
    setRefreshKey(k => k+1);
  }

  function toggleTodo(custNum, id) {
    const key   = `todos_${custNum}`;
    const todos = JSON.parse(localStorage.getItem(key) || "[]");
    const updated = todos.map(t => t.id===id ? {...t, done:!t.done} : t);
    localStorage.setItem(key, JSON.stringify(updated));
    refresh();
  }

  // What to show
  const filtered = data.filter(d => {
    if (filter === "open")  return d.todos.some(t => !t.done);
    if (filter === "notes") return !!d.notes;
    return d.todos.length > 0 || d.notes;
  }).sort((a,b) => {
    // Sort by open todos first
    const aOpen = a.todos.filter(t=>!t.done).length;
    const bOpen = b.todos.filter(t=>!t.done).length;
    return bOpen - aOpen;
  });

  const totalOpen = data.reduce((s,d) => s + d.todos.filter(t=>!t.done).length, 0);
  const totalDone = data.reduce((s,d) => s + d.todos.filter(t=> t.done).length, 0);

  return (
    <div>
      {/* ── AI Action Plan ── */}
      <div style={{ ...S.card, borderLeft:`4px solid ${color}`, marginBottom:"1rem", background:"linear-gradient(135deg,#F8FAFF,#EEF4FF)" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom: aiPlan||aiLoading?"0.75rem":"0" }}>
          <div>
            <div style={{ fontSize:"0.78rem", fontWeight:700, color:color }}>◈ AI Action Plan — {repName}</div>
            {aiPlanDate && <div style={{ fontSize:"0.65rem", color:MUTED, marginTop:1 }}>Last generated: {aiPlanDate}</div>}
            {!aiPlan && !aiLoading && <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:2 }}>Personalized plan based on your accounts, trends & AD programs</div>}
          </div>
          <button onClick={generateAIPlan} disabled={aiLoading}
            style={{ ...S.btn(color), fontSize:"0.7rem", opacity:aiLoading?0.7:1 }}>
            {aiLoading ? "◈ Analyzing…" : aiPlan ? "↺ Regenerate" : "◈ Generate Plan"}
          </button>
        </div>
        {aiLoading && (
          <div style={{ fontSize:"0.75rem", color:MUTED, textAlign:"center", padding:"0.75rem" }}>
            ◈ Analyzing {repName}'s accounts, sales trends & AD programs…
          </div>
        )}
        {aiPlan && !aiLoading && (() => {
          let parsed = null;
          try { parsed = JSON.parse(aiPlan); } catch {}
          if (!parsed) return <div style={{ fontSize:"0.78rem", color:TEXT, lineHeight:1.85, whiteSpace:"pre-wrap", paddingTop:"0.75rem" }}>{aiPlan}</div>;

          const priorityColor = p => p==="HIGH"?RED:p==="MEDIUM"?AMBER:GREEN;
          const catIcon = c => c==="Visit"?"🚗":c==="AD Program"?"🏆":c==="At Risk"?"⚠️":"📞";

          return (
            <div style={{ borderTop:`1px solid ${BORDER}`, paddingTop:"0.75rem" }}>
              {/* Summary */}
              {parsed.summary && (
                <div style={{ fontSize:"0.78rem", color:TEXT, lineHeight:1.75, marginBottom:"1rem", padding:"0.6rem 0.75rem", background:"#F4F7FB", borderRadius:6 }}>
                  {parsed.summary}
                </div>
              )}

              {/* Action items */}
              {parsed.actions?.length > 0 && (
                <div style={{ marginBottom:"0.85rem" }}>
                  <div style={{ fontSize:"0.68rem", fontWeight:700, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.5rem" }}>
                    Action Items — click to add to customer to-do
                  </div>
                  <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                    {parsed.actions.map((a,i) => (
                      <AIActionItem key={i} action={a} repName={repName} actionPlan={actionPlan} color={color} />
                    ))}
                  </div>
                </div>
              )}

              {/* Coaching insight */}
              {parsed.coaching && (
                <div style={{ padding:"0.6rem 0.75rem", background:"#EDE9FE", border:"1px solid #C4B5FD", borderRadius:6, fontSize:"0.75rem", color:"#5B21B6", lineHeight:1.7 }}>
                  💡 <strong>Coaching:</strong> {parsed.coaching}
                </div>
              )}
            </div>
          );
        })()}
      </div>

      {/* KPI row */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"Open To-Dos",     val:totalOpen,        color:totalOpen>0?RED:GREEN },
          { label:"Open Leads",      val:(leads||[]).filter(l=>l.status==="open"&&(l.assigned_to===currentUser?.id||l.assigned_to_name===repName)).length, color:GREEN },
          { label:"Customers w/Notes", val:data.filter(d=>d.notes).length, color:AMBER },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Filter + Refresh */}
      <div style={{ display:"flex", gap:6, marginBottom:"0.75rem", alignItems:"center" }}>
        {[["open","📋 Open"],["notes","📝 Notes"],["all","All"]].map(([v,l])=>(
          <button key={v} onClick={()=>setFilter(v)}
            style={{ fontSize:"0.7rem", fontWeight:filter===v?700:400, color:filter===v?"#fff":MUTED,
              background:filter===v?color:"#F4F7FB", border:`1px solid ${filter===v?color:BORDER}`,
              borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer" }}>{l}</button>
        ))}
        <button onClick={refresh} style={{ marginLeft:"auto", fontSize:"0.68rem", color:MUTED, background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"0.25rem 0.6rem", cursor:"pointer" }}>↺ Refresh</button>
      </div>

      {/* Empty state */}
      {filtered.length === 0 && (
        <div style={{ ...S.card, textAlign:"center", padding:"2rem", color:MUTED }}>
          <div style={{ fontSize:"1.5rem", marginBottom:8 }}>📋</div>
          <div style={{ fontWeight:600, marginBottom:4 }}>
            {filter==="open" ? "No open to-dos" : filter==="notes" ? "No notes yet" : "No to-dos or notes yet"}
          </div>
          <div style={{ fontSize:"0.72rem" }}>Open a customer tab → 📞 Calls to add to-dos and notes</div>
        </div>
      )}

      {/* Customer cards */}
      <div style={{ display:"flex", flexDirection:"column", gap:"0.75rem" }}>
        {filtered.map(({ ap, todos, notes }) => {
          const openTodos = todos.filter(t=>!t.done);
          const doneTodos = todos.filter(t=> t.done);
          return (
            <div key={ap.custNum} style={{ ...S.card, padding:"0.85rem 1rem" }}>
              {/* Header */}
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"0.65rem" }}>
                <div>
                  <div onClick={()=>onCustomerClick&&onCustomerClick(ap)}
                    style={{ fontSize:"0.85rem", fontWeight:700, color:AMBER, cursor:"pointer", display:"inline-flex", alignItems:"center", gap:5 }}
                    onMouseEnter={e=>e.target.style.textDecoration="underline"}
                    onMouseLeave={e=>e.target.style.textDecoration="none"}>
                    ↗ {ap.customer}
                  </div>
                  <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:1 }}>📍 {ap.city} · {ap.salesman}</div>
                </div>
                <div style={{ display:"flex", gap:5, alignItems:"center" }}>
                  {openTodos.length > 0 && <span style={{ background:RED, color:"#fff", borderRadius:8, padding:"1px 8px", fontSize:"0.65rem", fontWeight:700 }}>{openTodos.length} open</span>}
                  {notes && <span style={{ background:"#FEF3C7", color:"#92400E", borderRadius:8, padding:"1px 8px", fontSize:"0.65rem", fontWeight:600 }}>📝</span>}
                </div>
              </div>

              {/* Open to-dos */}
              {openTodos.length > 0 && (
                <div style={{ display:"flex", flexDirection:"column", gap:4, marginBottom: notes||doneTodos.length>0?"0.6rem":0 }}>
                  {openTodos.map(t => (
                    <TodoItem key={t.id} todo={t}
                      onToggle={()=>toggleTodo(ap.custNum, t.id)}
                      onDelete={()=>{
                        const key = `todos_${ap.custNum}`;
                        const updated = JSON.parse(localStorage.getItem(key)||"[]").filter(x=>x.id!==t.id);
                        localStorage.setItem(key, JSON.stringify(updated));
                        if (window.__pulseUser) syncTodosUp(window.__pulseUser.id, ap.custNum, ap.customer, ap.city, ap.salesman, updated);
                        refresh();
                      }}
                      onEdit={(newText)=>{
                        const key = `todos_${ap.custNum}`;
                        const updated = JSON.parse(localStorage.getItem(key)||"[]").map(x=>x.id===t.id?{...x,text:newText}:x);
                        localStorage.setItem(key, JSON.stringify(updated));
                        if (window.__pulseUser) syncTodosUp(window.__pulseUser.id, ap.custNum, ap.customer, ap.city, ap.salesman, updated);
                        refresh();
                      }}
                    />
                  ))}
                </div>
              )}

              {/* Completed to-dos (collapsed) */}
              {doneTodos.length > 0 && filter==="all" && (
                <div style={{ display:"flex", flexDirection:"column", gap:3, marginBottom: notes?"0.6rem":0, opacity:0.6 }}>
                  {doneTodos.map(t => (
                    <div key={t.id} style={{ display:"flex", alignItems:"flex-start", gap:8, padding:"0.3rem 0.6rem", background:"#F4F7FB", borderRadius:6 }}>
                      <input type="checkbox" checked={true} onChange={()=>toggleTodo(ap.custNum,t.id)} style={{ marginTop:3, cursor:"pointer", accentColor:GREEN, flexShrink:0 }} />
                      <div style={{ fontSize:"0.72rem", color:MUTED, textDecoration:"line-through", flex:1 }}>{t.text}</div>
                    </div>
                  ))}
                </div>
              )}

              {/* Notes snippet */}
              {notes && (filter==="notes"||filter==="all") && (
                <div style={{ padding:"0.5rem 0.65rem", background:"#F4F7FB", borderRadius:6, borderLeft:`3px solid ${AMBER}` }}>
                  <div style={{ fontSize:"0.65rem", color:MUTED, fontWeight:600, marginBottom:3, textTransform:"uppercase" }}>📝 Notes</div>
                  <div style={{ fontSize:"0.74rem", color:TEXT, lineHeight:1.6, whiteSpace:"pre-wrap",
                    overflow:"hidden", textOverflow:"ellipsis", display:"-webkit-box", WebkitLineClamp:3, WebkitBoxOrient:"vertical" }}>
                    {notes}
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Statesboro Tab ────────────────────────────────────────────────────────────
function StatesboroTab() {
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("total");
  const [deptFilter, setDeptFilter] = useState("All");
  const [expandedCust, setExpandedCust] = useState(null);

  const SOS_NUM = "3000104";
  const AG_OTR_IND = ["FARM TIRES","INDUSTRIAL TIRES","OFF THE ROAD TIRES"];

  const allAccounts = Object.entries(STATESBORO_DATA).map(([cnum, d]) => ({
    cnum, ...d,
    isSOS: cnum === SOS_NUM,
    totalGP: Object.values(d.depts).reduce((s,v)=>s+(v.gp||0),0),
    topDept: Object.entries(d.depts).sort((a,b)=>b[1].amount-a[1].amount)[0]?.[0] || "",
  }));

  // Dept filter options
  const allDepts = [...new Set(allAccounts.flatMap(a => Object.keys(a.depts)))].sort();
  const deptOptions = ["All", ...AG_OTR_IND, ...allDepts.filter(d => !AG_OTR_IND.includes(d))];

  const filtered = allAccounts
    .filter(a => !search || a.name.toLowerCase().includes(search.toLowerCase()))
    .filter(a => deptFilter === "All" || a.depts[deptFilter])
    .sort((a,b) => {
      if (a.isSOS && !b.isSOS) return -1;
      if (!a.isSOS && b.isSOS) return 1;
      if (sortBy === "total") return b.total - a.total;
      if (sortBy === "gp") return b.totalGP - a.totalGP;
      return a.name.localeCompare(b.name);
    });

  const grandTotal = allAccounts.reduce((s,a)=>s+a.total,0);
  const grandGP    = allAccounts.reduce((s,a)=>s+a.totalGP,0);
  const sosAcct    = allAccounts.find(a=>a.isSOS);

  return (
    <div>
      {/* KPI row */}
      <div style={S.kpiRow}>
        <div style={S.kpi("#0891B2")}>
          <div style={S.kpiVal}>{fmt(grandTotal)}</div>
          <div style={S.kpiLbl}>Statesboro YTD 2026</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>{fmt(grandGP)}</div>
          <div style={S.kpiLbl}>Total GP$</div>
        </div>
        <div style={S.kpi(AMBER)}>
          <div style={S.kpiVal}>{pct(grandTotal > 0 ? grandGP/grandTotal : 0)}</div>
          <div style={S.kpiLbl}>Blended GP%</div>
        </div>
        <div style={S.kpi(MUTED)}>
          <div style={S.kpiVal}>{allAccounts.length}</div>
          <div style={S.kpiLbl}>Active Accounts</div>
        </div>
      </div>

      {/* SOS callout */}
      {sosAcct && (
        <div style={{ ...S.card, borderLeft:"4px solid #D97706", marginBottom:"0.85rem", background:"#FFFBEB" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
            <div>
              <div style={{ fontSize:"0.8rem", fontWeight:700, color:"#D97706" }}>⭐ SOS TIRE & AUTO — Strategic Account</div>
              <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:2 }}>Customer #{SOS_NUM} · All departments tracked</div>
            </div>
            <div style={{ fontSize:"1.05rem", fontWeight:800, color:"#D97706" }}>{fmt(sosAcct.total)}</div>
          </div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:"0.75rem" }}>
            {Object.entries(sosAcct.depts).sort((a,b)=>b[1].amount-a[1].amount).map(([dept,v])=>(
              <div key={dept} style={{ background:"#FEF3C7", border:"1px solid #FDE68A", borderRadius:6, padding:"0.3rem 0.65rem", fontSize:"0.68rem" }}>
                <span style={{ fontWeight:700, color:"#92400E" }}>{dept}</span>
                <span style={{ color:MUTED, marginLeft:4 }}>{fmt(v.amount)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Filters */}
      <div style={{ display:"flex", gap:8, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search account..."
          style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.7rem", borderRadius:4, fontSize:"0.7rem", width:180 }} />
        <div style={{ position:"relative" }}>
          <select value={deptFilter} onChange={e=>setDeptFilter(e.target.value)}
            style={{ appearance:"none", background:"#FFFFFF", border:`2px solid ${deptFilter!=="All"?"#0891B2":BORDER}`, borderRadius:6,
              color:deptFilter!=="All"?"#0891B2":TEXT, fontWeight:deptFilter!=="All"?700:400,
              padding:"0.35rem 1.75rem 0.35rem 0.65rem", fontSize:"0.7rem", cursor:"pointer", outline:"none" }}>
            {deptOptions.map(d=><option key={d} value={d}>{d}</option>)}
          </select>
          <span style={{ position:"absolute", right:6, top:"50%", transform:"translateY(-50%)", pointerEvents:"none", color:MUTED, fontSize:"0.65rem" }}>▼</span>
        </div>
        <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
          style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.5rem", borderRadius:4, fontSize:"0.7rem" }}>
          <option value="total">Sort: Revenue</option>
          <option value="gp">Sort: GP$</option>
          <option value="name">Sort: Name</option>
        </select>
        <span style={{ color:MUTED, fontSize:"0.65rem", marginLeft:"auto" }}>{filtered.length} accounts · W1–W21</span>
      </div>

      {/* Account list */}
      <div style={S.card}>
        <div style={{ overflowX:"auto" }}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Account</th>
                <th style={{ ...S.th, textAlign:"right" }}>YTD Revenue</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP$</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP%</th>
                <th style={S.th}>Top Dept</th>
                <th style={S.th}>Segments</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((a,i) => (
                <>
                  <tr key={a.cnum}
                    onClick={()=>setExpandedCust(expandedCust===a.cnum ? null : a.cnum)}
                    style={{ cursor:"pointer", background: expandedCust===a.cnum ? "#EEF4FF" : "transparent" }}
                    onMouseEnter={e=>{ if(expandedCust!==a.cnum) e.currentTarget.style.background="#F4F7FB"; }}
                    onMouseLeave={e=>{ if(expandedCust!==a.cnum) e.currentTarget.style.background="transparent"; }}>
                    <td style={{ ...S.td, fontWeight:600, color: a.isSOS ? "#D97706" : AMBER }}>
                      {a.isSOS ? "⭐ " : "↗ "}{a.name}
                      <span style={{ fontSize:"0.62rem", color:MUTED, fontWeight:400, marginLeft:4 }}>#{a.cnum}</span>
                    </td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700 }}>{fmt(a.total)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:GREEN }}>{fmt(a.totalGP)}</td>
                    <td style={{ ...S.td, textAlign:"right", color: a.total>0 && a.totalGP/a.total < 0.08 ? RED : GREEN }}>
                      {a.total > 0 ? pct(a.totalGP/a.total) : "—"}
                    </td>
                    <td style={{ ...S.td, fontSize:"0.68rem" }}>{a.topDept}</td>
                    <td style={{ ...S.td }}>
                      <div style={{ display:"flex", flexWrap:"wrap", gap:3 }}>
                        {Object.keys(a.depts).map(d=>(
                          <span key={d} style={{ fontSize:"0.58rem", background: AG_OTR_IND.includes(d)?"#D1FAE5":"#F4F7FB",
                            color: AG_OTR_IND.includes(d)?GREEN:MUTED, borderRadius:6, padding:"1px 5px", border:`1px solid ${AG_OTR_IND.includes(d)?"#BBF7D0":BORDER}` }}>
                            {d}
                          </span>
                        ))}
                      </div>
                    </td>
                  </tr>
                  {expandedCust === a.cnum && (
                    <tr key={`${a.cnum}_exp`}>
                      <td colSpan={6} style={{ ...S.td, background:"#F8FAFF", padding:"0.75rem 1rem" }}>
                        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, marginBottom:"0.5rem", textTransform:"uppercase", letterSpacing:"0.08em" }}>Department Breakdown</div>
                        <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
                          {Object.entries(a.depts).sort((x,y)=>y[1].amount-x[1].amount).map(([dept,v])=>(
                            <div key={dept} style={{ background:"#FFFFFF", border:`1px solid ${AG_OTR_IND.includes(dept)?"#BBF7D0":BORDER}`, borderRadius:6, padding:"0.4rem 0.75rem", minWidth:140 }}>
                              <div style={{ fontSize:"0.68rem", fontWeight:700, color: AG_OTR_IND.includes(dept)?GREEN:TEXT }}>{dept}</div>
                              <div style={{ fontSize:"0.75rem", fontWeight:700, color:AMBER, marginTop:2 }}>{fmt(v.amount)}</div>
                              <div style={{ fontSize:"0.65rem", color:MUTED }}>{v.qty} units · GP {pct(v.amount>0?v.gp/v.amount:0)}</div>
                            </div>
                          ))}
                        </div>
                      </td>
                    </tr>
                  )}
                </>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── Specials Tab ──────────────────────────────────────────────────────────────
function SpecialsTab({ repName, color }) {
  const storageKey = `specials_${repName}`;
  const [specials, setSpecials] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [imgPreviews, setImgPreviews] = useState({});
  const fileInputRef = useRef(null);

  const DEPTS = ["FARM TIRES","INDUSTRIAL TIRES","OFF THE ROAD TIRES","RAD LT TRUCK",
    "TRUCK TIRES","PASSENGER TIRES","ST TRAILER","TUBES","VALVE STEMS",
    "WHEELS","WHEEL WEIGHTS","TIRE TOOLS","LAWN & GARDEN","ALL DEPARTMENTS"];

  const emptyForm = { title:"", desc:"", discount:"", startDate:"", endDate:"", depts:[], notes:"", imgData:"" };
  const [form, setForm] = useState(emptyForm);

  // Load from storage on mount
  useEffect(() => {
    async function load() {
      try {
        const res = (()=>{ const v=localStorage.getItem(storageKey); return v?{value:v}:null; })();
        if (res?.value) setSpecials(JSON.parse(res.value));
      } catch { /* no data yet */ }
      setLoaded(true);
    }
    load();
  }, [repName]);

  async function save(list) {
    setSpecials(list);
    try { localStorage.setItem(storageKey, JSON.stringify(list)); } catch {}
  }

  function handleImgUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => setForm(f => ({ ...f, imgData: ev.target.result }));
    reader.readAsDataURL(file);
  }

  function handlePdfUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => setForm(f => ({ ...f, imgData: ev.target.result, pdfName: file.name }));
    reader.readAsDataURL(file);
  }

  function toggleDept(dept) {
    setForm(f => ({
      ...f,
      depts: f.depts.includes(dept) ? f.depts.filter(d=>d!==dept) : [...f.depts, dept]
    }));
  }

  function openAdd() { setForm(emptyForm); setEditId(null); setShowForm(true); }
  function openEdit(sp) { setForm({...sp}); setEditId(sp.id); setShowForm(true); }
  function cancelForm() { setShowForm(false); setForm(emptyForm); setEditId(null); }

  function submitForm() {
    if (!form.title.trim()) { alert("Please enter a title for this special."); return; }
    let next;
    if (editId) {
      next = specials.map(s => s.id === editId ? { ...form, id: editId } : s);
    } else {
      next = [...specials, { ...form, id: Date.now().toString() }];
    }
    save(next);
    cancelForm();
  }

  function deleteSpecial(id) {
    if (!confirm("Remove this special?")) return;
    save(specials.filter(s => s.id !== id));
  }

  const isActive = (sp) => {
    const today = new Date().toISOString().slice(0,10);
    if (!sp.startDate && !sp.endDate) return true;
    if (sp.endDate && sp.endDate < today) return false;
    if (sp.startDate && sp.startDate > today) return false;
    return true;
  };

  const activeSpecials = specials.filter(isActive);
  const expiredSpecials = specials.filter(s => !isActive(s));

  if (!loaded) return <div style={{ color:MUTED, padding:"2rem", textAlign:"center" }}>Loading specials…</div>;

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"1rem", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontSize:"0.78rem", fontWeight:700, color }}>
            {repName.toUpperCase()} — ACTIVE SPECIALS
            <span style={{ marginLeft:8, background:activeSpecials.length>0?"#D1FAE5":"#F3F4F6", color:activeSpecials.length>0?GREEN:MUTED, borderRadius:10, padding:"1px 8px", fontSize:"0.68rem", fontWeight:700 }}>
              {activeSpecials.length} active
            </span>
          </div>
          <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:3 }}>Upload flyers or fill in the details for any specials running right now</div>
        </div>
        <button
          onClick={openAdd}
          style={{ ...S.btn(color), display:"flex", alignItems:"center", gap:6 }}>
          + Add Special
        </button>
      </div>

      {/* Add / Edit Form */}
      {showForm && (
        <div style={{ ...S.card, border:`2px solid ${color}`, marginBottom:"1.25rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color, marginBottom:"1rem" }}>
            {editId ? "✏ Edit Special" : "➕ New Special"}
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0.75rem", marginBottom:"0.75rem" }}>
            <div style={{ gridColumn:"1/-1" }}>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>TITLE *</label>
              <input value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))}
                placeholder="e.g. Farm Tire Blowout Sale"
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.8rem", boxSizing:"border-box" }} />
            </div>
            <div>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>DISCOUNT / PRICE</label>
              <input value={form.discount} onChange={e=>setForm(f=>({...f,discount:e.target.value}))}
                placeholder="e.g. 15% off, Buy 3 get 1, $50 rebate"
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.8rem", boxSizing:"border-box" }} />
            </div>
            <div>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>DESCRIPTION</label>
              <input value={form.desc} onChange={e=>setForm(f=>({...f,desc:e.target.value}))}
                placeholder="Brief details"
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.8rem", boxSizing:"border-box" }} />
            </div>
            <div>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>START DATE</label>
              <input type="date" value={form.startDate} onChange={e=>setForm(f=>({...f,startDate:e.target.value}))}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.8rem", boxSizing:"border-box" }} />
            </div>
            <div>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>END DATE</label>
              <input type="date" value={form.endDate} onChange={e=>setForm(f=>({...f,endDate:e.target.value}))}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.8rem", boxSizing:"border-box" }} />
            </div>
            <div style={{ gridColumn:"1/-1" }}>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:6 }}>APPLICABLE DEPARTMENTS</label>
              <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                {DEPTS.map(d => (
                  <button key={d} onClick={()=>toggleDept(d)}
                    style={{ padding:"0.25rem 0.65rem", borderRadius:20, border:`2px solid ${form.depts.includes(d)?color:BORDER}`,
                      background:form.depts.includes(d)?"#EEF4FF":"#FFFFFF", color:form.depts.includes(d)?color:MUTED,
                      fontWeight:form.depts.includes(d)?700:400, fontSize:"0.68rem", cursor:"pointer" }}>
                    {d}
                  </button>
                ))}
              </div>
            </div>
            <div style={{ gridColumn:"1/-1" }}>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:4 }}>NOTES / TERMS</label>
              <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}
                placeholder="Any fine print, restrictions, or call-to-action notes"
                rows={2}
                style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.78rem", resize:"vertical", boxSizing:"border-box" }} />
            </div>
            {/* Image / flyer upload */}
            <div style={{ gridColumn:"1/-1" }}>
              <label style={{ fontSize:"0.68rem", color:MUTED, display:"block", marginBottom:6 }}>UPLOAD FLYER / IMAGE <span style={{ color:MUTED, fontWeight:400 }}>(optional — JPG, PNG)</span></label>
              <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                <label style={{ ...S.btn(MUTED), cursor:"pointer", fontSize:"0.72rem" }}>
                  📎 Choose Image
                  <input type="file" accept="image/*" onChange={handleImgUpload} style={{ display:"none" }} />
                </label>
                {form.imgData && form.imgData.startsWith("data:image") && (
                  <div style={{ position:"relative" }}>
                    <img src={form.imgData} alt="preview" style={{ height:60, borderRadius:6, border:`1px solid ${BORDER}`, objectFit:"cover" }} />
                    <button onClick={()=>setForm(f=>({...f,imgData:""}))}
                      style={{ position:"absolute", top:-6, right:-6, background:RED, color:"#fff", border:"none", borderRadius:"50%", width:16, height:16, fontSize:10, cursor:"pointer", lineHeight:1 }}>×</button>
                  </div>
                )}
                {!form.imgData && <span style={{ fontSize:"0.68rem", color:MUTED }}>No image uploaded</span>}
              </div>
            </div>
          </div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end", paddingTop:"0.75rem", borderTop:`1px solid ${BORDER}` }}>
            <button onClick={cancelForm} style={{ ...S.btn(MUTED) }}>Cancel</button>
            <button onClick={submitForm} style={{ ...S.btn(color) }}>{editId ? "Save Changes" : "Add Special"}</button>
          </div>
        </div>
      )}

      {/* Active specials */}
      {activeSpecials.length === 0 && !showForm && (
        <div style={{ ...S.card, textAlign:"center", padding:"2.5rem 1rem", color:MUTED }}>
          <div style={{ fontSize:"2rem", marginBottom:"0.5rem" }}>🏷</div>
          <div style={{ fontSize:"0.8rem", fontWeight:600, marginBottom:4 }}>No active specials right now</div>
          <div style={{ fontSize:"0.72rem" }}>Click "+ Add Special" to post a deal or promotion</div>
        </div>
      )}

      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))", gap:"0.85rem", marginBottom: expiredSpecials.length > 0 ? "1.5rem" : 0 }}>
        {activeSpecials.map(sp => (
          <div key={sp.id} style={{ ...S.card, borderTop:`4px solid ${color}`, marginBottom:0, position:"relative" }}>
            {/* Action buttons */}
            <div style={{ position:"absolute", top:10, right:10, display:"flex", gap:4 }}>
              <button onClick={()=>openEdit(sp)} style={{ background:"#EEF4FF", border:`1px solid ${BORDER}`, color:AMBER, borderRadius:4, padding:"2px 8px", fontSize:"0.65rem", cursor:"pointer" }}>✏ Edit</button>
              <button onClick={()=>deleteSpecial(sp.id)} style={{ background:"#FEE2E2", border:`1px solid #FECDD3`, color:RED, borderRadius:4, padding:"2px 8px", fontSize:"0.65rem", cursor:"pointer" }}>✕</button>
            </div>
            {/* Flyer image */}
            {sp.imgData && sp.imgData.startsWith("data:image") && (
              <img src={sp.imgData} alt={sp.title} style={{ width:"100%", maxHeight:160, objectFit:"cover", borderRadius:6, marginBottom:"0.75rem", border:`1px solid ${BORDER}` }} />
            )}
            {/* Title + discount */}
            <div style={{ fontSize:"0.9rem", fontWeight:700, color:TEXT, marginBottom:4, paddingRight:80 }}>{sp.title}</div>
            {sp.discount && (
              <div style={{ display:"inline-block", background: color==="#D97706"?"#FEF3C7":"#EEF4FF", color, fontWeight:800, fontSize:"0.85rem", padding:"0.2rem 0.75rem", borderRadius:6, marginBottom:"0.5rem" }}>
                {sp.discount}
              </div>
            )}
            {sp.desc && <div style={{ fontSize:"0.76rem", color:TEXT, marginBottom:"0.5rem", lineHeight:1.6 }}>{sp.desc}</div>}
            {/* Date range */}
            {(sp.startDate || sp.endDate) && (
              <div style={{ fontSize:"0.68rem", color:MUTED, marginBottom:"0.5rem" }}>
                📅 {sp.startDate && `From ${sp.startDate}`}{sp.startDate && sp.endDate && " → "}{sp.endDate && `Ends ${sp.endDate}`}
              </div>
            )}
            {!sp.startDate && !sp.endDate && (
              <div style={{ fontSize:"0.68rem", color:GREEN, marginBottom:"0.5rem", fontWeight:600 }}>● No expiration</div>
            )}
            {/* Departments */}
            {sp.depts?.length > 0 && (
              <div style={{ display:"flex", flexWrap:"wrap", gap:4, marginBottom:"0.5rem" }}>
                {sp.depts.map(d=>(
                  <span key={d} style={{ fontSize:"0.62rem", background:"#F4F7FB", border:`1px solid ${BORDER}`, color:MUTED, borderRadius:8, padding:"1px 7px" }}>{d}</span>
                ))}
              </div>
            )}
            {/* Notes */}
            {sp.notes && (
              <div style={{ fontSize:"0.68rem", color:MUTED, borderTop:`1px solid ${BORDER}`, paddingTop:"0.4rem", marginTop:"0.4rem", fontStyle:"italic" }}>
                {sp.notes}
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Expired / upcoming */}
      {expiredSpecials.length > 0 && (
        <div>
          <div style={{ fontSize:"0.68rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.5rem" }}>Expired / Upcoming</div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))", gap:"0.6rem" }}>
            {expiredSpecials.map(sp=>(
              <div key={sp.id} style={{ ...S.card, marginBottom:0, opacity:0.55, borderTop:`3px solid ${BORDER}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                  <div style={{ fontSize:"0.78rem", fontWeight:600, color:MUTED }}>{sp.title}</div>
                  <button onClick={()=>deleteSpecial(sp.id)} style={{ background:"none", border:"none", color:MUTED, cursor:"pointer", fontSize:"0.75rem" }}>✕</button>
                </div>
                {sp.endDate && <div style={{ fontSize:"0.65rem", color:RED, marginTop:2 }}>Ended {sp.endDate}</div>}
                {sp.startDate && new Date(sp.startDate).toISOString().slice(0,10) > new Date().toISOString().slice(0,10) && (
                  <div style={{ fontSize:"0.65rem", color:AMBER, marginTop:2 }}>Starts {sp.startDate}</div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── AR Tab ────────────────────────────────────────────────────────────────────
function ARTab({ ar }) {
  const [repFilter, setRepFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("balance");
  const [subTab, setSubTab] = useState("aging");

  const reps = ["All", ...Array.from(new Set(ar.map(r => r.salesman).filter(Boolean))).sort()];

  const filtered = ar
    .filter(r => repFilter === "All" || r.salesman === repFilter)
    .filter(r => !search || r.name.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      if (sortBy === "balance") return b.balance - a.balance;
      if (sortBy === "over90") return b.dueOver90 - a.dueOver90;
      if (sortBy === "name") return a.name.localeCompare(b.name);
      return 0;
    });

  const totals = filtered.reduce((acc, r) => ({
    balance: acc.balance + r.balance,
    futDue: acc.futDue + r.futDue,
    curDue: acc.curDue + r.curDue,
    due1_30: acc.due1_30 + r.due1_30,
    due31_60: acc.due31_60 + r.due31_60,
    due61_90: acc.due61_90 + r.due61_90,
    dueOver90: acc.dueOver90 + r.dueOver90,
  }), { balance:0, futDue:0, curDue:0, due1_30:0, due31_60:0, due61_90:0, dueOver90:0 });

  const allTotals = ar.reduce((acc, r) => ({
    balance: acc.balance + r.balance,
    curDue: acc.curDue + r.curDue,
    due1_30: acc.due1_30 + r.due1_30,
    due31_60: acc.due31_60 + r.due31_60,
    due61_90: acc.due61_90 + r.due61_90,
    dueOver90: acc.dueOver90 + r.dueOver90,
  }), { balance:0, curDue:0, due1_30:0, due31_60:0, due61_90:0, dueOver90:0 });

  const riskAccounts = ar.filter(r => r.dueOver90 > 0).sort((a,b) => b.dueOver90 - a.dueOver90);
  const pastDueTotal = allTotals.due31_60 + allTotals.due61_90 + allTotals.dueOver90;
  const pastDuePct = allTotals.balance ? pastDueTotal / allTotals.balance : 0;

  // Aging breakdown for chart
  const agingBuckets = [
    { label: "Current", val: allTotals.curDue, color: GREEN },
    { label: "1-30 days", val: allTotals.due1_30, color: AMBER },
    { label: "31-60 days", val: allTotals.due31_60, color: "#F6AD55" },
    { label: "61-90 days", val: allTotals.due61_90, color: RED },
    { label: "90+ days", val: allTotals.dueOver90, color: "#FC4040" },
  ];

  // By rep summary
  const byRep = {};
  ar.forEach(r => {
    const rep = r.salesman || "Unknown";
    if (!byRep[rep]) byRep[rep] = { balance:0, curDue:0, due1_30:0, due31_60:0, due61_90:0, dueOver90:0, count:0 };
    byRep[rep].balance += r.balance;
    byRep[rep].curDue += r.curDue;
    byRep[rep].due1_30 += r.due1_30;
    byRep[rep].due31_60 += r.due31_60;
    byRep[rep].due61_90 += r.due61_90;
    byRep[rep].dueOver90 += r.dueOver90;
    byRep[rep].count++;
  });

  function agingBar(r) {
    const total = r.balance || 1;
    const segs = [
      { w: r.curDue/total, c: GREEN },
      { w: r.due1_30/total, c: AMBER },
      { w: r.due31_60/total, c: "#F6AD55" },
      { w: r.due61_90/total, c: RED },
      { w: r.dueOver90/total, c: "#FC4040" },
    ];
    return (
      <div style={{ display:"flex", height:6, borderRadius:3, overflow:"hidden", width:120 }}>
        {segs.map((s,i) => s.w > 0 && <div key={i} style={{ width:`${s.w*100}%`, background:s.c }} />)}
      </div>
    );
  }

  return (
    <div>
      {/* KPI row */}
      <div style={S.kpiRow}>
        <div style={S.kpi(AMBER)}>
          <div style={S.kpiVal}>{fmt(allTotals.balance)}</div>
          <div style={S.kpiLbl}>Total AR Balance</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>{fmt(allTotals.curDue)}</div>
          <div style={S.kpiLbl}>Current ({pct(allTotals.balance ? allTotals.curDue/allTotals.balance : 0)})</div>
        </div>
        <div style={S.kpi(RED)}>
          <div style={S.kpiVal}>{fmt(pastDueTotal)}</div>
          <div style={S.kpiLbl}>Past Due 31+ days ({pct(pastDuePct)})</div>
        </div>
        <div style={S.kpi("#FC4040")}>
          <div style={{ ...S.kpiVal, color: allTotals.dueOver90 > 0 ? RED : TEXT }}>{fmt(allTotals.dueOver90)}</div>
          <div style={S.kpiLbl}>Over 90 Days ({riskAccounts.length} accts)</div>
        </div>
      </div>

      {/* Sub tabs */}
      <div style={S.subNav}>
        <button style={S.subBtn(subTab==="aging")} onClick={()=>setSubTab("aging")}>📊 Aging Summary</button>
        <button style={S.subBtn(subTab==="accounts", AMBER)} onClick={()=>setSubTab("accounts")}>📋 All Accounts</button>
        <button style={S.subBtn(subTab==="risk", RED)} onClick={()=>setSubTab("risk")}>⚠ 90+ Day Risk</button>
        <button style={S.subBtn(subTab==="byrep", TEAL)} onClick={()=>setSubTab("byrep")}>👤 By Rep</button>
      </div>

      {/* Aging Summary */}
      {subTab === "aging" && (
        <div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"0.75rem" }}>
            <div style={S.card}>
              <div style={{ fontSize:"0.68rem", color:MUTED, letterSpacing:"0.1em", marginBottom:"1rem" }}>AGING BREAKDOWN</div>
              {agingBuckets.map(b => (
                <div key={b.label} style={{ marginBottom:"0.75rem" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.72rem", marginBottom:4 }}>
                    <span style={{ color:b.color }}>{b.label}</span>
                    <span style={{ fontWeight:600 }}>{fmt(b.val)} <span style={{ color:MUTED, fontSize:"0.65rem" }}>({pct(allTotals.balance ? b.val/allTotals.balance : 0)})</span></span>
                  </div>
                  <div style={{ height:8, background:BORDER, borderRadius:4 }}>
                    <div style={{ height:8, background:b.color, borderRadius:4, width:`${allTotals.balance ? (b.val/allTotals.balance)*100 : 0}%`, opacity:0.85 }} />
                  </div>
                </div>
              ))}
            </div>
            <div style={S.card}>
              <div style={{ fontSize:"0.68rem", color:MUTED, letterSpacing:"0.1em", marginBottom:"1rem" }}>AS OF MAY 4, 2026 · {ar.length} ACCOUNTS</div>
              <table style={S.table}>
                <thead>
                  <tr>
                    <th style={S.th}>Bucket</th>
                    <th style={{ ...S.th, textAlign:"right" }}>Amount</th>
                    <th style={{ ...S.th, textAlign:"right" }}>% of AR</th>
                    <th style={{ ...S.th, textAlign:"right" }}>Accts</th>
                  </tr>
                </thead>
                <tbody>
                  {agingBuckets.map(b => (
                    <tr key={b.label}>
                      <td style={{ ...S.td, color:b.color }}>{b.label}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:600 }}>{fmt(b.val)}</td>
                      <td style={{ ...S.td, textAlign:"right" }}>{pct(allTotals.balance ? b.val/allTotals.balance : 0)}</td>
                      <td style={{ ...S.td, textAlign:"right", color:MUTED }}>
                        {b.label === "Current" ? ar.filter(r=>r.curDue>0).length
                         : b.label === "1-30 days" ? ar.filter(r=>r.due1_30>0).length
                         : b.label === "31-60 days" ? ar.filter(r=>r.due31_60>0).length
                         : b.label === "61-90 days" ? ar.filter(r=>r.due61_90>0).length
                         : ar.filter(r=>r.dueOver90>0).length}
                      </td>
                    </tr>
                  ))}
                  <tr style={{ borderTop:`1px solid ${AMBER}` }}>
                    <td style={{ ...S.td, color:AMBER, fontWeight:700 }}>TOTAL</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700, color:AMBER }}>{fmt(allTotals.balance)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:MUTED }}>100%</td>
                    <td style={{ ...S.td, textAlign:"right", color:MUTED }}>{ar.length}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* All Accounts */}
      {subTab === "accounts" && (
        <div>
          <div style={{ display:"flex", gap:8, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
            <div style={{ display:"flex", gap:4 }}>
              {reps.map(r => <button key={r} style={S.subBtn(repFilter===r, REP_COLORS[r]||AMBER)} onClick={()=>setRepFilter(r)}>{r}</button>)}
            </div>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search customer..."
              style={{ background:BG2, border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.7rem", borderRadius:4, fontSize:"0.7rem", width:190 }} />
            <select value={sortBy} onChange={e=>setSortBy(e.target.value)}
              style={{ background:BG2, border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.5rem", borderRadius:4, fontSize:"0.68rem" }}>
              <option value="balance">Sort: Balance</option>
              <option value="over90">Sort: 90+ Days</option>
              <option value="name">Sort: Name</option>
            </select>
            <span style={{ color:MUTED, fontSize:"0.65rem", marginLeft:"auto" }}>{filtered.length} accounts · {fmt(totals.balance)} total</span>
          </div>
          <div style={S.card}>
            <div style={{ overflowX:"auto", maxHeight:520, overflowY:"auto" }}>
              <table style={S.table}>
                <thead style={{ position:"sticky", top:0, background:BG2 }}>
                  <tr>
                    <th style={S.th}>Customer</th>
                    <th style={S.th}>Rep</th>
                    <th style={{ ...S.th, textAlign:"right" }}>Balance</th>
                    <th style={{ ...S.th, width:130 }}>Aging Mix</th>
                    <th style={{ ...S.th, textAlign:"right" }}>Current</th>
                    <th style={{ ...S.th, textAlign:"right" }}>1-30</th>
                    <th style={{ ...S.th, textAlign:"right" }}>31-60</th>
                    <th style={{ ...S.th, textAlign:"right" }}>61-90</th>
                    <th style={{ ...S.th, textAlign:"right", color:RED }}>90+</th>
                    <th style={S.th}>Last Paid</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={i} style={{ background: r.dueOver90 > 5000 ? "rgba(252,129,129,0.04)" : "transparent" }}>
                      <td style={{ ...S.td, maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={r.name}>{r.name}</td>
                      <td style={{ ...S.td, color:REP_COLORS[r.salesman]||MUTED }}>{r.salesman}</td>
                      <td style={{ ...S.td, textAlign:"right", fontWeight:600 }}>{fmt(r.balance)}</td>
                      <td style={{ ...S.td }}>{agingBar(r)}</td>
                      <td style={{ ...S.td, textAlign:"right", color:r.curDue>0?GREEN:MUTED, fontSize:"0.68rem" }}>{r.curDue>0?fmt(r.curDue):"—"}</td>
                      <td style={{ ...S.td, textAlign:"right", color:r.due1_30>0?AMBER:MUTED, fontSize:"0.68rem" }}>{r.due1_30>0?fmt(r.due1_30):"—"}</td>
                      <td style={{ ...S.td, textAlign:"right", color:r.due31_60>0?"#F6AD55":MUTED, fontSize:"0.68rem" }}>{r.due31_60>0?fmt(r.due31_60):"—"}</td>
                      <td style={{ ...S.td, textAlign:"right", color:r.due61_90>0?RED:MUTED, fontSize:"0.68rem" }}>{r.due61_90>0?fmt(r.due61_90):"—"}</td>
                      <td style={{ ...S.td, textAlign:"right", color:r.dueOver90>0?"#FC4040":MUTED, fontWeight:r.dueOver90>0?700:400, fontSize:"0.68rem" }}>{r.dueOver90>0?fmt(r.dueOver90):"—"}</td>
                      <td style={{ ...S.td, color:MUTED, fontSize:"0.65rem" }}>{r.lastPaid}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop:`1px solid ${AMBER}` }}>
                    <td style={{ ...S.td, color:AMBER, fontWeight:700 }} colSpan={2}>TOTAL ({filtered.length})</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700, color:AMBER }}>{fmt(totals.balance)}</td>
                    <td style={S.td} />
                    <td style={{ ...S.td, textAlign:"right", fontWeight:600, color:GREEN }}>{fmt(totals.curDue)}</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:600, color:AMBER }}>{fmt(totals.due1_30)}</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:600, color:"#F6AD55" }}>{fmt(totals.due31_60)}</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:600, color:RED }}>{fmt(totals.due61_90)}</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700, color:"#FC4040" }}>{fmt(totals.dueOver90)}</td>
                    <td style={S.td} />
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* 90+ Risk */}
      {subTab === "risk" && (
        <div style={S.card}>
          <div style={{ fontSize:"0.68rem", color:RED, letterSpacing:"0.1em", marginBottom:"0.75rem" }}>
            ⚠ {riskAccounts.length} ACCOUNTS WITH 90+ DAY BALANCE · {fmt(allTotals.dueOver90)} TOTAL AT RISK
          </div>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Customer</th>
                <th style={S.th}>Rep</th>
                <th style={{ ...S.th, textAlign:"right" }}>Total Balance</th>
                <th style={{ ...S.th, textAlign:"right", color:RED }}>90+ Amount</th>
                <th style={{ ...S.th, textAlign:"right" }}>% of Balance</th>
                <th style={S.th}>Phone</th>
                <th style={S.th}>Last Paid</th>
              </tr>
            </thead>
            <tbody>
              {riskAccounts.map((r, i) => (
                <tr key={i}>
                  <td style={{ ...S.td, maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={r.name}>{r.name}</td>
                  <td style={{ ...S.td, color:REP_COLORS[r.salesman]||MUTED }}>{r.salesman}</td>
                  <td style={{ ...S.td, textAlign:"right" }}>{fmt(r.balance)}</td>
                  <td style={{ ...S.td, textAlign:"right", color:"#FC4040", fontWeight:700 }}>{fmt(r.dueOver90)}</td>
                  <td style={{ ...S.td, textAlign:"right", color: r.dueOver90/r.balance > 0.5 ? RED : AMBER }}>{pct(r.balance ? r.dueOver90/r.balance : 0)}</td>
                  <td style={{ ...S.td, color:MUTED, fontSize:"0.65rem" }}>{r.phone}</td>
                  <td style={{ ...S.td, color:MUTED, fontSize:"0.65rem" }}>{r.lastPaid}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* By Rep */}
      {subTab === "byrep" && (
        <div style={S.card}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Rep</th>
                <th style={{ ...S.th, textAlign:"right" }}>Accounts</th>
                <th style={{ ...S.th, textAlign:"right" }}>Total Balance</th>
                <th style={{ ...S.th, textAlign:"right" }}>Current</th>
                <th style={{ ...S.th, textAlign:"right" }}>1-30</th>
                <th style={{ ...S.th, textAlign:"right" }}>31-60</th>
                <th style={{ ...S.th, textAlign:"right" }}>61-90</th>
                <th style={{ ...S.th, textAlign:"right", color:RED }}>90+</th>
                <th style={{ ...S.th, textAlign:"right" }}>% Current</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(byRep).sort((a,b)=>b[1].balance-a[1].balance).map(([rep, d]) => (
                <tr key={rep}>
                  <td style={{ ...S.td, color:REP_COLORS[rep]||TEXT, fontWeight:600 }}>{rep||"Unknown"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:MUTED }}>{d.count}</td>
                  <td style={{ ...S.td, textAlign:"right", fontWeight:600 }}>{fmt(d.balance)}</td>
                  <td style={{ ...S.td, textAlign:"right", color:GREEN }}>{fmt(d.curDue)}</td>
                  <td style={{ ...S.td, textAlign:"right", color:AMBER }}>{d.due1_30>0?fmt(d.due1_30):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:"#F6AD55" }}>{d.due31_60>0?fmt(d.due31_60):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:RED }}>{d.due61_90>0?fmt(d.due61_90):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:"#FC4040", fontWeight:d.dueOver90>0?700:400 }}>{d.dueOver90>0?fmt(d.dueOver90):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color: d.balance&&d.curDue/d.balance>0.85?GREEN:AMBER }}>{pct(d.balance?d.curDue/d.balance:0)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── Car Dealer Tab ───────────────────────────────────────────────────────────
function CarDealerTab({ weekComp, customers, onCustomerClick }) {
  const color = "#DC2626";
  const [cityFilter, setCityFilter] = useState("All");
  const [repFilter, setRepFilter] = useState("All");
  const [search, setSearch]   = useState("");
  const [sortBy, setSortBy]   = useState("change");
  const [userLoc, setUserLoc] = useState(null);
  const [locStatus, setLocStatus] = useState("idle");

  // Build full car dealer account list
  const allAccounts = buildCarDealerAccounts(weekComp?.actionPlan, customers);

  // Coord map for distance sort
  const coordMap = {};
  (customers || SEED_CUSTOMERS).filter(c => c.lat && c.lon).forEach(c => { coordMap[c.num] = { lat: c.lat, lon: c.lon }; });

  function haversineCD(lat1, lon1, lat2, lon2) {
    const R = 3958.8, dLat=(lat2-lat1)*Math.PI/180, dLon=(lon2-lon1)*Math.PI/180;
    const a = Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
    return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
  }

  function requestLocation() {
    if (!navigator.geolocation) { setLocStatus("denied"); return; }
    setLocStatus("loading");
    navigator.geolocation.getCurrentPosition(
      pos => { setUserLoc({ lat: pos.coords.latitude, lon: pos.coords.longitude }); setLocStatus("ready"); setSortBy("closest"); },
      () => setLocStatus("denied")
    );
  }

  const cities = ["All", ...Array.from(new Set(allAccounts.map(a => a.city).filter(Boolean))).sort()];
  const matchedReps = ["All", ...Array.from(new Set(allAccounts.map(a => a.salesman))).filter(Boolean).sort()];

  const withDist = allAccounts.map(a => {
    const coords = coordMap[a.custNum];
    const dist = (userLoc && coords) ? haversineCD(userLoc.lat, userLoc.lon, coords.lat, coords.lon) : null;
    return { ...a, _dist: dist };
  });

  const filtered = withDist
    .filter(a => cityFilter === "All" || a.city === cityFilter)
    .filter(a => repFilter === "All" || a.salesman === repFilter)
    .filter(a => !search || a.customer.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      if (sortBy === "closest") {
        if (a._dist===null && b._dist===null) return 0;
        if (a._dist===null) return 1; if (b._dist===null) return -1;
        return a._dist - b._dist;
      }
      if (sortBy === "change") return Math.abs(b.change)-Math.abs(a.change);
      return b.sales2026 - a.sales2026;
    });

  // KPIs
  const total25   = filtered.reduce((s,a)=>s+a.sales2025, 0);
  const total26   = filtered.reduce((s,a)=>s+a.sales2026, 0);
  const growing   = filtered.filter(a=>(a.action||"").toUpperCase().includes("GROW")).length;
  const declining = filtered.filter(a=>(a.action||"").toUpperCase().includes("DECLIN")||(a.action||"").toUpperCase().includes("LOST")).length;
  const isFiltered = cityFilter !== "All" || repFilter !== "All";

  return (
    <div>
      {/* KPI row */}
      <div style={S.kpiRow}>
        <div style={S.kpi(color)}>
          <div style={S.kpiVal}>{fmt(total26)}</div>
          <div style={S.kpiLbl}>{isFiltered ? "Filtered Sales 2026" : "Total Sales 2026"}</div>
        </div>
        <div style={S.kpi(clr(total26-total25))}>
          <div style={{ ...S.kpiVal, color: clr(total26-total25) }}>{fmt(total26-total25)}</div>
          <div style={S.kpiLbl}>vs 2025 ({pct(total25 ? (total26-total25)/total25 : 0)})</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>{growing}</div>
          <div style={S.kpiLbl}>Growing</div>
        </div>
        <div style={S.kpi(RED)}>
          <div style={S.kpiVal}>{declining}</div>
          <div style={S.kpiLbl}>Declining</div>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display:"flex", gap:8, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
        <div style={{ fontSize:"0.75rem", fontWeight:700, color }}>
          CAR DEALER ACCOUNTS
          <span style={{ fontWeight:400, color:MUTED, marginLeft:6 }}>({filtered.length} of {allAccounts.length})</span>
        </div>

        {/* City dropdown */}
        <div style={{ position:"relative" }}>
          <select value={cityFilter} onChange={e=>setCityFilter(e.target.value)}
            style={{ appearance:"none", background:"#FFFFFF", border:`2px solid ${cityFilter!=="All"?color:BORDER}`, borderRadius:6, color:cityFilter!=="All"?color:TEXT, fontWeight:cityFilter!=="All"?700:400, padding:"0.4rem 2rem 0.4rem 0.75rem", fontSize:"0.75rem", cursor:"pointer", outline:"none", minWidth:150 }}>
            {cities.map(c=><option key={c} value={c}>{c==="All"?"📍 All Cities":`📍 ${c}`}</option>)}
          </select>
          <span style={{ position:"absolute", right:8, top:"50%", transform:"translateY(-50%)", pointerEvents:"none", color:cityFilter!=="All"?color:MUTED, fontSize:"0.7rem" }}>▼</span>
        </div>

        {/* Matched rep filter */}
        <div style={{ position:"relative" }}>
          <select value={repFilter} onChange={e=>setRepFilter(e.target.value)}
            style={{ appearance:"none", background:"#FFFFFF", border:`2px solid ${repFilter!=="All"?(REP_COLORS[repFilter]||BORDER):BORDER}`, borderRadius:6, color:repFilter!=="All"?(REP_COLORS[repFilter]||TEXT):TEXT, fontWeight:repFilter!=="All"?700:400, padding:"0.4rem 2rem 0.4rem 0.75rem", fontSize:"0.75rem", cursor:"pointer", outline:"none", minWidth:130 }}>
            {matchedReps.map(r=><option key={r} value={r}>{r==="All"?"👤 All Reps":`👤 ${r}`}</option>)}
          </select>
          <span style={{ position:"absolute", right:8, top:"50%", transform:"translateY(-50%)", pointerEvents:"none", fontSize:"0.7rem" }}>▼</span>
        </div>

        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search dealer…"
          style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.35rem 0.7rem", borderRadius:4, fontSize:"0.7rem", width:160 }} />

        <select value={sortBy}
          onChange={e=>{ if(e.target.value==="closest"&&locStatus==="idle") requestLocation(); else setSortBy(e.target.value); }}
          style={{ background:"#FFFFFF", border:`2px solid ${sortBy==="closest"?color:BORDER}`, color:sortBy==="closest"?color:TEXT, fontWeight:sortBy==="closest"?700:400, padding:"0.35rem 0.5rem", borderRadius:4, fontSize:"0.68rem" }}>
          <option value="change">Sort: $ Change</option>
          <option value="sales">Sort: 2026 Sales</option>
          <option value="closest">📍 Sort: Closest</option>
        </select>

        {locStatus==="loading" && <span style={{ fontSize:"0.68rem", color }}>📡 Getting location…</span>}
        {locStatus==="ready"   && <span style={{ fontSize:"0.68rem", color:GREEN }}>✓ Sorted by distance</span>}
        {locStatus==="denied"  && <span style={{ fontSize:"0.68rem", color:RED }}>Location denied</span>}
        {(cityFilter!=="All"||repFilter!=="All") && (
          <button onClick={()=>{ setCityFilter("All"); setRepFilter("All"); }}
            style={{ background:"none", border:"none", color:MUTED, cursor:"pointer", fontSize:"0.7rem" }}>✕ Clear filters</button>
        )}
      </div>

      {/* Table */}
      <div style={S.card}>
        <div style={{ overflowX:"auto" }}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Dealer</th>
                <th style={S.th}>City</th>
                <th style={S.th}>Matched Rep</th>
                {sortBy==="closest" && <th style={{ ...S.th, textAlign:"right", color }}>Miles</th>}
                <th style={{ ...S.th, textAlign:"right" }}>2025</th>
                <th style={{ ...S.th, textAlign:"right" }}>2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>$ Change</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP%</th>
                <th style={S.th}>Status</th>
                <th style={S.th}>Focus</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((a,i) => (
                <tr key={i}
                  onClick={()=>onCustomerClick&&onCustomerClick(a)}
                  style={{ cursor:onCustomerClick?"pointer":"default" }}
                  onMouseEnter={e=>{ if(onCustomerClick) e.currentTarget.style.background="#FFF5F5"; }}
                  onMouseLeave={e=>{ e.currentTarget.style.background="transparent"; }}>
                  <td style={{ ...S.td, maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", color:color, fontWeight:600 }} title={a.customer}>
                    ↗ {a.customer}
                    <span style={{ marginLeft:5, fontSize:"0.6rem", background:"#FEE2E2", color:"#DC2626", borderRadius:8, padding:"1px 5px", fontWeight:700 }}>CAR DEALER</span>
                  </td>
                  <td style={{ ...S.td, color:MUTED }}>{a.city}</td>
                  <td style={{ ...S.td, fontWeight:600, color:REP_COLORS[a.salesman]||MUTED }}>{a.salesman}</td>
                  {sortBy==="closest" && (
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700, color:a._dist!==null?color:MUTED }}>
                      {a._dist!==null?`${a._dist.toFixed(1)} mi`:"—"}
                    </td>
                  )}
                  <td style={{ ...S.td, textAlign:"right", color:MUTED }}>{a.sales2025>0?fmt(a.sales2025):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right" }}>{a.sales2026>0?fmt(a.sales2026):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:clr(a.change), fontWeight:600 }}>{a.change!==0?fmt(a.change):"—"}</td>
                  <td style={{ ...S.td, textAlign:"right", color:a.gpPct<0.08&&a.gpPct>0?RED:TEXT }}>{a.gpPct>0?pct(a.gpPct):"—"}</td>
                  <td style={S.td}>
                    {(a.action&&a.action!=="CAR DEALER") && (
                      <span style={{ fontSize:"0.68rem", background:
                        a.action.toUpperCase().includes("GROW")?GREEN:
                        a.action.toUpperCase().includes("LOST")||a.action.toUpperCase().includes("DECLIN")?RED:TEAL,
                        color:BG2, padding:"1px 5px", borderRadius:2, fontWeight:700 }}>
                        {a.action.toUpperCase().includes("GROW")?"GROWING":
                         a.action.toUpperCase().includes("LOST")?"LOST":
                         a.action.toUpperCase().includes("DECLIN")?"DECLINING":"MAINTAIN"}
                      </span>
                    )}
                  </td>
                  <td style={{ ...S.td, color:MUTED, fontSize:"0.68rem", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{a.focus!=="On track"?a.focus:""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


// ── Branches Sub-Tab ──────────────────────────────────────────────────────────
function BranchesTab({ branchData }) {
  const [period, setPeriod] = useState("ytd");   // q1 | q2 | ytd
  const [metric, setMetric] = useState("sales"); // sales | gp | gpPct

  const BRANCH_COLORS = { Byron:"#1E5FCC", Tifton:"#7C3AED", Statesboro:"#0891B2", Athens:"#D97706" };
  const branches = ["Byron","Tifton","Statesboro","Athens"];

  function getVals(b) {
    const bd = branchData.branches[b];
    if (period === "q1")  return { cur: bd.q1_2026, prev: bd.q1_2025, gCur: bd.q1_gp26, gPrev: bd.q1_gp25 };
    if (period === "q2")  return { cur: bd.q2_2026, prev: bd.q2_2025, gCur: bd.q2_gp26, gPrev: bd.q2_gp25 };
    return { cur: bd.q1_2026+bd.q2_2026, prev: bd.q1_2025+bd.q2_2025, gCur: bd.q1_gp26+bd.q2_gp26, gPrev: bd.q1_gp25+bd.q2_gp25 };
  }

  const branchRows = branches.map(b => {
    const { cur, prev, gCur, gPrev } = getVals(b);
    const chg = cur - prev;
    const chgPct = prev ? chg/prev : 0;
    const gpPct = cur ? gCur/cur : 0;
    return { b, cur, prev, chg, chgPct, gCur, gPrev, gpPct };
  });

  const totalCur  = branchRows.reduce((s,r)=>s+r.cur,0);
  const totalPrev = branchRows.reduce((s,r)=>s+r.prev,0);
  const totalGP   = branchRows.reduce((s,r)=>s+r.gCur,0);
  const totalChg  = totalCur - totalPrev;

  // Weekly trend data for chart (2026 only)
  const chartData = branchData.weeklySales.map(w => ({
    week: `W${w.week}`,
    Byron: Math.round(w.Byron),
    Tifton: Math.round(w.Tifton),
    Statesboro: Math.round(w.Statesboro),
    Athens: Math.round(w.Athens),
  }));

  // Tifton dept breakdown
  const tiftonDepts = (branchData.tiftonQ1Depts || []).filter(d => d.sales2026 > 500).slice(0,10);
  const maxDeptSale = Math.max(...tiftonDepts.map(d=>d.sales2026));

  const periodLabel = period==="q1" ? "Q1 2026" : period==="q2" ? "Q2 2026 (YTD)" : "YTD 2026 (Q1+Q2)";

  return (
    <div>
      {/* Period + Metric toggles */}
      <div style={{ display:"flex", gap:8, marginBottom:"1rem", flexWrap:"wrap", alignItems:"center" }}>
        <div style={{ display:"flex", gap:4 }}>
          {[["ytd","YTD"],["q1","Q1"],["q2","Q2 YTD"]].map(([k,l])=>(
            <button key={k} style={S.subBtn(period===k,AMBER)} onClick={()=>setPeriod(k)}>{l}</button>
          ))}
        </div>
        <div style={{ width:1, background:BORDER, margin:"0 4px", height:20 }} />
        <div style={{ display:"flex", gap:4 }}>
          {[["sales","Revenue"],["gp","GP $"],["gpPct","GP %"]].map(([k,l])=>(
            <button key={k} style={S.subBtn(metric===k,"#0891B2")} onClick={()=>setMetric(k)}>{l}</button>
          ))}
        </div>
        <span style={{ fontSize:"0.72rem", color:MUTED, marginLeft:"auto" }}>{periodLabel}</span>
      </div>

      {/* Branch KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.75rem", marginBottom:"1rem" }}>
        {branchRows.map(r=>(
          <div key={r.b} style={{ ...S.card, borderTop:`4px solid ${BRANCH_COLORS[r.b]}`, marginBottom:0, padding:"0.85rem 1rem" }}>
            <div style={{ fontSize:"0.75rem", fontWeight:700, color:BRANCH_COLORS[r.b], marginBottom:"0.5rem" }}>{r.b.toUpperCase()}</div>
            <div style={{ fontSize:"1.05rem", fontWeight:700, color:TEXT }}>
              {metric==="gpPct" ? pct(r.gpPct) : metric==="gp" ? fmt(r.gCur) : fmt(r.cur)}
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:6, marginTop:4 }}>
              <span style={{ fontSize:"0.7rem", color:clr(r.chg), fontWeight:600 }}>
                {metric==="gpPct" ? pct(r.gpPct - (r.gPrev/r.prev||0)) : metric==="gp" ? fmt(r.gCur-r.gPrev) : fmt(r.chg)}
              </span>
              <span style={{ fontSize:"0.68rem", color:clr(r.chgPct), background:r.chgPct>=0?"#D1FAE5":"#FEE2E2", padding:"1px 5px", borderRadius:8, fontWeight:600 }}>
                {pct(r.chgPct)}
              </span>
            </div>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginTop:3 }}>vs {period==="ytd"?"YTD 2025":period==="q1"?"Q1 2025":"Q2 2025"}: {metric==="gpPct"?pct(r.gPrev/r.prev||0):metric==="gp"?fmt(r.gPrev):fmt(r.prev)}</div>
            <div style={{ marginTop:"0.5rem", height:3, background:BORDER, borderRadius:2 }}>
              <div style={{ height:3, background:BRANCH_COLORS[r.b], borderRadius:2, width:`${(r.cur/Math.max(...branchRows.map(x=>x.cur)))*100}%` }} />
            </div>
          </div>
        ))}
      </div>

      {/* Company totals bar */}
      <div style={{ ...S.card, marginBottom:"1rem", display:"flex", gap:"2rem", flexWrap:"wrap" }}>
        {[
          { label:"Company Total 2026", val:fmt(totalCur), color:AMBER },
          { label:"vs 2025", val:fmt(totalChg), color:clr(totalChg) },
          { label:`Change %`, val:pct(totalPrev?(totalChg/totalPrev):0), color:clr(totalChg) },
          { label:"Total GP$", val:fmt(totalGP), color:GREEN },
          { label:"GP %", val:pct(totalCur?totalGP/totalCur:0), color:GREEN },
        ].map(k=>(
          <div key={k.label}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.65rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.08em", marginTop:2 }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Weekly trend chart */}
      <div style={{ ...S.card, marginBottom:"1rem" }}>
        <div style={{ fontSize:"0.7rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>Weekly Revenue by Branch — 2026 (W1–W18)</div>
        <div style={{ height:220 }}>
          <Recharts.ResponsiveContainer width="100%" height="100%">
            <Recharts.LineChart data={chartData} margin={{ top:4, right:16, bottom:0, left:0 }}>
              <Recharts.CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
              <Recharts.XAxis dataKey="week" tick={{ fontSize:9, fill:MUTED }} />
              <Recharts.YAxis tick={{ fontSize:9, fill:MUTED }} tickFormatter={v=>`$${(v/1000).toFixed(0)}K`} width={55} />
              <Recharts.Tooltip formatter={(v,n)=>[fmt(v),n]} contentStyle={{ fontSize:"0.72rem", border:`1px solid ${BORDER}`, borderRadius:6 }} />
              <Recharts.Legend iconSize={8} wrapperStyle={{ fontSize:"0.7rem" }} />
              {Object.entries(BRANCH_COLORS).map(([b,c])=>(
                <Recharts.Line key={b} type="monotone" dataKey={b} stroke={c} strokeWidth={2} dot={false} />
              ))}
            </Recharts.LineChart>
          </Recharts.ResponsiveContainer>
        </div>
      </div>

      {/* Tifton dept breakdown */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>Tifton — Department Breakdown (Q1 Actuals)</div>
        <div style={{ overflowX:"auto" }}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Department</th>
                <th style={{ ...S.th, textAlign:"right" }}>2025</th>
                <th style={{ ...S.th, textAlign:"right" }}>2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>$ Change</th>
                <th style={{ ...S.th, textAlign:"right" }}>% Change</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP$ 2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP%</th>
                <th style={{ ...S.th, width:100 }}>Trend</th>
              </tr>
            </thead>
            <tbody>
              {tiftonDepts.map((d,i)=>{
                const chg = d.sales2026-d.sales2025;
                const chgPct = d.sales2025 ? chg/d.sales2025 : 0;
                const gpPct = d.sales2026 ? d.gp2026/d.sales2026 : 0;
                return (
                  <tr key={i}>
                    <td style={{ ...S.td, fontWeight:600 }}>{d.dept}</td>
                    <td style={{ ...S.td, textAlign:"right", color:MUTED }}>{fmt(d.sales2025)}</td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:600 }}>{fmt(d.sales2026)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:clr(chg), fontWeight:600 }}>{fmt(chg)}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>
                      <span style={{ fontSize:"0.68rem", background:chgPct>=0?"#D1FAE5":"#FEE2E2", color:clr(chgPct), borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct(chgPct)}</span>
                    </td>
                    <td style={{ ...S.td, textAlign:"right" }}>{fmt(d.gp2026)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:gpPct<0.08?RED:GREEN, fontWeight:600 }}>{pct(gpPct)}</td>
                    <td style={S.td}>
                      <div style={{ height:6, background:BORDER, borderRadius:3 }}>
                        <div style={{ height:6, background:chg>=0?GREEN:RED, borderRadius:3, width:`${Math.min(100,(d.sales2026/maxDeptSale)*100)}%` }} />
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── QTD Sub-Tab ───────────────────────────────────────────────────────────────
function QTDTab({ branchData }) {
  const branches = ["Byron","Tifton","Statesboro","Athens"];
  const BRANCH_COLORS = { Byron:"#1E5FCC", Tifton:"#7C3AED", Statesboro:"#0891B2", Athens:"#D97706" };

  const q1_26 = branches.map(b=>branchData.branches[b].q1_2026);
  const q1_25 = branches.map(b=>branchData.branches[b].q1_2025);
  const q2_26 = branches.map(b=>branchData.branches[b].q2_2026);
  const q2_25 = branches.map(b=>branchData.branches[b].q2_2025);

  const totalQ1_26 = q1_26.reduce((s,v)=>s+v,0);
  const totalQ1_25 = q1_25.reduce((s,v)=>s+v,0);
  const totalQ2_26 = q2_26.reduce((s,v)=>s+v,0);
  const totalQ2_25 = q2_25.reduce((s,v)=>s+v,0);
  const totalYTD26 = totalQ1_26 + totalQ2_26;
  const totalYTD25 = totalQ1_25 + totalQ2_25;

  const totalQ1_gp26 = branches.reduce((s,b)=>s+branchData.branches[b].q1_gp26,0);
  const totalQ2_gp26 = branches.reduce((s,b)=>s+branchData.branches[b].q2_gp26,0);
  const ytdGP26 = totalQ1_gp26 + totalQ2_gp26;

  // Bar chart data for Q1 vs Q2 by branch
  const barData = branches.map(b => ({
    branch: b,
    'Q1 2026': Math.round(branchData.branches[b].q1_2026),
    'Q1 2025': Math.round(branchData.branches[b].q1_2025),
    'Q2 2026': Math.round(branchData.branches[b].q2_2026),
    'Q2 2025': Math.round(branchData.branches[b].q2_2025),
  }));

  const WEEK_IN_Q2 = branchData.weeklySales.filter(w=>w.week>=14).length;

  return (
    <div>
      {/* Header banner */}
      <div style={{ ...S.card, background:"linear-gradient(135deg,#EEF4FF,#F4F7FB)", borderLeft:`4px solid ${AMBER}`, marginBottom:"1rem" }}>
        <div style={{ fontSize:"0.68rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:6 }}>Quarter-to-Date Performance</div>
        <div style={{ display:"flex", gap:"2.5rem", flexWrap:"wrap" }}>
          {[
            { label:"Q1 2026 (Complete)", val:fmt(totalQ1_26), sub:`vs ${fmt(totalQ1_25)} 2025`, chg:(totalQ1_26-totalQ1_25)/totalQ1_25 },
            { label:`Q2 2026 (W14–W${13+WEEK_IN_Q2}, ${WEEK_IN_Q2} wks)`, val:fmt(totalQ2_26), sub:`vs ${fmt(totalQ2_25)} 2025`, chg:(totalQ2_26-totalQ2_25)/totalQ2_25 },
            { label:"YTD 2026 (Q1+Q2)", val:fmt(totalYTD26), sub:`vs ${fmt(totalYTD25)} 2025`, chg:(totalYTD26-totalYTD25)/totalYTD25 },
            { label:"YTD GP$", val:fmt(ytdGP26), sub:`GP% ${pct(ytdGP26/totalYTD26)}`, chg:null },
          ].map(k=>(
            <div key={k.label}>
              <div style={{ fontSize:"0.95rem", fontWeight:700, color:AMBER }}>{k.val}</div>
              <div style={{ fontSize:"0.65rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.08em", marginTop:2 }}>{k.label}</div>
              {k.chg!==null && <div style={{ fontSize:"0.7rem", color:clr(k.chg), fontWeight:600, marginTop:2 }}>{pct(k.chg)} vs prior year</div>}
              {k.sub && k.chg===null && <div style={{ fontSize:"0.7rem", color:MUTED, marginTop:2 }}>{k.sub}</div>}
            </div>
          ))}
        </div>
      </div>

      {/* Q1 vs Q2 branch bar chart */}
      <div style={{ ...S.card, marginBottom:"1rem" }}>
        <div style={{ fontSize:"0.7rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>Q1 vs Q2 — Revenue by Branch (2026 vs 2025)</div>
        <div style={{ height:240 }}>
          <Recharts.ResponsiveContainer width="100%" height="100%">
            <Recharts.BarChart data={barData} margin={{ top:4, right:8, bottom:0, left:0 }}>
              <Recharts.CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
              <Recharts.XAxis dataKey="branch" tick={{ fontSize:10, fill:MUTED }} />
              <Recharts.YAxis tick={{ fontSize:9, fill:MUTED }} tickFormatter={v=>`$${(v/1000000).toFixed(1)}M`} width={55} />
              <Recharts.Tooltip formatter={(v,n)=>[fmt(v),n]} contentStyle={{ fontSize:"0.72rem", border:`1px solid ${BORDER}`, borderRadius:6 }} />
              <Recharts.Legend iconSize={8} wrapperStyle={{ fontSize:"0.7rem" }} />
              <Recharts.Bar dataKey="Q1 2026" fill="#1E5FCC" radius={[2,2,0,0]} />
              <Recharts.Bar dataKey="Q1 2025" fill="#93C5FD" radius={[2,2,0,0]} />
              <Recharts.Bar dataKey="Q2 2026" fill="#7C3AED" radius={[2,2,0,0]} />
              <Recharts.Bar dataKey="Q2 2025" fill="#C4B5FD" radius={[2,2,0,0]} />
            </Recharts.BarChart>
          </Recharts.ResponsiveContainer>
        </div>
      </div>

      {/* Branch QTD breakdown table */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>Branch QTD Detail</div>
        <div style={{ overflowX:"auto" }}>
          <table style={S.table}>
            <thead>
              <tr>
                <th style={S.th}>Branch</th>
                <th style={{ ...S.th, textAlign:"right" }}>Q1 2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>Q1 vs 2025</th>
                <th style={{ ...S.th, textAlign:"right", background:"#EEF4FF" }}>Q2 2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>Q2 vs 2025</th>
                <th style={{ ...S.th, textAlign:"right" }}>YTD 2026</th>
                <th style={{ ...S.th, textAlign:"right" }}>YTD vs 2025</th>
                <th style={{ ...S.th, textAlign:"right" }}>GP % YTD</th>
              </tr>
            </thead>
            <tbody>
              {branches.map(b=>{
                const bd = branchData.branches[b];
                const q1chgPct = (bd.q1_2026-bd.q1_2025)/bd.q1_2025;
                const q2chgPct = (bd.q2_2026-bd.q2_2025)/bd.q2_2025;
                const ytd26 = bd.q1_2026+bd.q2_2026;
                const ytd25 = bd.q1_2025+bd.q2_2025;
                const ytdChg = (ytd26-ytd25)/ytd25;
                const ytdGP  = bd.q1_gp26+bd.q2_gp26;
                const gpPct  = ytd26 ? ytdGP/ytd26 : 0;
                return (
                  <tr key={b}>
                    <td style={{ ...S.td, fontWeight:700, color:BRANCH_COLORS[b] }}>{b}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>{fmt(bd.q1_2026)}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>
                      <span style={{ fontSize:"0.68rem", background:q1chgPct>=0?"#D1FAE5":"#FEE2E2", color:clr(q1chgPct), borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct(q1chgPct)}</span>
                    </td>
                    <td style={{ ...S.td, textAlign:"right", background:"#F8FAFF" }}>{fmt(bd.q2_2026)}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>
                      <span style={{ fontSize:"0.68rem", background:q2chgPct>=0?"#D1FAE5":"#FEE2E2", color:clr(q2chgPct), borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct(q2chgPct)}</span>
                    </td>
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700 }}>{fmt(ytd26)}</td>
                    <td style={{ ...S.td, textAlign:"right" }}>
                      <span style={{ fontSize:"0.68rem", background:ytdChg>=0?"#D1FAE5":"#FEE2E2", color:clr(ytdChg), borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct(ytdChg)}</span>
                    </td>
                    <td style={{ ...S.td, textAlign:"right", color:gpPct<0.10?RED:GREEN, fontWeight:600 }}>{pct(gpPct)}</td>
                  </tr>
                );
              })}
              <tr style={{ borderTop:`2px solid ${BORDER}` }}>
                <td style={{ ...S.td, fontWeight:700 }}>COMPANY</td>
                <td style={{ ...S.td, textAlign:"right", fontWeight:700 }}>{fmt(totalQ1_26)}</td>
                <td style={{ ...S.td, textAlign:"right" }}>
                  <span style={{ fontSize:"0.68rem", background:"#D1FAE5", color:GREEN, borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct((totalQ1_26-totalQ1_25)/totalQ1_25)}</span>
                </td>
                <td style={{ ...S.td, textAlign:"right", fontWeight:700, background:"#F8FAFF" }}>{fmt(totalQ2_26)}</td>
                <td style={{ ...S.td, textAlign:"right" }}>
                  <span style={{ fontSize:"0.68rem", background:(totalQ2_26-totalQ2_25)>=0?"#D1FAE5":"#FEE2E2", color:clr(totalQ2_26-totalQ2_25), borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct((totalQ2_26-totalQ2_25)/totalQ2_25)}</span>
                </td>
                <td style={{ ...S.td, textAlign:"right", fontWeight:700 }}>{fmt(totalYTD26)}</td>
                <td style={{ ...S.td, textAlign:"right" }}>
                  <span style={{ fontSize:"0.68rem", background:"#D1FAE5", color:GREEN, borderRadius:8, padding:"1px 6px", fontWeight:700 }}>{pct((totalYTD26-totalYTD25)/totalYTD25)}</span>
                </td>
                <td style={{ ...S.td, textAlign:"right", fontWeight:700, color:ytdGP26/totalYTD26<0.10?RED:GREEN }}>{pct(ytdGP26/totalYTD26)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}



// ── Ascenso Tab ───────────────────────────────────────────────────────────────
function AscensoTab({ ascenso, custName }) {
  // Tiers are company-wide dollar thresholds
  const totalAmt = ASCENSO_TOTAL;
  const tier     = getAdTier(totalAmt, ASCENSO_TIERS);
  const nextTier = getNextAdTier(totalAmt, ASCENSO_TIERS);
  const toNext   = nextTier ? nextTier.min - totalAmt : 0;
  const isClose  = nextTier && toNext <= nextTier.min * 0.15;
  const tColor   = tier.color;
  const avgPrice = ascenso.qty > 0 ? ascenso.amount / ascenso.qty : 0;
  const currentPayout = totalAmt * (tier.payout || 0);
  const nextPayout    = nextTier ? nextTier.min * nextTier.payout : null;
  const acctShare     = totalAmt > 0 ? ascenso.amount / totalAmt : 0;

  function tierProgress(amt, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => amt >= t.min);
    const cur = asc[curIdx];
    const nxt = asc[curIdx + 1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((amt - cur.min) / (nxt.min - cur.min)) * 100));
  }
  const progress = tierProgress(totalAmt, ASCENSO_TIERS);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:`4px solid ${tColor}`, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Ascenso (ASC) — Manufacturer Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>Agricultural, Industrial &amp; Specialty Tires</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#D97706", background:"#FEF3C7", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #FCD34D" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"Company Total",   val:fmt(totalAmt),                 color:tColor },
          { label:"This Acct",       val:fmt(ascenso.amount),           color:AMBER  },
          { label:"Est. Payout",     val:fmt(currentPayout),            color:GREEN  },
          { label:"Acct Contribution",val:pct(acctShare),               color:MUTED  },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({fmt(totalAmt)} company total)</span>
          {nextTier
            ? <span style={{ color:isClose?"#D97706":MUTED, fontWeight:isClose?700:400 }}>{toNext} units needed for {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>
          }
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{
            height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s ease",
            background: isClose ? "linear-gradient(90deg,#D97706,#FBBF24)" : `linear-gradient(90deg,${tColor},${tColor}99)`
          }} />
        </div>
        {/* Tier markers */}
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...ASCENSO_TIERS].reverse().map(t => {
            const active = ascenso.qty >= t.min;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background: active ? t.color : BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.62rem", color: active ? t.color : BORDER, fontWeight: active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.58rem", color:MUTED }}>{t.min>=1000?`$${(t.min/1000).toFixed(0)}K`:t.min>0?`$${t.min}`:""}</div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Activity detail */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>Purchase Activity</div>
        {/* Payout summary */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"0.6rem", marginBottom:"1rem" }}>
          {[
            { label:"Current Payout",  val:fmt(currentPayout),  sub:`${(tier.payout*100).toFixed(0)}% of ${fmt(totalAmt)}`, color:GREEN },
            { label:"Next Tier Payout",val:nextPayout?fmt(nextPayout):"★ Max", sub:nextTier?`${(nextTier.payout*100).toFixed(0)}% of ${fmt(nextTier.min)}`:"Top tier reached", color:nextTier?"#7C3AED":MUTED },
            { label:"Payout Gain",     val:nextTier?fmt(nextPayout-currentPayout):"—", sub:nextTier?`+${fmt(toNext)} needed`:"Already at max", color:nextTier?AMBER:MUTED },
          ].map(k=>(
            <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.6rem 0.75rem", borderLeft:`3px solid ${k.color}` }}>
              <div style={{ fontSize:"0.85rem", fontWeight:700, color:k.color }}>{k.val}</div>
              <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2 }}>{k.label}</div>
              <div style={{ fontSize:"0.6rem", color:MUTED, marginTop:1 }}>{k.sub}</div>
            </div>
          ))}
        </div>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.6rem" }}>This Account</div>
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {[
            { label:"Units Purchased",      val:ascenso.qty,       icon:"📦" },
            { label:"Revenue",              val:fmt(ascenso.amount), icon:"💰" },
            { label:"Avg Price/Unit",       val:`$${avgPrice.toFixed(2)}`, icon:"🏷" },
            { label:"Line Items (Details)", val:ascenso.details,   icon:"📋" },
            { label:"Invoices",             val:ascenso.invoices,  icon:"🧾" },
            { label:"Units per Invoice",    val:(ascenso.qty/(ascenso.invoices||1)).toFixed(1), icon:"📈" },
            { label:"% of Company Total",   val:pct(acctShare),    icon:"🏢" },
          ].map(row => (
            <div key={row.label} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"0.4rem 0.6rem", background:"#F4F7FB", borderRadius:6 }}>
              <span style={{ fontSize:"0.75rem", color:MUTED }}>{row.icon} {row.label}</span>
              <span style={{ fontSize:"0.78rem", fontWeight:700, color:TEXT }}>{row.val}</span>
            </div>
          ))}
        </div>

        {ascenso.details > ascenso.invoices && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#EFF6FF", border:`1px solid ${BORDER}`, borderRadius:6, fontSize:"0.72rem", color:MUTED }}>
            💡 {ascenso.details} line items across {ascenso.invoices} invoices — customer orders multiple SKUs per visit.
          </div>
        )}
      </div>
    </div>
  );
}





// ── Falken TBR Tab ────────────────────────────────────────────────────────────
function FalkenTBRTab({ falken, custName }) {
  const tier     = getAdTier(falken.ytd, FALKEN_TBR_TIERS);
  const nextTier = getNextAdTier(falken.ytd, FALKEN_TBR_TIERS);
  const toNext   = nextTier ? nextTier.min - falken.ytd : 0;
  const isClose  = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
  const tColor   = tier.color === "#6B7A99" ? MUTED : tier.color;

  function tierProgress(units, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => units >= t.min);
    const cur = asc[curIdx];
    const nxt = asc[curIdx + 1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((units - cur.min) / (nxt.min - cur.min)) * 100));
  }
  const progress = tierProgress(falken.ytd, FALKEN_TBR_TIERS);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:"4px solid #0891B2", marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Falken Fanatic — TBR Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>
              Fanatic ID #{falken.falkenId} · {falken.city} · as of 5/22/2026
            </div>
            <div style={{ fontSize:"0.68rem", color:"#0891B2", fontWeight:600, marginTop:4 }}>
              Truck / Bus / Radial Tires
            </div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#059669", background:"#D1FAE5", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #6EE7B7" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"YTD Units",  val: falken.ytd, color: tColor    },
          { label:"Q1 Units",   val: falken.q1,  color: MUTED     },
          { label:"Q2 Units",   val: falken.q2,  color: "#0891B2" },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({falken.q2} Q2 units)</span>
          {nextTier
            ? <span style={{ color:isClose?"#059669":MUTED, fontWeight:isClose?700:400 }}>{toNext} units to {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>
          }
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{
            height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s ease",
            background: isClose ? "linear-gradient(90deg,#059669,#34D399)" : "linear-gradient(90deg,#0891B2,#67E8F9)"
          }} />
        </div>
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...FALKEN_TBR_TIERS].reverse().map(t => {
            const active = falken.ytd >= t.min;
            const c2 = t.color === "#6B7A99" ? MUTED : t.color;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background:active?c2:BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.58rem", color:active?c2:BORDER, fontWeight:active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.55rem", color:MUTED }}>{t.min>0?`${t.min}+`:""}</div>
              </div>
            );
          })}
        </div>
        {isClose && nextTier && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#D1FAE5", border:"1px solid #6EE7B7", borderRadius:6, fontSize:"0.72rem", color:"#065F46", fontWeight:600 }}>
            ⚡ Only {toNext} more TBR units to reach {nextTier.label}!
          </div>
        )}
      </div>

      {/* Q1 vs Q2 */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.85rem" }}>Quarter Breakdown</div>
        <div style={{ display:"flex", gap:"1rem", alignItems:"flex-end" }}>
          {[
            { label:"Q1 (Jan–Mar)", val:falken.q1, color:MUTED },
            { label:"Q2 (Apr–Jun)", val:falken.q2, color:"#0891B2" },
          ].map(b => {
            const maxV = Math.max(falken.q1, falken.q2, 1);
            return (
              <div key={b.label} style={{ flex:1, textAlign:"center" }}>
                <div style={{ fontSize:"1rem", fontWeight:700, color:b.color, marginBottom:6 }}>{b.val}</div>
                <div style={{ height:80, background:BORDER, borderRadius:6, position:"relative", overflow:"hidden" }}>
                  <div style={{ position:"absolute", bottom:0, left:0, right:0, height:`${(Math.max(0,b.val)/maxV)*100}%`, background:b.color, borderRadius:6, transition:"height 0.5s ease" }} />
                </div>
                <div style={{ fontSize:"0.65rem", color:MUTED, marginTop:6 }}>{b.label}</div>
              </div>
            );
          })}
          <div style={{ flex:2, padding:"0.75rem", background:"#F4F7FB", borderRadius:8 }}>
            <div style={{ fontSize:"0.75rem", color:MUTED }}>
              {falken.q1 === 0
                ? "Q1 had no TBR activity — all units purchased in Q2"
                : `${falken.q2} Q2 units vs ${falken.q1} Q1 units`}
            </div>
            <div style={{ marginTop:8, fontSize:"0.68rem", color:"#0891B2", fontWeight:600 }}>
              💡 TBR tier thresholds are placeholder — update when confirmed
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Falken PLT Tab ────────────────────────────────────────────────────────────
function FalkenPLTTab({ falken, custName }) {
  const tier     = getAdTier(falken.ytd, FALKEN_PLT_TIERS);
  const nextTier = getNextAdTier(falken.ytd, FALKEN_PLT_TIERS);
  const toNext   = nextTier ? nextTier.min - falken.ytd : 0;
  const isClose  = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
  const tColor   = tier.color === "#6B7A99" ? MUTED : tier.color;
  const q2vsQ1   = falken.q1 > 0 ? ((falken.q2 - falken.q1) / falken.q1) : null;

  function tierProgress(units, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => units >= t.min);
    const cur = asc[curIdx];
    const nxt = asc[curIdx + 1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((units - cur.min) / (nxt.min - cur.min)) * 100));
  }
  const progress = tierProgress(falken.ytd, FALKEN_PLT_TIERS);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:"4px solid #059669", marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Falken Fanatic — PLT Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>
              Fanatic ID #{falken.falkenId} · {falken.city} · as of 5/14/2026
            </div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#059669", background:"#D1FAE5", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #6EE7B7" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"Q2 Units",    val: falken.q2,   color: tColor    },
          { label:"Q1 Units",    val: falken.q1,   color: MUTED     },
          { label:"Q2 vs Q1",   val: q2vsQ1 !== null ? pct(q2vsQ1) : "N/A",
            color: q2vsQ1 === null ? MUTED : q2vsQ1 >= 0 ? GREEN : RED },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({falken.ytd} units YTD)</span>
          {nextTier
            ? <span style={{ color:isClose?"#059669":MUTED, fontWeight:isClose?700:400 }}>{toNext} units to {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>
          }
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{
            height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s ease",
            background: isClose ? "linear-gradient(90deg,#059669,#34D399)" : "linear-gradient(90deg,#059669,#6EE7B7)"
          }} />
        </div>
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...FALKEN_PLT_TIERS].reverse().map(t => {
            const active = falken.ytd >= t.min;
            const c2 = t.color === "#6B7A99" ? MUTED : t.color;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background:active?c2:BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.6rem", color:active?c2:BORDER, fontWeight:active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.55rem", color:MUTED }}>{t.min>0?`${t.min}+`:""}</div>
              </div>
            );
          })}
        </div>
        {isClose && nextTier && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#D1FAE5", border:"1px solid #6EE7B7", borderRadius:6, fontSize:"0.72rem", color:"#065F46", fontWeight:600 }}>
            ⚡ Only {toNext} more Q2 units to reach {nextTier.label}!
          </div>
        )}
      </div>

      {/* Q1 vs Q2 breakdown */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.85rem" }}>Quarter Breakdown</div>
        <div style={{ display:"flex", gap:"1rem", alignItems:"flex-end", marginBottom:"1rem" }}>
          {[
            { label:"Q1 (Jan–Mar)", val:falken.q1, color:AMBER },
            { label:"Q2 (Apr–Jun)", val:falken.q2, color:"#059669" },
          ].map(b => {
            const maxV = Math.max(falken.q1, falken.q2, 1);
            return (
              <div key={b.label} style={{ flex:1, textAlign:"center" }}>
                <div style={{ fontSize:"1rem", fontWeight:700, color:b.color, marginBottom:6 }}>{b.val}</div>
                <div style={{ height:80, background:BORDER, borderRadius:6, position:"relative", overflow:"hidden" }}>
                  <div style={{ position:"absolute", bottom:0, left:0, right:0, height:`${(Math.max(0,b.val)/maxV)*100}%`, background:b.color, borderRadius:6, transition:"height 0.5s ease" }} />
                </div>
                <div style={{ fontSize:"0.65rem", color:MUTED, marginTop:6 }}>{b.label}</div>
              </div>
            );
          })}
          <div style={{ flex:2, padding:"0.75rem", background:"#F4F7FB", borderRadius:8 }}>
            {q2vsQ1 !== null ? (
              <>
                <div style={{ fontSize:"0.85rem", fontWeight:700, color:clr(q2vsQ1), marginBottom:4 }}>
                  {q2vsQ1 >= 0 ? "▲" : "▼"} {pct(Math.abs(q2vsQ1))} Q2 vs Q1
                </div>
                <div style={{ fontSize:"0.68rem", color:MUTED }}>
                  {q2vsQ1 >= 0 ? "Momentum building in Q2" : "Q2 pace is slower than Q1"}
                </div>
              </>
            ) : (
              <div style={{ fontSize:"0.72rem", color:MUTED }}>Q1 data not available for comparison</div>
            )}
            <div style={{ marginTop:8, fontSize:"0.68rem", color:MUTED, background:"#FEF3C7", padding:"4px 8px", borderRadius:4 }}>
              💡 TBR data will be added separately
            </div>
          </div>
        </div>


      </div>
    </div>
  );
}


// ── Yokohama Tab ──────────────────────────────────────────────────────────────
function YokohamaTab({ yoko, custName }) {
  const tier     = getAdTier(yoko.ytd, YOKOHAMA_TIERS);
  const nextTier = getNextAdTier(yoko.ytd, YOKOHAMA_TIERS);
  const toNext   = nextTier ? nextTier.min - yoko.ytd : 0;
  const isClose  = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
  const tColor   = tier.color === "#6B7A99" ? MUTED : tier.color;

  function tierProgress(units, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => units >= t.min);
    const cur = asc[curIdx]; const nxt = asc[curIdx+1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((units-cur.min)/(nxt.min-cur.min))*100));
  }
  const progress = tierProgress(yoko.ytd, YOKOHAMA_TIERS);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:"4px solid #7C3AED", marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Yokohama Dealer Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>Program #{yoko.progNum} · Rep: {yoko.rep} · as of {yoko.asOf}</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#7C3AED", background:"#EDE9FE", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #C4B5FD" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"QTD Units",     val: yoko.ytd,      color: tColor },
          { label:"Primary Units", val: yoko.primary,  color: yoko.priPct < 75 ? RED : AMBER },
          { label:"Secondary",     val: yoko.secondary||0, color: MUTED },
          { label:"Primary %",     val: `${yoko.priPct}%`, color: yoko.priPct >= 80 ? GREEN : yoko.priPct >= 75 ? "#D97706" : RED },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Primary % warning */}
      {yoko.priPct < 75 && (
        <div style={{ ...S.card, background:"#FEF2F2", border:"2px solid #FECACA", marginBottom:"0.85rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:RED }}>
            🚨 Primary purchase % is {yoko.priPct}% — below the 75% minimum required for Yokohama rebate payout.
          </div>
        </div>
      )}

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({yoko.ytd} units)</span>
          {nextTier
            ? <span style={{ color:isClose?"#7C3AED":MUTED, fontWeight:isClose?700:400 }}>{toNext} units to {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>}
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{ height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s",
            background: isClose ? "linear-gradient(90deg,#7C3AED,#A78BFA)" : "linear-gradient(90deg,#7C3AED,#C4B5FD)" }} />
        </div>
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...YOKOHAMA_TIERS].reverse().map(t => {
            const active = yoko.ytd >= t.min;
            const c2 = t.color === "#6B7A99" ? MUTED : t.color;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background:active?c2:BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.6rem", color:active?c2:BORDER, fontWeight:active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.55rem", color:MUTED }}>{t.min>0?`${t.min}+`:""}</div>
              </div>
            );
          })}
        </div>
        {yoko.toNext > 0 && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#EDE9FE", border:"1px solid #C4B5FD", borderRadius:6, fontSize:"0.72rem", color:"#5B21B6", fontWeight:600 }}>
            ⚡ {yoko.toNext} more units needed to reach {nextTier?.label} — tiers are placeholders, update when confirmed.
          </div>
        )}
      </div>
    </div>
  );
}

// ── BF BARNN Tab ──────────────────────────────────────────────────────────────
function BARNNTab({ barnn, custName }) {
  const tier     = getAdTier((barnn.total||barnn.ytd||0), BARNN_TIERS);
  const nextTier = getNextAdTier((barnn.total||barnn.ytd||0), BARNN_TIERS);
  const toNext   = nextTier ? nextTier.min - barnn.ytd : 0;
  const isClose  = nextTier && toNext <= Math.max(10, Math.round(nextTier.min * 0.15));
  const tColor   = tier.color === "#6B7A99" ? MUTED : tier.color;
  const changePct = barnn.changePct;

  function tierProgress(units, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => units >= t.min);
    const cur = asc[curIdx];
    const nxt = asc[curIdx + 1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((units - cur.min) / (nxt.min - cur.min)) * 100));
  }
  const progress = tierProgress((barnn.total||barnn.ytd||0), BARNN_TIERS);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:"4px solid #D97706", marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Bridgestone Firestone BARNN Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>
              Program # {barnn.progNum} · YTD through {barnn.asOf}
            </div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#D97706", background:"#FEF3C7", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #FCD34D" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"QTD Units",     val: barnn.ytd,   color: tColor },
          { label:"Primary Units", val: barnn.primary, color: (barnn.priPct||100) < 75 ? RED : AMBER },
          { label:"Secondary",     val: barnn.secondary||0, color: MUTED },
          { label:"Primary %",     val: `${barnn.priPct||0}%`, color: (barnn.priPct||100) >= 80 ? GREEN : (barnn.priPct||100) >= 75 ? "#D97706" : RED },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({(barnn.total||barnn.ytd||0)} units)</span>
          {nextTier
            ? <span style={{ color:isClose?"#D97706":MUTED, fontWeight:isClose?700:400 }}>{toNext} units to {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>
          }
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{
            height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s ease",
            background: isClose ? "linear-gradient(90deg,#D97706,#FBBF24)" : "linear-gradient(90deg,#D97706,#FCD34D)"
          }} />
        </div>
        {/* Tier markers */}
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...BARNN_TIERS].reverse().map(t => {
            const active = barnn.ytd >= t.min;
            const c2 = t.color === "#6B7A99" ? MUTED : t.color;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background:active?c2:BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.62rem", color:active?c2:BORDER, fontWeight:active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.58rem", color:MUTED }}>{t.min>0?`${t.min}+`:""}</div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Primary % warning */}
      {(barnn.priPct||100) < 75 && (
        <div style={{ ...S.card, background:"#FEF2F2", border:"2px solid #FECACA", marginBottom:"0.85rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:RED }}>
            🚨 Primary purchase % is {barnn.priPct||0}% — below the 75% threshold required for BARNN rebate payout.
          </div>
          <div style={{ fontSize:"0.72rem", color:"#991B1B", marginTop:4 }}>
            Customer needs to shift more purchases to Tire Distributors of GA as primary supplier to qualify.
          </div>
        </div>
      )}

      {/* To Next Tier */}
      {barnn.toNext > 0 && (
        <div style={{ ...S.card, background:"#FFFBEB", border:"1px solid #FDE68A", marginBottom:"0.85rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:"#D97706" }}>
            ⚡ {barnn.toNext} more units needed to reach {nextTier?.label || "next tier"} ({(nextTier?.min||0)} total)
          </div>
        </div>
      )}
      {barnn.toNext === 0 && barnn.ytd >= 75 && (
        <div style={{ ...S.card, background:"#F0FDF4", border:"1px solid #BBF7D0", marginBottom:"0.85rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:GREEN }}>✓ At or above current tier threshold</div>
        </div>
      )}

      {/* Program info */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.85rem" }}>Program Info</div>
        <div style={{ display:"flex", gap:"1.5rem", alignItems:"flex-end", marginBottom:"1rem" }}>
          {[
            { label:`PY (${parseInt(barnn.asOf)-1} YTD)`, val:barnn.py,  color:MUTED },
            { label:`YTD ${barnn.asOf}`,                   val:barnn.ytd, color:"#D97706" },
          ].map(b => {
            const maxVal = Math.max(barnn.ytd, barnn.py, 1);
            return (
              <div key={b.label} style={{ flex:1, textAlign:"center" }}>
                <div style={{ fontSize:"0.95rem", fontWeight:700, color:b.color, marginBottom:6 }}>{b.val}</div>
                <div style={{ height:80, background:BORDER, borderRadius:6, position:"relative", overflow:"hidden" }}>
                  <div style={{
                    position:"absolute", bottom:0, left:0, right:0,
                    height:`${(b.val/maxVal)*100}%`,
                    background:b.color, borderRadius:6, transition:"height 0.5s ease"
                  }} />
                </div>
                <div style={{ fontSize:"0.65rem", color:MUTED, marginTop:6 }}>{b.label}</div>
              </div>
            );
          })}
          <div style={{ flex:2, padding:"0.75rem", background:"#F4F7FB", borderRadius:8, alignSelf:"center" }}>
            <div style={{ fontSize:"0.78rem", fontWeight:700, color:clr(barnn.change), marginBottom:4 }}>
              {barnn.change > 0 ? "▲" : "▼"} {Math.abs(barnn.change)} units {barnn.change > 0 ? "growth" : "decline"}
            </div>
            <div style={{ fontSize:"0.68rem", color:MUTED }}>
              {barnn.py > 0
                ? `${(changePct*100).toFixed(0)}% vs prior year`
                : "New account this year — no prior year baseline"}
            </div>
            {isClose && nextTier && (
              <div style={{ marginTop:8, fontSize:"0.68rem", fontWeight:700, color:"#D97706" }}>
                ⚡ {toNext} more units = {nextTier.label} tier
              </div>
            )}
          </div>
        </div>


      </div>
    </div>
  );
}

// ── Americus Tab ──────────────────────────────────────────────────────────────
function AmericusTab({ americus, custName }) {
  const tier     = getAdTier(americus.ytd, AMERICUS_TIERS);
  const nextTier = getNextAdTier(americus.ytd, AMERICUS_TIERS);
  const toNext   = nextTier ? nextTier.min - americus.ytd : 0;
  const gap      = nextTier ? nextTier.min - tier.min : 0;
  const isClose  = nextTier && toNext <= Math.max(10, Math.round(gap * 0.15));
  const tColor   = tier.color;

  const vsLY     = americus.units2025 > 0 ? (americus.ytd - americus.units2025) / americus.units2025 : null;
  const q2pace   = americus.q1 > 0 ? americus.q2 / americus.q1 : null;

  function tierProgress(units, tiers) {
    const asc = [...tiers].reverse();
    const curIdx = asc.findIndex(t => units >= t.min);
    const cur = asc[curIdx];
    const nxt = asc[curIdx + 1];
    if (!nxt) return 100;
    return Math.min(100, Math.round(((units - cur.min) / (nxt.min - cur.min)) * 100));
  }
  const progress = tierProgress(americus.ytd, AMERICUS_TIERS);

  // Monthly breakdown for Q2
  const q2months = [
    { label:"April", val: americus.apr },
    { label:"May",   val: americus.may },
    { label:"June",  val: americus.jun },
  ];
  const maxMonth = Math.max(...q2months.map(m => m.val), 1);

  return (
    <div>
      {/* Header */}
      <div style={{ ...S.card, borderLeft:`4px solid ${tColor}`, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>Americus Partners Program</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>
              Dealer #{americus.amerNum} · Enrolled {americus.enrollYear} · Primary: {americus.primary ? "Yes" : "No"} · As of {americus.asOf}
            </div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.78rem", fontWeight:800, color:tColor, background:tColor+"22", padding:"0.25rem 0.85rem", borderRadius:20, border:`2px solid ${tColor}` }}>
              {tier.label}
            </span>
            {isClose && (
              <span style={{ fontSize:"0.72rem", fontWeight:700, color:"#7C3AED", background:"#EDE9FE", padding:"0.25rem 0.7rem", borderRadius:20, border:"2px solid #C4B5FD" }}>
                ⚡ {toNext} to {nextTier.label}!
              </span>
            )}
          </div>
        </div>
      </div>

      {/* KPI cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
        {[
          { label:"YTD Units",   val:americus.ytd,        color:tColor },
          { label:"Q1 Units",    val:americus.q1,         color:AMBER  },
          { label:"Q2 Units",    val:americus.q2,         color:TEAL   },
          { label:"2025 Full Yr",val:americus.units2025,  color:MUTED  },
        ].map(k=>(
          <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.65rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
            <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Tier progress */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>
          <span style={{ fontWeight:600 }}>Tier Progress — {tier.label} ({americus.ytd} units YTD)</span>
          {nextTier
            ? <span style={{ color:isClose?"#7C3AED":MUTED, fontWeight:isClose?700:400 }}>{toNext} units to {nextTier.label} ({nextTier.min}+)</span>
            : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>
          }
        </div>
        <div style={{ height:12, background:BORDER, borderRadius:6, overflow:"hidden", marginBottom:8 }}>
          <div style={{
            height:12, borderRadius:6, width:`${progress}%`, transition:"width 0.5s ease",
            background: isClose ? "linear-gradient(90deg,#7C3AED,#A78BFA)" : `linear-gradient(90deg,${tColor},${tColor}99)`
          }} />
        </div>
        {/* Tier markers */}
        <div style={{ display:"flex", justifyContent:"space-between" }}>
          {[...AMERICUS_TIERS].reverse().map(t => {
            const active = americus.ytd >= t.min;
            return (
              <div key={t.label} style={{ textAlign:"center", flex:1 }}>
                <div style={{ width:2, height:6, background: active ? t.color : BORDER, margin:"0 auto 3px" }} />
                <div style={{ fontSize:"0.62rem", color:active?t.color:BORDER, fontWeight:active?700:400 }}>{t.label}</div>
                <div style={{ fontSize:"0.58rem", color:MUTED }}>{t.min>0?`${t.min}+`:""}</div>
              </div>
            );
          })}
        </div>

        {/* vs LY and pace */}
        <div style={{ display:"flex", gap:"1.5rem", marginTop:"0.85rem", paddingTop:"0.75rem", borderTop:`1px solid ${BORDER}`, flexWrap:"wrap" }}>
          {vsLY !== null && (
            <div>
              <div style={{ fontSize:"0.85rem", fontWeight:700, color:clr(vsLY) }}>{pct(vsLY)}</div>
              <div style={{ fontSize:"0.62rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.06em" }}>vs 2025 Pace</div>
            </div>
          )}
          {q2pace !== null && (
            <div>
              <div style={{ fontSize:"0.85rem", fontWeight:700, color: q2pace >= 0.8 ? GREEN : q2pace >= 0.5 ? "#D97706" : RED }}>
                {pct(q2pace)}
              </div>
              <div style={{ fontSize:"0.62rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.06em" }}>Q2 vs Q1 Pace</div>
            </div>
          )}
          {vsLY !== null && americus.units2025 > 0 && (
            <div>
              <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>{americus.units2025}</div>
              <div style={{ fontSize:"0.62rem", color:MUTED, textTransform:"uppercase", letterSpacing:"0.06em" }}>
                Full Year 2025 ({getAdTier(americus.units2025, AMERICUS_TIERS).label})
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Q2 monthly breakdown */}
      <div style={S.card}>
        <div style={{ fontSize:"0.7rem", fontWeight:600, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.85rem" }}>Q2 Monthly Breakdown</div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"0.75rem" }}>
          {q2months.map(m => (
            <div key={m.label} style={{ textAlign:"center" }}>
              <div style={{ fontSize:"1rem", fontWeight:700, color: m.val > 0 ? AMBER : m.val < 0 ? RED : MUTED }}>{m.val}</div>
              <div style={{ fontSize:"0.68rem", color:MUTED, marginBottom:8 }}>{m.label}</div>
              <div style={{ height:60, background:BORDER, borderRadius:4, position:"relative", overflow:"hidden" }}>
                <div style={{
                  position:"absolute", bottom:0, left:0, right:0,
                  height:`${m.val > 0 ? (m.val/maxMonth)*100 : 0}%`,
                  background: m.val > 0 ? `linear-gradient(180deg,${tColor},${tColor}99)` : RED,
                  borderRadius:4, transition:"height 0.5s ease"
                }} />
              </div>
            </div>
          ))}
        </div>
        {americus.jun === 0 && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#FEF3C7", border:"1px solid #FCD34D", borderRadius:6, fontSize:"0.72rem", color:"#92400E" }}>
            ⚠ June data not yet reported (as of {americus.asOf})
          </div>
        )}
        {isClose && nextTier && (
          <div style={{ marginTop:"0.75rem", padding:"0.6rem 0.85rem", background:"#EDE9FE", border:"2px solid #C4B5FD", borderRadius:6, fontSize:"0.75rem", color:"#5B21B6", fontWeight:600 }}>
            ⚡ Only {toNext} units needed to reach {nextTier.label} ({nextTier.min} total) — push {Math.ceil(toNext / Math.max(1, q2months.filter(m=>m.val>0).length))} more units per active month.
          </div>
        )}
      </div>
    </div>
  );
}

// ── AD Program Tab ────────────────────────────────────────────────────────────
function ADProgramTab({ adProgram, custName }) {
  const pcr = adProgram.pcr;
  const tbr = adProgram.tbr;

  const pcrTier    = getAdTier(pcr.total, AD_PCR_TIERS);
  const pcrNext    = getNextAdTier(pcr.total, AD_PCR_TIERS);
  const tbrTier    = getAdTier(tbr.total, AD_TBR_TIERS);
  const tbrNext    = getNextAdTier(tbr.total, AD_TBR_TIERS);

  const pcrToNext  = pcrNext ? pcrNext.min - pcr.total : 0;
  const tbrToNext  = tbrNext ? tbrNext.min - tbr.total : 0;
  const pcrIsClose = pcrNext && pcrToNext <= Math.max(20, Math.round((pcrNext.min - (AD_PCR_TIERS.find(t=>t.min<=pcr.total&&t!==pcrNext)?.min||0)) * 0.15));
  const tbrIsClose = tbrNext && tbrToNext <= Math.max(20, Math.round((tbrNext.min - (AD_TBR_TIERS.find(t=>t.min<=tbr.total&&t!==tbrNext)?.min||0)) * 0.15));

  // Progress bar: position within current tier band
  function tierProgress(units, tiers) {
    const tiersSorted = [...tiers].reverse(); // ascending
    const curIdx = tiersSorted.findIndex(t => units >= t.min);
    const cur = tiersSorted[curIdx];
    const nxt = tiersSorted[curIdx + 1];
    if (!nxt) return 100; // at top tier
    const range = nxt.min - cur.min;
    const progress = units - cur.min;
    return Math.min(100, Math.round((progress / range) * 100));
  }

  const pcrProgress = tierProgress(pcr.total, AD_PCR_TIERS);
  const tbrProgress = tierProgress(tbr.total, AD_TBR_TIERS);
  const pcrColor = pcrTier.color === "#MUTED" ? MUTED : pcrTier.color;
  const tbrColor = tbrTier.color === "#MUTED" ? MUTED : tbrTier.color;

  return (
    <div>
      {/* Program header */}
      <div style={{ ...S.card, borderLeft:`4px solid ${AMBER}`, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
          <div>
            <div style={{ fontSize:"0.85rem", fontWeight:700, color:TEXT }}>{adProgram.program}</div>
            <div style={{ fontSize:"0.72rem", color:MUTED, marginTop:2 }}>District 360 · Branch: {adProgram.branch} · Toyo # {adProgram.toyoNum}</div>
          </div>
          <div style={{ fontSize:"0.68rem", color:MUTED, textAlign:"right" }}>
            <div>Primary Direct Dealer</div>
            <div style={{ color:TEXT, fontWeight:600 }}>Tire Distributors of GA — {adProgram.branch}</div>
          </div>
        </div>
      </div>

      {/* PCR / LTR Program */}
      <div style={{ ...S.card, marginBottom:"0.85rem" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em" }}>PCR / LTR Program</div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.72rem", fontWeight:800, color:pcrColor, background: pcrColor === MUTED ? "#F4F7FB" : pcrColor+"22", padding:"0.2rem 0.75rem", borderRadius:20, border:`2px solid ${pcrColor}` }}>
              {pcrTier.label}
            </span>
            {pcrIsClose && (
              <span style={{ fontSize:"0.7rem", fontWeight:700, color:"#D97706", background:"#FEF3C7", padding:"0.2rem 0.65rem", borderRadius:20, border:"2px solid #FCD34D", animation:"pulse 1s infinite" }}>
                ⚡ {pcrToNext} to {pcrNext.label}!
              </span>
            )}
          </div>
        </div>

        {/* Unit breakdown */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
          {[
            { label:"Primary Units",   val:pcr.primary,   color:AMBER },
            { label:"Secondary Units", val:pcr.secondary, color:MUTED },
            { label:"Total Units",     val:pcr.total,     color:pcrColor },
            { label:"Primary %",       val:`${pcr.pct}%`, color: pcr.pct >= 80 ? GREEN : pcr.pct >= 75 ? "#D97706" : RED },
          ].map(k => (
            <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.6rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
              <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
              <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
            </div>
          ))}
        </div>

        {/* Tier progress bar */}
        <div style={{ marginBottom:"0.5rem" }}>
          <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:6 }}>
            <span>Tier Progress — {pcrTier.label} ({pcr.total} units)</span>
            {pcrNext ? <span style={{ color:pcrIsClose?"#D97706":MUTED, fontWeight:pcrIsClose?700:400 }}>{pcrToNext} units to {pcrNext.label} ({pcrNext.min})</span> : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>}
          </div>
          <div style={{ height:10, background:BORDER, borderRadius:5, overflow:"hidden" }}>
            <div style={{ height:10, background: pcrIsClose ? "linear-gradient(90deg,#D97706,#FBBF24)" : pcrColor, borderRadius:5, width:`${pcrProgress}%`, transition:"width 0.5s ease" }} />
          </div>
          {/* Tier markers */}
          <div style={{ display:"flex", justifyContent:"space-between", marginTop:4 }}>
            {[...AD_PCR_TIERS].reverse().map(t => (
              <div key={t.label} style={{ textAlign:"center" }}>
                <div style={{ fontSize:"0.58rem", color: pcr.total >= t.min ? (t.color==="#MUTED"?MUTED:t.color) : BORDER, fontWeight: pcr.total >= t.min ? 700 : 400 }}>{t.label}</div>
                <div style={{ fontSize:"0.55rem", color:MUTED }}>{t.min}+</div>
              </div>
            ))}
          </div>
        </div>

        {/* PCR % warning */}
        {pcr.pct < 75 && (
          <div style={{ marginTop:"0.75rem", padding:"0.6rem 0.85rem", background:"#FEF2F2", border:"2px solid #FECACA", borderRadius:6, fontSize:"0.75rem", color:"#DC2626", fontWeight:700 }}>
            🚨 Primary purchase % is {pcr.pct}% — BELOW the 75% minimum required for Toyo rebate payout.
            Customer must shift {Math.ceil((pcr.total * 0.75) - pcr.primary)} more units to primary to qualify.
          </div>
        )}
        {pcr.pct >= 75 && pcr.pct < 80 && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#FFF7ED", border:"1px solid #FED7AA", borderRadius:6, fontSize:"0.72rem", color:"#D97706" }}>
            ⚠ Primary purchase % is {pcr.pct}% — above 75% rebate threshold but below 80%. Watch closely.
          </div>
        )}
        {pcr.pct >= 80 && (
          <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.75rem", background:"#F0FDF4", border:"1px solid #BBF7D0", borderRadius:6, fontSize:"0.72rem", color:"#059669", fontWeight:600 }}>
            ✓ Primary purchase % is {pcr.pct}% — above the 75% rebate threshold. Rebate eligible.
          </div>
        )}
      </div>

      {/* TBR Program */}
      <div style={S.card}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em" }}>TBR Program (Truck / Bus / Radial)</div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:"0.72rem", fontWeight:800, color:tbrColor, background: tbrColor === MUTED ? "#F4F7FB" : tbrColor+"22", padding:"0.2rem 0.75rem", borderRadius:20, border:`2px solid ${tbrColor}` }}>
              {tbrTier.label}
            </span>
            {tbrIsClose && tbr.total > 0 && (
              <span style={{ fontSize:"0.7rem", fontWeight:700, color:"#D97706", background:"#FEF3C7", padding:"0.2rem 0.65rem", borderRadius:20, border:"2px solid #FCD34D" }}>
                ⚡ {tbrToNext} to {tbrNext.label}!
              </span>
            )}
          </div>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"0.6rem", marginBottom:"1rem" }}>
          {[
            { label:"Primary Units",   val:Math.max(0,tbr.primary),   color:AMBER },
            { label:"Secondary Units", val:tbr.secondary, color:MUTED },
            { label:"Total Units",     val:tbr.total,     color:tbrColor },
          ].map(k => (
            <div key={k.label} style={{ background:"#F4F7FB", borderRadius:6, padding:"0.6rem 0.8rem", borderTop:`3px solid ${k.color}` }}>
              <div style={{ fontSize:"0.95rem", fontWeight:700, color:k.color }}>{k.val}</div>
              <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2, textTransform:"uppercase", letterSpacing:"0.06em" }}>{k.label}</div>
            </div>
          ))}
        </div>

        {tbr.total === 0 && (
          <div style={{ fontSize:"0.75rem", color:MUTED, textAlign:"center", padding:"0.5rem", background:"#F4F7FB", borderRadius:6 }}>
            No TBR units purchased — opportunity to introduce truck tire program
          </div>
        )}

        {tbr.total > 0 && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:"0.68rem", color:MUTED, marginBottom:6 }}>
              <span>Tier Progress — {tbrTier.label} ({tbr.total} units)</span>
              {tbrNext ? <span style={{ color:tbrIsClose?"#D97706":MUTED, fontWeight:tbrIsClose?700:400 }}>{tbrToNext} to {tbrNext.label} ({tbrNext.min})</span> : <span style={{ color:"#7C3AED", fontWeight:700 }}>Top Tier ★</span>}
            </div>
            <div style={{ height:10, background:BORDER, borderRadius:5, overflow:"hidden" }}>
              <div style={{ height:10, background:tbrIsClose ? "linear-gradient(90deg,#D97706,#FBBF24)" : tbrColor, borderRadius:5, width:`${tbrProgress}%`, transition:"width 0.5s ease" }} />
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Customer Detail Tab ───────────────────────────────────────────────────────

// ── Todo Item ─────────────────────────────────────────────────────────────────
function TodoItem({ todo, onToggle, onDelete, onEdit }) {
  const [editing, setEditing] = useState(false);
  const [editText, setEditText] = useState(todo.text);
  const inputRef = useRef(null);

  useEffect(() => {
    if (editing && inputRef.current) inputRef.current.focus();
  }, [editing]);

  function commitEdit() {
    if (editText.trim() && editText.trim() !== todo.text) {
      onEdit(editText.trim());
    } else {
      setEditText(todo.text);
    }
    setEditing(false);
  }

  return (
    <div style={{ display:"flex", alignItems:"flex-start", gap:8, padding:"0.45rem 0.6rem",
      background: todo.done ? "#F4F7FB" : "#FFFBEB",
      borderRadius:6,
      border: `1px solid ${todo.done ? BORDER : "#FDE68A"}`,
      opacity: todo.done ? 0.7 : 1,
      transition:"all 0.2s" }}>
      {/* Checkbox */}
      <input type="checkbox" checked={todo.done} onChange={onToggle}
        style={{ marginTop:3, cursor:"pointer", accentColor:AMBER, flexShrink:0 }} />
      {/* Text / edit */}
      <div style={{ flex:1, minWidth:0 }}>
        {editing ? (
          <input
            ref={inputRef}
            value={editText}
            onChange={e => setEditText(e.target.value)}
            onKeyDown={e => { if(e.key==="Enter") commitEdit(); if(e.key==="Escape") { setEditText(todo.text); setEditing(false); } }}
            onBlur={commitEdit}
            style={{ width:"100%", fontSize:"0.75rem", color:TEXT, background:"#FFFFFF",
              border:`1px solid ${AMBER}`, borderRadius:4, padding:"2px 6px", outline:"none", boxSizing:"border-box" }}
          />
        ) : (
          <div
            onDoubleClick={() => { if(!todo.done) { setEditing(true); } }}
            style={{ fontSize:"0.75rem", color:TEXT, textDecoration:todo.done?"line-through":"none",
              cursor:todo.done?"default":"text", wordBreak:"break-word" }}
            title={todo.done ? "" : "Double-click to edit"}>
            {todo.text}
          </div>
        )}
        <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2 }}>{todo.by} · {todo.date}</div>
      </div>
      {/* Actions */}
      <div style={{ display:"flex", gap:4, flexShrink:0, alignItems:"flex-start", marginTop:1 }}>
        {!todo.done && !editing && (
          <button onClick={()=>setEditing(true)}
            style={{ background:"none", border:"none", color:MUTED, cursor:"pointer", fontSize:"0.7rem", padding:"0 2px" }}
            title="Edit">✏</button>
        )}
        <button onClick={onDelete}
          style={{ background:"none", border:"none",
            color: todo.done ? RED : MUTED,
            cursor:"pointer", fontSize:"0.8rem", padding:"0 2px", fontWeight:todo.done?700:400 }}
          title="Delete">×</button>
      </div>
    </div>
  );
}

function CustomerDetailTab({ ap, customers, ar, weekComp, onClose, inactiveRecord, onMarkInactive, onMarkActive, currentUser, onLogActivity }) {
  const [subTab, setSubTab] = useState("overview");
  const [aiLoading, setAiLoading] = useState(false);
  const [callSummary, setCallSummary] = useState(null);
  const [showInactiveForm, setShowInactiveForm] = useState(false);
  const [inactiveReason, setInactiveReason] = useState("");
  const isInactive = !!inactiveRecord;

  // Persistent notes for this customer
  const notesKey = `notes_${ap.custNum}`;
  const [notes, setNotes] = useState(() => {
    try { return localStorage.getItem(notesKey) || ""; } catch { return ""; }
  });
  function saveNotes(val) {
    setNotes(val);
    try { localStorage.setItem(notesKey, val); } catch {}
    // Sync to Supabase (debounced — fire after 1.5s pause)
    if (window._notesSyncTimer) clearTimeout(window._notesSyncTimer);
    window._notesSyncTimer = setTimeout(() => {
      if (currentUser) {
        syncNotesUp(currentUser.id, ap.custNum, val);
        if (onLogActivity) onLogActivity("update_notes", `${ap.customer}`);
      }
    }, 1500);
  }

  // Load notes from Supabase on mount (overrides localStorage if newer)
  useEffect(() => {
    if (!currentUser) return;
    syncNotesDown(currentUser.id, ap.custNum).then(row => {
      if (row?.notes && row.notes !== notes) {
        setNotes(row.notes);
        try { localStorage.setItem(notesKey, row.notes); } catch {}
      }
    });
  }, [ap.custNum]);

  // To-Do items for this customer
  const todosKey = `todos_${ap.custNum}`;
  const [todos, setTodos] = useState(() => {
    try { return JSON.parse(localStorage.getItem(todosKey) || "[]"); } catch { return []; }
  });
  const [newTodoText, setNewTodoText] = useState("");
  function saveTodos(updated) {
    setTodos(updated);
    try { localStorage.setItem(todosKey, JSON.stringify(updated)); } catch {}
    // Sync to Supabase
    if (currentUser) {
      syncTodosUp(currentUser.id, ap.custNum, ap.customer, ap.city, ap.salesman, updated);
    }
  }

  // Load todos from Supabase on mount
  useEffect(() => {
    if (!currentUser) return;
    syncTodosDown(currentUser.id, ap.custNum).then(rows => {
      if (rows && rows.length > 0) {
        setTodos(rows);
        try { localStorage.setItem(todosKey, JSON.stringify(rows)); } catch {}
      }
    });
  }, [ap.custNum]);
  function addTodo() {
    if (!newTodoText.trim()) return;
    const updated = [{ id: Date.now(), text: newTodoText.trim(), done: false,
      date: new Date().toISOString().slice(0,10), by: currentUser?.name || ap.salesman }, ...todos];
    saveTodos(updated);
    setNewTodoText("");
    if (onLogActivity) onLogActivity("add_todo", `${ap.customer} — ${newTodoText.trim()}`);
  }
  function toggleTodo(id) {
    saveTodos(todos.map(t => t.id===id ? {...t, done:!t.done} : t));
  }
  function deleteTodo(id) {
    saveTodos(todos.filter(t => t.id!==id));
  }

  // Look up full customer record
  const cust = (customers || []).find(c => c.num === ap.custNum) || {};
  const arRecord = (ar || []).find(r => r.custNum === ap.custNum) || null;
  const repColor = REP_COLORS[ap.salesman] || AMBER;

  // Parse department focus details
  const focusLines = (ap.focus || "").split("|").map(f => f.trim()).filter(Boolean);
  const deptItems = focusLines.map(line => {
    const isDown = line.startsWith("DOWN:") || line.startsWith("LOST:");
    const isLow = line.startsWith("LOW GP:");
    const color = isDown ? RED : isLow ? "#D97706" : GREEN;
    return { line, color };
  });

  // GPS direction links
  const hasCoords = cust.lat && cust.lon;
  const gpsLinks = hasCoords ? [
    { label: "Google Maps", icon: "🗺", url: `https://www.google.com/maps/dir/?api=1&destination=${cust.lat},${cust.lon}` },
    { label: "Waze", icon: "🚗", url: `https://waze.com/ul?ll=${cust.lat},${cust.lon}&navigate=yes` },
    { label: "Apple Maps", icon: "🍎", url: `http://maps.apple.com/?daddr=${cust.lat},${cust.lon}` },
  ] : [];

  // Generate AI call summary
  async function generateCallSummary() {
    setAiLoading(true);
    setSubTab("calls");
    const arInfo = arRecord ? `AR Balance: $${arRecord.balance.toFixed(0)}, Current: $${arRecord.curDue.toFixed(0)}, 90+ days: $${arRecord.dueOver90.toFixed(0)}, Last paid: ${arRecord.lastPaid || "unknown"}` : "No AR data";
    const prompt = `You are a tire & ag supply sales assistant helping a rep prepare for a customer call.

Customer: ${ap.customer}
City: ${ap.city}
Rep: ${ap.salesman}
2025 Sales: $${ap.sales2025.toFixed(0)}
2026 Sales: $${ap.sales2026.toFixed(0)}
Change: $${ap.change.toFixed(0)} (${ap.gpPct ? (ap.gpPct*100).toFixed(1)+"% GP" : "N/A"})
Status: ${ap.action}
Top Department 2026: ${ap.topDept}
Most Declined Dept: ${ap.declinedDept || "None"}
Department Focus: ${ap.focus}
${arInfo}

Generate a NEXT VISIT PLAN for this account. Return JSON only: { "callPlan": "2-3 actionable sentences covering: what to discuss, which products/programs to push based on their top and declining departments, and any AR or relationship notes worth addressing." }`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      if (!res.ok) throw new Error(`API error ${res.status}: ${await res.text()}`);
      const data = await res.json();
      const raw = data.content?.[0]?.text || "{}";
      const clean = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      // normalise — handle both {callPlan:...} and plain string responses
      setCallSummary({ callPlan: parsed.callPlan || parsed.plan || String(parsed) });
    } catch(e) {
      setCallSummary({ error: `Error: ${e.message || 'Could not connect to AI. Check API key.'}` });
    }
    setAiLoading(false);
  }

  // Auto-generate on first "calls" tab open
  const planGenerated = useRef(false);
  useEffect(() => {
    // Auto-generate on first overview open
    if (subTab === "overview" && !callSummary && !aiLoading && !planGenerated.current) {
      planGenerated.current = true;
      generateCallSummary();
    }
  }, [subTab]);

  const adProgram = AD_PROGRAMS[String(ap.custNum)] || null;
  const ascensoProgram  = ASCENSO_PROGRAMS[String(ap.custNum)]  || null;
  const americusProgram = AMERICUS_PROGRAMS[String(ap.custNum)] || null;
  const barnnProgram     = BARNN_PROGRAMS[String(ap.custNum)]     || null;
  const yokoProgram      = YOKOHAMA_PROGRAMS[String(ap.custNum)]    || null;
  const falkenPltProgram = FALKEN_PLT_PROGRAMS[String(ap.custNum)] || null;
  const falkenTbrProgram = FALKEN_TBR_PROGRAMS[String(ap.custNum)] || null;
  const realCallLog = SEED_CALL_LOG[String(ap.custNum)] || null;
  const hasRealCalls = realCallLog && realCallLog.length > 0;

  const subtabs = [
    { id: "overview", label: "📊 Overview" },
    { id: "depts",    label: "📦 Departments" },
    { id: "ar",       label: "$ AR" },
    { id: "calls",    label: "📞 Call Log" },
    ...(adProgram ? [{ id: "ad", label: "🏆 AD Programs" }] : []),
    ...(ascensoProgram  ? [{ id: "ascenso",  label: "🌀 Ascenso"  }] : []),
    ...(americusProgram ? [{ id: "americus", label: "🔵 Americus" }] : []),
    ...(barnnProgram      ? [{ id: "barnn",    label: "🔶 BF BARNN"   }] : []),
    ...(yokoProgram       ? [{ id: "yoko",     label: "🔵 Yokohama"  }] : []),
    ...(falkenPltProgram  ? [{ id: "falkenplt", label: "🟢 Falken PLT" }] : []),
    ...(falkenTbrProgram  ? [{ id: "falkentbr", label: "🟢 Falken TBR" }] : []),
  ];

  return (
    <div>
      {/* Customer header */}
      <div style={{ ...S.card, borderLeft: `4px solid ${repColor}`, marginBottom: "0.75rem" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontSize: "1.05rem", fontWeight: 700, color: TEXT }}>{ap.customer}</div>
            <div style={{ fontSize: "0.78rem", color: MUTED, marginTop: 3 }}>
              📍 {cust.address ? `${cust.address}, ` : ""}{ap.city}{cust.state ? `, ${cust.state}` : ""}
              {cust.zip ? ` ${cust.zip}` : ""}
            </div>
            {cust.phone && <div style={{ fontSize: "0.78rem", color: MUTED, marginTop: 2 }}>📞 <a href={`tel:${cust.phone}`} style={{ color: AMBER, textDecoration: "none" }}>{cust.phone}</a></div>}
            <div style={{ fontSize: "0.72rem", color: repColor, fontWeight: 600, marginTop: 4 }}>Rep: {ap.salesman}</div>
          </div>
          {/* GPS Directions */}
          <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end" }}>
            {hasCoords ? (
              <>
                <div style={{ fontSize: "0.68rem", color: MUTED, marginBottom: 2 }}>Directions:</div>
                <div style={{ display: "flex", gap: 6 }}>
                  {gpsLinks.map(g => (
                    <a key={g.label} href={g.url} target="_blank" rel="noopener noreferrer"
                      style={{ display: "flex", alignItems: "center", gap: 4, padding: "0.35rem 0.75rem", background: "#EEF4FF", border: `1px solid ${BORDER}`, borderRadius: 6, fontSize: "0.72rem", color: AMBER, fontWeight: 600, textDecoration: "none", whiteSpace: "nowrap" }}>
                      {g.icon} {g.label}
                    </a>
                  ))}
                </div>
              </>
            ) : (
              <div style={{ fontSize: "0.68rem", color: MUTED }}>No coordinates on file</div>
            )}
            <button onClick={onClose} style={{ background: "none", border: `1px solid ${BORDER}`, color: MUTED, borderRadius: 4, padding: "0.25rem 0.6rem", cursor: "pointer", fontSize: "0.7rem", marginTop: 2 }}>× Close Tab</button>
            {!isInactive
              ? <button onClick={() => setShowInactiveForm(true)} style={{ background: "#FEF2F2", border: "1px solid #FECACA", color: RED, borderRadius: 4, padding: "0.25rem 0.6rem", cursor: "pointer", fontSize: "0.7rem", marginTop: 2 }}>⊘ Mark Inactive</button>
              : <button onClick={() => { if(window.confirm(`Reactivate ${ap.customer}?`)) onMarkActive(); }} style={{ background: "#F0FDF4", border: "1px solid #BBF7D0", color: GREEN, borderRadius: 4, padding: "0.25rem 0.6rem", cursor: "pointer", fontSize: "0.7rem", marginTop: 2 }}>✓ Reactivate</button>
            }
          </div>
        </div>

        {/* Quick KPIs */}
        {adProgram && (() => {
          const pcrTier = getAdTier(adProgram.pcr.total, AD_PCR_TIERS);
          const nextTier = getNextAdTier(adProgram.pcr.total, AD_PCR_TIERS);
          const toNext = nextTier ? nextTier.min - adProgram.pcr.total : 0;
          const isClose = nextTier && toNext <= Math.max(20, Math.round(nextTier.min * 0.15));
          return (
            <div style={{ marginTop:"0.75rem", padding:"0.5rem 0.85rem", background: isClose ? "#FEF3C7" : "#EEF4FF", border:`1px solid ${isClose?"#FCD34D":BORDER}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color: pcrTier.color === "#MUTED" ? MUTED : pcrTier.color }}>
                🏆 {adProgram.program} — {pcrTier.label}
              </span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>PCR/LTR: {adProgram.pcr.total} units</span>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color: adProgram.pcr.pct < 75 ? "#DC2626" : adProgram.pcr.pct < 80 ? "#D97706" : GREEN }}>
                {adProgram.pcr.pct}% primary {adProgram.pcr.pct < 75 ? "⚠ BELOW 75%" : ""}
              </span>
              {adProgram.tbr.total > 0 && <span style={{ fontSize:"0.68rem", color:MUTED }}>TBR: {adProgram.tbr.total} units</span>}
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#D97706" }}>⚡ {toNext} units to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("ad")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:AMBER, background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {ascensoProgram && (() => {
          const totalAmt = ASCENSO_TOTAL;
          const tier = getAdTier(totalAmt, ASCENSO_TIERS);
          const nextTier = getNextAdTier(totalAmt, ASCENSO_TIERS);
          const toNext = nextTier ? nextTier.min - totalAmt : 0;
          const isClose = nextTier && toNext <= nextTier.min * 0.15;
          const tColor = tier.color;
          const payout = totalAmt * (tier.payout || 0);
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background: isClose ? "#FEF3C7" : "#F0FDF4", border:`1px solid ${isClose?"#FCD34D":"#BBF7D0"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:tColor }}>🌀 Ascenso — Company {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>Co. Total: {fmt(totalAmt)} · This acct: {fmt(ascensoProgram.amount)}</span>
              {tier.payout > 0 && <span style={{ fontSize:"0.68rem", fontWeight:700, color:GREEN }}>Est. Payout: {fmt(payout)} ({(tier.payout*100).toFixed(0)}%)</span>}
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#D97706" }}>⚡ {fmt(toNext)} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("ascenso")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:"#059669", background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {americusProgram && (() => {
          const ap2 = americusProgram;
          const tier = getAdTier(ap2.ytd, AMERICUS_TIERS);
          const nextTier = getNextAdTier(ap2.ytd, AMERICUS_TIERS);
          const toNext = nextTier ? nextTier.min - ap2.ytd : 0;
          const gap = nextTier ? nextTier.min - tier.min : 0;
          const isClose = nextTier && toNext <= Math.max(10, Math.round(gap * 0.15));
          const tColor = tier.color;
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background: isClose?"#EDE9FE":"#EFF6FF", border:`1px solid ${isClose?"#C4B5FD":"#BFDBFE"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:tColor }}>🔵 Americus Partners — {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>YTD: {ap2.ytd} units · Q1: {ap2.q1} · Q2: {ap2.q2}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>2025: {ap2.units2025} units</span>
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#7C3AED" }}>⚡ {toNext} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("americus")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:"#0891B2", background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {yokoProgram && (() => {
          const yp = yokoProgram;
          const tier = getAdTier(yp.ytd, YOKOHAMA_TIERS);
          const nextTier = getNextAdTier(yp.ytd, YOKOHAMA_TIERS);
          const toNext = nextTier ? nextTier.min - yp.ytd : 0;
          const isClose = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
          const tColor = tier.color === "#6B7A99" ? MUTED : tier.color;
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background:isClose?"#EDE9FE":"#EFF6FF", border:`1px solid ${isClose?"#C4B5FD":"#BFDBFE"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:tColor }}>🔵 Yokohama — {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>QTD: {yp.ytd} units · Primary: {yp.primary} ({yp.priPct}%)</span>
              {yp.priPct < 75 && <span style={{ fontSize:"0.68rem", fontWeight:700, color:RED }}>⚠ Below 75%</span>}
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#7C3AED" }}>⚡ {toNext} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("yoko")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:tColor, background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {barnnProgram && (() => {
          const b = barnnProgram;
          const bTotal = b.total||b.ytd||0;
          const tier = getAdTier(bTotal, BARNN_TIERS);
          const nextTier = getNextAdTier(bTotal, BARNN_TIERS);
          const toNext = nextTier ? nextTier.min - bTotal : 0;
          const isClose = nextTier && toNext <= Math.max(10, Math.round(nextTier.min * 0.15));
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background:isClose?"#FFF7ED":"#FFFBEB", border:`1px solid ${isClose?"#FCD34D":"#FDE68A"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#D97706" }}>🔶 BF BARNN — {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>YTD: {b.ytd} units · PY: {b.py} · {b.asOf}</span>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:clr(b.change) }}>{b.change>0?"+":""}{b.change} ({b.changePct>0?"+":""}{(b.changePct*100).toFixed(0)}%)</span>
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#D97706" }}>⚡ {toNext} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("barnn")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:"#D97706", background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {falkenTbrProgram && (() => {
          const fp = falkenTbrProgram;
          const tier = getAdTier(fp.ytd, FALKEN_TBR_TIERS);
          const nextTier = getNextAdTier(fp.ytd, FALKEN_TBR_TIERS);
          const toNext = nextTier ? nextTier.min - fp.ytd : 0;
          const isClose = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
          const tColor = tier.color === "#6B7A99" ? MUTED : tier.color;
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background:isClose?"#F0FDF4":"#ECFDF5", border:`1px solid ${isClose?"#86EFAC":"#A7F3D0"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#059669" }}>🟢 Falken Fanatic TBR — {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>Q2: {fp.q2} units (Q1: {fp.q1})</span>
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#059669" }}>⚡ {toNext} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("falkentbr")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:"#059669", background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        {falkenPltProgram && (() => {
          const fp = falkenPltProgram;
          const tier = getAdTier(fp.ytd, FALKEN_PLT_TIERS);
          const nextTier = getNextAdTier(fp.ytd, FALKEN_PLT_TIERS);
          const toNext = nextTier ? nextTier.min - fp.ytd : 0;
          const isClose = nextTier && toNext <= Math.max(5, Math.round(nextTier.min * 0.15));
          const tColor = tier.color === "#6B7A99" ? MUTED : tier.color;
          return (
            <div style={{ marginTop:"0.5rem", padding:"0.5rem 0.85rem", background:isClose?"#F0FDF4":"#ECFDF5", border:`1px solid ${isClose?"#86EFAC":"#A7F3D0"}`, borderRadius:8, display:"flex", alignItems:"center", gap:12, flexWrap:"wrap" }}>
              <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#059669" }}>🟢 Falken Fanatic PLT — {tier.label}</span>
              <span style={{ fontSize:"0.68rem", color:MUTED }}>Q1: {fp.q1} · Q2: {fp.q2} · YTD: {fp.ytd}</span>
              {isClose && <span style={{ fontSize:"0.68rem", fontWeight:700, color:"#059669" }}>⚡ {toNext} to {nextTier.label}!</span>}
              <button onClick={()=>setSubTab("falkenplt")} style={{ marginLeft:"auto", fontSize:"0.65rem", color:"#059669", background:"none", border:`1px solid ${BORDER}`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>View Details →</button>
            </div>
          );
        })()}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: "0.6rem", marginTop: "1rem" }}>
          {[
            { label: "2026 Sales", val: fmt(ap.sales2026), color: AMBER },
            { label: "vs 2025", val: fmt(ap.change), color: clr(ap.change) },
            { label: "GP %", val: pct(ap.gpPct), color: ap.gpPct < 0.08 ? RED : GREEN },
            { label: "AR Balance", val: arRecord ? fmt(arRecord.balance) : "—", color: arRecord?.dueOver90 > 0 ? RED : TEXT },
          ].map(k => (
            <div key={k.label} style={{ background: "#F4F7FB", borderRadius: 6, padding: "0.6rem 0.8rem", borderTop: `3px solid ${k.color}` }}>
              <div style={{ fontSize: "0.95rem", fontWeight: 700, color: k.color }}>{k.val}</div>
              <div style={{ fontSize: "0.65rem", color: MUTED, marginTop: 2, textTransform: "uppercase", letterSpacing: "0.08em" }}>{k.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sub tabs */}
      <div style={S.subNav}>
        {subtabs.map(s => <button key={s.id} style={S.subBtn(subTab===s.id, AMBER)} onClick={() => setSubTab(s.id)}>{s.label}</button>)}
      </div>

      {/* Overview */}
      {subTab === "overview" && (
        <div style={S.card}>
          <div style={{ marginBottom: "1rem" }}>
            <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.5rem" }}>Account Status</div>
            <div style={{ display: "inline-block", padding: "0.3rem 0.8rem", borderRadius: 20, background: ap.action.toUpperCase().includes("GROW") ? "#D1FAE5" : ap.action.toUpperCase().includes("LOST") || ap.action.toUpperCase().includes("DECLIN") ? "#FEE2E2" : "#EFF6FF", color: ap.action.toUpperCase().includes("GROW") ? GREEN : ap.action.toUpperCase().includes("LOST") || ap.action.toUpperCase().includes("DECLIN") ? RED : AMBER, fontWeight: 700, fontSize: "0.78rem" }}>
              {ap.action}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
            <div>
              <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.5rem" }}>Top Department 2026</div>
              <div style={{ fontSize: "0.85rem", fontWeight: 600, color: GREEN }}>{ap.topDept || "—"}</div>
            </div>
            <div>
              <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.5rem" }}>Most Declined</div>
              <div style={{ fontSize: "0.85rem", fontWeight: 600, color: RED }}>{ap.declinedDept || "—"}</div>
            </div>
          </div>
          {focusLines.length > 0 && (
            <div style={{ marginTop: "1rem" }}>
              <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.5rem" }}>Focus Areas</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {deptItems.map((d, i) => (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "0.4rem 0.75rem", background: "#F4F7FB", borderRadius: 6, borderLeft: `3px solid ${d.color}` }}>
                    <span style={{ fontSize: "0.75rem", color: d.color, fontWeight: 600 }}>{d.line}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Next Visit Plan on Overview */}
          <div style={{ marginTop:"1rem", paddingTop:"1rem", borderTop:`1px solid ${BORDER}` }}>
            <div style={{ fontSize:"0.7rem", color:GREEN, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.5rem" }}>🎯 Next Visit Plan</div>
            {callSummary && !callSummary.error
              ? <div style={{ fontSize:"0.8rem", color:TEXT, lineHeight:1.8 }}>{callSummary.callPlan}
                  <button onClick={generateCallSummary} style={{ ...S.btn(MUTED), fontSize:"0.62rem", marginLeft:10 }}>↺</button>
                </div>
              : aiLoading
                ? <div style={{ fontSize:"0.75rem", color:MUTED }}>◈ Generating…</div>
                : <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                    <button onClick={generateCallSummary} style={{ ...S.btn(GREEN), fontSize:"0.7rem" }}>◈ Generate Plan</button>
                    <span style={{ fontSize:"0.68rem", color:MUTED }}>or go to 📞 Calls tab</span>
                  </div>
            }
          </div>
        </div>
      )}

      {/* Departments */}
      {subTab === "depts" && (
        <div style={S.card}>
          <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.75rem" }}>Department Analysis</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem", marginBottom: "1rem" }}>
            <div style={{ background: "#F0FDF4", border: `1px solid #BBF7D0`, borderRadius: 8, padding: "0.75rem 1rem" }}>
              <div style={{ fontSize: "0.68rem", color: GREEN, fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>▲ Top Performer</div>
              <div style={{ fontSize: "0.85rem", fontWeight: 600, color: TEXT }}>{ap.topDept || "—"}</div>
            </div>
            <div style={{ background: "#FFF1F2", border: `1px solid #FECDD3`, borderRadius: 8, padding: "0.75rem 1rem" }}>
              <div style={{ fontSize: "0.68rem", color: RED, fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>▼ Most Declined</div>
              <div style={{ fontSize: "0.85rem", fontWeight: 600, color: TEXT }}>{ap.declinedDept || "None"}</div>
            </div>
          </div>
          <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.5rem" }}>Action Items by Department</div>
          {deptItems.length > 0 ? deptItems.map((d, i) => (
            <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "0.6rem 0.75rem", background: "#F4F7FB", borderRadius: 6, borderLeft: `3px solid ${d.color}`, marginBottom: 6 }}>
              <span style={{ fontSize: "0.9rem" }}>{d.color === RED ? "⬇" : d.color === "#D97706" ? "⚠" : "✓"}</span>
              <span style={{ fontSize: "0.78rem", color: TEXT, lineHeight: 1.5 }}>{d.line}</span>
            </div>
          )) : <div style={{ color: MUTED, fontSize: "0.75rem" }}>No focus items — account is on track.</div>}
        </div>
      )}

      {/* AR */}
      {subTab === "ar" && (
        <div style={S.card}>
          <div style={{ fontSize: "0.7rem", color: MUTED, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "0.75rem" }}>Accounts Receivable — {ap.customer}</div>
          {arRecord ? (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "0.6rem", marginBottom: "1rem" }}>
                {[
                  { label: "Total Balance", val: fmt(arRecord.balance), color: AMBER },
                  { label: "Current", val: fmt(arRecord.curDue), color: GREEN },
                  { label: "1–30 Days", val: arRecord.due1_30 > 0 ? fmt(arRecord.due1_30) : "—", color: arRecord.due1_30 > 0 ? "#D97706" : MUTED },
                  { label: "31–60 Days", val: arRecord.due31_60 > 0 ? fmt(arRecord.due31_60) : "—", color: arRecord.due31_60 > 0 ? "#D97706" : MUTED },
                  { label: "61–90 Days", val: arRecord.due61_90 > 0 ? fmt(arRecord.due61_90) : "—", color: arRecord.due61_90 > 0 ? RED : MUTED },
                  { label: "90+ Days", val: arRecord.dueOver90 > 0 ? fmt(arRecord.dueOver90) : "—", color: arRecord.dueOver90 > 0 ? RED : MUTED },
                ].map(k => (
                  <div key={k.label} style={{ background: "#F4F7FB", borderRadius: 6, padding: "0.6rem 0.8rem" }}>
                    <div style={{ fontSize: "0.9rem", fontWeight: 700, color: k.color }}>{k.val}</div>
                    <div style={{ fontSize: "0.65rem", color: MUTED, marginTop: 2 }}>{k.label}</div>
                  </div>
                ))}
              </div>
              {arRecord.lastPaid && <div style={{ fontSize: "0.75rem", color: MUTED }}>Last payment received: <span style={{ color: TEXT, fontWeight: 600 }}>{arRecord.lastPaid}</span></div>}
              {arRecord.dueOver90 > 0 && (
                <div style={{ marginTop: "0.75rem", padding: "0.6rem 0.85rem", background: "#FFF1F2", border: `1px solid #FECDD3`, borderRadius: 6, fontSize: "0.75rem", color: RED, fontWeight: 600 }}>
                  ⚠ {fmt(arRecord.dueOver90)} is 90+ days past due — discuss collections on this visit.
                </div>
              )}
            </>
          ) : (
            <div style={{ color: MUTED, fontSize: "0.78rem" }}>No AR record found for this customer.</div>
          )}
        </div>
      )}

      {/* Calls */}
      {/* To-Dos + Notes — shown on calls tab */}
      {subTab === "calls" && (
        <>
          {/* To-Do items */}
          <div style={{ ...S.card, marginTop:"0.85rem" }}>
            <div style={{ fontSize:"0.7rem", fontWeight:700, color:AMBER, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.75rem" }}>
              ✅ To-Do
              {todos.filter(t=>!t.done).length > 0 && <span style={{ marginLeft:6, background:RED, color:"#fff", borderRadius:8, padding:"1px 7px", fontSize:"0.62rem" }}>{todos.filter(t=>!t.done).length}</span>}
            </div>
            {/* Add new */}
            <div style={{ display:"flex", gap:6, marginBottom:"0.75rem" }}>
              <input
                value={newTodoText}
                onChange={e => setNewTodoText(e.target.value)}
                onKeyDown={e => e.key==="Enter" && addTodo()}
                placeholder={`Add a to-do for ${ap.customer}...`}
                style={{ flex:1, background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.4rem 0.7rem", borderRadius:6, fontSize:"0.75rem", outline:"none" }}
                onFocus={e => e.target.style.borderColor=AMBER}
                onBlur={e => e.target.style.borderColor=BORDER}
              />
              <button onClick={addTodo} style={{ ...S.btn(AMBER), fontSize:"0.72rem", padding:"0.4rem 0.85rem" }}>+ Add</button>
            </div>
            {/* List */}
            {todos.length === 0 && <div style={{ fontSize:"0.72rem", color:MUTED, textAlign:"center", padding:"0.5rem" }}>No to-dos yet</div>}
            {todos.filter(t=>t.done).length > 0 && (
              <div style={{ display:"flex", justifyContent:"flex-end", marginBottom:4 }}>
                <button onClick={()=>saveTodos(todos.filter(t=>!t.done))}
                  style={{ fontSize:"0.65rem", color:MUTED, background:"#FEF2F2", border:`1px solid #FECACA`, borderRadius:4, padding:"2px 8px", cursor:"pointer" }}>
                  × Clear {todos.filter(t=>t.done).length} completed
                </button>
              </div>
            )}
            <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
              {todos.map(t => (
                <TodoItem key={t.id} todo={t}
                  onToggle={()=>toggleTodo(t.id)}
                  onDelete={()=>deleteTodo(t.id)}
                  onEdit={(newText)=>{
                    saveTodos(todos.map(x => x.id===t.id ? {...x, text:newText} : x));
                  }}
                />
              ))}
            </div>
          </div>

          {/* Rep Notes */}
          <div style={{ ...S.card, marginTop:"0.75rem" }}>
            <div style={{ fontSize:"0.7rem", fontWeight:700, color:MUTED, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.6rem" }}>
              📝 Rep Notes <span style={{ fontSize:"0.62rem", fontWeight:400, color:MUTED }}>(saved automatically)</span>
            </div>
            <textarea
              value={notes}
              onChange={e => saveNotes(e.target.value)}
              placeholder={`Notes for ${ap.customer}... (pricing, relationship notes, follow-ups)`}
              rows={4}
              style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT, padding:"0.6rem 0.75rem", borderRadius:6, fontSize:"0.78rem", resize:"vertical", boxSizing:"border-box", lineHeight:1.7, outline:"none" }}
              onFocus={e => e.target.style.borderColor=AMBER}
              onBlur={e => e.target.style.borderColor=BORDER}
            />
            {notes && <div style={{ fontSize:"0.62rem", color:GREEN, marginTop:4 }}>✓ Saved to this device</div>}
          </div>
        </>
      )}

      {subTab === "ad" && adProgram && (
        <ADProgramTab adProgram={adProgram} custName={ap.customer} />
      )}
      {subTab === "ascenso" && ascensoProgram && (
        <AscensoTab ascenso={ascensoProgram} custName={ap.customer} />
      )}
      {subTab === "americus" && americusProgram && (
        <AmericusTab americus={americusProgram} custName={ap.customer} />
      )}
      {subTab === "barnn" && barnnProgram && (
        <BARNNTab barnn={barnnProgram} custName={ap.customer} />
      )}
      {subTab === "yoko" && yokoProgram && (
        <YokohamaTab yoko={yokoProgram} custName={ap.customer} />
      )}
      {subTab === "falkenplt" && falkenPltProgram && (
        <FalkenPLTTab falken={falkenPltProgram} custName={ap.customer} />
      )}
      {subTab === "falkentbr" && falkenTbrProgram && (
        <FalkenTBRTab falken={falkenTbrProgram} custName={ap.customer} />
      )}

            {subTab === "calls" && (
        <div>
          {/* Next Visit Plan */}
          {!aiLoading && !callSummary && (
            <div style={{ ...S.card, background:"#F0FDF4", borderLeft:`4px solid ${GREEN}`, marginBottom:"0.75rem" }}>
              <div style={{ fontSize:"0.7rem", color:GREEN, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:6 }}>🎯 Next Visit Plan</div>
              <div style={{ display:"flex", alignItems:"center", gap:12 }}>
                <button onClick={generateCallSummary} style={{ ...S.btn(GREEN), fontSize:"0.72rem" }}>◈ Generate AI Plan</button>
                <span style={{ fontSize:"0.7rem", color:MUTED }}>AI analyses account data, sales trends & AR to build a visit plan</span>
              </div>
            </div>
          )}
          {aiLoading && (
            <div style={{ ...S.card, background:"#F0FDF4", borderLeft:`4px solid ${GREEN}`, marginBottom:"0.75rem", textAlign:"center", color:MUTED, fontSize:"0.75rem", padding:"1.25rem" }}>
              ◈ Generating visit plan…
            </div>
          )}
          {callSummary && !callSummary.error && (
            <div style={{ ...S.card, background:"#F0FDF4", borderLeft:`4px solid ${GREEN}`, marginBottom:"0.75rem" }}>
              <div style={{ fontSize:"0.7rem", color:GREEN, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:"0.5rem" }}>🎯 Next Visit Plan</div>
              <div style={{ fontSize:"0.8rem", color:TEXT, lineHeight:1.8 }}>{callSummary.callPlan}</div>
              <button onClick={generateCallSummary} style={{ ...S.btn(MUTED), fontSize:"0.65rem", marginTop:"0.6rem" }}>↺ Regenerate</button>
            </div>
          )}
          {callSummary?.error && <div style={{ ...S.card, color:RED, marginBottom:"0.75rem" }}>{callSummary.error}</div>}
        </div>
      )}
    </div>
  );
}

