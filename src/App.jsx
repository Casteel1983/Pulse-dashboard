import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Cell, AreaChart, Area, PieChart, Pie, Legend
} from "recharts";
import * as Recharts from "recharts";

// SEED_CUSTOMERS loaded asynchronously to keep App.jsx small
let _customersLoaded = false;
async function ensureCustomers(setFn) {
  if (_customersLoaded) return;
  try {
    const res  = await fetch('./customers.json');
    const data = await res.json();
    _customersLoaded = true;
    if (setFn) setFn(prev => ({ ...prev, customers: data }));
  } catch { /* fallback to empty */ }
}


const SEED_WEEK_COMP = {"weeks":[{"week":1,"sales2025":123427.64,"sales2026":45856.73,"change":-77570.91,"changePct":-0.628473,"gp2025":18015.72,"gp2026":7458.9,"gpChange":-10556.82,"locations":{"Byron":{"sales2025":163452.32,"sales2026":71316.71,"gp2025":21613.49,"gp2026":9927.56},"Tifton":{"sales2025":124055.64,"sales2026":45985.73,"gp2025":17983.88,"gp2026":7437.01},"Statesboro":{"sales2025":115720.85,"sales2026":46755.93,"gp2025":16574.42,"gp2026":5754.92},"Athens":{"sales2025":115170.69,"sales2026":73570.06,"gp2025":17369.69,"gp2026":11388.53}}},{"week":2,"sales2025":301822.72,"sales2026":382007.38,"change":80184.66,"changePct":0.265668,"gp2025":46329.9,"gp2026":54276.6,"gpChange":7946.7,"locations":{"Byron":{"sales2025":377044.43,"sales2026":570752.41,"gp2025":52249.02,"gp2026":76862.73},"Tifton":{"sales2025":303298.81,"sales2026":383463.71,"gp2025":46284.22,"gp2026":54183.13},"Statesboro":{"sales2025":462386.87,"sales2026":481873.11,"gp2025":60271.41,"gp2026":59412.46},"Athens":{"sales2025":273827.42,"sales2026":383818.48,"gp2025":39103.51,"gp2026":55728.0}}},{"week":3,"sales2025":245152.08,"sales2026":374483.37,"change":129331.29,"changePct":0.527555,"gp2025":38402.28,"gp2026":61237.1,"gpChange":22834.82,"locations":{"Byron":{"sales2025":425593.39,"sales2026":426682.0,"gp2025":58172.78,"gp2026":67610.08},"Tifton":{"sales2025":246446.08,"sales2026":376015.37,"gp2025":38342.58,"gp2026":61165.46},"Statesboro":{"sales2025":427419.41,"sales2026":404745.78,"gp2025":61195.15,"gp2026":48898.06},"Athens":{"sales2025":349566.16,"sales2026":342498.08,"gp2025":49966.03,"gp2026":47315.53}}},{"week":4,"sales2025":161565.13,"sales2026":350548.24,"change":188983.11,"changePct":1.169702,"gp2025":27003.95,"gp2026":50179.59,"gpChange":23175.64,"locations":{"Byron":{"sales2025":247061.2,"sales2026":522986.13,"gp2025":34166.35,"gp2026":70060.96},"Tifton":{"sales2025":162354.13,"sales2026":351882.3,"gp2025":26948.23,"gp2026":50151.79},"Statesboro":{"sales2025":250714.52,"sales2026":380511.43,"gp2025":34832.31,"gp2026":48248.13},"Athens":{"sales2025":268575.17,"sales2026":361138.66,"gp2025":37126.24,"gp2026":48640.81}}},{"week":5,"sales2025":261017.26,"sales2026":309344.09,"change":48326.83,"changePct":0.185148,"gp2025":39303.55,"gp2026":43783.31,"gpChange":4479.76,"locations":{"Byron":{"sales2025":320762.32,"sales2026":475548.29,"gp2025":41937.18,"gp2026":60561.96},"Tifton":{"sales2025":262329.26,"sales2026":308672.3,"gp2025":39271.71,"gp2026":43662.98},"Statesboro":{"sales2025":405300.67,"sales2026":415940.26,"gp2025":53454.33,"gp2026":49150.08},"Athens":{"sales2025":254913.44,"sales2026":264278.07,"gp2025":35273.09,"gp2026":33378.33}}},{"week":6,"sales2025":348604.74,"sales2026":348507.15,"change":-97.59,"changePct":-0.00028,"gp2025":50704.85,"gp2026":45700.15,"gpChange":-5004.7,"locations":{"Byron":{"sales2025":474252.62,"sales2026":476076.79,"gp2025":56572.2,"gp2026":61164.86},"Tifton":{"sales2025":350518.74,"sales2026":349455.92,"gp2025":50637.19,"gp2026":45500.63},"Statesboro":{"sales2025":493114.38,"sales2026":370412.24,"gp2025":61443.78,"gp2026":44444.13},"Athens":{"sales2025":354158.58,"sales2026":420047.08,"gp2025":45660.65,"gp2026":49245.91}}},{"week":7,"sales2025":306406.92,"sales2026":316725.6,"change":10318.68,"changePct":0.033676,"gp2025":43977.04,"gp2026":42577.08,"gpChange":-1399.96,"locations":{"Byron":{"sales2025":420629.79,"sales2026":465205.62,"gp2025":53125.83,"gp2026":57900.03},"Tifton":{"sales2025":308024.92,"sales2026":318158.47,"gp2025":43919.33,"gp2026":42513.41},"Statesboro":{"sales2025":461628.62,"sales2026":436169.39,"gp2025":42022.9,"gp2026":47512.59},"Athens":{"sales2025":386143.35,"sales2026":392013.04,"gp2025":49769.34,"gp2026":44201.35}}},{"week":8,"sales2025":303267.93,"sales2026":431601.88,"change":128333.95,"changePct":0.42317,"gp2025":36984.54,"gp2026":56432.01,"gpChange":19447.47,"locations":{"Byron":{"sales2025":421597.5,"sales2026":509094.67,"gp2025":54514.42,"gp2026":63853.94},"Tifton":{"sales2025":305073.85,"sales2026":433432.69,"gp2025":36920.78,"gp2026":56316.59},"Statesboro":{"sales2025":495206.73,"sales2026":471785.27,"gp2025":65989.66,"gp2026":51875.84},"Athens":{"sales2025":352936.66,"sales2026":423549.38,"gp2025":43955.8,"gp2026":51934.53}}},{"week":9,"sales2025":354550.14,"sales2026":452361.44,"change":97811.3,"changePct":0.275874,"gp2025":54039.46,"gp2026":65480.89,"gpChange":11441.43,"locations":{"Byron":{"sales2025":388503.35,"sales2026":545565.1,"gp2025":61624.26,"gp2026":66852.36},"Tifton":{"sales2025":356486.14,"sales2026":454495.49,"gp2025":53937.97,"gp2026":0},"Statesboro":{"sales2025":487011.7,"sales2026":440676.51,"gp2025":64815.44,"gp2026":36031.08},"Athens":{"sales2025":332534.01,"sales2026":447750.21,"gp2025":44801.32,"gp2026":54336.7}}},{"week":10,"sales2025":412345.78,"sales2026":468339.07,"change":55993.29,"changePct":0.135792,"gp2025":58080.16,"gp2026":65382.09,"gpChange":7301.93,"locations":{"Byron":{"sales2025":634848.31,"sales2026":578802.13,"gp2025":73784.03,"gp2026":77971.84},"Tifton":{"sales2025":414364.78,"sales2026":470118.83,"gp2025":57982.65,"gp2026":65184.07},"Statesboro":{"sales2025":571196.82,"sales2026":497235.86,"gp2025":61522.51,"gp2026":55352.06},"Athens":{"sales2025":375647.27,"sales2026":412614.08,"gp2025":47383.42,"gp2026":56149.13}}},{"week":11,"sales2025":299421.68,"sales2026":546099.8,"change":246678.12,"changePct":0.823849,"gp2025":47738.92,"gp2026":79416.32,"gpChange":31677.4,"locations":{"Byron":{"sales2025":475206.0,"sales2026":640012.25,"gp2025":66232.5,"gp2026":80560.68},"Tifton":{"sales2025":301002.68,"sales2026":543488.79,"gp2025":47663.3,"gp2026":78817.87},"Statesboro":{"sales2025":495203.93,"sales2026":504056.61,"gp2025":66603.02,"gp2026":55433.1},"Athens":{"sales2025":403830.44,"sales2026":501260.43,"gp2025":51959.3,"gp2026":70066.14}}},{"week":12,"sales2025":298950.36,"sales2026":424333.68,"change":125383.32,"changePct":0.419412,"gp2025":48950.19,"gp2026":60917.7,"gpChange":11967.51,"locations":{"Byron":{"sales2025":453355.91,"sales2026":601780.28,"gp2025":64771.83,"gp2026":67903.66},"Tifton":{"sales2025":300523.39,"sales2026":423218.69,"gp2025":48842.76,"gp2026":60222.15},"Statesboro":{"sales2025":536189.73,"sales2026":468570.64,"gp2025":71961.73,"gp2026":54358.2},"Athens":{"sales2025":364320.36,"sales2026":471417.78,"gp2025":51103.97,"gp2026":63914.76}}},{"week":13,"sales2025":313484.87,"sales2026":457493.26,"change":144008.39,"changePct":0.459379,"gp2025":46319.04,"gp2026":60789.56,"gpChange":14470.52,"locations":{"Byron":{"sales2025":397626.94,"sales2026":644955.08,"gp2025":57424.99,"gp2026":65278.52},"Tifton":{"sales2025":315203.87,"sales2026":459112.59,"gp2025":46243.42,"gp2026":60656.24},"Statesboro":{"sales2025":521325.91,"sales2026":537096.2,"gp2025":70354.11,"gp2026":58086.69},"Athens":{"sales2025":359990.21,"sales2026":499530.35,"gp2025":49554.54,"gp2026":60263.27}}},{"week":14,"sales2025":341112.78,"sales2026":522761.89,"change":181649.11,"changePct":0.532519,"gp2025":50418.76,"gp2026":66013.4,"gpChange":15594.64,"locations":{"Byron":{"sales2025":531766.41,"sales2026":695710.98,"gp2025":70423.4,"gp2026":88238.24},"Tifton":{"sales2025":342927.94,"sales2026":524966.89,"gp2025":50299.52,"gp2026":65927.83},"Statesboro":{"sales2025":515104.02,"sales2026":518940.04,"gp2025":70282.06,"gp2026":60444.4},"Athens":{"sales2025":404237.49,"sales2026":435410.47,"gp2025":55055.27,"gp2026":52300.73}}},{"week":15,"sales2025":380833.6,"sales2026":441159.9,"change":60326.3,"changePct":0.158406,"gp2025":54651.08,"gp2026":61749.54,"gpChange":7098.46,"locations":{"Byron":{"sales2025":591777.75,"sales2026":722977.83,"gp2025":87767.2,"gp2026":90146.06},"Tifton":{"sales2025":382489.6,"sales2026":441600.98,"gp2025":54545.61,"gp2026":61626.96},"Statesboro":{"sales2025":557293.32,"sales2026":487888.46,"gp2025":73158.27,"gp2026":54237.02},"Athens":{"sales2025":549425.64,"sales2026":372553.22,"gp2025":76248.34,"gp2026":50114.46}}},{"week":16,"sales2025":333294.04,"sales2026":393167.3,"change":59873.26,"changePct":0.179641,"gp2025":56625.53,"gp2026":53546.95,"gpChange":-3078.58,"locations":{"Byron":{"sales2025":596454.03,"sales2026":581034.42,"gp2025":96832.78,"gp2026":78029.98},"Tifton":{"sales2025":334890.98,"sales2026":394978.03,"gp2025":56567.76,"gp2026":53466.73},"Statesboro":{"sales2025":527256.33,"sales2026":514214.76,"gp2025":83874.62,"gp2026":56587.44},"Athens":{"sales2025":454735.18,"sales2026":385857.73,"gp2025":63725.16,"gp2026":49486.82}}},{"week":17,"sales2025":355870.61,"sales2026":420293.86,"change":64423.25,"changePct":0.18103,"gp2025":64902.25,"gp2026":58256.7,"gpChange":-6645.55,"locations":{"Byron":{"sales2025":484884.37,"sales2026":631906.61,"gp2025":83593.59,"gp2026":81159.27},"Tifton":{"sales2025":357674.61,"sales2026":422232.86,"gp2025":64782.85,"gp2026":58218.89},"Statesboro":{"sales2025":539417.6,"sales2026":519390.14,"gp2025":83810.45,"gp2026":61531.18},"Athens":{"sales2025":324248.22,"sales2026":453313.95,"gp2025":55300.1,"gp2026":56826.95}}},{"week":18,"sales2025":394844.47,"sales2026":441890.32,"change":47045.85,"changePct":0.11915,"gp2025":75014.24,"gp2026":60198.08,"gpChange":-14816.16,"locations":{"Byron":{"sales2025":519580.9,"sales2026":550851.99,"gp2025":83804.41,"gp2026":71336.72},"Tifton":{"sales2025":396848.53,"sales2026":443646.32,"gp2025":74924.75,"gp2026":60144.35},"Statesboro":{"sales2025":531111.92,"sales2026":521227.35,"gp2025":83353.85,"gp2026":60075.84},"Athens":{"sales2025":347877.2,"sales2026":511784.91,"gp2025":62474.46,"gp2026":61587.44}}},{"week":19,"sales2025":428994.06,"sales2026":468288.87,"change":39294.81,"changePct":0.091598,"gp2025":81852.17,"gp2026":66516.12,"gpChange":-15336.05,"locations":{"Byron":{"sales2025":595298.83,"sales2026":533288.93,"gp2025":100931.56,"gp2026":73801.54},"Tifton":{"sales2025":430969.01,"sales2026":470134.9,"gp2025":81700.88,"gp2026":66433.14},"Statesboro":{"sales2025":584155.89,"sales2026":501653.66,"gp2025":92927.32,"gp2026":59357.94},"Athens":{"sales2025":381173.58,"sales2026":551444.45,"gp2025":69387.3,"gp2026":72470.71}}},{"week":20,"sales2025":326178.17,"sales2026":399853.0,"change":73674.83,"changePct":0.225873,"gp2025":64917.52,"gp2026":57051.05,"gpChange":-7866.47,"locations":{"Byron":{"sales2025":488434.99,"sales2026":537770.24,"gp2025":82420.49,"gp2026":74998.44},"Tifton":{"sales2025":327879.41,"sales2026":401595.03,"gp2025":64832.19,"gp2026":57001.33},"Statesboro":{"sales2025":529184.8,"sales2026":463117.3,"gp2025":88137.64,"gp2026":54170.94},"Athens":{"sales2025":381262.85,"sales2026":399091.79,"gp2025":67066.79,"gp2026":51646.22}}},{"week":21,"sales2025":393701.7,"sales2026":377297.0,"change":-16404.7,"changePct":-0.041668,"gp2025":73092.23,"gp2026":58301.08,"gpChange":-14791.15,"locations":{"Byron":{"sales2025":550075.48,"sales2026":581257.34,"gp2025":97040.79,"gp2026":77922.52},"Tifton":{"sales2025":395583.7,"sales2026":378909.39,"gp2025":73030.54,"gp2026":58212.12},"Statesboro":{"sales2025":575972.34,"sales2026":475852.73,"gp2025":94703.25,"gp2026":58912.24},"Athens":{"sales2025":405443.52,"sales2026":407849.53,"gp2025":71959.77,"gp2026":56981.17}}},{"week":22,"sales2025":346389.91,"sales2026":286570.53,"change":-59819.38,"changePct":-0.172694,"gp2025":61512.47,"gp2026":46168.86,"gpChange":-15343.61,"locations":{"Byron":{"sales2025":444933.24,"sales2026":385149.63,"gp2025":74142.43,"gp2026":57098.92},"Tifton":{"sales2025":347990.91,"sales2026":287879.95,"gp2025":61416.95,"gp2026":46124.68},"Statesboro":{"sales2025":428728.05,"sales2026":425248.3,"gp2025":68453.95,"gp2026":50103.87},"Athens":{"sales2025":387834.9,"sales2026":396320.4,"gp2025":60949.68,"gp2026":53828.87}}},{"week":23,"sales2025":352730.05,"sales2026":455395.57,"change":102665.52,"changePct":0.29106,"gp2025":59573.52,"gp2026":66364.35,"gpChange":6790.83,"locations":{"Byron":{"sales2025":529599.33,"sales2026":553010.94,"gp2025":82691.56,"gp2026":82824.42},"Tifton":{"sales2025":352730.05,"sales2026":455395.57,"gp2025":59573.52,"gp2026":66364.35},"Statesboro":{"sales2025":541659.65,"sales2026":480859.78,"gp2025":83855.52,"gp2026":58476.32},"Athens":{"sales2025":473748.1,"sales2026":435054.06,"gp2025":77895.85,"gp2026":62859.14}}}],"depts":[{"dept":"RAD LT TRUCK","sales":3468938.76,"gp":510917.08,"gpPct":0.147283,"lineItems":2344,"assessment":"TOP PERFORMER"},{"dept":"TRUCK TIRES","sales":2072979.57,"gp":200698.0,"gpPct":0.096816,"lineItems":1079,"assessment":"STRONG - watch margin"},{"dept":"PASSENGER TIRES","sales":1891624.61,"gp":326216.53,"gpPct":0.172453,"lineItems":2379,"assessment":"TOP PERFORMER"},{"dept":"FARM TIRES","sales":545783.47,"gp":69392.06,"gpPct":0.127142,"lineItems":501,"assessment":"STRONG - watch margin"},{"dept":"ST TRAILER","sales":415407.26,"gp":86019.23,"gpPct":0.207072,"lineItems":929,"assessment":"STRONG - good margins"},{"dept":"OFF THE ROAD TIRES","sales":324788.09,"gp":26939.21,"gpPct":0.082944,"lineItems":103,"assessment":"STRONG - watch margin"},{"dept":"INDUSTRIAL TIRES","sales":152271.35,"gp":14776.23,"gpPct":0.097039,"lineItems":168,"assessment":"STRONG - watch margin"},{"dept":"TUBES","sales":104724.66,"gp":22504.7,"gpPct":0.214894,"lineItems":568,"assessment":"STRONG - good margins"},{"dept":"OUTSIDE PURCHASE","sales":43965.51,"gp":2049.93,"gpPct":0.046626,"lineItems":4,"assessment":"WATCH - GP needs improvement"},{"dept":"VALVE STEMS","sales":29617.68,"gp":9632.76,"gpPct":0.325237,"lineItems":148,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"WHEEL WEIGHTS","sales":17287.59,"gp":7864.34,"gpPct":0.454912,"lineItems":105,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"TIRE TOOLS","sales":14000.5,"gp":4090.89,"gpPct":0.292196,"lineItems":60,"assessment":"STRONG - good margins"},{"dept":"WHEELS","sales":11990.9,"gp":1750.6,"gpPct":0.145994,"lineItems":46,"assessment":"GOOD - Healthy margins"},{"dept":"LAWN & GARDEN","sales":8611.77,"gp":1691.73,"gpPct":0.196444,"lineItems":106,"assessment":"GOOD - Healthy margins"},{"dept":"PATCHES AND REPAIR","sales":7786.85,"gp":2655.98,"gpPct":0.341085,"lineItems":124,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"MOUNTING LUBE","sales":3503.36,"gp":1240.87,"gpPct":0.354194,"lineItems":43,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"ALIGNMENT SHIMS","sales":1767.14,"gp":429.79,"gpPct":0.243212,"lineItems":2,"assessment":"STRONG - good margins"},{"dept":"FREIGHT CHARGES","sales":644.35,"gp":359.35,"gpPct":0.557694,"lineItems":3,"assessment":"GREAT MARGIN - accessory builder"},{"dept":"ATV TIRES","sales":551.03,"gp":137.57,"gpPct":0.24966,"lineItems":2,"assessment":"STRONG - good margins"}],"actionPlan":[{"salesman":"House","city":"EL SEGUNDO","custNum":200976,"customer":"TIRESEASY-LLC (TIFTON WHSE)","sales2025":240163.32,"sales2026":1054590.97,"change":814427.65,"gpPct":0.141912,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.1%) | LOW GP: INDUSTRIAL TIRES (9.2%) | LOW GP: FARM TIRES (0.9%)"},{"salesman":"House","city":"TREVOSE","custNum":200953,"customer":"SIMPLE TIRE - TIFTON","sales2025":410682.38,"sales2026":648138.93,"change":237456.55,"gpPct":0.169633,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$57741) | LOW GP: INDUSTRIAL TIRES (5.3%) | LOW GP: FARM TIRES (9.4%) | DOWN: FARM TIRES (-$368)"},{"salesman":"Tiffany","city":"OCALA","custNum":200922,"customer":"ADVANCED TIRE SERVICE","sales2025":50788.35,"sales2026":353576.71,"change":302788.36,"gpPct":0.075772,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"DISCOUNTS/COUPONS","focus":"LOW GP: INDUSTRIAL TIRES (4.4%) | LOW GP: TRUCK TIRES (9.3%) | LOW GP: OFF THE ROAD TIRES (3.8%)"},{"salesman":"Austin","city":"VALDOSTA","custNum":200266,"customer":"FUSSELL TIRE & SERVICE","sales2025":251152.42,"sales2026":341332.76,"change":90180.34,"gpPct":0.075829,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: TRUCK TIRES (4.5%) | LOW GP: FARM TIRES (5.8%) | LOW GP: OFF THE ROAD TIRES (6.9%) | DOWN: OFF THE ROAD TIRES (-$23498) | LOST: OUTSIDE PURCHASE (was $1509) | LOST: FREIGHT CHARGES (was $275)"},{"salesman":"Larry","city":"TIFTON","custNum":200635,"customer":"DELTORO TIRE #2","sales2025":213753.67,"sales2026":287560.67,"change":73807.0,"gpPct":0.114846,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: TIRE TOOLS (-$442) | DOWN: OFF THE ROAD TIRES (-$1311)"},{"salesman":"Larry","city":"DOUGLAS","custNum":101161,"customer":"JMC TIRE CO.  INC.","sales2025":285131.65,"sales2026":217436.62,"change":-67695.03,"gpPct":0.114558,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: OFF THE ROAD TIRES (was $3263) | LOW GP: TRUCK TIRES (8.8%) | DOWN: RAD LT TRUCK (-$25926) | DOWN: FARM TIRES (-$11588) | LOST: INDUSTRIAL TIRES (was $2628) | LOST: WHEELS (was $500)"},{"salesman":"Larry","city":"TIFTON","custNum":101326,"customer":"TIFTON GENERAL TIRE","sales2025":182778.48,"sales2026":180647.74,"change":-2130.74,"gpPct":0.094796,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.9%) | DOWN: TRUCK TIRES (-$19181) | LOW GP: PASSENGER TIRES (7.1%) | LOW GP: RAD LT TRUCK (8.8%) | DOWN: INDUSTRIAL TIRES (-$410) | DOWN: WHEELS (-$454) | LOST: OFF THE ROAD TIRES (was $139)"},{"salesman":"Larry","city":"SYLVESTER","custNum":200220,"customer":"SINGLETARY & SON TIRE CO","sales2025":85173.08,"sales2026":153514.19,"change":68341.11,"gpPct":0.146837,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (9.7%)"},{"salesman":"Tiffany","city":"LAKELAND","custNum":200827,"customer":"LAKELAND TIRE DBA COOK & SONS","sales2025":110531.38,"sales2026":152767.5,"change":42236.12,"gpPct":0.157022,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: OFF THE ROAD TIRES (6.0%)"},{"salesman":"Unknown","city":"Unknown","custNum":3000389,"customer":"SOUTHERN TIRE MART, LLC (#134)","sales2025":29199.15,"sales2026":148805.75,"change":119606.6,"gpPct":0.136735,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: RAD LT TRUCK (6.0%) | LOW GP: PASSENGER TIRES (4.3%) | DOWN: FARM TIRES (-$5655) | LOW GP: OFF THE ROAD TIRES (8.8%) | LOST: INDUSTRIAL TIRES (was $1099)"},{"salesman":"House","city":"THOMASVILLE","custNum":200410,"customer":"EZDEALIN WHEELS AND TIRES","sales2025":122828.55,"sales2026":142918.32,"change":20089.77,"gpPct":0.20616,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (7.0%) | DOWN: PASSENGER TIRES (-$2189)"},{"salesman":"Larry","city":"FITZGERALD","custNum":101463,"customer":"SHELL RAPID LUBE (FITZGERALD)","sales2025":112199.96,"sales2026":121202.71,"change":9002.75,"gpPct":0.147902,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: INDUSTRIAL TIRES (-$803) | DOWN: TRUCK TIRES (-$9050) | LOST: OUTSIDE PURCHASE (was $214) | LOST: WHEELS (was $4294) | LOST: TIRE TOOLS (was $300)"},{"salesman":"Larry","city":"SYCAMORE","custNum":200891,"customer":"EJH WRECKER & TIRE SERVICE","sales2025":69701.85,"sales2026":119225.86,"change":49524.01,"gpPct":0.141565,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $237)"},{"salesman":"Larry","city":"TIFTON","custNum":101323,"customer":"ERIC'S TIRE SERVICE","sales2025":74622.85,"sales2026":114752.53,"change":40129.68,"gpPct":0.162442,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: INDUSTRIAL TIRES (9.6%)"},{"salesman":"Larry","city":"NASHVILLE","custNum":101415,"customer":"THE TIRE STORE","sales2025":80568.18,"sales2026":114751.56,"change":34183.38,"gpPct":0.155769,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TIRE TOOLS (was $9) | LOST: MOUNTING LUBE (was $23)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200198,"customer":"MOULTRIE TIRE","sales2025":67950.62,"sales2026":107850.61,"change":39899.99,"gpPct":0.172093,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: WHEELS (was $67)"},{"salesman":"Larry","city":"ASHBURN","custNum":101283,"customer":"CAMERON'S TOWING AND TIRE","sales2025":79857.84,"sales2026":107315.22,"change":27457.38,"gpPct":0.133667,"action":"GROW","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"DOUGLAS","custNum":101080,"customer":"AMERSON TIRE INC.","sales2025":154419.09,"sales2026":106972.51,"change":-47446.58,"gpPct":0.080502,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.6%) | DOWN: PASSENGER TIRES (-$12730) | LOW GP: FARM TIRES (8.3%) | LOW GP: TRUCK TIRES (9.6%) | DOWN: TRUCK TIRES (-$13696) | LOW GP: OFF THE ROAD TIRES (7.1%) | DOWN: OFF THE ROAD TIRES (-$9853) | LOST: PATCHES AND REPAIR (was $46) | LOST: LAWN & GARDEN (was $502)"},{"salesman":"Larry","city":"TIFTON","custNum":100282,"customer":"RUDY'S TIRE SERVICE","sales2025":118730.14,"sales2026":102470.82,"change":-16259.32,"gpPct":0.173845,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.5%) | LOST: LAWN & GARDEN (was $141) | LOST: WHEEL WEIGHTS (was $48)"},{"salesman":"House","city":"ELLENTON","custNum":101512,"customer":"ELLENTON TIRE AND AUTO","sales2025":98163.3,"sales2026":102363.89,"change":4200.59,"gpPct":0.124228,"action":"OK","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (5.9%) | LOW GP: PASSENGER TIRES (3.5%) | DOWN: PASSENGER TIRES (-$5773) | DOWN: OFF THE ROAD TIRES (-$1411) | LOST: VALVE STEMS (was $14) | LOST: LAWN & GARDEN (was $16) | LOST: TIRE TOOLS (was $23)"},{"salesman":"Tiffany","city":"MADISON","custNum":200636,"customer":"MTC SOUTH  INC.","sales2025":156914.18,"sales2026":101439.82,"change":-55474.36,"gpPct":0.103625,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.4%) | DOWN: TRUCK TIRES (-$61348) | DOWN: FARM TIRES (-$1866) | LOST: VALVE STEMS (was $21) | DOWN: MOUNTING LUBE (-$59)"},{"salesman":"House","city":"MOULTRIE","custNum":200690,"customer":"COLQUITT COUNTY TIRE LLC","sales2025":74467.17,"sales2026":98259.49,"change":23792.32,"gpPct":0.1467,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$2242)"},{"salesman":"Tiffany","city":"MAYO","custNum":200595,"customer":"W.R. WILLIAMS","sales2025":99727.99,"sales2026":94921.85,"change":-4806.14,"gpPct":0.137617,"action":"WATCH","topDept":"FARM TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6844) | LOW GP: TRUCK TIRES (9.7%) | LOW GP: OFF THE ROAD TIRES (7.2%) | DOWN: OFF THE ROAD TIRES (-$2979)"},{"salesman":"Larry","city":"SYCAMORE","custNum":200971,"customer":"ALLEN'S TIRE","sales2025":77503.43,"sales2026":90938.05,"change":13434.62,"gpPct":0.13742,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: TRUCK TIRES (9.7%) | LOW GP: FARM TIRES (5.0%) | DOWN: PASSENGER TIRES (-$5185)"},{"salesman":"Tiffany","city":"WHITE SPRINGS","custNum":200906,"customer":"E&H TIRE","sales2025":22811.3,"sales2026":88400.98,"change":65589.68,"gpPct":0.130649,"action":"GROW","topDept":"OFF THE ROAD TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$870) | DOWN: TIRE TOOLS (-$24)"},{"salesman":"House","city":"WARWICK","custNum":101108,"customer":"PARKER TIRE DIRECT","sales2025":85761.85,"sales2026":86550.25,"change":788.4,"gpPct":0.090175,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (8.4%) | LOW GP: TRUCK TIRES (9.0%) | DOWN: VALVE STEMS (-$124) | LOST: PATCHES AND REPAIR (was $5)"},{"salesman":"House","city":"EDISON","custNum":200166,"customer":"EDISON TIRE","sales2025":40661.64,"sales2026":85501.32,"change":44839.68,"gpPct":0.130421,"action":"GROW","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (9.7%) | DOWN: INDUSTRIAL TIRES (-$1177)"},{"salesman":"Larry","city":"ALBANY","custNum":200939,"customer":"RNR TIRE EXPRESS","sales2025":55806.73,"sales2026":84621.25,"change":28814.52,"gpPct":0.205916,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"DISCOUNTS/COUPONS","focus":"LOST: ST TRAILER (was $694)"},{"salesman":"Larry","city":"TIFTON","custNum":200560,"customer":"FIVE STAR TIRE SERVICE LLC","sales2025":68081.91,"sales2026":73411.64,"change":5329.73,"gpPct":0.191304,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $595)"},{"salesman":"House","city":"DAWSON","custNum":200719,"customer":"FOSTER EASY PAY TIRE CO.  INC.","sales2025":40276.99,"sales2026":68723.92,"change":28446.93,"gpPct":0.150759,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"TUBES","focus":"LOW GP: TRUCK TIRES (9.5%) | LOST: TUBES (was $113) | LOW GP: OFF THE ROAD TIRES (9.0%)"},{"salesman":"Larry","city":"DOUGLAS","custNum":101539,"customer":"COURSON'S TIRE OF DOUGLAS","sales2025":105350.43,"sales2026":65652.08,"change":-39698.35,"gpPct":0.136859,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$28687) | DOWN: PASSENGER TIRES (-$12657) | DOWN: TUBES (-$441) | LOW GP: TRUCK TIRES (7.2%) | DOWN: LAWN & GARDEN (-$73)"},{"salesman":"Larry","city":"DESOTO","custNum":100417,"customer":"CLARK BASS SERVICE","sales2025":62756.1,"sales2026":63775.67,"change":1019.57,"gpPct":0.094554,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: RAD LT TRUCK (6.7%) | DOWN: OFF THE ROAD TIRES (-$2425) | LOW GP: PASSENGER TIRES (4.8%) | DOWN: FARM TIRES (-$11389)"},{"salesman":"Larry","city":"CORDELE","custNum":200807,"customer":"CORDELE TIRE & WHEEL  LLC","sales2025":46562.69,"sales2026":61825.68,"change":15262.99,"gpPct":0.181404,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (9.1%)"},{"salesman":"House","city":"BAINBRIDGE","custNum":200162,"customer":"DELTA TIRE CO","sales2025":67758.46,"sales2026":60060.64,"change":-7697.82,"gpPct":0.127503,"action":"LOST","topDept":"FARM TIRES","declinedDept":"FARM TIRES","focus":"LOST: TRUCK TIRES (was $6004) | LOST: TUBES (was $579) | LOST: ST TRAILER (was $1647)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200293,"customer":"THOMASVILLE TIRE DEPT.","sales2025":53394.97,"sales2026":60045.8,"change":6650.83,"gpPct":0.120936,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: FARM TIRES (7.3%) | DOWN: TRUCK TIRES (-$3836)"},{"salesman":"House","city":"ROCHELLE","custNum":101025,"customer":"ROCHELLE TIRE","sales2025":32814.34,"sales2026":57725.29,"change":24910.95,"gpPct":0.142816,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"WHEELS","focus":"LOST: WHEELS (was $287)"},{"salesman":"Unknown","city":"Unknown","custNum":102273,"customer":"COMPLETE TIRE & SVC (CORDELE)","sales2025":92145.06,"sales2026":53825.31,"change":-38319.75,"gpPct":0.0536,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (7.9%) | DOWN: TRUCK TIRES (-$23380) | LOW GP: PASSENGER TIRES (4.9%) | DOWN: PASSENGER TIRES (-$8536) | LOW GP: RAD LT TRUCK (1.8%) | DOWN: RAD LT TRUCK (-$8082) | LOW GP: OFF THE ROAD TIRES (2.0%) | LOW GP: INDUSTRIAL TIRES (1.9%)"},{"salesman":"Larry","city":"DOUGLAS","custNum":200959,"customer":"SOUTHERN GEORGIA TIRE LLC","sales2025":73475.92,"sales2026":50567.4,"change":-22908.52,"gpPct":0.178277,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$17786) | LOW GP: TRUCK TIRES (9.8%)"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200883,"customer":"ROLLING BEAR TIRES LLC","sales2025":58206.69,"sales2026":50154.38,"change":-8052.31,"gpPct":0.192372,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6799) | DOWN: TRUCK TIRES (-$1153)"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200885,"customer":"PIERCE INDUSTRIAL TIRE LLC","sales2025":23498.43,"sales2026":49256.09,"change":25757.66,"gpPct":0.024159,"action":"GROW","topDept":"OFF THE ROAD TIRES","declinedDept":"TUBES","focus":"LOST: TUBES (was $301) | LOW GP: OFF THE ROAD TIRES (2.0%) | LOW GP: RAD LT TRUCK (4.0%) | LOW GP: FARM TIRES (8.0%) | LOW GP: PASSENGER TIRES (1.0%)"},{"salesman":"House","city":"ALAPAHA","custNum":101201,"customer":"TUCKERS SERVICE STATION","sales2025":65919.29,"sales2026":49079.97,"change":-16839.32,"gpPct":0.139573,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: RAD LT TRUCK (-$7568) | DOWN: FARM TIRES (-$10977) | DOWN: LAWN & GARDEN (-$252) | LOST: PATCHES AND REPAIR (was $203) | DOWN: TUBES (-$1664)"},{"salesman":"House","city":"QUINCY","custNum":200585,"customer":"QUINCY TIRE AND RECAPPING","sales2025":55596.38,"sales2026":47662.36,"change":-7934.02,"gpPct":0.202198,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$15390)"},{"salesman":"Larry","city":"SYLVESTER","custNum":101436,"customer":"ED'S TIRE","sales2025":29215.75,"sales2026":47320.01,"change":18104.26,"gpPct":0.1813,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$498) | LOST: FARM TIRES (was $78)"},{"salesman":"Larry","city":"TIFTON","custNum":100551,"customer":"SOUTHSIDE TIRE & AUTO SERVICE","sales2025":43668.68,"sales2026":42562.76,"change":-1105.92,"gpPct":0.202269,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: INDUSTRIAL TIRES (-$224)"},{"salesman":"House","city":"ALLENTOWN","custNum":200974,"customer":"PRIORITY TIRE (TIFTON WHSE)","sales2025":29934.93,"sales2026":41113.64,"change":11178.71,"gpPct":0.127464,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (6.6%) | LOST: TRUCK TIRES (was $5027) | DOWN: ST TRAILER (-$284) | LOW GP: INDUSTRIAL TIRES (9.9%)"},{"salesman":"House","city":"OMEGA","custNum":101439,"customer":"A.T. TIRE SERVICE","sales2025":23725.45,"sales2026":40603.18,"change":16877.73,"gpPct":0.154018,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $29)"},{"salesman":"House","city":"BENICIA","custNum":201038,"customer":"GIGA TIRES  LLC (TIFTON WHSE)","sales2025":48612.81,"sales2026":40195.93,"change":-8416.88,"gpPct":0.022541,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.9%) | LOW GP: TRUCK TIRES (8.4%) | DOWN: TRUCK TIRES (-$14513) | LOW GP: ST TRAILER (8.8%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200663,"customer":"BEN'S TIRE & AUTO","sales2025":39221.25,"sales2026":39282.85,"change":61.6,"gpPct":0.203274,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: WHEEL WEIGHTS (was $120) | LOST: INDUSTRIAL TIRES (was $57)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200810,"customer":"BOULEVARD TIRE CENTER","sales2025":27821.72,"sales2026":39150.11,"change":11328.39,"gpPct":0.12507,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"DOWN: TUBES (-$710) | DOWN: INDUSTRIAL TIRES (-$870) | LOST: ST TRAILER (was $458)"},{"salesman":"House","city":"DOERUN","custNum":200193,"customer":"MCLEAN TIRES INC","sales2025":22214.45,"sales2026":38767.06,"change":16552.61,"gpPct":0.150914,"action":"GROW","topDept":"FARM TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"DOWN: OFF THE ROAD TIRES (-$582)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200294,"customer":"TIRE KING OF VALDOSTA","sales2025":20273.69,"sales2026":37404.72,"change":17131.03,"gpPct":0.165877,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$827) | DOWN: TUBES (-$76) | DOWN: FARM TIRES (-$577)"},{"salesman":"House","city":"LAKE CITY","custNum":201032,"customer":"A-1 TIRE PLUS","sales2025":52626.6,"sales2026":36566.66,"change":-16059.94,"gpPct":0.139234,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: RAD LT TRUCK (-$4760) | DOWN: PASSENGER TIRES (-$6764) | DOWN: TRUCK TIRES (-$10107)"},{"salesman":"Tiffany","city":"MADISON","custNum":200691,"customer":"KENDA TRUCK CENTER","sales2025":24851.8,"sales2026":36329.05,"change":11477.25,"gpPct":0.08459,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: TRUCK TIRES (7.9%) | DOWN: PASSENGER TIRES (-$397) | LOST: OFF THE ROAD TIRES (was $1297) | LOST: ST TRAILER (was $272)"},{"salesman":"Larry","city":"DAWSON","custNum":200946,"customer":"ABBI'S 24 HOUR","sales2025":19013.41,"sales2026":36062.53,"change":17049.12,"gpPct":0.092876,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.4%) | LOW GP: RAD LT TRUCK (5.2%)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200628,"customer":"SOUTH GEORGIA TIRE","sales2025":33673.4,"sales2026":34669.56,"change":996.16,"gpPct":0.208594,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (6.7%) | DOWN: TRUCK TIRES (-$2221) | DOWN: TUBES (-$152) | LOST: OFF THE ROAD TIRES (was $548)"},{"salesman":"House","city":"NASHVILLE","custNum":101549,"customer":"NASHVILLE TIRE","sales2025":33762.9,"sales2026":34170.51,"change":407.61,"gpPct":0.22583,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: TUBES (-$514) | LOST: TRUCK TIRES (was $630) | LOST: INDUSTRIAL TIRES (was $595)"},{"salesman":"House","city":"MOULTRIE","custNum":200829,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2025":108500.46,"sales2026":33839.46,"change":-74661.0,"gpPct":0.111708,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.4%) | DOWN: TRUCK TIRES (-$59938) | LOST: FARM TIRES (was $8231) | LOST: OFF THE ROAD TIRES (was $5859) | DOWN: ST TRAILER (-$2049) | LOST: TUBES (was $131) | LOST: INDUSTRIAL TIRES (was $1713) | LOST: LAWN & GARDEN (was $32)"},{"salesman":"Tiffany","city":"NORMAN PARK","custNum":101066,"customer":"WARRIOR CREEK TIRE  LLC","sales2025":41518.67,"sales2026":33338.55,"change":-8180.12,"gpPct":0.164558,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: TUBES (-$668) | DOWN: FARM TIRES (-$12237) | LOW GP: PASSENGER TIRES (6.5%) | LOST: LAWN & GARDEN (was $89)"},{"salesman":"House","city":"EL SEGUNDO","custNum":201040,"customer":"TIRES EASY (NAP - TIFTON)","sales2025":31700.02,"sales2026":32488.67,"change":788.65,"gpPct":0.125274,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$13316)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":101146,"customer":"NE-RO TIRE & BRAKE SERVICE INC","sales2025":34789.51,"sales2026":32131.61,"change":-2657.9,"gpPct":0.11775,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TUBES (-$1059) | DOWN: INDUSTRIAL TIRES (-$3795) | DOWN: TRUCK TIRES (-$13689) | LOW GP: OFF THE ROAD TIRES (7.7%)"},{"salesman":"Larry","city":"ALMA","custNum":101507,"customer":"BURNETTE AUTOMOTIVE SERVICE","sales2025":9554.61,"sales2026":32131.47,"change":22576.86,"gpPct":0.151394,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"PEARSON","custNum":101297,"customer":"FOUR C'S LUBE","sales2025":47470.23,"sales2026":31704.45,"change":-15765.78,"gpPct":0.135408,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$1065) | DOWN: TRUCK TIRES (-$14420) | LOST: OFF THE ROAD TIRES (was $439) | LOST: INDUSTRIAL TIRES (was $710)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200991,"customer":"THE TIRE CENTRE OF FLORIDA LLC","sales2025":23662.43,"sales2026":30572.63,"change":6910.2,"gpPct":0.195044,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$1724) | LOST: INDUSTRIAL TIRES (was $57) | LOST: TUBES (was $10) | LOW GP: OFF THE ROAD TIRES (7.2%)"},{"salesman":"Larry","city":"LEESBURG","custNum":201053,"customer":"BERNEYS TIRE SERVICE","sales2025":13232.34,"sales2026":27882.85,"change":14650.51,"gpPct":0.12587,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"DONALSONVILLE","custNum":200246,"customer":"TRI COUNTY TIRE COMPANY","sales2025":24439.78,"sales2026":27762.36,"change":3322.58,"gpPct":0.127716,"action":"GROW","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (8.5%) | DOWN: TRUCK TIRES (-$5728) | DOWN: ST TRAILER (-$984) | DOWN: PASSENGER TIRES (-$316) | DOWN: INDUSTRIAL TIRES (-$823)"},{"salesman":"House","city":"ROCHELLE","custNum":101530,"customer":"MARTIN TIRE SERVICE","sales2025":28333.34,"sales2026":26955.69,"change":-1377.65,"gpPct":0.123028,"action":"WATCH","topDept":"FARM TIRES","declinedDept":"FARM TIRES","focus":"DOWN: TUBES (-$449) | LOST: LAWN & GARDEN (was $167) | DOWN: RAD LT TRUCK (-$3718) | DOWN: FARM TIRES (-$7124) | LOW GP: WHEELS (5.3%)"},{"salesman":"House","city":"CORDELE","custNum":200890,"customer":"PMT TRK. TRAILER & TIRE REPAIR","sales2025":26257.05,"sales2026":26615.44,"change":358.39,"gpPct":0.104638,"action":"OK","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.7%) | DOWN: TRUCK TIRES (-$7789) | LOW GP: RAD LT TRUCK (9.7%) | LOST: OFF THE ROAD TIRES (was $968)"},{"salesman":"House","city":"FRIDLEY","custNum":201074,"customer":"TIRE DEPOT CO. - TAG (TIFTON)","sales2025":25767.1,"sales2026":26345.37,"change":578.27,"gpPct":0.131769,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (7.0%)"},{"salesman":"Larry","city":"ALBANY","custNum":100301,"customer":"ALBANY GENERAL TIRE SERVICE","sales2025":39842.87,"sales2026":26049.25,"change":-13793.62,"gpPct":0.108832,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: RAD LT TRUCK (4.1%) | DOWN: FARM TIRES (-$840) | LOW GP: PASSENGER TIRES (6.2%) | DOWN: TRUCK TIRES (-$10338) | LOW GP: INDUSTRIAL TIRES (8.5%) | LOST: OFF THE ROAD TIRES (was $12334)"},{"salesman":"Larry","city":"ALBANY","custNum":100967,"customer":"SOUTHEASTERN COMMERCIAL TIRE","sales2025":60720.09,"sales2026":25139.31,"change":-35580.78,"gpPct":0.124212,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$37308) | LOW GP: RAD LT TRUCK (10.0%) | DOWN: TUBES (-$130) | DOWN: FARM TIRES (-$2851) | LOST: WHEEL WEIGHTS (was $12)"},{"salesman":"Larry","city":"MOULTRIE","custNum":201035,"customer":"DAVID'S AUTO SALES (MOULTRIE)","sales2025":10401.49,"sales2026":25112.48,"change":14710.99,"gpPct":0.2368,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"DOUGLAS","custNum":101295,"customer":"DAVID'S AUTO SALES / DOUGLAS","sales2025":24851.06,"sales2026":24730.14,"change":-120.92,"gpPct":0.23946,"action":"OK","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200674,"customer":"DASHER LLC","sales2025":31979.05,"sales2026":23546.61,"change":-8432.44,"gpPct":0.198658,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$7957)"},{"salesman":"House","city":"VALDOSTA","custNum":200631,"customer":"MARQUEZ TIRE SHOP","sales2025":35506.61,"sales2026":23391.51,"change":-12115.1,"gpPct":0.165687,"action":"LOST","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$11679) | DOWN: PASSENGER TIRES (-$2185)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200362,"customer":"RAY NORTON TIRE & AUTO","sales2025":9067.96,"sales2026":23164.8,"change":14096.84,"gpPct":0.184536,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":200868,"customer":"GOLDEN ENVIRONMENTAL","sales2025":23979.83,"sales2026":22992.61,"change":-987.22,"gpPct":0.187208,"action":"WATCH","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: ST TRAILER (-$520) | DOWN: RAD LT TRUCK (-$284)"},{"salesman":"Tiffany","city":"QUITMAN","custNum":200383,"customer":"WILLIAMS ALIGNMENT & TIRE","sales2025":13044.88,"sales2026":22873.62,"change":9828.74,"gpPct":0.155497,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"INDUSTRIAL TIRES","focus":"LOST: INDUSTRIAL TIRES (was $803)"},{"salesman":"Larry","city":"ALBANY","custNum":200592,"customer":"BERNEY'S TIRE SERVICE","sales2025":12857.67,"sales2026":22671.24,"change":9813.57,"gpPct":0.156623,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: TRUCK TIRES (9.8%)"},{"salesman":"House","city":"LIVE OAK","custNum":201008,"customer":"DBJ MOBILE TIRE SERVICE  INC.","sales2025":16819.88,"sales2026":22273.23,"change":5453.35,"gpPct":0.149361,"action":"GROW","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $739) | DOWN: RAD LT TRUCK (-$660)"},{"salesman":"Larry","city":"ALMA","custNum":200920,"customer":"PRECISION MAINTENANCE","sales2025":30668.91,"sales2026":22028.84,"change":-8640.07,"gpPct":0.147275,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$8286)"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200687,"customer":"LIVE OAK TIRE CENTER  LLC","sales2025":5522.82,"sales2026":21731.92,"change":16209.1,"gpPct":0.103023,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (5.8%) | LOW GP: FARM TIRES (10.0%)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200755,"customer":"N-T TIRE SERVICE","sales2025":25743.14,"sales2026":21610.7,"change":-4132.44,"gpPct":0.16968,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$1058)"},{"salesman":"House","city":"TIFTON","custNum":200709,"customer":"ASHLEY'S AUTOMOTIVE REPAIR","sales2025":12518.58,"sales2026":21027.95,"change":8509.37,"gpPct":0.118681,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: TUBES (-$25) | LOW GP: TRUCK TIRES (7.3%) | DOWN: ST TRAILER (-$303)"},{"salesman":"House","city":"MOULTRIE","custNum":201003,"customer":"NICHOLAS TIRES INC.","sales2025":26416.42,"sales2026":20704.9,"change":-5711.52,"gpPct":0.227654,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$4433) | DOWN: ST TRAILER (-$1354) | LOST: TRUCK TIRES (was $217)"},{"salesman":"Larry","city":"ADEL","custNum":200804,"customer":"FIVE STAR TIRE ****ADEL****","sales2025":30178.2,"sales2026":20402.93,"change":-9775.27,"gpPct":0.178684,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$8822) | DOWN: ST TRAILER (-$2439) | LOST: LAWN & GARDEN (was $107)"},{"salesman":"House","city":"BROXTON","custNum":201028,"customer":"CLEMENT USED TIRES","sales2025":22854.52,"sales2026":20077.27,"change":-2777.25,"gpPct":0.165107,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $1478) | LOST: TRUCK TIRES (was $239)"},{"salesman":"House","city":"WAYCROSS","custNum":200754,"customer":"JORGE USED TIRE SHOP","sales2025":26511.61,"sales2026":20010.42,"change":-6501.19,"gpPct":0.235525,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: RAD LT TRUCK (-$2080) | LOST: TRUCK TIRES (was $4080) | LOST: WHEELS (was $79) | LOST: LAWN & GARDEN (was $29) | LOST: FARM TIRES (was $549)"},{"salesman":"House","city":"DOUGLAS","custNum":101164,"customer":"DAVIS TIRE (DOUGLAS)","sales2025":7653.49,"sales2026":19645.52,"change":11992.03,"gpPct":0.165859,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200867,"customer":"24/7 DIESEL AND TIRE REPAIR","sales2025":18694.45,"sales2026":19435.36,"change":740.91,"gpPct":0.126465,"action":"OK","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":200915,"customer":"DAVID'S AUTO SALES / TIFTON","sales2025":27543.71,"sales2026":19259.85,"change":-8283.86,"gpPct":0.210511,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5166) | LOST: ST TRAILER (was $92)"},{"salesman":"Larry","city":"TIFTON","custNum":200973,"customer":"ERIC'S TIRE (REBEL ROAD)","sales2025":23927.07,"sales2026":19175.47,"change":-4751.6,"gpPct":0.160554,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$3141) | DOWN: RAD LT TRUCK (-$4678) | DOWN: ST TRAILER (-$235) | LOST: TUBES (was $10)"},{"salesman":"Tiffany","city":"HAHIRA","custNum":200270,"customer":"HAHIRA AUTOMOTIVE SERVICE","sales2025":23066.9,"sales2026":19093.27,"change":-3973.63,"gpPct":0.151728,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$6140) | DOWN: TUBES (-$34) | LOST: FARM TIRES (was $78)"},{"salesman":"House","city":"CLIMAX","custNum":201002,"customer":"ZAPATA'S TIRE","sales2025":27270.87,"sales2026":19086.14,"change":-8184.73,"gpPct":0.097664,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $79) | LOW GP: TRUCK TIRES (8.5%) | DOWN: TRUCK TIRES (-$8206) | LOST: MOUNTING LUBE (was $323)"},{"salesman":"Car Dealer","city":"TIFTON","custNum":200456,"customer":"PRINCE TOYOTA","sales2025":36504.99,"sales2026":18810.58,"change":-17694.41,"gpPct":0.042846,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: RAD LT TRUCK (-$14184) | LOW GP: PASSENGER TIRES (6.4%) | DOWN: PASSENGER TIRES (-$3604)"},{"salesman":"Anthony","city":"AUBURNDALE","custNum":2000024,"customer":"FORKLIFT TIRE OF CENTRAL FL","sales2025":0.0,"sales2026":18120.32,"change":18120.32,"gpPct":0.103881,"action":"NEW","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200655,"customer":"ECONOMY USED TIRE (ALBANY)","sales2025":6038.57,"sales2026":17922.76,"change":11884.19,"gpPct":0.106757,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: PASSENGER TIRES (7.4%) | LOST: PATCHES AND REPAIR (was $304) | LOST: TRUCK TIRES (was $457)"},{"salesman":"House","city":"MOULTRIE","custNum":200537,"customer":"TONY'S TIRE & ROAD SERVICE INC","sales2025":11155.89,"sales2026":17647.12,"change":6491.23,"gpPct":0.163593,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: TRUCK TIRES (9.0%) | DOWN: TRUCK TIRES (-$4816) | LOST: VALVE STEMS (was $38)"},{"salesman":"Larry","city":"NAHUNTA","custNum":200478,"customer":"82 TIRE & LUBE","sales2025":14830.51,"sales2026":17074.46,"change":2243.95,"gpPct":0.228281,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"House","city":"ASHBURN","custNum":201007,"customer":"SHORTY HUGHES TRUCKING  LLC","sales2025":33282.4,"sales2026":16949.89,"change":-16332.51,"gpPct":0.143903,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$16333) | LOST: ST TRAILER (was $35) | LOST: MOUNTING LUBE (was $39)"},{"salesman":"Tiffany","city":"PERRY","custNum":200673,"customer":"JB'S TIRE & REPAIR SVC.","sales2025":6104.06,"sales2026":16849.05,"change":10744.99,"gpPct":0.08378,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (4.9%) | LOW GP: PASSENGER TIRES (5.6%)"},{"salesman":"House","city":"CORDELE","custNum":200866,"customer":"LEMUS TIRE SHOP","sales2025":8090.65,"sales2026":16848.7,"change":8758.05,"gpPct":0.21822,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"WHEEL WEIGHTS","focus":"DOWN: VALVE STEMS (-$12) | DOWN: WHEEL WEIGHTS (-$35)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":101477,"customer":"DISCOUNT TIRE (ALMA)OSTEEN","sales2025":14235.17,"sales2026":16547.76,"change":2312.59,"gpPct":0.137627,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOW GP: PASSENGER TIRES (8.3%) | LOST: OUTSIDE PURCHASE (was $140) | DOWN: FARM TIRES (-$3287) | LOW GP: OFF THE ROAD TIRES (10.0%) | LOST: INDUSTRIAL TIRES (was $101)"},{"salesman":"House","city":"SYLVESTER","custNum":201052,"customer":"PRECISION DIESEL REPAIR LLC","sales2025":10406.51,"sales2026":16501.02,"change":6094.51,"gpPct":0.164974,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"DOWN: PASSENGER TIRES (-$408) | LOST: FARM TIRES (was $916)"},{"salesman":"House","city":"ROCHELLE","custNum":200880,"customer":"R&R TIRE CO.","sales2025":22136.9,"sales2026":16459.93,"change":-5676.97,"gpPct":0.112147,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$12875) | LOW GP: RAD LT TRUCK (9.0%) | DOWN: FARM TIRES (-$921)"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200961,"customer":"BABCOCK TIRE LLC","sales2025":25858.63,"sales2026":16109.96,"change":-9748.67,"gpPct":0.129545,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7789) | LOST: INDUSTRIAL TIRES (was $3182)"},{"salesman":"House","city":"LEESBURG","custNum":101525,"customer":"BMS DISCOUNT TIRES","sales2025":14169.93,"sales2026":15999.11,"change":1829.18,"gpPct":0.166365,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1120)"},{"salesman":"Larry","city":"MOULTRIE","custNum":200317,"customer":"BROTHERS TIRES","sales2025":21290.09,"sales2026":15751.81,"change":-5538.28,"gpPct":0.202162,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: ST TRAILER (-$533) | LOST: TRUCK TIRES (was $620) | LOST: LAWN & GARDEN (was $122)"},{"salesman":"Anthony","city":"WAYCROSS","custNum":200461,"customer":"WALKER JONES CHEVY-BUICK","sales2025":2455.18,"sales2026":15608.85,"change":13153.67,"gpPct":0.057086,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.9%)"},{"salesman":"House","city":"DOTHAN","custNum":200132,"customer":"WILKS A-ONE TIRE SALES","sales2025":8556.55,"sales2026":15545.13,"change":6988.58,"gpPct":0.108225,"action":"GROW","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $3924) | DOWN: OFF THE ROAD TIRES (-$1399) | LOST: RAD LT TRUCK (was $1064) | LOST: ST TRAILER (was $373)"},{"salesman":"House","city":"LAKELAND","custNum":200315,"customer":"BOBBY'S CITGO","sales2025":9246.37,"sales2026":15436.07,"change":6189.7,"gpPct":0.116465,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: PASSENGER TIRES (8.3%)"},{"salesman":"House","city":"PELHAM","custNum":2000013,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2025":332.71,"sales2026":15414.1,"change":15081.39,"gpPct":0.108533,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: FARM TIRES (9.8%) | LOW GP: INDUSTRIAL TIRES (8.0%)"},{"salesman":"House","city":"ARABI","custNum":200962,"customer":"GREENE'S TIRE SERVICE LLC","sales2025":8446.25,"sales2026":15056.15,"change":6609.9,"gpPct":0.166669,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$5552) | DOWN: TUBES (-$231)"},{"salesman":"Larry","city":"TIFTON","custNum":201062,"customer":"TIRE MASTERS LLC","sales2025":7182.96,"sales2026":15018.99,"change":7836.03,"gpPct":0.124694,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":100591,"customer":"RAINEY ALIGNMENT","sales2025":7084.44,"sales2026":14301.86,"change":7217.42,"gpPct":0.086422,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TUBES","focus":"LOW GP: RAD LT TRUCK (5.6%) | LOST: TUBES (was $8)"},{"salesman":"Unknown","city":"Unknown","custNum":2000049,"customer":"BEALL TIRE WHOLESALE, LLC","sales2025":208.48,"sales2026":14217.49,"change":14009.01,"gpPct":0.094026,"action":"GROW","topDept":"FARM TIRES","declinedDept":"","focus":"LOW GP: FARM TIRES (7.9%) | LOW GP: RAD LT TRUCK (2.7%) | LOW GP: OFF THE ROAD TIRES (8.0%)"},{"salesman":"House","city":"NEW YORK","custNum":201042,"customer":"TIRE AGENT CORP (TIFTON WHS)","sales2025":15754.03,"sales2026":14061.85,"change":-1692.18,"gpPct":0.089776,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.2%) | LOW GP: PASSENGER TIRES (8.4%)"},{"salesman":"Anthony","city":"TIFTON","custNum":101298,"customer":"TENNESON NISSAN","sales2025":6079.94,"sales2026":13673.33,"change":7593.39,"gpPct":0.132687,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ROCHELLE","custNum":101588,"customer":"STEPHENS BROTHERS","sales2025":9897.93,"sales2026":13343.89,"change":3445.96,"gpPct":0.141229,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $811) | DOWN: ST TRAILER (-$323) | LOST: TUBES (was $39) | LOST: FARM TIRES (was $613)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200967,"customer":"EDDIES AUTOMOTIVE AND TIRE","sales2025":2904.13,"sales2026":13001.22,"change":10097.09,"gpPct":0.156491,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $29)"},{"salesman":"House","city":"MOULTRIE","custNum":200947,"customer":"SUNSET TIRE & AUTOMOTIVE","sales2025":9251.94,"sales2026":12784.71,"change":3532.77,"gpPct":0.238654,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TIFTON","custNum":200331,"customer":"BUDGET CAR SALES","sales2025":38003.52,"sales2026":12565.57,"change":-25437.95,"gpPct":0.230756,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$7402) | DOWN: RAD LT TRUCK (-$18036)"},{"salesman":"House","city":"ALMA","custNum":201015,"customer":"LUBE KING & TIRES","sales2025":6905.77,"sales2026":12313.93,"change":5408.16,"gpPct":0.13906,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"LENOX","custNum":200406,"customer":"LENOX TIRE & SERVICE CENTER","sales2025":4318.58,"sales2026":11996.6,"change":7678.02,"gpPct":0.169936,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: WHEEL WEIGHTS (was $10)"},{"salesman":"Car Dealer","city":"CORDELE","custNum":2000025,"customer":"FORD CORDELE","sales2025":1817.12,"sales2026":11920.79,"change":10103.67,"gpPct":0.17733,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (6.4%)"},{"salesman":"House","city":"TIFTON","custNum":200407,"customer":"HOLLOWAY TRUCK & TRAILER REPAI","sales2025":3756.12,"sales2026":11853.86,"change":8097.74,"gpPct":0.179991,"action":"GROW","topDept":"ST TRAILER","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $78) | LOW GP: RAD LT TRUCK (6.8%)"},{"salesman":"Larry","city":"TIFTON","custNum":101371,"customer":"DIRTY SOUTH KUSTOMS","sales2025":12335.43,"sales2026":11788.11,"change":-547.32,"gpPct":0.047179,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (4.4%) | LOST: TRUCK TIRES (was $1761) | LOST: PASSENGER TIRES (was $112)"},{"salesman":"House","city":"MOULTRIE","custNum":200975,"customer":"SANTOS TIRE SHOP","sales2025":6085.42,"sales2026":11603.93,"change":5518.51,"gpPct":0.225541,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"On track"},{"salesman":"House","city":"CAIRO","custNum":200146,"customer":"BRACEWELL AUTOMOTIVE SERVICE","sales2025":9475.42,"sales2026":11581.72,"change":2106.3,"gpPct":0.075717,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (7.9%) | LOW GP: PASSENGER TIRES (7.2%) | LOST: ST TRAILER (was $231) | LOST: TUBES (was $13)"},{"salesman":"Anthony","city":"TIFTON","custNum":200832,"customer":"GRIFFIN FORD","sales2025":13146.27,"sales2026":11456.25,"change":-1690.02,"gpPct":0.167063,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$4375) | LOST: ST TRAILER (was $78)"},{"salesman":"Tiffany","city":"PELHAM","custNum":101544,"customer":"GODWIN TIRE & AUTO","sales2025":2174.93,"sales2026":11437.61,"change":9262.68,"gpPct":0.198977,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"PEARSON","custNum":200762,"customer":"POWER MAN TIRE SHOP","sales2025":30697.69,"sales2026":11377.97,"change":-19319.72,"gpPct":0.22574,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$4175) | DOWN: RAD LT TRUCK (-$11825) | DOWN: VALVE STEMS (-$83) | DOWN: ST TRAILER (-$2096) | LOST: TUBES (was $215) | LOST: TRUCK TIRES (was $1453) | DOWN: LAWN & GARDEN (-$290)"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200734,"customer":"ANDERSON FORD","sales2025":0.0,"sales2026":11375.81,"change":11375.81,"gpPct":0.171655,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Car Dealer","city":"MOULTRIE","custNum":200365,"customer":"ROBERT HUTSON LINCOLN","sales2025":12072.57,"sales2026":11184.27,"change":-888.3,"gpPct":0.173107,"action":"WATCH","topDept":"OFF THE ROAD TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1982) | LOST: ST TRAILER (was $149)"},{"salesman":"Unknown","city":"Unknown","custNum":2000052,"customer":"L&A TIRE, LLC","sales2025":-220.8,"sales2026":10806.53,"change":11027.33,"gpPct":0.170263,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"SYLVESTER","custNum":200972,"customer":"ERIC'S TIRE OF SYLVESTER","sales2025":16110.12,"sales2026":10769.93,"change":-5340.19,"gpPct":0.190806,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: LAWN & GARDEN (was $39) | DOWN: RAD LT TRUCK (-$3705) | LOST: TRUCK TIRES (was $696)"},{"salesman":"House","city":"LIVE OAK","custNum":200913,"customer":"LASHLEY'S HOMETOWN TIRE LLC","sales2025":14172.32,"sales2026":10710.02,"change":-3462.3,"gpPct":0.177108,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $105) | LOW GP: RAD LT TRUCK (0.9%) | DOWN: RAD LT TRUCK (-$1434) | LOST: LAWN & GARDEN (was $250) | LOST: ST TRAILER (was $298)"},{"salesman":"House","city":"ADEL","custNum":200416,"customer":"DENT'S SERVICE STATION","sales2025":13634.12,"sales2026":10673.52,"change":-2960.6,"gpPct":0.21874,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$883) | DOWN: PASSENGER TIRES (-$1939) | LOST: TUBES (was $4)"},{"salesman":"House","city":"TIFTON","custNum":101322,"customer":"GRIMES AUTO SERVICE","sales2025":11973.67,"sales2026":10658.44,"change":-1315.23,"gpPct":0.146948,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2908) | DOWN: ST TRAILER (-$137)"},{"salesman":"House","city":"CHICAGO","custNum":2000014,"customer":"UNITED TIRES ONLINE SALES -T","sales2025":0.0,"sales2026":10650.85,"change":10650.85,"gpPct":0.077144,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (9.8%) | LOW GP: PASSENGER TIRES (5.1%)"},{"salesman":"Larry","city":"ALBANY","custNum":200239,"customer":"T & S TIRE","sales2025":14031.21,"sales2026":10504.32,"change":-3526.89,"gpPct":0.250563,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2852)"},{"salesman":"House","city":"VALDOSTA","custNum":200327,"customer":"DRAPER TIRES & AUTOMOTIVE","sales2025":14794.03,"sales2026":10423.43,"change":-4370.6,"gpPct":0.099252,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (8.3%) | LOST: ST TRAILER (was $1311)"},{"salesman":"Larry","city":"NASHVILLE","custNum":200319,"customer":"BUCK'S AUTO REPAIR","sales2025":16471.63,"sales2026":10387.74,"change":-6083.89,"gpPct":0.124669,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$1683)"},{"salesman":"Car Dealer","city":"TIFTON","custNum":101374,"customer":"PRINCE CHEVY-OLDS  INC","sales2025":10608.56,"sales2026":10281.35,"change":-327.21,"gpPct":0.15976,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: FARM TIRES (was $183) | DOWN: TRUCK TIRES (-$550)"},{"salesman":"House","city":"VALDOSTA","custNum":200668,"customer":"24/7 TIRE","sales2025":13376.28,"sales2026":10242.38,"change":-3133.9,"gpPct":0.253589,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$3209)"},{"salesman":"House","city":"TIFTON","custNum":200901,"customer":"LENCHO'S & SON TIRE SHOP","sales2025":6820.85,"sales2026":10029.44,"change":3208.59,"gpPct":0.256205,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $295)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":2000044,"customer":"MURRAY'S TIRE & ROAD SERVICE","sales2025":2733.96,"sales2026":9964.22,"change":7230.26,"gpPct":0.175141,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.2%)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":2000012,"customer":"GLOBAL TRUCK & EQUIPMENT SALES","sales2025":440.0,"sales2026":9858.88,"change":9418.88,"gpPct":0.184705,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200474,"customer":"TIRE & WHEEL INC","sales2025":11679.56,"sales2026":9697.13,"change":-1982.43,"gpPct":0.038862,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (0.6%) | LOST: TRUCK TIRES (was $1480) | LOST: PASSENGER TIRES (was $693)"},{"salesman":"Larry","city":"CAMILLA","custNum":200965,"customer":"MARQUEZ TIRE SHOP LLC","sales2025":7891.6,"sales2026":9556.19,"change":1664.59,"gpPct":0.155108,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.5%) | DOWN: RAD LT TRUCK (-$2348)"},{"salesman":"House","city":"CORDELE","custNum":101208,"customer":"MIKE FRASER AUTO REPAIR","sales2025":14791.65,"sales2026":9379.01,"change":-5412.64,"gpPct":0.148423,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5691) | LOW GP: TRUCK TIRES (5.9%)"},{"salesman":"Anthony","city":"TIFTON","custNum":200580,"customer":"JEFF FENDER BUICK  GMC, CAD.","sales2025":639.36,"sales2026":9360.58,"change":8721.22,"gpPct":0.08101,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (8.0%) | LOST: PASSENGER TIRES (was $639)"},{"salesman":"House","city":"WHIGHAM","custNum":200268,"customer":"GIANT TIRE SALES/SERVICE","sales2025":16070.82,"sales2026":8983.83,"change":-7086.99,"gpPct":0.148164,"action":"LOST","topDept":"OFF THE ROAD TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: TRUCK TIRES (was $2932) | DOWN: RAD LT TRUCK (-$3319) | DOWN: TUBES (-$131)"},{"salesman":"Tiffany","city":"LENOX","custNum":101466,"customer":"WATTS REPAIR SERVICE","sales2025":13487.83,"sales2026":8957.18,"change":-4530.65,"gpPct":0.194827,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$4460) | LOST: ST TRAILER (was $1277)"},{"salesman":"Larry","city":"TIFTON","custNum":200409,"customer":"MASTER CRAFT IND.(NO PASS/LT)","sales2025":7898.08,"sales2026":8830.47,"change":932.39,"gpPct":0.198551,"action":"GROW","topDept":"INDUSTRIAL TIRES","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$2413)"},{"salesman":"Larry","city":"ALBANY","custNum":200601,"customer":"TUFF ENTERPRISES LLC","sales2025":12394.31,"sales2026":8792.16,"change":-3602.15,"gpPct":0.093201,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (9.6%) | DOWN: RAD LT TRUCK (-$4495) | LOW GP: PASSENGER TIRES (7.1%)"},{"salesman":"Car Dealer","city":"TIFTON","custNum":200831,"customer":"GRIFFIN CHRYSLER DODGE JEEP","sales2025":541.24,"sales2026":8741.42,"change":8200.18,"gpPct":0.142849,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALBANY","custNum":500373,"customer":"TOMAHAWK TIRE (ALBANY)","sales2025":4280.89,"sales2026":8738.26,"change":4457.37,"gpPct":0.213938,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$368) | LOW GP: TRUCK TIRES (8.9%)"},{"salesman":"House","city":"ALBANY","custNum":200949,"customer":"SOUTHERN TIRE MART","sales2025":7559.47,"sales2026":8638.25,"change":1078.78,"gpPct":0.160838,"action":"GROW","topDept":"INDUSTRIAL TIRES","declinedDept":"FARM TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1532) | LOST: FARM TIRES (was $3868) | DOWN: TRUCK TIRES (-$1007)"},{"salesman":"Car Dealer","city":"NASHVILLE","custNum":2000009,"customer":"KING FORD OF NASHVILLE","sales2025":3920.44,"sales2026":8551.04,"change":4630.6,"gpPct":0.232158,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"NASHVILLE","custNum":200941,"customer":"D&S WHEELS & DEALS LLC","sales2025":11668.0,"sales2026":8503.69,"change":-3164.31,"gpPct":0.12359,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$4022)"},{"salesman":"Anthony","city":"FITZGERALD","custNum":200833,"customer":"FITZGERALD CHRYSLER DODGE RAM","sales2025":1925.62,"sales2026":8501.62,"change":6576.0,"gpPct":0.129414,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"THOMASVILLE","custNum":200277,"customer":"IMPORT SERVICE & SALES","sales2025":9247.05,"sales2026":8380.87,"change":-866.18,"gpPct":0.204412,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$840)"},{"salesman":"Unknown","city":"Unknown","custNum":2000055,"customer":"PARKER TIRE - TIFTON","sales2025":1375.14,"sales2026":8277.63,"change":6902.49,"gpPct":0.016559,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.8%)"},{"salesman":"Tiffany","city":"OCALA","custNum":2000039,"customer":"ADVANCED TIRE SERVICE","sales2025":1268.58,"sales2026":8207.26,"change":6938.68,"gpPct":0.144105,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"FITZGERALD","custNum":101519,"customer":"MARK TAYLOR DBA/MTAA ENT.","sales2025":8947.23,"sales2026":8018.76,"change":-928.47,"gpPct":0.133024,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$589) | DOWN: RAD LT TRUCK (-$1117)"},{"salesman":"Unknown","city":"Unknown","custNum":2000046,"customer":"KELLEY MANUFACTURING CO.","sales2025":0.0,"sales2026":7964.46,"change":7964.46,"gpPct":0.18627,"action":"NEW","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALBANY","custNum":200679,"customer":"RICHARD'S AUTO CARE & TIRE SVC","sales2025":6019.31,"sales2026":7894.19,"change":1874.88,"gpPct":0.235302,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: ST TRAILER (-$564)"},{"salesman":"Larry","city":"ALBANY","custNum":200752,"customer":"RAINEY USED CARS (ALBANY)","sales2025":6461.31,"sales2026":7886.88,"change":1425.57,"gpPct":0.142041,"action":"GROW","topDept":"FARM TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (9.3%)"},{"salesman":"House","city":"QUINCY","custNum":200599,"customer":"W&L TIRE & WHEEL CO. INC.","sales2025":11568.58,"sales2026":7789.26,"change":-3779.32,"gpPct":0.219145,"action":"LOST","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$5137) | LOST: PASSENGER TIRES (was $2248) | LOW GP: INDUSTRIAL TIRES (7.6%) | DOWN: INDUSTRIAL TIRES (-$1634) | LOST: OFF THE ROAD TIRES (was $132)"},{"salesman":"House","city":"TIFTON","custNum":200918,"customer":"SOUTH GEORGIA TRUCKING SVC LLC","sales2025":5124.41,"sales2026":7788.7,"change":2664.29,"gpPct":0.159122,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1243) | DOWN: ST TRAILER (-$521) | LOST: INDUSTRIAL TIRES (was $333)"},{"salesman":"House","city":"ALBANY","custNum":200546,"customer":"SOWEGA TIRE OF ALBANY","sales2025":6553.02,"sales2026":7476.56,"change":923.54,"gpPct":0.137074,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.7%)"},{"salesman":"House","city":"NASHVILLE","custNum":200451,"customer":"HARROD BROTHERS","sales2025":7631.56,"sales2026":7008.24,"change":-623.32,"gpPct":0.109879,"action":"WATCH","topDept":"FARM TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (3.7%) | DOWN: ST TRAILER (-$426) | LOST: TRUCK TIRES (was $2315)"},{"salesman":"Unknown","city":"Unknown","custNum":2000060,"customer":"NISSAN OF TIFTON","sales2025":235.0,"sales2026":6922.02,"change":6687.02,"gpPct":0.172526,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (5.3%)"},{"salesman":"House","city":"LEESBURG","custNum":200191,"customer":"LEE COUNTY AUTO SERVICE","sales2025":6778.82,"sales2026":6919.31,"change":140.49,"gpPct":0.112178,"action":"OK","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (10.0%)"},{"salesman":"Larry","city":"PEARSON","custNum":200388,"customer":"PEARSON TIRE & LUBE","sales2025":4505.58,"sales2026":6796.64,"change":2291.06,"gpPct":0.186253,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"LAWN & GARDEN","focus":"LOST: LAWN & GARDEN (was $39)"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200777,"customer":"ANDERSON BUICK GMC","sales2025":3385.62,"sales2026":6786.56,"change":3400.94,"gpPct":0.141415,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $1162)"},{"salesman":"House","city":"THOMASVILLE","custNum":200607,"customer":"B AND B SERVICE CENTER  INC.","sales2025":4777.55,"sales2026":6553.8,"change":1776.25,"gpPct":0.240038,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1134) | LOW GP: INDUSTRIAL TIRES (9.0%)"},{"salesman":"House","city":"ADEL","custNum":201037,"customer":"AFTER HOURS TIRE SERVICE","sales2025":9003.18,"sales2026":6516.84,"change":-2486.34,"gpPct":0.190521,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1919) | LOW GP: PASSENGER TIRES (7.3%) | DOWN: PASSENGER TIRES (-$724) | LOST: LAWN & GARDEN (was $19) | LOST: FARM TIRES (was $952)"},{"salesman":"House","city":"LAKE CITY","custNum":201048,"customer":"RRO 24 HR ROADSIDE ASSISTANCE","sales2025":11391.77,"sales2026":6441.24,"change":-4950.53,"gpPct":0.163453,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: TRUCK TIRES (8.7%) | DOWN: PASSENGER TIRES (-$873) | LOST: INDUSTRIAL TIRES (was $491) | LOST: FARM TIRES (was $439)"},{"salesman":"Larry","city":"CORDELE","custNum":100907,"customer":"SMITH'S DIESEL REPAIR","sales2025":8678.57,"sales2026":6428.18,"change":-2250.39,"gpPct":0.11887,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$3405) | LOST: ST TRAILER (was $231)"},{"salesman":"Larry","city":"ALBANY","custNum":101479,"customer":"BILL THOMPSON TIRE SERVICES","sales2025":8.0,"sales2026":6389.82,"change":6381.82,"gpPct":0.065822,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (3.4%) | LOW GP: RAD LT TRUCK (3.2%) | LOW GP: TRUCK TIRES (8.0%)"},{"salesman":"Larry","city":"CORDELE","custNum":101241,"customer":"NEW PETTIS TIRE","sales2025":7257.96,"sales2026":6372.81,"change":-885.15,"gpPct":0.130782,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOST: LAWN & GARDEN (was $101) | LOW GP: PASSENGER TIRES (8.2%) | LOST: TUBES (was $310) | DOWN: TRUCK TIRES (-$672) | DOWN: FARM TIRES (-$1088)"},{"salesman":"House","city":"ALMA","custNum":200897,"customer":"ALMA TIRE & AUTO REPAIR","sales2025":6574.8,"sales2026":6299.87,"change":-274.93,"gpPct":0.109896,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (8.4%) | DOWN: RAD LT TRUCK (-$3328) | LOW GP: FARM TIRES (1.9%) | LOST: LAWN & GARDEN (was $593)"},{"salesman":"Larry","city":"TIFTON","custNum":200483,"customer":"PERRIN FARM EQUIPMENT","sales2025":7431.76,"sales2026":6280.62,"change":-1151.14,"gpPct":0.125394,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$1908) | DOWN: PASSENGER TIRES (-$1114) | LOST: ST TRAILER (was $506) | LOW GP: INDUSTRIAL TIRES (7.2%)"},{"salesman":"House","city":"CAMILLA","custNum":200396,"customer":"PATE TIRE & SERVICE LLC","sales2025":9685.63,"sales2026":6142.23,"change":-3543.4,"gpPct":0.146461,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: TRUCK TIRES (-$1695) | LOST: RAD LT TRUCK (was $2041) | LOST: ST TRAILER (was $338) | LOST: TUBES (was $222) | LOW GP: OFF THE ROAD TIRES (8.4%) | LOST: FARM TIRES (was $833)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":2000018,"customer":"ALL PRO DIESEL  LLC","sales2025":0.0,"sales2026":6105.39,"change":6105.39,"gpPct":0.218115,"action":"NEW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000047,"customer":"SUNPOINT TIRES","sales2025":0.0,"sales2026":6019.22,"change":6019.22,"gpPct":0.081614,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.3%)"},{"salesman":"Larry","city":"TIFTON","custNum":100842,"customer":"LOVE AVE. SERVICE CTR.","sales2025":4429.78,"sales2026":5894.14,"change":1464.36,"gpPct":0.121076,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$363)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":101565,"customer":"HARRY B ANDERSON","sales2025":0.0,"sales2026":5876.88,"change":5876.88,"gpPct":0.170665,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MOULTRIE","custNum":200896,"customer":"SOUTHERN AUTO SPECIALIST","sales2025":4230.81,"sales2026":5852.71,"change":1621.9,"gpPct":0.205301,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$353) | LOST: TUBES (was $20)"},{"salesman":"House","city":"THOMASVILLE","custNum":200358,"customer":"PONDER'S AUTOMOTIVE INC","sales2025":9398.3,"sales2026":5792.14,"change":-3606.16,"gpPct":0.048901,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (3.9%) | DOWN: PASSENGER TIRES (-$3496) | LOW GP: RAD LT TRUCK (6.3%) | LOST: ST TRAILER (was $347)"},{"salesman":"Car Dealer","city":"VALDOSTA","custNum":200933,"customer":"GRIFFIN CDJR VALDOSTA","sales2025":4957.24,"sales2026":5779.81,"change":822.57,"gpPct":0.212383,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"House","city":"LIVE OAK","custNum":201018,"customer":"CRAWLEY'S AUTOMOTIVE & TIRE","sales2025":2396.54,"sales2026":5512.19,"change":3115.65,"gpPct":0.185153,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"Anthony","city":"QUITMAN","custNum":200783,"customer":"CASS BURCH CHEVROLET","sales2025":0.0,"sales2026":5365.84,"change":5365.84,"gpPct":0.095167,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (7.4%)"},{"salesman":"House","city":"ALBANY","custNum":200761,"customer":"MARIO NEW AND USED TIRE SHOP","sales2025":7208.73,"sales2026":5347.4,"change":-1861.33,"gpPct":0.211607,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$1468) | DOWN: RAD LT TRUCK (-$1176) | LOST: PATCHES AND REPAIR (was $18)"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200790,"customer":"PRINCE AUTO. VALDOSTA  BUICK","sales2025":0.0,"sales2026":5164.5,"change":5164.5,"gpPct":0.12206,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"ALBANY","custNum":200214,"customer":"PETERSON TIRE & AUTO CENTER","sales2025":740.94,"sales2026":5158.84,"change":4417.9,"gpPct":0.234281,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Austin","city":"MOULTRIE","custNum":2000030,"customer":"BEASON EQUIPMENT CO","sales2025":0.0,"sales2026":5135.11,"change":5135.11,"gpPct":0.104109,"action":"NEW","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"WARWICK","custNum":200760,"customer":"MALLARD'S SERVICE CENTER","sales2025":6868.62,"sales2026":5066.83,"change":-1801.79,"gpPct":0.195346,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1792)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":2000029,"customer":"SOUTHERN TIRE MART @ PILOT","sales2025":0.0,"sales2026":4792.35,"change":4792.35,"gpPct":0.031855,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (6.9%) | LOW GP: TRUCK TIRES (4.4%) | LOW GP: RAD LT TRUCK (1.6%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000051,"customer":"STONE'S OUTDOOR POWER CENTER, LLC","sales2025":0.0,"sales2026":4776.72,"change":4776.72,"gpPct":0.088404,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (1.4%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200439,"customer":"BEAR TIRE SERVICE","sales2025":1756.63,"sales2026":4672.57,"change":2915.94,"gpPct":0.227247,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: FARM TIRES (7.0%)"},{"salesman":"House","city":"MADISON","custNum":2000038,"customer":"FAST TIRE SERVICE","sales2025":0.0,"sales2026":4663.36,"change":4663.36,"gpPct":0.071108,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (7.1%)"},{"salesman":"Larry","city":"BROXTON","custNum":100107,"customer":"JOHNSON AUTO & TIRE","sales2025":6594.39,"sales2026":4593.14,"change":-2001.25,"gpPct":0.243119,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$507) | DOWN: RAD LT TRUCK (-$1190) | LOST: TRUCK TIRES (was $298)"},{"salesman":"Larry","city":"DOUGLAS","custNum":2000042,"customer":"ROJAS AUTO REPAIR","sales2025":78.51,"sales2026":4572.37,"change":4493.86,"gpPct":0.262727,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200352,"customer":"MALUDA AUTO SALES","sales2025":6083.01,"sales2026":4475.56,"change":-1607.45,"gpPct":0.246146,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1765) | LOST: ST TRAILER (was $208)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200543,"customer":"GATOR TIRE","sales2025":5921.67,"sales2026":4389.73,"change":-1531.94,"gpPct":0.245099,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$813) | DOWN: ST TRAILER (-$343) | LOST: TUBES (was $10)"},{"salesman":"House","city":"CORDELE","custNum":100741,"customer":"MASSEY'S MUFFLER","sales2025":4670.15,"sales2026":4381.76,"change":-288.39,"gpPct":0.214752,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: LAWN & GARDEN (was $197) | DOWN: RAD LT TRUCK (-$1339) | LOST: PATCHES AND REPAIR (was $4)"},{"salesman":"House","city":"CORDELE","custNum":200940,"customer":"PERRY BROS. OIL (CORDELE)","sales2025":5271.96,"sales2026":4316.89,"change":-955.07,"gpPct":0.116717,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $692) | LOW GP: TRUCK TIRES (8.9%)"},{"salesman":"Car Dealer","city":"ALBANY","custNum":200765,"customer":"ALBANY CHRYSLER DODGE JEEP RAM","sales2025":5844.95,"sales2026":4313.09,"change":-1531.86,"gpPct":0.258299,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$3085)"},{"salesman":"Larry","city":"HOBOKEN","custNum":200683,"customer":"HICKOX AUTO DEALERS","sales2025":1440.71,"sales2026":4254.88,"change":2814.17,"gpPct":0.119646,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (9.6%)"},{"salesman":"House","city":"ALBANY","custNum":200213,"customer":"PREMIER AUTOWORKS","sales2025":590.48,"sales2026":4237.99,"change":3647.51,"gpPct":0.090201,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$307) | LOW GP: RAD LT TRUCK (5.8%)"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":200878,"customer":"THOMAS TIRE REPAIR & ROAD SVC","sales2025":3377.8,"sales2026":4215.2,"change":837.4,"gpPct":0.174625,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"ST TRAILER","focus":"On track"},{"salesman":"Larry","city":"ALMA","custNum":2000022,"customer":"R&R AUTO SERVICE & REPAIR","sales2025":644.06,"sales2026":4180.67,"change":3536.61,"gpPct":0.188957,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"WILLACOOCHEE","custNum":201014,"customer":"D&R AUTO SALES & SALVAGE PARTS","sales2025":3656.03,"sales2026":4146.74,"change":490.71,"gpPct":0.235027,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"House","city":"MONTICELLO","custNum":2000007,"customer":"KEATON & SON TIRE LLC","sales2025":2163.61,"sales2026":4138.34,"change":1974.73,"gpPct":0.203947,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $60)"},{"salesman":"Larry","city":"TIFTON","custNum":200956,"customer":"TENNESON COLLISION CENTER","sales2025":2277.72,"sales2026":4006.44,"change":1728.72,"gpPct":0.101731,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (7.0%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000057,"customer":"WOLFES WAY LLC","sales2025":0.0,"sales2026":3984.65,"change":3984.65,"gpPct":0.13419,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"DAWSON","custNum":200664,"customer":"ABR COMMERCIAL TRUCK & AUTO","sales2025":11948.71,"sales2026":3832.46,"change":-8116.25,"gpPct":0.128817,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7413) | LOST: RAD LT TRUCK (was $459) | LOST: ST TRAILER (was $249)"},{"salesman":"Tiffany","city":"BROOKSVILLE","custNum":2000040,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":3779.04,"change":3779.04,"gpPct":0.174912,"action":"NEW","topDept":"OUTSIDE PURCHASE","declinedDept":"","focus":"LOW GP: OUTSIDE PURCHASE (5.0%)"},{"salesman":"Larry","city":"SYLVESTER","custNum":2000017,"customer":"JORDAN AUTOMOTIVE & TIRES","sales2025":0.0,"sales2026":3754.02,"change":3754.02,"gpPct":0.087464,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (4.3%)"},{"salesman":"House","city":"MADISON","custNum":200624,"customer":"WALLACE MOTORS","sales2025":524.7,"sales2026":3721.7,"change":3197.0,"gpPct":0.122076,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (6.0%)"},{"salesman":"House","city":"ADEL","custNum":200401,"customer":"ADEL TIRE CO","sales2025":1322.48,"sales2026":3704.0,"change":2381.52,"gpPct":0.256061,"action":"GROW","topDept":"FARM TIRES","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$159)"},{"salesman":"Larry","city":"TIFTON","custNum":200503,"customer":"BB'S AUTOMOTIVE","sales2025":3065.43,"sales2026":3669.73,"change":604.3,"gpPct":0.176678,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.8%) | LOST: TRUCK TIRES (was $596)"},{"salesman":"Larry","city":"DOUGLAS","custNum":200907,"customer":"KNOLLWOOD TIRE & WHEEL","sales2025":12111.89,"sales2026":3657.72,"change":-8454.17,"gpPct":0.165789,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$6579) | LOST: TRUCK TIRES (was $1742) | LOST: ST TRAILER (was $133)"},{"salesman":"House","city":"ALBANY","custNum":200870,"customer":"D&K USED TIRES","sales2025":484.48,"sales2026":3601.83,"change":3117.35,"gpPct":0.237149,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"SYLVESTER","custNum":200715,"customer":"R&M AUTO TRUCKING  INC","sales2025":4156.96,"sales2026":3600.55,"change":-556.41,"gpPct":0.207203,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"DOWN: ST TRAILER (-$783)"},{"salesman":"Tiffany","city":"GAINESVILLE","custNum":2000021,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":3533.93,"change":3533.93,"gpPct":0.251957,"action":"NEW","topDept":"VALVE STEMS","declinedDept":"","focus":"On track"},{"salesman":"House","city":"JACKSONVILLE","custNum":201070,"customer":"SNIDER INDUSTRIAL","sales2025":10603.01,"sales2026":3530.64,"change":-7072.37,"gpPct":0.147901,"action":"LOST","topDept":"INDUSTRIAL TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $7401) | LOST: RAD LT TRUCK (was $908) | LOST: PASSENGER TIRES (was $412) | LOST: ST TRAILER (was $444) | LOST: FARM TIRES (was $1437)"},{"salesman":"House","city":"SYLVESTER","custNum":200681,"customer":"CITY OF SYLVESTER","sales2025":2810.44,"sales2026":3513.59,"change":703.15,"gpPct":0.103593,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"FARM TIRES","focus":"LOST: RAD LT TRUCK (was $427) | LOW GP: PASSENGER TIRES (9.6%) | LOW GP: TRUCK TIRES (6.2%) | LOST: INDUSTRIAL TIRES (was $431) | LOST: FARM TIRES (was $566)"},{"salesman":"House","city":"ALBANY","custNum":200311,"customer":"AUTOMOTIVE NECESSITIES","sales2025":861.93,"sales2026":3490.16,"change":2628.23,"gpPct":0.014698,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOW GP: RAD LT TRUCK (1.5%) | LOST: ST TRAILER (was $48)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200979,"customer":"DISCOUNT TIRE & AUTO SHOP","sales2025":10139.57,"sales2026":3388.13,"change":-6751.44,"gpPct":0.250329,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$5029) | DOWN: RAD LT TRUCK (-$2083) | LOST: ST TRAILER (was $120)"},{"salesman":"House","city":"LENOX","custNum":200417,"customer":"HIGHWAY TIRE & DIESEL","sales2025":549.8,"sales2026":3343.13,"change":2793.33,"gpPct":0.185012,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$309)"},{"salesman":"House","city":"VALDOSTA","custNum":200355,"customer":"MOTION WHEELS HUBCAPS,&TIRES","sales2025":3325.1,"sales2026":3334.53,"change":9.43,"gpPct":0.176469,"action":"OK","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"DOWN: PASSENGER TIRES (-$429) | DOWN: PATCHES AND REPAIR (-$33) | LOST: TRUCK TIRES (was $681) | LOST: LAWN & GARDEN (was $17)"},{"salesman":"House","city":"CAIRO","custNum":200658,"customer":"POWE AUTOMOTIVE","sales2025":3903.76,"sales2026":3289.68,"change":-614.08,"gpPct":0.046558,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: PASSENGER TIRES (4.3%) | DOWN: PASSENGER TIRES (-$1909) | LOW GP: RAD LT TRUCK (4.7%)"},{"salesman":"House","city":"QUITMAN","custNum":201023,"customer":"PEASE ON THE GO 24/7","sales2025":814.09,"sales2026":3258.4,"change":2444.31,"gpPct":0.232668,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200606,"customer":"AFFORDABLE TIRE SERVICE LLC","sales2025":4553.26,"sales2026":3233.73,"change":-1319.53,"gpPct":0.198263,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1301) | LOST: ST TRAILER (was $68)"},{"salesman":"Larry","city":"DOUGLAS","custNum":200510,"customer":"M & R TRUCK ACCESSORIES","sales2025":5075.05,"sales2026":3224.24,"change":-1850.81,"gpPct":0.085053,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (5.6%) | DOWN: RAD LT TRUCK (-$1284)"},{"salesman":"House","city":"VALDOSTA","custNum":201055,"customer":"HERNANDEZ TIRES SHOP","sales2025":7044.34,"sales2026":3131.91,"change":-3912.43,"gpPct":0.243474,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$2045) | DOWN: RAD LT TRUCK (-$1198) | DOWN: ST TRAILER (-$416) | LOST: TRUCK TIRES (was $252)"},{"salesman":"House","city":"MOULTRIE","custNum":200307,"customer":"ARREDONDO TIRE SERVICE","sales2025":3599.62,"sales2026":3126.07,"change":-473.55,"gpPct":0.219011,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$976) | DOWN: ST TRAILER (-$492)"},{"salesman":"Larry","city":"HORTENSE","custNum":200502,"customer":"ATKINSON TIRE","sales2025":746.61,"sales2026":3116.12,"change":2369.51,"gpPct":0.044713,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (6.0%)"},{"salesman":"House","city":"ASHBURN","custNum":200803,"customer":"SOUTH MAIN GARAGE","sales2025":3395.77,"sales2026":3058.37,"change":-337.4,"gpPct":0.151027,"action":"WATCH","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$575) | DOWN: RAD LT TRUCK (-$1202)"},{"salesman":"Anthony","city":"FITZGERALD","custNum":101300,"customer":"FITZGERALD FORD AND LINCOLN","sales2025":1497.88,"sales2026":3057.49,"change":1559.61,"gpPct":0.081174,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.8%)"},{"salesman":"House","city":"ALBANY","custNum":200743,"customer":"SOUTHERN SALES & RENTALS  LLC","sales2025":1050.54,"sales2026":3035.81,"change":1985.27,"gpPct":0.255688,"action":"GROW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200756,"customer":"THE SHOP OF ALBANY  LLC","sales2025":4409.62,"sales2026":2945.24,"change":-1464.38,"gpPct":0.208353,"action":"LOST","topDept":"ST TRAILER","declinedDept":"PASSENGER TIRES","focus":"DOWN: TRUCK TIRES (-$1271) | DOWN: PASSENGER TIRES (-$1603)"},{"salesman":"House","city":"VALDOSTA","custNum":200717,"customer":"JW AUTOMOTIVE","sales2025":4271.03,"sales2026":2913.92,"change":-1357.11,"gpPct":0.234258,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1848)"},{"salesman":"Larry","city":"OCILLA","custNum":101181,"customer":"SOUTH GA LUBE CENTER","sales2025":3978.12,"sales2026":2909.56,"change":-1068.56,"gpPct":0.184991,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $410)"},{"salesman":"Unknown","city":"Unknown","custNum":2000053,"customer":"JOINER CONTRACTING","sales2025":-246.96,"sales2026":2865.58,"change":3112.54,"gpPct":0.218081,"action":"GROW","topDept":"ST TRAILER","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (6.7%)"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200932,"customer":"RNR TIRE EXPRESS","sales2025":16694.07,"sales2026":2735.1,"change":-13958.97,"gpPct":0.179299,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$11935) | DOWN: PASSENGER TIRES (-$2024)"},{"salesman":"House","city":"ALBANY","custNum":200357,"customer":"PONDER AUTO REPAIR","sales2025":504.64,"sales2026":2726.83,"change":2222.19,"gpPct":0.207769,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $421)"},{"salesman":"Larry","city":"TIFTON","custNum":201016,"customer":"O&C AUTO REPAIR","sales2025":2381.49,"sales2026":2647.88,"change":266.39,"gpPct":0.242794,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$874)"},{"salesman":"House","city":"DOUGLAS","custNum":200648,"customer":"JOE'S AUTO REPAIR  LLC","sales2025":1295.76,"sales2026":2639.02,"change":1343.26,"gpPct":0.148961,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000050,"customer":"SUN BELT FLEET SERVICES","sales2025":0.0,"sales2026":2623.94,"change":2623.94,"gpPct":0.042783,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: TRUCK TIRES (5.9%) | LOW GP: PASSENGER TIRES (4.8%)"},{"salesman":"House","city":"SYCAMORE","custNum":101305,"customer":"SYCAMORE SALES & SALVAGE LLC","sales2025":3238.76,"sales2026":2621.24,"change":-617.52,"gpPct":0.280931,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: ST TRAILER (-$177) | LOST: FARM TIRES (was $351)"},{"salesman":"Anthony","city":"ALBANY","custNum":200838,"customer":"ALBANY CHRYSLER DODGE JEEP RAM","sales2025":1110.17,"sales2026":2583.76,"change":1473.59,"gpPct":0.112882,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (7.5%)"},{"salesman":"House","city":"ALBANY","custNum":200828,"customer":"HENRY'S ALIGNMNET","sales2025":2890.85,"sales2026":2566.58,"change":-324.27,"gpPct":0.158062,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1006)"},{"salesman":"Larry","city":"TIFTON","custNum":201017,"customer":"PINEDA'S AUTOMOTIVE","sales2025":2793.4,"sales2026":2564.75,"change":-228.65,"gpPct":0.147305,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$600)"},{"salesman":"Unknown","city":"Unknown","custNum":102274,"customer":"SNIDER FLEET SOLUTIONS","sales2025":0.0,"sales2026":2487.74,"change":2487.74,"gpPct":0.155487,"action":"NEW","topDept":"OFF THE ROAD TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"WAYCROSS","custNum":200459,"customer":"LOW COUNTRY TIRE LLC","sales2025":2419.79,"sales2026":2460.17,"change":40.38,"gpPct":0.197084,"action":"OK","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$847)"},{"salesman":"Unknown","city":"Unknown","custNum":101984,"customer":"KOUNTRY BOI TIRES, LLC","sales2025":589.23,"sales2026":2458.2,"change":1868.97,"gpPct":0.112155,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"LAKE CITY","custNum":200923,"customer":"ADVANCED TIRE SERVICE","sales2025":0.0,"sales2026":2387.13,"change":2387.13,"gpPct":0.15576,"action":"NEW","topDept":"INDUSTRIAL TIRES","declinedDept":"","focus":"LOW GP: INDUSTRIAL TIRES (5.0%)"},{"salesman":"House","city":"HAHIRA","custNum":201064,"customer":"CHAD'S AUTO REPAIR","sales2025":733.68,"sales2026":2375.68,"change":1642.0,"gpPct":0.111092,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $193) | LOW GP: RAD LT TRUCK (5.9%)"},{"salesman":"House","city":"LIVE OAK","custNum":200898,"customer":"PRECISION AUTO & MUFFLER LLC","sales2025":1917.68,"sales2026":2315.8,"change":398.12,"gpPct":0.255212,"action":"GROW","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1918)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200836,"customer":"KING MUFFLER","sales2025":317.51,"sales2026":2203.6,"change":1886.09,"gpPct":0.293076,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $297)"},{"salesman":"Unknown","city":"Unknown","custNum":100489,"customer":"LUMBER CITY ENT/BURKETT TIRE","sales2025":11569.42,"sales2026":2194.89,"change":-9374.53,"gpPct":0.103326,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"OFF THE ROAD TIRES","focus":"LOW GP: RAD LT TRUCK (9.6%) | LOST: TRUCK TIRES (was $2964) | LOST: OFF THE ROAD TIRES (was $4350) | LOST: TUBES (was $1939)"},{"salesman":"Anthony","city":"SYLVESTER","custNum":2000027,"customer":"FORD SYLVESTER","sales2025":0.0,"sales2026":2182.26,"change":2182.26,"gpPct":0.077099,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"FITZGERALD","custNum":200621,"customer":"SOUTH GA LUBE (FITZGERALD)","sales2025":5826.39,"sales2026":2148.68,"change":-3677.71,"gpPct":0.208854,"action":"LOST","topDept":"INDUSTRIAL TIRES","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: WHEELS (was $1349) | DOWN: FARM TIRES (-$1473) | LOST: OFF THE ROAD TIRES (was $2365)"},{"custNum":200539,"customer":"MIKE BURCH FORD (BLACKSHEAR)","sales2025":985.16,"sales2026":2140.8,"change":1155.64,"gpPct":0.0996,"salesman":"House","city":"","topDept":"RAD LT TRUCK","action":"GROW"},{"salesman":"Unknown","city":"Unknown","custNum":2000048,"customer":"SAN LUIS CAR SALES","sales2025":0.0,"sales2026":2123.13,"change":2123.13,"gpPct":0.184562,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MONTICELLO","custNum":200753,"customer":"AUTO TECH OF MIAMI INC.","sales2025":1043.28,"sales2026":2110.5,"change":1067.22,"gpPct":0.234996,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: WHEELS (was $44)"},{"salesman":"House","city":"TIFTON","custNum":200718,"customer":"JOBBER ACCT (TIFTON)","sales2025":4916.56,"sales2026":2079.02,"change":-2837.54,"gpPct":0.097652,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1470) | LOW GP: RAD LT TRUCK (9.1%) | DOWN: RAD LT TRUCK (-$370) | DOWN: ST TRAILER (-$358) | LOW GP: VALVE STEMS (9.1%)"},{"salesman":"House","city":"LEESBURG","custNum":101524,"customer":"MASTER BODY WORKS","sales2025":423.36,"sales2026":2063.16,"change":1639.8,"gpPct":0.144501,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"LAKE CITY","custNum":200895,"customer":"TIRE MART OF LAKE CITY","sales2025":1329.48,"sales2026":2057.28,"change":727.8,"gpPct":0.014971,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (1.5%)"},{"salesman":"Larry","city":"TIFTON","custNum":201043,"customer":"BILL'S TRAILER SERVICE","sales2025":2653.11,"sales2026":2040.69,"change":-612.42,"gpPct":0.085217,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"ST TRAILER","focus":"LOW GP: TRUCK TIRES (6.4%) | DOWN: ST TRAILER (-$905) | LOST: RAD LT TRUCK (was $183)"},{"salesman":"Anthony","city":"TIFTON","custNum":200527,"customer":"GRIFFIN CHRY/DOD/JEEP/RAM","sales2025":8331.44,"sales2026":2038.78,"change":-6292.66,"gpPct":0.158281,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $3022) | DOWN: RAD LT TRUCK (-$3043) | LOST: ST TRAILER (was $228)"},{"salesman":"Unknown","city":"Unknown","custNum":101128,"customer":"SIMPLE TIRE - *BILLING ACCT*","sales2025":0.0,"sales2026":2021.67,"change":2021.67,"gpPct":0.040061,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (4.0%)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200793,"customer":"ALL SEASON AUTO REPAIR","sales2025":1440.12,"sales2026":2007.2,"change":567.08,"gpPct":0.099935,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$939) | LOW GP: TRUCK TIRES (7.9%)"},{"salesman":"House","city":"CUTHBERT","custNum":200506,"customer":"DEVANE TIRE & SERVICE LLC","sales2025":1664.23,"sales2026":1989.37,"change":325.14,"gpPct":0.111302,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (9.9%)"},{"salesman":"House","city":"PLANT CITY","custNum":2000028,"customer":"SUNPOINT TIRES & ROAD SERVICE","sales2025":0.0,"sales2026":1957.92,"change":1957.92,"gpPct":0.100168,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"LOW GP: TRUCK TIRES (8.5%)"},{"salesman":"Unknown","city":"Unknown","custNum":101999,"customer":"SIMPLE TIRE (NAT. ACCT. D.R.)","sales2025":14036.9,"sales2026":1931.81,"change":-12105.09,"gpPct":-0.005019,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: PASSENGER TIRES (9.4%) | DOWN: RAD LT TRUCK (-$9712)"},{"salesman":"Larry","city":"TIFTON","custNum":200608,"customer":"OFFROAD POWERSPORTS","sales2025":834.44,"sales2026":1899.83,"change":1065.39,"gpPct":0.092956,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (3.9%)"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200588,"customer":"LANGDALE HYUNDAI OF SOUTH GA","sales2025":2251.49,"sales2026":1855.24,"change":-396.25,"gpPct":0.211207,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $259)"},{"salesman":"House","city":"VALDOSTA","custNum":200475,"customer":"AZALEA CITY AUTO SALES/SERVICE","sales2025":4648.2,"sales2026":1833.75,"change":-2814.45,"gpPct":0.205998,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1811) | LOST: ST TRAILER (was $257)"},{"salesman":"Tiffany","city":"QUINCY","custNum":201005,"customer":"MIDWAY ENTERPRISE FL  LLC","sales2025":3847.59,"sales2026":1819.51,"change":-2028.08,"gpPct":0.229496,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1578) | DOWN: RAD LT TRUCK (-$384)"},{"salesman":"House","city":"LAKE CITY","custNum":201058,"customer":"AFTER 5 COMM. TIRE & OFF ROAD","sales2025":16909.88,"sales2026":1742.9,"change":-15166.98,"gpPct":0.229307,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"DOWN: TRUCK TIRES (-$7204) | LOST: TUBES (was $368) | LOST: ST TRAILER (was $2343) | DOWN: RAD LT TRUCK (-$2364) | LOST: INDUSTRIAL TIRES (was $488) | LOST: PASSENGER TIRES (was $2400)"},{"salesman":"Larry","city":"ADEL","custNum":200366,"customer":"ROUNTREE PERFORMANCE","sales2025":3435.36,"sales2026":1722.4,"change":-1712.96,"gpPct":0.051202,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1394) | LOW GP: RAD LT TRUCK (3.4%)"},{"salesman":"House","city":"VALDOSTA","custNum":200654,"customer":"ECONOMY USED TIRE (VALDOSTA)","sales2025":257.68,"sales2026":1720.09,"change":1462.41,"gpPct":0.052695,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"MOUNTING LUBE","focus":"LOST: MOUNTING LUBE (was $10) | LOW GP: RAD LT TRUCK (4.9%) | LOW GP: PASSENGER TIRES (3.9%)"},{"salesman":"House","city":"HOMERVILLE","custNum":200391,"customer":"CLINCH BRAKE & ALIGNMENT","sales2025":0.0,"sales2026":1708.81,"change":1708.81,"gpPct":0.259175,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"LIVE OAK","custNum":200659,"customer":"TOWN & COUNTRY TIRE","sales2025":66.99,"sales2026":1698.37,"change":1631.38,"gpPct":0.231375,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"TUBES","focus":"LOST: TUBES (was $67)"},{"salesman":"House","city":"ADEL","custNum":200958,"customer":"BULLARD DIESEL & AUTO","sales2025":2708.98,"sales2026":1681.02,"change":-1027.96,"gpPct":0.157512,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: PASSENGER TIRES (was $273) | LOST: INDUSTRIAL TIRES (was $145)"},{"salesman":"House","city":"HOMERVILLE","custNum":200806,"customer":"WALKERS AUTO & OUTDOOR  INC","sales2025":5368.2,"sales2026":1654.18,"change":-3714.02,"gpPct":0.236462,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$1835) | DOWN: PASSENGER TIRES (-$1879)"},{"salesman":"House","city":"QUITMAN","custNum":200356,"customer":"NEELY'S SERVICE CENTER","sales2025":3484.79,"sales2026":1634.86,"change":-1849.93,"gpPct":0.152588,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1083) | DOWN: RAD LT TRUCK (-$755) | LOST: ST TRAILER (was $114)"},{"salesman":"Larry","city":"TIFTON","custNum":200468,"customer":"MCKEE'S AUTO CENTER  INC","sales2025":1052.81,"sales2026":1504.97,"change":452.16,"gpPct":0.233182,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: PATCHES AND REPAIR (was $8) | LOW GP: TRUCK TIRES (7.3%) | DOWN: TRUCK TIRES (-$436)"},{"salesman":"Unknown","city":"Unknown","custNum":2000069,"customer":"IRWIN COUNTY CUSTOMS & REPAIR","sales2025":2032.89,"sales2026":1497.31,"change":-535.58,"gpPct":0.384006,"action":"LOST","topDept":"WHEEL WEIGHTS","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200609,"customer":"LIBERTY AUTO CARE CENTER","sales2025":0.0,"sales2026":1491.38,"change":1491.38,"gpPct":0.305945,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TIFTON","custNum":101152,"customer":"BROOKS BODY SHOP","sales2025":284.0,"sales2026":1460.0,"change":1176.0,"gpPct":0.044164,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOW GP: RAD LT TRUCK (4.4%) | LOST: PASSENGER TIRES (was $284)"},{"salesman":"Car Dealer","city":"TIFTON","custNum":200455,"customer":"PRINCE HONDA","sales2025":1519.64,"sales2026":1458.12,"change":-61.52,"gpPct":0.240791,"action":"WATCH","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$804)"},{"salesman":"Tiffany","city":"DOTHAN","custNum":101878,"customer":"TRI STATE COMMERCIAL TIRE LLC","sales2025":1103.9,"sales2026":1452.12,"change":348.22,"gpPct":0.081557,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"DOWN: FARM TIRES (-$1020) | LOW GP: RAD LT TRUCK (7.5%)"},{"salesman":"House","city":"ALBANY","custNum":201056,"customer":"LEE'S AUTO SHOP","sales2025":4505.32,"sales2026":1428.8,"change":-3076.52,"gpPct":0.198929,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$1387) | LOW GP: PASSENGER TIRES (4.2%) | DOWN: PASSENGER TIRES (-$1689)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200385,"customer":"WILLIAMS AUTOMOTIVE","sales2025":1262.74,"sales2026":1392.73,"change":129.99,"gpPct":0.2737,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"On track"},{"salesman":"Larry","city":"CORDELE","custNum":101491,"customer":"QUALITY AUTO & R.V. SERVICE","sales2025":4319.6,"sales2026":1388.24,"change":-2931.36,"gpPct":0.267619,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$1628) | DOWN: RAD LT TRUCK (-$787) | LOST: ST TRAILER (was $517)"},{"salesman":"Unknown","city":"Unknown","custNum":100224,"customer":"WARD'S SERVICE CENTER","sales2025":0.0,"sales2026":1305.35,"change":1305.35,"gpPct":0.132026,"action":"NEW","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":2000020,"customer":"T&D TIRE","sales2025":6411.51,"sales2026":1270.64,"change":-5140.87,"gpPct":0.208601,"action":"LOST","topDept":"TUBES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":100125,"customer":"MACON COMMERCIAL TIRE, INC.","sales2025":0.0,"sales2026":1246.54,"change":1246.54,"gpPct":0.122996,"action":"NEW","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200370,"customer":"SMITH TIRE COMPANY","sales2025":0.0,"sales2026":1224.0,"change":1224.0,"gpPct":0.039706,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (4.0%)"},{"salesman":"House","city":"LENOX","custNum":200462,"customer":"QUALITY FEEDSTUFFS  INC","sales2025":3260.63,"sales2026":1219.31,"change":-2041.32,"gpPct":0.18075,"action":"LOST","topDept":"FARM TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2230) | LOST: WHEELS (was $175) | LOST: ST TRAILER (was $120)"},{"salesman":"House","city":"OMEGA","custNum":201039,"customer":"CERVANTES AUTO SALES","sales2025":1010.94,"sales2026":1157.09,"change":146.15,"gpPct":0.170203,"action":"GROW","topDept":"ST TRAILER","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $593)"},{"salesman":"Larry","city":"CORDELE","custNum":2000043,"customer":"CENTRAL GA TIRE LLC","sales2025":0.0,"sales2026":1150.04,"change":1150.04,"gpPct":0.303642,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MADISON","custNum":200671,"customer":"STEWARTS AUTO SERVICE CENTER","sales2025":0.0,"sales2026":1130.95,"change":1130.95,"gpPct":0.248473,"action":"NEW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"CORDELE","custNum":200562,"customer":"BEST CARS OF CORDELE  LLC","sales2025":467.26,"sales2026":1092.62,"change":625.36,"gpPct":0.307488,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TIFTON","custNum":101252,"customer":"DNA DIESEL & AUTOMOTIVE REPAIR","sales2025":2356.9,"sales2026":1076.88,"change":-1280.02,"gpPct":0.111099,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: RAD LT TRUCK (was $200) | LOW GP: TRUCK TIRES (7.6%) | DOWN: TRUCK TIRES (-$1288)"},{"salesman":"House","city":"NORMAN PARK","custNum":200642,"customer":"E.G. AUTO SALES","sales2025":2432.66,"sales2026":1029.34,"change":-1403.32,"gpPct":0.245147,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$432) | DOWN: PASSENGER TIRES (-$905) | LOST: TUBES (was $10) | LOST: FARM TIRES (was $57)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200434,"customer":"CAPITAL AUTO PARTS  INC","sales2025":0.0,"sales2026":1025.06,"change":1025.06,"gpPct":0.140499,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":200490,"customer":"T.C.A. IRRIGATION","sales2025":0.0,"sales2026":1021.42,"change":1021.42,"gpPct":0.181904,"action":"NEW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TIFTON","custNum":200805,"customer":"JOEY HALL AUTO SALES LLC","sales2025":573.83,"sales2026":988.0,"change":414.17,"gpPct":-0.060121,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $57)"},{"salesman":"House","city":"TIFTON","custNum":200759,"customer":"AADCO","sales2025":867.72,"sales2026":984.38,"change":116.66,"gpPct":0.276834,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$476)"},{"salesman":"House","city":"DOUGLAS","custNum":200720,"customer":"LUBE MASTERS","sales2025":0.0,"sales2026":968.36,"change":968.36,"gpPct":0.258251,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":102049,"customer":"RAFFIELD TIRE (TRUCK CENTER)","sales2025":0.0,"sales2026":957.96,"change":957.96,"gpPct":0.127521,"action":"NEW","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"TALLAHASSEE","custNum":200990,"customer":"PATTON'S ALIGNMENT & BRAKE SVC","sales2025":1554.4,"sales2026":928.0,"change":-626.4,"gpPct":0.039526,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.0%) | DOWN: RAD LT TRUCK (-$626)"},{"salesman":"House","city":"ALBANY","custNum":200600,"customer":"CHARLOT TRUCKING & TIRE SVC.","sales2025":0.0,"sales2026":925.74,"change":925.74,"gpPct":0.305561,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":2000015,"customer":"EDWIN'S TIRES LLC","sales2025":651.12,"sales2026":896.4,"change":245.28,"gpPct":0.225825,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000059,"customer":"LAKE CITY TIRE SHOP","sales2025":1109.18,"sales2026":894.59,"change":-214.59,"gpPct":0.320705,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"VALDOSTA","custNum":200935,"customer":"FROMETA USED CAR & TIRE CENTER","sales2025":1437.02,"sales2026":887.68,"change":-549.34,"gpPct":0.157061,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $800) | LOST: ST TRAILER (was $309)"},{"salesman":"House","city":"THOMASVILLE","custNum":201012,"customer":"AG PRO FUEL","sales2025":521.54,"sales2026":883.37,"change":361.83,"gpPct":0.075076,"action":"GROW","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOW GP: RAD LT TRUCK (7.1%) | LOW GP: TRUCK TIRES (7.9%)"},{"salesman":"Unknown","city":"Unknown","custNum":2000065,"customer":"MOSS MOTORS LLC","sales2025":0.0,"sales2026":861.16,"change":861.16,"gpPct":0.270472,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"NASHVILLE","custNum":200908,"customer":"O'STEEN CHRYSLER DODGE JEEP","sales2025":223.16,"sales2026":845.43,"change":622.27,"gpPct":0.193795,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (8.2%)"},{"salesman":"Anthony","city":"VALDOSTA","custNum":200348,"customer":"LANGDALE KIA (SERVICE)","sales2025":2756.25,"sales2026":845.43,"change":-1910.82,"gpPct":0.290574,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$307) | DOWN: RAD LT TRUCK (-$1604)"},{"salesman":"House","city":"ADEL","custNum":200980,"customer":"SLYDER'S GARAGE","sales2025":850.2,"sales2026":840.48,"change":-9.72,"gpPct":0.229298,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$287)"},{"salesman":"House","city":"BROXTON","custNum":101364,"customer":"KENNY'S AUTO AND TRUCK SALVAGE","sales2025":100.91,"sales2026":810.12,"change":709.21,"gpPct":0.185787,"action":"GROW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"MOULTRIE","custNum":200242,"customer":"THOMAS TIRE RECAPPING INC.","sales2025":20236.64,"sales2026":800.36,"change":-19436.28,"gpPct":0.1711,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$3891) | LOST: RAD LT TRUCK (was $10705) | LOST: TRUCK TIRES (was $4326) | LOST: ST TRAILER (was $139)"},{"salesman":"House","city":"LIVE OAK","custNum":200900,"customer":"GILLETTES AUTO","sales2025":1872.92,"sales2026":782.06,"change":-1090.86,"gpPct":0.235289,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: RAD LT TRUCK (-$384) | LOST: ST TRAILER (was $277) | DOWN: PASSENGER TIRES (-$429)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200579,"customer":"RODS CAR & TRUCK ACC.","sales2025":339.56,"sales2026":764.68,"change":425.12,"gpPct":0.078203,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"LOW GP: PASSENGER TIRES (4.2%)"},{"salesman":"Larry","city":"TIFTON","custNum":200676,"customer":"ALL PURPOSE AUTO CENTER","sales2025":3870.71,"sales2026":734.5,"change":-3136.21,"gpPct":0.389667,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1009) | LOST: TRUCK TIRES (was $582)"},{"salesman":"House","city":"ADEL","custNum":200393,"customer":"BRUISER'S TIRE & TOWING","sales2025":0.0,"sales2026":721.58,"change":721.58,"gpPct":0.194074,"action":"NEW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"DOUGLAS","custNum":200778,"customer":"WOODY FOLSOM CDJR (DOUGLAS)","sales2025":0.0,"sales2026":716.64,"change":716.64,"gpPct":0.025508,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (2.6%)"},{"salesman":"House","city":"CORDELE","custNum":200912,"customer":"LANE'S TRK & TRL REPAIR & AUTO","sales2025":1258.7,"sales2026":715.68,"change":-543.02,"gpPct":0.202046,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $538)"},{"salesman":"House","city":"FLORHAM PARK","custNum":201041,"customer":"TIRETREADS LLC (TIFTON ACCT)","sales2025":535.53,"sales2026":704.53,"change":169.0,"gpPct":0.261834,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$338)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200886,"customer":"C&L PERFORMANCE INC","sales2025":1716.09,"sales2026":695.88,"change":-1020.21,"gpPct":0.243102,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$585) | LOST: PASSENGER TIRES (was $362) | LOST: TUBES (was $10) | LOST: FARM TIRES (was $64)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200291,"customer":"SINGLETARY TIRE PROS","sales2025":769.32,"sales2026":674.24,"change":-95.08,"gpPct":0.282881,"action":"LOST","topDept":"PATCHES AND REPAIR","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $769)"},{"salesman":"House","city":"BLACKSHEAR","custNum":201036,"customer":"GODWIN & SON REPAIR & SALES","sales2025":845.74,"sales2026":662.7,"change":-183.04,"gpPct":0.21005,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$241) | LOST: ST TRAILER (was $59)"},{"salesman":"Unknown","city":"Unknown","custNum":2000064,"customer":"COUNTY LINE MOTORS LLC","sales2025":0.0,"sales2026":655.92,"change":655.92,"gpPct":0.281071,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"MOULTRIE","custNum":200427,"customer":"SAUNDERS AUTO REPAIR","sales2025":0.0,"sales2026":655.6,"change":655.6,"gpPct":0.213362,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":201031,"customer":"RENO'S QUALITY COLLISION","sales2025":5705.57,"sales2026":643.52,"change":-5062.05,"gpPct":0.147165,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$346) | LOST: RAD LT TRUCK (was $3270) | LOST: TRUCK TIRES (was $1924)"},{"salesman":"Unknown","city":"Unknown","custNum":2000063,"customer":"GATEWAY TIRE","sales2025":0.0,"sales2026":579.63,"change":579.63,"gpPct":-0.092179,"action":"NEW","topDept":"FARM TIRES","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"TIFTON","custNum":200744,"customer":"TIFTON COMMERCIAL","sales2025":1561.91,"sales2026":578.63,"change":-983.28,"gpPct":0.263225,"action":"LOST","topDept":"TUBES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$719) | LOST: ST TRAILER (was $412) | LOST: RAD LT TRUCK (was $123)"},{"salesman":"Larry","city":"MOULTRIE","custNum":101513,"customer":"GAY'S TIRE SERVICE","sales2025":133.86,"sales2026":568.77,"change":434.91,"gpPct":0.282821,"action":"GROW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200964,"customer":"MARTINEZ AUTO SERVICE","sales2025":0.0,"sales2026":551.4,"change":551.4,"gpPct":0.295285,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200699,"customer":"METTS PERFORMANCE & AUTOMOTIVE","sales2025":0.0,"sales2026":542.76,"change":542.76,"gpPct":0.238024,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":2000068,"customer":"STONE ENTERPRISES - FAUSSETTS LLC","sales2025":0.0,"sales2026":541.83,"change":541.83,"gpPct":0.242235,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"CAIRO","custNum":200148,"customer":"AUTO & TRUCK CARE SPECIAL","sales2025":481.6,"sales2026":465.1,"change":-16.5,"gpPct":0.278779,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"On track"},{"salesman":"House","city":"DOUGLAS","custNum":101366,"customer":"B & M AUTOMOTIVE SERVICE","sales2025":4573.22,"sales2026":457.85,"change":-4115.37,"gpPct":0.155684,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$1105) | LOST: RAD LT TRUCK (was $3010)"},{"salesman":"House","city":"BLACKSHEAR","custNum":200465,"customer":"DIXON SERVICE CENTER","sales2025":0.0,"sales2026":414.49,"change":414.49,"gpPct":0.196362,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"SYLVESTER","custNum":1999999,"customer":"GRIFFIN CHEVROLET OF SYLVESTER","sales2025":1737.92,"sales2026":408.24,"change":-1329.68,"gpPct":0.246375,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1426)"},{"salesman":"House","city":"ALBANY","custNum":201034,"customer":"ECONOMIC NICHOLAS TIRE","sales2025":874.62,"sales2026":407.92,"change":-466.7,"gpPct":0.206021,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $692)"},{"salesman":"Larry","city":"TIFTON","custNum":200501,"customer":"LARRY'S BODY SHOP","sales2025":896.92,"sales2026":404.16,"change":-492.76,"gpPct":0.265588,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$493)"},{"salesman":"House","city":"ALBANY","custNum":200795,"customer":"WINCHESTER PAINT & BODY","sales2025":1798.9,"sales2026":401.84,"change":-1397.06,"gpPct":0.103333,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: ST TRAILER (was $114) | DOWN: PASSENGER TIRES (-$260) | LOST: RAD LT TRUCK (was $1226)"},{"salesman":"House","city":"CAIRO","custNum":200287,"customer":"RIDLEY'S AUTOMOTIVE","sales2025":1349.64,"sales2026":399.16,"change":-950.48,"gpPct":0.20493,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$950)"},{"salesman":"Unknown","city":"Unknown","custNum":100165,"customer":"MOSLEY TIRE ALIGN.& BRAKE CTR.","sales2025":0.0,"sales2026":381.2,"change":381.2,"gpPct":0.163169,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"On track"},{"salesman":"House","city":"BAINBRIDGE","custNum":200230,"customer":"SOUTHERN TIRE & BATTERY","sales2025":1872.13,"sales2026":375.06,"change":-1497.07,"gpPct":0.323468,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1429) | LOST: PASSENGER TIRES (was $68)"},{"salesman":"House","city":"BAINBRIDGE","custNum":2000010,"customer":"SOUTHERN AUTOMOTIVE SVC & REP","sales2025":2938.36,"sales2026":365.0,"change":-2573.36,"gpPct":0.044164,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (4.4%) | DOWN: RAD LT TRUCK (-$2573)"},{"salesman":"Larry","city":"ALBANY","custNum":200644,"customer":"SKIP'S AUTOMOTIVE","sales2025":0.0,"sales2026":362.32,"change":362.32,"gpPct":0.280691,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"LAKE CITY","custNum":200916,"customer":"GATEWAY DIESEL  AUTO & MOBILE","sales2025":2187.3,"sales2026":328.7,"change":-1858.6,"gpPct":0.122117,"action":"LOST","topDept":"TRUCK TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: ST TRAILER (was $124) | LOST: PASSENGER TIRES (was $637) | DOWN: TRUCK TIRES (-$1098)"},{"salesman":"Tiffany","city":"JASPER","custNum":200876,"customer":"SUWANNEE VALLEY TIRE","sales2025":2500.0,"sales2026":322.44,"change":-2177.56,"gpPct":0.013925,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $2500) | LOW GP: RAD LT TRUCK (1.4%)"},{"salesman":"House","city":"DOTHAN","custNum":200927,"customer":"MASTER TIRE","sales2025":809.44,"sales2026":322.39,"change":-487.05,"gpPct":0.073451,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOW GP: RAD LT TRUCK (7.3%) | DOWN: RAD LT TRUCK (-$487)"},{"salesman":"House","city":"LEESBURG","custNum":200192,"customer":"MCGEHEE'S TIRE & AUTO","sales2025":0.0,"sales2026":295.64,"change":295.64,"gpPct":0.073332,"action":"NEW","topDept":"RAD LT TRUCK","declinedDept":"","focus":"LOW GP: RAD LT TRUCK (0.8%)"},{"salesman":"House","city":"ALBANY","custNum":101497,"customer":"PERFORMANCE MOTORSPORT","sales2025":317.7,"sales2026":285.98,"change":-31.72,"gpPct":0.30121,"action":"WATCH","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $176)"},{"salesman":"House","city":"ALMA","custNum":101476,"customer":"HART'S SERVICE STATION","sales2025":1555.84,"sales2026":284.91,"change":-1270.93,"gpPct":0.271559,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"DOWN: RAD LT TRUCK (-$1271)"},{"custNum":200533,"customer":"FLOWERS IMPORTS LLC (HONDA)","sales2025":1036.77,"sales2026":248.0,"change":-788.77,"gpPct":0.054,"salesman":"Car Dealer","city":"","topDept":"PASSENGER TIRES","action":"LOST"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200921,"customer":"AUTO DOCTOR DIESEL & REPAIR","sales2025":5520.26,"sales2026":242.46,"change":-5277.8,"gpPct":0.355234,"action":"LOST","topDept":"WHEEL WEIGHTS","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $550) | LOST: RAD LT TRUCK (was $4971)"},{"salesman":"House","city":"TIFTON","custNum":200821,"customer":"MARK'S BODY SHOP-TBR ONLY","sales2025":0.0,"sales2026":235.93,"change":235.93,"gpPct":0.137371,"action":"NEW","topDept":"TRUCK TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":201025,"customer":"A-1 WRECKER SERVICE","sales2025":10099.37,"sales2026":213.28,"change":-9886.09,"gpPct":0.278132,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$2244) | LOST: RAD LT TRUCK (was $6241) | LOST: ST TRAILER (was $769) | LOST: TRUCK TIRES (was $633)"},{"salesman":"Unknown","city":"Unknown","custNum":101776,"customer":"COLUMBUS TIRE CO","sales2025":0.0,"sales2026":203.3,"change":203.3,"gpPct":0.583965,"action":"NEW","topDept":"WHEELS","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"BLACKSHEAR","custNum":200712,"customer":"TANNER AUTO REPAIR PLUS  LLC","sales2025":7193.61,"sales2026":186.14,"change":-7007.47,"gpPct":0.17836,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $5248) | DOWN: PASSENGER TIRES (-$1759)"},{"salesman":"House","city":"LAKE PARK","custNum":2000004,"customer":"ODELL AUTOMOTIVE","sales2025":0.0,"sales2026":185.0,"change":185.0,"gpPct":0.176865,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"WAYCROSS","custNum":200484,"customer":"ROBBIE ROBERSON FORD","sales2025":1293.68,"sales2026":184.22,"change":-1109.46,"gpPct":0.080067,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $164) | LOW GP: RAD LT TRUCK (8.0%) | DOWN: RAD LT TRUCK (-$945)"},{"salesman":"House","city":"DOTHAN","custNum":200241,"customer":"THE RIM SHOP INC","sales2025":642.55,"sales2026":179.71,"change":-462.84,"gpPct":0.110845,"action":"LOST","topDept":"RAD LT TRUCK","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $404)"},{"salesman":"House","city":"POULAN","custNum":200504,"customer":"PLATINUM RECOVERY SERVICES LLC","sales2025":0.0,"sales2026":163.42,"change":163.42,"gpPct":0.212826,"action":"NEW","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"House","city":"WAYCROSS","custNum":201024,"customer":"LIBERTY CAR WASH & TIRE","sales2025":393.56,"sales2026":132.92,"change":-260.64,"gpPct":0.255191,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $494)"},{"salesman":"Tiffany","city":"WAYCROSS","custNum":200496,"customer":"MILLER TIRE CO.","sales2025":0.0,"sales2026":130.8,"change":130.8,"gpPct":0.278287,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"VALDOSTA","custNum":200914,"customer":"BESTDRIVE COMMERCIAL TIRE CTR","sales2025":4022.35,"sales2026":106.0,"change":-3916.35,"gpPct":0.362075,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"TRUCK TIRES","focus":"LOST: RAD LT TRUCK (was $1364) | LOST: INDUSTRIAL TIRES (was $230) | LOST: FARM TIRES (was $128) | LOST: TUBES (was $20) | LOST: TRUCK TIRES (was $2280)"},{"salesman":"Larry","city":"FITZGERALD","custNum":2000023,"customer":"FABOS AUTO SALES LLC","sales2025":0.0,"sales2026":104.08,"change":104.08,"gpPct":0.259416,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":101196,"customer":"BEASLEY AUTO & TRUCK REPAIR","sales2025":0.0,"sales2026":85.0,"change":85.0,"gpPct":0.297647,"action":"NEW","topDept":"WHEEL WEIGHTS","declinedDept":"","focus":"On track"},{"salesman":"Larry","city":"SYLVESTER","custNum":200500,"customer":"SHELL RAPID LUBE (SYLVESTER)","sales2025":465.28,"sales2026":78.96,"change":-386.32,"gpPct":0.215046,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"PASSENGER TIRES","focus":"DOWN: PASSENGER TIRES (-$403)"},{"salesman":"Anthony","city":"DOUGLAS","custNum":101256,"customer":"ROBERT FENDER CHEVROLET","sales2025":0.0,"sales2026":60.13,"change":60.13,"gpPct":0.276401,"action":"NEW","topDept":"PASSENGER TIRES","declinedDept":"","focus":"On track"},{"salesman":"Unknown","city":"Unknown","custNum":102396,"customer":"CONTENDER SERVICE LLC","sales2025":0.0,"sales2026":49.36,"change":49.36,"gpPct":0.260332,"action":"NEW","topDept":"TUBES","declinedDept":"","focus":"On track"},{"salesman":"House","city":"NASHVILLE","custNum":200622,"customer":"MOORE'S ACCESSORIES & OFFROAD","sales2025":2802.91,"sales2026":20.86,"change":-2782.05,"gpPct":0.298178,"action":"LOST","topDept":"TUBES","declinedDept":"RAD LT TRUCK","focus":"DOWN: TUBES (-$122) | LOST: RAD LT TRUCK (was $2660)"},{"salesman":"Anthony","city":"TIFTON","custNum":200554,"customer":"HONDA OF SOUTH GEORGIA","sales2025":825.6,"sales2026":19.72,"change":-805.88,"gpPct":0.330629,"action":"LOST","topDept":"MOUNTING LUBE","declinedDept":"RAD LT TRUCK","focus":"LOST: ST TRAILER (was $206) | LOST: RAD LT TRUCK (was $619)"},{"salesman":"House","city":"ALBANY","custNum":200750,"customer":"LAWN PERFORMANCE  LLC","sales2025":206.1,"sales2026":0.0,"change":-206.1,"gpPct":0.279683,"action":"LOST","topDept":"ST TRAILER","declinedDept":"","focus":"On track"},{"salesman":"Anthony","city":"ALBANY","custNum":200360,"customer":"PRINCE CHEVY BUICK GMC","sales2025":281.88,"sales2026":0.0,"change":-281.88,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $282)"},{"salesman":"Anthony","city":"CAIRO","custNum":200626,"customer":"HOBSON CHEVROLET BUICK","sales2025":383.48,"sales2026":0.0,"change":-383.48,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $383)"},{"salesman":"Anthony","city":"CORDELE","custNum":200882,"customer":"SUNBELT FORD CORDELE (AMI)","sales2025":690.84,"sales2026":0.0,"change":-690.84,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $691)"},{"salesman":"Anthony","city":"CORDELE","custNum":101568,"customer":"SUNBELT FORD OF CORDELE  INC","sales2025":7878.73,"sales2026":0.0,"change":-7878.73,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $5409) | LOST: PASSENGER TIRES (was $2469)"},{"salesman":"Anthony","city":"DOTHAN","custNum":200943,"customer":"BONDY'S NISSAN INC.","sales2025":482.63,"sales2026":0.0,"change":-482.63,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $483)"},{"salesman":"Anthony","city":"NASHVILLE","custNum":200688,"customer":"NASHVILLE FORD","sales2025":7675.92,"sales2026":0.0,"change":-7675.92,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $3150) | LOST: RAD LT TRUCK (was $4526)"},{"salesman":"Anthony","city":"SYLVESTER","custNum":200232,"customer":"SUNBELT FORD INC","sales2025":3522.1,"sales2026":0.0,"change":-3522.1,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: PASSENGER TIRES (was $776) | LOST: RAD LT TRUCK (was $2692) | LOST: ST TRAILER (was $55)"},{"salesman":"Anthony","city":"TIFTON","custNum":200524,"customer":"GRIFFIN FORD LINCOLN  INC","sales2025":4936.42,"sales2026":0.0,"change":-4936.42,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $4079) | LOST: ST TRAILER (was $96) | LOST: PASSENGER TIRES (was $761)"},{"salesman":"House","city":"ADEL","custNum":200822,"customer":"UNITED TIRE LLC","sales2025":140.52,"sales2026":0.0,"change":-140.52,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $130) | LOST: TUBES (was $11)"},{"salesman":"House","city":"ALBANY","custNum":2000008,"customer":"ALBANY MOTORCARS","sales2025":-100.0,"sales2026":0.0,"change":100.0,"gpPct":0.0,"action":"GROW","topDept":"None","declinedDept":"","focus":"On track"},{"salesman":"House","city":"ALBANY","custNum":200675,"customer":"AUTO SOLUTIONS LLC","sales2025":234.24,"sales2026":0.0,"change":-234.24,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $234)"},{"salesman":"House","city":"ALBANY","custNum":200411,"customer":"DJ'S CAR WASH","sales2025":3028.51,"sales2026":0.0,"change":-3028.51,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: ST TRAILER (was $114) | LOST: PASSENGER TIRES (was $1747) | LOST: RAD LT TRUCK (was $492) | LOST: FARM TIRES (was $675)"},{"salesman":"House","city":"ALBANY","custNum":200586,"customer":"STEEDLEY'S TRANSMISSION  INC","sales2025":1025.48,"sales2026":0.0,"change":-1025.48,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $1025)"},{"salesman":"House","city":"ALBANY","custNum":200136,"customer":"SUNBELT FORD ALBANY  INC (AMI)","sales2025":1481.0,"sales2026":0.0,"change":-1481.0,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1481)"},{"salesman":"House","city":"ALMA","custNum":200985,"customer":"ROBERTS AUTO SERVICE LLC","sales2025":902.32,"sales2026":0.0,"change":-902.32,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $282) | LOST: PASSENGER TIRES (was $621)"},{"salesman":"House","city":"BAINBRIDGE","custNum":200297,"customer":"WHOLESALE BATTERY","sales2025":152.36,"sales2026":0.0,"change":-152.36,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $152)"},{"salesman":"House","city":"BLACKSHEAR","custNum":200714,"customer":"C&S AUTO SERVICE INC.","sales2025":158.65,"sales2026":0.0,"change":-158.65,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TUBES","focus":"LOST: TUBES (was $129) | LOST: ST TRAILER (was $30)"},{"salesman":"House","city":"CAMILLA","custNum":101545,"customer":"TOMMY'S TIRE","sales2025":1585.52,"sales2026":0.0,"change":-1585.52,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1586) | LOST: FARM TIRES (was $1047)"},{"salesman":"House","city":"COOLIDGE","custNum":201073,"customer":"GTO TIRE SERVICE & AUTO LLC","sales2025":1058.9,"sales2026":0.0,"change":-1058.9,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $663) | LOST: RAD LT TRUCK (was $396)"},{"salesman":"House","city":"DOTHAN","custNum":200217,"customer":"SCOTT STEVENS TIRE & SERVICE","sales2025":158.78,"sales2026":0.0,"change":-158.78,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $159)"},{"salesman":"House","city":"JACKSONVILLE","custNum":200889,"customer":"CONLAN TIRE CO.","sales2025":1124.56,"sales2026":0.0,"change":-1124.56,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $1125)"},{"salesman":"House","city":"ORLANDO","custNum":2000005,"customer":"GOODYEAR COMMERCIAL TIRE & SVC","sales2025":1563.2,"sales2026":0.0,"change":-1563.2,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1563)"},{"salesman":"House","city":"PELHAM","custNum":200695,"customer":"PELHAM TIRE & EQUIPMENT COMP.","sales2025":883.24,"sales2026":0.0,"change":-883.24,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $883)"},{"salesman":"House","city":"QUITMAN","custNum":101566,"customer":"HARVEY'S GARAGE & MUFFLER","sales2025":200.0,"sales2026":0.0,"change":-200.0,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"DISCOUNTS/COUPONS","focus":"LOST: DISCOUNTS/COUPONS (was $200)"},{"salesman":"House","city":"SYLVESTER","custNum":200860,"customer":"SUNBELT FORD INC. (AMI ACCT)","sales2025":17.54,"sales2026":0.0,"change":-17.54,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $169)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200998,"customer":"FRIENDLY AUTO SALES","sales2025":541.86,"sales2026":0.0,"change":-541.86,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $542)"},{"salesman":"House","city":"TALLAHASSEE","custNum":200796,"customer":"TRUCK N CAR CONCEPTS","sales2025":1915.96,"sales2026":0.0,"change":-1915.96,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1916)"},{"salesman":"House","city":"THOMASVILLE","custNum":200641,"customer":"AUTO AIR OF THOMASVILLE","sales2025":54.85,"sales2026":0.0,"change":-54.85,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $55)"},{"salesman":"House","city":"THOMASVILLE","custNum":200526,"customer":"GERMAN IMPORT SERVICE","sales2025":70.04,"sales2026":0.0,"change":-70.04,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"ST TRAILER","focus":"LOST: TUBES (was $8) | LOST: ST TRAILER (was $62)"},{"salesman":"House","city":"VALDOSTA","custNum":200298,"customer":"WISENBAKER'S TIRE & BRAKE","sales2025":103.18,"sales2026":0.0,"change":-103.18,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"ST TRAILER","focus":"LOST: ST TRAILER (was $103)"},{"salesman":"House","city":"VALDOSTA","custNum":200300,"customer":"Z TIRE EXPRESS","sales2025":4664.9,"sales2026":0.0,"change":-4664.9,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $653) | LOST: PASSENGER TIRES (was $2100) | LOST: ST TRAILER (was $1912)"},{"salesman":"House","city":"WHIGHAM","custNum":200238,"customer":"CROSSROADS TIRE & ACC LLC","sales2025":4041.93,"sales2026":0.0,"change":-4041.93,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: FARM TIRES (was $495) | LOST: RAD LT TRUCK (was $1318) | LOST: TRUCK TIRES (was $2008) | LOST: PASSENGER TIRES (was $221)"},{"salesman":"Larry","city":"ALBANY","custNum":200288,"customer":"RIGHT CHOICE AUTO","sales2025":302.92,"sales2026":0.0,"change":-302.92,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $303)"},{"salesman":"Larry","city":"BAINBRIDGE","custNum":200184,"customer":"JIMMY'S AUTO SALES","sales2025":689.08,"sales2026":0.0,"change":-689.08,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: PASSENGER TIRES (was $439) | LOST: ST TRAILER (was $250)"},{"salesman":"Larry","city":"CAMILLA","custNum":200280,"customer":"JW PERFORMANCE & AUTO","sales2025":2819.76,"sales2026":0.0,"change":-2819.76,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2820)"},{"salesman":"Larry","city":"FITZGERALD","custNum":100993,"customer":"D & G PERFORMANCE","sales2025":2712.92,"sales2026":0.0,"change":-2712.92,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2713)"},{"salesman":"Larry","city":"SYLVESTER","custNum":2000002,"customer":"EG AGRI PARTS LLC","sales2025":16535.6,"sales2026":0.0,"change":-16535.6,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"FARM TIRES","focus":"LOST: FARM TIRES (was $12513) | LOST: TRUCK TIRES (was $4421)"},{"salesman":"Larry","city":"TIFTON","custNum":200869,"customer":"AFFORDABLE DIESEL REPAIR","sales2025":8138.56,"sales2026":0.0,"change":-8138.56,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $8021) | LOST: RAD LT TRUCK (was $118)"},{"salesman":"Larry","city":"TIFTON","custNum":201051,"customer":"EXPRESS OIL CHANGE #3168","sales2025":2913.64,"sales2026":0.0,"change":-2913.64,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1975) | LOST: PASSENGER TIRES (was $939)"},{"salesman":"Tiffany","city":"GREENVILLE","custNum":200684,"customer":"OTR SERVICES  INC.","sales2025":1913.04,"sales2026":0.0,"change":-1913.04,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $1748) | LOST: TUBES (was $165)"},{"salesman":"Tiffany","city":"LAKE BUTLER","custNum":201046,"customer":"BIELLING'S TIRE INC.","sales2025":4106.0,"sales2026":0.0,"change":-4106.0,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $4106)"},{"salesman":"Tiffany","city":"PERRY","custNum":200596,"customer":"CRIBBS TIRE","sales2025":593.18,"sales2026":0.0,"change":-593.18,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $593)"},{"salesman":"Tiffany","city":"PERRY","custNum":200698,"customer":"YARBROUGH TIRE CO.  INC.","sales2025":1755.68,"sales2026":0.0,"change":-1755.68,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"PASSENGER TIRES","focus":"LOST: RAD LT TRUCK (was $818) | LOST: PASSENGER TIRES (was $938)"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200970,"customer":"BIG PINE REPAIR","sales2025":-100.0,"sales2026":0.0,"change":100.0,"gpPct":0.0,"action":"GROW","topDept":"None","declinedDept":"","focus":"On track"},{"salesman":"Tiffany","city":"THOMASVILLE","custNum":200285,"customer":"PRECISION AUTOCRAFT INC","sales2025":577.84,"sales2026":0.0,"change":-577.84,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $578)"},{"salesman":"Unknown","city":"Unknown","custNum":500420,"customer":"BLUE LEVEL SERVICES LLC","sales2025":1683.64,"sales2026":0.0,"change":-1683.64,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $1684)"},{"salesman":"Unknown","city":"Unknown","custNum":100395,"customer":"BUBBA'S MOBILE TIRE","sales2025":1688.26,"sales2026":0.0,"change":-1688.26,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1688)"},{"salesman":"Unknown","city":"Unknown","custNum":102361,"customer":"C. MCDOWELL TRUCK PARTS & SRV.","sales2025":9600.0,"sales2026":0.0,"change":-9600.0,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"OFF THE ROAD TIRES","focus":"LOST: OFF THE ROAD TIRES (was $9600)"},{"salesman":"Unknown","city":"Unknown","custNum":102319,"customer":"D&W ELITE AUTO SALES","sales2025":1063.86,"sales2026":0.0,"change":-1063.86,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $1064)"},{"salesman":"Unknown","city":"Unknown","custNum":101115,"customer":"ECONOMY USED TIRE (WAREHOUSE)","sales2025":81136.01,"sales2026":0.0,"change":-81136.01,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $81136)"},{"salesman":"Unknown","city":"Unknown","custNum":101691,"customer":"JK MOBILE SERVICE LLC","sales2025":746.49,"sales2026":0.0,"change":-746.49,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"TRUCK TIRES","focus":"LOST: TRUCK TIRES (was $746)"},{"salesman":"Unknown","city":"Unknown","custNum":102035,"customer":"MARSHALLVILLE TIRE & LUBE","sales2025":2350.06,"sales2026":0.0,"change":-2350.06,"gpPct":0.0,"action":"LOST","topDept":"None","declinedDept":"RAD LT TRUCK","focus":"LOST: RAD LT TRUCK (was $2179) | LOST: PASSENGER TIRES (was $171)"},{"salesman":"House","city":"ADEL","custNum":200647,"customer":"FAUSETTS TIRE CO.","sales2025":19193.02,"sales2026":-104.68,"change":-19297.7,"gpPct":0.0,"action":"LOST","topDept":"PASSENGER TIRES","declinedDept":"RAD LT TRUCK","focus":"DOWN: PASSENGER TIRES (-$4336) | LOST: TUBES (was $391) | LOST: TRUCK TIRES (was $234) | LOST: LAWN & GARDEN (was $96) | LOST: FARM TIRES (was $43) | LOST: ST TRAILER (was $1813)"}],"periods":{"wtd":{"label":"WTD \u2014 W23 (Jun 2\u20136, 2026)","data":[{"custNum":200922,"customer":"ADVANCED TIRE SERVICE","sales2026":57153.95,"sales2025":8400.73,"gp2026":2805.28,"gp2025":728.43,"change":48753.22,"changePct":5.8035,"salesman":"Tiffany"},{"custNum":200953,"customer":"SIMPLE TIRE - TIFTON","sales2026":47609.16,"sales2025":32198.85,"gp2026":9347.67,"gp2025":5439.73,"change":15410.31,"changePct":0.4786,"salesman":"House"},{"custNum":200976,"customer":"TIRESEASY-LLC (TIFTON WHSE)","sales2026":30916.26,"sales2025":7758.15,"gp2026":4641.01,"gp2025":1450.95,"change":23158.11,"changePct":2.985,"salesman":"House"},{"custNum":200266,"customer":"FUSSELL TIRE & SERVICE","sales2026":28052.02,"sales2025":5070.7,"gp2026":1945.27,"gp2025":942.63,"change":22981.32,"changePct":4.5322,"salesman":"Austin"},{"custNum":200635,"customer":"DELTORO TIRE #2","sales2026":16525.99,"sales2025":16739.09,"gp2026":2100.71,"gp2025":2081.11,"change":-213.1,"changePct":-0.0127,"salesman":"Larry"},{"custNum":200410,"customer":"EZDEALIN WHEELS AND TIRES","sales2026":14284.13,"sales2025":1623.08,"gp2026":3405.67,"gp2025":490.96,"change":12661.05,"changePct":7.8006,"salesman":"House"},{"custNum":200885,"customer":"PIERCE INDUSTRIAL TIRE LLC","sales2026":10108.69,"sales2025":641.58,"gp2026":670.33,"gp2025":145.55,"change":9467.11,"changePct":14.7559,"salesman":"Tiffany"},{"custNum":200198,"customer":"MOULTRIE TIRE","sales2026":8435.55,"sales2025":5244.97,"gp2026":1417.51,"gp2025":1043.63,"change":3190.58,"changePct":0.6083,"salesman":"Larry"},{"custNum":101080,"customer":"AMERSON TIRE INC.","sales2026":8055.11,"sales2025":4226.1,"gp2026":765.75,"gp2025":455.98,"change":3829.01,"changePct":0.906,"salesman":"Larry"},{"custNum":200220,"customer":"SINGLETARY & SON TIRE CO","sales2026":6842.13,"sales2025":4530.72,"gp2026":1196.68,"gp2025":1158.7,"change":2311.41,"changePct":0.5102,"salesman":"Larry"},{"custNum":200807,"customer":"CORDELE TIRE & WHEEL LLC","sales2026":6749.19,"sales2025":3157.6,"gp2026":866.66,"gp2025":834.34,"change":3591.59,"changePct":1.1374,"salesman":"Larry"},{"custNum":200595,"customer":"W.R. WILLIAMS","sales2026":6160.08,"sales2025":2241.41,"gp2026":861.81,"gp2025":422.07,"change":3918.67,"changePct":1.7483,"salesman":"Tiffany"},{"custNum":200827,"customer":"LAKELAND TIRE DBA COOK & SONS","sales2026":5910.78,"sales2025":5958.38,"gp2026":999.12,"gp2025":1462.42,"change":-47.6,"changePct":-0.008,"salesman":"Tiffany"},{"custNum":101161,"customer":"JMC TIRE CO. INC.","sales2026":5705.4,"sales2025":18554.51,"gp2026":874.12,"gp2025":2363.74,"change":-12849.11,"changePct":-0.6925,"salesman":"Larry"},{"custNum":100282,"customer":"RUDYS TIRE SERVICE","sales2026":5337.37,"sales2025":7514.74,"gp2026":680.28,"gp2025":1216.7,"change":-2177.37,"changePct":-0.2897,"salesman":"Larry"},{"custNum":101326,"customer":"TIFTON GENERAL TIRE","sales2026":4959.04,"sales2025":12205.09,"gp2026":626.72,"gp2025":914.68,"change":-7246.05,"changePct":-0.5937,"salesman":"Larry"},{"custNum":200890,"customer":"PMT TRK. TRAILER & TIRE REPAIR","sales2026":4949.9,"sales2025":2438.7,"gp2026":691.6,"gp2025":452.61,"change":2511.2,"changePct":1.0297,"salesman":"House"},{"custNum":101463,"customer":"SHELL RAPID LUBE (FITZGERALD)","sales2026":4859.22,"sales2025":13426.67,"gp2026":1018.48,"gp2025":1072.0,"change":-8567.45,"changePct":-0.6381,"salesman":"Larry"},{"custNum":200636,"customer":"MTC SOUTH INC.","sales2026":4717.84,"sales2025":3726.48,"gp2026":713.12,"gp2025":474.3,"change":991.36,"changePct":0.266,"salesman":"Tiffany"},{"custNum":101439,"customer":"A.T. TIRE SERVICE","sales2026":4652.73,"sales2025":674.91,"gp2026":838.21,"gp2025":100.98,"change":3977.82,"changePct":5.8939,"salesman":"House"},{"custNum":200880,"customer":"R&R TIRE CO.","sales2026":4434.16,"sales2025":0,"gp2026":333.6,"gp2025":0,"change":4434.16,"changePct":0,"salesman":"House"},{"custNum":200883,"customer":"ROLLING BEAR TIRES LLC","sales2026":4374.5,"sales2025":1451.5,"gp2026":563.1,"gp2025":283.45,"change":2923.0,"changePct":2.0138,"salesman":"Tiffany"},{"custNum":201032,"customer":"A-1 TIRE PLUS","sales2026":3846.89,"sales2025":2159.15,"gp2026":468.28,"gp2025":-117.75,"change":1687.74,"changePct":0.7817,"salesman":"House"},{"custNum":101436,"customer":"EDS TIRE","sales2026":3713.64,"sales2025":2656.32,"gp2026":486.87,"gp2025":636.48,"change":1057.32,"changePct":0.398,"salesman":"Larry"},{"custNum":101512,"customer":"ELLENTON TIRE AND AUTO","sales2026":3633.86,"sales2025":6608.45,"gp2026":646.36,"gp2025":1671.74,"change":-2974.59,"changePct":-0.4501,"salesman":"House"},{"custNum":200663,"customer":"BENS TIRE & AUTO","sales2026":3600.26,"sales2025":905.7,"gp2026":901.68,"gp2025":242.68,"change":2694.56,"changePct":2.9751,"salesman":"Tiffany"},{"custNum":100301,"customer":"ALBANY GENERAL TIRE SERVICE","sales2026":3399.19,"sales2025":853.24,"gp2026":494.04,"gp2025":96.6,"change":2545.95,"changePct":2.9839,"salesman":"Larry"},{"custNum":101415,"customer":"THE TIRE STORE","sales2026":3276.8,"sales2025":3018.62,"gp2026":467.92,"gp2025":617.28,"change":258.18,"changePct":0.0855,"salesman":"Larry"},{"custNum":200891,"customer":"EJH WRECKER & TIRE SERVICE","sales2026":3150.82,"sales2025":4298.46,"gp2026":496.86,"gp2025":698.34,"change":-1147.64,"changePct":-0.267,"salesman":"Larry"},{"custNum":101283,"customer":"CAMERONS TOWING AND TIRE","sales2026":3142.41,"sales2025":2713.29,"gp2026":439.24,"gp2025":401.69,"change":429.12,"changePct":0.1582,"salesman":"Larry"},{"custNum":200906,"customer":"E&H TIRE","sales2026":2875.32,"sales2025":227.25,"gp2026":168.58,"gp2025":58.51,"change":2648.07,"changePct":11.6527,"salesman":"Tiffany"},{"custNum":101201,"customer":"TUCKERS SERVICE STATION","sales2026":2774.73,"sales2025":1283.49,"gp2026":445.29,"gp2025":72.31,"change":1491.24,"changePct":1.1619,"salesman":"House"},{"custNum":200973,"customer":"ERICS TIRE (REBEL ROAD)","sales2026":2711.37,"sales2025":1582.78,"gp2026":572.37,"gp2025":263.91,"change":1128.59,"changePct":0.713,"salesman":"Larry"},{"custNum":200971,"customer":"ALLENS TIRE","sales2026":2564.87,"sales2025":3658.71,"gp2026":402.2,"gp2025":513.82,"change":-1093.84,"changePct":-0.299,"salesman":"Larry"},{"custNum":200974,"customer":"PRIORITY TIRE (TIFTON WHSE)","sales2026":2458.0,"sales2025":0,"gp2026":942.5,"gp2025":0,"change":2458.0,"changePct":0,"salesman":"House"},{"custNum":200193,"customer":"MCLEAN TIRES INC","sales2026":2439.73,"sales2025":923.7,"gp2026":320.34,"gp2025":169.34,"change":1516.03,"changePct":1.6413,"salesman":"House"},{"custNum":200867,"customer":"24/7 DIESEL AND TIRE REPAIR","sales2026":2421.08,"sales2025":1133.4,"gp2026":255.11,"gp2025":252.4,"change":1287.68,"changePct":1.1361,"salesman":"Tiffany"},{"custNum":2000052,"customer":"L&A TIRE LLC","sales2026":2302.45,"sales2025":0,"gp2026":385.35,"gp2025":0,"change":2302.45,"changePct":0,"salesman":"House"},{"custNum":200829,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2026":2268.89,"sales2025":3922.28,"gp2026":130.3,"gp2025":430.06,"change":-1653.39,"changePct":-0.4215,"salesman":"House"},{"custNum":200270,"customer":"HAHIRA AUTOMOTIVE SERVICE","sales2026":2231.89,"sales2025":478.38,"gp2026":394.85,"gp2025":75.78,"change":1753.51,"changePct":3.6655,"salesman":"Tiffany"},{"custNum":2000044,"customer":"MURRAYS TIRE & ROAD SERVICE","sales2026":2141.97,"sales2025":0,"gp2026":213.68,"gp2025":0,"change":2141.97,"changePct":0,"salesman":"Tiffany"},{"custNum":200539,"customer":"MIKE BURCH FORD (BLACKSHEAR)","sales2026":2140.8,"sales2025":0,"gp2026":213.3,"gp2025":0,"change":2140.8,"changePct":0,"salesman":"House"},{"custNum":200690,"customer":"COLQUITT COUNTY TIRE LLC","sales2026":2034.1,"sales2025":7340.73,"gp2026":451.62,"gp2025":818.87,"change":-5306.63,"changePct":-0.7229,"salesman":"House"},{"custNum":101297,"customer":"FOUR CS LUBE","sales2026":2028.14,"sales2025":1272.83,"gp2026":236.56,"gp2025":261.34,"change":755.31,"changePct":0.5934,"salesman":"Larry"},{"custNum":200719,"customer":"FOSTER EASY PAY TIRE CO. INC.","sales2026":2020.46,"sales2025":2158.05,"gp2026":349.36,"gp2025":405.99,"change":-137.59,"changePct":-0.0638,"salesman":"House"},{"custNum":200868,"customer":"GOLDEN ENVIRONMENTAL","sales2026":1973.82,"sales2025":899.9,"gp2026":582.06,"gp2025":306.62,"change":1073.92,"changePct":1.1934,"salesman":"Larry"},{"custNum":101549,"customer":"NASHVILLE TIRE","sales2026":1951.72,"sales2025":2274.29,"gp2026":446.07,"gp2025":639.96,"change":-322.57,"changePct":-0.1418,"salesman":"House"},{"custNum":201052,"customer":"PRECISION DIESEL REPAIR LLC","sales2026":1915.42,"sales2025":0,"gp2026":296.38,"gp2025":0,"change":1915.42,"changePct":0,"salesman":"House"},{"custNum":101323,"customer":"ERICS TIRE SERVICE","sales2026":1902.44,"sales2025":4072.46,"gp2026":422.69,"gp2025":950.67,"change":-2170.02,"changePct":-0.5329,"salesman":"Larry"},{"custNum":101025,"customer":"ROCHELLE TIRE","sales2026":1866.22,"sales2025":2648.75,"gp2026":445.03,"gp2025":132.47,"change":-782.53,"changePct":-0.2954,"salesman":"House"},{"custNum":2000060,"customer":"NISSAN OF TIFTON","sales2026":1860.0,"sales2025":0,"gp2026":325.9,"gp2025":0,"change":1860.0,"changePct":0,"salesman":"House"},{"custNum":201015,"customer":"LUBE KING & TIRES","sales2026":1799.38,"sales2025":0,"gp2026":336.51,"gp2025":0,"change":1799.38,"changePct":0,"salesman":"House"},{"custNum":200293,"customer":"THOMASVILLE TIRE DEPT.","sales2026":1797.04,"sales2025":1817.07,"gp2026":333.8,"gp2025":152.82,"change":-20.03,"changePct":-0.011,"salesman":"Tiffany"},{"custNum":200474,"customer":"TIRE & WHEEL INC","sales2026":1773.44,"sales2025":0,"gp2026":53.92,"gp2025":0,"change":1773.44,"changePct":0,"salesman":"Tiffany"},{"custNum":101479,"customer":"BILL THOMPSON TIRE SERVICES","sales2026":1762.82,"sales2025":221.34,"gp2026":287.86,"gp2025":41.01,"change":1541.48,"changePct":6.9643,"salesman":"Larry"},{"custNum":200162,"customer":"DELTA TIRE CO","sales2026":1742.34,"sales2025":1970.24,"gp2026":373.96,"gp2025":197.02,"change":-227.9,"changePct":-0.1157,"salesman":"House"},{"custNum":200585,"customer":"QUINCY TIRE AND RECAPPING","sales2026":1664.67,"sales2025":2967.63,"gp2026":303.45,"gp2025":781.32,"change":-1302.96,"changePct":-0.4391,"salesman":"House"},{"custNum":101539,"customer":"COURSONS TIRE OF DOUGLAS","sales2026":1613.57,"sales2025":3062.81,"gp2026":484.29,"gp2025":551.96,"change":-1449.24,"changePct":-0.4732,"salesman":"Larry"},{"custNum":200146,"customer":"BRACEWELL AUTOMOTIVE SERVICE","sales2026":1580.0,"sales2025":0,"gp2026":45.56,"gp2025":0,"change":1580.0,"changePct":0,"salesman":"House"},{"custNum":200709,"customer":"ASHLEYS AUTOMOTIVE REPAIR","sales2026":1560.86,"sales2025":515.86,"gp2026":124.07,"gp2025":102.0,"change":1045.0,"changePct":2.0257,"salesman":"House"},{"custNum":200560,"customer":"FIVE STAR TIRE SERVICELLC","sales2026":1533.02,"sales2025":3398.24,"gp2026":331.76,"gp2025":865.27,"change":-1865.22,"changePct":-0.5489,"salesman":"Larry"},{"custNum":200315,"customer":"BOBBYS CITGO","sales2026":1440.46,"sales2025":1481.72,"gp2026":69.45,"gp2025":256.34,"change":-41.26,"changePct":-0.0278,"salesman":"House"},{"custNum":200407,"customer":"HOLLOWAY TRUCK & TRAILER REPAI","sales2026":1413.14,"sales2025":0,"gp2026":382.64,"gp2025":0,"change":1413.14,"changePct":0,"salesman":"House"},{"custNum":200918,"customer":"SOUTH GEORGIA TRUCKING SVC LLC","sales2026":1409.33,"sales2025":0,"gp2026":240.46,"gp2025":0,"change":1409.33,"changePct":0,"salesman":"House"},{"custNum":200132,"customer":"WILKS A-ONE TIRE SALES","sales2026":1390.01,"sales2025":0,"gp2026":252.15,"gp2025":0,"change":1390.01,"changePct":0,"salesman":"House"},{"custNum":200383,"customer":"WILLIAMS ALIGNMENT & TIRE","sales2026":1374.98,"sales2025":248.36,"gp2026":312.36,"gp2025":55.06,"change":1126.62,"changePct":4.5362,"salesman":"Tiffany"},{"custNum":200396,"customer":"PATE TIRE & SERVICE LLC","sales2026":1342.99,"sales2025":0,"gp2026":288.44,"gp2025":0,"change":1342.99,"changePct":0,"salesman":"House"},{"custNum":100417,"customer":"CLARK BASS SERVICE","sales2026":1324.78,"sales2025":966.41,"gp2026":152.74,"gp2025":133.33,"change":358.37,"changePct":0.3708,"salesman":"Larry"},{"custNum":200810,"customer":"BOULEVARD TIRE CENTER","sales2026":1310.0,"sales2025":822.31,"gp2026":123.18,"gp2025":202.16,"change":487.69,"changePct":0.5931,"salesman":"House"},{"custNum":200478,"customer":"82 TIRE & LUBE","sales2026":1279.87,"sales2025":1495.35,"gp2026":321.95,"gp2025":397.74,"change":-215.48,"changePct":-0.1441,"salesman":"Larry"},{"custNum":200406,"customer":"LENOX TIRE & SERVICE CENTER","sales2026":1267.23,"sales2025":2035.31,"gp2026":180.92,"gp2025":399.97,"change":-768.08,"changePct":-0.3774,"salesman":"Tiffany"},{"custNum":200836,"customer":"KING MUFFLER","sales2026":1255.04,"sales2025":0,"gp2026":421.4,"gp2025":0,"change":1255.04,"changePct":0,"salesman":"Tiffany"},{"custNum":2000014,"customer":"UNITED TIRES ONLINE SALES -T","sales2026":1188.0,"sales2025":0,"gp2026":348.0,"gp2025":0,"change":1188.0,"changePct":0,"salesman":"House"},{"custNum":200537,"customer":"TONYS TIRE & ROAD SERVICEINC","sales2026":1175.22,"sales2025":1041.34,"gp2026":185.98,"gp2025":307.88,"change":133.88,"changePct":0.1286,"salesman":"House"},{"custNum":100551,"customer":"SOUTHSIDE TIRE & AUTO SERVICE","sales2026":1128.74,"sales2025":2216.45,"gp2026":262.39,"gp2025":590.21,"change":-1087.71,"changePct":-0.4907,"salesman":"Larry"},{"custNum":101295,"customer":"DAVIDS AUTO SALES / DOUGLAS","sales2026":1110.62,"sales2025":419.1,"gp2026":296.35,"gp2025":125.8,"change":691.52,"changePct":1.65,"salesman":"Larry"},{"custNum":101477,"customer":"DISCOUNT TIRE (ALMA)OSTEEN","sales2026":1104.0,"sales2025":260.96,"gp2026":44.68,"gp2025":73.08,"change":843.04,"changePct":3.2305,"salesman":"House"},{"custNum":200673,"customer":"JBS TIRE & REPAIR SVC.","sales2026":1085.28,"sales2025":1018.36,"gp2026":96.42,"gp2025":257.62,"change":66.92,"changePct":0.0657,"salesman":"Tiffany"},{"custNum":201074,"customer":"TIRE DEPOT CO. - TAG (TIFTON)","sales2026":1079.52,"sales2025":6580.88,"gp2026":160.75,"gp2025":393.2,"change":-5501.36,"changePct":-0.836,"salesman":"House"},{"custNum":200687,"customer":"LIVE OAK TIRE CENTER LLC","sales2026":1028.51,"sales2025":0,"gp2026":124.73,"gp2025":0,"change":1028.51,"changePct":0,"salesman":"Tiffany"},{"custNum":2000025,"customer":"FORD CORDELE","sales2026":984.0,"sales2025":0,"gp2026":98.4,"gp2025":0,"change":984.0,"changePct":0,"salesman":"Car Dealer"},{"custNum":200915,"customer":"DAVIDS AUTO SALES / TIFTON","sales2026":966.16,"sales2025":491.76,"gp2026":197.92,"gp2025":135.46,"change":474.4,"changePct":0.9647,"salesman":"House"},{"custNum":500373,"customer":"TOMAHAWK TIRE (ALBANY)","sales2026":959.13,"sales2025":381.43,"gp2026":232.68,"gp2025":117.82,"change":577.7,"changePct":1.5146,"salesman":"Larry"},{"custNum":200755,"customer":"N-T TIRE SERVICE","sales2026":957.54,"sales2025":1827.6,"gp2026":248.0,"gp2025":319.27,"change":-870.06,"changePct":-0.4761,"salesman":"Larry"},{"custNum":200592,"customer":"BERNEYS TIRE SERVICE","sales2026":955.4,"sales2025":1600.86,"gp2026":182.18,"gp2025":321.04,"change":-645.46,"changePct":-0.4032,"salesman":"Larry"},{"custNum":201023,"customer":"PEASE ON THE GO 24/7","sales2026":951.64,"sales2025":0,"gp2026":129.88,"gp2025":0,"change":951.64,"changePct":0,"salesman":"House"},{"custNum":100842,"customer":"LOVE AVE. SERVICE CTR.","sales2026":936.0,"sales2025":0,"gp2026":86.72,"gp2025":0,"change":936.0,"changePct":0,"salesman":"House"},{"custNum":200831,"customer":"GRIFFIN CHRYSLER DODGE JEEP","sales2026":883.12,"sales2025":12.5,"gp2026":222.43,"gp2025":7.83,"change":870.62,"changePct":69.6496,"salesman":"Car Dealer"},{"custNum":200991,"customer":"THE TIRE CENTRE OF FLORIDA LLC","sales2026":878.36,"sales2025":1027.28,"gp2026":197.58,"gp2025":296.2,"change":-148.92,"changePct":-0.145,"salesman":"House"},{"custNum":201053,"customer":"BERNEYS TIRE SERVICE","sales2026":861.96,"sales2025":1223.2,"gp2026":2.5,"gp2025":326.8,"change":-361.24,"changePct":-0.2953,"salesman":"Larry"},{"custNum":101146,"customer":"NE-RO TIRE & BRAKE SERVICEINC","sales2026":854.99,"sales2025":0,"gp2026":163.17,"gp2025":0,"change":854.99,"changePct":0,"salesman":"Tiffany"},{"custNum":101530,"customer":"MARTIN TIRE SERVICE","sales2026":832.68,"sales2025":0,"gp2026":91.62,"gp2025":0,"change":832.68,"changePct":0,"salesman":"House"},{"custNum":101525,"customer":"BMS DISCOUNT TIRES","sales2026":753.36,"sales2025":1455.52,"gp2026":180.94,"gp2025":189.6,"change":-702.16,"changePct":-0.4824,"salesman":"House"},{"custNum":200762,"customer":"POWER MAN TIRE SHOP","sales2026":751.21,"sales2025":1179.56,"gp2026":61.05,"gp2025":287.78,"change":-428.35,"changePct":-0.3631,"salesman":"Larry"},{"custNum":2000055,"customer":"PARKER TIRE - TIFTON","sales2026":734.12,"sales2025":0,"gp2026":137.86,"gp2025":0,"change":734.12,"changePct":0,"salesman":"House"},{"custNum":101108,"customer":"PARKER TIRE DIRECT","sales2026":732.57,"sales2025":4946.07,"gp2026":140.79,"gp2025":575.12,"change":-4213.5,"changePct":-0.8519,"salesman":"House"},{"custNum":200761,"customer":"MARIO NEW AND USED TIRE SHOP","sales2026":732.42,"sales2025":0,"gp2026":231.84,"gp2025":0,"change":732.42,"changePct":0,"salesman":"House"},{"custNum":101519,"customer":"MARK TAYLOR DBA/MTAA ENT.","sales2026":728.22,"sales2025":0,"gp2026":80.56,"gp2025":0,"change":728.22,"changePct":0,"salesman":"Larry"},{"custNum":200940,"customer":"PERRY BROS. OIL (CORDELE)","sales2026":717.93,"sales2025":466.98,"gp2026":54.87,"gp2025":69.24,"change":250.95,"changePct":0.5374,"salesman":"House"},{"custNum":200754,"customer":"JORGE USED TIRE SHOP","sales2026":706.97,"sales2025":1606.14,"gp2026":197.93,"gp2025":529.52,"change":-899.17,"changePct":-0.5598,"salesman":"House"},{"custNum":200913,"customer":"LASHLEYS HOMETOWN TIRE LLC","sales2026":688.71,"sales2025":0,"gp2026":147.34,"gp2025":0,"change":688.71,"changePct":0,"salesman":"House"},{"custNum":201038,"customer":"GIGA TIRES LLC (TIFTON WHSE)","sales2026":683.36,"sales2025":5299.49,"gp2026":-13.68,"gp2025":448.59,"change":-4616.13,"changePct":-0.8711,"salesman":"House"},{"custNum":201040,"customer":"TIRES EASY (NAP - TIFTON)","sales2026":641.56,"sales2025":0,"gp2026":77.22,"gp2025":0,"change":641.56,"changePct":0,"salesman":"House"},{"custNum":200949,"customer":"SOUTHERN TIRE MART","sales2026":640.74,"sales2025":0,"gp2026":112.38,"gp2025":0,"change":640.74,"changePct":0,"salesman":"House"},{"custNum":101322,"customer":"GRIMES AUTO SERVICE","sales2026":636.0,"sales2025":653.47,"gp2026":16.52,"gp2025":71.15,"change":-17.47,"changePct":-0.0267,"salesman":"House"},{"custNum":200294,"customer":"TIRE KING OF VALDOSTA","sales2026":633.93,"sales2025":1953.24,"gp2026":162.91,"gp2025":534.53,"change":-1319.31,"changePct":-0.6754,"salesman":"Tiffany"},{"custNum":200631,"customer":"MARQUEZ TIRE SHOP","sales2026":631.5,"sales2025":1492.96,"gp2026":140.2,"gp2025":329.32,"change":-861.46,"changePct":-0.577,"salesman":"House"},{"custNum":2000069,"customer":"IRWIN COUNTY CUSTOMS & REPAIR","sales2026":626.96,"sales2025":0,"gp2026":179.3,"gp2025":0,"change":626.96,"changePct":0,"salesman":"House"},{"custNum":200679,"customer":"RICHARDS AUTO CARE & TIRE SVC","sales2026":621.15,"sales2025":636.68,"gp2026":103.7,"gp2025":192.58,"change":-15.53,"changePct":-0.0244,"salesman":"Larry"},{"custNum":200674,"customer":"DASHER LLC","sales2026":597.01,"sales2025":1040.09,"gp2026":174.83,"gp2025":179.69,"change":-443.08,"changePct":-0.426,"salesman":"House"},{"custNum":201062,"customer":"TIRE MASTERS LLC","sales2026":579.94,"sales2025":409.17,"gp2026":51.49,"gp2025":24.55,"change":170.77,"changePct":0.4174,"salesman":"Larry"},{"custNum":200655,"customer":"ECONOMY USED TIRE (ALBANY)","sales2026":577.52,"sales2025":251.26,"gp2026":175.36,"gp2025":65.08,"change":326.26,"changePct":1.2985,"salesman":"House"},{"custNum":200959,"customer":"SOUTHERN GEORGIA TIRE LLC","sales2026":576.04,"sales2025":4835.62,"gp2026":162.51,"gp2025":812.28,"change":-4259.58,"changePct":-0.8809,"salesman":"Larry"},{"custNum":201017,"customer":"PINEDAS AUTOMOTIVE","sales2026":546.42,"sales2025":918.6,"gp2026":123.56,"gp2025":75.8,"change":-372.18,"changePct":-0.4052,"salesman":"Larry"},{"custNum":201035,"customer":"DAVIDS AUTO SALES (MOULTRIE)","sales2026":525.12,"sales2025":962.8,"gp2026":123.76,"gp2025":289.78,"change":-437.68,"changePct":-0.4546,"salesman":"Larry"},{"custNum":2000039,"customer":"ADVANCED TIRE SERVICE","sales2026":521.75,"sales2025":0,"gp2026":66.05,"gp2025":0,"change":521.75,"changePct":0,"salesman":"Tiffany"},{"custNum":201037,"customer":"AFTER HOURS TIRE SERVICE","sales2026":509.88,"sales2025":1309.76,"gp2026":36.21,"gp2025":127.0,"change":-799.88,"changePct":-0.6107,"salesman":"House"},{"custNum":101066,"customer":"WARRIOR CREEK TIRE LLC","sales2026":504.6,"sales2025":2689.36,"gp2026":77.97,"gp2025":744.9,"change":-2184.76,"changePct":-0.8124,"salesman":"Tiffany"},{"custNum":200362,"customer":"RAY NORTON TIRE & AUTO","sales2026":501.52,"sales2025":1552.68,"gp2026":130.03,"gp2025":428.8,"change":-1051.16,"changePct":-0.677,"salesman":"Tiffany"},{"custNum":201003,"customer":"NICHOLAS TIRES INC.","sales2026":496.6,"sales2025":0,"gp2026":128.28,"gp2025":0,"change":496.6,"changePct":0,"salesman":"House"},{"custNum":200191,"customer":"LEE COUNTY AUTO SERVICE","sales2026":494.97,"sales2025":0,"gp2026":211.7,"gp2025":0,"change":494.97,"changePct":0,"salesman":"House"},{"custNum":201048,"customer":"RRO 24 HR ROADSIDE ASSISTANCE","sales2026":492.99,"sales2025":840.66,"gp2026":42.04,"gp2025":224.42,"change":-347.67,"changePct":-0.4136,"salesman":"House"},{"custNum":200307,"customer":"ARREDONDO TIRE SERVICE","sales2026":490.08,"sales2025":139.84,"gp2026":137.36,"gp2025":45.6,"change":350.24,"changePct":2.5046,"salesman":"House"},{"custNum":200975,"customer":"SANTOS TIRE SHOP","sales2026":485.0,"sales2025":152.66,"gp2026":119.98,"gp2025":46.35,"change":332.34,"changePct":2.177,"salesman":"House"},{"custNum":200752,"customer":"RAINEY USED CARS (ALBANY)","sales2026":484.0,"sales2025":0,"gp2026":188.08,"gp2025":0,"change":484.0,"changePct":0,"salesman":"House"},{"custNum":200979,"customer":"DISCOUNT TIRE & AUTO SHOP","sales2026":480.88,"sales2025":1469.59,"gp2026":125.36,"gp2025":341.28,"change":-988.71,"changePct":-0.6728,"salesman":"House"},{"custNum":101507,"customer":"BURNETTE AUTOMOTIVE SERVICE","sales2026":470.06,"sales2025":556.01,"gp2026":135.56,"gp2025":129.58,"change":-85.95,"changePct":-0.1546,"salesman":"Larry"},{"custNum":200365,"customer":"ROBERT HUTSON LINCOLN","sales2026":464.92,"sales2025":0,"gp2026":103.32,"gp2025":0,"change":464.92,"changePct":0,"salesman":"Car Dealer"},{"custNum":2000059,"customer":"LAKE CITY TIRE SHOP","sales2026":457.49,"sales2025":0,"gp2026":146.63,"gp2025":0,"change":457.49,"changePct":0,"salesman":"House"},{"custNum":200416,"customer":"DENTS SERVICE STATION","sales2026":451.6,"sales2025":1302.24,"gp2026":138.29,"gp2025":285.98,"change":-850.64,"changePct":-0.6532,"salesman":"House"},{"custNum":200317,"customer":"BROTHERS TIRES","sales2026":440.86,"sales2025":2295.63,"gp2026":57.11,"gp2025":269.87,"change":-1854.77,"changePct":-0.808,"salesman":"Larry"},{"custNum":200939,"customer":"RNR TIRE EXPRESS","sales2026":428.16,"sales2025":1652.38,"gp2026":251.75,"gp2025":506.77,"change":-1224.22,"changePct":-0.7409,"salesman":"House"},{"custNum":201042,"customer":"TIRE AGENT CORP (TIFTON WHS)","sales2026":411.3,"sales2025":0,"gp2026":54.14,"gp2025":0,"change":411.3,"changePct":0,"salesman":"House"},{"custNum":2000053,"customer":"JOINER CONTRACTING","sales2026":403.1,"sales2025":0,"gp2026":48.8,"gp2025":0,"change":403.1,"changePct":0,"salesman":"House"},{"custNum":2000015,"customer":"EDWINS TIRES LLC","sales2026":402.92,"sales2025":0,"gp2026":92.48,"gp2025":0,"change":402.92,"changePct":0,"salesman":"Larry"},{"custNum":200760,"customer":"MALLARDS SERVICE CENTER","sales2026":400.19,"sales2025":0,"gp2026":49.49,"gp2025":0,"change":400.19,"changePct":0,"salesman":"Larry"},{"custNum":200242,"customer":"THOMAS TIRE COMPANY LLC","sales2026":395.51,"sales2025":-500.0,"gp2026":112.69,"gp2025":-22.8,"change":895.51,"changePct":0,"salesman":"Larry"},{"custNum":201014,"customer":"D&R AUTO SALES & SALVAGE PARTS","sales2026":388.28,"sales2025":61.96,"gp2026":98.24,"gp2025":18.37,"change":326.32,"changePct":5.2666,"salesman":"Larry"},{"custNum":200319,"customer":"BUCKS AUTO REPAIR","sales2026":374.08,"sales2025":-549.64,"gp2026":112.6,"gp2025":-181.8,"change":923.72,"changePct":0,"salesman":"Larry"},{"custNum":200277,"customer":"IMPORT SERVICE & SALES","sales2026":372.18,"sales2025":148.3,"gp2026":76.88,"gp2025":45.08,"change":223.88,"changePct":1.5096,"salesman":"House"},{"custNum":201031,"customer":"RENOS QUALITY COLLISION","sales2026":362.0,"sales2025":0,"gp2026":-1.18,"gp2025":0,"change":362.0,"changePct":0,"salesman":"House"},{"custNum":2000049,"customer":"BEALL TIRE WHOLESALE LLC","sales2026":349.54,"sales2025":0,"gp2026":66.32,"gp2025":0,"change":349.54,"changePct":0,"salesman":"House"},{"custNum":200803,"customer":"SOUTH MAIN GARAGE","sales2026":347.8,"sales2025":180.06,"gp2026":77.36,"gp2025":66.18,"change":167.74,"changePct":0.9316,"salesman":"House"},{"custNum":201070,"customer":"SNIDER INDUSTRIAL","sales2026":345.8,"sales2025":0,"gp2026":83.34,"gp2025":0,"change":345.8,"changePct":0,"salesman":"House"},{"custNum":200456,"customer":"PRINCE TOYOTA","sales2026":342.04,"sales2025":2795.96,"gp2026":91.28,"gp2025":150.96,"change":-2453.92,"changePct":-0.8777,"salesman":"Car Dealer"},{"custNum":200717,"customer":"JW AUTOMOTIVE","sales2026":335.52,"sales2025":219.44,"gp2026":69.64,"gp2025":64.04,"change":116.08,"changePct":0.529,"salesman":"House"},{"custNum":200439,"customer":"BEAR TIRE SERVICE","sales2026":329.93,"sales2025":0,"gp2026":87.0,"gp2025":0,"change":329.93,"changePct":0,"salesman":"Tiffany"},{"custNum":200455,"customer":"PRINCE HONDA","sales2026":306.0,"sales2025":0,"gp2026":62.64,"gp2025":0,"change":306.0,"changePct":0,"salesman":"Car Dealer"},{"custNum":200681,"customer":"CITY OF SYLVESTER","sales2026":301.88,"sales2025":0,"gp2026":80.24,"gp2025":0,"change":301.88,"changePct":0,"salesman":"House"},{"custNum":101999,"customer":"SIMPLE TIRE (NAT. ACCT. D.R.)","sales2026":298.08,"sales2025":1586.44,"gp2026":12.84,"gp2025":89.7,"change":-1288.36,"changePct":-0.8121,"salesman":"House"},{"custNum":102273,"customer":"COMPLETE TIRE & SVC (CORDELE)","sales2026":292.99,"sales2025":1516.43,"gp2026":71.32,"gp2025":92.55,"change":-1223.44,"changePct":-0.8068,"salesman":"House"},{"custNum":200562,"customer":"BEST CARS OF CORDELE LLC","sales2026":286.8,"sales2025":0,"gp2026":104.92,"gp2025":0,"change":286.8,"changePct":0,"salesman":"House"},{"custNum":200451,"customer":"HARROD BROTHERS","sales2026":283.02,"sales2025":0,"gp2026":113.02,"gp2025":0,"change":283.02,"changePct":0,"salesman":"House"},{"custNum":2000013,"customer":"TIRE SOLUTIONS & VEH. REPAIRS","sales2026":275.22,"sales2025":0,"gp2026":89.16,"gp2025":0,"change":275.22,"changePct":0,"salesman":"House"},{"custNum":101466,"customer":"WATTS REPAIR SERVICE","sales2026":271.84,"sales2025":0,"gp2026":39.88,"gp2025":0,"change":271.84,"changePct":0,"salesman":"House"},{"custNum":200947,"customer":"SUNSET TIRE & AUTOMOTIVE","sales2026":268.0,"sales2025":279.68,"gp2026":110.2,"gp2025":91.2,"change":-11.68,"changePct":-0.0418,"salesman":"House"},{"custNum":2000042,"customer":"ROJAS AUTO REPAIR","sales2026":264.86,"sales2025":0,"gp2026":-70.64,"gp2025":0,"change":264.86,"changePct":0,"salesman":"House"},{"custNum":101305,"customer":"SYCAMORE SALES & SALVAGE LLC","sales2026":264.08,"sales2025":333.6,"gp2026":54.24,"gp2025":129.64,"change":-69.52,"changePct":-0.2084,"salesman":"House"},{"custNum":200533,"customer":"FLOWERS IMPORTS LLC (HONDA)","sales2026":248.0,"sales2025":0,"gp2026":13.4,"gp2025":0,"change":248.0,"changePct":0,"salesman":"Car Dealer"},{"custNum":200546,"customer":"SOWEGA TIRE OF ALBANY","sales2026":241.4,"sales2025":0,"gp2026":75.36,"gp2025":0,"change":241.4,"changePct":0,"salesman":"House"},{"custNum":101181,"customer":"SOUTH GA LUBE CENTER","sales2026":222.78,"sales2025":0,"gp2026":38.68,"gp2025":0,"change":222.78,"changePct":0,"salesman":"Larry"},{"custNum":200866,"customer":"LEMUS TIRE SHOP","sales2026":219.38,"sales2025":459.22,"gp2026":41.78,"gp2025":154.08,"change":-239.84,"changePct":-0.5223,"salesman":"House"},{"custNum":201043,"customer":"BILLS TRAILER SERVICE","sales2026":219.1,"sales2025":0,"gp2026":18.33,"gp2025":0,"change":219.1,"changePct":0,"salesman":"Larry"},{"custNum":200933,"customer":"GRIFFIN CDJR VALDOSTA","sales2026":214.62,"sales2025":0,"gp2026":57.58,"gp2025":0,"change":214.62,"changePct":0,"salesman":"Car Dealer"},{"custNum":200765,"customer":"ALBANY CHRYSLER DODGE JEEP RAM","sales2026":208.0,"sales2025":931.64,"gp2026":40.48,"gp2025":239.64,"change":-723.64,"changePct":-0.7767,"salesman":"Car Dealer"},{"custNum":2000020,"customer":"T&D TIRE","sales2026":207.56,"sales2025":0,"gp2026":51.86,"gp2025":0,"change":207.56,"changePct":0,"salesman":"Larry"},{"custNum":2000007,"customer":"KEATON & SON TIRE LLC","sales2026":206.64,"sales2025":121.95,"gp2026":45.27,"gp2025":43.17,"change":84.69,"changePct":0.6945,"salesman":"House"},{"custNum":200795,"customer":"WINCHESTER PAINT & BODY","sales2026":203.84,"sales2025":219.96,"gp2026":59.88,"gp2025":77.68,"change":-16.12,"changePct":-0.0733,"salesman":"House"},{"custNum":200896,"customer":"SOUTHERN AUTO SPECIALIST","sales2026":196.0,"sales2025":868.2,"gp2026":29.72,"gp2025":248.88,"change":-672.2,"changePct":-0.7742,"salesman":"House"},{"custNum":200628,"customer":"SOUTH GEORGIA TIRE","sales2026":189.13,"sales2025":3424.71,"gp2026":54.31,"gp2025":514.43,"change":-3235.58,"changePct":-0.9448,"salesman":"House"},{"custNum":200475,"customer":"AZALEA CITY AUTO SALES/SERVICE","sales2026":177.28,"sales2025":242.41,"gp2026":43.56,"gp2025":79.55,"change":-65.13,"changePct":-0.2687,"salesman":"House"},{"custNum":201039,"customer":"CERVANTES AUTO SALES","sales2026":172.5,"sales2025":0,"gp2026":45.24,"gp2025":0,"change":172.5,"changePct":0,"salesman":"House"},{"custNum":2000022,"customer":"R&R AUTO SERVICE & REPAIR","sales2026":163.3,"sales2025":0,"gp2026":56.8,"gp2025":0,"change":163.3,"changePct":0,"salesman":"Larry"},{"custNum":200980,"customer":"SLYDERS GARAGE","sales2026":156.0,"sales2025":118.68,"gp2026":31.06,"gp2025":34.77,"change":37.32,"changePct":0.3145,"salesman":"House"},{"custNum":2000009,"customer":"KING FORD OF NASHVILLE","sales2026":149.14,"sales2025":306.84,"gp2026":39.18,"gp2025":61.36,"change":-157.7,"changePct":-0.5139,"salesman":"Car Dealer"},{"custNum":200327,"customer":"DRAPER TIRES & AUTOMOTIVE","sales2026":146.06,"sales2025":549.8,"gp2026":45.75,"gp2025":51.84,"change":-403.74,"changePct":-0.7343,"salesman":"House"},{"custNum":200664,"customer":"ABR COMMERCIAL TRUCK & AUTO","sales2026":145.13,"sales2025":247.51,"gp2026":40.59,"gp2025":38.1,"change":-102.38,"changePct":-0.4136,"salesman":"House"},{"custNum":101524,"customer":"MASTER BODY WORKS","sales2026":123.52,"sales2025":0,"gp2026":31.3,"gp2025":0,"change":123.52,"changePct":0,"salesman":"House"},{"custNum":201008,"customer":"DBJ MOBILE TIRE SERVICE INC.","sales2026":118.86,"sales2025":1550.32,"gp2026":29.89,"gp2025":389.53,"change":-1431.46,"changePct":-0.9233,"salesman":"House"},{"custNum":200715,"customer":"R&M AUTO TRUCKING INC","sales2026":118.8,"sales2025":266.28,"gp2026":19.16,"gp2025":88.6,"change":-147.48,"changePct":-0.5539,"salesman":"House"},{"custNum":100741,"customer":"MASSEYS MUFFLER","sales2026":118.0,"sales2025":0,"gp2026":5.02,"gp2025":0,"change":118.0,"changePct":0,"salesman":"House"},{"custNum":101544,"customer":"GODWIN TIRE & AUTO","sales2026":117.23,"sales2025":0,"gp2026":38.54,"gp2025":0,"change":117.23,"changePct":0,"salesman":"Tiffany"},{"custNum":200388,"customer":"PEARSON TIRE & LUBE","sales2026":109.22,"sales2025":0,"gp2026":29.94,"gp2025":0,"change":109.22,"changePct":0,"salesman":"Larry"},{"custNum":200718,"customer":"JOBBER ACCT (TIFTON)","sales2026":101.48,"sales2025":390.16,"gp2026":43.6,"gp2025":62.82,"change":-288.68,"changePct":-0.7399,"salesman":"House"},{"custNum":200676,"customer":"ALL PURPOSE AUTO CENTER","sales2026":98.88,"sales2025":0,"gp2026":38.78,"gp2025":0,"change":98.88,"changePct":0,"salesman":"Larry"},{"custNum":200468,"customer":"MCKEES AUTO CENTER INC","sales2026":78.06,"sales2025":0,"gp2026":20.18,"gp2025":0,"change":78.06,"changePct":0,"salesman":"Larry"},{"custNum":200543,"customer":"GATOR TIRE","sales2026":63.24,"sales2025":0,"gp2026":16.52,"gp2025":0,"change":63.24,"changePct":0,"salesman":"Tiffany"},{"custNum":200972,"customer":"ERICS TIRE OF SYLVESTER","sales2026":62.0,"sales2025":181.65,"gp2026":15.28,"gp2025":47.32,"change":-119.65,"changePct":-0.6587,"salesman":"Larry"},{"custNum":200759,"customer":"AADCO","sales2026":61.95,"sales2025":0,"gp2026":15.12,"gp2025":0,"change":61.95,"changePct":0,"salesman":"House"},{"custNum":100107,"customer":"JOHNSON AUTO & TIRE","sales2026":60.34,"sales2025":758.04,"gp2026":17.06,"gp2025":231.13,"change":-697.7,"changePct":-0.9204,"salesman":"Larry"},{"custNum":200300,"customer":"Z TIRE EXPRESS","sales2026":0,"sales2025":-296.0,"gp2026":0,"gp2025":-79.04,"change":296.0,"changePct":0,"salesman":"House"},{"custNum":200916,"customer":"GATEWAY DIESEL AUTO & MOBILE","sales2026":0,"sales2025":288.0,"gp2026":0,"gp2025":96.0,"change":-288.0,"changePct":-1.0,"salesman":"House"},{"custNum":200941,"customer":"D&S WHEELS & DEALS LLC","sales2026":0,"sales2025":1578.81,"gp2026":0,"gp2025":81.13,"change":-1578.81,"changePct":-1.0,"salesman":"House"},{"custNum":101371,"customer":"DIRTY SOUTH KUSTOMS","sales2026":0,"sales2025":639.88,"gp2026":0,"gp2025":37.28,"change":-639.88,"changePct":-1.0,"salesman":"House"},{"custNum":200668,"customer":"24/7 TIRE","sales2026":0,"sales2025":173.38,"gp2026":0,"gp2025":66.0,"change":-173.38,"changePct":-1.0,"salesman":"House"},{"custNum":100967,"customer":"SOUTHEASTERN COMMERCIAL TIRE","sales2026":0,"sales2025":419.16,"gp2026":0,"gp2025":62.9,"change":-419.16,"changePct":-1.0,"salesman":"House"},{"custNum":200386,"customer":"SHANE'S TIRE & AUTO","sales2026":0,"sales2025":217.1,"gp2026":0,"gp2025":71.34,"change":-217.1,"changePct":-1.0,"salesman":"House"},{"custNum":200331,"customer":"BUDGET CAR SALES","sales2026":0,"sales2025":4241.24,"gp2026":0,"gp2025":991.29,"change":-4241.24,"changePct":-1.0,"salesman":"House"},{"custNum":201055,"customer":"HERNANDEZ TIRES SHOP","sales2026":0,"sales2025":400.0,"gp2026":0,"gp2025":31.32,"change":-400.0,"changePct":-1.0,"salesman":"House"},{"custNum":200352,"customer":"MALUDA AUTO SALES","sales2026":0,"sales2025":190.48,"gp2026":0,"gp2025":57.68,"change":-190.48,"changePct":-1.0,"salesman":"House"},{"custNum":200920,"customer":"PRECISION MAINTENANCE","sales2026":0,"sales2025":2369.32,"gp2026":0,"gp2025":292.76,"change":-2369.32,"changePct":-1.0,"salesman":"House"},{"custNum":200287,"customer":"RIDLEY'S AUTOMOTIVE","sales2026":0,"sales2025":240.0,"gp2026":0,"gp2025":80.0,"change":-240.0,"changePct":-1.0,"salesman":"House"},{"custNum":201016,"customer":"O&C AUTO REPAIR","sales2026":0,"sales2025":103.08,"gp2026":0,"gp2025":31.62,"change":-103.08,"changePct":-1.0,"salesman":"House"},{"custNum":101491,"customer":"QUALITY AUTO & R.V. SERVICE","sales2026":0,"sales2025":236.08,"gp2026":0,"gp2025":74.88,"change":-236.08,"changePct":-1.0,"salesman":"House"},{"custNum":200599,"customer":"W&L TIRE & WHEEL CO. INC.","sales2026":0,"sales2025":332.71,"gp2026":0,"gp2025":89.27,"change":-332.71,"changePct":-1.0,"salesman":"House"},{"custNum":200935,"customer":"FROMETA USED CAR & TIRE CENTER","sales2026":0,"sales2025":625.44,"gp2026":0,"gp2025":151.32,"change":-625.44,"changePct":-1.0,"salesman":"House"},{"custNum":101298,"customer":"TENNESON NISSAN","sales2026":0,"sales2025":462.16,"gp2026":0,"gp2025":81.78,"change":-462.16,"changePct":-1.0,"salesman":"House"},{"custNum":200712,"customer":"TANNER AUTO REPAIR PLUS, LLC","sales2026":0,"sales2025":1066.54,"gp2026":0,"gp2025":351.07,"change":-1066.54,"changePct":-1.0,"salesman":"House"},{"custNum":201028,"customer":"CLEMENT USED TIRES","sales2026":0,"sales2025":1339.96,"gp2026":0,"gp2025":394.37,"change":-1339.96,"changePct":-1.0,"salesman":"House"},{"custNum":101366,"customer":"B & M AUTOMOTIVE SERVICE","sales2026":0,"sales2025":206.08,"gp2026":0,"gp2025":38.44,"change":-206.08,"changePct":-1.0,"salesman":"House"},{"custNum":200793,"customer":"ALL SEASON AUTO REPAIR","sales2026":0,"sales2025":987.28,"gp2026":0,"gp2025":98.72,"change":-987.28,"changePct":-1.0,"salesman":"House"},{"custNum":200833,"customer":"FITZGERALD CHRYSLER DODGE RAM","sales2026":0,"sales2025":665.24,"gp2026":0,"gp2025":103.96,"change":-665.24,"changePct":-1.0,"salesman":"House"},{"custNum":201025,"customer":"A-1 WRECKER SERVICE","sales2026":0,"sales2025":330.24,"gp2026":0,"gp2025":105.88,"change":-330.24,"changePct":-1.0,"salesman":"House"},{"custNum":200606,"customer":"AFFORDABLE TIRE SERVICE LLC","sales2026":0,"sales2025":828.9,"gp2026":0,"gp2025":266.36,"change":-828.9,"changePct":-1.0,"salesman":"House"},{"custNum":200411,"customer":"DJ'S CAR WASH","sales2026":0,"sales2025":390.2,"gp2026":0,"gp2025":113.16,"change":-390.2,"changePct":-1.0,"salesman":"House"},{"custNum":200647,"customer":"FAUSETTS TIRE CO.","sales2026":0,"sales2025":1310.86,"gp2026":0,"gp2025":117.93,"change":-1310.86,"changePct":-1.0,"salesman":"House"},{"custNum":200804,"customer":"FIVE STAR TIRE ****ADEL****","sales2026":0,"sales2025":2780.48,"gp2026":0,"gp2025":711.58,"change":-2780.48,"changePct":-1.0,"salesman":"House"},{"custNum":101513,"customer":"GAY'S TIRE SERVICE","sales2026":0,"sales2025":131.39,"gp2026":0,"gp2025":46.71,"change":-131.39,"changePct":-1.0,"salesman":"House"},{"custNum":200621,"customer":"SOUTH GA LUBE (FITZGERALD)","sales2026":0,"sales2025":194.2,"gp2026":0,"gp2025":49.05,"change":-194.2,"changePct":-1.0,"salesman":"House"},{"custNum":200503,"customer":"BBS AUTOMOTIVE","sales2026":0,"sales2025":65.66,"gp2026":0,"gp2025":19.95,"change":-65.66,"changePct":-1.0,"salesman":"House"},{"custNum":200588,"customer":"LANGDALE HYUNDAI OF SOUTH GA","sales2026":0,"sales2025":51.54,"gp2026":0,"gp2025":15.81,"change":-51.54,"changePct":-1.0,"salesman":"House"},{"custNum":100907,"customer":"SMITHS DIESEL REPAIR","sales2026":0,"sales2025":532.6,"gp2026":0,"gp2025":103.18,"change":-532.6,"changePct":-1.0,"salesman":"House"},{"custNum":200348,"customer":"LANGDALE KIA (SERVICE)","sales2026":0,"sales2025":244.28,"gp2026":0,"gp2025":80.0,"change":-244.28,"changePct":-1.0,"salesman":"House"},{"custNum":101208,"customer":"MIKE FRASER AUTO REPAIR","sales2026":0,"sales2025":1030.88,"gp2026":0,"gp2025":157.14,"change":-1030.88,"changePct":-1.0,"salesman":"House"},{"custNum":200744,"customer":"TIFTON COMMERCIAL","sales2026":0,"sales2025":64.67,"gp2026":0,"gp2025":14.72,"change":-64.67,"changePct":-1.0,"salesman":"House"},{"custNum":201036,"customer":"GODWIN & SON REPAIR & SALES","sales2026":0,"sales2025":242.24,"gp2026":0,"gp2025":72.28,"change":-242.24,"changePct":-1.0,"salesman":"House"},{"custNum":200232,"customer":"SUNBELT FORD INC","sales2026":0,"sales2025":350.76,"gp2026":0,"gp2025":119.0,"change":-350.76,"changePct":-1.0,"salesman":"House"},{"custNum":200148,"customer":"AUTO & TRUCK CARE SPECIAL","sales2026":0,"sales2025":1029.16,"gp2026":0,"gp2025":195.84,"change":-1029.16,"changePct":-1.0,"salesman":"House"},{"custNum":200366,"customer":"ROUNTREE PERFORMANCE","sales2026":0,"sales2025":465.9,"gp2026":0,"gp2025":109.44,"change":-465.9,"changePct":-1.0,"salesman":"House"},{"custNum":100591,"customer":"RAINEY ALIGNMENT","sales2026":0,"sales2025":503.92,"gp2026":0,"gp2025":72.52,"change":-503.92,"changePct":-1.0,"salesman":"House"},{"custNum":200734,"customer":"ANDERSON FORD","sales2026":0,"sales2025":2541.76,"gp2026":0,"gp2025":224.64,"change":-2541.76,"changePct":-1.0,"salesman":"House"},{"custNum":101411,"customer":"LONE MOUNTAIN TRUCK LEASING","sales2026":0,"sales2025":208.48,"gp2026":0,"gp2025":50.44,"change":-208.48,"changePct":-1.0,"salesman":"House"},{"custNum":200527,"customer":"GRIFFIN CHRY/DOD/JEEP/RAM","sales2026":0,"sales2025":1509.28,"gp2026":0,"gp2025":344.48,"change":-1509.28,"changePct":-1.0,"salesman":"House"},{"custNum":200901,"customer":"LENCHOS & SON TIRE SHOP","sales2026":0,"sales2025":702.34,"gp2026":0,"gp2025":232.98,"change":-702.34,"changePct":-1.0,"salesman":"House"},{"custNum":3000389,"customer":"SOUTHERN TIRE MART LLC (#134)","sales2026":0,"sales2025":-1218.22,"gp2026":0,"gp2025":1.59,"change":1218.22,"changePct":0,"salesman":"House"},{"custNum":200706,"customer":"CARL'S SERVICE STATION LLC","sales2026":0,"sales2025":859.24,"gp2026":0,"gp2025":253.6,"change":-859.24,"changePct":-1.0,"salesman":"House"},{"custNum":200932,"customer":"RNR TIRE EXPRESS","sales2026":0,"sales2025":84.65,"gp2026":0,"gp2025":16.93,"change":-84.65,"changePct":-1.0,"salesman":"House"},{"custNum":101164,"customer":"DAVIS TIRE (DOUGLAS)","sales2026":0,"sales2025":603.2,"gp2026":0,"gp2025":186.43,"change":-603.2,"changePct":-1.0,"salesman":"House"},{"custNum":200753,"customer":"AUTO TECH OF MIAMI INC.","sales2026":0,"sales2025":208.48,"gp2026":0,"gp2025":50.44,"change":-208.48,"changePct":-1.0,"salesman":"House"},{"custNum":101152,"customer":"BROOKS BODY SHOP","sales2026":0,"sales2025":278.24,"gp2026":0,"gp2025":92.6,"change":-278.24,"changePct":-1.0,"salesman":"House"},{"custNum":200166,"customer":"EDISON TIRE","sales2026":0,"sales2025":4886.58,"gp2026":0,"gp2025":1156.33,"change":-4886.58,"changePct":-1.0,"salesman":"House"},{"custNum":200720,"customer":"LUBE MASTERS","sales2026":0,"sales2025":118.18,"gp2026":0,"gp2025":27.78,"change":-118.18,"changePct":-1.0,"salesman":"House"},{"custNum":101300,"customer":"FITZGERALD FORD AND LINCOLN","sales2026":0,"sales2025":123.92,"gp2026":0,"gp2025":36.74,"change":-123.92,"changePct":-1.0,"salesman":"House"},{"custNum":200239,"customer":"T & S TIRE","sales2026":-77.28,"sales2025":208.48,"gp2026":5.35,"gp2025":50.44,"change":-285.76,"changePct":-1.3707,"salesman":"Larry"},{"custNum":200750,"customer":"LAWN PERFORMANCE LLC","sales2026":-189.0,"sales2025":0,"gp2026":-52.86,"gp2025":0,"change":-189.0,"changePct":0,"salesman":"House"}]}},"apVersion":"W23"};

const SEED_AR = [{"salesman":"Larry","custNum":200635,"shortName":"DEL","name":"DELTORO TIRE #2","phone":"229-396-5510","balance":86292.47,"futDue":17854.66,"curDue":2449.22,"due1_30":65988.59,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"Tiffany","custNum":200636,"shortName":"MTC","name":"MTC SOUTH, INC.","phone":"850-251-5393","balance":50405.03,"futDue":4717.84,"curDue":12392.94,"due1_30":21737.52,"due31_60":11556.73,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":101512,"shortName":"ELLENTON","name":"ELLENTON TIRE AND AUTO","phone":"229-324-2475","balance":44577.46,"futDue":3960.53,"curDue":24271.91,"due1_30":16345.02,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200268,"shortName":"GIANT","name":"GIANT TIRE SALES/SERVICE","phone":"229-762-3230","balance":7626.91,"futDue":0.0,"curDue":678.91,"due1_30":6948.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14 00:00:00"},{"salesman":"Larry","custNum":100967,"shortName":"SOUTHEAST","name":"SOUTHEASTERN COMMERCIAL TIRE","phone":"229-888-3300","balance":14829.76,"futDue":0.0,"curDue":8711.34,"due1_30":6118.42,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-10 00:00:00"},{"salesman":"Larry","custNum":200198,"shortName":"MOULTRIE","name":"MOULTRIE TIRE","phone":"229-985-5619","balance":34076.08,"futDue":8923.73,"curDue":10140.19,"due1_30":3814.34,"due31_60":11197.82,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Larry","custNum":200868,"shortName":"GOLDEN","name":"GOLDEN ENVIRONMENTAL","phone":"229-382-0309","balance":9727.92,"futDue":1891.1,"curDue":4032.02,"due1_30":3804.8,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"Larry","custNum":100282,"shortName":"RUDY","name":"RUDY'S TIRE SERVICE","phone":"229-382-5324","balance":27689.6,"futDue":5494.03,"curDue":18818.0,"due1_30":3377.57,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200991,"shortName":"THE","name":"THE TIRE CENTRE OF FLORIDA LLC","phone":"850-671-4181","balance":6890.48,"futDue":878.36,"curDue":3012.12,"due1_30":3000.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"Tiffany","custNum":200293,"shortName":"THOM","name":"THOMASVILLE TIRE DEPT.","phone":"229-228-0260","balance":2719.3,"futDue":54.46,"curDue":34.36,"due1_30":2603.92,"due31_60":26.56,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":101588,"shortName":"STEPH","name":"STEPHENS BROTHERS","phone":"229-425-1055","balance":5457.46,"futDue":0.0,"curDue":3441.22,"due1_30":2016.24,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-01 00:00:00"},{"salesman":"Larry","custNum":201062,"shortName":"TIRE","name":"TIRE MASTERS LLC","phone":"229-445-7500","balance":2851.19,"futDue":582.94,"curDue":597.35,"due1_30":1670.9,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"Larry","custNum":200409,"shortName":"MASTER","name":"MASTER CRAFT IND.(NO PASS/LT)","phone":"229-386-0610","balance":2517.56,"futDue":0.0,"curDue":1343.82,"due1_30":1173.74,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22 00:00:00"},{"salesman":"House","custNum":200829,"shortName":"TIRE","name":"TIRE SOLUTIONS & VEH. REPAIRS","phone":"229-985-8473","balance":43414.79,"futDue":1101.19,"curDue":653.93,"due1_30":664.58,"due31_60":676.56,"due61_90":699.74,"dueOver90":39618.79,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Car Dealer","custNum":2000024,"shortName":"FORKLIFT","name":"FORKLIFT TIRE OF CENTRAL FL","phone":"863-559-9353","balance":623.9,"futDue":0.0,"curDue":9.22,"due1_30":614.68,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14 00:00:00"},{"salesman":"Tiffany","custNum":2000028,"shortName":"SUNPOINT","name":"SUNPOINT TIRES & ROAD SERVICE","phone":"863-272-8823","balance":485.74,"futDue":0.0,"curDue":7.18,"due1_30":478.56,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-06 00:00:00"},{"salesman":"House","custNum":2000013,"shortName":"TIRE","name":"TIRE SOLUTIONS & VEH. REPAIRS","phone":"229-294-2801","balance":31897.83,"futDue":464.63,"curDue":458.12,"due1_30":467.24,"due31_60":489.36,"due61_90":958.54,"dueOver90":29059.94,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200427,"shortName":"SAUNDERS","name":"SAUNDERS AUTO REPAIR","phone":"229-616-1041","balance":331.55,"futDue":20.59,"curDue":0.0,"due1_30":310.96,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-30 00:00:00"},{"salesman":"House","custNum":200942,"shortName":"RNR","name":"RNR TIRE EXPRESS","phone":"813-977-9800","balance":14163.17,"futDue":290.94,"curDue":13584.75,"due1_30":287.48,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":200214,"shortName":"PETERSON","name":"PETERSON TIRE & AUTO CENTER","phone":"229-352-5136","balance":25.1,"futDue":-255.27,"curDue":0.37,"due1_30":280.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Larry","custNum":100551,"shortName":"SOUTH","name":"SOUTHSIDE TIRE & AUTO SERVICE","phone":"229-387-6283","balance":147.4,"futDue":-1.34,"curDue":0.0,"due1_30":148.74,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":101241,"shortName":"PETTIS","name":"NEW PETTIS TIRE","phone":"229-273-1153","balance":6758.43,"futDue":0.0,"curDue":195.72,"due1_30":77.2,"due31_60":1338.99,"due61_90":3573.36,"dueOver90":1573.16,"lastPaid":"2026-02-03 00:00:00"},{"salesman":"House","custNum":200238,"shortName":"CROSSROADS","name":"CROSSROADS TIRE & ACC LLC","phone":"nan","balance":3832.3,"futDue":0.0,"curDue":110.83,"due1_30":54.21,"due31_60":53.42,"due61_90":52.64,"dueOver90":3561.2,"lastPaid":"2025-10-07 00:00:00"},{"salesman":"House","custNum":200718,"shortName":"JOBBER #2","name":"JOBBER ACCT (TIFTON)","phone":"nan","balance":115.16,"futDue":110.6,"curDue":0.0,"due1_30":4.56,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-22 00:00:00"},{"salesman":"House","custNum":2000038,"shortName":"FAST","name":"FAST TIRE SERVICE","phone":"386-324-1853","balance":1.14,"futDue":0.0,"curDue":0.02,"due1_30":1.12,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-13 00:00:00"},{"salesman":"House","custNum":201056,"shortName":"LEE","name":"LEE'S AUTO SHOP","phone":"813-368-4143","balance":853.66,"futDue":0.0,"curDue":24.86,"due1_30":0.0,"due31_60":828.8,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-20 00:00:00"},{"salesman":"House","custNum":201040,"shortName":"TIRES","name":"TIRES EASY (NAP - TIFTON)","phone":"844-347-0789","balance":3334.69,"futDue":643.56,"curDue":94.18,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":2596.95,"lastPaid":""},{"salesman":"Tiffany","custNum":200867,"shortName":"24","name":"24/7 DIESEL AND TIRE REPAIR","phone":"229-375-8667","balance":2431.78,"futDue":2277.21,"curDue":154.57,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"House","custNum":200668,"shortName":"247","name":"24/7 TIRE","phone":"229-247-8473","balance":2348.81,"futDue":0.19,"curDue":2348.62,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Larry","custNum":200478,"shortName":"82","name":"82 TIRE & LUBE","phone":"912-462-7357","balance":-1205.11,"futDue":-1205.11,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":201032,"shortName":"A1","name":"A-1 TIRE PLUS","phone":"386-867-5495","balance":1117.44,"futDue":1117.44,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200759,"shortName":"AADCO","name":"AADCO","phone":"229-382-8080","balance":62.95,"futDue":62.95,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"House","custNum":200664,"shortName":"ABR","name":"ABR COMMERCIAL TRUCK & AUTO","phone":"229-995-2169","balance":73.39,"futDue":73.39,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200401,"shortName":"ADEL","name":"ADEL TIRE CO","phone":"229-896-3086","balance":-58.31,"futDue":-58.31,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"Tiffany","custNum":200922,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-236-8825","balance":168352.22,"futDue":124938.78,"curDue":43413.44,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21 00:00:00"},{"salesman":"Tiffany","custNum":200923,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"386-406-6745","balance":838.12,"futDue":321.82,"curDue":516.3,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21 00:00:00"},{"salesman":"Tiffany","custNum":2000021,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-559-0708","balance":2280.24,"futDue":2280.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21 00:00:00"},{"salesman":"Tiffany","custNum":2000039,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-236-8825","balance":1378.44,"futDue":1378.44,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Tiffany","custNum":2000040,"shortName":"ADVANCED","name":"ADVANCED TIRE SERVICE","phone":"352-691-7771","balance":2890.26,"futDue":531.73,"curDue":2358.53,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-21 00:00:00"},{"salesman":"House","custNum":201037,"shortName":"AFTER","name":"AFTER HOURS TIRE SERVICE","phone":"229-589-6702","balance":512.06,"futDue":511.88,"curDue":0.18,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":201058,"shortName":"AFTER","name":"AFTER 5 COMM. TIRE & OFF ROAD","phone":"352-451-3863","balance":-67.1,"futDue":-67.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-12 00:00:00"},{"salesman":"Larry","custNum":100301,"shortName":"ALBANY","name":"ALBANY GENERAL TIRE SERVICE","phone":"229-436-2485","balance":6821.37,"futDue":3416.19,"curDue":3405.18,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"Car Dealer","custNum":200765,"shortName":"ALBANY","name":"ALBANY CHRYSLER DODGE JEEP RAM","phone":"229-233-7769","balance":1119.9,"futDue":212.0,"curDue":907.9,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Larry","custNum":200676,"shortName":"ALL","name":"ALL PURPOSE AUTO CENTER","phone":"229-326-0597","balance":-20.26,"futDue":-20.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Tiffany","custNum":2000018,"shortName":"ALL","name":"ALL PRO DIESEL, LLC","phone":"386-438-9912","balance":-2.04,"futDue":-2.04,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":200971,"shortName":"ALLEN","name":"ALLEN'S TIRE","phone":"229-567-3390","balance":5979.51,"futDue":2582.87,"curDue":3396.64,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":200897,"shortName":"ALMA","name":"ALMA TIRE & AUTO REPAIR","phone":"912-286-9140","balance":-671.28,"futDue":-671.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20 00:00:00"},{"salesman":"Larry","custNum":101080,"shortName":"AMERSON","name":"AMERSON TIRE INC.","phone":"912-393-3674","balance":5814.16,"futDue":5814.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"Car Dealer","custNum":200734,"shortName":"ANDERSON","name":"ANDERSON FORD","phone":"912-384-2600","balance":303.88,"futDue":-1.0,"curDue":304.88,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"House","custNum":200307,"shortName":"ARR","name":"ARREDONDO TIRE SERVICE","phone":"229-529-0882","balance":-0.94,"futDue":-0.94,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"House","custNum":200709,"shortName":"ASHLEY","name":"ASHLEY'S AUTOMOTIVE REPAIR","phone":"229-396-4640","balance":456.64,"futDue":456.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":101439,"shortName":"AT","name":"A.T. TIRE SERVICE","phone":"229-891-5428","balance":1545.94,"futDue":1049.62,"curDue":496.32,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200311,"shortName":"AUTO","name":"AUTOMOTIVE NECESSITIES","phone":"229-434-1237","balance":-752.72,"futDue":-752.72,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"House","custNum":200753,"shortName":"AUTO","name":"AUTO TECH OF MIAMI INC.","phone":"850-997-0200","balance":-2.15,"futDue":-2.15,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"House","custNum":200475,"shortName":"AZALEA","name":"AZALEA CITY AUTO SALES/SERVICE","phone":"229-293-7337","balance":-0.14,"futDue":-0.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Larry","custNum":200503,"shortName":"BB","name":"BB'S AUTOMOTIVE","phone":"229-382-4572","balance":-0.57,"futDue":-0.57,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"House","custNum":200607,"shortName":"BB","name":"B AND B SERVICE CENTER, INC.","phone":"229-236-2886","balance":66.94,"futDue":-41.83,"curDue":108.77,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"Tiffany","custNum":200439,"shortName":"BEAR","name":"BEAR TIRE SERVICE","phone":"229-242-1910","balance":-125.24,"futDue":-125.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Austin","custNum":2000030,"shortName":"BEASON","name":"BEASON EQUIPMENT CO","phone":"229-985-9785","balance":2907.14,"futDue":0.0,"curDue":2907.14,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07 00:00:00"},{"salesman":"Tiffany","custNum":200663,"shortName":"BEN","name":"BEN'S TIRE & AUTO","phone":"229-242-8777","balance":72.2,"futDue":72.2,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":200592,"shortName":"BERNEY","name":"BERNEY'S TIRE SERVICE","phone":"229-435-0412","balance":4086.11,"futDue":963.4,"curDue":3122.71,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"Larry","custNum":201053,"shortName":"BERNEY","name":"BERNEYS TIRE SERVICE","phone":"229-435-0413","balance":8801.39,"futDue":865.96,"curDue":7935.43,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"House","custNum":200562,"shortName":"BEST","name":"BEST CARS OF CORDELE, LLC","phone":"229-273-2378","balance":4.0,"futDue":4.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Larry","custNum":101479,"shortName":"BILL","name":"BILL THOMPSON TIRE SERVICES","phone":"229-435-7753","balance":4021.83,"futDue":3428.82,"curDue":593.01,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":201043,"shortName":"BILL","name":"BILL'S TRAILER SERVICE","phone":"229-396-4391","balance":220.1,"futDue":220.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":101366,"shortName":"BM","name":"B & M AUTOMOTIVE SERVICE","phone":"912-384-6115","balance":-60.68,"futDue":-60.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-06 00:00:00"},{"salesman":"House","custNum":101525,"shortName":"BMS","name":"BMS DISCOUNT TIRES","phone":"229-234-0033","balance":321.88,"futDue":153.36,"curDue":168.52,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"House","custNum":200315,"shortName":"BOB","name":"BOBBY'S CITGO","phone":"229-482-2724","balance":5126.45,"futDue":1446.46,"curDue":3679.99,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"House","custNum":200809,"shortName":"BOULEVARD","name":"BOULEVARD TIRE CENTER","phone":"386-734-6447","balance":8825.97,"futDue":1310.0,"curDue":7515.97,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"House","custNum":200146,"shortName":"BRACE","name":"BRACEWELL AUTOMOTIVE SERVICE","phone":"229-377-1771","balance":4316.52,"futDue":1588.0,"curDue":2728.52,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"Larry","custNum":200317,"shortName":"BRO","name":"BROTHERS TIRES","phone":"229-319-2919","balance":94.58,"futDue":94.58,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":101152,"shortName":"BROOKS","name":"BROOKS BODY SHOP","phone":"229-386-1800","balance":-8.11,"futDue":-8.11,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-27 00:00:00"},{"salesman":"House","custNum":201049,"shortName":"BROUSSARD","name":"BROUSSARD ACCESSORIES LLC","phone":"229-496-8004","balance":-63.31,"futDue":-63.31,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-06 00:00:00"},{"salesman":"Larry","custNum":200319,"shortName":"BUCK","name":"BUCK'S AUTO REPAIR","phone":"229-686-2290","balance":-125.23,"futDue":-125.23,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":200331,"shortName":"BUDGET","name":"BUDGET CAR SALES","phone":"229-388-0020","balance":1105.63,"futDue":0.0,"curDue":1105.63,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-13 00:00:00"},{"salesman":"Larry","custNum":101507,"shortName":"BURNETTE","name":"BURNETTE AUTOMOTIVE SERVICE","phone":"912-632-2713","balance":8562.77,"futDue":474.06,"curDue":8088.71,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-13 00:00:00"},{"salesman":"Larry","custNum":101283,"shortName":"CAMERON","name":"CAMERON'S TOWING AND TIRE","phone":"229-567-2437","balance":1180.07,"futDue":682.32,"curDue":497.75,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200600,"shortName":"CHARLOT","name":"CHARLOT TRUCKING & TIRE SVC.","phone":"229-881-4561","balance":-6.53,"futDue":-6.53,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":200681,"shortName":"CITY","name":"CITY OF SYLVESTER","phone":"229-776-8504","balance":291.22,"futDue":-456.78,"curDue":748.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Tiffany","custNum":200886,"shortName":"CL","name":"C&L PERFORMANCE INC","phone":"912-283-0071","balance":113.29,"futDue":0.0,"curDue":113.29,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-10 00:00:00"},{"salesman":"Larry","custNum":100417,"shortName":"CLARK","name":"CLARK BASS SERVICE","phone":"229-874-4685","balance":13324.75,"futDue":1331.78,"curDue":11992.97,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":200807,"shortName":"CORDELE","name":"CORDELE TIRE & WHEEL, LLC","phone":"229-417-5099","balance":1109.59,"futDue":1109.59,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":101539,"shortName":"COURSON","name":"COURSON'S TIRE OF DOUGLAS","phone":"912-383-6188","balance":1631.57,"futDue":1631.57,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":201018,"shortName":"CRAWLEY","name":"CRAWLEY'S AUTOMOTIVE & TIRE","phone":"386-658-2007","balance":-52.86,"futDue":-52.86,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":200674,"shortName":"DASH","name":"DASHER LLC","phone":"229-469-6353","balance":152.13,"futDue":150.13,"curDue":2.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Larry","custNum":101295,"shortName":"DAVID","name":"DAVID'S AUTO SALES / DOUGLAS","phone":"912-384-8570","balance":2102.78,"futDue":2102.78,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07 00:00:00"},{"salesman":"Larry","custNum":201035,"shortName":"DAVID","name":"DAVID'S AUTO SALES (MOULTRIE)","phone":"229-217-4959","balance":533.12,"futDue":533.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":201008,"shortName":"DBJ","name":"DBJ MOBILE TIRE SERVICE, INC.","phone":"386-219-6036","balance":487.73,"futDue":0.0,"curDue":487.73,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200162,"shortName":"DELTA","name":"DELTA TIRE CO","phone":"229-246-2750","balance":13746.7,"futDue":1930.85,"curDue":11815.85,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"House","custNum":200416,"shortName":"DENT","name":"DENT'S SERVICE STATION","phone":"229-896-4160","balance":-0.93,"futDue":-0.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200506,"shortName":"DEV","name":"DEVANE TIRE & SERVICE LLC","phone":"229-310-9586","balance":1000.44,"futDue":0.0,"curDue":1000.44,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-24 00:00:00"},{"salesman":"House","custNum":200979,"shortName":"DISCOUNT","name":"DISCOUNT TIRE & AUTO SHOP","phone":"850-544-7234","balance":519.35,"futDue":519.35,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Larry","custNum":201014,"shortName":"DR","name":"D&R AUTO SALES & SALVAGE PARTS","phone":"912-534-6543","balance":-6.68,"futDue":-6.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200327,"shortName":"DRAPER","name":"DRAPER TIRES & AUTOMOTIVE","phone":"229-247-4531","balance":-111.14,"futDue":-111.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":201034,"shortName":"ECONOMIC","name":"ECONOMIC NICHOLAS TIRE","phone":"229-454-5123","balance":-12.32,"futDue":-12.32,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-06 00:00:00"},{"salesman":"Larry","custNum":101436,"shortName":"ED","name":"ED'S TIRE","phone":"229-776-6952","balance":8497.33,"futDue":3754.64,"curDue":4742.69,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Tiffany","custNum":200967,"shortName":"EDD","name":"EDDIES AUTOMOTIVE AND TIRE","phone":"229-232-4958","balance":-179.28,"futDue":-179.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":200166,"shortName":"EDISON","name":"EDISON TIRE","phone":"229-835-2077","balance":5873.36,"futDue":0.0,"curDue":5873.36,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":2000015,"shortName":"EDWIN","name":"EDWIN'S TIRES LLC","phone":"229-429-6141","balance":-12.6,"futDue":-12.6,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200642,"shortName":"EG","name":"E.G. AUTO SALES","phone":"229-769-5011","balance":-4.0,"futDue":-170.3,"curDue":166.3,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"Tiffany","custNum":200906,"shortName":"EH","name":"E&H TIRE","phone":"386-397-3333","balance":-192.9,"futDue":-192.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":200891,"shortName":"EJH","name":"EJH WRECKER & TIRE SERVICE","phone":"229-566-3334","balance":451.73,"futDue":451.73,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":101323,"shortName":"ERIC","name":"ERIC'S TIRE SERVICE","phone":"229-472-1543","balance":10246.59,"futDue":1923.44,"curDue":8323.15,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Larry","custNum":200972,"shortName":"ERIC","name":"ERIC'S TIRE OF SYLVESTER","phone":"229-821-2000","balance":1094.37,"futDue":63.0,"curDue":1031.37,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Larry","custNum":200973,"shortName":"ERIC","name":"ERIC'S TIRE (REBEL ROAD)","phone":"229-396-4946","balance":5799.86,"futDue":2741.37,"curDue":3058.49,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"House","custNum":200410,"shortName":"EZ","name":"EZDEALIN WHEELS AND TIRES","phone":"877-247-2230","balance":11381.93,"futDue":11381.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200647,"shortName":"FAUSETTS","name":"FAUSETTS TIRE CO.","phone":"229-896-7481","balance":-22.2,"futDue":-22.2,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Car Dealer","custNum":101300,"shortName":"FITZGERALD","name":"FITZGERALD FORD AND LINCOLN","phone":"229-423-8787","balance":794.68,"futDue":0.0,"curDue":794.68,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15 00:00:00"},{"salesman":"Larry","custNum":200560,"shortName":"FIVE","name":"FIVE STAR TIRE SERVICE,LLC","phone":"229-396-5412","balance":825.86,"futDue":180.62,"curDue":645.24,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Car Dealer","custNum":200533,"shortName":"FLOW","name":"FLOWERS IMPORTS LLC (HONDA)","phone":"229-225-1144","balance":250.0,"futDue":250.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-11-11 00:00:00"},{"salesman":"Car Dealer","custNum":2000025,"shortName":"FORD","name":"FORD CORDELE","phone":"229-276-0607","balance":2807.3,"futDue":1026.5,"curDue":1780.8,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200719,"shortName":"FOSTER","name":"FOSTER EASY PAY TIRE CO., INC.","phone":"229-995-2167","balance":17156.37,"futDue":2033.46,"curDue":15122.91,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":101297,"shortName":"FOUR","name":"FOUR C'S LUBE","phone":"912-422-6866","balance":167.16,"futDue":167.16,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Tiffany","custNum":200935,"shortName":"FROMETA","name":"FROMETA USED CAR & TIRE CENTER","phone":"229-333-8225","balance":-99.36,"futDue":-99.36,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-04 00:00:00"},{"salesman":"Austin","custNum":200266,"shortName":"FUSS","name":"FUSSELL TIRE & SERVICE","phone":"229-259-0034","balance":66486.38,"futDue":27977.48,"curDue":38508.9,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"Tiffany","custNum":200543,"shortName":"GATOR","name":"GATOR TIRE","phone":"912-670-9397","balance":-9.77,"futDue":-9.77,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200900,"shortName":"GILLETTES","name":"GILLETTES AUTO","phone":"386-362-5171","balance":-17.64,"futDue":-17.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-30 00:00:00"},{"salesman":"Tiffany","custNum":101544,"shortName":"GODWIN","name":"GODWIN TIRE & AUTO","phone":"229-294-9553","balance":118.23,"futDue":0.0,"curDue":118.23,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"House","custNum":200962,"shortName":"GREENE","name":"GREENE'S TIRE SERVICE LLC","phone":"229-699-0748","balance":-19.88,"futDue":-19.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"Car Dealer","custNum":200831,"shortName":"GRIFFFIN","name":"GRIFFIN CHRYSLER DODGE JEEP","phone":"229-382-0440","balance":927.8,"futDue":579.68,"curDue":348.12,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Car Dealer","custNum":200832,"shortName":"GRIFFIN","name":"GRIFFIN FORD","phone":"229-382-1300","balance":527.8,"futDue":0.0,"curDue":527.8,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Car Dealer","custNum":200933,"shortName":"GRIFFIN","name":"GRIFFIN CDJR VALDOSTA","phone":"229-471-4645","balance":445.38,"futDue":216.62,"curDue":228.76,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Car Dealer","custNum":1999999,"shortName":"GRIFFIN","name":"GRIFFIN CHEVROLET OF SYLVESTER","phone":"229-776-3473","balance":98.52,"futDue":0.0,"curDue":98.52,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Tiffany","custNum":200270,"shortName":"HAHIRA","name":"HAHIRA AUTOMOTIVE SERVICE","phone":"229-794-2429","balance":6105.82,"futDue":2249.89,"curDue":3855.93,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"House","custNum":200451,"shortName":"HARROD","name":"HARROD BROTHERS","phone":"229-686-3959","balance":3098.4,"futDue":285.02,"curDue":2813.38,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"Tiffany","custNum":101565,"shortName":"HARRY","name":"HARRY B ANDERSON","phone":"229-242-5945","balance":1379.04,"futDue":0.0,"curDue":1379.04,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":201069,"shortName":"HEAVY","name":"HEAVY DUTY TIRE","phone":"904-833-7750","balance":-6.6,"futDue":-6.6,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-01-16 00:00:00"},{"salesman":"Larry","custNum":200683,"shortName":"HICKOX","name":"HICKOX AUTO DEALERS","phone":"912-281-3922","balance":-0.5,"futDue":-0.5,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"House","custNum":200407,"shortName":"HOLLO","name":"HOLLOWAY TRUCK & TRAILER REPAI","phone":"229-387-8795","balance":619.9,"futDue":619.9,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Car Dealer","custNum":200554,"shortName":"HONDA","name":"HONDA OF SOUTH GEORGIA","phone":"229-396-4050","balance":19.72,"futDue":0.0,"curDue":19.72,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-20 00:00:00"},{"salesman":"House","custNum":200277,"shortName":"IMP","name":"IMPORT SERVICE & SALES","phone":"229-226-9844","balance":1304.99,"futDue":378.18,"curDue":926.81,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-08 00:00:00"},{"salesman":"Tiffany","custNum":200673,"shortName":"JB","name":"JB'S TIRE & REPAIR SVC.","phone":"850-584-2400","balance":206.67,"futDue":206.67,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Car Dealer","custNum":200580,"shortName":"JEFF","name":"JEFF FENDER BUICK, GMC, CAD.","phone":"229-386-1985","balance":4979.64,"futDue":0.0,"curDue":4979.64,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Larry","custNum":101161,"shortName":"JMC","name":"JMC TIRE CO., INC.","phone":"912-384-4940","balance":-2179.24,"futDue":-2179.24,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200648,"shortName":"JOE","name":"JOE'S AUTO REPAIR, LLC","phone":"912-384-2010","balance":-78.86,"futDue":-78.86,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15 00:00:00"},{"salesman":"House","custNum":200805,"shortName":"JOEY","name":"JOEY HALL AUTO SALES LLC","phone":"229-382-6900","balance":-0.92,"futDue":-0.92,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-03 00:00:00"},{"salesman":"Larry","custNum":100107,"shortName":"JOHNSON","name":"JOHNSON AUTO & TIRE","phone":"912-359-2452","balance":1029.58,"futDue":60.34,"curDue":969.24,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"Larry","custNum":2000017,"shortName":"JORDAN","name":"JORDAN AUTOMOTIVE & TIRES","phone":"229-439-6881","balance":-6.0,"futDue":-6.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"House","custNum":200754,"shortName":"JORGE","name":"JORGE USED TIRE SHOP","phone":"912-309-5729","balance":-1.19,"futDue":-1.19,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":2000007,"shortName":"KEATON","name":"KEATON & SON TIRE LLC","phone":"850-284-9344","balance":592.72,"futDue":209.64,"curDue":383.08,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Tiffany","custNum":200691,"shortName":"KENDA","name":"KENDA TRUCK CENTER","phone":"850-929-3700","balance":30261.87,"futDue":0.0,"curDue":30261.87,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14 00:00:00"},{"salesman":"Tiffany","custNum":200836,"shortName":"KING","name":"KING MUFFLER","phone":"912-285-4939","balance":1273.38,"futDue":1273.38,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Car Dealer","custNum":2000009,"shortName":"KING","name":"KING FORD OF NASHVILLE","phone":"229-686-2058","balance":-1474.74,"futDue":-1474.74,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Tiffany","custNum":200827,"shortName":"LAKELAND","name":"LAKELAND TIRE DBA COOK & SONS","phone":"229-482-1000","balance":686.3,"futDue":686.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200912,"shortName":"LANE","name":"LANE'S TRK & TRL REPAIR & AUTO","phone":"229-322-6338","balance":-0.1,"futDue":-0.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-06 00:00:00"},{"salesman":"Car Dealer","custNum":200348,"shortName":"LANG","name":"LANGDALE KIA (SERVICE)","phone":"229-242-3835","balance":75.02,"futDue":0.0,"curDue":75.02,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-14 00:00:00"},{"salesman":"Larry","custNum":200501,"shortName":"LAR","name":"LARRY'S BODY SHOP","phone":"229-386-4523","balance":3.78,"futDue":0.0,"curDue":3.78,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"House","custNum":200913,"shortName":"LASHLEY","name":"LASHLEY'S HOMETOWN TIRE LLC","phone":"386-209-4919","balance":-2.93,"futDue":-2.93,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":200191,"shortName":"LEE","name":"LEE COUNTY AUTO SERVICE","phone":"229-759-2001","balance":-370.71,"futDue":-370.71,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200866,"shortName":"LEMUS","name":"LEMUS TIRE SHOP","phone":"229-417-4737","balance":-17.77,"futDue":-17.77,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":200901,"shortName":"LENCHO","name":"LENCHO'S & SON TIRE SHOP","phone":"229-699-1259","balance":238.12,"futDue":-10.68,"curDue":248.8,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Tiffany","custNum":200406,"shortName":"LENOX","name":"LENOX TIRE & SERVICE CENTER","phone":"229-546-4119","balance":1272.23,"futDue":1272.23,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"House","custNum":201024,"shortName":"LIBERTY","name":"LIBERTY CAR WASH & TIRE","phone":"912-387-0408","balance":1.0,"futDue":0.0,"curDue":1.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Tiffany","custNum":200687,"shortName":"LIVE","name":"LIVE OAK TIRE CENTER, LLC","phone":"386-362-1972","balance":1366.75,"futDue":1028.51,"curDue":338.24,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":200459,"shortName":"LOW","name":"LOW COUNTRY TIRE LLC","phone":"912-590-6331","balance":-0.02,"futDue":-0.02,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"House","custNum":200352,"shortName":"MAL","name":"MALUDA AUTO SALES","phone":"229-244-9163","balance":-0.01,"futDue":-0.01,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"Larry","custNum":200760,"shortName":"MALLARD","name":"MALLARD'S SERVICE CENTER","phone":"478-960-5246","balance":321.1,"futDue":321.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"House","custNum":200761,"shortName":"MARIO","name":"MARIO NEW AND USED TIRE SHOP","phone":"229-591-5893","balance":-73.14,"futDue":-73.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":101519,"shortName":"MARK","name":"MARK TAYLOR DBA/MTAA ENT.","phone":"229-425-4601","balance":677.57,"futDue":677.57,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"House","custNum":200631,"shortName":"MARQ","name":"MARQUEZ TIRE SHOP","phone":"229-921-3897","balance":368.36,"futDue":-342.59,"curDue":710.95,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Larry","custNum":200965,"shortName":"MARQ","name":"MARQUEZ TIRE SHOP LLC","phone":"229-886-8657","balance":-505.8,"futDue":-505.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":101530,"shortName":"MARTIN","name":"MARTIN TIRE SERVICE","phone":"229-425-4377","balance":834.68,"futDue":834.68,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"House","custNum":200964,"shortName":"MARTINEZ","name":"MARTINEZ AUTO SERVICE","phone":"972-801-7027","balance":-0.52,"futDue":-0.52,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-20 00:00:00"},{"salesman":"House","custNum":100741,"shortName":"MASSEY","name":"MASSEY'S MUFFLER","phone":"229-273-4339","balance":119.0,"futDue":119.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":101524,"shortName":"MASTER","name":"MASTER BODY WORKS","phone":"229-439-8833","balance":213.85,"futDue":125.52,"curDue":88.33,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Larry","custNum":200468,"shortName":"MCKEE","name":"MCKEE'S AUTO CENTER, INC","phone":"229-382-7642","balance":79.06,"futDue":79.06,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200193,"shortName":"MCLEAN","name":"MCLEAN TIRES INC","phone":"229-782-7428","balance":8494.95,"futDue":2449.73,"curDue":6045.22,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Tiffany","custNum":201005,"shortName":"MIDWAY","name":"MIDWAY ENTERPRISE FL, LLC","phone":"850-875-2444","balance":-159.42,"futDue":-159.42,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-07 00:00:00"},{"salesman":"House","custNum":101208,"shortName":"MIKE","name":"MIKE FRASER AUTO REPAIR","phone":"229-273-0652","balance":-60.64,"futDue":-60.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"House","custNum":200622,"shortName":"MOORE","name":"MOORE'S ACCESSORIES & OFFROAD","phone":"229-223-6389","balance":-143.35,"futDue":-143.35,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-09 00:00:00"},{"salesman":"House","custNum":200355,"shortName":"MOTION","name":"MOTION WHEELS,HUBCAPS,&TIRES","phone":"229-253-0080","balance":-33.1,"futDue":-33.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-05 00:00:00"},{"salesman":"Larry","custNum":200510,"shortName":"MR","name":"M & R TRUCK ACCESSORIES","phone":"912-384-2362","balance":-0.3,"futDue":-0.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-28 00:00:00"},{"salesman":"Tiffany","custNum":2000044,"shortName":"MURRAY","name":"MURRAY'S TIRE & ROAD SERVICE","phone":"386-752-5688","balance":1088.38,"futDue":1088.38,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Car Dealer","custNum":200688,"shortName":"NASH","name":"NASHVILLE FORD","phone":"229-686-2058","balance":-4.88,"futDue":-4.88,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-07-01 00:00:00"},{"salesman":"House","custNum":101549,"shortName":"NASHVILLE","name":"NASHVILLE TIRE","phone":"229-686-1900","balance":369.26,"futDue":369.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200356,"shortName":"NEEL","name":"NEELY'S SERVICE CENTER","phone":"229-263-4454","balance":308.5,"futDue":0.0,"curDue":308.5,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-08 00:00:00"},{"salesman":"Tiffany","custNum":101146,"shortName":"NERO","name":"NE-RO TIRE & BRAKE SERVICE,INC","phone":"229-244-8353","balance":11857.0,"futDue":1659.74,"curDue":10197.26,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"House","custNum":201003,"shortName":"NICHOLAS","name":"NICHOLAS TIRES INC.","phone":"229-850-0289","balance":431.27,"futDue":-1.53,"curDue":432.8,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":200755,"shortName":"NT","name":"N-T TIRE SERVICE","phone":"229-891-2234","balance":434.14,"futDue":434.14,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":201016,"shortName":"OC","name":"O&C AUTO REPAIR","phone":"229-821-1221","balance":-10.74,"futDue":-10.74,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"Car Dealer","custNum":200908,"shortName":"OSTEEN","name":"O'STEEN CHRYSLER DODGE JEEP","phone":"229-686-2068","balance":-44.89,"futDue":-225.16,"curDue":180.27,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-17 00:00:00"},{"salesman":"House","custNum":101108,"shortName":"PARKER","name":"PARKER TIRE DIRECT","phone":"229-457-3818","balance":5679.36,"futDue":1480.22,"curDue":4199.14,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-21 00:00:00"},{"salesman":"House","custNum":200396,"shortName":"PAT","name":"PATE TIRE & SERVICE LLC","phone":"229-890-7389","balance":1344.99,"futDue":1344.99,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Larry","custNum":200388,"shortName":"PEAR","name":"PEARSON TIRE & LUBE","phone":"912-422-6820","balance":-324.73,"futDue":-324.73,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"House","custNum":201023,"shortName":"PEASE","name":"PEASE ON THE GO 24/7","phone":"229-539-6995","balance":-56.02,"futDue":-56.02,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":101497,"shortName":"PERF","name":"PERFORMANCE MOTORSPORT","phone":"229-438-5248","balance":-3.06,"futDue":-3.06,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Tiffany","custNum":200885,"shortName":"PIERCE","name":"PIERCE INDUSTRIAL TIRE LLC","phone":"912-807-3685","balance":12479.85,"futDue":10112.69,"curDue":2367.16,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Larry","custNum":201017,"shortName":"PINEDA","name":"PINEDA'S AUTOMOTIVE","phone":"229-382-1583","balance":121.55,"futDue":121.55,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":200890,"shortName":"PMT","name":"PMT TRK. TRAILER & TIRE REPAIR","phone":"229-457-5167","balance":1231.01,"futDue":1231.01,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200358,"shortName":"POND","name":"PONDER'S AUTOMOTIVE INC","phone":"229-228-5779","balance":556.0,"futDue":0.0,"curDue":556.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20 00:00:00"},{"salesman":"House","custNum":200658,"shortName":"POWE","name":"POWE AUTOMOTIVE","phone":"229-397-0459","balance":-223.82,"futDue":-223.82,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-01 00:00:00"},{"salesman":"Larry","custNum":200762,"shortName":"POWER","name":"POWER MAN TIRE SHOP","phone":"912-381-4065","balance":4025.62,"futDue":756.21,"curDue":3269.41,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Larry","custNum":200920,"shortName":"PRECISION","name":"PRECISION MAINTENANCE","phone":"912-253-5237","balance":-395.26,"futDue":-395.26,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":201052,"shortName":"PRECISION","name":"PRECISION DIESEL REPAIR LLC","phone":"229-472-7268","balance":1925.42,"futDue":1925.42,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"House","custNum":200213,"shortName":"PREMIER","name":"PREMIER AUTOWORKS","phone":"229-435-2886","balance":-406.0,"futDue":-406.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20 00:00:00"},{"salesman":"Car Dealer","custNum":101374,"shortName":"PRINCE","name":"PRINCE CHEVY-OLDS, INC","phone":"229-386-4050","balance":4919.48,"futDue":0.0,"curDue":4919.48,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"Car Dealer","custNum":200455,"shortName":"PRINCE","name":"PRINCE HONDA","phone":"229-386-1400","balance":470.98,"futDue":309.0,"curDue":161.98,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15 00:00:00"},{"salesman":"Car Dealer","custNum":200456,"shortName":"PRINCE","name":"PRINCE TOYOTA","phone":"229-386-4052","balance":5236.52,"futDue":346.04,"curDue":4890.48,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-26 00:00:00"},{"salesman":"House","custNum":200462,"shortName":"QUALITY","name":"QUALITY FEEDSTUFFS, INC","phone":"229-686-2770","balance":462.94,"futDue":0.0,"curDue":462.94,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-23 00:00:00"},{"salesman":"House","custNum":200585,"shortName":"QUI","name":"QUINCY TIRE AND RECAPPING","phone":"850-627-6050","balance":9433.83,"futDue":1664.67,"curDue":7769.16,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Larry","custNum":101372,"shortName":"RAIN","name":"RAINEY USED CARS (BRONWOOD)","phone":"229-695-9153","balance":4277.34,"futDue":728.08,"curDue":3549.26,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-18 00:00:00"},{"salesman":"Tiffany","custNum":200362,"shortName":"RAY","name":"RAY NORTON TIRE & AUTO","phone":"229-247-1555","balance":4395.82,"futDue":506.52,"curDue":3889.3,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"House","custNum":201031,"shortName":"RENO","name":"RENO'S QUALITY COLLISION","phone":"229-563-0752","balance":393.12,"futDue":393.12,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-24 00:00:00"},{"salesman":"Larry","custNum":200679,"shortName":"RICHARD","name":"RICHARD'S AUTO CARE & TIRE SVC","phone":"229-407-8600","balance":-48.83,"futDue":-48.83,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":200288,"shortName":"RIGHT","name":"RIGHT CHOICE AUTO","phone":"nan","balance":-0.08,"futDue":-0.08,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-04-10 00:00:00"},{"salesman":"House","custNum":200241,"shortName":"RIM","name":"THE RIM SHOP INC","phone":"334-793-9292","balance":-13.64,"futDue":-13.64,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-02-26 00:00:00"},{"salesman":"House","custNum":200715,"shortName":"RM","name":"R&M AUTO TRUCKING, INC","phone":"229-206-2102","balance":120.8,"futDue":120.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Car Dealer","custNum":101256,"shortName":"ROBERT","name":"ROBERT FENDER CHEVROLET","phone":"912-292-9005","balance":61.13,"futDue":0.0,"curDue":61.13,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-08-27 00:00:00"},{"salesman":"Car Dealer","custNum":200365,"shortName":"ROBERT","name":"ROBERT HUTSON LINCOLN","phone":"229-985-6603","balance":969.6,"futDue":468.92,"curDue":500.68,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-27 00:00:00"},{"salesman":"House","custNum":101025,"shortName":"ROCHELLE","name":"ROCHELLE TIRE","phone":"229-365-7943","balance":2424.82,"futDue":1883.22,"curDue":541.6,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Tiffany","custNum":200883,"shortName":"ROLLING","name":"ROLLING BEAR TIRES LLC","phone":"912-387-6642","balance":1485.26,"futDue":1377.1,"curDue":108.16,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200880,"shortName":"RR","name":"R&R TIRE CO.","phone":"229-805-4245","balance":-7.27,"futDue":-7.27,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":2000022,"shortName":"RR","name":"R&R AUTO SERVICE & REPAIR","phone":"912-347-5301","balance":-231.06,"futDue":-231.06,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":201048,"shortName":"RRO","name":"RRO 24 HR ROADSIDE ASSISTANCE","phone":"386-965-0117","balance":-113.2,"futDue":-113.2,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"House","custNum":200975,"shortName":"SANTOS","name":"SANTOS TIRE SHOP","phone":"229-921-6310","balance":157.73,"futDue":157.73,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"Larry","custNum":101463,"shortName":"SHELL","name":"SHELL RAPID LUBE (FITZGERALD)","phone":"229-424-9348","balance":25262.07,"futDue":4905.22,"curDue":20356.85,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":200220,"shortName":"SING","name":"SINGLETARY & SON TIRE CO","phone":"229-776-5535","balance":5396.99,"futDue":2277.2,"curDue":3119.79,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Tiffany","custNum":200291,"shortName":"SING","name":"SINGLETARY TIRE PROS","phone":"229-226-2842","balance":317.85,"futDue":0.0,"curDue":317.85,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-20 00:00:00"},{"salesman":"House","custNum":200980,"shortName":"SLYDER","name":"SLYDER'S GARAGE","phone":"229-560-0388","balance":158.0,"futDue":158.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":201070,"shortName":"SNIDER","name":"SNIDER INDUSTRIAL","phone":"904-383-1143","balance":345.8,"futDue":345.8,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-08 00:00:00"},{"salesman":"Larry","custNum":101181,"shortName":"SOUTH","name":"SOUTH GA LUBE CENTER","phone":"229-468-4435","balance":387.77,"futDue":224.78,"curDue":162.99,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Larry","custNum":200621,"shortName":"SOUTH","name":"SOUTH GA LUBE (FITZGERALD)","phone":"229-345-1704","balance":831.89,"futDue":0.0,"curDue":831.89,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"House","custNum":200803,"shortName":"SOUTH","name":"SOUTH MAIN GARAGE","phone":"229-566-3880","balance":-0.3,"futDue":-0.3,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200918,"shortName":"SOUTH","name":"SOUTH GEORGIA TRUCKING SVC LLC","phone":"229-472-1056","balance":1416.33,"futDue":1416.33,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-13 00:00:00"},{"salesman":"Larry","custNum":200224,"shortName":"SOUTHERN","name":"SOUTHERN EXPRESS LUBE","phone":"229-777-0932","balance":702.94,"futDue":0.0,"curDue":702.94,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-14 00:00:00"},{"salesman":"House","custNum":200230,"shortName":"SOUTHERN","name":"SOUTHERN TIRE & BATTERY","phone":"229-246-4925","balance":377.06,"futDue":0.0,"curDue":377.06,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-05-13 00:00:00"},{"salesman":"House","custNum":200742,"shortName":"SOUTHERN","name":"SOUTHERN SALES & RENTALS, LLC","phone":"706-546-9760","balance":613.9,"futDue":0.0,"curDue":613.9,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-11 00:00:00"},{"salesman":"House","custNum":200896,"shortName":"SOUTHERN","name":"SOUTHERN AUTO SPECIALIST","phone":"229-217-0558","balance":198.0,"futDue":198.0,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-21 00:00:00"},{"salesman":"House","custNum":200949,"shortName":"SOUTHERN","name":"SOUTHERN TIRE MART","phone":"229-920-3030","balance":4070.39,"futDue":2841.36,"curDue":1229.03,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-02 00:00:00"},{"salesman":"Larry","custNum":200959,"shortName":"SOUTHERN","name":"SOUTHERN GEORGIA TIRE LLC","phone":"912-292-0001","balance":-598.4,"futDue":-598.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":201063,"shortName":"SOUTHERN","name":"SOUTHERN TIRE EXPORTERS","phone":"404-819-5113","balance":-6.1,"futDue":-6.1,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-03-13 00:00:00"},{"salesman":"Tiffany","custNum":2000029,"shortName":"SOUTHERN","name":"SOUTHERN TIRE MART @ PILOT","phone":"229-244-3179","balance":2156.3,"futDue":0.0,"curDue":2156.3,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-20 00:00:00"},{"salesman":"House","custNum":200546,"shortName":"SOW","name":"SOWEGA TIRE OF ALBANY","phone":"229-888-1881","balance":345.99,"futDue":244.99,"curDue":101.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"Car Dealer","custNum":200232,"shortName":"SUNBELT","name":"SUNBELT FORD INC","phone":"229-776-7691","balance":-1014.5,"futDue":-1014.5,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-10-13 00:00:00"},{"salesman":"Tiffany","custNum":200876,"shortName":"SUWANNEE","name":"SUWANNEE VALLEY TIRE","phone":"386-792-2420","balance":322.44,"futDue":0.0,"curDue":322.44,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-12 00:00:00"},{"salesman":"House","custNum":101305,"shortName":"SYCAMORE","name":"SYCAMORE SALES & SALVAGE LLC","phone":"229-567-2005","balance":-14.28,"futDue":-14.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"Tiffany","custNum":200712,"shortName":"TANNER","name":"TANNER AUTO REPAIR PLUS, LLC","phone":"912-807-8277","balance":-93.82,"futDue":-93.82,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-16 00:00:00"},{"salesman":"Larry","custNum":200490,"shortName":"TCA","name":"T.C.A. IRRIGATION","phone":"229-387-7097","balance":401.09,"futDue":0.0,"curDue":401.09,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-15 00:00:00"},{"salesman":"Larry","custNum":2000020,"shortName":"TD","name":"T&D TIRE","phone":"229-456-3436","balance":209.56,"futDue":209.56,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-15 00:00:00"},{"salesman":"Larry","custNum":200956,"shortName":"TENNESON","name":"TENNESON COLLISION CENTER","phone":"229-391-9318","balance":403.19,"futDue":0.0,"curDue":403.19,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"Larry","custNum":101415,"shortName":"THE","name":"THE TIRE STORE","phone":"229-686-2073","balance":25818.16,"futDue":3310.8,"curDue":22507.36,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"House","custNum":200756,"shortName":"THE","name":"THE SHOP OF ALBANY, LLC","phone":"229-573-7066","balance":-2.37,"futDue":-2.37,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Larry","custNum":200242,"shortName":"THOMAS","name":"THOMAS TIRE COMPANY, LLC","phone":"229-985-1839","balance":808.36,"futDue":400.51,"curDue":407.85,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-12-30 00:00:00"},{"salesman":"Larry","custNum":101326,"shortName":"TIFTON","name":"TIFTON GENERAL TIRE","phone":"229-382-6013","balance":36437.8,"futDue":4983.04,"curDue":31454.76,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Tiffany","custNum":200294,"shortName":"TIRE","name":"TIRE KING OF VALDOSTA","phone":"229-247-1345","balance":19314.27,"futDue":10990.64,"curDue":8323.63,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-04-24 00:00:00"},{"salesman":"Tiffany","custNum":200474,"shortName":"TIRE","name":"TIRE & WHEEL INC","phone":"912-449-6164","balance":1777.44,"futDue":1777.44,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"Larry","custNum":500373,"shortName":"TOMAHAWK","name":"TOMAHAWK TIRE (ALBANY)","phone":"229-439-6594","balance":102.85,"futDue":102.85,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200537,"shortName":"TONY","name":"TONY'S TIRE & ROAD SERVICE,INC","phone":"229-890-9989","balance":298.4,"futDue":298.4,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Tiffany","custNum":200659,"shortName":"TOWN","name":"TOWN & COUNTRY TIRE","phone":"386-362-4535","balance":601.41,"futDue":0.0,"curDue":601.41,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-01-14 00:00:00"},{"salesman":"Tiffany","custNum":200246,"shortName":"TRI","name":"TRI COUNTY TIRE COMPANY","phone":"229-524-2654","balance":3420.51,"futDue":86.05,"curDue":3334.46,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"Larry","custNum":200239,"shortName":"TS","name":"T & S TIRE","phone":"229-888-0696","balance":-82.28,"futDue":-82.28,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"House","custNum":101201,"shortName":"TUCKER","name":"TUCKERS SERVICE STATION","phone":"229-532-6097","balance":321.41,"futDue":321.41,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-04 00:00:00"},{"salesman":"Larry","custNum":200601,"shortName":"TUFF","name":"TUFF ENTERPRISES LLC","phone":"229-883-8700","balance":2983.36,"futDue":0.0,"curDue":2983.36,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"House","custNum":200624,"shortName":"WAL","name":"WALLACE MOTORS","phone":"850-973-1230","balance":565.97,"futDue":0.0,"curDue":565.97,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-29 00:00:00"},{"salesman":"Car Dealer","custNum":200461,"shortName":"WALKER","name":"WALKER JONES CHEVY-BUICK","phone":"912-490-1314","balance":1145.28,"futDue":0.0,"curDue":1145.28,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-28 00:00:00"},{"salesman":"House","custNum":200806,"shortName":"WALKERS","name":"WALKERS AUTO & OUTDOOR, INC","phone":"912-487-0084","balance":158.84,"futDue":0.0,"curDue":158.84,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Tiffany","custNum":101066,"shortName":"WARRIOR","name":"WARRIOR CREEK TIRE, LLC","phone":"229-798-0923","balance":2883.56,"futDue":504.6,"curDue":2378.96,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-01 00:00:00"},{"salesman":"House","custNum":200132,"shortName":"WILKS","name":"WILKS A-ONE TIRE SALES","phone":"334-792-2225","balance":8583.08,"futDue":7880.23,"curDue":702.85,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-19 00:00:00"},{"salesman":"Tiffany","custNum":200383,"shortName":"WILL","name":"WILLIAMS ALIGNMENT & TIRE","phone":"229-263-4797","balance":-355.87,"futDue":-355.87,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-05 00:00:00"},{"salesman":"House","custNum":200599,"shortName":"WL","name":"W&L TIRE & WHEEL CO. INC.","phone":"850-627-8830","balance":2347.15,"futDue":0.0,"curDue":2347.15,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-05-12 00:00:00"},{"salesman":"Tiffany","custNum":200595,"shortName":"WR","name":"W.R. WILLIAMS","phone":"386-294-1888","balance":19338.0,"futDue":6285.5,"curDue":13052.5,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2026-06-03 00:00:00"},{"salesman":"House","custNum":200300,"shortName":"Z","name":"Z TIRE EXPRESS","phone":"229-244-2084","balance":-116.56,"futDue":-116.56,"curDue":0.0,"due1_30":0.0,"due31_60":0.0,"due61_90":0.0,"dueOver90":0.0,"lastPaid":"2025-09-19 00:00:00"}]

// ── File Parsers ──────────────────────────────────────────────────────────────
function readSheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

const SEED_BRANCH_DATA = {"weeklySales":[{"week":1,"Byron":71317.0,"Tifton":45986.0,"Statesboro":46756.0,"Athens":73570.0,"Byron25":163452.0,"Tifton25":124056.0,"Statesboro25":115721.0,"Athens25":115171.0},{"week":2,"Byron":570752.0,"Tifton":383464.0,"Statesboro":481873.0,"Athens":383818.0,"Byron25":377044.0,"Tifton25":303299.0,"Statesboro25":462387.0,"Athens25":273827.0},{"week":3,"Byron":426682.0,"Tifton":376015.0,"Statesboro":404746.0,"Athens":342498.0,"Byron25":425593.0,"Tifton25":246446.0,"Statesboro25":427419.0,"Athens25":349566.0},{"week":4,"Byron":522986.0,"Tifton":351882.0,"Statesboro":380511.0,"Athens":361139.0,"Byron25":247061.0,"Tifton25":162354.0,"Statesboro25":250715.0,"Athens25":268575.0},{"week":5,"Byron":475548.0,"Tifton":308672.0,"Statesboro":415940.0,"Athens":264278.0,"Byron25":320762.0,"Tifton25":262329.0,"Statesboro25":405301.0,"Athens25":254913.0},{"week":6,"Byron":476077.0,"Tifton":349456.0,"Statesboro":370412.0,"Athens":420047.0,"Byron25":474253.0,"Tifton25":350519.0,"Statesboro25":493114.0,"Athens25":354159.0},{"week":7,"Byron":465206.0,"Tifton":318158.0,"Statesboro":436169.0,"Athens":392013.0,"Byron25":420630.0,"Tifton25":308025.0,"Statesboro25":461629.0,"Athens25":386143.0},{"week":8,"Byron":509095.0,"Tifton":433433.0,"Statesboro":471785.0,"Athens":423549.0,"Byron25":421598.0,"Tifton25":305074.0,"Statesboro25":495207.0,"Athens25":352937.0},{"week":9,"Byron":545565.0,"Tifton":454495.0,"Statesboro":440677.0,"Athens":447750.0,"Byron25":388503.0,"Tifton25":356486.0,"Statesboro25":487012.0,"Athens25":332534.0},{"week":10,"Byron":578802.0,"Tifton":470119.0,"Statesboro":497236.0,"Athens":412614.0,"Byron25":634848.0,"Tifton25":414365.0,"Statesboro25":571197.0,"Athens25":375647.0},{"week":11,"Byron":640012.0,"Tifton":543489.0,"Statesboro":504057.0,"Athens":501260.0,"Byron25":475206.0,"Tifton25":301003.0,"Statesboro25":495204.0,"Athens25":403830.0},{"week":12,"Byron":601780.0,"Tifton":423219.0,"Statesboro":468571.0,"Athens":471418.0,"Byron25":453356.0,"Tifton25":300523.0,"Statesboro25":536190.0,"Athens25":364320.0},{"week":13,"Byron":644955.0,"Tifton":459113.0,"Statesboro":537096.0,"Athens":499530.0,"Byron25":397627.0,"Tifton25":315204.0,"Statesboro25":521326.0,"Athens25":359990.0},{"week":14,"Byron":695711.0,"Tifton":524967.0,"Statesboro":518940.0,"Athens":435410.0,"Byron25":531766.0,"Tifton25":342928.0,"Statesboro25":515104.0,"Athens25":404237.0},{"week":15,"Byron":722978.0,"Tifton":441601.0,"Statesboro":487888.0,"Athens":372553.0,"Byron25":591778.0,"Tifton25":382490.0,"Statesboro25":557293.0,"Athens25":549426.0},{"week":16,"Byron":581034.0,"Tifton":394978.0,"Statesboro":514215.0,"Athens":385858.0,"Byron25":596454.0,"Tifton25":334891.0,"Statesboro25":527256.0,"Athens25":454735.0},{"week":17,"Byron":631907.0,"Tifton":422233.0,"Statesboro":519390.0,"Athens":453314.0,"Byron25":484884.0,"Tifton25":357675.0,"Statesboro25":539418.0,"Athens25":324248.0},{"week":18,"Byron":550852.0,"Tifton":443646.0,"Statesboro":521227.0,"Athens":511785.0,"Byron25":519581.0,"Tifton25":396849.0,"Statesboro25":531112.0,"Athens25":347877.0},{"week":19,"Byron":533289.0,"Tifton":470135.0,"Statesboro":501654.0,"Athens":551444.0,"Byron25":595299.0,"Tifton25":430969.0,"Statesboro25":584156.0,"Athens25":381174.0},{"week":20,"Byron":537770.0,"Tifton":401595.0,"Statesboro":463117.0,"Athens":399092.0,"Byron25":488435.0,"Tifton25":327879.0,"Statesboro25":529185.0,"Athens25":381263.0},{"week":21,"Byron":581257.0,"Tifton":378909.0,"Statesboro":475853.0,"Athens":407850.0,"Byron25":550075.0,"Tifton25":395584.0,"Statesboro25":575972.0,"Athens25":405444.0},{"week":22,"Byron":385150.0,"Tifton":287880.0,"Statesboro":425248.0,"Athens":396320.0,"Byron25":444933.0,"Tifton25":347991.0,"Statesboro25":428728.0,"Athens25":387835.0},{"week":23,"Byron":553011,"Byron25":529599,"Byron_gp":82824,"Byron_gp25":82692,"Tifton":455396,"Tifton25":352730,"Tifton_gp":66364,"Tifton_gp25":59574,"Statesboro":480860,"Statesboro25":541660,"Statesboro_gp":58476,"Statesboro_gp25":83856,"Athens":435054,"Athens25":473748,"Athens_gp":62859,"Athens_gp25":77896}],"branches":{"Byron":{"q1_2025":5266440.63,"q1_2026":6822661.29,"q1_gp25":705844.68,"q1_gp26":861094.27,"q2_2025":5266298.78,"q2_2026":5479075.08,"q2_gp25":849992.41,"q2_gp26":740971.02},"Tifton":{"q1_2025":3825950.57,"q1_2026":5138331.47,"q1_gp25":566747.56,"q1_gp26":720485.34,"q2_2025":3593716.46,"q2_2026":4000509.33,"q2_gp25":629905.03,"q2_gp26":564225.82},"Statesboro":{"q1_2025":5835303.76,"q1_2026":5667773.51,"q1_gp25":745163.56,"q1_gp26":639736.99,"q2_2025":5217000.3,"q2_2026":4696448.24,"q2_gp25":808434.14,"q2_gp26":548717.54},"Athens":{"q1_2025":4287894.3,"q1_2026":5220784.25,"q1_gp25":577373.95,"q1_gp26":672226.68,"q2_2025":4013706.14,"q2_2026":4121381.95,"q2_gp25":645715.67,"q2_gp26":542438.81}},"tiftonQ1Depts":[{"dept":"1 - BYRON","sales2025":5266440.63,"sales2026":6822661.29,"gp2025":705844.68,"gp2026":861094.27},{"dept":"3 - STATESBORO","sales2025":5835303.76,"sales2026":5667773.51,"gp2025":745163.56,"gp2026":639736.99},{"dept":"5 - ATHENS","sales2025":4287894.3,"sales2026":5220784.25,"gp2025":577373.95,"gp2026":672226.68},{"dept":"2 - TIFTON","sales2025":3825950.57,"sales2026":5138331.47,"gp2025":566747.56,"gp2026":720485.34},{"dept":"RAD LT TRUCK","sales2025":1519974.15,"sales2026":2014306.08,"gp2025":222260.8,"gp2026":292854.91},{"dept":"TRUCK TIRES","sales2025":1039514.78,"sales2026":1133774.88,"gp2025":124413.63,"gp2026":111761.18},{"dept":"PASSENGER TIRES","sales2025":703217.54,"sales2026":1118847.12,"gp2025":138720.89,"gp2026":189943.94},{"dept":"FARM TIRES","sales2025":263687.6,"sales2026":278661.49,"gp2025":39236.24,"gp2026":33798.0},{"dept":"ST TRAILER","sales2025":111818.91,"sales2026":210417.35,"gp2025":21358.76,"gp2026":43126.35},{"dept":"OFF THE ROAD TIRES","sales2025":94922.14,"sales2026":167685.84,"gp2025":8105.2,"gp2026":14697.57},{"dept":"INDUSTRIAL TIRES","sales2025":26644.71,"sales2026":100424.94,"gp2025":3463.2,"gp2026":9692.53},{"dept":"TUBES","sales2025":41187.21,"sales2026":56254.35,"gp2025":8705.12,"gp2026":12141.58},{"dept":"GA EPD FEE","sales2025":20908.04,"sales2026":22857.42,"gp2025":0.04,"gp2026":0.42},{"dept":"VALVE STEMS","sales2025":527.18,"sales2026":16268.09,"gp2025":252.4,"gp2026":5252.01},{"dept":"WHEEL WEIGHTS","sales2025":206.21,"sales2026":9524.49,"gp2025":91.52,"gp2026":4431.67},{"dept":"TIRE TOOLS","sales2025":249.53,"sales2026":9513.72,"gp2025":67.84,"gp2026":2833.74},{"dept":"PATCHES AND REPAIR","sales2025":519.39,"sales2026":4511.29,"gp2025":161.42,"gp2026":1482.06},{"dept":"LAWN & GARDEN","sales2025":2881.42,"sales2026":4320.11,"gp2025":581.49,"gp2026":848.16},{"dept":"WHEELS","sales2025":1876.75,"sales2026":2216.06,"gp2025":285.71,"gp2026":635.91},{"dept":"MOUNTING LUBE","sales2025":260.79,"sales2026":1659.5,"gp2025":72.42,"gp2026":686.18},{"dept":"ALIGNMENT SHIMS","sales2025":0.0,"sales2026":719.96,"gp2025":0.0,"gp2026":243.44},{"dept":"FREIGHT CHARGES","sales2025":0.0,"sales2026":171.0,"gp2025":0.0,"gp2026":76.0},{"dept":"OUTSIDE PURCHASE","sales2025":354.31,"sales2026":96.0,"gp2025":66.38,"gp2026":0.0}]};

const SEED_CUSTOMERS = [];



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

const FILE_SLOT_GROUPS = [
  {
    label: "📦 Master Upload",
    desc: "Upload everything at once for your weekly meeting",
    slots: [
      { id:"master", label:"⬡ Master Upload", desc:"Pulse_Master_Upload_Wxx.xlsx — AR, WTD, CustomerComp, all AD programs in one click", isMaster:true },
    ]
  },
  {
    label: "📊 Core Data",
    desc: "Individual uploads for core sales data",
    slots: [
      { id:"customers", label:"Customer List", desc:"Account list with salesperson assignment" },
      { id:"ar",        label:"AR / Aging",    desc:"Balances · 30/60/90/120+ day aging" },
      { id:"weekComp",  label:"Week Comp",     desc:"Week-to-Week Customer Comp by Department" },
      { id:"sales",     label:"Sales Data",    desc:"Current & prior year sales" },
    ]
  },
  {
    label: "🏆 AD Programs",
    desc: "Upload individual AD program reports as they come in — no need to wait for master upload day",
    slots: [
      { id:"toyo",      label:"Toyo AD",        desc:"Associate Dealer report — paste or upload anytime" },
      { id:"americus",  label:"Americus",        desc:"Month-end program update" },
      { id:"ascenso",   label:"Ascenso",         desc:"HITS Sales Analysis — Sales by Customer (ASCENSO)" },
      { id:"falkenPLT", label:"Falken PLT",      desc:"Falken Fanatic PLT quarterly report" },
      { id:"falkenTBR", label:"Falken TBR",      desc:"Falken Fanatic TBR quarterly report" },
      { id:"barnn",     label:"BF BARNN",        desc:"Bridgestone/Firestone BARNN dealer report" },
      { id:"yokohama",  label:"Yokohama",        desc:"Yokohama dealer program — quarterly" },
    ]
  },
];
// Flat list for backward-compat
const FILE_SLOTS = FILE_SLOT_GROUPS.flatMap(g=>g.slots);


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
  return sheets.includes("AR") && sheets.includes("CustomerComp") &&
    (sheets.includes("WTD") || sheets.includes("QTD") || sheets.includes("Summary"));
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

// ── QTD sheet — same structure as WTD, drives branch Q2 totals ──────────────
function parseMasterQTD(wb) {
  if (!wb.Sheets["QTD"]) return null;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets["QTD"], { header:1, defval:null });

  const BRANCH_KEYS = {
    "1 - BYRON":"Byron","2 - TIFTON":"Tifton",
    "3 - STATESBORO":"Statesboro","5 - ATHENS":"Athens","TOTAL":"Total"
  };
  const branchData={}, tiftonDepts=[], byronDepts=[];
  let section=null;

  function num(v){ return typeof v==="number"?v:parseFloat(String(v||"0").replace(/[$,%]/g,""))||0; }

  for (const row of rows) {
    const raw0  = String(row[0]||"").trim();
    const cell0 = raw0.toUpperCase();

    // ── Branch rows first (exact match) ─────────────────────────────────────
    const bKey = Object.keys(BRANCH_KEYS).find(k=>cell0===k.toUpperCase());
    if (bKey) {
      const bName=BRANCH_KEYS[bKey];
      if (bName!=="Total") {
        branchData[bName]={
          sales2025:num(row[1]), sales2026:num(row[2]),
          gp2025:num(row[5]),    gp2026:num(row[6]),
        };
      }
      continue;
    }

    // ── Section headers ──────────────────────────────────────────────────────
    if (raw0.startsWith("①")||(cell0.includes("TIFTON")&&cell0.includes("DEPARTMENT"))){ section="tifton"; continue; }
    if (raw0.startsWith("②")||cell0.includes("BRANCH COMPARISON"))                     { section="branch"; continue; }
    if (raw0.startsWith("③")||(cell0.includes("BYRON")&&cell0.includes("DEPARTMENT"))) { section="byron";  continue; }

    // ── Skip headers/totals/empty ────────────────────────────────────────────
    if (!raw0||raw0.length<3) continue;
    if (cell0==="DEPARTMENT"||cell0==="BRANCH") continue;
    if (cell0.includes("SUBTOTAL")) continue;

    // ── Department rows ──────────────────────────────────────────────────────
    const s1=num(row[1]),s2=num(row[2]),g1=num(row[5]),g2=num(row[6]);
    if (s1===0&&s2===0) continue;
    const entry={dept:raw0, sales2025:s1, sales2026:s2, gp2025:g1, gp2026:g2,
      salesChange:s2-s1, gpPct2026:s2>0?g2/s2:0};
    if (section==="tifton") tiftonDepts.push(entry);
    else if (section==="byron") byronDepts.push(entry);
  }

  return Object.keys(branchData).length>0
    ? { branchData, tiftonDepts, byronDepts }
    : null;
}

// ── Summary sheet — explicit QTD/YTD/weekly values entered by user ──────────
function parseMasterSummary(wb) {
  if (!wb.Sheets["Summary"]) return null;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets["Summary"], { header:1, defval:null });
  const result = { weekNum:null, weekDate:"", tifton:{}, branches:{}, weekly:{} };
  const num = v => typeof v==="number" ? v : parseFloat(String(v||"0").replace(/[$,%,]/g,""))||0;

  for (let i=0; i<rows.length; i++) {
    const row  = rows[i];
    const col0 = String(row[0]||"").trim();
    const col0U = col0.toUpperCase();

    if (col0U === "WEEK NUMBER")      { result.weekNum  = num(row[1]); continue; }
    if (col0U === "WEEK ENDING DATE") { result.weekDate = String(row[1]||""); continue; }

    // Tifton rows (col A = metric, col B = 2025, col C = 2026)
    if (col0U.includes("Q1 SALES"))           { result.tifton.q1_2025=num(row[1]); result.tifton.q1_2026=num(row[2]); }
    if (col0U.includes("Q2 SALES"))           { result.tifton.q2_2025=num(row[1]); result.tifton.q2_2026=num(row[2]); }
    if (col0U.includes("Q1 GP"))              { result.tifton.q1_gp25=num(row[1]); result.tifton.q1_gp26=num(row[2]); }
    if (col0U.includes("Q2 GP"))              { result.tifton.q2_gp25=num(row[1]); result.tifton.q2_gp26=num(row[2]); }
    if (col0U === "TIFTON WEEKLY SALES")      { result.weekly.s25=num(row[1]);  result.weekly.s26=num(row[2]); }
    if (col0U === "TIFTON WEEKLY GP$")        { result.weekly.g25=num(row[1]);  result.weekly.g26=num(row[2]); }

    // Branch rows (col A = branch name, col B-E = Q2 values)
    for (const br of ["Byron","Statesboro","Athens"]) {
      if (col0U === br.toUpperCase()) {
        result.branches[br] = {
          q2_2025:num(row[1]), q2_2026:num(row[2]),
          q2_gp25:num(row[3]), q2_gp26:num(row[4]),
        };
      }
    }
  }
  return (result.weekNum || Object.keys(result.tifton).length > 0) ? result : null;
}

// ── WTD sheet ─────────────────────────────────────────────────────────────────
function parseMasterWTD(wb) {
  if (!wb.Sheets["WTD"]) return null;
  const rows = getSheetRows(wb, "WTD", 0);

  const BRANCH_KEYS = {
    "1 - BYRON":     "Byron",
    "2 - TIFTON":    "Tifton",
    "3 - STATESBORO":"Statesboro",
    "5 - ATHENS":    "Athens",
    "TOTAL":         "Total",
  };
  const branchData  = {};
  const tiftonDepts = [];
  const byronDepts  = [];
  let section = null;

  function num(v) { return typeof v==="string" ? parseFloat(v.replace(/[$,%]/g,""))||0 : Number(v||0); }

  function autoMetric(s1, s2, g1, g2) {
    const salesChg=s2-s1, gpChg=g2-g1;
    const gp1pct=s1>0?g1/s1:0, gp2pct=s2>0?g2/s2:0;
    if (salesChg>0 && gpChg>0)  return "★ Growth + GP up";
    if (salesChg>0 && gpChg<0)  return gp2pct<gp1pct-0.03?"★ Growth + margin compression":"✓ Growth";
    if (salesChg<0 && Math.abs(salesChg)>5000) return "⚠ Largest $ loss";
    if (salesChg<0) return "⚠ Decline";
    return "→ Flat";
  }

  function autoComment(dept, s1, s2, g1, g2) {
    const sDelta=s2-s1;
    const sPct=s1>0?((s2-s1)/s1*100).toFixed(1):"N/A";
    const gpChg=g2-g1, gpPctChg=Math.max(g1,1)>0?(Math.abs(gpChg)/Math.max(g1,1)*100).toFixed(1):0;
    const gp1=s1>0?(g1/s1*100).toFixed(1):"0", gp2=s2>0?(g2/s2*100).toFixed(1):"0";
    return `${dept}: Revenue ${sDelta>=0?"up":"down"} ${Math.abs(sPct)}% ($${Math.abs(sDelta).toLocaleString("en-US",{maximumFractionDigits:0})}). GP$ ${gpChg>=0?"up":"down"} ${gpPctChg}%. Margin ${gp1}%→${gp2}%.`;
  }

  for (const row of rows) {
    const raw0  = String(row[0] || "").trim();
    const cell0 = raw0.toUpperCase();

    // ── 1. BRANCH DATA — check exact match first (before section headers) ────
    const branchKey = Object.keys(BRANCH_KEYS).find(k => cell0 === k.toUpperCase());
    if (branchKey) {
      const bName = BRANCH_KEYS[branchKey];
      if (bName !== "Total") {
        const s1=num(row[1]), s2=num(row[2]);
        const g1=num(row[5]), g2=num(row[6]);  // GP$ Range 1 = col5, Range 2 = col6
        branchData[bName] = {
          sales2025: s1, sales2026: s2,
          gp2025:    g1, gp2026:    g2,
          salesChange:    s2-s1,
          salesChangePct: s1>0?(s2-s1)/s1:0,
          metric: autoMetric(s1,s2,g1,g2),
        };
      }
      continue;
    }

    // ── 2. SECTION HEADERS — only match the ① ② ③ structured headers ────────
    if (raw0.startsWith("①") || (cell0.includes("TIFTON") && cell0.includes("DEPARTMENT"))) {
      section = "tifton"; continue;
    }
    if (raw0.startsWith("②") || cell0.includes("BRANCH COMPARISON")) {
      section = "branch"; continue;
    }
    if (raw0.startsWith("③") || (cell0.includes("BYRON") && cell0.includes("DEPARTMENT"))) {
      section = "byron"; continue;
    }

    // ── 3. SKIP header rows, empty rows, subtotal rows ────────────────────────
    if (!raw0 || raw0.length < 3) continue;
    if (cell0 === "DEPARTMENT" || cell0 === "BRANCH" || cell0 === "LOCATION") continue;
    if (cell0.includes("SUBTOTAL")) continue;

    // ── 4. DEPARTMENT ROWS (Tifton or Byron sections) ─────────────────────────
    const s1=num(row[1]), s2=num(row[2]);
    if (s1===0 && s2===0) continue;
    const g1=num(row[5]), g2=num(row[6]);
    const entry = {
      dept: raw0,
      sales2025:s1, sales2026:s2, gp2025:g1, gp2026:g2,
      salesChange:s2-s1, salesChangePct:s1>0?(s2-s1)/s1:0,
      gpChange:g2-g1, gpChangePct:g1>0?(g2-g1)/g1:0,
      gpPct2025:s1>0?g1/s1:0, gpPct2026:s2>0?g2/s2:0,
      metric:  String(row[13]||"").trim() || autoMetric(s1,s2,g1,g2),
      comment: String(row[14]||"").trim() || autoComment(raw0,s1,s2,g1,g2),
    };
    if (section==="tifton") tiftonDepts.push(entry);
    else if (section==="byron") byronDepts.push(entry);
  }

  return {
    branchData:  Object.keys(branchData).length > 0 ? branchData : null,
    tiftonDepts: tiftonDepts.length > 0 ? tiftonDepts : null,
    byronDepts:  byronDepts.length  > 0 ? byronDepts  : null,
  };
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
    summary:    parseMasterSummary(wb),
    qtd:        parseMasterQTD(wb),
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
  const [fileData, setFileData] = useState(() => {
    // Load persisted data from localStorage, fall back to seeds
    let weekComp = SEED_WEEK_COMP;
    let ar = SEED_AR;
    let adData = {};
    try {
      const savedWeeks = localStorage.getItem("pulse_weeks");
      if (savedWeeks) {
        const parsed = JSON.parse(savedWeeks);
        // Merge saved weeks on top of seed weeks
        const mergedWeeks = [...SEED_WEEK_COMP.weeks];
        parsed.forEach(w => {
          const idx = mergedWeeks.findIndex(s => s.week === w.week);
          if (idx >= 0) mergedWeeks[idx] = w;
          else mergedWeeks.push(w);
        });
        mergedWeeks.sort((a,b) => a.week - b.week);
        weekComp = { ...SEED_WEEK_COMP, weeks: mergedWeeks };
      }
      // Only load localStorage AP if it matches seed version — prevents stale data overriding seed
      const savedAPVersion = localStorage.getItem("pulse_ap_version");
      if (savedAPVersion === SEED_WEEK_COMP.apVersion) {
        const savedAP = localStorage.getItem("pulse_action_plan");
        if (savedAP) weekComp = { ...weekComp, actionPlan: JSON.parse(savedAP) };
        const savedDepts = localStorage.getItem("pulse_depts");
        if (savedDepts) weekComp = { ...weekComp, depts: JSON.parse(savedDepts) };
      } else {
        // Seed is newer — clear stale localStorage data
        localStorage.removeItem("pulse_action_plan");
        localStorage.removeItem("pulse_depts");
        localStorage.setItem("pulse_ap_version", SEED_WEEK_COMP.apVersion);
      }
      const savedAR = localStorage.getItem("pulse_ar_data");
      if (savedAR) ar = JSON.parse(savedAR);
      const savedAD = localStorage.getItem("pulse_ad_data");
      if (savedAD) adData = JSON.parse(savedAD);
      // Rebuild branchData from SEED + stored WTD contributions NEWER than seed
      const storedWTD = JSON.parse(localStorage.getItem("pulse_wtd_weekly") || "{}");
      const seedWeek  = SEED_BRANCH_DATA.seedWeek || 0;
      // Only apply contributions for weeks the seed doesn't already include
      const newContribs = Object.entries(storedWTD).filter(([wk]) => Number(wk) > seedWeek);
      if (newContribs.length > 0) {
        const bd = JSON.parse(JSON.stringify(SEED_BRANCH_DATA));
        newContribs.forEach(([,wkContrib]) => {
          Object.entries(wkContrib).forEach(([bName, v]) => {
            if (bd.branches?.[bName]) {
              bd.branches[bName].q2_2026 = (bd.branches[bName].q2_2026 || 0) + (v.s26 || 0);
              bd.branches[bName].q2_gp26 = (bd.branches[bName].q2_gp26 || 0) + (v.g26 || 0);
              bd.branches[bName].q2_2025 = (bd.branches[bName].q2_2025 || 0) + (v.s25 || 0);
              bd.branches[bName].q2_gp25 = (bd.branches[bName].q2_gp25 || 0) + (v.g25 || 0);
            }
          });
        });
        // Rebuild weeklySales from SEED + stored weekly chart entries
        const storedWkEntries = JSON.parse(localStorage.getItem("pulse_ws_entries") || "{}");
        const uploadedWkNums  = Object.keys(storedWkEntries).map(Number);
        bd.weeklySales = [
          ...(SEED_BRANCH_DATA.weeklySales||[]).filter(w => !uploadedWkNums.includes(w.week)),
          ...Object.values(storedWkEntries),
        ].sort((a,b) => a.week - b.week);
        adData.branchData = bd;
      }
    } catch {}
    return { weekComp, customers: SEED_CUSTOMERS, ar, ...adData };
  });
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

  async function handleUpload(slotId, wb, filename="") {
    let parsed;

    // ── MASTER WORKBOOK DETECTION ─────────────────────────────────────────
    if (isMasterWorkbook(wb)) {
      const m = parseMasterWorkbook(wb);
      const summary = [];

      // ── Week number and branch totals: Summary sheet takes priority ──────────
      const sumSheet = m.summary;
      let weekNum = sumSheet?.weekNum || null;
      if (!weekNum) {
        const wkMatch = filename.match(/W(\d{1,2})/i);
        weekNum = wkMatch ? parseInt(wkMatch[1]) : null;
      }
      if (!weekNum) {
        const prompted = window.prompt("Which week is this upload for? Enter week number (e.g. 23):");
        weekNum = prompted ? parseInt(prompted) : null;
      }

      // ── Branch Q2/QTD update — always runs, QTD sheet gives definitive numbers ─
      const qtdData = m.qtd?.branchData ? m.qtd : null;
      // Always update branchData — QTD sheet → Summary fallback → WTD weekly calc
      {
        const newBD = JSON.parse(JSON.stringify(SEED_BRANCH_DATA));

        if (qtdData) {
          // ── QTD sheet: read branch Row 2 (Range 2 = current QTD) for each branch
          Object.entries(qtdData.branchData).forEach(([bName, d]) => {
            if (newBD.branches[bName]) {
              newBD.branches[bName].q2_2026 = d.sales2026 || 0;
              newBD.branches[bName].q2_2025 = d.sales2025 || 0;
              newBD.branches[bName].q2_gp26 = d.gp2026    || 0;
              newBD.branches[bName].q2_gp25 = d.gp2025    || 0;
            }
          });
          const tifQTD = qtdData.branchData.Tifton || {};
          summary.push(`QTD (from sheet): Tifton $${((tifQTD.sales2026||0)/1000000).toFixed(2)}M · YTD $${(((newBD.branches.Tifton.q1_2026||0)+(tifQTD.sales2026||0))/1000000).toFixed(2)}M`);
        } else {
          // ── Summary sheet fallback ────────────────────────────────────────
          const tif = sumSheet.tifton;
          if (tif.q1_2026) newBD.branches.Tifton.q1_2026 = tif.q1_2026;
          if (tif.q1_2025) newBD.branches.Tifton.q1_2025 = tif.q1_2025;
          if (tif.q2_2026) newBD.branches.Tifton.q2_2026 = tif.q2_2026;
          if (tif.q2_2025) newBD.branches.Tifton.q2_2025 = tif.q2_2025;
          if (tif.q1_gp26) newBD.branches.Tifton.q1_gp26 = tif.q1_gp26;
          if (tif.q1_gp25) newBD.branches.Tifton.q1_gp25 = tif.q1_gp25;
          if (tif.q2_gp26) newBD.branches.Tifton.q2_gp26 = tif.q2_gp26;
          if (tif.q2_gp25) newBD.branches.Tifton.q2_gp25 = tif.q2_gp25;
          Object.entries(sumSheet.branches||{}).forEach(([br,v])=>{
            if (newBD.branches[br]){
              if(v.q2_2026) newBD.branches[br].q2_2026=v.q2_2026;
              if(v.q2_2025) newBD.branches[br].q2_2025=v.q2_2025;
              if(v.q2_gp26) newBD.branches[br].q2_gp26=v.q2_gp26;
              if(v.q2_gp25) newBD.branches[br].q2_gp25=v.q2_gp25;
            }
          });
          summary.push(`QTD (summary): Tifton $${((tif.q2_2026||0)/1000000).toFixed(2)}M`);
        }

        // ── Weekly chart entry (WTD weekly data for the per-week chart) ──────
        const wtdBD = m.wtd?.branchData;
        if (weekNum && wtdBD) {
          const wsList = [...(SEED_BRANCH_DATA.weeklySales||[])];
          const wsIdx  = wsList.findIndex(w=>w.week===weekNum);
          const wsE    = { week:weekNum };
          Object.entries(wtdBD).forEach(([bName,d])=>{
            wsE[bName]        = Math.round(d.sales2026||0);
            wsE[`${bName}25`] = Math.round(d.sales2025||0);
          });
          if (wsIdx>=0) wsList[wsIdx]=wsE; else wsList.push(wsE);
          wsList.sort((a,b)=>a.week-b.week);
          newBD.weeklySales = wsList;
          localStorage.setItem("pulse_ws_entries",
            JSON.stringify(Object.fromEntries(wsList.filter(w=>w.week>22).map(w=>[w.week,w]))));
        }

        localStorage.setItem("pulse_branch_data", JSON.stringify(newBD));
        setFileData(prev => ({ ...prev, branchData: newBD }));
      }

      // ── Build week entry from WTD — Tifton only ──────────────────────────────
      if (weekNum && m.wtd?.branchData) {
        const bd  = m.wtd.branchData;
        const tif = bd["Tifton"] || {};

        const newWeekEntry = {
          week:      weekNum,
          sales2025: tif.sales2025 || 0,
          sales2026: tif.sales2026 || 0,
          change:    (tif.sales2026||0) - (tif.sales2025||0),
          changePct: tif.sales2025 > 0 ? ((tif.sales2026||0)-(tif.sales2025||0))/tif.sales2025 : 0,
          gp2025:    tif.gp2025 || 0,
          gp2026:    tif.gp2026 || 0,
          gpChange:  (tif.gp2026||0) - (tif.gp2025||0),
          locations: bd,
          tiftonDepts: m.wtd.tiftonDepts || [],
          byronDepts:  m.wtd.byronDepts  || [],
        };

        setFileData(prev => {
          const existing = prev.weekComp || SEED_WEEK_COMP;
          let mergedWeeks = [...(existing.weeks || [])];
          const idx = mergedWeeks.findIndex(w => w.week === weekNum);
          if (idx >= 0) mergedWeeks[idx] = newWeekEntry;
          else mergedWeeks.push(newWeekEntry);
          mergedWeeks.sort((a,b) => a.week - b.week);
          // Persist ALL uploaded weeks (not seeds) to localStorage
          const uploadedWeeks = mergedWeeks.filter(w =>
            !SEED_WEEK_COMP.weeks.find(s => s.week === w.week && s.sales2026 === w.sales2026)
          );
          try { localStorage.setItem("pulse_weeks", JSON.stringify(mergedWeeks)); } catch {}
          // Detect as new week for AI banner
          if (idx < 0) {
            localStorage.setItem("pulse_new_weeks", JSON.stringify([weekNum]));
            setNewWeeksDetected([weekNum]);
          }
          return { ...prev, weekComp: { ...existing, weeks: mergedWeeks } };
        });
        summary.push(`W${weekNum} added to Overview`);
      }

      // AR
      if (m.ar && m.ar.length > 0) {
        setFileData(prev => ({ ...prev, ar: m.ar }));
        try { localStorage.setItem("pulse_ar_data", JSON.stringify(m.ar)); } catch {}
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
          try { localStorage.setItem("pulse_action_plan", JSON.stringify(mergedAP)); } catch {}
          return { ...prev, weekComp: { ...existing, actionPlan: mergedAP } };
        });
        summary.push(`CustomerComp: ${m.customerComp.actionPlan.length} accounts`);

        // ── Store as WTD or QTD period for Customer Comp view ─────────────
        // Build per-customer period list from the uploaded CC data
        if (m.customerComp?.actionPlan?.length > 0) {
          const periodEntries = m.customerComp.actionPlan.map(a => ({
            custNum:   a.custNum,
            customer:  a.customer,
            sales2026: a.sales2026 || 0,
            sales2025: a.sales2025 || 0,
            gp2026:    (a.sales2026||0) * (a.gpPct||0),
            change:    (a.sales2026||0) - (a.sales2025||0),
            changePct: a.sales2025>0 ? ((a.sales2026||0)-(a.sales2025||0))/a.sales2025 : 0,
            salesman:  a.salesman || "House",
          })).sort((a,b) => b.sales2026 - a.sales2026);

          // Detect period type by total volume
          // QTD (whole quarter) will be much larger than WTD (one week)
          const totalVol = periodEntries.reduce((s,r) => s+r.sales2026, 0);
          const seedWeeklyAvg = (SEED_BRANCH_DATA.branches?.Tifton?.q2_2026||0) / 23;
          const isQTD = totalVol > seedWeeklyAvg * 3;  // more than 3 weeks = likely QTD

          const periodType = isQTD ? "qtd" : "wtd";
          const periodLabel = isQTD
            ? `QTD — Q2 2026 (${new Date().toLocaleDateString("en-US",{month:"short",day:"numeric"})})`
            : `WTD — W${weekNum||"?"} (${new Date().toLocaleDateString("en-US",{month:"short",day:"numeric"})})`;

          setFileData(prev => {
            const wc = prev.weekComp || SEED_WEEK_COMP;
            const periods = { ...(wc.periods||{}) };
            periods[periodType] = { label: periodLabel, data: periodEntries };
            try {
              const pd = JSON.parse(localStorage.getItem("pulse_periods") || "{}");
              pd[periodType] = { label: periodLabel, data: periodEntries };
              localStorage.setItem("pulse_periods", JSON.stringify(pd));
            } catch {}
            return { ...prev, weekComp: { ...wc, periods } };
          });
          summary.push(`${periodType.toUpperCase()} comp: ${periodEntries.length} customers`);
        }
      }

      // ── Update branchData QTD/YTD from WTD branch comparison ────────────────
      if (m.wtd?.branchData && weekNum) {
        try {
          // Start from current branchData (already in fileData) or SEED
          const curBD = (() => {
            try {
              const s = localStorage.getItem("pulse_branch_data");
              return s ? JSON.parse(s) : SEED_BRANCH_DATA;
            } catch { return SEED_BRANCH_DATA; }
          })();
          const newBD = JSON.parse(JSON.stringify(curBD)); // deep copy

          // Per-week tracking to avoid double-count on re-upload
          const WTD_WK_KEY = "pulse_wtd_weekly";
          const wtdWeekly  = (() => {
            try { return JSON.parse(localStorage.getItem(WTD_WK_KEY) || "{}"); } catch { return {}; }
          })();

          // Remove previous contribution for this week (if re-upload)
          const prev = wtdWeekly[weekNum] || {};
          Object.entries(prev).forEach(([br, v]) => {
            if (newBD.branches?.[br]) {
              newBD.branches[br].q2_2026 -= (v.s26 || 0);
              newBD.branches[br].q2_gp26 -= (v.g26 || 0);
              newBD.branches[br].q2_2025 -= (v.s25 || 0);
              newBD.branches[br].q2_gp25 -= (v.g25 || 0);
            }
          });

          // Add this week's WTD contribution for each branch
          const newContrib = {};
          Object.entries(m.wtd.branchData).forEach(([bName, d]) => {
            if (newBD.branches?.[bName]) {
              const s26 = d.sales2026 || 0, g26 = d.gp2026 || 0;
              const s25 = d.sales2025 || 0, g25 = d.gp2025 || 0;
              newBD.branches[bName].q2_2026 = (newBD.branches[bName].q2_2026 || 0) + s26;
              newBD.branches[bName].q2_gp26 = (newBD.branches[bName].q2_gp26 || 0) + g26;
              newBD.branches[bName].q2_2025 = (newBD.branches[bName].q2_2025 || 0) + s25;
              newBD.branches[bName].q2_gp25 = (newBD.branches[bName].q2_gp25 || 0) + g25;
              newContrib[bName] = { s26, g26, s25, g25 };
            }
          });
          wtdWeekly[weekNum] = newContrib;

          // ── Rebuild weeklySales: always start from SEED + all stored weekly entries ──
          // This ensures weeklySales is never lost if localStorage has partial data
          const storedWkEntries = JSON.parse(localStorage.getItem("pulse_ws_entries") || "{}");
          // Add/replace this week's entry
          const wsEntry = { week: weekNum };
          Object.entries(m.wtd.branchData).forEach(([bName, d]) => {
            wsEntry[bName]          = Math.round(d.sales2026 || 0);
            wsEntry[`${bName}25`]   = Math.round(d.sales2025 || 0);
            wsEntry[`${bName}_gp`]  = Math.round(d.gp2026    || 0);
            wsEntry[`${bName}_gp25`]= Math.round(d.gp2025    || 0);
          });
          storedWkEntries[weekNum] = wsEntry;
          localStorage.setItem("pulse_ws_entries", JSON.stringify(storedWkEntries));
          // Rebuild weeklySales = SEED W1-W22 + all stored uploaded weeks
          const seedWS  = SEED_BRANCH_DATA.weeklySales || [];
          const uploadedWkNums = Object.keys(storedWkEntries).map(Number);
          const wsList  = [
            ...seedWS.filter(w => !uploadedWkNums.includes(w.week)),
            ...Object.values(storedWkEntries),
          ].sort((a,b) => a.week - b.week);
          newBD.weeklySales = wsList;

          localStorage.setItem(WTD_WK_KEY, JSON.stringify(wtdWeekly));
          localStorage.setItem("pulse_branch_data", JSON.stringify(newBD));
          setFileData(prev => ({ ...prev, branchData: newBD }));
          summary.push(`QTD updated — Tifton Q2: $${((newBD.branches?.Tifton?.q2_2026||0)/1000000).toFixed(2)}M`);
        } catch(e) { console.warn("Branch update error:", e.message); }
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
        // Persist AD program data
        try {
          const existing = JSON.parse(localStorage.getItem("pulse_ad_data") || "{}");
          localStorage.setItem("pulse_ad_data", JSON.stringify({ ...existing, ...adUpdates }));
        } catch {}
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


    // ── INDIVIDUAL AD PROGRAM UPLOADS ─────────────────────────────────────────
    // Each reads first sheet of the uploaded file (not a named sheet)
    function readFirstSheet(wb) {
      return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:null });
    }

    if (slotId === "toyo") {
      // Same format as Toyo sheet — filter Tifton rows
      const rows = readFirstSheet(wb);
      const result = {};
      for (const row of rows) {
        if (!row || !String(row[5]||"").includes("Tifton")) continue;
        const toyoNum = String(row[2]||"").trim();
        if (!toyoNum || toyoNum === "Number") continue;
        result[toyoNum] = {
          toyoNum, dealerName: String(row[3]||"").trim(),
          pcr:{ primary:Number(row[6]||0), secondary:Number(row[7]||0), total:Number(row[8]||0),
                pct:parseFloat(String(row[9]||"0").replace("%",""))||0 },
          tbr:{ primary:Number(row[10]||0), secondary:Number(row[11]||0), total:Number(row[12]||0) },
        };
      }
      if (Object.keys(result).length > 0) {
        setFileData(prev=>({...prev, toyoData:result}));
        setNotice(`✓ Toyo updated — ${Object.keys(result).length} Tifton dealers`);
      } else setNotice("⚠ No Tifton dealers found in Toyo file — check Primary Dealer column");
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
      return;
    }

    if (slotId === "americus") {
      const rows = readFirstSheet(wb);
      const result = {};
      for (const row of rows) {
        const dn = row[0];
        if (!dn || typeof dn !== "number") continue;
        const name = String(row[1]||"").trim();
        if (!name || name.toLowerCase().includes("total")) continue;
        result[String(dn)] = { dealerNum:dn, dealerName:name,
          units2025:Number(row[6]||0), jan:Number(row[7]||0), feb:Number(row[8]||0),
          mar:Number(row[9]||0), q1:Number(row[10]||0), apr:Number(row[11]||0),
          may:Number(row[12]||0), jun:Number(row[13]||0), q2:Number(row[14]||0),
          ytd:Number(row[15]||0), asOf:new Date().toLocaleDateString("en-US",{month:"short",year:"numeric"}) };
      }
      if (Object.keys(result).length>0) { setFileData(prev=>({...prev,americusData:result}));
        setNotice(`✓ Americus updated — ${Object.keys(result).length} dealers`); }
      else setNotice("⚠ Americus file could not be parsed");
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
      return;
    }

    if (slotId === "ascenso") {
      const rows = readFirstSheet(wb);
      const result = {};
      for (const row of rows) {
        const cn = row[0];
        if (!cn || typeof cn !== "number") continue;
        result[String(cn)] = { custNum:cn, name:String(row[1]||"").trim(),
          qty:Number(row[3]||0), amount:Number(row[4]||0) };
      }
      if (Object.keys(result).length>0) { setFileData(prev=>({...prev,ascensoData:result}));
        setNotice(`✓ Ascenso updated — ${Object.keys(result).length} customers, $${Object.values(result).reduce((s,r)=>s+r.amount,0).toLocaleString("en-US",{maximumFractionDigits:0})}`); }
      else setNotice("⚠ Ascenso file could not be parsed");
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
      return;
    }

    if (slotId === "falkenPLT" || slotId === "falkenTBR") {
      const rows = readFirstSheet(wb);
      const result = [];
      for (const row of rows) {
        const fid = row[0];
        if (!fid || typeof fid !== "number") continue;
        result.push({ fanId:fid, segment:String(row[1]||"").trim(), parentId:row[2],
          dealer:String(row[3]||"").trim(), city:String(row[7]||"").trim(),
          territory:String(row[9]||"").trim(),
          q1:Number(row[10]||0), q2:Number(row[11]||0), q3:Number(row[12]||0),
          q4:Number(row[13]||0), ytd:Number(row[14]||0) });
      }
      const key = slotId === "falkenPLT" ? "falkenPLT" : "falkenTBR";
      if (result.length>0) { setFileData(prev=>({...prev,[key]:result}));
        setNotice(`✓ Falken ${slotId==="falkenPLT"?"PLT":"TBR"} updated — ${result.length} dealers`); }
      else setNotice(`⚠ Falken ${slotId==="falkenPLT"?"PLT":"TBR"} file could not be parsed`);
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
      return;
    }

    if (slotId === "barnn") {
      const rows = readFirstSheet(wb);
      const result = { primary:[], secondary:[] };
      for (const row of rows) {
        const parent=row[0], acctNum=row[1], name=String(row[2]||"").trim();
        if (!acctNum || typeof acctNum !== "number" || name.toLowerCase().includes("total")) continue;
        const entry = { parent:String(parent||""), acctNum, name,
          bs:Number(row[5]||0), fs:Number(row[8]||0), total:Number(row[15]||0),
          role: String(parent)==="960788"?"Primary":"Secondary" };
        if (String(parent)==="960788") result.primary.push(entry);
        else result.secondary.push(entry);
      }
      const count = result.primary.length + result.secondary.length;
      if (count>0) { setFileData(prev=>({...prev,barnnData:result}));
        setNotice(`✓ BARNN updated — ${result.primary.length} primary, ${result.secondary.length} secondary accounts`); }
      else setNotice("⚠ BARNN file could not be parsed");
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
      return;
    }

    if (slotId === "yokohama") {
      const rows = readFirstSheet(wb);
      const result = [];
      for (const row of rows) {
        const dn = row[0];
        if (!dn || typeof dn !== "number") continue;
        result.push({ dealerNum:dn, dealerName:String(row[1]||"").trim(),
          salesRep:String(row[2]||"").trim(), primary:Number(row[4]||0),
          priPct:parseFloat(String(row[5]||"0").replace("%",""))||0,
          hasSecondary:String(row[6]||"").trim().toUpperCase()==="Y",
          secondary:Number(row[8]||0), qtd:Number(row[10]||0), toNext:Number(row[11]||0) });
      }
      if (result.length>0) { setFileData(prev=>({...prev,yokohamaData:result}));
        setNotice(`✓ Yokohama updated — ${result.length} dealers`); }
      else setNotice("⚠ Yokohama file could not be parsed");
      setTimeout(()=>setNotice(""),5000);
      const nd={...uploadDates,[slotId]:new Date().toISOString()};
      setUploadDates(nd); try{localStorage.setItem("upload_dates",JSON.stringify(nd));}catch{}
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
        try { localStorage.setItem("pulse_ar_data", JSON.stringify(arParsed)); } catch {}
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
        {tab === "overview" && <OverviewTab weekComp={weekComp} branchData={fileData.branchData || SEED_BRANCH_DATA} onAskAI={goAI} onCustomerClick={openCustomer} customers={fileData.customers || SEED_CUSTOMERS} user={currentUser} />}
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
function FileSetup({ fileData, onUpload, uploadDates }) {
  const fmt = iso => iso ? new Date(iso).toLocaleDateString("en-US",{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"}) : null;

  return (
    <div>
      <div style={{ fontSize:"0.78rem", fontWeight:700, color:TEXT, marginBottom:"1rem" }}>
        ⬡ Data Upload
      </div>
      {FILE_SLOT_GROUPS.map(group => (
        <div key={group.label} style={{ marginBottom:"1.25rem" }}>
          {/* Group header */}
          <div style={{ display:"flex", alignItems:"baseline", gap:8, marginBottom:"0.45rem" }}>
            <div style={{ fontSize:"0.73rem", fontWeight:700, color:TEXT }}>{group.label}</div>
            <div style={{ fontSize:"0.63rem", color:MUTED }}>{group.desc}</div>
          </div>
          {/* Slots in this group */}
          <div style={{ display:"flex", flexDirection:"column", gap:"0.45rem" }}>
            {group.slots.map(slot => {
              const uploaded = uploadDates?.[slot.id];
              return (
                <div key={slot.id}
                  style={{ ...S.card,
                    border: slot.isMaster ? "2px solid #1E5FCC" : `1px solid ${BORDER}`,
                    background: slot.isMaster ? "#EFF6FF" : undefined,
                    padding:"0.6rem 0.85rem",
                    display:"flex", justifyContent:"space-between", alignItems:"center",
                    gap:8, flexWrap:"wrap" }}>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:"0.75rem", fontWeight:700,
                      color: slot.isMaster ? "#1E5FCC" : TEXT }}>{slot.label}</div>
                    <div style={{ fontSize:"0.63rem", color:MUTED, marginTop:1 }}>{slot.desc}</div>
                    {uploaded && (
                      <div style={{ fontSize:"0.6rem", color:"#059669", marginTop:2 }}>
                        ✓ Uploaded {fmt(uploaded)}
                      </div>
                    )}
                  </div>
                  <label style={{ cursor:"pointer", flexShrink:0 }}>
                    <div style={{ fontSize:"0.7rem", fontWeight:700,
                      color: slot.isMaster ? "#fff" : "#1E5FCC",
                      background: slot.isMaster ? "#1E5FCC" : "#EFF6FF",
                      border:`1px solid ${slot.isMaster?"#1E5FCC":"#BFDBFE"}`,
                      borderRadius:6, padding:"0.35rem 0.85rem",
                      whiteSpace:"nowrap", userSelect:"none" }}>
                      {uploaded ? "↺ Re-upload" : "↑ Upload"}
                    </div>
                    <input type="file" accept=".xlsx,.xls" style={{ display:"none" }}
                      onChange={e => {
                        const file = e.target.files[0]; if (!file) return;
                        const reader = new FileReader();
                        reader.onload = ev => {
                          const wb = XLSX.read(ev.target.result, { type:"array" });
                          onUpload(slot.id, wb, file.name);
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
      ))}
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
  const [showUpload, setShowUpload]   = useState(false);
  const [transcript, setTranscript]   = useState("");
  const [meetingDate, setMeetingDate] = useState(new Date().toISOString().slice(0,10));
  const [generating, setGenerating]   = useState(false);
  const [showPrevious, setShowPrevious] = useState(false);

  function saveMeetings(updated) {
    setMeetings(updated);
    try { localStorage.setItem(MSR_KEY, JSON.stringify(updated)); } catch {}
  }

  // Latest meeting = first in array; rest are archived
  const [current, ...archived] = meetings.sort((a,b) => new Date(b.date)-new Date(a.date));

  async function generateSummary() {
    if (!transcript.trim()) return;
    setGenerating(true);
    const prompt = `You are analyzing a tire distribution company's Monday Sales Review meeting transcript.

TRANSCRIPT:
${transcript.slice(0, 12000)}

Generate a structured meeting summary as JSON only (no markdown, no backticks):
{
  "headline": "One sentence capturing the most important takeaway",
  "weekNum": <week number if mentioned or null>,
  "sections": [
    { "label": "emoji + Section Title", "color": "#hexcolor", "bullets": ["bullet", "bullet"] }
  ],
  "tiftonNotes": "Any Tifton-specific items, action items, or callouts",
  "briefSummary": "2-3 sentence plain-english recap of what happened this week vs last week",
  "createdBy": "AI Summary"
}
Section labels: 📊 Performance, 🏆 Departments, 📅 YTD, 📦 Inventory, 💰 AR Collections, 🎯 Promotions, 🔑 Key Directives.
Colors: red=#DC2626 (down), green=#059669 (up), teal=#0891B2 (info), amber=#D97706 (caution), purple=#7C3AED (inventory), blue=#1E5FCC (strategy).
Be specific — use actual numbers, names, and account names from the transcript.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json", "x-api-key":ANTHROPIC_KEY,
          "anthropic-version":"2023-06-01", "anthropic-dangerous-direct-browser-access":"true" },
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:2000,
          messages:[{role:"user",content:prompt}] })
      });
      const data = await res.json();
      const raw  = data.content?.[0]?.text || "{}";
      const parsed = JSON.parse(raw.replace(/```json|```/g,"").trim());
      const newMeeting = {
        id:          `msr_${Date.now()}`,
        date:        meetingDate,
        weekNum:     parsed.weekNum || null,
        title:       `W${parsed.weekNum||"?"} Review — ${new Date(meetingDate).toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}`,
        headline:    parsed.headline || "",
        sections:    parsed.sections || [],
        tiftonNotes: parsed.tiftonNotes || "",
        briefSummary:parsed.briefSummary || "",
        createdBy:   "AI Summary",
      };
      saveMeetings([newMeeting, ...meetings]);
      setShowUpload(false);
      setTranscript("");
    } catch(e) { alert(`Summary failed: ${e.message}`); }
    setGenerating(false);
  }

  function deleteMeeting(id) {
    if (!window.confirm("Remove this meeting summary?")) return;
    saveMeetings(meetings.filter(m=>m.id!==id));
  }

  function MeetingCard({ m, compact=false }) {
    const [open, setOpen] = useState(!compact);
    return (
      <div style={{ ...S.card, marginBottom: compact?"0.5rem":"1rem",
        border: compact?`1px solid ${BORDER}`:"none",
        padding: compact?"0.65rem 0.85rem":undefined }}>
        {compact && (
          <div onClick={()=>setOpen(!open)}
            style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
              cursor:"pointer", userSelect:"none" }}>
            <div>
              <span style={{ fontSize:"0.75rem", fontWeight:700, color:TEXT }}>{m.title}</span>
              {m.briefSummary && (
                <div style={{ fontSize:"0.68rem", color:MUTED, marginTop:2,
                  maxWidth:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                  {m.briefSummary}
                </div>
              )}
            </div>
            <span style={{ fontSize:"0.75rem", color:MUTED, marginLeft:8, flexShrink:0 }}>
              {open?"▾":"▸"}
            </span>
          </div>
        )}
        {(!compact || open) && (
          <div style={{ marginTop: compact&&open?"0.65rem":0 }}>
            {/* Headline card */}
            {!compact && (
              <div style={{ padding:"0.75rem 1rem",
                background:"linear-gradient(135deg,#EFF6FF,#F0FDF4)",
                border:"1px solid #BFDBFE", borderRadius:8, marginBottom:"1rem" }}>
                <div style={{ fontSize:"0.63rem", color:"#1E5FCC", fontWeight:700,
                  textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>
                  {m.title} · {new Date(m.date).toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"})}
                </div>
                <div style={{ fontSize:"0.82rem", fontWeight:700, color:TEXT, lineHeight:1.6 }}>
                  {m.headline}
                </div>
              </div>
            )}

            {/* Brief summary (compact mode shows below toggle) */}
            {compact && m.briefSummary && (
              <div style={{ fontSize:"0.73rem", color:TEXT, lineHeight:1.65,
                marginBottom:"0.65rem", padding:"0.5rem 0.65rem",
                background:"#F8FAFC", borderRadius:6 }}>
                {m.briefSummary}
              </div>
            )}

            {/* Sections */}
            <div style={{ display:"flex", flexDirection:"column", gap:"0.6rem", marginBottom:"0.75rem" }}>
              {(m.sections||[]).map((sec,i) => (
                <div key={i} style={{ ...S.card, borderLeft:`4px solid ${sec.color}`,
                  padding:"0.65rem 0.85rem" }}>
                  <div style={{ fontSize:"0.68rem", fontWeight:700, color:sec.color,
                    textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:"0.4rem" }}>
                    {sec.label}
                  </div>
                  {(sec.bullets||[]).map((b,j)=>(
                    <div key={j} style={{ display:"flex", gap:8, fontSize:"0.74rem",
                      color:TEXT, lineHeight:1.65, marginBottom:2 }}>
                      <span style={{ color:sec.color, flexShrink:0 }}>▸</span>
                      <span>{b}</span>
                    </div>
                  ))}
                </div>
              ))}
            </div>

            {/* Tifton notes */}
            {m.tiftonNotes && (
              <div style={{ ...S.card, background:"#FFFBEB", borderLeft:`4px solid ${AMBER}`,
                padding:"0.65rem 0.85rem", marginBottom:"0.5rem" }}>
                <div style={{ fontSize:"0.67rem", fontWeight:700, color:AMBER,
                  textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:"0.4rem" }}>
                  📍 Tifton Notes
                </div>
                <div style={{ fontSize:"0.74rem", color:TEXT, lineHeight:1.7 }}>{m.tiftonNotes}</div>
              </div>
            )}

            {/* Footer */}
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
              fontSize:"0.6rem", color:MUTED }}>
              <span>{m.createdBy} · {new Date(m.date).toLocaleDateString()}</span>
              <button onClick={()=>deleteMeeting(m.id)}
                style={{ fontSize:"0.6rem", color:RED, background:"none",
                  border:"none", cursor:"pointer", padding:0 }}>
                × Remove
              </button>
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
        marginBottom:"0.85rem", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontSize:"0.82rem", fontWeight:700, color:"#1E5FCC" }}>📋 Monday Sales Review</div>
          <div style={{ fontSize:"0.68rem", color:MUTED }}>
            {meetings.length} meeting{meetings.length!==1?"s":""} · {archived.length} archived
          </div>
        </div>
        <button onClick={()=>setShowUpload(!showUpload)}
          style={{ ...S.btn("#1E5FCC"), fontSize:"0.7rem", fontWeight:700 }}>
          {showUpload ? "× Cancel" : "+ Add Meeting"}
        </button>
      </div>

      {/* Upload form */}
      {showUpload && (
        <div style={{ ...S.card, border:"2px solid #1E5FCC", marginBottom:"1rem" }}>
          <div style={{ fontSize:"0.75rem", fontWeight:700, color:"#1E5FCC", marginBottom:"0.75rem" }}>
            ◈ Add Meeting Summary
          </div>
          <div style={{ marginBottom:"0.6rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>Meeting Date</div>
            <input type="date" value={meetingDate} onChange={e=>setMeetingDate(e.target.value)}
              style={{ background:"#FFFFFF", border:`1px solid ${BORDER}`, color:TEXT,
                padding:"0.4rem 0.65rem", borderRadius:6, fontSize:"0.75rem" }} />
          </div>
          <div style={{ marginBottom:"0.75rem" }}>
            <div style={{ fontSize:"0.65rem", color:MUTED, marginBottom:3 }}>
              Paste meeting transcript or notes
            </div>
            <textarea value={transcript} onChange={e=>setTranscript(e.target.value)}
              placeholder="Paste the full meeting transcript here..."
              rows={7}
              style={{ width:"100%", background:"#FFFFFF", border:`1px solid ${BORDER}`,
                color:TEXT, padding:"0.5rem 0.75rem", borderRadius:6, fontSize:"0.75rem",
                resize:"vertical", outline:"none", boxSizing:"border-box", lineHeight:1.7 }}
              onFocus={e=>e.target.style.borderColor="#1E5FCC"}
              onBlur={e=>e.target.style.borderColor=BORDER}
            />
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:3 }}>
              {transcript.length.toLocaleString()} characters
            </div>
          </div>
          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <button onClick={()=>{setShowUpload(false);setTranscript("");}}
              style={S.btn(MUTED)}>Cancel</button>
            <button onClick={generateSummary} disabled={!transcript.trim()||generating}
              style={{ ...S.btn("#1E5FCC"), background:"#1E5FCC", color:"#fff",
                opacity:!transcript.trim()||generating?0.5:1 }}>
              {generating?"◈ Analyzing…":"◈ Generate Summary"}
            </button>
          </div>
        </div>
      )}

      {/* Latest meeting — always fully expanded */}
      {current && <MeetingCard m={current} compact={false} />}

      {/* Previous meetings — collapsible */}
      {archived.length > 0 && (
        <div>
          <button onClick={()=>setShowPrevious(!showPrevious)}
            style={{ width:"100%", display:"flex", justifyContent:"space-between",
              alignItems:"center", padding:"0.6rem 0.85rem",
              background:"#F8FAFC", border:`1px solid ${BORDER}`,
              borderRadius:8, cursor:"pointer", marginBottom: showPrevious?"0.5rem":0,
              fontSize:"0.75rem", fontWeight:700, color:TEXT, userSelect:"none" }}>
            <span>
              {showPrevious?"▾":"▸"}&nbsp; Previous Meetings
              <span style={{ fontSize:"0.65rem", color:MUTED, fontWeight:400, marginLeft:8 }}>
                {archived.length} archived
              </span>
            </span>
            <span style={{ fontSize:"0.65rem", color:MUTED, fontWeight:400 }}>
              {archived.length > 0 && new Date(archived[0].date).toLocaleDateString("en-US",{month:"short",day:"numeric"})}
              {archived.length > 1 && ` – ${new Date(archived[archived.length-1].date).toLocaleDateString("en-US",{month:"short",day:"numeric"})}`}
            </span>
          </button>
          {showPrevious && archived.map(m => (
            <MeetingCard key={m.id} m={m} compact={true} />
          ))}
        </div>
      )}
    </div>
  );
}


function OverviewTab({ weekComp, branchData: propBranchData, onAskAI, onCustomerClick, customers, user }) {
  const [subTab, setSubTab] = useState("weekcomp");
  const branchData = propBranchData || SEED_BRANCH_DATA;

  const weeks = weekComp?.weeks || [];
  // YTD = sum of Q1 + Q2 from branchData (Tifton only, updated by WTD uploads)
  const tifBD = branchData?.branches?.Tifton || {};
  const totalYTD25 = (tifBD.q1_2025||0) + (tifBD.q2_2025||0);
  const totalYTD26 = (tifBD.q1_2026||0) + (tifBD.q2_2026||0);
  const totalGP26  = (tifBD.q1_gp26||0) + (tifBD.q2_gp26||0);
  const totalGP25  = (tifBD.q1_gp25||0) + (tifBD.q2_gp25||0);
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
      {subTab === "depts"    && <DeptView depts={weekComp?.depts || []} weekComp={weekComp} />}
      {subTab === "action"   && <ActionPlanView actionPlan={weekComp?.actionPlan || []} onCustomerClick={onCustomerClick} customers={customers} />}
      {subTab === "comp"     && <CustomerCompView weekComp={weekComp} onCustomerClick={onCustomerClick} customers={customers} user={user} />}
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
function DeptView({ depts, weekComp }) {
  const [aiAnalysis, setAiAnalysis] = useState(() => {
    try { return localStorage.getItem("pulse_dept_ai") || ""; } catch { return ""; }
  });
  const [loading, setLoading] = useState(false);

  const clean = (depts||[]).filter(d =>
    d.sales > 0
    && !String(d.dept||"").startsWith("•")
    && !String(d.dept||"").toLowerCase().includes("departments")
    && !String(d.dept||"").toLowerCase().includes("doing well")
    && !String(d.dept||"").toLowerCase().includes("to focus")
  );
  const sorted = [...clean].sort((a,b) => b.sales - a.sales);
  const maxSales = sorted[0]?.sales || 1;
  const totalSales = sorted.reduce((s,d)=>s+d.sales,0);
  const totalGP    = sorted.reduce((s,d)=>s+d.gp,0);

  async function analyzeWithAI() {
    setLoading(true);
    const deptSummary = sorted.slice(0,12).map(d=>
      `${d.dept}: $${(d.sales/1000).toFixed(1)}K sales, GP ${(d.gpPct*100).toFixed(1)}%, ${d.assessment}`
    ).join('\n');
    const prompt = `You are a sales manager for Tire Distributors of Georgia - Tifton branch.
Analyze this department breakdown and give actionable recommendations.

DEPARTMENT DATA (YTD 2026):
Total Sales: $${(totalSales/1000000).toFixed(2)}M  Total GP: $${(totalGP/1000).toFixed(1)}K  Overall GP%: ${(totalGP/totalSales*100).toFixed(1)}%

${deptSummary}

Give a 5-6 bullet analysis focused on:
1. Top opportunities to grow margin
2. Departments losing GP and why
3. Quick wins for the team this week
4. Any concerns to flag
Keep each bullet to 1-2 sentences, specific and actionable.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":ANTHROPIC_KEY,
          "anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:600,messages:[{role:"user",content:prompt}]})
      });
      const data = await res.json();
      const text = data.content?.[0]?.text || "";
      setAiAnalysis(text);
      try { localStorage.setItem("pulse_dept_ai", text); } catch {}
    } catch(e) { setAiAnalysis(`Analysis failed: ${e.message}`); }
    setLoading(false);
  }

  return (
    <div>
      {/* AI Analysis card */}
      <div style={{ ...S.card, marginBottom:"0.85rem", background:"#F8FAFF",
        border:"1px solid #BFDBFE" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center",
          marginBottom: aiAnalysis?"0.6rem":"0" }}>
          <div style={{ fontSize:"0.72rem", fontWeight:700, color:"#1E5FCC" }}>
            ◈ AI Department Analysis
          </div>
          <button onClick={analyzeWithAI} disabled={loading}
            style={{ fontSize:"0.68rem", fontWeight:700, color:"#fff", background:"#1E5FCC",
              border:"none", borderRadius:6, padding:"0.3rem 0.75rem", cursor:"pointer",
              opacity:loading?0.6:1 }}>
            {loading ? "Analyzing…" : aiAnalysis ? "↺ Refresh" : "Analyze"}
          </button>
        </div>
        {aiAnalysis && (
          <div style={{ fontSize:"0.73rem", color:TEXT, lineHeight:1.75, whiteSpace:"pre-wrap" }}>
            {aiAnalysis}
          </div>
        )}
      </div>

      {/* Summary KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"0.6rem",
        marginBottom:"0.85rem" }}>
        <div style={S.kpi(TEAL)}>
          <div style={S.kpiVal}>${(totalSales/1000000).toFixed(2)}M</div>
          <div style={S.kpiLbl}>YTD Sales</div>
        </div>
        <div style={S.kpi(GREEN)}>
          <div style={S.kpiVal}>${(totalGP/1000).toFixed(1)}K</div>
          <div style={S.kpiLbl}>YTD GP$</div>
        </div>
        <div style={S.kpi(sorted[0]?.gpPct > 0.15 ? GREEN : AMBER)}>
          <div style={S.kpiVal}>{(totalGP/totalSales*100).toFixed(1)}%</div>
          <div style={S.kpiLbl}>Overall GP%</div>
        </div>
      </div>

      {/* Department table */}
      <div style={S.card}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>Department</th>
              <th style={{ ...S.th, width:130 }}>Sales Mix</th>
              <th style={{ ...S.th, textAlign:"right" }}>YTD Sales</th>
              <th style={{ ...S.th, textAlign:"right" }}>GP$</th>
              <th style={{ ...S.th, textAlign:"right" }}>GP%</th>
              <th style={{ ...S.th, textAlign:"right" }}>Lines</th>
              <th style={S.th}>Assessment</th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((d,i) => (
              <tr key={d.dept} style={{ background: i%2===0?"transparent":"#F8FAFC" }}>
                <td style={{ ...S.td, fontWeight:600 }}>{d.dept}</td>
                <td style={S.td}>
                  <div style={{ position:"relative", height:8, background:"#E2E8F0", borderRadius:4 }}>
                    <div style={{ position:"absolute", top:0, left:0, height:8, borderRadius:4,
                      width:`${Math.min(100,(d.sales/maxSales)*100)}%`,
                      background: d.gpPct>0.20?"#059669":d.gpPct>0.12?"#0891B2":"#D97706" }} />
                  </div>
                  <div style={{ fontSize:"0.6rem", color:MUTED, marginTop:2 }}>
                    {(d.sales/totalSales*100).toFixed(1)}% of total
                  </div>
                </td>
                <td style={{ ...S.td, textAlign:"right", fontWeight:600 }}>
                  ${d.sales>=1000000?(d.sales/1000000).toFixed(2)+"M":(d.sales/1000).toFixed(1)+"K"}
                </td>
                <td style={{ ...S.td, textAlign:"right" }}>
                  ${d.gp>=1000?(d.gp/1000).toFixed(1)+"K":d.gp.toFixed(0)}
                </td>
                <td style={{ ...S.td, textAlign:"right",
                  color: d.gpPct>0.20?GREEN:d.gpPct>0.12?TEAL:RED,
                  fontWeight:700 }}>
                  {(d.gpPct*100).toFixed(1)}%
                </td>
                <td style={{ ...S.td, textAlign:"right", color:MUTED }}>
                  {(d.lineItems||0).toLocaleString()}
                </td>
                <td style={{ ...S.td, fontSize:"0.68rem",
                  color: d.gpPct>0.20?"#059669":d.gpPct>0.12?"#0891B2":"#D97706" }}>
                  {d.assessment}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── Action Plan View ──────────────────────────────────────────────────────────

// ── Customer Comparison Tabs ──────────────────────────────────────────────────
function CustomerCompView({ weekComp, onCustomerClick, customers, user }) {
  const [period, setPeriod] = useState("wtd");
  const [search, setSearch] = useState("");
  const [repFilter, setRepFilter] = useState("all");
  const [sortCol, setSortCol] = useState("sales2026");
  const [sortDir, setSortDir] = useState(-1);

  const rep = user?.name || "Admin";
  const isAdmin = user?.role === "Admin";

  // ── Data sources ─────────────────────────────────────────────────────────
  const ytdData    = weekComp?.actionPlan || [];
  const periods    = weekComp?.periods    || {};
  const periodKeys = Object.keys(periods).sort().reverse();

  // ── Fixed tab order: WTD → QTD → YTD ───────────────────────────────────────
  const TAB_ORDER = [
    { key:"wtd", label:"WTD",  color:"#0891B2", bg:"#ECFEFF" },
    { key:"qtd", label:"QTD",  color:"#7C3AED", bg:"#F5F3FF" },  // auto-summed from weekly uploads
    { key:"ytd", label:"YTD",  color:"#1E5FCC", bg:"#EFF6FF" },
  ];

  // ── Current dataset ─────────────────────────────────────────────────────
  // YTD  = full-year actionPlan (Jan 1 – today)
  // QTD  = Q2-only CC report (Apr 1 – today) — needs separate upload
  // WTD  = current week CC report
  const rawData = period === "ytd" ? ytdData :
                  period === "qtd" ? (periods?.qtd?.data || []) :
                  periods[period]?.data || [];
  const noData  = period !== "ytd" && rawData.length === 0;

  // Normalize field names (period data uses s26/s25, ytd uses sales2026/sales2025)
  const normalize = row => ({
    custNum:    row.custNum,
    customer:   row.customer || row.name || "",
    sales2026:  row.sales2026 ?? row.s26 ?? 0,
    sales2025:  row.sales2025 ?? row.s25 ?? 0,
    gp2026:     row.gp2026   ?? row.g26 ?? 0,
    gp2025:     row.gp2025   ?? row.g25 ?? 0,
    change:     row.change   ?? ((row.sales2026??row.s26??0) - (row.sales2025??row.s25??0)),
    changePct:  row.changePct ?? 0,
    salesman:   row.salesman || "House",
    action:     row.action   || "",
    city:       row.city     || "",
  });

  // Reps for filter
  const allReps = ["all", ...new Set(rawData.map(r => r.salesman || "House")
    .filter(Boolean).sort())];

  // Filter & sort
  let rows = rawData.map(normalize).filter(r => {
    if (!isAdmin && r.salesman !== rep) return false;
    if (repFilter !== "all" && r.salesman !== repFilter) return false;
    if (search) {
      const q = search.toLowerCase();
      return r.customer.toLowerCase().includes(q) ||
             String(r.custNum).includes(q);
    }
    return true;
  });

  rows = [...rows].sort((a,b) => {
    const av = a[sortCol] ?? 0, bv = b[sortCol] ?? 0;
    return sortDir * (typeof av==="string" ? av.localeCompare(bv) : bv - av);
  });

  function ThCol({ col, label, right=false }) {
    const active = sortCol === col;
    return (
      <th style={{ ...S.th, textAlign:right?"right":"left", cursor:"pointer",
        color: active?"#1E5FCC":undefined,
        background: active?"#EFF6FF":undefined }}
        onClick={()=>{ if(sortCol===col) setSortDir(d=>d*-1); else { setSortCol(col); setSortDir(-1); } }}>
        {label}{active ? (sortDir<0?" ▾":" ▴") : ""}
      </th>
    );
  }

  const fmt = v => v>=1000000 ? `$${(v/1000000).toFixed(2)}M` :
                   v>=1000    ? `$${(v/1000).toFixed(1)}K`    :
                   `$${v.toFixed(0)}`;

  const totalS26 = rows.reduce((s,r)=>s+r.sales2026,0);
  const totalS25 = rows.reduce((s,r)=>s+r.sales2025,0);
  const totalChg = totalS26 - totalS25;
  const growing  = rows.filter(r=>r.sales2026 > r.sales2025*1.05).length;
  const declining= rows.filter(r=>r.sales2026 < r.sales2025*0.90).length;

  return (
    <div>
      {/* Period sub-tabs — fixed order WTD → QTD → YTD */}
      <div style={{ display:"flex", gap:6, marginBottom:"0.85rem", flexWrap:"wrap",
        alignItems:"center" }}>
        {TAB_ORDER.map(tab => {
          const active  = period === tab.key;
          const hasData = tab.key === "ytd" || (periods[tab.key]?.data?.length > 0);
          return (
            <button key={tab.key} onClick={()=>setPeriod(tab.key)}
              style={{ fontSize:"0.72rem", fontWeight:700, padding:"0.35rem 1rem",
                borderRadius:6, border: active?"none":`1px solid ${tab.color}33`,
                cursor:"pointer", position:"relative",
                background: active ? tab.color : tab.bg,
                color:       active ? "#fff"    : tab.color,
                opacity:     hasData ? 1 : 0.55 }}>
              {tab.label}
              {!hasData && (
                <span style={{ fontSize:"0.55rem", marginLeft:4, opacity:0.8 }}>
                  (no data)
                </span>
              )}
            </button>
          );
        })}
        {/* Any extra uploaded periods (future weeks) */}
        {Object.entries(periods)
          .filter(([k]) => !["wtd","qtd"].includes(k))
          .sort(([a],[b])=>a.localeCompare(b))
          .map(([pk,pv]) => (
            <button key={pk} onClick={()=>setPeriod(pk)}
              style={{ fontSize:"0.7rem", fontWeight:700, padding:"0.35rem 0.9rem",
                borderRadius:6, border:"none", cursor:"pointer",
                background: period===pk?"#059669":"#ECFDF5",
                color:       period===pk?"#fff":"#059669" }}>
              {pv.label || pk.toUpperCase()}
            </button>
          ))}
      </div>

      {/* No-data message for QTD */}
      {noData && (
        <div style={{ ...S.card, background:"#FFFBEB", border:"1px solid #FCD34D",
          padding:"0.85rem 1rem", marginBottom:"0.85rem",
          fontSize:"0.75rem", color:"#92400E" }}>
          <strong>No {period.toUpperCase()} data yet.</strong>
          {period === "qtd" && (
            <span> Run your Customer Comp report in HITS for <strong>Apr 1 – Jun 6, 2026 vs Apr 1 – Jun 6, 2025</strong>, paste it into the CustomerComp tab in your master workbook, and upload. That populates this QTD view with Q2-only numbers.</span>
          )}
        </div>
      )}

      {/* Filters */}
      <div style={{ display:"flex", gap:8, marginBottom:"0.75rem", flexWrap:"wrap", alignItems:"center" }}>
        <input value={search} onChange={e=>setSearch(e.target.value)}
          placeholder="Search customer…"
          style={{ flex:1, minWidth:140, background:"#fff", border:`1px solid ${BORDER}`,
            color:TEXT, padding:"0.35rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none" }}
          onFocus={e=>e.target.style.borderColor="#1E5FCC"}
          onBlur={e=>e.target.style.borderColor=BORDER} />
        {isAdmin && (
          <select value={repFilter} onChange={e=>setRepFilter(e.target.value)}
            style={{ background:"#fff", border:`1px solid ${BORDER}`, color:TEXT,
              padding:"0.35rem 0.65rem", borderRadius:6, fontSize:"0.75rem", outline:"none" }}>
            {allReps.map(r=>(
              <option key={r} value={r}>{r==="all"?"All Reps":r}</option>
            ))}
          </select>
        )}
        <div style={{ fontSize:"0.68rem", color:MUTED, whiteSpace:"nowrap" }}>
          {rows.length} customers
        </div>
      </div>

      {/* KPI summary */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"0.5rem",
        marginBottom:"0.85rem" }}>
        {[
          {label:"2026 Sales", val:fmt(totalS26), color:TEAL},
          {label:"2025 Sales", val:fmt(totalS25), color:MUTED},
          {label:"Change",     val:(totalChg>=0?"+":"")+fmt(Math.abs(totalChg)),
           color:totalChg>=0?GREEN:RED},
          {label:"▲ Growing / ▼ Declining",
           val:`${growing} / ${declining}`, color:AMBER},
        ].map(k=>(
          <div key={k.label} style={S.kpi(k.color)}>
            <div style={{ fontSize:"0.82rem", fontWeight:800, color:k.color }}>{k.val}</div>
            <div style={{ fontSize:"0.62rem", color:MUTED, marginTop:2 }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Table */}
      <div style={{ ...S.card, padding:0, overflow:"hidden" }}>
        <div style={{ overflowX:"auto" }}>
          <table style={{ ...S.table, minWidth:700 }}>
            <thead>
              <tr>
                <ThCol col="customer"  label="Customer" />
                {isAdmin && <ThCol col="salesman" label="Rep" />}
                <ThCol col="sales2026" label="2026 Sales" right />
                <ThCol col="sales2025" label="2025 Sales" right />
                <ThCol col="change"    label="$ Change"   right />
                <ThCol col="changePct" label="% Change"   right />
                {period==="ytd" && <ThCol col="action" label="Status" />}
              </tr>
            </thead>
            <tbody>
              {rows.map((r,i)=>{
                const pct  = r.sales2025>0 ? (r.sales2026-r.sales2025)/r.sales2025 : 0;
                const chgColor = r.change>0?GREEN:r.change<0?RED:MUTED;
                const actionColor = {GROW:GREEN,OK:TEAL,WATCH:AMBER,LOST:RED,NEW:"#7C3AED"}[r.action]||MUTED;
                return (
                  <tr key={r.custNum} style={{ background:i%2===0?"transparent":"#F8FAFC",
                    cursor:"pointer" }}
                    onClick={()=>onCustomerClick && onCustomerClick(r.custNum)}>
                    <td style={{ ...S.td, fontWeight:600 }}>{r.customer}</td>
                    {isAdmin && <td style={{ ...S.td, fontSize:"0.68rem", color:MUTED }}>{r.salesman}</td>}
                    <td style={{ ...S.td, textAlign:"right", fontWeight:700 }}>{fmt(r.sales2026)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:MUTED }}>{fmt(r.sales2025)}</td>
                    <td style={{ ...S.td, textAlign:"right", color:chgColor, fontWeight:600 }}>
                      {r.change>=0?"+":""}{fmt(Math.abs(r.change))}
                    </td>
                    <td style={{ ...S.td, textAlign:"right", color:chgColor }}>
                      {r.sales2025>0?(pct>=0?"+":"")+(pct*100).toFixed(1)+"%":"NEW"}
                    </td>
                    {period==="ytd" && (
                      <td style={{ ...S.td }}>
                        <span style={{ fontSize:"0.65rem", fontWeight:700, color:actionColor,
                          background:actionColor+"22", borderRadius:4,
                          padding:"0.15rem 0.5rem" }}>
                          {r.action}
                        </span>
                      </td>
                    )}
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
  const [acctPeriod, setAcctPeriod] = useState("qtd"); // wtd | qtd | ytd

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


  // ── Pick dataset based on selected period ──────────────────────────────
  const wtdData    = weekComp?.periods?.wtd?.data || [];
  const baseAccounts = acctPeriod === "wtd"
    ? wtdData.map(r => ({
        ...r,
        sales2026: r.sales2026 ?? 0,
        sales2025: r.sales2025 ?? 0,
        action: r.sales2026 > r.sales2025*1.05 ? "GROW" :
                r.sales2026 < r.sales2025*0.90 ? "LOST" : "WATCH",
        topDept: r.topDept || "",
      }))
    : (weekComp?.actionPlan || []);

  const ownAccounts = baseAccounts.filter(a =>
    (a.salesman||"House").toLowerCase() === repName.toLowerCase());

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

      {/* ── WTD / QTD / YTD period selector ── */}
      <div style={{ display:"flex", gap:6, marginBottom:"0.75rem", flexWrap:"wrap" }}>
        {[
          { key:"wtd", label:"WTD",  color:"#0891B2", bg:"#ECFEFF",
            hasData: (weekComp?.periods?.wtd?.data||[]).filter(r=>
              (r.salesman||"House").toLowerCase()===repName.toLowerCase()).length > 0 },
          { key:"qtd", label:"QTD",  color:"#7C3AED", bg:"#F5F3FF", hasData:true },
          { key:"ytd", label:"YTD",  color:"#1E5FCC", bg:"#EFF6FF", hasData:true },
        ].map(tab => (
          <button key={tab.key} onClick={()=>setAcctPeriod(tab.key)}
            style={{ fontSize:"0.72rem", fontWeight:700, padding:"0.35rem 1rem",
              borderRadius:6, cursor:"pointer",
              border: acctPeriod===tab.key ? "none" : `1px solid ${tab.color}33`,
              background: acctPeriod===tab.key ? tab.color : tab.bg,
              color:       acctPeriod===tab.key ? "#fff"    : tab.color,
              opacity:     tab.hasData ? 1 : 0.5 }}>
            {tab.label}
            {!tab.hasData && <span style={{fontSize:"0.6rem",marginLeft:4}}>(no data)</span>}
          </button>
        ))}
        <span style={{ fontSize:"0.65rem", color:MUTED, alignSelf:"center", marginLeft:4 }}>
          {acctPeriod==="wtd" ? "Current week" :
           acctPeriod==="qtd" ? "Q2 Apr 1–Jun 6" : "Full year Jan 1–Jun 6"}
        </span>
      </div>

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

