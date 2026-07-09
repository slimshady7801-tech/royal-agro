import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  Plus, Trash2, Users, ShoppingCart, Wallet, Search, X, Download,
  ArrowLeft, Building2, Receipt, Minus, Upload, ChevronDown, ChevronUp,
  Settings as SettingsIcon, LayoutDashboard, Boxes, AlertTriangle, Pencil,
  MapPin, User, RotateCcw, Calculator, ClipboardList, TrendingUp, Share2, Calendar,
} from "lucide-react";

// ----------------------------- Catalog ---------------------------------
const DISCOUNT = 0.35;

const packSize = (pk) => {
  const s = String(pk || "").toLowerCase();
  const xm = s.match(/x\s*(\d+)/);
  const mul = xm ? parseInt(xm[1]) : 1;
  const m = s.match(/([\d.]+)\s*(kg|ltr|ml|gm|nos)/);
  if (!m) return null;
  let v = parseFloat(m[1]);
  const u = m[2];
  if (u === "ml" || u === "gm") v /= 1000;
  if (u === "nos") return null;
  return +(v * mul).toFixed(4);
};

const RAW = [
  ["Cymoxanil + Mancozeb 72 WP","Fungicide","600 GM",1690],
  ["Cymoxanil + Mancozeb 72 WP","Fungicide","25 KG",66000],
  ["Finer 50 SC (Azoxystrobin + Tebuconazole)","Fungicide","120 ML",675],
  ["Kanglixin 18.7 (Pyraclostrobin + Dimethomorph)","Fungicide","200 GM",550],
  ["Metalaxyl + Mancozeb 72 WP","Fungicide","250 GM",715],
  ["Metalaxyl + Mancozeb 72 WP","Fungicide","1 KG",2575],
  ["Metalaxyl + Mancozeb 72 WP","Fungicide","25 KG",61500],
  ["Maidian 56 SC (Azoxystrobin + Chlorothalonil)","Fungicide","400 ML",1200],
  ["Pyrazole 45 WP (Pyraclostrobin + Tebuconazole)","Fungicide","100 GM",385],
  ["Puching Star 47 WP (Kasugamycin + Copper Oxychloride)","Fungicide","250 GM",1050],
  ["Puching Star 47 WP (Kasugamycin + Copper Oxychloride)","Fungicide","25 KG",92500],
  ["Recado Super 32 SC (Azoxystrobin + Propiconazole)","Fungicide","100 ML",560],
  ["Sulphur 80 WDG","Fungicide","1 KG",1030],
  ["Sulphur 80 WDG","Fungicide","2 KG",1960],
  ["Sulphur 80 WDG","Fungicide","25 KG",23000],
  ["Thiophenate Methyl 70 WP","Fungicide","100 GM",350],
  ["Thiophenate Methyl 70 WP","Fungicide","400 GM",1100],
  ["Thiophenate Methyl 70 WP","Fungicide","25 KG",62000],
  ["Triadimefon 25 WP","Fungicide","100 GM",290],
  ["Acetochlor 50 EC","Herbicide","100 ML",265],
  ["Acetochlor 50 EC","Herbicide","1 LTR",1380],
  ["Atrazine 38 SC","Herbicide","500 ML",650],
  ["Butachlor 60 EC","Herbicide","800 ML",1200],
  ["Full Control 50 WP (Mesotrione + Atrazine)","Herbicide","500 GM",950],
  ["Full Control 50 WP (Mesotrione + Atrazine)","Herbicide","1 KG",1800],
  ["Gazonner 15 EC (Quizalofop)","Herbicide","500 ML",1400],
  ["Gengwei 55 SC","Herbicide","500 ML",860],
  ["Gengwei 55 SC","Herbicide","1 LTR",1595],
  ["Gengwei Super 33.5 SC","Herbicide","750 ML",1530],
  ["Glyphosate 48 SL","Herbicide","1 LTR",1285],
  ["Click 72.4 SE (Acetochlor + Atrazine)","Herbicide","750 ML",1630],
  ["Matrix 30 WP (Bispyribac Sodium + Bensulfuron)","Herbicide","100 GM",870],
  ["Paraquat 20 SC","Herbicide","1 LTR",1440],
  ["Pendimethalin 33 SC","Herbicide","1 LTR",1700],
  ["Recall 42 EC (Pendimethalin + Acetochlor)","Herbicide","800 ML",1300],
  ["Ridge 30 SC (Topramezone)","Herbicide","35 ML",800],
  ["Specter 20 EC (Cyhalofop-Butyl + Metamifop)","Herbicide","150 ML",770],
  ["Specter 20 EC (Cyhalofop-Butyl + Metamifop)","Herbicide","800 ML",3500],
  ["Terminator 10 WP (Pyrazosulfuron-ethyl)","Herbicide","100 GM",350],
  ["Winsta 30 WP (Bispyribac Sodium + Bensulfuron)","Herbicide","100 GM",800],
  ["Yinongle 20 WP (Bensulfuron + Acetochlor)","Herbicide","300 GM",725],
  ["Orcus 75 WDG (Halosulfuron)","Herbicide","20 GM",700],
  ["Abamectin 1.8 EC","Insecticide","400 ML",600],
  ["Acephate 75 SP","Insecticide","250 GM",730],
  ["Acetamiprid 20 SP","Insecticide","250 GM",490],
  ["Bifenthrin 10 EC","Insecticide","500 ML",1040],
  ["Bifenthrin 10 EC","Insecticide","1 LTR",1940],
  ["Bravo 5 WDG (Emamectin)","Insecticide","150 GM",615],
  ["Buprofezin 25 WP","Insecticide","900 GM",1325],
  ["Buprofezin 25 WP","Insecticide","25 KG",33800],
  ["Ceedo 20 SC (Clothianidin)","Insecticide","200 ML",430],
  ["Ceedo 20 SC (Clothianidin)","Insecticide","1 LTR",1500],
  ["Chlorpyrifos 40 EC","Insecticide","250 ML",595],
  ["Chlorpyrifos 40 EC","Insecticide","1 LTR",1925],
  ["Cypermethrin 10 EC","Insecticide","250 ML",505],
  ["Cypermethrin 10 EC","Insecticide","800 ML",1550],
  ["Cypermethrin + Profenofos 440 EC","Insecticide","500 ML",1450],
  ["Cypermethrin + Profenofos 440 EC","Insecticide","1 LTR",2750],
  ["Dimethoate 40 EC","Insecticide","400 ML",950],
  ["Dino 20 SG (Dinotefuran)","Insecticide","100 GM",385],
  ["Dumei 50 WDG (Flonicamid)","Insecticide","120 GM",1400],
  ["Emamectin 1.9 EC","Insecticide","400 ML",780],
  ["Emamectin 1.9 EC","Insecticide","1 LTR",1750],
  ["Fipronil 5 SC","Insecticide","480 ML",975],
  ["Imidacloprid 25 WP","Insecticide","200 GM",460],
  ["Instant Super 50 WDG (Nitenpyram)","Insecticide","80 GM",575],
  ["Lambda Cyhalothrin 2.5 EC","Insecticide","250 ML",400],
  ["Lambda Cyhalothrin 2.5 EC","Insecticide","1 LTR",1150],
  ["Lancer 10 EC (Lambda-Cyhalothrin)","Insecticide","80 ML",375],
  ["Lancer 10 EC (Lambda-Cyhalothrin)","Insecticide","240 ML",820],
  ["Leera 80 WDG (Diafenthiuron)","Insecticide","150 GM",1275],
  ["Lufenuron 5 EC","Insecticide","400 ML",660],
  ["Lufenuron 5 EC","Insecticide","1 LTR",1430],
  ["Pattern 10 WDG (Chlorfenapyr)","Insecticide","350 GM",710],
  ["Pyriproxifen 10.8 EC","Insecticide","500 ML",925],
  ["Shidding 21 EC (Lambda + Triazophos)","Insecticide","800 ML",1850],
  ["Sudao 11.6 SC (Emamectin + Chlorantraniliprole)","Insecticide","100 ML",560],
  ["Thiamethoxam 25 WP","Insecticide","50 GM",200],
  ["Valor 12 SC (Abamectin + Thiamethoxam)","Insecticide","200 ML",500],
  ["Valor 12 SC (Abamectin + Thiamethoxam)","Insecticide","400 ML",825],
  ["Veyong Jinteng 60 WG (Pymetrozine + Dinotefuron)","Insecticide","100 GM",850],
  ["Woolmer 3 SC (Lufenuron + Emamectin)","Insecticide","400 ML",585],
  ["Woolmer 3 SC (Lufenuron + Emamectin)","Insecticide","1 LTR",1230],
  ["Xuanli Jing 45 WDG (Emamectin + Lufenuron)","Insecticide","40 GM",490],
  ["Comfly 26 FS (Fludioxonil + Metalaxyl-M + Clothianidin)","Seed","80 ML",660],
  ["Bio Organo Phosphate (BOP)","Nutrient","50 KG",3400],
  ["Maxima 40+7% (Humic Acid Granules)","Nutrient","8 KG",2000],
  ["Maxima 40+7% (Humic Acid Granules)","Nutrient","20 KG",4800],
  ["Roxy (Organic Matter)","Nutrient","25 KG",2370],
  ["Torch 10% (Humic Acid Liquid + Fulvic)","Nutrient","4 LTR",840],
  ["Torch 10% (Humic Acid Liquid + Fulvic)","Nutrient","20 LTR",2875],
  ["Torch 10% (Humic Acid Liquid + Fulvic)","Nutrient","200 LTR",26700],
  ["Addox 5% (Boron Liquid)","Nutrient","500 ML",540],
  ["Addox 5% (Boron Liquid)","Nutrient","20 LTR",17000],
  ["Addox 5% (Boron Liquid)","Nutrient","200 LTR",156000],
  ["Adena 20% (Liquid Nitrogen)","Nutrient","1 LTR",475],
  ["Adena 20% (Liquid Nitrogen)","Nutrient","4 LTR",1640],
  ["Adena 20% (Liquid Nitrogen)","Nutrient","20 LTR",5800],
  ["Adena 20% (Liquid Nitrogen)","Nutrient","200 LTR",43000],
  ["Adforce 5% (Zinc Chelated)","Nutrient","2 KG",1700],
  ["Adpotash 30% (Potash Liquid) Foliar","Nutrient","1 LTR",1080],
  ["Adpotash 30% (Potash Liquid) Fertigation","Nutrient","3 LTR",2200],
  ["Adpotash 30% (Potash Liquid) Fertigation","Nutrient","20 LTR",10750],
  ["Adpotash 30% (Potash Liquid) Fertigation","Nutrient","200 LTR",92000],
  ["Brito 10% G (Zinc 7% + Ferrous 3%)","Nutrient","5 KG",1400],
  ["Hamada 20% (Liquid Phosphorus)","Nutrient","3 LTR",2500],
  ["Hamada 20% (Liquid Phosphorus)","Nutrient","5 LTR",3700],
  ["Hamada 20% (Liquid Phosphorus)","Nutrient","20 LTR",12600],
  ["Hamada 20% (Liquid Phosphorus)","Nutrient","200 LTR",110000],
  ["Mega Grow-TB 10% (Gibberellic Acid)","Nutrient","10 GM x 10",180],
  ["Prima (Multi Micronutrients Zn 6% + Fe 2% + Mn 2%)","Nutrient","500 ML",510],
  ["Prima (Multi Micronutrients Zn 6% + Fe 2% + Mn 2%)","Nutrient","20 LTR",14000],
  ["Prima (Multi Micronutrients Zn 6% + Fe 2% + Mn 2%)","Nutrient","200 LTR",124000],
  ["Vigro Plus (NPK 10:10:10)","Nutrient","500 ML",465],
  ["Vigro Plus (NPK 10:10:10)","Nutrient","20 LTR",12100],
  ["Vigro Plus (NPK 10:10:10)","Nutrient","200 LTR",105000],
  ["Zenebien 10% (Zinc Liquid)","Nutrient","3 LTR",1140],
  ["Zenebien 10% (Zinc Liquid)","Nutrient","20 LTR",6250],
  ["Zenebien 10% (Zinc Liquid)","Nutrient","200 LTR",60000],
  ["Zenith 21% (Zinc Sulphate)","Nutrient","5 KG",3100],
  ["Aditive (Bio Stimulant)","Nutrient","500 ML",600],
  ["Aditive (Bio Stimulant)","Nutrient","1 LTR",1080],
  ["Aditive (Bio Stimulant)","Nutrient","3 LTR",3000],
  ["Adoratin 11.6 SC (CTPR + Emamectin)","Adler","100 ML",600],
  ["Adoxachlor 20 SC (CTPR + Indoxacarb)","Adler","100 ML",1190],
  ["Adoxy 35 SC (Oxine Copper + Tetramycin)","Adler","200 ML",1850],
  ["Dinomid 30 SC (Flonicamid + Dinotefuron)","Adler","200 ML",1430],
  ["Dinoprole 30 SC (CTPR + Dinotefuron)","Adler","100 ML",770],
  ["Jago (CTPR + Methoxyfenozide)","Adler","50 ML",770],
  ["Tolfen 15 SC (Tolfenpyrad + Spinosad)","Adler","200 ML",2000],
];

const CATALOG = RAW.map(([name, cat, pack, rate], i) => ({
  id: "prod-" + i, name, cat, pack, rate, dealer: Math.round(rate * (1 - DISCOUNT)),
}));
const CATS = ["All", "Fungicide", "Herbicide", "Insecticide", "Seed", "Nutrient", "Adler"];
const PACK_SIZE = {};
CATALOG.forEach((p) => { PACK_SIZE[p.id] = packSize(p.pack); });
const CAT_BY_ID = {};
CATALOG.forEach((p) => { CAT_BY_ID[p.id] = p; });

// Seed data extracted from ROYAL cost sheet (Apr 2026) and stock report (06-Jul-2026)
// COST_SEED[id] = [costPerLtrKg, packingCost]; cost-with-packing = cpu*packSize + packing
const COST_SEED = {"prod-0":[1410,65],"prod-1":[1410,250],"prod-2":[2225,85],"prod-3":[1400,55],"prod-4":[1425,60],"prod-5":[1425,75],"prod-6":[1425,250],"prod-7":[1350,90],"prod-8":[1450,55],"prod-9":[2200,70],"prod-10":[2200,250],"prod-11":[1975,85],"prod-12":[430,75],"prod-13":[430,95],"prod-14":[430,250],"prod-15":[1300,55],"prod-16":[1300,65],"prod-17":[1300,250],"prod-18":[850,55],"prod-19":[775,85],"prod-20":[775,100],"prod-21":[550,90],"prod-22":[750,100],"prod-23":[975,65],"prod-24":[975,75],"prod-25":[1425,90],"prod-26":[835,90],"prod-27":[835,100],"prod-28":[920,100],"prod-29":[600,100],"prod-30":[990,100],"prod-32":[735,100],"prod-33":[950,100],"prod-34":[935,100],"prod-35":[405,30],"prod-39":[3025,150],"prod-40":[1000,100],"prod-42":[750,90],"prod-43":[1400,55],"prod-44":[825,60],"prod-45":[1115,90],"prod-46":[1115,100],"prod-47":[1750,55],"prod-48":[775,75],"prod-49":[775,250],"prod-50":[735,90],"prod-51":[735,100],"prod-52":[1115,90],"prod-53":[1115,100],"prod-54":[790,90],"prod-55":[790,100],"prod-56":[1550,90],"prod-57":[1550,100],"prod-58":[1050,90],"prod-59":[1450,55],"prod-60":[6250,60],"prod-61":[1000,90],"prod-62":[1000,100],"prod-63":[950,90],"prod-64":[1050,60],"prod-65":[3100,55],"prod-66":[580,90],"prod-67":[580,100],"prod-68":[1300,85],"prod-69":[1300,90],"prod-70":[4000,55],"prod-71":[825,90],"prod-72":[825,100],"prod-73":[395,10],"prod-74":[925,90],"prod-75":[1075,100],"prod-76":[2350,110],"prod-77":[925,55],"prod-78":[925,90],"prod-79":[925,90],"prod-80":[3895,55],"prod-81":[615,90],"prod-82":[615,100],"prod-83":[4600,55],"prod-84":[3150,95],"prod-85":[37,250],"prod-86":[120,125],"prod-87":[120,200],"prod-88":[41,250],"prod-89":[30,110],"prod-90":[30,1250],"prod-91":[30,4250],"prod-92":[380,90],"prod-93":[380,1250],"prod-94":[380,4250],"prod-95":[115,115],"prod-96":[115,250],"prod-97":[115,1250],"prod-98":[115,4250],"prod-99":[380,95],"prod-100":[400,100],"prod-101":[225,0],"prod-102":[225,1250],"prod-103":[225,4250],"prod-104":[145,0],"prod-105":[280,240],"prod-106":[280,525],"prod-107":[280,1250],"prod-108":[280,4250],"prod-109":[100,55],"prod-110":[300,90],"prod-111":[300,1250],"prod-112":[300,4250],"prod-113":[250,90],"prod-114":[250,1250],"prod-115":[250,4250],"prod-116":[145,0],"prod-117":[145,1250],"prod-118":[145,4250],"prod-119":[300,0],"prod-120":[395,90],"prod-121":[395,100],"prod-122":[395,0],"prod-123":[2350,110],"prod-124":[5400,95],"prod-126":[405,30],"prod-127":[405,30],"prod-129":[4500,95]};
const STOCK_SEED = {"prod-0":48,"prod-2":215,"prod-3":128,"prod-4":50,"prod-5":35,"prod-6":0,"prod-7":131,"prod-8":36,"prod-9":259,"prod-10":0,"prod-11":172,"prod-12":1747,"prod-13":106,"prod-14":14,"prod-16":2,"prod-18":467,"prod-19":1103,"prod-20":595,"prod-21":282,"prod-22":226,"prod-23":0,"prod-24":157,"prod-25":62,"prod-26":632,"prod-27":245,"prod-29":489,"prod-30":211,"prod-31":1106,"prod-32":253,"prod-33":94,"prod-34":194,"prod-35":458,"prod-36":698,"prod-37":129,"prod-38":1761,"prod-39":626,"prod-41":260,"prod-42":839,"prod-43":76,"prod-44":706,"prod-45":325,"prod-46":488,"prod-47":826,"prod-48":79,"prod-49":0,"prod-50":1034,"prod-51":400,"prod-52":781,"prod-53":267,"prod-56":0,"prod-57":0,"prod-58":306,"prod-59":1342,"prod-60":50,"prod-61":202,"prod-62":302,"prod-63":437,"prod-64":0,"prod-65":173,"prod-66":388,"prod-67":450,"prod-70":68,"prod-71":683,"prod-72":661,"prod-73":80,"prod-74":435,"prod-75":0,"prod-76":1074,"prod-77":829,"prod-80":214,"prod-81":1122,"prod-82":1058,"prod-83":216,"prod-86":39,"prod-88":64,"prod-90":128,"prod-91":0,"prod-92":-32,"prod-93":0,"prod-97":0,"prod-104":-21,"prod-110":0,"prod-116":178,"prod-120":60,"prod-121":153,"prod-124":0,"prod-126":0,"prod-127":0,"prod-128":0,"prod-129":0};
const SEED_DATE = "2026-07-06";

// ------------------------- Matching helpers -----------------------------
const norm = (s) =>
  String(s == null ? "" : s).toLowerCase().replace(/\(.*?\)/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/([0-9])([a-z])/g, "$1 $2").replace(/([a-z])([0-9])/g, "$1 $2")
    .trim().replace(/\s+/g, " ");
const parseNum = (v) => {
  const n = parseFloat(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
};
const CAT_KEYS = CATALOG.map((p) => {
  const base = norm(p.name);
  return { id: p.id, base, full: (base + " " + norm(p.pack)).trim() };
});
function matchCatId(name, pack) {
  const rn = norm(name);
  const rfull = (rn + (pack ? " " + norm(pack) : "")).trim();
  let hit = CAT_KEYS.find((k) => k.full === rfull) || CAT_KEYS.find((k) => k.full === rn);
  if (hit) return hit.id;
  const baseM = CAT_KEYS.filter((k) => k.base === rn || k.base === rfull);
  if (baseM.length === 1) return baseM[0].id;
  const tokens = new Set(rfull.split(" ").filter(Boolean));
  let best = null, bestScore = 0;
  for (const k of CAT_KEYS) {
    const kt = k.full.split(" ");
    const inter = kt.filter((t) => tokens.has(t)).length;
    const uni = new Set([...tokens, ...kt]).size;
    const sc = uni ? inter / uni : 0;
    if (sc > bestScore) { bestScore = sc; best = k; }
  }
  return best && bestScore >= 0.6 ? best.id : null;
}

// Parse a pasted stock report. When copied from a PDF, columns often glue together
// (e.g. "1,457658" = balance 1,457 + next column 658). We reconstruct the true closing
// balance from the accounting identity (Balance = Opening + Purchase + Production + Transfer + Sale)
// and reject any row that doesn't reconcile, so a bad paste never produces wrong stock.
function parseStockText(text) {
  const PACK = /(\d+(?:\.\d+)?)(kg|gm|ml|ltr|nos)\b/ig;
  const SKIP = /MULTAN|Product Name|Page \d|Pack Bal|^Adj\.|FERTILIZER|FUNJICIDE|FUNGICIDE|HERBICIDE|INSECTICIDE|GRANULE|HOUSE HOLD|MICRO NUTR|OTHERS|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday/i;
  function balOf(toks, nums) {
    if (nums.length >= 6) { const b = nums[0]; return b === nums[1] + nums[2] + nums[3] + nums[4] + nums[5] ? b : null; }
    if (nums.length === 5) {
      const S = nums[1] + nums[2] + nums[3] + nums[4];
      const g = String(toks[0]).replace(/,/g, "").replace(/^-/, "");
      for (let k = 1; k < g.length; k++) { const b = parseInt(g.slice(0, k), 10), o = parseInt(g.slice(k), 10); if (!isNaN(b) && !isNaN(o) && o >= 0 && b >= 0 && b === o + S) return b; }
      if (nums[0] === nums[1] + nums[2] + nums[3] + nums[4]) return nums[0];
      return null;
    }
    if (nums.length === 4) return nums[0] === nums[1] + nums[2] + nums[3] ? nums[0] : null;
    return null;
  }
  const map = {}, unmatched = []; let lines = 0, skipped = 0;
  for (const raw of String(text).split(/\r?\n/)) {
    if (SKIP.test(raw)) continue;
    PACK.lastIndex = 0; let m, last = null; while ((m = PACK.exec(raw))) last = m;
    if (!last) continue;
    const name = raw.slice(0, last.index).replace(/^\s*(-?[\d,]+\s+)+/, "").trim();
    if (!name) continue;
    const cluster = raw.slice(last.index + last[0].length);
    const toks = cluster.match(/-?[\d,]+/g); if (!toks || !toks.length) continue;
    const nums = toks.map((t) => parseInt(t.replace(/,/g, ""), 10));
    if (nums.some(isNaN)) continue;
    const bal = balOf(toks, nums);
    lines++;
    if (bal == null) { skipped++; continue; }
    const id = matchCatId(name, last[0]);
    if (id) map[id] = (map[id] || 0) + bal;
    else unmatched.push(name + " " + last[0]);
  }
  const rounded = {};
  Object.keys(map).forEach((k) => { rounded[k] = Math.round(map[k]); });
  return { map: rounded, matched: Object.keys(rounded).length, unmatched, lines, skipped };
}

// Column-detecting importer for CSV/XLSX stock files.
function buildImport(rows) {
  if (!rows || !rows.length) return { map: {}, matched: 0, unmatched: [] };
  const header = (rows[0] || []).map((c) => String(c).toLowerCase());
  const findCol = (re) => header.findIndex((h) => re.test(h));
  let nameCol = findCol(/product|name|item|descrip|particular|article|goods/);
  const packCol = findCol(/pack|packing|size|uom|volume|weight/);
  let valCol = findCol(/clos|balance|\bbal\b|stock|qty|quantity|avail|on.?hand|inhand|remain/);
  if (nameCol < 0) nameCol = 0;
  const colCount = Math.max(...rows.map((r) => r.length));
  if (valCol < 0) {
    let best = -1, bestC = -1;
    for (let c = 0; c < colCount; c++) {
      if (c === nameCol || c === packCol) continue;
      let cnt = 0;
      for (let i = 1; i < rows.length; i++) if (parseNum((rows[i] || [])[c]) != null) cnt++;
      if (cnt > bestC) { bestC = cnt; best = c; }
    }
    valCol = best >= 0 ? best : 1;
  }
  const start = parseNum((rows[0] || [])[valCol]) == null ? 1 : 0;
  const map = {}, unmatched = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i] || [];
    const nm = String(r[nameCol] == null ? "" : r[nameCol]).trim();
    if (!nm) continue;
    const rawVal = r[valCol];
    if (rawVal === "" || rawVal == null) continue;
    const val = parseNum(rawVal);
    if (val == null) continue;
    const pk = packCol >= 0 ? String(r[packCol] == null ? "" : r[packCol]).trim() : "";
    const id = matchCatId(nm, pk);
    if (id) map[id] = val; else unmatched.push(nm + (pk ? " " + pk : ""));
  }
  return { map, matched: Object.keys(map).length, unmatched };
}
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read"));
    if (/\.csv$/i.test(file.name)) {
      reader.onload = () => { try { resolve(Papa.parse(String(reader.result), { skipEmptyLines: true }).data); } catch (e) { reject(e); } };
      reader.readAsText(file);
    } else {
      reader.onload = () => {
        try {
          const wb = XLSX.read(new Uint8Array(reader.result), { type: "array" });
          const sh = wb.Sheets[wb.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" }));
        } catch (e) { reject(e); }
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

// ----------------------------- Utils ------------------------------------
const fmt = (n) => "Rs " + Math.round(n).toLocaleString("en-US");
const plain = (n) => Math.round(n).toLocaleString("en-US");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => new Date().toISOString().slice(0, 10);
const prettyDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  const mm = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+m - 1];
  return (+d) + " " + mm + " " + y;
};
// Positive = iso is in the past (overdue by N days), 0 = today, negative = upcoming
const daysBetween = (iso) => Math.round((new Date(today() + "T00:00:00") - new Date(iso + "T00:00:00")) / 86400000);
const followUpStatus = (d) => {
  if (!d.followUp) return null;
  const diff = daysBetween(d.followUp);
  if (diff > 0) return { label: diff === 1 ? "1 day overdue" : diff + " days overdue", tone: "red" };
  if (diff === 0) return { label: "Follow up today", tone: "amber" };
  return { label: "Follow up " + prettyDate(d.followUp), tone: "slate" };
};
const KEY = "royal-agro-ledger-v2";
const hasStore = typeof window !== "undefined" && window.storage;
async function dbLoad() { if (!hasStore) return null; try { const r = await window.storage.get(KEY); return r && r.value ? JSON.parse(r.value) : null; } catch { return null; } }
async function dbSave(d) { if (!hasStore) return; try { await window.storage.set(KEY, JSON.stringify(d)); } catch { /* ignore */ } }
function downloadFile(filename, text, mime) {
  const blob = new Blob([text], { type: mime || "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
const csvCell = (v) => { const s = String(v == null ? "" : v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };

const SEED_META = { updated: SEED_DATE, matched: Object.keys(STOCK_SEED).length, unmatched: [], seeded: true };
const DEFAULTS = {
  dealers: [], txns: [], stock: STOCK_SEED, costing: COST_SEED, names: {},
  stockMeta: SEED_META, settings: { lowStock: 10 },
};
let flashTimeoutId = null; // module-scope: survives re-renders without needing useRef

// Products grouped for the costing editor (cost/Ltr·Kg is shared across a product's packs)
const COST_GROUPS = (() => {
  const g = {};
  CATALOG.forEach((p) => { (g[p.name] = g[p.name] || []).push(p); });
  return Object.keys(g).map((name) => ({ name, cat: g[name][0].cat, skus: g[name] }));
})();

// =============================== App ====================================
export default function App() {
  const [data, setData] = useState(DEFAULTS);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("glance");
  const [detailId, setDetailId] = useState(null);
  const [tab, setTab] = useState("ledger");
  const [msg, setMsg] = useState("");
  const [msgUndo, setMsgUndo] = useState(null); // undo callback for the current toast, or null
  const [shareModal, setShareModal] = useState(null); // { title, text } | null
  const [showSettings, setShowSettings] = useState(false);
  const [showCosting, setShowCosting] = useState(false);
  const [costSearch, setCostSearch] = useState("");

  const [dealerSearch, setDealerSearch] = useState("");
  const [fArea, setFArea] = useState("All");
  const [fRep, setFRep] = useState("All");

  const [formOpen, setFormOpen] = useState(false);
  const [editId, setEditId] = useState(null);
  const [f, setF] = useState({ name: "", phone: "", area: "", rep: "", remarks: "", followUp: "" });

  const [cart, setCart] = useState([]);
  const [oDate, setODate] = useState(today());
  const [oNote, setONote] = useState("");
  const [pSearch, setPSearch] = useState("");
  const [pCat, setPCat] = useState("All");

  const [payAmt, setPayAmt] = useState("");
  const [payDate, setPayDate] = useState(today());
  const [payMethod, setPayMethod] = useState("Cash");
  const [payNote, setPayNote] = useState("");

  const [openRows, setOpenRows] = useState({});
  const [stockSearch, setStockSearch] = useState("");
  const [showUnmatched, setShowUnmatched] = useState(false);
  const [showPaste, setShowPaste] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [confirmAsk, setConfirmAsk] = useState(null);   // { msg, action }
  const [editOrderId, setEditOrderId] = useState(null); // order being edited
  const [payEdit, setPayEdit] = useState(null);         // payment being edited
  const [repView, setRepView] = useState(null);         // salesperson name
  const [todayDate, setTodayDate] = useState(today());
  const [prodView, setProdView] = useState(null);       // catId being viewed/edited
  const [ocDealerId, setOcDealerId] = useState("");     // "" | dealerId | "__new__"
  const [ocNew, setOcNew] = useState({ name: "", area: "", phone: "", rep: "" });
  const [ocPayAmt, setOcPayAmt] = useState("");
  const [ocPayMethod, setOcPayMethod] = useState("Cash");

  useEffect(() => { (async () => {
    const d = await dbLoad();
    if (d) {
      const merged = { ...DEFAULTS, ...d };
      merged.settings = { ...DEFAULTS.settings, ...(d.settings || {}) };
      merged.costing = { ...COST_SEED, ...(d.costing || {}) };
      merged.names = { ...(d.names || {}) };
      if (!d.stockMeta) { merged.stock = STOCK_SEED; merged.stockMeta = SEED_META; }
      setData(merged);
    }
    setLoaded(true);
  })(); }, []);

  const persist = (next) => {
    setData((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      dbSave(resolved);
      return resolved;
    });
  };
  const flash = (t, undo) => {
    setMsg(t); setMsgUndo(() => undo || null);
    if (flashTimeoutId) clearTimeout(flashTimeoutId);
    flashTimeoutId = setTimeout(() => { setMsg(""); setMsgUndo(null); }, undo ? 5500 : 2800);
  };
  const threshold = data.settings?.lowStock ?? 10;

  // hidden cost-with-packing for a SKU (used only for flagging + costing editor)
  const costWith = (id) => {
    const c = data.costing[id];
    if (!c) return null;
    const cpu = c[0], ps = PACK_SIZE[id];
    if (cpu == null || ps == null) return null;
    return Math.round(cpu * ps + (c[1] || 0));
  };
  // profit on an order = sum over costed items of (rate - cost-with-packing) * qty
  const orderProfit = (t) => t.items.reduce((s, i) => { const c = costWith(i.catId); return c != null ? s + (i.rate - c) * i.qty : s; }, 0);

  // display name for a product (custom override or catalog default)
  const pname = (id) => (data.names && data.names[id]) || (CAT_BY_ID[id] ? CAT_BY_ID[id].name : "");
  const setName = (id, val) => { const n = { ...(data.names || {}) }; if (val.trim()) n[id] = val; else delete n[id]; persist({ ...data, names: n }); };
  const setStockVal = (id, val) => { const s = { ...data.stock }; if (val === "") delete s[id]; else s[id] = Math.max(0, Math.round(parseNum(val) || 0)); persist({ ...data, stock: s }); };
  const setCostField = (id, idx, val) => {
    const cur = data.costing[id] ? [data.costing[id][0], data.costing[id][1]] : [null, 0];
    cur[idx] = val === "" ? (idx === 0 ? null : 0) : parseNum(val);
    persist({ ...data, costing: { ...data.costing, [id]: cur } });
  };
  const setGroupCpu = (id, val) => {
    const nm = CAT_BY_ID[id] ? CAT_BY_ID[id].name : null;
    if (!nm) return;
    const v = val === "" ? null : parseNum(val);
    const next = { ...data.costing };
    CATALOG.forEach((p) => { if (p.name === nm) { const cur = next[p.id] || [null, 0]; next[p.id] = [v, cur[1] || 0]; } });
    persist({ ...data, costing: next });
  };
  const nextOrderNo = () => {
    const ns = data.txns.filter((t) => t.type === "order" && t.orderNo).map((t) => parseInt(String(t.orderNo).replace(/\D/g, "")) || 0);
    return (ns.length ? Math.max(...ns) : 1000) + 1;
  };
  const resetComposer = () => { setCart([]); setOcDealerId(""); setOcNew({ name: "", area: "", phone: "", rep: "" }); setOcPayAmt(""); setONote(""); setODate(today()); setPSearch(""); setEditOrderId(null); };
  const saveComposerOrder = () => {
    if (!cart.length) { flash("Add at least one product to the order"); return; }
    let dealerId = ocDealerId;
    let next = { ...data };
    if (ocDealerId === "__new__") {
      const name = ocNew.name.trim();
      if (!name) { flash("Enter the new dealer's name"); return; }
      dealerId = uid();
      next = { ...next, dealers: [...next.dealers, { id: dealerId, name, phone: ocNew.phone.trim(), area: ocNew.area.trim(), rep: ocNew.rep.trim(), remarks: "", created: Date.now() }] };
    } else if (!dealerId) { flash("Choose a dealer for this order"); return; }
    const items = cart.map((c) => { const cost = costWith(c.id); return { catId: c.id, name: c.name, pack: c.pack, qty: c.qty, rate: c.rate, line: c.rate * c.qty, belowCost: cost != null && c.rate < cost }; });
    const total = items.reduce((s, i) => s + i.line, 0);
    const orderNo = nextOrderNo();
    const txns = [...next.txns, { id: uid(), dealerId, type: "order", orderNo, date: oDate || today(), note: oNote.trim(), items, total, created: Date.now() }];
    const amt = parseFloat(ocPayAmt);
    if (amt && amt > 0) txns.push({ id: uid(), dealerId, type: "payment", date: oDate || today(), method: ocPayMethod, note: "Against order #" + orderNo, amount: amt, created: Date.now() + 1 });
    persist({ ...next, txns });
    resetComposer();
    flash("Order #" + orderNo + " saved");
    setDetailId(dealerId);
  };

  // ---- dealers ----
  const openForm = (dealer) => {
    if (dealer) { setEditId(dealer.id); setF({ name: dealer.name, phone: dealer.phone || "", area: dealer.area || "", rep: dealer.rep || "", remarks: dealer.remarks || "", followUp: dealer.followUp || "" }); }
    else { setEditId(null); setF({ name: "", phone: "", area: "", rep: "", remarks: "", followUp: "" }); }
    setFormOpen(true);
  };
  const saveForm = () => {
    const name = f.name.trim(); if (!name) return;
    const rec = { name, phone: f.phone.trim(), area: f.area.trim(), rep: f.rep.trim(), remarks: f.remarks.trim(), followUp: f.followUp || "" };
    if (editId) persist({ ...data, dealers: data.dealers.map((d) => (d.id === editId ? { ...d, ...rec } : d)) });
    else persist({ ...data, dealers: [...data.dealers, { id: uid(), ...rec, created: Date.now() }] });
    setFormOpen(false); flash(editId ? "Dealer updated" : "Dealer added");
  };
  const removeDealer = (id) => setConfirmAsk({
    msg: "Delete this dealer and all their records? This cannot be undone.",
    action: () => { persist({ ...data, dealers: data.dealers.filter((d) => d.id !== id), txns: data.txns.filter((t) => t.dealerId !== id) }); setDetailId(null); },
  });

  // ---- transactions ----
  const saveOrder = () => {
    if (!cart.length) return;
    const items = cart.map((c) => {
      const cost = costWith(c.id);
      return { catId: c.id, name: c.name, pack: c.pack, qty: c.qty, rate: c.rate, line: c.rate * c.qty, belowCost: cost != null && c.rate < cost };
    });
    const total = items.reduce((s, i) => s + i.line, 0);
    if (editOrderId) {
      persist({ ...data, txns: data.txns.map((t) => (t.id === editOrderId ? { ...t, date: oDate || today(), note: oNote.trim(), items, total } : t)) });
    } else {
      persist({ ...data, txns: [...data.txns, { id: uid(), dealerId: detailId, type: "order", date: oDate || today(), note: oNote.trim(), items, total, created: Date.now() }] });
    }
    setCart([]); setONote(""); setODate(today()); setPSearch(""); const wasEdit = editOrderId; setEditOrderId(null); setTab("ledger"); flash(wasEdit ? "Order updated" : "Order saved");
  };
  const editOrder = (t) => {
    setCart(t.items.map((i) => { const p = CATALOG.find((x) => x.id === i.catId); return { id: i.catId, name: i.name, pack: i.pack, dealer: p ? p.dealer : i.rate, rate: i.rate, qty: i.qty }; }));
    setEditOrderId(t.id); setODate(t.date); setONote(t.note || ""); setTab("order");
  };
  const cancelEditOrder = () => { setEditOrderId(null); setCart([]); setONote(""); setODate(today()); };
  const savePayEdit = () => {
    if (!payEdit) return;
    const amt = parseFloat(payEdit.amount); if (!amt || amt <= 0) return;
    persist({ ...data, txns: data.txns.map((t) => (t.id === payEdit.id ? { ...t, amount: amt, date: payEdit.date || today(), method: payEdit.method, note: (payEdit.note || "").trim() } : t)) });
    setPayEdit(null); flash("Payment updated");
  };
  const savePayment = () => {
    const amt = parseFloat(payAmt); if (!amt || amt <= 0) return;
    persist({ ...data, txns: [...data.txns, { id: uid(), dealerId: detailId, type: "payment", date: payDate || today(), method: payMethod, note: payNote.trim(), amount: amt, created: Date.now() }] });
    setPayAmt(""); setPayNote(""); setPayDate(today()); setTab("ledger"); flash("Payment recorded");
  };
  const removeTxn = (id) => {
    const t = data.txns.find((x) => x.id === id);
    if (!t) return;
    persist((prev) => ({ ...prev, txns: prev.txns.filter((x) => x.id !== id) }));
    flash((t.type === "order" ? "Order" : "Payment") + " deleted", () => {
      persist((prev) => (prev.txns.some((x) => x.id === t.id) ? prev : { ...prev, txns: [...prev.txns, t] }));
    });
  };

  const balanceOf = (id) => data.txns.filter((t) => t.dealerId === id).reduce((b, t) => b + (t.type === "order" ? t.total : -t.amount), 0);
  const ledgerFor = (id) => {
    const list = data.txns.filter((t) => t.dealerId === id).slice()
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.created - b.created));
    let run = 0;
    return list.map((t) => { run += t.type === "order" ? t.total : -t.amount; return { ...t, bal: run }; }).reverse();
  };

  // ---- cart ----
  const addToCart = (p) => setCart((prev) => {
    const i = prev.findIndex((c) => c.id === p.id);
    if (i >= 0) { const cp = [...prev]; cp[i] = { ...cp[i], qty: cp[i].qty + 1 }; return cp; }
    return [...prev, { id: p.id, name: p.name, pack: p.pack, dealer: p.dealer, rate: p.dealer, qty: 1 }];
  });
  const setQty = (id, q) => setCart((p) => p.map((c) => (c.id === id ? { ...c, qty: Math.max(1, q || 1) } : c)));
  const setRate = (id, r) => setCart((p) => p.map((c) => (c.id === id ? { ...c, rate: r } : c)));
  const delCartItem = (id) => setCart((p) => p.filter((c) => c.id !== id));

  // ---- stock imports ----
  const onStockFile = async (e) => {
    const file = e.target.files && e.target.files[0]; if (!file) return; e.target.value = "";
    try {
      const rows = await parseFile(file);
      const { map, matched, unmatched } = buildImport(rows);
      persist({ ...data, stock: map, stockMeta: { updated: today(), matched, unmatched } });
      flash(matched ? "Stock updated: " + matched + " products" + (unmatched.length ? ", " + unmatched.length + " unmatched" : "") : "No products matched — check the file columns");
    } catch { flash("Could not read that file"); }
  };
  const loadPaste = () => {
    const { map, matched, unmatched, lines } = parseStockText(pasteText);
    if (!lines) { flash("No stock rows found in the pasted text"); return; }
    persist({ ...data, stock: map, stockMeta: { updated: today(), matched, unmatched } });
    setShowPaste(false); setPasteText("");
    flash(matched ? "Stock updated: " + matched + " products" + (unmatched.length ? ", " + unmatched.length + " not in catalog" : "") : "No products matched the catalog");
  };

  // ---- costing editor ----
  const groupCpu = (skus) => { const c = data.costing[skus[0].id]; return c && c[0] != null ? c[0] : ""; };
  const setCpu = (grp, val) => {
    const v = val === "" ? null : parseNum(val);
    const next = { ...data.costing };
    grp.skus.forEach((p) => { const cur = next[p.id] || [null, 0]; next[p.id] = [v, cur[1] || 0]; });
    persist({ ...data, costing: next });
  };
  const setPacking = (id, val) => {
    const v = val === "" ? 0 : (parseNum(val) || 0);
    const cur = data.costing[id] || [null, 0];
    persist({ ...data, costing: { ...data.costing, [id]: [cur[0], v] } });
  };
  const setThreshold = (v) => persist({ ...data, settings: { ...data.settings, lowStock: Math.max(0, parseInt(v) || 0) } });

  // ---- derived ----
  const areas = useMemo(() => [...new Set(data.dealers.map((d) => (d.area || "").trim()).filter(Boolean))].sort(), [data.dealers]);
  const reps = useMemo(() => [...new Set(data.dealers.map((d) => (d.rep || "").trim()).filter(Boolean))].sort(), [data.dealers]);
  const dealersSorted = useMemo(() => {
    const q = dealerSearch.trim().toLowerCase();
    return data.dealers.filter((d) =>
      (fArea === "All" || (d.area || "").trim() === fArea) &&
      (fRep === "All" || (d.rep || "").trim() === fRep) &&
      (!q || d.name.toLowerCase().includes(q) || (d.phone || "").includes(q))
    ).slice().sort((a, b) => a.name.localeCompare(b.name));
  }, [data.dealers, dealerSearch, fArea, fRep]);
  const filteredProducts = useMemo(() => {
    const q = pSearch.trim().toLowerCase();
    return CATALOG.filter((p) => (pCat === "All" || p.cat === pCat) && (!q || p.name.toLowerCase().includes(q) || p.pack.toLowerCase().includes(q)));
  }, [pSearch, pCat]);
  const lowStock = useMemo(() =>
    CATALOG.filter((p) => data.stock[p.id] != null && data.stock[p.id] <= threshold)
      .map((p) => ({ ...p, qty: data.stock[p.id] })).sort((a, b) => a.qty - b.qty), [data.stock, threshold]);
  const stockList = useMemo(() => {
    const q = stockSearch.trim().toLowerCase();
    return CATALOG.filter((p) => !q || pname(p.id).toLowerCase().includes(q) || p.pack.toLowerCase().includes(q))
      .map((p) => ({ ...p, qty: data.stock[p.id] }));
  }, [data.stock, data.names, stockSearch]);
  const costGroupsFiltered = useMemo(() => {
    const q = costSearch.trim().toLowerCase();
    return q ? COST_GROUPS.filter((g) => g.name.toLowerCase().includes(q)) : COST_GROUPS;
  }, [costSearch]);
  const glance = useMemo(() => {
    let outstanding = 0, advance = 0;
    const byArea = {}, byRep = {}, debtors = [];
    data.dealers.forEach((d) => {
      const b = balanceOf(d.id);
      const ak = (d.area || "").trim() || "Unassigned";
      const rk = (d.rep || "").trim() || "Unassigned";
      byArea[ak] = byArea[ak] || { out: 0, n: 0 }; byArea[ak].n++;
      byRep[rk] = byRep[rk] || { out: 0, n: 0 }; byRep[rk].n++;
      if (b > 0) {
        outstanding += b; byArea[ak].out += b; byRep[rk].out += b;
        const pays = data.txns.filter((t) => t.dealerId === d.id && t.type === "payment");
        const lastPay = pays.length ? pays.reduce((mx, t) => (t.date > mx ? t.date : mx), pays[0].date) : null;
        debtors.push({ name: d.name, area: d.area, bal: b, id: d.id, lastPay });
      }
      else if (b < 0) advance += -b;
    });
    const belowCostOrders = data.txns.filter((t) => t.type === "order" && t.items.some((i) => i.belowCost)).length;
    const sortObj = (o) => Object.entries(o).sort((a, b) => b[1].out - a[1].out);
    return { outstanding, advance, byArea: sortObj(byArea), byRep: sortObj(byRep), debtors: debtors.sort((a, b) => b.bal - a.bal).slice(0, 5), belowCostOrders };
  }, [data]);
  const followUps = useMemo(() =>
    data.dealers.map((d) => ({ d, diff: d.followUp ? daysBetween(d.followUp) : null }))
      .filter((x) => x.diff != null && x.diff >= 0)
      .sort((a, b) => b.diff - a.diff), [data.dealers]);

  const repInfo = useMemo(() => {
    if (!repView) return null;
    const ds = data.dealers.filter((d) => (((d.rep || "").trim()) || "Unassigned") === repView);
    let out = 0, sales = 0, paid = 0, profit = 0;
    const rows = ds.map((d) => {
      const tx = data.txns.filter((t) => t.dealerId === d.id);
      let b = 0, s = 0, p = 0, pr = 0;
      tx.forEach((t) => { if (t.type === "order") { b += t.total; s += t.total; pr += orderProfit(t); } else { b -= t.amount; p += t.amount; } });
      if (b > 0) out += b;
      sales += s; paid += p; profit += pr;
      return { id: d.id, name: d.name, area: d.area, bal: b, sales: s };
    }).sort((a, b) => b.bal - a.bal);
    return { rows, out, sales, paid, profit, count: ds.length };
  }, [data, repView]);

  const cartTotal = cart.reduce((s, c) => s + c.rate * c.qty, 0);
  const active = data.dealers.find((d) => d.id === detailId);

  // ---- exports ----
  const buildRows = (dealer) => {
    const list = data.txns.filter((t) => t.dealerId === dealer.id).slice()
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.created - b.created));
    let run = 0; const out = [];
    list.forEach((t) => {
      run += t.type === "order" ? t.total : -t.amount;
      const details = t.type === "order"
        ? t.items.map((i) => i.name + " " + i.pack + " x" + i.qty + " @" + Math.round(i.rate)).join("; ") + (t.note ? " [" + t.note + "]" : "")
        : "Payment (" + t.method + ")" + (t.note ? " [" + t.note + "]" : "");
      out.push([t.date, t.type === "order" ? "Purchase" : "Payment", details,
        t.type === "order" ? Math.round(t.total) : "", t.type === "payment" ? Math.round(t.amount) : "", Math.round(run)]);
    });
    return { out, closing: run };
  };
  const exportAll = () => {
    const rows = [["Dealer", "Area", "Salesperson", "Phone", "Date", "Type", "Details", "Purchase (Rs)", "Payment (Rs)", "Balance (Rs)"]];
    data.dealers.slice().sort((a, b) => a.name.localeCompare(b.name)).forEach((d) => {
      buildRows(d).out.forEach((r) => rows.push([d.name, d.area || "", d.rep || "", d.phone || "", ...r]));
    });
    downloadFile("Royal-Agro-All-Ledgers-" + today() + ".csv", rows.map((r) => r.map(csvCell).join(",")).join("\n"));
    flash("Exported all ledgers");
  };
  const exportStatement = (d) => {
    const { out, closing } = buildRows(d);
    const rows = [["Statement - " + d.name], [((d.area || "") + "  " + (d.phone || "")).trim()], [d.rep ? "Salesperson: " + d.rep : ""],
      ["Generated " + prettyDate(today())], [],
      ["Date", "Type", "Details", "Purchase (Rs)", "Payment (Rs)", "Balance (Rs)"], ...out, [],
      ["", "", "Closing balance", "", "", Math.round(closing)]];
    downloadFile("Statement-" + d.name.replace(/[^a-z0-9]+/gi, "-") + "-" + today() + ".csv", rows.map((r) => r.map(csvCell).join(",")).join("\n"));
    flash("Statement exported");
  };
  const openShare = (title, text) => setShareModal({ title, text });
  const doNativeShare = async () => {
    if (!shareModal) return;
    if (navigator.share) {
      try { await navigator.share({ title: shareModal.title, text: shareModal.text }); return; }
      catch (e) { if (e && e.name === "AbortError") return; }
    }
    doCopyShare();
  };
  const doCopyShare = async () => {
    if (!shareModal) return;
    try { await navigator.clipboard.writeText(shareModal.text); flash("Copied — paste it in WhatsApp or SMS"); }
    catch { flash("Couldn't copy automatically — select the text below and copy it"); }
  };
  const shareStatement = (d) => {
    const list = ledgerFor(d.id); // newest first
    const bal = balanceOf(d.id);
    const recent = list.slice(0, 8).slice().reverse();
    const lines = recent.map((t) => "• " + prettyDate(t.date) + " " + (t.type === "order" ? "Purchase Rs " + plain(t.total) : "Payment Rs " + plain(t.amount)) + " (Bal Rs " + plain(t.bal) + ")");
    const header = "*Royal Agro Services*\n" + d.name + (d.area ? " — " + d.area : "") + (d.rep ? "\nSalesperson: " + d.rep : "") + "\n";
    const body = lines.length ? "\nRecent entries:\n" + lines.join("\n") + "\n" : "\nNo entries yet.\n";
    const closingLine = "\n*Outstanding balance: Rs " + plain(Math.abs(bal)) + "*" + (bal > 0 ? " — please clear at your earliest convenience" : bal < 0 ? " advance held on account" : " — fully settled");
    openShare(d.name + " — statement", header + body + closingLine + "\n\nRoyal Agro Ledger · " + prettyDate(today()));
  };
  const backup = () => { downloadFile("Royal-Agro-Backup-" + today() + ".json", JSON.stringify(data, null, 2), "application/json"); flash("Backup downloaded"); };
  const onRestore = (e) => {
    const file = e.target.files && e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try { const p = JSON.parse(reader.result); if (!p.dealers || !p.txns) throw new Error();
        setConfirmAsk({ msg: "Replace all current data with this backup?", action: () => {
          persist({ ...DEFAULTS, ...p, settings: { ...DEFAULTS.settings, ...(p.settings || {}) }, costing: { ...COST_SEED, ...(p.costing || {}) } });
          flash("Backup restored");
        } });
      } catch { flash("Could not read that backup file"); }
    };
    reader.readAsText(file); e.target.value = "";
  };
  const openDealer = (id) => { setDetailId(id); setTab("ledger"); setCart([]); setPSearch(""); setONote(""); setEditOrderId(null); setRepView(null); };

  const BalancePill = ({ id }) => {
    const b = balanceOf(id);
    const cls = b > 0 ? "text-red-700 bg-red-50 border-red-200" : b < 0 ? "text-emerald-700 bg-emerald-50 border-emerald-200" : "text-slate-600 bg-slate-100 border-slate-200";
    return <span className={"inline-flex items-center rounded-full border font-semibold px-2 py-0.5 text-xs " + cls}>{b === 0 ? "Settled" : (b > 0 ? "Owes " : "Advance ") + fmt(Math.abs(b))}</span>;
  };
  const stockBadge = (id) => {
    const q = data.stock[id];
    if (q == null) return null;
    const cls = q <= 0 ? "bg-red-100 text-red-700" : q <= threshold ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700";
    return <span className={"text-xs px-1.5 py-0.5 rounded " + cls}>Stock {plain(q)}</span>;
  };

  if (!loaded) return <div className="min-h-screen flex items-center justify-center bg-slate-50 text-slate-500">Loading…</div>;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      {/* Header */}
      <div className="bg-emerald-700 text-white sticky top-0 z-20 shadow-sm">
        <div className="max-w-2xl mx-auto px-4 py-3 flex items-center gap-3">
          {(detailId || repView) ? (
            <button onClick={() => (detailId ? setDetailId(null) : setRepView(null))} className="shrink-0 -ml-1 p-1.5 rounded-lg hover:bg-emerald-600"><ArrowLeft size={20} /></button>
          ) : (
            <div className="shrink-0 h-9 w-9 rounded-lg bg-emerald-600 flex items-center justify-center"><Building2 size={20} /></div>
          )}
          <div className="min-w-0 flex-1">
            <div className="font-bold leading-tight truncate">{detailId && active ? active.name : repView ? repView : "Royal Agro Services"}</div>
            <div className="text-emerald-100 text-xs truncate">
              {detailId && active ? ([active.area, active.rep && "Rep: " + active.rep, active.phone].filter(Boolean).join(" · ") || "Dealer account") : repView ? "Salesperson performance" : "Dealer Ledger · 35% off invoice"}
            </div>
          </div>
          {!detailId && !repView && <button onClick={() => setShowSettings(true)} className="shrink-0 p-1.5 rounded-lg hover:bg-emerald-600"><SettingsIcon size={20} /></button>}
        </div>
        {!detailId && !repView && (
          <div className="max-w-2xl mx-auto px-3 pb-2 flex gap-1 text-sm font-semibold overflow-x-auto">
            {[{ k: "today", label: "Today", icon: TrendingUp }, { k: "order", label: "Order", icon: ShoppingCart }, { k: "glance", label: "Overview", icon: LayoutDashboard }, { k: "dealers", label: "Dealers", icon: Users }, { k: "stock", label: "Stock", icon: Boxes }].map((t) => {
              const Icon = t.icon;
              return <button key={t.k} onClick={() => setView(t.k)} className={"shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-lg " + (view === t.k ? "bg-white text-emerald-700" : "text-emerald-100 hover:bg-emerald-600")}><Icon size={15} />{t.label}</button>;
            })}
          </div>
        )}
      </div>

      {msg && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-50 bg-slate-900 text-white text-sm pl-4 pr-2 py-2 rounded-full shadow-lg max-w-xs flex items-center gap-2">
          <span className="truncate">{msg}</span>
          {msgUndo && <button onClick={() => { const fn = msgUndo; setMsg(""); setMsgUndo(null); fn(); }} className="font-semibold text-emerald-400 px-2.5 py-1 rounded-full hover:bg-slate-700 shrink-0">Undo</button>}
        </div>
      )}

      <div className="max-w-2xl mx-auto px-4 py-4">
        {detailId && active ? (
          <>
            <div className="bg-white rounded-2xl border border-slate-200 p-4 mb-3">
              <div className="flex items-start justify-between gap-3">
                <div>
                  <div className="text-slate-500 text-xs mb-1">Current balance</div>
                  {(() => {
                    const b = balanceOf(active.id);
                    const color = b > 0 ? "text-red-700" : b < 0 ? "text-emerald-700" : "text-slate-700";
                    return <><div className={"text-3xl font-bold tabular-nums " + color}>{fmt(Math.abs(b))}</div>
                      <div className="text-xs text-slate-500 mt-0.5">{b > 0 ? "Dealer owes you" : b < 0 ? "Advance / credit" : "Fully settled"}</div></>;
                  })()}
                </div>
                <div className="flex gap-1">
                  <button onClick={() => openForm(active)} title="Edit dealer" className="p-2 rounded-lg text-slate-400 hover:text-emerald-700 hover:bg-emerald-50"><Pencil size={17} /></button>
                  <button onClick={() => removeDealer(active.id)} title="Delete dealer" className="p-2 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50"><Trash2 size={17} /></button>
                </div>
              </div>
              {active.remarks && <div className="mt-2 text-xs text-slate-600 bg-slate-50 rounded-lg px-2.5 py-1.5">{active.remarks}</div>}
              {active.followUp && (() => {
                const fu = followUpStatus(active);
                return (
                  <div className={"mt-2 text-xs rounded-lg px-2.5 py-1.5 flex items-center justify-between gap-2 " + (fu.tone === "red" ? "bg-red-50 text-red-700" : fu.tone === "amber" ? "bg-amber-50 text-amber-700" : "bg-slate-50 text-slate-600")}>
                    <span className="inline-flex items-center gap-1"><Calendar size={12} /> {fu.label}</span>
                    <button onClick={() => persist((prev) => ({ ...prev, dealers: prev.dealers.map((d) => (d.id === active.id ? { ...d, followUp: "" } : d)) }))} className="font-semibold underline shrink-0">Mark done</button>
                  </div>
                );
              })()}
              <div className="mt-3 flex items-center gap-3">
                <button onClick={() => exportStatement(active)} className="inline-flex items-center gap-1 text-sm text-emerald-700 font-semibold"><Download size={15} /> Export</button>
                <button onClick={() => shareStatement(active)} className="inline-flex items-center gap-1 text-sm text-emerald-700 font-semibold"><Share2 size={15} /> Share</button>
              </div>
            </div>

            <div className="grid grid-cols-3 gap-1 bg-slate-100 p-1 rounded-xl mb-3 text-sm font-semibold">
              {[{ k: "ledger", label: "Ledger", icon: Receipt }, { k: "order", label: "New order", icon: ShoppingCart }, { k: "payment", label: "Payment", icon: Wallet }].map((t) => {
                const Icon = t.icon;
                return <button key={t.k} onClick={() => setTab(t.k)} className={"flex items-center justify-center gap-1.5 py-2 rounded-lg " + (tab === t.k ? "bg-white shadow text-emerald-700" : "text-slate-500")}><Icon size={15} />{t.label}</button>;
              })}
            </div>

            {tab === "ledger" && (
              <div className="space-y-2">
                {ledgerFor(active.id).length === 0 ? (
                  <div className="text-center text-slate-500 text-sm py-10 border border-dashed border-slate-300 rounded-xl">No entries yet. Add an order or a payment.</div>
                ) : ledgerFor(active.id).map((t) => {
                  const isOrder = t.type === "order"; const open = openRows[t.id]; const flagged = isOrder && t.items.some((i) => i.belowCost);
                  return (
                    <div key={t.id} className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                      <div className="p-3 flex items-start justify-between gap-3">
                        <div className="min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className={"text-xs font-semibold px-2 py-0.5 rounded-full " + (isOrder ? "bg-amber-50 text-amber-700" : "bg-emerald-50 text-emerald-700")}>{isOrder ? "Purchase" : "Payment"}</span>
                            <span className="text-xs text-slate-500">{prettyDate(t.date)}</span>
                            {flagged && <span className="text-xs font-semibold text-red-700 inline-flex items-center gap-0.5"><AlertTriangle size={12} /> below cost</span>}
                          </div>
                          {isOrder ? (
                            <button onClick={() => setOpenRows((o) => ({ ...o, [t.id]: !o[t.id] }))} className="mt-1 text-sm text-slate-600 inline-flex items-center gap-1">{t.items.length} item{t.items.length > 1 ? "s" : ""}{open ? <ChevronUp size={14} /> : <ChevronDown size={14} />}</button>
                          ) : <div className="mt-1 text-sm text-slate-600">{t.method}{t.note ? " · " + t.note : ""}</div>}
                        </div>
                        <div className="text-right shrink-0">
                          <div className={"font-bold tabular-nums " + (isOrder ? "text-red-700" : "text-emerald-700")}>{isOrder ? "+" : "−"}{fmt(isOrder ? t.total : t.amount)}</div>
                          <div className="text-xs text-slate-500 tabular-nums">Bal {fmt(t.bal)}</div>
                          <div className="mt-1 flex items-center gap-3 justify-end">
                            <button onClick={() => (isOrder ? editOrder(t) : setPayEdit({ id: t.id, amount: String(t.amount), date: t.date, method: t.method, note: t.note || "" }))} className="text-xs text-slate-400 hover:text-emerald-700">edit</button>
                            <button onClick={() => removeTxn(t.id)} className="text-xs text-slate-400 hover:text-red-600">delete</button>
                          </div>
                        </div>
                      </div>
                      {isOrder && open && (
                        <div className="border-t border-slate-100 bg-slate-50 px-3 py-2 space-y-1">
                          {t.note ? <div className="text-xs text-slate-500 mb-1">Note: {t.note}</div> : null}
                          {t.items.map((i, idx) => (
                            <div key={idx} className="flex justify-between text-xs text-slate-600 gap-2">
                              <span className="min-w-0 truncate">{i.name} · {i.pack} × {i.qty} @ {plain(i.rate)}{i.belowCost && <span className="text-red-600 font-semibold"> ⚠</span>}</span>
                              <span className="tabular-nums shrink-0">{fmt(i.line)}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            {tab === "order" && (
              <div>
                {editOrderId && (
                  <div className="mb-3 flex items-center justify-between bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-xs text-amber-800">
                    <span className="font-medium">Editing an existing order — adjust items and save</span>
                    <button onClick={cancelEditOrder} className="font-semibold underline">Cancel</button>
                  </div>
                )}
                {cart.length > 0 && (
                  <div className="bg-white rounded-xl border border-slate-200 p-3 mb-3">
                    <div className="text-sm font-semibold mb-2">Order items</div>
                    <div className="space-y-3">
                      {cart.map((c) => {
                        const cost = costWith(c.id); const below = cost != null && c.rate < cost;
                        return (
                          <div key={c.id} className={"rounded-lg border p-2 " + (below ? "border-red-300 bg-red-50" : "border-slate-100")}>
                            <div className="flex items-start justify-between gap-2">
                              <div className="min-w-0">
                                <div className="text-sm font-medium truncate">{c.name}</div>
                                <div className="text-xs text-slate-500 flex items-center gap-1.5 flex-wrap">{c.pack} · list {plain(c.dealer)} {stockBadge(c.id)}</div>
                              </div>
                              <button onClick={() => delCartItem(c.id)} className="text-slate-400 hover:text-red-600 shrink-0"><X size={16} /></button>
                            </div>
                            <div className="flex items-center gap-2 mt-2">
                              <div className="flex items-center border border-slate-200 rounded-lg bg-white">
                                <button onClick={() => setQty(c.id, c.qty - 1)} className="px-2 py-1.5 text-slate-600"><Minus size={14} /></button>
                                <input value={c.qty} onChange={(e) => setQty(c.id, parseInt(e.target.value.replace(/\D/g, "")) || 1)} inputMode="numeric" className="w-9 text-center text-sm outline-none tabular-nums" />
                                <button onClick={() => setQty(c.id, c.qty + 1)} className="px-2 py-1.5 text-slate-600"><Plus size={14} /></button>
                              </div>
                              <div className="flex items-center gap-1 flex-1">
                                <span className="text-xs text-slate-500">Rate</span>
                                <input value={c.rate} onChange={(e) => setRate(c.id, parseNum(e.target.value) ?? 0)} inputMode="decimal" className={"w-20 px-2 py-1.5 rounded-lg border text-sm outline-none tabular-nums " + (below ? "border-red-400 text-red-700" : "border-slate-200 focus:border-emerald-400")} />
                                {c.rate !== c.dealer && <button onClick={() => setRate(c.id, c.dealer)} title="Reset to list rate" className="text-slate-400 hover:text-emerald-700"><RotateCcw size={14} /></button>}
                              </div>
                              <div className="w-20 text-right text-sm font-semibold tabular-nums">{fmt(c.rate * c.qty)}</div>
                            </div>
                            {below && <div className="text-xs text-red-700 font-semibold mt-1 inline-flex items-center gap-1"><AlertTriangle size={12} /> Rate is below cost</div>}
                          </div>
                        );
                      })}
                    </div>
                    <div className="border-t border-slate-100 mt-3 pt-3 space-y-2">
                      <div className="flex gap-2">
                        <input type="date" value={oDate} onChange={(e) => setODate(e.target.value)} className="flex-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                        <input value={oNote} onChange={(e) => setONote(e.target.value)} placeholder="Bill # / note" className="flex-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                      </div>
                      <div className="flex items-center justify-between"><span className="text-slate-600 text-sm">Order total</span><span className="text-lg font-bold tabular-nums">{fmt(cartTotal)}</span></div>
                      <button onClick={saveOrder} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-semibold text-sm">{editOrderId ? "Update order" : "Save order to ledger"}</button>
                    </div>
                  </div>
                )}

                <div className="bg-white rounded-xl border border-slate-200 p-3">
                  <div className="relative mb-2">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input value={pSearch} onChange={(e) => setPSearch(e.target.value)} placeholder="Search product (name or pack)" className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                  </div>
                  <div className="flex gap-1.5 overflow-x-auto pb-2 mb-1">
                    {CATS.map((c) => <button key={c} onClick={() => setPCat(c)} className={"shrink-0 text-xs px-3 py-1.5 rounded-full border " + (pCat === c ? "bg-emerald-600 text-white border-emerald-600" : "bg-white border-slate-200 text-slate-600")}>{c}</button>)}
                  </div>
                  <div className="max-h-96 overflow-y-auto -mx-1 divide-y divide-slate-100">
                    {filteredProducts.map((p) => (
                      <button key={p.id} onClick={() => addToCart(p)} className="w-full text-left px-1 py-2.5 flex items-center justify-between gap-2 hover:bg-slate-50 rounded">
                        <div className="min-w-0">
                          <div className="text-sm font-medium truncate">{p.name}</div>
                          <div className="text-xs text-slate-500 flex items-center gap-1.5">{p.pack} · <span className="line-through">{plain(p.rate)}</span> {stockBadge(p.id)}</div>
                        </div>
                        <div className="text-right shrink-0">
                          <div className="text-sm font-semibold text-emerald-700 tabular-nums">{fmt(p.dealer)}</div>
                          <div className="text-xs text-emerald-600 inline-flex items-center gap-0.5"><Plus size={11} /> add</div>
                        </div>
                      </button>
                    ))}
                    {filteredProducts.length === 0 && <div className="text-center text-slate-500 text-sm py-6">No products match.</div>}
                  </div>
                </div>
              </div>
            )}

            {tab === "payment" && (
              <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-3">
                <div className="text-sm font-semibold flex items-center gap-1"><Wallet size={15} /> Record a payment</div>
                <div><label className="text-xs text-slate-500">Amount received (Rs)</label>
                  <input value={payAmt} onChange={(e) => setPayAmt(e.target.value.replace(/[^\d.]/g, ""))} inputMode="decimal" placeholder="0" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-lg font-semibold outline-none focus:border-emerald-400 tabular-nums" /></div>
                <div className="flex gap-2">
                  <div className="flex-1"><label className="text-xs text-slate-500">Date</label>
                    <input type="date" value={payDate} onChange={(e) => setPayDate(e.target.value)} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
                  <div className="flex-1"><label className="text-xs text-slate-500">Method</label>
                    <select value={payMethod} onChange={(e) => setPayMethod(e.target.value)} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 bg-white">
                      <option>Cash</option><option>Bank Transfer</option><option>Cheque</option><option>Easypaisa/JazzCash</option><option>Other</option></select></div>
                </div>
                <div><label className="text-xs text-slate-500">Note (optional)</label>
                  <input value={payNote} onChange={(e) => setPayNote(e.target.value)} placeholder="e.g. cheque no., against Bill #" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
                <button onClick={savePayment} disabled={!parseFloat(payAmt)} className="w-full bg-emerald-600 disabled:bg-slate-300 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-semibold text-sm">Record payment</button>
              </div>
            )}
          </>
        ) : repView ? (
          <div className="space-y-3">
            <div className="bg-white rounded-2xl border border-slate-200 p-4">
              <div className="text-slate-500 text-xs">Pending balance</div>
              <div className="text-3xl font-bold text-red-700 tabular-nums">{fmt(repInfo.out)}</div>
              <div className="text-xs text-slate-500 mt-0.5">{repInfo.count} dealer{repInfo.count !== 1 ? "s" : ""} handled by {repView}</div>
            </div>
            <div className="grid grid-cols-3 gap-2">
              <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Total sales</div><div className="text-base font-bold tabular-nums">{plain(repInfo.sales)}</div></div>
              <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Collected</div><div className="text-base font-bold text-emerald-700 tabular-nums">{plain(repInfo.paid)}</div></div>
              <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Profit</div><div className="text-base font-bold text-emerald-700 tabular-nums">{plain(repInfo.profit)}</div></div>
            </div>
            <div className="bg-white rounded-xl border border-slate-200 p-3">
              <div className="text-sm font-semibold mb-2">Dealers (by balance)</div>
              {repInfo.rows.length === 0 ? <div className="text-center text-slate-500 text-sm py-4">No dealers assigned to this salesperson.</div> : (
                <div className="space-y-1.5">{repInfo.rows.map((d) => (
                  <button key={d.id} onClick={() => openDealer(d.id)} className="w-full flex items-center justify-between gap-2 text-left hover:bg-slate-50 rounded-lg px-1 -mx-1 py-1">
                    <div className="min-w-0"><div className="text-sm truncate">{d.name}{d.area ? <span className="text-slate-400 text-xs"> · {d.area}</span> : null}</div><div className="text-xs text-slate-400">sales {plain(d.sales)}</div></div>
                    <span className={"text-sm font-semibold tabular-nums shrink-0 " + (d.bal > 0 ? "text-red-700" : d.bal < 0 ? "text-emerald-700" : "text-slate-500")}>{d.bal === 0 ? "settled" : (d.bal > 0 ? "" : "adv ") + fmt(Math.abs(d.bal))}</span>
                  </button>))}</div>
              )}
            </div>
            <p className="text-xs text-slate-400 px-1">Profit is revenue minus cost on costed items and stays private to this screen.</p>
          </div>
        ) : view === "order" ? (
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <div className="text-sm font-semibold text-slate-700 flex items-center gap-1"><ShoppingCart size={15} /> New order</div>
              <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 rounded-full px-2 py-0.5">#{nextOrderNo()}</span>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-2">
              <label className="text-xs text-slate-500">Dealer</label>
              <select value={ocDealerId} onChange={(e) => setOcDealerId(e.target.value)} className="w-full px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 bg-white">
                <option value="">Choose a dealer…</option>
                <option value="__new__">➕ Add a new dealer</option>
                {data.dealers.slice().sort((a, b) => a.name.localeCompare(b.name)).map((d) => (
                  <option key={d.id} value={d.id}>{d.name}{d.area ? " — " + d.area : ""}</option>
                ))}
              </select>
              {ocDealerId && ocDealerId !== "__new__" && (() => {
                const d = data.dealers.find((x) => x.id === ocDealerId);
                if (!d) return null;
                return (
                  <div className="text-xs text-slate-600 bg-slate-50 rounded-lg px-3 py-2 space-y-0.5">
                    {d.area ? <div><span className="text-slate-400">Area:</span> {d.area}</div> : null}
                    {d.rep ? <div><span className="text-slate-400">Salesperson:</span> {d.rep}</div> : null}
                    {d.phone ? <div><span className="text-slate-400">Phone:</span> {d.phone}</div> : null}
                    <div><span className="text-slate-400">Balance:</span> {fmt(balanceOf(d.id))}</div>
                  </div>
                );
              })()}
              {ocDealerId === "__new__" && (
                <div className="space-y-2 pt-1">
                  <input value={ocNew.name} onChange={(e) => setOcNew({ ...ocNew, name: e.target.value })} placeholder="Dealer / shop name" className="w-full px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                  <div className="flex gap-2">
                    <input value={ocNew.area} onChange={(e) => setOcNew({ ...ocNew, area: e.target.value })} list="oc-areas" placeholder="Area" className="flex-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                    <input value={ocNew.phone} onChange={(e) => setOcNew({ ...ocNew, phone: e.target.value })} inputMode="tel" placeholder="Phone" className="flex-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                  </div>
                  <input value={ocNew.rep} onChange={(e) => setOcNew({ ...ocNew, rep: e.target.value })} list="oc-reps" placeholder="Salesperson (choose existing or type new)" className="w-full px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                  <datalist id="oc-areas">{areas.map((a) => <option key={a} value={a} />)}</datalist>
                  <datalist id="oc-reps">{reps.map((r) => <option key={r} value={r} />)}</datalist>
                </div>
              )}
            </div>

            {cart.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2">Products</div>
                <div className="space-y-3">
                  {cart.map((c) => {
                    const cost = costWith(c.id); const below = cost != null && c.rate < cost;
                    return (
                      <div key={c.id} className={"rounded-lg border p-2 " + (below ? "border-red-300 bg-red-50" : "border-slate-100")}>
                        <div className="flex items-start justify-between gap-2">
                          <div className="min-w-0"><div className="text-sm font-medium truncate">{pname(c.id)}</div><div className="text-xs text-slate-500 flex items-center gap-1.5 flex-wrap">{c.pack} · list {plain(c.dealer)} {stockBadge(c.id)}</div></div>
                          <button onClick={() => delCartItem(c.id)} className="text-slate-400 hover:text-red-600 shrink-0"><X size={16} /></button>
                        </div>
                        <div className="flex items-center gap-2 mt-2">
                          <div className="flex items-center border border-slate-200 rounded-lg bg-white">
                            <button onClick={() => setQty(c.id, c.qty - 1)} className="px-2 py-1.5 text-slate-600"><Minus size={14} /></button>
                            <input value={c.qty} onChange={(e) => setQty(c.id, parseInt(e.target.value.replace(/\D/g, "")) || 1)} inputMode="numeric" className="w-9 text-center text-sm outline-none tabular-nums" />
                            <button onClick={() => setQty(c.id, c.qty + 1)} className="px-2 py-1.5 text-slate-600"><Plus size={14} /></button>
                          </div>
                          <div className="flex items-center gap-1 flex-1">
                            <span className="text-xs text-slate-500">Rate</span>
                            <input value={c.rate} onChange={(e) => setRate(c.id, parseNum(e.target.value) ?? 0)} inputMode="decimal" className={"w-20 px-2 py-1.5 rounded-lg border text-sm outline-none tabular-nums " + (below ? "border-red-400 text-red-700" : "border-slate-200 focus:border-emerald-400")} />
                            {c.rate !== c.dealer && <button onClick={() => setRate(c.id, c.dealer)} title="Reset to list rate" className="text-slate-400 hover:text-emerald-700"><RotateCcw size={14} /></button>}
                          </div>
                          <div className="w-20 text-right text-sm font-semibold tabular-nums">{fmt(c.rate * c.qty)}</div>
                        </div>
                        {below && <div className="text-xs text-red-700 font-semibold mt-1 inline-flex items-center gap-1"><AlertTriangle size={12} /> Rate is below cost</div>}
                      </div>
                    );
                  })}
                </div>
                <div className="flex items-center justify-between border-t border-slate-100 mt-3 pt-3"><span className="text-slate-600 text-sm">Order total</span><span className="text-lg font-bold tabular-nums">{fmt(cartTotal)}</span></div>
              </div>
            )}

            <div className="bg-white rounded-xl border border-slate-200 p-3">
              <div className="relative mb-2">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input value={pSearch} onChange={(e) => setPSearch(e.target.value)} placeholder="Add a product…" className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
              </div>
              <div className="flex gap-1.5 overflow-x-auto pb-2 mb-1">
                {CATS.map((c) => <button key={c} onClick={() => setPCat(c)} className={"shrink-0 text-xs px-3 py-1.5 rounded-full border " + (pCat === c ? "bg-emerald-600 text-white border-emerald-600" : "bg-white border-slate-200 text-slate-600")}>{c}</button>)}
              </div>
              <div className="max-h-80 overflow-y-auto -mx-1 divide-y divide-slate-100">
                {filteredProducts.map((p) => (
                  <button key={p.id} onClick={() => addToCart(p)} className="w-full text-left px-1 py-2.5 flex items-center justify-between gap-2 hover:bg-slate-50 rounded">
                    <div className="min-w-0"><div className="text-sm font-medium truncate">{pname(p.id)}</div><div className="text-xs text-slate-500 flex items-center gap-1.5">{p.pack} {stockBadge(p.id)}</div></div>
                    <div className="text-right shrink-0"><div className="text-sm font-semibold text-emerald-700 tabular-nums">{fmt(p.dealer)}</div><div className="text-xs text-emerald-600 inline-flex items-center gap-0.5"><Plus size={11} /> add</div></div>
                  </button>
                ))}
                {filteredProducts.length === 0 && <div className="text-center text-slate-500 text-sm py-6">No products match.</div>}
              </div>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-2">
              <div className="flex gap-2">
                <input type="date" value={oDate} onChange={(e) => setODate(e.target.value)} className="flex-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                <input value={oNote} onChange={(e) => setONote(e.target.value)} placeholder="Bill # / note" className="flex-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
              </div>
              <div className="flex gap-2">
                <input value={ocPayAmt} onChange={(e) => setOcPayAmt(e.target.value.replace(/[^\d.]/g, ""))} inputMode="decimal" placeholder="Payment now (optional)" className="flex-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 tabular-nums" />
                <select value={ocPayMethod} onChange={(e) => setOcPayMethod(e.target.value)} className="px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 bg-white"><option>Cash</option><option>Bank Transfer</option><option>Cheque</option><option>Easypaisa/JazzCash</option><option>Other</option></select>
              </div>
            </div>

            <div className="flex gap-2">
              <button onClick={saveComposerOrder} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-semibold text-sm">Save order</button>
              <button onClick={resetComposer} className="px-4 py-2.5 rounded-lg text-sm border border-slate-200">Clear</button>
            </div>
          </div>
        ) : view === "today" ? (
          <div className="space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div className="text-sm font-semibold text-slate-700">Daily performance</div>
              <input type="date" value={todayDate} onChange={(e) => setTodayDate(e.target.value)} className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400" />
            </div>
            {(() => {
              const dOrders = data.txns.filter((t) => t.type === "order" && t.date === todayDate);
              const dPays = data.txns.filter((t) => t.type === "payment" && t.date === todayDate);
              const dSales = dOrders.reduce((s, t) => s + t.total, 0);
              const dColl = dPays.reduce((s, t) => s + t.amount, 0);
              const dProfit = dOrders.reduce((s, t) => s + orderProfit(t), 0);
              const nm = (id) => { const d = data.dealers.find((x) => x.id === id); return d ? d.name : "—"; };
              return (
                <>
                  <div className="grid grid-cols-3 gap-2">
                    <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Sales</div><div className="text-lg font-bold tabular-nums">{plain(dSales)}</div></div>
                    <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Collection</div><div className="text-lg font-bold text-emerald-700 tabular-nums">{plain(dColl)}</div></div>
                    <div className="bg-emerald-600 rounded-xl p-3 text-white"><div className="text-emerald-100 text-xs">Profit</div><div className="text-lg font-bold tabular-nums">{plain(dProfit)}</div></div>
                  </div>
                  <div className="flex items-center justify-between gap-2">
                    <div className="text-xs text-slate-500">{dOrders.length} order{dOrders.length !== 1 ? "s" : ""} · {dPays.length} payment{dPays.length !== 1 ? "s" : ""} on {prettyDate(todayDate)}</div>
                    <button onClick={() => openShare("Daily summary — " + prettyDate(todayDate),
                      "*Royal Agro Services — Daily Summary*\n" + prettyDate(todayDate) + "\n\n" +
                      "Sales: Rs " + plain(dSales) + " (" + dOrders.length + " order" + (dOrders.length !== 1 ? "s" : "") + ")\n" +
                      "Collection: Rs " + plain(dColl) + " (" + dPays.length + " payment" + (dPays.length !== 1 ? "s" : "") + ")\n" +
                      "Profit: Rs " + plain(dProfit) +
                      (dOrders.length ? "\n\nSales:\n" + dOrders.map((t) => "• " + nm(t.dealerId) + " — Rs " + plain(t.total)).join("\n") : "") +
                      (dPays.length ? "\n\nCollections:\n" + dPays.map((t) => "• " + nm(t.dealerId) + " — Rs " + plain(t.amount)).join("\n") : "") +
                      "\n\n— Royal Agro Ledger")}
                      className="shrink-0 inline-flex items-center gap-1 text-xs font-semibold text-emerald-700"><Share2 size={13} /> Share</button>
                  </div>
                  <div className="bg-white rounded-xl border border-slate-200 p-3">
                    <div className="text-sm font-semibold mb-2 flex items-center gap-1"><ShoppingCart size={14} /> Sales</div>
                    {dOrders.length === 0 ? <div className="text-slate-400 text-sm py-2">No sales recorded for this day.</div> : (
                      <div className="space-y-1.5">{dOrders.map((t) => {
                        const pr = orderProfit(t);
                        return (
                          <button key={t.id} onClick={() => openDealer(t.dealerId)} className="w-full flex items-center justify-between gap-2 text-left hover:bg-slate-50 rounded-lg px-1 -mx-1 py-1">
                            <span className="text-sm truncate">{nm(t.dealerId)} <span className="text-slate-400 text-xs">· {t.items.length} item{t.items.length > 1 ? "s" : ""}</span></span>
                            <span className="text-right shrink-0"><span className="text-sm font-semibold tabular-nums">{fmt(t.total)}</span><span className="block text-xs text-emerald-600 tabular-nums">+{plain(pr)} profit</span></span>
                          </button>);
                      })}</div>
                    )}
                  </div>
                  <div className="bg-white rounded-xl border border-slate-200 p-3">
                    <div className="text-sm font-semibold mb-2 flex items-center gap-1"><Wallet size={14} /> Collection</div>
                    {dPays.length === 0 ? <div className="text-slate-400 text-sm py-2">No payments recorded for this day.</div> : (
                      <div className="space-y-1.5">{dPays.map((t) => (
                        <button key={t.id} onClick={() => openDealer(t.dealerId)} className="w-full flex items-center justify-between gap-2 text-left hover:bg-slate-50 rounded-lg px-1 -mx-1 py-1">
                          <span className="text-sm truncate">{nm(t.dealerId)} <span className="text-slate-400 text-xs">· {t.method}</span></span>
                          <span className="text-sm font-semibold text-emerald-700 tabular-nums shrink-0">{fmt(t.amount)}</span>
                        </button>))}</div>
                    )}
                  </div>
                </>
              );
            })()}
          </div>
        ) : view === "glance" ? (
          <div className="space-y-3">
            <div className="bg-white rounded-2xl border border-slate-200 p-4">
              <div className="text-slate-500 text-xs">Total outstanding (receivables)</div>
              <div className="text-3xl font-bold text-red-700 tabular-nums">{fmt(glance.outstanding)}</div>
              <div className="text-xs text-slate-500 mt-0.5">across {data.dealers.length} dealer{data.dealers.length !== 1 ? "s" : ""}</div>
            </div>
            <div className="grid grid-cols-3 gap-2">
              <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Advance held</div><div className="text-lg font-bold text-emerald-700 tabular-nums">{plain(glance.advance)}</div></div>
              <button onClick={() => setView("stock")} className="bg-white rounded-xl border border-slate-200 p-3 text-left"><div className="text-slate-500 text-xs">Low stock</div><div className={"text-lg font-bold tabular-nums " + (lowStock.length ? "text-amber-600" : "text-slate-700")}>{lowStock.length}</div></button>
              <div className="bg-white rounded-xl border border-slate-200 p-3"><div className="text-slate-500 text-xs">Below-cost sales</div><div className={"text-lg font-bold tabular-nums " + (glance.belowCostOrders ? "text-red-700" : "text-slate-700")}>{glance.belowCostOrders}</div></div>
            </div>
            {glance.debtors.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2">Top outstanding dealers</div>
                <div className="space-y-2">{glance.debtors.map((d) => (
                  <button key={d.id} onClick={() => openDealer(d.id)} className="w-full flex items-center justify-between gap-2 text-left">
                    <div className="min-w-0">
                      <div className="text-sm truncate">{d.name}{d.area ? <span className="text-slate-400 text-xs"> · {d.area}</span> : null}</div>
                      <div className="text-xs text-slate-400">{d.lastPay ? daysBetween(d.lastPay) + "d since last payment" : "No payments yet"}</div>
                    </div>
                    <span className="text-sm font-semibold text-red-700 tabular-nums shrink-0">{fmt(d.bal)}</span></button>))}</div>
              </div>
            )}
            {followUps.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2 flex items-center gap-1 text-amber-700"><Calendar size={14} /> Follow-ups due ({followUps.length})</div>
                <div className="space-y-1.5">{followUps.map(({ d, diff }) => (
                  <button key={d.id} onClick={() => openDealer(d.id)} className="w-full flex items-center justify-between gap-2 text-left">
                    <span className="text-sm truncate">{d.name}{d.area ? <span className="text-slate-400 text-xs"> · {d.area}</span> : null}</span>
                    <span className={"text-xs font-semibold shrink-0 " + (diff > 0 ? "text-red-600" : "text-amber-600")}>{diff === 0 ? "Today" : diff + "d overdue"}</span>
                  </button>))}</div>
              </div>
            )}
            {glance.byArea.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2 flex items-center gap-1"><MapPin size={14} /> Outstanding by area</div>
                <div className="space-y-1.5">{glance.byArea.map(([k, v]) => (
                  <div key={k} className="flex items-center justify-between gap-2 text-sm"><span className="text-slate-600 truncate">{k} <span className="text-slate-400 text-xs">({v.n})</span></span><span className="font-semibold tabular-nums shrink-0">{fmt(v.out)}</span></div>))}</div>
              </div>
            )}
            {glance.byRep.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2 flex items-center gap-1"><User size={14} /> By salesperson <span className="text-slate-400 text-xs font-normal">(tap for details)</span></div>
                <div className="space-y-1.5">{glance.byRep.map(([k, v]) => (
                  <button key={k} onClick={() => setRepView(k)} className="w-full flex items-center justify-between gap-2 text-sm text-left hover:bg-slate-50 rounded-lg px-1 -mx-1 py-0.5">
                    <span className="text-slate-600 truncate flex items-center gap-1">{k} <span className="text-slate-400 text-xs">({v.n})</span></span>
                    <span className="flex items-center gap-1 shrink-0"><span className="font-semibold tabular-nums">{fmt(v.out)}</span><ChevronDown size={13} className="-rotate-90 text-slate-400" /></span>
                  </button>))}</div>
              </div>
            )}
            {lowStock.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-3">
                <div className="text-sm font-semibold mb-2 flex items-center gap-1 text-amber-700"><AlertTriangle size={14} /> Low stock (≤ {threshold})</div>
                <div className="space-y-1.5">{lowStock.slice(0, 8).map((p) => (
                  <div key={p.id} className="flex items-center justify-between gap-2 text-sm"><span className="text-slate-600 truncate">{p.name} · {p.pack}</span><span className={"font-semibold tabular-nums shrink-0 " + (p.qty <= 0 ? "text-red-700" : "text-amber-700")}>{plain(p.qty)}</span></div>))}
                  {lowStock.length > 8 && <button onClick={() => setView("stock")} className="text-xs text-emerald-700 font-semibold pt-1">+{lowStock.length - 8} more →</button>}</div>
              </div>
            )}
            {data.dealers.length === 0 && (
              <div className="text-center text-slate-500 text-sm py-6 border border-dashed border-slate-300 rounded-xl">Add dealers to see receivables here. Stock is already loaded from your 6 Jul report.</div>
            )}
          </div>
        ) : view === "dealers" ? (
          <>
            <div className="flex gap-2 mb-2">
              <div className="relative flex-1">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input value={dealerSearch} onChange={(e) => setDealerSearch(e.target.value)} placeholder="Search dealers" className="w-full pl-9 pr-3 py-2.5 rounded-xl border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400" />
              </div>
              <button onClick={() => openForm(null)} className="shrink-0 inline-flex items-center gap-1 bg-emerald-600 text-white px-3 py-2.5 rounded-xl text-sm font-semibold hover:bg-emerald-700"><Plus size={16} /> Dealer</button>
            </div>
            {(areas.length > 0 || reps.length > 0) && (
              <div className="flex gap-2 mb-3">
                <select value={fArea} onChange={(e) => setFArea(e.target.value)} className="flex-1 px-3 py-2 rounded-xl border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400">
                  <option value="All">All areas</option>{areas.map((a) => <option key={a} value={a}>{a}</option>)}</select>
                <select value={fRep} onChange={(e) => setFRep(e.target.value)} className="flex-1 px-3 py-2 rounded-xl border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400">
                  <option value="All">All salespeople</option>{reps.map((r) => <option key={r} value={r}>{r}</option>)}</select>
              </div>
            )}
            {dealersSorted.length === 0 ? (
              <div className="text-center text-slate-500 text-sm py-12 border border-dashed border-slate-300 rounded-xl">{data.dealers.length === 0 ? "No dealers yet. Add your first dealer to start a ledger." : "No dealers match the filters."}</div>
            ) : (
              <div className="space-y-2">{dealersSorted.map((d) => (
                <button key={d.id} onClick={() => openDealer(d.id)} className="w-full text-left bg-white rounded-xl border border-slate-200 p-3 flex items-center justify-between gap-3 hover:border-emerald-300 active:bg-slate-50">
                  <div className="min-w-0">
                    <div className="font-semibold truncate">{d.name}</div>
                    <div className="text-slate-500 text-xs truncate">{[d.area, d.rep && "Rep: " + d.rep, d.phone].filter(Boolean).join(" · ") || "—"}</div>
                    {d.followUp && (() => { const fu = followUpStatus(d); return <div className={"text-xs font-semibold mt-0.5 " + (fu.tone === "red" ? "text-red-600" : fu.tone === "amber" ? "text-amber-600" : "text-slate-400")}>{fu.label}</div>; })()}
                  </div>
                  <BalancePill id={d.id} />
                </button>))}</div>
            )}
          </>
        ) : (
          <>
            <div className="bg-white rounded-xl border border-slate-200 p-3 mb-3">
              <div className="flex items-center justify-between gap-2">
                <div>
                  <div className="text-sm font-semibold">Inventory</div>
                  <div className="text-xs text-slate-500">{data.stockMeta ? "Updated " + prettyDate(data.stockMeta.updated) + " · " + data.stockMeta.matched + " products" + (data.stockMeta.seeded ? " (from your report)" : "") : "No stock loaded yet"}</div>
                </div>
              </div>
              <div className="flex gap-2 mt-3">
                <button onClick={() => setShowPaste((s) => !s)} className="flex-1 inline-flex items-center justify-center gap-1 bg-emerald-600 text-white px-3 py-2 rounded-lg text-sm font-semibold hover:bg-emerald-700"><ClipboardList size={15} /> Paste report</button>
                <label className="flex-1 inline-flex items-center justify-center gap-1 border border-slate-200 px-3 py-2 rounded-lg text-sm font-semibold hover:bg-slate-50 cursor-pointer"><Upload size={15} /> Upload CSV/Excel<input type="file" accept=".csv,.xlsx,.xls" onChange={onStockFile} className="hidden" /></label>
              </div>
              {showPaste && (
                <div className="mt-3">
                  <textarea value={pasteText} onChange={(e) => setPasteText(e.target.value)} rows={5} placeholder="Open your stock report PDF, select all, copy, and paste the text here…" className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs outline-none focus:border-emerald-400 resize-none" />
                  <div className="flex gap-2 mt-2">
                    <button onClick={loadPaste} disabled={!pasteText.trim()} className="flex-1 bg-emerald-600 disabled:bg-slate-300 text-white py-2 rounded-lg text-sm font-semibold">Load pasted stock</button>
                    <button onClick={() => { setShowPaste(false); setPasteText(""); }} className="px-4 py-2 rounded-lg text-sm border border-slate-200">Cancel</button>
                  </div>
                </div>
              )}
              <p className="text-xs text-slate-500 mt-2">Each update replaces the stock snapshot. Balances stocked under multiple companies are summed. Products not in your invoice catalog are skipped.</p>
              {data.stockMeta && data.stockMeta.unmatched && data.stockMeta.unmatched.length > 0 && (
                <div className="mt-2">
                  <button onClick={() => setShowUnmatched((s) => !s)} className="text-xs text-amber-700 font-semibold inline-flex items-center gap-1">{data.stockMeta.unmatched.length} rows not in catalog {showUnmatched ? <ChevronUp size={12} /> : <ChevronDown size={12} />}</button>
                  {showUnmatched && <div className="mt-1 text-xs text-slate-500 max-h-32 overflow-y-auto bg-slate-50 rounded-lg p-2 space-y-0.5">{data.stockMeta.unmatched.map((u, i) => <div key={i} className="truncate">{u}</div>)}</div>}
                </div>
              )}
            </div>

            {lowStock.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 mb-3">
                <div className="flex items-center justify-between gap-2 mb-2">
                  <div className="text-sm font-semibold text-amber-800 flex items-center gap-1"><AlertTriangle size={14} /> Low stock ({lowStock.length}) · at or below {threshold}</div>
                  <button onClick={() => openShare("Reorder list",
                    "*Royal Agro Services — Reorder List*\n" + prettyDate(today()) + "\n" + lowStock.length + " product(s) at or below " + threshold + ":\n\n" +
                    lowStock.map((p) => "• " + pname(p.id) + " (" + p.pack + ") — " + plain(p.qty) + " left").join("\n") +
                    "\n\n— Royal Agro Ledger")}
                    className="shrink-0 inline-flex items-center gap-1 text-xs font-semibold text-amber-800"><Share2 size={12} /> Share</button>
                </div>
                <div className="space-y-1">{lowStock.map((p) => (
                  <div key={p.id} className="flex items-center justify-between gap-2 text-sm"><span className="text-slate-700 truncate">{pname(p.id)} · {p.pack}</span><span className={"font-semibold tabular-nums shrink-0 " + (p.qty <= 0 ? "text-red-700" : "text-amber-700")}>{plain(p.qty)}</span></div>))}</div>
              </div>
            )}

            {data.stockMeta && (
              <>
                <div className="relative mb-2">
                  <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input value={stockSearch} onChange={(e) => setStockSearch(e.target.value)} placeholder="Search stock" className="w-full pl-9 pr-3 py-2.5 rounded-xl border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400" />
                </div>
                <div className="bg-white rounded-xl border border-slate-200 divide-y divide-slate-100">
                  {stockList.length === 0 ? <div className="text-center text-slate-500 text-sm py-6">No products match.</div> :
                    stockList.map((p) => (
                      <button key={p.id} onClick={() => setProdView(p.id)} className="w-full text-left px-3 py-2.5 flex items-center justify-between gap-2 hover:bg-slate-50">
                        <div className="min-w-0"><div className="text-sm truncate">{pname(p.id)}</div><div className="text-xs text-slate-500">{p.pack}</div></div>
                        <span className="flex items-center gap-1.5 shrink-0">
                          <span className={"text-sm font-semibold tabular-nums " + (p.qty == null ? "text-slate-300" : p.qty <= 0 ? "text-red-700" : p.qty <= threshold ? "text-amber-700" : "text-slate-700")}>{p.qty == null ? "—" : plain(p.qty)}</span>
                          <Pencil size={13} className="text-slate-300" />
                        </span>
                      </button>))}
                </div>
              </>
            )}
          </>
        )}
      </div>
      <div className="h-8" />

      {/* ---------------------------- Dealer form modal ---------------------------- */}
      {formOpen && (
        <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center">
          <div className="absolute inset-0 bg-black opacity-40" onClick={() => setFormOpen(false)}></div>
          <div className="relative bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-4 max-h-screen overflow-y-auto">
            <div className="flex items-center justify-between mb-3"><div className="font-semibold flex items-center gap-1.5"><Users size={17} /> {editId ? "Edit dealer" : "New dealer"}</div><button onClick={() => setFormOpen(false)} className="p-1 text-slate-400 hover:text-slate-700"><X size={18} /></button></div>
            <div className="space-y-2">
              <div><label className="text-xs text-slate-500">Dealer / shop name</label>
                <input value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
              <div className="flex gap-2">
                <div className="flex-1"><label className="text-xs text-slate-500">Area</label>
                  <input value={f.area} onChange={(e) => setF({ ...f, area: e.target.value })} list="areas-list" placeholder="e.g. Multan" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                  <datalist id="areas-list">{areas.map((a) => <option key={a} value={a} />)}</datalist></div>
                <div className="flex-1"><label className="text-xs text-slate-500">Phone</label>
                  <input value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} inputMode="tel" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
              </div>
              <div><label className="text-xs text-slate-500">Salesperson</label>
                <input value={f.rep} onChange={(e) => setF({ ...f, rep: e.target.value })} list="reps-list" placeholder="Field rep name" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                <datalist id="reps-list">{reps.map((r) => <option key={r} value={r} />)}</datalist></div>
              <div><label className="text-xs text-slate-500">Remarks</label>
                <textarea value={f.remarks} onChange={(e) => setF({ ...f, remarks: e.target.value })} rows={2} placeholder="Any notes about this dealer" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 resize-none" /></div>
              <div><label className="text-xs text-slate-500">Next follow-up (optional)</label>
                <input type="date" value={f.followUp} onChange={(e) => setF({ ...f, followUp: e.target.value })} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                <p className="text-xs text-slate-400 mt-1">Shows this dealer under Follow-ups on Overview once due.</p></div>
              <div className="flex gap-2 pt-1">
                <button onClick={saveForm} disabled={!f.name.trim()} className="flex-1 bg-emerald-600 disabled:bg-slate-300 text-white py-2.5 rounded-lg text-sm font-semibold">{editId ? "Save changes" : "Add dealer"}</button>
                <button onClick={() => setFormOpen(false)} className="px-4 py-2.5 rounded-lg text-sm border border-slate-200">Cancel</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------- Settings modal ---------------------------- */}
      {showSettings && (
        <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center">
          <div className="absolute inset-0 bg-black opacity-40" onClick={() => setShowSettings(false)}></div>
          <div className="relative bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-4 max-h-screen overflow-y-auto">
            <div className="flex items-center justify-between mb-3"><div className="font-semibold flex items-center gap-1.5"><SettingsIcon size={17} /> Settings & data</div><button onClick={() => setShowSettings(false)} className="p-1 text-slate-400 hover:text-slate-700"><X size={18} /></button></div>
            <div className="space-y-4">
              <div>
                <div className="text-sm font-semibold mb-1">Low-stock threshold</div>
                <div className="flex items-center gap-2"><input type="number" value={threshold} onChange={(e) => setThreshold(e.target.value)} className="w-24 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 tabular-nums" /><span className="text-xs text-slate-500">flag products at or below this quantity</span></div>
              </div>
              <div>
                <div className="text-sm font-semibold mb-1">Product costing</div>
                <p className="text-xs text-slate-500 mb-2">Edit cost per Ltr/Kg and packing — cost-with-packing recalculates automatically and is used only to flag below-cost pricing. These figures never appear in ledgers, orders, statements, or CSV exports.</p>
                <button onClick={() => { setShowSettings(false); setShowCosting(true); }} className="inline-flex items-center gap-1 text-sm bg-slate-800 text-white rounded-lg px-3 py-2 hover:bg-slate-900"><Calculator size={15} /> Edit costing</button>
              </div>
              <div>
                <div className="text-sm font-semibold mb-2">Backup & export</div>
                <div className="flex flex-wrap gap-2">
                  <button onClick={exportAll} className="inline-flex items-center gap-1 text-sm border border-slate-200 rounded-lg px-3 py-2 hover:bg-slate-50"><Download size={15} /> Export all (CSV)</button>
                  <button onClick={backup} className="inline-flex items-center gap-1 text-sm border border-slate-200 rounded-lg px-3 py-2 hover:bg-slate-50"><Download size={15} /> Backup (JSON)</button>
                  <label className="inline-flex items-center gap-1 text-sm border border-slate-200 rounded-lg px-3 py-2 hover:bg-slate-50 cursor-pointer"><Upload size={15} /> Restore<input type="file" accept="application/json,.json" onChange={onRestore} className="hidden" /></label>
                </div>
                <p className="text-xs text-slate-500 mt-2">{hasStore ? "Records are saved on this device/account. Back up regularly." : "Storage isn't available here — records last only for this session. Back up before closing."}</p>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------- Costing editor ---------------------------- */}
      {showCosting && (
        <div className="fixed inset-0 z-40 flex flex-col bg-slate-50">
          <div className="bg-slate-800 text-white sticky top-0 shadow">
            <div className="max-w-2xl mx-auto px-4 py-3 flex items-center gap-3">
              <button onClick={() => setShowCosting(false)} className="-ml-1 p-1.5 rounded-lg hover:bg-slate-700"><ArrowLeft size={20} /></button>
              <div className="flex-1 min-w-0"><div className="font-bold leading-tight flex items-center gap-1.5"><Calculator size={18} /> Product costing</div><div className="text-slate-300 text-xs">Private · edit cost/Ltr·Kg → cost-with-packing updates</div></div>
            </div>
          </div>
          <div className="flex-1 overflow-y-auto">
            <div className="max-w-2xl mx-auto px-4 py-4">
              <div className="relative mb-3">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input value={costSearch} onChange={(e) => setCostSearch(e.target.value)} placeholder="Search product" className="w-full pl-9 pr-3 py-2.5 rounded-xl border border-slate-200 bg-white text-sm outline-none focus:border-emerald-400" />
              </div>
              <div className="space-y-3">
                {costGroupsFiltered.map((g) => (
                  <div key={g.name} className="bg-white rounded-xl border border-slate-200 p-3">
                    <div className="flex items-start justify-between gap-2 mb-2">
                      <div className="min-w-0"><div className="text-sm font-semibold leading-snug">{g.name}</div><div className="text-xs text-slate-400">{g.cat}</div></div>
                      <div className="shrink-0 text-right">
                        <label className="text-xs text-slate-500 block">Cost / Ltr·Kg</label>
                        <input value={groupCpu(g.skus)} onChange={(e) => setCpu(g, e.target.value)} inputMode="decimal" placeholder="—" className="w-24 mt-0.5 px-2 py-1.5 rounded-lg border border-slate-300 text-sm text-right outline-none focus:border-emerald-500 tabular-nums" />
                      </div>
                    </div>
                    <div className="divide-y divide-slate-100">
                      {g.skus.map((p) => {
                        const c = data.costing[p.id]; const packing = c && c[1] != null ? c[1] : "";
                        const cwp = costWith(p.id);
                        const margin = cwp != null && p.dealer ? Math.round((p.dealer - cwp) / p.dealer * 100) : null;
                        const mc = margin == null ? "text-slate-400" : margin < 0 ? "text-red-700" : margin < 12 ? "text-amber-700" : "text-emerald-700";
                        return (
                          <div key={p.id} className="py-2 flex items-center gap-2 text-sm flex-wrap">
                            <div className="w-16 text-xs text-slate-600 shrink-0">{p.pack}</div>
                            <div className="flex items-center gap-1">
                              <span className="text-xs text-slate-400">pack</span>
                              <input value={packing} onChange={(e) => setPacking(p.id, e.target.value)} inputMode="decimal" placeholder="0" className="w-16 px-2 py-1 rounded border border-slate-200 text-xs text-right outline-none focus:border-emerald-400 tabular-nums" />
                            </div>
                            <div className="text-xs text-slate-500">= cost <span className="font-semibold text-slate-700 tabular-nums">{cwp == null ? "—" : plain(cwp)}</span></div>
                            <div className="ml-auto text-xs text-right">
                              <div className="text-slate-400">sells {plain(p.dealer)}</div>
                              <div className={"font-semibold tabular-nums " + mc}>{margin == null ? "set cost" : margin + "% margin"}</div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
                {costGroupsFiltered.length === 0 && <div className="text-center text-slate-500 text-sm py-8">No products match.</div>}
              </div>
              <div className="h-8" />
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------- Confirm dialog ---------------------------- */}
      {confirmAsk && (
        <div className="fixed inset-0 z-50 flex items-center justify-center">
          <div className="absolute inset-0 bg-black opacity-40" onClick={() => setConfirmAsk(null)}></div>
          <div className="relative bg-white w-full max-w-xs mx-4 rounded-2xl p-4">
            <div className="text-sm text-slate-700">{confirmAsk.msg}</div>
            <div className="flex gap-2 mt-4">
              <button onClick={() => { const a = confirmAsk.action; setConfirmAsk(null); if (a) a(); }} className="flex-1 bg-red-600 hover:bg-red-700 text-white py-2 rounded-lg text-sm font-semibold">Delete</button>
              <button onClick={() => setConfirmAsk(null)} className="flex-1 border border-slate-200 py-2 rounded-lg text-sm">Cancel</button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------- Edit payment modal ---------------------------- */}
      {payEdit && (
        <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center">
          <div className="absolute inset-0 bg-black opacity-40" onClick={() => setPayEdit(null)}></div>
          <div className="relative bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-4 max-h-screen overflow-y-auto">
            <div className="flex items-center justify-between mb-3"><div className="font-semibold flex items-center gap-1.5"><Wallet size={17} /> Edit payment</div><button onClick={() => setPayEdit(null)} className="p-1 text-slate-400 hover:text-slate-700"><X size={18} /></button></div>
            <div className="space-y-2">
              <div><label className="text-xs text-slate-500">Amount (Rs)</label>
                <input value={payEdit.amount} onChange={(e) => setPayEdit({ ...payEdit, amount: e.target.value.replace(/[^\d.]/g, "") })} inputMode="decimal" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-lg font-semibold outline-none focus:border-emerald-400 tabular-nums" /></div>
              <div className="flex gap-2">
                <div className="flex-1"><label className="text-xs text-slate-500">Date</label>
                  <input type="date" value={payEdit.date} onChange={(e) => setPayEdit({ ...payEdit, date: e.target.value })} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
                <div className="flex-1"><label className="text-xs text-slate-500">Method</label>
                  <select value={payEdit.method} onChange={(e) => setPayEdit({ ...payEdit, method: e.target.value })} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 bg-white">
                    <option>Cash</option><option>Bank Transfer</option><option>Cheque</option><option>Easypaisa/JazzCash</option><option>Other</option></select></div>
              </div>
              <div><label className="text-xs text-slate-500">Note</label>
                <input value={payEdit.note} onChange={(e) => setPayEdit({ ...payEdit, note: e.target.value })} className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" /></div>
              <div className="flex gap-2 pt-1">
                <button onClick={savePayEdit} disabled={!parseFloat(payEdit.amount)} className="flex-1 bg-emerald-600 disabled:bg-slate-300 text-white py-2.5 rounded-lg text-sm font-semibold">Save changes</button>
                <button onClick={() => setPayEdit(null)} className="px-4 py-2.5 rounded-lg text-sm border border-slate-200">Cancel</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------- Share modal ---------------------------- */}
      {shareModal && (
        <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center">
          <div className="absolute inset-0 bg-black opacity-40" onClick={() => setShareModal(null)}></div>
          <div className="relative bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-4 max-h-screen overflow-y-auto">
            <div className="flex items-center justify-between mb-3">
              <div className="font-semibold flex items-center gap-1.5 truncate"><Share2 size={17} className="shrink-0" /> {shareModal.title}</div>
              <button onClick={() => setShareModal(null)} className="p-1 text-slate-400 hover:text-slate-700 shrink-0"><X size={18} /></button>
            </div>
            <textarea readOnly value={shareModal.text} rows={10} onClick={(e) => e.target.select()} className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs outline-none resize-none bg-slate-50 font-mono" />
            <div className="flex gap-2 mt-3">
              <button onClick={doNativeShare} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg text-sm font-semibold inline-flex items-center justify-center gap-1.5"><Share2 size={15} /> Share</button>
              <button onClick={doCopyShare} className="px-4 py-2.5 rounded-lg text-sm border border-slate-200 font-semibold">Copy</button>
            </div>
            <p className="text-xs text-slate-400 mt-2">Share opens WhatsApp or any app on your phone. Copy pastes it yourself.</p>
          </div>
        </div>
      )}

      {/* ---------------------------- Product detail ---------------------------- */}
      {prodView && CAT_BY_ID[prodView] && (() => {
        const p = CAT_BY_ID[prodView];
        const cwp = costWith(prodView);
        const margin = cwp != null && p.dealer ? Math.round((p.dealer - cwp) / p.dealer * 100) : null;
        const c = data.costing[prodView] || [null, 0];
        const qty = data.stock[prodView];
        return (
          <div className="fixed inset-0 z-40 flex flex-col bg-slate-50">
            <div className="bg-emerald-700 text-white sticky top-0 shadow">
              <div className="max-w-2xl mx-auto px-4 py-3 flex items-center gap-3">
                <button onClick={() => setProdView(null)} className="-ml-1 p-1.5 rounded-lg hover:bg-emerald-600"><ArrowLeft size={20} /></button>
                <div className="flex-1 min-w-0"><div className="font-bold leading-tight truncate">{pname(prodView)}</div><div className="text-emerald-100 text-xs">{p.cat} · {p.pack}</div></div>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto">
              <div className="max-w-2xl mx-auto px-4 py-4 space-y-3">
                <div className="bg-white rounded-xl border border-slate-200 p-3">
                  <label className="text-xs text-slate-500">Display name</label>
                  <div className="flex items-center gap-2 mt-1">
                    <input value={pname(prodView)} onChange={(e) => setName(prodView, e.target.value)} className="flex-1 px-3 py-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400" />
                    {data.names && data.names[prodView] && <button onClick={() => setName(prodView, "")} title="Reset to catalog name" className="text-slate-400 hover:text-emerald-700"><RotateCcw size={16} /></button>}
                  </div>
                  <div className="text-xs text-slate-400 mt-1">Catalog name: {p.name}</div>
                </div>

                <div className="bg-white rounded-xl border border-slate-200 p-3">
                  <label className="text-xs text-slate-500">Stock on hand ({p.pack})</label>
                  <input value={qty == null ? "" : qty} onChange={(e) => setStockVal(prodView, e.target.value.replace(/[^\d]/g, ""))} inputMode="numeric" placeholder="not set" className="w-full mt-1 px-3 py-2.5 rounded-lg border border-slate-200 text-lg font-semibold outline-none focus:border-emerald-400 tabular-nums" />
                  <div className="text-xs text-slate-400 mt-1">Low-stock alert triggers at or below {threshold}.</div>
                </div>

                <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-3">
                  <div className="text-sm font-semibold flex items-center gap-1"><Calculator size={15} /> Costing <span className="text-xs font-normal text-slate-400">(private)</span></div>
                  <div className="flex gap-2">
                    <div className="flex-1">
                      <label className="text-xs text-slate-500">Cost / Ltr·Kg</label>
                      <input value={c[0] == null ? "" : c[0]} onChange={(e) => setGroupCpu(prodView, e.target.value)} inputMode="decimal" placeholder="—" className="w-full mt-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 tabular-nums" />
                    </div>
                    <div className="flex-1">
                      <label className="text-xs text-slate-500">Packing cost</label>
                      <input value={c[1] == null ? "" : c[1]} onChange={(e) => setCostField(prodView, 1, e.target.value)} inputMode="decimal" placeholder="0" className="w-full mt-1 px-3 py-2 rounded-lg border border-slate-200 text-sm outline-none focus:border-emerald-400 tabular-nums" />
                    </div>
                  </div>
                  <div className="grid grid-cols-3 gap-2 text-center">
                    <div className="bg-slate-50 rounded-lg py-2"><div className="text-xs text-slate-400">Cost</div><div className="text-sm font-bold tabular-nums">{cwp == null ? "—" : plain(cwp)}</div></div>
                    <div className="bg-slate-50 rounded-lg py-2"><div className="text-xs text-slate-400">Sells</div><div className="text-sm font-bold tabular-nums">{plain(p.dealer)}</div></div>
                    <div className="bg-slate-50 rounded-lg py-2"><div className="text-xs text-slate-400">Margin</div><div className={"text-sm font-bold tabular-nums " + (margin == null ? "text-slate-400" : margin < 0 ? "text-red-700" : margin < 12 ? "text-amber-700" : "text-emerald-700")}>{margin == null ? "—" : margin + "%"}</div></div>
                  </div>
                  <div className="text-xs text-slate-400">Cost / Ltr·Kg is shared across all pack sizes of this product. Invoice rate before discount: {plain(p.rate)}.</div>
                </div>
              </div>
              <div className="h-8" />
            </div>
          </div>
        );
      })()}
    </div>
  );
}
