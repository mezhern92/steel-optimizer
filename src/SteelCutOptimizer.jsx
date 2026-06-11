import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

/* ============================================================================
   STEEL OPTIMIZER — combined Plates + Sections (Stage 1)
   Theme: amber / dark engineering.  Imports only xlsx.
============================================================================ */

const DENSITY = 7850; // kg/m3
const PART_COLORS = ["#F59E0B","#3B82F6","#10B981","#EF4444","#8B5CF6","#06B6D4","#F97316","#84CC16","#EC4899","#14B8A6","#6366F1","#D97706","#059669","#DC2626","#7C3AED","#0EA5E9","#65A30D","#DB2777"];
const fmtTon = (kg) => (kg / 1000).toFixed(kg / 1000 >= 100 ? 0 : 2);
const fmtMm = (mm) => (mm >= 1000 ? `${(mm / 1000).toFixed(3).replace(/\.?0+$/, "")}m` : `${Math.round(mm)}mm`);
const fmtKg = (kg) => (kg >= 1000 ? `${(kg / 1000).toFixed(2)} t` : `${kg.toFixed(1)} kg`);
const genId = () => Math.random().toString(36).slice(2, 8).toUpperCase();

/* ─── SECTION DATABASE (expanded in Stage 2) ─────────────────────────────── */
const STEEL_DB = [
  ["IPE 80","IPE",6.0],["IPE 100","IPE",8.1],["IPE 120","IPE",10.4],["IPE 140","IPE",12.9],["IPE 160","IPE",15.8],["IPE 180","IPE",18.8],["IPE 200","IPE",22.4],["IPE 220","IPE",26.2],["IPE 240","IPE",30.7],["IPE 270","IPE",36.1],["IPE 300","IPE",42.2],["IPE 330","IPE",49.1],["IPE 360","IPE",57.1],["IPE 400","IPE",66.3],["IPE 450","IPE",77.6],["IPE 500","IPE",90.7],["IPE 550","IPE",106.0],["IPE 600","IPE",122.0],
  ["HEA 100","HEA",16.7],["HEA 120","HEA",19.9],["HEA 140","HEA",24.7],["HEA 160","HEA",30.4],["HEA 180","HEA",35.5],["HEA 200","HEA",42.3],["HEA 220","HEA",50.5],["HEA 240","HEA",60.3],["HEA 260","HEA",68.2],["HEA 280","HEA",76.4],["HEA 300","HEA",88.3],["HEA 320","HEA",97.6],["HEA 340","HEA",105.0],["HEA 360","HEA",112.0],["HEA 400","HEA",125.0],["HEA 450","HEA",140.0],["HEA 500","HEA",155.0],["HEA 550","HEA",166.0],["HEA 600","HEA",178.0],["HEA 650","HEA",190.0],["HEA 700","HEA",204.0],["HEA 800","HEA",224.0],["HEA 900","HEA",252.0],["HEA 1000","HEA",272.0],
  ["HEB 100","HEB",20.4],["HEB 120","HEB",26.7],["HEB 140","HEB",33.7],["HEB 160","HEB",42.6],["HEB 180","HEB",51.2],["HEB 200","HEB",61.3],["HEB 220","HEB",71.5],["HEB 240","HEB",83.2],["HEB 260","HEB",93.0],["HEB 280","HEB",103.0],["HEB 300","HEB",117.0],["HEB 320","HEB",127.0],["HEB 340","HEB",134.0],["HEB 360","HEB",142.0],["HEB 400","HEB",155.0],["HEB 450","HEB",171.0],["HEB 500","HEB",187.0],["HEB 550","HEB",199.0],["HEB 600","HEB",212.0],["HEB 650","HEB",225.0],["HEB 700","HEB",241.0],["HEB 800","HEB",262.0],["HEB 900","HEB",291.0],["HEB 1000","HEB",314.0],
  ["127x76x13 UB","UB",13.0],["152x89x16 UB","UB",16.0],["178x102x19 UB","UB",19.0],["203x102x23 UB","UB",23.1],["203x133x25 UB","UB",25.1],["203x133x30 UB","UB",30.0],["254x102x22 UB","UB",22.0],["254x102x25 UB","UB",25.2],["254x102x28 UB","UB",28.3],["254x146x31 UB","UB",31.1],["254x146x37 UB","UB",37.0],["254x146x43 UB","UB",43.0],["305x102x25 UB","UB",24.8],["305x102x28 UB","UB",28.2],["305x102x33 UB","UB",32.8],["305x127x37 UB","UB",37.0],["305x127x42 UB","UB",41.9],["305x127x48 UB","UB",48.1],["305x165x40 UB","UB",40.3],["305x165x46 UB","UB",46.1],["305x165x54 UB","UB",54.0],["356x127x33 UB","UB",33.1],["356x127x39 UB","UB",39.1],["356x171x45 UB","UB",45.0],["356x171x51 UB","UB",51.0],["356x171x57 UB","UB",57.0],["356x171x67 UB","UB",67.1],["406x140x39 UB","UB",39.0],["406x140x46 UB","UB",46.0],["406x178x54 UB","UB",54.1],["406x178x60 UB","UB",60.1],["406x178x67 UB","UB",67.1],["406x178x74 UB","UB",74.2],["457x152x52 UB","UB",52.3],["457x152x60 UB","UB",59.8],["457x152x67 UB","UB",67.2],["457x152x74 UB","UB",74.2],["457x152x82 UB","UB",82.1],["457x191x67 UB","UB",67.1],["457x191x74 UB","UB",74.3],["457x191x82 UB","UB",82.0],["457x191x89 UB","UB",89.3],["457x191x98 UB","UB",98.3],["533x210x82 UB","UB",82.2],["533x210x92 UB","UB",92.1],["533x210x101 UB","UB",101.0],["533x210x109 UB","UB",109.0],["533x210x122 UB","UB",122.0],["610x229x101 UB","UB",101.0],["610x229x113 UB","UB",113.0],["610x229x125 UB","UB",125.0],["610x229x140 UB","UB",140.0],["610x305x149 UB","UB",149.0],["610x305x179 UB","UB",179.0],["610x305x238 UB","UB",238.0],["686x254x125 UB","UB",125.0],["686x254x140 UB","UB",140.0],["686x254x152 UB","UB",152.0],["686x254x170 UB","UB",170.0],["762x267x134 UB","UB",134.0],["762x267x147 UB","UB",147.0],["762x267x173 UB","UB",173.0],["762x267x197 UB","UB",197.0],["838x292x176 UB","UB",176.0],["838x292x194 UB","UB",194.0],["838x292x226 UB","UB",226.0],["914x305x201 UB","UB",201.0],["914x305x224 UB","UB",224.0],["914x305x253 UB","UB",253.0],["914x305x289 UB","UB",289.0],
  ["152x152x23 UC","UC",23.0],["152x152x30 UC","UC",30.0],["152x152x37 UC","UC",37.0],["152x152x44 UC","UC",44.0],["203x203x46 UC","UC",46.1],["203x203x52 UC","UC",52.0],["203x203x60 UC","UC",60.0],["203x203x71 UC","UC",71.0],["203x203x86 UC","UC",86.1],["203x203x100 UC","UC",100.0],["254x254x73 UC","UC",73.1],["254x254x89 UC","UC",88.9],["254x254x107 UC","UC",107.0],["254x254x132 UC","UC",132.0],["254x254x167 UC","UC",167.0],["305x305x97 UC","UC",96.9],["305x305x118 UC","UC",117.0],["305x305x137 UC","UC",137.0],["305x305x158 UC","UC",158.0],["305x305x198 UC","UC",198.0],["305x305x240 UC","UC",240.0],["305x305x283 UC","UC",283.0],["356x368x129 UC","UC",129.0],["356x368x153 UC","UC",153.0],["356x368x177 UC","UC",177.0],["356x368x202 UC","UC",202.0],["356x406x235 UC","UC",235.0],["356x406x287 UC","UC",287.0],["356x406x340 UC","UC",340.0],["356x406x393 UC","UC",393.0],["356x406x467 UC","UC",467.0],["356x406x551 UC","UC",551.0],["356x406x634 UC","UC",634.0],
  ["HW 100x100","JIS-HW",17.2],["HW 125x125","JIS-HW",23.8],["HW 150x150","JIS-HW",31.5],["HW 175x175","JIS-HW",40.4],["HW 200x200","JIS-HW",49.9],["HW 250x250","JIS-HW",72.4],["HW 300x300","JIS-HW",94.0],["HW 350x350","JIS-HW",137.0],["HW 400x400","JIS-HW",172.0],
  ["HM 150x100","JIS-HM",21.1],["HM 200x150","JIS-HM",30.6],["HM 250x175","JIS-HM",44.1],["HM 300x200","JIS-HM",56.8],["HM 350x250","JIS-HM",79.7],["HM 400x300","JIS-HM",107.0],["HM 450x300","JIS-HM",124.0],["HM 500x300","JIS-HM",128.0],
  ["HN 100x50","JIS-HN",9.3],["HN 125x60","JIS-HN",13.1],["HN 150x75","JIS-HN",14.0],["HN 175x90","JIS-HN",18.0],["HN 200x100","JIS-HN",20.9],["HN 250x125","JIS-HN",29.0],["HN 300x150","JIS-HN",36.7],["HN 350x175","JIS-HN",49.4],["HN 400x200","JIS-HN",65.4],["HN 450x200","JIS-HN",74.9],["HN 500x200","JIS-HN",88.5],["HN 600x200","JIS-HN",102.0],["HN 700x300","JIS-HN",166.0],["HN 800x300","JIS-HN",191.0],["HN 900x300","JIS-HN",213.0],
  ["RHS 50x30x3","RHS",3.45],["RHS 60x40x3","RHS",4.39],["RHS 80x40x4","RHS",6.71],["RHS 100x50x4","RHS",8.59],["RHS 100x60x5","RHS",11.30],["RHS 120x60x5","RHS",12.80],["RHS 120x80x5","RHS",14.40],["RHS 150x100x5","RHS",18.70],["RHS 160x80x6","RHS",20.70],["RHS 200x100x6","RHS",26.40],["RHS 200x120x8","RHS",37.30],["RHS 250x150x8","RHS",47.70],["RHS 300x200x8","RHS",60.30],
  ["SHS 40x40x3","SHS",3.41],["SHS 50x50x3","SHS",4.35],["SHS 50x50x4","SHS",5.64],["SHS 60x60x4","SHS",6.90],["SHS 70x70x4","SHS",8.13],["SHS 80x80x5","SHS",11.30],["SHS 90x90x5","SHS",12.80],["SHS 100x100x5","SHS",14.40],["SHS 100x100x6","SHS",17.00],["SHS 120x120x6","SHS",20.70],["SHS 150x150x8","SHS",34.20],["SHS 200x200x8","SHS",46.50],["SHS 250x250x10","SHS",72.70],
  ["CHS 33.7x3.2","CHS",2.41],["CHS 42.4x3.2","CHS",3.09],["CHS 48.3x3.2","CHS",3.56],["CHS 60.3x3.6","CHS",5.03],["CHS 76.1x4.0","CHS",7.11],["CHS 88.9x4.0","CHS",8.38],["CHS 114.3x5.0","CHS",13.50],["CHS 139.7x5.0","CHS",16.60],["CHS 168.3x6.0","CHS",24.00],["CHS 219.1x6.0","CHS",31.50],["CHS 273.0x8.0","CHS",52.30],
  ["C 100x50x2.0","C-COLD",3.13],["C 120x50x2.0","C-COLD",3.45],["C 140x60x2.0","C-COLD",4.08],["C 150x65x2.0","C-COLD",4.40],["C 150x65x2.5","C-COLD",5.46],["C 200x65x2.0","C-COLD",5.18],["C 200x75x2.5","C-COLD",6.87],["C 250x75x2.5","C-COLD",7.66],["C 300x90x3.0","C-COLD",11.10],
  ["Z 100x50x2.0","Z-COLD",3.13],["Z 120x50x2.0","Z-COLD",3.45],["Z 140x60x2.0","Z-COLD",4.08],["Z 150x65x2.0","Z-COLD",4.40],["Z 150x65x2.5","Z-COLD",5.46],["Z 200x65x2.5","Z-COLD",6.45],["Z 200x75x2.5","Z-COLD",6.87],["Z 250x75x3.0","Z-COLD",9.12],["Z 300x90x3.0","Z-COLD",11.10],
  // UPN (European channels, tapered flange)
  ["UPN 80","UPN",8.64],["UPN 100","UPN",10.6],["UPN 120","UPN",13.4],["UPN 140","UPN",16.0],["UPN 160","UPN",18.8],["UPN 180","UPN",22.0],["UPN 200","UPN",25.3],["UPN 220","UPN",29.4],["UPN 240","UPN",33.2],["UPN 260","UPN",37.9],["UPN 280","UPN",41.8],["UPN 300","UPN",46.2],["UPN 320","UPN",59.5],["UPN 350","UPN",60.6],["UPN 380","UPN",63.1],["UPN 400","UPN",71.8],
  // UPE (European channels, parallel flange)
  ["UPE 80","UPE",7.9],["UPE 100","UPE",9.82],["UPE 120","UPE",12.1],["UPE 140","UPE",14.5],["UPE 160","UPE",17.0],["UPE 180","UPE",19.7],["UPE 200","UPE",22.8],["UPE 220","UPE",26.6],["UPE 240","UPE",30.2],["UPE 270","UPE",35.2],["UPE 300","UPE",44.4],["UPE 330","UPE",53.2],["UPE 360","UPE",61.2],["UPE 400","UPE",72.2],
  // PFC (UK parallel flange channels)
  ["100x50x10 PFC","PFC",10.2],["125x65x15 PFC","PFC",14.8],["150x75x18 PFC","PFC",17.9],["150x90x24 PFC","PFC",23.9],["180x75x20 PFC","PFC",20.3],["180x90x26 PFC","PFC",26.1],["200x75x23 PFC","PFC",23.4],["200x90x30 PFC","PFC",29.7],["230x75x26 PFC","PFC",25.7],["230x90x32 PFC","PFC",32.2],["260x75x28 PFC","PFC",27.6],["260x90x35 PFC","PFC",34.8],["300x90x41 PFC","PFC",41.4],["300x100x46 PFC","PFC",45.5],["380x100x54 PFC","PFC",54.0],["430x100x64 PFC","PFC",64.4],
  // HEM European wide-flange (heavy)
  ["HEM 100","HEM",41.8],["HEM 120","HEM",52.1],["HEM 140","HEM",63.2],["HEM 160","HEM",76.2],["HEM 180","HEM",88.9],["HEM 200","HEM",103],["HEM 220","HEM",117],["HEM 240","HEM",157],["HEM 260","HEM",172],["HEM 280","HEM",189],["HEM 300","HEM",238],["HEM 320","HEM",245],["HEM 340","HEM",248],["HEM 360","HEM",250],["HEM 400","HEM",256],["HEM 450","HEM",263],["HEM 500","HEM",270],["HEM 550","HEM",278],["HEM 600","HEM",285],["HEM 650","HEM",293],["HEM 700","HEM",301],["HEM 800","HEM",317],["HEM 900","HEM",333],["HEM 1000","HEM",349],
  // IPN European I-beam (narrow, tapered flange)
  ["IPN 80","IPN",5.94],["IPN 100","IPN",8.34],["IPN 120","IPN",11.1],["IPN 140","IPN",14.3],["IPN 160","IPN",17.9],["IPN 180","IPN",21.9],["IPN 200","IPN",26.2],["IPN 220","IPN",31.1],["IPN 240","IPN",36.2],["IPN 260","IPN",41.9],["IPN 280","IPN",47.9],["IPN 300","IPN",54.2],["IPN 320","IPN",61.0],["IPN 340","IPN",68.0],["IPN 360","IPN",76.1],["IPN 380","IPN",84.0],["IPN 400","IPN",92.4],["IPN 450","IPN",115],["IPN 500","IPN",141],["IPN 550","IPN",166],["IPN 600","IPN",199],
  // W American wide-flange (kg/m from lb/ft)
  ["W6x9","W",13.4],["W6x12","W",17.9],["W6x15","W",22.3],["W6x16","W",23.8],["W6x20","W",29.8],["W6x25","W",37.2],["W8x10","W",14.9],["W8x13","W",19.3],["W8x15","W",22.3],["W8x18","W",26.8],["W8x21","W",31.3],["W8x24","W",35.7],
  ["W8x28","W",41.7],["W8x31","W",46.1],["W8x35","W",52.1],["W8x40","W",59.5],["W8x48","W",71.4],["W8x58","W",86.3],["W8x67","W",99.7],["W10x12","W",17.9],["W10x15","W",22.3],["W10x17","W",25.3],["W10x19","W",28.3],["W10x22","W",32.7],
  ["W10x26","W",38.7],["W10x30","W",44.6],["W10x33","W",49.1],["W10x39","W",58.0],["W10x45","W",67.0],["W10x49","W",72.9],["W10x54","W",80.4],["W10x60","W",89.3],["W10x68","W",101.2],["W10x77","W",114.6],["W10x88","W",131.0],["W10x100","W",148.8],
  ["W10x112","W",166.7],["W12x14","W",20.8],["W12x16","W",23.8],["W12x19","W",28.3],["W12x22","W",32.7],["W12x26","W",38.7],["W12x30","W",44.6],["W12x35","W",52.1],["W12x40","W",59.5],["W12x45","W",67.0],["W12x50","W",74.4],["W12x53","W",78.9],
  ["W12x58","W",86.3],["W12x65","W",96.7],["W12x72","W",107.1],["W12x79","W",117.6],["W12x87","W",129.5],["W12x96","W",142.9],["W12x106","W",157.7],["W12x120","W",178.6],["W12x136","W",202.4],["W12x152","W",226.2],["W12x170","W",253.0],["W12x190","W",282.8],
  ["W12x210","W",312.5],["W12x230","W",342.3],["W12x252","W",375.0],["W12x279","W",415.2],["W12x305","W",453.9],["W12x336","W",500.0],["W14x22","W",32.7],["W14x26","W",38.7],["W14x30","W",44.6],["W14x34","W",50.6],["W14x38","W",56.6],["W14x43","W",64.0],
  ["W14x48","W",71.4],["W14x53","W",78.9],["W14x61","W",90.8],["W14x68","W",101.2],["W14x74","W",110.1],["W14x82","W",122.0],["W14x90","W",133.9],["W14x99","W",147.3],["W14x109","W",162.2],["W14x120","W",178.6],["W14x132","W",196.4],["W14x145","W",215.8],
  ["W14x159","W",236.6],["W14x176","W",261.9],["W14x193","W",287.2],["W14x211","W",314.0],["W14x233","W",346.7],["W14x257","W",382.5],["W14x283","W",421.2],["W14x311","W",462.8],["W14x342","W",509.0],["W14x370","W",550.6],["W14x398","W",592.3],["W14x426","W",634.0],
  ["W14x455","W",677.1],["W14x500","W",744.1],["W14x550","W",818.5],["W14x605","W",900.3],["W14x665","W",989.6],["W14x730","W",1086.4],["W16x26","W",38.7],["W16x31","W",46.1],["W16x36","W",53.6],["W16x40","W",59.5],["W16x45","W",67.0],["W16x50","W",74.4],
  ["W16x57","W",84.8],["W16x67","W",99.7],["W16x77","W",114.6],["W16x89","W",132.4],["W16x100","W",148.8],["W18x35","W",52.1],["W18x40","W",59.5],["W18x46","W",68.5],["W18x50","W",74.4],["W18x55","W",81.8],["W18x60","W",89.3],["W18x65","W",96.7],
  ["W18x71","W",105.7],["W18x76","W",113.1],["W18x86","W",128.0],["W18x97","W",144.4],["W18x106","W",157.7],["W18x119","W",177.1],["W18x130","W",193.5],["W18x143","W",212.8],["W18x158","W",235.1],["W21x44","W",65.5],["W21x50","W",74.4],["W21x57","W",84.8],
  ["W21x62","W",92.3],["W21x68","W",101.2],["W21x73","W",108.6],["W21x83","W",123.5],["W21x93","W",138.4],["W21x101","W",150.3],["W21x111","W",165.2],["W21x122","W",181.6],["W21x132","W",196.4],["W21x147","W",218.8],["W24x55","W",81.8],["W24x62","W",92.3],
  ["W24x68","W",101.2],["W24x76","W",113.1],["W24x84","W",125.0],["W24x94","W",139.9],["W24x103","W",153.3],["W24x104","W",154.8],["W24x117","W",174.1],["W24x131","W",194.9],["W24x146","W",217.3],["W24x162","W",241.1],["W27x84","W",125.0],["W27x94","W",139.9],
  ["W27x102","W",151.8],["W27x114","W",169.7],["W27x129","W",192.0],["W27x146","W",217.3],["W27x161","W",239.6],["W27x178","W",264.9],["W30x90","W",133.9],["W30x99","W",147.3],["W30x108","W",160.7],["W30x116","W",172.6],["W30x124","W",184.5],["W30x132","W",196.4],
  ["W30x148","W",220.2],["W30x173","W",257.5],["W30x191","W",284.2],["W30x211","W",314.0],["W33x118","W",175.6],["W33x130","W",193.5],["W33x141","W",209.8],["W33x152","W",226.2],["W33x169","W",251.5],["W33x201","W",299.1],["W33x221","W",328.9],["W33x241","W",358.6],
  ["W36x135","W",200.9],["W36x150","W",223.2],["W36x160","W",238.1],["W36x170","W",253.0],["W36x182","W",270.8],["W36x194","W",288.7],["W36x210","W",312.5],["W36x231","W",343.8],["W36x232","W",345.3],["W36x247","W",367.6],["W36x256","W",381.0],["W36x262","W",389.9],
  ["W36x282","W",419.7],["W36x302","W",449.4],["W36x361","W",537.2],
  // C American standard channel (kg/m from lb/ft)
  ["C3x4.1","C-AMER",6.1],["C3x5","C-AMER",7.4],["C3x6","C-AMER",8.9],["C4x5.4","C-AMER",8.0],["C4x7.25","C-AMER",10.8],["C5x6.7","C-AMER",10.0],["C5x9","C-AMER",13.4],["C6x8.2","C-AMER",12.2],["C6x10.5","C-AMER",15.6],["C6x13","C-AMER",19.3],["C7x9.8","C-AMER",14.6],["C7x12.25","C-AMER",18.2],
  ["C7x14.75","C-AMER",22.0],["C8x11.5","C-AMER",17.1],["C8x13.75","C-AMER",20.5],["C8x18.75","C-AMER",27.9],["C9x13.4","C-AMER",19.9],["C9x15","C-AMER",22.3],["C9x20","C-AMER",29.8],["C10x15.3","C-AMER",22.8],["C10x20","C-AMER",29.8],["C10x25","C-AMER",37.2],["C10x30","C-AMER",44.6],["C12x20.7","C-AMER",30.8],
  ["C12x25","C-AMER",37.2],["C12x30","C-AMER",44.6],["C15x33.9","C-AMER",50.4],["C15x40","C-AMER",59.5],["C15x50","C-AMER",74.4],
  // MC American miscellaneous channel
  ["MC6x12","MC",17.9],["MC6x15.1","MC",22.5],["MC8x18.7","MC",27.8],["MC8x21.4","MC",31.8],["MC10x22","MC",32.7],["MC10x25","MC",37.2],["MC12x31","MC",46.1],["MC12x35","MC",52.1],["MC18x42.7","MC",63.5],["MC18x58","MC",86.3],
  // UBP UK bearing piles
  ["203x203x45 UBP","UBP",45],["254x254x63 UBP","UBP",63.4],["254x254x71 UBP","UBP",71],["254x254x85 UBP","UBP",85.1],["305x305x79 UBP","UBP",78.9],["305x305x88 UBP","UBP",88],["305x305x95 UBP","UBP",94.9],["305x305x110 UBP","UBP",110],["305x305x126 UBP","UBP",126],["305x305x149 UBP","UBP",149],["305x305x186 UBP","UBP",186],["305x305x223 UBP","UBP",223],["356x368x109 UBP","UBP",109],["356x368x133 UBP","UBP",133],["356x368x152 UBP","UBP",152],["356x368x174 UBP","UBP",174],
  // Equal angles (metric)
  ["L 20x20x3","L",0.87],["L 25x25x3","L",1.11],["L 25x25x4","L",1.44],["L 30x30x3","L",1.34],["L 30x30x4","L",1.76],["L 40x40x4","L",2.39],["L 40x40x5","L",2.94],["L 45x45x5","L",3.34],["L 50x50x5","L",3.73],["L 50x50x6","L",4.43],
  ["L 60x60x6","L",5.37],["L 60x60x8","L",7.03],["L 65x65x6","L",5.84],["L 70x70x6","L",6.31],["L 70x70x7","L",7.31],["L 75x75x6","L",6.78],["L 75x75x8","L",8.92],["L 80x80x8","L",9.55],["L 80x80x10","L",11.78],["L 90x90x8","L",10.8],
  ["L 90x90x10","L",13.34],["L 100x100x10","L",14.91],["L 100x100x12","L",17.71],["L 120x120x10","L",18.05],["L 120x120x12","L",21.48],["L 150x150x12","L",27.13],["L 150x150x15","L",33.56],["L 200x200x16","L",48.23],["L 200x200x20","L",59.66],["L 200x200x24","L",70.84],
  // Unequal angles (metric)
  ["L 75x50x6","L",5.6],["L 75x50x8","L",7.35],["L 100x65x7","L",8.68],["L 100x65x10","L",12.17],["L 100x75x8","L",10.49],["L 100x75x10","L",12.95],["L 125x75x8","L",12.06],["L 125x75x10","L",14.91],["L 150x75x9","L",15.26],
  ["L 150x75x10","L",16.88],["L 150x90x10","L",18.05],["L 150x90x12","L",21.48],["L 150x100x10","L",18.84],["L 150x100x12","L",22.42],["L 200x100x10","L",22.76],["L 200x100x12","L",27.13],["L 200x150x12","L",31.84],["L 200x150x15","L",39.45],
  // RHS rectangular hollow (expanded)
  ["RHS 40x20x2","RHS",1.65],["RHS 40x20x3","RHS",2.3],["RHS 50x25x3","RHS",3.01],["RHS 50x30x3","RHS",3.24],["RHS 60x40x3","RHS",4.18],["RHS 60x40x4","RHS",5.35],["RHS 70x50x3","RHS",5.13],["RHS 80x40x3","RHS",5.13],["RHS 80x40x4","RHS",6.6],
  ["RHS 80x60x4","RHS",7.86],["RHS 90x50x3","RHS",6.07],["RHS 100x40x4","RHS",7.86],["RHS 100x50x3","RHS",6.54],["RHS 100x50x4","RHS",8.49],["RHS 100x50x5","RHS",10.32],["RHS 100x60x4","RHS",9.11],["RHS 100x60x5","RHS",11.1],["RHS 120x60x4","RHS",10.37],
  ["RHS 120x60x5","RHS",12.67],["RHS 120x80x5","RHS",14.24],["RHS 120x80x6","RHS",16.74],["RHS 140x80x5","RHS",15.81],["RHS 150x100x5","RHS",18.17],["RHS 150x100x6","RHS",21.45],["RHS 160x80x6","RHS",20.51],["RHS 200x100x5","RHS",22.09],["RHS 200x100x6","RHS",26.16],
  ["RHS 200x100x8","RHS",33.95],["RHS 200x120x8","RHS",36.46],["RHS 250x150x6","RHS",35.58],["RHS 250x150x8","RHS",46.51],["RHS 250x150x10","RHS",56.96],["RHS 300x200x8","RHS",59.07],["RHS 300x200x10","RHS",72.66],["RHS 400x200x10","RHS",88.36],["RHS 400x200x12","RHS",104.64],
  ["RHS 400x300x12","RHS",123.48],["RHS 500x300x12","RHS",142.32],["RHS 500x300x16","RHS",186.02],
  // SHS square hollow (expanded)
  ["SHS 20x20x2","SHS",1.02],["SHS 25x25x2","SHS",1.34],["SHS 25x25x3","SHS",1.83],["SHS 30x30x2","SHS",1.65],["SHS 30x30x3","SHS",2.3],["SHS 40x40x2","SHS",2.28],["SHS 40x40x3","SHS",3.24],["SHS 40x40x4","SHS",4.09],["SHS 50x50x3","SHS",4.18],
  ["SHS 50x50x4","SHS",5.35],["SHS 50x50x5","SHS",6.39],["SHS 60x60x3","SHS",5.13],["SHS 60x60x4","SHS",6.6],["SHS 60x60x5","SHS",7.96],["SHS 70x70x4","SHS",7.86],["SHS 70x70x5","SHS",9.53],["SHS 70x70x6","SHS",11.09],["SHS 80x80x4","SHS",9.11],
  ["SHS 80x80x5","SHS",11.1],["SHS 80x80x6","SHS",12.97],["SHS 90x90x5","SHS",12.67],["SHS 90x90x6","SHS",14.86],["SHS 100x100x4","SHS",11.63],["SHS 100x100x5","SHS",14.24],["SHS 100x100x6","SHS",16.74],["SHS 100x100x8","SHS",21.39],["SHS 120x120x5","SHS",17.38],
  ["SHS 120x120x6","SHS",20.51],["SHS 120x120x8","SHS",26.41],["SHS 140x140x6","SHS",24.28],["SHS 140x140x8","SHS",31.43],["SHS 150x150x6","SHS",26.16],["SHS 150x150x8","SHS",33.95],["SHS 150x150x10","SHS",41.26],["SHS 160x160x8","SHS",36.46],["SHS 180x180x8","SHS",41.48],
  ["SHS 200x200x8","SHS",46.51],["SHS 200x200x10","SHS",56.96],["SHS 200x200x12","SHS",66.96],["SHS 250x250x8","SHS",59.07],["SHS 250x250x10","SHS",72.66],["SHS 250x250x12","SHS",85.8],["SHS 300x300x10","SHS",88.36],["SHS 300x300x12","SHS",104.64],["SHS 350x350x12","SHS",123.48],
  ["SHS 400x400x12","SHS",142.32],["SHS 400x400x16","SHS",186.02],
  // CHS circular hollow (expanded)
  ["CHS 21.3x2.6","CHS",1.2],["CHS 26.9x2.6","CHS",1.56],["CHS 26.9x3.2","CHS",1.87],["CHS 33.7x2.6","CHS",1.99],["CHS 33.7x3.2","CHS",2.41],["CHS 42.4x2.6","CHS",2.55],["CHS 42.4x3.2","CHS",3.09],["CHS 48.3x2.6","CHS",2.93],
  ["CHS 48.3x3.2","CHS",3.56],["CHS 48.3x4","CHS",4.37],["CHS 60.3x3.2","CHS",4.51],["CHS 60.3x3.6","CHS",5.03],["CHS 60.3x5","CHS",6.82],["CHS 76.1x3.2","CHS",5.75],["CHS 76.1x4","CHS",7.11],["CHS 76.1x5","CHS",8.77],
  ["CHS 88.9x4","CHS",8.38],["CHS 88.9x5","CHS",10.35],["CHS 114.3x5","CHS",13.48],["CHS 114.3x6.3","CHS",16.78],["CHS 139.7x5","CHS",16.61],["CHS 139.7x6.3","CHS",20.73],["CHS 168.3x5","CHS",20.14],["CHS 168.3x6","CHS",24.02],
  ["CHS 168.3x6.3","CHS",25.17],["CHS 193.7x6.3","CHS",29.12],["CHS 219.1x6","CHS",31.53],["CHS 219.1x6.3","CHS",33.06],["CHS 244.5x6.3","CHS",37.01],["CHS 273x6.3","CHS",41.44],["CHS 273x8","CHS",52.28],["CHS 323.9x8","CHS",62.32],
  ["CHS 355.6x8","CHS",68.58],["CHS 355.6x10","CHS",85.23],["CHS 406.4x10","CHS",97.76],["CHS 457x10","CHS",110.24],["CHS 508x10","CHS",122.81],
  // JIS HW wide-flange — extra (heavy) series
  ["HW 200x204","JIS-HW",56.2],["HW 250x255","JIS-HW",82.2],["HW 300x305","JIS-HW",106],["HW 350x357","JIS-HW",160],["HW 400x408","JIS-HW",197],["HW 414x405","JIS-HW",232],["HW 428x407","JIS-HW",283],["HW 458x417","JIS-HW",415],["HW 498x432","JIS-HW",605],
  // JIS HM medium-flange — deep series
  ["HM 550x300","JIS-HM",137],["HM 600x300","JIS-HM",151],["HM 600x302","JIS-HM",175],["HM 700x300","JIS-HM",185],["HM 800x300","JIS-HM",210],["HM 900x300","JIS-HM",243],
  // JIS HN narrow-flange — extra sizes
  ["HN 198x99","JIS-HN",18.2],["HN 248x124","JIS-HN",25.1],["HN 298x149","JIS-HN",32.0],["HN 346x174","JIS-HN",41.4],["HN 396x199","JIS-HN",56.6],["HN 446x199","JIS-HN",66.7],["HN 496x199","JIS-HN",79.5],["HN 596x199","JIS-HN",94.6],["HN 606x201","JIS-HN",120],["HN 692x300","JIS-HN",166],["HN 792x300","JIS-HN",191],["HN 892x299","JIS-HN",213],
  // JIS I-beams (G3192)
  ["I 100x75x5x8","JIS-I",9.3],["I 125x75x5.5x9.5","JIS-I",13.2],["I 150x75x5.5x9.5","JIS-I",14.0],["I 150x125x8.5x14","JIS-I",28.0],["I 180x100x6x10","JIS-I",21.4],["I 200x100x7x10","JIS-I",26.0],
  ["I 200x150x9x16","JIS-I",50.4],["I 250x125x7.5x12.5","JIS-I",38.3],["I 250x125x10x19","JIS-I",55.5],["I 300x150x8x13","JIS-I",48.3],["I 300x150x10x18.5","JIS-I",65.5],["I 300x150x11.5x22","JIS-I",76.8],
  ["I 350x150x9x15","JIS-I",58.5],["I 350x150x12x24","JIS-I",87.2],["I 400x150x10x18","JIS-I",72.0],["I 400x150x12.5x25","JIS-I",95.8],["I 450x175x11x20","JIS-I",91.7],["I 450x175x13x26","JIS-I",115],
  ["I 600x190x13x25","JIS-I",133],["I 600x190x16x35","JIS-I",176],
  // JIS channels 溝形鋼 (G3192)
  ["C 75x40x5x7","JIS-C",6.92],["C 100x50x5x7.5","JIS-C",9.36],["C 125x65x6x8","JIS-C",13.4],["C 150x75x6.5x10","JIS-C",18.6],["C 150x75x9x12.5","JIS-C",24.0],["C 180x75x7x10.5","JIS-C",21.4],
  ["C 200x80x7.5x11","JIS-C",24.6],["C 200x90x8x13.5","JIS-C",30.3],["C 250x90x9x13","JIS-C",34.6],["C 250x90x11x14.5","JIS-C",40.2],["C 300x90x9x13","JIS-C",38.1],["C 300x90x10x15.5","JIS-C",43.8],
  ["C 300x90x12x16","JIS-C",48.6],["C 380x100x10.5x16","JIS-C",54.5],["C 380x100x13x16.5","JIS-C",62.0],["C 380x100x13x20","JIS-C",67.3],
  // JIS equal angles 等辺山形鋼 (G3192)
  ["L 25x25x3","L",1.11],["L 30x30x3","L",1.34],["L 40x40x3","L",1.81],["L 40x40x5","L",2.94],["L 45x45x4","L",2.7],["L 45x45x5","L",3.34],["L 50x50x4","L",3.01],["L 50x50x6","L",4.43],
  ["L 60x60x4","L",3.64],["L 60x60x5","L",4.51],["L 65x65x5","L",4.91],["L 65x65x6","L",5.84],["L 65x65x8","L",7.66],["L 70x70x6","L",6.31],["L 75x75x6","L",6.78],["L 75x75x9","L",9.96],
  ["L 80x80x6","L",7.25],["L 90x90x7","L",9.51],["L 90x90x10","L",13.34],["L 100x100x7","L",10.61],["L 100x100x10","L",14.91],["L 100x100x13","L",19.08],["L 120x120x8","L",14.57],["L 130x130x9","L",17.73],
  ["L 130x130x12","L",23.36],["L 150x150x12","L",27.13],["L 150x150x15","L",33.56],["L 150x150x19","L",41.91],["L 175x175x12","L",31.84],["L 175x175x15","L",39.45],["L 200x200x15","L",45.33],["L 200x200x20","L",59.66],
  ["L 200x200x25","L",73.59],["L 250x250x25","L",93.22],["L 250x250x35","L",127.76],
  // JIS unequal angles 不等辺山形鋼 (G3192)
  ["L 90x75x9","L",11.02],["L 100x75x7","L",9.23],["L 100x75x10","L",12.95],["L 125x75x7","L",10.61],["L 125x75x10","L",14.91],["L 125x90x10","L",16.09],["L 125x90x13","L",20.61],["L 150x90x9","L",16.32],
  ["L 150x90x12","L",21.48],["L 150x100x9","L",17.03],["L 150x100x12","L",22.42],["L 200x90x9","L",19.85],["L 200x90x14","L",30.33],["L 200x100x10","L",22.76],["L 200x100x15","L",33.56],
  // JIS lipped (light gauge) channels リップ溝形鋼 (G3350)
  ["LC 60x30x10x1.6","JIS-LIP",1.45],["LC 75x45x15x1.6","JIS-LIP",2.06],["LC 75x45x15x2.0","JIS-LIP",2.56],["LC 100x50x20x1.6","JIS-LIP",2.69],["LC 100x50x20x2.0","JIS-LIP",3.34],["LC 100x50x20x2.3","JIS-LIP",3.8],
  ["LC 125x50x20x2.0","JIS-LIP",3.69],["LC 125x50x20x2.3","JIS-LIP",4.21],["LC 150x65x20x2.0","JIS-LIP",4.51],["LC 150x65x20x2.3","JIS-LIP",5.14],["LC 150x75x20x2.3","JIS-LIP",5.5],["LC 200x75x20x2.3","JIS-LIP",6.29],
  ["LC 200x75x25x3.2","JIS-LIP",8.51],["LC 250x75x25x3.2","JIS-LIP",9.55],["LC 250x75x25x4.5","JIS-LIP",13.1],["LC 300x90x30x3.2","JIS-LIP",11.8],
  // JIS Channels (mm) — type JIS-C
  ["C 75x40","JIS-C",6.92],["C 100x50","JIS-C",9.36],["C 125x65","JIS-C",13.4],["C 150x75","JIS-C",18.6],["C 150x75H","JIS-C",24.0],["C 180x75","JIS-C",21.4],["C 200x80","JIS-C",24.6],["C 200x90","JIS-C",30.3],["C 250x90","JIS-C",34.6],["C 250x90H","JIS-C",40.2],["C 300x90","JIS-C",38.1],["C 300x90M","JIS-C",43.8],["C 300x90H","JIS-C",48.6],["C 380x100","JIS-C",54.5],["C 380x100M","JIS-C",62.0],["C 380x100H","JIS-C",67.3],
  // JIS I-beams (mm) — type JIS-I
  ["I 100x75","JIS-I",9.3],["I 125x75","JIS-I",11.5],["I 150x75","JIS-I",14.0],["I 150x125","JIS-I",28.0],["I 180x100","JIS-I",18.4],["I 200x100","JIS-I",21.3],["I 200x150","JIS-I",38.3],["I 250x125","JIS-I",29.6],["I 250x125H","JIS-I",38.3],["I 300x150","JIS-I",36.7],["I 300x150H","JIS-I",48.3],["I 350x150","JIS-I",44.1],["I 350x150H","JIS-I",63.8],["I 400x150","JIS-I",56.6],["I 400x150H","JIS-I",72.0],["I 450x175","JIS-I",76.5],["I 450x175H","JIS-I",91.7],["I 600x190","JIS-I",133],["I 600x190H","JIS-I",176],
  // Russian GOST 8239 I-beams (Двутавр)
  ["GOST I10","GOST-I",9.46],["GOST I12","GOST-I",11.5],["GOST I14","GOST-I",13.7],["GOST I16","GOST-I",15.9],["GOST I18","GOST-I",18.4],["GOST I20","GOST-I",21.0],
  ["GOST I22","GOST-I",24.0],["GOST I24","GOST-I",27.3],["GOST I27","GOST-I",31.5],["GOST I30","GOST-I",36.5],["GOST I33","GOST-I",42.2],["GOST I36","GOST-I",48.6],
  ["GOST I40","GOST-I",57.0],["GOST I45","GOST-I",66.5],["GOST I50","GOST-I",78.5],["GOST I55","GOST-I",92.6],["GOST I60","GOST-I",108.0],
  // Russian GOST 8240 channels (Швеллер)
  ["GOST C5","GOST-C",4.84],["GOST C6.5","GOST-C",5.9],["GOST C8","GOST-C",7.05],["GOST C10","GOST-C",8.59],["GOST C12","GOST-C",10.4],["GOST C14","GOST-C",12.3],
  ["GOST C14a","GOST-C",13.3],["GOST C16","GOST-C",14.2],["GOST C16a","GOST-C",15.3],["GOST C18","GOST-C",16.3],["GOST C18a","GOST-C",17.4],["GOST C20","GOST-C",18.4],
  ["GOST C20a","GOST-C",19.8],["GOST C22","GOST-C",21.0],["GOST C22a","GOST-C",22.6],["GOST C24","GOST-C",24.0],["GOST C24a","GOST-C",25.8],["GOST C27","GOST-C",27.7],
  ["GOST C30","GOST-C",31.8],["GOST C33","GOST-C",36.5],["GOST C36","GOST-C",41.9],["GOST C40","GOST-C",48.3],
  // Russian GOST 26020 wide-flange beams (Б, parallel flange)
  ["GOST 20B1","GOST-B",21.3],["GOST 25B1","GOST-B",25.7],["GOST 25B2","GOST-B",29.6],["GOST 30B1","GOST-B",32.9],["GOST 30B2","GOST-B",36.6],["GOST 35B1","GOST-B",38.9],
  ["GOST 35B2","GOST-B",43.3],["GOST 40B1","GOST-B",48.1],["GOST 40B2","GOST-B",53.0],["GOST 45B1","GOST-B",59.8],["GOST 45B2","GOST-B",66.3],["GOST 50B1","GOST-B",73.0],
  ["GOST 50B2","GOST-B",80.1],["GOST 55B1","GOST-B",89.0],["GOST 55B2","GOST-B",96.8],["GOST 60B1","GOST-B",108.0],["GOST 60B2","GOST-B",117.0],["GOST 70B1","GOST-B",129.0],
  ["GOST 80B1","GOST-B",159.0],["GOST 90B1","GOST-B",183.0],["GOST 100B1","GOST-B",214.0],
  // Russian GOST 26020 column beams (К)
  ["GOST 20K1","GOST-K",41.4],["GOST 20K2","GOST-K",49.9],["GOST 25K1","GOST-K",56.5],["GOST 25K2","GOST-K",67.6],["GOST 30K1","GOST-K",73.8],["GOST 30K2","GOST-K",84.8],
  ["GOST 35K1","GOST-K",109.0],["GOST 35K2","GOST-K",127.0],["GOST 40K1","GOST-K",145.0],["GOST 40K2","GOST-K",163.0],
  // Chinese GB/T 706 I-beams (工字钢)
  ["GB I10","GB-I",11.3],["GB I12.6","GB-I",14.2],["GB I14","GB-I",16.9],["GB I16","GB-I",20.5],["GB I18","GB-I",24.1],["GB I20a","GB-I",27.9],
  ["GB I20b","GB-I",31.1],["GB I22a","GB-I",33.0],["GB I22b","GB-I",36.5],["GB I25a","GB-I",38.1],["GB I25b","GB-I",42.0],["GB I28a","GB-I",43.5],
  ["GB I28b","GB-I",47.9],["GB I32a","GB-I",52.7],["GB I32b","GB-I",57.7],["GB I32c","GB-I",62.8],["GB I36a","GB-I",60.0],["GB I36b","GB-I",65.7],
  ["GB I36c","GB-I",71.3],["GB I40a","GB-I",67.6],["GB I40b","GB-I",73.9],["GB I40c","GB-I",80.2],["GB I45a","GB-I",80.4],["GB I45b","GB-I",87.5],
  ["GB I45c","GB-I",94.6],["GB I50a","GB-I",93.7],["GB I50b","GB-I",101.5],["GB I50c","GB-I",109.4],["GB I56a","GB-I",106.3],["GB I56b","GB-I",115.1],
  ["GB I56c","GB-I",123.9],["GB I63a","GB-I",121.4],["GB I63b","GB-I",131.3],["GB I63c","GB-I",141.2],
  // Chinese GB/T 706 channels (槽钢)
  ["GB C5","GB-C",5.44],["GB C6.3","GB-C",6.63],["GB C8","GB-C",8.04],["GB C10","GB-C",10.0],["GB C12.6","GB-C",12.3],["GB C14a","GB-C",14.5],
  ["GB C14b","GB-C",16.7],["GB C16a","GB-C",17.2],["GB C16b","GB-C",19.8],["GB C18a","GB-C",20.2],["GB C18b","GB-C",23.0],["GB C20a","GB-C",22.6],
  ["GB C20b","GB-C",25.8],["GB C22a","GB-C",25.0],["GB C22b","GB-C",28.5],["GB C25a","GB-C",27.4],["GB C25b","GB-C",31.3],["GB C25c","GB-C",35.3],
  ["GB C28a","GB-C",31.4],["GB C28b","GB-C",35.8],["GB C28c","GB-C",40.2],["GB C32a","GB-C",38.1],["GB C32b","GB-C",43.1],["GB C32c","GB-C",48.1],
  ["GB C36a","GB-C",47.8],["GB C36b","GB-C",53.5],["GB C36c","GB-C",59.1],["GB C40a","GB-C",58.9],["GB C40b","GB-C",65.2],["GB C40c","GB-C",71.5],
  // Additional cold-formed C purlins
  ["C 250x75x2.0","C-COLD",6.13],["C 250x75x3.0","C-COLD",9.12],["C 300x75x3.0","C-COLD",10.3],["C 300x90x2.5","C-COLD",9.34],["C 350x100x3.0","C-COLD",12.5],["C 400x100x3.0","C-COLD",13.7],
  ["C 400x100x4.0","C-COLD",18.1],
].map(([name, type, kgm]) => ({ name, type, kgm }));

const STEEL_GRADES = ["S235JR","S275JR","S355JR","S355J2","S420","S460","A36","A572-50","A992","Q235","Q345","Q355","SS400","ST37"];

/* ─── SECTION HELPERS ────────────────────────────────────────────────────── */
function isExcludedProfile(name) {
  if (!name) return false;
  const n = String(name).trim().toUpperCase().replace(/\s+/g, "");
  return /^(PL|PLT|PLAT|PLATE|FLT|FL|FB|FLAT|SHEET|SHT|GUSSET|GPL|BPL|CHK|CHEQ|CLEAT|STIFF|CAPPL|BASEPL)(?=$|[\d*x×.\-_/])/i.test(n) || /\bPLATE\b/i.test(name);
}
function hollowKgm(type, d) {
  const t = d.t; if (!t || t <= 0) return null;
  if (type === "CHS") { const D = d.a; if (!D) return null; return Math.round(Math.PI * (D - t) * t * 1e-6 * DENSITY * 100) / 100; }
  const H = d.a, B = d.b || d.a; if (!H || !B) return null;
  const area = (H * B) - ((H - 2 * t) * (B - 2 * t)) - (4 - Math.PI) * (2 * t) * (2 * t);
  return Math.round(area * 1e-6 * DENSITY * 100) / 100;
}
function normalizeHollow(raw) {
  let n = String(raw).trim().toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "");
  n = n.replace(/^(CF|HF|CFC|HFC)(RHS|SHS|CHS)/, "$2");
  let m;
  if ((m = n.match(/^SHS(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?:X(\d+(?:\.\d+)?))?$/))) { const a = +m[1], t = m[3] !== undefined ? +m[3] : +m[2]; return { type: "SHS", canonical: `SHS ${a}x${a}x${t}`, kgm: hollowKgm("SHS", { a, b: a, t }) }; }
  if ((m = n.match(/^RHS(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)$/))) { const fd = { a: +m[1], b: +m[2], t: +m[3] }; return { type: "RHS", canonical: `RHS ${fd.a}x${fd.b}x${fd.t}`, kgm: hollowKgm("RHS", fd) }; }
  if ((m = n.match(/^(?:CHS|PIPE|DIA|Ø|O|D)(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)$/))) { const fd = { a: +m[1], t: +m[2] }; return { type: "CHS", canonical: `CHS ${fd.a}x${fd.t}`, kgm: hollowKgm("CHS", fd) }; }
  return null;
}
function normalizeAngle(raw) {
  const n = String(raw).trim().toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "");
  let m;
  if ((m = n.match(/^[LA](\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)(?:X(\d+(?:\.\d+)?))?$/))) {
    const a = +m[1], b = m[3] !== undefined ? +m[2] : a, t = m[3] !== undefined ? +m[3] : +m[2];
    if (!(a > 0 && b > 0 && t > 0) || t > Math.min(a, b)) return null;
    return { type: "L", canonical: `L ${a}x${b}x${t}`, kgm: Math.round(t * (a + b - t) * 1e-6 * DENSITY * 100) / 100 };
  }
  return null;
}
function aliasToCanonical(raw) {
  const n = String(raw).trim().toUpperCase().replace(/\s+/g, ""); let m;
  if ((m = n.match(/^HE(\d+)(A|B|M)$/))) return `HE${m[2]} ${m[1]}`;
  if ((m = n.match(/^(HEA|HEB|HEM)(\d+)$/))) return `${m[1]} ${m[2]}`;
  if ((m = n.match(/^(IPE|IPN|UB|UC|HW|HM|HN|UPN|UPE|PFC|UBP)(\d.*)$/))) return `${m[1]} ${m[2]}`;
  return null;
}
function parseBuiltUp(name) {
  if (!name) return null;
  const m = name.trim().match(/^(?:BU|PG)?\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)$/i);
  if (!m) return null;
  const D = +m[1], Bf = +m[2], tf = +m[3], tw = +m[4];
  if (!(D > 0 && Bf > 0 && tf > 0 && tw > 0) || tf >= D / 2) return null;
  const kgm = ((2 * Bf * tf + (D - 2 * tf) * tw) * 1e-6) * DENSITY;
  return { kgm: Math.round(kgm * 10) / 10, D, Bf, tf, tw, builtUp: true };
}
function findSection(name) {
  if (!name || isExcludedProfile(name)) return null;
  const n = name.trim().toUpperCase().replace(/\s+/g, " ");
  let hit = STEEL_DB.find(s => s.name.toUpperCase() === n);
  if (hit) return hit;
  const compact = n.replace(/\s+/g, "");
  hit = STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === compact);
  if (hit) return hit;
  // separator-agnostic match: unify -, ×, * to x and drop spaces (handles Tekla "H-200x200x8x12", "I-300x150x10x18.5")
  const norm = v => v.toUpperCase().replace(/×/g, "X").replace(/[*∗]/g, "X").replace(/[-_]/g, "").replace(/\s+/g, "");
  const nn = norm(name);
  hit = STEEL_DB.find(s => norm(s.name) === nn);
  if (hit) return hit;
  // Tekla often prefixes JIS H-shapes with "H" instead of HW/HM/HN — match on leading HxB.
  const hm = nn.match(/^H(\d+X\d+)/);
  if (hm) { const lead = hm[1]; hit = STEEL_DB.find(s => { const sn = norm(s.name); return /^H[WMN]/.test(sn) && sn.slice(2).startsWith(lead) && (sn.slice(2) === lead || sn.slice(2 + lead.length).startsWith("X") || sn.slice(2) === lead); }); if (hit) return hit; }
  const alias = aliasToCanonical(name);
  if (alias) { const a = STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === alias.replace(/\s+/g, "")); if (a) return a; }
  const h = normalizeHollow(name); if (h && h.kgm) return STEEL_DB.find(s => s.name.toUpperCase().replace(/\s+/g, "") === h.canonical.toUpperCase().replace(/\s+/g, "")) || { name: h.canonical, type: h.type, kgm: h.kgm, computed: true };
  const ang = normalizeAngle(name); if (ang && ang.kgm) return { name: ang.canonical, type: ang.type, kgm: ang.kgm, computed: true };
  const bu = parseBuiltUp(name); if (bu) return { name: name.trim().toUpperCase(), type: "BUILT-UP", kgm: bu.kgm, builtUp: true };
  return null;
}
function searchSections(q, limit = 12) {
  if (!q) return [];
  const Q = q.trim().toUpperCase().replace(/\s+/g, " ");
  const starts = [], contains = [];
  for (const s of STEEL_DB) { const up = s.name.toUpperCase(); if (up.startsWith(Q)) starts.push(s); else if (up.replace(/\s/g, "").includes(Q.replace(/\s/g, ""))) contains.push(s); }
  return [...starts, ...contains].slice(0, limit);
}

/* ─── L-SHAPE OFFCUT HELPERS ─────────────────────────────────────────────
   An L-offcut = full rectangle A (width) × B (length) with a corner notch
   C (width) × D (length) removed. The largest axis-aligned rectangle that
   fits inside the L is the better of the two leftover rectangles:
     • (A − C) × B   (full-height strip beside the notch)
     • A × (B − D)   (full-width strip below the notch)
   We nest only into that usable rectangle — never claims metal that isn't there. */
function lUsableRect(A, B, C, D) {
  const r1 = { w: A - C, h: B };       // beside the notch
  const r2 = { w: A, h: B - D };       // below the notch
  return (r1.w * r1.h >= r2.w * r2.h) ? r1 : r2;
}

/* Pack parts (with qty) into ONE usable rectangle. Returns {placements, leftover}.
   leftover = the same parts list with reduced quantities for whatever didn't fit. */
function nestIntoRect(rectW, rectH, parts, kerf, allowRotation) {
  // expand
  let pieces = [];
  parts.forEach(p => { for (let q = 0; q < p.qty; q++) pieces.push({ w: p.width, h: p.length, id: p.id, label: p.label }); });
  pieces.sort((a, b) => (b.w * b.h) - (a.w * a.h));
  const placements = []; const placedCount = {}; const notPlaced = [];
  let shelf = { usedW: 0, rowH: 0 }, shelfY = 0;
  pieces.forEach(part => {
    const tryPlace = (w, h, rot) => { if (shelf.usedW + w + (shelf.usedW > 0 ? kerf : 0) <= rectW && shelfY + h <= rectH) { placements.push({ x: shelf.usedW, y: shelfY, w, h, rotated: rot, id: part.id, label: part.label }); shelf.usedW += w + kerf; shelf.rowH = Math.max(shelf.rowH, h); return true; } return false; };
    let done = false;
    if (tryPlace(part.w, part.h, false)) done = true;
    else if (allowRotation && tryPlace(part.h, part.w, true)) done = true;
    else { shelfY += shelf.rowH + kerf; shelf = { usedW: 0, rowH: 0 }; if (shelfY < rectH) { if (tryPlace(part.w, part.h, false)) done = true; else if (allowRotation && tryPlace(part.h, part.w, true)) done = true; } }
    if (done) placedCount[part.id] = (placedCount[part.id] || 0) + 1; else notPlaced.push(part);
  });
  const leftover = parts.map(p => ({ ...p, qty: p.qty - (placedCount[p.id] || 0) })).filter(p => p.qty > 0);
  return { placements, leftover, placedAny: placements.length > 0 };
}

/* Find reusable leftover on a sheet by scanning the FREE area. Uses a y-band
   sweep: for every horizontal band bounded by part edges, find the empty x-gaps.
   This is guaranteed never to overlap a placed part. Reports the biggest empty
   rectangle, or an L when there's a clean right + bottom corner region. */
function findOffcuts(placements, m, usableW, usableH, reuseMin) {
  const L = m, R = m + usableW, T = m, B = m + usableH;
  const yTops = [...new Set([T, ...placements.map(p => p.y + p.h)])].filter(y => y >= T && y < B).sort((a, b) => a - b);
  const yBots = [...new Set([B, ...placements.map(p => p.y)])].filter(y => y > T && y <= B).sort((a, b) => a - b);
  let best = null;
  for (const y1 of yTops) for (const y2 of yBots) {
    const h = y2 - y1; if (h < reuseMin) continue;
    // x-intervals blocked by any part overlapping this y-band
    const blockers = placements
      .filter(p => p.y < y2 - 0.5 && p.y + p.h > y1 + 0.5)
      .map(p => [Math.max(L, p.x), Math.min(R, p.x + p.w)])
      .filter(([a, b]) => b > a)
      .sort((a, b) => a[0] - b[0]);
    let cursor = L;
    for (const [bx0, bx1] of blockers) {
      if (bx0 - cursor >= reuseMin) { const w = bx0 - cursor, area = w * h; if (!best || area > best.area) best = { x: cursor, y: y1, w, h, area }; }
      cursor = Math.max(cursor, bx1);
    }
    if (R - cursor >= reuseMin) { const w = R - cursor, area = w * h; if (!best || area > best.area) best = { x: cursor, y: y1, w, h, area }; }
  }
  // L-shape = sheet minus the top-left used bounding box (both strips guaranteed empty).
  const usedRight = placements.length ? Math.max(...placements.map(p => p.x + p.w)) : L;
  const usedBottom = placements.length ? Math.max(...placements.map(p => p.y + p.h)) : T;
  const rightW = R - usedRight, bottomH = B - usedBottom, notchW = usedRight - L, notchH = usedBottom - T;
  if (rightW >= reuseMin && bottomH >= reuseMin && notchW > 0 && notchH > 0) {
    const Larea = usableW * usableH - notchW * notchH;
    const r1 = { w: rightW, h: usableH }, r2 = { w: usableW, h: bottomH };
    const usable = (r1.w * r1.h >= r2.w * r2.h) ? r1 : r2;
    if (!best || Larea >= best.area) return [{ shape: "L", x: L, y: T, A: Math.round(usableW), B: Math.round(usableH), notchW: Math.round(notchW), notchH: Math.round(notchH), w: Math.round(usable.w), h: Math.round(usable.h) }];
  }
  if (best) return [{ shape: "rect", x: best.x, y: best.y, A: Math.round(best.w), B: Math.round(best.h), w: Math.round(best.w), h: Math.round(best.h) }];
  return [];
}

/* ─── PLATE NESTING ENGINE (with reusable offcuts + oversize detection) ──── */
function splitPart(W, L, uw, uh, allowRotation, pref) {
  // Split a part too big for the sheet into welded sub-pieces, divided EQUALLY so
  // the weld lands sensibly and there are never tiny slivers.
  //   pref "welds" → fewest pieces / fewest welds (bigger pieces, may pack looser)
  //   pref "pack"  → smaller equal pieces that tile the sheet height better (less scrap)
  const layout = (a, b, colDiv, rowDiv) => {
    const cols = Math.max(1, colDiv(a), Math.ceil(a / uw)), rows = Math.max(1, rowDiv(b), Math.ceil(b / uh));
    const cw = a / cols, rh = b / rows, list = [];
    for (let c = 0; c < cols; c++) for (let r = 0; r < rows; r++) list.push({ w: Math.round(cw), h: Math.round(rh) });
    return list;
  };
  const minPieces = (a, b) => layout(a, b, () => 1, () => 1);
  // packing layout: choose row/col counts so each piece ≈ half the sheet (tiles better)
  const packLayout = (a, b) => layout(a, b, x => Math.ceil(x / (uw / 2)), y => Math.ceil(y / (uh / 2)));
  const make = (a, b) => pref === "pack" ? packLayout(a, b) : minPieces(a, b);
  const o1 = make(W, L);
  if (allowRotation) { const o2 = make(L, W); if (o2.length < o1.length) return o2; }
  return o1;
}

function nestPlates(sheetW, sheetH, parts, kerf, margin, allowRotation, reuseMin, splicePref) {
  const sheets = [];
  const fitsRaw = (w, h) => (w <= sheetW && h <= sheetH) || (allowRotation && h <= sheetW && w <= sheetH);

  // Auto-clamp the edge margin from the parts that already fit, so a near-sheet part isn't blocked.
  let m = margin;
  parts.forEach(p => { if (fitsRaw(p.width, p.length)) { const longP = Math.max(p.width, p.length), shortP = Math.min(p.width, p.length); const longS = Math.max(sheetW, sheetH), shortS = Math.min(sheetW, sheetH); if (longP <= longS) m = Math.min(m, Math.floor((longS - longP) / 2)); if (shortP <= shortS) m = Math.min(m, Math.floor((shortS - shortP) / 2)); } });
  m = Math.max(0, m);
  const usableW = sheetW - m * 2, usableH = sheetH - m * 2;

  // Build the piece queue, SPLICING any part bigger than the sheet into welded sub-pieces.
  let queue = [];
  const spliceMap = {}; // label -> {label, W, L, count, qty}
  parts.forEach(p => {
    for (let q = 0; q < p.qty; q++) {
      if (fitsRaw(p.width, p.length)) { queue.push({ w: p.width, h: p.length, id: p.id, label: p.label }); }
      else {
        const subs = splitPart(p.width, p.length, usableW, usableH, allowRotation, splicePref);
        subs.forEach((s, idx) => queue.push({ w: s.w, h: s.h, id: p.id, label: `${p.label}·${idx + 1}/${subs.length}`, spliced: true }));
        const key = p.label; if (!spliceMap[key]) spliceMap[key] = { label: p.label, W: p.width, L: p.length, count: subs.length, qty: 0 }; spliceMap[key].qty += 1;
      }
    }
  });

  queue.sort((a, b) => (b.w * b.h) - (a.w * a.h));
  let guard = 0;
  while (queue.length > 0 && guard < 20000) {
    guard++;
    const placements = []; let shelf = { usedW: 0, rowH: 0 }, shelfY = m; const remaining = []; const rows = [];
    queue.forEach(part => {
      const tryPlace = (w, h, rot) => { if (shelf.usedW + w + kerf <= usableW + kerf && h <= (usableH - (shelfY - m))) { placements.push({ x: m + shelf.usedW, y: shelfY, w, h, rotated: rot, id: part.id, label: part.label, spliced: part.spliced }); shelf.usedW += w + kerf; shelf.rowH = Math.max(shelf.rowH, h); return true; } return false; };
      let done = false;
      if (tryPlace(part.w, part.h, false)) done = true;
      else if (allowRotation && tryPlace(part.h, part.w, true)) done = true;
      else { rows.push({ y: shelfY, usedW: shelf.usedW, rowH: shelf.rowH }); shelfY += shelf.rowH + kerf; shelf = { usedW: 0, rowH: 0 }; if (shelfY - m < usableH) { if (tryPlace(part.w, part.h, false)) done = true; else if (allowRotation && tryPlace(part.h, part.w, true)) done = true; } }
      if (!done) remaining.push(part);
    });
    if (shelf.usedW > 0 || shelf.rowH > 0) rows.push({ y: shelfY, usedW: shelf.usedW, rowH: shelf.rowH });
    if (placements.length === 0 && remaining.length > 0) { const p = remaining.shift(); sheets.push({ placements: [{ x: m, y: m, w: Math.min(p.w, usableW), h: Math.min(p.h, usableH), rotated: false, id: p.id, label: p.label, spliced: p.spliced }], offcuts: [] }); }
    else {
      sheets.push({ placements, offcuts: findOffcuts(placements, m, usableW, usableH, reuseMin) });
    }
    queue = remaining;
  }
  sheets.splices = Object.values(spliceMap);
  sheets.marginUsed = m;
  return sheets;
}

/* ─── SECTION (1D BAR) NESTING ENGINE ────────────────────────────────────── */
function nestBars(items, kerf) {
  const pieces = []; const spliceMap = {}; let fullBarsFromSplices = 0;
  items.forEach(it => {
    for (let i = 0; i < it.qty; i++) {
      if (it.length > it.stock) {
        const bars = Math.ceil(it.length / it.stock); const fullBars = bars - 1; fullBarsFromSplices += fullBars;
        const remainder = it.length - fullBars * it.stock;
        if (remainder > 0) pieces.push({ length: remainder, label: it.label || "", stock: it.stock, spliced: true });
        const key = `${it.length}`; if (!spliceMap[key]) spliceMap[key] = { length: it.length, qty: 0, bars, stock: it.stock }; spliceMap[key].qty += 1;
      } else pieces.push({ length: it.length, label: it.label || "", stock: it.stock });
    }
  });
  pieces.sort((a, b) => b.length - a.length);
  const bins = [];
  for (let i = 0; i < fullBarsFromSplices; i++) { const stock = items[0].stock; bins.push({ stockLength: stock, cuts: [{ length: stock, label: "(spliced run)", spliced: true }], remaining: 0, spliceBar: true }); }
  for (const p of pieces) {
    let best = null, bestRem = Infinity;
    for (const b of bins) { if (b.spliceBar || b.stockLength !== p.stock) continue; const need = p.length + kerf; if (b.remaining >= need && (b.remaining - need) < bestRem) { bestRem = b.remaining - need; best = b; } }
    if (best) { best.cuts.push({ length: p.length, label: p.label, spliced: p.spliced }); best.remaining -= p.length + kerf; }
    else bins.push({ stockLength: p.stock, cuts: [{ length: p.length, label: p.label, spliced: p.spliced }], remaining: p.stock - p.length - kerf });
  }
  const totalStock = bins.reduce((s, b) => s + b.stockLength, 0);
  const totalNet = bins.reduce((s, b) => s + b.cuts.reduce((ss, c) => ss + c.length, 0), 0);
  const totalWaste = bins.reduce((s, b) => s + b.remaining, 0);
  const wastePct = totalStock ? ((totalWaste / totalStock) * 100).toFixed(1) : "0";
  return { bins, splices: Object.values(spliceMap), summary: { stockCount: bins.length, totalStock, totalNet, totalWaste, wastePct, utilPct: (100 - parseFloat(wastePct)).toFixed(1) } };
}

/* ════════════════════════════════════════════════════════════════════════
   3D COVER — rotating I-beam + plate (Canvas 2D perspective, amber theme)
════════════════════════════════════════════════════════════════════════ */
function Cover3D({ onStart }) {
  const ref = useRef(null), raf = useRef(null), mouse = useRef({ x: 0, y: 0, tx: 0, ty: 0 });
  useEffect(() => {
    const canvas = ref.current; if (!canvas) return; const ctx = canvas.getContext("2d");
    let t = 0;
    const resize = () => { const dpr = Math.min(window.devicePixelRatio || 1, 2); canvas.width = canvas.offsetWidth * dpr; canvas.height = canvas.offsetHeight * dpr; ctx.setTransform(dpr, 0, 0, dpr, 0, 0); };
    resize(); window.addEventListener("resize", resize);
    const onMove = e => { const r = canvas.getBoundingClientRect(); mouse.current.tx = (e.clientX - r.left) / r.width - 0.5; mouse.current.ty = (e.clientY - r.top) / r.height - 0.5; };
    window.addEventListener("mousemove", onMove);
    const proj = (x, y, z, cx, cy, rx, ry) => { const f = 1200, cx1 = Math.cos(rx), sx = Math.sin(rx), cy1 = Math.cos(ry), sy = Math.sin(ry); const y1 = y * cx1 - z * sx, z1 = y * sx + z * cx1; const x2 = x * cy1 + z1 * sy, z2 = -x * sy + z1 * cy1; const s = f / (f + z2 + 460); return { x: cx + x2 * s, y: cy + y1 * s, s }; };
    const beamProfile = [[-1.1,-1.5],[1.1,-1.5],[1.1,-1.15],[0.15,-1.15],[0.15,1.15],[1.1,1.15],[1.1,1.5],[-1.1,1.5],[-1.1,1.15],[-0.15,1.15],[-0.15,-1.15],[-1.1,-1.15]];
    const render = () => {
      const W = canvas.offsetWidth, H = canvas.offsetHeight;
      mouse.current.x += (mouse.current.tx - mouse.current.x) * .05; mouse.current.y += (mouse.current.ty - mouse.current.y) * .05;
      const my = mouse.current.y;
      ctx.clearRect(0, 0, W, H);
      const bg = ctx.createLinearGradient(0, 0, W, H); bg.addColorStop(0, "#0e1117"); bg.addColorStop(0.5, "#141a22"); bg.addColorStop(1, "#0a0d12"); ctx.fillStyle = bg; ctx.fillRect(0, 0, W, H);
      const glow = ctx.createRadialGradient(W * .5, H * .46, 0, W * .5, H * .46, W * .6); glow.addColorStop(0, "rgba(245,158,11,.13)"); glow.addColorStop(.5, "rgba(245,158,11,.04)"); glow.addColorStop(1, "rgba(0,0,0,0)"); ctx.fillStyle = glow; ctx.fillRect(0, 0, W, H);
      ctx.strokeStyle = "rgba(245,158,11,.05)"; ctx.lineWidth = 1; const gs = 46; for (let x = 0; x < W; x += gs) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); } for (let y = 0; y < H; y += gs) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
      for (let i = 0; i < 45; i++) { const px = (i * 137.5 + t * 6) % W, py = (i * 79.3 + Math.sin(t * .3 + i) * 30 + t * 2) % H; ctx.fillStyle = `rgba(245,200,120,${(Math.sin(t * .4 + i) * .5 + .5) * .25})`; ctx.beginPath(); ctx.arc(px, py, .9, 0, 7); ctx.fill(); }
      // Pure vertical (Y) axis spin — no tilt, like a turntable seen head-on.
      const rx = 0, ry = t * 0.5;
      // ── CENTERED I-BEAM, brushed-metal faces ──
      const bcx = W * 0.5, bcy = H * 0.5, sc = Math.min(W, H) * 0.11, depth = 9;
      const front = beamProfile.map(([x, y]) => proj(x * sc, y * sc, depth * sc * 0.5, bcx, bcy, rx, ry));
      const back = beamProfile.map(([x, y]) => proj(x * sc, y * sc, -depth * sc * 0.5, bcx, bcy, rx, ry));
      // light direction follows rotation so metal "catches" the light as it turns
      const lit = (base) => { const s = 0.55 + 0.45 * Math.abs(Math.cos(ry)); return base.map(ch => Math.min(255, Math.round(ch * s))); };
      for (let i = 0; i < beamProfile.length; i++) {
        const j = (i + 1) % beamProfile.length;
        ctx.beginPath(); ctx.moveTo(front[i].x, front[i].y); ctx.lineTo(front[j].x, front[j].y); ctx.lineTo(back[j].x, back[j].y); ctx.lineTo(back[i].x, back[i].y); ctx.closePath();
        const [r1, g1, b1] = lit([110, 116, 128]), [r2, g2, b2] = lit([170, 180, 192]), [r3, g3, b3] = lit([74, 82, 92]);
        const g = ctx.createLinearGradient(front[i].x, front[i].y, back[j].x, back[j].y);
        g.addColorStop(0, `rgb(${r1},${g1},${b1})`); g.addColorStop(.45, `rgb(${r2},${g2},${b2})`); g.addColorStop(.55, "#e8edf2"); g.addColorStop(1, `rgb(${r3},${g3},${b3})`);
        ctx.fillStyle = g; ctx.fill(); ctx.strokeStyle = "rgba(20,26,34,.55)"; ctx.lineWidth = .6; ctx.stroke();
      }
      // front face: brushed steel + moving specular streak
      ctx.save();
      ctx.beginPath(); front.forEach((p, i) => i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)); ctx.closePath(); ctx.clip();
      const fg = ctx.createLinearGradient(front[0].x, front[0].y - sc, front[0].x, front[0].y + sc * 1.6);
      fg.addColorStop(0, "#cfd7e0"); fg.addColorStop(.3, "#eef2f6"); fg.addColorStop(.5, "#b8c2cd"); fg.addColorStop(.72, "#dde3ea"); fg.addColorStop(1, "#9aa4b0");
      ctx.fillStyle = fg; ctx.fill();
      // brushed horizontal striations
      ctx.globalAlpha = .12; ctx.strokeStyle = "#ffffff"; ctx.lineWidth = .6;
      for (let yy = bcy - sc * 1.6; yy < bcy + sc * 1.6; yy += 3) { ctx.beginPath(); ctx.moveTo(bcx - sc * 1.4, yy); ctx.lineTo(bcx + sc * 1.4, yy + (Math.random() - .5) * 1.5); ctx.stroke(); }
      ctx.globalAlpha = 1;
      // specular sweep tied to rotation
      const sweep = (Math.sin(ry) * .5 + .5);
      const sp = ctx.createLinearGradient(bcx - sc * 1.4, 0, bcx + sc * 1.4, 0);
      sp.addColorStop(Math.max(0, sweep - .18), "rgba(255,255,255,0)"); sp.addColorStop(sweep, "rgba(255,255,255,.5)"); sp.addColorStop(Math.min(1, sweep + .18), "rgba(255,255,255,0)");
      ctx.fillStyle = sp; ctx.fillRect(bcx - sc * 1.6, bcy - sc * 1.7, sc * 3.2, sc * 3.4);
      ctx.restore();
      ctx.beginPath(); front.forEach((p, i) => i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)); ctx.closePath(); ctx.strokeStyle = "#f59e0b"; ctx.lineWidth = 1.6; ctx.stroke();
      t += 0.016; raf.current = requestAnimationFrame(render);
    };
    render();
    return () => { cancelAnimationFrame(raf.current); window.removeEventListener("resize", resize); window.removeEventListener("mousemove", onMove); };
  }, []);
  return (
    <div style={{ position: "relative", width: "100%", height: "100vh", minHeight: 560, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <canvas ref={ref} style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }} />
      <div style={{ position: "absolute", inset: 0, background: "radial-gradient(ellipse 50% 42% at 50% 52%, rgba(10,13,18,.5) 0%, rgba(10,13,18,.25) 45%, rgba(10,13,18,0) 72%)", zIndex: 1, pointerEvents: "none" }} />
      <div style={{ position: "relative", zIndex: 2, textAlign: "center", padding: "0 24px", pointerEvents: "none" }}>
        <div style={{ display: "inline-block", background: "rgba(245,158,11,.12)", border: "1px solid rgba(245,158,11,.4)", borderRadius: 30, padding: "5px 18px", marginBottom: 22, fontFamily: "'Space Mono', monospace", fontSize: 11, letterSpacing: 3, color: "#fbbf24", textTransform: "uppercase" }}>Steel Fabrication Intelligence</div>
        <h1 style={{ margin: "0 0 16px", lineHeight: 1.04, fontFamily: "'Playfair Display', Georgia, serif", fontWeight: 900, fontSize: "clamp(36px,6vw,72px)", color: "#f8fafc", textShadow: "0 4px 30px rgba(0,0,0,.6)" }}>Steel <span style={{ color: "#f59e0b" }}>Optimizer</span></h1>
        <p style={{ margin: "0 0 38px", color: "#cbd5e1", fontSize: "clamp(14px,2vw,18px)", fontFamily: "'Playfair Display', Georgia, serif", fontWeight: 600 }}>Plates &amp; Sections &nbsp;•&nbsp; Minimum waste &nbsp;•&nbsp; Instant tonnage</p>
        <button onClick={onStart} style={{ pointerEvents: "auto", background: "linear-gradient(135deg,#f59e0b,#d97706)", border: "none", color: "#1a1206", padding: "15px 50px", borderRadius: 8, fontSize: 16, fontWeight: 800, cursor: "pointer", letterSpacing: 1, fontFamily: "'Space Mono', monospace", boxShadow: "0 10px 34px rgba(245,158,11,.4)" }}>START OPTIMIZING →</button>
      </div>
      <div style={{ position: "absolute", bottom: 28, left: "50%", transform: "translateX(-50%)", color: "rgba(245,200,120,.55)", fontSize: 11, letterSpacing: 2, fontFamily: "'Space Mono', monospace", textTransform: "uppercase", zIndex: 2 }}>↓ scroll to start</div>
    </div>
  );
}

/* ─── PLATE NESTING CANVAS ───────────────────────────────────────────────── */
function PlateCanvas({ sheets, sheetW, sheetH, colorMap, thickness }) {
  const [active, setActive] = useState(0); const [zoom, setZoom] = useState(1); const cv = useRef(null);
  useEffect(() => { setActive(0); }, [sheets]);
  const sObj = sheets[active] || { placements: [], offcuts: [] };
  const sheet = sObj.placements || [], offcuts = sObj.offcuts || [];
  const SCALE = Math.min(520 / sheetW, 360 / sheetH) * zoom;
  useEffect(() => {
    const c = cv.current; if (!c) return; const ctx = c.getContext("2d");
    const W = sheetW * SCALE, H = sheetH * SCALE; c.width = W + 44; c.height = H + 44;
    ctx.fillStyle = "#0f1318"; ctx.fillRect(0, 0, c.width, c.height); ctx.save(); ctx.translate(22, 22);
    ctx.fillStyle = "#1b2530"; ctx.fillRect(0, 0, W, H); ctx.strokeStyle = "#f59e0b88"; ctx.lineWidth = 1.5; ctx.strokeRect(0, 0, W, H);
    ctx.strokeStyle = "rgba(120,130,140,.2)"; ctx.lineWidth = .5; const step = Math.max(20, Math.round(100 * SCALE));
    for (let x = step; x < W; x += step) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, H); ctx.stroke(); }
    for (let y = step; y < H; y += step) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
    sheet.forEach((p, i) => { const px = p.x * SCALE, py = p.y * SCALE, pw = p.w * SCALE, ph = p.h * SCALE; const color = colorMap[p.id] || PART_COLORS[i % PART_COLORS.length]; ctx.fillStyle = color + "55"; ctx.fillRect(px, py, pw, ph); ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.strokeRect(px, py, pw, ph);
      if (p.spliced) { ctx.save(); ctx.beginPath(); ctx.rect(px, py, pw, ph); ctx.clip(); ctx.strokeStyle = "rgba(255,255,255,.35)"; ctx.lineWidth = 1; for (let d = -ph; d < pw; d += 8) { ctx.beginPath(); ctx.moveTo(px + d, py); ctx.lineTo(px + d + ph, py + ph); ctx.stroke(); } ctx.restore(); ctx.strokeStyle = "#fbbf24"; ctx.lineWidth = 2; ctx.setLineDash([5, 3]); ctx.strokeRect(px, py, pw, ph); ctx.setLineDash([]); }
      const fs = Math.max(8, Math.min(11, pw / 6, ph / 3)); ctx.fillStyle = "rgba(255,255,255,.95)"; ctx.font = `${fs}px 'Space Mono', monospace`; ctx.textAlign = "center"; ctx.textBaseline = "middle"; if (pw > 22 && ph > 13) { ctx.fillText(p.rotated ? `${p.label}↺` : p.label, px + pw / 2, py + ph / 2 - (ph > 26 ? 6 : 0)); if (ph > 26) { ctx.font = `${Math.max(7, fs - 1)}px 'Space Mono', monospace`; ctx.fillStyle = p.spliced ? "#fcd34d" : "rgba(230,210,170,.8)"; ctx.fillText(p.spliced ? `WELD ${Math.round(p.w)}×${Math.round(p.h)}` : `${Math.round(p.w)}×${Math.round(p.h)}`, px + pw / 2, py + ph / 2 + 7); } } });
    // Reusable leftover (green): draw the real shape — L-shape or rectangle — plus its usable rectangle dashed.
    offcuts.forEach(o => {
      ctx.save();
      if (o.shape === "L") {
        const ox = o.x * SCALE, oy = o.y * SCALE, A = o.A * SCALE, B = o.B * SCALE, nW = o.notchW * SCALE, nH = o.notchH * SCALE;
        ctx.beginPath(); ctx.moveTo(ox + nW, oy); ctx.lineTo(ox + A, oy); ctx.lineTo(ox + A, oy + B); ctx.lineTo(ox, oy + B); ctx.lineTo(ox, oy + nH); ctx.lineTo(ox + nW, oy + nH); ctx.closePath();
        ctx.fillStyle = "rgba(16,185,129,.18)"; ctx.fill(); ctx.strokeStyle = "#10b981"; ctx.lineWidth = 1.4; ctx.stroke();
      } else {
        const ox = o.x * SCALE, oy = o.y * SCALE, A = o.A * SCALE, B = o.B * SCALE;
        ctx.fillStyle = "rgba(16,185,129,.18)"; ctx.fillRect(ox, oy, A, B); ctx.strokeStyle = "#10b981"; ctx.lineWidth = 1.4; ctx.strokeRect(ox, oy, A, B);
      }
      ctx.restore();
      if (o.A * SCALE > 30 && o.B * SCALE > 16) { ctx.fillStyle = "#6ee7b7"; ctx.font = `${Math.max(8, Math.min(10, o.A * SCALE / 10))}px 'Space Mono', monospace`; ctx.textAlign = "center"; ctx.textBaseline = "middle"; const cx = (o.x + (o.shape === "L" ? o.A : o.A) / 2) * SCALE; const cy = (o.y + o.B / 2) * SCALE; ctx.fillText(o.shape === "L" ? "♻ L-OFFCUT" : "♻ OFFCUT", cx, cy - 5); ctx.font = `${Math.max(7, Math.min(9, o.A * SCALE / 12))}px 'Space Mono', monospace`; ctx.fillText(`use ${o.w}×${o.h}`, cx, cy + 6); }
    });
    ctx.fillStyle = "rgba(245,200,120,.8)"; ctx.font = "10px 'Space Mono', monospace"; ctx.textAlign = "center"; ctx.fillText(`${sheetW} mm`, W / 2, H + 15);
    ctx.save(); ctx.translate(-9, H / 2); ctx.rotate(-Math.PI / 2); ctx.textAlign = "center"; ctx.fillText(`${sheetH} mm`, 0, 0); ctx.restore(); ctx.restore();
  }, [sheet, offcuts, sheetW, sheetH, SCALE, colorMap]);
  const used = sheet.reduce((s, p) => s + p.w * p.h, 0), total = sheetW * sheetH, util = Math.round((used / total) * 100);
  return (
    <div style={{ fontFamily: "'Space Mono', monospace" }}>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        {sheets.map((sh, i) => { const pct = Math.round(((sh.placements || []).reduce((s, p) => s + p.w * p.h, 0) / total) * 100); return <button key={i} onClick={() => setActive(i)} style={{ padding: "5px 12px", borderRadius: 4, border: "1px solid", borderColor: active === i ? "#f59e0b" : "#2d3748", background: active === i ? "rgba(245,158,11,.15)" : "rgba(15,19,24,.8)", color: active === i ? "#f59e0b" : "#94a3b8", fontSize: 11, cursor: "pointer" }}>Sheet {i + 1} · {pct}%</button>; })}
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}><button onClick={() => setZoom(z => Math.min(z + .2, 3))} style={ZB}>+</button><button onClick={() => setZoom(z => Math.max(z - .2, .4))} style={ZB}>−</button></div>
      </div>
      <div style={{ overflow: "auto", borderRadius: 6, border: "1px solid #2d3748", maxHeight: 440 }}><canvas ref={cv} style={{ display: "block" }} /></div>
      {offcuts.length > 0 && <div style={{ marginTop: 8, padding: "8px 12px", background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.3)", borderRadius: 6, fontSize: 11, color: "#6ee7b7" }}>♻ Reusable leftover{offcuts.length > 1 ? "s" : ""}: {offcuts.map(o => `${o.shape === "L" ? "L-shape, use " : ""}${o.w}×${o.h}mm`).join(" · ")}</div>}
      <div style={{ marginTop: 10, display: "flex", gap: 18, flexWrap: "wrap", padding: "10px 14px", background: "rgba(15,19,24,.6)", borderRadius: 6, border: "1px solid #2d3748" }}>
        <Mini label="Sheet" value={`${active + 1} / ${sheets.length}`} /><Mini label="Thickness" value={`${thickness} mm`} /><Mini label="Parts" value={sheet.length} /><Mini label="Utilization" value={`${util}%`} accent={util > 80 ? "#10B981" : util > 60 ? "#F59E0B" : "#EF4444"} />
      </div>
    </div>
  );
}
function Mini({ label, value, accent }) { return <div style={{ textAlign: "center" }}><div style={{ fontSize: 10, color: "#64748b", marginBottom: 2, letterSpacing: 1 }}>{label}</div><div style={{ fontSize: 14, color: accent || "#cbd5e1", fontWeight: 700 }}>{value}</div></div>; }

/* ─── SECTION CUT BAR ────────────────────────────────────────────────────── */
function CutBar({ bin, index }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
        <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#94a3b8" }}>BAR #{String(index + 1).padStart(2, "0")} — {fmtMm(bin.stockLength)}</span>
        <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#f59e0b" }}>{(((bin.stockLength - bin.remaining) / bin.stockLength) * 100).toFixed(1)}%</span>
      </div>
      <div style={{ display: "flex", height: 28, borderRadius: 4, overflow: "hidden", border: "1px solid #2d3748", background: "#0f1318" }}>
        {bin.cuts.map((cut, ci) => { const w = (cut.length / bin.stockLength) * 100; return <div key={ci} title={`${cut.label ? cut.label + ": " : ""}${fmtMm(cut.length)}`} style={{ width: `${w}%`, background: PART_COLORS[ci % PART_COLORS.length], opacity: .85, borderRight: "1px solid #0f1318", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" }}>{w > 6 && <span style={{ fontSize: 9, fontFamily: "'Space Mono', monospace", color: "#000", fontWeight: 700, padding: "0 2px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{cut.label || fmtMm(cut.length)}</span>}</div>; })}
        {bin.remaining > 0 && <div style={{ width: `${(bin.remaining / bin.stockLength) * 100}%`, background: "#1e293b", borderLeft: "2px dashed #334155", display: "flex", alignItems: "center", justifyContent: "center" }}>{(bin.remaining / bin.stockLength) > 0.04 && <span style={{ fontSize: 9, fontFamily: "'Space Mono', monospace", color: "#475569" }}>{fmtMm(bin.remaining)}</span>}</div>}
      </div>
    </div>
  );
}

/* ─── L-SHAPE INTERACTIVE DIAGRAM (labeled, eliminates ambiguity) ─────────── */
function LShapeDiagram({ A, B, C, D, t }) {
  const W = 240, H = 240, pad = 40;
  const a = Math.max(A, 1), b = Math.max(B, 1), c = Math.min(Math.max(C, 0), a - 1), d = Math.min(Math.max(D, 0), b - 1);
  const sc = Math.min((W - pad * 2) / a, (H - pad * 2) / b);
  const ox = pad, oy = pad, aw = a * sc, bh = b * sc, cw = c * sc, dh = d * sc;
  const pts = [[ox, oy + dh], [ox + aw - cw, oy + dh], [ox + aw - cw, oy], [ox + aw, oy], [ox + aw, oy + bh], [ox, oy + bh]];
  const usable = lUsableRect(a, b, c, d);
  let ux, uy, uw, uh;
  if (usable.w === a) { ux = ox; uy = oy + dh; uw = aw; uh = (b - d) * sc; } else { ux = ox; uy = oy; uw = (a - c) * sc; uh = bh; }
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", maxWidth: 280, background: "rgba(15,19,24,.6)", borderRadius: 8, border: "1px solid #2d3748" }}>
      <polygon points={pts.map(p => p.join(",")).join(" ")} fill="rgba(245,158,11,.12)" stroke="#f59e0b" strokeWidth="1.5" />
      <rect x={ux} y={uy} width={uw} height={uh} fill="rgba(16,185,129,.22)" stroke="#10b981" strokeWidth="1.4" strokeDasharray="5 3" />
      <text x={ux + uw / 2} y={uy + uh / 2} fill="#6ee7b7" fontSize="9" textAnchor="middle" fontFamily="monospace">usable</text>
      <text x={ox + aw / 2} y={oy + bh + 16} fill="#fbbf24" fontSize="11" textAnchor="middle" fontFamily="monospace">A = {Math.round(A)}</text>
      <text x={ox - 10} y={oy + bh / 2} fill="#fbbf24" fontSize="11" textAnchor="middle" fontFamily="monospace" transform={`rotate(-90 ${ox - 10} ${oy + bh / 2})`}>B = {Math.round(B)}</text>
      <text x={ox + aw - cw / 2} y={oy - 6} fill="#93c5fd" fontSize="10" textAnchor="middle" fontFamily="monospace">C = {Math.round(C)}</text>
      <text x={ox + aw + 12} y={oy + dh / 2} fill="#93c5fd" fontSize="10" textAnchor="middle" fontFamily="monospace" transform={`rotate(90 ${ox + aw + 12} ${oy + dh / 2})`}>D = {Math.round(D)}</text>
      <text x={W / 2} y={H - 6} fill="#94a3b8" fontSize="9" textAnchor="middle" fontFamily="monospace">t = {t} mm · green dashed = nested area</text>
    </svg>
  );
}

/* ════════════════════════════════════════════════════════════════════════
   MAIN APP
════════════════════════════════════════════════════════════════════════ */
export default function App() {
  const wsRef = useRef(null);
  const [module, setModule] = useState(null);
  const scrollWs = () => wsRef.current?.scrollIntoView({ behavior: "smooth" });
  return (
    <div style={{ background: "#0a0d12", minHeight: "100vh", color: "#cbd5e1", fontFamily: "'DM Sans', system-ui, sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Space+Mono:wght@400;700&family=Playfair+Display:wght@700;900&display=swap" rel="stylesheet" />
      <Cover3D onStart={scrollWs} />
      <div ref={wsRef} style={{ maxWidth: 1180, margin: "0 auto", padding: "44px 20px 80px" }}>
        {module === null && <ModuleChooser onPick={setModule} />}
        {module === "plates" && <PlatesModule onBack={() => setModule(null)} />}
        {module === "sections" && <SectionsModule onBack={() => setModule(null)} />}
      </div>
    </div>
  );
}

function ModuleChooser({ onPick }) {
  return (
    <div>
      <SectionTitle>Choose what to optimize</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 20, marginTop: 18 }}>
        <Chooser icon="🅸" title="Steel Sections" desc="Beams, columns, angles, Z & C cold-formed sections, and built-up sections. Cut from stock-length bars with splicing for over-length members." onClick={() => onPick("sections")} />
        <Chooser icon="▭" title="Steel Plates & Sheets" desc="Plate parts nested onto stock sheets (2D). Reuse leftover offcuts first, then minimum new sheets, with total tonnage." onClick={() => onPick("plates")} />
      </div>
    </div>
  );
}

/* ─── PLATES MODULE ──────────────────────────────────────────────────────── */
function PlatesModule({ onBack }) {
  const [inputMode, setInputMode] = useState(null);
  const DEFAULT_SHEET = { w: 1220, h: 2440 };
  const [stock, setStock] = useState([{ thickness: 8, w: 1220, h: 2440 }, { thickness: 12, w: 1220, h: 2440 }, { thickness: 20, w: 1220, h: 2440 }]);
  const [offcutStock, setOffcutStock] = useState([]); // reused offcuts {label, thickness, shape:'rect'|'L', A,B,C,D}
  const [parts, setParts] = useState([{ id: "P1", label: "P1", length: 500, width: 300, thickness: 12, qty: 4 }, { id: "P2", label: "P2", length: 380, width: 250, thickness: 12, qty: 6 }, { id: "P3", label: "P3", length: 600, width: 400, thickness: 20, qty: 3 }, { id: "P4", label: "P4", length: 280, width: 180, thickness: 8, qty: 8 }]);
  const [material, setMaterial] = useState("S235JR");
  const [kerf, setKerf] = useState(3); const [margin, setMargin] = useState(10); const [reuseMin, setReuseMin] = useState(300); const [allowRotation, setAllowRotation] = useState(true); const [splicePref, setSplicePref] = useState("welds");
  const [results, setResults] = useState(null); const [colorMap, setColorMap] = useState({}); const [excelMsg, setExcelMsg] = useState(""); const [columnMap, setColumnMap] = useState(null); const [isOpt, setIsOpt] = useState(false);

  const stockThk = [...new Set(stock.map(s => s.thickness).filter(t => t > 0))].sort((a, b) => a - b);
  const sheetFor = thk => { const h = stock.find(s => +s.thickness === +thk); return h ? { w: h.w, h: h.h } : DEFAULT_SHEET; };
  const addStock = () => setStock(s => [...s, { thickness: 0, w: 1220, h: 2440 }]);
  const upStock = (i, f, v) => setStock(s => s.map((r, ri) => ri === i ? { ...r, [f]: +v || 0 } : r));
  const delStock = i => setStock(s => s.filter((_, ri) => ri !== i));
  const addOffcut = () => setOffcutStock(o => [...o, { label: `Offcut ${o.length + 1}`, thickness: stockThk[0] || 8, shape: "rect", A: 600, B: 400, C: 0, D: 0 }]);
  const upOffcut = (i, f, v) => setOffcutStock(o => o.map((r, ri) => ri === i ? { ...r, [f]: (f === "label" || f === "shape") ? v : +v || 0 } : r));
  const delOffcut = i => setOffcutStock(o => o.filter((_, ri) => ri !== i));
  const addPart = () => { const id = `P${parts.length + 1}`; setParts(p => [...p, { id, label: id, length: 400, width: 200, thickness: stockThk[0] || 10, qty: 1 }]); };
  const upPart = (i, f, v) => setParts(p => p.map((r, ri) => ri === i ? { ...r, [f]: f === "label" ? v : +v || 0 } : r));
  const delPart = i => setParts(p => p.filter((_, ri) => ri !== i));

  const norm = s => String(s == null ? "" : s).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const detectCol = (h, al) => { const H = h.map(norm); for (let i = 0; i < H.length; i++) for (const a of al) if (H[i].toLowerCase() === a.toLowerCase()) return h[i]; for (let i = 0; i < H.length; i++) for (const a of al) if (H[i].toLowerCase().includes(a.toLowerCase())) return h[i]; return null; };
  const PLATE_PFX = /^(PLATE|PLT|PL|FLT|FLAT|FL|FB|P)\s*[-_]?\s*\d/i;
  const HOT = /^(SHS|RHS|CHS|IPE|IPN|HEA|HEB|HEM|HE|UBP|UB|UC|PFC|UPN|UPE|JIS|HSS|HD|HP|UA|EA|RSA|RSJ|L|W|S|C|MC|BR|RD)/i;
  const isPlate = p => PLATE_PFX.test(String(p || "").trim());
  const isHot = p => { const s = String(p || "").trim(); if (isPlate(s)) return false; return HOT.test(s); };
  const parsePlate = raw => { const p = String(raw || "").trim().toUpperCase(); if (!p) return null; const m = p.match(/^(?:PLATE|PLT|PL|FLT|FLAT|FL|FB|P)\s*[-_]?\s*(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (m) { const a = +m[1], b = +m[2]; return { thickness: Math.min(a, b), width: Math.max(a, b) }; } if (!isHot(p)) { const b = p.match(/^(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (b) { const x = +b[1], y = +b[2]; return { thickness: Math.min(x, y), width: Math.max(x, y) }; } } return null; };

  function extract(headers, rows, map) {
    const out = []; const idx = c => c ? headers.indexOf(c) : -1; const pi = idx(map.profile), qi = idx(map.qty), li = idx(map.length), wi = idx(map.width), ti = idx(map.thickness);
    rows.forEach((row, ri) => {
      const cells = row.map(c => norm(c)); const joined = cells.join(" ");
      if (/^\s*(grand\s+)?total\b/i.test(joined) || /total\s+for|members/i.test(joined)) return;
      if (map.profile) { const pc = cells[headers.indexOf(map.profile)] || ""; if (/^total$/i.test(pc.trim())) return; }
      let thickness = 0, width = 0, length = 0, qty = 0;
      const pcell = pi >= 0 ? cells[pi] : (cells.find(c => isPlate(c) || isHot(c)) || "");
      if (pcell) { if (isHot(pcell)) return; const pl = parsePlate(pcell); if (pl) { thickness = pl.thickness; width = pl.width; } }
      if (li >= 0) length = parseFloat(cells[li]) || 0; if (qi >= 0) qty = parseInt(cells[qi]) || 0; if (ti >= 0 && !thickness) thickness = parseFloat(cells[ti]) || 0; if (wi >= 0 && !width) width = parseFloat(cells[wi]) || 0;
      const havePlate = thickness > 0 && width > 0;
      if (!havePlate && !(length > 0 && width > 0)) {
        const nums = cells.flatMap(c => { const pr = c.match(/(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)/); if (pr) return [+pr[1], +pr[2]]; const n = parseFloat(c.replace(/[^\d.]/g, "")); return isFinite(n) && c.match(/\d/) ? [n] : []; }).filter(n => n > 0 && n < 100000);
        if (nums.length >= 2) { const sorted = [...nums].sort((a, b) => a - b); if (!thickness) { const tc = sorted.find(n => n >= 3 && n <= 80); thickness = tc || 10; } const big = sorted.filter(n => n !== thickness); if (big.length >= 2) { width = width || big[0]; length = length || big[big.length - 1]; } else if (big.length === 1) { length = length || big[0]; width = width || big[0]; } }
      }
      if (!qty) { const qc = cells.find(c => /^\d{1,4}$/.test(c) && +c <= 999 && +c !== length && +c !== width && +c !== thickness); qty = qc ? +qc : 1; }
      if (!thickness) thickness = 10;
      if (length > 0 && width > 0) out.push({ id: `P${ri + 1}`, label: `P${ri + 1}`, length: Math.round(length * 10) / 10, width: Math.round(width * 10) / 10, thickness, qty: qty || 1 });
    });
    return out;
  }
  const parseExcel = useCallback(file => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" }); const ws = wb.Sheets[wb.SheetNames[0]]; const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let hr = -1; for (let i = 0; i < Math.min(raw.length, 25); i++) { const r = raw[i].map(c => norm(c).toLowerCase()); if (r.some(c => /profile|length|width|qty|no\.?$|thick|partpos|part pos/i.test(c))) { hr = i; break; } } if (hr === -1) hr = 0;
        const headers = raw[hr].map(c => norm(c)); const rows = raw.slice(hr + 1).filter(r => r.some(c => c !== ""));
        const map = { profile: detectCol(headers, ["profile", "section", "size"]), qty: detectCol(headers, ["no.", "no", "qty", "number", "count", "quantity", "pcs"]), length: detectCol(headers, ["length (mm)", "length(mm)", "length", "l (mm)"]), width: detectCol(headers, ["width (mm)", "width(mm)", "width", "breadth"]), thickness: detectCol(headers, ["thickness", "thk", "t (mm)"]) };
        const out = extract(headers, rows, map);
        if (out.length) { setParts(out); setExcelMsg(`✓ Imported ${out.length} plate part type${out.length > 1 ? "s" : ""}. Hot-rolled sections & subtotal rows ignored.`); setColumnMap(null); }
        else { setColumnMap({ headers, map, rows }); setExcelMsg("⚠ Couldn't auto-detect columns. Map them below."); }
      } catch { setExcelMsg("⚠ Could not read this file."); }
    };
    reader.readAsBinaryString(file);
  }, []);
  const onDrop = useCallback(e => { e.preventDefault(); const f = e.dataTransfer?.files?.[0] || e.target?.files?.[0]; if (f) parseExcel(f); }, [parseExcel]);

  const optimize = () => {
    setIsOpt(true);
    setTimeout(() => {
      const valid = parts.filter(p => p.length > 0 && p.width > 0 && p.qty > 0);
      const cm = {}; valid.forEach((p, i) => cm[p.id] = PART_COLORS[i % PART_COLORS.length]); setColorMap(cm);
      const byThk = {}; valid.forEach(p => { const k = p.thickness || 10; (byThk[k] = byThk[k] || []).push(p); });
      const groups = Object.keys(byThk).map(Number).sort((a, b) => b - a).map(thk => {
        let gp = byThk[thk]; const { w: sw, h: sh } = sheetFor(thk);
        // 1) Use reused offcuts of this thickness FIRST, so new sheets are the minimum needed.
        const myOffcuts = offcutStock.filter(o => +o.thickness === +thk && o.A > 0 && o.B > 0);
        const offcutPlans = [];
        myOffcuts.forEach(o => {
          const rect = o.shape === "L" ? lUsableRect(o.A, o.B, o.C, o.D) : { w: o.A, h: o.B };
          if (rect.w <= 0 || rect.h <= 0) return;
          const res = nestIntoRect(rect.w, rect.h, gp, kerf, allowRotation);
          if (res.placedAny) { offcutPlans.push({ label: o.label, shape: o.shape, rectW: rect.w, rectH: rect.h, placements: res.placements }); gp = res.leftover; }
        });
        // 2) Whatever remains goes onto NEW standard sheets (the minimum).
        const gsheets = nestPlates(sw, sh, gp, kerf, margin, allowRotation, reuseMin, splicePref);
        const usedArea = gsheets.reduce((s, st) => s + (st.placements || []).reduce((a, p) => a + p.w * p.h, 0), 0);
        const totalArea = gsheets.length * sw * sh; const wastePct = gsheets.length ? Math.round(((totalArea - usedArea) / totalArea) * 100) : 0;
        const allThkParts = byThk[thk];
        const partVol = allThkParts.reduce((s, p) => s + p.length * p.width * thk * p.qty, 0); const partWeight = (partVol / 1e9) * DENSITY; const sheetWeight = (gsheets.length * sw * sh * thk / 1e9) * DENSITY;
        const offcuts = []; gsheets.forEach((st, si) => (st.offcuts || []).forEach(o => { const area = o.shape === "L" ? (o.A * o.B - o.notchW * o.notchH) : (o.A * o.B); offcuts.push({ sheet: si + 1, ...o, area, weight: (area * thk / 1e9) * DENSITY }); }));
        const splices = (gsheets.splices || []).map(s => ({ label: s.label, W: s.W, L: s.L, count: s.count, qty: s.qty }));
        const reusedUsedCount = offcutPlans.length;
        return { thickness: thk, sw, sh, sheets: gsheets, sheetCount: gsheets.length, partCount: allThkParts.reduce((s, p) => s + p.qty, 0), usedArea, totalArea, wastePct, utilPct: gsheets.length ? 100 - wastePct : 100, partWeight, sheetWeight, wasteWeight: sheetWeight - partWeight, offcuts, offcutWeight: offcuts.reduce((s, o) => s + o.weight, 0), splices, marginUsed: gsheets.marginUsed, offcutPlans, reusedUsedCount };
      });
      const totals = groups.reduce((a, g) => ({ sheets: a.sheets + g.sheetCount, parts: a.parts + g.partCount, sheetWeight: a.sheetWeight + g.sheetWeight, partWeight: a.partWeight + g.partWeight, wasteWeight: a.wasteWeight + g.wasteWeight, totalArea: a.totalArea + g.totalArea, usedArea: a.usedArea + g.usedArea, offcutCount: a.offcutCount + g.offcuts.length, offcutWeight: a.offcutWeight + g.offcutWeight, spliceCount: a.spliceCount + g.splices.reduce((ss, s) => ss + s.qty, 0), reusedUsed: a.reusedUsed + g.reusedUsedCount }), { sheets: 0, parts: 0, sheetWeight: 0, partWeight: 0, wasteWeight: 0, totalArea: 0, usedArea: 0, offcutCount: 0, offcutWeight: 0, spliceCount: 0, reusedUsed: 0 });
      totals.utilPct = Math.round((totals.usedArea / totals.totalArea) * 100); totals.wastePct = 100 - totals.utilPct;
      setResults({ groups, totals }); setIsOpt(false);
    }, 350);
  };

  if (results) return <PlateResults results={results} colorMap={colorMap} material={material} reuseMin={reuseMin} onBack={() => setResults(null)} />;

  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}><SectionTitle>Steel Plates — Workspace</SectionTitle><BackChip onClick={onBack} label="↩ Modules" /></div>
      {inputMode === null && (
        <div>
          <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 15, marginBottom: 18 }}>How would you like to enter your cutting list?</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 20 }}>
            <Chooser icon="✏️" title="Manual Entry" desc="Type parts into a simple table." onClick={() => setInputMode("manual")} />
            <Chooser icon="📊" title="Upload Tekla / Excel" desc="Drop a material list. Plates read, sections ignored, messy sheets cleaned." onClick={() => setInputMode("excel")} />
          </div>
        </div>
      )}
      {inputMode !== null && (
        <Card title="📐 Stock & Cutting Settings">
          <Label>Standard stock sheets — set a sheet size per thickness</Label>
          <div style={{ overflowX: "auto", marginTop: 6 }}>
            <table style={{ borderCollapse: "collapse", fontSize: 13 }}><thead><tr style={{ borderBottom: "1px solid #2d3748" }}>{["Thickness (mm)", "Width (mm)", "Length (mm)", ""].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: "#64748b", fontWeight: 600, fontSize: 11, letterSpacing: 1 }}>{h}</th>)}</tr></thead>
              <tbody>{stock.map((s, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "5px 10px" }}><input type="number" value={s.thickness} onChange={e => upStock(i, "thickness", e.target.value)} style={{ ...CI, width: 90, borderColor: "#d9770699" }} /></td><td style={{ padding: "5px 10px" }}><input type="number" value={s.w} onChange={e => upStock(i, "w", e.target.value)} style={{ ...CI, width: 90 }} /></td><td style={{ padding: "5px 10px" }}><input type="number" value={s.h} onChange={e => upStock(i, "h", e.target.value)} style={{ ...CI, width: 90 }} /></td><td style={{ padding: "5px 10px" }}><button onClick={() => delStock(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 14 }}>✕</button></td></tr>)}</tbody>
            </table>
          </div>
          <button onClick={addStock} style={{ marginTop: 10, padding: "7px 16px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 12 }}>+ Add Thickness</button>
          <div style={{ marginTop: 8, fontSize: 11, color: "#94a3b8", fontFamily: "'Space Mono', monospace" }}>If parts are longer than 2440 mm, set a bigger sheet here (e.g. 1500×3000 or 2000×6000) for that thickness — otherwise those parts can't be nested.</div>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-end", marginTop: 18, paddingTop: 16, borderTop: "1px solid #1a2230" }}>
            <div><Label>Material Grade</Label><select value={material} onChange={e => setMaterial(e.target.value)} style={SEL}>{STEEL_GRADES.map(m => <option key={m}>{m}</option>)}</select></div>
            <div><Label>Cutting Gap / Kerf (mm)</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div>
            <div><Label>Edge Margin (mm)</Label><input type="number" value={margin} onChange={e => setMargin(+e.target.value)} style={{ ...IN, width: 80 }} /></div>
            <div><Label>Reusable Offcut Min (mm)</Label><input type="number" value={reuseMin} onChange={e => setReuseMin(+e.target.value)} style={{ ...IN, width: 100 }} /></div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, paddingBottom: 6 }}><input type="checkbox" id="rot" checked={allowRotation} onChange={e => setAllowRotation(e.target.checked)} style={{ width: 16, height: 16, accentColor: "#f59e0b" }} /><label htmlFor="rot" style={{ color: "#94a3b8", fontSize: 13 }}>Allow 90° Rotation</label></div>
            <div><Label>Splice strategy (parts bigger than sheet)</Label><select value={splicePref} onChange={e => setSplicePref(e.target.value)} style={{ ...SEL, minWidth: 220 }}><option value="welds">Fewest welds (bigger pieces)</option><option value="pack">Best material use (more, smaller pieces)</option></select></div>
          </div>
        </Card>
      )}
      {inputMode !== null && (
        <Card title="♻ Reused Offcuts from Previous Projects (optional)" style={{ marginTop: 20 }}>
          <div style={{ fontSize: 12, color: "#6ee7b7", marginBottom: 12, lineHeight: 1.6 }}>Add leftover plates you already have. They are used <b>first</b>, so new standard sheets are only the <b>minimum</b> you still need to buy. Pick the shape and enter the labeled dimensions — the diagram shows exactly what each letter means.</div>
          {offcutStock.length === 0 && <div style={{ fontSize: 12, color: "#64748b", fontStyle: "italic", marginBottom: 12 }}>No reused offcuts added. Click below to add one.</div>}
          {offcutStock.map((o, i) => { const use = o.shape === "L" ? lUsableRect(o.A, o.B, o.C, o.D) : { w: o.A, h: o.B }; return (
            <div key={i} style={{ display: "flex", gap: 18, flexWrap: "wrap", alignItems: "flex-start", padding: "14px 0", borderTop: i > 0 ? "1px solid #1a2230" : "none" }}>
              <div style={{ flex: "1 1 320px", minWidth: 280 }}>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 10 }}>
                  <input value={o.label} onChange={e => upOffcut(i, "label", e.target.value)} style={{ ...CI, width: 150 }} />
                  <select value={o.shape} onChange={e => upOffcut(i, "shape", e.target.value)} style={{ ...CI, width: 130, cursor: "pointer" }}><option value="rect">Rectangle</option><option value="L">L-shape</option></select>
                  <button onClick={() => delOffcut(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 14 }}>✕ remove</button>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <div><Label>Thickness (mm)</Label><input type="number" value={o.thickness} onChange={e => upOffcut(i, "thickness", e.target.value)} style={{ ...CI, width: 80, borderColor: "#d9770699" }} /></div>
                  <div><Label>A — width (mm)</Label><input type="number" value={o.A} onChange={e => upOffcut(i, "A", e.target.value)} style={{ ...CI, width: 80 }} /></div>
                  <div><Label>B — length (mm)</Label><input type="number" value={o.B} onChange={e => upOffcut(i, "B", e.target.value)} style={{ ...CI, width: 80 }} /></div>
                  {o.shape === "L" && <div><Label>C — notch width</Label><input type="number" value={o.C} onChange={e => upOffcut(i, "C", e.target.value)} style={{ ...CI, width: 80, borderColor: "#3b82f699" }} /></div>}
                  {o.shape === "L" && <div><Label>D — notch length</Label><input type="number" value={o.D} onChange={e => upOffcut(i, "D", e.target.value)} style={{ ...CI, width: 80, borderColor: "#3b82f699" }} /></div>}
                </div>
                <div style={{ marginTop: 10, fontSize: 11, color: "#6ee7b7", fontFamily: "'Space Mono', monospace" }}>Usable rectangle fed to nesting: {Math.round(use.w)} × {Math.round(use.h)} mm @ {o.thickness} mm</div>
              </div>
              <LShapeDiagram A={o.A} B={o.B} C={o.shape === "L" ? o.C : 0} D={o.shape === "L" ? o.D : 0} t={o.thickness} />
            </div>
          ); })}
          <button onClick={addOffcut} style={{ marginTop: 12, padding: "8px 18px", background: "rgba(16,185,129,.15)", border: "1px dashed #10b981", color: "#6ee7b7", borderRadius: 4, cursor: "pointer", fontSize: 12 }}>♻ Add Reused Offcut</button>
        </Card>
      )}
      {inputMode === "manual" && (
        <Card title="✏️ Manual Part Entry" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} label="↺ Change method" />}>
          <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}><thead><tr style={{ borderBottom: "1px solid #2d3748" }}>{["ID", "Length (mm)", "Width (mm)", "Thickness (mm)", "Qty", ""].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 8px", color: "#64748b", fontWeight: 600, fontSize: 11, letterSpacing: 1 }}>{h}</th>)}</tr></thead>
            <tbody>{parts.map((p, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "5px 4px" }}><input value={p.label} onChange={e => upPart(i, "label", e.target.value)} style={{ ...CI, width: 50 }} /></td><td style={{ padding: "5px 4px" }}><input type="number" value={p.length} onChange={e => upPart(i, "length", e.target.value)} style={{ ...CI, width: 80 }} /></td><td style={{ padding: "5px 4px" }}><input type="number" value={p.width} onChange={e => upPart(i, "width", e.target.value)} style={{ ...CI, width: 80 }} /></td><td style={{ padding: "5px 4px" }}><select value={p.thickness} onChange={e => upPart(i, "thickness", e.target.value)} style={{ ...CI, width: 90, borderColor: "#d9770699", cursor: "pointer" }}>{!stockThk.includes(p.thickness) && <option value={p.thickness}>{p.thickness} (not in stock)</option>}{stockThk.map(t => <option key={t} value={t}>{t} mm</option>)}</select></td><td style={{ padding: "5px 4px" }}><input type="number" value={p.qty} onChange={e => upPart(i, "qty", e.target.value)} style={{ ...CI, width: 55 }} /></td><td style={{ padding: "5px 4px" }}><button onClick={() => delPart(i)} style={{ background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 14 }}>✕</button></td></tr>)}</tbody>
          </table></div>
          <button onClick={addPart} style={{ marginTop: 12, padding: "8px 18px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 12, width: "100%" }}>+ Add Part</button>
        </Card>
      )}
      {inputMode === "excel" && (
        <Card title="📊 Upload Tekla / Excel" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} label="↺ Change method" />}>
          <div style={{ marginBottom: 14, padding: "10px 14px", borderRadius: 6, background: "rgba(245,158,11,.08)", border: "1px solid rgba(245,158,11,.25)", color: "#fcd34d", fontSize: 12, lineHeight: 1.6 }}>💡 Thickness comes from the file; the matching stock sheet size is taken from your list above. Unknown thickness falls back to 1220×2440.</div>
          <div onDrop={onDrop} onDragOver={e => e.preventDefault()} onClick={() => document.getElementById("pxlsx").click()} style={{ border: "2px dashed #2d3748", borderRadius: 8, padding: "48px 20px", textAlign: "center", background: "rgba(15,19,24,.5)", cursor: "pointer" }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}>📋</div><div style={{ color: "#94a3b8", fontSize: 15, marginBottom: 6 }}>Drop Tekla list / Excel / CSV / TXT</div><div style={{ color: "#475569", fontSize: 12 }}>or click to browse</div><div style={{ marginTop: 14, color: "#3a4452", fontSize: 10, lineHeight: 1.7 }}>Reads Profile column (PLT/FLT/PL → thickness × width). Hot-rolled ignored.</div>
          </div>
          <input id="pxlsx" type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={onDrop} />
          {excelMsg && <div style={{ marginTop: 10, padding: "8px 12px", borderRadius: 4, background: excelMsg.startsWith("✓") ? "rgba(16,185,129,.15)" : "rgba(245,158,11,.15)", border: `1px solid ${excelMsg.startsWith("✓") ? "rgba(16,185,129,.4)" : "rgba(245,158,11,.4)"}`, color: excelMsg.startsWith("✓") ? "#6ee7b7" : "#fcd34d", fontSize: 12 }}>{excelMsg}</div>}
          {columnMap && <div style={{ marginTop: 12, padding: 12, background: "rgba(15,19,24,.8)", borderRadius: 6, border: "1px solid #2d3748" }}><div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 8 }}>Map columns:</div>{["profile", "qty", "length", "width", "thickness"].map(f => <div key={f} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}><span style={{ fontSize: 11, color: "#64748b", width: 70 }}>{f}</span><select value={columnMap.map[f] || ""} onChange={e => setColumnMap(c => ({ ...c, map: { ...c.map, [f]: e.target.value } }))} style={{ ...SEL, flex: 1 }}><option value="">—</option>{columnMap.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></div>)}<button onClick={() => { const out = extract(columnMap.headers, columnMap.rows, columnMap.map); if (out.length) { setParts(out); setExcelMsg(`✓ Imported ${out.length} plate parts`); setColumnMap(null); } else setExcelMsg("⚠ Still couldn't extract."); }} style={BTN}>Apply Mapping</button></div>}
          {parts.length > 0 && <div style={{ marginTop: 14 }}><div style={{ fontSize: 11, color: "#475569", marginBottom: 6, letterSpacing: 1 }}>PARTS READY ({parts.length})</div><div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>{parts.slice(0, 16).map((p, i) => <div key={i} style={{ padding: "3px 8px", borderRadius: 3, background: (colorMap[p.id] || PART_COLORS[i % PART_COLORS.length]) + "22", border: `1px solid ${(colorMap[p.id] || PART_COLORS[i % PART_COLORS.length])}55`, fontSize: 10, color: "#94a3b8", fontFamily: "'Space Mono', monospace" }}>{p.label}: {p.length}×{p.width}×{p.thickness} ×{p.qty}</div>)}{parts.length > 16 && <div style={{ fontSize: 10, color: "#475569", alignSelf: "center" }}>+{parts.length - 16} more</div>}</div></div>}
        </Card>
      )}
      {inputMode !== null && <div style={{ textAlign: "center", marginTop: 32 }}><button onClick={optimize} disabled={isOpt} style={OPT_BTN(isOpt)}>{isOpt ? "⚙ Optimizing…" : "⚡ OPTIMIZE PLATES"}</button></div>}
    </>
  );
}

function PlateResults({ results, colorMap, material, reuseMin, onBack }) {
  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}><SectionTitle>Plate Results</SectionTitle><BackChip onClick={onBack} label="← Back to input" /></div>
      {results.totals.spliceCount > 0 && (
        <div style={{ marginBottom: 20, padding: "16px 18px", background: "rgba(245,158,11,.1)", border: "1px solid rgba(245,158,11,.45)", borderRadius: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#fbbf24", marginBottom: 8, fontFamily: "'Space Mono', monospace" }}>⚙ {results.totals.spliceCount} part{results.totals.spliceCount > 1 ? "s are" : " is"} larger than the sheet — cut &amp; welded from multiple pieces</div>
          <div style={{ fontSize: 13, color: "#fcd34d", lineHeight: 1.6 }}>Each is split across sheets and joined with a welded seam (no bigger sheet needed). <b>Confirm a welded seam is acceptable for these pieces</b> before fabrication:</div>
          <div style={{ marginTop: 10, display: "flex", flexWrap: "wrap", gap: 6 }}>{results.groups.flatMap(g => g.splices.map((s, k) => <span key={`${g.thickness}-${k}`} style={{ padding: "4px 10px", borderRadius: 4, background: "rgba(245,158,11,.15)", border: "1px solid rgba(245,158,11,.4)", color: "#fcd34d", fontSize: 11, fontFamily: "'Space Mono', monospace" }}>{s.label}: {Math.round(s.W)}×{Math.round(s.L)}mm @ {g.thickness}mm → {s.count} pcs{s.qty > 1 ? ` ×${s.qty}` : ""}</span>))}</div>
        </div>
      )}
      {results.totals.reusedUsed > 0 && (
        <div style={{ marginBottom: 20, padding: "14px 18px", background: "rgba(16,185,129,.1)", border: "1px solid rgba(16,185,129,.4)", borderRadius: 10, fontSize: 13, color: "#6ee7b7", lineHeight: 1.6 }}>
          ♻ {results.totals.reusedUsed} reused offcut{results.totals.reusedUsed > 1 ? "s" : ""} from previous projects {results.totals.reusedUsed > 1 ? "were" : "was"} used first. The sheet counts below are the <b>minimum new standard sheets</b> you still need to buy.
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12, marginBottom: 24 }}>
        {[{ l: "Total Parts", v: results.totals.parts, i: "◼" }, { l: "Total Sheets", v: results.totals.sheets, i: "📄", a: "#3B82F6" }, { l: "Thickness Groups", v: results.groups.length, i: "≡", a: "#8B5CF6" }, { l: "Utilization", v: `${results.totals.utilPct}%`, i: "📊", a: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" }, { l: "Total Purchase Weight", v: `${fmtTon(results.totals.sheetWeight)} t`, i: "🧾", a: "#f59e0b" }, { l: "Net Parts Weight", v: `${fmtTon(results.totals.partWeight)} t`, i: "⚖" }, { l: "Scrap / Offcut Weight", v: `${fmtTon(results.totals.wasteWeight)} t`, i: "🗑", a: "#EF4444" }, { l: "Reusable Offcuts", v: `${results.totals.offcutCount} (${fmtTon(results.totals.offcutWeight)} t)`, i: "♻", a: "#10B981" }].map(c => <StatCard key={c.l} {...c} />)}
      </div>
      {results.groups.map((g, gi) => <Card key={gi} title={`🗂 Nesting Layout — ${g.thickness} mm Plate (${g.sheetCount} sheet${g.sheetCount > 1 ? "s" : ""})`} style={{ marginBottom: 20 }}><PlateCanvas sheets={g.sheets} sheetW={g.sw} sheetH={g.sh} colorMap={colorMap} thickness={g.thickness} /></Card>)}
      {results.totals.offcutCount > 0 && (
        <Card title="♻ Reusable Offcut Register" style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 12, color: "#6ee7b7", marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 6, padding: "8px 12px" }}>{results.totals.offcutCount} reusable leftover{results.totals.offcutCount > 1 ? "s" : ""} ≈ {fmtTon(results.totals.offcutWeight)} t to keep for future jobs (≥ {reuseMin} mm/side). For L-shaped leftovers, <b>USABLE</b> is the largest rectangle you can cut from it.</div>
          <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(16,185,129,.12)", borderBottom: "1px solid rgba(16,185,129,.3)" }}>{["THICKNESS", "FROM SHEET", "SHAPE", "OVERALL", "USABLE RECT", "WEIGHT", "STATUS"].map((h, i) => <th key={h} style={{ padding: "10px 12px", textAlign: i > 4 ? "right" : "left", color: "#6ee7b7", fontSize: 11, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
            <tbody>{results.groups.flatMap(g => g.offcuts.map((o, oi) => <tr key={`${g.thickness}-${oi}`} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{g.thickness} mm</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>Sheet {o.sheet}</td><td style={{ padding: "9px 12px" }}>{o.shape === "L" ? <span style={{ color: "#fbbf24", fontWeight: 700 }}>⌐ L-shape</span> : <span style={{ color: "#94a3b8" }}>▭ Rect</span>}</td><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{o.shape === "L" ? `${o.A}×${o.B} − notch ${o.notchW}×${o.notchH}` : `${o.A} × ${o.B} mm`}</td><td style={{ padding: "9px 12px", color: "#6ee7b7", fontWeight: 700 }}>{o.w} × {o.h} mm</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#94a3b8" }}>{o.weight >= 1 ? `${o.weight.toFixed(1)} kg` : `${(o.weight * 1000).toFixed(0)} g`}</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#10B981", fontWeight: 700 }}>✓ keep</td></tr>))}</tbody>
          </table></div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 14 }}>
            <button onClick={() => exportOffcutPDF(results, material, reuseMin)} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>📄 Export Offcut PDF</button>
            <button onClick={() => exportOffcutExcel(results, material, reuseMin)} style={{ ...EXP, background: "linear-gradient(135deg,#10b981,#059669)", color: "#04140d" }}>📊 Export Offcut Excel</button>
          </div>
        </Card>
      )}
      <Card title="💾 Export" style={{ marginBottom: 28 }}><div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}><button onClick={() => exportPlatePDF(results, material, reuseMin)} style={EXP}>📄 Export PDF Report</button><button onClick={() => exportPlateExcel(results, material, reuseMin)} style={EXP}>📊 Export Excel Summary</button></div></Card>
      <ProcurementBlock results={results} material={material} />
    </>
  );
}

function ProcurementBlock({ results, material }) {
  return (
    <div style={{ border: "2px solid rgba(245,158,11,.35)", borderRadius: 12, overflow: "hidden", boxShadow: "0 0 60px rgba(245,158,11,.08)" }}>
      <div style={{ background: "linear-gradient(135deg,rgba(245,158,11,.25),rgba(217,119,6,.2))", borderBottom: "1px solid rgba(245,158,11,.3)", padding: "20px 28px", display: "flex", alignItems: "center", gap: 14 }}><span style={{ fontSize: 28 }}>📦</span><div><div style={{ fontSize: 11, letterSpacing: 3, color: "#fbbf24", textTransform: "uppercase", marginBottom: 3, fontFamily: "'Space Mono', monospace" }}>Final Decision-Ready Output</div><div style={{ fontSize: 22, fontWeight: 700, color: "#f8fafc", fontFamily: "'Playfair Display', serif" }}>Material Procurement Summary</div></div><div style={{ marginLeft: "auto", textAlign: "right" }}><div style={{ fontSize: 10, color: "#94a3b8", letterSpacing: 1, fontFamily: "'Space Mono', monospace" }}>MATERIAL</div><div style={{ fontSize: 16, color: "#f59e0b", fontWeight: 700 }}>{material}</div></div></div>
      <div style={{ padding: "26px 28px", background: "rgba(10,13,18,.92)" }}>
        <div style={{ marginBottom: 22, padding: "18px 20px", background: "rgba(245,158,11,.08)", border: "1px solid rgba(245,158,11,.3)", borderRadius: 10 }}>
          <div style={{ fontSize: 11, letterSpacing: 2, color: "#fbbf24", textTransform: "uppercase", marginBottom: 12, fontFamily: "'Space Mono', monospace" }}>What you need to buy</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>{results.groups.map((g, gi) => <div key={gi} style={{ display: "flex", alignItems: "center", gap: 12, fontSize: 17, color: "#f8fafc", fontFamily: "'Playfair Display', serif" }}><span style={{ color: "#f59e0b", fontSize: 18 }}>✓</span><span>You need <b style={{ color: "#fbbf24", fontSize: 20 }}>{g.sheetCount}</b> sheet{g.sheetCount > 1 ? "s" : ""} of <b style={{ color: "#f59e0b" }}>{g.thickness} mm</b> <span style={{ color: "#94a3b8", fontSize: 14, fontFamily: "'Space Mono', monospace" }}>({g.sw} × {g.sh} mm, {material})</span></span></div>)}</div>
        </div>
        <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(245,158,11,.15)", borderBottom: "2px solid rgba(245,158,11,.3)" }}>{["THICKNESS", "SHEET SIZE", "SHEETS REQ.", "PURCHASE WT (full sheets)", "PARTS", "UTILIZATION", "NET PARTS WT", "SCRAP WT"].map((h, i) => <th key={h} style={{ padding: "13px 14px", textAlign: i < 2 ? "left" : i === 7 ? "right" : "center", color: "#fbbf24", fontSize: 11, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
          <tbody>{results.groups.map((g, gi) => <tr key={gi} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "15px 14px" }}><span style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "rgba(245,158,11,.12)", border: "1px solid rgba(245,158,11,.3)", borderRadius: 4, padding: "4px 12px" }}><span style={{ color: "#f59e0b" }}>✓</span><span style={{ color: "#f8fafc", fontWeight: 700, fontSize: 15 }}>{g.thickness} mm</span></span></td><td style={{ padding: "15px 14px", color: "#94a3b8" }}>{g.sw} × {g.sh} mm</td><td style={{ padding: "15px 14px", textAlign: "center" }}><span style={{ display: "inline-block", background: "rgba(59,130,246,.18)", border: "1px solid rgba(59,130,246,.4)", borderRadius: 6, padding: "6px 16px", color: "#93c5fd", fontWeight: 800, fontSize: 18 }}>{g.sheetCount} sheet{g.sheetCount > 1 ? "s" : ""}</span></td><td style={{ padding: "15px 14px", textAlign: "center", color: "#f59e0b", fontWeight: 700 }}>{fmtTon(g.sheetWeight)} t</td><td style={{ padding: "15px 14px", textAlign: "center", color: "#94a3b8" }}>{g.partCount}</td><td style={{ padding: "15px 14px", textAlign: "center", fontWeight: 700, color: g.utilPct > 80 ? "#10B981" : g.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{g.utilPct}%</td><td style={{ padding: "15px 14px", textAlign: "center", color: "#94a3b8" }}>{fmtTon(g.partWeight)} t</td><td style={{ padding: "15px 14px", textAlign: "right", color: "#f08a8a" }}>{fmtTon(g.wasteWeight)} t</td></tr>)}
            <tr style={{ background: "rgba(245,158,11,.12)", borderTop: "2px solid rgba(245,158,11,.4)" }}><td style={{ padding: "16px 14px", fontWeight: 800, color: "#f8fafc" }}>TOTAL</td><td style={{ padding: "16px 14px", color: "#94a3b8" }}>{results.groups.length} thickness{results.groups.length > 1 ? "es" : ""}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, fontSize: 18, color: "#93c5fd" }}>{results.totals.sheets} sheet{results.totals.sheets > 1 ? "s" : ""}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#f59e0b" }}>{fmtTon(results.totals.sheetWeight)} t</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 700, color: "#f8fafc" }}>{results.totals.parts}</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: results.totals.utilPct > 80 ? "#10B981" : results.totals.utilPct > 60 ? "#F59E0B" : "#EF4444" }}>{results.totals.utilPct}%</td><td style={{ padding: "16px 14px", textAlign: "center", fontWeight: 800, color: "#f8fafc" }}>{fmtTon(results.totals.partWeight)} t</td><td style={{ padding: "16px 14px", textAlign: "right", fontWeight: 800, color: "#f08a8a" }}>{fmtTon(results.totals.wasteWeight)} t</td></tr>
          </tbody></table></div>
        <div style={{ marginTop: 12, padding: "10px 14px", background: "rgba(15,19,24,.5)", border: "1px solid #1a2230", borderRadius: 6, fontSize: 11, color: "#94a3b8", lineHeight: 1.7, fontFamily: "'Space Mono', monospace" }}><span style={{ color: "#f59e0b" }}>● Purchase Wt</span> = gross full sheets bought · <span style={{ color: "#10B981" }}>● Net Parts Wt</span> = finished cut plates · <span style={{ color: "#f08a8a" }}>● Scrap Wt</span> = offcut (Purchase − Net)</div>
        <div style={{ marginTop: 22, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 16 }}>
          <BigTon label="Total Purchase Weight" value={fmtTon(results.totals.sheetWeight)} note={`gross — all ${results.totals.sheets} full sheets`} color="#f59e0b" big />
          <BigTon label="Net Parts Weight" value={fmtTon(results.totals.partWeight)} note="finished cut plates only" color="#10B981" />
          <BigTon label="Scrap / Offcut Weight" value={fmtTon(results.totals.wasteWeight)} note={`${results.totals.wastePct}% of purchase`} color="#EF4444" />
        </div>
        <div style={{ marginTop: 18, textAlign: "right", fontFamily: "'Space Mono', monospace", fontSize: 10, color: "#475569" }}>Generated by Steel Optimizer · {new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })} · {material}</div>
      </div>
    </div>
  );
}

/* ─── SECTIONS MODULE ────────────────────────────────────────────────────── */
function SectionsModule({ onBack }) {
  const [inputMode, setInputMode] = useState(null);
  const DEFAULT_STOCK = 12000;
  const [rows, setRows] = useState([{ id: genId(), profile: "", grade: "S355", lengths: "", stock: "" }]);
  const [activeAuto, setActiveAuto] = useState(null);
  const [cutList, setCutList] = useState([]); const [kerf, setKerf] = useState(3); const [results, setResults] = useState(null); const [error, setError] = useState(""); const [uploadInfo, setUploadInfo] = useState(null);
  const fileRef = useRef(null);

  const upRow = (id, f, v) => setRows(p => p.map(r => r.id === id ? { ...r, [f]: v } : r));
  const addRow = () => setRows(p => [...p, { id: genId(), profile: "", grade: "S355", lengths: "", stock: "" }]);
  const delRow = id => setRows(p => p.filter(r => r.id !== id));
  const parseLengths = raw => { if (!raw) return []; const toks = String(raw).trim().split(/[\s,;]+/).filter(Boolean); const out = []; for (const t of toks) { const m = t.match(/^(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+)$/i); if (m) { const len = +m[1], q = +m[2]; if (len > 0 && q > 0) for (let i = 0; i < q; i++) out.push(len); } else { const len = parseFloat(t); if (len > 0) out.push(len); } if (out.length >= 200) break; } return out.slice(0, 200); };
  const buildManual = () => { const list = []; rows.forEach(r => { const lens = parseLengths(r.lengths); const stock = parseInt(r.stock) > 0 ? parseInt(r.stock) : DEFAULT_STOCK; const counts = {}; lens.forEach(l => counts[l] = (counts[l] || 0) + 1); Object.entries(counts).forEach(([len, qty]) => list.push({ profile: (r.profile || "(unspecified)").toUpperCase(), grade: r.grade || "", length: +len, qty, stock })); }); return list; };

  const SECTION_RX = /\b(?:IPE|IPN|HEA|HEB|HEM|HE|UBP|UB|UC|HW|HM|HN|RHS|SHS|CHS|UPN|UPE|PFC|PIPE)\s?\d+(?:\s?[x×*]\s?\d+(?:\.\d+)?)*\b|\b\d+x\d+x\d+\s?(?:UB|UC|UBP|PFC)\b|\b(?:HW|HM|HN|C|L|I|H|LC)\s?\d+[x×*]\d+(?:[x×*]\d+(?:\.\d+)?){0,2}\b|\bW\d+x\d+\b|\b(?:GOST|GB)\s?[ICKB]?\s?\d+(?:\.\d+)?[a-cAB12]?\b/i;
  const norm = s => String(s == null ? "" : s).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const looksProfile = c => { if (!c) return false; if (isExcludedProfile(c)) return false; if (findSection(c)) return true; return SECTION_RX.test(c); };
  const num = v => { if (v == null || v === "") return null; let s = String(v).trim().replace(/\s/g, ""); if (s.includes(",") && s.includes(".")) { if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", "."); else s = s.replace(/,/g, ""); } else if (s.includes(",")) s = /,\d{1,2}$/.test(s) ? s.replace(",", ".") : s.replace(/,/g, ""); const n = parseFloat(s); return isFinite(n) ? n : null; };
  function autoRead(wb) {
    const out = []; let skipped = 0;
    wb.SheetNames.forEach(name => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: "" });
      let hr = -1, cm = null, inM = false;
      for (let i = 0; i < Math.min(data.length, 30); i++) { const cells = (data[i] || []).map(c => norm(c).toLowerCase()); const f = re => cells.findIndex(c => re.test(c)); const pi = f(/profile|section|member|designation/), li = f(/length/); if (pi !== -1 && li !== -1) { hr = i; const lh = cells[li]; inM = /\(\s*m\s*\)|\[\s*m\s*\]/.test(lh) && !/mm/.test(lh); cm = { profile: pi, grade: f(/grade|material/), qty: f(/\bqty\b|quantity|\bnos?\b|count|number|pcs/), length: li }; break; } }
      const toMm = v => v == null ? null : (inM ? v * 1000 : v);
      if (cm) {
        for (let i = hr + 1; i < data.length; i++) { const r = data[i] || []; const rawP = String(r[cm.profile] == null ? "" : r[cm.profile]).trim(); if (!rawP || /^total/i.test(rawP)) continue; if (isExcludedProfile(rawP)) { skipped++; continue; } if (!looksProfile(rawP)) continue; const length = toMm(cm.length !== -1 ? num(r[cm.length]) : null); if (!length || length < 5 || length > 40000) continue; const qty = (cm.qty !== -1 && num(r[cm.qty])) ? Math.round(num(r[cm.qty])) : 1; const grade = cm.grade !== -1 ? String(r[cm.grade] || "").trim() : ""; out.push({ profile: rawP.toUpperCase(), grade: grade.toUpperCase(), length, qty }); }
      } else {
        for (const r of data) { if (!r || !r.length) continue; const cells = r.map(c => c == null ? "" : String(c).trim()); if (cells.some(c => /^total/i.test(c))) continue; let profile = ""; for (const c of cells) { if (!profile) { if (isExcludedProfile(c)) { profile = "__PLATE__"; break; } const sm = c.match(SECTION_RX); if (sm) profile = sm[0].replace(/\s+/g, " ").toUpperCase(); } } if (profile === "__PLATE__") { skipped++; continue; } const lengths = [], smalls = []; cells.forEach(c => { const m = c.replace(/,/g, "").match(/-?\d+(\.\d+)?/g); if (m) m.forEach(x => { const v = parseFloat(x); if (v >= 200 && v <= 40000) lengths.push(v); else if (Number.isInteger(v) && v >= 1 && v <= 999) smalls.push(v); }); }); if (lengths.length && profile) out.push({ profile, grade: "", length: lengths[lengths.length - 1], qty: smalls.length ? smalls[smalls.length - 1] : 1 }); }
      }
    });
    const merged = {}; out.forEach(r => { const k = `${r.profile}||${r.grade}||${r.length}`; if (!merged[k]) merged[k] = { ...r }; else merged[k].qty += r.qty; }); const res = Object.values(merged); res._skipped = skipped; return res;
  }
  const handleFile = async f => { setError(""); setUploadInfo(null); try { const buf = await f.arrayBuffer(); const wb = XLSX.read(buf, { type: "array" }); const det = autoRead(wb); const skipped = det._skipped || 0; const list = det.map(r => ({ ...r, stock: DEFAULT_STOCK })); if (!list.length) { setError("No structural sections detected. If this file is only plates, use the Plates module."); return; } setCutList(list); setUploadInfo({ name: f.name, rows: list.length, pieces: list.reduce((s, r) => s + r.qty, 0), skipped }); } catch (e) { setError("Failed to read: " + e.message); } };
  const canonProfileKey = raw => { const sec = findSection(raw); if (sec) return sec.name.toUpperCase().replace(/\s+/g, ""); return String(raw || "").toUpperCase().replace(/×/g, "X").replace(/\*/g, "X").replace(/\s+/g, "").replace(/(\d)\.0+(?=\D|$)/g, "$1"); };
  const runOpt = list => { setError(""); const valid = list.filter(c => c.length > 0 && c.qty > 0); if (!valid.length) { setError("No valid cuts."); return; } const groups = {}; valid.forEach(c => { const k = `${canonProfileKey(c.profile)}||${(c.grade || "").toUpperCase().trim()}`; if (!groups[k]) groups[k] = { items: [], displayProfile: c.profile, grade: c.grade }; groups[k].items.push(c); }); let grandKg = 0; const gres = Object.values(groups).map(({ items, displayProfile, grade }) => { const sec = findSection(displayProfile); const profile = sec ? sec.name : displayProfile; const kgm = sec ? sec.kgm : null; const o = nestBars(items, kerf); const netKg = kgm ? (o.summary.totalNet / 1000) * kgm : null; const stockKg = kgm ? (o.summary.totalStock / 1000) * kgm : null; if (stockKg) grandKg += stockKg; return { profile, grade, kgm, netKg, stockKg, computed: sec?.computed, ...o }; }); setResults({ groups: gres, grandKg }); };

  if (results) return <SectionResults results={results} kerf={kerf} onBack={() => setResults(null)} />;

  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}><SectionTitle>Steel Sections — Workspace</SectionTitle><BackChip onClick={onBack} label="↩ Modules" /></div>
      {inputMode === null && (
        <div>
          <div style={{ textAlign: "center", color: "#94a3b8", fontSize: 15, marginBottom: 18 }}>How would you like to enter your cutting list?</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(260px,1fr))", gap: 20 }}>
            <Chooser icon="✏️" title="Manual Entry" desc="Type a profile (IPE, UB, HEA, SHS…) and pick from the live list; enter lengths." onClick={() => setInputMode("manual")} />
            <Chooser icon="📊" title="Upload Tekla / Excel" desc="Drop a material list. Sections detected; plates ignored." onClick={() => setInputMode("excel")} />
          </div>
        </div>
      )}
      {inputMode === "manual" && (
        <Card title="✏️ Manual Cutting List" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} label="↺ Change method" />}>
          <div style={{ overflowX: "visible" }}><table style={{ width: "100%", borderCollapse: "separate", borderSpacing: "0 10px" }}><thead><tr>{["Profile", "Grade", "Lengths (mm)", "Stock length (mm)", ""].map((h, i) => <th key={h} style={{ textAlign: "left", padding: "0 8px", color: "#64748b", fontSize: 11, letterSpacing: 1, width: i === 0 ? "24%" : i === 1 ? "14%" : i === 3 ? "18%" : i === 4 ? 40 : "auto" }}>{h}</th>)}</tr></thead>
            <tbody>{rows.map(r => { const count = parseLengths(r.lengths).length; const matches = activeAuto === r.id ? searchSections(r.profile) : []; const sec = findSection(r.profile); return (
              <tr key={r.id}>
                <td style={{ position: "relative", padding: "0 8px" }}><input style={INL} placeholder="IPE 300" value={r.profile} onChange={e => { upRow(r.id, "profile", e.target.value); setActiveAuto(r.id); }} onFocus={() => setActiveAuto(r.id)} onBlur={() => setTimeout(() => setActiveAuto(a => a === r.id ? null : a), 150)} />{sec && <div style={{ fontSize: 11, color: sec.computed ? "#60a5fa" : "#22c55e", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{sec.kgm} kg/m · {sec.type}{sec.computed ? " (computed)" : ""}</div>}{matches.length > 0 && <div style={{ position: "absolute", top: "100%", left: 8, right: 8, zIndex: 50, background: "#0f1318", border: "1px solid #2d3748", borderRadius: 6, maxHeight: 260, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,.5)" }}>{matches.map(m => <div key={m.name} onMouseDown={() => { upRow(r.id, "profile", m.name); setActiveAuto(null); }} style={{ padding: "11px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", borderBottom: "1px solid #1a2230", fontFamily: "'Space Mono', monospace", fontSize: 13 }}><span style={{ color: "#cbd5e1" }}>{m.name}</span><span style={{ color: "#f59e0b" }}>{m.kgm} kg/m</span></div>)}</div>}</td>
                <td style={{ padding: "0 8px" }}><input list="grades" style={INL} placeholder="S355" value={r.grade} onChange={e => upRow(r.id, "grade", e.target.value)} /></td>
                <td style={{ padding: "0 8px" }}><textarea style={{ ...INL, minHeight: 52, resize: "vertical", lineHeight: 1.6 }} rows={2} placeholder="3000 4500 6000x3 …" value={r.lengths} onChange={e => upRow(r.id, "lengths", e.target.value)} /><div style={{ fontSize: 11, color: count > 200 ? "#ef4444" : "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>{count} pieces</div></td>
                <td style={{ padding: "0 8px" }}><input style={INL} type="number" placeholder="12000" value={r.stock} onChange={e => upRow(r.id, "stock", e.target.value)} /><div style={{ fontSize: 11, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 5 }}>blank = 12m</div></td>
                <td style={{ padding: "0 8px", verticalAlign: "top" }}>{rows.length > 1 && <button style={{ background: "#1a2230", border: "1px solid #2d3748", borderRadius: 4, color: "#94a3b8", cursor: "pointer", fontSize: 18, width: 40, height: 48 }} onClick={() => delRow(r.id)}>×</button>}</td>
              </tr>); })}</tbody>
          </table></div>
          <datalist id="grades">{STEEL_GRADES.map(g => <option key={g} value={g} />)}</datalist>
          <button style={{ marginTop: 8, padding: "12px 22px", background: "rgba(245,158,11,.15)", border: "1px dashed #d97706", color: "#fbbf24", borderRadius: 4, cursor: "pointer", fontSize: 13 }} onClick={addRow}>+ ADD PROFILE</button>
          <p style={{ fontSize: 11, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 14, lineHeight: 1.6 }}>Type IPE, HEA, HEB, UB, UC, JIS HW/HM/HN, RHS, SHS, CHS, cold-formed C/Z. Built-up girder: type <span style={{ color: "#60a5fa" }}>900x400x20x15</span>. Use <span style={{ color: "#f59e0b" }}>6000x3</span> for 3 pieces.</p>
          <div style={{ marginTop: 16 }}><Label>Cutting Gap / Kerf (mm)</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div>
        </Card>
      )}
      {inputMode === "excel" && (
        <Card title="📊 Upload Tekla / Excel" style={{ marginTop: 20 }} action={<BackChip onClick={() => setInputMode(null)} label="↺ Change method" />}>
          <div onClick={() => fileRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }} style={{ border: "2px dashed #2d3748", borderRadius: 10, padding: "44px 24px", textAlign: "center", cursor: "pointer", background: "rgba(15,19,24,.5)", transition: "border-color .2s" }} onMouseEnter={e => e.currentTarget.style.borderColor = "#f59e0b"} onMouseLeave={e => e.currentTarget.style.borderColor = "#2d3748"}><div style={{ fontSize: 44, marginBottom: 14 }}>📊</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 18, fontWeight: 700, color: "#f8fafc", marginBottom: 8 }}>Upload your material list</div><div style={{ fontSize: 14, color: "#cbd5e1", lineHeight: 1.6, maxWidth: 460, margin: "0 auto 6px" }}>Drop a <b style={{ color: "#fbbf24" }}>Tekla material list</b> — or <b>any</b> material list that has steel <b>lengths and quantities</b>.</div><div style={{ fontSize: 12, color: "#94a3b8", marginTop: 6 }}>Accepts .xlsx · .xls · .csv · .txt — or click to browse</div><div style={{ fontSize: 11, color: "#475569", marginTop: 12, lineHeight: 1.7 }}>Sections are detected automatically · plate rows ignored (use the Plates module for those) · stock length defaults to 12 m and is editable.</div><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (f) handleFile(f); }} /></div>
          {uploadInfo && <div style={{ marginTop: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#22c55e", marginBottom: 12 }}>✓ {uploadInfo.name} — {uploadInfo.rows} sections, {uploadInfo.pieces} pieces{uploadInfo.skipped ? <span style={{ color: "#94a3b8" }}> · {uploadInfo.skipped} plate row{uploadInfo.skipped > 1 ? "s" : ""} ignored</span> : null}</div><div style={{ marginBottom: 14 }}><Label>Cutting Gap / Kerf (mm)</Label><input type="number" value={kerf} onChange={e => setKerf(+e.target.value)} style={{ ...IN, width: 90 }} /></div><table style={{ width: "100%", borderCollapse: "collapse" }}><thead><tr>{["Profile", "Grade", "Length", "Qty", "kg/m"].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 10px", borderBottom: "1px solid #2d3748", color: "#475569", fontSize: 10, letterSpacing: 1, fontFamily: "'Space Mono', monospace" }}>{h}</th>)}</tr></thead><tbody>{cutList.slice(0, 40).map((r, i) => { const sec = findSection(r.profile); return <tr key={i}><td style={TD}>{r.profile}</td><td style={TD}>{r.grade || "—"}</td><td style={TD}>{fmtMm(r.length)}</td><td style={TD}>{r.qty}</td><td style={{ ...TD, color: sec ? "#22c55e" : "#475569" }}>{sec ? sec.kgm : "—"}</td></tr>; })}</tbody></table>{cutList.length > 40 && <div style={{ fontSize: 11, color: "#475569", fontFamily: "'Space Mono', monospace", marginTop: 8 }}>…and {cutList.length - 40} more</div>}</div>}
        </Card>
      )}
      {error && <div style={{ background: "#450a0a", border: "1px solid #7f1d1d", borderRadius: 6, padding: "12px 16px", color: "#fca5a5", fontSize: 12, fontFamily: "'Space Mono', monospace", margin: "12px 0" }}>⚠ {error}</div>}
      {inputMode === "manual" && <div style={{ textAlign: "center", marginTop: 24 }}><button onClick={() => { const l = buildManual(); if (!l.length) { setError("Enter at least one length."); return; } runOpt(l); }} style={OPT_BTN(false)}>⚡ OPTIMIZE SECTIONS</button></div>}
      {inputMode === "excel" && uploadInfo && <div style={{ textAlign: "center", marginTop: 24 }}><button onClick={() => runOpt(cutList)} style={OPT_BTN(false)}>⚡ OPTIMIZE SECTIONS</button></div>}
    </>
  );
}

function SectionResults({ results, kerf, onBack }) {
  const totalStock = results.groups.reduce((s, g) => s + g.summary.stockCount, 0);
  const stockMm = results.groups.reduce((s, g) => s + g.summary.totalStock, 0);
  const wasteMm = results.groups.reduce((s, g) => s + g.summary.totalWaste, 0);
  const wastePct = stockMm ? ((wasteMm / stockMm) * 100).toFixed(1) : "0";
  const wasteKg = results.groups.reduce((s, g) => s + (g.kgm ? (g.summary.totalWaste / 1000) * g.kgm : 0), 0);
  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}><SectionTitle>Section Results</SectionTitle><BackChip onClick={onBack} label="← Back to input" /></div>
      <div style={{ background: "linear-gradient(135deg,#f59e0b,#d97706)", borderRadius: 12, padding: "22px 26px", marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, boxShadow: "0 8px 30px rgba(245,158,11,.3)" }}><div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#1a1206", opacity: .7, letterSpacing: ".1em" }}>TOTAL STEEL WEIGHT TO PROCURE</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 38, fontWeight: 900, color: "#1a1206" }}>{fmtKg(results.grandKg)}</div></div><div style={{ textAlign: "right" }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#1a1206", opacity: .7 }}>{results.groups.length} PROFILE GROUP{results.groups.length > 1 ? "S" : ""}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#1a1206", fontWeight: 700 }}>{totalStock} STOCK BARS</div></div></div>
      {results.groups.map((g, gi) => (
        <div key={gi} style={{ marginBottom: 32 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, paddingBottom: 8, borderBottom: "1px solid #1e293b", flexWrap: "wrap" }}><span style={{ fontFamily: "'Playfair Display', serif", fontSize: 22, fontWeight: 700, color: "#f8fafc" }}>{g.profile}</span>{g.grade && <span style={TAG}>{g.grade}</span>}{g.kgm && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#94a3b8" }}>{g.kgm} kg/m{g.computed ? " (computed)" : ""}</span>}{g.stockKg != null && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 14, color: "#f59e0b", fontWeight: 700 }}>{fmtKg(g.stockKg)}</span>}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(110px,1fr))", gap: 8, marginBottom: 16 }}>{[{ v: g.summary.stockCount, l: "Bars to Buy" }, { v: fmtMm(g.summary.totalStock), l: "Total Length" }, { v: fmtMm(g.summary.totalWaste), l: "Offcut" }, { v: `${g.summary.wastePct}%`, l: "Waste" }, { v: `${g.summary.utilPct}%`, l: "Utilization" }, { v: g.stockKg != null ? fmtKg(g.stockKg) : "—", l: "Weight" }].map(s => <div key={s.l} style={ST}><div style={STV}>{s.v}</div><div style={STL}>{s.l}</div></div>)}</div>
          <div style={{ background: "#1c1600", border: "1px solid #f59e0b", borderRadius: 10, padding: 18, marginBottom: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#f59e0b", letterSpacing: ".15em", marginBottom: 14 }}>⬡ WHAT TO ORDER</div>{Object.entries(g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {})).map(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : null; return <div key={len} style={{ display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap", marginBottom: 10, paddingBottom: 10, borderBottom: "1px dashed #3a2e0a" }}><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#94a3b8" }}>YOU NEED</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 30, color: "#f59e0b", lineHeight: 1 }}>{qty}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#94a3b8" }}>×</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: 22, color: "#f8fafc" }}>{g.profile}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#cbd5e1" }}>@ {fmtMm(parseInt(len))}</span>{g.grade && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#94a3b8" }}>grade {g.grade}</span>}{kg && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#22c55e", fontWeight: 700 }}>{fmtKg(kg)}</span>}</div>; })}</div>
          {g.splices && g.splices.length > 0 && <div style={{ background: "#0c1f14", border: "1px solid #15803d", borderRadius: 10, padding: 18, marginBottom: 16 }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#4ade80", letterSpacing: ".1em", marginBottom: 10 }}>⚙ LONG MEMBERS — SPLICED FROM {fmtMm(g.splices[0].stock)} BARS</div>{g.splices.map((sp, oi) => <div key={oi} style={{ fontSize: 13, color: "#bbf7d0", marginBottom: 4 }}><span style={{ fontFamily: "'Space Mono', monospace", fontWeight: 700 }}>{sp.qty} × </span>{fmtMm(sp.length)}<span style={{ color: "#4ade80", fontSize: 11 }}> — each from {sp.bars} × {fmtMm(sp.stock)} bars</span></div>)}</div>}
          <Card title="Visual Cut Plan">{g.bins.length ? g.bins.map((b, i) => <CutBar key={i} bin={b} index={i} />) : <div style={{ color: "#64748b", fontFamily: "'Space Mono', monospace", fontSize: 12 }}>No plan.</div>}</Card>
        </div>
      ))}
      <div style={{ background: "#1c1600", border: "2px solid #f59e0b", borderRadius: 12, padding: "20px 24px", marginBottom: 16 }}>
        <div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#f59e0b", letterSpacing: ".15em", marginBottom: 14 }}>✦ SUMMARY — WHAT YOU NEED TO BUY</div>
        {results.groups.map((g, gi) => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); const parts = Object.entries(byLen).map(([len, qty]) => `${qty} × ${fmtMm(parseInt(len))}`).join(", "); return <div key={gi} style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap", padding: "8px 0", borderBottom: gi < results.groups.length - 1 ? "1px dashed #3a2e0a" : "none" }}><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 900, fontSize: 26, color: "#f59e0b", lineHeight: 1 }}>{g.summary.stockCount}</span><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#cbd5e1" }}>bar{g.summary.stockCount > 1 ? "s" : ""} of</span><span style={{ fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: 20, color: "#f8fafc" }}>{g.profile}</span>{g.grade && <span style={{ fontFamily: "'Space Mono', monospace", fontSize: 12, color: "#94a3b8" }}>({g.grade})</span>}<span style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#64748b" }}>— {parts}</span>{g.stockKg != null && <span style={{ marginLeft: "auto", fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#22c55e", fontWeight: 700 }}>{fmtKg(g.stockKg)}</span>}</div>; })}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 14, paddingTop: 12, borderTop: "1px solid #3a2e0a" }}><span style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#cbd5e1", fontWeight: 700 }}>TOTAL: {totalStock} bars across {results.groups.length} profile{results.groups.length > 1 ? "s" : ""}</span><span style={{ fontFamily: "'Playfair Display', serif", fontSize: 22, fontWeight: 900, color: "#f59e0b" }}>{fmtKg(results.grandKg)}</span></div>
      </div>
      <div style={{ background: parseFloat(wastePct) <= 8 ? "linear-gradient(135deg,#16a34a,#15803d)" : parseFloat(wastePct) <= 20 ? "linear-gradient(135deg,#f59e0b,#d97706)" : "linear-gradient(135deg,#dc2626,#991b1b)", borderRadius: 12, padding: "20px 24px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}><div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#fff", opacity: .85, letterSpacing: ".1em" }}>WASTE AFTER CUTTING</div><div style={{ fontFamily: "'Playfair Display', serif", fontSize: 40, fontWeight: 900, color: "#fff" }}>{wastePct}%</div></div><div style={{ textAlign: "right" }}><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 11, color: "#fff", opacity: .85 }}>OFFCUT LENGTH</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 18, color: "#fff", fontWeight: 700 }}>{fmtMm(Math.round(wasteMm))}</div><div style={{ fontFamily: "'Space Mono', monospace", fontSize: 13, color: "#fff", opacity: .9, marginTop: 4 }}>{fmtKg(wasteKg)} scrap/reuse</div></div></div>
      {(() => {
        const REUSE_MIN = 1000; // a leftover bar >= 1 m is worth keeping
        const rows = [];
        results.groups.forEach(g => g.bins.forEach((b, i) => { if (b.remaining >= REUSE_MIN) rows.push({ profile: g.profile, grade: g.grade, bar: i + 1, stock: b.stockLength, len: Math.round(b.remaining), kg: g.kgm ? (b.remaining / 1000) * g.kgm : null }); }));
        const totalKg = rows.reduce((s, r) => s + (r.kg || 0), 0);
        if (!rows.length) return null;
        return (
          <Card title="♻ Reusable Bar Offcuts — leftover steel for the next project" style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 12, color: "#6ee7b7", marginBottom: 12, background: "rgba(16,185,129,.08)", border: "1px solid rgba(16,185,129,.25)", borderRadius: 6, padding: "8px 12px" }}>{rows.length} leftover bar end{rows.length > 1 ? "s" : ""} ≥ 1 m ≈ <b>{fmtKg(totalKg)}</b> — keep these to cut from on your next job instead of buying new.</div>
            <div style={{ overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "'Space Mono', monospace" }}><thead><tr style={{ background: "rgba(16,185,129,.12)", borderBottom: "1px solid rgba(16,185,129,.3)" }}>{["PROFILE", "GRADE", "FROM BAR", "STOCK LENGTH", "REUSABLE LEFTOVER", "WEIGHT", "STATUS"].map((h, i) => <th key={h} style={{ padding: "10px 12px", textAlign: i > 4 ? "right" : "left", color: "#6ee7b7", fontSize: 11, letterSpacing: 1, fontWeight: 700 }}>{h}</th>)}</tr></thead>
              <tbody>{rows.map((r, i) => <tr key={i} style={{ borderBottom: "1px solid #1a2230" }}><td style={{ padding: "9px 12px", color: "#cbd5e1" }}>{r.profile}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{r.grade || "—"}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>Bar {r.bar}</td><td style={{ padding: "9px 12px", color: "#94a3b8" }}>{fmtMm(r.stock)}</td><td style={{ padding: "9px 12px", color: "#6ee7b7", fontWeight: 700 }}>{fmtMm(r.len)}</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#94a3b8" }}>{r.kg != null ? fmtKg(r.kg) : "—"}</td><td style={{ padding: "9px 12px", textAlign: "right", color: "#10B981", fontWeight: 700 }}>✓ keep</td></tr>)}</tbody>
            </table></div>
          </Card>
        );
      })()}
      <Card title="💾 Export"><div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}><button onClick={() => exportSectionPDF(results, kerf, { totalStock, wastePct, wasteMm, wasteKg })} style={EXP}>📄 Export PDF Report</button><button onClick={() => exportSectionExcel(results)} style={EXP}>📊 Export Excel Summary</button></div></Card>
    </>
  );
}

/* ─── SHARED UI ──────────────────────────────────────────────────────────── */
function SectionTitle({ children }) { return <h2 style={{ fontSize: 24, fontWeight: 900, color: "#f8fafc", margin: 0, fontFamily: "'Playfair Display', serif", borderLeft: "4px solid #f59e0b", paddingLeft: 14 }}>{children}</h2>; }
function Chooser({ icon, title, desc, onClick }) { return <button onClick={onClick} style={{ textAlign: "left", padding: "28px 26px", borderRadius: 12, background: "#131920", border: "1px solid #1e293b", cursor: "pointer", transition: "all .2s", fontFamily: "inherit", color: "#cbd5e1", minHeight: 200, display: "flex", flexDirection: "column" }} onMouseEnter={e => { e.currentTarget.style.borderColor = "#f59e0b"; e.currentTarget.style.transform = "translateY(-3px)"; }} onMouseLeave={e => { e.currentTarget.style.borderColor = "#1e293b"; e.currentTarget.style.transform = "translateY(0)"; }}><div style={{ fontSize: 40 }}>{icon}</div><div style={{ fontSize: 22, fontWeight: 700, color: "#f8fafc", margin: "16px 0 10px", fontFamily: "'Playfair Display', serif" }}>{title}</div><div style={{ fontSize: 14, color: "#94a3b8", lineHeight: 1.5, flex: 1 }}>{desc}</div><span style={{ color: "#f59e0b", fontFamily: "'Space Mono', monospace", fontSize: 13, fontWeight: 700, marginTop: 12 }}>Choose →</span></button>; }
function Card({ title, children, style, action }) { return <div style={{ background: "#131920", border: "1px solid #1e293b", borderRadius: 10, overflow: "hidden", ...style }}><div style={{ padding: "12px 18px", borderBottom: "1px solid #1e293b", background: "rgba(245,158,11,.04)", display: "flex", alignItems: "center", justifyContent: "space-between" }}><span style={{ fontSize: 13, fontWeight: 700, color: "#fbbf24", letterSpacing: .5, fontFamily: "'Space Mono', monospace" }}>{title}</span>{action}</div><div style={{ padding: 18 }}>{children}</div></div>; }
function Label({ children }) { return <div style={{ fontSize: 10, color: "#64748b", letterSpacing: 1, textTransform: "uppercase", marginBottom: 5, fontFamily: "'Space Mono', monospace" }}>{children}</div>; }
function BackChip({ onClick, label }) { return <button onClick={onClick} style={{ padding: "5px 14px", borderRadius: 20, border: "1px solid #2d3748", background: "rgba(15,19,24,.6)", color: "#94a3b8", cursor: "pointer", fontSize: 11, fontFamily: "'Space Mono', monospace" }}>{label}</button>; }
function StatCard({ l, v, i, a }) { return <div style={{ padding: "16px 14px", borderRadius: 8, background: "#0f1318", border: "1px solid #1e293b", textAlign: "center" }}><div style={{ fontSize: 20, marginBottom: 4 }}>{i}</div><div style={{ fontSize: 20, fontWeight: 800, color: a || "#cbd5e1", fontFamily: "'Space Mono', monospace" }}>{v}</div><div style={{ fontSize: 10, color: "#64748b", marginTop: 4, letterSpacing: 1 }}>{l}</div></div>; }
function BigTon({ label, value, note, color, big }) { return <div style={{ padding: "20px 22px", borderRadius: 10, background: big ? `${color}1a` : "#0f1318", border: `1px solid ${color}55`, textAlign: "center" }}><div style={{ fontSize: 10, color: "#64748b", letterSpacing: 1.5, textTransform: "uppercase", marginBottom: 8, fontFamily: "'Space Mono', monospace" }}>{label}</div><div style={{ fontFamily: "'Space Mono', monospace", fontWeight: 800, color, fontSize: big ? 44 : 32, lineHeight: 1 }}>{value}<span style={{ fontSize: big ? 18 : 14, marginLeft: 4 }}>Ton</span></div><div style={{ fontSize: 11, color: "#64748b", marginTop: 8, fontFamily: "'Space Mono', monospace" }}>{note}</div></div>; }

const IN = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "7px 10px", borderRadius: 4, fontSize: 13, width: 110, fontFamily: "'Space Mono', monospace", outline: "none" };
const ZB = { width: 28, height: 28, borderRadius: 4, border: "1px solid #2d3748", background: "rgba(15,19,24,.8)", color: "#94a3b8", cursor: "pointer", fontSize: 16, lineHeight: 1 };
const INL = { background: "#0f1318", border: "1.5px solid #334155", color: "#cbd5e1", padding: "14px 14px", borderRadius: 4, fontSize: 15, width: "100%", boxSizing: "border-box", fontFamily: "'Space Mono', monospace", outline: "none" };
const SEL = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "7px 10px", borderRadius: 4, fontSize: 13, fontFamily: "'Space Mono', monospace", cursor: "pointer", outline: "none" };
const CI = { background: "#0f1318", border: "1px solid #2d3748", color: "#cbd5e1", padding: "5px 7px", borderRadius: 3, fontSize: 12, fontFamily: "'Space Mono', monospace", outline: "none" };
const BTN = { padding: "8px 18px", borderRadius: 4, border: "1px solid #d97706", background: "rgba(245,158,11,.15)", color: "#fbbf24", cursor: "pointer", fontSize: 12, marginTop: 6 };
const EXP = { padding: "11px 24px", borderRadius: 6, border: "none", background: "linear-gradient(135deg,#f59e0b,#d97706)", color: "#1a1206", cursor: "pointer", fontSize: 13, fontWeight: 700, fontFamily: "'Space Mono', monospace" };
const TAG = { display: "inline-block", padding: "4px 10px", borderRadius: 20, fontSize: 11, fontFamily: "'Space Mono', monospace", background: "rgba(245,158,11,.15)", border: "1px solid #f59e0b", color: "#f59e0b" };
const TD = { padding: "9px 10px", borderBottom: "1px solid #1a2230", color: "#cbd5e1", fontFamily: "'Space Mono', monospace", fontSize: 12 };
const ST = { background: "#0f1318", border: "1px solid #1e293b", borderRadius: 8, padding: "16px 14px", textAlign: "center" };
const STV = { fontSize: 20, fontWeight: 900, fontFamily: "'Space Mono', monospace", color: "#f59e0b", lineHeight: 1.1 };
const STL = { fontSize: 9, color: "#64748b", letterSpacing: ".08em", marginTop: 4, textTransform: "uppercase" };
const OPT_BTN = isOpt => ({ padding: "16px 60px", fontSize: 16, fontWeight: 800, background: isOpt ? "rgba(120,80,10,.4)" : "linear-gradient(135deg,#f59e0b,#d97706)", border: "none", color: "#1a1206", borderRadius: 8, cursor: isOpt ? "default" : "pointer", boxShadow: "0 8px 30px rgba(245,158,11,.4)", letterSpacing: 2, fontFamily: "'Space Mono', monospace" });

/* ─── EXPORTS ────────────────────────────────────────────────────────────── */
function downloadHTML(html, name) { const win = window.open("", "_blank"); if (win && win.document) { win.document.open(); win.document.write(html); win.document.close(); } else { const blob = new Blob([html], { type: "text/html" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 4000); } }
function saveWorkbook(wb, name) { try { XLSX.writeFile(wb, name); } catch { try { const out = XLSX.write(wb, { bookType: "xlsx", type: "array" }); const blob = new Blob([out], { type: "application/octet-stream" }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 4000); } catch (e) { alert("Export failed: " + (e.message || e)); } } }

function exportPlateExcel(results, material, reuseMin) {
  const wb = XLSX.utils.book_new();
  const summary = [["Steel Optimizer — Plate Procurement"], [], ["Material", material], [],
    ["Thickness (mm)", "Sheet Size", "Sheets Req.", "Purchase Wt (t)", "Parts", "Utilization %", "Net Parts Wt (t)", "Scrap Wt (t)"],
    ...results.groups.map(g => [g.thickness, `${g.sw}x${g.sh}`, g.sheetCount, (g.sheetWeight / 1000).toFixed(2), g.partCount, g.utilPct, (g.partWeight / 1000).toFixed(2), (g.wasteWeight / 1000).toFixed(2)]),
    ["TOTAL", `${results.groups.length} thicknesses`, results.totals.sheets, (results.totals.sheetWeight / 1000).toFixed(2), results.totals.parts, results.totals.utilPct, (results.totals.partWeight / 1000).toFixed(2), (results.totals.wasteWeight / 1000).toFixed(2)],
    [], ["TOTAL PURCHASE WEIGHT - gross full sheets (Ton)", (results.totals.sheetWeight / 1000).toFixed(2)],
    ["NET PARTS WEIGHT - finished plates only (Ton)", (results.totals.partWeight / 1000).toFixed(2)],
    ["SCRAP / OFFCUT WEIGHT (Ton)", (results.totals.wasteWeight / 1000).toFixed(2)],
    ["REUSABLE OFFCUTS", `${results.totals.offcutCount} pieces`, `${(results.totals.offcutWeight / 1000).toFixed(2)} t`]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "Summary");
  const off = [[`Reusable Offcut Register (each side >= ${reuseMin} mm)`], [], ["Thickness (mm)", "From Sheet", "Offcut W (mm)", "Offcut L (mm)", "Weight (kg)", "Status"],
    ...results.groups.flatMap(g => g.offcuts.map(o => [g.thickness, `Sheet ${o.sheet}`, o.w, o.h, o.weight.toFixed(2), "keep for future jobs"]))];
  if (results.totals.offcutCount === 0) off.push(["No reusable offcuts — all leftovers below the minimum size."]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(off), "Reusable Offcuts");
  saveWorkbook(wb, "SteelOptimizer_Plates.xlsx");
}

function exportOffcutExcel(results, material, reuseMin) {
  const wb = XLSX.utils.book_new();
  const rows = [["Reusable Offcut Register"], ["Material", material], [`Minimum size kept (each side)`, `${reuseMin} mm`], [],
    ["#", "Thickness (mm)", "From Sheet", "Shape", "Overall W (mm)", "Overall L (mm)", "Notch W (mm)", "Notch L (mm)", "Usable Rect W (mm)", "Usable Rect L (mm)", "Weight (kg)", "Status"]];
  let n = 0; results.groups.forEach(g => g.offcuts.forEach(o => { n++; rows.push([n, g.thickness, `Sheet ${o.sheet}`, o.shape === "L" ? "L-shape" : "Rectangle", o.A, o.B, o.shape === "L" ? o.notchW : 0, o.shape === "L" ? o.notchH : 0, o.w, o.h, o.weight.toFixed(2), "keep for future jobs"]); }));
  if (n === 0) rows.push(["—", "No reusable offcuts above the minimum size."]);
  rows.push([], ["TOTAL REUSABLE OFFCUTS", `${results.totals.offcutCount} pieces`, "", "", "", (results.totals.offcutWeight).toFixed(1) + " kg", `${fmtTon(results.totals.offcutWeight)} t`]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Reusable Offcuts");
  saveWorkbook(wb, "SteelOptimizer_Offcuts.xlsx");
}

function offcutSVG(o) {
  // Draw the real leftover shape (L or rectangle) with the usable rectangle shaded.
  const W = 150, H = 130, pad = 18;
  const A = o.A, B = o.B, sc = Math.min((W - pad * 2) / A, (H - pad * 2) / B);
  const aw = A * sc, bh = B * sc, ox = (W - aw) / 2, oy = (H - bh) / 2;
  let body;
  if (o.shape === "L") {
    const nW = o.notchW * sc, nH = o.notchH * sc;
    const pts = [[ox + nW, oy], [ox + aw, oy], [ox + aw, oy + bh], [ox, oy + bh], [ox, oy + nH], [ox + nW, oy + nH]];
    body = `<polygon points="${pts.map(p => p.map(v => v.toFixed(1)).join(",")).join(" ")}" fill="#d8f3e6" stroke="#0a7a52" stroke-width="1.5"/>`;
    // usable rect: bottom strip or right strip
    const useBottom = o.w === A;
    const ux = ox, uy = useBottom ? oy + nH : oy, uw = useBottom ? aw : (aw - nW), uh = useBottom ? (bh - nH) : bh;
    body += `<rect x="${ux.toFixed(1)}" y="${uy.toFixed(1)}" width="${uw.toFixed(1)}" height="${uh.toFixed(1)}" fill="#10b981" fill-opacity="0.35" stroke="#10b981" stroke-width="1" stroke-dasharray="4 2"/>`;
  } else {
    body = `<rect x="${ox.toFixed(1)}" y="${oy.toFixed(1)}" width="${aw.toFixed(1)}" height="${bh.toFixed(1)}" fill="#10b981" fill-opacity="0.3" stroke="#0a7a52" stroke-width="1.5"/>`;
  }
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body}<text x="${W / 2}" y="${H - 4}" font-size="9" text-anchor="middle" fill="#0a7a52" font-family="monospace">${A}×${B}${o.shape === "L" ? " L" : ""}</text></svg>`;
}

function exportOffcutPDF(results, material, reuseMin) {
  const rowsHtml = results.groups.flatMap(g => g.offcuts.map((o, k) => `<tr><td>${g.thickness} mm</td><td>Sheet ${o.sheet}</td><td>${o.shape === "L" ? "L-shape" : "Rectangle"}</td><td>${o.shape === "L" ? `${o.A} x ${o.B} (notch ${o.notchW} x ${o.notchH})` : `${o.A} x ${o.B}`}</td><td><b>${o.w} x ${o.h} mm</b></td><td>${o.weight >= 1 ? o.weight.toFixed(1) + " kg" : (o.weight * 1000).toFixed(0) + " g"}</td><td>keep for future jobs</td></tr>`)).join("");
  const cards = results.groups.flatMap(g => g.offcuts.map(o => `<div class="card">${offcutSVG(o)}<div class="cap"><b>${g.thickness} mm · Sheet ${o.sheet}</b><br/>${o.shape === "L" ? "L-shape" : "Rectangle"} — usable <b>${o.w}×${o.h}</b><br/>${o.weight >= 1 ? o.weight.toFixed(1) + " kg" : (o.weight * 1000).toFixed(0) + " g"}</div></div>`)).join("");
  const html = `<html><head><title>Steel Optimizer — Reusable Offcuts</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#04140d}
  h1{color:#0a7a52;border-bottom:2px solid #bfe6d2;padding-bottom:10px}
  h2{color:#0a7a52;font-size:15px;margin-top:26px}
  table{width:100%;border-collapse:collapse;margin-top:14px;font-family:monospace;font-size:13px}
  th{background:#e8f8ef;color:#0a7a52;padding:10px;text-align:left;font-size:11px}
  td{padding:10px;border-bottom:1px solid #eee}
  tr.total td{background:#e8f8ef;font-weight:bold;border-top:2px solid #10b981}
  .banner{margin-top:14px;padding:14px 16px;background:#effaf3;border:2px solid #0a7a52;border-radius:8px;font-size:14px}
  .grid{display:flex;flex-wrap:wrap;gap:14px;margin-top:12px}
  .card{border:1px solid #cdeadd;border-radius:8px;padding:10px;text-align:center;background:#fbfffd;width:170px}
  .cap{font-family:monospace;font-size:10px;color:#234;margin-top:4px;line-height:1.5}
  button{margin-top:24px;padding:12px 28px;background:#0a7a52;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer}
  @media print{button{display:none}}
  </style></head><body>
  <h1>&#9851; Reusable Offcut Register</h1>
  <p>Material: <b>${material}</b> &middot; ${new Date().toLocaleDateString("en-GB")} &middot; minimum kept side &ge; ${reuseMin} mm</p>
  <div class="banner">${results.totals.offcutCount} reusable offcut${results.totals.offcutCount !== 1 ? "s" : ""} &asymp; <b>${fmtTon(results.totals.offcutWeight)} t</b> to keep for future jobs. Green dashed = the largest rectangle you can cut from each leftover.</div>
  <h2>Leftover shapes</h2>
  <div class="grid">${cards || "<p>No reusable offcuts above the minimum size.</p>"}</div>
  <h2>Register</h2>
  <table><tr><th>Thickness</th><th>From Sheet</th><th>Shape</th><th>Overall</th><th>Usable Rect</th><th>Weight</th><th>Status</th></tr>
  ${rowsHtml || `<tr><td colspan="7">No reusable offcuts above the minimum size.</td></tr>`}
  <tr class="total"><td colspan="5">TOTAL — ${results.totals.offcutCount} pieces</td><td>${fmtTon(results.totals.offcutWeight)} t</td><td></td></tr></table>
  <button onclick="window.print()">Print / Save as PDF</button></body></html>`;
  downloadHTML(html, "SteelOptimizer_Offcuts.html");
}

function sheetSVG(group, sObj, partColors) {
  const sw = group.sw, sh = group.sh, W = 150, H = Math.round(150 * sh / sw);
  const sc = W / sw;
  let body = `<rect x="0" y="0" width="${W}" height="${H}" fill="#0f1318" stroke="#b45309" stroke-width="1"/>`;
  (sObj.placements || []).forEach((p, i) => {
    const x = p.x * sc, y = p.y * sc, w = p.w * sc, h = p.h * sc, col = partColors[p.id] || "#3b82f6";
    body += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="${col}" fill-opacity="0.45" stroke="${col}" stroke-width="0.6"/>`;
    if (p.spliced) body += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="none" stroke="#f59e0b" stroke-width="0.8" stroke-dasharray="3 2"/>`;
  });
  (sObj.offcuts || []).forEach(o => {
    if (o.shape === "L") {
      const ox = o.x * sc, oy = o.y * sc, A = o.A * sc, B = o.B * sc, nW = o.notchW * sc, nH = o.notchH * sc;
      const pts = [[ox + nW, oy], [ox + A, oy], [ox + A, oy + B], [ox, oy + B], [ox, oy + nH], [ox + nW, oy + nH]];
      body += `<polygon points="${pts.map(p => p.map(v => v.toFixed(1)).join(",")).join(" ")}" fill="#10b981" fill-opacity="0.22" stroke="#10b981" stroke-width="0.8"/>`;
    } else {
      body += `<rect x="${(o.x * sc).toFixed(1)}" y="${(o.y * sc).toFixed(1)}" width="${(o.A * sc).toFixed(1)}" height="${(o.B * sc).toFixed(1)}" fill="#10b981" fill-opacity="0.22" stroke="#10b981" stroke-width="0.8"/>`;
    }
  });
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body}</svg>`;
}

function exportPlatePDF(results, material, reuseMin) {
  const offRows = results.groups.flatMap(g => g.offcuts.map(o => `<tr><td>${g.thickness} mm</td><td>Sheet ${o.sheet}</td><td>${o.shape === "L" ? "L-shape" : "Rect"}</td><td>${o.shape === "L" ? `${o.A}×${o.B} (notch ${o.notchW}×${o.notchH})` : `${o.A}×${o.B}`}</td><td><b>${o.w} × ${o.h} mm</b></td><td>${o.weight.toFixed(1)} kg</td><td>keep for future jobs</td></tr>`)).join("");
  const PCOL = ["#3b82f6", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#14b8a6", "#f97316", "#06b6d4", "#a855f7"];
  const layouts = results.groups.map(g => { let ci = 0; const cmap = {}; (g.sheets || []).forEach(s => (s.placements || []).forEach(p => { if (!(p.id in cmap)) cmap[p.id] = PCOL[ci++ % PCOL.length]; })); const tiles = (g.sheets || []).map((s, si) => `<div class="tile">${sheetSVG(g, s, cmap)}<div class="tcap">Sheet ${si + 1} · ${g.thickness}mm</div></div>`).join(""); return `<h2>${g.thickness} mm — ${g.sheetCount} sheet${g.sheetCount > 1 ? "s" : ""} (${g.sw}×${g.sh})</h2><div class="tiles">${tiles}</div>`; }).join("");
  const html = `<html><head><title>Steel Optimizer — Plate Report</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#1a1206}
  h1{color:#b45309;border-bottom:2px solid #f0d9b0;padding-bottom:10px}
  h2{color:#0a7a52;font-size:15px;margin-top:24px;border-bottom:1px solid #e6dcc8;padding-bottom:6px}
  table{width:100%;border-collapse:collapse;margin-top:14px;font-family:monospace;font-size:13px}
  th{background:#fdf4e3;color:#b45309;padding:10px;text-align:left;font-size:11px}
  td{padding:10px;border-bottom:1px solid #eee}
  tr.total td{background:#fdf4e3;font-weight:bold;border-top:2px solid #f59e0b}
  table.off th{background:#e8f8ef;color:#0a7a52}
  .purchase{margin-top:20px;padding:18px;background:#fdf4e3;border:2px solid #f59e0b;border-radius:8px;font-size:18px}
  .purchase b{font-size:30px;color:#b45309}
  .buylist{margin:14px 0;padding:16px 18px;background:#fffaf0;border:1px solid #f0d9b0;border-radius:8px}
  .buylist h3{margin:0 0 10px;color:#b45309;font-size:13px;letter-spacing:1px;text-transform:uppercase}
  .buylist .row{font-size:16px;margin:6px 0}.buylist .n{font-weight:bold;color:#92400e;font-size:19px}
  .reuse{margin-top:12px;padding:14px 16px;background:#effaf3;border:2px solid #0a7a52;border-radius:8px;font-size:14px}
  .tiles{display:flex;flex-wrap:wrap;gap:10px;margin-top:10px}
  .tile{text-align:center}.tcap{font-family:monospace;font-size:9px;color:#555;margin-top:3px}
  .legend{font-size:11px;color:#555;margin-top:6px}.legend b{color:#0a7a52}
  button{margin-top:24px;padding:12px 28px;background:#b45309;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer}
  @media print{button{display:none}}
  </style></head><body>
  <h1>Steel Optimizer — Plate Procurement Summary</h1>
  <p>Material: <b>${material}</b> &nbsp;&middot;&nbsp; ${new Date().toLocaleDateString("en-GB")}</p>
  <div class="buylist"><h3>What you need to buy</h3>
  ${results.groups.map(g => `<div class="row">&#10003; You need <span class="n">${g.sheetCount}</span> sheet${g.sheetCount > 1 ? "s" : ""} of <b>${g.thickness} mm</b> (${g.sw}x${g.sh} mm, ${material})</div>`).join("")}</div>
  <table><tr><th>Thickness</th><th>Sheet Size</th><th>Sheets</th><th>Purchase Wt</th><th>Parts</th><th>Utilization</th><th>Net Parts Wt</th><th>Scrap Wt</th></tr>
  ${results.groups.map(g => `<tr><td>${g.thickness} mm</td><td>${g.sw}x${g.sh}</td><td>${g.sheetCount}</td><td>${fmtTon(g.sheetWeight)} t</td><td>${g.partCount}</td><td>${g.utilPct}%</td><td>${fmtTon(g.partWeight)} t</td><td>${fmtTon(g.wasteWeight)} t</td></tr>`).join("")}
  <tr class="total"><td>TOTAL</td><td>${results.groups.length} thk</td><td>${results.totals.sheets}</td><td>${fmtTon(results.totals.sheetWeight)} t</td><td>${results.totals.parts}</td><td>${results.totals.utilPct}%</td><td>${fmtTon(results.totals.partWeight)} t</td><td>${fmtTon(results.totals.wasteWeight)} t</td></tr></table>
  <p style="font-size:12px;color:#777">Purchase Wt = gross full sheets &middot; Net Parts Wt = finished cut plates &middot; Scrap Wt = offcut (Purchase - Net)</p>
  <div class="purchase">Total Purchase Weight (gross): <b>${fmtTon(results.totals.sheetWeight)} Ton</b> &nbsp;(${results.totals.sheets} sheets) &middot; Net: ${fmtTon(results.totals.partWeight)} t &middot; Scrap: ${fmtTon(results.totals.wasteWeight)} t</div>
  <h2>Nesting layouts</h2>
  <div class="legend">Amber dashed = welded/spliced piece &middot; <b>Green = reusable leftover</b> (L-shape or rectangle).</div>
  ${layouts}
  <h2>&#9851; Reusable Offcut Register (each side >= ${reuseMin} mm)</h2>
  ${results.totals.offcutCount > 0 ? `<div class="reuse">${results.totals.offcutCount} reusable offcut${results.totals.offcutCount > 1 ? "s" : ""} &asymp; ${fmtTon(results.totals.offcutWeight)} t to keep for future jobs.</div><table class="off"><tr><th>Thickness</th><th>From Sheet</th><th>Shape</th><th>Overall</th><th>Usable Rect</th><th>Weight</th><th>Status</th></tr>${offRows}</table>` : `<p style="color:#777">No reusable offcuts — all leftovers below the minimum size.</p>`}
  <button onclick="window.print()">Print / Save as PDF</button></body></html>`;
  downloadHTML(html, "SteelOptimizer_Plates.html");
}

function exportSectionExcel(results) {
  const wb = XLSX.utils.book_new();
  const proc = [["Steel Optimizer — Section Procurement"], [], ["Profile", "Grade", "Market Length (mm)", "Qty to Order", "kg/m", "Total Weight (t)"]];
  results.groups.forEach(g => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); Object.entries(byLen).forEach(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : 0; proc.push([g.profile, g.grade || "", parseInt(len), qty, g.kgm || "", (kg / 1000).toFixed(3)]); }); });
  proc.push([], ["TOTAL STEEL TO PROCURE (Ton)", "", "", "", "", (results.grandKg / 1000).toFixed(2)]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(proc), "Procurement");
  const plan = [["Profile", "Grade", "Bar #", "Stock Length (mm)", "Cuts (mm)", "Pieces", "Offcut (mm)", "Utilization %"]];
  results.groups.forEach(g => g.bins.forEach((b, i) => plan.push([g.profile, g.grade || "", i + 1, b.stockLength, b.cuts.map(c => Math.round(c.length)).join(" + "), b.cuts.length, Math.round(b.remaining), Math.round(((b.stockLength - b.remaining) / b.stockLength) * 1000) / 10])));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(plan), "Cutting Plan");
  const reuse = [["Reusable Bar Offcuts (leftover >= 1 m)"], [], ["Profile", "Grade", "From Bar", "Stock Length (mm)", "Reusable Leftover (mm)", "Weight (kg)"]];
  let rk = 0; results.groups.forEach(g => g.bins.forEach((b, i) => { if (b.remaining >= 1000) { rk++; reuse.push([g.profile, g.grade || "", `Bar ${i + 1}`, b.stockLength, Math.round(b.remaining), g.kgm ? ((b.remaining / 1000) * g.kgm).toFixed(1) : ""]); } }));
  if (rk === 0) reuse.push(["No reusable bar offcuts >= 1 m."]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(reuse), "Reusable Offcuts");
  saveWorkbook(wb, "SteelOptimizer_Sections.xlsx");
}

function exportSectionPDF(results, kerf, ow) {
  const procRows = results.groups.map(g => { const byLen = g.bins.reduce((a, b) => { a[b.stockLength] = (a[b.stockLength] || 0) + 1; return a; }, {}); return Object.entries(byLen).map(([len, qty]) => { const kg = g.kgm ? (parseInt(len) / 1000) * g.kgm * qty : null; return `<tr><td>${g.profile}</td><td>${g.grade || "—"}</td><td>${fmtMm(parseInt(len))}</td><td>${qty}</td><td>${g.kgm || "—"}</td><td>${kg ? fmtKg(kg) : "—"}</td></tr>`; }).join(""); }).join("");
  const planRows = results.groups.map(g => g.bins.map((b, i) => `<tr><td>${g.profile}</td><td>#${i + 1}</td><td>${fmtMm(b.stockLength)}</td><td>${b.cuts.map(c => fmtMm(c.length)).join(" + ")}</td><td>${fmtMm(b.remaining)}</td><td>${(((b.stockLength - b.remaining) / b.stockLength) * 100).toFixed(1)}%</td></tr>`).join("")).join("");
  const html = `<html><head><title>Steel Optimizer — Section Report</title><style>
  body{font-family:Georgia,serif;padding:32px;color:#1a1206}
  h1{color:#b45309;border-bottom:2px solid #f0d9b0;padding-bottom:10px}
  h2{color:#92400e;font-size:15px;margin-top:24px;border-bottom:1px solid #e6dcc8;padding-bottom:6px}
  table{width:100%;border-collapse:collapse;margin-top:12px;font-family:monospace;font-size:13px}
  th{background:#fdf4e3;color:#b45309;padding:9px;text-align:left;font-size:11px}
  td{padding:9px;border-bottom:1px solid #eee}
  .purchase{margin-top:18px;padding:18px;background:#fdf4e3;border:2px solid #f59e0b;border-radius:8px;font-size:18px}
  .purchase b{font-size:30px;color:#b45309}
  button{margin-top:24px;padding:12px 28px;background:#b45309;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer}
  @media print{button{display:none}}
  </style></head><body>
  <h1>Steel Optimizer — Section Procurement & Cutting</h1>
  <p>${new Date().toLocaleDateString("en-GB")} &nbsp;&middot;&nbsp; kerf ${kerf}mm</p>
  <div class="purchase">Total Steel to Procure: <b>${fmtKg(results.grandKg)}</b> &nbsp;&middot; ${ow.totalStock} bars &middot; waste ${ow.wastePct}% (${fmtKg(ow.wasteKg)})</div>
  <h2>Procurement — What to Order</h2>
  <table><tr><th>Profile</th><th>Grade</th><th>Market Length</th><th>Qty</th><th>kg/m</th><th>Total Weight</th></tr>${procRows}</table>
  <h2>Cutting Plan</h2>
  <table><tr><th>Profile</th><th>Bar #</th><th>Stock Length</th><th>Cuts</th><th>Offcut</th><th>Util</th></tr>${planRows}</table>
  <button onclick="window.print()">Print / Save as PDF</button></body></html>`;
  downloadHTML(html, "SteelOptimizer_Sections.html");
}
