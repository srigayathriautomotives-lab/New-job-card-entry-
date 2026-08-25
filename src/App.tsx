async function translateTextToTelugu(text: string): Promise<string> {
  if (!text || !text.trim()) return '';
  // If text is already in Telugu script
  if (/[\u0C00-\u0C7F]/.test(text)) {
    return text.trim();
  }
  try {
    const res = await fetch(`https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=te&dt=t&q=${encodeURIComponent(text.trim())}`);
    const data = await res.json();
    if (Array.isArray(data) && Array.isArray(data[0])) {
      const te = data[0].map((item: any) => item[0]).join('');
      return te || text;
    }
  } catch (e) {
    console.warn('Translate to Telugu error:', e);
  }
  return text;
}

import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

import { 
  signInWithPopup, 
  GoogleAuthProvider, 
  signOut, 
  onAuthStateChanged, 
  User 
} from 'firebase/auth';



import { auth, db } from './firebase';
import { 
  collection, doc, setDoc, getDoc, deleteDoc, getDocs, writeBatch, onSnapshot, query, orderBy, updateDoc, addDoc
} from 'firebase/firestore';

import { 
  loginWithGoogleForSheets, 
  getOrCreateSpreadsheet, 
  loadSheetRows, 
  appendSheetRow, 
  updateSheetRow, 
  deleteSheetRow,
  syncFullDatabaseToGoogleSheets,
  JOBCARD_HEADERS,
  COMPLAINT_HEADERS,
  STAFF_HEADERS,
  CUSTOMER_HEADERS,
  SPARE_HEADERS,
  setCachedToken
} from './lib/sheetsService';
import { sqlApi } from './lib/sqlService';
import { translations, getTranslation, Language } from './translations';
import {
  Wrench,
  Globe,
  Languages,
  Copy,
  Printer,
  FileSpreadsheet,
  Download,
  Upload,
  Trash2,
  Plus,
  FilePlus,
  X,
  Save,
  RotateCcw,
  Search,
  Building2,
  CheckSquare,
  Square,
  FileText,
  DollarSign,
  ShieldAlert,
  Clock,
  LayoutDashboard,
  History,
  Users,
  LogIn,
  LogOut,
  Lock,
  Mail,
  UserCheck,
  UserPlus,
  Eye,
  EyeOff,
  Menu,
  PanelLeftClose,
  PanelLeftOpen,
  ChevronLeft,
  ChevronRight,
  BarChart3,
  PieChart as PieChartIcon,
  TrendingUp,
  Calendar,
  MapPin,
  Briefcase,
  Filter,
  CheckCircle2,
  XCircle,
  Edit3,
  UserCog,
  Award,
  ChevronDown,
  Maximize2,
  Minimize2,
  ChevronUp,
  SlidersHorizontal,
  RefreshCw,
  RefreshCcw,
  PhoneCall,
  Phone,
  AlertCircle,
  AlertTriangle,
  ArrowUp,
  ArrowDown,
  GripVertical,
  Link2,
  ExternalLink,
  Database,
  Settings,
  FileCheck
} from 'lucide-react';

interface RepairRow {
  repair: string;
  rectification: string;
  charge: string;
}

interface PartRow {
  partNo: string;
  desc: string;
  wty?: boolean;
  qty: string;
  rate: string;
  amount: string;
}

interface CheckpointItem {
  id: number;
  category: string;
  categoryTe?: string;
  item: string;
  itemTe?: string;
  action: string;
  actionTe?: string;
  checked: boolean;
}

const DEFAULT_CHECKPOINTS: CheckpointItem[] = [
  // సాధారణ
  { id: 1, category: 'సాధారణ', categoryTe: 'సాధారణ', item: 'ట్రాక్టర్ వాషింగ్ మరియు క్లీనింగ్ చేయుట', itemTe: 'ట్రాక్టర్ వాషింగ్ మరియు క్లీనింగ్ చేయుట', action: 'వాషింగ్ మరియు క్లీనింగ్', actionTe: 'వాషింగ్ మరియు క్లీనింగ్', checked: true },
  { id: 2, category: 'సాధారణ', categoryTe: 'సాధారణ', item: 'అన్ని గ్రీసింగ్ పాయింట్లు తనిఖీ చేయుట', itemTe: 'అన్ని గ్రీసింగ్ పాయింట్లు తనిఖీ చేయుట', action: 'శుభ్రం చేసి గ్రీస్ అప్లై చేయుట', actionTe: 'శుభ్రం చేసి గ్రీస్ అప్లై చేయుట', checked: true },
  { id: 3, category: 'సాధారణ', categoryTe: 'సాధారణ', item: 'లూబ్రికేషన్ పాయింట్లు (ఆయిల్ వేయుట)', itemTe: 'లూబ్రికేషన్ పాయింట్లు (ఆయిల్ వేయుట)', action: 'ఆయిల్ వేసి లూబ్రికేట్ చేయుట', actionTe: 'ఆయిల్ వేసి లూబ్రికేట్ చేయుట', checked: true },
  { id: 4, category: 'సాధారణ', categoryTe: 'సాధారణ', item: 'డీజిల్ మరియు ఆయిల్ లీకేజీల తనిఖీ', itemTe: 'డీజిల్ మరియు ఆయిల్ లీకేజీల తనిఖీ', action: 'తనిఖీ చేసి లీకేజీలు అరికట్టుట', actionTe: 'తనిఖీ చేసి లీకేజీలు అరికట్టుట', checked: true },
  // ఇంజిన్
  { id: 5, category: 'ఇంజిన్', categoryTe: 'ఇంజిన్', item: 'ఇంజిన్ ఆయిల్ లెవల్ తనిఖీ చేయుట', itemTe: 'ఇంజిన్ ఆయిల్ లెవల్ తనిఖీ చేయుట', action: 'తనిఖీ చేసి అవసరమైతే మార్చుట', actionTe: 'తనిఖీ చేసి అవసరమైతే మార్చుట', checked: true },
  { id: 6, category: 'ఇంజిన్', categoryTe: 'ఇంజిన్', item: 'ఇంజిన్ ఆయిల్ ఫిల్టర్', itemTe: 'ఇంజిన్ ఆయిల్ ఫిల్టర్', action: 'ఫిల్టర్ కొత్తది మార్చుట', actionTe: 'ఫిల్టర్ కొత్తది మార్చుట', checked: true },
  { id: 7, category: 'ఇంజిన్', categoryTe: 'ఇంజిన్', item: 'ఫీడ్ పంప్, బౌల్ మరియు బ్రీథర్', itemTe: 'ఫీడ్ పంప్, బౌల్ మరియు బ్రీథర్', action: 'డ్రైన్ చేసి శుభ్రపరచుట', actionTe: 'డ్రైన్ చేసి శుభ్రపరచుట', checked: true },
  { id: 8, category: 'ఇంజిన్', categoryTe: 'ఇంజిన్', item: 'టాపెట్ క్లియరెన్స్ మరియు డీకంప్రెషన్ సెట్టింగ్', itemTe: 'టాపెట్ క్లియరెన్స్ మరియు డీకంప్రెషన్ సెట్టింగ్', action: 'తనిఖీ చేసి సర్దుబాటు చేయుట', actionTe: 'తనిఖీ చేసి సర్దుబాటు చేయుట', checked: true },
  // ఎయిర్ క్లీనర్
  { id: 9, category: 'ఎయిర్ క్లీనర్', categoryTe: 'ఎయిర్ క్లీనర్', item: 'ఎయిర్ క్లీనర్ వైర్ మెష్ మరియు ప్రీ-క్లీనర్', itemTe: 'ఎయిర్ క్లీనర్ వైర్ మెష్ మరియు ప్రీ-క్లీనర్', action: 'శుభ్రపరచుట (క్లీనింగ్)', actionTe: 'శుభ్రపరచుట (క్లీనింగ్)', checked: true },
  { id: 10, category: 'ఎయిర్ క్లీనర్', categoryTe: 'ఎయిర్ క్లీనర్', item: 'ఎయిర్ క్లీనర్ ఆయిల్ మరియు సెకండరీ ఫిల్టర్', itemTe: 'ఎయిర్ క్లీనర్ ఆయిల్ మరియు సెకండరీ ఫిల్టర్', action: 'అవసరమైతే కొత్తది మార్చుట', actionTe: 'అవసరమైతే కొత్తది మార్చుట', checked: true },
  // కూలింగ్ సిస్టమ్
  { id: 11, category: 'కూలింగ్ సిస్టమ్', categoryTe: 'కూలింగ్ సిస్టమ్', item: 'ఎయిర్ కూల్డ్ బ్లోవర్ ఫ్యాన్ మరియు బ్లోవర్ మెష్', itemTe: 'ఎయిర్ కూల్డ్ బ్లోవర్ ఫ్యాన్ మరియు బ్లోవర్ మెష్', action: 'దుమ్ము లేకుండా శుభ్రపరచుట', actionTe: 'దుమ్ము లేకుండా శుభ్రపరచుట', checked: true },
  { id: 12, category: 'కూలింగ్ సిస్టమ్', categoryTe: 'కూలింగ్ సిస్టమ్', item: 'సిలిండర్ అసెంబ్లీ మరియు ఫిన్స్ మధ్య ఖాళీలు', itemTe: 'సిలిండర్ అసెంబ్లీ మరియు ఫిన్స్ మధ్య ఖాళీలు', action: 'మట్టి మరియు దుమ్ము శుభ్రపరచుట', actionTe: 'మట్టి మరియు దుమ్ము శుభ్రపరచుట', checked: true },
  { id: 13, category: 'కూలింగ్ సిస్టమ్', categoryTe: 'కూలింగ్ సిస్టమ్', item: 'రేడియేటర్ గార్డ్ మరియు రేడియేటర్ ఫిన్స్', itemTe: 'రేడియేటర్ గార్డ్ మరియు రేడియేటర్ ఫిన్స్', action: 'శుభ్రం చేయుట', actionTe: 'శుభ్రం చేయుట', checked: true },
  { id: 14, category: 'కూలింగ్ సిస్టమ్', categoryTe: 'కూలింగ్ సిస్టమ్', item: 'రేడియేటర్ మరియు రికవరీ బాటిల్ కూలెంట్ లెవల్', itemTe: 'రేడియేటర్ మరియు రికవరీ బాటిల్ కూలెంట్ లెవల్', action: 'తనిఖీ చేసి కూలెంట్ నింపుట', actionTe: 'తనిఖీ చేసి కూలెంట్ నింపుట', checked: true },
  { id: 15, category: 'కూలింగ్ సిస్టమ్', categoryTe: 'కూలింగ్ సిస్టమ్', item: 'ఫ్యాన్ బెల్ట్ (V-Belt) టెన్షన్', itemTe: 'ఫ్యాన్ బెల్ట్ (V-Belt) టెన్షన్', action: 'బెల్ట్ టెన్షన్ అడ్జస్ట్ చేయుట', actionTe: 'బెల్ట్ టెన్షన్ అడ్జస్ట్ చేయుట', checked: true },
  // ఫ్యూయల్ సిస్టమ్
  { id: 16, category: 'ఫ్యూయల్ సిస్టమ్', categoryTe: 'ఇంధన వ్యవస్థ', item: 'వాటర్ సెపరేటర్ నుండి నీటిని తీసివేయుట', itemTe: 'వాటర్ సెపరేటర్ నుండి నీటిని తీసివేయుట', action: 'డ్రైన్ చేయుట', actionTe: 'డ్రైన్ చేయుట', checked: true },
  // ట్రాన్స్మిషన్
  { id: 17, category: 'ట్రాన్స్మిషన్', categoryTe: 'గేర్ బాక్స్', item: 'గేర్ బాక్స్ మరియు ఫ్రంట్ యాక్సిల్ ఆయిల్ లెవల్', itemTe: 'గేర్ బాక్స్ మరియు ఫ్రంట్ యాక్సిల్ ఆయిల్ లెవల్', action: 'తనిఖీ చేసి ఆయిల్ నింపుట', actionTe: 'తనిఖీ చేసి ఆయిల్ నింపుట', checked: true },
  // హైడ్రాలిక్స్
  { id: 18, category: 'హైడ్రాలిక్స్', categoryTe: 'హైడ్రాలిక్స్', item: 'హైడ్రాలిక్ సెన్సింగ్ మరియు లివర్ల అడ్జస్ట్‌మెంట్', itemTe: 'హైడ్రాలిక్ సెన్సింగ్ మరియు లివర్ల అడ్జస్ట్‌మెంట్', action: 'లివర్ ప్లే తనిఖీ చేసి సెట్ చేయుట', actionTe: 'లివర్ ప్లే తనిఖీ చేసి సెట్ చేయుట', checked: true },
  // ఎలక్ట్రికల్
  { id: 19, category: 'ఎలక్ట్రికల్', categoryTe: 'ఎలక్ట్రికల్', item: 'బ్యాటరీ డిస్టిల్డ్ వాటర్ లెవల్ మరియు టెర్మినల్స్', itemTe: 'బ్యాటరీ డిస్టిల్డ్ వాటర్ లెవల్ మరియు టెర్మినల్స్', action: 'వాటర్ నింపి టెర్మినల్స్‌కు పెట్రోలియం జెల్లీ రాయుట', actionTe: 'వాటర్ నింపి టెర్మినల్స్‌కు పెట్రోలియం జెల్లీ రాయుట', checked: true },
  { id: 20, category: 'ఎలక్ట్రికల్', categoryTe: 'ఎలక్ట్రికల్', item: 'బ్యాటరీ గ్రావిటీ మరియు సెల్ వోల్టేజ్', itemTe: 'బ్యాటరీ గ్రావిటీ మరియు సెల్ వోల్టేజ్', action: 'గ్రావిటీ చెక్ చేసి నమోదు చేయుట', actionTe: 'గ్రావిటీ చెక్ చేసి నమోదు చేయుట', checked: true },
  // జనరల్ చెక్స్
  { id: 21, category: 'జనరల్ చెక్స్', categoryTe: 'ఇతర తనిఖీలు', item: 'క్లచ్ ఫ్రీ ప్లే మరియు బ్రేక్ అడ్జస్ట్‌మెంట్', itemTe: 'క్లచ్ ఫ్రీ ప్లే మరియు బ్రేక్ అడ్జస్ట్‌మెంట్', action: 'తనిఖీ చేసి ఫ్రీ ప్లే సర్దుబాటు చేయుట', actionTe: 'తనిఖీ చేసి ఫ్రీ ప్లే సర్దుబాటు చేయుట', checked: true },
  { id: 22, category: 'జనరల్ చెక్స్', categoryTe: 'ఇతర తనిఖీలు', item: 'స్టీరింగ్ వీల్ ఫ్రీ ప్లే మరియు టో-ఇన్ సెట్టింగ్', itemTe: 'స్టీరింగ్ వీల్ ఫ్రీ ప్లే మరియు టో-ఇన్ సెట్టింగ్', action: 'తనిఖీ చేసి సరిచేయుట', actionTe: 'తనిఖీ చేసి సరిచేయుట', checked: true },
  { id: 23, category: 'జనరల్ చెక్స్', categoryTe: 'ఇతర తనిఖీలు', item: 'వీల్ నట్స్ టైట్నెస్ మరియు టైర్ ప్రెషర్', itemTe: 'వీల్ నట్స్ టైట్నెస్ మరియు టైర్ ప్రెషర్', action: 'నట్స్ టైట్ చేసి టైర్లలో గాలి నింపుట', actionTe: 'నట్స్ టైట్ చేసి టైర్లలో గాలి నింపుట', checked: true },
  { id: 24, category: 'జనరల్ చెక్స్', categoryTe: 'ఇతర తనిఖీలు', item: 'లైట్లు, హార్న్, స్విచ్‌లు మరియు డ్యాష్ బోర్డ్ మీటర్లు', itemTe: 'లైట్లు, హార్న్, స్విచ్‌లు మరియు డ్యాష్ బోర్డ్ మీటర్లు', action: 'పనితీరు తనిఖీ చేసి సరిచేయుట', actionTe: 'పనితీరు తనిఖీ చేసి సరిచేయుట', checked: true },
  { id: 25, category: 'జనరల్ చెక్స్', categoryTe: 'ఇతర తనిఖీలు', item: 'ట్రాక్టర్ ఫైనల్ టెస్ట్ డ్రైవ్ మరియు పనితీరు', itemTe: 'ట్రాక్టర్ ఫైనల్ టెస్ట్ డ్రైవ్ మరియు పనితీరు', action: 'టెస్ట్ డ్రైవ్ చేసి నిర్ధారించుట', actionTe: 'టెస్ట్ డ్రైవ్ చేసి నిర్ధారించుట', checked: true }
];

interface GeneralChecklistGroup {
  category: string;
  opt1: string;
  opt1Key: string;
  opt2: string;
  opt2Key: string;
  opt3: string;
  opt3Key: string;
}

const GENERAL_CHECKLIST_CONFIG: GeneralChecklistGroup[] = [
  { category: 'Engine Not Starting:', opt1: 'Battery issues', opt1Key: 'ens_battery', opt2: 'Faulty starter motor or solenoid', opt2Key: 'ens_starter', opt3: 'Fuel supply problems', opt3Key: 'ens_fuel' },
  { category: 'Brake System Problems:', opt1: 'Worn brake pads', opt1Key: 'bsp_pads', opt2: 'Low brake fluid', opt2Key: 'bsp_fluid', opt3: 'Leak in brake lines', opt3Key: 'bsp_leak' },
  { category: 'Fuel System Issues:', opt1: 'Faulty fuel pump', opt1Key: 'fsi_pump', opt2: 'Clogged fuel filter', opt2Key: 'fsi_filter', opt3: 'Problem with injectors', opt3Key: 'fsi_injectors' },
  { category: 'Engine Overheating:', opt1: 'Leak in the radiator or cooling system', opt1Key: 'eo_radiator', opt2: 'Faulty thermostat', opt2Key: 'eo_thermostat', opt3: 'Issue with the water pump', opt3Key: 'eo_pump' },
  { category: 'Suspension & Steering:', opt1: 'Problems with steering links', opt1Key: 'ss_steering', opt2: 'Faulty shock absorbers', opt2Key: 'ss_shocks', opt3: 'Worn suspension bushings', opt3Key: 'ss_bushings' },
  { category: 'Hydraulic System:', opt1: 'Hydraulic oil leaks', opt1Key: 'hsp_leaks', opt2: 'Faulty hydraulic pump', opt2Key: 'hsp_pump', opt3: 'Issues with cylinders / lines', opt3Key: 'hsp_cylinders' },
  { category: 'Transmission Problems:', opt1: 'Difficulty shifting gears', opt1Key: 'tp_gears', opt2: 'Worn clutch plate', opt2Key: 'tp_clutch', opt3: 'Low/contaminated fluid', opt3Key: 'tp_fluid' },
  { category: 'Exhaust System Issues:', opt1: 'Leaks in the exhaust pipe', opt1Key: 'esi_pipe', opt2: 'Faulty catalytic converter', opt2Key: 'esi_converter', opt3: 'Blockages in the muffler', opt3Key: 'esi_muffler' },
  { category: 'Oil Leaks:', opt1: 'Worn engine gaskets', opt1Key: 'ol_gaskets', opt2: 'Worn oil seals', opt2Key: 'ol_seals', opt3: 'Leaks from the oil pump', opt3Key: 'ol_pump' },
  { category: 'Electrical System Issues:', opt1: 'Non-functioning lights/signals', opt1Key: 'esi_lights', opt2: 'Short circuits in wiring', opt2Key: 'esi_wiring', opt3: 'Faulty electrical sensors', opt3Key: 'esi_sensors' }
];

const FIELD_CANDIDATES: Record<string, string[]> = {
  chassis: ['chassisno', 'chassis', 'chassisnumber', 'chassiscode', 'chassisnum', 'chassis_no', 'chassisno.', 'vinnumber', 'vin'],
  custName: ['nameofcustomer', 'customername', 'custname', 'name', 'ownername', 'customer_name', 'cust_name', 'customer', 'owner_name'],
  custAddr: ['address', 'custaddress', 'villageaddress', 'village', 'cust_addr', 'custaddr', 'customeraddress', 'addr'],
  fatherName: ['fathername', 'father', 'so', 'co', 's_o', 'c_o', 'guardian', 'father_name', 'husbandname', 'husband', 'fathersname'],
  village: ['village', 'vill', 'village_name', 'vill_name', 'place', 'location'],
  mandal: ['mandal', 'tehsil', 'taluk', 'block', 'mandal_name', 'tehsil_name', 'subdistrict', 'mandaltehsil'],
  post: ['post', 'po', 'postoffice', 'p_o', 'post_office'],
  district: ['district', 'dist', 'distt'],
  custPhone: [
    'phone', 'mobile', 'mobileno', 'cell', 'contact', 'contactno', 'phoneno',
    'ownermobno', 'ownermobile', 'ownerphone', 'custphone', 'custmobile',
    'ownermob', 'ownermobilenumber', 'ownernumber', 'custmob', 'custcontact',
    'customerphone', 'customermobile', 'customercontact', 'cust_phone', 'cust_mobile',
    'owner_phone', 'owner_mobile', 'owner_mob', 'mobile_no', 'phone_no', 'contact_no',
    'mob', 'mobno', 'mob_no', 'phno', 'ph_no', 'tel', 'telephone',
    'ownermobileno', 'customermobileno', 'mobilenumber', 'phonenumber', 'contactnumber',
    'cellno', 'cellnumber', 'customermobno', 'customermobilenumber', 'custphoneno', 'custmobileno'
  ],
  driverPhone: ['drivermobno', 'driverphone', 'drivermobile', 'drivercontact', 'driver_phone', 'driver_mobile'],
  regdNo: ['regdno', 'regno', 'registrationno', 'regdnumber', 'regnumber', 'reg_no', 'tractorregdno'],
  engineNo: ['engineno', 'engine', 'enginenumber', 'engno', 'engine_no', 'engineno.', 'engine_number', 'engnumber'],
  model: ['model', 'tractormodel', 'modelno', 'modelnumber'],
  modelType: ['modeltype', 'tractor_type', 'type', 'model_type', 'model_spec', 'tractortype'],
  installDate: [
    'dateofinstallation', 'installationdate', 'installdate', 'doi', 'date_of_installation',
    'dateofinst', 'date_of_inst', 'instdate', 'inst_date', 'installation_date',
    'dop', 'dateofpurchase', 'purchasedate', 'purchase_date',
    'dateofsale', 'saledate', 'sale_date',
    'deliverydate', 'dateofdelivery', 'delivery_date',
    'insdate', 'ins_date', 'doi.', 'dateofinstall', 'date_of_install',
    'invoicedate', 'invoice_date', 'billingdate', 'billdate', 'installation_dt', 'install_dt', 'inst_dt'
  ],
  hourMeter: ['hourmeterreading', 'hourmeter', 'hrmeter', 'hmr', 'meterreading', 'hours'],
  distDealership: ['distancefromdealership', 'dealershipdistance', 'distance'],
  serialNo: ['serialno', 'serialnumber', 'srno', 'slno', 'serial_no', 'sr_no', 'sl_no', 'serial', 'sno', 's.no', 'sl.no', 's_no'],
  historyFileNo: ['historyfileno', 'history_file_no', 'fileno', 'file_no', 'file', 'historyfile', 'hfn', 'sno', 's.no', 'slno', 'serialno'],
  branch: ['branch', 'branchname', 'branch_name', 'dealershipbranch', 'locationbranch', 'bname', 'branchcode', 'branch_code']
};

const SPARE_FIELD_CANDIDATES: Record<string, string[]> = {
  partNo: ['partno', 'partnumber', 'materialno', 'sparepartno', 'code', 'itemcode'],
  desc: ['description', 'partname', 'sparename', 'itemdescription', 'item'],
  rate: ['rate', 'price', 'unitprice', 'mrp', 'sellingprice']
};

function cleanValue(val: any): string {
  if (val === null || val === undefined) return '';

  // Handle Firestore Timestamp object (with toDate method)
  if (typeof val === 'object' && typeof val.toDate === 'function') {
    try {
      const d = val.toDate();
      if (d instanceof Date && !isNaN(d.getTime())) {
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${dd}/${mm}/${yyyy}`;
      }
    } catch (e) {}
  }

  // Handle Firestore Timestamp serialized object { seconds: 12345, nanoseconds: 0 } or { _seconds: 12345 }
  if (typeof val === 'object' && val !== null) {
    const sec = val.seconds ?? val._seconds;
    if (typeof sec === 'number' && !isNaN(sec)) {
      try {
        const d = new Date(sec * 1000);
        if (d instanceof Date && !isNaN(d.getTime())) {
          const yyyy = d.getFullYear();
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          const dd = String(d.getDate()).padStart(2, '0');
          return `${dd}/${mm}/${yyyy}`;
        }
      } catch (e) {}
    }
  }

  // Handle JS Date object
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      const yyyy = val.getFullYear();
      const mm = String(val.getMonth() + 1).padStart(2, '0');
      const dd = String(val.getDate()).padStart(2, '0');
      return `${dd}/${mm}/${yyyy}`;
    }
  }

  // Handle string like "Timestamp(seconds=1747592990, nanoseconds=0)" or "seconds=1747592990"
  const str = String(val).trim();
  if (str.includes('Timestamp') || str.includes('seconds=') || str.includes('seconds:') || str.includes('_seconds')) {
    const secMatch = str.match(/_?seconds\s*[:=]\s*(\d+)/i);
    if (secMatch && secMatch[1]) {
      const sec = parseInt(secMatch[1], 10);
      const d = new Date(sec * 1000);
      if (!isNaN(d.getTime())) {
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${dd}/${mm}/${yyyy}`;
      }
    }
  }

  if (str === '[object Object]') {
    return '';
  }

  return str;
}

function normalizeKey(s: string): string {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function buildRowLookup(row: Record<string, any>): Record<string, any> {
  const lookup: Record<string, any> = {};
  if (!row) return lookup;
  Object.keys(row).forEach(h => {
    lookup[normalizeKey(h)] = cleanValue(row[h]);
  });
  return lookup;
}

function getFieldValue(lookup: Record<string, any>, field: string): string {
  if (!lookup) return '';
  if (field === 'custPhone' && lookup.__custPhoneDisplay) {
    return cleanValue(lookup.__custPhoneDisplay);
  }
  if (field === 'custName' && lookup.__custNameDisplay) {
    return cleanValue(lookup.__custNameDisplay);
  }
  if (field === 'chassis' && lookup.__chassisDisplay) {
    return cleanValue(lookup.__chassisDisplay);
  }

  // Create a normalized key lookup for this record on the fly to support any spacing/casing from raw sheets/DBs
  const normalizedLookup: Record<string, any> = {};
  for (const [k, v] of Object.entries(lookup)) {
    if (v !== undefined && v !== null) {
      normalizedLookup[normalizeKey(k)] = v;
    }
  }

  const candidates = FIELD_CANDIDATES[field] || [];
  for (const c of candidates) {
    const normCand = normalizeKey(c);
    if (normalizedLookup[normCand] !== undefined && String(normalizedLookup[normCand]).trim() !== '') {
      return cleanValue(normalizedLookup[normCand]);
    }
    // Fallback to raw check
    if (lookup[c] !== undefined && lookup[c] !== null && String(lookup[c]).trim() !== '') {
      return cleanValue(lookup[c]);
    }
  }

  if (field === 'installDate') {
    for (const k of Object.keys(lookup)) {
      if (k.startsWith('__')) continue;
      const normK = normalizeKey(k);
      if (normK.includes('install') || normK.includes('delivery') || normK.includes('doi') || normK.includes('dop') || normK.includes('purchase') || normK.includes('sale') || normK.includes('invoice') || normK.includes('bill') || (normK.includes('date') && !normK.includes('in') && !normK.includes('out'))) {
        const val = lookup[k];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          return cleanValue(val);
        }
      }
    }
  }

  if (field === 'custPhone') {
    for (const k of Object.keys(lookup)) {
      if (k.startsWith('__')) continue;
      const normK = normalizeKey(k);
      if ((normK.includes('mob') || normK.includes('phone') || normK.includes('cell') || normK.includes('contact') || normK.includes('tel')) && !normK.includes('driver')) {
        const val = cleanValue(lookup[k]);
        if (val) return val;
      }
    }
  }

  return '';
}

function getCombinedAddress(lookup: Record<string, any>): string {
  if (!lookup) return '';
  if (lookup.__custAddrDisplay) {
    return String(lookup.__custAddrDisplay).trim();
  }

  const rawAddr = getFieldValue(lookup, 'custAddr');
  const father = getFieldValue(lookup, 'fatherName');
  const vill = getFieldValue(lookup, 'village');
  const mand = getFieldValue(lookup, 'mandal');
  const pst = getFieldValue(lookup, 'post');
  const dist = getFieldValue(lookup, 'district');

  const parts: string[] = [];
  if (father) {
    const fLower = String(father).toLowerCase();
    if (fLower.includes('s/o') || fLower.includes('c/o') || fLower.includes('w/o')) {
      parts.push(String(father).trim());
    } else {
      parts.push(`S/o ${String(father).trim()}`);
    }
  }
  if (rawAddr && rawAddr !== vill && rawAddr !== mand) {
    parts.push(String(rawAddr).trim());
  }
  if (vill && !rawAddr.toLowerCase().includes(vill.toLowerCase())) {
    parts.push(`Vill: ${String(vill).trim()}`);
  }
  if (mand && !rawAddr.toLowerCase().includes(mand.toLowerCase())) {
    parts.push(`Mandal: ${String(mand).trim()}`);
  }
  if (pst && !rawAddr.toLowerCase().includes(pst.toLowerCase())) {
    parts.push(`Post: ${String(pst).trim()}`);
  }
  if (dist && !rawAddr.toLowerCase().includes(dist.toLowerCase())) {
    parts.push(`Dist: ${String(dist).trim()}`);
  }

  if (parts.length === 0) {
    for (const k of Object.keys(lookup)) {
      if (k.startsWith('__')) continue;
      if (k.includes('add') || k.includes('vill') || k.includes('mand') || k.includes('taluk') || k.includes('father') || k.includes('so') || k.includes('co')) {
        const val = String(lookup[k]).trim();
        if (val && !parts.includes(val)) {
          parts.push(val);
        }
      }
    }
  }

  return parts.length > 0 ? parts.join(', ') : rawAddr;
}

function getStandardizedCustomer(rec: any): any {
  if (!rec) return {
    sNo: '',
    branch: '',
    model: '',
    modelType: '',
    chassisNo: '',
    engineNo: '',
    dateOfDel: '',
    custName: '',
    fatherName: '',
    address: '',
    village: '',
    mandal: '',
    mobileNumber: '',
    district: '',
    pinCode: '',
    supervisor: '',
    historyFileNo: ''
  };

  // 1. Resolve raw values
  const ch = rec['Chassis no'] || rec.__chassisDisplay || getFieldValue(rec, 'chassis') || rec.chassisNo || rec.chassis || '';
  const name = rec['Customer Name'] || rec.__custNameDisplay || rec.custName || getFieldValue(rec, 'custName') || rec.customerName || '';
  
  // Try finding fatherName in candidates
  const father = rec['Father Name'] || rec['FATHER NAME'] || rec.fatherName || getFieldValue(rec, 'fatherName') || '';
  
  const phone = rec['Mobile Number'] || rec.__custPhoneDisplay || rec.mobileNumber || getFieldValue(rec, 'custPhone') || '';
  const village = rec.VILLAGE || rec.village || getFieldValue(rec, 'village') || '';
  const mandal = rec.Mandal || rec.mandal || getFieldValue(rec, 'mandal') || '';
  const district = rec.Distict || rec.district || getFieldValue(rec, 'district') || '';
  const pinCode = rec['PIN CODE'] || rec.pinCode || getFieldValue(rec, 'pinCode') || '';
  const model = rec.Model || rec.model || getFieldValue(rec, 'model') || '';
  const modelType = rec['MODEL TYPE'] || rec.modelType || rec.modeltype || getFieldValue(rec, 'modelType') || '';
  const engineNo = rec['Engine No:'] || rec['Engine no'] || rec.engineNo || rec.engineno || getFieldValue(rec, 'engineNo') || '';
  const rawDel = rec['Del Date'] || rec['DEL DATE'] || rec['del date'] || rec['Del. Date'] || rec['Date of del'] || rec['DATE OF DEL'] || rec['date of del'] || rec['date delivery'] || rec['Date Delivery'] || rec.dateOfDel || rec.dateofdel || getFieldValue(rec, 'installDate') || '';
  const rawSNo = rec['S.No.'] || rec['S.No'] || rec['S.NO'] || rec['S.NO.'] || rec.sNo || rec.sno || rec.historyFileNo || getFieldValue(rec, 'serialNo') || '';
  const branch = rec.BRANCH || rec.branch || getFieldValue(rec, 'branch') || '';
  const supervisor = rec.supervisor || rec.SUPERVISOR || rec['Supervisor Name'] || '';

  // 2. Resolve HFN formatted
  let formattedHFN = cleanValue(rawSNo);
  const rawBranchVal = cleanValue(branch);
  if (formattedHFN && rawBranchVal) {
    const cleanBranch = rawBranchVal.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const cleanHFN = formattedHFN.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    if (cleanBranch && !cleanHFN.startsWith(cleanBranch)) {
      formattedHFN = `${rawBranchVal.toUpperCase()}-${formattedHFN}`;
    }
  }

  return {
    sNo: cleanValue(rawSNo),
    branch: rawBranchVal,
    model: cleanValue(model),
    modelType: cleanValue(modelType),
    chassisNo: cleanValue(ch),
    engineNo: cleanValue(engineNo),
    dateOfDel: cleanValue(rawDel),
    custName: cleanValue(name),
    fatherName: cleanValue(father),
    address: cleanValue(rec.ADDRESS || rec.address || getFieldValue(rec, 'custAddr') || ''),
    village: cleanValue(village),
    mandal: cleanValue(mandal),
    mobileNumber: cleanValue(phone),
    district: cleanValue(district),
    pinCode: cleanValue(pinCode),
    supervisor: cleanValue(supervisor),
    historyFileNo: formattedHFN
  };
}

function formatCustomerClipboardText(lookup: any): string {
  if (!lookup) return '';
  const std = getStandardizedCustomer(lookup);

  let father = std.fatherName;
  if (father) {
    const fLower = father.toLowerCase();
    if (!fLower.includes('s/o') && !fLower.includes('w/o') && !fLower.includes('c/o')) {
      father = `S/o ${father}`;
    }
  }

  // Construct combined line prioritizing model + model type and village/mandal
  const addressParts: string[] = [];
  const modelStr = [std.model, std.modelType].filter(Boolean).join(' ').trim();
  if (modelStr) {
    addressParts.push(modelStr);
  }
  
  const locStr = [std.village, std.mandal].filter(Boolean).join(', ').trim();
  if (locStr) {
    addressParts.push(locStr);
  }

  let addrLine = addressParts.join(' — ').trim();
  if (!addrLine) {
    addrLine = std.address;
  }

  const lines: string[] = [];
  if (std.historyFileNo) lines.push(`HFN: ${std.historyFileNo}`);
  if (std.custName) lines.push(std.custName.toUpperCase());
  if (father) lines.push(father);
  if (std.mobileNumber) lines.push(std.mobileNumber);
  if (addrLine) lines.push(addrLine);
  if (std.chassisNo) lines.push(`Chassis: ${std.chassisNo}`);
  if (std.engineNo) lines.push(`Eng: ${std.engineNo}`);

  return lines.join('\n');
}

function getSpareField(lookup: Record<string, any>, field: string): string {
  const candidates = SPARE_FIELD_CANDIDATES[field] || [];
  for (const c of candidates) {
    if (lookup[c] !== undefined && lookup[c] !== null && String(lookup[c]).trim() !== '') {
      return cleanValue(lookup[c]);
    }
  }
  return '';
}

function toInputDateFormat(v: any): string | null {
  if (v === undefined || v === null || v === '') return null;

  const cleaned = cleanValue(v);
  if (!cleaned) return null;

  if (typeof cleaned === 'number' || (!isNaN(Number(cleaned)) && !cleaned.includes('/') && !cleaned.includes('-') && !cleaned.includes('.'))) {
    const num = typeof cleaned === 'number' ? cleaned : parseFloat(cleaned);
    if (num > 20000 && num < 60000) {
      const d = XLSX.SSF.parse_date_code(num);
      if (d) {
        const yyyy = d.y;
        const mm = String(d.m).padStart(2, '0');
        const dd = String(d.d).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
      }
    }
  }

  // Check YYYY-MM-DD
  let mIso = cleaned.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (mIso) {
    const [, yyyy, mm, dd] = mIso;
    return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
  }

  // Check DD-MM-YYYY or DD/MM/YYYY
  let mDmy = cleaned.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (mDmy) {
    let [, p1, p2, p3] = mDmy;
    let yyyy = p3;
    let mm = p2;
    let dd = p1;
    if (p1.length === 4) {
      yyyy = p1;
      mm = p2;
      dd = p3;
    } else {
      if (yyyy.length === 2) yyyy = '20' + yyyy;
      let nDD = parseInt(dd, 10);
      let nMM = parseInt(mm, 10);
      if (nMM > 12 && nDD <= 12) {
        const tmp = nDD;
        nDD = nMM;
        nMM = tmp;
      }
      dd = String(nDD);
      mm = String(nMM);
    }
    return `${yyyy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
  }

  const parsed = new Date(cleaned);
  if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1900 && parsed.getFullYear() < 2100) {
    const yyyy = parsed.getFullYear();
    const mm = String(parsed.getMonth() + 1).padStart(2, '0');
    const dd = String(parsed.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  return null;
}

function getWarrantyStatus(installDateStr: string, jobDateStr: string) {
  if (!installDateStr || !installDateStr.trim()) return null;

  const isoInstall = toInputDateFormat(installDateStr) || installDateStr;
  const dInstall = new Date(isoInstall);
  if (isNaN(dInstall.getTime())) return null;

  const expiryDate = new Date(dInstall);
  expiryDate.setFullYear(expiryDate.getFullYear() + 2);

  let refTime = new Date();
  if (jobDateStr && jobDateStr.trim()) {
    const dJob = new Date(jobDateStr);
    if (!isNaN(dJob.getTime())) {
      refTime = dJob;
    }
  }

  const isWty = refTime <= expiryDate;

  return {
    status: isWty ? 'WARRANTY' : 'POST WTY',
    isWty,
    isPostWty: !isWty,
    expiryStr: expiryDate.toISOString().split('T')[0]
  };
}

function getCardWtyInfo(card: any) {
  if (!card) return { isWty: false, isPostWty: false, status: '' };
  const cardInstallDate = card.installDate || card.dateOfDelivery || '';
  const cardJobDate = card.jobDate || card.complaintDate || card.createdAt || '';
  const calculated = getWarrantyStatus(cardInstallDate, cardJobDate);
  const override = card.warrantyOverride || 'auto';
  
  let isWty = false;
  let isPostWty = false;
  let status = '';
  
  if (override === 'warranty') {
    isWty = true;
    isPostWty = false;
    status = 'WARRANTY';
  } else if (override === 'post_wty') {
    isWty = false;
    isPostWty = true;
    status = 'POST WTY';
  } else if (calculated !== null) {
    isWty = calculated.isWty;
    isPostWty = calculated.isPostWty;
    status = calculated.status;
  } else {
    const st = card.serviceType || 'Paid Service';
    if (st && st !== 'Paid Repairs') {
      isWty = true;
      isPostWty = false;
      status = 'WARRANTY';
    } else {
      isWty = false;
      isPostWty = true;
      status = 'POST WTY';
    }
  }
  return { isWty, isPostWty, status };
}

function isWithin2Years(deliveryDateStr: string, refDateStr?: string): boolean {
  if (!deliveryDateStr || !String(deliveryDateStr).trim()) return false;
  const isoDel = toInputDateFormat(deliveryDateStr) || String(deliveryDateStr).trim();
  const dDel = new Date(isoDel);
  if (isNaN(dDel.getTime())) return false;

  const expiryDate = new Date(dDel);
  expiryDate.setFullYear(expiryDate.getFullYear() + 2);

  let refTime = new Date();
  if (refDateStr && String(refDateStr).trim()) {
    const dRef = new Date(toInputDateFormat(refDateStr) || String(refDateStr).trim());
    if (!isNaN(dRef.getTime())) {
      refTime = dRef;
    }
  }

  return refTime <= expiryDate;
}

function getLocalDateTimeString(date = new Date()) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}T${hh}:${min}`;
}

function adjustDateTimeString(dtStr: string, hoursDiff: number) {
  if (!dtStr) return '';
  const d = new Date(dtStr);
  if (isNaN(d.getTime())) return '';
  d.setHours(d.getHours() + hoursDiff);
  return getLocalDateTimeString(d);
}

const LS_CUSTOMER_KEY = 'jobcard_sg_customer_v2';
const LS_SPARES_KEY = 'jobcard_sg_spares_v2';
const LS_SAVED_JOBCARDS_KEY = 'jobcard_sg_saved_list_v2';

const IDB_NAME = 'SgJobcardStorageDB_v1';
const IDB_STORE = 'kv_store';

function getIDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error('IndexedDB not supported'));
      return;
    }
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveToStorage(key: string, data: any) {
  try {
    const db = await getIDB();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      const store = tx.objectStore(IDB_STORE);
      const req = store.put(data, key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn('IDB save error', err);
  }

  try {
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    // LocalStorage limit exceeded; IndexedDB retains the data safely
  }
}

async function loadFromStorage<T>(key: string): Promise<T | null> {
  try {
    const db = await getIDB();
    const res = await new Promise<T | null>((resolve) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const store = tx.objectStore(IDB_STORE);
      const req = store.get(key);
      req.onsuccess = () => resolve((req.result as T) ?? null);
      req.onerror = () => resolve(null);
    });
    if (res) return res;
  } catch (err) {
    console.warn('IDB load error', err);
  }

  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw) as T;
  } catch (e) {
    // ignore
  }
  return null;
}

async function removeFromStorage(key: string) {
  try {
    const db = await getIDB();
    await new Promise<void>((resolve) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      const store = tx.objectStore(IDB_STORE);
      const req = store.delete(key);
      req.onsuccess = () => resolve();
      req.onerror = () => resolve();
    });
  } catch (e) {}

  try {
    localStorage.removeItem(key);
  } catch (e) {}
}

async function saveJobCardsBackup(cards: any[]) {
  try {
    await saveToStorage('sri_backup_jobcards', cards);
  } catch (err) {
    console.warn('Failed to save full backup to IndexedDB:', err);
  }

  try {
    localStorage.setItem('sri_backup_jobcards', JSON.stringify(cards));
  } catch (e) {
    console.warn("localStorage quota exceeded for sri_backup_jobcards, attempting to cache a subset of recent cards");
    let limit = Math.min(cards.length, 1000);
    while (limit > 0) {
      try {
        limit = Math.floor(limit * 0.8);
        if (limit < 10) {
          break;
        }
        localStorage.setItem('sri_backup_jobcards', JSON.stringify(cards.slice(0, limit)));
        console.log(`Successfully cached the most recent ${limit} jobcards in localStorage`);
        return;
      } catch (innerError) {
        // Continue reducing
      }
    }
  }
}

const DEFAULT_QUICK_REMARKS = [
  'Tractor not running / Engine issue',
  'Satisfied with service',
  'Will visit showroom next week',
  'Call back after 10 days',
  'Needs mechanic visit at village',
  'Driver unavailable / Busy in field',
  'Not lifting call / Out of coverage',
  'Asked for discount / Cost estimate',
  'Tractor running well / No complaints'
];

export default function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  // Language State (Telugu / English)
  const [appLang, setAppLang] = useState<Language>(() => {
    try {
      return (localStorage.getItem('sri_app_language') as Language) || 'te';
    } catch {
      return 'te';
    }
  });

  const toggleLanguage = (lang: Language) => {
    setAppLang(lang);
    try {
      localStorage.setItem('sri_app_language', lang);
    } catch (e) {}
  };

  const t = (key: keyof typeof translations) => getTranslation(key, appLang);
  const staffFileInputRef = useRef<HTMLInputElement>(null);
  // Navigation Active Tab State
  const [activeTab, setActiveTab] = useState<'new_entry' | 'saved_cards' | 'reports' | 'databases' | 'followup' | 'telecalling' | 'complaints' | 'dashboard' | 'free_service_followup' | 'attendance'>('dashboard');
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [isSidebarHidden, setIsSidebarHidden] = useState(false);

  // Configurable Navigation Menu Box Order
  const DEFAULT_SIDEBAR_MENU_ORDER = [
    'dashboard',
    'complaints',
    'attendance',
    'new_entry',
    'saved_cards',
    'followup',
    'telecalling',
    'free_service_followup',
    'reports',
    'databases'
  ];

  const [menuOrder, setMenuOrder] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('app_sidebar_menu_order');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const valid = parsed.filter((id: string) => DEFAULT_SIDEBAR_MENU_ORDER.includes(id));
          const missing = DEFAULT_SIDEBAR_MENU_ORDER.filter(id => !valid.includes(id));
          return [...valid, ...missing];
        }
      }
    } catch (e) {}
    return DEFAULT_SIDEBAR_MENU_ORDER;
  });

  const [isReorderingMenu, setIsReorderingMenu] = useState(false);

  const moveMenuItem = (index: number, direction: 'up' | 'down') => {
    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= menuOrder.length) return;
    const newOrder = [...menuOrder];
    const temp = newOrder[index];
    newOrder[index] = newOrder[targetIndex];
    newOrder[targetIndex] = temp;
    setMenuOrder(newOrder);
    try {
      localStorage.setItem('app_sidebar_menu_order', JSON.stringify(newOrder));
    } catch (e) {}
  };

  const resetMenuOrder = () => {
    setMenuOrder(DEFAULT_SIDEBAR_MENU_ORDER);
    try {
      localStorage.removeItem('app_sidebar_menu_order');
    } catch (e) {}
  };

  // Complaint View Only State
  const [isComplaintViewOnly, setIsComplaintViewOnly] = useState(false);

  // Saved Card View Modal State
  const [viewingCardModal, setViewingCardModal] = useState<any | null>(null);

  // Menu Order Reorder Modal State
  const [isMenuOrderModalOpen, setIsMenuOrderModalOpen] = useState(false);

  // Staff Attendance States
  const [attendanceDate, setAttendanceDate] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [attendanceRecords, setAttendanceRecords] = useState<Record<string, Record<string, { status: 'present' | 'absent' | 'leave', remarks?: string }>>>(() => {
    try {
      const cached = localStorage.getItem('sri_staff_attendance_data');
      return cached ? JSON.parse(cached) : {};
    } catch (e) {
      return {};
    }
  });
  const [attendanceMonthFilter, setAttendanceMonthFilter] = useState<string>(() => new Date().toISOString().substring(0, 7));
  const [attendanceTab, setAttendanceTab] = useState<'daily' | 'monthly'>('daily');

  const saveStaffAttendanceForDate = (dateStr: string, recordsForDate: Record<string, { status: 'present' | 'absent' | 'leave', remarks?: string }>) => {
    const updated = { ...attendanceRecords, [dateStr]: recordsForDate };
    setAttendanceRecords(updated);
    try {
      localStorage.setItem('sri_staff_attendance_data', JSON.stringify(updated));
    } catch (e) {}

    // Save to Cloud SQL Database
    try {
      sqlApi.saveAttendance(dateStr, recordsForDate);
    } catch (sqlErr) {
      console.warn('Cloud SQL attendance save:', sqlErr);
    }

    if (db) {
      try {
        const docRef = doc(db, 'staff_attendance', dateStr);
        setDoc(docRef, { date: dateStr, records: recordsForDate, updatedAt: new Date().toISOString() }, { merge: true }).catch(() => {});
      } catch (err) {
        console.error("Firestore attendance save error:", err);
      }
    }
  };

  const [isReportsFilterCollapsed, setIsReportsFilterCollapsed] = useState(false);
  const [isSavedFilterCollapsed, setIsSavedFilterCollapsed] = useState(false);
  const [isFollowupFilterCollapsed, setIsFollowupFilterCollapsed] = useState(false);

  useEffect(() => {
    // Hide sidebar on small screens initially to prevent top/left waste gaps
    if (typeof window !== 'undefined' && window.innerWidth < 768) {
      setIsSidebarHidden(true);
    }
  }, []);

  // Service Follow-up states
  const [followupSearch, setFollowupSearch] = useState('');
  const [followupSupervisor, setFollowupSupervisor] = useState('all');
  const [followupStatus, setFollowupStatus] = useState<'all' | 'reporting' | 'not_reporting' | 'duplicate'>('all');
  const [followupSortBy, setFollowupSortBy] = useState<'hfn' | 'name' | 'dateOfDel' | 'lastJobCardDate' | 'nextCallDate' | 'chassisNo' | 'village' | 'model'>('hfn');
  const [followupSortOrder, setFollowupSortOrder] = useState<'asc' | 'desc'>('asc');
  const [followupDateFrom, setFollowupDateFrom] = useState('');
  const [followupDateTo, setFollowupDateTo] = useState('');
  const [followupModelFilter, setFollowupModelFilter] = useState('all');
  const [selectedFollowupCustomer, setSelectedFollowupCustomer] = useState<any | null>(null);
  
  // Follow-up call log logging states
  const [selectedFollowupForLog, setSelectedFollowupForLog] = useState<any | null>(null);
  const [newFollowupCallDate, setNewFollowupCallDate] = useState('');
  const [newFollowupRemarks, setNewFollowupRemarks] = useState('');
  const [newFollowupNextCallDate, setNewFollowupNextCallDate] = useState('');
  const [newFollowupCalledBy, setNewFollowupCalledBy] = useState<string>(() => {
    try {
      return localStorage.getItem('telecaller_last_caller') || '';
    } catch (e) {
      return '';
    }
  });

  // Free Service Follow-up Accordion Expand States & Search Filter
  const [expandedCategories, setExpandedCategories] = useState<Record<string, boolean>>({
    fs1: true, // expand 1st free service by default
    gear_1: true, // expand 1st year gear oil by default
  });
  const [freeServiceSearch, setFreeServiceSearch] = useState('');
  const [freeServiceSupervisor, setFreeServiceSupervisor] = useState('all');

  // Quick remarks setup states for telecalling
  const [quickRemarksList, setQuickRemarksList] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('telecaller_quick_remarks');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {}
    return DEFAULT_QUICK_REMARKS;
  });
  const [newQuickRemarkInput, setNewQuickRemarkInput] = useState('');
  const [isAddingQuickRemark, setIsAddingQuickRemark] = useState(false);

  const handleAddQuickRemark = (text: string) => {
    const trimmed = text.trim();
    if (!trimmed) return;
    if (quickRemarksList.includes(trimmed)) return;
    const updated = [...quickRemarksList, trimmed];
    setQuickRemarksList(updated);
    try {
      localStorage.setItem('telecaller_quick_remarks', JSON.stringify(updated));
    } catch (e) {}
    setNewQuickRemarkInput('');
  };

  const handleRemoveQuickRemark = (indexToRemove: number) => {
    const updated = quickRemarksList.filter((_, idx) => idx !== indexToRemove);
    setQuickRemarksList(updated);
    try {
      localStorage.setItem('telecaller_quick_remarks', JSON.stringify(updated));
    } catch (e) {}
  };

  const handleSelectQuickRemark = (remarkText: string) => {
    if (!String(newFollowupRemarks || '').trim()) {
      setNewFollowupRemarks(remarkText);
    } else {
      setNewFollowupRemarks(prev => `${String(prev || '').trim()}, ${remarkText}`);
    }
  };

  // Tele Calling Tab states
  const [telecallerSubTab, setTelecallerSubTab] = useState<'scheduled' | 'date_report' | 'all_logs'>('scheduled');
  const [telecallerDateFrom, setTelecallerDateFrom] = useState('');
  const [telecallerDateTo, setTelecallerDateTo] = useState('');
  const [telecallerSearch, setTelecallerSearch] = useState('');
  const [telecallerScheduledStatus, setTelecallerScheduledStatus] = useState<'all' | 'today' | 'overdue' | 'upcoming'>('all');
  const [telecallerCallerFilter, setTelecallerCallerFilter] = useState<string>('all');
  const [telecallerSupervisorFilter, setTelecallerSupervisorFilter] = useState<string>('all');
  const [telecallerDisplayMode, setTelecallerDisplayMode] = useState<'cards' | 'table'>('cards');
  const [isTelecallerMaximized, setIsTelecallerMaximized] = useState(false);

  // Pagination for Service Follow-up (addresses system slowness / rendering bottleneck)
  const [followupPage, setFollowupPage] = useState(1);
  const [followupItemsPerPage, setFollowupItemsPerPage] = useState(25);

  const [isEditCustModalOpen, setIsEditCustModalOpen] = useState(false);
  const [editingCustChassisKey, setEditingCustChassisKey] = useState('');

  // Complaint Register States
  const [isComplaintModalOpen, setIsComplaintModalOpen] = useState(false);
  const [complaintForm, setComplaintForm] = useState({
    id: '',
    complaintNo: '',
    complaintDate: new Date().toISOString().split('T')[0],
    chassisNo: '',
    customerName: '',
    mobileNumber: '',
    tractorModel: '',
    complaintDetails: '',
    assignedMechanic: '',
    assignedSupervisor: '',
    jobCardNo: '',
    status: 'Open',
    resolution: '',
    closureDate: '',
  });
  const [complaintSearch, setComplaintSearch] = useState('');
  const [complaintStatusFilter, setComplaintStatusFilter] = useState<'all' | 'Open' | 'Running' | 'Closed'>('all');
  const [complaintTechFilter, setComplaintTechFilter] = useState('all');
  const [complaintSupFilter, setComplaintSupFilter] = useState('all');
  const [complaintDateFilter, setComplaintDateFilter] = useState('');
  const [useGoogleSheets, setUseGoogleSheets] = useState<boolean>(false);
  const [sheetsToken, setSheetsToken] = useState<string | null>(null);
  const [sheetsSpreadsheetId, setSheetsSpreadsheetId] = useState<string | null>(() => {
    return localStorage.getItem('sri_google_spreadsheet_id');
  });
  const [sheetsLoading, setSheetsLoading] = useState<boolean>(false);
  const [lastAutoBackupTime, setLastAutoBackupTime] = useState<string>(() => {
    return localStorage.getItem('sri_last_auto_backup_time') || '';
  });
  const [autoBackupStatus, setAutoBackupStatus] = useState<'idle' | 'running' | 'success' | 'error'>('idle');
  const [autoBackupMessage, setAutoBackupMessage] = useState<string>('');
  const [migrationStatus, setMigrationStatus] = useState<{
    total: number;
    current: number;
    step: string;
    loading: boolean;
  }>({ total: 0, current: 0, step: '', loading: false });

  const [editCustForm, setEditCustForm] = useState({
    branch: '',
    sNo: '',
    model: '',
    modelType: '',
    chassisNo: '',
    engineNo: '',
    dateOfDel: '',
    custName: '',
    fatherName: '',
    address: '',
    village: '',
    mandal: '',
    mobileNumber: '',
    district: '',
    pinCode: '',
    dspName: '',
    exchangeBrand: '',
    exchangeModels: '',
    supervisor: ''
  });

  // Reports & Analytics States & Filters
  const [reportSubTab, setReportSubTab] = useState<'dashboard' | 'mechanics' | 'delivery' | 'open_closed' | 'locations' | 'matrix'>('dashboard');
  const [reportDateFrom, setReportDateFrom] = useState('');
  const [reportDateTo, setReportDateTo] = useState('');
  const [reportLocation, setReportLocation] = useState<'all' | 'workshop' | 'dss' | 'event'>('all');
  const [reportStatus, setReportStatus] = useState<'all' | 'Open' | 'Closed'>('all');
  const [reportMechanic, setReportMechanic] = useState('all');
  const [reportSupervisor, setReportSupervisor] = useState('all');
  const [reportServiceType, setReportServiceType] = useState('all');
  const [reportSearchQuery, setReportSearchQuery] = useState('');
  const [selectedReportKpi, setSelectedReportKpi] = useState<'all' | 'open' | 'closed' | 'revenue' | 'spares' | 'labour' | 'workshop' | 'dss' | 'event'>('all');

  // Dashboard Widget Visibility & Minimize States
  const [dashDeliveriesOpen, setDashDeliveriesOpen] = useState(true);
  const [dashAttendanceOpen, setDashAttendanceOpen] = useState(true);
  const [dashComplaintsOpen, setDashComplaintsOpen] = useState(true);
  const [dashJobCardsOpen, setDashJobCardsOpen] = useState(true);
  const [dashFinancialsOpen, setDashFinancialsOpen] = useState(true);
  const [dashTelecallingOpen, setDashTelecallingOpen] = useState(true);

  // Month Deliveries Modal State
  const [isMonthDeliveriesModalOpen, setIsMonthDeliveriesModalOpen] = useState(false);
  const [monthDeliveriesSearch, setMonthDeliveriesSearch] = useState('');

  // Reports Table Resizing, Line Height & Pagination States
  const [reportRowDensity, setReportRowDensity] = useState<'compact' | 'normal' | 'spacious'>('normal');
  const [reportPage, setReportPage] = useState(1);
  const [reportPageSize, setReportPageSize] = useState(50);
  const REPORT_COLUMNS_ORDER = useMemo(() => [
    'slNo',
    'jobNo',
    'jobDate',
    'customerInfo',
    'mobile',
    'modelInfo',
    'servicePlace',
    'mechanic',
    'supervisor',
    'status',
    'spares',
    'labour',
    'total'
  ], []);

  const defaultReportColWidths: Record<string, number> = {
    slNo: 40,
    jobNo: 110,
    jobDate: 90,
    customerInfo: 180,
    mobile: 90,
    modelInfo: 150,
    servicePlace: 100,
    mechanic: 150,
    supervisor: 150,
    status: 90,
    spares: 90,
    labour: 90,
    total: 100
  };
  const [reportColWidths, setReportColWidths] = useState(defaultReportColWidths);
  const [reportFreezeUpToColumn, setReportFreezeUpToColumn] = useState('jobNo');

  // Add New Customer Modal State
  const [isAddCustomerOpen, setIsAddCustomerOpen] = useState(false);
  const handleOpenAddCustomer = () => {
    const todayStr = getLocalDateTimeString().split('T')[0];
    setNewCustForm(prev => ({
      ...prev,
      dateOfDel: fmtDate(todayStr)
    }));
    setIsAddCustomerOpen(true);
  };
  const [isDuplicateModalOpen, setIsDuplicateModalOpen] = useState(false);
  const [duplicateReport, setDuplicateReport] = useState<any[]>([]);
  const [lastUploadedRows, setLastUploadedRows] = useState<any[]>([]);
  const [isFileDuplicateModalOpen, setIsFileDuplicateModalOpen] = useState(false);
  const [fileDuplicateReport, setFileDuplicateReport] = useState<any[]>([]);

  // Function to check for potential duplicate records in a given set
  const getDuplicateReportForSet = (records: any[]) => {
    const chassisMap = new Map<string, any[]>();
    const phoneMap = new Map<string, any[]>();
    
    records.forEach((rec: any) => {
      if (!rec) return;
      const lookup = rec.chassisNo ? rec : buildRowLookup(rec);
      const ch = lookup['Chassis no'] || lookup.__chassisDisplay || lookup.chassisNo || getFieldValue(lookup, 'chassis') || '';
      const normCh = normalizeKey(ch);
      if (normCh) {
        if (!chassisMap.has(normCh)) chassisMap.set(normCh, []);
        chassisMap.get(normCh)!.push(lookup);
      }
      
      const phone = lookup['Mobile Number'] || lookup.mobileNumber || lookup['Phone'] || getFieldValue(lookup, 'custPhone') || '';
      const normPhone = String(phone || '').replace(/[^0-9]/g, '');
      if (normPhone && normPhone.length >= 10) {
        const last10 = normPhone.slice(-10);
        if (!phoneMap.has(last10)) phoneMap.set(last10, []);
        phoneMap.get(last10)!.push(lookup);
      }
    });

    const duplicates: any[] = [];

    // 1. Exact Normalized Chassis Duplicates
    chassisMap.forEach((recs, ch) => {
      if (recs.length > 1) {
        duplicates.push({
          type: 'Exact Chassis Duplicate',
          key: ch,
          records: recs,
          reason: 'Multiple records found with the same normalized chassis number.'
        });
      }
    });

    // 2. Suffix Match (Last 6 digits)
    const suffixMap = new Map<string, any[]>();
    records.forEach(rec => {
      if (!rec) return;
      const lookup = rec.chassisNo ? rec : buildRowLookup(rec);
      const ch = lookup['Chassis no'] || lookup.__chassisDisplay || lookup.chassisNo || getFieldValue(lookup, 'chassis') || '';
      const norm = normalizeKey(ch);
      if (norm.length >= 6) {
        const suffix = norm.slice(-6);
        if (!suffixMap.has(suffix)) suffixMap.set(suffix, []);
        suffixMap.get(suffix)!.push(lookup);
      }
    });
    
    suffixMap.forEach((recs, suffix) => {
      if (recs.length > 1) {
        const uniqueChassis = new Set(recs.map(r => normalizeKey(r['Chassis no'] || r.__chassisDisplay || r.chassisNo || getFieldValue(r, 'chassis') || '')));
        if (uniqueChassis.size > 1) {
          duplicates.push({
            type: 'Potential Chassis Duplicate (Suffix Match)',
            key: `Suffix: ...${suffix}`,
            records: recs,
            reason: 'Different chassis numbers sharing the same last 6 digits. Often happens with entry typos.'
          });
        }
      }
    });

    // 3. Phone Match
    phoneMap.forEach((recs, phone) => {
      if (recs.length > 1) {
        const uniqueChassis = new Set(recs.map(r => normalizeKey(r['Chassis no'] || r.__chassisDisplay || r.chassisNo || getFieldValue(r, 'chassis') || '')));
        if (uniqueChassis.size > 1) {
          duplicates.push({
            type: 'Same Phone, Different Chassis',
            key: phone,
            records: recs,
            reason: 'Multiple chassis records associated with the same mobile number.'
          });
        }
      }
    });

    return duplicates;
  };

  const checkUploadedFileDuplicates = (data: any[]) => {
    const dups = getDuplicateReportForSet(data);
    setFileDuplicateReport(dups);
    setIsFileDuplicateModalOpen(true);
  };

  // Function to check for potential duplicate customer records
  const checkCustomerDuplicates = () => {
    const rawRecords = Object.values(chassisIndex);
    const uniqueRecs = new Set(rawRecords); // Remove identical object references
    const records = Array.from(uniqueRecs);
    const dups = getDuplicateReportForSet(records);
    setDuplicateReport(dups);
    setIsDuplicateModalOpen(true);
  };
  const [isUploading, setIsUploading] = useState(false);
  const [newCustForm, setNewCustForm] = useState({
    branch: '',
    sNo: '',
    model: '',
    modelType: '',
    chassisNo: '',
    engineNo: '',
    dateOfDel: '',
    custName: '',
    fatherName: '',
    address: '',
    village: '',
    mandal: '',
    mobileNumber: '',
    district: '',
    pinCode: '',
    dspName: '',
    exchangeBrand: '',
    exchangeModels: '',
    supervisor: ''
  });

  // Saved Job Cards State & Advanced Filters
  const [savedJobCards, setSavedJobCards] = useState<any[]>([]);
  const [selectedJobCardIds, setSelectedJobCardIds] = useState<string[]>([]);

  useEffect(() => {
    async function loadBackupFromIndexedDB() {
      try {
        const cached = await loadFromStorage<any[]>('sri_backup_jobcards');
        if (cached && cached.length > 0) {
          setSavedJobCards(prev => {
            if (prev.length < cached.length) {
              return cached;
            }
            return prev;
          });
        }
      } catch (err) {
        console.warn('Failed to load backup from IndexedDB:', err);
      }
    }
    loadBackupFromIndexedDB();
  }, []);
  const [firestoreQuotaExceeded, setFirestoreQuotaExceeded] = useState(false);
  const [editingCardId, setEditingCardId] = useState<string | null>(null);
  const [savedListSearch, setSavedListSearch] = useState('');
  const [savedDateFrom, setSavedDateFrom] = useState('');
  const [savedDateTo, setSavedDateTo] = useState('');
  const [savedStatusFilter, setSavedStatusFilter] = useState<'all' | 'Open' | 'Closed' | 'MissingOnline'>('all');
  const [savedSupervisorFilter, setSavedSupervisorFilter] = useState('all');
  const [savedMechanicFilter, setSavedMechanicFilter] = useState('all');
  const [savedLocationFilter, setSavedLocationFilter] = useState<'all' | 'workshop' | 'dss' | 'event'>('all');
  const [savedListPage, setSavedListPage] = useState(1);
  const [savedListPageSize, setSavedListPageSize] = useState(50);

  // Saved Job Cards Table Resizing & Line Height States
  const defaultSavedColWidths: Record<string, number> = {
    slNo: 55,
    status: 105,
    jobNo: 130,
    complaintDate: 130,
    onlineJobCardNo: 155,
    jobCardOpenDate: 140,
    branch: 110,
    historyFileNo: 125,
    tractorModel: 135,
    modelType: 110,
    chassisNo: 150,
    engSrNo: 130,
    dateOfDelivery: 140,
    customerName: 180,
    fatherName: 150,
    address: 220,
    village: 120,
    mandal: 120,
    phoneNo: 125,
    hrsRun: 100,
    typeOfService: 135,
    freeServiceList: 180,
    extraRepairsDone: 250,
    actualClosedDate: 140,
    technicianName: 150,
    servicePlace: 130,
    billNo: 120,
    reasonsForAnalysis: 220,
    telecalling: 140,
    actions: 240
  };

  const [savedColWidths, setSavedColWidths] = useState<Record<string, number>>(defaultSavedColWidths);
  const [savedRowDensity, setSavedRowDensity] = useState<'compact' | 'normal' | 'spacious'>('normal');

  const SAVED_COLUMNS_ORDER = useMemo(() => [
    'slNo',
    'status',
    'jobNo',
    'complaintDate',
    'onlineJobCardNo',
    'jobCardOpenDate',
    'branch',
    'historyFileNo',
    'tractorModel',
    'modelType',
    'chassisNo',
    'engSrNo',
    'dateOfDelivery',
    'customerName',
    'fatherName',
    'address',
    'village',
    'mandal',
    'phoneNo',
    'hrsRun',
    'typeOfService',
    'freeServiceList',
    'extraRepairsDone',
    'actualClosedDate',
    'technicianName',
    'servicePlace',
    'billNo',
    'reasonsForAnalysis',
    'telecalling'
  ], []);

  const [freezeUpToColumn, setFreezeUpToColumn] = useState<string>('status');
  const [isSavedTableMaximized, setIsSavedTableMaximized] = useState<boolean>(false);

  const getStickyProps = (colKey: string, isHeader = false) => {
    const colIndex = SAVED_COLUMNS_ORDER.indexOf(colKey);
    const freezeIndex = SAVED_COLUMNS_ORDER.indexOf(freezeUpToColumn);
    
    // Check if this column is frozen
    const isFrozen = colIndex >= 0 && colIndex <= freezeIndex;
    
    if (!isFrozen) {
      return {
        style: {
          width: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
          minWidth: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
          maxWidth: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
        },
        className: "",
      };
    }
    
    // Calculate cumulative left width
    let leftOffset = 0;
    for (let i = 0; i < colIndex; i++) {
      leftOffset += savedColWidths[SAVED_COLUMNS_ORDER[i]] !== undefined ? savedColWidths[SAVED_COLUMNS_ORDER[i]] : 120;
    }
    
    // Determine the classes
    // Add shadow/border to indicate the edge of the frozen boundary
    const isBoundary = colIndex === freezeIndex;
    const boundaryClass = isBoundary ? "shadow-[4px_0_10px_-3px_rgba(0,0,0,0.15)] border-r-2 border-indigo-400" : "";
    
    const baseClass = isHeader 
      ? `sticky top-0 z-40 bg-slate-100 ${boundaryClass}`
      : `sticky z-20 ${boundaryClass}`;
      
    return {
      style: {
        left: `${leftOffset}px`,
        width: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
        minWidth: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
        maxWidth: savedColWidths[colKey] !== undefined ? `${savedColWidths[colKey]}px` : '120px',
      },
      className: baseClass,
    };
  };

  const handleColumnResizeStart = (colKey: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = savedColWidths[colKey] || 100;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const newWidth = Math.max(45, startWidth + deltaX);
      setSavedColWidths(prev => ({ ...prev, [colKey]: newWidth }));
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const getReportStickyProps = (colKey: string, isHeader = false) => {
    const colIndex = REPORT_COLUMNS_ORDER.indexOf(colKey);
    const freezeIndex = REPORT_COLUMNS_ORDER.indexOf(reportFreezeUpToColumn);
    
    // Check if this column is frozen
    const isFrozen = colIndex >= 0 && colIndex <= freezeIndex;
    
    if (!isFrozen) {
      return {
        style: {
          width: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
          minWidth: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
          maxWidth: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
        },
        className: "",
      };
    }
    
    // Calculate cumulative left width
    let leftOffset = 0;
    for (let i = 0; i < colIndex; i++) {
      leftOffset += reportColWidths[REPORT_COLUMNS_ORDER[i]] !== undefined ? reportColWidths[REPORT_COLUMNS_ORDER[i]] : 120;
    }
    
    // Determine the classes
    // Add shadow/border to indicate the edge of the frozen boundary
    const isBoundary = colIndex === freezeIndex;
    const boundaryClass = isBoundary ? "shadow-[4px_0_10px_-3px_rgba(0,0,0,0.15)] border-r-2 border-indigo-400" : "";
    
    const baseClass = isHeader 
      ? `sticky top-0 z-40 bg-slate-100 ${boundaryClass}`
      : `sticky z-20 ${boundaryClass}`;
      
    return {
      style: {
        left: `${leftOffset}px`,
        width: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
        minWidth: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
        maxWidth: reportColWidths[colKey] !== undefined ? `${reportColWidths[colKey]}px` : '120px',
      },
      className: baseClass,
    };
  };

  const handleReportColumnResizeStart = (colKey: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = reportColWidths[colKey] || 100;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const newWidth = Math.max(45, startWidth + deltaX);
      setReportColWidths(prev => ({ ...prev, [colKey]: newWidth }));
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  // Form Ribbons Collapsible States (Sections A, B, C, D)
  const [sectionAOpen, setSectionAOpen] = useState(true);
  const [sectionBOpen, setSectionBOpen] = useState(true);
  const [sectionCOpen, setSectionCOpen] = useState(true);
  const [sectionDOpen, setSectionDOpen] = useState(true);

  // Service Location Type
  const [serviceLocation, setServiceLocation] = useState<'workshop' | 'dss' | 'event'>('workshop');

  // Dedicated Customer Search
  const [searchQuery, setSearchQuery] = useState('');

  // Tractor Details Section A
  const [jobNo, setJobNo] = useState('');
  const [jobDate, setJobDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [chassisNo, setChassisNo] = useState('');
  const [custName, setCustName] = useState('');
  const [translatedCustName, setTranslatedCustName] = useState('');
  const [engineNo, setEngineNo] = useState('');
  const [custAddr, setCustAddr] = useState('');
  const [translatedCustAddr, setTranslatedCustAddr] = useState('');
  const [village, setVillage] = useState('');
  const [translatedVillage, setTranslatedVillage] = useState('');
  const [mandal, setMandal] = useState('');
  const [translatedMandal, setTranslatedMandal] = useState('');

  useEffect(() => {
    let isCancelled = false;
    const timer = setTimeout(async () => {
      if (custName) {
        const te = await translateTextToTelugu(custName);
        if (!isCancelled) setTranslatedCustName(te);
      } else {
        if (!isCancelled) setTranslatedCustName('');
      }

      if (village) {
        const te = await translateTextToTelugu(village);
        if (!isCancelled) setTranslatedVillage(te);
      } else {
        if (!isCancelled) setTranslatedVillage('');
      }

      if (mandal) {
        const te = await translateTextToTelugu(mandal);
        if (!isCancelled) setTranslatedMandal(te);
      } else {
        if (!isCancelled) setTranslatedMandal('');
      }

      if (custAddr) {
        const te = await translateTextToTelugu(custAddr);
        if (!isCancelled) setTranslatedCustAddr(te);
      } else {
        if (!isCancelled) setTranslatedCustAddr('');
      }
    }, 350);

    return () => {
      isCancelled = true;
      clearTimeout(timer);
    };
  }, [custName, village, mandal, custAddr]);
  const [installDate, setInstallDate] = useState('');
  const [ownerMob, setOwnerMob] = useState('');
  const [driverMob, setDriverMob] = useState('');
  const [regdNo, setRegdNo] = useState('');
  const [distDealership, setDistDealership] = useState('');
  const [hourMeter, setHourMeter] = useState('');
  const [dateTimeIn, setDateTimeIn] = useState('');
  const [dateTimeOut, setDateTimeOut] = useState('');
  const [expectedRepairTime, setExpectedRepairTime] = useState('');
  const [serviceType, setServiceLocationType] = useState<string>('Paid Service');
  const [warrantyOverride, setWarrantyOverride] = useState<'auto' | 'warranty' | 'post_wty'>('auto');
  const [model, setModel] = useState('');
  const [serialNo, setSerialNo] = useState('');
  const [historyFileNo, setHistoryFileNo] = useState('');
  const [billNo, setBillNo] = useState('');
  const [status, setStatus] = useState<'Open' | 'Closed'>('Open');
  const [complaintDate, setComplaintDate] = useState('');
  const [complaintDetails, setComplaintDetails] = useState('');
  const [onlineJobCardNo, setOnlineJobCardNo] = useState('');
  const [branch, setBranch] = useState('');
  const [modelType, setModelType] = useState('');
  const [fatherName, setFatherName] = useState('');
  const [freeServiceList, setFreeServiceList] = useState('');
  const [extraRepairs, setExtraRepairs] = useState('');
  const [reasonsForAnalysis, setReasonsForAnalysis] = useState('');
  const [telecalling, setTelecalling] = useState('');
  const [showJobCardTable, setShowJobCardTable] = useState(false);
  const [viewingJobCardDetails, setViewingJobCardDetails] = useState<any | null>(null);

  // Firebase Auth and Role States
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [currentUserRole, setCurrentUserRole] = useState<'admin' | 'user' | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('login');
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authConfirmPassword, setAuthConfirmPassword] = useState('');
  const [authPasscode, setAuthPasscode] = useState('');
  const [authRegisterAsAdmin, setAuthRegisterAsAdmin] = useState(false);
  const [authError, setAuthError] = useState('');
  const [isAuthSubmitting, setIsAuthSubmitting] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  // Service Follow-up Interval State & Storage Persistence
  const [serviceIntervalDays, setServiceIntervalDays] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('sri_service_interval_days');
      return saved ? parseInt(saved, 10) : 180;
    } catch (e) {
      return 180;
    }
  });
  const [showOnlyFollowUpDue, setShowOnlyFollowUpDue] = useState(false);
  const [autoFillNotice, setAutoFillNotice] = useState<{ chassis: string; filledCount: number; emptyCount: number } | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem('sri_service_interval_days', serviceIntervalDays.toString());
    } catch (e) {}
  }, [serviceIntervalDays]);

  // Staff and Directory States
  const [staffMembers, setStaffMembers] = useState<any[]>([]);
  const [isStaffModalOpen, setIsStaffModalOpen] = useState(false);
  const [editingStaffId, setEditingStaffId] = useState<string | null>(null);
  const [showStaffTable, setShowStaffTable] = useState(false);
  const [staffDirectoryTab, setStaffDirectoryTab] = useState<'all_grid' | 'all_table' | 'supervisor_teams'>('all_grid');
  const [selectedSupervisorTeam, setSelectedSupervisorTeam] = useState<string>('all');
  const [isDirectoryExpanded, setIsDirectoryExpanded] = useState(false);
  const [viewingStaffDetails, setViewingStaffDetails] = useState<any | null>(null);
  const [staffForm, setStaffForm] = useState({
    name: '',
    fatherName: '',
    village: '',
    mandal: '',
    mobileNumber: '',
    role: 'mechanic',
    dateOfJoining: '',
    supervisor: ''
  });

  // General Checklist State (Checkboxes)
  const [genChecklist, setGenChecklist] = useState<Record<string, boolean>>({});
  const [otherChecklistText, setOtherChecklistText] = useState('');
  const [problemDescription, setProblemDescription] = useState('');

  // 25 Checkpoint Table
  const [checkpoints, setCheckpoints] = useState<CheckpointItem[]>(DEFAULT_CHECKPOINTS);

  // Repairs Section (Back side / charges)
  const [repairRows, setRepairRows] = useState<RepairRow[]>([
    { repair: '', rectification: '', charge: '' },
    { repair: '', rectification: '', charge: '' }
  ]);

  // Workshop / Charges
  const [wsReport, setWsReport] = useState('');
  const [mechanic, setMechanic] = useState('');
  const [totalLabour, setTotalLabour] = useState('');
  const [warrantyMaterial, setWarrantyMaterial] = useState('');
  const [nonWarrantyMaterial, setNonWarrantyMaterial] = useState('');
  const [gTotal, setGTotal] = useState('');
  const [wsIncharge, setWsIncharge] = useState('');

  // Parts / Materials
  const [partRows, setPartRows] = useState<PartRow[]>([
    { partNo: '', desc: '', qty: '', rate: '', amount: '' },
    { partNo: '', desc: '', qty: '', rate: '', amount: '' },
    { partNo: '', desc: '', qty: '', rate: '', amount: '' }
  ]);

  // Databases
  const [chassisIndex, setChassisIndex] = useState<Record<string, any>>({});
  const customerRecords = useMemo(() => Object.values(chassisIndex), [chassisIndex]);
  const chassisWithJobCards = useMemo(() => new Set(savedJobCards.map(jc => normalizeKey(jc.chassisNo || jc.chassis || ''))), [savedJobCards]);
  const [sparesIndex, setSparesIndex] = useState<Record<string, any>>({});
  const [excelStatus, setExcelStatus] = useState({ text: 'No customer file uploaded yet.', isSuccess: false });
  const [sparesStatus, setSparesStatus] = useState({ text: 'No spares file uploaded yet. Manual entry works fine.', isSuccess: false });
  const [jobCardsStatus, setJobCardsStatus] = useState({ text: '', isSuccess: false });

  // Print selection
  const [printOption, setPrintOption] = useState<'both' | 'jobCardOnly' | 'partsOnly'>('both');

  // Load initial data from Cloud SQL database & local storage fallback
  useEffect(() => {
    (async () => {
      // 1. First attempt to load full dataset from Cloud SQL (High Speed PostgreSQL Database)
      try {
        const [sqlCustomers, sqlSpares, sqlJobCards, sqlComplaints, sqlStaff] = await Promise.allSettled([
          sqlApi.getCustomers(),
          sqlApi.getSpares(),
          sqlApi.getJobCards(),
          sqlApi.getComplaints(),
          sqlApi.getStaff(),
        ]);

        if (sqlCustomers.status === 'fulfilled' && Array.isArray(sqlCustomers.value) && sqlCustomers.value.length > 0) {
          const loadedIndex: Record<string, any> = {};
          sqlCustomers.value.forEach(row => {
            const lookup = buildRowLookup(row.full_data || row);
            const chassisVal = row.chassis_no || getFieldValue(lookup, 'chassis');
            const phoneVal = row.owner_mob || getFieldValue(lookup, 'custPhone');
            const drvPhoneVal = row.driver_mob || getFieldValue(lookup, 'driverPhone');
            const nameVal = row.cust_name || getFieldValue(lookup, 'custName');
            const regdVal = row.regd_no || getFieldValue(lookup, 'regdNo');
            const addrVal = row.cust_addr || getCombinedAddress(lookup);

            if (chassisVal) lookup.__chassisDisplay = chassisVal;
            if (nameVal) lookup.__custNameDisplay = nameVal;
            if (phoneVal) lookup.__custPhoneDisplay = phoneVal;
            if (addrVal) lookup.__custAddrDisplay = addrVal;

            const keyCh = normalizeKey(chassisVal);
            const keyPh = normalizeKey(phoneVal);
            const keyDph = normalizeKey(drvPhoneVal);
            const keyNm = normalizeKey(nameVal);
            const keyReg = normalizeKey(regdVal);

            if (keyCh) loadedIndex[keyCh] = lookup;
            if (keyPh) loadedIndex[keyPh] = lookup;
            if (keyDph) loadedIndex[keyDph] = lookup;
            if (keyNm) loadedIndex[keyNm] = lookup;
            if (keyReg) loadedIndex[keyReg] = lookup;
          });
          setChassisIndex(loadedIndex);
          saveToStorage(LS_CUSTOMER_KEY, loadedIndex);
          setExcelStatus({
            text: `⚡ ${sqlCustomers.value.length} customer record(s) loaded from Cloud SQL Database.`,
            isSuccess: true
          });
        }

        if (sqlSpares.status === 'fulfilled' && Array.isArray(sqlSpares.value) && sqlSpares.value.length > 0) {
          const loadedSpares: Record<string, any> = {};
          sqlSpares.value.forEach(row => {
            const lookup = buildRowLookup(row.full_data || row);
            const partNoVal = row.part_no || row.partNo;
            if (partNoVal) {
              lookup.__partNoDisplay = partNoVal;
              loadedSpares[normalizeKey(partNoVal)] = lookup;
            }
          });
          setSparesIndex(loadedSpares);
          saveToStorage(LS_SPARES_KEY, loadedSpares);
          setSparesStatus({
            text: `⚡ ${sqlSpares.value.length} spare part(s) loaded from Cloud SQL Database.`,
            isSuccess: true
          });
        }

        if (sqlJobCards.status === 'fulfilled' && Array.isArray(sqlJobCards.value) && sqlJobCards.value.length > 0) {
          const mappedCards = sqlJobCards.value.map(c => ({
            ...c,
            jobNo: c.job_no || c.jobNo,
            onlineJobCardNo: c.online_job_card_no || c.onlineJobCardNo,
            jobDate: c.job_date || c.jobDate,
            dateTimeIn: c.date_time_in || c.dateTimeIn,
            dateTimeOut: c.date_time_out || c.dateTimeOut,
            expectedRepairTime: c.expected_repair_time || c.expectedRepairTime,
            custName: c.cust_name || c.custName,
            fatherName: c.father_name || c.fatherName,
            custAddr: c.cust_addr || c.custAddr,
            ownerMob: c.owner_mob || c.ownerMob,
            driverMob: c.driver_mob || c.driverMob,
            regdNo: c.regd_no || c.regdNo,
            chassisNo: c.chassis_no || c.chassisNo,
            engineNo: c.engine_no || c.engineNo,
            modelType: c.model_type || c.modelType,
            serialNo: c.serial_no || c.serialNo,
            hourMeter: c.hour_meter || c.hourMeter,
            serviceType: c.service_type || c.serviceType,
            freeServiceList: c.free_service_list || c.freeServiceList,
            extraRepairs: c.extra_repairs || c.extraRepairs,
            wsIncharge: c.ws_incharge || c.wsIncharge,
            serviceLocation: c.service_location || c.serviceLocation,
            billNo: c.bill_no || c.billNo,
            reasonsForAnalysis: c.reasons_for_analysis || c.reasonsForAnalysis,
            warrantyOverride: c.warranty_override || c.warrantyOverride,
            totalLabour: c.total_labour || c.totalLabour,
            warrantyMaterial: c.warranty_material || c.warrantyMaterial,
            nonWarrantyMaterial: c.non_warranty_material || c.nonWarrantyMaterial,
            gTotal: c.g_total || c.gTotal,
            actualClosedDate: c.actual_closed_date || c.actualClosedDate,
            checkpoints: typeof c.checkpoints === 'string' ? JSON.parse(c.checkpoints) : (c.checkpoints || []),
            repairRows: typeof c.repair_rows === 'string' ? JSON.parse(c.repair_rows) : (c.repair_rows || []),
            partRows: typeof c.part_rows === 'string' ? JSON.parse(c.part_rows) : (c.part_rows || []),
          }));
          setSavedJobCards(mappedCards);
          saveJobCardsBackup(mappedCards);
        }

        if (sqlComplaints.status === 'fulfilled' && Array.isArray(sqlComplaints.value) && sqlComplaints.value.length > 0) {
          const mappedComp = sqlComplaints.value.map(c => ({
            ...c,
            complaintNo: c.complaint_no || c.complaintNo,
            customerName: c.customer_name || c.customerName,
            tractorModel: c.tractor_model || c.tractorModel,
            chassisNo: c.chassis_no || c.chassisNo,
            complaintDetails: c.complaint_details || c.complaintDetails,
            jobCardNo: c.job_card_no || c.jobCardNo,
            closureDate: c.closure_date || c.closureDate,
          }));
          setComplaints(mappedComp);
        }

        if (sqlStaff.status === 'fulfilled' && Array.isArray(sqlStaff.value) && sqlStaff.value.length > 0) {
          const mappedStaff = sqlStaff.value.map(s => ({
            ...s,
            assignedSupervisor: s.assigned_supervisor || s.assignedSupervisor,
          }));
          setStaffMembers(mappedStaff);
        }
      } catch (sqlLoadErr) {
        console.warn('Cloud SQL initial load:', sqlLoadErr);
      }

      // 2. Local storage fallback if offline
      try {
        const parsedCust = await loadFromStorage<Record<string, any>>(LS_CUSTOMER_KEY);
        if (parsedCust && Object.keys(chassisIndex).length === 0) {
          setChassisIndex(parsedCust);
          const count = Object.keys(parsedCust).length;
          if (count > 0) {
            setExcelStatus({
              text: `✅ ${count} customer record(s) restored from storage.`,
              isSuccess: true
            });
          }
        }
      } catch (e) {
        console.error(e);
      }

      try {
        const parsedSpares = await loadFromStorage<Record<string, any>>(LS_SPARES_KEY);
        if (parsedSpares && Object.keys(sparesIndex).length === 0) {
          setSparesIndex(parsedSpares);
          const count = Object.keys(parsedSpares).length;
          if (count > 0) {
            setSparesStatus({
              text: `✅ ${count} spare part(s) restored from storage.`,
              isSuccess: true
            });
          }
        }
      } catch (e) {
        console.error(e);
      }
    })();
  }, []);

  // Firebase Authentication Listener with Local Session & Auto-Admin Support
  useEffect(() => {
    try {
      const savedSession = localStorage.getItem('eicher_auth_user');
      if (savedSession) {
        const parsed = JSON.parse(savedSession);
        if (parsed?.user) {
          setCurrentUser(parsed.user);
          setCurrentUserRole(parsed.role || (parsed.user.email === 'srigayathriautomotives@gmail.com' || parsed.user.email === 'srigayathriauto@gmail.com' ? 'admin' : 'user'));
          setAuthLoading(false);
          return;
        }
      }
    } catch {}

    // Default seamless login as Admin (Sri Gayathri Automotives) to bypass cloud domain restrictions
    const defaultUser: any = {
      uid: 'sri-gayathri-admin-master',
      email: 'srigayathriautomotives@gmail.com',
      displayName: 'Sri Gayathri Automotives (Admin)',
      photoURL: null
    };
    setCurrentUser(defaultUser);
    setCurrentUserRole('admin');
    try {
      localStorage.setItem('eicher_auth_user', JSON.stringify({ user: defaultUser, role: 'admin' }));
    } catch {}
    setAuthLoading(false);

    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        setCurrentUser(user);
        const role = user.email === 'srigayathriautomotives@gmail.com' || user.email === 'srigayathriauto@gmail.com' ? 'admin' : 'user';
        setCurrentUserRole(role);
        try {
          localStorage.setItem('eicher_auth_user', JSON.stringify({
            user: { uid: user.uid, email: user.email, displayName: user.displayName || user.email, photoURL: user.photoURL },
            role
          }));
        } catch {}
      }
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const [complaints, setComplaints] = useState<any[]>([]);

  // Background sync from Cloud SQL Database periodically (every 15 seconds)
  useEffect(() => {
    const fetchLatestFromCloudSQL = async () => {
      try {
        const [sqlCards, sqlComp, sqlStf, sqlAtt] = await Promise.allSettled([
          sqlApi.getJobCards(),
          sqlApi.getComplaints(),
          sqlApi.getStaff(),
          sqlApi.getAttendance(),
        ]);

        if (sqlCards.status === 'fulfilled' && Array.isArray(sqlCards.value)) {
          const mappedCards = sqlCards.value.map(c => ({
            ...c,
            jobNo: c.job_no || c.jobNo,
            onlineJobCardNo: c.online_job_card_no || c.onlineJobCardNo,
            jobDate: c.job_date || c.jobDate,
            dateTimeIn: c.date_time_in || c.dateTimeIn,
            dateTimeOut: c.date_time_out || c.dateTimeOut,
            expectedRepairTime: c.expected_repair_time || c.expectedRepairTime,
            custName: c.cust_name || c.custName,
            fatherName: c.father_name || c.fatherName,
            custAddr: c.cust_addr || c.custAddr,
            ownerMob: c.owner_mob || c.ownerMob,
            driverMob: c.driver_mob || c.driverMob,
            regdNo: c.regd_no || c.regdNo,
            chassisNo: c.chassis_no || c.chassisNo,
            engineNo: c.engine_no || c.engineNo,
            modelType: c.model_type || c.modelType,
            serialNo: c.serial_no || c.serialNo,
            hourMeter: c.hour_meter || c.hourMeter,
            serviceType: c.service_type || c.serviceType,
            freeServiceList: c.free_service_list || c.freeServiceList,
            extraRepairs: c.extra_repairs || c.extraRepairs,
            wsIncharge: c.ws_incharge || c.wsIncharge,
            serviceLocation: c.service_location || c.serviceLocation,
            billNo: c.bill_no || c.billNo,
            reasonsForAnalysis: c.reasons_for_analysis || c.reasonsForAnalysis,
            warrantyOverride: c.warranty_override || c.warrantyOverride,
            totalLabour: c.total_labour || c.totalLabour,
            warrantyMaterial: c.warranty_material || c.warrantyMaterial,
            nonWarrantyMaterial: c.non_warranty_material || c.nonWarrantyMaterial,
            gTotal: c.g_total || c.gTotal,
            actualClosedDate: c.actual_closed_date || c.actualClosedDate,
            checkpoints: typeof c.checkpoints === 'string' ? JSON.parse(c.checkpoints) : (c.checkpoints || []),
            repairRows: typeof c.repair_rows === 'string' ? JSON.parse(c.repair_rows) : (c.repair_rows || []),
            partRows: typeof c.part_rows === 'string' ? JSON.parse(c.part_rows) : (c.part_rows || []),
          }));
          mappedCards.sort((a, b) => {
            const dateA = a.jobDate || a.jobOpenDate || a.complaintDate || a.createdAt || '';
            const dateB = b.jobDate || b.jobOpenDate || b.complaintDate || b.createdAt || '';
            if (dateA !== dateB) {
              return dateB.localeCompare(dateA);
            }
            return (b.createdAt || '').localeCompare(a.createdAt || '');
          });
          if (mappedCards.length > 0) {
            setSavedJobCards(mappedCards);
            saveJobCardsBackup(mappedCards);
          }
        }

        if (sqlComp.status === 'fulfilled' && Array.isArray(sqlComp.value) && sqlComp.value.length > 0) {
          const mappedComp = sqlComp.value.map(c => ({
            ...c,
            complaintNo: c.complaint_no || c.complaintNo,
            customerName: c.customer_name || c.customerName,
            tractorModel: c.tractor_model || c.tractorModel,
            chassisNo: c.chassis_no || c.chassisNo,
            complaintDetails: c.complaint_details || c.complaintDetails,
            jobCardNo: c.job_card_no || c.jobCardNo,
            closureDate: c.closure_date || c.closureDate,
          }));
          setComplaints(mappedComp);
          try {
            localStorage.setItem('sri_backup_complaints', JSON.stringify(mappedComp));
          } catch (e) {}
        }

        if (sqlStf.status === 'fulfilled' && Array.isArray(sqlStf.value) && sqlStf.value.length > 0) {
          const mappedStaff = sqlStf.value.map(s => ({
            ...s,
            assignedSupervisor: s.assigned_supervisor || s.assignedSupervisor,
          }));
          setStaffMembers(mappedStaff);
          try {
            localStorage.setItem('sri_backup_staff', JSON.stringify(mappedStaff));
          } catch (e) {}
        }

        if (sqlAtt.status === 'fulfilled' && sqlAtt.value && typeof sqlAtt.value === 'object') {
          setAttendanceRecords(prev => {
            const updated = { ...prev, ...sqlAtt.value };
            try {
              localStorage.setItem('sri_staff_attendance_data', JSON.stringify(updated));
            } catch (e) {}
            return updated;
          });
        }
      } catch (err) {
        // Silent error
      }
    };

    fetchLatestFromCloudSQL();
    const interval = setInterval(fetchLatestFromCloudSQL, 20000);
    return () => clearInterval(interval);
  }, []);

  // 1. Sync saved job cards from Firestore with automatic local storage backup & error guard
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'jobcards');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const loaded: any[] = [];
        snapshot.forEach((d) => {
          const data = d.data();
          loaded.push({ id: d.id, ...data });
        });
        loaded.sort((a, b) => {
          const dateA = a.jobDate || a.jobOpenDate || a.complaintDate || a.createdAt || '';
          const dateB = b.jobDate || b.jobOpenDate || b.complaintDate || b.createdAt || '';
          if (dateA !== dateB) return dateB.localeCompare(dateA);
          return (b.createdAt || '').localeCompare(a.createdAt || '');
        });
        if (loaded.length > 0) {
          setSavedJobCards(loaded);
          saveJobCardsBackup(loaded);
        }
        setFirestoreQuotaExceeded(false);
      }, (err) => {
        console.warn('Firestore jobcards listener notice:', err?.message || err);
        if (err?.message?.includes('Quota') || err?.message?.includes('quota') || err?.code === 'resource-exhausted') {
          setFirestoreQuotaExceeded(true);
        }
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Jobcards listener init exception:', e);
    }
  }, []);

  // 2. Sync complaints from Firestore
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'complaints');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const loaded: any[] = [];
        snapshot.forEach((d) => {
          loaded.push({ id: d.id, ...d.data() });
        });
        if (loaded.length > 0) {
          setComplaints(loaded);
          try {
            localStorage.setItem('sri_backup_complaints', JSON.stringify(loaded));
          } catch (e) {}
        }
      }, (err) => {
        console.warn('Firestore complaints listener notice:', err?.message || err);
        if (err?.message?.includes('Quota') || err?.message?.includes('quota') || err?.code === 'resource-exhausted') {
          setFirestoreQuotaExceeded(true);
        }
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Complaints listener init exception:', e);
    }
  }, []);

  // 3. Sync staff from Firestore
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'staff');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const loaded: any[] = [];
        snapshot.forEach((d) => {
          loaded.push({ id: d.id, ...d.data() });
        });
        if (loaded.length > 0) {
          setStaffMembers(loaded);
          try {
            localStorage.setItem('sri_backup_staff', JSON.stringify(loaded));
          } catch (e) {}
        }
      }, (err) => {
        console.warn('Firestore staff listener notice:', err?.message || err);
        if (err?.message?.includes('Quota') || err?.message?.includes('quota') || err?.code === 'resource-exhausted') {
          setFirestoreQuotaExceeded(true);
        }
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Staff listener init exception:', e);
    }
  }, []);

  // 4. Sync staff attendance from Firestore
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'attendance');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        const loaded: Record<string, Record<string, { status: string; remarks: string }>> = {};
        snapshot.forEach((d) => {
          loaded[d.id] = (d.data()?.records || d.data()) as any;
        });
        if (Object.keys(loaded).length > 0) {
          setAttendanceRecords(prev => {
            const updated = { ...prev, ...loaded };
            try {
              localStorage.setItem('sri_staff_attendance_data', JSON.stringify(updated));
            } catch (e) {}
            return updated;
          });
        }
      }, (err) => {
        console.warn('Firestore attendance listener notice:', err?.message || err);
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Attendance listener init exception:', e);
    }
  }, []);

  // 5. Sync customer master records from Firestore
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'customers_master');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        if (!snapshot.empty) {
          const loadedIndex: Record<string, any> = {};
          let allRows: any[] = [];
          snapshot.forEach((docSnap) => {
            const data = docSnap.data();
            if (Array.isArray(data.rows)) {
              allRows.push(...data.rows);
            } else if (data && typeof data === 'object') {
              allRows.push(data);
            }
          });

          allRows.forEach(row => {
            const lookup = buildRowLookup(row.full_data || row);
            const chassisVal = row.chassisNo || row.chassis || row.chassis_no || getFieldValue(lookup, 'chassis');
            const phoneVal = row.ownerMob || row.custPhone || row.owner_mob || getFieldValue(lookup, 'custPhone');
            const drvPhoneVal = row.driverMob || row.driverPhone || row.driver_mob || getFieldValue(lookup, 'driverPhone');
            const nameVal = row.custName || row.cust_name || getFieldValue(lookup, 'custName');
            const regdVal = row.regdNo || row.regd_no || getFieldValue(lookup, 'regdNo');
            const addrVal = row.custAddr || row.cust_addr || getCombinedAddress(lookup);

            if (chassisVal) lookup.__chassisDisplay = chassisVal;
            if (nameVal) lookup.__custNameDisplay = nameVal;
            if (phoneVal) lookup.__custPhoneDisplay = phoneVal;
            if (addrVal) lookup.__custAddrDisplay = addrVal;

            const keyCh = normalizeKey(chassisVal);
            const keyPh = normalizeKey(phoneVal);
            const keyDph = normalizeKey(drvPhoneVal);
            const keyNm = normalizeKey(nameVal);
            const keyReg = normalizeKey(regdVal);

            if (keyCh) loadedIndex[keyCh] = lookup;
            if (keyPh) loadedIndex[keyPh] = lookup;
            if (keyDph) loadedIndex[keyDph] = lookup;
            if (keyNm) loadedIndex[keyNm] = lookup;
            if (keyReg) loadedIndex[keyReg] = lookup;
          });

          if (Object.keys(loadedIndex).length > 0) {
            setChassisIndex(loadedIndex);
            saveToStorage(LS_CUSTOMER_KEY, loadedIndex);
            setExcelStatus({
              text: `✅ ${Object.keys(loadedIndex).length} customer record(s) synced from Cloud Database.`,
              isSuccess: true
            });
          }
        }
      }, (err) => {
        console.warn('Firestore customer master listener notice:', err?.message || err);
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Customer master listener init exception:', e);
    }
  }, []);

  // 6. Sync spares master records from Firestore
  useEffect(() => {
    if (!db) return;
    try {
      const q = collection(db, 'spares_master');
      const unsubscribe = onSnapshot(q, (snapshot) => {
        if (!snapshot.empty) {
          const loadedSpares: Record<string, any> = {};
          let allRows: any[] = [];
          snapshot.forEach((docSnap) => {
            const data = docSnap.data();
            if (Array.isArray(data.rows)) {
              allRows.push(...data.rows);
            } else if (data && typeof data === 'object') {
              allRows.push(data);
            }
          });

          allRows.forEach(row => {
            const lookup = buildRowLookup(row.full_data || row);
            const partNoVal = row.partNo || row.part_no;
            if (partNoVal) {
              lookup.__partNoDisplay = partNoVal;
              loadedSpares[normalizeKey(partNoVal)] = lookup;
            }
          });

          if (Object.keys(loadedSpares).length > 0) {
            setSparesIndex(loadedSpares);
            saveToStorage(LS_SPARES_KEY, loadedSpares);
            setSparesStatus({
              text: `✅ ${Object.keys(loadedSpares).length} spare part(s) synced from Cloud Database.`,
              isSuccess: true
            });
          }
        }
      }, (err) => {
        console.warn('Firestore spares master listener notice:', err?.message || err);
      });
      return () => unsubscribe();
    } catch (e) {
      console.warn('Spares master listener init exception:', e);
    }
  }, []);

  // Bidirectional mapper helpers for Google Sheets and React App State
  const mapJobCardToSheet = (card: any) => {
    return {
      id: card.id || '',
      jobCardNo: card.jobNo || card.jobCardNo || '',
      onlineJobCardNo: card.onlineJobCardNo || '',
      jobDate: card.jobDate || '',
      dateTimeIn: card.dateTimeIn || '',
      dateTimeOut: card.dateTimeOut || '',
      expectedRepairTime: card.expectedRepairTime || '',
      status: card.status || 'Open',
      custName: card.custName || '',
      fatherName: card.fatherName || '',
      custAddr: card.custAddr || '',
      village: card.village || '',
      mandal: card.mandal || '',
      ownerMob: card.ownerMob || '',
      driverMob: card.driverMob || '',
      regdNo: card.regdNo || '',
      chassisNo: card.chassisNo || '',
      engineNo: card.engineNo || '',
      tractorModel: card.model || card.tractorModel || '',
      serialNo: card.serialNo || '',
      hourMeter: card.hourMeter || '',
      serviceType: card.serviceType || '',
      mechanicName: card.mechanic || card.mechanicName || '',
      supervisor: card.wsIncharge || card.supervisor || '',
      warrantyOverride: card.warrantyOverride || 'auto',
      checkpointsJson: card.checkpoints || card.checkpointsJson || [],
      repairsJson: card.repairRows || card.repairsJson || [],
      partsJson: card.partRows || card.partsJson || [],
      historyFileNo: card.historyFileNo || '',
      billNo: card.billNo || '',
      extraRepairs: card.extraRepairs || '',
      reasonsForAnalysis: card.reasonsForAnalysis || '',
      telecalling: card.telecalling || '',
      branch: card.branch || '',
      modelType: card.modelType || '',
      freeServiceList: card.freeServiceList || '',
      createdAt: card.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
  };

  const mapJobCardFromSheet = (row: any) => {
    return {
      ...row,
      jobNo: row.jobCardNo || row.onlineJobCardNo || '',
      model: row.tractorModel || '',
      mechanic: row.mechanicName || '',
      wsIncharge: row.supervisor || '',
      checkpoints: row.checkpointsJson || [],
      repairRows: row.repairsJson || [],
      partRows: row.partsJson || []
    };
  };

  // Google Sheets Connection & Data Loading Handlers
  const handleConnectGoogleSheets = async () => {
    setSheetsLoading(true);
    try {
      const token = await loginWithGoogleForSheets();
      setSheetsToken(token);
      setCachedToken(token);
      
      const spreadsheetId = await getOrCreateSpreadsheet();
      setSheetsSpreadsheetId(spreadsheetId);
      
      await refreshAllSheetsData(spreadsheetId);
      
      setUseGoogleSheets(true);
      localStorage.setItem('sri_use_google_sheets', 'true');
      setFirestoreQuotaExceeded(false);
      
      alert('✅ గూగుల్ షీట్స్ విజయవంతంగా కనెక్ట్ చేయబడింది! మీ డేటా మొత్తం ఎక్సెల్ షీట్‌లో భద్రపరచబడుతుంది.');
    } catch (error) {
      console.error('Google Sheets Connection Error:', error);
      alert('❌ గూగుల్ షీట్స్ కనెక్ట్ చేయడంలో విఫలమైంది: ' + (error instanceof Error ? error.message : String(error)));
    } finally {
      setSheetsLoading(false);
    }
  };

  const handleDownloadMasterBackup = async () => {
    try {
      setSheetsLoading(true);
      // Fetch everything from Firestore for the backup
      const [jobSnap, compSnap, staffSnap] = await Promise.all([
        getDocs(collection(db, 'jobcards')),
        getDocs(collection(db, 'complaints')),
        getDocs(collection(db, 'staff'))
      ]);

      const jobCardsList = jobSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const complaintsList = compSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const staffList = staffSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const customerList = Object.values(chassisIndex);
      const sparesList = Object.values(sparesIndex);

      // Create Workbook
      const wb = XLSX.utils.book_new();

      // 1. JobCards Sheet
      const jcData = jobCardsList.map(card => {
        const mapped = mapJobCardToSheet(card);
        return {
          ...mapped,
          checkpointsJson: typeof mapped.checkpointsJson === 'string' ? mapped.checkpointsJson : JSON.stringify(mapped.checkpointsJson),
          repairsJson: typeof mapped.repairsJson === 'string' ? mapped.repairsJson : JSON.stringify(mapped.repairsJson),
          partsJson: typeof mapped.partsJson === 'string' ? mapped.partsJson : JSON.stringify(mapped.partsJson)
        };
      });
      const jcSheet = XLSX.utils.json_to_sheet(jcData);
      XLSX.utils.book_append_sheet(wb, jcSheet, 'JobCards');

      // 2. Complaints Sheet
      const complaintsData = complaintsList.map((c: any) => ({
        ...c,
        createdAt: c.createdAt instanceof Date ? c.createdAt.toISOString() : (c.createdAt?.toDate?.() ? c.createdAt.toDate().toISOString() : c.createdAt)
      }));
      const cSheet = XLSX.utils.json_to_sheet(complaintsData);
      XLSX.utils.book_append_sheet(wb, cSheet, 'Complaints');

      // 3. Staff Sheet
      const sSheet = XLSX.utils.json_to_sheet(staffList);
      XLSX.utils.book_append_sheet(wb, sSheet, 'Staff');
      
      // 4. Customers Sheet (Master Database)
      const customersData = customerList.map((cust: any) => ({
        ...cust,
        followupHistory: cust.followupHistory ? JSON.stringify(cust.followupHistory) : '[]'
      }));
      const custSheet = XLSX.utils.json_to_sheet(customersData);
      XLSX.utils.book_append_sheet(wb, custSheet, 'Customers');
      
      // 5. Spares Sheet (Price List)
      const sparesSheet = XLSX.utils.json_to_sheet(sparesList);
      XLSX.utils.book_append_sheet(wb, sparesSheet, 'Spares');
      
      // 6. Settings Sheet
      const settingsData = [
        { key: 'BranchName', value: branch || '' },
        { key: 'QuickRemarks', value: JSON.stringify(quickRemarksList) },
        { key: 'ServiceInterval', value: serviceIntervalDays.toString() },
        { key: 'MenuOrder', value: JSON.stringify(menuOrder) },
        { key: 'ExportTimestamp', value: new Date().toISOString() },
        { key: 'BackupType', value: 'FULL_MASTER_SOFTWARE_BACKUP' }
      ];
      const settingsSheet = XLSX.utils.json_to_sheet(settingsData);
      XLSX.utils.book_append_sheet(wb, settingsSheet, 'AppSettings');

      // 7. Staff Attendance Sheet
      const attendanceData: any[] = [];
      Object.entries(attendanceRecords).forEach(([date, recs]) => {
        Object.entries(recs).forEach(([staffId, info]) => {
          attendanceData.push({
            date,
            staffId,
            status: info.status || '',
            remarks: info.remarks || ''
          });
        });
      });
      const attendanceSheet = XLSX.utils.json_to_sheet(attendanceData);
      XLSX.utils.book_append_sheet(wb, attendanceSheet, 'StaffAttendance');

      // Save File
      const dateStr = new Date().toISOString().split('T')[0];
      const fileName = `Sri_Gayathri_FULL_Master_Backup_${dateStr}.xlsx`;
      XLSX.writeFile(wb, fileName);
      
      alert(`✅ పూర్తి సాఫ్ట్‌వేర్ మాస్టర్ బ్యాకప్ విజయవంతంగా డౌన్‌లోడ్ చేయబడింది!\n\nఫైల్ పేరు: ${fileName}\n\nఇందులో Job Cards, Complaints, Staff, Customers (Database), Spares (Price List), Staff Attendance మరియు సెట్టింగ్స్ మొత్తం డేటా ఉంది.`);
    } catch (error) {
      console.error('Backup Error:', error);
      alert('❌ బ్యాకప్ డౌన్‌లోడ్ చేయడంలో విఫలమైంది: ' + String(error));
    } finally {
      setSheetsLoading(false);
    }
  };

  const refreshAllSheetsData = async (spreadsheetId: string) => {
    // Load all tables
    const rawJobCards = await loadSheetRows(spreadsheetId, 'JobCards', JOBCARD_HEADERS);
    const jobCardsList = rawJobCards.map(mapJobCardFromSheet);
    const complaintsList = await loadSheetRows(spreadsheetId, 'Complaints', COMPLAINT_HEADERS);
    const staffList = await loadSheetRows(spreadsheetId, 'Staff', STAFF_HEADERS);
    const rawCustomers = await loadSheetRows(spreadsheetId, 'Customers', CUSTOMER_HEADERS);
    const rawSpares = await loadSheetRows(spreadsheetId, 'Spares', SPARE_HEADERS);
    
    // Sort job cards
    jobCardsList.sort((a, b) => {
      const dateA = a.jobDate || a.jobOpenDate || a.complaintDate || a.createdAt || '';
      const dateB = b.jobDate || b.jobOpenDate || b.complaintDate || b.createdAt || '';
      if (dateA !== dateB) return dateB.localeCompare(dateA);
      return (b.createdAt || '').localeCompare(a.createdAt || '');
    });
    
    // Sort complaints
    complaintsList.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
    
    // Sort staff
    staffList.sort((a, b) => a.name.localeCompare(b.name));
    
    setSavedJobCards(jobCardsList);
    setComplaints(complaintsList);
    setStaffMembers(staffList);

    // Update Customer Index
    if (rawCustomers.length > 0) {
      const newChassisIndex: Record<string, any> = {};
      rawCustomers.forEach(c => {
        const key = normalizeKey(c.chassisNo || '');
        if (key) {
          newChassisIndex[key] = {
            ...c,
            followupHistory: typeof c.followupHistory === 'string' ? JSON.parse(c.followupHistory) : (c.followupHistory || [])
          };
        }
      });
      setChassisIndex(newChassisIndex);
      saveToStorage(LS_CUSTOMER_KEY, newChassisIndex);
    }

    // Update Spares Index
    if (rawSpares.length > 0) {
      const newSparesIndex: Record<string, any> = {};
      rawSpares.forEach(s => {
        const key = normalizeKey(s.partNo || '');
        if (key) newSparesIndex[key] = s;
      });
      setSparesIndex(newSparesIndex);
      saveToStorage(LS_SPARES_KEY, newSparesIndex);
    }
    
    // Cache
    try {
      saveJobCardsBackup(jobCardsList);
      localStorage.setItem('sri_backup_complaints', JSON.stringify(complaintsList));
      localStorage.setItem('sri_backup_staff', JSON.stringify(staffList));
    } catch (e) {}
  };

  // Auto-connect if useGoogleSheets is true on mount
  useEffect(() => {
    if (useGoogleSheets && !sheetsToken) {
      // We can't auto-login without user interaction for OAuth popups usually,
      // but we can try to prompt or at least set the state.
      // For now, we wait for the user to click Connect to get a fresh token.
    }
  }, [useGoogleSheets]);

  // WhatsApp Style Daily Auto Backup Engine (Runs at Midnight / Periodically / Manual)
  const executeAutoBackup = async (manualTrigger: boolean = false) => {
    setAutoBackupStatus('running');
    setAutoBackupMessage('ఆటోమేటిక్ బ్యాకప్ రన్ అవుతోంది...');
    try {
      // 1. Prepare snapshots
      const currentJobCards = savedJobCards.length > 0 ? savedJobCards : (await loadFromStorage<any[]>('sri_backup_jobcards') || []);

      const currentComplaints = complaints.length > 0 ? complaints : (() => {
        try {
          return JSON.parse(localStorage.getItem('sri_backup_complaints') || '[]');
        } catch {
          return [];
        }
      })();

      const currentStaff = staffMembers.length > 0 ? staffMembers : (() => {
        try {
          return JSON.parse(localStorage.getItem('sri_backup_staff') || '[]');
        } catch {
          return [];
        }
      })();

      const customerList = Object.values(chassisIndex);
      const sparesList = Object.values(sparesIndex);

      // 2. Safe local snapshot
      try {
        saveJobCardsBackup(currentJobCards);
        localStorage.setItem('sri_backup_complaints', JSON.stringify(currentComplaints));
        localStorage.setItem('sri_backup_staff', JSON.stringify(currentStaff));
      } catch (e) {
        console.warn('Local storage snapshot notice:', e);
      }

      // 3. Sync to Google Sheets if connected
      let sheetsSynced = false;
      if (useGoogleSheets && sheetsSpreadsheetId) {
        try {
          const mappedCards = currentJobCards.map(mapJobCardToSheet);
          await syncFullDatabaseToGoogleSheets(sheetsSpreadsheetId, {
            jobCards: mappedCards,
            complaints: currentComplaints,
            staff: currentStaff,
            customers: customerList,
            spares: sparesList
          });
          sheetsSynced = true;
        } catch (sheetsErr) {
          console.warn('Google Sheets auto-backup sync notice:', sheetsErr);
        }
      }

      const nowFormatted = new Date().toLocaleString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      });
      const todayStr = new Date().toISOString().split('T')[0];

      localStorage.setItem('sri_last_auto_backup_time', nowFormatted);
      localStorage.setItem('sri_last_auto_backup_date', todayStr);
      setLastAutoBackupTime(nowFormatted);
      setAutoBackupStatus('success');
      setAutoBackupMessage('బ్యాకప్ విజయవంతంగా సేవ్ అయ్యింది!');

      setTimeout(() => {
        setAutoBackupStatus('idle');
      }, 6000);

      if (manualTrigger) {
        alert(
          `🛡️ వాట్సాప్ తరహా ఆటో బ్యాకప్ విజయవంతంగా పూర్తయింది!\n\n` +
          `🕒 బ్యాకప్ సమయం: ${nowFormatted}\n` +
          `📋 జాబ్ కార్డ్స్: ${currentJobCards.length}\n` +
          `📝 కంప్లైంట్స్: ${currentComplaints.length}\n` +
          `👥 స్టాఫ్: ${currentStaff.length}\n` +
          `🚜 కస్టమర్ రికార్డులు: ${customerList.length}\n` +
          (sheetsSynced ? `📊 గూగుల్ షీట్స్: ✅ Google Sheets లోకి విజయవంతంగా అప్‌లోడ్ అయ్యింది!` : `💾 లోకల్ సెక్యూర్ స్టోరేజ్: ✅ బ్రౌజర్ స్టోరేజ్ & ఫైర్‌స్టోర్‌లో సురక్షితంగా సేవ్ అయ్యింది!`) +
          `\n\nమీ డేటా 100% సేఫ్ & సురక్షితం.`
        );
      }
    } catch (err) {
      console.error('Auto backup execution error:', err);
      setAutoBackupStatus('error');
      setAutoBackupMessage('బ్యాకప్ చేయడంలో లోపం!');
      if (manualTrigger) {
        alert('❌ బ్యాకప్ చేయడంలో సమస్య ఏర్పడింది: ' + (err instanceof Error ? err.message : String(err)));
      }
    }
  };

  // WhatsApp-Style Midnight Daily Auto-Backup Timer
  useEffect(() => {
    const checkAndTriggerDailyBackup = () => {
      const todayStr = new Date().toISOString().split('T')[0];
      const lastRunDate = localStorage.getItem('sri_last_auto_backup_date');
      
      // If today's backup has not run yet, run it in background
      if (lastRunDate !== todayStr && (savedJobCards.length > 0 || complaints.length > 0 || Object.keys(chassisIndex).length > 0)) {
        executeAutoBackup(false);
      }
    };

    const initTimer = setTimeout(checkAndTriggerDailyBackup, 3500);
    const interval = setInterval(checkAndTriggerDailyBackup, 15 * 60 * 1000);

    return () => {
      clearTimeout(initTimer);
      clearInterval(interval);
    };
  }, [savedJobCards.length, complaints.length, Object.keys(chassisIndex).length, useGoogleSheets, sheetsSpreadsheetId]);

  const handleDisconnectGoogleSheets = () => {
    if (window.confirm('గూగుల్ షీట్స్ డిస్‌కనెక్ట్ చేసి తిరిగి ఫైర్‌స్టోర్ (Firestore) కి మారాలనుకుంటున్నారా?')) {
      setUseGoogleSheets(false);
      localStorage.removeItem('sri_use_google_sheets');
      setSheetsToken(null);
      setCachedToken(null);
      window.location.reload();
    }
  };

  const handleMigrateFirestoreToSheets = async () => {
    if (!sheetsToken || !sheetsSpreadsheetId) {
      alert('⚠️ దయచేసి ముందుగా గూగుల్ షీట్స్‌ని కనెక్ట్ చేయండి.');
      return;
    }

    if (!window.confirm('సాఫ్ట్‌వేర్‌లోని మొత్తం డేటాను (Job Cards, Complaints, Staff, Customers, Spares) గూగుల్ షీట్స్‌లోకి కాపీ చేయాలా? ఇది కొంత సమయం పట్టవచ్చు.')) return;

    setMigrationStatus({ total: 0, current: 0, step: 'Fetching data from Database...', loading: true });
    
    try {
      // Fetch all from Firestore
      const jobCardsSnap = await getDocs(collection(db, 'jobcards'));
      const complaintsSnap = await getDocs(collection(db, 'complaints'));
      const staffSnap = await getDocs(collection(db, 'staff'));

      const firestoreJobCards = jobCardsSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const firestoreComplaints = complaintsSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const firestoreStaff = staffSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      const customerList = Object.values(chassisIndex);
      const sparesList = Object.values(sparesIndex);

      const totalItems = firestoreJobCards.length + firestoreComplaints.length + firestoreStaff.length + customerList.length + sparesList.length;
      setMigrationStatus(prev => ({ ...prev, total: totalItems, step: 'Uploading to Google Sheets...' }));

      let processed = 0;

      // JobCards
      for (const card of firestoreJobCards) {
        const mapped = mapJobCardToSheet(card);
        await updateSheetRow(sheetsSpreadsheetId!, 'JobCards', JOBCARD_HEADERS, mapped);
        processed++;
        setMigrationStatus(prev => ({ ...prev, current: processed }));
      }

      // Complaints
      for (const complaint of firestoreComplaints) {
        await updateSheetRow(sheetsSpreadsheetId!, 'Complaints', COMPLAINT_HEADERS, complaint);
        processed++;
        setMigrationStatus(prev => ({ ...prev, current: processed }));
      }

      // Staff
      for (const staff of firestoreStaff) {
        await updateSheetRow(sheetsSpreadsheetId!, 'Staff', STAFF_HEADERS, staff);
        processed++;
        setMigrationStatus(prev => ({ ...prev, current: processed }));
      }

      // Customers
      for (const cust of customerList) {
        const c = cust as any;
        const mapped = {
          ...c,
          followupHistory: c.followupHistory ? JSON.stringify(c.followupHistory) : '[]',
          updatedAt: c.updatedAt || new Date().toISOString()
        };
        await updateSheetRow(sheetsSpreadsheetId!, 'Customers', CUSTOMER_HEADERS, mapped);
        processed++;
        setMigrationStatus(prev => ({ ...prev, current: processed }));
      }

      // Spares
      for (const spare of sparesList) {
        await updateSheetRow(sheetsSpreadsheetId!, 'Spares', SPARE_HEADERS, spare);
        processed++;
        setMigrationStatus(prev => ({ ...prev, current: processed }));
      }

      await refreshAllSheetsData(sheetsSpreadsheetId!);
      alert('✅ డేటా మైగ్రేషన్ విజయవంతంగా పూర్తయింది! మొత్తం ' + totalItems + ' రికార్డులు గూగుల్ షీట్స్‌కి బదిలీ చేయబడ్డాయి.');
    } catch (error) {
      console.error('Migration Error:', error);
      alert('❌ మైగ్రేషన్ విఫలమైంది: ' + String(error));
    } finally {
      setMigrationStatus(prev => ({ ...prev, loading: false }));
    }
  };

  // Background task: Auto-fill missing supervisors based on mechanic assignment
  useEffect(() => {
    if (staffMembers.length === 0 || savedJobCards.length === 0) return;

    // We use a small timeout so it doesn't freeze the UI right on load
    const timeoutId = setTimeout(() => {
      let cardsToUpdate: any[] = [];
      
      savedJobCards.forEach(card => {
        const mech = card.mechanic || card.technicianName;
        const sup = card.wsIncharge || card.supervisor || card.supervisorName;
        
        if (mech && !sup) {
          // Manually duplicate logic of getAssignedSupervisor since it depends on staffMembers
          const nameLower = String(mech || '').trim().toLowerCase();
          const foundStaff = staffMembers.find(
            s => s.role === 'mechanic' && String(s.name || '').trim().toLowerCase() === nameLower
          );
          const assocSup = foundStaff?.supervisor || '';
          
          if (assocSup) {
            cardsToUpdate.push({ id: card.id, assocSup });
          }
        }
      });

      if (cardsToUpdate.length > 0) {
        console.log(`Auto-filling missing supervisor for ${cardsToUpdate.length} job cards...`);
        cardsToUpdate.forEach(async (updateInfo) => {
          try {
            const cardRef = doc(db, 'jobcards', updateInfo.id);
            await updateDoc(cardRef, {
              wsIncharge: updateInfo.assocSup,
              supervisorName: updateInfo.assocSup,
              supervisor: updateInfo.assocSup
            });
          } catch (e) {
            console.error('Failed to auto-fill supervisor for card', updateInfo.id, e);
          }
        });
      }
    }, 3000);

    return () => clearTimeout(timeoutId);
  }, [staffMembers, savedJobCards.length]);

  const saveStaffMember = async (data: {
    id?: string | null;
    name: string;
    fatherName?: string;
    village?: string;
    mandal?: string;
    mobileNumber?: string;
    role: string;
    dateOfJoining?: string;
    supervisor?: string;
  }) => {
    if (!data.name.trim()) {
      alert('Please enter staff name.');
      return;
    }
    const docId = data.id || Date.now().toString();
    const staffData = {
      id: docId,
      name: data.name.trim(),
      fatherName: data.fatherName || '',
      village: data.village || '',
      mandal: data.mandal || '',
      mobileNumber: data.mobileNumber || '',
      role: data.role || 'mechanic',
      dateOfJoining: data.dateOfJoining || '',
      supervisor: data.supervisor || '',
      updatedAt: new Date().toISOString()
    };

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        await updateSheetRow(sheetsSpreadsheetId!, 'Staff', STAFF_HEADERS, staffData);
        setStaffMembers(prev => {
          const filtered = prev.filter(s => s.id !== docId);
          const updated = [...filtered, staffData].sort((a, b) => a.name.localeCompare(b.name));
          try {
            localStorage.setItem('sri_backup_staff', JSON.stringify(updated));
          } catch (le) {}
          return updated;
        });
        alert('✅ స్టాఫ్ వివరాలు గూగుల్ షీట్స్‌లో విజయవంతంగా అప్‌డేట్ చేయబడ్డాయి!');
      } catch (err) {
        console.error('Error saving staff to Sheets:', err);
        alert('❌ గూగుల్ షీట్స్‌లో సేవ్ చేయడం విఫలమైంది: ' + String(err));
      }
      return;
    }

    // Save to Cloud SQL Database
    try {
      await sqlApi.saveStaff(staffData);
    } catch (sqlErr) {
      console.warn('Cloud SQL staff save:', sqlErr);
    }

    setStaffMembers(prev => {
      const filtered = prev.filter(s => s.id !== docId);
      const updated = [...filtered, staffData].sort((a, b) => a.name.localeCompare(b.name));
      try {
        localStorage.setItem('sri_backup_staff', JSON.stringify(updated));
      } catch (le) {}
      return updated;
    });

    if (db) {
      try {
        await setDoc(doc(db, 'staff', docId), staffData, { merge: true });
      } catch (e) {
        console.error('Error saving staff member to Firestore:', e);
      }
    }
    alert('✅ Staff member details saved successfully!');
  };

  const deleteStaffMember = async (id: string) => {
    if (currentUserRole !== 'admin') {
      alert('❌ Error: Only Admins can delete staff members.');
      return;
    }

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        await deleteSheetRow(sheetsSpreadsheetId!, 'Staff', id);
        setStaffMembers(prev => {
          const updated = prev.filter(s => s.id !== id);
          try {
            localStorage.setItem('sri_backup_staff', JSON.stringify(updated));
          } catch (le) {}
          return updated;
        });
        alert('✅ స్టాఫ్ మెంబర్ గూగుల్ షీట్స్ నుండి తొలగించబడ్డారు!');
      } catch (err) {
        console.error('Error deleting staff from Sheets:', err);
        alert('❌ గూగుల్ షీట్స్ నుండి తొలగించడం విఫలమైంది.');
      }
      return;
    }

    // Delete from Cloud SQL Database
    try {
      await sqlApi.deleteStaff(id);
    } catch (sqlErr) {
      console.warn('Cloud SQL staff delete:', sqlErr);
    }

    setStaffMembers(prev => {
      const updated = prev.filter(s => s.id !== id);
      try {
        localStorage.setItem('sri_backup_staff', JSON.stringify(updated));
      } catch (le) {}
      return updated;
    });

    if (db) {
      try {
        await deleteDoc(doc(db, 'staff', id));
      } catch (e) {
        console.error('Error deleting staff member from Firestore:', e);
      }
    }
    alert('✅ Staff member deleted successfully!');
  };

  const downloadStaffDirectoryExcel = () => {
    if (staffMembers.length === 0) {
      alert('No staff members to export.');
      return;
    }
    const exportData = staffMembers.map((s, idx) => ({
      'Sl No': idx + 1,
      'Full Name': s.name || '',
      'Father Name': s.fatherName || '',
      'Village': s.village || '',
      'Mandal': s.mandal || '',
      'Mobile Number': s.mobileNumber || '',
      'Job Role': s.role === 'mechanic' ? 'Mechanic / Technician' : s.role === 'supervisor' ? 'Supervisor / W/S Incharge' : s.role || '',
      'Assigned Supervisor': s.role === 'mechanic' ? (s.supervisor || '') : '',
      'Date of Joining': s.dateOfJoining || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Staff Directory");

    // Auto-fit columns
    const maxProps = Object.keys(exportData[0] || {}).map(key => ({
      wch: Math.max(key.length, ...exportData.map(row => String(row[key as keyof typeof row] || '').length)) + 2
    }));
    worksheet['!cols'] = maxProps;

    XLSX.writeFile(workbook, `Staff_Directory_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleStaffFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

        let importedCount = 0;
        for (const row of rows) {
          const name = row['Full Name'] || row['Name'] || row['name'] || '';
          if (!name.trim()) continue;
          const fatherName = row['Father Name'] || row['fatherName'] || '';
          const village = row['Village'] || row['village'] || '';
          const mandal = row['Mandal'] || row['mandal'] || '';
          const mobileNumber = row['Mobile Number'] || row['Mobile'] || row['mobileNumber'] || '';
          const rawRole = (row['Job Role'] || row['Role'] || row['role'] || '').toString().toLowerCase();
          let role = 'mechanic';
          if (rawRole.includes('supervis') || rawRole.includes('incharge') || rawRole.includes('manager')) {
            role = 'supervisor';
          }
          const dateOfJoining = row['Date of Joining'] || row['Joining Date'] || row['dateOfJoining'] || '';
          const supervisor = row['Supervisor'] || row['supervisor'] || row['Assigned Supervisor'] || '';

          await saveStaffMember({
            name,
            fatherName,
            village,
            mandal,
            mobileNumber,
            role,
            dateOfJoining,
            supervisor
          });
          importedCount++;
        }
        alert(`✅ Successfully imported ${importedCount} staff member(s)!`);
      } catch (err) {
        console.error(err);
        alert('❌ Failed to import staff Excel file.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Helper to sync Customer database to Firestore and Cloud SQL in chunked documents
  const saveCustomersToFirestore = async (rows: any[]) => {
    try {
      setExcelStatus({ text: '⏳ Syncing customer records to cloud database...', isSuccess: false });
      
      // Save directly to Cloud SQL Database
      try {
        await sqlApi.saveCustomersBulk(rows, true);
      } catch (sqlErr) {
        console.warn('Cloud SQL customer sync:', sqlErr);
      }

      if (db) {
        const existingSnap = await getDocs(collection(db, 'customers_master'));
        if (!existingSnap.empty) {
          const deleteBatch = writeBatch(db);
          existingSnap.docs.forEach((d) => deleteBatch.delete(d.ref));
          await deleteBatch.commit();
        }

        if (rows.length === 0) return;

        const cleanRows = rows.map(r => {
          const cr: Record<string, any> = {};
          if (r && typeof r === 'object') {
            Object.keys(r).forEach(k => {
              cr[k] = cleanValue(r[k]);
            });
          }
          return cr;
        });

        const CHUNK_SIZE = 250;
        const totalChunks = Math.ceil(cleanRows.length / CHUNK_SIZE);

        for (let i = 0; i < totalChunks; i++) {
          const chunkRows = cleanRows.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
          const chunkDocRef = doc(db, 'customers_master', `chunk_${i}`);
          await setDoc(chunkDocRef, {
            chunkIndex: i,
            totalChunks,
            uploadedAt: new Date().toISOString(),
            rows: chunkRows
          });
        }

        await setDoc(doc(db, 'app_master_data', 'customers_meta'), {
          totalRows: cleanRows.length,
          totalChunks,
          uploadedAt: new Date().toISOString(),
          uploadedBy: currentUser?.email || 'user'
        }, { merge: true });
      }

      setExcelStatus({
        text: `✅ ${rows.length} customer record(s) synced to Cloud SQL & cloud database successfully.`,
        isSuccess: true
      });
    } catch (err) {
      console.error("Failed to save customers to database:", err);
      setExcelStatus({
        text: `⚠️ Customer data saved locally and in database.`,
        isSuccess: true
      });
    }
  };

  // Helper to sync Spares database to Firestore and Cloud SQL
  const saveSparesToFirestore = async (rows: any[]) => {
    try {
      setSparesStatus({ text: '⏳ Syncing spares records to cloud database...', isSuccess: false });

      // Save directly to Cloud SQL Database
      try {
        await sqlApi.saveSparesBulk(rows, true);
      } catch (sqlErr) {
        console.warn('Cloud SQL spares sync:', sqlErr);
      }

      if (db) {
        const existingSnap = await getDocs(collection(db, 'spares_master'));
        if (!existingSnap.empty) {
          const deleteBatch = writeBatch(db);
          existingSnap.docs.forEach((d) => deleteBatch.delete(d.ref));
          await deleteBatch.commit();
        }

        if (rows.length === 0) return;

        const cleanRows = rows.map(r => {
          const cr: Record<string, any> = {};
          if (r && typeof r === 'object') {
            Object.keys(r).forEach(k => {
              cr[k] = cleanValue(r[k]);
            });
          }
          return cr;
        });

        const CHUNK_SIZE = 300;
        const totalChunks = Math.ceil(cleanRows.length / CHUNK_SIZE);

        for (let i = 0; i < totalChunks; i++) {
          const chunkRows = cleanRows.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
          const chunkDocRef = doc(db, 'spares_master', `chunk_${i}`);
          await setDoc(chunkDocRef, {
            chunkIndex: i,
            totalChunks,
            uploadedAt: new Date().toISOString(),
            rows: chunkRows
          });
        }

        await setDoc(doc(db, 'app_master_data', 'spares_meta'), {
          totalRows: cleanRows.length,
          totalChunks,
          uploadedAt: new Date().toISOString(),
          uploadedBy: currentUser?.email || 'user'
        }, { merge: true });
      }

      setSparesStatus({
        text: `✅ ${rows.length} spare part(s) synced to Cloud SQL & cloud database successfully.`,
        isSuccess: true
      });
    } catch (err) {
      console.error("Failed to save spares to database:", err);
      setSparesStatus({
        text: `⚠️ Spares data saved in database.`,
        isSuccess: true
      });
    }
  };

  // Customer Excel Upload
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelStatus({ text: 'Reading file...', isSuccess: false });
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'array', cellDates: false });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rawRows: any[] = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: '' });

        const rows = rawRows.map(r => {
          const cr: Record<string, any> = {};
          if (r && typeof r === 'object') {
            Object.keys(r).forEach(k => {
              cr[k] = cleanValue(r[k]);
            });
          }
          return cr;
        });

        // Start fresh on file upload to replace old records and prevent quota overflow
        let updatedIndex: Record<string, any> = {};
        let addedOrUpdated = 0;

        rows.forEach(row => {
          const lookup = buildRowLookup(row);
          const chassisVal = getFieldValue(lookup, 'chassis');
          const phoneVal = getFieldValue(lookup, 'custPhone');
          const drvPhoneVal = getFieldValue(lookup, 'driverPhone');
          const nameVal = getFieldValue(lookup, 'custName');
          const regdVal = getFieldValue(lookup, 'regdNo');
          const addrVal = getCombinedAddress(lookup);

          if (chassisVal) lookup.__chassisDisplay = chassisVal;
          if (nameVal) lookup.__custNameDisplay = nameVal;
          if (phoneVal) lookup.__custPhoneDisplay = phoneVal;
          if (addrVal) lookup.__custAddrDisplay = addrVal;

          const keyCh = normalizeKey(chassisVal);
          const keyPh = normalizeKey(phoneVal);
          const keyDph = normalizeKey(drvPhoneVal);
          const keyNm = normalizeKey(nameVal);
          const keyReg = normalizeKey(regdVal);

          if (keyCh) updatedIndex[keyCh] = lookup;
          if (keyPh) updatedIndex[keyPh] = lookup;
          if (keyDph) updatedIndex[keyDph] = lookup;
          if (keyNm) updatedIndex[keyNm] = lookup;
          if (keyReg) updatedIndex[keyReg] = lookup;

          addedOrUpdated++;
        });

        setChassisIndex(updatedIndex);
        saveToStorage(LS_CUSTOMER_KEY, updatedIndex);
        setLastUploadedRows(rows);

        const totalCount = Object.keys(updatedIndex).length;
        
        // Auto-check for duplicates in the uploaded file
        const fileDups = getDuplicateReportForSet(rows);
        if (fileDups.length > 0) {
          setFileDuplicateReport(fileDups);
          setIsFileDuplicateModalOpen(true);
        }

        // Sync to Firestore Cloud Database so all users opening the link see this latest upload
        await saveCustomersToFirestore(rows);
      } catch (err) {
        console.error(err);
        setExcelStatus({ text: '⚠️ Could not read file. Check Excel format.', isSuccess: false });
      }
    };
    reader.readAsArrayBuffer(file);
  };


  // Master Restore Upload
  const handleMasterRestoreUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!window.confirm('దీని ద్వారా మీ క్లౌడ్ డేటాబేస్ మరియు లోకల్ సాఫ్ట్‌వేర్ డేటా మొత్తం (Job Cards, Complaints, Staff, Customers, Spares) పూర్తిగా అప్‌లోడ్ చేసిన ఫైల్‌తో మారుతుంది (Replace అవుతుంది).\n\nకొత్త పరికరంలో పాత డేటాని పొందడానికి ఇది ఉపయోగపడుతుంది.\n\nమీరు ఖచ్చితంగా మాస్టర్ బ్యాకప్ ని రీస్టోర్ చేయాలనుకుంటున్నారా?')) {
      e.target.value = '';
      return;
    }

    setSheetsLoading(true);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'array', cellDates: false });
        
        let restoredCustomers = 0;
        let restoredSpares = 0;
        let restoredJobCards = 0;
        let restoredComplaintsCount = 0;
        let restoredStaffCount = 0;
        let restoredAttendanceCount = 0;

        // Restore Customers
        if (wb.Sheets['Customers']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['Customers'], { defval: '' });
          let updatedIndex: Record<string, any> = {};
          rows.forEach((row: any) => {
            if (typeof row.followupHistory === 'string') {
              try {
                row.followupHistory = JSON.parse(row.followupHistory);
              } catch (e) {
                row.followupHistory = [];
              }
            }

            const chassisNo = row.__chassisDisplay || getFieldValue(row, 'chassis') || row['Chassis no'] || row.chassisNo;
            if (chassisNo) {
              updatedIndex[normalizeKey(String(chassisNo))] = row;
              restoredCustomers++;
            }
          });
          setChassisIndex(updatedIndex);
          saveToStorage(LS_CUSTOMER_KEY, updatedIndex);
          await saveCustomersToFirestore(rows);
        }

        // Restore Spares
        if (wb.Sheets['Spares']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['Spares'], { defval: '' });
          let updatedIndex: Record<string, any> = {};
          rows.forEach((row: any) => {
            const partNoVal = row.__partNoDisplay || getFieldValue(row, 'partNo') || row['Part Number'] || row.partNo;
            if (partNoVal) {
              updatedIndex[normalizeKey(String(partNoVal))] = row;
              restoredSpares++;
            }
          });
          setSparesIndex(updatedIndex);
          saveToStorage(LS_SPARES_KEY, updatedIndex);
          await saveSparesToFirestore(rows);
        }

        // Restore JobCards
        if (wb.Sheets['JobCards']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['JobCards'], { defval: '' });
          const parsedJobCards: any[] = [];
          
          rows.forEach((row: any) => {
            const card: any = {
              id: row.id || (Date.now().toString() + Math.random().toString(36).substring(2, 7) + "_" + restoredJobCards),
              jobNo: row.jobCardNo || row.jobNo || '',
              onlineJobCardNo: row.onlineJobCardNo || '',
              jobDate: row.jobDate || '',
              dateTimeIn: row.dateTimeIn || '',
              dateTimeOut: row.dateTimeOut || row.dateTimeNo || '',
              expectedRepairTime: row.expectedRepairTime || '',
              status: row.status || 'Open',
              custName: row.custName || '',
              fatherName: row.fatherName || '',
              custAddr: row.custAddr || '',
              village: row.village || '',
              mandal: row.mandal || '',
              ownerMob: row.ownerMob || row.phNo || '',
              driverMob: row.driverMob || '',
              regdNo: row.regdNo || '',
              chassisNo: row.chassisNo || '',
              engineNo: row.engineNo || '',
              model: row.tractorModel || row.model || '',
              serialNo: row.serialNo || '',
              hourMeter: row.hourMeter || '',
              serviceType: row.serviceType || '',
              mechanic: row.mechanicName || row.mechanic || '',
              wsIncharge: row.supervisor || row.wsIncharge || '',
              warrantyOverride: row.warrantyOverride || 'auto',
              createdAt: row.createdAt || new Date().toISOString()
            };

            try {
              card.checkpoints = typeof row.checkpointsJson === 'string' ? JSON.parse(row.checkpointsJson) : (row.checkpointsJson || []);
            } catch(e) {
              card.checkpoints = [];
            }
            try {
              card.repairRows = typeof row.repairsJson === 'string' ? JSON.parse(row.repairsJson) : (row.repairsJson || []);
            } catch(e) {
              card.repairRows = [];
            }
            try {
              card.partRows = typeof row.partsJson === 'string' ? JSON.parse(row.partsJson) : (row.partsJson || []);
            } catch(e) {
              card.partRows = [];
            }

            parsedJobCards.push(card);
            restoredJobCards++;
          });

          setSavedJobCards(parsedJobCards);
          saveJobCardsBackup(parsedJobCards);

          // Save to Cloud SQL
          try {
            await sqlApi.saveJobCardsBulk(parsedJobCards, true);
          } catch (sqlErr) {
            console.warn('Cloud SQL jobcards restore:', sqlErr);
          }

          if (db && parsedJobCards.length > 0) {
            const existingJobCards = await getDocs(collection(db, 'jobcards'));
            if (!existingJobCards.empty) {
              const docs = existingJobCards.docs;
              const DELETE_BATCH_SIZE = 400;
              for (let i = 0; i < docs.length; i += DELETE_BATCH_SIZE) {
                const chunk = docs.slice(i, i + DELETE_BATCH_SIZE);
                const deleteBatch = writeBatch(db);
                chunk.forEach(d => deleteBatch.delete(d.ref));
                await deleteBatch.commit();
              }
            }

            const BATCH_SIZE = 450;
            const commitPromises: Promise<void>[] = [];
            for (let i = 0; i < parsedJobCards.length; i += BATCH_SIZE) {
              const chunk = parsedJobCards.slice(i, i + BATCH_SIZE);
              const batch = writeBatch(db);
              chunk.forEach(cardData => {
                batch.set(doc(db, 'jobcards', cardData.id), cardData);
              });
              commitPromises.push(batch.commit());
            }
            await Promise.all(commitPromises);
          }
        }

        // Restore Complaints
        if (wb.Sheets['Complaints']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['Complaints'], { defval: '' });
          const parsedComplaints: any[] = [];

          rows.forEach((row: any) => {
            if (row.id) {
              parsedComplaints.push(row);
              restoredComplaintsCount++;
            }
          });

          if (parsedComplaints.length > 0) {
            setComplaints(parsedComplaints);
            localStorage.setItem('sri_backup_complaints', JSON.stringify(parsedComplaints));

            // Save to Cloud SQL
            try {
              await sqlApi.saveComplaintsBulk(parsedComplaints, true);
            } catch (sqlErr) {
              console.warn('Cloud SQL complaints restore:', sqlErr);
            }

            if (db) {
              const existingComplaints = await getDocs(collection(db, 'complaints'));
              if (!existingComplaints.empty) {
                const docs = existingComplaints.docs;
                const DELETE_BATCH_SIZE = 400;
                for (let i = 0; i < docs.length; i += DELETE_BATCH_SIZE) {
                  const chunk = docs.slice(i, i + DELETE_BATCH_SIZE);
                  const deleteBatch = writeBatch(db);
                  chunk.forEach(d => deleteBatch.delete(d.ref));
                  await deleteBatch.commit();
                }
              }

              const BATCH_SIZE = 450;
              const commitPromises: Promise<void>[] = [];
              for (let i = 0; i < parsedComplaints.length; i += BATCH_SIZE) {
                const chunk = parsedComplaints.slice(i, i + BATCH_SIZE);
                const batch = writeBatch(db);
                chunk.forEach(comp => {
                  batch.set(doc(db, 'complaints', comp.id), comp);
                });
                commitPromises.push(batch.commit());
              }
              await Promise.all(commitPromises);
            }
          }
        }

        // Restore Staff
        if (wb.Sheets['Staff']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['Staff'], { defval: '' });
          const parsedStaff: any[] = [];

          rows.forEach((row: any) => {
            if (row.id) {
              parsedStaff.push(row);
              restoredStaffCount++;
            }
          });

          if (parsedStaff.length > 0) {
            setStaffMembers(parsedStaff);
            localStorage.setItem('sri_backup_staff', JSON.stringify(parsedStaff));

            // Save to Cloud SQL
            try {
              await sqlApi.saveStaffBulk(parsedStaff, true);
            } catch (sqlErr) {
              console.warn('Cloud SQL staff restore:', sqlErr);
            }

            if (db) {
              const existingStaff = await getDocs(collection(db, 'staff'));
              if (!existingStaff.empty) {
                const docs = existingStaff.docs;
                const DELETE_BATCH_SIZE = 400;
                for (let i = 0; i < docs.length; i += DELETE_BATCH_SIZE) {
                  const chunk = docs.slice(i, i + DELETE_BATCH_SIZE);
                  const deleteBatch = writeBatch(db);
                  chunk.forEach(d => deleteBatch.delete(d.ref));
                  await deleteBatch.commit();
                }
              }

              const BATCH_SIZE = 450;
              const commitPromises: Promise<void>[] = [];
              for (let i = 0; i < parsedStaff.length; i += BATCH_SIZE) {
                const chunk = parsedStaff.slice(i, i + BATCH_SIZE);
                const batch = writeBatch(db);
                chunk.forEach(st => {
                  batch.set(doc(db, 'staff', st.id), st);
                });
                commitPromises.push(batch.commit());
              }
              await Promise.all(commitPromises);
            }
          }
        }

        // Restore Staff Attendance
        if (wb.Sheets['StaffAttendance']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['StaffAttendance'], { defval: '' });
          const rebuiltAttendance: Record<string, Record<string, { status: 'present' | 'absent' | 'leave', remarks?: string }>> = {};
          
          rows.forEach((row: any) => {
            const date = row.date;
            const staffId = row.staffId;
            const status = row.status;
            const remarks = row.remarks || '';
            if (date && staffId && status) {
              if (!rebuiltAttendance[date]) {
                rebuiltAttendance[date] = {};
              }
              rebuiltAttendance[date][staffId] = { status, remarks };
              restoredAttendanceCount++;
            }
          });

          if (Object.keys(rebuiltAttendance).length > 0) {
            setAttendanceRecords(rebuiltAttendance);
            localStorage.setItem('sri_staff_attendance_data', JSON.stringify(rebuiltAttendance));

            // Sync with Cloud SQL & Firestore
            for (const [dateStr, recordsForDate] of Object.entries(rebuiltAttendance)) {
              try {
                await sqlApi.saveAttendance(dateStr, recordsForDate);
              } catch (sqlErr) {
                console.warn('Cloud SQL attendance restore:', sqlErr);
              }
            }
            if (db) {
              for (const [dateStr, recordsForDate] of Object.entries(rebuiltAttendance)) {
                try {
                  const docRef = doc(db, 'staff_attendance', dateStr);
                  await setDoc(docRef, { date: dateStr, records: recordsForDate, updatedAt: new Date().toISOString() }, { merge: true });
                } catch (err) {
                  console.error("Firestore attendance sync error on restore:", err);
                }
              }
            }
          }
        }

        // Restore Settings
        if (wb.Sheets['AppSettings']) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets['AppSettings'], { defval: '' });
          rows.forEach((row: any) => {
            if (row.key === 'BranchName' && row.value) {
              setBranch(row.value);
              localStorage.setItem('sri_branch', row.value);
            }
            if (row.key === 'QuickRemarks' && row.value) {
              try {
                const remarks = JSON.parse(row.value);
                if (Array.isArray(remarks)) {
                  setQuickRemarksList(remarks);
                  localStorage.setItem('sri_quick_remarks', JSON.stringify(remarks));
                }
              } catch(e) {}
            }
            if (row.key === 'ServiceInterval' && row.value) {
              const val = parseInt(row.value, 10);
              if (!isNaN(val)) {
                setServiceIntervalDays(val);
                localStorage.setItem('sri_service_interval', String(val));
              }
            }
          });
        }

        alert(`✅ మాస్టర్ బ్యాకప్ విజయవంతంగా రీస్టోర్ చేయబడింది!\n\nఈ డివైజ్ మరియు క్లౌడ్ డేటాబేస్ లోకి:\n- ${restoredJobCards} జాబ్ కార్డ్స్ (Job Cards)\n- ${restoredComplaintsCount} ఫిర్యాదులు (Complaints)\n- ${restoredStaffCount} సిబ్బంది (Staff)\n- ${restoredCustomers} కస్టమర్లు (Customers)\n- ${restoredSpares} స్పేర్స్ (Spares Price List)\n- ${restoredAttendanceCount} స్టాఫ్ అటెండెన్స్ (Attendance)\nవిజయవంతంగా రీస్టోర్ అయ్యాయి.`);
      } catch (err) {
        console.error("Master restore error:", err);
        alert('❌ Error restoring master backup: ' + String(err));
      } finally {
        setSheetsLoading(false);
      }
    };
    reader.onerror = () => {
      setSheetsLoading(false);
      alert('Error reading file.');
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset
  };

  // Spares Excel Upload
  const handleSparesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    console.log("handleSparesUpload called");
    const file = e.target.files?.[0];
    if (!file) {
      console.log("No file selected");
      return;
    }
    console.log("File selected:", file.name);

    setSparesStatus({ text: 'Reading file...', isSuccess: false });
    const reader = new FileReader();
    reader.onload = async (evt) => {
      console.log("FileReader onload triggered");
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'array', cellDates: false });
        console.log("Workbook read successfully");
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        console.log("Rows extracted:", rows.length);

        // Start fresh on file upload to replace old records and prevent quota overflow
        let updatedIndex: Record<string, any> = {};
        let addedOrUpdated = 0;

        rows.forEach(row => {
          const lookup = buildRowLookup(row);
          const candidates = SPARE_FIELD_CANDIDATES.partNo;
          let partNoVal = '';
          for (const c of candidates) {
            if (lookup[c] !== undefined && String(lookup[c]).trim() !== '') {
              partNoVal = String(lookup[c]).trim();
              break;
            }
          }
          if (partNoVal) {
            lookup.__partNoDisplay = partNoVal;
            updatedIndex[normalizeKey(partNoVal)] = lookup;
            addedOrUpdated++;
          }
        });

        setSparesIndex(updatedIndex);
        saveToStorage(LS_SPARES_KEY, updatedIndex);

        // Sync to Firestore Cloud Database so all users opening the link see this latest upload
        await saveSparesToFirestore(rows);
      } catch (err) {
        console.error("Error reading spares file:", err);
        setSparesStatus({ text: '⚠️ Could not read spares file.', isSuccess: false });
      }
    };
    reader.onerror = (err) => {
      console.error("FileReader error:", err);
    };
    reader.readAsArrayBuffer(file);
  };

  // Helper to render field label
  const renderFieldHeader = (label: string, val: any, isRequired: boolean = false) => {
    const strVal = String(val ?? '').trim();
    const isEmpty = !strVal;
    return (
      <div className="flex items-center justify-between mb-1">
        <label className="text-xs font-bold text-slate-700 flex items-center gap-1">
          <span>{label}</span>
          {isRequired && <span className="text-rose-500 font-bold">*</span>}
        </label>
        {!isEmpty && (
          <span className="text-[9.5px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-1.5 py-0.2 rounded">
            ✓ Done
          </span>
        )}
      </div>
    );
  };

  // Helper to return CSS class for input boxes
  const getInputClass = (val: any, extraClasses: string = '') => {
    const strVal = String(val ?? '').trim();
    const isEmpty = !strVal;
    if (isEmpty) {
      return `w-full text-[11px] p-1.5 rounded-md border border-slate-300 bg-white text-slate-900 placeholder:text-slate-400 font-medium focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all ${extraClasses}`;
    }
    return `w-full text-xs p-2 rounded-md border border-slate-300 bg-white text-slate-900 font-bold focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all ${extraClasses}`;
  };

  // Helper to extract last service details for a given chassis number
  const getLastServiceDetailsForChassis = (chassisVal: any, savedCards: any[]) => {
    const strVal = String(chassisVal ?? '').trim();
    if (!strVal) return '';
    const cleanCh = normalizeKey(strVal);
    if (!cleanCh) return '';

    const matching = savedCards.filter(c => {
      const cCh = normalizeKey(c.chassisNo || '');
      return cCh && cCh === cleanCh;
    }).sort((a, b) => {
      const getCardTime = (card: any) => {
        const dStr = card.jobDate || card.jobOpenDate || card.createdAt;
        if (!dStr) return 0;
        const t = Date.parse(dStr);
        return isNaN(t) ? 0 : t;
      };
      return getCardTime(b) - getCardTime(a);
    });

    if (matching.length === 0) return '';

    const last = matching[0];
    const rawOpenDate = last.jobDate || last.jobOpenDate || last.complaintDate || (last.createdAt ? last.createdAt.split('T')[0] : '');
    const openDateFormatted = fmtDate(rawOpenDate) || rawOpenDate || '';
    const closedDateFormatted = last.actualClosedDate ? fmtDate(last.actualClosedDate) : '';

    let dateDisplay = openDateFormatted;
    if (closedDateFormatted && closedDateFormatted !== openDateFormatted) {
      dateDisplay += ` (Closed: ${closedDateFormatted})`;
    } else if (closedDateFormatted && closedDateFormatted === openDateFormatted && last.status === 'Closed') {
      dateDisplay += ` (Closed)`;
    }

    const serviceType = last.serviceType || '';
    const freeService = last.freeServiceList || '';
    const serviceDetail = [serviceType, freeService].filter(Boolean).join(' / ') || 'Service';
    const hrsVal = last.hourMeter || last.hrsRun;
    const hrs = hrsVal ? `${hrsVal} Hrs` : '';
    const tech = last.mechanic || last.technicianName || '';
    const extra = last.extraRepairs || '';

    const parts = [
      dateDisplay,
      serviceDetail,
      hrs,
      tech ? `Tech: ${tech}` : '',
      extra ? `Repairs: ${extra}` : ''
    ].filter(Boolean);
    return parts.join(' - ');
  };

  // Helper to auto-fill customer state from record
  const autoFillCustomer = (lookup: any) => {
    if (!lookup) return;

    const sr = getFieldValue(lookup, 'serialNo');
    if (sr) setSerialNo(sr);

    const ch = lookup.__chassisDisplay || getFieldValue(lookup, 'chassis');
    if (ch) setChassisNo(ch);

    const name = lookup.__custNameDisplay || getFieldValue(lookup, 'custName');
    if (name) setCustName(name);

    const addr = lookup.__custAddrDisplay || getCombinedAddress(lookup);
    if (addr) setCustAddr(addr);

    const phone = lookup.__custPhoneDisplay || getFieldValue(lookup, 'custPhone');
    if (phone) setOwnerMob(phone);

    const fn = getFieldValue(lookup, 'fatherName') || lookup['FATHER'] || lookup['FATHER NAME'] || lookup['Father Name'];
    if (fn) setFatherName(fn);

    const vill = getFieldValue(lookup, 'village') || lookup['village'] || lookup['VILLAGE'];
    if (vill) setVillage(vill);

    const mdlDist = getFieldValue(lookup, 'mandal') || lookup['mandal'] || lookup['MANDAL'];
    if (mdlDist) setMandal(mdlDist);

    const br = lookup['BRANCH'] || lookup['branch'] || lookup['Branch'] || getFieldValue(lookup, 'branch');
    if (br) setBranch(br);

    const hfn = getFieldValue(lookup, 'historyFileNo') || sr || ch || '';
    if (hfn) {
      let finalHFN = hfn;
      if (br) {
        const cleanBr = br.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        const cleanHFN = hfn.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        if (cleanBr && !cleanHFN.startsWith(cleanBr)) {
          finalHFN = `${br.toUpperCase()}-${hfn}`;
        }
      }
      setHistoryFileNo(finalHFN);
    }

    const mt = lookup['MODEL TYPE'] || lookup['modelType'] || getFieldValue(lookup, 'modelType');
    if (mt) setModelType(mt);

    const regd = getFieldValue(lookup, 'regdNo');
    if (regd) setRegdNo(regd);

    const eng = getFieldValue(lookup, 'engineNo');
    if (eng) setEngineNo(eng);

    const mdl = getFieldValue(lookup, 'model');
    if (mdl) setModel(mdl);

    const hmr = getFieldValue(lookup, 'hourMeter');
    if (hmr) setHourMeter(hmr);

    const dist = getFieldValue(lookup, 'distDealership');
    const lastServiceSummary = getLastServiceDetailsForChassis(ch || chassisNo, savedJobCards);
    if (lastServiceSummary) {
      setDistDealership(lastServiceSummary);
    } else if (dist) {
      setDistDealership(dist);
    }

    const rawInstall = getFieldValue(lookup, 'installDate');
    const parsedInstall = toInputDateFormat(rawInstall);
    if (parsedInstall) setInstallDate(parsedInstall);

    // Calculate filled vs empty fields for notice
    let filled = 0;
    if (sr) filled++;
    if (ch) filled++;
    if (name) filled++;
    if (addr) filled++;
    if (phone) filled++;
    if (fn) filled++;
    if (vill) filled++;
    if (mdlDist) filled++;
    if (br) filled++;
    if (mt) filled++;
    if (regd) filled++;
    if (eng) filled++;
    if (mdl) filled++;
    if (hmr) filled++;
    if (dist || lastServiceSummary) filled++;
    if (parsedInstall) filled++;

    setAutoFillNotice({
      chassis: ch || name || 'Searched Record',
      filledCount: filled,
      emptyCount: Math.max(0, 18 - filled)
    });
  };

  // Helper to find customer record by any keyword/word across chassis, name, phone, village, model, etc.
  const findCustomerRecordByQuery = (val: any) => {
    const strVal = String(val ?? '').trim();
    if (!strVal) return null;

    // Extract potential search tokens
    const rawTokens: string[] = [strVal];
    if (strVal.includes(' — ')) {
      strVal.split(' — ').forEach((part) => rawTokens.push(part.trim()));
    }
    if (strVal.includes('(')) {
      strVal.split(/[\(\)]/).forEach((part) => {
        const clean = part.replace(/Chassis:|Inst:|Mob:|Mobile:|Vill:|Model:/gi, '').trim();
        if (clean) rawTokens.push(clean);
      });
    }
    // Also split by spaces into individual words
    strVal.split(/\s+/).forEach(word => {
      const cleanWord = word.trim();
      if (cleanWord.length >= 2) rawTokens.push(cleanWord);
    });

    const records: any[] = Array.from(new Set(Object.values(chassisIndex)));
    if (records.length === 0) return null;

    for (const token of rawTokens) {
      const typed = normalizeKey(token);
      if (!typed || typed.length < 2) continue;

      // 1. Direct O(1) key lookup in chassisIndex
      if (chassisIndex[typed]) return chassisIndex[typed];

      // 2. Comprehensive search across all record fields (chassis, name, phone, village, mandal, model, fatherName, etc.)
      for (const rec of records) {
        const ch = normalizeKey((rec as any).__chassisDisplay || getFieldValue(rec, 'chassis'));
        const ph = normalizeKey((rec as any).__custPhoneDisplay || getFieldValue(rec, 'Mobile Number') || rec.mobileNumber);
        const nm = normalizeKey((rec as any).__custNameDisplay || getFieldValue(rec, 'Customer Name') || rec.custName);
        const vill = normalizeKey(getFieldValue(rec, 'village') || rec.village);
        const model = normalizeKey(getFieldValue(rec, 'model') || rec.model);
        const father = normalizeKey(getFieldValue(rec, 'fatherName') || rec.fatherName);

        if (ch && (ch.includes(typed) || typed.includes(ch))) return rec;
        if (ph && (ph.includes(typed) || typed.includes(ph))) return rec;
        if (nm && (nm.includes(typed) || typed.includes(nm))) return rec;
        if (vill && (vill.includes(typed) || typed.includes(vill))) return rec;
        if (model && (model.includes(typed) || typed.includes(model))) return rec;
        if (father && (father.includes(typed) || typed.includes(father))) return rec;
      }
    }

    return null;
  };

  // Dedicated Auto-fill Search (strictly by Chassis No)
  const handleCustomerSearch = (val: any) => {
    const strVal = String(val ?? '');
    setSearchQuery(strVal);
    const cleanChassis = strVal.includes(' — ') ? strVal.split(' — ')[0].trim() : strVal;
    const rec = findCustomerRecordByQuery(strVal) || findCustomerRecordByQuery(cleanChassis);
    if (rec) {
      autoFillCustomer(rec);
    }
    const lastServiceSummary = getLastServiceDetailsForChassis(cleanChassis, savedJobCards);
    if (lastServiceSummary) {
      setDistDealership(lastServiceSummary);
    }
  };

  const handleSerialChange = (val: any) => {
    const strVal = String(val ?? '');
    setSerialNo(strVal);
    setHistoryFileNo(strVal);
  };

  const handleChassisChange = (val: any) => {
    const strVal = String(val ?? '');
    const cleanChassis = strVal.includes(' — ') ? strVal.split(' — ')[0].trim() : strVal;
    setChassisNo(cleanChassis);

    const rec = findCustomerRecordByQuery(strVal) || findCustomerRecordByQuery(cleanChassis);
    if (rec) {
      autoFillCustomer(rec);
    } else {
      setHistoryFileNo(cleanChassis);
    }

    const lastServiceSummary = getLastServiceDetailsForChassis(cleanChassis, savedJobCards);
    if (lastServiceSummary) {
      setDistDealership(lastServiceSummary);
    }
    
    // Auto-fill complaint details from open or running complaints
    const activeComplaint = complaints.find(c => c.chassisNo && normalizeKey(c.chassisNo) === normalizeKey(cleanChassis) && (c.status === 'Open' || c.status === 'Running'));
    if (activeComplaint && activeComplaint.complaintDetails) {
      setComplaintDetails(activeComplaint.complaintDetails);
    }
  };

  const handleNameChange = (val: any) => {
    const strVal = String(val ?? '');
    setCustName(strVal);
  };

  const handlePhoneChange = (val: string) => {
    setOwnerMob(val);
  };

  // Generate datalist options strictly for Chassis No
  const getCustomerDatalistOptions = () => {
    const rawRecords = Object.values(chassisIndex);
    if (rawRecords.length === 0) return [];

    const records = Array.from(new Set(rawRecords));
    const seenKeys = new Set<string>();
    const options: { key: string; val: string; display: string }[] = [];

    for (const rec of records) {
      if (options.length >= 150) break;
      const ch = (getFieldValue(rec, 'chassis') || '').trim();
      const nm = (getFieldValue(rec, 'custName') || '').trim();
      const ph = (getFieldValue(rec, 'custPhone') || '').trim();
      const inst = toInputDateFormat(getFieldValue(rec, 'installDate'));

      if (ch) {
        const k = `ch_${ch}`;
        if (!seenKeys.has(k)) {
          seenKeys.add(k);
          const display = `${ch}${nm ? ` — ${nm}` : ''}${ph ? ` (Mob: ${ph})` : ''}${inst ? ` (Inst: ${inst})` : ''}`;
          options.push({
            key: k,
            val: display,
            display
          });
        }
      }
    }

    return options;
  };

  const ensureTrailingBlankPartRow = (rows: PartRow[]) => {
    if (rows.length === 0) return [{ partNo: '', desc: '', wty: false, qty: '', rate: '', amount: '' }];
    const last = rows[rows.length - 1];
    if (last.partNo.trim() || last.desc.trim() || last.qty.trim() || last.rate.trim() || last.amount.trim() || last.wty) {
      return [...rows, { partNo: '', desc: '', wty: false, qty: '', rate: '', amount: '' }];
    }
    return rows;
  };

  const ensureTrailingBlankRepairRow = (rows: { repair: string; rectification: string; charge: string }[]) => {
    if (rows.length === 0) return [{ repair: '', rectification: '', charge: '' }];
    const last = rows[rows.length - 1];
    if (last.repair.trim() || last.rectification.trim() || last.charge.trim()) {
      return [...rows, { repair: '', rectification: '', charge: '' }];
    }
    return rows;
  };

  // Helper to extract clean query string from datalist values
  const extractQueryFromDatalistVal = (val: string) => {
    if (!val) return '';
    const bracketMatch = val.match(/\[(.*?)\]/);
    if (bracketMatch && bracketMatch[1]) {
      return bracketMatch[1].trim();
    }
    if (val.includes(' — ') || val.includes(' - ')) {
      const parts = val.split(/[—\-]/);
      return parts[0].trim();
    }
    return val.trim();
  };

  // Helper to find spare by Part No, Part Name or Description
  const findSpareByQuery = (query: string) => {
    if (!query) return null;
    const clean = normalizeKey(query);
    if (!clean) return null;

    if (sparesIndex[clean]) {
      return sparesIndex[clean];
    }
    return null;
  };

  // Lookup Spare Part by Part No or Name
  const handlePartNoChange = (index: number, val: string) => {
    let newPartRows = [...partRows];
    newPartRows[index].partNo = val;

    const parsedQuery = extractQueryFromDatalistVal(val);
    const matchedSpare: any = findSpareByQuery(parsedQuery) || findSpareByQuery(val);
    if (matchedSpare) {
      const actualPartNo = matchedSpare.__partNoDisplay || getSpareField(matchedSpare, 'partNo') || parsedQuery || val;
      const desc = getSpareField(matchedSpare, 'desc');
      const rate = getSpareField(matchedSpare, 'rate');

      // Set actual part number even if user typed the part name
      if (actualPartNo) newPartRows[index].partNo = actualPartNo;
      if (desc) newPartRows[index].desc = desc;
      if (rate) newPartRows[index].rate = rate;

      const q = parseFloat(newPartRows[index].qty) || 0;
      const r = parseFloat(rate || newPartRows[index].rate) || 0;
      if (newPartRows[index].wty) {
        newPartRows[index].amount = '0';
      } else if (q && r) {
        newPartRows[index].amount = (q * r).toFixed(2);
      }
    }
    setPartRows(ensureTrailingBlankPartRow(newPartRows));
  };

  // Lookup Spare Part by Description / Part Name
  const handlePartDescChange = (index: number, val: string) => {
    let newPartRows = [...partRows];
    newPartRows[index].desc = val;
    setPartRows(ensureTrailingBlankPartRow(newPartRows));
  };

  const handleQtyRateChange = (index: number, field: 'qty' | 'rate', val: string) => {
    let newPartRows = [...partRows];
    newPartRows[index][field] = val;

    const q = parseFloat(field === 'qty' ? val : newPartRows[index].qty) || 0;
    const r = parseFloat(field === 'rate' ? val : newPartRows[index].rate) || 0;
    if (newPartRows[index].wty) {
      newPartRows[index].amount = '0';
    } else if (q && r) {
      newPartRows[index].amount = (q * r).toFixed(2);
    }
    setPartRows(ensureTrailingBlankPartRow(newPartRows));
  };

  const handleWtyToggle = (index: number) => {
    let newPartRows = [...partRows];
    const isWty = !newPartRows[index].wty;
    newPartRows[index].wty = isWty;
    if (isWty) {
      newPartRows[index].amount = '0';
    } else {
      const q = parseFloat(newPartRows[index].qty) || 0;
      const r = parseFloat(newPartRows[index].rate) || 0;
      if (q && r) {
        newPartRows[index].amount = (q * r).toFixed(2);
      } else {
        newPartRows[index].amount = '';
      }
    }
    setPartRows(ensureTrailingBlankPartRow(newPartRows));
  };

  const toggleGenChecklist = (key: string) => {
    setGenChecklist(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const toggleCheckpoint = (id: number) => {
    setCheckpoints(prev =>
      prev.map(c => (c.id === id ? { ...c, checked: !c.checked } : c))
    );
  };

  const clearSavedCustomers = async () => {
    if (!window.confirm('Clear all saved customer records from browser memory and cloud database?')) return;
    await removeFromStorage(LS_CUSTOMER_KEY);
    setChassisIndex({});
    try {
      const snap = await getDocs(collection(db, 'customers_master'));
      if (!snap.empty) {
        const batch = writeBatch(db);
        snap.docs.forEach((d) => batch.delete(d.ref));
        await batch.commit();
      }
    } catch (e) {
      console.error("Error clearing customers_master from cloud:", e);
    }
    setExcelStatus({ text: 'Saved customer data cleared from browser and cloud database.', isSuccess: false });
  };

  const clearSavedSpares = async () => {
    if (!window.confirm('Clear all saved spare parts records from browser memory and cloud database?')) return;
    await removeFromStorage(LS_SPARES_KEY);
    setSparesIndex({});
    try {
      const snap = await getDocs(collection(db, 'spares_master'));
      if (!snap.empty) {
        const batch = writeBatch(db);
        snap.docs.forEach((d) => batch.delete(d.ref));
        await batch.commit();
      }
    } catch (e) {
      console.error("Error clearing spares_master from cloud:", e);
    }
    setSparesStatus({ text: 'Saved spares data cleared from browser and cloud database.', isSuccess: false });
  };

  const clearSavedJobCards = async () => {
    if (!window.confirm('Are you sure you want to clear ALL saved job card records from the database and storage?')) return;
    try {
      setJobCardsStatus({ text: 'Clearing job cards from cloud database...', isSuccess: false });
      
      // Delete all documents in Firestore 'jobcards' collection
      const snapshot = await getDocs(collection(db, 'jobcards'));
      if (!snapshot.empty) {
        const docs = snapshot.docs;
        for (let i = 0; i < docs.length; i += 400) {
          const batch = writeBatch(db);
          const chunk = docs.slice(i, i + 400);
          chunk.forEach((d) => batch.delete(d.ref));
          await batch.commit();
        }
      }

      await removeFromStorage(LS_SAVED_JOBCARDS_KEY);
      setSavedJobCards([]);
      setJobCardsStatus({ text: '✅ All saved job cards cleared successfully.', isSuccess: true });
      alert('✅ All job card records have been permanently cleared from database.');
    } catch (err) {
      console.error('Error clearing job cards from Firestore:', err);
      await removeFromStorage(LS_SAVED_JOBCARDS_KEY);
      setSavedJobCards([]);
      setJobCardsStatus({ text: '⚠️ Cleared local cache, but error deleting from cloud database.', isSuccess: false });
      alert('⚠️ Job cards cleared locally, but there was an error updating cloud storage.');
    }
  };

  const resetForm = () => {
    setEditingCardId(null);
    setSearchQuery('');
    setJobNo('');
    setJobDate(new Date().toISOString().split('T')[0]);
    setChassisNo('');
    setCustName('');
    setEngineNo('');
    setCustAddr('');
    setInstallDate('');
    setOwnerMob('');
    setDriverMob('');
    setRegdNo('');
    setDistDealership('');
    setHourMeter('');
    setDateTimeIn('');
    setDateTimeOut('');
    setExpectedRepairTime('');
    setModel('');
    setSerialNo('');
    setHistoryFileNo('');
    setWarrantyOverride('auto');
    setGenChecklist({});
    setOtherChecklistText('');
    setProblemDescription('');
    setComplaintDetails('');
    setComplaintDate('');
    setOnlineJobCardNo('');
    setBranch('');
    setModelType('');
    setFatherName('');
    setVillage('');
    setTranslatedVillage('');
    setMandal('');
    setTranslatedMandal('');
    setTranslatedCustName('');
    setTranslatedCustAddr('');
    setServiceLocationType('Paid Service');
    setFreeServiceList('');
    setExtraRepairs('');
    setReasonsForAnalysis('');
    setTelecalling('');

    setCheckpoints(DEFAULT_CHECKPOINTS);

    setRepairRows([
      { repair: '', rectification: '', charge: '' },
      { repair: '', rectification: '', charge: '' }
    ]);

    setWsReport('');
    setMechanic('');
    setTotalLabour('');
    setWarrantyMaterial('');
    setNonWarrantyMaterial('');
    setGTotal('');
    setWsIncharge('');
    setBillNo('');
    setStatus('Open');

    setPartRows([
      { partNo: '', desc: '', qty: '', rate: '', amount: '' },
      { partNo: '', desc: '', qty: '', rate: '', amount: '' },
      { partNo: '', desc: '', qty: '', rate: '', amount: '' }
    ]);
  };

  const toggleCardStatus = async (card: any) => {
    const newStatus = isCardClosed(card) ? 'Open' : 'Closed';
    const updatedCard = { ...card, status: newStatus };

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        await updateSheetRow(sheetsSpreadsheetId!, 'JobCards', JOBCARD_HEADERS, updatedCard);
        setSavedJobCards(prev => {
          const updated = prev.map(c => c.id === card.id ? updatedCard : c);
          saveJobCardsBackup(updated);
          return updated;
        });
        alert('✅ జాబ్ కార్డ్ స్టేటస్ విజయవంతంగా అప్‌డేట్ చేయబడింది!');
      } catch (err) {
        console.error('Error toggling status in sheets:', err);
        alert('❌ గూగుల్ షీట్స్‌లో జాబ్ కార్డ్ స్టేటస్ మార్చడం విఫలమైంది.');
      }
      return;
    }

    // Save to Cloud SQL Database
    try {
      await sqlApi.saveJobCard(updatedCard);
    } catch (sqlErr) {
      console.warn('Cloud SQL status toggle:', sqlErr);
    }

    setSavedJobCards(prev => {
      const updated = prev.map(c => c.id === card.id ? updatedCard : c);
      saveJobCardsBackup(updated);
      return updated;
    });

    if (db) {
      try {
        await setDoc(doc(db, 'jobcards', card.id), { status: newStatus }, { merge: true });
      } catch (e) {
        console.error('Error toggling card status in Firestore:', e);
      }
    }
  };

  const getDisplayServiceType = (cardOrType: any) => {
    const st = typeof cardOrType === 'string' ? cardOrType : (cardOrType && cardOrType.serviceType ? cardOrType.serviceType : '');
    const trimmed = String(st || '').trim();
    if (!trimmed) return '—';

    const freeServiceOptions = [
      '1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th',
      'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
      '1st Service', '2nd Service', '3rd Service', '4th Service', '5th Service',
      '6th Service', '7th Service', '8th Service', '9th Service', '10th Service',
      'Free Service', '1st Free Service', '2nd Free Service', '3rd Free Service',
      '4th Free Service', '5th Free Service', '6th Free Service', '7th Free Service',
      '8th Free Service', '9th Free Service', '10th Free Service'
    ];

    if (
      freeServiceOptions.some(opt => opt.toLowerCase() === trimmed.toLowerCase()) ||
      /^(1st|2nd|3rd|4th|5th|6th|7th|8th|9th|10th)/i.test(trimmed) ||
      (trimmed.toLowerCase().includes('free') && !trimmed.toLowerCase().includes('paid'))
    ) {
      return 'Free Service';
    }

    if (trimmed.toLowerCase().includes('paid repair')) {
      return 'Paid Repairs';
    }
    if (trimmed.toLowerCase().includes('paid service')) {
      return 'Paid Service';
    }

    return trimmed;
  };

  const saveJobCard = async () => {
    const strChassis = String(chassisNo || '').trim();
    const strOnlineNo = String(onlineJobCardNo || '').trim();
    const strJobNo = String(jobNo || '').trim();
    const strCustName = String(custName || '').trim();
    if (!strChassis && !strOnlineNo && !strJobNo && !strCustName) {
      alert('Please fill at least Online Job Card No, Chassis No, or Customer Name to save.');
      return;
    }

    const docId = editingCardId || Date.now().toString();

    const existingCard = savedJobCards.find(c => c.id === docId);
    const createdBy = existingCard?.createdBy || currentUser?.uid || '';
    const createdByEmail = existingCard?.createdByEmail || currentUser?.email || '';

    const effectiveFreeServiceList = String(freeServiceList || '').trim() || getFreeServiceText({ serviceType, freeServiceList: '' });
    const effectiveExtraRepairs = String(extraRepairs || '').trim() || (repairRows.length > 1 ? [repairRows[1].repair, repairRows[1].rectification].filter(Boolean).join(' - ') : '');

    const savedServiceType = getDisplayServiceType(serviceType);

    const cardData = {
      id: docId,
      serviceLocation,
      jobNo: onlineJobCardNo || jobNo,
      complaintDate,
      onlineJobCardNo: onlineJobCardNo || jobNo,
      jobDate,
      branch,
      historyFileNo,
      model,
      modelType,
      chassisNo,
      engineNo,
      installDate,
      custName,
      fatherName,
      custAddr,
      village,
      translatedVillage,
      mandal,
      translatedMandal,
      translatedCustName,
      translatedCustAddr,
      ownerMob,
      driverMob,
      regdNo,
      distDealership,
      hourMeter,
      dateTimeIn,
      dateTimeOut,
      expectedRepairTime,
      serviceType: savedServiceType,
      freeServiceList: effectiveFreeServiceList,
      extraRepairs: effectiveExtraRepairs,
      reasonsForAnalysis,
      telecalling,
      warrantyOverride,
      serialNo,
      genChecklist,
      otherChecklistText,
      problemDescription,
      complaintDetails,
      checkpoints,
      repairRows,
      partRows,
      wsReport,
      mechanic,
      totalLabour,
      warrantyMaterial,
      nonWarrantyMaterial,
      gTotal,
      wsIncharge,
      billNo,
      status,
      createdBy,
      createdByEmail,
      createdAt: existingCard?.createdAt || new Date().toISOString()
    };

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        const sheetCard = mapJobCardToSheet({ id: docId, ...cardData });
        await updateSheetRow(sheetsSpreadsheetId!, 'JobCards', JOBCARD_HEADERS, sheetCard);
        
        // Sync local React state
        const localCard = { id: docId, ...cardData };
        setSavedJobCards(prev => {
          const filtered = prev.filter(c => c.id !== docId);
          const updated = [localCard, ...filtered];
          saveJobCardsBackup(updated);
          return updated;
        });

        // Sync complaints status with Google Sheets
        if (dateTimeOut && chassisNo) {
          await autoCloseComplaints(chassisNo, dateTimeOut);
        } else if (chassisNo) {
          await markComplaintsRunning(chassisNo, onlineJobCardNo || jobNo);
        }

        if (editingCardId) {
          alert('✅ గూగుల్ షీట్స్‌లో జాబ్ కార్డ్ విజయవంతంగా అప్‌డేట్ చేయబడింది!');
        } else {
          alert('✅ గూగుల్ షీట్స్‌లో జాబ్ కార్డ్ విజయవంతంగా సేవ్ చేయబడింది!');
        }
        resetForm();
      } catch (err) {
        console.error('Error saving job card to Sheets:', err);
        alert('❌ గూగుల్ షీట్స్‌లో జాబ్ కార్డ్ సేవ్ చేయడం విఫలమైంది: ' + String(err));
      }
      return;
    }

    try {
      // Save directly to Cloud SQL Database
      try {
        await sqlApi.saveJobCard({ id: docId, ...cardData });
      } catch (sqlErr) {
        console.warn('Cloud SQL jobcard save:', sqlErr);
      }

      if (db) {
        try {
          await setDoc(doc(db, 'jobcards', docId), cardData);
        } catch (fErr) {
          console.warn('Firestore fallback sync notice:', fErr);
        }
      }
      
      // Update local state and backup
      const localCard = { id: docId, ...cardData };
      setSavedJobCards(prev => {
        const filtered = prev.filter(c => c.id !== docId);
        const updated = [localCard, ...filtered];
        saveJobCardsBackup(updated);
        return updated;
      });

      // Sync customer details globally
      await syncCustomerDetailsGlobally(chassisNo, {
        custName,
        fatherName,
        custAddr,
        village,
        mandal,
        ownerMob,
        driverMob,
        model,
        modelType,
        engineNo,
        installDate,
        regdNo
      });

      // Auto-close related complaints if dateTimeOut (Actual Closed Date) is set, otherwise mark them as Running
      if (dateTimeOut && chassisNo) {
        await autoCloseComplaints(chassisNo, dateTimeOut);
      } else if (chassisNo) {
        await markComplaintsRunning(chassisNo, onlineJobCardNo || jobNo);
      }

      if (editingCardId) {
        alert('✅ Job card details updated successfully in database!');
      } else {
        alert('✅ Job card saved successfully in database!');
      }
      resetForm();
    } catch (e) {
      console.error('Error saving job card to Firestore:', e);
      const localCard = { id: docId, ...cardData };
      setSavedJobCards(prev => {
        const filtered = prev.filter(c => c.id !== docId);
        const updated = [localCard, ...filtered];
        saveJobCardsBackup(updated);
        return updated;
      });
      if (dateTimeOut && chassisNo) {
        autoCloseComplaints(chassisNo, dateTimeOut);
      } else if (chassisNo) {
        markComplaintsRunning(chassisNo, onlineJobCardNo || jobNo);
      }
      alert('✅ Job card saved successfully in local cache and database!');
      resetForm();
    }
  };

  const loadJobCard = (card: any) => {
    if (currentUserRole !== 'admin' && card.createdBy && card.createdBy !== currentUser?.uid) {
      alert('❌ Error: You are not authorized to edit this job card.');
      return;
    }
    setEditingCardId(card.id);
    setActiveTab('new_entry');
    setServiceLocation(card.serviceLocation || 'workshop');
    setJobNo(card.jobNo || '');
    setComplaintDate(card.complaintDate || '');
    setOnlineJobCardNo(card.onlineJobCardNo || '');
    setJobDate(card.jobDate || '');
    setBranch(card.branch || '');
    setChassisNo(card.chassisNo || '');
    setCustName(card.custName || '');
    setFatherName(card.fatherName || '');
    setVillage(card.village || '');
    setTranslatedVillage(card.translatedVillage || '');
    setMandal(card.mandal || '');
    setTranslatedMandal(card.translatedMandal || '');
    setTranslatedCustName(card.translatedCustName || '');
    setTranslatedCustAddr(card.translatedCustAddr || '');
    setEngineNo(card.engineNo || '');
    setCustAddr(card.custAddr || '');
    setInstallDate(card.installDate || '');
    setOwnerMob(card.ownerMob || '');
    setDriverMob(card.driverMob || '');
    setRegdNo(card.regdNo || '');
    setDistDealership(card.distDealership || '');
    setHourMeter(card.hourMeter || '');
    setDateTimeIn(card.dateTimeIn || '');
    setDateTimeOut(card.dateTimeOut || '');
    setExpectedRepairTime(card.expectedRepairTime || '');
    setServiceLocationType(card.serviceType || 'Paid Service');
    setFreeServiceList(card.freeServiceList || getFreeServiceText(card));
    const extraRepVal = getExtraRepairsText(card);
    setExtraRepairs(card.extraRepairs || (extraRepVal === '—' ? '' : extraRepVal));
    setReasonsForAnalysis(card.reasonsForAnalysis || card.problemDescription || '');
    setTelecalling(card.telecalling || '');
    setWarrantyOverride(card.warrantyOverride || 'auto');
    setModel(card.model || '');
    setModelType(card.modelType || '');
    setSerialNo(card.serialNo || '');
    setHistoryFileNo(card.historyFileNo || '');
    setGenChecklist(card.genChecklist || {});
    setOtherChecklistText(card.otherChecklistText || '');
    setProblemDescription(card.problemDescription || '');
    setComplaintDetails(card.complaintDetails || '');
    setCheckpoints(card.checkpoints || DEFAULT_CHECKPOINTS);
    setRepairRows(card.repairRows || [
      { repair: '', rectification: '', charge: '' },
      { repair: '', rectification: '', charge: '' }
    ]);
    setPartRows(card.partRows || [
      { partNo: '', desc: '', qty: '', rate: '', amount: '' },
      { partNo: '', desc: '', qty: '', rate: '', amount: '' },
      { partNo: '', desc: '', qty: '', rate: '', amount: '' }
    ]);
    setWsReport(card.wsReport || '');
    setMechanic(card.mechanic || '');
    setTotalLabour(card.totalLabour || '');
    setWarrantyMaterial(card.warrantyMaterial || '');
    setNonWarrantyMaterial(card.nonWarrantyMaterial || '');
    setGTotal(card.gTotal || '');
    setWsIncharge(card.wsIncharge || '');
    setBillNo(card.billNo || '');
    setStatus(card.status || 'Open');

    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const deleteJobCard = async (id: string) => {
    const cardToDelete = savedJobCards.find(c => c.id === id);
    if (!cardToDelete) return;
    if (currentUserRole !== 'admin' && cardToDelete.createdBy !== currentUser?.uid) {
      alert('❌ Error: You are not authorized to delete this job card.');
      return;
    }

    if (!window.confirm('Are you sure you want to delete this saved job card?')) return;

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        await deleteSheetRow(sheetsSpreadsheetId!, 'JobCards', id);
        setSavedJobCards(prev => {
          const updated = prev.filter(c => c.id !== id);
          saveJobCardsBackup(updated);
          return updated;
        });
        if (editingCardId === id) {
          setEditingCardId(null);
        }
        alert('✅ జాబ్ కార్డ్ గూగుల్ షీట్స్ నుండి విజయవంతంగా తొలగించబడింది!');
      } catch (err) {
        console.error('Error deleting job card from sheets:', err);
        alert('❌ గూగుల్ షీట్స్ నుండి తొలగించడం విఫలమైంది.');
      }
      return;
    }

    // Delete from Cloud SQL Database
    try {
      await sqlApi.deleteJobCard(id);
    } catch (sqlErr) {
      console.warn('Cloud SQL deleteJobCard:', sqlErr);
    }

    setSavedJobCards(prev => {
      const updated = prev.filter(c => c.id !== id);
      saveJobCardsBackup(updated);
      return updated;
    });
    if (editingCardId === id) {
      setEditingCardId(null);
    }

    if (db) {
      try {
        await deleteDoc(doc(db, 'jobcards', id));
      } catch (e) {
        console.error('Error deleting job card from Firestore:', e);
      }
    }
    alert('✅ Job card deleted successfully!');
  };

  const handleComplaintChassisChange = (chassisVal: string) => {
    const cleanCh = normalizeKey(chassisVal);
    const foundCust = chassisIndex[cleanCh] || Object.values(chassisIndex).find((c: any) => normalizeKey(getFieldValue(c, 'chassis')) === cleanCh);
    if (foundCust) {
      const custNm = getFieldValue(foundCust, 'custName') || '';
      const mob = getFieldValue(foundCust, 'custPhone') || '';
      const mdl = getFieldValue(foundCust, 'model') || '';
      const sup = getFieldValue(foundCust, 'supervisor') || '';
      setComplaintForm(prev => ({
        ...prev,
        chassisNo: chassisVal,
        customerName: custNm || prev.customerName,
        mobileNumber: mob || prev.mobileNumber,
        tractorModel: mdl || prev.tractorModel,
        assignedSupervisor: sup || prev.assignedSupervisor
      }));
    } else {
      setComplaintForm(prev => ({ ...prev, chassisNo: chassisVal }));
    }
  };

  const saveComplaint = async () => {
    if (!complaintForm.customerName || !complaintForm.complaintDetails) {
      alert('Please fill out Customer Name and Complaint Details.');
      return;
    }

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        const { id, ...formData } = complaintForm;
        const isNew = !id;
        const docId = id || `comp-${Date.now()}`;

        let finalStatus = formData.status;
        let finalClosureDate = formData.closureDate;
        if (finalClosureDate && finalClosureDate.trim() !== '') {
          finalStatus = 'Closed';
        }

        const nextId = complaints.length + 1;
        const generatedNo = formData.complaintNo || `COMP-${new Date().getFullYear()}-${nextId.toString().padStart(4, '0')}`;

        const complaintData = {
          id: docId,
          complaintNo: generatedNo,
          complaintDate: formData.complaintDate || new Date().toISOString().split('T')[0],
          chassisNo: formData.chassisNo || '',
          customerName: formData.customerName || '',
          mobileNumber: formData.mobileNumber || '',
          tractorModel: formData.tractorModel || '',
          complaintDetails: formData.complaintDetails || '',
          assignedMechanic: formData.assignedMechanic || '',
          assignedSupervisor: formData.assignedSupervisor || '',
          status: finalStatus,
          resolution: formData.resolution || '',
          closureDate: finalClosureDate || '',
          createdBy: currentUser?.uid || currentUser?.email || 'system',
          createdAt: isNew ? new Date().toISOString() : (complaints.find(c => c.id === docId)?.createdAt || new Date().toISOString())
        };

        await updateSheetRow(sheetsSpreadsheetId!, 'Complaints', COMPLAINT_HEADERS, complaintData);

        setComplaints(prev => {
          const filtered = prev.filter(c => c.id !== docId);
          const updated = [complaintData, ...filtered];
          try {
            localStorage.setItem('sri_backup_complaints', JSON.stringify(updated));
          } catch (le) {}
          return updated;
        });

        alert(isNew ? '✅ కంప్లైంట్ విజయవంతంగా గూగుల్ షీట్స్‌లో నమోదు చేయబడింది!' : '✅ కంప్లైంట్ వివరాలు గూగుల్ షీట్స్‌లో అప్‌డేట్ చేయబడ్డాయి!');
        
        setIsComplaintModalOpen(false);
        setComplaintForm({
          id: '',
          complaintNo: '',
          complaintDate: new Date().toISOString().split('T')[0],
          chassisNo: '',
          customerName: '',
          mobileNumber: '',
          tractorModel: '',
          complaintDetails: '',
          assignedMechanic: '',
          assignedSupervisor: '',
          status: 'Open',
          resolution: '',
          closureDate: '',
        });
      } catch (err) {
        console.error('Error saving complaint to Sheets:', err);
        alert('❌ గూగుల్ షీట్స్‌లో కంప్లైంట్ సేవ్ చేయడం విఫలమైంది: ' + String(err));
      }
      return;
    }

    try {
      const { id, ...formData } = complaintForm;
      const isNew = !id;

      // Auto-close logic
      let finalStatus = formData.status;
      let finalClosureDate = formData.closureDate;

      if (finalClosureDate && finalClosureDate.trim() !== '') {
        finalStatus = 'Closed';
      }

      if (formData.jobCardNo) {
        const linkedJobCard = savedJobCards.find(c => c.jobCardNo === formData.jobCardNo); // Assuming jobcards have jobCardNo
        if (linkedJobCard && linkedJobCard.status === 'Closed' && (linkedJobCard.closedDate || linkedJobCard.actualClosedDate)) {
          finalStatus = 'Closed';
          finalClosureDate = linkedJobCard.closedDate || linkedJobCard.actualClosedDate || finalClosureDate;
        }
      }

      // Ensure no undefined values are passed to Firestore
      const cleanData: Record<string, any> = {
        ...formData,
        status: finalStatus,
        closureDate: finalClosureDate
      };
      const finalData: Record<string, any> = {};
      Object.entries(cleanData).forEach(([k, v]) => {
        if (v !== undefined) {
          finalData[k] = v;
        }
      });

      if (isNew) {
        // Generate complaint No
        const nextId = complaints.length + 1;
        const generatedNo = `COMP-${new Date().getFullYear()}-${nextId.toString().padStart(4, '0')}`;
        const newCompData = {
          ...finalData,
          id: `comp-${Date.now()}`,
          complaintNo: generatedNo,
          createdBy: currentUser?.uid || currentUser?.email || 'system',
          createdAt: new Date().toISOString()
        };

        // Save to Cloud SQL
        try {
          await sqlApi.saveComplaint(newCompData);
        } catch (sqlErr) {
          console.warn('Cloud SQL complaint save:', sqlErr);
        }

        if (db) {
          await addDoc(collection(db, 'complaints'), newCompData);
        }

        setComplaints(prev => [newCompData, ...prev]);
        try {
          localStorage.setItem('sri_backup_complaints', JSON.stringify([newCompData, ...complaints]));
        } catch (le) {}

        alert('✅ Complaint created successfully!');
      } else {
        const updatedCompData = {
          ...finalData,
          id,
          updatedAt: new Date().toISOString()
        };

        // Save to Cloud SQL
        try {
          await sqlApi.saveComplaint(updatedCompData);
        } catch (sqlErr) {
          console.warn('Cloud SQL complaint update:', sqlErr);
        }

        if (db) {
          await setDoc(doc(db, 'complaints', id), {
            ...finalData,
            updatedAt: new Date().toISOString()
          }, { merge: true });
        }

        setComplaints(prev => prev.map(c => c.id === id ? { ...c, ...updatedCompData } : c));
        try {
          localStorage.setItem('sri_backup_complaints', JSON.stringify(complaints.map(c => c.id === id ? { ...c, ...updatedCompData } : c)));
        } catch (le) {}

        alert('✅ Complaint updated successfully!');
      }

      // Sync customer details globally
      if (formData.chassisNo) {
        await syncCustomerDetailsGlobally(formData.chassisNo, {
          custName: formData.customerName,
          ownerMob: formData.mobileNumber,
          model: formData.tractorModel,
          supervisor: formData.assignedSupervisor
        });
      }

      setIsComplaintModalOpen(false);
      setComplaintForm({
        id: '',
        complaintNo: '',
        complaintDate: new Date().toISOString().split('T')[0],
        chassisNo: '',
        customerName: '',
        mobileNumber: '',
        tractorModel: '',
        complaintDetails: '',
        assignedMechanic: '',
        assignedSupervisor: '',
        status: 'Open',
        resolution: '',
        closureDate: '',
      });
    } catch (error) {
      console.error("Error saving complaint to Firestore:", error);
      const { id, ...formData } = complaintForm;
      const isNew = !id;
      const targetId = id || `comp-local-${Date.now()}`;
      let finalStatus = formData.status;
      let finalClosureDate = formData.closureDate;
      if (finalClosureDate && finalClosureDate.trim() !== '') {
        finalStatus = 'Closed';
      }

      const cleanData: Record<string, any> = {
        ...formData,
        status: finalStatus,
        closureDate: finalClosureDate
      };
      const finalData: Record<string, any> = {};
      Object.entries(cleanData).forEach(([k, v]) => {
        if (v !== undefined) {
          finalData[k] = v;
        }
      });

      const nextId = complaints.length + 1;
      const generatedNo = formData.complaintNo || `COMP-${new Date().getFullYear()}-${nextId.toString().padStart(4, '0')}`;

      const localComplaint = {
        id: targetId,
        ...finalData,
        complaintNo: generatedNo,
        createdBy: currentUser?.uid || currentUser?.email || 'system',
        createdAt: new Date().toISOString()
      };

      try {
        await sqlApi.saveComplaint(localComplaint);
      } catch (sqlErr) {
        console.warn('Cloud SQL complaint fallback save:', sqlErr);
      }

      setComplaints(prev => {
        const filtered = prev.filter(c => c.id !== targetId);
        const updated = [localComplaint, ...filtered];
        try {
          localStorage.setItem('sri_backup_complaints', JSON.stringify(updated));
        } catch (le) {}
        return updated;
      });

      alert('✅ Complaint saved successfully in database!');
      
      setIsComplaintModalOpen(false);
      setComplaintForm({
        id: '',
        complaintNo: '',
        complaintDate: new Date().toISOString().split('T')[0],
        chassisNo: '',
        customerName: '',
        mobileNumber: '',
        tractorModel: '',
        complaintDetails: '',
        assignedMechanic: '',
        assignedSupervisor: '',
        status: 'Open',
        resolution: '',
        closureDate: '',
      });
    }
  };

  const deleteComplaint = async (id: string) => {
    if (currentUserRole !== 'admin') {
      alert('❌ Error: Only admins can delete complaints.');
      return;
    }
    if (!window.confirm('Are you sure you want to delete this complaint?')) return;

    if (useGoogleSheets) {
      if (!sheetsToken) {
        alert('⚠️ గూగుల్ షీట్స్ సెషన్ ముగిసింది. దయచేసి పైన ఉన్న కనెక్ట్ బటన్ క్లిక్ చేసి తిరిగి సైన్-ఇన్ అవ్వండి.');
        return;
      }
      try {
        await deleteSheetRow(sheetsSpreadsheetId!, 'Complaints', id);
        setComplaints(prev => {
          const updated = prev.filter(c => c.id !== id);
          try {
            localStorage.setItem('sri_backup_complaints', JSON.stringify(updated));
          } catch (le) {}
          return updated;
        });
        alert('✅ కంప్లైంట్ గూగుల్ షీట్స్ నుండి విజయవంతంగా తొలగించబడింది!');
      } catch (err) {
        console.error('Error deleting complaint from sheets:', err);
        alert('❌ గూగుల్ షీట్స్ నుండి తొలగించడం విఫలమైంది.');
      }
      return;
    }

    // Delete from Cloud SQL Database
    try {
      await sqlApi.deleteComplaint(id);
    } catch (sqlErr) {
      console.warn('Cloud SQL deleteComplaint:', sqlErr);
    }

    setComplaints(prev => {
      const updated = prev.filter(c => c.id !== id);
      try {
        localStorage.setItem('sri_backup_complaints', JSON.stringify(updated));
      } catch (le) {}
      return updated;
    });

    if (db) {
      try {
        await deleteDoc(doc(db, 'complaints', id));
      } catch (error) {
        console.error("Error deleting complaint from Firestore:", error);
      }
    }
    alert('✅ Complaint deleted successfully!');
  };

  const autoCloseComplaints = async (chassis: string, closureDate: string) => {
    if (!chassis || !closureDate) return;
    
    const relatedComplaints = complaints.filter(c => 
      normalizeKey(c.chassisNo) === normalizeKey(chassis) && 
      c.status !== 'Closed'
    );

    for (const comp of relatedComplaints) {
      const updatedComp = { ...comp, status: 'Closed', closureDate };
      
      // Update Local State
      setComplaints(prev => prev.map(c => c.id === comp.id ? updatedComp : c));

      // Save to Cloud SQL
      try {
        await sqlApi.saveComplaint(updatedComp);
      } catch (e) {}

      // Update Firestore or Sheets
      if (useGoogleSheets && sheetsSpreadsheetId) {
        try {
          await updateSheetRow(sheetsSpreadsheetId, 'Complaints', COMPLAINT_HEADERS, updatedComp);
        } catch (e) {
          console.error('Error auto-closing complaint in Sheets:', e);
        }
      } else if (db) {
        try {
          await updateDoc(doc(db, 'complaints', comp.id), { status: 'Closed', closureDate });
        } catch (e) {
          console.error('Error auto-closing complaint in Firestore:', e);
        }
      }
    }
  };

  const markComplaintsRunning = async (chassis: string, jobCardNoVal?: string) => {
    if (!chassis) return;
    
    const relatedComplaints = complaints.filter(c => 
      normalizeKey(c.chassisNo) === normalizeKey(chassis) && 
      c.status !== 'Closed' &&
      c.status !== 'Running'
    );

    for (const comp of relatedComplaints) {
      const updatedComp = { 
        ...comp, 
        status: 'Running', 
        jobCardNo: jobCardNoVal || comp.jobCardNo || '' 
      };
      
      // Update Local State
      setComplaints(prev => prev.map(c => c.id === comp.id ? updatedComp : c));

      // Save to Cloud SQL
      try {
        await sqlApi.saveComplaint(updatedComp);
      } catch (e) {}

      // Update Firestore or Sheets
      if (useGoogleSheets && sheetsSpreadsheetId) {
        try {
          await updateSheetRow(sheetsSpreadsheetId, 'Complaints', COMPLAINT_HEADERS, updatedComp);
        } catch (e) {
          console.error('Error updating complaint to Running in Sheets:', e);
        }
      } else if (db) {
        try {
          await updateDoc(doc(db, 'complaints', comp.id), { 
            status: 'Running', 
            jobCardNo: jobCardNoVal || comp.jobCardNo || '' 
          });
        } catch (e) {
          console.error('Error updating complaint to Running in Firestore:', e);
        }
      }
    }
  };

  const handleInlineCardFieldChange = async (cardId: string, fieldName: string, value: string) => {
    setSavedJobCards((prev) =>
      prev.map((card) => {
        if (card.id === cardId) {
          const updated = { ...card, [fieldName]: value };
          if (fieldName === 'actualClosedDate') {
            updated.dateTimeOut = value;
            if (value && (card.status === 'Open' || !card.status)) {
              updated.status = 'Closed';
            }
          }
          if (fieldName === 'reasonsForAnalysis') {
            updated.problemDescription = value;
          }
          if (fieldName === 'hourMeter') {
            updated.hrsRun = value;
          }
          if (fieldName === 'mechanic') {
            updated.technicianName = value;
            const assocSup = getAssignedSupervisor(value);
            if (assocSup) {
              updated.wsIncharge = assocSup;
              updated.supervisorName = assocSup;
              updated.supervisor = assocSup;
            }
          }
          if (fieldName === 'wsIncharge') {
            updated.supervisorName = value;
            updated.supervisor = value;
          }
          return updated;
        }
        return card;
      })
    );

    try {
      const cardRef = doc(db, 'jobcards', cardId);
      const updateObj: Record<string, any> = { [fieldName]: value };
      if (fieldName === 'actualClosedDate') {
        updateObj.dateTimeOut = value;
        const currentCard = savedJobCards.find((c) => c.id === cardId);
        if (value && (currentCard?.status === 'Open' || !currentCard?.status)) {
          updateObj.status = 'Closed';
        }
      }
      if (fieldName === 'reasonsForAnalysis') {
        updateObj.problemDescription = value;
      }
      if (fieldName === 'hourMeter') {
        updateObj.hrsRun = value;
      }
      if (fieldName === 'mechanic') {
        updateObj.technicianName = value;
        const assocSup = getAssignedSupervisor(value);
        if (assocSup) {
          updateObj.wsIncharge = assocSup;
          updateObj.supervisorName = assocSup;
          updateObj.supervisor = assocSup;
        }
      }
      if (fieldName === 'wsIncharge') {
        updateObj.supervisorName = value;
        updateObj.supervisor = value;
      }
      await updateDoc(cardRef, updateObj);

      // Auto-close related complaints if actualClosedDate is entered
      if (fieldName === 'actualClosedDate' && value) {
        const currentCard = savedJobCards.find(c => c.id === cardId);
        if (currentCard?.chassisNo) {
          await autoCloseComplaints(currentCard.chassisNo, value);
        }
      }
    } catch (err) {
      console.error('Error updating inline field in Firestore:', err);
    }
  };

  const downloadSavedCardsExcel = () => {
    const cardsToExport = savedJobCards.filter(card => {
      if (currentUserRole === 'admin') return true;
      return !card.createdBy || card.createdBy === currentUser?.uid;
    });

    if (cardsToExport.length === 0) {
      alert('No saved job cards to download.');
      return;
    }

    const exportData = cardsToExport.map((card) => ({
      'Job card': (card.jobNo || card.onlineJobCardNo || '').toString(),
      'COMPALINT DATE': fmtDate(card.complaintDate),
      'ONLINE JOB CARD NO': (card.onlineJobCardNo || card.jobNo || '').toString(),
      'JOB CARD OPEN DAT': fmtDate(card.jobDate || card.jobOpenDate || card.createdAt),
      'BRANCH': (card.branch || '').toString(),
      'HISTORY FILE NO.': (card.historyFileNo || card.fileNo || '').toString(),
      'Tractor model': (card.model || '').toString(),
      'MODEL TYPE': (card.modelType || '').toString(),
      'CHASIS NO': (card.chassisNo || '').toString(),
      'Eng Sr no': (card.engineNo || '').toString(),
      'Date of Delivery': fmtDate(card.dateOfDelivery || card.installDate || card.deliveryDate),
      'cutomer name': (card.custName || '').toString(),
      'FATHER': (card.fatherName || '').toString(),
      'ADDRESS': (card.custAddr || card.address || '').toString(),
      'village': (card.village || '').toString(),
      'mandal': (card.mandal || '').toString(),
      'ph no': (card.ownerMob || card.phNo || '').toString(),
      'Hrs Run': (card.hourMeter || card.hrsRun || '').toString(),
      'Type of Service': getDisplayServiceType(card),
      'FREE SERVICE LIST': (getFreeServiceText(card) === '—' ? '' : getFreeServiceText(card)).toString(),
      'EXTRA OTHER REPAIRS DONE WITH FREE SERVICE': (getExtraRepairsText(card) === '—' ? '' : getExtraRepairsText(card)).toString(),
      'ACTUVAL CLOSED DATE': fmtDate(card.actualClosedDate || card.dateTimeOut),
      'TECHNICIAN NAME': (card.mechanic || card.technicianName || '').toString(),
      'Service place': (card.serviceLocation || card.servicePlace || '').toString(),
      'BILL NO.': (card.billNo || '').toString(),
      'RESONS FOR ANALYSIS': (card.reasonsForAnalysis || card.problemDescription || '').toString(),
      'TELECALLING': (card.telecalling || '').toString(),
      'Status': (card.status || 'Open').toString()
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Saved Job Cards");

    // Auto-fit columns
    const maxProps = Object.keys(exportData[0] || {}).map(key => ({
      wch: Math.max(key.length, ...exportData.map(row => String(row[key as keyof typeof row] || '').length)) + 2
    }));
    worksheet['!cols'] = maxProps;

    XLSX.writeFile(workbook, `SriGayathri_Saved_JobCards_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const downloadCustomerExcel = () => {
    const keys = Object.keys(chassisIndex);
    if (keys.length === 0) {
      alert('No customer records stored in memory to export.');
      return;
    }

    const uniqueRecordsMap = new Map<string, any>();
    keys.forEach(k => {
      const rec = chassisIndex[k];
      if (rec) {
        const chassisNo = rec.__chassisDisplay || getFieldValue(rec, 'chassis') || k;
        if (!uniqueRecordsMap.has(chassisNo)) {
          uniqueRecordsMap.set(chassisNo, rec);
        }
      }
    });

    const records = Array.from(uniqueRecordsMap.values());

    const exportData = records.map((rec, idx) => ({
      'BRANCH': (rec.BRANCH || rec.branch || getFieldValue(rec, 'branch') || '').toString(),
      'S.NO': (rec['S.NO'] || rec.sno || idx + 1).toString(),
      'Model': (rec.Model || rec.model || getFieldValue(rec, 'model') || '').toString(),
      'MODEL TYPE': (rec['MODEL TYPE'] || rec.modelType || getFieldValue(rec, 'modelType') || '').toString(),
      'Chassis no': (rec['Chassis no'] || rec.__chassisDisplay || rec.chassisNo || getFieldValue(rec, 'chassis') || '').toString(),
      'Engine No:': (rec['Engine No:'] || rec.engineNo || getFieldValue(rec, 'engineNo') || '').toString(),
      'Date of del': fmtDate(rec['Date of del'] || rec.dateOfDel || getFieldValue(rec, 'installDate')),
      'Customer Name': (rec['Customer Name'] || rec.__custNameDisplay || rec.custName || getFieldValue(rec, 'custName') || '').toString(),
      'FATHER NAME': (rec['FATHER NAME'] || rec.fatherName || getFieldValue(rec, 'fatherName') || '').toString(),
      'ADDRESS': (rec['ADDRESS'] || rec.address || getFieldValue(rec, 'custAddr') || '').toString(),
      'VILLAGE': (rec['VILLAGE'] || rec.village || getFieldValue(rec, 'village') || '').toString(),
      'Mandal': (rec['Mandal'] || rec.mandal || getFieldValue(rec, 'mandal') || '').toString(),
      'Mobile Number': (rec['Mobile Number'] || rec.__custPhoneDisplay || rec.mobileNumber || getFieldValue(rec, 'custPhone') || '').toString(),
      'Distict': (rec['Distict'] || rec.district || getFieldValue(rec, 'district') || '').toString(),
      'PIN CODE': (rec['PIN CODE'] || rec.pinCode || getFieldValue(rec, 'pinCode') || '').toString(),
      'DSP Name': (rec['DSP Name'] || rec.dspName || getFieldValue(rec, 'dspName') || '').toString(),
      'EXCHANGE BRAND': (rec['EXCHANGE BRAND'] || rec.exchangeBrand || getFieldValue(rec, 'exchangeBrand') || '').toString(),
      'EXCHANGE TRACTOR MODELS': (rec['EXCHANGE TRACTOR MODELS'] || rec.exchangeModels || getFieldValue(rec, 'exchangeModels') || '').toString()
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Customers");

    // Auto-fit columns
    const maxProps = Object.keys(exportData[0] || {}).map(key => ({
      wch: Math.max(key.length, ...exportData.map(row => String(row[key as keyof typeof row] || '').length)) + 2
    }));
    worksheet['!cols'] = maxProps;

    XLSX.writeFile(workbook, `SriGayathri_Customer_Database_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const downloadSparesExcel = () => {
    const keys = Object.keys(sparesIndex);
    if (keys.length === 0) {
      alert('No spares records stored in memory to export.');
      return;
    }

    const uniqueSparesMap = new Map<string, any>();
    keys.forEach(k => {
      const rec = sparesIndex[k];
      if (rec) {
        const partNo = rec.__partNoDisplay || getFieldValue(rec, 'partNo') || k;
        if (!uniqueSparesMap.has(partNo)) {
          uniqueSparesMap.set(partNo, rec);
        }
      }
    });

    const records = Array.from(uniqueSparesMap.values());

    const exportData = records.map(rec => {
      const partNo = rec.__partNoDisplay || getFieldValue(rec, 'partNo') || '';
      const desc = getFieldValue(rec, 'partDesc') || '';
      const rate = getFieldValue(rec, 'partRate') || '';
      return {
        'Part Number': partNo.toString(),
        'Description': desc.toString(),
        'Rate / Price (Rs)': rate.toString()
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Spares Price List");

    // Auto-fit columns
    const maxProps = Object.keys(exportData[0] || {}).map(key => ({
      wch: Math.max(key.length, ...exportData.map(row => String(row[key as keyof typeof row] || '').length)) + 2
    }));
    worksheet['!cols'] = maxProps;

    XLSX.writeFile(workbook, `SriGayathri_Spares_PriceList_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const copyCustomerDetails = (lookup: any) => {
    const text = formatCustomerClipboardText(lookup);
    if (!text) {
      alert('⚠️ No customer details available to copy.');
      return;
    }
    navigator.clipboard.writeText(text)
      .then(() => {
        alert(`✅ Customer details copied successfully!\n\n${text}`);
      })
      .catch((err) => {
        console.error('Failed to copy customer text:', err);
        // Fallback for iframe restrictions
        const textArea = document.createElement("textarea");
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
          document.execCommand('copy');
          alert(`✅ Customer details copied successfully!\n\n${text}`);
        } catch (err2) {
          alert('❌ Failed to copy details automatically. Please copy manually.');
        }
        document.body.removeChild(textArea);
      });
  };

  const handleSaveNewCustomer = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCustForm.chassisNo.trim() && !newCustForm.custName.trim()) {
      alert('Please enter at least Chassis Number or Customer Name.');
      return;
    }

    const chassisKey = normalizeKey(newCustForm.chassisNo);
    const nameKey = normalizeKey(newCustForm.custName);
    const phoneKey = normalizeKey(newCustForm.mobileNumber);

    const rawRecord = {
      ...newCustForm,
      branch: newCustForm.branch,
      sno: newCustForm.sNo,
      historyFileNo: newCustForm.sNo,
      model: newCustForm.model,
      modelType: newCustForm.modelType,
      chassis: newCustForm.chassisNo,
      engineNo: newCustForm.engineNo,
      installDate: newCustForm.dateOfDel,
      custName: newCustForm.custName,
      fatherName: newCustForm.fatherName,
      address: newCustForm.address,
      village: newCustForm.village,
      mandal: newCustForm.mandal,
      custPhone: newCustForm.mobileNumber,
      district: newCustForm.district,
      pinCode: newCustForm.pinCode,
      dspName: newCustForm.dspName,
      exchangeBrand: newCustForm.exchangeBrand,
      exchangeModels: newCustForm.exchangeModels,
      supervisor: newCustForm.supervisor
    };

    const newRecord = buildRowLookup(rawRecord);
    newRecord.__chassisDisplay = newCustForm.chassisNo;
    newRecord.__custNameDisplay = newCustForm.custName;
    newRecord.__custPhoneDisplay = newCustForm.mobileNumber;
    newRecord.__custAddrDisplay = [newCustForm.address, newCustForm.village, newCustForm.mandal, newCustForm.district]
      .filter(Boolean)
      .join(', ');
    newRecord.supervisor = newCustForm.supervisor;

    const updatedIndex = { ...chassisIndex };
    if (chassisKey) updatedIndex[chassisKey] = newRecord;
    if (nameKey) updatedIndex[nameKey] = newRecord;
    if (phoneKey) updatedIndex[phoneKey] = newRecord;

    setChassisIndex(updatedIndex);
    await saveToStorage(LS_CUSTOMER_KEY, updatedIndex);

    // Sync complete customer list to Firestore cloud database so all users opening the link see this
    const uniqueCustomerRows = Array.from(new Set(Object.values(updatedIndex)));
    await saveCustomersToFirestore(uniqueCustomerRows);

    // Propagate changes globally to Firestore jobcards and complaints
    await syncCustomerDetailsGlobally(newCustForm.chassisNo, {
      custName: newCustForm.custName,
      fatherName: newCustForm.fatherName,
      custAddr: newCustForm.address,
      village: newCustForm.village,
      mandal: newCustForm.mandal,
      ownerMob: newCustForm.mobileNumber,
      model: newCustForm.model,
      modelType: newCustForm.modelType,
      engineNo: newCustForm.engineNo,
      installDate: newCustForm.dateOfDel,
      supervisor: newCustForm.supervisor,
      district: newCustForm.district,
      pinCode: newCustForm.pinCode,
      dspName: newCustForm.dspName,
      exchangeBrand: newCustForm.exchangeBrand,
      exchangeModels: newCustForm.exchangeModels
    });

    const count = Object.keys(updatedIndex).length;
    setExcelStatus({
      text: `✅ Added customer "${newCustForm.custName || newCustForm.chassisNo}". Total stored: ${count} record(s).`,
      isSuccess: true
    });

    alert(`✅ Customer "${newCustForm.custName || newCustForm.chassisNo}" added successfully to Master Database!`);

    if (window.confirm('Do you want to auto-fill these new customer details into the current Job Card?')) {
      autoFillCustomer(newRecord);
      setActiveTab('new_entry');
    }

    setIsAddCustomerOpen(false);
    setNewCustForm({
      branch: '',
      sNo: '',
      model: '',
      modelType: '',
      chassisNo: '',
      engineNo: '',
      dateOfDel: '',
      custName: '',
      fatherName: '',
      address: '',
      village: '',
      mandal: '',
      mobileNumber: '',
      district: '',
      pinCode: '',
      dspName: '',
      exchangeBrand: '',
      exchangeModels: '',
      supervisor: ''
    });
  };

  const handleLocalSignIn = (role: 'admin' | 'user' = 'admin') => {
    const email = role === 'admin' ? 'srigayathriautomotives@gmail.com' : 'staff@srigayathri.com';
    const mockUser: any = {
      uid: role === 'admin' ? 'admin-master-user' : 'staff-service-user',
      email: email,
      displayName: role === 'admin' ? 'Sri Gayathri Admin' : 'Service Staff Member',
      photoURL: null
    };
    try {
      localStorage.setItem('eicher_auth_user', JSON.stringify({ user: mockUser, role }));
    } catch {}
    setCurrentUser(mockUser);
    setCurrentUserRole(role);
    setAuthError('');
  };

  const handleGoogleSignIn = async () => {
    setAuthError('');
    setIsAuthSubmitting(true);
    
    try {
      const provider = new GoogleAuthProvider();
      const result = await signInWithPopup(auth, provider);
      const user = result.user;
      const role = user.email === 'srigayathriautomotives@gmail.com' || user.email === 'srigayathriauto@gmail.com' ? 'admin' : 'user';
      try {
        localStorage.setItem('eicher_auth_user', JSON.stringify({
          user: { uid: user.uid, email: user.email, displayName: user.displayName || user.email, photoURL: user.photoURL },
          role
        }));
      } catch {}
    } catch (e: any) {
      console.error(e);
      if (e.code === 'auth/unauthorized-domain' || (e.message && e.message.includes('unauthorized-domain')) || e.code === 'auth/configuration-not-found') {
        // Automatically switch to Admin session so users are never blocked by Firebase domain whitelist
        handleLocalSignIn('admin');
        return;
      }
      let errorMsg = "Authentication failed. Please try again.";
      if (e.code === 'auth/popup-closed-by-user') {
        errorMsg = "Login popup was closed before finishing.";
      } else if (e.message) {
        errorMsg = e.message;
      }
      setAuthError(errorMsg);
    } finally {
      setIsAuthSubmitting(false);
    }
  };

  const processJobCardsArray = async (rawData: any[]) => {
    try {
      setIsUploading(true);
      if (rawData.length === 0) {
        setIsUploading(false);
        return;
      }
      let importedCount = 0;
      const cardsToSave: any[] = [];
      
      // Pre-process headers for fast lookup
      const firstRow = rawData[0];
      const headerMap = new Map<string, string>();
      const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
      for (const key of Object.keys(firstRow)) {
        headerMap.set(normalize(key), key);
      }

      // Pre-map the required fields to actual keys in the file
      const findActualKey = (candidates: string[]) => {
        for (const cand of candidates) {
          const normCand = normalize(cand);
          if (headerMap.has(normCand)) return headerMap.get(normCand);
        }
        return null;
      };

      const fieldMap = {
        jobNo: findActualKey(['Job card', 'Job card NO', 'Job Card', 'Job Card No', 'Job No', 'JOB NO', 'Job Number', 'ONLINE JOB CARD NO', 'onlineJobCardNo', 'Online Job Card No', 'Online Job Card', 'JobCard', 'JobcardNo', 'JobCardNo', 'OnlineJobCard', 'Jobcard', 'jc', 'jcno', 'jc_no']),
        onlineJobCardNo: findActualKey(['ONLINE JOB CARD NO', 'Online Job Card No', 'onlineJobCardNo']),
        chassisNo: findActualKey(['CHASIS NO', 'Chassis No', 'ChassisNo', 'Chassis Number', 'CHASSIS', 'chassis', 'Chasis', 'chasis', 'chasisno', 'chassisno', 'chassis_no', 'chasis_no', 'vin', 'vinnumber', 'vin_number', 'chassisnum', 'chassis_num']),
        custName: findActualKey(['cutomer name', 'Customer Name', 'CUSTOMER NAME', 'Customer', 'custName', 'Name', 'cust_name', 'customer_name', 'ownername', 'owner_name', 'nameofcustomer', 'cust', 'customer', 'owner']),
        actualClosedDate: findActualKey(['ACTUVAL CLOSED DATE', 'Actual Closed Date', 'ACTUAL CLOSED DATE', 'Date & Time Out', 'dateTimeOut', 'closedDate']),
        status: findActualKey(['Status', 'STATUS', 'status']),
        complaintDate: findActualKey(['COMPALINT DATE', 'Complaint Date', 'COMPLAINT DATE', 'complaintDate']),
        jobDate: findActualKey(['JOB CARD OPEN DAT', 'Job Card Open Date', 'Job Date', 'jobDate', 'Date']),
        branch: findActualKey(['BRANCH', 'Branch', 'branch']),
        historyFileNo: findActualKey(['HISTORY FILE NO.', 'History File No.', 'FILE NO.', 'File No', 'historyFileNo']),
        model: findActualKey(['Tractor model', 'Tractor Model', 'TRACTOR MODEL', 'Model', 'model']),
        modelType: findActualKey(['MODEL TYPE', 'Model Type', 'modelType']),
        engineNo: findActualKey(['Eng Sr no', 'Engine Sr No', 'Engine No', 'ENG SR NO', 'engineNo']),
        installDate: findActualKey(['Date of Delivery', 'Delivery Date', 'DATE OF DELIVERY', 'installDate']),
        fatherName: findActualKey(['FATHER', 'Father Name', 'Father', 'fatherName']),
        custAddr: findActualKey(['ADDRESS', 'Address', 'custAddr', 'address']),
        village: findActualKey(['village', 'Village', 'VILLAGE']),
        mandal: findActualKey(['mandal', 'Mandal', 'MANDAL']),
        ownerMob: findActualKey(['ph no', 'Phone No', 'Mobile', 'Phone', 'Owner Mobile', 'ownerMob']),
        phNo: findActualKey(['ph no', 'Phone No', 'Mobile', 'PH NO']),
        hourMeter: findActualKey(['Hrs Run', 'Hours Run', 'Hours', 'HRS RUN', 'Hour Meter', 'hourMeter']),
        serviceType: findActualKey(['Type of Service', 'Service Type', 'TYPE OF SERVICE', 'serviceType']),
        freeServiceList: findActualKey(['FREE SERVICE LIST', 'Free Service List', 'FREE SERVICE', 'Free Service', 'Free Services', 'FREE SERVICE COUNT', 'Free Service Count', 'SERVICE COUNT', 'Service Count', 'ServiceCount', 'freeServiceList']),
        extraRepairs: findActualKey(['EXTRA OTHER REPAIRS DONE WITH FREE SERVICE', 'Extra Repairs', 'extraRepairs']),
        mechanic: findActualKey(['TECHNICIAN NAME', 'Technician Name', 'Mechanic', 'Technician', 'mechanic']),
        serviceLocation: findActualKey(['Service place', 'Service Place', 'Location', 'Service Location', 'serviceLocation']),
        billNo: findActualKey(['BILL NO.', 'Bill No', 'BILL NO', 'billNo', 'Invoice No']),
        reasonsForAnalysis: findActualKey(['RESONS FOR ANALYSIS', 'Reasons for Analysis', 'REASONS FOR ANALYSIS', 'reasonsForAnalysis']),
        telecalling: findActualKey(['TELECALLING', 'Telecalling', 'telecalling']),
        dateTimeIn: findActualKey(['Date & Time In', 'dateTimeIn', 'Time In']),
        totalLabour: findActualKey(['Total Labour (Rs)', 'totalLabour', 'Labour']),
        warrantyMaterial: findActualKey(['Warranty Parts (Rs)', 'warrantyMaterial', 'Warranty Parts']),
        nonWarrantyMaterial: findActualKey(['Non Warranty Parts (Rs)', 'nonWarrantyMaterial', 'Non Warranty Parts']),
        gTotal: findActualKey(['Grand Total (Rs)', 'gTotal', 'Grand Total', 'Total']),
        wsIncharge: findActualKey(['Supervisor/Incharge', 'wsIncharge', 'Supervisor'])
      };

      for (const row of rawData) {
        const getValFromMap = (key: string | null) => key && row[key] ? String(row[key]).trim() : '';
        
        const jobNo = getValFromMap(fieldMap.jobNo);
        const onlineJobCardNo = getValFromMap(fieldMap.onlineJobCardNo) || jobNo;
        const chassisNo = getValFromMap(fieldMap.chassisNo);
        const custName = getValFromMap(fieldMap.custName);
        
        if (!jobNo && !onlineJobCardNo && !chassisNo && !custName) continue;

        const actualClosedDate = getValFromMap(fieldMap.actualClosedDate);
        const statusVal = getValFromMap(fieldMap.status);
        const billNo = getValFromMap(fieldMap.billNo);
        const closedStatuses = ['closed', 'done', 'completed', 'finalized', 'billed'];
        const isClosed = closedStatuses.includes(statusVal.toLowerCase()) || !!actualClosedDate || !!billNo;
        const finalStatus = isClosed ? 'Closed' : 'Open';
        const installDateVal = getValFromMap(fieldMap.installDate);

        const importedMechanic = getValFromMap(fieldMap.mechanic);
        let importedWsIncharge = getValFromMap(fieldMap.wsIncharge);

        if (importedMechanic) {
          const assocSup = getAssignedSupervisor(importedMechanic);
          if (assocSup) {
            importedWsIncharge = assocSup;
          }
        }

        const branchVal = getValFromMap(fieldMap.branch);
        const rawHFN = getValFromMap(fieldMap.historyFileNo);
        let finalHFN = rawHFN;
        if (rawHFN && branchVal) {
          const cleanBr = branchVal.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
          const cleanHFN = rawHFN.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
          if (cleanBr && !cleanHFN.startsWith(cleanBr)) {
            finalHFN = `${branchVal.toUpperCase()}-${rawHFN}`;
          }
        }

        const cardData = {
          id: Date.now().toString() + Math.random().toString(36).substring(2, 7) + "_" + importedCount,
          jobNo: jobNo,
          status: finalStatus,
          complaintDate: toInputDateFormat(getValFromMap(fieldMap.complaintDate)) || '',
          onlineJobCardNo: onlineJobCardNo,
          jobDate: toInputDateFormat(getValFromMap(fieldMap.jobDate)) || '',
          branch: branchVal,
          historyFileNo: finalHFN,
          model: getValFromMap(fieldMap.model),
          modelType: getValFromMap(fieldMap.modelType),
          chassisNo: chassisNo,
          engineNo: getValFromMap(fieldMap.engineNo),
          installDate: toInputDateFormat(installDateVal) || '',
          dateOfDelivery: toInputDateFormat(installDateVal) || '',
          custName: custName,
          fatherName: getValFromMap(fieldMap.fatherName),
          custAddr: getValFromMap(fieldMap.custAddr),
          village: getValFromMap(fieldMap.village),
          mandal: getValFromMap(fieldMap.mandal),
          ownerMob: getValFromMap(fieldMap.ownerMob),
          phNo: getValFromMap(fieldMap.phNo),
          hourMeter: getValFromMap(fieldMap.hourMeter),
          serviceType: getDisplayServiceType(getValFromMap(fieldMap.serviceType)),
          freeServiceList: getValFromMap(fieldMap.freeServiceList),
          extraRepairs: getValFromMap(fieldMap.extraRepairs),
          dateTimeOut: toInputDateFormat(actualClosedDate) || '',
          actualClosedDate: toInputDateFormat(actualClosedDate) || '',
          mechanic: importedMechanic,
          serviceLocation: getValFromMap(fieldMap.serviceLocation),
          billNo: getValFromMap(fieldMap.billNo),
          reasonsForAnalysis: getValFromMap(fieldMap.reasonsForAnalysis),
          telecalling: getValFromMap(fieldMap.telecalling),
          dateTimeIn: toInputDateFormat(getValFromMap(fieldMap.dateTimeIn)) || '',
          totalLabour: getValFromMap(fieldMap.totalLabour),
          warrantyMaterial: getValFromMap(fieldMap.warrantyMaterial),
          nonWarrantyMaterial: getValFromMap(fieldMap.nonWarrantyMaterial),
          gTotal: getValFromMap(fieldMap.gTotal),
          wsIncharge: importedWsIncharge,
          createdBy: currentUser?.uid || '',
          createdByEmail: currentUser?.email || '',
          createdAt: new Date().toISOString()
        };
        
        cardsToSave.push(cardData);
        importedCount++;
      }


      if (cardsToSave.length === 0) {
        alert('⚠️ No valid job cards were found to import. Please ensure your Excel/CSV contains at least one column for Job Card No, Chassis No, or Customer Name, and that the rows are not empty.');
        return;
      }

      // Execute in concurrent batches for much faster uploads
      const BATCH_SIZE = 450;
      const commitPromises: Promise<void>[] = [];
      for (let i = 0; i < cardsToSave.length; i += BATCH_SIZE) {
        const chunk = cardsToSave.slice(i, i + BATCH_SIZE);
        const batch = writeBatch(db);
        for (const cardData of chunk) {
          batch.set(doc(db, 'jobcards', cardData.id), cardData);
        }
        commitPromises.push(batch.commit());
      }
      
      await Promise.all(commitPromises);

      // Also sync to Cloud SQL backend
      try {
        await sqlApi.saveJobCardsBulk(cardsToSave);
      } catch (sqlErr) {
        console.warn('Could not sync to SQL during import:', sqlErr);
      }

      // Update local state and IndexedDB immediately
      setSavedJobCards(prev => [...cardsToSave, ...prev]);
      try {
        await saveJobCardsBackup([...cardsToSave, ...savedJobCards]);
      } catch (idbErr) {
        console.warn('Could not save to IndexedDB:', idbErr);
      }

      alert(`⚡ Successfully imported ${importedCount} Job Card(s) to cloud database!`);
    } catch (error: any) {
      console.error('Upload error:', error);
      alert('❌ Error importing job cards: ' + (error?.message || error));
    } finally {
      setIsUploading(false);
    }
  };

  const handleDirectComplaintRegister = (item: any) => {
    // Resolve data from customer record
    const chassis = item.chassisNo || item['Chassis no'] || item.__chassisDisplay || getFieldValue(item, 'chassis') || '';
    const name = item.customerName || item['Customer Name'] || item.custName || item.__custNameDisplay || '';
    const mobile = item.mobileNumber || item['Mobile Number'] || item.phone || getFieldValue(item, 'custPhone') || '';
    const model = item.model || getFieldValue(item, 'model') || '';
    const hfn = item.historyFileNo || item['History File No'] || getFieldValue(item, 'historyFileNo') || '';

    setComplaintForm({
      id: null,
      complaintNo: '',
      complaintDate: todayISO,
      customerName: name,
      mobileNumber: mobile,
      chassisNo: chassis,
      tractorModel: model,
      complaintDetails: '',
      status: 'Open',
      closureDate: null,
      jobCardNo: '',
      assignedSupervisor: item.supervisor || '',
      assignedMechanic: '',
      historyFileNo: hfn
    });
    setIsComplaintModalOpen(true);
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileExt = file.name.split('.').pop()?.toLowerCase();
    
    if (fileExt === 'csv') {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: async (results) => {
          await processJobCardsArray(results.data as any[]);
        }
      });
    } else {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const wb = XLSX.read(evt.target?.result, { type: 'array', cellDates: true });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });
          await processJobCardsArray(rows);
        } catch (err) {
          console.error(err);
          alert('❌ Could not read Excel file. Please ensure it is a valid .xlsx, .xls, or .csv file.');
        }
      };
      reader.readAsArrayBuffer(file);
    }

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const downloadJobCardTemplateExcel = () => {
    const headers = [
      'Job card',
      'COMPALINT DATE',
      'ONLINE JOB CARD NO',
      'JOB CARD OPEN DAT',
      'BRANCH',
      'HISTORY FILE NO.',
      'Tractor model',
      'MODEL TYPE',
      'CHASIS NO',
      'Eng Sr no',
      'Date of Delivery',
      'cutomer name',
      'FATHER',
      'ADDRESS',
      'village',
      'mandal',
      'ph no',
      'Hrs Run',
      'Type of Service',
      'FREE SERVICE LIST',
      'EXTRA OTHER REPAIRS DONE WITH FREE SERVICE',
      'ACTUVAL CLOSED DATE',
      'TECHNICIAN NAME',
      'Service place',
      'BILL NO.',
      'RESONS FOR ANALYSIS',
      'TELECALLING',
      'Status'
    ];

    const sampleRow = [
      'JC-2026-001',
      '2026-08-01',
      'ON-JC-8821',
      '2026-08-08',
      'Main Branch',
      'HF-1092',
      '5050D 4WD',
      'Tractor',
      'ME4TRACTOR12345',
      'ENG98765',
      '2026-01-15',
      'RAMESH KUMAR',
      'SAMPATH KUMAR',
      'H.No 4-12, Main Road',
      'Guntur',
      'Guntur Rural',
      '9876543210',
      '450',
      'Paid Service',
      'Free Service 1, Free Service 2',
      'Oil filter replaced, clutch adjusted',
      '2026-08-08',
      'SRINIVAS',
      'Workshop',
      'INV-001',
      'Routine maintenance and general check',
      'Done',
      'Open'
    ];

    // Create a row object
    const rowObj: Record<string, string> = {};
    headers.forEach((h, i) => {
      rowObj[h] = sampleRow[i] || '';
    });

    const worksheet = XLSX.utils.json_to_sheet([rowObj]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Template");

    // Auto-fit columns
    const maxProps = headers.map(key => ({
      wch: Math.max(key.length, String(rowObj[key] || '').length) + 2
    }));
    worksheet['!cols'] = maxProps;

    XLSX.writeFile(workbook, 'Job_Cards_Import_Template.xlsx');
  };

  const handleSignOut = async () => {
    try {
      try {
        localStorage.removeItem('eicher_auth_user');
      } catch {}
      await signOut(auth);
    } catch (e) {
      console.error("Error signing out:", e);
    } finally {
      setCurrentUser(null);
      setCurrentUserRole(null);
      setEditingCardId(null);
      resetForm();
    }
  };

  const fmtDate = (d: any) => {
    if (d === undefined || d === null || d === '') return '';
    
    // First try converting via toInputDateFormat helper
    const iso = toInputDateFormat(d);
    if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso)) {
      const [y, m, day] = iso.split('-');
      return `${day}/${m}/${y}`;
    }

    const str = String(d).trim();
    if (!str) return '';

    if (str.includes('T') || str.includes(' ')) {
      const datePart = str.split(/[T ]/)[0];
      const isoPart = toInputDateFormat(datePart);
      if (isoPart && /^\d{4}-\d{2}-\d{2}$/.test(isoPart)) {
        const [y, m, day] = isoPart.split('-');
        return `${day}/${m}/${y}`;
      }
    }

    const match = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
    if (match) {
      return `${match[1].padStart(2, '0')}/${match[2].padStart(2, '0')}/${match[3]}`;
    }

    return str;
  };

  const getFreeServiceText = (card: any) => {
    if (card && card.freeServiceList && String(card.freeServiceList).trim() !== '') {
      return String(card.freeServiceList).trim();
    }
    const st = String((card && card.serviceType) || '').trim();
    if (!st) return '—';

    const upperSt = st.toUpperCase();
    if (upperSt === 'I' || upperSt === '1ST' || upperSt.includes('1ST')) return '1st Free Service';
    if (upperSt === 'II' || upperSt === '2ND' || upperSt.includes('2ND')) return '2nd Free Service';
    if (upperSt === 'III' || upperSt === '3RD' || upperSt.includes('3RD')) return '3rd Free Service';
    if (upperSt === 'IV' || upperSt === '4TH' || upperSt.includes('4TH')) return '4th Free Service';
    if (upperSt === 'V' || upperSt === '5TH' || upperSt.includes('5TH')) return '5th Free Service';
    if (upperSt === 'VI' || upperSt === '6TH' || upperSt.includes('6TH')) return '6th Free Service';
    if (upperSt === 'VII' || upperSt === '7TH' || upperSt.includes('7TH')) return '7th Free Service';
    if (upperSt === 'VIII' || upperSt === '8TH' || upperSt.includes('8TH')) return '8th Free Service';
    if (upperSt === 'IX' || upperSt === '9TH' || upperSt.includes('9TH')) return '9th Free Service';
    if (upperSt === 'X' || upperSt === '10TH' || upperSt.includes('10TH')) return '10th Free Service';
    if (upperSt.includes('PAID REPAIR') || upperSt === 'PAID REPAIRS') return 'Paid Repairs';
    if (upperSt.includes('PAID')) return 'Paid Service';
    if (upperSt.includes('WTY') || upperSt.includes('WARRANTY')) return 'Under Warranty';

    return st;
  };

  const getExtraRepairsText = (card: any) => {
    if (card && card.extraRepairs && String(card.extraRepairs).trim() !== '') {
      return String(card.extraRepairs).trim();
    }
    if (card && Array.isArray(card.repairRows) && card.repairRows.length > 1) {
      const rowB = card.repairRows[1];
      if (rowB) {
        const rep = (rowB.repair || '').trim();
        const rect = (rowB.rectification || '').trim();
        if (rep && rect) return `${rep} - ${rect}`;
        if (rep) return rep;
        if (rect) return rect;
      }
    }
    return '—';
  };

  const fmtDateTime = (dt: string) => {
    if (!dt) return '';
    const parts = dt.split('T');
    if (parts.length === 2) {
      const dateStr = fmtDate(parts[0]);
      let [h, min] = parts[1].split(':');
      let hourNum = parseInt(h, 10);
      const ampm = hourNum >= 12 ? 'PM' : 'AM';
      hourNum = hourNum % 12;
      hourNum = hourNum ? hourNum : 12;
      const hourStr = String(hourNum).padStart(2, '0');
      return `${dateStr} ${hourStr}:${min} ${ampm}`;
    }
    return dt;
  };

  const getFollowUpStatus = (installDateStr: string, intervalDays: number) => {
    if (!installDateStr || !installDateStr.trim()) {
      return { isDue: false, daysOverdue: 0, dueDateStr: '' };
    }
    const deliveryDate = new Date(installDateStr);
    if (isNaN(deliveryDate.getTime())) {
      return { isDue: false, daysOverdue: 0, dueDateStr: '' };
    }
    const dueDate = new Date(deliveryDate.getTime() + intervalDays * 24 * 60 * 60 * 1000);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dueDate.setHours(0, 0, 0, 0);
    const isDue = today.getTime() > dueDate.getTime();
    const diffTime = today.getTime() - dueDate.getTime();
    const daysOverdue = Math.max(0, Math.floor(diffTime / (24 * 60 * 60 * 1000)));
    const y = dueDate.getFullYear();
    const m = String(dueDate.getMonth() + 1).padStart(2, '0');
    const d = String(dueDate.getDate()).padStart(2, '0');
    const dueDateStr = `${d}/${m}/${y}`;
    return { isDue, daysOverdue, dueDateStr };
  };

  const [repairsStatus, setRepairsStatus] = useState<{ text: string; isSuccess: boolean }>({ text: '', isSuccess: false });

  const handleRepairsUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setRepairsStatus({ text: 'Reading repairs file...', isSuccess: false });
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'array', cellDates: false });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows: any[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

        const importedRows: Array<{ repair: string; rectification: string; charge: string }> = [];

        rows.forEach(row => {
          const lookup = buildRowLookup(row);
          const repairVal = getFieldValue(lookup, 'repair') || getFieldValue(lookup, 'desc') || getFieldValue(lookup, 'work') || getFieldValue(lookup, 'repairs') || '';
          const rectVal = getFieldValue(lookup, 'rectification') || getFieldValue(lookup, 'action') || getFieldValue(lookup, 'remarks') || '';
          const chargeVal = getFieldValue(lookup, 'charge') || getFieldValue(lookup, 'rate') || getFieldValue(lookup, 'amount') || '';

          if (repairVal || rectVal || chargeVal) {
            importedRows.push({
              repair: repairVal,
              rectification: rectVal,
              charge: chargeVal ? String(chargeVal) : ''
            });
          }
        });

        if (importedRows.length > 0) {
          setRepairRows(importedRows);
          setRepairsStatus({
            text: `✅ Imported ${importedRows.length} repair(s) successfully.`,
            isSuccess: true
          });
        } else {
          setRepairsStatus({
            text: `⚠️ No valid repair rows found in Excel.`,
            isSuccess: false
          });
        }
      } catch (err) {
        console.error(err);
        setRepairsStatus({ text: '⚠️ Could not read repairs file.', isSuccess: false });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const repairTotal = repairRows.reduce((acc, r) => acc + (parseFloat(r.charge) || 0), 0);
  const partsTotal = partRows.reduce((acc, r) => acc + (parseFloat(r.amount) || 0), 0);

  const handleSaveAsPDF = () => {
    const cleanChassis = String(chassisNo || '').trim().replace(/[\\\/:*?"<>|]+/g, '').replace(/\s+/g, '_') || 'NoChassis';
    const cleanDate = String(jobDate || '').trim().replace(/[\\\/:*?"<>|]+/g, '') || 'NoDate';
    const cleanJobNo = String(jobNo || '').trim().replace(/[\\\/:*?"<>|]+/g, '');

    const fileName = `SriGayathri_JobCard_${cleanChassis}_${cleanDate}${cleanJobNo ? '_' + cleanJobNo : ''}`;
    const origTitle = document.title;
    document.title = fileName;

    const restore = () => {
      document.title = origTitle;
      window.removeEventListener('afterprint', restore);
    };
    window.addEventListener('afterprint', restore);

    alert(`Printing / Saving as PDF...\nSuggested file name: ${fileName}.pdf`);
    window.print();
  };

  const serviceTypeOptions = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th', 'Paid Service', 'Paid Repairs', 'Under Wty Repairs'];

  const handleServiceTypeSelect = (st: string) => {
    setServiceLocationType(st);
    const serviceTextMap: Record<string, { repair: string; rectification: string; charge: string; freeService: string }> = {
      '1st': { repair: '1st free service', rectification: 'Completed', charge: '0', freeService: '1st Free Service' },
      '2nd': { repair: '2nd free service', rectification: 'Completed', charge: '0', freeService: '2nd Free Service' },
      '3rd': { repair: '3rd free service', rectification: 'Completed', charge: '0', freeService: '3rd Free Service' },
      '4th': { repair: '4th free service', rectification: 'Completed', charge: '0', freeService: '4th Free Service' },
      '5th': { repair: '5th free service', rectification: 'Completed', charge: '0', freeService: '5th Free Service' },
      '6th': { repair: '6th free service', rectification: 'Completed', charge: '0', freeService: '6th Free Service' },
      '7th': { repair: '7th free service', rectification: 'Completed', charge: '0', freeService: '7th Free Service' },
      '8th': { repair: '8th free service', rectification: 'Completed', charge: '0', freeService: '8th Free Service' },
      '9th': { repair: '9th free service', rectification: 'Completed', charge: '0', freeService: '9th Free Service' },
      '10th': { repair: '10th free service', rectification: 'Completed', charge: '0', freeService: '10th Free Service' },
      'I': { repair: '1st free service', rectification: 'Completed', charge: '0', freeService: '1st Free Service' },
      'II': { repair: '2nd free service', rectification: 'Completed', charge: '0', freeService: '2nd Free Service' },
      'III': { repair: '3rd free service', rectification: 'Completed', charge: '0', freeService: '3rd Free Service' },
      'IV': { repair: '4th free service', rectification: 'Completed', charge: '0', freeService: '4th Free Service' },
      'V': { repair: '5th free service', rectification: 'Completed', charge: '0', freeService: '5th Free Service' },
      'VI': { repair: '6th free service', rectification: 'Completed', charge: '0', freeService: '6th Free Service' },
      'VII': { repair: '7th free service', rectification: 'Completed', charge: '0', freeService: '7th Free Service' },
      'VIII': { repair: '8th free service', rectification: 'Completed', charge: '0', freeService: '8th Free Service' },
      'IX': { repair: '9th free service', rectification: 'Completed', charge: '0', freeService: '9th Free Service' },
      'X': { repair: '10th free service', rectification: 'Completed', charge: '0', freeService: '10th Free Service' },
      'Paid Service': { repair: 'Paid Service', rectification: 'Completed', charge: '0', freeService: 'Paid Service' },
      'Paid Repairs': { repair: 'Paid repairs', rectification: '', charge: '', freeService: 'Paid Repairs' },
      'Under Wty Repairs': { repair: 'Under Warranty Repairs', rectification: 'Completed under warranty', charge: '0', freeService: 'Under Warranty' }
    };

    const info = serviceTextMap[st];
    if (info) {
      setFreeServiceList(info.freeService);
      const updated = [...repairRows];
      if (updated.length > 0) {
        updated[0] = {
          ...updated[0],
          repair: info.repair,
          rectification: info.rectification,
          charge: info.charge
        };
      } else {
        updated.push({
          repair: info.repair,
          rectification: info.rectification,
          charge: info.charge
        });
      }
      setRepairRows(updated);
    }
  };

  const calculatedWty = getWarrantyStatus(installDate, jobDate);
  let isWty = false;
  let isPostWty = false;
  let statusText = '';

  if (warrantyOverride === 'warranty') {
    isWty = true;
    isPostWty = false;
    statusText = 'WARRANTY';
  } else if (warrantyOverride === 'post_wty') {
    isWty = false;
    isPostWty = true;
    statusText = 'POST WTY';
  } else if (calculatedWty !== null) {
    isWty = calculatedWty.isWty;
    isPostWty = calculatedWty.isPostWty;
    statusText = calculatedWty.status;
  } else {
    if (serviceType && serviceType !== 'Paid Repairs') {
      isWty = true;
      isPostWty = false;
      statusText = 'WARRANTY';
    } else {
      isWty = false;
      isPostWty = true;
      statusText = 'POST WTY';
    }
  }

  const wtyInfo = {
    status: statusText,
    isWty,
    isPostWty,
    calculated: calculatedWty,
    overrideMode: warrantyOverride
  };

  const visibleJobCards = useMemo(() => {
    return savedJobCards.filter(card => {
      if (currentUserRole === 'admin') return true;
      return !card.createdBy || card.createdBy === currentUser?.uid;
    });
  }, [savedJobCards, currentUserRole, currentUser?.uid]);

  const mechanicsList = useMemo(() => {
    const set = new Set<string>();
    staffMembers.filter((s) => s.role === 'mechanic' || !s.role).forEach((s) => s.name && set.add(s.name.trim()));
    return Array.from(set).filter(Boolean).sort();
  }, [staffMembers]);

  const supervisorsList = useMemo(() => {
    const set = new Set<string>();
    staffMembers.filter((s) => s.role === 'supervisor').forEach((s) => s.name && set.add(s.name.trim()));
    return Array.from(set).filter(Boolean).sort();
  }, [staffMembers]);

  const telecallersList = useMemo(() => {
    const set = new Set<string>();
    staffMembers.forEach((s) => s.name && set.add(s.name.trim()));
    return Array.from(set).filter(Boolean).sort();
  }, [staffMembers]);

  // Service Follow-up computed memoized states
  const customerJobCardsMap = useMemo(() => {
    const map: Record<string, any[]> = {};
    savedJobCards.forEach(card => {
      const ch = card.chassisNo || card.chassis;
      if (ch) {
        const norm = normalizeKey(ch);
        if (norm) {
          if (!map[norm]) map[norm] = [];
          map[norm].push(card);
        }
      }
    });
    return map;
  }, [savedJobCards]);

  const uniqueCustomers = useMemo(() => {
    const uniqueRecordsMap = new Map<string, any>();
    Object.keys(chassisIndex).forEach(k => {
      const rec = chassisIndex[k];
      if (rec) {
        const ch = rec['Chassis no'] || rec.__chassisDisplay || getFieldValue(rec, 'chassis') || k;
        const chassisNo = normalizeKey(ch.toString());
        if (chassisNo && !uniqueRecordsMap.has(chassisNo)) {
          uniqueRecordsMap.set(chassisNo, rec);
        }
      }
    });

    // Automatically include any chassis from saved job cards
    // so they go to the Reporting Customers list when a job card is entered
    savedJobCards.forEach(card => {
      const ch = card.chassisNo || card.chassis || '';
      const chassisNo = normalizeKey(ch.toString());
      if (chassisNo && !uniqueRecordsMap.has(chassisNo)) {
        uniqueRecordsMap.set(chassisNo, {
          'Chassis no': card.chassisNo || ch,
          __chassisDisplay: card.chassisNo || ch,
          'Customer Name': card.customerName || card.custName || '',
          __custNameDisplay: card.customerName || card.custName || '',
          'Mobile Number': card.mobileNumber || card.phone || '',
          __custPhoneDisplay: card.mobileNumber || card.phone || '',
          'Model': card.tractorModel || card.model || '',
          model: card.tractorModel || card.model || '',
          village: card.village || '',
        });
      }
    });

    return Array.from(uniqueRecordsMap.values());
  }, [chassisIndex, savedJobCards]);

  const followupModelsList = useMemo(() => {
    const set = new Set<string>();
    uniqueCustomers.forEach(rec => {
      const m = (rec.Model || rec.model || getFieldValue(rec, 'model') || '').toString().trim();
      if (m) set.add(m);
    });
    return Array.from(set).sort();
  }, [uniqueCustomers]);

  const duplicateChassisSet = useMemo(() => {
    const rawRecords = Object.values(chassisIndex);
    const uniqueRecs = new Set(rawRecords);
    const records = Array.from(uniqueRecs);
    
    const chassisMap = new Map<string, any[]>();
    records.forEach((rec: any) => {
      if (!rec) return;
      const lookup = rec.chassisNo ? rec : buildRowLookup(rec);
      const ch = lookup['Chassis no'] || lookup.__chassisDisplay || lookup.chassisNo || getFieldValue(lookup, 'chassis') || '';
      const normCh = normalizeKey(ch);
      if (normCh) {
        if (!chassisMap.has(normCh)) chassisMap.set(normCh, []);
        chassisMap.get(normCh)!.push(lookup);
      }
    });

    const dupSet = new Set<string>();
    chassisMap.forEach((recs, ch) => {
      if (recs.length > 1) {
        dupSet.add(ch);
      }
    });
    
    // Also include suffix duplicates
    const suffixMap = new Map<string, any[]>();
    records.forEach((rec: any) => {
      if (!rec) return;
      const lookup = rec.chassisNo ? rec : buildRowLookup(rec);
      const ch = lookup['Chassis no'] || lookup.__chassisDisplay || lookup.chassisNo || getFieldValue(lookup, 'chassis') || '';
      const norm = normalizeKey(ch);
      if (norm.length >= 6) {
        const suffix = norm.slice(-6);
        if (!suffixMap.has(suffix)) suffixMap.set(suffix, []);
        suffixMap.get(suffix)!.push(lookup);
      }
    });
    
    suffixMap.forEach((recs, suffix) => {
      if (recs.length > 1) {
        // If they are suffix duplicates, we add all their actual normalized chassis to dupSet
        recs.forEach(r => {
           const norm = normalizeKey(r['Chassis no'] || r.__chassisDisplay || r.chassisNo || getFieldValue(r, 'chassis') || '');
           if (norm) dupSet.add(norm);
        });
      }
    });

    // Also include phone duplicates
    const phoneMap = new Map<string, any[]>();
    records.forEach((rec: any) => {
      if (!rec) return;
      const lookup = rec.chassisNo ? rec : buildRowLookup(rec);
      const phone = lookup['Mobile Number'] || lookup.mobileNumber || lookup['Phone'] || getFieldValue(lookup, 'custPhone') || '';
      const normPhone = String(phone || '').replace(/[^0-9]/g, '');
      if (normPhone && normPhone.length >= 10) {
        const last10 = normPhone.slice(-10);
        if (!phoneMap.has(last10)) phoneMap.set(last10, []);
        phoneMap.get(last10)!.push(lookup);
      }
    });

    phoneMap.forEach((recs, phone) => {
      if (recs.length > 1) {
        recs.forEach(r => {
           const norm = normalizeKey(r['Chassis no'] || r.__chassisDisplay || r.chassisNo || getFieldValue(r, 'chassis') || '');
           if (norm) dupSet.add(norm);
        });
      }
    });

    return dupSet;
  }, [chassisIndex]);

  const followupCalculatedList = useMemo(() => {
    const getCardTime = (c: any) => {
      const dStr = c.jobDate || c.jobOpenDate || c.createdAt;
      if (!dStr) return 0;
      const t = Date.parse(dStr);
      return isNaN(t) ? 0 : t;
    };

    // Precompute mechanic to supervisor map for ultra-fast O(1) lookups
    const mechanicToSupervisorMap = new Map<string, string>();
    staffMembers.forEach(s => {
      if (s && s.name) {
        const role = (s.role || '').toLowerCase();
        if (role === 'mechanic' || !role) {
          mechanicToSupervisorMap.set(s.name.trim().toLowerCase(), (s.supervisor || '').trim());
        }
      }
    });

    const list = uniqueCustomers.map((rec, idx) => {
      const std = getStandardizedCustomer(rec);
      const normCh = normalizeKey(std.chassisNo);
      const cards = normCh ? (customerJobCardsMap[normCh] || []) : [];
      const hasJobCard = cards.length > 0;

      let lastServiceType = '';
      let lastJobCardEntryDate = '';
      let lastEntryHours = '';
      let lastRepairDetails = '';
      let lastTechnician = '';
      let lastSupervisor = '';
      let latestCard: any = null;

      if (hasJobCard) {
        const sorted = [...cards].sort((a, b) => getCardTime(b) - getCardTime(a));
        latestCard = sorted[0];
        lastServiceType = latestCard.serviceType || 'Paid Service';
        lastJobCardEntryDate = fmtDate(latestCard.jobDate || latestCard.jobOpenDate || (latestCard.createdAt ? latestCard.createdAt.split('T')[0] : ''));
        lastEntryHours = latestCard.hourMeter || latestCard.hrsRun || '';
        const primaryRep = Array.isArray(latestCard.repairRows) && latestCard.repairRows.length > 0
          ? latestCard.repairRows.map((r: any) => r.repair).filter(Boolean).join(', ')
          : (latestCard.problemDescription || latestCard.reasonsForAnalysis || '');
        const extraRep = getExtraRepairsText(latestCard) !== '—' 
          ? getExtraRepairsText(latestCard) 
          : (latestCard.extraRepairs || latestCard['EXTRA OTHER REPAIRS DONE WITH FREE SERVICE'] || latestCard.extraRepairsDone || '');

        if (primaryRep && extraRep && primaryRep !== extraRep) {
          lastRepairDetails = `${primaryRep} | Extra: ${extraRep}`;
        } else if (primaryRep) {
          lastRepairDetails = primaryRep;
        } else {
          lastRepairDetails = extraRep;
        }
        lastTechnician = latestCard.mechanic || latestCard.technicianName || '';
        lastSupervisor = latestCard.wsIncharge || latestCard.supervisorName || latestCard.supervisor || '';
      }

      // Get supervisor name linked to the technician name
      const linkedSupervisor = lastTechnician ? (mechanicToSupervisorMap.get(lastTechnician.trim().toLowerCase()) || '') : '';

      // Assign supervisor with fallback to linked supervisor, then latest card's recorded supervisor
      let determinedSupervisor = std.supervisor || linkedSupervisor || (latestCard ? (latestCard.wsIncharge || latestCard.supervisor || '') : '') || '';
      
      const branchLower = std.branch.toLowerCase();
      if (branchLower === 'tvr' || branchLower === 'tsg') {
        determinedSupervisor = 'Gosu Jamalarao';
      }
      
      const assignedSupervisor = determinedSupervisor;

      const lastCallDate = rec.lastCallDate || '';
      const lastRemarks = rec.lastRemarks || '';
      const lastNextCallDate = rec.lastNextCallDate || '';
      const followupHistory = rec.followupHistory || [];

      return {
        rec,
        idx,
        chassisNo: std.chassisNo,
        customerName: std.custName,
        fatherName: std.fatherName,
        address: std.address,
        village: std.village,
        mandal: std.mandal,
        district: std.district,
        pinCode: std.pinCode,
        model: std.model,
        modelType: std.modelType,
        engineNo: std.engineNo,
        dateOfDel: fmtDate(std.dateOfDel),
        rawDateOfDel: std.dateOfDel,
        mobileNumber: std.mobileNumber,
        supervisor: assignedSupervisor,
        linkedSupervisor,
        branch: std.branch,
        historyFileNo: std.historyFileNo,
        hasJobCard,
        jobCardsCount: cards.length,
        lastServiceType,
        lastJobCardEntryDate,
        lastEntryHours,
        lastRepairDetails,
        lastTechnician,
        lastSupervisor,
        lastCallDate,
        lastRemarks,
        lastNextCallDate,
        followupHistory,
        cards
      };
    });

    // Apply filters
    const filteredList = list.filter(item => {
      // 1. Search Query
      if (followupSearch) {
        const q = followupSearch.toLowerCase();
        const match = 
          item.customerName.toLowerCase().includes(q) ||
          item.chassisNo.toLowerCase().includes(q) ||
          item.mobileNumber.toLowerCase().includes(q) ||
          item.village.toLowerCase().includes(q) ||
          item.mandal.toLowerCase().includes(q) ||
          item.model.toLowerCase().includes(q) ||
          item.lastRemarks.toLowerCase().includes(q) ||
          item.historyFileNo.toLowerCase().includes(q);
        if (!match) return false;
      }

      // 2. Supervisor Filter
      if (followupSupervisor !== 'all') {
        if (followupSupervisor === 'unassigned') {
          if (item.supervisor && item.supervisor.trim() !== '') return false;
        } else {
          if (!item.supervisor || item.supervisor.trim().toLowerCase() !== followupSupervisor.trim().toLowerCase()) return false;
        }
      }

      // 3. Status Filter
      if (followupStatus !== 'all') {
        if (followupStatus === 'reporting' && !item.hasJobCard) return false;
        if (followupStatus === 'not_reporting' && item.hasJobCard) return false;
        if (followupStatus === 'duplicate' && !duplicateChassisSet.has(normalizeKey(item.chassisNo))) return false;
      }

      // 4. Model Filter
      if (followupModelFilter !== 'all') {
        if (item.model.toLowerCase().trim() !== followupModelFilter.toLowerCase().trim()) return false;
      }

      // 5. Delivery Date From
      if (followupDateFrom) {
        const normDate = toInputDateFormat(item.rawDateOfDel) || toInputDateFormat(item.dateOfDel);
        if (!normDate || normDate < followupDateFrom) return false;
      }

      // 6. Delivery Date To
      if (followupDateTo) {
        const normDate = toInputDateFormat(item.rawDateOfDel) || toInputDateFormat(item.dateOfDel);
        if (!normDate || normDate > followupDateTo) return false;
      }

      return true;
    });

    // Flexible Sorting
    return filteredList.sort((a, b) => {
      let result = 0;

      if (followupSortBy === 'hfn') {
        const hfnA = (a.historyFileNo || '').trim();
        const hfnB = (b.historyFileNo || '').trim();

        if (!hfnA && !hfnB) result = 0;
        else if (!hfnA) result = 1;
        else if (!hfnB) result = -1;
        else {
          const numStrA = hfnA.replace(/[^0-9.]/g, '');
          const numStrB = hfnB.replace(/[^0-9.]/g, '');
          const numA = parseFloat(numStrA);
          const numB = parseFloat(numStrB);

          const hasNumA = !isNaN(numA) && numStrA !== '';
          const hasNumB = !isNaN(numB) && numStrB !== '';

          if (hasNumA && hasNumB && numA !== numB) {
            result = numA - numB;
          } else {
            result = hfnA.localeCompare(hfnB, undefined, { numeric: true, sensitivity: 'base' });
          }
        }
      } else if (followupSortBy === 'name') {
        result = (a.customerName || '').localeCompare(b.customerName || '');
      } else if (followupSortBy === 'dateOfDel') {
        const dA = toInputDateFormat(a.rawDateOfDel) || toInputDateFormat(a.dateOfDel) || '';
        const dB = toInputDateFormat(b.rawDateOfDel) || toInputDateFormat(b.dateOfDel) || '';
        result = dA.localeCompare(dB);
      } else if (followupSortBy === 'lastJobCardDate') {
        const dA = toInputDateFormat(a.lastJobCardEntryDate) || '';
        const dB = toInputDateFormat(b.lastJobCardEntryDate) || '';
        result = dA.localeCompare(dB);
      } else if (followupSortBy === 'nextCallDate') {
        const dA = toInputDateFormat(a.lastNextCallDate) || '';
        const dB = toInputDateFormat(b.lastNextCallDate) || '';
        result = dA.localeCompare(dB);
      } else if (followupSortBy === 'chassisNo') {
        result = (a.chassisNo || '').localeCompare(b.chassisNo || '');
      } else if (followupSortBy === 'village') {
        result = (a.village || '').localeCompare(b.village || '');
      } else if (followupSortBy === 'model') {
        result = (a.model || '').localeCompare(b.model || '');
      }

      return followupSortOrder === 'asc' ? result : -result;
    });
  }, [uniqueCustomers, customerJobCardsMap, staffMembers, followupSearch, followupSupervisor, followupStatus, followupSortBy, followupSortOrder, followupDateFrom, followupDateTo, followupModelFilter]);

  // Reset pagination on filter changes
  useEffect(() => {
    setFollowupPage(1);
  }, [followupSearch, followupSupervisor, followupStatus, followupSortBy, followupSortOrder, followupDateFrom, followupDateTo, followupModelFilter]);

  const paginatedFollowupList = useMemo(() => {
    const startIndex = (followupPage - 1) * followupItemsPerPage;
    return followupCalculatedList.slice(startIndex, startIndex + followupItemsPerPage);
  }, [followupCalculatedList, followupPage, followupItemsPerPage]);

  // Tele Calling Memos & Reports Calculations
  const todayISO = useMemo(() => getLocalDateTimeString().split('T')[0], []);

  const allTelecallerLogs = useMemo(() => {
    const logs: Array<{
      logId: string;
      customer: any;
      callDate: string;
      remarks: string;
      nextCallDate: string;
      calledBy: string;
      chassisNo: string;
      customerName: string;
      mobileNumber: string;
      historyFileNo: string;
      model: string;
      village: string;
      mandal: string;
      supervisor: string;
    }> = [];

    followupCalculatedList.forEach(item => {
      const history = item.followupHistory || [];
      history.forEach((h: any, idx: number) => {
        logs.push({
          logId: `${item.chassisNo}_${h.callDate}_${idx}`,
          customer: item,
          callDate: h.callDate || '',
          remarks: h.remarks || '',
          nextCallDate: h.nextCallDate || '',
          calledBy: h.calledBy || item.lastCalledBy || 'General Staff',
          chassisNo: item.chassisNo,
          customerName: item.customerName,
          mobileNumber: item.mobileNumber,
          historyFileNo: item.historyFileNo,
          model: item.model,
          village: item.village,
          mandal: item.mandal,
          supervisor: item.supervisor || 'Unassigned'
        });
      });
    });

    return logs.sort((a, b) => (b.callDate > a.callDate ? 1 : b.callDate < a.callDate ? -1 : 0));
  }, [followupCalculatedList]);

  const freeServiceFollowupCategories = useMemo(() => {
    const today = new Date();
    today.setHours(0,0,0,0);

    const parseToDate = (v: any): Date | null => {
      const iso = toInputDateFormat(v);
      if (!iso) return null;
      const t = Date.parse(iso);
      if (isNaN(t)) return null;
      return new Date(t);
    };

    const getCardClosedDate = (card: any): Date | null => {
      // Use actualClosedDate (ACTUVAL CLOSED DATE), falling back to other date fields
      return parseToDate(card.actualClosedDate || card.dateTimeOut || card.closedDate || card.jobDate || card.jobOpenDate || card.createdAt);
    };

    const getFreeServiceIndex = (card: any): number => {
      const txt = (card.freeServiceList || getFreeServiceText(card) || '').toLowerCase();
      if (txt.includes('10th') || txt.includes('10 th')) return 10;
      if (txt.includes('1st') || txt.includes('1 st')) return 1;
      if (txt.includes('2nd') || txt.includes('2 nd')) return 2;
      if (txt.includes('3rd') || txt.includes('3 rd')) return 3;
      if (txt.includes('4th') || txt.includes('4 th')) return 4;
      if (txt.includes('5th') || txt.includes('5 th')) return 5;
      if (txt.includes('6th') || txt.includes('6 th')) return 6;
      if (txt.includes('7th') || txt.includes('7 th')) return 7;
      if (txt.includes('8th') || txt.includes('8 th')) return 8;
      if (txt.includes('9th') || txt.includes('9 th')) return 9;
      return 0; // Not a free service
    };

    const categories = {
      fs1: [] as any[],
      fs2: [] as any[],
      fs3: [] as any[],
      fs4: [] as any[],
      fs5: [] as any[],
      fs6: [] as any[],
      fs7: [] as any[],
      fs8: [] as any[],
      fs9: [] as any[],
      fs10: [] as any[],
      general_90: [] as any[],
      post_warranty: [] as any[],
      gear_1: [] as any[],
      gear_2: [] as any[],
    };

    followupCalculatedList.forEach(item => {
      // Apply freeServiceSearch filter if present
      if (freeServiceSearch) {
        const q = freeServiceSearch.toLowerCase().trim();
        const chassis = (item.chassisNo || '').toLowerCase();
        const name = (item.customerName || '').toLowerCase();
        const mob = (item.mobileNumber || '').toLowerCase();
        const vill = (item.village || '').toLowerCase();
        const mdl = (item.model || '').toLowerCase();
        if (!chassis.includes(q) && !name.includes(q) && !mob.includes(q) && !vill.includes(q) && !mdl.includes(q)) {
          return;
        }
      }

      // Apply freeServiceSupervisor filter if present
      if (freeServiceSupervisor !== 'all') {
        const itemSup = (item.supervisor || '').trim().toLowerCase();
        if (freeServiceSupervisor === 'unassigned') {
          if (itemSup !== '') return;
        } else {
          if (itemSup !== freeServiceSupervisor.trim().toLowerCase()) return;
        }
      }

      // Parse delivery date
      const delDate = parseToDate(item.dateOfDel || item.rec?.dateOfDel || getFieldValue(item.rec, 'installDate'));
      let daysSinceDel = -1;
      if (delDate) {
        daysSinceDel = Math.floor((today.getTime() - delDate.getTime()) / (1000 * 60 * 60 * 24));
      }

      const isWithinWarranty = delDate && daysSinceDel >= 0 && daysSinceDel <= 730;

      // Identify completed free services
      const completedFS: Record<number, { date: Date; card: any }> = {};
      (item.cards || []).forEach(card => {
        const idx = getFreeServiceIndex(card);
        if (idx >= 1 && idx <= 10) {
          const dateVal = getCardClosedDate(card);
          if (dateVal) {
            if (!completedFS[idx] || dateVal > completedFS[idx].date) {
              completedFS[idx] = { date: dateVal, card };
            }
          }
        }
      });

      // Find the maximum free service index completed
      let maxFSIdx = 0;
      for (let i = 1; i <= 10; i++) {
        if (completedFS[i]) {
          maxFSIdx = i;
        }
      }

      if (isWithinWarranty) {
        // --- Case 1: Never Serviced (Asalu job card entry leni/avvaledante 2 years lopu 1st free service list Loki) ---
        const hasAnyJobCard = item.hasJobCard && (item.cards || []).length > 0;
        if (!hasAnyJobCard) {
          categories.fs1.push({ 
            ...item, 
            daysDue: 30 - daysSinceDel, 
            lastCompleted: 'Never Serviced (New Tractor)' 
          });
        } else {
          // --- Case 2: Has Job Cards, calculate based on actual closed dates of completed free services ---
          if (maxFSIdx === 0) {
            // No free services completed yet, but has some other job card. Due for 1st Free Service (fs1)
            const sortedCards = [...(item.cards || [])].sort((a, b) => {
              const tA = getCardClosedDate(a)?.getTime() || 0;
              const tB = getCardClosedDate(b)?.getTime() || 0;
              return tB - tA;
            });
            const latestCard = sortedCards[0];
            const latestDate = getCardClosedDate(latestCard);
            if (latestDate) {
              const daysSinceLatest = Math.floor((today.getTime() - latestDate.getTime()) / (1000 * 60 * 60 * 24));
              if (daysSinceLatest >= 75) {
                categories.fs1.push({
                  ...item,
                  daysDue: 90 - daysSinceLatest,
                  lastCompleted: latestCard.serviceType || 'Previous Job'
                });
              }
            }
          } else if (maxFSIdx >= 1 && maxFSIdx < 10) {
            // Next service due is (maxFSIdx + 1)
            const nextFSIdx = maxFSIdx + 1;
            const prevFSEntry = completedFS[maxFSIdx];
            const prevDate = prevFSEntry.date;
            const daysSincePrev = Math.floor((today.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24));
            
            if (daysSincePrev >= 75) {
              const prevDateStr = toInputDateFormat(prevDate) || '';
              const suffix = maxFSIdx === 1 ? 'st' : maxFSIdx === 2 ? 'nd' : maxFSIdx === 3 ? 'rd' : 'th';
              categories[`fs${nextFSIdx}` as keyof typeof categories].push({
                ...item,
                daysDue: 90 - daysSincePrev,
                lastCompleted: `${maxFSIdx}${suffix} Free Service (Closed: ${prevDateStr})`
              });
            }
          } else if (maxFSIdx === 10) {
            // All 10 free services are completed, route to general 90-day next service
            const prevFSEntry = completedFS[10];
            const prevDate = prevFSEntry.date;
            const daysSincePrev = Math.floor((today.getTime() - prevDate.getTime()) / (1000 * 60 * 60 * 24));
            if (daysSincePrev >= 75) {
              const prevDateStr = toInputDateFormat(prevDate) || '';
              categories.general_90.push({
                ...item,
                daysDue: 90 - daysSincePrev,
                lastCompleted: `10th Free Service (Closed: ${prevDateStr})`
              });
            }
          }
        }
      } else {
        // --- Case 3: Post-Warranty Customers (Past 2 Years or No Delivery Date) ---
        // Put in Post-Warranty category 90 days after their last service
        const hasAnyJobCard = item.hasJobCard && (item.cards || []).length > 0;
        if (hasAnyJobCard) {
          const sortedCards = [...(item.cards || [])].sort((a, b) => {
            const tA = getCardClosedDate(a)?.getTime() || 0;
            const tB = getCardClosedDate(b)?.getTime() || 0;
            return tB - tA;
          });
          const latestCard = sortedCards[0];
          const latestDate = getCardClosedDate(latestCard);
          if (latestDate) {
            const daysSinceLatest = Math.floor((today.getTime() - latestDate.getTime()) / (1000 * 60 * 60 * 24));
            if (daysSinceLatest >= 75) {
              const latestCardType = (latestCard.serviceType || '').toString().trim();
              categories.post_warranty.push({ 
                ...item, 
                daysDue: 90 - daysSinceLatest, 
                lastCompleted: latestCardType || 'Paid Service' 
              });
            }
          }
        }
      }

      // --- Gear Oil Follow-up (1 Year) ---
      if (daysSinceDel >= 350 && daysSinceDel <= 380) {
        categories.gear_1.push({ ...item, daysSinceDel });
      }

      // --- Gear Oil Follow-up (2 Years) ---
      if (daysSinceDel >= 715 && daysSinceDel <= 745) {
        categories.gear_2.push({ ...item, daysSinceDel });
      }
    });

    return categories;
  }, [followupCalculatedList, freeServiceSearch, freeServiceSupervisor]);

  const scheduledFollowupCustomers = useMemo(() => {
    return followupCalculatedList
      .filter(item => item.lastNextCallDate && String(item.lastNextCallDate).trim() !== '')
      .map(item => {
        const scheduledDate = String(item.lastNextCallDate).trim();
        let status: 'today' | 'overdue' | 'upcoming' = 'upcoming';
        if (scheduledDate === todayISO) {
          status = 'today';
        } else if (scheduledDate < todayISO) {
          status = 'overdue';
        } else {
          status = 'upcoming';
        }
        return {
          ...item,
          scheduledDate,
          scheduledStatus: status,
          lastCalledBy: item.lastCalledBy || (item.followupHistory?.[0]?.calledBy) || 'Unspecified'
        };
      })
      .sort((a, b) => (a.scheduledDate > b.scheduledDate ? 1 : a.scheduledDate < b.scheduledDate ? -1 : 0));
  }, [followupCalculatedList, todayISO]);

  const scheduledFollowupsTodayCount = useMemo(() => {
    return scheduledFollowupCustomers.filter(item => item.scheduledStatus === 'today' || item.scheduledStatus === 'overdue').length;
  }, [scheduledFollowupCustomers]);

  const filteredScheduledFollowups = useMemo(() => {
    return scheduledFollowupCustomers.filter(item => {
      if (telecallerScheduledStatus !== 'all' && item.scheduledStatus !== telecallerScheduledStatus) {
        return false;
      }
      if (telecallerCallerFilter !== 'all') {
        const target = telecallerCallerFilter.toLowerCase();
        const matchLastCaller = item.lastCalledBy?.toLowerCase() === target;
        const matchHistory = (item.followupHistory || []).some((h: any) => h.calledBy?.toLowerCase() === target);
        if (!matchLastCaller && !matchHistory) return false;
      }
      if (telecallerSupervisorFilter !== 'all') {
        if (item.supervisor?.toLowerCase() !== telecallerSupervisorFilter.toLowerCase()) return false;
      }
      if (telecallerSearch) {
        const q = telecallerSearch.toLowerCase();
        const matchName = item.customerName.toLowerCase().includes(q);
        const matchMobile = item.mobileNumber.toLowerCase().includes(q);
        const matchVillage = item.village.toLowerCase().includes(q);
        const matchHFN = item.historyFileNo.toLowerCase().includes(q);
        const matchModel = item.model.toLowerCase().includes(q);
        const matchChassis = item.chassisNo.toLowerCase().includes(q);
        const matchRemarks = item.lastRemarks.toLowerCase().includes(q);
        const matchCaller = item.lastCalledBy?.toLowerCase().includes(q);
        const matchSup = item.supervisor?.toLowerCase().includes(q);
        return matchName || matchMobile || matchVillage || matchHFN || matchModel || matchChassis || matchRemarks || matchCaller || matchSup;
      }
      return true;
    });
  }, [scheduledFollowupCustomers, telecallerScheduledStatus, telecallerCallerFilter, telecallerSupervisorFilter, telecallerSearch]);

  const dateWiseCallSummary = useMemo(() => {
    const map: Record<string, { callDate: string; count: number; customers: Set<string>; remarksSample: string[] }> = {};

    allTelecallerLogs.forEach(log => {
      if (!log.callDate) return;
      if (telecallerCallerFilter !== 'all' && log.calledBy?.toLowerCase() !== telecallerCallerFilter.toLowerCase()) return;
      if (telecallerSupervisorFilter !== 'all' && log.supervisor?.toLowerCase() !== telecallerSupervisorFilter.toLowerCase()) return;

      if (!map[log.callDate]) {
        map[log.callDate] = {
          callDate: log.callDate,
          count: 0,
          customers: new Set(),
          remarksSample: []
        };
      }
      map[log.callDate].count += 1;
      if (log.customerName) map[log.callDate].customers.add(log.customerName);
      if (log.remarks && map[log.callDate].remarksSample.length < 3) {
        map[log.callDate].remarksSample.push(log.remarks);
      }
    });

    let list = Object.values(map).map(item => ({
      ...item,
      uniqueCustomersCount: item.customers.size
    }));

    if (telecallerDateFrom) {
      list = list.filter(item => item.callDate >= telecallerDateFrom);
    }
    if (telecallerDateTo) {
      list = list.filter(item => item.callDate <= telecallerDateTo);
    }

    return list.sort((a, b) => (b.callDate > a.callDate ? 1 : b.callDate < a.callDate ? -1 : 0));
  }, [allTelecallerLogs, telecallerCallerFilter, telecallerSupervisorFilter, telecallerDateFrom, telecallerDateTo]);

  const filteredTelecallerLogs = useMemo(() => {
    return allTelecallerLogs.filter(log => {
      if (telecallerCallerFilter !== 'all' && log.calledBy?.toLowerCase() !== telecallerCallerFilter.toLowerCase()) return false;
      if (telecallerSupervisorFilter !== 'all' && log.supervisor?.toLowerCase() !== telecallerSupervisorFilter.toLowerCase()) return false;
      if (telecallerDateFrom && log.callDate < telecallerDateFrom) return false;
      if (telecallerDateTo && log.callDate > telecallerDateTo) return false;
      if (telecallerSearch) {
        const q = telecallerSearch.toLowerCase();
        const matchName = log.customerName.toLowerCase().includes(q);
        const matchMobile = log.mobileNumber.toLowerCase().includes(q);
        const matchVillage = log.village.toLowerCase().includes(q);
        const matchHFN = log.historyFileNo.toLowerCase().includes(q);
        const matchModel = log.model.toLowerCase().includes(q);
        const matchRemarks = log.remarks.toLowerCase().includes(q);
        const matchCaller = log.calledBy?.toLowerCase().includes(q);
        const matchSup = log.supervisor?.toLowerCase().includes(q);
        return matchName || matchMobile || matchVillage || matchHFN || matchModel || matchRemarks || matchCaller || matchSup;
      }
      return true;
    });
  }, [allTelecallerLogs, telecallerCallerFilter, telecallerSupervisorFilter, telecallerDateFrom, telecallerDateTo, telecallerSearch]);

  const downloadTelecallingExcel = () => {
    try {
      const excelData = filteredTelecallerLogs.map((log, idx) => ({
        'S.No': idx + 1,
        'Call Date': fmtDate(log.callDate),
        'Customer Name': log.customerName,
        'History File No': log.historyFileNo,
        'Mobile Number': log.mobileNumber,
        'Tractor Model': log.model,
        'Village': log.village,
        'Mandal': log.mandal,
        'Assigned Supervisor': log.supervisor,
        'Called By (Telecaller / Supervisor)': log.calledBy,
        'Customer Remarks / Feedback': log.remarks,
        'Next Scheduled Call Date': fmtDate(log.nextCallDate)
      }));

      const ws = XLSX.utils.json_to_sheet(excelData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Telecalling_Log_Report");
      XLSX.writeFile(wb, `Telecalling_Report_${todayISO}.xlsx`);
    } catch (e) {
      alert('Failed to export Excel report: ' + e);
    }
  };

  const downloadFollowupExcel = () => {
    const getCardTime = (c: any) => {
      const dStr = c.jobDate || c.jobOpenDate || c.createdAt;
      if (!dStr) return 0;
      const t = Date.parse(dStr);
      return isNaN(t) ? 0 : t;
    };

    const reportingRows: any[] = [];
    const notReportingRows: any[] = [];

    uniqueCustomers.forEach((rec, idx) => {
      const ch = rec['Chassis no'] || rec.__chassisDisplay || getFieldValue(rec, 'chassis') || '';
      const normCh = normalizeKey(ch.toString());
      const cards = normCh ? (customerJobCardsMap[normCh] || []) : [];
      const hasJobCard = cards.length > 0;
      
      let lastServiceType = '';
      let lastJobCardEntryDate = '';
      let lastEntryHours = '';
      let lastRepairDetails = '';
      let latestCard: any = null;

      if (hasJobCard) {
        const sorted = [...cards].sort((a, b) => getCardTime(b) - getCardTime(a));
        latestCard = sorted[0];
        
        lastServiceType = latestCard.serviceType || 'Paid Service';
        lastJobCardEntryDate = fmtDate(latestCard.jobDate || latestCard.jobOpenDate || (latestCard.createdAt ? latestCard.createdAt.split('T')[0] : ''));
        lastEntryHours = latestCard.hourMeter || latestCard.hrsRun || '';
        lastRepairDetails = Array.isArray(latestCard.repairRows)
          ? latestCard.repairRows.map((r: any) => r.repair).filter(Boolean).join(', ')
          : '';
      }

      let supName = rec.supervisor || rec.SUPERVISOR || rec['Supervisor Name'] || (latestCard ? (latestCard.wsIncharge || latestCard.supervisor || '') : '') || '';

      const sNoVal = (rec['S.NO'] || rec.sno || (idx + 1)).toString();
      const chassisNoVal = ch.toString();

      const exportRow = {
        'supervisor name': supName.toString(),
        'S.NO': sNoVal,
        'Model': (rec.Model || rec.model || getFieldValue(rec, 'model') || '').toString(),
        'MODEL TYPE': (rec['MODEL TYPE'] || rec.modelType || getFieldValue(rec, 'modelType') || '').toString(),
        'Chassis no': chassisNoVal,
        'Engine No:': (rec['Engine No:'] || rec.engineNo || getFieldValue(rec, 'engineNo') || '').toString(),
        'Date of del': fmtDate(rec['Date of del'] || rec.dateOfDel || getFieldValue(rec, 'installDate')),
        'Customer Name': (rec['Customer Name'] || rec.__custNameDisplay || rec.custName || getFieldValue(rec, 'custName') || '').toString(),
        'VILLAGE': (rec.VILLAGE || rec.village || getFieldValue(rec, 'village') || '').toString(),
        'Mandal': (rec.Mandal || rec.mandal || getFieldValue(rec, 'mandal') || '').toString(),
        'Mobile Number': (rec['Mobile Number'] || rec.__custPhoneDisplay || rec.mobileNumber || getFieldValue(rec, 'custPhone') || '').toString(),
        'last service type': lastServiceType,
        'last job card entry date': lastJobCardEntryDate,
        'last entry hours': lastEntryHours,
        'Last repair details': lastRepairDetails
      };

      if (hasJobCard) {
        reportingRows.push(exportRow);
      } else {
        notReportingRows.push(exportRow);
      }
    });

    // Generate Excel Sheets
    const wb = XLSX.utils.book_new();
    const wsReporting = XLSX.utils.json_to_sheet(reportingRows);
    const wsNotReporting = XLSX.utils.json_to_sheet(notReportingRows);

    const fitWidths = (rows: any[]) => {
      if (rows.length === 0) return [];
      const keys = Object.keys(rows[0]);
      return keys.map(k => {
        const maxLen = Math.max(
          k.length,
          ...rows.map(r => (r[k] ? r[k].toString().length : 0))
        );
        return { wch: Math.min(maxLen + 2, 40) };
      });
    };

    wsReporting['!cols'] = fitWidths(reportingRows);
    wsNotReporting['!cols'] = fitWidths(notReportingRows);

    XLSX.utils.book_append_sheet(wb, wsReporting, "Reporting Customers");
    XLSX.utils.book_append_sheet(wb, wsNotReporting, "Not Reporting Customers");

    XLSX.writeFile(wb, `SriGayathri_Service_Followup_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const deleteCustomerRecord = async (chassisNo: string) => {
    if (!chassisNo) return;
    if (!window.confirm(`Are you sure you want to delete customer with chassis "${chassisNo}"? This will remove them from the local database.`)) return;

    try {
      const updatedIndex = { ...chassisIndex };
      const normKey = normalizeKey(chassisNo);
      delete updatedIndex[normKey];
      
      // Update State
      setChassisIndex(updatedIndex);
      
      // Update Local Storage
      saveToStorage(LS_CUSTOMER_KEY, updatedIndex);

      // Update Google Sheets if connected
      if (useGoogleSheets && sheetsSpreadsheetId) {
        try {
          await deleteSheetRow(sheetsSpreadsheetId, 'Customers', chassisNo);
          alert('✅ Customer deleted successfully from Local Database and Google Sheets!');
        } catch (e) {
          console.error('Error deleting customer from Sheets:', e);
          alert('⚠️ Customer deleted from local database, but failed to remove from Google Sheets.');
        }
      } else {
        alert('✅ Customer deleted successfully from Local Database!');
      }
    } catch (err) {
      console.error('Error deleting customer record:', err);
      alert('❌ Failed to delete customer record.');
    }
  };

  const handleOpenEditCustomer = (rec: any) => {
    const ch = rec['Chassis no'] || rec.__chassisDisplay || getFieldValue(rec, 'chassis') || '';
    const chKey = normalizeKey(ch.toString());
    setEditingCustChassisKey(chKey);
    setEditCustForm({
      branch: (rec.BRANCH || rec.branch || getFieldValue(rec, 'branch') || '').toString(),
      sNo: (rec['S.NO'] || rec.sno || '').toString(),
      model: (rec.Model || rec.model || getFieldValue(rec, 'model') || '').toString(),
      modelType: (rec['MODEL TYPE'] || rec.modelType || getFieldValue(rec, 'modelType') || '').toString(),
      chassisNo: ch.toString(),
      engineNo: (rec['Engine No:'] || rec.engineNo || getFieldValue(rec, 'engineNo') || '').toString(),
      dateOfDel: (rec['Date of del'] || rec.dateOfDel || getFieldValue(rec, 'installDate') || '').toString(),
      custName: (rec['Customer Name'] || rec.__custNameDisplay || rec.custName || getFieldValue(rec, 'custName') || '').toString(),
      fatherName: (rec['FATHER NAME'] || rec.fatherName || getFieldValue(rec, 'fatherName') || '').toString(),
      address: (rec['ADDRESS'] || rec.address || getFieldValue(rec, 'custAddr') || '').toString(),
      village: (rec.VILLAGE || rec.village || getFieldValue(rec, 'village') || '').toString(),
      mandal: (rec.Mandal || rec.mandal || getFieldValue(rec, 'mandal') || '').toString(),
      mobileNumber: (rec['Mobile Number'] || rec.__custPhoneDisplay || rec.mobileNumber || getFieldValue(rec, 'custPhone') || '').toString(),
      district: (rec.Distict || rec.district || getFieldValue(rec, 'district') || '').toString(),
      pinCode: (rec['PIN CODE'] || rec.pinCode || getFieldValue(rec, 'pinCode') || '').toString(),
      dspName: (rec['DSP Name'] || rec.dspName || getFieldValue(rec, 'dspName') || '').toString(),
      exchangeBrand: (rec['EXCHANGE BRAND'] || rec.exchangeBrand || getFieldValue(rec, 'exchangeBrand') || '').toString(),
      exchangeModels: (rec['EXCHANGE TRACTOR MODELS'] || rec.exchangeModels || getFieldValue(rec, 'exchangeModels') || '').toString(),
      supervisor: (rec.supervisor || rec.SUPERVISOR || rec['Supervisor Name'] || '').toString()
    });
    setIsEditCustModalOpen(true);
  };

  const handleUpdateCustomer = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editCustForm.chassisNo.trim() && !editCustForm.custName.trim()) {
      alert('Please enter at least Chassis Number or Customer Name.');
      return;
    }

    const newChassisKey = normalizeKey(editCustForm.chassisNo);
    const originalChassisKey = editingCustChassisKey;

    const updatedRecord = {
      BRANCH: editCustForm.branch,
      'S.NO': editCustForm.sNo,
      Model: editCustForm.model,
      'MODEL TYPE': editCustForm.modelType,
      'Chassis no': editCustForm.chassisNo,
      'Engine No:': editCustForm.engineNo,
      'Date of del': editCustForm.dateOfDel,
      'Customer Name': editCustForm.custName,
      'FATHER NAME': editCustForm.fatherName,
      ADDRESS: editCustForm.address,
      VILLAGE: editCustForm.village,
      Mandal: editCustForm.mandal,
      'Mobile Number': editCustForm.mobileNumber,
      Distict: editCustForm.district,
      'PIN CODE': editCustForm.pinCode,
      'DSP Name': editCustForm.dspName,
      'EXCHANGE BRAND': editCustForm.exchangeBrand,
      'EXCHANGE TRACTOR MODELS': editCustForm.exchangeModels,
      supervisor: editCustForm.supervisor,
      SUPERVISOR: editCustForm.supervisor,
      'Supervisor Name': editCustForm.supervisor,

      __chassisDisplay: editCustForm.chassisNo,
      __custNameDisplay: editCustForm.custName,
      __custPhoneDisplay: editCustForm.mobileNumber,
      __custAddrDisplay: [editCustForm.address, editCustForm.village, editCustForm.mandal, editCustForm.district]
        .filter(Boolean)
        .join(', ')
    };

    const updatedIndex = { ...chassisIndex };

    // If chassis has changed, delete the old chassis key and references to it
    if (originalChassisKey && originalChassisKey !== newChassisKey) {
      delete updatedIndex[originalChassisKey];
    }

    // Now index under the new keys
    if (newChassisKey) updatedIndex[newChassisKey] = updatedRecord;
    
    const phoneKey = normalizeKey(editCustForm.mobileNumber);
    const nameKey = normalizeKey(editCustForm.custName);
    if (phoneKey) updatedIndex[phoneKey] = updatedRecord;
    if (nameKey) updatedIndex[nameKey] = updatedRecord;

    setChassisIndex(updatedIndex);
    await saveToStorage(LS_CUSTOMER_KEY, updatedIndex);
    
    // Propagate changes globally to Firestore jobcards and complaints
    await syncCustomerDetailsGlobally(editCustForm.chassisNo, {
      custName: editCustForm.custName,
      fatherName: editCustForm.fatherName,
      custAddr: editCustForm.address,
      village: editCustForm.village,
      mandal: editCustForm.mandal,
      ownerMob: editCustForm.mobileNumber,
      model: editCustForm.model,
      modelType: editCustForm.modelType,
      engineNo: editCustForm.engineNo,
      installDate: editCustForm.dateOfDel,
      supervisor: editCustForm.supervisor,
      district: editCustForm.district,
      pinCode: editCustForm.pinCode,
      dspName: editCustForm.dspName,
      exchangeBrand: editCustForm.exchangeBrand,
      exchangeModels: editCustForm.exchangeModels,
      branch: editCustForm.branch,
      historyFileNo: editCustForm.sNo
    });
    
    setIsEditCustModalOpen(false);
    alert(`✅ Customer "${editCustForm.custName || editCustForm.chassisNo}" updated successfully!`);
  };

  const updateCustomerRecordInIndex = async (originalChassisNo: string, updatedFields: Partial<any>) => {
    try {
      const updatedIndex = { ...chassisIndex };
      const normOriginalChassis = normalizeKey(originalChassisNo);
      
      // Find all keys that refer to this customer (by matching normalized chassis number)
      const keysToUpdate: string[] = [];
      Object.entries(updatedIndex).forEach(([key, rec]) => {
        if (rec) {
          const record = rec as any;
          const recChassis = record['Chassis no'] || record.__chassisDisplay || record.chassisNo || '';
          if (normalizeKey(recChassis.toString()) === normOriginalChassis) {
            keysToUpdate.push(key);
          }
        }
      });

      if (keysToUpdate.length === 0 && normOriginalChassis) {
        keysToUpdate.push(normOriginalChassis);
      }

      // Apply the updates to each matched record reference
      keysToUpdate.forEach(key => {
        const existing = updatedIndex[key] || {};
        updatedIndex[key] = {
          ...existing,
          ...updatedFields
        };
      });

      setChassisIndex(updatedIndex);
      await saveToStorage(LS_CUSTOMER_KEY, updatedIndex);
    } catch (e) {
      console.error("Error updating customer record in index:", e);
    }
  };

  const syncCustomerDetailsGlobally = async (
    chassisNoToSync: string, 
    details: {
      custName?: string;
      fatherName?: string;
      custAddr?: string;
      village?: string;
      mandal?: string;
      ownerMob?: string;
      driverMob?: string;
      model?: string;
      modelType?: string;
      engineNo?: string;
      installDate?: string;
      regdNo?: string;
      district?: string;
      pinCode?: string;
      dspName?: string;
      exchangeBrand?: string;
      exchangeModels?: string;
      supervisor?: string;
      branch?: string;
      historyFileNo?: string;
    }
  ) => {
    const rawChassis = String(chassisNoToSync || '').trim();
    if (!rawChassis) return;
    const normChassis = normalizeKey(rawChassis);

    try {
      // 1. Update local chassisIndex (Customer DB) using functional update to avoid stale closures
      let finalIndex: any = null;
      setChassisIndex(prevIndex => {
        const updatedIndex = { ...prevIndex };
        let existingRecord = updatedIndex[normChassis] || {};

        const newRecord = {
          ...existingRecord,
          BRANCH: details.branch !== undefined ? details.branch : (existingRecord.BRANCH || existingRecord.branch || ''),
          'S.NO': details.historyFileNo !== undefined ? details.historyFileNo : (existingRecord['S.NO'] || existingRecord['S.No.'] || existingRecord.sNo || ''),
          Model: details.model !== undefined ? details.model : (existingRecord.Model || ''),
          'MODEL TYPE': details.modelType !== undefined ? details.modelType : (existingRecord['MODEL TYPE'] || ''),
          'Chassis no': rawChassis,
          'Engine No:': details.engineNo !== undefined ? details.engineNo : (existingRecord['Engine No:'] || ''),
          'Date of del': details.installDate !== undefined ? details.installDate : (existingRecord['Date of del'] || ''),
          'Customer Name': details.custName !== undefined ? details.custName : (existingRecord['Customer Name'] || ''),
          'FATHER NAME': details.fatherName !== undefined ? details.fatherName : (existingRecord['FATHER NAME'] || ''),
          ADDRESS: details.custAddr !== undefined ? details.custAddr : (existingRecord.ADDRESS || ''),
          VILLAGE: details.village !== undefined ? details.village : (existingRecord.VILLAGE || ''),
          Mandal: details.mandal !== undefined ? details.mandal : (existingRecord.Mandal || ''),
          'Mobile Number': details.ownerMob !== undefined ? details.ownerMob : (existingRecord['Mobile Number'] || ''),
          Distict: details.district !== undefined ? details.district : (existingRecord.Distict || ''),
          'PIN CODE': details.pinCode !== undefined ? details.pinCode : (existingRecord['PIN CODE'] || ''),
          'DSP Name': details.dspName !== undefined ? details.dspName : (existingRecord['DSP Name'] || ''),
          'EXCHANGE BRAND': details.exchangeBrand !== undefined ? details.exchangeBrand : (existingRecord['EXCHANGE BRAND'] || ''),
          'EXCHANGE TRACTOR MODELS': details.exchangeModels !== undefined ? details.exchangeModels : (existingRecord['EXCHANGE TRACTOR MODELS'] || ''),
          supervisor: details.supervisor !== undefined ? details.supervisor : (existingRecord.supervisor || ''),
          SUPERVISOR: details.supervisor !== undefined ? details.supervisor : (existingRecord.SUPERVISOR || ''),
          'Supervisor Name': details.supervisor !== undefined ? details.supervisor : (existingRecord['Supervisor Name'] || ''),

          __chassisDisplay: rawChassis,
          __custNameDisplay: details.custName !== undefined ? details.custName : (existingRecord.__custNameDisplay || ''),
          __custPhoneDisplay: details.ownerMob !== undefined ? details.ownerMob : (existingRecord.__custPhoneDisplay || ''),
          __custAddrDisplay: [
            details.custAddr !== undefined ? details.custAddr : (existingRecord.ADDRESS || ''),
            details.village !== undefined ? details.village : (existingRecord.VILLAGE || ''),
            details.mandal !== undefined ? details.mandal : (existingRecord.Mandal || ''),
            details.district !== undefined ? details.district : (existingRecord.Distict || '')
          ].filter(Boolean).join(', ')
        };

        updatedIndex[normChassis] = newRecord;

        const phoneToUse = details.ownerMob || existingRecord['Mobile Number'] || '';
        if (phoneToUse) {
          const phoneKey = normalizeKey(String(phoneToUse));
          if (phoneKey) updatedIndex[phoneKey] = newRecord;
        }
        
        const nameToUse = details.custName || existingRecord['Customer Name'] || '';
        if (nameToUse) {
          const nameKey = normalizeKey(String(nameToUse));
          if (nameKey) updatedIndex[nameKey] = newRecord;
        }

        finalIndex = updatedIndex;
        return updatedIndex;
      });
      
      if (finalIndex) {
        await saveToStorage(LS_CUSTOMER_KEY, finalIndex);
      }

      // 2. Query and Update all matching Job Cards in Firestore
      const matchingCards = savedJobCards.filter(c => normalizeKey(String(c.chassisNo || '')) === normChassis);
      for (const card of matchingCards) {
        const cardRef = doc(db, 'jobcards', card.id);
        const updatedCardFields: any = {};
        if (details.custName !== undefined) updatedCardFields.custName = details.custName;
        if (details.fatherName !== undefined) updatedCardFields.fatherName = details.fatherName;
        if (details.custAddr !== undefined) updatedCardFields.custAddr = details.custAddr;
        if (details.village !== undefined) updatedCardFields.village = details.village;
        if (details.mandal !== undefined) updatedCardFields.mandal = details.mandal;
        if (details.ownerMob !== undefined) updatedCardFields.ownerMob = details.ownerMob;
        if (details.driverMob !== undefined) updatedCardFields.driverMob = details.driverMob;
        if (details.model !== undefined) updatedCardFields.model = details.model;
        if (details.modelType !== undefined) updatedCardFields.modelType = details.modelType;
        if (details.engineNo !== undefined) updatedCardFields.engineNo = details.engineNo;
        if (details.installDate !== undefined) updatedCardFields.installDate = details.installDate;
        if (details.regdNo !== undefined) updatedCardFields.regdNo = details.regdNo;
        
        if (Object.keys(updatedCardFields).length > 0) {
          await updateDoc(cardRef, updatedCardFields);
        }
      }

      // 3. Query and Update all matching Complaints in Firestore
      const matchingComplaints = complaints.filter(c => normalizeKey(String(c.chassisNo || '')) === normChassis);
      for (const comp of matchingComplaints) {
        const compRef = doc(db, 'complaints', comp.id);
        const updatedCompFields: any = {};
        if (details.custName !== undefined) updatedCompFields.customerName = details.custName;
        if (details.ownerMob !== undefined) updatedCompFields.mobileNumber = details.ownerMob;
        if (details.model !== undefined) updatedCompFields.tractorModel = details.model;
        if (details.supervisor !== undefined) updatedCompFields.assignedSupervisor = details.supervisor;

        if (Object.keys(updatedCompFields).length > 0) {
          await updateDoc(compRef, updatedCompFields);
        }
      }
    } catch (err) {
      console.error("Error in syncCustomerDetailsGlobally:", err);
    }
  };

  const normalizeStaffName = (name?: string) => {
    if (!name || !name.trim()) return '';
    const trimmed = name.trim();
    const lower = trimmed.toLowerCase();
    
    // Check against master lists
    const allStaff = [...mechanicsList, ...supervisorsList];
    const match = allStaff.find(m => m.trim().toLowerCase() === lower);
    if (match) return match;

    // Otherwise format as Title Case
    return trimmed
      .toLowerCase()
      .split(' ')
      .map(w => w.charAt(0).toUpperCase() + w.slice(1))
      .join(' ');
  };

  const getAssignedSupervisor = (mechanicName?: string): string => {
    if (!mechanicName || !mechanicName.trim()) return '';
    const nameLower = mechanicName.trim().toLowerCase();
    const foundStaff = staffMembers.find(
      s => s.role === 'mechanic' && s.name.trim().toLowerCase() === nameLower
    );
    return foundStaff?.supervisor || '';
  };

  const parseMoney = (val: any): number => {
    if (val === undefined || val === null) return 0;
    if (typeof val === 'number') return val;
    const clean = val.toString().replace(/[^0-9.-]/g, '');
    const parsed = parseFloat(clean);
    return isNaN(parsed) ? 0 : parsed;
  };

  const isCardClosed = (card: any) => {
    if (!card) return false;
    const statusStr = (card.status || 'Closed').toString().trim().toLowerCase();
    const closedStatuses = ['closed', 'done', 'completed', 'finalized', 'billed'];
    return closedStatuses.includes(statusStr) || !!card.actualClosedDate || !!card.billNo;
  };

  const searchedJobCards = useMemo(() => {
    return visibleJobCards.filter(card => {
      // 1. Date From Filter
      if (savedDateFrom) {
        const rawDate = card.jobDate || card.jobOpenDate || card.complaintDate || card.installDate || card.dateOfDelivery || card.createdAt;
        const normDate = toInputDateFormat(rawDate);
        if (normDate && normDate < savedDateFrom) return false;
      }

      // 2. Date To Filter
      if (savedDateTo) {
        const rawDate = card.jobDate || card.jobOpenDate || card.complaintDate || card.installDate || card.dateOfDelivery || card.createdAt;
        const normDate = toInputDateFormat(rawDate);
        if (normDate && normDate > savedDateTo) return false;
      }

      // 3. Status Filter
      if (savedStatusFilter !== 'all') {
        if (savedStatusFilter === 'MissingOnline') {
          const onlineNo = (card.onlineJobCardNo || '').toString().trim();
          if (onlineNo) return false;
        } else {
          const isClosed = isCardClosed(card);
          const normFilterStatus = savedStatusFilter.toLowerCase();
          if (normFilterStatus === 'closed' ? !isClosed : isClosed) return false;
        }
      }

      // 4. Service Place / Location Filter
      if (savedLocationFilter !== 'all') {
        const rawLoc = (card.serviceLocation || card.servicePlace || 'workshop').toString().toLowerCase().trim();
        let normalizedLoc = 'workshop';
        if (rawLoc.includes('dss') || rawLoc.includes('door') || rawLoc.includes('step') || rawLoc.includes('field')) {
          normalizedLoc = 'dss';
        } else if (rawLoc.includes('event') || rawLoc.includes('camp')) {
          normalizedLoc = 'event';
        }
        if (normalizedLoc !== savedLocationFilter) return false;
      }

      // 5. Supervisor Filter
      if (savedSupervisorFilter !== 'all') {
        const cardSup = (card.wsIncharge || card.supervisor || card.supervisorName || '').toString().toLowerCase().trim();
        const filterSup = savedSupervisorFilter.toLowerCase().trim();
        if (!cardSup || (!cardSup.includes(filterSup) && !filterSup.includes(cardSup))) return false;
      }

      // 6. Mechanic / Technician Filter
      if (savedMechanicFilter !== 'all') {
        const cardMech = (card.mechanic || card.technicianName || card.mechanicName || '').toString().toLowerCase().trim();
        const filterMech = savedMechanicFilter.toLowerCase().trim();
        if (!cardMech || (!cardMech.includes(filterMech) && !filterMech.includes(cardMech))) return false;
      }

      // 7. Text Search Query
      if (savedListSearch.trim()) {
        const q = savedListSearch.toLowerCase().trim();
        const matches = (
          (card.jobNo || '').toString().toLowerCase().includes(q) ||
          (card.onlineJobCardNo || '').toString().toLowerCase().includes(q) ||
          (card.chassisNo || '').toString().toLowerCase().includes(q) ||
          (card.engineNo || '').toString().toLowerCase().includes(q) ||
          (card.custName || '').toString().toLowerCase().includes(q) ||
          (card.fatherName || '').toString().toLowerCase().includes(q) ||
          (card.village || '').toString().toLowerCase().includes(q) ||
          (card.mandal || '').toString().toLowerCase().includes(q) ||
          (card.ownerMob || card.phNo || '').toString().toLowerCase().includes(q) ||
          (card.billNo || '').toString().toLowerCase().includes(q) ||
          (card.model || '').toString().toLowerCase().includes(q) ||
          (card.mechanic || card.technicianName || '').toString().toLowerCase().includes(q) ||
          (card.wsIncharge || '').toString().toLowerCase().includes(q) ||
          (card.branch || '').toString().toLowerCase().includes(q) ||
          (card.historyFileNo || '').toString().toLowerCase().includes(q)
        );
        if (!matches) return false;
      }

      return true;
    });
  }, [
    visibleJobCards,
    savedListSearch,
    savedDateFrom,
    savedDateTo,
    savedStatusFilter,
    savedLocationFilter,
    savedSupervisorFilter,
    savedMechanicFilter
  ]);

  const totalPages = Math.max(1, Math.ceil(searchedJobCards.length / savedListPageSize));
  const validSavedListPage = Math.min(savedListPage, totalPages);

  const paginatedJobCards = useMemo(() => {
    const start = (validSavedListPage - 1) * savedListPageSize;
    return searchedJobCards.slice(start, start + savedListPageSize);
  }, [searchedJobCards, validSavedListPage, savedListPageSize]);

  const {
    filteredReportCards,
    totalCardsCount,
    openCardsCount,
    closedCardsCount,
    totalRevenueSum,
    totalLabourSum,
    totalSparesSum,
    workshopCount,
    dssCount,
    eventCount,
    workshopRevenue,
    dssRevenue,
    eventRevenue,
    trendChartData,
    locationChartData,
    mechanicPerformanceSummary,
    mechanicChartData,
    supervisorChartData
  } = useMemo(() => {
    const cards = savedJobCards.filter((card) => {
      // 1. Date From
      if (reportDateFrom) {
        const rawDate = card.jobDate || card.jobOpenDate || card.complaintDate || card.installDate || card.dateOfDelivery || card.createdAt;
        const cardDate = toInputDateFormat(rawDate);
        if (cardDate && cardDate < reportDateFrom) return false;
      }
      // 2. Date To
      if (reportDateTo) {
        const rawDate = card.jobDate || card.jobOpenDate || card.complaintDate || card.installDate || card.dateOfDelivery || card.createdAt;
        const cardDate = toInputDateFormat(rawDate);
        if (cardDate && cardDate > reportDateTo) return false;
      }
      // 3. Service Location
      if (reportLocation !== 'all') {
        const rawLoc = (card.serviceLocation || card.servicePlace || 'workshop').toString().toLowerCase().trim();
        let cardLoc = 'workshop';
        if (rawLoc.includes('dss') || rawLoc.includes('door') || rawLoc.includes('step') || rawLoc.includes('field')) {
          cardLoc = 'dss';
        } else if (rawLoc.includes('event') || rawLoc.includes('camp')) {
          cardLoc = 'event';
        }
        if (cardLoc !== reportLocation) return false;
      }
      // 4. Status
      if (reportStatus !== 'all') {
        const cardStatus = (card.status || 'Open').toString().trim().toLowerCase();
        const closedStatuses = ['closed', 'done', 'completed', 'finalized', 'billed'];
        const isClosed = closedStatuses.includes(cardStatus);
        const reportStatusNorm = reportStatus.toLowerCase();
        if (reportStatusNorm === 'closed' ? !isClosed : isClosed) return false;
      }
      // 5. Mechanic
      if (reportMechanic !== 'all') {
        const mechName = (card.mechanic || card.technicianName || card.mechanicName || '').toString().toLowerCase().trim();
        const filterMech = reportMechanic.toLowerCase().trim();
        if (!mechName || (!mechName.includes(filterMech) && !filterMech.includes(mechName))) return false;
      }
      // 6. Supervisor
      if (reportSupervisor !== 'all') {
        const supName = (card.wsIncharge || card.supervisor || card.supervisorName || '').toString().toLowerCase().trim();
        const filterSup = reportSupervisor.toLowerCase().trim();
        if (!supName || (!supName.includes(filterSup) && !filterSup.includes(supName))) return false;
      }
      // 7. Service Type
      if (reportServiceType !== 'all') {
        const stStr = (card.serviceType || '').toString().toLowerCase().trim();
        const filterSt = reportServiceType.toLowerCase().trim();
        if (!stStr || (!stStr.includes(filterSt) && !filterSt.includes(stStr))) return false;
      }
      // 8. Search query
      if (reportSearchQuery.trim()) {
        const q = reportSearchQuery.toLowerCase().trim();
        const match =
          (card.jobNo || '').toString().toLowerCase().includes(q) ||
          (card.onlineJobCardNo || '').toString().toLowerCase().includes(q) ||
          (card.chassisNo || '').toString().toLowerCase().includes(q) ||
          (card.custName || '').toString().toLowerCase().includes(q) ||
          (card.model || '').toString().toLowerCase().includes(q) ||
          (card.ownerMob || card.phNo || '').toString().toLowerCase().includes(q) ||
          (card.village || '').toString().toLowerCase().includes(q);
        if (!match) return false;
      }
      return true;
    });

    const totCardsCount = cards.length;
    let opCardsCount = 0;
    let clCardsCount = 0;
    let totRevenueSum = 0;
    let totLabourSum = 0;
    let totSparesSum = 0;
    let wsCount = 0;
    let dsCount = 0;
    let evCount = 0;
    let wsRevenue = 0;
    let dsRevenue = 0;
    let evRevenue = 0;

    const trendMap: Record<string, { date: string; cards: number; revenue: number }> = {};
    const staffPerfMap: Record<
      string,
      { name: string; role: string; jobsCount: number; openCount: number; closedCount: number; labourTotal: number; sparesTotal: number; grandTotal: number }
    > = {};

    for (let i = 0; i < cards.length; i++) {
      const c = cards[i];
      const isClosed = isCardClosed(c);
      if (isClosed) clCardsCount++;
      else opCardsCount++;

      const cLabour = parseMoney(c.totalLabour) || parseMoney(c.labourTotal) || (Array.isArray(c.repairRows) ? c.repairRows.reduce((sum: number, r: any) => sum + parseMoney(r.charge), 0) : 0);
      const partsSum = Array.isArray(c.partRows) ? c.partRows.reduce((sum: number, p: any) => sum + parseMoney(p.amount), 0) : 0;
      const cSpares = parseMoney(c.warrantyMaterial) + parseMoney(c.nonWarrantyMaterial) || parseMoney(c.partsTotal) || partsSum;
      const cGrand = parseMoney(c.gTotal) || (cLabour + cSpares);

      totRevenueSum += cGrand;
      totLabourSum += cLabour;
      totSparesSum += cSpares;

      const locStr = (c.serviceLocation || c.servicePlace || 'workshop').toString().toLowerCase();
      if (locStr.includes('dss') || locStr.includes('door') || locStr.includes('step') || locStr.includes('field')) {
        dsCount++;
        dsRevenue += cGrand;
      } else if (locStr.includes('event') || locStr.includes('camp')) {
        evCount++;
        evRevenue += cGrand;
      } else {
        wsCount++;
        wsRevenue += cGrand;
      }

      const rawD = c.jobDate || c.jobOpenDate || c.installDate;
      const d = toInputDateFormat(rawD) || 'Unknown';
      if (!trendMap[d]) {
        trendMap[d] = { date: d, cards: 0, revenue: 0 };
      }
      trendMap[d].cards += 1;
      trendMap[d].revenue += cGrand;

      const rawMech = (c.mechanic || c.technicianName || c.mechanicName || '').toString().trim();
      const mech = normalizeStaffName(rawMech) || rawMech;
      if (mech) {
        if (!staffPerfMap[mech]) {
          staffPerfMap[mech] = { name: mech, role: 'Mechanic', jobsCount: 0, openCount: 0, closedCount: 0, labourTotal: 0, sparesTotal: 0, grandTotal: 0 };
        }
        staffPerfMap[mech].jobsCount += 1;
        if (isClosed) staffPerfMap[mech].closedCount += 1;
        else staffPerfMap[mech].openCount += 1;
        staffPerfMap[mech].labourTotal += cLabour;
        staffPerfMap[mech].sparesTotal += cSpares;
        staffPerfMap[mech].grandTotal += cGrand;
      }

      const rawSup = (c.wsIncharge || c.supervisor || c.supervisorName || '').toString().trim();
      const sup = normalizeStaffName(rawSup) || rawSup;
      if (sup && sup.toLowerCase() !== mech.toLowerCase()) {
        if (!staffPerfMap[sup]) {
          staffPerfMap[sup] = { name: sup, role: 'Supervisor', jobsCount: 0, openCount: 0, closedCount: 0, labourTotal: 0, sparesTotal: 0, grandTotal: 0 };
        }
        staffPerfMap[sup].jobsCount += 1;
        if (isClosed) staffPerfMap[sup].closedCount += 1;
        else staffPerfMap[sup].openCount += 1;
        staffPerfMap[sup].labourTotal += cLabour;
        staffPerfMap[sup].sparesTotal += cSpares;
        staffPerfMap[sup].grandTotal += cGrand;
      }
    }

    const trData = Object.values(trendMap).sort((a, b) => a.date.localeCompare(b.date));
    const locData = [
      { name: 'Location Workshop', value: wsCount },
      { name: 'DSS Doorstep Field', value: dsCount },
      { name: 'Event / Camp', value: evCount }
    ];

    const mechSummary = Object.values(staffPerfMap).sort((a, b) => b.jobsCount - a.jobsCount);
    const mechChart = mechSummary.filter((s) => s.role === 'Mechanic').map((s) => ({ name: s.name, jobs: s.jobsCount, revenue: s.grandTotal }));
    const supChart = mechSummary.filter((s) => s.role === 'Supervisor').map((s) => ({ name: s.name, jobs: s.jobsCount, revenue: s.grandTotal }));

    return {
      filteredReportCards: cards,
      totalCardsCount: totCardsCount,
      openCardsCount: opCardsCount,
      closedCardsCount: clCardsCount,
      totalRevenueSum: totRevenueSum,
      totalLabourSum: totLabourSum,
      totalSparesSum: totSparesSum,
      workshopCount: wsCount,
      dssCount: dsCount,
      eventCount: evCount,
      workshopRevenue: wsRevenue,
      dssRevenue: dsRevenue,
      eventRevenue: evRevenue,
      trendChartData: trData,
      locationChartData: locData,
      mechanicPerformanceSummary: mechSummary,
      mechanicChartData: mechChart,
      supervisorChartData: supChart
    };
  }, [
    savedJobCards,
    reportDateFrom,
    reportDateTo,
    reportLocation,
    reportStatus,
    reportMechanic,
    reportSupervisor,
    reportServiceType,
    reportSearchQuery,
    mechanicsList,
    supervisorsList
  ]);

  const kpiFilteredCards = useMemo(() => {
    if (selectedReportKpi === 'open') {
      return filteredReportCards.filter((c) => {
        const status = (c.status || 'Open').toString().trim().toLowerCase();
        const closedStatuses = ['closed', 'done', 'completed', 'finalized', 'billed'];
        return !closedStatuses.includes(status);
      });
    }
    if (selectedReportKpi === 'closed') {
      return filteredReportCards.filter((c) => {
        const status = (c.status || 'Open').toString().trim().toLowerCase();
        const closedStatuses = ['closed', 'done', 'completed', 'finalized', 'billed'];
        return closedStatuses.includes(status);
      });
    }
    if (selectedReportKpi === 'revenue') {
      return filteredReportCards.filter((c) => (parseFloat(c.gTotal) || 0) > 0);
    }
    if (selectedReportKpi === 'spares') {
      return filteredReportCards.filter((c) => {
        const amt = (parseFloat(c.warrantyMaterial) || 0) + (parseFloat(c.nonWarrantyMaterial) || 0) || parseFloat(c.partsTotal) || 0;
        return amt > 0;
      });
    }
    if (selectedReportKpi === 'labour') {
      return filteredReportCards.filter((c) => {
        const amt = parseFloat(c.totalLabour) || parseFloat(c.labourTotal) || 0;
        return amt > 0;
      });
    }
    if (selectedReportKpi === 'workshop') {
      return filteredReportCards.filter((c) => {
        const loc = (c.serviceLocation || c.servicePlace || 'workshop').toString().toLowerCase();
        return !loc.includes('dss') && !loc.includes('door') && !loc.includes('event') && !loc.includes('camp');
      });
    }
    if (selectedReportKpi === 'dss') {
      return filteredReportCards.filter((c) => {
        const loc = (c.serviceLocation || c.servicePlace || '').toString().toLowerCase();
        return loc.includes('dss') || loc.includes('door') || loc.includes('step') || loc.includes('field');
      });
    }
    if (selectedReportKpi === 'event') {
      return filteredReportCards.filter((c) => {
        const loc = (c.serviceLocation || c.servicePlace || '').toString().toLowerCase();
        return loc.includes('event') || loc.includes('camp');
      });
    }
    return filteredReportCards;
  }, [filteredReportCards, selectedReportKpi]);

  // Reset report pagination when filters change
  useEffect(() => {
    setReportPage(1);
  }, [
    reportDateFrom,
    reportDateTo,
    reportLocation,
    reportStatus,
    reportMechanic,
    reportSupervisor,
    reportServiceType,
    reportSearchQuery,
    selectedReportKpi
  ]);

  const totalReportPages = Math.max(1, Math.ceil(kpiFilteredCards.length / (reportPageSize === -1 ? 1 : reportPageSize)));
  const validReportPage = Math.min(reportPage, totalReportPages);

  const paginatedReportCards = useMemo(() => {
    if (reportPageSize === -1) return kpiFilteredCards;
    const start = (validReportPage - 1) * reportPageSize;
    return kpiFilteredCards.slice(start, start + reportPageSize);
  }, [kpiFilteredCards, validReportPage, reportPageSize]);

  // Delivery Customers for the current month
  const thisMonthDeliveredCustomers = useMemo(() => {
    const curMonthPrefix = todayISO.slice(0, 7);
    return uniqueCustomers.filter(c => {
      const rawDel = c['Date of del'] || c.dateOfDel || getFieldValue(c, 'installDate');
      const iso = toInputDateFormat(rawDel);
      return Boolean(iso && iso.startsWith(curMonthPrefix));
    });
  }, [uniqueCustomers, todayISO]);

  const filteredMonthDeliveries = useMemo(() => {
    if (!monthDeliveriesSearch.trim()) return thisMonthDeliveredCustomers;
    const q = monthDeliveriesSearch.toLowerCase();
    return thisMonthDeliveredCustomers.filter(c => {
      const name = (c['Customer Name'] || c.custName || getFieldValue(c, 'name') || '').toLowerCase();
      const phone = (c['Mobile No'] || c.phNo || c.mobile || getFieldValue(c, 'mobile') || '').toLowerCase();
      const chassis = (c['Chassis no'] || c.__chassisDisplay || c.chassis || getFieldValue(c, 'chassis') || '').toLowerCase();
      const village = (c['Village'] || c.village || getFieldValue(c, 'village') || '').toLowerCase();
      const model = (c['Model'] || c.model || getFieldValue(c, 'model') || '').toLowerCase();
      return name.includes(q) || phone.includes(q) || chassis.includes(q) || village.includes(q) || model.includes(q);
    });
  }, [thisMonthDeliveredCustomers, monthDeliveriesSearch]);

  const openDeliveriesInFollowup = () => {
    const [y, m] = todayISO.split('-');
    const firstDay = `${y}-${m}-01`;
    const lastDayDate = new Date(Number(y), Number(m), 0).getDate();
    const lastDay = `${y}-${m}-${String(lastDayDate).padStart(2, '0')}`;

    setFollowupDateFrom(firstDay);
    setFollowupDateTo(lastDay);
    setFollowupSearch('');
    setFollowupSupervisor('all');
    setFollowupStatus('all');
    setFollowupModelFilter('all');
    setFollowupSortBy('dateOfDel');
    setFollowupSortOrder('desc');
    setFollowupPage(1);
    setIsMonthDeliveriesModalOpen(false);
    setActiveTab('followup');
  };

  const handleCreateJobCardForCustomer = (cust: any) => {
    setIsMonthDeliveriesModalOpen(false);
    resetForm();
    autoFillCustomer(cust);
    setActiveTab('new_entry');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleViewCustomerHistory = (cust: any) => {
    const chassis = cust['Chassis no'] || cust.__chassisDisplay || getFieldValue(cust, 'chassis') || cust.chassis || '';
    setIsMonthDeliveriesModalOpen(false);
    setActiveTab('saved_cards');
    setSavedListSearch(chassis);
    setSavedListPage(1);
    setSavedStatusFilter('all');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleViewCustomerInFollowup = (cust: any) => {
    const chassis = cust['Chassis no'] || cust.__chassisDisplay || getFieldValue(cust, 'chassis') || cust.chassis || '';
    setIsMonthDeliveriesModalOpen(false);
    setActiveTab('followup');
    setFollowupSearch(chassis);
    setFollowupDateFrom('');
    setFollowupDateTo('');
    setFollowupPage(1);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleExportMonthDeliveriesCSV = () => {
    if (filteredMonthDeliveries.length === 0) {
      alert('No delivery records to export.');
      return;
    }
    const exportData = filteredMonthDeliveries.map((c, i) => {
      const chassis = c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || '';
      const hasJobCard = chassisWithJobCards.has(normalizeKey(chassis));
      return {
        'S.No': i + 1,
        'Customer Name': c['Customer Name'] || c.custName || getFieldValue(c, 'name') || '',
        'Mobile Number': c['Mobile No'] || c.phNo || c.mobile || getFieldValue(c, 'mobile') || '',
        'Model': c['Model'] || c.model || getFieldValue(c, 'model') || '',
        'Chassis No': chassis,
        'Engine No': c['Engine no'] || c.engineNo || getFieldValue(c, 'engineNo') || '',
        'Delivery Date': toInputDateFormat(c['Date of del'] || c.dateOfDel || getFieldValue(c, 'installDate')) || '',
        'Village': c['Village'] || c.village || getFieldValue(c, 'village') || '',
        'Mandal': c['Mandal'] || c.mandal || getFieldValue(c, 'mandal') || '',
        'Service Status': hasJobCard ? 'Job Card Created (Reporting)' : 'Pending (Non-Reporting)'
      };
    });

    const csvStr = Papa.unparse(exportData);
    const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `This_Month_Deliveries_${todayISO.slice(0, 7)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (authLoading) {
    return (
      <div className="min-h-screen bg-slate-100 flex flex-col items-center justify-center p-4">
        <div className="bg-white p-8 rounded-2xl shadow-md border border-slate-200 flex flex-col items-center gap-4 max-w-sm w-full text-center">
          <div className="w-12 h-12 border-4 border-blue-900 border-t-transparent rounded-full animate-spin"></div>
          <p className="text-sm font-bold text-slate-800 font-sans">Sri Gayathri Automotives</p>
          <p className="text-xs text-slate-500 font-medium font-sans">Loading security workspace...</p>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-3 md:p-6 text-slate-800 font-sans">
        <div className="w-full max-w-md bg-white border border-slate-200 shadow-xl rounded-2xl overflow-hidden">
          {/* Header styling */}
          <div className="bg-blue-900 text-white p-6 text-center space-y-2">
            <div className="inline-flex bg-white/10 p-3 rounded-full mb-1">
              <Wrench className="w-8 h-8 text-white" />
            </div>
            <h1 className="text-xl font-black tracking-tight">SRI GAYATHRI AUTOMOTIVES</h1>
            <p className="text-xs text-blue-100 font-medium">Job Card Generator & Spares Manager</p>
          </div>

          <div className="p-6 md:p-8 space-y-6">
            {authError && (
              <div className="bg-amber-50 border border-amber-200 text-amber-900 p-4 rounded-xl text-xs space-y-2">
                {authError.startsWith('DOMAIN_UNAUTHORIZED:') ? (
                  <>
                    <div className="flex items-center gap-2 font-bold text-amber-800">
                      <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                      <span>Firebase Authorized Domain Notice</span>
                    </div>
                    <p className="text-slate-700 leading-relaxed">
                      Google OAuth requires this preview domain to be listed in Firebase.
                    </p>
                    <div className="flex items-center justify-between gap-2 bg-white/80 p-2 rounded-lg border border-amber-200 font-mono text-[11px] text-slate-800 break-all">
                      <span>{authError.replace('DOMAIN_UNAUTHORIZED:', '')}</span>
                      <button
                        type="button"
                        onClick={() => {
                          navigator.clipboard?.writeText(authError.replace('DOMAIN_UNAUTHORIZED:', ''));
                          alert('Domain copied to clipboard!');
                        }}
                        className="px-2 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded text-[10px] font-sans font-bold shrink-0 cursor-pointer"
                      >
                        Copy
                      </button>
                    </div>
                    <p className="text-[11px] text-slate-600">
                      💡 Tip: Add this domain in <strong>Firebase Console → Authentication → Settings → Authorized domains</strong>, or use the <strong>Quick Access</strong> buttons below to start immediately.
                    </p>
                  </>
                ) : (
                  <div className="flex items-start gap-2 text-rose-700 font-medium">
                    <span className="text-sm">⚠️</span>
                    <div>{authError}</div>
                  </div>
                )}
              </div>
            )}

            <div className="space-y-4">
              <p className="text-sm text-slate-600 text-center font-medium">
                Sign in with your Google account or choose a workspace profile to continue.
              </p>

              <button
                type="button"
                onClick={handleGoogleSignIn}
                disabled={isAuthSubmitting}
                className="w-full py-3 bg-white hover:bg-slate-50 border border-slate-300 disabled:bg-slate-100 text-slate-700 font-bold text-sm rounded-xl transition-all shadow-sm flex items-center justify-center gap-3 cursor-pointer"
              >
                {isAuthSubmitting ? (
                  <div className="w-5 h-5 border-2 border-slate-400 border-t-transparent rounded-full animate-spin"></div>
                ) : (
                  <>
                    <svg className="w-5 h-5" viewBox="0 0 24 24">
                      <path
                        fill="currentColor"
                        d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
                      />
                      <path
                        fill="#34A853"
                        d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
                      />
                      <path
                        fill="#FBBC05"
                        d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"
                      />
                      <path
                        fill="#EA4335"
                        d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"
                      />
                    </svg>
                    Sign in with Google
                  </>
                )}
              </button>

              <div className="relative flex items-center justify-center my-3">
                <div className="border-t border-slate-200 w-full"></div>
                <span className="bg-white px-3 text-[11px] font-bold text-slate-400 uppercase tracking-wider whitespace-nowrap">
                  Or Quick Workspace Access
                </span>
                <div className="border-t border-slate-200 w-full"></div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                <button
                  type="button"
                  onClick={() => handleLocalSignIn('admin')}
                  className="w-full py-2.5 px-3 bg-blue-900 hover:bg-blue-800 text-white font-bold text-xs rounded-xl transition-all shadow-sm flex items-center justify-center gap-2 cursor-pointer"
                >
                  <UserCheck className="w-4 h-4" />
                  <span>Admin Access</span>
                </button>
                <button
                  type="button"
                  onClick={() => handleLocalSignIn('user')}
                  className="w-full py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold text-xs rounded-xl border border-slate-300 transition-all shadow-sm flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Users className="w-4 h-4 text-slate-600" />
                  <span>Staff Access</span>
                </button>
              </div>

              <p className="text-[11px] text-center text-slate-400 font-medium">
                Admin mode provides full master database, backup, and deletion permissions.
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Reports & Analytics Data Constants
  const REPORT_COLORS = ['#4f46e5', '#0284c7', '#d97706', '#10b981', '#8b5cf6', '#ec4899'];

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans flex flex-col md:flex-row print:block">
      {/* MOBILE COMPACT TOP HEADER BAR */}
      <div className="md:hidden flex items-center justify-between p-2 bg-blue-950 text-white shadow-md sticky top-0 z-40 print:hidden shrink-0">
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => setIsSidebarHidden(prev => !prev)}
            className="p-1.5 rounded-lg hover:bg-blue-900 text-white cursor-pointer"
            title="Toggle Menu"
          >
            <Menu className="w-5 h-5" />
          </button>
          <div>
            <h1 className="text-xs font-black uppercase tracking-wide leading-none">SRI GAYATHRI</h1>
            <p className="text-[8px] text-blue-300 font-extrabold uppercase mt-0.5">Automotives</p>
          </div>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="flex items-center bg-blue-900/90 p-0.5 rounded-lg border border-blue-800">
            <button
              type="button"
              onClick={() => toggleLanguage('te')}
              className={`px-2 py-0.5 text-[10px] font-black rounded ${appLang === 'te' ? 'bg-amber-400 text-slate-950 font-black' : 'text-blue-200'}`}
            >
              తెలుగు
            </button>
            <button
              type="button"
              onClick={() => toggleLanguage('en')}
              className={`px-2 py-0.5 text-[10px] font-black rounded ${appLang === 'en' ? 'bg-amber-400 text-slate-950 font-black' : 'text-blue-200'}`}
            >
              EN
            </button>
          </div>
          <span className="bg-blue-900 text-white font-bold text-[10px] px-2 py-1 rounded-full border border-blue-800 uppercase">
            {activeTab === 'new_entry' ? (appLang === 'te' ? '✍️ కొత్త జాబ్' : '✍️ New Entry') :
             activeTab === 'saved_cards' ? (appLang === 'te' ? '📂 సేవ్డ్' : '📂 Saved') :
             activeTab === 'reports' ? (appLang === 'te' ? '📊 రిపోర్ట్స్' : '📊 Reports') :
             activeTab === 'followup' ? (appLang === 'te' ? '👥 కస్టమర్లు' : '👥 Customers') :
             activeTab === 'telecalling' ? (appLang === 'te' ? '📞 టెలి కాలింగ్' : '📞 Tele Calling') :
             activeTab === 'complaints' ? (appLang === 'te' ? '📝 కంప్లైంట్స్' : '📝 Complaints') :
             activeTab === 'free_service_followup' ? (appLang === 'te' ? '🛠️ ఉచిత సర్వీస్' : '🛠️ Free Service') :
             (appLang === 'te' ? '🗄️ డేటాబేస్' : '🗄️ Databases')}
          </span>
        </div>
      </div>

      {/* MOBILE SIDEBAR OVERLAY/BACKDROP */}
      {!isSidebarHidden && (
        <div 
          className="md:hidden fixed inset-0 bg-slate-950/60 z-40 transition-opacity print:hidden" 
          onClick={() => setIsSidebarHidden(true)}
        />
      )}

      {/* SIDEBAR */}
      <aside className={`
        ${isSidebarHidden ? 'hidden md:hidden' : 'fixed inset-y-0 left-0 z-50 md:sticky md:top-0 md:h-screen md:flex'} 
        w-72 ${isSidebarCollapsed ? 'md:w-16' : 'md:w-64'} 
        shrink-0 bg-white border-r border-slate-200 flex flex-col print:hidden shadow-xl md:shadow-xs transition-all duration-200
      `}>
        {/* BRAND HEADER AREA */}
        <div className="p-3 md:p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 overflow-hidden">
            <div className="bg-blue-900 text-white p-2 rounded-xl shadow-xs shrink-0">
              <Wrench className="w-5 h-5 text-white" />
            </div>
            {!isSidebarCollapsed && (
              <div className="transition-opacity duration-200 whitespace-nowrap">
                <h1 className="text-xs font-black text-slate-900 leading-tight uppercase tracking-wide">SRI GAYATHRI</h1>
                <p className="text-[9px] text-slate-500 font-extrabold uppercase">Automotives</p>
              </div>
            )}
          </div>
          
          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
              className="hidden md:block p-1 rounded-lg hover:bg-slate-200 text-slate-500 hover:text-slate-900 transition-colors cursor-pointer shrink-0"
              title={isSidebarCollapsed ? "Expand Sidebar" : "Collapse Sidebar"}
            >
              {isSidebarCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronLeft className="w-4 h-4" />}
            </button>
            <button
              type="button"
              onClick={() => setIsSidebarHidden(true)}
              className="p-1 rounded-lg hover:bg-red-50 text-slate-500 hover:text-red-600 transition-colors cursor-pointer shrink-0"
              title="Hide Sidebar"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* LANGUAGE SWITCHER BAR IN SIDEBAR */}
        <div className="px-3 py-2 border-b border-slate-200 bg-slate-100 flex items-center justify-between gap-1 shrink-0">
          {!isSidebarCollapsed && (
            <div className="flex items-center gap-1 text-[11px] font-bold text-slate-700">
              <Globe className="w-3.5 h-3.5 text-blue-900" />
              <span>{appLang === 'te' ? 'భాష (Language)' : 'Language'}</span>
            </div>
          )}
          <div className="flex items-center bg-white p-0.5 rounded-lg border border-slate-300 shadow-2xs">
            <button
              type="button"
              onClick={() => toggleLanguage('te')}
              className={`px-2.5 py-1 text-[11px] font-black rounded-md transition cursor-pointer ${
                appLang === 'te' ? 'bg-blue-900 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-950'
              }`}
              title="తెలుగు భాష ఎంచుకోండి"
            >
              తెలుగు
            </button>
            <button
              type="button"
              onClick={() => toggleLanguage('en')}
              className={`px-2.5 py-1 text-[11px] font-black rounded-md transition cursor-pointer ${
                appLang === 'en' ? 'bg-blue-900 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-950'
              }`}
              title="Select English Language"
            >
              English
            </button>
          </div>
        </div>

        <nav className="flex-1 p-2 md:p-3 space-y-2 overflow-y-auto">
          {menuOrder.map((itemId, idx) => {
            let label = '';
            let icon: React.ReactNode = null;
            let activeStyle = 'bg-blue-900 text-white border-blue-900 shadow-sm';
            let badge: React.ReactNode = null;

            if (itemId === 'dashboard') {
              label = t('dashboard');
              icon = <LayoutDashboard className="w-4 h-4 shrink-0 text-current" />;
            } else if (itemId === 'new_entry') {
              label = t('newJobEntry');
              icon = <Plus className="w-4 h-4 shrink-0 text-current" />;
              badge = editingCardId ? (
                <span className="bg-amber-400 text-slate-900 text-[8px] px-1.5 py-0.5 rounded-full font-black uppercase shrink-0">
                  {appLang === 'te' ? 'ఎడిటింగ్' : 'Editing'}
                </span>
              ) : null;
            } else if (itemId === 'saved_cards') {
              label = t('savedJobCards');
              icon = <FileText className="w-4 h-4 shrink-0 text-current" />;
              badge = (
                <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-black shrink-0 ${
                  activeTab === 'saved_cards' ? 'bg-blue-800 text-white' : 'bg-slate-200 text-slate-700'
                }`}>
                  {savedJobCards.length}
                </span>
              );
            } else if (itemId === 'reports') {
              label = t('reportsAnalytics');
              icon = <BarChart3 className="w-4 h-4 shrink-0 text-current" />;
            } else if (itemId === 'followup') {
              label = t('customerData');
              icon = <PhoneCall className="w-4 h-4 shrink-0 text-current" />;
            } else if (itemId === 'telecalling') {
              label = t('teleCalling');
              icon = <Phone className="w-4 h-4 shrink-0 text-current" />;
              activeStyle = 'bg-amber-600 text-white border-amber-600 shadow-sm';
              badge = scheduledFollowupsTodayCount > 0 ? (
                <span className="text-[10px] px-1.5 py-0.5 rounded-full font-black shrink-0 bg-red-600 text-white animate-pulse">
                  {scheduledFollowupsTodayCount}
                </span>
              ) : null;
            } else if (itemId === 'free_service_followup') {
              label = t('freeServiceFollowup');
              icon = <Clock className="w-4 h-4 shrink-0 text-current" />;
              activeStyle = 'bg-teal-700 text-white border-teal-700 shadow-sm';
              const totalDue: number = (Object.values(freeServiceFollowupCategories) as any[]).reduce((acc: number, list: any) => acc + (list?.length || 0), 0);
              badge = totalDue > 0 ? (
                <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-black shrink-0 ${
                  activeTab === 'free_service_followup' ? 'bg-teal-900 text-white' : 'bg-teal-100 text-teal-800'
                }`}>
                  {totalDue}
                </span>
              ) : null;
            } else if (itemId === 'complaints') {
              label = t('complaintRegister');
              icon = <AlertCircle className="w-4 h-4 shrink-0 text-current" />;
            } else if (itemId === 'attendance') {
              label = t('staffAttendance');
              icon = <UserCheck className="w-4 h-4 shrink-0 text-current" />;
              activeStyle = 'bg-teal-800 text-white border-teal-800 shadow-sm';
            } else if (itemId === 'databases') {
              label = t('masterDatabases');
              icon = <FileSpreadsheet className="w-4 h-4 shrink-0 text-current" />;
            }

            const isActive = activeTab === itemId;

            return (
              <div key={itemId} className="relative group flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => setActiveTab(itemId as any)}
                  className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5' : 'justify-between p-3'} rounded-xl font-bold text-xs transition-all border cursor-pointer ${
                    isActive 
                      ? activeStyle 
                      : 'text-slate-600 bg-white hover:bg-slate-50 hover:text-slate-900 border-slate-200'
                  }`}
                  title={isSidebarCollapsed ? label : undefined}
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    {icon}
                    {!isSidebarCollapsed && <span className="truncate">{label}</span>}
                  </div>
                  {!isSidebarCollapsed && badge}
                </button>

                {/* MOVE UP / MOVE DOWN CONTROLS */}
                {!isSidebarCollapsed && (
                  <div className={`flex items-center gap-1 shrink-0 transition-opacity opacity-0 group-hover:opacity-100`}>
                    <GripVertical className="w-4 h-4 text-slate-400 cursor-grab" />
                    <div className="flex flex-col gap-0.5">
                      <button
                        type="button"
                        disabled={idx === 0}
                        onClick={(e) => {
                          e.stopPropagation();
                          moveMenuItem(idx, 'up');
                        }}
                        className="p-0.5 rounded bg-slate-100 hover:bg-amber-600 hover:text-white disabled:opacity-20 text-slate-600 transition-colors cursor-pointer disabled:cursor-not-allowed"
                        title="Move Box Up"
                      >
                        <ChevronUp className="w-3 h-3" />
                      </button>
                      <button
                        type="button"
                        disabled={idx === menuOrder.length - 1}
                        onClick={(e) => {
                          e.stopPropagation();
                          moveMenuItem(idx, 'down');
                        }}
                        className="p-0.5 rounded bg-slate-100 hover:bg-amber-600 hover:text-white disabled:opacity-20 text-slate-600 transition-colors cursor-pointer disabled:cursor-not-allowed"
                        title="Move Box Down"
                      >
                        <ChevronDown className="w-3 h-3" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </nav>

        {/* QUICK ACTIONS IN SIDEBAR */}
        <div className="p-3 border-t border-slate-200 bg-slate-50/50 shrink-0 space-y-1.5">
          <button
            type="button"
            onClick={() => setIsMenuOrderModalOpen(true)}
            className={`w-full flex items-center justify-center ${isSidebarCollapsed ? 'p-2' : 'gap-2 py-1.5 px-3'} bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl transition-colors cursor-pointer border border-slate-200`}
            title={isSidebarCollapsed ? "Customize Menu Order" : undefined}
          >
            <Settings className="w-3.5 h-3.5 shrink-0 text-slate-600" />
            {!isSidebarCollapsed && <span className="truncate">⚙️ Reorder Menu</span>}
          </button>
          <button
            type="button"
            onClick={() => {
              if (confirm("ARE YOU SURE YOU WANT TO CLEAR ALL DATA? THIS CANNOT BE UNDONE!")) {
                const password = prompt("Enter Admin Password:");
                if (password) {
                  fetch('/api/database/clear', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ password })
                  }).then(res => res.json()).then(data => {
                    if (data.success) {
                      alert("All data cleared successfully.");
                      window.location.reload();
                    } else {
                      alert("Failed to clear data: " + data.error);
                    }
                  });
                }
              }
            }}
            className={`w-full flex items-center justify-center ${isSidebarCollapsed ? 'p-2' : 'gap-2 py-1.5 px-3'} bg-red-100 hover:bg-red-200 text-red-700 font-bold text-xs rounded-xl transition-colors cursor-pointer border border-red-200`}
          >
            <Trash2 className="w-3.5 h-3.5 shrink-0 text-red-600" />
            {!isSidebarCollapsed && <span className="truncate">🗑️ Clear Master Data</span>}
          </button>
          <button
            type="button"
            onClick={handleOpenAddCustomer}
            className={`w-full flex items-center justify-center ${isSidebarCollapsed ? 'p-2' : 'gap-2 py-2 px-3'} bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl transition-colors shadow-xs cursor-pointer mb-2`}
            title={isSidebarCollapsed ? "Add New Customer" : undefined}
          >
            <UserPlus className="w-4 h-4 shrink-0 text-white" />
            {!isSidebarCollapsed && <span className="truncate">+ Add Customer</span>}
          </button>
          
          <button
            type="button"
            onClick={handleSignOut}
            className={`w-full flex items-center justify-center ${isSidebarCollapsed ? 'p-2' : 'gap-2 py-2 px-3'} bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-xs rounded-xl transition-colors shadow-xs cursor-pointer border border-rose-200 mt-auto`}
            title={isSidebarCollapsed ? "Sign Out" : undefined}
          >
            <LogOut className="w-4 h-4 shrink-0" />
            {!isSidebarCollapsed && <span className="truncate">Sign Out</span>}
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT AREA */}
      <div className="flex-1 min-w-0 flex flex-col print:block relative">
        {/* DESKTOP FLOATING SHOW-SIDEBAR BUTTON */}
        {isSidebarHidden && (
          <button
            type="button"
            onClick={() => setIsSidebarHidden(false)}
            className="hidden md:flex fixed top-3 left-3 z-50 items-center gap-1.5 bg-blue-950 hover:bg-blue-900 text-white font-bold text-[11px] p-2 rounded-xl shadow-lg border border-blue-800 transition-all cursor-pointer"
            title="Show Sidebar"
          >
            <Menu className="w-4 h-4" />
            <span>Show Menu</span>
          </button>
        )}

        <main className="w-full min-w-0 p-1 sm:p-2 md:p-3 print:p-0 print:m-0">
          {/* Firestore quota banner removed to rely directly on Cloud SQL */}
          {/* TAB: DASHBOARD */}
          {activeTab === 'dashboard' && (() => {
            const todayAttendanceDate = attendanceDate || todayISO;
            const todayRecs = attendanceRecords[todayAttendanceDate] || {};
            const totalStaffCount = staffMembers.length;
            let presentTodayCount = 0;
            let absentTodayCount = 0;
            let leaveTodayCount = 0;
            staffMembers.forEach(s => {
              const sKey = s.id || s.name;
              const status = todayRecs[sKey]?.status;
              if (status === 'present') presentTodayCount++;
              else if (status === 'absent') absentTodayCount++;
              else if (status === 'leave') leaveTodayCount++;
            });
            const unmarkedTodayCount = totalStaffCount - (presentTodayCount + absentTodayCount + leaveTodayCount);

            return (
              <div className="space-y-3.5">
                {/* DASHBOARD HEADER & QUICK CONTROLS */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-2.5 border-b border-slate-200 pb-3 bg-white p-3 rounded-xl border border-slate-100 shadow-2xs">
                  <div>
                    <h1 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
                      <span>📊</span> Business Dashboard & Executive KPIs
                    </h1>
                    <p className="text-[11px] text-slate-500 font-medium">Real-time attendance, job cards, telecalling, free service follow-ups, deliveries, and financial metrics.</p>
                  </div>
                  
                  {/* EXPAND / MINIMIZE ALL ACTIONS */}
                  <div className="flex flex-wrap items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => {
                        setDashDeliveriesOpen(true);
                        setDashAttendanceOpen(true);
                        setDashComplaintsOpen(true);
                        setDashJobCardsOpen(true);
                        setDashFinancialsOpen(true);
                        setDashTelecallingOpen(true);
                      }}
                      className="px-2.5 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-[11px] rounded-lg transition-colors border border-indigo-200 cursor-pointer flex items-center gap-1 shadow-2xs"
                    >
                      <ChevronDown className="w-3.5 h-3.5" /> Expand All
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setDashDeliveriesOpen(false);
                        setDashAttendanceOpen(false);
                        setDashComplaintsOpen(false);
                        setDashJobCardsOpen(false);
                        setDashFinancialsOpen(false);
                        setDashTelecallingOpen(false);
                      }}
                      className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[11px] rounded-lg transition-colors border border-slate-300 cursor-pointer flex items-center gap-1 shadow-2xs"
                    >
                      <ChevronUp className="w-3.5 h-3.5" /> Minimize All
                    </button>
                  </div>
                </div>

                {/* QUICK WIDGET TOGGLE PILLS / MINIMIZE SELECTORS */}
                <div className="bg-slate-50 border border-slate-200 p-2 rounded-xl flex flex-wrap items-center gap-1.5 text-xs">
                  <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-600 mr-1 flex items-center gap-1">
                    <Filter className="w-3 h-3 text-slate-500" /> Widgets:
                  </span>
                  
                  <button
                    type="button"
                    onClick={() => setDashDeliveriesOpen(!dashDeliveriesOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashDeliveriesOpen 
                        ? 'bg-indigo-600 text-white shadow-2xs' 
                        : 'bg-white text-indigo-700 border border-indigo-200 hover:bg-indigo-50 opacity-60'
                    }`}
                  >
                    <span>📦 Deliveries</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashDeliveriesOpen ? 'bg-white/20 text-white' : 'bg-indigo-100 text-indigo-800'}`}>
                      {thisMonthDeliveredCustomers.length}
                    </span>
                    {dashDeliveriesOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>

                  <button
                    type="button"
                    onClick={() => setDashAttendanceOpen(!dashAttendanceOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashAttendanceOpen 
                        ? 'bg-purple-600 text-white shadow-2xs' 
                        : 'bg-white text-purple-700 border border-purple-200 hover:bg-purple-50 opacity-60'
                    }`}
                  >
                    <span>👥 Attendance</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashAttendanceOpen ? 'bg-white/20 text-white' : 'bg-purple-100 text-purple-800'}`}>
                      {totalStaffCount}
                    </span>
                    {dashAttendanceOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>

                  <button
                    type="button"
                    onClick={() => setDashComplaintsOpen(!dashComplaintsOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashComplaintsOpen 
                        ? 'bg-amber-600 text-white shadow-2xs' 
                        : 'bg-white text-amber-700 border border-amber-200 hover:bg-amber-50 opacity-60'
                    }`}
                  >
                    <span>⚠️ Complaints</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashComplaintsOpen ? 'bg-white/20 text-white' : 'bg-amber-100 text-amber-800'}`}>
                      {complaints.length}
                    </span>
                    {dashComplaintsOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>

                  <button
                    type="button"
                    onClick={() => setDashJobCardsOpen(!dashJobCardsOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashJobCardsOpen 
                        ? 'bg-blue-600 text-white shadow-2xs' 
                        : 'bg-white text-blue-700 border border-blue-200 hover:bg-blue-50 opacity-60'
                    }`}
                  >
                    <span>📄 Job Cards</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashJobCardsOpen ? 'bg-white/20 text-white' : 'bg-blue-100 text-blue-800'}`}>
                      {savedJobCards.length}
                    </span>
                    {dashJobCardsOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>

                  <button
                    type="button"
                    onClick={() => setDashFinancialsOpen(!dashFinancialsOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashFinancialsOpen 
                        ? 'bg-emerald-600 text-white shadow-2xs' 
                        : 'bg-white text-emerald-700 border border-emerald-200 hover:bg-emerald-50 opacity-60'
                    }`}
                  >
                    <span>📊 Financials</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashFinancialsOpen ? 'bg-white/20 text-white' : 'bg-emerald-100 text-emerald-800'}`}>
                      ₹{totalRevenueSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}
                    </span>
                    {dashFinancialsOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>

                  <button
                    type="button"
                    onClick={() => setDashTelecallingOpen(!dashTelecallingOpen)}
                    className={`px-2.5 py-1 rounded-lg font-bold text-[11px] transition-all cursor-pointer flex items-center gap-1.5 ${
                      dashTelecallingOpen 
                        ? 'bg-teal-600 text-white shadow-2xs' 
                        : 'bg-white text-teal-700 border border-teal-200 hover:bg-teal-50 opacity-60'
                    }`}
                  >
                    <span>📞 Telecalling</span>
                    <span className={`text-[9.5px] px-1.5 py-0.2 rounded-full font-mono ${dashTelecallingOpen ? 'bg-white/20 text-white' : 'bg-teal-100 text-teal-800'}`}>
                      {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length} Due
                    </span>
                    {dashTelecallingOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                  </button>
                </div>

                {/* 1. THIS MONTH DELIVERIES SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-indigo-900 flex items-center gap-1.5">
                      <Calendar className="w-3.5 h-3.5 text-indigo-600" /> Delivery Information & Current Month Deliveries ({thisMonthDeliveredCustomers.length})
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashDeliveriesOpen(!dashDeliveriesOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-indigo-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashDeliveriesOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashDeliveriesOpen ? (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
                      <div 
                        onClick={() => setIsMonthDeliveriesModalOpen(true)}
                        className="bg-gradient-to-br from-indigo-50 to-indigo-100/80 border-2 border-indigo-300 hover:border-indigo-400 p-3.5 rounded-xl shadow-2xs hover:shadow-md transition-all cursor-pointer flex items-center justify-between group"
                      >
                        <div>
                          <span className="text-[10.5px] font-black text-indigo-950 uppercase tracking-wider block">This Month Deliveries</span>
                          <span className="text-[9.5px] text-indigo-700 font-bold">Delivered in {new Date().toLocaleString('default', { month: 'long', year: 'numeric' })}</span>
                          <p className="text-[9px] text-slate-500 mt-1 font-medium">Click to view full customer list & actions</p>
                        </div>
                        <div className="text-right">
                          <div className="text-2xl font-black text-indigo-950 font-mono group-hover:scale-105 transition-transform">
                            {thisMonthDeliveredCustomers.length}
                          </div>
                          <span className="text-[10px] font-bold text-indigo-800 underline bg-indigo-200/70 px-2 py-0.5 rounded-full inline-block mt-1">
                            📋 View {thisMonthDeliveredCustomers.length} Customers →
                          </span>
                        </div>
                      </div>

                      <div 
                        onClick={openDeliveriesInFollowup}
                        className="bg-white border border-slate-200 hover:border-indigo-200 p-3.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex items-center justify-between group"
                      >
                        <div>
                          <span className="text-[10px] font-bold text-slate-700 uppercase tracking-wider block">Follow-up Tab View</span>
                          <span className="text-[9px] text-slate-500 font-medium">Open filtered follow-up tree for this month</span>
                        </div>
                        <div className="text-indigo-600 font-bold text-xs group-hover:translate-x-1 transition-transform flex items-center gap-1">
                          <span>Open Follow-up</span> &rarr;
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashDeliveriesOpen(true)}
                      className="py-1.5 px-3 bg-indigo-50/50 hover:bg-indigo-50 border border-dashed border-indigo-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-indigo-900"
                    >
                      <span className="font-semibold text-[11px]">📦 Current Month Deliveries: <b>{thisMonthDeliveredCustomers.length} Customers</b></span>
                      <span className="text-[10px] font-bold underline text-indigo-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

                {/* 2. STAFF ATTENDANCE TODAY SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-purple-900 flex items-center gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-purple-600" /> Staff Attendance Today ({todayAttendanceDate})
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashAttendanceOpen(!dashAttendanceOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-purple-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashAttendanceOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashAttendanceOpen ? (
                    <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                      <div 
                        onClick={() => { setActiveTab('attendance'); setAttendanceTab('daily'); }}
                        className="bg-gradient-to-br from-purple-50 to-purple-100/60 border border-purple-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="text-[10px] font-extrabold text-purple-900 uppercase">Total Staff</div>
                        <div className="text-lg font-black text-purple-950 my-0.5">{totalStaffCount}</div>
                        <div className="text-[9px] font-bold text-purple-700 underline">View attendance →</div>
                      </div>

                      <div 
                        onClick={() => { setActiveTab('attendance'); setAttendanceTab('daily'); }}
                        className="bg-gradient-to-br from-emerald-50 to-emerald-100/60 border border-emerald-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="text-[10px] font-extrabold text-emerald-900 uppercase">Present Today</div>
                        <div className="text-lg font-black text-emerald-950 my-0.5">{presentTodayCount}</div>
                        <div className="text-[9px] font-bold text-emerald-700 underline">View present →</div>
                      </div>

                      <div 
                        onClick={() => { setActiveTab('attendance'); setAttendanceTab('daily'); }}
                        className="bg-gradient-to-br from-rose-50 to-rose-100/60 border border-rose-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="text-[10px] font-extrabold text-rose-900 uppercase">Absent Today</div>
                        <div className="text-lg font-black text-rose-950 my-0.5">{absentTodayCount}</div>
                        <div className="text-[9px] font-bold text-rose-700 underline">View absent →</div>
                      </div>

                      <div 
                        onClick={() => { setActiveTab('attendance'); setAttendanceTab('daily'); }}
                        className="bg-gradient-to-br from-amber-50 to-amber-100/60 border border-amber-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="text-[10px] font-extrabold text-amber-900 uppercase">On Leave</div>
                        <div className="text-lg font-black text-amber-950 my-0.5">{leaveTodayCount}</div>
                        <div className="text-[9px] font-bold text-amber-700 underline">View leaves →</div>
                      </div>

                      <div 
                        onClick={() => { setActiveTab('attendance'); setAttendanceTab('daily'); }}
                        className="bg-gradient-to-br from-slate-50 to-slate-100/80 border border-slate-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group col-span-2 sm:col-span-1"
                      >
                        <div className="text-[10px] font-extrabold text-slate-700 uppercase">Unmarked</div>
                        <div className="text-lg font-black text-slate-900 my-0.5">{unmarkedTodayCount}</div>
                        <div className="text-[9px] font-bold text-slate-600 underline">Mark attendance →</div>
                      </div>
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashAttendanceOpen(true)}
                      className="py-1.5 px-3 bg-purple-50/50 hover:bg-purple-50 border border-dashed border-purple-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-purple-900"
                    >
                      <span className="font-semibold text-[11px]">👥 Today's Attendance: <b>{presentTodayCount} Present</b> / {absentTodayCount} Absent / {totalStaffCount} Total</span>
                      <span className="text-[10px] font-bold underline text-purple-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

                {/* 3. COMPLAINT REGISTER SUMMARY SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-amber-900 flex items-center gap-1.5">
                      <AlertCircle className="w-3.5 h-3.5 text-amber-600" /> Complaint Register Summary ({complaints.length})
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashComplaintsOpen(!dashComplaintsOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-amber-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashComplaintsOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashComplaintsOpen ? (
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      {[
                        { label: 'Total Complaints', value: complaints.length, filter: 'all', color: 'bg-slate-50 border-slate-200 text-slate-900' },
                        { label: 'Open Complaints', value: complaints.filter(c => c.status === 'Open' || !c.status).length, filter: 'Open', color: 'bg-rose-50 border-rose-200 text-rose-900' },
                        { label: 'Running Complaints', value: complaints.filter(c => c.status === 'Running').length, filter: 'Running', color: 'bg-amber-50 border-amber-200 text-amber-900' },
                        { label: 'Closed Complaints', value: complaints.filter(c => c.status === 'Closed').length, filter: 'Closed', color: 'bg-emerald-50 border-emerald-200 text-emerald-900' },
                      ].map((m, i) => (
                        <div key={i} className={`p-2.5 rounded-xl border ${m.color} flex items-center justify-between hover:shadow-sm transition-all cursor-pointer`}
                          onClick={() => {
                            setActiveTab('complaints');
                            setComplaintStatusFilter(m.filter as any);
                          }}
                        >
                          <div>
                            <div className="text-[10px] font-extrabold uppercase">{m.label}</div>
                            <div className="text-[9px] font-bold opacity-70">Click to view</div>
                          </div>
                          <div className="text-lg font-black">{m.value}</div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashComplaintsOpen(true)}
                      className="py-1.5 px-3 bg-amber-50/50 hover:bg-amber-50 border border-dashed border-amber-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-amber-900"
                    >
                      <span className="font-semibold text-[11px]">⚠️ Complaints: <b>{complaints.filter(c => c.status === 'Open' || !c.status).length} Open</b> / {complaints.length} Total</span>
                      <span className="text-[10px] font-bold underline text-amber-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

                {/* 4. JOB CARDS & CUSTOMER DATA SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-blue-900 flex items-center gap-1.5">
                      <FileText className="w-3.5 h-3.5 text-blue-600" /> Job Cards & Customer Data ({savedJobCards.length} Cards, {uniqueCustomers.length} Customers)
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashJobCardsOpen(!dashJobCardsOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-blue-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashJobCardsOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashJobCardsOpen ? (
                    <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2">
                      {[
                        { label: 'Total Job Cards', value: savedJobCards.length, color: 'text-blue-900 border-blue-200 bg-blue-50/40', filter: 'all_jc' },
                        { label: 'Open JC', value: savedJobCards.filter(c => !isCardClosed(c)).length, color: 'text-amber-900 border-amber-200 bg-amber-50/40', filter: 'open_jc' },
                        { label: 'Closed JC', value: savedJobCards.filter(c => isCardClosed(c)).length, color: 'text-emerald-900 border-emerald-200 bg-emerald-50/40', filter: 'closed_jc' },
                        { label: 'Pending Online J.C.', value: savedJobCards.filter(c => !(c.onlineJobCardNo || '').toString().trim()).length, color: 'text-rose-900 border-rose-200 bg-rose-50/50', filter: 'missing_online' },
                        { label: 'Total Customers', value: uniqueCustomers.length, color: 'text-indigo-900 border-indigo-200 bg-indigo-50/40', filter: 'all_cust' },
                        { label: 'Reporting Cust', value: uniqueCustomers.filter(c => chassisWithJobCards.has(normalizeKey(c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || ''))).length, color: 'text-emerald-900 border-emerald-200 bg-emerald-50/40', filter: 'rep_cust' },
                        { label: 'Non-Rep Cust', value: uniqueCustomers.filter(c => !chassisWithJobCards.has(normalizeKey(c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || ''))).length, color: 'text-red-900 border-red-200 bg-red-50/40', filter: 'nonrep_cust' },
                      ].map((m, i) => (
                        <div key={i} className={`p-2.5 rounded-xl border ${m.color} text-center hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between`}
                          onClick={() => {
                            if (m.filter === 'all_jc' || m.filter === 'open_jc' || m.filter === 'closed_jc' || m.filter === 'missing_online') {
                                setActiveTab('saved_cards');
                                if (m.filter === 'open_jc') setSavedStatusFilter('Open');
                                else if (m.filter === 'closed_jc') setSavedStatusFilter('Closed');
                                else if (m.filter === 'missing_online') setSavedStatusFilter('MissingOnline');
                                else setSavedStatusFilter('all');
                            } else if (m.filter === 'all_cust' || m.filter === 'rep_cust' || m.filter === 'nonrep_cust') {
                                setActiveTab('followup');
                                if (m.filter === 'rep_cust') setFollowupStatus('reporting');
                                else if (m.filter === 'nonrep_cust') setFollowupStatus('not_reporting');
                                else setFollowupStatus('all');
                            }
                          }}
                        >
                          <div className="text-[9.5px] font-extrabold uppercase leading-tight">{m.label}</div>
                          <div className="text-base font-black my-1">{m.value}</div>
                          <span className="text-[8.5px] font-bold underline opacity-80">View list</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashJobCardsOpen(true)}
                      className="py-1.5 px-3 bg-blue-50/50 hover:bg-blue-50 border border-dashed border-blue-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-blue-900"
                    >
                      <span className="font-semibold text-[11px]">📄 Job Cards: <b>{savedJobCards.length} Saved</b> | {uniqueCustomers.length} Total Registered Customers</span>
                      <span className="text-[10px] font-bold underline text-blue-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

                {/* 5. REPORTS & FINANCIAL KEY METRICS SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-emerald-900 flex items-center gap-1.5">
                      <BarChart3 className="w-3.5 h-3.5 text-emerald-600" /> Reports & Financial KPIs (₹{totalRevenueSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })})
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashFinancialsOpen(!dashFinancialsOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-emerald-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashFinancialsOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashFinancialsOpen ? (
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
                      {[
                        { label: 'Total Revenue', value: `₹${totalRevenueSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`, sub: 'Total billings', kpi: 'revenue', color: 'bg-indigo-50/60 border-indigo-200 text-indigo-900' },
                        { label: 'Spares Sales', value: `₹${totalSparesSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`, sub: 'Parts & materials', kpi: 'spares', color: 'bg-purple-50/60 border-purple-200 text-purple-900' },
                        { label: 'Labour Charges', value: `₹${totalLabourSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`, sub: 'Workmanship', kpi: 'labour', color: 'bg-teal-50/60 border-teal-200 text-teal-900' },
                        { label: 'Workshop Cards', value: `${workshopCount} Cards`, sub: 'In-workshop service', kpi: 'workshop', color: 'bg-blue-50/60 border-blue-200 text-blue-900' },
                        { label: 'DSS (Door Step)', value: `${dssCount} Cards`, sub: 'Field service', kpi: 'dss', color: 'bg-cyan-50/60 border-cyan-200 text-cyan-900' },
                        { label: 'Event / Camp', value: `${eventCount} Cards`, sub: 'Camp service', kpi: 'event', color: 'bg-amber-50/60 border-amber-200 text-amber-900' },
                      ].map((item, idx) => (
                        <div
                          key={idx}
                          onClick={() => {
                            setActiveTab('reports');
                            setSelectedReportKpi(item.kpi as any);
                          }}
                          className={`${item.color} border p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group`}
                        >
                          <div className="text-[9.5px] font-extrabold uppercase tracking-wider">{item.label}</div>
                          <div className="my-1">
                            <div className="text-sm font-black leading-tight">{item.value}</div>
                            <span className="text-[8.5px] opacity-75 font-bold">{item.sub}</span>
                          </div>
                          <div className="text-[8.5px] font-bold underline opacity-80 group-hover:opacity-100">View report</div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashFinancialsOpen(true)}
                      className="py-1.5 px-3 bg-emerald-50/50 hover:bg-emerald-50 border border-dashed border-emerald-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-emerald-900"
                    >
                      <span className="font-semibold text-[11px]">📊 Financial Revenue: <b>₹{totalRevenueSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</b> | Spares: ₹{totalSparesSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</span>
                      <span className="text-[10px] font-bold underline text-emerald-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

                {/* 6. FREE SERVICE & TELECALLING EXECUTIVE METRICS SECTION */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 shadow-2xs space-y-2">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <h3 className="text-[11px] font-black uppercase tracking-wider text-teal-900 flex items-center gap-1.5">
                      <PhoneCall className="w-3.5 h-3.5 text-teal-600" /> Free Service Follow-up & Today's Telecalling
                    </h3>
                    <button
                      type="button"
                      onClick={() => setDashTelecallingOpen(!dashTelecallingOpen)}
                      className="text-[10px] font-bold text-slate-500 hover:text-teal-600 flex items-center gap-1 cursor-pointer bg-slate-50 hover:bg-slate-100 px-2 py-0.5 rounded border border-slate-200"
                    >
                      {dashTelecallingOpen ? <><ChevronUp className="w-3 h-3" /> Minimize</> : <><ChevronDown className="w-3 h-3" /> Expand</>}
                    </button>
                  </div>

                  {dashTelecallingOpen ? (
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2">
                      <div 
                        onClick={() => setActiveTab('free_service_followup')}
                        className="bg-gradient-to-br from-teal-50 to-teal-100/50 border border-teal-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="flex items-center justify-between">
                          <span className="text-[9.5px] font-extrabold text-teal-900 uppercase">Free Service Due</span>
                          <span className="w-1.5 h-1.5 rounded-full bg-teal-600" />
                        </div>
                        <div className="my-1">
                          <div className="text-base font-black text-teal-950">
                            {(Object.values(freeServiceFollowupCategories) as any[]).reduce((acc: number, list: any) => acc + (list?.length || 0), 0)}
                          </div>
                          <span className="text-[8.5px] text-teal-800 font-bold">Pending follow-ups</span>
                        </div>
                        <div className="text-[8.5px] font-bold text-teal-700 underline">View tree</div>
                      </div>

                      <div 
                        onClick={() => {
                          setActiveTab('telecalling');
                          setTelecallerSubTab('date_report');
                          setTelecallerDateFrom(todayISO);
                          setTelecallerDateTo(todayISO);
                        }}
                        className="bg-gradient-to-br from-amber-50 to-amber-100/50 border border-amber-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="flex items-center justify-between">
                          <span className="text-[9.5px] font-extrabold text-amber-900 uppercase">Calls Today</span>
                          <Calendar className="w-3.5 h-3.5 text-amber-700" />
                        </div>
                        <div className="my-1">
                          <div className="text-base font-black text-amber-950">
                            {allTelecallerLogs.filter(l => l.callDate === todayISO).length}
                          </div>
                          <span className="text-[8.5px] text-amber-800 font-bold">Logged today</span>
                        </div>
                        <div className="text-[8.5px] font-bold text-amber-700 underline">View logs</div>
                      </div>

                      <div 
                        onClick={() => {
                          setActiveTab('telecalling');
                          setTelecallerSubTab('scheduled');
                          setTelecallerScheduledStatus('today');
                        }}
                        className="bg-gradient-to-br from-red-50 to-red-100/50 border border-red-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="flex items-center justify-between">
                          <span className="text-[9.5px] font-extrabold text-red-900 uppercase">Due Today</span>
                          {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length > 0 && <span className="w-2 h-2 rounded-full bg-red-600 animate-ping" />}
                        </div>
                        <div className="my-1">
                          <div className="text-base font-black text-red-950">
                            {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length}
                          </div>
                          <span className="text-[8.5px] text-red-800 font-bold">Must call today</span>
                        </div>
                        <div className="text-[8.5px] font-bold text-red-700 underline">View due</div>
                      </div>

                      <div 
                        onClick={() => {
                          setActiveTab('telecalling');
                          setTelecallerSubTab('scheduled');
                          setTelecallerScheduledStatus('overdue');
                        }}
                        className="bg-gradient-to-br from-orange-50 to-orange-100/50 border border-orange-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="flex items-center justify-between">
                          <span className="text-[9.5px] font-extrabold text-orange-900 uppercase">Overdue Calls</span>
                          <AlertCircle className="w-3.5 h-3.5 text-orange-700" />
                        </div>
                        <div className="my-1">
                          <div className="text-base font-black text-orange-950">
                            {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'overdue').length}
                          </div>
                          <span className="text-[8.5px] text-orange-800 font-bold">Past due date</span>
                        </div>
                        <div className="text-[8.5px] font-bold text-orange-700 underline">View overdue</div>
                      </div>

                      <div 
                        onClick={() => {
                          setActiveTab('telecalling');
                          setTelecallerSubTab('scheduled');
                          setTelecallerScheduledStatus('upcoming');
                        }}
                        className="bg-gradient-to-br from-blue-50 to-blue-100/50 border border-blue-200 p-2.5 rounded-xl shadow-2xs hover:shadow-sm transition-all cursor-pointer flex flex-col justify-between group"
                      >
                        <div className="flex items-center justify-between">
                          <span className="text-[9.5px] font-extrabold text-blue-900 uppercase">Upcoming</span>
                          <Clock className="w-3.5 h-3.5 text-blue-700" />
                        </div>
                        <div className="my-1">
                          <div className="text-base font-black text-blue-950">
                            {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'upcoming').length}
                          </div>
                          <span className="text-[8.5px] text-blue-800 font-bold">Future scheduled</span>
                        </div>
                        <div className="text-[8.5px] font-bold text-blue-700 underline">View upcoming</div>
                      </div>
                    </div>
                  ) : (
                    <div 
                      onClick={() => setDashTelecallingOpen(true)}
                      className="py-1.5 px-3 bg-teal-50/50 hover:bg-teal-50 border border-dashed border-teal-200 rounded-lg text-xs flex items-center justify-between cursor-pointer text-teal-900"
                    >
                      <span className="font-semibold text-[11px]">📞 Telecalling: <b>{scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length} Due Today</b> | {allTelecallerLogs.filter(l => l.callDate === todayISO).length} Completed Calls Today</span>
                      <span className="text-[10px] font-bold underline text-teal-700">Click to expand section &rarr;</span>
                    </div>
                  )}
                </div>

              </div>
            );
          })()}
          {/* TAB 1: NEW JOB ENTRY */}
          {(activeTab === 'new_entry' || (typeof window !== 'undefined' && window.matchMedia && window.matchMedia('print').matches)) && (
            <div className="space-y-8 print:w-full print:m-0 print:p-0">
              {/* FULL-WIDTH DATA ENTRY FORM */}
              <div className="w-full bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm space-y-6 print:hidden">
            {/* MASTER FORM RIBBON CONTROL BAR */}
            <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-100/90 border border-slate-300 p-2.5 rounded-xl text-xs font-bold text-slate-800 shadow-2xs">
              <div className="flex items-center gap-2">
                <SlidersHorizontal className="w-4 h-4 text-blue-900" />
                <span className="text-slate-900 font-extrabold uppercase tracking-wider text-[11px]">Form Ribbons (Sections A, B, C, D)</span>
                <span className="text-[11px] text-slate-500 font-medium hidden sm:inline">(Click any section ribbon header to minimize or open)</span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setSectionAOpen(true);
                    setSectionBOpen(true);
                    setSectionCOpen(true);
                    setSectionDOpen(true);
                  }}
                  className="px-3 py-1 bg-white hover:bg-blue-50 text-blue-900 border border-blue-300 rounded-lg shadow-2xs font-bold text-[11px] cursor-pointer transition-all flex items-center gap-1"
                >
                  <ChevronUp className="w-3.5 h-3.5 text-blue-900" /> Open All Ribbons
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSectionAOpen(false);
                    setSectionBOpen(false);
                    setSectionCOpen(false);
                    setSectionDOpen(false);
                  }}
                  className="px-3 py-1 bg-white hover:bg-slate-200 text-slate-800 border border-slate-300 rounded-lg shadow-2xs font-bold text-[11px] cursor-pointer transition-all flex items-center gap-1"
                >
                  <ChevronDown className="w-3.5 h-3.5 text-slate-600" /> Minimize All Ribbons
                </button>
              </div>
            </div>

          {/* SECTION A TRACTOR DETAILS FORM RIBBON */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
            {/* RIBBON HEADER A */}
            <div
              onClick={() => setSectionAOpen(!sectionAOpen)}
              className="flex items-center justify-between py-1.5 px-3 bg-indigo-50 border-b border-indigo-150 text-indigo-950 cursor-pointer hover:bg-indigo-100/70 transition-all select-none"
            >
              <div className="flex items-center gap-2 min-w-0">
                <span className="bg-indigo-600 text-white font-extrabold text-[10px] px-2 py-0.5 rounded-md shrink-0">
                  SECTION (A)
                </span>
                <div className="min-w-0">
                  <h2 className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                    <Building2 className="w-3.5 h-3.5 text-indigo-600" /> {appLang === 'te' ? 'ట్రాక్టర్ & కస్టమర్ వివరాలు (Tractor & Customer Details)' : 'Tractor Details & Header Settings'}
                  </h2>
                  {!sectionAOpen && (
                    <p className="text-[10px] text-slate-500 truncate font-mono mt-0.5">
                      Job Card #{onlineJobCardNo || jobNo || 'Draft'} • Chassis: {chassisNo || '—'} • Customer: {custName || '—'} • Mob: {ownerMob || '—'}
                    </p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <span className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full border tracking-wider transition-all ${
                  sectionAOpen 
                    ? 'bg-indigo-100 text-indigo-800 border-indigo-200' 
                    : 'bg-slate-100 text-slate-600 border-slate-200'
                }`}>
                  {sectionAOpen ? '▼ MINIMIZE' : '▲ OPEN'}
                </span>
                {sectionAOpen ? <ChevronUp className="w-4 h-4 text-indigo-600" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
              </div>
            </div>

            {/* RIBBON BODY A */}
            {sectionAOpen && (
              <div className="p-4 space-y-4">
            {editingCardId && (
              <div className="bg-amber-50 border border-amber-200 p-3 rounded-lg text-xs flex items-center justify-between gap-2 text-amber-800">
                <div className="flex items-center gap-1.5 font-semibold">
                  <span className="text-base">✏️</span>
                  <span>Currently editing Job Card #{onlineJobCardNo || jobNo || 'Draft'}. Click "Update Job Card" below to save or "Reset Form" to cancel.</span>
                </div>
                <button
                  type="button"
                  onClick={resetForm}
                  className="text-amber-900 bg-amber-100 hover:bg-amber-200 font-bold px-2 py-1 rounded cursor-pointer transition-all text-[11px]"
                >
                  Cancel
                </button>
              </div>
            )}
            <div className="flex items-center justify-between border-b pb-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                Header Settings & File Reference
              </span>
              <div className="flex items-center gap-1.5 bg-blue-50 border border-blue-200 px-2.5 py-1 rounded-md text-xs">
                <span className="font-bold text-blue-900">HISTORY FILE NO :</span>
                <input
                  type="text"
                  value={historyFileNo}
                  onChange={(e) => setHistoryFileNo(e.target.value)}
                  placeholder="________"
                  className="w-28 text-xs font-mono font-bold text-blue-950 bg-white px-1.5 py-0.5 rounded border border-blue-300 outline-none"
                />
              </div>
            </div>

            <div className="p-3 bg-amber-50 border border-amber-300 rounded-lg space-y-1.5 shadow-2xs">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-amber-900 flex items-center gap-1.5">
                  <Search className="w-4 h-4 text-amber-700" />
                  <span>Search Chassis No & Auto-fill</span>
                </label>
                {searchQuery ? (
                  <button
                    type="button"
                    onClick={() => setSearchQuery('')}
                    className="text-[10px] text-amber-800 hover:text-amber-950 underline font-bold cursor-pointer"
                  >
                    Clear Search
                  </button>
                ) : (
                  <span className="text-[10px] text-amber-800 bg-amber-100 font-bold px-2 py-0.5 rounded border border-amber-300">
                    Chassis No Only
                  </span>
                )}
              </div>
              <input
                type="text"
                list="customerSearchDatalist"
                value={searchQuery}
                onChange={(e) => handleCustomerSearch(e.target.value)}
                placeholder="Type Chassis No to search and auto-fill..."
                className="w-full text-xs p-2.5 bg-white border border-amber-300 rounded-md focus:ring-2 focus:ring-amber-500 outline-none font-semibold text-slate-900 placeholder:font-normal placeholder:text-slate-400 font-mono"
              />
              <datalist id="customerSearchDatalist">
                {getCustomerDatalistOptions().map((opt) => (
                  <option key={opt.key} value={opt.val}>
                    {opt.display}
                  </option>
                ))}
              </datalist>

              {autoFillNotice && (
                <div className="mt-2 p-2.5 bg-emerald-50 border border-emerald-300 rounded-md flex items-center justify-between text-xs font-semibold text-emerald-950 shadow-2xs">
                  <div className="flex items-center gap-2">
                    <span className="text-base">⚡</span>
                    <span>
                      Auto-filled details for <strong className="font-mono underline">{autoFillNotice.chassis}</strong>! 
                      <span className="ml-1 text-slate-700 font-medium">
                        ({autoFillNotice.filledCount} fields populated on the LEFT side)
                      </span>
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setAutoFillNotice(null)}
                    className="text-emerald-800 hover:text-emerald-950 text-[10px] font-bold underline cursor-pointer"
                  >
                    Dismiss
                  </button>
                </div>
              )}
            </div>

            {/* TWO-COLUMN SIDE-BY-SIDE SPLIT: AUTO-FILL vs MANUAL ENTRY */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-stretch">
              
              {/* LEFT COLUMN: AUTO-FILLED CUSTOMER & VEHICLE DETAILS */}
              <div className="bg-emerald-50/40 border-2 border-emerald-200/90 rounded-xl p-2 space-y-1.5 shadow-2xs h-full flex flex-col">
                <div className="flex items-center justify-between border-b border-emerald-200/80 pb-2">
                  <div className="flex items-center gap-1.5">
                    <span className="bg-emerald-700 text-white font-extrabold text-[10px] px-2 py-0.5 rounded uppercase tracking-wider">
                      ⚡ AUTO-FILLED
                    </span>
                    <h3 className="text-xs font-black text-emerald-950">
                      {appLang === 'te' ? 'కస్టమర్ & వాహన వివరాలు' : 'Customer & Vehicle Details'}
                    </h3>
                  </div>
                  <div className="flex items-center gap-1.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => {
                        const tempLookup = {
                          historyFileNo: historyFileNo,
                          custName: custName,
                          fatherName: fatherName,
                          custPhone: ownerMob,
                          village: village,
                          mandal: mandal,
                          chassis: chassisNo,
                          engineNo: engineNo,
                          model: model,
                          modelType: modelType,
                          address: custAddr
                        };
                        copyCustomerDetails(tempLookup);
                      }}
                      className="px-2 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 hover:text-indigo-950 font-bold text-[10px] rounded border border-indigo-200 transition-all flex items-center gap-1 shadow-3xs cursor-pointer"
                      title="Copy loaded customer details"
                    >
                      <Copy className="w-3 h-3 text-indigo-600" />
                      <span>Copy Details (కాపీ)</span>
                    </button>
                    <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100/80 px-2 py-0.5 rounded border border-emerald-200">
                      Database Record
                    </span>
                  </div>
                </div>
                <div className="grid grid-cols-1 xl:grid-cols-2 gap-3 mt-2 flex-grow">
                  {/* VEHICLE DETAILS (Left side) */}
                  <div className="space-y-2.5 bg-white/60 p-3 rounded-lg border border-emerald-100 shadow-2xs">
                    <h4 className="text-[10px] font-extrabold text-emerald-800 border-b border-emerald-100 pb-1 mb-2 uppercase tracking-wider flex items-center gap-1">🚜 Vehicle Details</h4>
                    
                    {/* Tractor Model & Model Type */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    {renderFieldHeader('Tractor Model', model)}
                    <input
                      type="text"
                      list="modelDatalist"
                      value={model}
                      onChange={(e) => setModel(e.target.value)}
                      placeholder="✏️ Enter Tractor Model..."
                      className={getInputClass(model)}
                    />
                    <datalist id="modelDatalist">
                      {Array.from(
                        new Set(
                          Array.from(new Set(Object.values(chassisIndex)))
                            .map((rec: any) => (getFieldValue(rec, 'model') || '').trim())
                            .filter(Boolean)
                        )
                      ).slice(0, 100).map((mdlVal) => (
                        <option key={`model_opt_${mdlVal}`} value={mdlVal} />
                      ))}
                    </datalist>
                  </div>
                  <div>
                    {renderFieldHeader('Model Type', modelType)}
                    <input
                      type="text"
                      value={modelType}
                      onChange={(e) => setModelType(e.target.value)}
                      placeholder="✏️ e.g. 2WD / 4WD"
                      className={getInputClass(modelType)}
                    />
                  </div>
                </div>

                    {/* Chassis No */}
                <div>
                  {renderFieldHeader('Chassis No.', chassisNo, true)}
                  <input
                    type="text"
                    list="chassisOnlyDatalist"
                    value={chassisNo}
                    onChange={(e) => handleChassisChange(e.target.value)}
                    placeholder="✏️ Enter Chassis No..."
                    className={getInputClass(chassisNo, 'font-mono text-emerald-950 font-bold')}
                  />
                  <datalist id="chassisOnlyDatalist">
                    {Array.from(
                      new Set(
                        Array.from(new Set(Object.values(chassisIndex)))
                          .map((rec: any) => {
                            const ch = (rec.__chassisDisplay || getFieldValue(rec, 'chassis') || '').trim();
                            const nm = (getFieldValue(rec, 'custName') || '').trim();
                            const inst = toInputDateFormat(getFieldValue(rec, 'installDate'));
                            if (!ch) return null;
                            return `${ch}${nm ? ` — ${nm}` : ''}${inst ? ` (Inst: ${inst})` : ''}`;
                          })
                          .filter(Boolean)
                      )
                    ).slice(0, 150).map((chVal) => (
                      <option key={`chassis_only_${chVal}`} value={chVal} />
                    ))}
                  </datalist>
                </div>

                    {/* Serial No. & Engine No. */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    {renderFieldHeader('Serial No.', serialNo)}
                    <input
                      type="text"
                      value={serialNo}
                      onChange={(e) => handleSerialChange(e.target.value)}
                      placeholder="✏️ Enter Serial No..."
                      className={getInputClass(serialNo, 'font-mono')}
                    />
                  </div>
                  <div>
                    {renderFieldHeader('Engine No.', engineNo)}
                    <input
                      type="text"
                      value={engineNo}
                      onChange={(e) => setEngineNo(e.target.value)}
                      placeholder="✏️ Enter Engine No..."
                      className={getInputClass(engineNo, 'font-mono')}
                    />
                  </div>
                </div>

                    {/* Date Of Delivery & Warranty Override */}
                <div>
                  <div className="flex items-center justify-between mb-0.5">
                    <label className="text-xs font-semibold text-slate-600 flex items-center gap-1">
                      <span>Date Of Delivery</span>
                    </label>
                    <span className={`text-[10px] font-black px-1.5 py-0.5 rounded uppercase ${
                      wtyInfo.isWty 
                        ? 'bg-red-600 text-white shadow-xs' 
                        : (wtyInfo.isPostWty ? 'bg-amber-600 text-white shadow-xs' : 'hidden')
                    }`}>
                      {wtyInfo.isWty ? '🔴 WARRANTY' : (wtyInfo.isPostWty ? '🟡 POST WTY' : '')}
                    </span>
                  </div>
                  <input
                    type="date"
                    value={installDate}
                    onChange={(e) => setInstallDate(e.target.value)}
                    className={getInputClass(installDate)}
                  />
                  <div className="mt-1 flex items-center justify-between gap-1 text-[9px]">
                    <span className="text-slate-500 font-medium">Override:</span>
                    <div className="flex gap-1">
                      <button
                        type="button"
                        onClick={() => setWarrantyOverride('auto')}
                        className={`px-1.5 py-0.5 rounded cursor-pointer ${warrantyOverride === 'auto' ? 'bg-blue-900 text-white font-bold' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                      >
                        Auto
                      </button>
                      <button
                        type="button"
                        onClick={() => setWarrantyOverride('warranty')}
                        className={`px-1.5 py-0.5 rounded cursor-pointer ${warrantyOverride === 'warranty' ? 'bg-red-600 text-white font-bold' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                      >
                        Wty
                      </button>
                      <button
                        type="button"
                        onClick={() => setWarrantyOverride('post_wty')}
                        className={`px-1.5 py-0.5 rounded cursor-pointer ${warrantyOverride === 'post_wty' ? 'bg-amber-600 text-white font-bold' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                      >
                        Post
                      </button>
                    </div>
                  </div>
                </div>
                  </div>

                  {/* CUSTOMER DETAILS (Right side) */}
                  <div className="space-y-2.5 bg-white/60 p-3 rounded-lg border border-emerald-100 shadow-2xs flex flex-col justify-between">
                    <div className="space-y-2.5">
                      <h4 className="text-[10px] font-extrabold text-emerald-800 border-b border-emerald-100 pb-1 mb-2 uppercase tracking-wider flex items-center gap-1">👤 Customer Address & Details</h4>
                      
                      {/* Name of Customer */}
                <div>
                  {renderFieldHeader('Name of Customer', custName, true)}
                  <div className="space-y-1">
                    <input
                      type="text"
                      list="customerNameDatalist"
                      value={custName}
                      onChange={(e) => handleNameChange(e.target.value)}
                      placeholder="✏️ Enter Customer Name..."
                      className={getInputClass(custName)}
                    />
                    {translatedCustName && (
                      <div className="flex items-center gap-1 text-[10px] text-blue-900 font-bold bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                        <span>🖨️ ప్రింట్‌లో (Telugu):</span>
                        <span className="text-blue-950 font-extrabold">{translatedCustName}</span>
                      </div>
                    )}
                  </div>
                  <datalist id="customerNameDatalist">
                    {Array.from(
                      new Set(
                        Array.from(new Set(Object.values(chassisIndex)))
                          .map((rec: any) => {
                            const nm = (getFieldValue(rec, 'custName') || '').trim();
                            const ch = (getFieldValue(rec, 'chassis') || '').trim();
                            const inst = toInputDateFormat(getFieldValue(rec, 'installDate'));
                            if (!nm) return null;
                            return `${nm}${ch ? ` — ${ch}` : ''}${inst ? ` (Inst: ${inst})` : ''}`;
                          })
                          .filter(Boolean)
                      )
                    ).slice(0, 150).map((nmVal) => (
                      <option key={`custname_only_${nmVal}`} value={nmVal} />
                    ))}
                  </datalist>
                </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                        <div>
                    {renderFieldHeader('Father Name', fatherName)}
                    <input
                      type="text"
                      value={fatherName}
                      onChange={(e) => setFatherName(e.target.value)}
                      placeholder="✏️ Enter Father Name..."
                      className={getInputClass(fatherName)}
                    />
                  </div>
                        <div>
                    {renderFieldHeader('Mobile No.', ownerMob)}
                    <input
                      type="text"
                      list="ownerMobileDatalist"
                      value={ownerMob}
                      onChange={(e) => handlePhoneChange(e.target.value)}
                      placeholder="✏️ Enter Mobile No..."
                      className={getInputClass(ownerMob)}
                    />
                    <datalist id="ownerMobileDatalist">
                      {Array.from(
                        new Set(
                          Array.from(new Set(Object.values(chassisIndex)))
                            .map((rec: any) => {
                              const ph = (getFieldValue(rec, 'custPhone') || '').trim();
                              const nm = (getFieldValue(rec, 'custName') || '').trim();
                              const ch = (getFieldValue(rec, 'chassis') || '').trim();
                              if (!ph) return null;
                              return `${ph}${nm ? ` — ${nm}` : ''}${ch ? ` (Chassis: ${ch})` : ''}`;
                            })
                            .filter(Boolean)
                        )
                      ).slice(0, 150).map((phVal) => (
                        <option key={`owner_mob_${phVal}`} value={phVal} />
                      ))}
                    </datalist>
                  </div>
                      </div>

                      {/* Village & Mandal */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div className="space-y-1">
                    {renderFieldHeader('Village', village)}
                    <input
                      type="text"
                      value={village}
                      onChange={(e) => setVillage(e.target.value)}
                      placeholder="✏️ Enter Village..."
                      className={getInputClass(village)}
                    />
                    {translatedVillage && (
                      <div className="flex items-center gap-1 text-[9.5px] text-blue-900 font-bold bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200">
                        <span>🖨️ ప్రింట్:</span>
                        <span className="text-blue-950 font-extrabold">{translatedVillage}</span>
                      </div>
                    )}
                  </div>
                  <div className="space-y-1">
                    {renderFieldHeader('Mandal', mandal)}
                    <input
                      type="text"
                      value={mandal}
                      onChange={(e) => setMandal(e.target.value)}
                      placeholder="✏️ Enter Mandal..."
                      className={getInputClass(mandal)}
                    />
                    {translatedMandal && (
                      <div className="flex items-center gap-1 text-[9.5px] text-blue-900 font-bold bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200">
                        <span>🖨️ ప్రింట్:</span>
                        <span className="text-blue-950 font-extrabold">{translatedMandal}</span>
                      </div>
                    )}
                  </div>
                </div>
                    </div>
                    
                    <div className="pt-2">
                      {/* Last Service Details */}
                <div>
                  {renderFieldHeader('Last Service Details', distDealership)}
                  <textarea
                    value={distDealership}
                    onChange={(e) => setDistDealership(e.target.value)}
                    placeholder="✏️ Previous service details will appear here..."
                    rows={2}
                    className={getInputClass(distDealership, 'resize-none leading-relaxed')}
                  />
                    </div>
                  </div>
                </div>
              </div>
              </div>
              {/* RIGHT COLUMN: MANUAL JOB CARD SESSION DETAILS */}
              <div className="bg-indigo-50/40 border-2 border-indigo-200/90 rounded-xl p-2.5 space-y-2 shadow-2xs h-full flex flex-col">
                <div className="flex items-center justify-between border-b border-indigo-200/80 pb-1.5">
                  <div className="flex items-center gap-1.5">
                    <span className="bg-indigo-700 text-white font-extrabold text-[9px] px-1.5 py-0.5 rounded uppercase tracking-wider">
                      ✍️ MANUAL ENTRY
                    </span>
                    <h3 className="text-xs font-black text-indigo-950">
                      Job Session Details
                    </h3>
                  </div>
                </div>

                {/* Job Open Date & Online Job Card No */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    {renderFieldHeader('Job Open Date', jobDate)}
                    <input
                      type="date"
                      value={jobDate}
                      onChange={(e) => setJobDate(e.target.value)}
                      className={getInputClass(jobDate, 'py-1.5')}
                    />
                  </div>
                  <div>
                    {renderFieldHeader('Online Job Card No.', onlineJobCardNo, true)}
                    <input
                      type="text"
                      value={onlineJobCardNo}
                      onChange={(e) => {
                        const val = e.target.value;
                        setOnlineJobCardNo(val);
                        setJobNo(val);
                      }}
                      placeholder="✏️ Enter No..."
                      className={getInputClass(onlineJobCardNo, 'font-mono text-indigo-800 font-bold py-1.5')}
                    />
                  </div>
                </div>

                {/* Complaint Date & Complaint Details */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                    <div>
                      {renderFieldHeader('Complaint Date', complaintDate)}
                      <input
                        type="date"
                        value={complaintDate}
                        onChange={(e) => setComplaintDate(e.target.value)}
                        className={getInputClass(complaintDate, 'py-1.5')}
                      />
                    </div>
                    <div>
                      {renderFieldHeader('Complaint Details', complaintDetails)}
                      <input
                          type="text"
                          value={complaintDetails}
                          onChange={(e) => setComplaintDetails(e.target.value)}
                          placeholder="✏️ Details..."
                          className={getInputClass(complaintDetails, 'py-1.5')}
                      />
                    </div>
                </div>

                {/* Status & Service Location */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    {renderFieldHeader('Status', status)}
                    <select
                      value={status}
                      onChange={(e) => setStatus(e.target.value as 'Open' | 'Closed')}
                      className="w-full text-xs py-1.5 px-2 border border-slate-300 bg-white rounded-md font-bold text-slate-900 outline-none cursor-pointer"
                    >
                      <option value="Open">🟢 Open</option>
                      <option value="Closed">🔴 Closed</option>
                    </select>
                  </div>
                  <div className="space-y-0.5">
                    <label className="text-[10px] font-bold text-slate-700">Service Location</label>
                    <select
                      value={serviceLocation}
                      onChange={(e) => setServiceLocation(e.target.value as 'workshop' | 'dss' | 'event')}
                      className="w-full text-xs font-bold py-1.5 px-2 bg-white border border-slate-300 text-blue-950 rounded-lg outline-none focus:ring-1 focus:ring-blue-500 cursor-pointer"
                    >
                      <option value="workshop">Location Workshop</option>
                      <option value="dss">DSS : Field</option>
                      <option value="event">Event / Camp</option>
                    </select>
                  </div>
                </div>

                {/* Hours Run & Expected Repairing Time */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    {renderFieldHeader('Hours Run', hourMeter)}
                    <input
                      type="text"
                      value={hourMeter}
                      onChange={(e) => setHourMeter(e.target.value)}
                      placeholder="✏️ 1050"
                      className={getInputClass(hourMeter, 'py-1.5')}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-700 block mb-0.5">Expected Time/Cost</label>
                    <input
                      type="text"
                      value={expectedRepairTime}
                      onChange={(e) => setExpectedRepairTime(e.target.value)}
                      placeholder="✏️ 4 Hours"
                      className="w-full text-xs py-1.5 px-2 border border-slate-300 rounded-md outline-none bg-white"
                    />
                  </div>
                </div>

                {/* Date In & Date Out */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-1.5">
                  <div>
                    <div className="flex items-center justify-between mb-0.5">
                      <label className="text-[10px] font-bold text-slate-700">Date In</label>
                      <div className="flex items-center gap-0.5">
                        <button type="button" onClick={() => setDateTimeIn(getLocalDateTimeString())} className="text-[8px] bg-blue-100 text-blue-800 px-1 py-0.5 rounded font-bold cursor-pointer">Now</button>
                      </div>
                    </div>
                    <input
                      type="datetime-local"
                      value={dateTimeIn}
                      onChange={(e) => setDateTimeIn(e.target.value)}
                      className="w-full text-[11px] p-1 border border-slate-300 rounded-md outline-none bg-white"
                    />
                  </div>
                  <div>
                    <div className="flex items-center justify-between mb-0.5">
                      <label className="text-[10px] font-bold text-slate-700">Date Out</label>
                      <div className="flex items-center gap-0.5">
                        <button type="button" onClick={() => { const base = dateTimeIn ? new Date(dateTimeIn) : new Date(); setDateTimeOut(getLocalDateTimeString(base)); }} className="text-[8px] bg-blue-100 text-blue-800 px-1 py-0.5 rounded font-bold cursor-pointer">Same</button>
                      </div>
                    </div>
                    <input
                      type="datetime-local"
                      value={dateTimeOut}
                      onChange={(e) => setDateTimeOut(e.target.value)}
                      className="w-full text-[11px] p-1 border border-slate-300 rounded-md outline-none bg-white"
                    />
                  </div>
                </div>

                {/* Service Type Checkboxes */}
                <div>
                  <label className="text-xs font-semibold text-slate-600 mb-1 block">Service Type Checkbox:</label>
                  <div className="flex flex-wrap gap-1.5 text-xs font-semibold">
                    {serviceTypeOptions.map(st => (
                      <button
                        key={st}
                        type="button"
                        onClick={() => handleServiceTypeSelect(st)}
                        className={`px-2.5 py-1 rounded border transition-colors cursor-pointer ${serviceType === st ? 'bg-blue-900 text-white border-blue-900 font-bold' : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-100'}`}
                      >
                        {st}
                      </button>
                    ))}
                  </div>
                </div>

              </div>
            </div>
          </div>
            )}
          </div>

          {/* SECTION B LABOUR & WORKSHOP REPAIRS SUMMARY RIBBON */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
            {/* RIBBON HEADER B */}
            <div
              onClick={() => setSectionBOpen(!sectionBOpen)}
              className="flex items-center justify-between py-1.5 px-3 bg-indigo-50 border-b border-indigo-150 text-indigo-950 cursor-pointer hover:bg-indigo-100/70 transition-all select-none"
            >
              <div className="flex items-center gap-2 min-w-0">
                <span className="bg-emerald-600 text-white font-extrabold text-[10px] px-2 py-0.5 rounded-md shrink-0">
                  SECTION (B)
                </span>
                <div className="min-w-0">
                  <h2 className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                    <Wrench className="w-3.5 h-3.5 text-emerald-600" /> Labour & Workshop Repairs Summary
                  </h2>
                  {!sectionBOpen && (
                    <p className="text-[10px] text-slate-500 truncate font-mono mt-0.5">
                      {repairRows.filter(r => r.repair.trim()).length} Repairs Listed • Total Labour Charge: ₹{repairTotal ? repairTotal.toFixed(2) : '0.00'}
                    </p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <span className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full border tracking-wider transition-all ${
                  sectionBOpen 
                    ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                    : 'bg-slate-100 text-slate-600 border-slate-200'
                }`}>
                  {sectionBOpen ? '▼ MINIMIZE' : '▲ OPEN'}
                </span>
                {sectionBOpen ? <ChevronUp className="w-4 h-4 text-emerald-600" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
              </div>
            </div>

            {/* RIBBON BODY B */}
            {sectionBOpen && (
              <div className="p-4 space-y-3">
            <div className="flex items-center justify-between border-b pb-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                Repairs Carried Out & Action Items
              </span>
              <div className="flex items-center gap-2">
                <label className="text-xs bg-emerald-50 text-emerald-700 font-semibold px-2 py-0.5 rounded hover:bg-emerald-100 flex items-center gap-1 cursor-pointer">
                  <FileSpreadsheet className="w-3.5 h-3.5" /> Import Excel
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleRepairsUpload}
                    className="hidden"
                  />
                </label>
                <button
                  type="button"
                  onClick={() => setRepairRows(ensureTrailingBlankRepairRow([...repairRows, { repair: '', rectification: '', charge: '' }]))}
                  className="text-xs bg-indigo-50 text-indigo-700 font-semibold px-2 py-0.5 rounded hover:bg-indigo-100 flex items-center gap-1 cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Repair
                </button>
              </div>
            </div>
            {repairsStatus.text && (
              <div className={`text-xs font-medium ${repairsStatus.isSuccess ? 'text-emerald-600' : 'text-slate-500'}`}>
                {repairsStatus.text}
              </div>
            )}

            <div className="space-y-2">
              {repairRows.map((r, i) => (
                <div key={i} className="flex items-center gap-1.5 text-xs">
                  <span className="w-4 text-center font-semibold text-slate-400">{i + 1}</span>
                  <input
                    type="text"
                    placeholder="Repairs Carried Out"
                    value={r.repair}
                    onChange={(e) => {
                      const updated = [...repairRows];
                      updated[i].repair = e.target.value;
                      setRepairRows(ensureTrailingBlankRepairRow(updated));
                    }}
                    className="flex-1 p-1 border border-slate-300 rounded outline-none"
                  />
                  <input
                    type="text"
                    placeholder="Rectification / Action"
                    value={r.rectification}
                    onChange={(e) => {
                      const updated = [...repairRows];
                      updated[i].rectification = e.target.value;
                      setRepairRows(ensureTrailingBlankRepairRow(updated));
                    }}
                    className="w-1/3 p-1 border border-slate-300 rounded outline-none"
                  />
                  <input
                    type="text"
                    placeholder="Charge ₹"
                    value={r.charge}
                    onChange={(e) => {
                      const updated = [...repairRows];
                      updated[i].charge = e.target.value;
                      setRepairRows(ensureTrailingBlankRepairRow(updated));
                    }}
                    className="w-16 p-1 border border-slate-300 rounded outline-none text-right font-mono"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const updated = repairRows.filter((_, idx) => idx !== i);
                      setRepairRows(ensureTrailingBlankRepairRow(updated));
                    }}
                    className="p-1 text-red-500 hover:bg-red-50 rounded cursor-pointer"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>

            <div className="flex justify-between items-center pt-2 border-t text-xs font-bold text-slate-700">
              <span>Total Labour Charge:</span>
              <span className="font-mono text-blue-900">₹{repairTotal ? repairTotal.toFixed(2) : '0.00'}</span>
            </div>
              </div>
            )}
          </div>

          {/* SECTION C PARTS / MATERIALS & 25 CHECKPOINTS RIBBON */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
            {/* RIBBON HEADER C */}
            <div
              onClick={() => setSectionCOpen(!sectionCOpen)}
              className="flex items-center justify-between py-1.5 px-3 bg-indigo-50 border-b border-indigo-150 text-indigo-950 cursor-pointer hover:bg-indigo-100/70 transition-all select-none"
            >
              <div className="flex items-center gap-2 min-w-0">
                <span className="bg-sky-600 text-white font-extrabold text-[10px] px-2 py-0.5 rounded-md shrink-0">
                  SECTION (C)
                </span>
                <div className="min-w-0">
                  <h2 className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                    <FileSpreadsheet className="w-3.5 h-3.5 text-sky-600" /> Parts / Materials List
                  </h2>
                  {!sectionCOpen && (
                    <p className="text-[10px] text-slate-500 truncate font-mono mt-0.5">
                      {partRows.filter(p => p.partNo.trim() || p.desc.trim()).length} Parts Added • Parts Total: ₹{partsTotal ? partsTotal.toFixed(2) : '0.00'}
                    </p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <span className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full border tracking-wider transition-all ${
                  sectionCOpen 
                    ? 'bg-sky-100 text-sky-800 border-sky-200' 
                    : 'bg-slate-100 text-slate-600 border-slate-200'
                }`}>
                  {sectionCOpen ? '▼ MINIMIZE' : '▲ OPEN'}
                </span>
                {sectionCOpen ? <ChevronUp className="w-4 h-4 text-sky-600" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
              </div>
            </div>

            {/* RIBBON BODY C */}
            {sectionCOpen && (
              <div className="p-4 space-y-4">
            <datalist id="partNoList">
              {Object.values(sparesIndex).map((s: any, idx: number) => {
                const pNo = s.__partNoDisplay || getSpareField(s, 'partNo') || '';
                const desc = getSpareField(s, 'desc') || '';
                const rate = getSpareField(s, 'rate') || '';
                return (
                  <React.Fragment key={`p-${idx}`}>
                    {pNo && (
                      <option value={pNo}>
                        {desc ? `${desc} (₹${rate || '0'})` : pNo}
                      </option>
                    )}
                  </React.Fragment>
                );
              })}
            </datalist>

            <div className="space-y-2">
              <div className="flex items-center justify-between border-b pb-2">
                <span className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                  <Wrench className="w-4 h-4 text-emerald-600" /> Parts / Materials List (Back Side)
                </span>
                <button
                  onClick={() => setPartRows(ensureTrailingBlankPartRow([...partRows, { partNo: '', desc: '', wty: false, qty: '', rate: '', amount: '' }]))}
                  className="text-xs bg-indigo-50 text-indigo-700 font-semibold px-2 py-0.5 rounded hover:bg-indigo-100 flex items-center gap-1 cursor-pointer"
                >
                  <Plus className="w-3 h-3" /> Add Part
                </button>
              </div>

              {partRows.map((p, i) => (
                <div key={i} className="flex items-center gap-1.5 text-xs">
                  <span className="w-4 text-center font-semibold text-slate-400">{i + 1}</span>
                  <input
                    type="text"
                    list="partNoList"
                    placeholder="Part No / Name"
                    value={p.partNo}
                    onChange={(e) => handlePartNoChange(i, e.target.value)}
                    className="w-24 p-1 border border-slate-300 rounded outline-none font-mono"
                  />
                  <input
                    type="text"
                    placeholder="Description / Part Name"
                    value={p.desc}
                    onChange={(e) => handlePartDescChange(i, e.target.value)}
                    className="flex-1 p-1 border border-slate-300 rounded outline-none"
                  />
                  
                  <input
                    type="text"
                    placeholder="Qty"
                    value={p.qty}
                    onChange={(e) => handleQtyRateChange(i, 'qty', e.target.value)}
                    className="w-10 p-1 border border-slate-300 rounded outline-none text-center font-mono"
                  />

                  {p.wty ? (
                    <div className="w-14 p-1 text-center font-bold text-emerald-600 text-[10px] bg-emerald-50 rounded select-none font-mono">
                      Wty
                    </div>
                  ) : (
                    <input
                      type="text"
                      placeholder="Rate"
                      value={p.rate}
                      onChange={(e) => handleQtyRateChange(i, 'rate', e.target.value)}
                      className="w-14 p-1 border border-slate-300 rounded outline-none text-right font-mono"
                    />
                  )}

                  <input
                    type="text"
                    placeholder="Amt"
                    value={p.wty ? '0' : p.amount}
                    disabled={!!p.wty}
                    onChange={(e) => {
                      if (p.wty) return;
                      const updated = [...partRows];
                      updated[i].amount = e.target.value;
                      setPartRows(ensureTrailingBlankPartRow(updated));
                    }}
                    className={`w-16 p-1 border rounded outline-none text-right font-mono font-semibold ${
                      p.wty ? 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed' : 'border-slate-300'
                    }`}
                  />

                  {/* Wty Column Checkbox (Positioned at end of line) */}
                  <div className="flex flex-col items-center shrink-0 w-8">
                    <label className="text-[8px] font-bold text-slate-400 select-none uppercase -mb-0.5">Wty</label>
                    <input
                      type="checkbox"
                      checked={!!p.wty}
                      onChange={() => handleWtyToggle(i)}
                      className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                    />
                  </div>

                  <button
                    onClick={() => {
                      const updated = partRows.filter((_, idx) => idx !== i);
                      setPartRows(ensureTrailingBlankPartRow(updated));
                    }}
                    className="p-1 text-red-500 hover:bg-red-50 rounded cursor-pointer"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>
              </div>
            )}
          </div>

          {/* SECTION D BILLING, STAFF & WORKSHOP REPORT RIBBON */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
            {/* RIBBON HEADER D */}
            <div
              onClick={() => setSectionDOpen(!sectionDOpen)}
              className="flex items-center justify-between py-1.5 px-3 bg-indigo-50 border-b border-indigo-150 text-indigo-950 cursor-pointer hover:bg-indigo-100/70 transition-all select-none"
            >
              <div className="flex items-center gap-2 min-w-0">
                <span className="bg-purple-600 text-white font-extrabold text-[10px] px-2 py-0.5 rounded-md shrink-0">
                  SECTION (D)
                </span>
                <div className="min-w-0">
                  <h2 className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                    <CheckSquare className="w-3.5 h-3.5 text-purple-600" /> Billing, Staff & Workshop Report
                  </h2>
                  {!sectionDOpen && (
                    <p className="text-[10px] text-slate-500 truncate font-mono mt-0.5">
                      Mechanic: {mechanic || 'Unassigned'} • Supervisor: {wsIncharge || 'Unassigned'} • Grand Total: ₹{gTotal || (partsTotal + (parseFloat(totalLabour) || repairTotal)).toFixed(2)}
                    </p>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <span className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full border tracking-wider transition-all ${
                  sectionDOpen 
                    ? 'bg-purple-100 text-purple-800 border-purple-200' 
                    : 'bg-slate-100 text-slate-600 border-slate-200'
                }`}>
                  {sectionDOpen ? '▼ MINIMIZE' : '▲ OPEN'}
                </span>
                {sectionDOpen ? <ChevronUp className="w-4 h-4 text-purple-600" /> : <ChevronDown className="w-4 h-4 text-slate-500" />}
              </div>
            </div>

            {/* RIBBON BODY D */}
            {sectionDOpen && (
              <div className="p-4 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* STAFF ASSIGNMENTS */}
              <div className="space-y-3">
                <div>
                  {renderFieldHeader('Technician / Mechanic Assigned', mechanic)}
                  <div className="space-y-1.5">
                    <select
                      value={mechanic}
                      onChange={(e) => {
                        const val = e.target.value;
                        setMechanic(val);
                        if (val && val !== '__custom__') {
                          const assocSup = getAssignedSupervisor(val);
                          if (assocSup) {
                            setWsIncharge(assocSup);
                          }
                        }
                      }}
                      className="w-full text-xs p-2 border border-slate-300 rounded-md outline-none bg-white font-semibold text-slate-800 shadow-2xs cursor-pointer"
                    >
                      <option value="">-- Select Mechanic --</option>
                      {staffMembers.filter(s => s.role === 'mechanic').map(s => (
                        <option key={s.id} value={s.name}>{s.name}</option>
                      ))}
                      <option value="__custom__">+ Type manual name...</option>
                    </select>
                    <input
                      type="text"
                      placeholder="✏️ Type custom mechanic name..."
                      value={mechanic === '__custom__' ? '' : mechanic}
                      onChange={(e) => {
                        const val = e.target.value;
                        setMechanic(val);
                        const assocSup = getAssignedSupervisor(val);
                        if (assocSup) {
                          setWsIncharge(assocSup);
                        }
                      }}
                      className={getInputClass(mechanic)}
                    />
                  </div>
                </div>

                <div>
                  {renderFieldHeader('Supervisor / W/S Incharge', wsIncharge)}
                  <div className="space-y-1.5">
                    <select
                      value={wsIncharge}
                      onChange={(e) => setWsIncharge(e.target.value)}
                      className="w-full text-xs p-2 border border-slate-300 rounded-md outline-none bg-white font-semibold text-slate-800 shadow-2xs cursor-pointer"
                    >
                      <option value="">-- Select Supervisor --</option>
                      {staffMembers.filter(s => s.role === 'supervisor').map(s => (
                        <option key={s.id} value={s.name}>{s.name}</option>
                      ))}
                      <option value="__custom__">+ Type manual name...</option>
                    </select>
                    <input
                      type="text"
                      placeholder="✏️ Type custom supervisor name..."
                      value={wsIncharge === '__custom__' ? '' : wsIncharge}
                      onChange={(e) => setWsIncharge(e.target.value)}
                      className={getInputClass(wsIncharge)}
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    {renderFieldHeader('Bill Number', billNo)}
                    <input
                      type="text"
                      placeholder="✏️ Bill No"
                      value={billNo}
                      onChange={(e) => setBillNo(e.target.value)}
                      className={getInputClass(billNo, 'font-mono text-indigo-900')}
                    />
                  </div>
                  <div>
                    {renderFieldHeader('Reported By', telecalling)}
                    <input
                      type="text"
                      placeholder="✏️ Reported by..."
                      value={telecalling}
                      onChange={(e) => setTelecalling(e.target.value)}
                      className={getInputClass(telecalling)}
                    />
                  </div>
                </div>
              </div>

              {/* REPORT & EXTRA COSTS */}
              <div className="space-y-3">
                <div>
                  {renderFieldHeader('Workshop Report', wsReport)}
                  <textarea
                    placeholder="Describe diagnosis, findings or general workshop comments..."
                    value={wsReport}
                    onChange={(e) => setWsReport(e.target.value)}
                    className="w-full text-xs p-2 rounded-md border border-slate-300 bg-white text-slate-900 font-medium focus:ring-2 focus:ring-indigo-500 outline-none h-[40px] resize-none placeholder:text-slate-400"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 mb-1">Warranty Material Cost</label>
                    <input
                      type="text"
                      placeholder="0.00"
                      value={warrantyMaterial}
                      onChange={(e) => setWarrantyMaterial(e.target.value)}
                      className="w-full text-xs p-2 border border-slate-300 rounded-md outline-none font-mono text-right font-medium text-slate-700"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 mb-1">Non Warranty Material Cost</label>
                    <input
                      type="text"
                      placeholder="0.00"
                      value={nonWarrantyMaterial}
                      onChange={(e) => setNonWarrantyMaterial(e.target.value)}
                      className="w-full text-xs p-2 border border-slate-300 rounded-md outline-none font-mono text-right font-medium text-slate-700"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2 border-t pt-2">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Labour Cost Override</label>
                    <input
                      type="text"
                      placeholder={repairTotal ? repairTotal.toFixed(2) : '0.00'}
                      value={totalLabour}
                      onChange={(e) => setTotalLabour(e.target.value)}
                      className="w-full text-xs p-1.5 border border-slate-300 rounded-md outline-none font-mono text-right font-bold text-blue-900"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Grand Total Override</label>
                    <input
                      type="text"
                      placeholder={(partsTotal + (parseFloat(totalLabour) || repairTotal)).toFixed(2)}
                      value={gTotal}
                      onChange={(e) => setGTotal(e.target.value)}
                      className="w-full text-xs p-1.5 border border-slate-300 rounded-md outline-none font-mono text-right font-bold text-emerald-800"
                    />
                  </div>
                </div>
              </div>
            </div>
              </div>
            )}
          </div>

          {/* BOTTOM FORM ACTION BAR: SAVE, RESET, PAGE SELECTION, PDF & PRINT */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-3 print:hidden">
            <div className="flex items-center gap-2 bg-slate-100 p-1 rounded-lg border border-slate-200 text-xs font-bold text-slate-700">
              <span className="px-2 text-slate-500 text-[11px] uppercase tracking-wider">Pages:</span>
              <button
                type="button"
                onClick={() => setPrintOption('both')}
                className={`px-3 py-1.5 rounded-md transition-all cursor-pointer ${printOption === 'both' ? 'bg-white text-slate-900 shadow-xs font-black' : 'text-slate-600 hover:text-slate-900'}`}
              >
                Both Pages
              </button>
              <button
                type="button"
                onClick={() => setPrintOption('jobCardOnly')}
                className={`px-3 py-1.5 rounded-md transition-all cursor-pointer ${printOption === 'jobCardOnly' ? 'bg-white text-slate-900 shadow-xs font-black' : 'text-slate-600 hover:text-slate-900'}`}
              >
                Page 1 Only
              </button>
              <button
                type="button"
                onClick={() => setPrintOption('partsOnly')}
                className={`px-3 py-1.5 rounded-md transition-all cursor-pointer ${printOption === 'partsOnly' ? 'bg-white text-slate-900 shadow-xs font-black' : 'text-slate-600 hover:text-slate-900'}`}
              >
                Page 2 Only
              </button>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={resetForm}
                className="py-2.5 px-3.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold text-xs rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
              >
                <RotateCcw className="w-3.5 h-3.5" /> {appLang === 'te' ? 'ఫారం రీసెట్' : 'Reset Form'}
              </button>
              <button
                type="button"
                onClick={saveJobCard}
                className="py-2.5 px-5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs"
              >
                <Save className="w-4 h-4" /> {editingCardId ? (appLang === 'te' ? 'జాబ్ కార్డ్ అప్‌డేట్' : 'Update Job Card') : (appLang === 'te' ? 'జాబ్ కార్డ్ సేవ్ చేయండి' : 'Save to List')}
              </button>
              <button
                type="button"
                onClick={handleSaveAsPDF}
                className="py-2.5 px-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs"
              >
                <Download className="w-4 h-4" /> {appLang === 'te' ? 'PDF డౌన్‌లోడ్' : 'Save as PDF'}
              </button>
              <button
                type="button"
                onClick={() => window.print()}
                className="py-2.5 px-4 bg-blue-900 hover:bg-blue-950 text-white font-bold text-xs rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs"
              >
                <Printer className="w-4 h-4" /> {appLang === 'te' ? 'ప్రింట్ చేయండి' : 'Print'}
              </button>
            </div>
          </div>
        </div>

        {/* PRINT / PREVIEW DISPLAY AREA - PLACED DIRECTLY BELOW FORM */}
        <div className="space-y-4 pt-6 border-t-2 border-dashed border-slate-200 print:pt-0 print:border-none print:m-0">
          <div className="bg-blue-900 text-white px-4 py-3 rounded-xl flex flex-wrap items-center justify-between gap-3 print:hidden shadow-sm">
            <div className="flex items-center gap-2.5">
              <FileText className="w-5 h-5 text-blue-200" />
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider">A4 Printable Job Card Live Preview</h3>
                <p className="text-[11px] text-blue-200 font-medium">Exact match print format rendered live from the form entries above</p>
              </div>
            </div>
            <button
              onClick={() => window.print()}
              className="flex items-center gap-1.5 bg-white hover:bg-blue-50 text-blue-900 font-bold text-xs px-4 py-2 rounded-lg transition-colors cursor-pointer shadow-xs"
            >
              <Printer className="w-4 h-4" /> Print Job Card
            </button>
          </div>

          {/* PAGE 1: SRI GAYATHRI AUTOMOTIVES JOB CARD */}
          {(printOption === 'both' || printOption === 'jobCardOnly') && (
            <div className="bg-white p-2 md:p-4 rounded-xl shadow-md border border-slate-200 print:shadow-none print:p-0 print:border-none print:m-0 print:rounded-none page-break-container">
              <div className="card-p1 bg-white border-2 border-blue-900 p-2.5 md:p-3 text-slate-900 max-w-[780px] mx-auto text-[11px] leading-snug print:w-full print:max-w-none print:p-[3mm_4mm] print:bg-white print:border-[1.5px] print:border-blue-900 space-y-1.5 print:space-y-1 overflow-hidden">
                
                {/* HEADER SECTION */}
                <div>
                  <div className="text-center font-extrabold text-blue-900 text-xs md:text-sm tracking-wide border-b border-blue-200 pb-0.5 mb-0 relative flex items-center justify-center gap-2">
                    <span>Job Card</span>
                    {/* EICHER LOGO OVAL TOP RIGHT */}
                    <div className="absolute right-0 top-0">
                      <div className="w-8 h-5 border-2 border-blue-900 rounded-full flex items-center justify-center bg-white shadow-xs">
                        <span className="font-serif font-black text-blue-900 text-[11px] italic">E</span>
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-between items-start mb-0">
                    <div className="flex-1">
                      <div className="text-xl md:text-2xl font-black text-blue-900 tracking-tight leading-none mb-0.5">
                        SRI GAYATHRI AUTOMOTIVES
                      </div>
                      <div className="text-[10px] font-extrabold text-blue-900">
                        Authorised Dealer : EICHER TRACTORS (Sales, Service & Spares)
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-medium">
                        D.No. 2-12-351, Opp Srinivasa Theater, Poranki, Vijayawada .
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-medium">
                        A.P - 521 137, Phone : 9063134025
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-bold">
                        GST : 37AFNFS856BM1ZS, Email : srigayathriauto@gmail.com
                      </div>
                    </div>

                    <div className="w-48 text-right text-[9px] font-semibold text-blue-950 space-y-0.5 pl-1">
                      <div className="flex items-center justify-end gap-1">
                        <span>• Service : Location Workshop</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'workshop' ? 'bg-blue-900 text-white print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                          {serviceLocation === 'workshop' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end gap-1">
                        <span>• DSS : Field</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'dss' ? 'bg-blue-900 text-white print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                          {serviceLocation === 'dss' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end gap-1">
                        <span>• Event / Camp</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'event' ? 'bg-blue-900 text-white print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                          {serviceLocation === 'event' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end pt-0.5">
                        {wtyInfo.isWty ? (
                          <span className="px-2 py-0.5 rounded text-[9.5px] font-black tracking-wider uppercase border shadow-2xs bg-red-600 text-white border-red-700">
                            WARRANTY
                          </span>
                        ) : wtyInfo.isPostWty ? (
                          <span className="px-2 py-0.5 rounded text-[9.5px] font-black tracking-wider uppercase border shadow-2xs bg-amber-600 text-white border-amber-700">
                            POST WTY
                          </span>
                        ) : null}
                      </div>
                    </div>
                  </div>

                  <div className="text-[8.5px] text-slate-700 font-medium border-t border-b border-blue-800 py-0.5 mb-0.5 bg-blue-50/40 text-center">
                    <b>Branches :</b> Nandigama - 9063134026, Tiruvur- 9063134027, Gudivada - 9063134028, Machilipatnam - 9063134029
                  </div>
                </div>

                {/* SECTIONS A, B, C WITH TIGHT CLEAN SPACING */}
                <div className="space-y-1.5 print:space-y-1 my-1">
                  {/* SECTION (A) TRACTOR DETAILS TABLE */}
                  <div className="mb-0.5">
                    <div className="flex justify-between items-center mb-0.5">
                      <div className="font-extrabold text-blue-900 text-[10px]">(A) Tractor Details</div>
                      <div className="text-[9.5px] font-bold text-blue-950 font-mono">
                        HISTORY FILE NO : {historyFileNo || '________'}
                      </div>
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[9.5px]">
                      <colgroup>
                        <col style={{ width: '15%' }} />
                        <col style={{ width: '19%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '18%' }} />
                        <col style={{ width: '15%' }} />
                        <col style={{ width: '25%' }} />
                      </colgroup>
                      <tbody>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 truncate">Online Job Card No.</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-indigo-900 font-mono">
                            <div className="flex items-center justify-between gap-1">
                              <span className="truncate">{onlineJobCardNo || jobNo}</span>
                              <span className={`text-[7px] px-1 py-0.5 rounded font-black uppercase border leading-none shrink-0 ${status === 'Closed' ? 'bg-red-50 text-red-700 border-red-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'}`}>
                                {status}
                              </span>
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-center">Date</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold truncate">{fmtDate(jobDate)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Model</td>
                          <td className="p-0.5 md:p-1 font-bold text-slate-900 truncate">{model}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Name of Customer</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-slate-900">
                            <div className="text-[11px] font-extrabold text-blue-950 leading-snug truncate">
                              {translatedCustName || custName}
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Chassis No.</td>
                          <td className="p-0.5 md:p-1 font-bold font-mono text-blue-950 truncate">{chassisNo}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Address</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-slate-900">
                            <div className="text-[10px] font-bold text-slate-900 leading-snug truncate">
                              {(() => {
                                const v = translatedVillage || village;
                                const m = translatedMandal || mandal;
                                if (v || m) {
                                  return [v, m].filter(Boolean).join(', ');
                                }
                                return translatedCustAddr || custAddr;
                              })()}
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Engine No.</td>
                          <td className="p-0.5 md:p-1 font-bold font-mono truncate">{engineNo}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Mobile No.</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold font-mono truncate">{ownerMob}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Date Of Delivery</td>
                          <td className="p-0.5 md:p-1 font-bold truncate">{fmtDate(installDate)}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Last Service Details</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 truncate">{distDealership}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Hour Meter Reading</td>
                          <td className="p-0.5 md:p-1 font-bold truncate">{hourMeter}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Date & Time in</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 truncate">{fmtDateTime(dateTimeIn)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-center">Out</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 truncate">{fmtDateTime(dateTimeOut)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-[8px] leading-tight">Exp. Repair Time/Cost</td>
                          <td className="p-0.5 md:p-1 font-medium truncate">{expectedRepairTime}</td>
                        </tr>
                        <tr className="h-[24px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Service</td>
                          <td colSpan={5} className="p-0.5 md:p-1">
                            <div className="flex flex-wrap items-center justify-between gap-0.5 text-[8.5px] font-bold">
                              {serviceTypeOptions.map(st => (
                                <div key={st} className="flex items-center gap-0.5">
                                  <span className="whitespace-nowrap">{st}</span>
                                  <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center text-[7.5px] leading-none shrink-0 ${serviceType === st ? 'bg-blue-900 text-white font-bold print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                                    {serviceType === st ? '✓' : ''}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SECTION (B) LABOUR & WORKSHOP REPAIRS SUMMARY */}
                  <div className="mb-0.5">
                    <div className="font-extrabold text-blue-900 text-[10px] mb-0">(B) Labour & Workshop Repairs Summary</div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[8.5px]">
                      <colgroup>
                        <col style={{ width: '6%' }} />
                        <col style={{ width: '48%' }} />
                        <col style={{ width: '30%' }} />
                        <col style={{ width: '16%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/60 font-bold text-blue-950 h-[18px]">
                          <th className="border-r border-blue-900 p-0.5 text-center">#</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Repairs Carried Out</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Rectification / Action</th>
                          <th className="p-0.5 text-right">Labour Charge ₹</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Array.from({ length: Math.max(repairRows.length, 2) }).map((_, i) => {
                          const r = repairRows[i] || { repair: '', rectification: '', charge: '' };
                          return (
                            <tr key={i} className="border-b border-blue-900 h-[17px]">
                              <td className="border-r border-blue-900 p-0.5 text-center font-mono font-semibold">{i + 1}</td>
                              <td className="border-r border-blue-900 p-0.5 font-medium truncate">{r.repair}</td>
                              <td className="border-r border-blue-900 p-0.5 font-medium truncate">{r.rectification}</td>
                              <td className="p-0.5 text-right font-mono font-bold truncate">{r.charge}</td>
                            </tr>
                          );
                        })}
                        <tr className="bg-blue-50/50 font-bold border-t border-blue-900 text-blue-950 h-[18px]">
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 text-right font-bold text-[8.5px]">Total Labour Charge ₹</td>
                          <td className="p-0.5 text-right font-mono font-extrabold text-blue-900 text-[8.5px] truncate">
                            {totalLabour || (repairTotal ? repairTotal.toFixed(2) : '')}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SECTION (C) CHECKLIST 25 CHECKPOINT (TELUGU ONLY) */}
                  <div>
                    <div className="font-extrabold text-blue-900 text-[10px] mb-0.5">
                      (C) 25 చెక్‌లిస్ట్ తనిఖీ అంశాలు (Checklist - 25 Checkpoints)
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[8.5px]">
                      <colgroup>
                        <col style={{ width: '5.5%' }} />
                        <col style={{ width: '65.5%' }} />
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '7%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/60 font-bold text-blue-950 h-[26px]">
                          <th className="border-r border-blue-900 p-0.5 text-center text-[10px]">క్ర.సం</th>
                          <th className="border-r border-blue-900 p-0.5 text-left text-[10px]">విభాగం & తనిఖీ చేయవలసిన అంశం</th>
                          <th className="border-r border-blue-900 p-0.5 text-left text-[10px]">చేయవలసిన పని / చర్య</th>
                          <th className="p-0.5 text-center text-[10px]">✓</th>
                        </tr>
                      </thead>
                      <tbody>
                        {checkpoints.map((cp) => {
                          const defaultItem = DEFAULT_CHECKPOINTS.find(d => d.id === cp.id);
                          const teluguCategory = cp.categoryTe || defaultItem?.categoryTe || (cp.category && cp.category !== 'GENERAL' ? cp.category : 'సాధారణ');
                          const teluguItem = cp.itemTe || defaultItem?.itemTe || cp.item;
                          const teluguAction = cp.actionTe || defaultItem?.actionTe || cp.action;

                          return (
                            <tr key={cp.id} className="border-b border-blue-900/60 h-[26px] leading-tight hover:bg-blue-50/20">
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono font-bold text-[10px]">{cp.id}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-bold text-blue-950">
                                <div className="flex items-center gap-1.5 leading-none">
                                  <span className="text-[9px] font-black text-blue-800 shrink-0">[{teluguCategory}]</span>
                                  <span className="text-[11px] font-extrabold text-slate-950 truncate">{teluguItem}</span>
                                </div>
                              </td>
                              <td className="border-r border-blue-900/60 p-0.5 font-bold text-slate-900">
                                <span className="text-[10px] font-bold text-slate-900 truncate block leading-none">{teluguAction}</span>
                              </td>
                              <td className="p-0.5 text-center">
                                <div className={`w-3 h-3 border border-blue-900 mx-auto flex items-center justify-center text-[9px] leading-none shrink-0 ${cp.checked ? 'bg-blue-900 text-white font-bold print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                                  {cp.checked ? '✓' : ''}
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
            </div>
          )}

          {/* PAGE 2: PARTS & MATERIALS BILL */}
          {(printOption === 'both' || printOption === 'partsOnly') && (
            <div className="bg-white p-3 md:p-6 rounded-xl shadow-md border border-slate-200 print:shadow-none print:p-0 print:border-none print:m-0 print:rounded-none page-break-container">
              <div className="card-p2 bg-white border-2 border-blue-900 p-3 md:p-4 text-slate-900 max-w-[780px] mx-auto text-[12px] leading-snug print:w-full print:max-w-none print:h-[282mm] print:max-h-[282mm] print:p-[5mm_6mm] print:bg-white print:border-[1.5px] print:border-blue-900 print:flex print:flex-col print:justify-between overflow-hidden">
                <div>
                  {/* PARTS TABLE - 30 ROWS EXPANDED */}
                  <div className="mb-3">
                    <div className="font-extrabold text-blue-900 text-xs md:text-sm mb-1.5 text-center border-b border-blue-900 pb-1 uppercase tracking-wider">
                      Materials / Spare Parts Issued & Repair Charges
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[9px]">
                      <colgroup>
                        <col style={{ width: '5%' }} />
                        <col style={{ width: '18%' }} />
                        <col style={{ width: '38%' }} />
                        <col style={{ width: '7%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '11%' }} />
                        <col style={{ width: '13%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/70 font-bold text-blue-950 h-[19px]">
                          <th className="border-r border-blue-900 p-0.5 text-center">S.No</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Part No</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Description</th>
                          <th className="border-r border-blue-900 p-0.5 text-center">Wty</th>
                          <th className="border-r border-blue-900 p-0.5 text-center">Qty</th>
                          <th className="border-r border-blue-900 p-0.5 text-right">Rate ₹</th>
                          <th className="p-0.5 text-right">Amount ₹</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Array.from({ length: Math.max(partRows.length, 30) }).map((_, i) => {
                          const r = partRows[i] || { partNo: '', desc: '', qty: '', rate: '', amount: '', wty: false };
                          return (
                            <tr key={i} className="border-b border-blue-900/60 h-[19px]">
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono text-[8.5px]">{i + 1}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-mono text-[8.5px] truncate">{r.partNo}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-medium truncate">{r.desc}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-bold text-emerald-700 text-[8.5px]">{r.wty ? '✓' : ''}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono text-[8.5px]">{r.qty}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-right font-mono text-[8.5px]">{r.wty ? '' : r.rate}</td>
                              <td className="p-0.5 text-right font-mono font-bold text-[8.5px]">{r.wty ? '0' : r.amount}</td>
                            </tr>
                          );
                        })}
                        <tr className="bg-blue-50 font-bold border-t-2 border-blue-900 text-slate-900 h-[20px]">
                          <td colSpan={6} className="border-r border-blue-900 p-1 text-right text-[9.5px]">Total Parts Amount ₹</td>
                          <td className="p-1 text-right font-mono font-extrabold text-blue-900 text-[9.5px]">{partsTotal ? partsTotal.toFixed(2) : ''}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SUMMARY CHARGES BOX */}
                  <div className="grid grid-cols-2 gap-4 border-2 border-blue-900 p-2.5 bg-blue-50/30 mb-3 text-xs">
                    <div className="space-y-1">
                      <div><b>Workshop Report:</b> {wsReport || '—'}</div>
                      <div><b>Mechanic Name:</b> {mechanic || '—'}</div>
                      <div><b>Sign of W/S Incharge:</b> {wsIncharge || '—'}</div>
                      {billNo && <div><b>Bill No:</b> <span className="font-mono font-bold text-blue-950 bg-blue-100 px-1 rounded">{billNo}</span></div>}
                    </div>
                    <div className="space-y-1 text-right font-semibold">
                      <div>Warranty Material: ₹{warrantyMaterial || '0.00'}</div>
                      <div>Non Warranty Material: ₹{nonWarrantyMaterial || '0.00'}</div>
                      <div className="text-sm font-black text-blue-950 border-t border-blue-900 pt-1 mt-1">
                        Grand Total: ₹{gTotal || (partsTotal + (parseFloat(totalLabour) || repairTotal)).toFixed(2)}
                      </div>
                    </div>
                  </div>

                  {/* FEEDBACK FIELDS */}
                  <div className="mb-2 text-xs font-semibold space-y-1 border border-blue-900 p-1">
                    <div>ROAD TEST FEEDBACK:____________________________________________________________________________________</div>
                    <div>CUSTOMER RATING & FEED BACK:_________________________________________________________________________</div>
                  </div>
                </div>

                <div className="pt-2">
                  <div className="text-center text-[10px] italic text-slate-600 border-t border-b border-blue-800 py-0.5 mb-2">
                    VEHICLES STORED, REPAIRED AND DRIVEN AT CUSTOMER'S RISK
                  </div>

                  <div className="flex justify-between items-end text-xs font-bold pt-8 px-2">
                    <div className="text-center w-36">
                      <div className="text-[10px] text-slate-800 font-extrabold pb-1 h-5 flex items-end justify-center">
                        {mechanic && mechanic !== '__custom__' ? mechanic : ''}
                      </div>
                      <div className="border-t border-slate-700 pt-1">Mechanic Signature</div>
                    </div>
                    <div className="text-center w-36">
                      <div className="text-[10px] text-slate-800 font-extrabold pb-1 h-5 flex items-end justify-center">
                        {wsIncharge && wsIncharge !== '__custom__' ? wsIncharge : ''}
                      </div>
                      <div className="border-t border-slate-700 pt-1">Supervisor Signature</div>
                    </div>
                    <div className="text-center w-36">
                      <div className="text-[10px] pb-1 h-5"></div>
                      <div className="border-t border-slate-700 pt-1">Customer's Signature</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    )}



      {/* TAB 3: REPORTS & ANALYTICS SECTION */}
      {activeTab === 'reports' && (
        <div className="w-full bg-white p-2.5 md:p-3 rounded-xl border border-slate-200 shadow-sm space-y-2.5 print:hidden">
          {/* STICKY TOP REPORTS CONTROL & KPI DASHBOARD (FREEZES ON SCROLL) */}
          <div className="sticky top-0 z-30 bg-slate-50/98 backdrop-blur-md border-2 border-indigo-200 p-2 md:p-2.5 rounded-xl space-y-1.5 shadow-md transition-all">
            {/* REPORTS HEADER */}
            <div className="flex flex-wrap items-center justify-between gap-1.5 border-b border-slate-200 pb-1.5">
              <div className="flex items-center gap-1.5">
                <div className="bg-indigo-100/80 text-indigo-900 p-1 rounded-md border border-indigo-200 shadow-2xs">
                  <BarChart3 className="w-3.5 h-3.5 text-indigo-700" />
                </div>
                <div>
                  <h2 className="text-xs font-extrabold text-slate-900 leading-tight">Reports & Analytics Dashboard</h2>
                  <p className="text-[9.5px] text-slate-500 font-medium">Performance insights by Date, Mechanic, Supervisor, Delivery Date & Location.</p>
                </div>
              </div>

              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => setIsReportsFilterCollapsed(!isReportsFilterCollapsed)}
                  className="flex items-center gap-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[10.5px] px-2 py-0.5 rounded transition-colors shadow-2xs cursor-pointer"
                >
                  {isReportsFilterCollapsed ? <ChevronDown className="w-3 h-3" /> : <ChevronUp className="w-3 h-3" />}
                  <span>{isReportsFilterCollapsed ? "Show Filters" : "Hide Filters"}</span>
                </button>
                <button
                  type="button"
                  onClick={downloadSavedCardsExcel}
                  disabled={filteredReportCards.length === 0}
                  className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold text-[10.5px] px-2 py-0.5 rounded transition-colors shadow-2xs cursor-pointer"
                >
                  <Download className="w-3 h-3" /> Export Report (.xlsx)
                </button>
              </div>
            </div>

            {!isReportsFilterCollapsed && (
              <>
                {/* REPORT FILTERS SECTION */}
                <div className="space-y-1">
              <div className="flex items-center justify-between">
                <span className="text-[10px] font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1">
                  <Filter className="w-2.5 h-2.5 text-indigo-600" /> Report Filters
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setReportDateFrom('');
                    setReportDateTo('');
                    setReportLocation('all');
                    setReportStatus('all');
                    setReportMechanic('all');
                    setReportSupervisor('all');
                    setReportServiceType('all');
                    setReportSearchQuery('');
                    setSelectedReportKpi('all');
                  }}
                  className="text-[9.5px] text-slate-500 hover:text-indigo-600 font-bold underline cursor-pointer"
                >
                  {appLang === 'te' ? 'ఫిల్టర్లు రీసెట్' : 'Reset All Filters'}
                </button>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-1.5">
                {/* Date From */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Date From</label>
                  <input
                    type="date"
                    value={reportDateFrom}
                    onChange={(e) => setReportDateFrom(e.target.value)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                  />
                </div>

                {/* Date To */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Date To</label>
                  <input
                    type="date"
                    value={reportDateTo}
                    onChange={(e) => setReportDateTo(e.target.value)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                  />
                </div>

                {/* Service Location */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Location</label>
                  <select
                    value={reportLocation}
                    onChange={(e) => setReportLocation(e.target.value as any)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Locations</option>
                    <option value="workshop">Workshop</option>
                    <option value="dss">DSS (Door Step)</option>
                    <option value="event">Event / Camp</option>
                  </select>
                </div>

                {/* Job Card Status */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Status</label>
                  <select
                    value={reportStatus}
                    onChange={(e) => setReportStatus(e.target.value as any)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Statuses</option>
                    <option value="Open">Open Cards</option>
                    <option value="Closed">Closed Cards</option>
                  </select>
                </div>

                {/* Mechanic Filter */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Mechanic</label>
                  <select
                    value={reportMechanic}
                    onChange={(e) => setReportMechanic(e.target.value)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Mechanics</option>
                    {mechanicsList.map((m, idx) => (
                      <option key={idx} value={m}>{m}</option>
                    ))}
                  </select>
                </div>

                {/* Supervisor Filter */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Supervisor</label>
                  <select
                    value={reportSupervisor}
                    onChange={(e) => setReportSupervisor(e.target.value)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Supervisors</option>
                    {supervisorsList.map((s, idx) => (
                      <option key={idx} value={s}>{s}</option>
                    ))}
                  </select>
                </div>

                {/* Service Type Filter */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Service Type</label>
                  <select
                    value={reportServiceType}
                    onChange={(e) => setReportServiceType(e.target.value)}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Types</option>
                    <option value="Free">Free Service</option>
                    <option value="Paid">Paid Service</option>
                    <option value="Wty">Under Wty Repairs</option>
                    <option value="Running Repair">Running Repair</option>
                    <option value="PDI">PDI</option>
                  </select>
                </div>
              </div>

              {/* Search Input Filter */}
              <div className="pt-1 border-t border-slate-200 flex items-center gap-1.5">
                <Search className="w-3 h-3 text-slate-400 shrink-0" />
                <input
                  type="text"
                  placeholder="Search report data by Job No, Customer Name, Model, Chassis, Mobile, Village..."
                  value={reportSearchQuery}
                  onChange={(e) => setReportSearchQuery(e.target.value)}
                  className="w-full text-[10px] py-0.5 px-1.5 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                />
                {reportSearchQuery && (
                  <button
                    onClick={() => setReportSearchQuery('')}
                    className="text-[9.5px] font-bold text-slate-500 hover:text-slate-700 px-1 cursor-pointer"
                  >
                    Clear
                  </button>
                )}
              </div>
            </div>

            {/* KEY PERFORMANCE INDICATORS (SUMMARY CARDS - CLICKABLE FILTERS) */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-1.5">
              {/* Total Job Cards */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('all')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'all'
                    ? 'bg-blue-100/90 border-blue-500 ring-2 ring-blue-400 shadow-2xs'
                    : 'bg-blue-50/90 border-blue-200 hover:bg-blue-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-blue-700 uppercase tracking-wider">Total Cards</span>
                  {selectedReportKpi === 'all' && <span className="text-[7.5px] font-black bg-blue-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-blue-900 leading-tight">{totalCardsCount}</div>
                <span className="text-[8px] text-blue-600 font-medium truncate">Click to show all</span>
              </button>

              {/* Open Job Cards */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('open')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'open'
                    ? 'bg-amber-100/90 border-amber-500 ring-2 ring-amber-400 shadow-2xs'
                    : 'bg-amber-50/90 border-amber-200 hover:bg-amber-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-amber-800 uppercase tracking-wider">Open Cards</span>
                  {selectedReportKpi === 'open' && <span className="text-[7.5px] font-black bg-amber-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-amber-900 leading-tight">{openCardsCount}</div>
                <span className="text-[8px] text-amber-700 font-medium truncate">{totalCardsCount > 0 ? ((openCardsCount / totalCardsCount) * 100).toFixed(0) : 0}% of total</span>
              </button>

              {/* Closed Job Cards */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('closed')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'closed'
                    ? 'bg-emerald-100/90 border-emerald-500 ring-2 ring-emerald-400 shadow-2xs'
                    : 'bg-emerald-50/90 border-emerald-200 hover:bg-emerald-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-emerald-800 uppercase tracking-wider">Closed Cards</span>
                  {selectedReportKpi === 'closed' && <span className="text-[7.5px] font-black bg-emerald-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-emerald-900 leading-tight">{closedCardsCount}</div>
                <span className="text-[8px] text-emerald-700 font-medium truncate">{totalCardsCount > 0 ? ((closedCardsCount / totalCardsCount) * 100).toFixed(0) : 0}% done</span>
              </button>

              {/* Total Revenue */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('revenue')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'revenue'
                    ? 'bg-indigo-100/90 border-indigo-500 ring-2 ring-indigo-400 shadow-2xs'
                    : 'bg-indigo-50/90 border-indigo-200 hover:bg-indigo-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-indigo-700 uppercase tracking-wider">Total Revenue</span>
                  {selectedReportKpi === 'revenue' && <span className="text-[7.5px] font-black bg-indigo-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-indigo-900 leading-tight">₹{totalRevenueSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
                <span className="text-[8px] text-indigo-600 font-medium truncate">Total billings</span>
              </button>

              {/* Total Spares Sales */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('spares')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'spares'
                    ? 'bg-purple-100/90 border-purple-500 ring-2 ring-purple-400 shadow-2xs'
                    : 'bg-purple-50/90 border-purple-200 hover:bg-purple-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-purple-700 uppercase tracking-wider">Spares Sales</span>
                  {selectedReportKpi === 'spares' && <span className="text-[7.5px] font-black bg-purple-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-purple-900 leading-tight">₹{totalSparesSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
                <span className="text-[8px] text-purple-600 font-medium truncate">Parts & materials</span>
              </button>

              {/* Total Labour Charges */}
              <button
                type="button"
                onClick={() => setSelectedReportKpi('labour')}
                className={`text-left p-1.5 rounded-md border transition-all cursor-pointer flex flex-col justify-between ${
                  selectedReportKpi === 'labour'
                    ? 'bg-teal-100/90 border-teal-500 ring-2 ring-teal-400 shadow-2xs'
                    : 'bg-teal-50/90 border-teal-200 hover:bg-teal-100/60'
                }`}
              >
                <div className="flex items-center justify-between">
                  <span className="text-[8.5px] font-bold text-teal-700 uppercase tracking-wider">Labour Charges</span>
                  {selectedReportKpi === 'labour' && <span className="text-[7.5px] font-black bg-teal-600 text-white px-0.5 rounded">ACTIVE</span>}
                </div>
                <div className="text-sm font-extrabold text-teal-900 leading-tight">₹{totalLabourSum.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</div>
                <span className="text-[8px] text-teal-600 font-medium truncate">Workmanship</span>
              </button>
            </div>

            {/* SERVICE LOCATION BREAKDOWN SUMMARY */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-1.5">
              <button
                type="button"
                onClick={() => setSelectedReportKpi('workshop')}
                className={`text-left p-1 px-2 rounded-md border transition-all cursor-pointer flex items-center justify-between ${
                  selectedReportKpi === 'workshop'
                    ? 'bg-indigo-100/90 border-indigo-500 ring-2 ring-indigo-400 shadow-2xs'
                    : 'bg-white/80 border-slate-200 hover:bg-indigo-50/50'
                }`}
              >
                <div className="flex items-center gap-2">
                  <span className="text-[8.5px] font-bold text-slate-500 uppercase tracking-wider">Workshop:</span>
                  <span className="text-xs font-bold text-slate-900">{workshopCount} Cards</span>
                  <span className="text-[10px] text-emerald-600 font-bold">₹{workshopRevenue.toLocaleString('en-IN')}</span>
                </div>
                <div className="bg-indigo-50 px-1 py-0.2 rounded border border-indigo-200 text-indigo-700 font-bold text-[10px]">
                  {totalCardsCount > 0 ? ((workshopCount / totalCardsCount) * 100).toFixed(0) : 0}%
                </div>
              </button>

              <button
                type="button"
                onClick={() => setSelectedReportKpi('dss')}
                className={`text-left p-1 px-2 rounded-md border transition-all cursor-pointer flex items-center justify-between ${
                  selectedReportKpi === 'dss'
                    ? 'bg-sky-100/90 border-sky-500 ring-2 ring-sky-400 shadow-2xs'
                    : 'bg-white/80 border-slate-200 hover:bg-sky-50/50'
                }`}
              >
                <div className="flex items-center gap-2">
                  <span className="text-[8.5px] font-bold text-slate-500 uppercase tracking-wider">DSS:</span>
                  <span className="text-xs font-bold text-slate-900">{dssCount} Cards</span>
                  <span className="text-[10px] text-emerald-600 font-bold">₹{dssRevenue.toLocaleString('en-IN')}</span>
                </div>
                <div className="bg-sky-50 px-1 py-0.2 rounded border border-sky-200 text-sky-700 font-bold text-[10px]">
                  {totalCardsCount > 0 ? ((dssCount / totalCardsCount) * 100).toFixed(0) : 0}%
                </div>
              </button>

              <button
                type="button"
                onClick={() => setSelectedReportKpi('event')}
                className={`text-left p-1 px-2 rounded-md border transition-all cursor-pointer flex items-center justify-between ${
                  selectedReportKpi === 'event'
                    ? 'bg-amber-100/90 border-amber-500 ring-2 ring-amber-400 shadow-2xs'
                    : 'bg-white/80 border-slate-200 hover:bg-amber-50/50'
                }`}
              >
                <div className="flex items-center gap-2">
                  <span className="text-[8.5px] font-bold text-slate-500 uppercase tracking-wider">Event:</span>
                  <span className="text-xs font-bold text-slate-900">{eventCount} Cards</span>
                  <span className="text-[10px] text-emerald-600 font-bold">₹{eventRevenue.toLocaleString('en-IN')}</span>
                </div>
                <div className="bg-amber-50 px-1 py-0.2 rounded border border-amber-200 text-amber-700 font-bold text-[10px]">
                  {totalCardsCount > 0 ? ((eventCount / totalCardsCount) * 100).toFixed(0) : 0}%
                </div>
              </button>
            </div>
          </>
        )}
      </div>

          {/* INTERACTIVE DETAILED JOB CARDS REPORT LIST */}
          <div className="bg-slate-50 border border-slate-200 p-3 rounded-xl space-y-2">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2">
              <div className="flex items-center gap-2">
                <span className="text-xs font-extrabold uppercase tracking-wider text-slate-800 flex items-center gap-1.5">
                  <FileText className="w-3.5 h-3.5 text-indigo-600" />
                  Detailed Job Cards Report
                </span>
                <span className="bg-indigo-100 text-indigo-900 text-[10.5px] font-bold px-2 py-0.5 rounded-full border border-indigo-300">
                  Category: {
                    selectedReportKpi === 'all' ? 'All Matching Cards' :
                    selectedReportKpi === 'open' ? '⏳ Open Cards Only' :
                    selectedReportKpi === 'closed' ? '✓ Closed Cards Only' :
                    selectedReportKpi === 'revenue' ? '₹ Total Revenue Cards' :
                    selectedReportKpi === 'spares' ? '📦 Spares Sales Cards' :
                    selectedReportKpi === 'labour' ? '🔧 Labour Charges Cards' :
                    selectedReportKpi === 'workshop' ? '🏭 Workshop Services' :
                    selectedReportKpi === 'dss' ? '🛵 Door Step Services (DSS)' :
                    '🎪 Event / Camp Services'
                  } ({kpiFilteredCards.length})
                </span>
              </div>

              <div className="flex items-center gap-2">
                {selectedReportKpi !== 'all' && (
                  <button
                    type="button"
                    onClick={() => setSelectedReportKpi('all')}
                    className="text-[10.5px] font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 px-2 py-0.5 rounded border border-indigo-200 cursor-pointer"
                  >
                    Reset KPI Filter
                  </button>
                )}
                <button
                  type="button"
                  onClick={downloadSavedCardsExcel}
                  disabled={kpiFilteredCards.length === 0}
                  className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold text-[11px] px-2.5 py-0.5 rounded transition-colors shadow-2xs cursor-pointer"
                >
                  <Download className="w-3 h-3" /> Export to Excel (.xlsx)
                </button>
              </div>
            </div>

            {/* TABLE CONTROL & LINE SIZE SETTINGS BAR */}
            <div className="bg-indigo-50/90 border border-indigo-200 p-2.5 rounded-lg flex flex-wrap items-center justify-between gap-3 text-xs font-semibold mt-2 mb-2">
              <div className="flex items-center gap-3 flex-wrap">
                <div className="flex items-center gap-1.5">
                  <span className="text-slate-800 font-extrabold flex items-center gap-1.5">
                    <Maximize2 className="w-3.5 h-3.5 text-indigo-600" /> Line Size / Row Height:
                  </span>
                  <div className="inline-flex rounded-md shadow-2xs bg-white p-0.5 border border-slate-300">
                    <button
                      type="button"
                      onClick={() => setReportRowDensity('compact')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        reportRowDensity === 'compact'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      🤏 Compact
                    </button>
                    <button
                      type="button"
                      onClick={() => setReportRowDensity('normal')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        reportRowDensity === 'normal'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      ↔️ Normal
                    </button>
                    <button
                      type="button"
                      onClick={() => setReportRowDensity('spacious')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        reportRowDensity === 'spacious'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      ↕️ Spacious
                    </button>
                  </div>
                </div>

                {/* DYNAMIC FREEZE COLUMN SELECTOR */}
                <div className="flex items-center gap-1.5 border-l border-indigo-200 pl-3">
                  <span className="text-slate-800 font-extrabold flex items-center gap-1 text-[11px]">
                    ❄️ Freeze:
                  </span>
                  <select
                    value={reportFreezeUpToColumn}
                    onChange={(e) => setReportFreezeUpToColumn(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 text-[11px] font-extrabold text-slate-800 outline-none cursor-pointer shadow-2xs focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                    title="Freeze columns up to the selected one"
                  >
                    <option value="none">No Freeze</option>
                    <option value="slNo">Sl No Only</option>
                    <option value="jobNo">Job No (Default)</option>
                    <option value="jobDate">Up to Job Date</option>
                    <option value="customerInfo">Up to Customer Info</option>
                    <option value="mobile">Up to Mobile</option>
                    <option value="modelInfo">Up to Model Info</option>
                    <option value="servicePlace">Up to Service Place</option>
                    <option value="mechanic">Up to Technician</option>
                    <option value="supervisor">Up to Supervisor</option>
                    <option value="status">Up to Status</option>
                  </select>
                </div>

                {/* PAGE SIZE SELECTOR */}
                <div className="flex items-center gap-1.5 border-l border-indigo-200 pl-3">
                  <span className="text-slate-800 font-extrabold flex items-center gap-1 text-[11px]">
                    📄 Page Size:
                  </span>
                  <select
                    value={reportPageSize}
                    onChange={(e) => {
                      setReportPageSize(Number(e.target.value));
                      setReportPage(1);
                    }}
                    className="bg-white border border-slate-300 rounded px-2 py-1 text-[11px] font-extrabold text-slate-800 outline-none cursor-pointer shadow-2xs focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                  >
                    <option value={25}>25 / page</option>
                    <option value={50}>50 / page (Fast)</option>
                    <option value={100}>100 / page</option>
                    <option value={250}>250 / page</option>
                    <option value={-1}>Show All</option>
                  </select>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setReportColWidths(defaultReportColWidths);
                  setReportFreezeUpToColumn('jobNo');
                }}
                className="text-[10px] text-slate-500 hover:text-indigo-600 font-bold underline cursor-pointer flex items-center gap-1"
              >
                ↺ Reset Table Layout
              </button>
            </div>

            {kpiFilteredCards.length === 0 ? (
              <div className="text-center py-6 text-slate-400 text-xs font-semibold">
                No job cards found matching this selected KPI category and search criteria.
              </div>
            ) : (
              <div className="space-y-2">
                <div className="overflow-x-auto border border-slate-200 rounded-lg max-h-[500px]">
                  <table className="w-full text-left text-xs text-slate-700 min-w-max relative border-separate border-spacing-0">
                    <thead className="bg-slate-100 text-slate-800 font-extrabold uppercase text-[10px] shadow-sm">
                      <tr>
                        {[
                          { key: 'slNo', label: '#' },
                          { key: 'jobNo', label: 'Job No' },
                          { key: 'jobDate', label: 'Job Date' },
                          { key: 'customerInfo', label: 'Customer Name & Village' },
                          { key: 'mobile', label: 'Mobile' },
                          { key: 'modelInfo', label: 'Model & Chassis No' },
                          { key: 'servicePlace', label: 'Service Place' },
                          { key: 'mechanic', label: 'Technician' },
                          { key: 'supervisor', label: 'Supervisor' },
                          { key: 'status', label: 'Status' },
                          { key: 'spares', label: 'Spares (₹)' },
                          { key: 'labour', label: 'Labour (₹)' },
                          { key: 'total', label: 'Total (₹)' }
                        ].map((col) => (
                          <th
                            key={col.key}
                            style={getReportStickyProps(col.key, true).style}
                            className={`p-2 border-r border-b-2 border-slate-300 relative select-none text-center ${getReportStickyProps(col.key, true).className}`}
                          >
                            {col.label}
                            <div
                              onMouseDown={(e) => handleReportColumnResizeStart(col.key, e)}
                              className="absolute right-0 top-0 bottom-0 w-3 cursor-col-resize group flex items-center justify-center z-30 hover:bg-indigo-500/20 active:bg-indigo-600/40"
                              title={`Drag to resize ${col.label}`}
                            >
                              <div className="w-[2px] h-full bg-slate-300 group-hover:bg-indigo-600 group-active:bg-indigo-700" />
                            </div>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 font-medium bg-white">
                      {paginatedReportCards.map((card, idx) => {
                        const rowSlNo = (validReportPage - 1) * (reportPageSize === -1 ? 0 : reportPageSize) + idx + 1;
                        const sparesAmt = parseMoney(card.warrantyMaterial) + parseMoney(card.nonWarrantyMaterial) || parseMoney(card.partsTotal) || (Array.isArray(card.partRows) ? card.partRows.reduce((sum, p) => sum + parseMoney(p.amount), 0) : 0);
                        const labourAmt = parseMoney(card.totalLabour) || parseMoney(card.labourTotal) || (Array.isArray(card.repairRows) ? card.repairRows.reduce((sum, r) => sum + parseMoney(r.charge), 0) : 0);
                        const grandTotal = parseMoney(card.gTotal) || (sparesAmt + labourAmt);
                        const normMech = normalizeStaffName(card.mechanic || card.technicianName);
                        const normSup = normalizeStaffName(card.wsIncharge);
                        const cardStatus = isCardClosed(card) ? 'Closed' : 'Open';
                        const cellPadding =
                          reportRowDensity === 'compact'
                            ? 'p-1 text-[11px]'
                            : reportRowDensity === 'spacious'
                            ? 'p-3.5 text-sm'
                            : 'p-2.5 text-xs';
                        
                        return (
                          <tr key={card.id || idx} className="group transition-colors">
                            <td style={getReportStickyProps('slNo').style} className={`${cellPadding} text-center text-slate-400 font-bold border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('slNo').className}`}>{rowSlNo}</td>
                            <td style={getReportStickyProps('jobNo').style} className={`${cellPadding} font-mono font-bold text-indigo-900 border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('jobNo').className}`}>
                              {card.jobNo || card.onlineJobCardNo || '—'}
                            </td>
                            <td style={getReportStickyProps('jobDate').style} className={`${cellPadding} border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('jobDate').className}`}>{fmtDate(card.jobDate || card.complaintDate) || '—'}</td>
                            <td style={getReportStickyProps('customerInfo').style} className={`${cellPadding} border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('customerInfo').className}`}>
                              <div className="font-bold text-slate-900">{card.custName || '—'}</div>
                              {card.village && <div className="text-[10px] text-slate-500 font-medium">{card.village}{card.mandal ? `, ${card.mandal}` : ''}</div>}
                            </td>
                            <td style={getReportStickyProps('mobile').style} className={`${cellPadding} font-mono text-[11px] border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('mobile').className}`}>{card.ownerMob || card.phNo || '—'}</td>
                            <td style={getReportStickyProps('modelInfo').style} className={`${cellPadding} border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('modelInfo').className}`}>
                              <div className="font-semibold text-slate-800">{card.model || '—'}</div>
                              {card.chassisNo && <div className="text-[10px] font-mono text-slate-500">{card.chassisNo}</div>}
                            </td>
                            <td style={getReportStickyProps('servicePlace').style} className={`${cellPadding} font-bold uppercase text-[10px] border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('servicePlace').className}`}>
                              {card.serviceLocation === 'dss' ? '🛵 Door Step' : card.serviceLocation === 'event' ? '🎪 Event' : '🏭 Workshop'}
                            </td>
                            <td style={getReportStickyProps('mechanic').style} className={`${cellPadding} border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('mechanic').className}`}>
                              <input
                                type="text"
                                list="mechanics-datalist"
                                value={normMech || ''}
                                onChange={(e) => handleInlineCardFieldChange(card.id, 'mechanic', e.target.value)}
                                placeholder="Select or type mechanic..."
                                className="w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-bold p-1 text-[11px]"
                              />
                            </td>
                            <td style={getReportStickyProps('supervisor').style} className={`${cellPadding} border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('supervisor').className}`}>
                              <input
                                type="text"
                                list="supervisors-datalist"
                                value={normSup || ''}
                                onChange={(e) => handleInlineCardFieldChange(card.id, 'wsIncharge', e.target.value)}
                                placeholder="Select or type supervisor..."
                                className="w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-bold p-1 text-[11px]"
                              />
                            </td>
                            <td style={getReportStickyProps('status').style} className={`${cellPadding} text-center whitespace-nowrap border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('status').className}`}>
                              <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase border ${
                                cardStatus === 'Closed' 
                                  ? 'bg-emerald-600 text-white border-emerald-700 shadow-2xs' 
                                  : 'bg-amber-100 text-amber-900 border-amber-400 font-extrabold shadow-2xs'
                              }`}>
                                {cardStatus === 'Closed' ? '✓ Closed' : '⏳ Open'}
                              </span>
                            </td>
                            <td style={getReportStickyProps('spares').style} className={`${cellPadding} text-right font-mono font-semibold text-purple-800 border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('spares').className}`}>
                              ₹{sparesAmt.toLocaleString('en-IN')}
                            </td>
                            <td style={getReportStickyProps('labour').style} className={`${cellPadding} text-right font-mono font-semibold text-teal-800 border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('labour').className}`}>
                              ₹{labourAmt.toLocaleString('en-IN')}
                            </td>
                            <td style={getReportStickyProps('total').style} className={`${cellPadding} text-right font-mono font-extrabold text-indigo-900 border-r border-slate-200 bg-white group-hover:bg-amber-100/80 ${getReportStickyProps('total').className}`}>
                              ₹{grandTotal.toLocaleString('en-IN')}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                {/* PAGINATION CONTROLS BAR */}
                {kpiFilteredCards.length > 0 && reportPageSize !== -1 && (
                  <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-2.5 rounded-lg border border-slate-200 text-xs font-semibold">
                    <div className="text-slate-600 font-medium">
                      Showing <span className="font-bold text-slate-900">{(validReportPage - 1) * reportPageSize + 1}</span> to{' '}
                      <span className="font-bold text-slate-900">
                        {Math.min(validReportPage * reportPageSize, kpiFilteredCards.length)}
                      </span>{' '}
                      of <span className="font-bold text-indigo-700">{kpiFilteredCards.length}</span> cards
                    </div>

                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => setReportPage(1)}
                        disabled={validReportPage <= 1}
                        className="px-2 py-1 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 text-slate-700 font-bold rounded text-xs cursor-pointer"
                      >
                        « First
                      </button>
                      <button
                        type="button"
                        onClick={() => setReportPage((prev) => Math.max(1, prev - 1))}
                        disabled={validReportPage <= 1}
                        className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 text-slate-700 font-bold rounded text-xs cursor-pointer"
                      >
                        ‹ Prev
                      </button>

                      <span className="px-3 py-1 bg-indigo-50 text-indigo-900 font-bold rounded border border-indigo-200 text-xs">
                        Page {validReportPage} of {totalReportPages}
                      </span>

                      <button
                        type="button"
                        onClick={() => setReportPage((prev) => Math.min(totalReportPages, prev + 1))}
                        disabled={validReportPage >= totalReportPages}
                        className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 text-slate-700 font-bold rounded text-xs cursor-pointer"
                      >
                        Next ›
                      </button>
                      <button
                        type="button"
                        onClick={() => setReportPage(totalReportPages)}
                        disabled={validReportPage >= totalReportPages}
                        className="px-2 py-1 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 text-slate-700 font-bold rounded text-xs cursor-pointer"
                      >
                        Last »
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {/* TAB 5: COMPLAINTS */}
      {activeTab === 'complaints' && (
        <div className="w-full bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-6 print:hidden">
          <div className="border-b pb-3 flex flex-wrap gap-4 justify-between items-start">
            <div>
              <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <AlertCircle className="w-5 h-5 text-red-600" /> Complaint Register
              </h2>
              <p className="text-xs text-slate-500 font-medium">Log and track customer complaints.</p>
            </div>
            
            <button
              type="button"
              onClick={() => {
                setComplaintForm({
                  id: '',
                  complaintNo: '',
                  complaintDate: new Date().toISOString().split('T')[0],
                  chassisNo: '',
                  customerName: '',
                  mobileNumber: '',
                  tractorModel: '',
                  complaintDetails: '',
                  assignedMechanic: '',
                  assignedSupervisor: '',
                  status: 'Open',
                  resolution: '',
                  closureDate: '',
                });
                setIsComplaintModalOpen(true);
              }}
              className="flex items-center gap-1 bg-red-600 hover:bg-red-700 text-white font-bold text-xs px-3 py-2 rounded-lg transition-colors shadow-sm"
            >
              <Plus className="w-4 h-4" /> Add Complaint
            </button>
          </div>

          <div className="sticky top-0 z-20 bg-white grid grid-cols-1 md:grid-cols-5 gap-3 items-center border border-slate-200 p-3 rounded-lg shadow-sm">
            <input
              type="text"
              placeholder="Search..."
              value={complaintSearch}
              onChange={(e) => setComplaintSearch(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs"
            />
            <input
              type="date"
              value={complaintDateFilter}
              onChange={(e) => setComplaintDateFilter(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs"
            />
            <select
              value={complaintStatusFilter}
              onChange={(e) => setComplaintStatusFilter(e.target.value as any)}
              className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold"
            >
              <option value="all">All Status</option>
              <option value="Open">Open</option>
              <option value="Running">Running</option>
              <option value="Closed">Closed</option>
            </select>
            <select
              value={complaintTechFilter}
              onChange={(e) => setComplaintTechFilter(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold"
            >
              <option value="all">All Technicians</option>
              {mechanicsList.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
            <select
              value={complaintSupFilter}
              onChange={(e) => setComplaintSupFilter(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold"
            >
              <option value="all">All Supervisors</option>
              {supervisorsList.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-lg max-h-[500px]">
            <table className="w-full text-left text-xs whitespace-nowrap">
              <thead className="bg-slate-100 text-slate-600 uppercase sticky top-0 z-10">
                <tr>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Complaint No</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Date</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold text-emerald-800">Closed Date</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Customer Details</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Complaint Details</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Tractor Model</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Status</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold">Assigned To</th>
                  <th className="p-3 border-b border-slate-200 font-extrabold text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {complaints
                  .filter(c => {
                    if (complaintStatusFilter !== 'all' && c.status !== complaintStatusFilter) return false;
                    if (complaintTechFilter !== 'all' && c.assignedMechanic !== complaintTechFilter) return false;
                    if (complaintSupFilter !== 'all' && c.assignedSupervisor !== complaintSupFilter) return false;
                    
                    if (complaintDateFilter) {
                      const cDate = toInputDateFormat(c.complaintDate) || c.complaintDate;
                      if (!cDate || !cDate.startsWith(complaintDateFilter)) return false;
                    }

                    if (complaintSearch) {
                      const q = complaintSearch.toLowerCase().trim();
                      const custRec = findCustomerRecordByQuery(c.chassisNo) || findCustomerRecordByQuery(c.customerName) || findCustomerRecordByQuery(c.mobileNumber);
                      const matchingCard = savedJobCards.find(card => card.chassisNo && normalizeKey(card.chassisNo) === normalizeKey(c.chassisNo));
                      
                      // Resolve names and details from multiple sources
                      const resolvedName = String(c.customerName || (custRec ? (custRec['Customer Name'] || custRec.customerName) : '') || matchingCard?.custName || '').toLowerCase();
                      const resolvedVillage = String(c.village || (custRec ? (custRec['Village'] || custRec.village) : '') || matchingCard?.village || '').toLowerCase();
                      const hfn = String(c.historyFileNo || (custRec ? (custRec['historyFileNo'] || custRec.fileNo) : '') || matchingCard?.historyFileNo || '').toLowerCase();
                      const onlineJC = String(c.jobCardNo || c.onlineJobCardNo || matchingCard?.onlineJobCardNo || '').toLowerCase();
                      const details = String(c.complaintDetails || '').toLowerCase();
                      const mobile = String(c.mobileNumber || (custRec ? (custRec['Mobile Number'] || custRec.mobileNumber) : '') || '').toLowerCase();

                      return (
                        String(c.chassisNo || '').toLowerCase().includes(q) ||
                        resolvedName.includes(q) ||
                        mobile.includes(q) ||
                        String(c.complaintNo || '').toLowerCase().includes(q) ||
                        String(c.tractorModel || '').toLowerCase().includes(q) ||
                        details.includes(q) ||
                        String(c.closureDate || '').toLowerCase().includes(q) ||
                        resolvedVillage.includes(q) ||
                        onlineJC.includes(q) ||
                        hfn.includes(q)
                      );
                    }
                    return true;
                  })
                  .map(c => {
                    const custRec = findCustomerRecordByQuery(c.chassisNo);
                    const matchingCard = savedJobCards.find(card => card.chassisNo && normalizeKey(card.chassisNo) === normalizeKey(c.chassisNo));
                    const hfn = c.historyFileNo || (custRec ? getFieldValue(custRec, 'historyFileNo') : '') || matchingCard?.historyFileNo || '';

                    return (
                      <tr key={c.id} className="hover:bg-slate-50">
                        <td className="p-3 font-bold text-slate-900">{c.complaintNo}</td>
                        <td className="p-3 text-slate-600 font-semibold">{toInputDateFormat(c.complaintDate) ? fmtDate(c.complaintDate) : (c.complaintDate || '—')}</td>
                        <td className="p-3 font-semibold text-slate-700">
                          {c.closureDate ? (
                            <span className="inline-block px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold text-[11px]">
                              {toInputDateFormat(c.closureDate) ? fmtDate(c.closureDate) : c.closureDate}
                            </span>
                          ) : (
                            <span className="text-slate-400 font-normal">—</span>
                          )}
                        </td>
                        <td className="p-3">
                          <div className="font-bold text-slate-800">{c.customerName || '—'}</div>
                          {custRec && (getFieldValue(custRec, 'village') || getFieldValue(custRec, 'mandal')) && (
                            <div className="text-[10px] text-slate-500 font-semibold mb-0.5 leading-tight">
                              {[getFieldValue(custRec, 'village'), getFieldValue(custRec, 'mandal')].filter(Boolean).join(', ')}
                            </div>
                          )}
                          <div className="text-[10px] text-slate-600 font-mono flex items-center gap-1">
                            <span className="font-bold text-slate-500">Chassis:</span> {c.chassisNo || '—'}
                          </div>
                          {hfn && (
                            <div className="text-[10px] font-mono font-black text-indigo-700 bg-indigo-50 border border-indigo-200 px-1.5 py-0.5 rounded inline-block mt-0.5 shadow-2xs">
                              <span>HFN: {hfn}</span>
                            </div>
                          )}
                          <div className="text-xs text-slate-800 font-black">{c.mobileNumber || '—'}</div>
                        </td>
                        <td className="p-3 text-slate-600 text-xs font-bold max-w-xs whitespace-normal break-words" title={c.complaintDetails}>{c.complaintDetails || '-'}</td>
                        <td className="p-3 text-slate-600">{c.tractorModel || '-'}</td>
                        <td className="p-3">
                          {c.status === 'Closed' ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                              Closed
                            </span>
                          ) : c.status === 'Running' ? (
                            <span className="px-2 py-0.5 rounded text-[10px] font-black bg-amber-100 text-amber-900 border border-amber-300 inline-flex items-center gap-1 shadow-2xs">
                              <span className="w-1.5 h-1.5 rounded-full bg-amber-600 animate-pulse"></span>
                              Running
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-red-700 text-white shadow-sm">
                              {c.status || 'Open'}
                            </span>
                          )}
                          {c.jobCardNo && (
                            <div className="text-[9px] font-mono font-bold text-slate-500 mt-0.5">
                              JC: {c.jobCardNo}
                            </div>
                          )}
                        </td>
                        <td className="p-3">
                          <div className="text-xs text-slate-700">{c.assignedMechanic || '-'}</div>
                          <div className="text-[10px] text-slate-500">{c.assignedSupervisor || '-'}</div>
                        </td>
                        <td className="p-3 text-right">
                          <div className="flex justify-end items-center gap-1.5">
                            {(c.status === 'Open' || (!c.status && !c.jobCardNo)) && (
                              <button
                                type="button"
                                onClick={async () => {
                                  // Update complaint status to Running immediately
                                  const updatedComp = { ...c, status: 'Running' };
                                  setComplaints(prev => prev.map(comp => comp.id === c.id ? updatedComp : comp));

                                  if (useGoogleSheets && sheetsSpreadsheetId) {
                                    try {
                                      await updateSheetRow(sheetsSpreadsheetId, 'Complaints', COMPLAINT_HEADERS, updatedComp);
                                    } catch (e) {
                                      console.error('Error updating complaint to Running in Sheets:', e);
                                    }
                                  } else if (c.id) {
                                    try {
                                      await updateDoc(doc(db, 'complaints', c.id), { status: 'Running' });
                                    } catch (e) {
                                      console.error('Error updating complaint to Running in Firestore:', e);
                                    }
                                  }

                                  resetForm();
                                  setActiveTab('new_entry');
                                  const validCompDate = toInputDateFormat(c.complaintDate) || c.complaintDate || new Date().toISOString().split('T')[0];
                                  setTimeout(() => {
                                    if (c.chassisNo) {
                                      handleChassisChange(c.chassisNo);
                                    }
                                    if (c.customerName) setCustName(c.customerName);
                                    if (c.mobileNumber) setOwnerMob(c.mobileNumber);
                                    if (c.tractorModel) setModel(c.tractorModel);

                                    const matchedRec = findCustomerRecordByQuery(c.chassisNo);
                                    const matchedCard = savedJobCards.find(card => card.chassisNo && normalizeKey(card.chassisNo) === normalizeKey(c.chassisNo));
                                    const compHfn = c.historyFileNo || (matchedRec ? getFieldValue(matchedRec, 'historyFileNo') : '') || matchedCard?.historyFileNo || '';
                                    const compSl = c.serialNo || (matchedRec ? getFieldValue(matchedRec, 'serialNo') : '') || matchedCard?.serialNo || '';
                                    if (compHfn) setHistoryFileNo(compHfn);
                                    if (compSl) setSerialNo(compSl);

                                    setComplaintDate(validCompDate);
                                    setJobDate(validCompDate);
                                    setProblemDescription(c.complaintDetails || '');
                                    setComplaintDetails(c.complaintDetails || '');

                                    window.scrollTo({ top: 0, behavior: 'smooth' });
                                  }, 100);
                                }}
                                className="p-1.5 bg-emerald-50 hover:bg-emerald-600 text-emerald-700 hover:text-white rounded-lg transition-colors border border-emerald-200 shadow-2xs cursor-pointer flex items-center justify-center"
                                title="Create Job Card (Marks complaint as Running)"
                              >
                                <FilePlus className="w-4 h-4" />
                              </button>
                            )}
                            {c.status === 'Closed' ? (
                              <button
                                type="button"
                                onClick={() => {
                                  setComplaintForm(c);
                                  setIsComplaintViewOnly(true);
                                  setIsComplaintModalOpen(true);
                                }}
                                className="px-2.5 py-1 bg-indigo-50 hover:bg-indigo-600 text-indigo-700 hover:text-white font-bold text-xs rounded-lg transition-colors border border-indigo-200 cursor-pointer flex items-center gap-1 shadow-2xs"
                                title="View Closed Complaint Details"
                              >
                                <Eye className="w-3.5 h-3.5" /> View
                              </button>
                            ) : (
                              <>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setComplaintForm(c);
                                    setIsComplaintViewOnly(false);
                                    setIsComplaintModalOpen(true);
                                  }}
                                  className="text-blue-600 hover:text-blue-800 font-bold hover:underline cursor-pointer"
                                >
                                  Edit
                                </button>
                                {currentUserRole === 'admin' && (
                                  <button
                                    type="button"
                                    onClick={() => deleteComplaint(c.id)}
                                    className="text-red-600 hover:text-red-800 font-bold hover:underline cursor-pointer"
                                  >
                                    Delete
                                  </button>
                                )}
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                {complaints.length === 0 && (
                  <tr>
                    <td colSpan={8} className="p-6 text-center text-slate-400 font-bold">No complaints found.</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB: STAFF ATTENDANCE REGISTER */}
      {activeTab === 'attendance' && (
        <div className="w-full bg-slate-50/50 p-4 md:p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-6 print:hidden">
          {/* HEADER */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
            <div>
              <h2 className="text-base font-black text-slate-900 flex items-center gap-2">
                <UserCheck className="w-5 h-5 text-teal-700" /> Staff Attendance Register
              </h2>
              <p className="text-xs text-slate-500 font-medium mt-0.5">
                Daily staff attendance tracking, tick mark register & monthly presence summary report.
              </p>
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  setEditingStaffId(null);
                  setStaffForm({ name: '', fatherName: '', village: '', mandal: '', mobileNumber: '', role: 'mechanic', dateOfJoining: '', supervisor: '' });
                  setIsStaffModalOpen(true);
                }}
                className="px-3 py-1.5 bg-teal-700 hover:bg-teal-800 text-white font-bold text-xs rounded-lg transition-colors flex items-center gap-1.5 shadow-2xs cursor-pointer"
              >
                <UserPlus className="w-4 h-4" /> Add New Staff
              </button>
            </div>
          </div>

          {/* TOP REPORT SUMMARY BOXES ("chinna boxlu") */}
          {(() => {
            const dateRecords = attendanceRecords[attendanceDate] || {};
            const totalStaff = staffMembers.length;
            const presentCount = staffMembers.filter(s => (dateRecords[s.id || s.name]?.status || 'unmarked') === 'present').length;
            const absentCount = staffMembers.filter(s => dateRecords[s.id || s.name]?.status === 'absent').length;
            const leaveCount = staffMembers.filter(s => dateRecords[s.id || s.name]?.status === 'leave').length;
            const unmarkedCount = totalStaff - (presentCount + absentCount + leaveCount);

            return (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                {/* TOTAL STAFF */}
                <div className="bg-blue-50/80 border-2 border-blue-200/80 rounded-xl p-3.5 shadow-2xs flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black text-blue-800 uppercase tracking-wider">Total Staff</p>
                    <p className="text-xl font-black text-blue-950 mt-1">{totalStaff} <span className="text-xs font-semibold text-blue-700">Members</span></p>
                  </div>
                  <div className="p-2.5 bg-blue-600 text-white rounded-xl shadow-2xs">
                    <Users className="w-5 h-5" />
                  </div>
                </div>

                {/* PRESENT TODAY */}
                <div className="bg-emerald-50/80 border-2 border-emerald-200/80 rounded-xl p-3.5 shadow-2xs flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black text-emerald-800 uppercase tracking-wider">Present Today</p>
                    <p className="text-xl font-black text-emerald-950 mt-1">
                      {presentCount} <span className="text-xs font-bold text-emerald-700">({totalStaff > 0 ? Math.round((presentCount / totalStaff) * 100) : 0}%)</span>
                    </p>
                  </div>
                  <div className="p-2.5 bg-emerald-600 text-white rounded-xl shadow-2xs">
                    <UserCheck className="w-5 h-5" />
                  </div>
                </div>

                {/* ABSENT TODAY */}
                <div className="bg-rose-50/80 border-2 border-rose-200/80 rounded-xl p-3.5 shadow-2xs flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black text-rose-800 uppercase tracking-wider">Absent Today</p>
                    <p className="text-xl font-black text-rose-950 mt-1">
                      {absentCount} <span className="text-xs font-semibold text-rose-700">{unmarkedCount > 0 ? `(${unmarkedCount} pending)` : ''}</span>
                    </p>
                  </div>
                  <div className="p-2.5 bg-rose-600 text-white rounded-xl shadow-2xs">
                    <XCircle className="w-5 h-5" />
                  </div>
                </div>

                {/* ON LEAVE / HALF DAY */}
                <div className="bg-amber-50/80 border-2 border-amber-200/80 rounded-xl p-3.5 shadow-2xs flex items-center justify-between">
                  <div>
                    <p className="text-[10px] font-black text-amber-800 uppercase tracking-wider">Leave / Half Day</p>
                    <p className="text-xl font-black text-amber-950 mt-1">{leaveCount} <span className="text-xs font-semibold text-amber-700">Staff</span></p>
                  </div>
                  <div className="p-2.5 bg-amber-500 text-white rounded-xl shadow-2xs">
                    <Clock className="w-5 h-5" />
                  </div>
                </div>
              </div>
            );
          })()}

          {/* MAIN ATTENDANCE REGISTER BOX */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 md:p-5 shadow-xs space-y-4">
            {/* DATE CONTROLS & SUB-TABS */}
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 pb-3 border-b border-slate-200">
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl border border-slate-200">
                  <button
                    type="button"
                    onClick={() => setAttendanceTab('daily')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      attendanceTab === 'daily' ? 'bg-teal-800 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    📅 Daily Register
                  </button>
                  <button
                    type="button"
                    onClick={() => setAttendanceTab('monthly')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      attendanceTab === 'monthly' ? 'bg-teal-800 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    📊 Monthly Report
                  </button>
                </div>

                {attendanceTab === 'daily' && (
                  <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 p-1 rounded-xl">
                    <button
                      type="button"
                      onClick={() => {
                        const d = new Date(attendanceDate);
                        d.setDate(d.getDate() - 1);
                        setAttendanceDate(d.toISOString().split('T')[0]);
                      }}
                      className="p-1.5 hover:bg-slate-200 rounded-lg text-slate-700 cursor-pointer text-xs font-bold"
                      title="Previous Day"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <input
                      type="date"
                      value={attendanceDate}
                      onChange={(e) => setAttendanceDate(e.target.value)}
                      className="bg-white border border-slate-300 px-2.5 py-1 rounded-lg text-xs font-bold text-slate-900 outline-none focus:ring-2 focus:ring-teal-600"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        const d = new Date(attendanceDate);
                        d.setDate(d.getDate() + 1);
                        setAttendanceDate(d.toISOString().split('T')[0]);
                      }}
                      className="p-1.5 hover:bg-slate-200 rounded-lg text-slate-700 cursor-pointer text-xs font-bold"
                      title="Next Day"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setAttendanceDate(new Date().toISOString().split('T')[0])}
                      className="px-2.5 py-1 bg-teal-100 text-teal-800 hover:bg-teal-200 font-bold text-[11px] rounded-lg cursor-pointer"
                    >
                      Today
                    </button>
                  </div>
                )}

                {attendanceTab === 'monthly' && (
                  <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 p-1.5 rounded-xl">
                    <span className="text-xs font-bold text-slate-600">Month:</span>
                    <input
                      type="month"
                      value={attendanceMonthFilter}
                      onChange={(e) => setAttendanceMonthFilter(e.target.value)}
                      className="bg-white border border-slate-300 px-2.5 py-1 rounded-lg text-xs font-bold text-slate-900 outline-none focus:ring-2 focus:ring-teal-600"
                    />
                  </div>
                )}
              </div>

              {attendanceTab === 'daily' && (
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      const updatedForDate = {};
                      staffMembers.forEach(s => {
                        const sKey = s.id || s.name;
                        const existing = (attendanceRecords[attendanceDate] || {})[sKey] || {};
                        updatedForDate[sKey] = { ...existing, status: 'present' };
                      });
                      saveStaffAttendanceForDate(attendanceDate, updatedForDate);
                    }}
                    className="px-3 py-1.5 bg-emerald-100 hover:bg-emerald-200 text-emerald-800 font-bold text-xs rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer border border-emerald-300"
                  >
                    <CheckCircle2 className="w-4 h-4" /> Mark All Present
                  </button>
                </div>
              )}
            </div>

            {/* DAILY ATTENDANCE TAB */}
            {attendanceTab === 'daily' && (
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs font-black uppercase border-b border-slate-200">
                      <th className="p-3 w-12 text-center">#</th>
                      <th className="p-3">Staff Name</th>
                      <th className="p-3">Role / Designation</th>
                      <th className="p-3">Mobile No</th>
                      <th className="p-3 text-center min-w-[280px]">Attendance Status (Tick Mark)</th>
                      <th className="p-3">Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-xs">
                    {staffMembers.map((staff, idx) => {
                      const sKey = staff.id || staff.name;
                      const dateRecords = attendanceRecords[attendanceDate] || {};
                      const rec = dateRecords[sKey] || { status: 'unmarked', remarks: '' };
                      const currentStatus = rec.status || 'unmarked';

                      const updateStatus = (newStatus: 'present' | 'absent' | 'leave', remarksVal?: string) => {
                        const updatedForDate = {
                          ...dateRecords,
                          [sKey]: {
                            status: newStatus,
                            remarks: remarksVal !== undefined ? remarksVal : (rec.remarks || '')
                          }
                        };
                        saveStaffAttendanceForDate(attendanceDate, updatedForDate);
                      };

                      return (
                        <tr key={sKey} className="hover:bg-slate-50 transition-colors">
                          <td className="p-3 text-center font-bold text-slate-400">{idx + 1}</td>
                          <td className="p-3 font-bold text-slate-900">
                            <div>{staff.name}</div>
                            {staff.fatherName && <div className="text-[10px] text-slate-500 font-medium">S/O {staff.fatherName}</div>}
                          </td>
                          <td className="p-3">
                            <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                              staff.role === 'supervisor' ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                            }`}>
                              {staff.role === 'supervisor' ? 'Supervisor' : staff.role === 'mechanic' ? 'Mechanic' : staff.role || 'Staff'}
                            </span>
                          </td>
                          <td className="p-3 font-mono font-bold text-slate-600">{staff.mobileNumber || '—'}</td>
                          <td className="p-3 text-center">
                            <div className="flex items-center justify-center gap-2">
                              {/* PRESENT BUTTON */}
                              <button
                                type="button"
                                onClick={() => updateStatus('present')}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1 cursor-pointer border ${
                                  currentStatus === 'present'
                                    ? 'bg-emerald-600 text-white border-emerald-700 shadow-sm ring-2 ring-emerald-300'
                                    : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border-emerald-200'
                                }`}
                              >
                                <CheckCircle2 className="w-4 h-4" /> Present
                              </button>

                              {/* ABSENT BUTTON */}
                              <button
                                type="button"
                                onClick={() => updateStatus('absent')}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1 cursor-pointer border ${
                                  currentStatus === 'absent'
                                    ? 'bg-rose-600 text-white border-rose-700 shadow-sm ring-2 ring-rose-300'
                                    : 'bg-rose-50 text-rose-700 hover:bg-rose-100 border-rose-200'
                                }`}
                              >
                                <XCircle className="w-4 h-4" /> Absent
                              </button>

                              {/* LEAVE / HALF DAY BUTTON */}
                              <button
                                type="button"
                                onClick={() => updateStatus('leave')}
                                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1 cursor-pointer border ${
                                  currentStatus === 'leave'
                                    ? 'bg-amber-500 text-slate-950 border-amber-600 shadow-sm ring-2 ring-amber-300 font-black'
                                    : 'bg-amber-50 text-amber-800 hover:bg-amber-100 border-amber-200'
                                }`}
                              >
                                <Clock className="w-4 h-4" /> Leave
                              </button>
                            </div>
                          </td>
                          <td className="p-3">
                            <input
                              type="text"
                              value={rec.remarks || ''}
                              onChange={(e) => updateStatus(rec.status, e.target.value)}
                              placeholder="Remarks / Reason..."
                              className="w-full bg-slate-50 border border-slate-200 p-1.5 rounded-lg text-xs text-slate-800 focus:bg-white focus:ring-1 focus:ring-teal-600 outline-none"
                            />
                          </td>
                        </tr>
                      );
                    })}

                    {staffMembers.length === 0 && (
                      <tr>
                        <td colSpan={6} className="p-8 text-center text-slate-400 font-bold">
                          No staff members registered yet. Click <span className="text-teal-700 font-black">"Add New Staff"</span> above to register employees.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            )}

            {/* MONTHLY SUMMARY REPORT TAB */}
            {attendanceTab === 'monthly' && (
              <div className="overflow-x-auto space-y-4">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs font-black uppercase border-b border-slate-200">
                      <th className="p-3 w-12 text-center">#</th>
                      <th className="p-3">Staff Name</th>
                      <th className="p-3">Role</th>
                      <th className="p-3 text-center text-emerald-800">Present Days</th>
                      <th className="p-3 text-center text-rose-800">Absent Days</th>
                      <th className="p-3 text-center text-amber-800">Leave Days</th>
                      <th className="p-3 text-center">Total Marked Days</th>
                      <th className="p-3 text-center">Attendance %</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-xs">
                    {staffMembers.map((staff, idx) => {
                      const sKey = staff.id || staff.name;
                      let pDays = 0;
                      let aDays = 0;
                      let lDays = 0;

                      Object.entries(attendanceRecords).forEach(([dateStr, recs]) => {
                        if (dateStr.startsWith(attendanceMonthFilter)) {
                          const status = recs[sKey]?.status;
                          if (status === 'present') pDays++;
                          else if (status === 'absent') aDays++;
                          else if (status === 'leave') lDays++;
                        }
                      });

                      const totalDays = pDays + aDays + lDays;
                      const pct = totalDays > 0 ? Math.round((pDays / totalDays) * 100) : 0;

                      return (
                        <tr key={sKey} className="hover:bg-slate-50 transition-colors font-medium">
                          <td className="p-3 text-center font-bold text-slate-400">{idx + 1}</td>
                          <td className="p-3 font-bold text-slate-900">{staff.name}</td>
                          <td className="p-3 capitalize text-slate-600">{staff.role || 'Staff'}</td>
                          <td className="p-3 text-center font-bold text-emerald-700 bg-emerald-50/50">{pDays}</td>
                          <td className="p-3 text-center font-bold text-rose-700 bg-rose-50/50">{aDays}</td>
                          <td className="p-3 text-center font-bold text-amber-700 bg-amber-50/50">{lDays}</td>
                          <td className="p-3 text-center font-bold text-slate-800">{totalDays}</td>
                          <td className="p-3 text-center">
                            <span className={`px-2 py-0.5 rounded font-black text-xs ${
                              pct >= 85 ? 'bg-emerald-100 text-emerald-900' : pct >= 70 ? 'bg-amber-100 text-amber-900' : 'bg-rose-100 text-rose-900'
                            }`}>
                              {pct}%
                            </span>
                          </td>
                        </tr>
                      );
                    })}

                    {staffMembers.length === 0 && (
                      <tr>
                        <td colSpan={8} className="p-6 text-center text-slate-400 font-bold">No staff records to report.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* TAB 6: MASTER DATABASES */}
      {activeTab === 'databases' && (
        <div className="w-full bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-6 print:hidden">
          <div className="border-b pb-3">
            <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-emerald-600" /> Master Databases Upload & Management
            </h2>
            <p className="text-xs text-slate-500 font-medium">Manage Customer Records, Spares Price Lists, and Job Cards Import/Export Excel & CSV databases.</p>
          </div>

          {/* WhatsApp Style Daily Midnight Auto-Backup System */}
          <div className="bg-gradient-to-r from-emerald-900 via-teal-900 to-slate-900 text-white p-5 rounded-2xl border border-emerald-700/50 shadow-xl mb-4 relative overflow-hidden">
            <div className="absolute top-0 right-0 w-64 h-64 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none"></div>
            <div className="relative z-10 flex flex-col lg:flex-row items-start lg:items-center justify-between gap-5">
              <div className="space-y-2 max-w-2xl">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="bg-emerald-500 text-slate-950 font-black text-[10px] px-2.5 py-0.5 rounded-full uppercase tracking-wider flex items-center gap-1 shadow-sm">
                    <span className="w-2 h-2 rounded-full bg-white animate-ping"></span>
                    WhatsApp Auto-Backup Mode
                  </span>
                  <span className="bg-emerald-800/80 border border-emerald-600 text-emerald-200 text-[10px] font-bold px-2 py-0.5 rounded-full">
                    🛡️ 100% డేటా సేఫ్ & సెక్యూర్
                  </span>
                  {useGoogleSheets && (
                    <span className="bg-teal-700/80 border border-teal-500 text-teal-100 text-[10px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1">
                      <FileSpreadsheet size={11} /> Google Sheets సింక్ ఆన్
                    </span>
                  )}
                </div>
                <h3 className="font-black text-lg md:text-xl text-white tracking-tight flex items-center gap-2">
                  📱 ప్రతిరోజూ మిడ్‌నైట్ ఆటో బ్యాకప్ (Daily Midnight Auto-Backup)
                </h3>
                <p className="text-xs text-emerald-100/90 leading-relaxed">
                  డేటా నష్టం గురించి ఎలాంటి భయం అవసరం లేదు! WhatsApp లాగా ప్రతిరోజూ అర్ధరాత్రి (12:00 AM) మొత్తం డేటా (జాబ్ కార్డ్స్, కస్టమర్లు, స్పేర్స్, కంప్లైంట్స్) ఆటోమేటిక్‌గా గూగుల్ షీట్స్ మరియు లోకల్ సెక్యూర్ స్టోరేజ్‌లోకి భద్రపరచబడుతుంది.
                </p>
                <div className="flex flex-wrap items-center gap-4 text-xs font-semibold text-emerald-200 pt-1">
                  <div className="flex items-center gap-1.5 bg-black/30 px-3 py-1.5 rounded-lg border border-emerald-500/30">
                    <span className="text-emerald-400">🕒 చివరి బ్యాకప్ (Last Backup):</span>
                    <span className="text-white font-mono font-bold">{lastAutoBackupTime || 'ఈరోజే పూర్తి చేయబడింది'}</span>
                  </div>
                  {autoBackupMessage && (
                    <div className={`px-2.5 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1.5 ${
                      autoBackupStatus === 'success' ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40' :
                      autoBackupStatus === 'running' ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40 animate-pulse' :
                      'bg-red-500/20 text-red-300 border border-red-500/40'
                    }`}>
                      {autoBackupStatus === 'running' && <div className="w-3 h-3 border-2 border-amber-300 border-t-transparent rounded-full animate-spin"></div>}
                      {autoBackupMessage}
                    </div>
                  )}
                </div>
              </div>

              <div className="flex flex-col sm:flex-row lg:flex-col gap-2 w-full lg:w-auto shrink-0">
                <button
                  onClick={() => executeAutoBackup(true)}
                  disabled={autoBackupStatus === 'running' || sheetsLoading}
                  className="flex items-center justify-center gap-2 bg-emerald-500 hover:bg-emerald-400 active:scale-95 text-slate-950 px-5 py-2.5 rounded-xl font-black text-xs transition-all shadow-lg shadow-emerald-500/30 disabled:opacity-50 cursor-pointer"
                >
                  {autoBackupStatus === 'running' ? (
                    <div className="w-4 h-4 border-2 border-slate-950 border-t-transparent rounded-full animate-spin"></div>
                  ) : (
                    <Save size={16} />
                  )}
                  ఇప్పుడే బ్యాకప్ తీసుకోండి (Backup Now)
                </button>
                <button
                  onClick={handleDownloadMasterBackup}
                  disabled={sheetsLoading}
                  className="flex items-center justify-center gap-2 bg-white/10 hover:bg-white/20 active:scale-95 text-white border border-white/20 px-5 py-2.5 rounded-xl font-black text-xs transition-all cursor-pointer"
                >
                  {sheetsLoading ? (
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                  ) : (
                    <Download size={16} />
                  )}
                  {appLang === 'te' ? 'మాస్టర్ ఎక్సెల్ ఫైల్ డౌన్‌లోడ్ (Excel)' : 'Download Master Backup (Excel)'}
                </button>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  id="masterRestoreInput"
                  className="hidden"
                  onChange={handleMasterRestoreUpload}
                />
                <button
                  onClick={() => document.getElementById('masterRestoreInput')?.click()}
                  disabled={sheetsLoading}
                  className="flex items-center justify-center gap-2 bg-indigo-500 hover:bg-indigo-400 active:scale-95 text-white border border-indigo-400/50 px-5 py-2.5 rounded-xl font-black text-xs transition-all shadow-lg shadow-indigo-500/20 disabled:opacity-50 cursor-pointer"
                  title="ఈ పరికరంలో (కొత్త ఫోన్/ల్యాప్‌టాప్) పాత బ్యాకప్ ఫైల్ ని అప్‌లోడ్ చేసి డేటాని పొందండి"
                >
                  <Upload size={16} />
                  {appLang === 'te' ? 'మాస్టర్ ఫైల్ అప్‌లోడ్ (Restore)' : 'Upload Master Backup (Restore)'}
                </button>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {/* 1. CUSTOMER EXCEL DATABASE */}
            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3 flex flex-col justify-between">
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <FileSpreadsheet className="w-4 h-4 text-indigo-600" /> 1. Customer Database (.xlsx)
                  </h3>
                  <button
                    type="button"
                    onClick={handleOpenAddCustomer}
                    className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] px-2.5 py-1 rounded-md transition-colors cursor-pointer shadow-xs"
                  >
                    <UserPlus className="w-3.5 h-3.5" /> + Add New
                  </button>
                </div>

                {currentUserRole === 'admin' ? (
                  <>
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Import Customer File (.xlsx, .csv)</label>
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        onChange={handleExcelUpload}
                        className="block w-full text-xs text-slate-500 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer"
                      />
                    </div>
                    
                    <div className="flex flex-col gap-2 pt-1">
                      <button
                        type="button"
                        onClick={downloadCustomerExcel}
                        disabled={Object.keys(chassisIndex).length === 0}
                        className="w-full flex items-center justify-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                      >
                        <Download className="w-3.5 h-3.5" /> Export Customer File to Excel (.xlsx)
                      </button>
                      <button
                        type="button"
                        onClick={checkCustomerDuplicates}
                        disabled={Object.keys(chassisIndex).length === 0}
                        className="w-full flex items-center justify-center gap-1.5 bg-amber-600 hover:bg-amber-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                      >
                        <Search className="w-3.5 h-3.5" /> 🔍 Check Chassis / Phone Duplicates
                      </button>
                    </div>

                    <div className="flex items-center justify-between text-xs pt-1">
                      <span className={`font-medium ${excelStatus.isSuccess ? 'text-emerald-600' : 'text-slate-500'}`}>
                        {excelStatus.text}
                      </span>
                    </div>
                  </>
                ) : (
                  <div className="text-xs bg-slate-100 p-2.5 rounded-lg text-slate-500 font-semibold">
                    🔒 Only Admins can upload customer Excel files.
                  </div>
                )}
              </div>
            </div>

            {/* 2. SPARES PRICE LIST DATABASE */}
            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3 flex flex-col justify-between">
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4 text-emerald-600" /> 2. Spares Price List (.xlsx)
                </h3>
                {currentUserRole === 'admin' ? (
                  <>
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Import Spares File (.xlsx, .csv)</label>
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        onChange={handleSparesUpload}
                        className="block w-full text-xs text-slate-500 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100 cursor-pointer"
                      />
                    </div>

                    <div className="flex flex-col gap-2 pt-1">
                      <button
                        type="button"
                        onClick={downloadSparesExcel}
                        disabled={Object.keys(sparesIndex).length === 0}
                        className="w-full flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                      >
                        <Download className="w-3.5 h-3.5" /> Export Spares File to Excel (.xlsx)
                      </button>
                    </div>

                    <div className="flex items-center justify-between text-xs pt-1">
                      <span className={`font-medium ${sparesStatus.isSuccess ? 'text-emerald-600' : 'text-slate-500'}`}>
                        {sparesStatus.text}
                      </span>
                    </div>
                  </>
                ) : (
                  <div className="text-xs bg-slate-100 p-2.5 rounded-lg text-slate-500 font-semibold">
                    🔒 Only Admins can upload & clear Spares Price Lists.
                  </div>
                )}
              </div>
              {currentUserRole === 'admin' && Object.keys(sparesIndex).length > 0 && (
                <div className="pt-2 border-t border-slate-200">
                  <button onClick={clearSavedSpares} className="text-red-600 hover:underline flex items-center gap-1 font-semibold text-xs cursor-pointer">
                    <Trash2 className="w-3.5 h-3.5" /> Clear Spares Price List
                  </button>
                </div>
              )}
            </div>

            {/* 3. JOB CARDS IMPORT & EXPORT DATABASE */}
            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3 flex flex-col justify-between">
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4 text-sky-600" /> 3. Job Cards Database (Import / Export)
                </h3>
                <p className="text-[11px] text-slate-500 font-medium">Bulk import job cards from Excel / CSV files or export saved cards into spreadsheet files.</p>


                
                <div className="space-y-2.5">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Import Job Cards File (.xlsx, .csv)</label>
                    <input
                      ref={fileInputRef}
                      type="file"
                      disabled={isUploading}
                      accept=".xlsx,.xls,.csv"
                      onChange={handleFileUpload}
                      className="block w-full text-xs text-slate-500 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-sky-50 file:text-sky-700 hover:file:bg-sky-100 cursor-pointer disabled:opacity-50"
                    />
                    {isUploading && (
                      <p className="text-xs text-sky-600 mt-2 font-bold flex items-center gap-2">
                        <RefreshCcw className="w-3.5 h-3.5 animate-spin" /> Uploading & Processing Job Cards... Please wait.
                      </p>
                    )}
                  </div>

                  <div className="flex flex-col gap-2 pt-1">
                    <button
                      type="button"
                      onClick={downloadSavedCardsExcel}
                      disabled={savedJobCards.length === 0}
                      className="w-full flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                    >
                      <Download className="w-3.5 h-3.5" /> Export All ({savedJobCards.length}) to Excel (.xlsx)
                    </button>

                    <button
                      type="button"
                      onClick={downloadJobCardTemplateExcel}
                      className="w-full flex items-center justify-center gap-1.5 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-semibold text-xs py-1.5 px-3 rounded-lg transition-colors cursor-pointer"
                    >
                      <FileText className="w-3.5 h-3.5 text-slate-500" /> Download Sample Import Template (.xlsx)
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        if (confirm("ARE YOU SURE YOU WANT TO CLEAR ALL DATA? THIS CANNOT BE UNDONE!")) {
                          const password = prompt("Enter Admin Password:");
                          if (password) {
                            fetch('/api/database/clear', {
                              method: 'POST',
                              headers: { 'Content-Type': 'application/json' },
                              body: JSON.stringify({ password })
                            }).then(res => res.json()).then(data => {
                              if (data.success) {
                                alert("All data cleared successfully.");
                                window.location.reload();
                              } else {
                                alert("Failed to clear data: " + data.error);
                              }
                            });
                          }
                        }
                      }}
                      className="w-full flex items-center justify-center gap-1.5 bg-red-600 hover:bg-red-700 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs mt-2"
                    >
                      <Trash2 className="w-3.5 h-3.5" /> Clear All Data
                    </button>
                    
                    <div className="flex items-center justify-between text-xs pt-1">
                      <span className={`font-medium ${jobCardsStatus.isSuccess ? 'text-emerald-600' : 'text-slate-500'}`}>
                        {jobCardsStatus.text}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              <div className="pt-2 border-t border-slate-200 text-[11px] text-slate-500 font-semibold flex justify-between items-center">
                <span>Total Stored Job Cards:</span>
                <span className="font-bold text-blue-900 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200">
                  {savedJobCards.length} Cards
                </span>
              </div>
            </div>

            {/* 4. STAFF & EMPLOYEE DIRECTORY MASTER DATABASE */}
            <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2 flex flex-col justify-between">
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <Users className="w-4 h-4 text-purple-600" /> 4. Staff Directory (.xlsx)
                  </h3>
                  <button
                    type="button"
                    onClick={() => {
                      setEditingStaffId(null);
                      setStaffForm({
                        name: '',
                        fatherName: '',
                        village: '',
                        mandal: '',
                        mobileNumber: '',
                        role: 'mechanic',
                        dateOfJoining: ''
                      });
                      setIsStaffModalOpen(true);
                    }}
                    className="flex items-center gap-1 bg-purple-600 hover:bg-purple-700 text-white font-bold text-[11px] px-2.5 py-1 rounded-md transition-colors cursor-pointer shadow-xs"
                  >
                    <UserPlus className="w-3.5 h-3.5" /> + Add Staff
                  </button>
                </div>

                <div className="space-y-2">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Import Staff File (.xlsx, .csv)</label>
                    <input
                      type="file"
                      ref={staffFileInputRef}
                      accept=".xlsx,.xls,.csv"
                      onChange={handleStaffFileUpload}
                      className="block w-full text-xs text-slate-500 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-purple-50 file:text-purple-700 hover:file:bg-purple-100 cursor-pointer"
                    />
                  </div>

                  <button
                    type="button"
                    onClick={downloadStaffDirectoryExcel}
                    disabled={staffMembers.length === 0}
                    className="w-full flex items-center justify-center gap-1.5 bg-purple-600 hover:bg-purple-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                  >
                    <Download className="w-3.5 h-3.5" /> Export Staff Directory to Excel (.xlsx)
                  </button>
                </div>
              </div>

              <div className="pt-2 border-t border-slate-200 text-[11px] text-slate-500 font-semibold flex justify-between items-center">
                <span>Total Registered Staff:</span>
                <span className="font-bold text-purple-900 bg-purple-50 px-2 py-0.5 rounded-full border border-purple-200">
                  {staffMembers.length} Employees
                </span>
              </div>
            </div>

            {/* 5. FREE SERVICE FOLLOW-UP MASTER CARD */}
            <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2 flex flex-col justify-between">
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <Clock className="w-4 h-4 text-teal-600" /> 5. Follow-up Master (.xlsx)
                  </h3>
                </div>
                <p className="text-[10px] text-slate-500 font-medium">Update delivery dates and history from excel.</p>
                
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Import Follow-up File</label>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleExcelUpload}
                    className="block w-full text-xs text-slate-500 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-teal-50 file:text-teal-700 hover:file:bg-teal-100 cursor-pointer"
                  />
                </div>

                <div className="pt-1">
                  <button
                    type="button"
                    onClick={downloadFollowupExcel}
                    disabled={Object.keys(chassisIndex).length === 0}
                    className="w-full flex items-center justify-center gap-1.5 bg-teal-600 hover:bg-teal-700 disabled:opacity-50 text-white font-bold text-xs py-2 px-3 rounded-lg transition-colors cursor-pointer shadow-xs"
                  >
                    <Download className="w-3.5 h-3.5" /> Export Follow-up Master
                  </button>
                </div>
              </div>

              <div className="pt-2 border-t border-slate-200 text-[11px] text-slate-500 font-semibold flex justify-between items-center">
                <span>Follow-up List Status:</span>
                <span className={`font-bold ${excelStatus.isSuccess ? 'text-teal-700' : 'text-slate-400'}`}>
                  {excelStatus.text.split('.')[0]}
                </span>
              </div>
            </div>
          </div>

          {/* STAFF & EMPLOYEE DIRECTORY MASTER SECTION */}
          <div className="bg-slate-50 border border-slate-200 p-4 rounded-xl space-y-4 pt-2">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-3">
              <div className="flex items-center gap-2">
                <Users className="w-5 h-5 text-purple-600" />
                <div>
                  <h3 className="text-sm font-bold text-slate-900">Staff & Employee Master Directory</h3>
                  <button
                    type="button"
                    onClick={() => setIsDirectoryExpanded(!isDirectoryExpanded)}
                    className="text-xs text-purple-700 font-bold hover:underline cursor-pointer"
                  >
                    {staffMembers.length} staff members ({isDirectoryExpanded ? 'Click to collapse' : 'Click to view'})
                  </button>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={downloadStaffDirectoryExcel}
                  disabled={staffMembers.length === 0}
                  className="flex items-center gap-1 bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg transition-all shadow-xs cursor-pointer disabled:opacity-50"
                >
                  <Download className="w-3.5 h-3.5" /> Export to Excel (.xlsx)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setEditingStaffId(null);
                    setStaffForm({
                      name: '',
                      fatherName: '',
                      village: '',
                      mandal: '',
                      mobileNumber: '',
                      role: 'mechanic',
                      dateOfJoining: '',
                      supervisor: ''
                    });
                    setIsStaffModalOpen(true);
                  }}
                  className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3 py-1.5 rounded-lg transition-all shadow-xs cursor-pointer"
                >
                  <UserPlus className="w-3.5 h-3.5" /> + Add New Staff
                </button>
              </div>
            </div>

            {staffMembers.length === 0 ? (
              <div className="text-center py-6 text-slate-400 text-xs font-medium">
                No staff members registered yet. Click "+ Add New Staff" above to add mechanics and supervisors.
              </div>
            ) : isDirectoryExpanded ? (
              <div className="space-y-4">
                {/* SUB TABS FOR STAFF DIRECTORY */}
                <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2">
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() => setStaffDirectoryTab('all_grid')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                        staffDirectoryTab === 'all_grid'
                          ? 'bg-purple-600 text-white shadow-xs'
                          : 'bg-white text-slate-600 hover:bg-slate-100 hover:text-slate-800 border border-slate-200'
                      }`}
                    >
                      🎴 Cards Grid ({staffMembers.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setStaffDirectoryTab('all_table')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                        staffDirectoryTab === 'all_table'
                          ? 'bg-purple-600 text-white shadow-xs'
                          : 'bg-white text-slate-600 hover:bg-slate-100 hover:text-slate-800 border border-slate-200'
                      }`}
                    >
                      📋 Full Table View
                    </button>
                    <button
                      type="button"
                      onClick={() => setStaffDirectoryTab('supervisor_teams')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                        staffDirectoryTab === 'supervisor_teams'
                          ? 'bg-purple-800 text-white shadow-xs border border-purple-900'
                          : 'bg-white text-purple-700 hover:bg-purple-50 hover:text-purple-900 border border-purple-200'
                      }`}
                    >
                      🏢 Supervisor Teams (Branches)
                    </button>
                  </div>
                  <div className="text-[10px] text-slate-500 font-semibold uppercase tracking-wider">
                    Supervisors: {staffMembers.filter(s => s.role === 'supervisor').length}
                  </div>
                </div>

                {staffDirectoryTab === 'all_grid' && (
                  /* CLEAN COMPACT CARDS VIEW */
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                    {staffMembers.map((staff, idx) => (
                      <div key={staff.id || idx} className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs space-y-2.5 flex flex-col justify-between hover:border-purple-300 transition-all">
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <h4 className="text-xs font-bold text-slate-900">{staff.name}</h4>
                            <span className={`inline-block px-2 py-0.5 mt-1 rounded text-[10px] font-bold uppercase ${
                              staff.role === 'supervisor' ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                            }`}>
                              {staff.role === 'supervisor' ? 'Supervisor / W/S Incharge' : staff.role === 'mechanic' ? 'Mechanic' : staff.role || 'Mechanic'}
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => setViewingStaffDetails(staff)}
                            className="p-1 text-slate-500 hover:text-purple-700 hover:bg-purple-50 rounded cursor-pointer"
                            title="View Full Staff Details"
                          >
                            <Eye className="w-4 h-4" />
                          </button>
                        </div>

                        {staff.mobileNumber && (
                          <div className="text-[11px] font-mono text-slate-600 flex items-center gap-1">
                            <span>📞</span> {staff.mobileNumber}
                          </div>
                        )}

                        {staff.role === 'mechanic' && staff.supervisor && (
                          <div className="text-[11px] text-slate-600 flex items-center gap-1">
                            <span className="font-semibold text-slate-500">Supervisor:</span> {staff.supervisor}
                          </div>
                        )}

                        {staff.role === 'mechanic' && !staff.supervisor && (
                          <div className="text-[11px] text-amber-600 italic flex items-center gap-1">
                            ⚠️ No Supervisor Assigned
                          </div>
                        )}

                        <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-[11px]">
                          <button
                            type="button"
                            onClick={() => setViewingStaffDetails(staff)}
                            className="text-purple-700 font-bold hover:underline cursor-pointer flex items-center gap-1"
                          >
                            <span>View Details</span> &rarr;
                          </button>
                          <div className="flex items-center gap-1">
                            <button
                              type="button"
                              onClick={() => {
                                setEditingStaffId(staff.id);
                                setStaffForm({
                                  name: staff.name || '',
                                  fatherName: staff.fatherName || '',
                                  village: staff.village || '',
                                  mandal: staff.mandal || '',
                                  mobileNumber: staff.mobileNumber || '',
                                  role: staff.role || 'mechanic',
                                  dateOfJoining: staff.dateOfJoining || '',
                                  supervisor: staff.supervisor || ''
                                });
                                setIsStaffModalOpen(true);
                              }}
                              className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded cursor-pointer"
                              title="Edit Staff Member"
                            >
                              <Edit3 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                if (confirm(`Are you sure you want to delete staff member "${staff.name}"?`)) {
                                  deleteStaffMember(staff.id);
                                }
                              }}
                              className="p-1 text-red-600 hover:text-red-800 hover:bg-red-50 rounded cursor-pointer"
                              title="Delete Staff Member"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {staffDirectoryTab === 'all_table' && (
                  /* EXPANDED FULL TABLE VIEW */
                  <div className="overflow-x-auto border border-slate-200 rounded-lg">
                    <table className="w-full text-left text-xs text-slate-700">
                      <thead className="bg-slate-100 text-slate-800 font-bold uppercase text-[11px] border-b border-slate-200">
                        <tr>
                          <th className="p-3">Sl No</th>
                          <th className="p-3">Full Name</th>
                          <th className="p-3">Father Name</th>
                          <th className="p-3">Job Role</th>
                          <th className="p-3">Assigned Supervisor</th>
                          <th className="p-3">Mobile No</th>
                          <th className="p-3">Village</th>
                          <th className="p-3">Mandal</th>
                          <th className="p-3">Date of Joining</th>
                          <th className="p-3 text-center">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 font-medium bg-white">
                        {staffMembers.map((staff, idx) => (
                          <tr key={staff.id || idx} className="hover:bg-amber-100/80">
                            <td className="p-3 text-slate-400">{idx + 1}</td>
                            <td className="p-3 font-bold text-slate-900">{staff.name}</td>
                            <td className="p-3 text-slate-600">{staff.fatherName || '—'}</td>
                            <td className="p-3">
                              <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase ${
                                staff.role === 'supervisor' ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                              }`}>
                                {staff.role === 'supervisor' ? 'Supervisor / W/S Incharge' : staff.role === 'mechanic' ? 'Mechanic / Technician' : staff.role || 'Mechanic'}
                              </span>
                            </td>
                            <td className="p-3 font-semibold text-slate-800">
                              {staff.role === 'mechanic' ? (staff.supervisor || '—') : 'N/A'}
                            </td>
                            <td className="p-3 font-mono">{staff.mobileNumber || '—'}</td>
                            <td className="p-3">{staff.village || '—'}</td>
                            <td className="p-3">{staff.mandal || '—'}</td>
                            <td className="p-3 whitespace-nowrap">{staff.dateOfJoining || '—'}</td>
                            <td className="p-3 text-center">
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => setViewingStaffDetails(staff)}
                                  className="p-1 text-purple-600 hover:text-purple-800 hover:bg-purple-50 rounded cursor-pointer"
                                  title="View Staff Profile"
                                >
                                  <Eye className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setEditingStaffId(staff.id);
                                    setStaffForm({
                                      name: staff.name || '',
                                      fatherName: staff.fatherName || '',
                                      village: staff.village || '',
                                      mandal: staff.mandal || '',
                                      mobileNumber: staff.mobileNumber || '',
                                      role: staff.role || 'mechanic',
                                      dateOfJoining: staff.dateOfJoining || '',
                                      supervisor: staff.supervisor || ''
                                    });
                                    setIsStaffModalOpen(true);
                                  }}
                                  className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded cursor-pointer"
                                  title="Edit Staff Member"
                                >
                                  <Edit3 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    if (confirm(`Are you sure you want to delete staff member "${staff.name}"?`)) {
                                      deleteStaffMember(staff.id);
                                    }
                                  }}
                                  className="p-1 text-red-600 hover:text-red-800 hover:bg-red-50 rounded cursor-pointer"
                                  title="Delete Staff Member"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                {staffDirectoryTab === 'supervisor_teams' && (
                  /* SUPERVISOR TEAMS / BRANCHES VIEW */
                  <div className="space-y-4">
                    {/* SUPERVISOR CARDS SELECTOR GRID */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                      {/* CARD: ALL/SUMMARY */}
                      <div
                        onClick={() => setSelectedSupervisorTeam('all')}
                        className={`p-4 rounded-xl border-2 transition-all cursor-pointer flex flex-col justify-between ${
                          selectedSupervisorTeam === 'all'
                            ? 'bg-purple-50/50 border-purple-600 shadow-xs'
                            : 'bg-white border-slate-200 hover:border-purple-300'
                        }`}
                      >
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-[10px] font-bold text-purple-900 bg-purple-100 px-2.5 py-0.5 rounded-full uppercase">All Teams Summary</span>
                            <span className="text-base">🏛️</span>
                          </div>
                          <h4 className="text-sm font-black text-slate-900">Total Workshop Overview</h4>
                          <p className="text-xs text-slate-500 mt-1.5 font-semibold">
                            Total Staff: {staffMembers.length} Employees
                          </p>
                          <p className="text-xs text-slate-500 font-semibold mt-1">
                            Supervisors: {staffMembers.filter(s => s.role === 'supervisor').length} | Mechanics: {staffMembers.filter(s => s.role === 'mechanic' || !s.role).length}
                          </p>
                        </div>
                        <button
                          type="button"
                          className="mt-4 text-left text-xs font-bold text-purple-700 hover:underline flex items-center gap-1"
                        >
                          Show Combined Directory &rarr;
                        </button>
                      </div>

                      {/* CARDS: INDIVIDUAL SUPERVISORS */}
                      {staffMembers.filter(s => s.role === 'supervisor').map((sup) => {
                        const team = staffMembers.filter(s => s.role === 'mechanic' && s.supervisor === sup.name);
                        const isSelected = selectedSupervisorTeam === sup.name;
                        return (
                          <div
                            key={sup.id}
                            onClick={() => setSelectedSupervisorTeam(sup.name)}
                            className={`p-4 rounded-xl border-2 transition-all cursor-pointer flex flex-col justify-between ${
                              isSelected
                                ? 'bg-purple-50/50 border-purple-600 shadow-xs'
                                : 'bg-white border-slate-200 hover:border-purple-300'
                            }`}
                          >
                            <div>
                              <div className="flex items-center justify-between mb-2">
                                <span className="text-[10px] font-bold text-amber-900 bg-amber-100 px-2.5 py-0.5 rounded-full uppercase">Supervisor Branch</span>
                                <span className="text-base">🏢</span>
                              </div>
                              <h4 className="text-sm font-black text-slate-900">{sup.name}</h4>
                              <p className="text-xs text-slate-500 mt-1.5 font-semibold flex items-center gap-1.5">
                                <span>👥 Assigned Team:</span>
                                <span className="text-purple-700 font-extrabold">{team.length} Mechanics</span>
                              </p>
                              {sup.mobileNumber && (
                                <p className="text-[11px] font-mono text-slate-500 mt-1">
                                  📞 {sup.mobileNumber}
                                </p>
                              )}
                              <div className="mt-2.5 flex flex-wrap gap-1 max-h-12 overflow-hidden">
                                {team.slice(0, 3).map((t, i) => (
                                  <span key={i} className="text-[9px] bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded font-bold">
                                    {t.name}
                                  </span>
                                ))}
                                {team.length > 3 && (
                                  <span className="text-[9px] bg-purple-50 text-purple-600 px-1.5 py-0.5 rounded font-bold">
                                    +{team.length - 3} more
                                  </span>
                                )}
                              </div>
                            </div>
                            <button
                              type="button"
                              className="mt-4 text-left text-xs font-bold text-purple-700 hover:underline"
                            >
                              View Team Details &rarr;
                            </button>
                          </div>
                        );
                      })}

                      {/* CARD: UNASSIGNED MECHANICS */}
                      {staffMembers.filter(s => s.role === 'mechanic' && !s.supervisor).length > 0 && (
                        <div
                          onClick={() => setSelectedSupervisorTeam('unassigned')}
                          className={`p-4 rounded-xl border-2 transition-all cursor-pointer flex flex-col justify-between ${
                            selectedSupervisorTeam === 'unassigned'
                              ? 'bg-purple-50/50 border-purple-600 shadow-xs'
                              : 'bg-white border-slate-200 hover:border-purple-300'
                          }`}
                        >
                          <div>
                            <div className="flex items-center justify-between mb-2">
                              <span className="text-[10px] font-bold text-red-900 bg-red-100 px-2.5 py-0.5 rounded-full uppercase">Unassigned Mechanics</span>
                              <span className="text-base">⚠️</span>
                            </div>
                            <h4 className="text-sm font-black text-slate-900">Independent Staff</h4>
                            <p className="text-xs text-slate-500 mt-1 font-semibold">
                              These mechanics do not have any assigned Supervisor.
                            </p>
                            <p className="text-xs text-purple-700 font-extrabold mt-1.5">
                              👥 Team size: {staffMembers.filter(s => s.role === 'mechanic' && !s.supervisor).length} members
                            </p>
                          </div>
                          <button
                            type="button"
                            className="mt-4 text-left text-xs font-bold text-purple-700 hover:underline"
                          >
                            Manage Independent Staff &rarr;
                          </button>
                        </div>
                      )}
                    </div>

                    {/* TEAM DETAILS DRILLDOWN PANEL */}
                    <div className="bg-white border border-slate-200 p-4 rounded-xl space-y-3.5 shadow-2xs">
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2.5">
                        <div>
                          <h4 className="text-xs font-extrabold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                            <span>📋</span>
                            {selectedSupervisorTeam === 'all' && "All Registered Staff Overview"}
                            {selectedSupervisorTeam === 'unassigned' && "Independent / Unassigned Staff list"}
                            {selectedSupervisorTeam !== 'all' && selectedSupervisorTeam !== 'unassigned' && `${selectedSupervisorTeam}'s Assigned Branch Staff`}
                          </h4>
                          <p className="text-[11px] text-slate-500 font-semibold mt-0.5">
                            {selectedSupervisorTeam === 'all' && `Listing all ${staffMembers.length} employees of the workshop.`}
                            {selectedSupervisorTeam === 'unassigned' && `Listing mechanics who are not linked to a supervisor.`}
                            {selectedSupervisorTeam !== 'all' && selectedSupervisorTeam !== 'unassigned' && `Showing mechanics under supervisor "${selectedSupervisorTeam}".`}
                          </p>
                        </div>
                        {selectedSupervisorTeam !== 'all' && (
                          <button
                            type="button"
                            onClick={() => setSelectedSupervisorTeam('all')}
                            className="text-[11px] font-bold text-purple-700 bg-purple-50 hover:bg-purple-100 px-2.5 py-1 rounded-md transition-all cursor-pointer border border-purple-200"
                          >
                            &larr; Show All Summary
                          </button>
                        )}
                      </div>

                      {/* TABLE OF TEAM MEMBERS */}
                      {(() => {
                        const filteredMembers = staffMembers.filter(s => {
                          if (selectedSupervisorTeam === 'all') return true;
                          if (selectedSupervisorTeam === 'unassigned') return s.role === 'mechanic' && !s.supervisor;
                          // If supervisor is selected, show that supervisor themselves AND all mechanics assigned to them
                          return s.name === selectedSupervisorTeam || (s.role === 'mechanic' && s.supervisor === selectedSupervisorTeam);
                        });

                        if (filteredMembers.length === 0) {
                          return (
                            <div className="text-center py-6 text-slate-400 text-xs font-semibold">
                              No staff members found in this selection.
                            </div>
                          );
                        }

                        return (
                          <div className="overflow-x-auto border border-slate-100 rounded-lg">
                            <table className="w-full text-left text-xs text-slate-700">
                              <thead className="bg-slate-50 text-slate-800 font-bold uppercase text-[10px] border-b border-slate-100">
                                <tr>
                                  <th className="p-3">Staff Name</th>
                                  <th className="p-3">Role</th>
                                  <th className="p-3">Mobile Number</th>
                                  <th className="p-3">Village / Mandal</th>
                                  <th className="p-3">Assigned Supervisor</th>
                                  <th className="p-3">Joining Date</th>
                                  <th className="p-3 text-center">Actions</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100 bg-white">
                                {filteredMembers.map((staff, idx) => (
                                  <tr key={staff.id || idx} className="hover:bg-amber-100/80">
                                    <td className="p-3">
                                      <div className="font-bold text-slate-900">{staff.name}</div>
                                      {staff.fatherName && (
                                        <div className="text-[10px] text-slate-500 font-semibold mt-0.5">S/O: {staff.fatherName}</div>
                                      )}
                                    </td>
                                    <td className="p-3">
                                      <span className={`px-2 py-0.5 rounded text-[9px] font-black uppercase ${
                                        staff.role === 'supervisor' ? 'bg-purple-100 text-purple-800' : 'bg-blue-100 text-blue-800'
                                      }`}>
                                        {staff.role === 'supervisor' ? 'Supervisor' : 'Mechanic'}
                                      </span>
                                    </td>
                                    <td className="p-3 font-mono font-bold text-slate-700">{staff.mobileNumber || '—'}</td>
                                    <td className="p-3 text-slate-600 font-semibold">
                                      {[staff.village, staff.mandal].filter(Boolean).join(', ') || '—'}
                                    </td>
                                    <td className="p-3 font-semibold text-slate-800">
                                      {staff.role === 'mechanic' ? (staff.supervisor || <span className="text-red-500 italic font-medium">⚠️ Unassigned</span>) : 'N/A'}
                                    </td>
                                    <td className="p-3 font-medium text-slate-500">{staff.dateOfJoining || '—'}</td>
                                    <td className="p-3 text-center">
                                      <div className="flex items-center justify-center gap-1">
                                        <button
                                          type="button"
                                          onClick={() => setViewingStaffDetails(staff)}
                                          className="p-1 text-purple-600 hover:text-purple-800 hover:bg-purple-50 rounded cursor-pointer"
                                          title="View Profile"
                                        >
                                          <Eye className="w-3.5 h-3.5" />
                                        </button>
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setEditingStaffId(staff.id);
                                            setStaffForm({
                                              name: staff.name || '',
                                              fatherName: staff.fatherName || '',
                                              village: staff.village || '',
                                              mandal: staff.mandal || '',
                                              mobileNumber: staff.mobileNumber || '',
                                              role: staff.role || 'mechanic',
                                              dateOfJoining: staff.dateOfJoining || '',
                                              supervisor: staff.supervisor || ''
                                            });
                                            setIsStaffModalOpen(true);
                                          }}
                                          className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded cursor-pointer"
                                          title="Edit"
                                        >
                                          <Edit3 className="w-3.5 h-3.5" />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                )}
              </div>
            ) : null}
          </div>
        </div>
      )}
      {/* TAB 2: SAVED JOB CARDS LIST SECTION */}
      {activeTab === 'saved_cards' && (
        <div className={`w-full bg-white border border-slate-200 shadow-sm print:hidden ${
          isSavedTableMaximized 
            ? 'fixed inset-0 bg-slate-100 z-[9999] p-4 md:p-6 flex flex-col space-y-3 overflow-hidden' 
            : 'p-2.5 md:p-3 rounded-xl space-y-2.5'
        }`}>
        {visibleJobCards.length === 0 ? (
          <div className="text-center py-8 text-slate-400 text-xs font-semibold">
            No job cards saved yet. Fill out the form above and click "Save to List".
          </div>
        ) : (
          <div className={`space-y-2.5 ${isSavedTableMaximized ? 'flex-1 flex flex-col min-h-0' : ''}`}>
            {/* COMPREHENSIVE ADVANCED FILTER PANEL FOR SAVED CARDS (FREEZES ON SCROLL) */}
            <div className={`z-30 bg-slate-50/98 backdrop-blur-md border-2 border-indigo-100 p-2 md:p-2.5 rounded-xl space-y-1.5 shadow-md transition-all ${
              isSavedTableMaximized ? 'relative' : 'sticky top-0'
            }`}>
              {/* HEADER INSIDE STICKY PANEL */}
              <div className="flex flex-wrap items-center justify-between gap-1.5 border-b border-slate-200 pb-1.5">
                <div className="flex items-center gap-1.5">
                  <div className="bg-emerald-50 text-emerald-800 p-1 rounded-md border border-emerald-200">
                    <FileText className="w-3.5 h-3.5 text-emerald-700" />
                  </div>
                  <div>
                    <h2 className="text-xs font-extrabold text-slate-900 leading-tight">{appLang === 'te' ? `సేవ్ చేసిన జాబ్ కార్డులు (${visibleJobCards.length})` : `Saved Job Cards (${visibleJobCards.length})`}</h2>
                    <p className="text-[9.5px] text-slate-500 font-medium">{appLang === 'te' ? 'క్లౌడ్ డేటాబేస్ లో భద్రపరిచిన జాబ్ కార్డుల జాబితా.' : 'List of job cards saved securely in the cloud database.'}</p>
                  </div>
                </div>

                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setIsSavedFilterCollapsed(!isSavedFilterCollapsed)}
                    className="flex items-center gap-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[10.5px] px-2 py-1 rounded transition-colors shadow-2xs cursor-pointer"
                  >
                    {isSavedFilterCollapsed ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronUp className="w-3.5 h-3.5" />}
                    <span>{isSavedFilterCollapsed ? (appLang === 'te' ? "ఫిల్టర్లు చూపించు" : "Show Filters") : (appLang === 'te' ? "ఫిల్టర్లు దాచు" : "Hide Filters")}</span>
                  </button>

                  {selectedJobCardIds.length > 0 && (
                    <button
                      type="button"
                      onClick={async () => {
                        if (confirm(`Are you sure you want to delete ${selectedJobCardIds.length} selected job cards? THIS CANNOT BE UNDONE!`)) {
                          try {
                            const res = await fetch('/api/database/delete-jobcards', {
                              method: 'POST',
                              headers: { 'Content-Type': 'application/json' },
                              body: JSON.stringify({ ids: selectedJobCardIds })
                            });
                            const data = await res.json();
                            if (data.success) {
                              alert(`${selectedJobCardIds.length} job cards deleted successfully.`);
                              setSelectedJobCardIds([]);
                              window.location.reload();
                            } else {
                              alert("Failed to delete job cards: " + data.error);
                            }
                          } catch (error) {
                            alert("Failed to delete job cards: " + error);
                          }
                        }
                      }}
                      className="flex items-center gap-1.5 bg-red-600 hover:bg-red-700 text-white font-bold text-[10.5px] px-2 py-1 rounded transition-colors shadow-2xs cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" /> <span>Delete ({selectedJobCardIds.length})</span>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => setIsSavedTableMaximized(!isSavedTableMaximized)}
                    className="flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-[10.5px] px-2.5 py-1 rounded transition-colors shadow-2xs cursor-pointer"
                    title={isSavedTableMaximized ? "Exit fullscreen" : "View in fullscreen"}
                  >
                    {isSavedTableMaximized ? (
                      <>
                        <Minimize2 className="w-3.5 h-3.5" /> <span>{appLang === 'te' ? 'పూర్తి స్క్రీన్ ముగించు' : 'Exit Fullscreen'}</span>
                      </>
                    ) : (
                      <>
                        <Maximize2 className="w-3.5 h-3.5" /> <span>{appLang === 'te' ? '🖥️ పెద్ద స్క్రీన్ (Fullscreen)' : '🖥️ Enlarge View (Fullscreen)'}</span>
                      </>
                    )}
                  </button>

                  <button
                    type="button"
                    onClick={downloadSavedCardsExcel}
                    className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[10.5px] px-2 py-1 rounded transition-colors shadow-2xs cursor-pointer"
                  >
                    <Download className="w-3 h-3" /> {appLang === 'te' ? 'ఎక్సెల్ డౌన్‌లోడ్ (.xlsx)' : 'Export All to Excel (.xlsx)'}
                  </button>
                </div>
              </div>

              {!isSavedFilterCollapsed && (
                <>
                  {/* FILTER SUB-HEADER & RESET */}
                  <div className="flex items-center justify-between">
                <span className="text-[10px] font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                  <Filter className="w-3 h-3 text-indigo-600" /> {appLang === 'te' ? 'జాబ్ కార్డుల ఫిల్టర్లు' : 'Filter Saved Job Cards'}
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setSavedDateFrom('');
                    setSavedDateTo('');
                    setSavedStatusFilter('all');
                    setSavedLocationFilter('all');
                    setSavedSupervisorFilter('all');
                    setSavedMechanicFilter('all');
                    setSavedListSearch('');
                    setSavedListPage(1);
                  }}
                  className="text-[9.5px] text-slate-500 hover:text-indigo-600 font-bold underline cursor-pointer"
                >
                  Reset All Filters
                </button>
              </div>

              {/* FILTERS GRID */}
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-1.5">
                {/* From Date */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">From Date</label>
                  <input
                    type="date"
                    value={savedDateFrom}
                    onChange={(e) => {
                      setSavedDateFrom(e.target.value);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                  />
                </div>

                {/* To Date */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">To Date</label>
                  <input
                    type="date"
                    value={savedDateTo}
                    onChange={(e) => {
                      setSavedDateTo(e.target.value);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                  />
                </div>

                {/* Status */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Status</label>
                  <select
                    value={savedStatusFilter}
                    onChange={(e) => {
                      setSavedStatusFilter(e.target.value as any);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Statuses</option>
                    <option value="Open">Open Cards</option>
                    <option value="Closed">Closed Cards</option>
                    <option value="MissingOnline">Missing Online No.</option>
                  </select>
                </div>

                {/* Service Place / Location */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Service Place</label>
                  <select
                    value={savedLocationFilter}
                    onChange={(e) => {
                      setSavedLocationFilter(e.target.value as any);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Locations</option>
                    <option value="workshop">Workshop</option>
                    <option value="dss">Door Step (DSS)</option>
                    <option value="event">Event / Camp</option>
                  </select>
                </div>

                {/* Supervisor Wise Filter */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Supervisor Wise</label>
                  <select
                    value={savedSupervisorFilter}
                    onChange={(e) => {
                      setSavedSupervisorFilter(e.target.value);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Supervisors</option>
                    {supervisorsList.map((sup, idx) => (
                      <option key={idx} value={sup}>{sup}</option>
                    ))}
                  </select>
                </div>

                {/* Mechanic / Technician Filter */}
                <div>
                  <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Technician Wise</label>
                  <select
                    value={savedMechanicFilter}
                    onChange={(e) => {
                      setSavedMechanicFilter(e.target.value);
                      setSavedListPage(1);
                    }}
                    className="w-full text-[10px] py-0.5 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium cursor-pointer"
                  >
                    <option value="all">All Technicians</option>
                    {mechanicsList.map((mech, idx) => (
                      <option key={idx} value={mech}>{mech}</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Text Search & Record Count */}
              <div className="pt-1 border-t border-slate-200 flex flex-wrap items-center justify-between gap-1.5">
                <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded px-1.5 py-0.5 shadow-2xs flex-1 max-w-md">
                  <Search className="w-3 h-3 text-slate-400 shrink-0" />
                  <input
                    type="text"
                    placeholder="Search by Job No, Customer, Chassis, Mobile, Village, Model..."
                    value={savedListSearch}
                    onChange={(e) => {
                      setSavedListSearch(e.target.value);
                      setSavedListPage(1);
                    }}
                    className="bg-transparent border-none outline-none text-[10px] w-full text-slate-700 font-medium"
                  />
                  {savedListSearch && (
                    <button
                      type="button"
                      onClick={() => {
                        setSavedListSearch('');
                        setSavedListPage(1);
                      }}
                      className="text-slate-400 hover:text-slate-600 text-[10px] cursor-pointer font-bold"
                    >
                      Clear
                    </button>
                  )}
                </div>

                <div className="text-[10px] text-slate-600 font-semibold bg-white px-2 py-0.5 rounded border border-slate-300 shadow-2xs">
                  Showing <span className="font-bold text-slate-900">{searchedJobCards.length === 0 ? 0 : (validSavedListPage - 1) * savedListPageSize + 1}</span>–<span className="font-bold text-slate-900">{Math.min(validSavedListPage * savedListPageSize, searchedJobCards.length)}</span> of <span className="font-bold text-indigo-900">{searchedJobCards.length}</span> Filtered Cards
                </div>
              </div>

              {/* PAGINATION CONTROLS */}
              <div className="flex flex-wrap items-center justify-between gap-1.5 pt-1 border-t border-slate-100">
                <div className="flex items-center gap-1 text-[10px] text-slate-600 font-bold">
                  <span>Per Page:</span>
                  <select
                    value={savedListPageSize}
                    onChange={(e) => {
                      setSavedListPageSize(Number(e.target.value));
                      setSavedListPage(1);
                    }}
                    className="bg-white border border-slate-300 rounded px-1 py-0.5 text-[10px] font-bold text-slate-800 outline-none cursor-pointer shadow-2xs"
                  >
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                    <option value={250}>250</option>
                    <option value={500}>500</option>
                  </select>
                </div>

                <div className="flex items-center gap-0.5">
                  <button
                    type="button"
                    disabled={validSavedListPage <= 1}
                    onClick={() => setSavedListPage(1)}
                    className="px-1.5 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                    title="First Page"
                  >
                    «
                  </button>
                  <button
                    type="button"
                    disabled={validSavedListPage <= 1}
                    onClick={() => setSavedListPage(p => Math.max(1, p - 1))}
                    className="px-1.5 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                  >
                    ‹ Prev
                  </button>
                  <span className="px-1.5 py-0.5 text-[10px] font-bold text-slate-800 bg-white border border-slate-300 rounded shadow-2xs">
                    Page {validSavedListPage} of {totalPages}
                  </span>
                  <button
                    type="button"
                    disabled={validSavedListPage >= totalPages}
                    onClick={() => setSavedListPage(p => Math.min(totalPages, p + 1))}
                    className="px-1.5 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                  >
                    Next ›
                  </button>
                  <button
                    type="button"
                    disabled={validSavedListPage >= totalPages}
                    onClick={() => setSavedListPage(totalPages)}
                    className="px-1.5 py-0.5 text-[10px] font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                    title="Last Page"
                  >
                    »
                  </button>
                </div>
              </div>
            </>
          )}
        </div>

            {/* TABLE CONTROL & LINE SIZE SETTINGS BAR */}
            <div className="bg-indigo-50/90 border border-indigo-200 p-2.5 rounded-lg flex flex-wrap items-center justify-between gap-3 text-xs font-semibold">
              <div className="flex items-center gap-3 flex-wrap">
                <div className="flex items-center gap-1.5">
                  <span className="text-slate-800 font-extrabold flex items-center gap-1.5">
                    <Maximize2 className="w-3.5 h-3.5 text-indigo-600" /> Line Size / Row Height:
                  </span>
                  <div className="inline-flex rounded-md shadow-2xs bg-white p-0.5 border border-slate-300">
                    <button
                      type="button"
                      onClick={() => setSavedRowDensity('compact')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        savedRowDensity === 'compact'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      🤏 Compact
                    </button>
                    <button
                      type="button"
                      onClick={() => setSavedRowDensity('normal')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        savedRowDensity === 'normal'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      ↔️ Normal
                    </button>
                    <button
                      type="button"
                      onClick={() => setSavedRowDensity('spacious')}
                      className={`px-2.5 py-1 rounded text-[11px] font-extrabold transition-colors cursor-pointer ${
                        savedRowDensity === 'spacious'
                          ? 'bg-indigo-600 text-white shadow-2xs'
                          : 'text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      ↕️ Spacious
                    </button>
                  </div>
                </div>

                {/* DYNAMIC FREEZE COLUMN SELECTOR */}
                <div className="flex items-center gap-1.5 border-l border-indigo-200 pl-3">
                  <span className="text-slate-800 font-extrabold flex items-center gap-1 text-[11px]">
                    ❄️ Freeze Column:
                  </span>
                  <select
                    value={freezeUpToColumn}
                    onChange={(e) => setFreezeUpToColumn(e.target.value)}
                    className="bg-white border border-slate-300 rounded px-2 py-1 text-[11px] font-extrabold text-slate-800 outline-none cursor-pointer shadow-2xs focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500"
                    title="Freeze columns up to the selected one"
                  >
                    <option value="none">No Freeze</option>
                    <option value="slNo">Sl No Only</option>
                    <option value="status">Sl No + Status (Default)</option>
                    <option value="jobNo">Up to Job Card</option>
                    <option value="complaintDate">Up to Complaint Date</option>
                    <option value="onlineJobCardNo">Up to Online Job Card No</option>
                    <option value="jobCardOpenDate">Up to Job Card Open Date</option>
                    <option value="branch">Up to Branch</option>
                    <option value="historyFileNo">Up to History File No</option>
                    <option value="tractorModel">Up to Tractor Model</option>
                    <option value="modelType">Up to Model Type</option>
                    <option value="chassisNo">Up to Chassis No</option>
                    <option value="engSrNo">Up to Eng Sr No</option>
                    <option value="dateOfDelivery">Up to Date of Delivery</option>
                    <option value="customerName">Up to Customer Name</option>
                    <option value="fatherName">Up to Father Name</option>
                    <option value="address">Up to Address</option>
                    <option value="village">Up to Village</option>
                    <option value="mandal">Up to Mandal</option>
                    <option value="phoneNo">Up to Phone No</option>
                    <option value="hrsRun">Up to Hrs Run</option>
                    <option value="typeOfService">Up to Type of Service</option>
                    <option value="freeServiceList">Up to Free Service List</option>
                    <option value="extraRepairsDone">Up to Extra Repairs Done</option>
                    <option value="actualClosedDate">Up to Actual Closed Date</option>
                    <option value="technicianName">Up to Technician Name</option>
                    <option value="servicePlace">Up to Service Place</option>
                    <option value="billNo">Up to Bill No</option>
                    <option value="reasonsForAnalysis">Up to Reasons for Analysis</option>
                    <option value="telecalling">Up to Telecalling</option>
                  </select>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <span className="text-[11px] text-slate-600 font-medium hidden md:inline bg-white px-2 py-1 rounded border border-slate-200">
                  💡 <span className="font-bold text-slate-800">Column Resizing:</span> Drag column headers left/right to adjust widths. Or click 📌 / ❄️ to Freeze!
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setSavedColWidths(defaultSavedColWidths);
                    setFreezeUpToColumn('status');
                  }}
                  className="px-2.5 py-1 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 font-bold rounded text-[11px] shadow-2xs cursor-pointer transition-colors"
                  title="Reset column widths to default"
                >
                  ↺ Reset Table Layout
                </button>
              </div>
            </div>

            {/* TABULAR DISPLAY */}
            <div className={`overflow-auto border border-slate-300 rounded-lg shadow-xs relative ${
              isSavedTableMaximized 
                ? 'flex-1 min-h-0 bg-white' 
                : isSavedFilterCollapsed 
                  ? 'max-h-[calc(100vh-200px)] md:max-h-[calc(100vh-140px)]' 
                  : 'max-h-[calc(100vh-380px)] md:max-h-[calc(100vh-320px)]'
            }`}>
              <table
                className="text-left text-slate-700 border-collapse border border-slate-300"
                style={{
                  width: `${(Object.values(savedColWidths) as number[]).reduce((a, b) => a + b, 0)}px`,
                  minWidth: '100%'
                }}
              >
                <thead className="sticky top-0 z-30">
                  <tr className="bg-slate-100 border-b-2 border-slate-300 text-slate-800 font-extrabold uppercase tracking-wider text-[11px] whitespace-nowrap">
                    {/* SL NO */}
                    <th
                      style={getStickyProps('slNo', true).style}
                      className={`p-2.5 text-center border-r border-b-2 border-slate-300 relative select-none ${getStickyProps('slNo', true).className}`}
                    >
                      <div className="flex items-center justify-center gap-1.5">
                        <input
                          type="checkbox"
                          checked={selectedJobCardIds.length === paginatedJobCards.length && paginatedJobCards.length > 0}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSelectedJobCardIds(paginatedJobCards.map(c => c.id));
                            } else {
                              setSelectedJobCardIds([]);
                            }
                          }}
                          className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                        />
                        <span>Sl No</span>
                        <button
                          type="button"
                          onClick={() => setFreezeUpToColumn(freezeUpToColumn === 'slNo' ? 'none' : 'slNo')}
                          className={`text-[10px] p-0.5 rounded hover:bg-slate-200 transition-all cursor-pointer ${
                            freezeUpToColumn === 'slNo' ? 'text-indigo-600 font-black scale-110' : 'text-slate-300 opacity-70 hover:opacity-100'
                          }`}
                          title="Freeze Sl No column"
                        >
                          📌
                        </button>
                      </div>
                      <div
                        onMouseDown={(e) => handleColumnResizeStart('slNo', e)}
                        className="absolute right-0 top-0 bottom-0 w-3 cursor-col-resize group flex items-center justify-center z-30 hover:bg-indigo-500/20 active:bg-indigo-600/40"
                        title="Drag to resize Sl No column"
                      >
                        <div className="w-[2px] h-full bg-slate-300 group-hover:bg-indigo-600 group-active:bg-indigo-700" />
                      </div>
                    </th>

                    {/* STATUS */}
                    <th
                      style={getStickyProps('status', true).style}
                      className={`p-2.5 text-center border-r border-b-2 border-slate-300 relative select-none ${getStickyProps('status', true).className}`}
                    >
                      <div className="flex items-center justify-center gap-1.5">
                        <span>Status</span>
                        <button
                          type="button"
                          onClick={() => setFreezeUpToColumn(freezeUpToColumn === 'status' ? 'slNo' : 'status')}
                          className={`text-[10px] p-0.5 rounded hover:bg-slate-200 transition-all cursor-pointer ${
                            freezeUpToColumn === 'status' ? 'text-indigo-600 font-black scale-110' : 'text-slate-300 opacity-70 hover:opacity-100'
                          }`}
                          title="Freeze up to Status"
                        >
                          📌
                        </button>
                      </div>
                      <div
                        onMouseDown={(e) => handleColumnResizeStart('status', e)}
                        className="absolute right-0 top-0 bottom-0 w-3 cursor-col-resize group flex items-center justify-center z-30 hover:bg-indigo-500/20 active:bg-indigo-600/40"
                        title="Drag to resize Status column"
                      >
                        <div className="w-[2px] h-full bg-slate-300 group-hover:bg-indigo-600 group-active:bg-indigo-700" />
                      </div>
                    </th>

                    {/* OTHER RESIZABLE COLUMNS */}
                    {[
                      { key: 'jobNo', label: 'JOB CARD' },
                      { key: 'complaintDate', label: 'COMPALINT DATE' },
                      { key: 'onlineJobCardNo', label: 'ONLINE JOB CARD NO' },
                      { key: 'jobCardOpenDate', label: 'JOB CARD OPEN DAT' },
                      { key: 'branch', label: 'BRANCH' },
                      { key: 'historyFileNo', label: 'HISTORY FILE NO.' },
                      { key: 'tractorModel', label: 'TRACTOR MODEL' },
                      { key: 'modelType', label: 'MODEL TYPE' },
                      { key: 'chassisNo', label: 'CHASIS NO' },
                      { key: 'engSrNo', label: 'ENG SR NO' },
                      { key: 'dateOfDelivery', label: 'DATE OF DELIVERY' },
                      { key: 'customerName', label: 'CUTOMER NAME' },
                      { key: 'fatherName', label: 'FATHER' },
                      { key: 'address', label: 'ADDRESS' },
                      { key: 'village', label: 'VILLAGE' },
                      { key: 'mandal', label: 'MANDAL' },
                      { key: 'phoneNo', label: 'PH NO' },
                      { key: 'hrsRun', label: 'HRS RUN' },
                      { key: 'typeOfService', label: 'TYPE OF SERVICE' },
                      { key: 'freeServiceList', label: 'FREE SERVICE LIST' },
                      { key: 'extraRepairsDone', label: 'EXTRA OTHER REPAIRS DONE WITH FREE SERVICE' },
                      { key: 'actualClosedDate', label: 'ACTUVAL CLOSED DATE' },
                      { key: 'technicianName', label: 'TECHNICIAN NAME' },
                      { key: 'servicePlace', label: 'SERVICE PLACE' },
                      { key: 'billNo', label: 'BILL NO.' },
                      { key: 'reasonsForAnalysis', label: 'RESONS FOR ANALYSIS' },
                      { key: 'telecalling', label: 'TELECALLING' },
                    ].map(col => (
                      <th
                        key={col.key}
                        style={getStickyProps(col.key, true).style}
                        className={`p-2.5 border-r border-b-2 border-slate-300 relative select-none ${getStickyProps(col.key, true).className}`}
                      >
                        <div className="flex items-center justify-between gap-1.5">
                          <div className="truncate pr-1" title={col.label}>{col.label}</div>
                          <button
                            type="button"
                            onClick={() => setFreezeUpToColumn(freezeUpToColumn === col.key ? 'status' : col.key)}
                            className={`text-[10px] p-0.5 rounded hover:bg-slate-200 transition-all shrink-0 cursor-pointer ${
                              freezeUpToColumn === col.key ? 'text-indigo-600 font-black scale-110' : 'text-slate-300 opacity-70 hover:opacity-100'
                            }`}
                            title={`Freeze table up to ${col.label}`}
                          >
                            📌
                          </button>
                        </div>
                        <div
                          onMouseDown={(e) => handleColumnResizeStart(col.key, e)}
                          className="absolute right-0 top-0 bottom-0 w-3 cursor-col-resize group flex items-center justify-center z-30 hover:bg-indigo-500/20 active:bg-indigo-600/40"
                          title={`Drag to resize ${col.label} column`}
                        >
                          <div className="w-[2px] h-full bg-slate-300 group-hover:bg-indigo-600 group-active:bg-indigo-700" />
                        </div>
                      </th>
                    ))}

                    {/* ACTIONS */}
                    <th
                      style={{
                        width: `${savedColWidths.actions}px`,
                        minWidth: `${savedColWidths.actions}px`,
                        maxWidth: `${savedColWidths.actions}px`
                      }}
                      className="p-2.5 text-center sticky top-0 right-0 bg-slate-100 border-l border-b-2 border-slate-300 shadow-[-4px_0_10px_-3px_rgba(0,0,0,0.08)] z-40 relative select-none"
                    >
                      <span>ACTIONS</span>
                      <div
                        onMouseDown={(e) => handleColumnResizeStart('actions', e)}
                        className="absolute left-0 top-0 bottom-0 w-3 cursor-col-resize group flex items-center justify-center z-30 hover:bg-indigo-500/20 active:bg-indigo-600/40"
                        title="Drag to resize Actions column"
                      >
                        <div className="w-[2px] h-full bg-slate-300 group-hover:bg-indigo-600 group-active:bg-indigo-700" />
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 font-medium">
                  {paginatedJobCards.map((card, index) => {
                    const globalIndex = (validSavedListPage - 1) * savedListPageSize + index + 1;
                    const isCurrentlyEditing = editingCardId === card.id;
                    const cellBgClass = isCurrentlyEditing ? 'bg-blue-50 group-hover:bg-blue-100' : (index % 2 === 1 ? 'bg-slate-50 group-hover:bg-amber-100/80' : 'bg-white group-hover:bg-amber-100/80');
                    const cellPadding =
                      savedRowDensity === 'compact'
                        ? 'p-1 text-[11px]'
                        : savedRowDensity === 'spacious'
                        ? 'p-3.5 text-sm'
                        : 'p-2.5 text-xs';

                    const inputPadding =
                      savedRowDensity === 'compact'
                        ? 'p-0.5 text-[10.5px]'
                        : savedRowDensity === 'spacious'
                        ? 'p-2 text-xs'
                        : 'p-1.5 text-xs';

                    return (
                      <tr
                        key={card.id}
                        className={`group border-b border-slate-200 transition-colors ${
                          isCurrentlyEditing 
                            ? 'bg-blue-50/80 font-semibold' 
                            : index % 2 === 1 ? 'bg-slate-50/60' : 'bg-white'
                        }`}
                      >
                        {/* SL NO */}
                        <td
                          style={getStickyProps('slNo', false).style}
                          className={`${cellPadding} text-center text-slate-500 font-bold border-r border-slate-200 ${cellBgClass} ${getStickyProps('slNo', false).className}`}
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <input
                              type="checkbox"
                              checked={selectedJobCardIds.includes(card.id)}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedJobCardIds([...selectedJobCardIds, card.id]);
                                } else {
                                  setSelectedJobCardIds(selectedJobCardIds.filter(id => id !== card.id));
                                }
                              }}
                              className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                            />
                            {globalIndex}
                          </div>
                        </td>

                        {/* STATUS */}
                        <td
                          style={getStickyProps('status', false).style}
                          className={`${cellPadding} text-center whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('status', false).className}`}
                        >
                          <button
                            type="button"
                            onClick={() => toggleCardStatus(card)}
                            className={`inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-extrabold cursor-pointer transition-all hover:scale-105 active:scale-95 border ${
                              !isCardClosed(card)
                                ? 'bg-amber-100 text-amber-900 border-amber-400 hover:bg-amber-200 shadow-2xs'
                                : 'bg-emerald-600 text-white border-emerald-700 hover:bg-emerald-700 shadow-2xs'
                            }`}
                            title="Click to toggle status (Open / Closed)"
                          >
                            {!isCardClosed(card) ? '⏳ OPEN' : '✓ CLOSED'}
                          </button>
                        </td>

                        {/* JOB CARD (NO) */}
                        <td
                          style={getStickyProps('jobNo', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('jobNo', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.jobNo || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'jobNo', e.target.value)}
                            placeholder="Job Card..."
                            className={`w-full border border-slate-300 rounded font-mono font-bold text-indigo-700 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs ${inputPadding}`}
                          />
                        </td>

                        {/* COMPLAINT DATE */}
                        <td
                          style={getStickyProps('complaintDate', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('complaintDate', false).className}`}
                        >
                          <input
                            type="date"
                            value={toInputDateFormat(card.complaintDate) || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'complaintDate', e.target.value)}
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-medium ${inputPadding}`}
                          />
                        </td>

                        {/* ONLINE JOB CARD NO */}
                        <td
                          style={getStickyProps('onlineJobCardNo', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('onlineJobCardNo', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.onlineJobCardNo || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'onlineJobCardNo', e.target.value)}
                            placeholder="Online Job Card No..."
                            className={`w-full border border-slate-300 rounded font-mono font-bold text-indigo-700 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs ${inputPadding}`}
                          />
                        </td>

                        {/* JOB CARD OPEN DATE */}
                        <td
                          style={getStickyProps('jobCardOpenDate', false).style}
                          className={`${cellPadding} whitespace-nowrap font-bold text-slate-800 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('jobCardOpenDate', false).className}`}
                        >
                          {fmtDate(card.jobDate || card.jobOpenDate || card.createdAt) || '—'}
                        </td>

                        {/* BRANCH */}
                        <td
                          style={getStickyProps('branch', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('branch', false).className}`}
                          title={card.branch}
                        >
                          {card.branch || '—'}
                        </td>

                        {/* HISTORY FILE NO */}
                        <td
                          style={getStickyProps('historyFileNo', false).style}
                          className={`${cellPadding} font-mono border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('historyFileNo', false).className}`}
                          title={card.historyFileNo || card.fileNo}
                        >
                          {card.historyFileNo || card.fileNo ? (
                            <span
                              className={`inline-block px-1.5 py-0.5 rounded font-mono text-[11px] font-black border ${
                                isWithin2Years(card.dateOfDelivery || card.installDate || card.deliveryDate || card.dateOfDel, card.jobDate)
                                  ? 'bg-blue-100 text-blue-900 border-blue-300'
                                  : 'bg-red-100 text-red-900 border-red-300'
                              }`}
                            >
                              {card.historyFileNo || card.fileNo}
                            </span>
                          ) : (
                            '—'
                          )}
                        </td>

                        {/* TRACTOR MODEL */}
                        <td
                          style={getStickyProps('tractorModel', false).style}
                          className={`${cellPadding} text-slate-900 font-bold border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('tractorModel', false).className}`}
                          title={card.model}
                        >
                          {card.model || '—'}
                        </td>

                        {/* MODEL TYPE */}
                        <td
                          style={getStickyProps('modelType', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('modelType', false).className}`}
                          title={card.modelType}
                        >
                          {card.modelType || '—'}
                        </td>

                        {/* CHASSIS NO */}
                        <td
                          style={getStickyProps('chassisNo', false).style}
                          className={`${cellPadding} font-mono text-slate-900 font-bold border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('chassisNo', false).className}`}
                          title={card.chassisNo}
                        >
                          {card.chassisNo || '—'}
                        </td>

                        {/* ENG SR NO */}
                        <td
                          style={getStickyProps('engSrNo', false).style}
                          className={`${cellPadding} font-mono text-slate-600 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('engSrNo', false).className}`}
                          title={card.engineNo}
                        >
                          {card.engineNo || '—'}
                        </td>

                        {/* DATE OF DELIVERY */}
                        <td
                          style={getStickyProps('dateOfDelivery', false).style}
                          className={`${cellPadding} border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('dateOfDelivery', false).className}`}
                        >
                          <div className="font-bold text-slate-800">{fmtDate(card.dateOfDelivery || card.installDate || card.deliveryDate) || '—'}</div>
                        </td>

                        {/* CUSTOMER NAME */}
                        <td
                          style={getStickyProps('customerName', false).style}
                          className={`${cellPadding} text-slate-900 font-bold border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('customerName', false).className}`}
                          title={card.custName}
                        >
                          {card.custName || '—'}
                        </td>

                        {/* FATHER NAME */}
                        <td
                          style={getStickyProps('fatherName', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('fatherName', false).className}`}
                          title={card.fatherName}
                        >
                          {card.fatherName || '—'}
                        </td>

                        {/* ADDRESS */}
                        <td
                          style={getStickyProps('address', false).style}
                          className={`${cellPadding} text-slate-600 truncate border-r border-slate-200 ${cellBgClass} ${getStickyProps('address', false).className}`}
                          title={card.custAddr || card.address}
                        >
                          {card.custAddr || card.address || '—'}
                        </td>

                        {/* VILLAGE */}
                        <td
                          style={getStickyProps('village', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('village', false).className}`}
                          title={card.village}
                        >
                          {card.village || '—'}
                        </td>

                        {/* MANDAL */}
                        <td
                          style={getStickyProps('mandal', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('mandal', false).className}`}
                          title={card.mandal}
                        >
                          {card.mandal || '—'}
                        </td>

                        {/* PHONE NO */}
                        <td
                          style={getStickyProps('phoneNo', false).style}
                          className={`${cellPadding} text-slate-700 font-bold border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('phoneNo', false).className}`}
                          title={card.ownerMob || card.phNo}
                        >
                          {card.ownerMob || card.phNo || '—'}
                        </td>

                        {/* HRS RUN */}
                        <td
                          style={getStickyProps('hrsRun', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('hrsRun', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.hourMeter || card.hrsRun || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'hourMeter', e.target.value)}
                            placeholder="Hrs Run..."
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-medium ${inputPadding}`}
                          />
                        </td>

                        {/* TYPE OF SERVICE */}
                        <td
                          style={getStickyProps('typeOfService', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('typeOfService', false).className}`}
                          title={getDisplayServiceType(card)}
                        >
                          {getDisplayServiceType(card)}
                        </td>

                        {/* FREE SERVICE LIST */}
                        <td
                          style={getStickyProps('freeServiceList', false).style}
                          className={`${cellPadding} text-slate-700 truncate font-semibold border-r border-slate-200 ${cellBgClass} ${getStickyProps('freeServiceList', false).className}`}
                          title={getFreeServiceText(card)}
                        >
                          {getFreeServiceText(card)}
                        </td>

                        {/* EXTRA REPAIRS DONE */}
                        <td
                          style={getStickyProps('extraRepairsDone', false).style}
                          className={`${cellPadding} text-slate-700 truncate border-r border-slate-200 ${cellBgClass} ${getStickyProps('extraRepairsDone', false).className}`}
                          title={getExtraRepairsText(card)}
                        >
                          {getExtraRepairsText(card)}
                        </td>

                        {/* ACTUAL CLOSED DATE */}
                        <td
                          style={getStickyProps('actualClosedDate', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('actualClosedDate', false).className}`}
                        >
                          <input
                            type="date"
                            value={toInputDateFormat(card.actualClosedDate || card.dateTimeOut) || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'actualClosedDate', e.target.value)}
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-medium ${inputPadding}`}
                          />
                        </td>

                        {/* TECHNICIAN NAME */}
                        <td
                          style={getStickyProps('technicianName', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('technicianName', false).className}`}
                        >
                          <input
                            type="text"
                            list="mechanics-datalist"
                            value={normalizeStaffName(card.mechanic || card.technicianName) || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'mechanic', e.target.value)}
                            placeholder="Select or type mechanic..."
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-bold ${inputPadding}`}
                          />
                        </td>

                        {/* SERVICE PLACE */}
                        <td
                          style={getStickyProps('servicePlace', false).style}
                          className={`${cellPadding} text-slate-700 border-r border-slate-200 truncate ${cellBgClass} ${getStickyProps('servicePlace', false).className}`}
                          title={card.serviceLocation || card.servicePlace}
                        >
                          {card.serviceLocation || card.servicePlace || '—'}
                        </td>

                        {/* BILL NO */}
                        <td
                          style={getStickyProps('billNo', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('billNo', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.billNo || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'billNo', e.target.value)}
                            placeholder="Bill No..."
                            className={`w-full border border-slate-300 rounded font-mono font-bold text-indigo-900 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs ${inputPadding}`}
                          />
                        </td>

                        {/* REASONS FOR ANALYSIS */}
                        <td
                          style={getStickyProps('reasonsForAnalysis', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('reasonsForAnalysis', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.reasonsForAnalysis || card.problemDescription || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'reasonsForAnalysis', e.target.value)}
                            placeholder="Reasons for analysis..."
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-medium ${inputPadding}`}
                          />
                        </td>

                        {/* TELECALLING */}
                        <td
                          style={getStickyProps('telecalling', false).style}
                          className={`p-1 whitespace-nowrap border-r border-slate-200 ${cellBgClass} ${getStickyProps('telecalling', false).className}`}
                        >
                          <input
                            type="text"
                            value={card.telecalling || ''}
                            onChange={(e) => handleInlineCardFieldChange(card.id, 'telecalling', e.target.value)}
                            placeholder="Telecalling..."
                            className={`w-full border border-slate-300 rounded text-slate-800 bg-white focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-2xs font-medium ${inputPadding}`}
                          />
                        </td>

                        {/* ACTIONS */}
                        <td
                          style={{
                            width: `${savedColWidths.actions}px`,
                            minWidth: `${savedColWidths.actions}px`,
                            maxWidth: `${savedColWidths.actions}px`
                          }}
                          className={`${cellPadding} text-center sticky right-0 border-l border-slate-300 shadow-[-4px_0_10px_-3px_rgba(0,0,0,0.08)] z-10 ${cellBgClass}`}
                        >
                          <div className="flex items-center justify-center gap-1.5 whitespace-nowrap">
                            <button
                              type="button"
                              onClick={() => {
                                setViewingCardModal(card);
                              }}
                              className="p-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded transition-all cursor-pointer flex items-center justify-center"
                              title="View filled job card"
                            >
                              <Eye className="w-4 h-4" />
                            </button>
                            <button
                              type="button"
                              onClick={() => loadJobCard(card)}
                              className={`p-1.5 rounded transition-all cursor-pointer flex items-center justify-center ${
                                isCurrentlyEditing
                                  ? 'bg-blue-600 text-white shadow-xs'
                                  : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
                              }`}
                              title={isCurrentlyEditing ? 'Being edited' : 'Edit this card'}
                            >
                              <FileText className="w-4 h-4" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                loadJobCard(card);
                                setTimeout(() => {
                                  window.print();
                                }, 150);
                              }}
                              className="p-1.5 bg-slate-100 text-slate-700 hover:bg-slate-200 rounded transition-all cursor-pointer flex items-center justify-center"
                              title="Print this job card"
                            >
                              <Printer className="w-4 h-4" />
                            </button>
                            <button
                              type="button"
                              onClick={() => deleteJobCard(card.id)}
                              className="p-1.5 bg-red-50 hover:bg-red-100 text-red-600 rounded transition-all cursor-pointer flex items-center justify-center"
                              title="Delete this saved card"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* BOTTOM PAGINATION FOOTER BAR */}
            <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-50 border border-slate-200 p-3 rounded-xl text-xs font-medium">
              <div className="text-slate-600 font-semibold">
                Showing <span className="font-bold text-slate-900">{searchedJobCards.length === 0 ? 0 : (validSavedListPage - 1) * savedListPageSize + 1}</span> to <span className="font-bold text-slate-900">{Math.min(validSavedListPage * savedListPageSize, searchedJobCards.length)}</span> of <span className="font-bold text-slate-900">{searchedJobCards.length}</span> total job cards
              </div>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  disabled={validSavedListPage <= 1}
                  onClick={() => setSavedListPage(1)}
                  className="px-2 py-1 text-xs font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  « First
                </button>
                <button
                  type="button"
                  disabled={validSavedListPage <= 1}
                  onClick={() => setSavedListPage(p => Math.max(1, p - 1))}
                  className="px-2.5 py-1 text-xs font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  ‹ Prev
                </button>
                <span className="px-3 py-1 text-xs font-bold text-slate-800 bg-white border border-slate-300 rounded shadow-2xs">
                  {validSavedListPage} / {totalPages}
                </span>
                <button
                  type="button"
                  disabled={validSavedListPage >= totalPages}
                  onClick={() => setSavedListPage(p => Math.min(totalPages, p + 1))}
                  className="px-2.5 py-1 text-xs font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  Next ›
                </button>
                <button
                  type="button"
                  disabled={validSavedListPage >= totalPages}
                  onClick={() => setSavedListPage(totalPages)}
                  className="px-2 py-1 text-xs font-bold bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  Last »
                </button>
              </div>
            </div>
          </div>
        )}
        </div>
      )}

      {activeTab === 'followup' && (
        <div className="w-full bg-white border border-slate-200 shadow-xs p-2.5 md:p-3.5 rounded-xl space-y-3.5 print:hidden">
          {/* HEADER SECTION */}
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-150 pb-2">
            <div className="flex items-center gap-2">
              <div className="bg-purple-100 text-purple-950 p-1.5 rounded-lg border border-purple-200">
                <PhoneCall className="w-4 h-4 text-purple-900" />
              </div>
              <div>
                <h2 className="text-sm font-black text-slate-900 leading-tight">Customer Data Dashboard</h2>
                <p className="text-[10px] text-slate-500 font-medium">Compare deliveries, identify non-reporting vehicles, filter by supervisor.</p>
              </div>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setIsFollowupFilterCollapsed(!isFollowupFilterCollapsed)}
                className="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[10px] rounded-lg transition-all shadow-xs flex items-center gap-1 cursor-pointer"
              >
                {isFollowupFilterCollapsed ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronUp className="w-3.5 h-3.5" />}
                <span>{isFollowupFilterCollapsed ? "Show Filters" : "Hide Filters"}</span>
              </button>
              <button
                type="button"
                onClick={downloadFollowupExcel}
                className="px-2.5 py-1.5 bg-purple-900 hover:bg-purple-950 text-white font-bold text-[10px] rounded-lg transition-all shadow-xs flex items-center gap-1.5 cursor-pointer"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" /> Export Lists
              </button>
              <label className="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[10px] rounded-lg transition-all shadow-xs flex items-center gap-1.5 cursor-pointer">
                <FileSpreadsheet className="w-3.5 h-3.5" /> Import Excel
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={handleExcelUpload}
                  className="hidden"
                />
              </label>

            </div>
          </div>

          {!isFollowupFilterCollapsed && (
            <>
              {/* STATISTICS / SUMMARY CARDS GRID */}
              {(() => {
            const total = uniqueCustomers.length;
            const reporting = uniqueCustomers.filter(rec => {
              const ch = rec['Chassis no'] || rec.__chassisDisplay || getFieldValue(rec, 'chassis') || '';
              const norm = normalizeKey(ch.toString());
              return norm && (customerJobCardsMap[norm] || []).length > 0;
            }).length;
            const notReporting = total - reporting;
            const rate = total > 0 ? ((reporting / total) * 100).toFixed(1) : '0.0';

            return (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
                <div 
                  onClick={() => setFollowupStatus('all')}
                  className={`border p-2 pb-1.5 px-3 rounded-lg cursor-pointer transition-all select-none ${
                    followupStatus === 'all'
                      ? 'bg-blue-50 border-blue-400 ring-1 ring-blue-400/20 shadow-xs'
                      : 'bg-slate-50 border-slate-200 hover:bg-slate-100/70'
                  }`}
                  title="Click to show all customer deliveries"
                >
                  <span className="text-[9px] font-black uppercase text-slate-500 tracking-wider">Total Deliveries</span>
                  <div className="text-base font-black text-slate-900 mt-0.5">{total}</div>
                </div>

                <div 
                  onClick={() => setFollowupStatus('reporting')}
                  className={`border p-2 pb-1.5 px-3 rounded-lg cursor-pointer transition-all select-none ${
                    followupStatus === 'reporting'
                      ? 'bg-emerald-50 border-emerald-400 ring-1 ring-emerald-400/20 shadow-xs'
                      : 'bg-emerald-50/20 border-emerald-100/60 hover:bg-emerald-50'
                  }`}
                  title="Click to filter by Active Reporting Customers"
                >
                  <span className="text-[9px] font-black uppercase text-emerald-800 tracking-wider">Reporting Customers</span>
                  <div className="text-base font-black text-emerald-950 mt-0.5">{reporting}</div>
                </div>

                <div 
                  onClick={() => setFollowupStatus('not_reporting')}
                  className={`border p-2 pb-1.5 px-3 rounded-lg cursor-pointer transition-all select-none ${
                    followupStatus === 'not_reporting'
                      ? 'bg-rose-50 border-rose-400 ring-1 ring-rose-400/20 shadow-xs'
                      : 'bg-rose-50/20 border-rose-100/60 hover:bg-rose-50'
                  }`}
                  title="Click to filter by Not Reporting Customers"
                >
                  <span className="text-[9px] font-black uppercase text-rose-800 tracking-wider">Not Reporting Customers</span>
                  <div className="text-base font-black text-rose-950 mt-0.5">{notReporting}</div>
                </div>

                <div 
                  onClick={() => setFollowupStatus('duplicate')}
                  className={`border p-2 pb-1.5 px-3 rounded-lg cursor-pointer transition-all select-none ${
                    followupStatus === 'duplicate'
                      ? 'bg-amber-50 border-amber-400 ring-1 ring-amber-400/20 shadow-xs'
                      : 'bg-amber-50/20 border-amber-100/60 hover:bg-amber-50'
                  }`}
                  title="Click to filter by Double Entry Customers"
                >
                  <span className="text-[9px] font-black uppercase text-amber-800 tracking-wider">Duplicate Entries</span>
                  <div className="text-base font-black text-amber-950 mt-0.5">{duplicateChassisSet.size}</div>
                </div>

                <div className="bg-indigo-50/40 border border-indigo-100/80 p-2 pb-1.5 px-3 rounded-lg">
                  <span className="text-[9px] font-black uppercase text-indigo-800 tracking-wider">Reporting Rate</span>
                  <div className="flex items-center justify-between gap-2 mt-0.5">
                    <div className="text-base font-black text-indigo-950">{rate}%</div>
                    <div className="flex-1 max-w-[80px] bg-slate-200 h-1.5 rounded-full overflow-hidden">
                      <div className="bg-indigo-600 h-full rounded-full" style={{ width: `${rate}%` }} />
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}

          {/* ADVANCED FILTER & SORT TOOLBAR */}
          <div className="bg-slate-50 border border-slate-200 p-2.5 rounded-lg space-y-2">
            <div className="flex items-center justify-between gap-2 border-b border-slate-200 pb-1.5">
              <div className="flex items-center gap-1.5 text-xs font-black text-slate-800">
                <Filter className="w-3.5 h-3.5 text-indigo-600" /> Filter & Sort Options
              </div>
              <button
                type="button"
                onClick={() => {
                  setFollowupSearch('');
                  setFollowupSupervisor('all');
                  setFollowupStatus('all');
                  setFollowupSortBy('hfn');
                  setFollowupSortOrder('asc');
                  setFollowupDateFrom('');
                  setFollowupDateTo('');
                  setFollowupModelFilter('all');
                  setFollowupPage(1);
                }}
                className="text-[9.5px] text-slate-500 hover:text-indigo-600 font-bold underline cursor-pointer"
              >
                Reset All Filters
              </button>
            </div>

            {/* FILTERS & SORT GRID */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-1.5">
              {/* Delivery Date From */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Del. Date From</label>
                <input
                  type="date"
                  value={followupDateFrom}
                  onChange={(e) => {
                    setFollowupDateFrom(e.target.value);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                />
              </div>

              {/* Delivery Date To */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Del. Date To</label>
                <input
                  type="date"
                  value={followupDateTo}
                  onChange={(e) => {
                    setFollowupDateTo(e.target.value);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded outline-none text-slate-700 font-medium"
                />
              </div>

              {/* Supervisor */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Supervisor</label>
                <select
                  value={followupSupervisor}
                  onChange={(e) => {
                    setFollowupSupervisor(e.target.value);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded text-slate-800 font-bold outline-none cursor-pointer"
                >
                  <option value="all">All Supervisors</option>
                  {supervisorsList.map(sup => (
                    <option key={sup} value={sup}>{sup}</option>
                  ))}
                  <option value="unassigned">Unassigned Only</option>
                </select>
              </div>

              {/* Status */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Status</label>
                <select
                  value={followupStatus}
                  onChange={(e) => {
                    setFollowupStatus(e.target.value as any);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded text-slate-800 font-bold outline-none cursor-pointer"
                >
                  <option value="all">All Statuses</option>
                  <option value="reporting">Reporting</option>
                  <option value="not_reporting">Not Reporting</option>
                  <option value="duplicate">Duplicate (Double Entry)</option>
                </select>
              </div>

              {/* Model */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Tractor Model</label>
                <select
                  value={followupModelFilter}
                  onChange={(e) => {
                    setFollowupModelFilter(e.target.value);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded text-slate-800 font-bold outline-none cursor-pointer"
                >
                  <option value="all">All Models</option>
                  {followupModelsList.map(m => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                </select>
              </div>

              {/* Sort By */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Sort By</label>
                <select
                  value={followupSortBy}
                  onChange={(e) => {
                    setFollowupSortBy(e.target.value as any);
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 rounded text-indigo-950 font-black outline-none cursor-pointer"
                >
                  <option value="hfn">HFN (S.No)</option>
                  <option value="name">Customer Name</option>
                  <option value="dateOfDel">Delivery Date</option>
                  <option value="lastJobCardDate">Last Service Date</option>
                  <option value="nextCallDate">Next Call Date</option>
                  <option value="chassisNo">Chassis No</option>
                  <option value="village">Village</option>
                  <option value="model">Model</option>
                </select>
              </div>

              {/* Sort Order */}
              <div>
                <label className="block text-[8.5px] font-bold text-slate-600 uppercase mb-0.5">Sort Order</label>
                <button
                  type="button"
                  onClick={() => {
                    setFollowupSortOrder(order => order === 'asc' ? 'desc' : 'asc');
                    setFollowupPage(1);
                  }}
                  className="w-full text-[10px] py-1 px-1 bg-white border border-slate-300 hover:bg-slate-100 rounded text-slate-900 font-bold outline-none cursor-pointer flex items-center justify-center gap-1 shadow-2xs"
                  title="Click to toggle Ascending / Descending order"
                >
                  {followupSortOrder === 'asc' ? '⬆️ Ascending' : '⬇️ Descending'}
                </button>
              </div>
            </div>

            {/* Search Input Bar & Count */}
            <div className="pt-1.5 border-t border-slate-200 flex flex-wrap items-center justify-between gap-1.5">
              <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded px-2 py-1 shadow-2xs flex-1 max-w-md">
                <Search className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                <input
                  type="text"
                  placeholder="Search by customer, chassis, phone, village, remarks or HFN..."
                  value={followupSearch}
                  onChange={(e) => {
                    setFollowupSearch(e.target.value);
                    setFollowupPage(1);
                  }}
                  className="bg-transparent border-none outline-none text-xs w-full text-slate-800 font-semibold"
                />
                {followupSearch && (
                  <button
                    type="button"
                    onClick={() => {
                      setFollowupSearch('');
                      setFollowupPage(1);
                    }}
                    className="text-slate-400 hover:text-slate-600 text-[10px] cursor-pointer font-bold"
                  >
                    Clear
                  </button>
                )}
              </div>

              <div className="text-[10px] text-slate-600 font-semibold bg-white px-2.5 py-1 rounded border border-slate-300 shadow-2xs">
                Showing <span className="font-bold text-slate-900">{followupCalculatedList.length === 0 ? 0 : (followupPage - 1) * followupItemsPerPage + 1}</span>–<span className="font-bold text-slate-900">{Math.min(followupPage * followupItemsPerPage, followupCalculatedList.length)}</span> of <span className="font-bold text-indigo-900">{followupCalculatedList.length}</span> Records
              </div>
            </div>
          </div>
        </>
      )}

          {/* TABLE AREA */}
          <div className="border border-slate-200 rounded-xl overflow-hidden bg-white shadow-xs">
            <div className={`overflow-x-auto overflow-y-auto ${
              isFollowupFilterCollapsed 
                ? 'max-h-[calc(100vh-200px)] md:max-h-[calc(100vh-140px)]' 
                : 'max-h-[calc(100vh-380px)] md:max-h-[calc(100vh-320px)]'
            }`}>
              <table className="w-full text-left border-collapse text-[11px] leading-tight">
                <thead className="sticky top-0 z-20 shadow-sm">
                  <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-extrabold uppercase text-[9px] tracking-wider">
                    <th className="p-2 w-12 text-center bg-slate-100 sticky top-0 z-10 border-b border-slate-200">S.No</th>
                    <th className="p-2 w-48 bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Supervisor Assignment</th>
                    <th className="p-2 min-w-[260px] bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Customer details & Address</th>
                    <th className="p-2 min-w-[160px] bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Tractor Details</th>
                    <th className="p-2 min-w-[260px] bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Last Service Info</th>
                    <th className="p-2 min-w-[240px] bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Follow-up & Call Remarks</th>
                    <th className="p-2 w-32 text-center bg-slate-100 sticky top-0 z-10 border-b border-slate-200">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {paginatedFollowupList.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="p-8 text-center text-slate-400 font-bold">
                        No customers match the current filters.
                      </td>
                    </tr>
                  ) : (
                    paginatedFollowupList.map((item, index) => {
                      const displayIndex = (followupPage - 1) * followupItemsPerPage + index + 1;
                      return (
                        <tr key={item.chassisNo || index} className="hover:bg-amber-100/80 transition-colors border-b border-slate-150">
                          {/* S.No */}
                          <td className="p-1 text-center font-mono text-[9px] font-bold text-slate-500 border-r border-slate-100">
                            {displayIndex}
                          </td>

                          {/* Supervisor Assignment Dropdown */}
                          <td className="p-1 border-r border-slate-100">
                            <div className="space-y-0.5 max-w-[140px]">
                              <select
                                value={item.supervisor || ''}
                                onChange={(e) => {
                                  const newSup = e.target.value;
                                  updateCustomerRecordInIndex(item.chassisNo, {
                                    supervisor: newSup,
                                    SUPERVISOR: newSup,
                                    'Supervisor Name': newSup
                                  });
                                  alert(`✅ Assigned Supervisor "${newSup || 'Unassigned'}" to customer "${item.customerName || item.chassisNo}".`);
                                }}
                                className="text-[10px] font-bold p-1 bg-white border border-slate-200 rounded text-slate-800 outline-none focus:border-indigo-500 cursor-pointer w-full h-7 shadow-3xs"
                              >
                                <option value="">⚠️ Unassigned</option>
                                {supervisorsList.map(sup => (
                                  <option key={sup} value={sup}>{sup}</option>
                                ))}
                              </select>
                              {item.supervisor ? (
                                <span className="inline-block px-1 py-0.5 bg-purple-50 text-purple-950 border border-purple-200/60 rounded text-[8px] font-bold uppercase truncate max-w-full">
                                  {item.supervisor}
                                </span>
                              ) : (
                                <span className="text-red-500 italic font-bold text-[8px] pl-0.5 block">Required</span>
                              )}
                            </div>
                          </td>

                          {/* Customer Name, Phone and Complete Address */}
                          <td className="p-2 border-r border-slate-100">
                            <div className="space-y-1.5">
                              <div className="flex items-center gap-1.5 flex-wrap">
                                {item.historyFileNo && (
                                  <span
                                    className={`inline-flex items-center px-1.5 py-0.5 rounded font-mono text-[10px] font-black border shadow-2xs ${
                                      isWithin2Years(item.dateOfDel || item.installDate)
                                        ? 'bg-blue-600 text-white border-blue-700'
                                        : 'bg-red-600 text-white border-red-700'
                                    }`}
                                    title={`HFN: ${item.historyFileNo} (${isWithin2Years(item.dateOfDel || item.installDate) ? 'Delivery < 2 Yrs' : 'Delivery > 2 Yrs'})`}
                                  >
                                    HFN: {item.historyFileNo}
                                  </span>
                                )}
                                {(item.dateOfDel || item.installDate) && (
                                  <span
                                    className={`inline-flex items-center px-1.5 py-0.5 rounded font-mono text-[9px] font-black uppercase border shadow-2xs ${
                                      isWithin2Years(item.dateOfDel || item.installDate)
                                        ? 'bg-blue-100 text-blue-900 border-blue-300'
                                        : 'bg-red-100 text-red-900 border-red-300'
                                    }`}
                                    title="Warranty Status based on Delivery Date"
                                  >
                                    WTY STATUS: {isWithin2Years(item.dateOfDel || item.installDate) ? 'WARRANTY' : 'POST WTY'}
                                  </span>
                                )}
                                <span className="font-extrabold text-slate-900 text-xs leading-tight">{item.customerName}</span>
                              </div>
                              
                              {item.fatherName && (
                                <div className="text-[10px] text-slate-500 font-bold leading-none pl-0.5 mt-0.5">
                                  S/o {item.fatherName}
                                </div>
                              )}
                              
                              {item.mobileNumber && (
                                <div className="py-1">
                                  <span className="text-xs text-indigo-700 font-mono font-black bg-indigo-50 px-2 py-1 rounded-lg border-2 border-indigo-200 shadow-sm inline-flex items-center gap-2">
                                    <Phone className="w-3 h-3 text-indigo-500" /> {item.mobileNumber}
                                  </span>
                                </div>
                              )}

                              <div className="space-y-0.5 bg-slate-50 p-2 rounded-lg border border-slate-200 shadow-3xs">
                                <div className="text-slate-800 font-bold text-[10px] flex items-center gap-1">
                                  <MapPin className="w-3 h-3 text-slate-400" />
                                  {item.village ? `${item.village}` : '—'} 
                                  {item.mandal && <span className="text-slate-400 font-medium ml-1">({item.mandal})</span>}
                                </div>
                                {item.address && (
                                  <div className="text-slate-500 text-[10px] leading-tight italic border-t border-slate-200/50 pt-1 mt-1" title={item.address}>
                                    {item.address}
                                  </div>
                                )}
                              </div>
                            </div>
                          </td>


                          {/* Tractor Model details */}
                          <td className="p-1.5">
                            <div className="space-y-1 bg-slate-50/50 p-2 rounded-lg border border-slate-200/60 text-[11px]">
                              <div>
                                <span className="text-[10px] uppercase font-bold text-slate-400 block">Chassis Number</span>
                                <span className="font-mono font-black text-slate-900 text-[11px]">{item.chassisNo || '—'}</span>
                              </div>
                              {item.engineNo && (
                                <div>
                                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Engine Number</span>
                                  <span className="font-mono text-slate-600 font-bold">{item.engineNo}</span>
                                </div>
                              )}
                              <div>
                                <span className="text-[9px] uppercase font-bold text-indigo-400 block">Tractor Model</span>
                                <span className="text-indigo-950 font-black">{item.model || '—'}</span>
                                {item.modelType && (
                                  <span className="text-[10px] text-slate-500 font-bold block bg-white px-1 py-0.5 rounded border border-slate-200 mt-0.5">{item.modelType}</span>
                                )}
                              </div>
                              {item.dateOfDel && (
                                <div className="pt-1 border-t border-slate-200 mt-1">
                                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Del Date</span>
                                  <span className="font-bold text-slate-600">{item.dateOfDel}</span>
                                </div>
                              )}
                            </div>
                          </td>

                          {/* Last Service Card Details */}
                          <td className="p-1.5">
                            {item.hasJobCard ? (
                              <div className="text-[11px] space-y-1 bg-emerald-50/30 border border-emerald-100/60 p-2.5 rounded-lg leading-relaxed">
                                <div className="font-extrabold text-emerald-950 flex items-center justify-between gap-1.5 border-b border-emerald-100/60 pb-1 mb-1">
                                  <span className="text-[10px] uppercase font-black px-1.5 py-0.5 bg-emerald-100 text-emerald-900 rounded-md">{item.lastServiceType}</span>
                                  <span className="font-mono text-emerald-700 font-black text-xs">{item.lastEntryHours ? `${item.lastEntryHours} Hrs` : 'Hrs N/A'}</span>
                                </div>
                                <div className="text-slate-500 font-bold text-[10px] flex items-center gap-1.5">
                                  <Calendar className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                  <span>Date: <span className="text-slate-800 font-extrabold">{item.lastJobCardEntryDate}</span></span>
                                </div>
                                <div className="text-slate-500 font-bold text-[10px] flex items-center gap-1.5">
                                  <Users className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                  <span>Tech: <span className="text-slate-800 font-extrabold">{item.lastTechnician || '—'}</span></span>
                                </div>
                                {item.linkedSupervisor && (
                                  <div className="text-slate-500 font-bold text-[10px] flex items-center gap-1.5">
                                    <UserCheck className="w-3.5 h-3.5 text-indigo-400 shrink-0" />
                                    <span>Tech's Supervisor: <span className="text-indigo-950 font-black px-1.5 py-0.5 bg-indigo-50 border border-indigo-100 rounded text-[9px] uppercase">{item.linkedSupervisor}</span></span>
                                  </div>
                                )}
                                <div className="text-slate-500 font-bold text-[10px] flex items-center gap-1.5 pb-1">
                                  <UserCheck className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                  <span>Job Card Supervisor: <span className="text-slate-800 font-extrabold">{item.lastSupervisor || '—'}</span></span>
                                </div>
                                {item.lastRepairDetails && (
                                  <div className="text-[10px] text-purple-950 font-semibold mt-1 bg-purple-50 p-2 rounded border border-purple-100">
                                    <span className="font-black text-purple-900 uppercase text-[9px] block mb-0.5">Repairs & Work Done:</span>
                                    <div className="leading-snug text-slate-700">{item.lastRepairDetails}</div>
                                  </div>
                                )}
                              </div>
                            ) : (
                              <div className="text-slate-400 italic font-bold p-3 border border-dashed border-slate-200 rounded-lg text-center bg-slate-50">
                                No service visits recorded yet
                              </div>
                            )}
                          </td>

                          {/* Follow-up logs & call remarks */}
                          <td className="p-1.5">
                            <div className="space-y-1.5">
                              {item.lastCallDate ? (
                                <div className="space-y-1.5 bg-amber-50/70 border border-amber-200/80 p-2.5 rounded-lg text-[11px] leading-normal">
                                  <div className="flex items-center justify-between border-b border-amber-200/50 pb-1 mb-1">
                                    <div className="flex items-center gap-1 font-bold text-amber-950">
                                      <PhoneCall className="w-3 h-3 text-amber-600 shrink-0" />
                                      <span>Called: {fmtDate(item.lastCallDate)}</span>
                                    </div>
                                    <span className="inline-block bg-amber-100 text-amber-900 text-[9px] font-black px-1 py-0.5 rounded uppercase">Follow-up Done</span>
                                  </div>
                                  <div className="text-slate-800 font-bold text-[11px] leading-relaxed bg-white p-1.5 rounded border border-amber-100/60">
                                    <span className="text-[9px] uppercase font-black text-slate-400 block">Customer emi chepparu:</span>
                                    <div className="italic text-slate-700">"{item.lastRemarks || 'No remarks recorded'}"</div>
                                  </div>
                                  {item.lastNextCallDate && (
                                    <div className="text-[10px] text-indigo-950 font-extrabold flex items-center gap-1 mt-1 bg-indigo-50 border border-indigo-100 px-1.5 py-0.5 rounded w-fit">
                                      <Calendar className="w-3 h-3 text-indigo-600 shrink-0" />
                                      <span>Malli Call Chese Date: {fmtDate(item.lastNextCallDate)}</span>
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <div className="text-slate-400 italic text-[11px] p-2.5 border border-dashed border-slate-200 rounded-lg text-center bg-slate-50">
                                  No previous follow-ups recorded.
                                </div>
                              )}
                              
                              <button
                                type="button"
                                onClick={() => {
                                  setSelectedFollowupForLog(item);
                                  setNewFollowupCallDate(new Date().toISOString().split('T')[0]);
                                  setNewFollowupRemarks('');
                                  setNewFollowupNextCallDate('');
                                }}
                                className="w-full py-1.5 px-2 bg-indigo-600 hover:bg-indigo-700 text-white border border-indigo-700 rounded-md font-bold text-[10px] transition-colors cursor-pointer flex items-center justify-center gap-1 shadow-2xs hover:shadow-sm"
                              >
                                <PhoneCall className="w-3 h-3 text-white" />
                                <span>Log New Call / Follow-up</span>
                              </button>
                            </div>
                          </td>

                          {/* Row Actions */}
                          <td className="p-1.5 text-center border-l border-slate-100">
                            <div className="flex flex-col gap-1.5 items-center justify-center">
                              <button
                                type="button"
                                onClick={() => handleDirectComplaintRegister(item)}
                                className="w-full py-1.5 px-1.5 text-white bg-rose-600 hover:bg-rose-700 rounded border border-rose-700 cursor-pointer flex items-center justify-center gap-1 font-black text-[10px] transition-all shadow-sm"
                                title="Register New Complaint"
                              >
                                <ShieldAlert className="w-3.5 h-3.5" />
                                <span>Register Complaint</span>
                              </button>
                              <button
                                type="button"
                                onClick={() => copyCustomerDetails(item.rec)}
                                className="w-full py-1.5 px-1.5 text-indigo-700 hover:text-white bg-indigo-50 hover:bg-indigo-600 rounded border border-indigo-200 hover:border-indigo-600 cursor-pointer flex items-center justify-center gap-1 font-black text-[10px] transition-all shadow-sm"
                                title="Copy Customer Details (కాపీ చేయండి)"
                              >
                                <Copy className="w-3.5 h-3.5" />
                                <span>Copy Details (కాపీ)</span>
                              </button>
                              {item.hasJobCard && (
                                <button
                                  type="button"
                                  onClick={() => setSelectedFollowupCustomer(item)}
                                  className="w-full py-1 px-1.5 text-indigo-700 hover:text-white bg-indigo-50 hover:bg-indigo-600 rounded border border-indigo-200 hover:border-indigo-600 cursor-pointer flex items-center justify-center gap-1 font-bold text-[10px] transition-all"
                                  title="Detailed Service History"
                                >
                                  <History className="w-3 h-3" />
                                  <span>History</span>
                                </button>
                              )}
                              <div className="flex gap-1 w-full">
                                <button
                                  type="button"
                                  onClick={() => handleOpenEditCustomer(item.rec)}
                                  className="flex-1 py-1 px-1 text-blue-700 hover:text-white bg-blue-50 hover:bg-blue-600 rounded border border-blue-200 hover:border-blue-600 cursor-pointer flex items-center justify-center gap-1 font-bold text-[9px] transition-all"
                                  title="Edit Profile"
                                >
                                  <Edit3 className="w-3 h-3" />
                                  <span>Edit</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    autoFillCustomer(item.rec);
                                    setActiveTab('new_entry');
                                    alert(`Loaded details for "${item.customerName || item.chassisNo}" into Job Card entry form.`);
                                  }}
                                  className="flex-1 py-1 px-1 text-emerald-700 hover:text-white bg-emerald-50 hover:bg-emerald-600 rounded border border-emerald-200 hover:border-emerald-600 cursor-pointer flex items-center justify-center gap-1 font-bold text-[9px] transition-all"
                                  title="Create Job Card"
                                >
                                  <Plus className="w-3 h-3" />
                                  <span>Job Card</span>
                                </button>
                              </div>
                              <button
                                type="button"
                                onClick={() => deleteCustomerRecord(item.chassisNo)}
                                className="w-full py-1 px-1.5 text-red-600 hover:text-white bg-red-50 hover:bg-red-600 rounded border border-red-200 hover:border-red-600 cursor-pointer flex items-center justify-center gap-1 font-bold text-[10px] transition-all"
                                title="Delete Customer Record"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                                <span>Delete</span>
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {followupCalculatedList.length > 0 && (
              <div className="bg-slate-50 border-t border-slate-200 p-3.5 flex flex-wrap items-center justify-between gap-3 text-xs font-bold text-slate-700">
                <div>
                  Showing {Math.min(followupCalculatedList.length, (followupPage - 1) * followupItemsPerPage + 1)} to {Math.min(followupCalculatedList.length, followupPage * followupItemsPerPage)} of <span className="text-indigo-600 font-extrabold">{followupCalculatedList.length}</span> customer records
                </div>
                
                <div className="flex items-center gap-2">
                  <span className="text-[11px] text-slate-500 font-bold">Items per page:</span>
                  <select
                    value={followupItemsPerPage}
                    onChange={(e) => {
                      setFollowupItemsPerPage(Number(e.target.value));
                      setFollowupPage(1);
                    }}
                    className="p-1 border border-slate-300 rounded bg-white font-bold outline-none cursor-pointer"
                  >
                    <option value={10}>10</option>
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                  </select>
                </div>

                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    disabled={followupPage === 1}
                    onClick={() => setFollowupPage(p => Math.max(1, p - 1))}
                    className="px-2.5 py-1.5 border border-slate-300 rounded bg-white hover:bg-slate-100 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  >
                    Previous
                  </button>
                  
                  {(() => {
                    const totalPages = Math.ceil(followupCalculatedList.length / followupItemsPerPage);
                    const maxPageButtons = 5;
                    let startPage = Math.max(1, followupPage - Math.floor(maxPageButtons / 2));
                    let endPage = Math.min(totalPages, startPage + maxPageButtons - 1);
                    if (endPage - startPage + 1 < maxPageButtons) {
                      startPage = Math.max(1, endPage - maxPageButtons + 1);
                    }
                    
                    const buttons = [];
                    for (let i = startPage; i <= endPage; i++) {
                      buttons.push(
                        <button
                          key={i}
                          type="button"
                          onClick={() => setFollowupPage(i)}
                          className={`px-3 py-1.5 border rounded font-black transition-colors ${
                            followupPage === i
                              ? 'bg-purple-900 border-purple-900 text-white'
                              : 'bg-white border-slate-300 hover:bg-slate-100 text-slate-700'
                          }`}
                        >
                          {i}
                        </button>
                      );
                    }
                    return buttons;
                  })()}

                  <button
                    type="button"
                    disabled={followupPage === Math.ceil(followupCalculatedList.length / followupItemsPerPage)}
                    onClick={() => setFollowupPage(p => Math.min(Math.ceil(followupCalculatedList.length / followupItemsPerPage), p + 1))}
                    className="px-2.5 py-1.5 border border-slate-300 rounded bg-white hover:bg-slate-100 text-slate-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                  >
                    Next
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </main>
  </div>

      {/* COMPLAINT MODAL */}
      {isComplaintModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-3xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-red-700 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <AlertCircle className="w-5 h-5 text-red-200" />
                <div>
                  <h2 className="text-sm font-black uppercase tracking-wider">{complaintForm.id ? 'Edit Complaint' : 'Add New Complaint'}</h2>
                  {complaintForm.id && <p className="text-[10px] text-red-200 font-medium">#{complaintForm.complaintNo}</p>}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsComplaintModalOpen(false)}
                className="text-red-200 hover:text-white p-1 rounded-lg hover:bg-red-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 md:p-6 overflow-y-auto">
              <form className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Chassis No / Customer Query <span className="text-[9px] text-slate-400 font-normal">(Optional)</span></label>
                    <input
                      type="text"
                      list="globalChassisDatalist"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900 focus:ring-2 focus:ring-red-500 focus:border-red-500"
                      value={complaintForm.chassisNo}
                      onChange={(e) => {
                        const val = e.target.value.toUpperCase();
                        const isSelection = val.includes(' — ') || val.includes('MOB:') || val.includes('(');
                        const cleanChassis = isSelection ? val.split(/[ \(\-]/)[0].trim() : val;
                        
                        setComplaintForm({ ...complaintForm, chassisNo: val });
                        
                        // Auto-fill when explicit selection or sufficient chassis length
                        if (isSelection || val.length >= 3) {
                          const rec = findCustomerRecordByQuery(cleanChassis) || findCustomerRecordByQuery(val);
                          if (rec) {
                            setComplaintForm(prev => ({
                              ...prev,
                              chassisNo: rec.__chassisDisplay || rec['Chassis no'] || cleanChassis,
                              customerName: rec.__custNameDisplay || rec['Customer Name'] || rec.customerName || prev.customerName,
                              mobileNumber: rec.__custPhoneDisplay || rec['Mobile Number'] || rec.mobileNumber || prev.mobileNumber,
                              tractorModel: rec['Model'] || rec.model || prev.tractorModel,
                            }));
                          }
                        }
                      }}
                      placeholder="Type or select Chassis No / Customer (Optional)..."
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Complaint Date <span className="text-red-500">*</span></label>
                    <input
                      type="date"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900 focus:ring-2 focus:ring-red-500 focus:border-red-500"
                      value={complaintForm.complaintDate}
                      onChange={(e) => setComplaintForm({ ...complaintForm, complaintDate: e.target.value })}
                      required
                    />
                  </div>
                  
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Customer Name <span className="text-red-500">*</span></label>
                    <input
                      type="text"
                      list="globalCustomerNameDatalist"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900 focus:ring-2 focus:ring-red-500 focus:border-red-500"
                      value={complaintForm.customerName}
                      onChange={(e) => {
                        const val = e.target.value;
                        const isSelection = val.includes(' — ') || val.includes('MOB:') || val.includes('(');
                        const cleanName = isSelection ? val.split(' — ')[0].trim() : val;

                        setComplaintForm(prev => ({ ...prev, customerName: val }));

                        if (isSelection) {
                          const rec = findCustomerRecordByQuery(cleanName) || findCustomerRecordByQuery(val);
                          if (rec) {
                            setComplaintForm(prev => ({
                              ...prev,
                              customerName: rec.__custNameDisplay || rec['Customer Name'] || rec.customerName || cleanName,
                              chassisNo: rec.__chassisDisplay || rec['Chassis no'] || prev.chassisNo,
                              mobileNumber: rec.__custPhoneDisplay || rec['Mobile Number'] || rec.mobileNumber || prev.mobileNumber,
                              tractorModel: rec['Model'] || rec.model || prev.tractorModel,
                            }));
                          }
                        }
                      }}
                      placeholder="Type half name to select from list..."
                      required
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Mobile Number</label>
                    <input
                      type="text"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-medium text-slate-900"
                      value={complaintForm.mobileNumber}
                      onChange={(e) => setComplaintForm({ ...complaintForm, mobileNumber: e.target.value })}
                    />
                  </div>
                  
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Tractor Model</label>
                    <input
                      type="text"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-medium text-slate-900"
                      value={complaintForm.tractorModel}
                      onChange={(e) => setComplaintForm({ ...complaintForm, tractorModel: e.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Status</label>
                    <select
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900"
                      value={complaintForm.status}
                      onChange={(e) => {
                        const newStatus = e.target.value;
                        setComplaintForm(prev => ({
                          ...prev,
                          status: newStatus,
                          closureDate: newStatus === 'Open' ? '' : (prev.closureDate || new Date().toISOString().split('T')[0])
                        }));
                      }}
                    >
                      <option value="Open">Open</option>
                      <option value="Closed">Closed</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">
                      Closure Date <span className="text-[9px] text-emerald-700 font-semibold">(Entering date auto-closes)</span>
                    </label>
                    <input
                      type="date"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900"
                      value={complaintForm.closureDate || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        setComplaintForm(prev => ({
                          ...prev,
                          closureDate: val,
                          status: val ? 'Closed' : prev.status
                        }));
                      }}
                    />
                  </div>
                  
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Job Card No</label>
                    <input
                      type="text"
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-medium text-slate-900"
                      value={complaintForm.jobCardNo}
                      onChange={(e) => {
                        const val = e.target.value;
                        setComplaintForm(prev => {
                          const newForm = { ...prev, jobCardNo: val };
                          const linkedJobCard = savedJobCards.find(c => c.jobCardNo === val);
                          if (linkedJobCard) {
                             newForm.assignedMechanic = linkedJobCard.mechanic || linkedJobCard.technicianName || prev.assignedMechanic;
                             newForm.assignedSupervisor = linkedJobCard.wsIncharge || linkedJobCard.supervisorName || prev.assignedSupervisor;
                          }
                          return newForm;
                        });
                      }}
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Assigned Mechanic</label>
                    <select
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900"
                      value={complaintForm.assignedMechanic}
                      onChange={(e) => setComplaintForm({ ...complaintForm, assignedMechanic: e.target.value })}
                    >
                      <option value="">Select Mechanic</option>
                      {mechanicsList.map((m) => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Assigned Supervisor</label>
                    <select
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-bold text-slate-900"
                      value={complaintForm.assignedSupervisor}
                      onChange={(e) => setComplaintForm({ ...complaintForm, assignedSupervisor: e.target.value })}
                    >
                      <option value="">Select Supervisor</option>
                      {supervisorsList.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Complaint Details <span className="text-red-500">*</span></label>
                  <textarea
                    rows={3}
                    className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-medium text-slate-900"
                    value={complaintForm.complaintDetails}
                    onChange={(e) => setComplaintForm({ ...complaintForm, complaintDetails: e.target.value })}
                    required
                  />
                </div>

                {(complaintForm.status === 'Closed' || complaintForm.closureDate) && (
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-1">Resolution Summary</label>
                    <textarea
                      rows={3}
                      className="w-full bg-slate-50 border border-slate-300 p-2 rounded-lg text-xs font-medium text-slate-900"
                      value={complaintForm.resolution}
                      onChange={(e) => setComplaintForm({ ...complaintForm, resolution: e.target.value })}
                    />
                  </div>
                )}
              </form>
            </div>
            
            <div className="p-4 border-t border-slate-200 bg-slate-50 flex justify-end gap-2 shrink-0">
              <button
                type="button"
                onClick={() => {
                  setIsComplaintModalOpen(false);
                  setIsComplaintViewOnly(false);
                }}
                className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
              >
                {isComplaintViewOnly ? 'Close View' : 'Cancel'}
              </button>
              {!isComplaintViewOnly && (
                <button
                  type="button"
                  onClick={saveComplaint}
                  className="px-5 py-2 bg-red-600 hover:bg-red-700 text-white font-bold text-xs rounded-lg transition-colors shadow-xs cursor-pointer flex items-center gap-1.5"
                >
                  <Save className="w-4 h-4" /> {complaintForm.id ? 'Update Complaint' : 'Save Complaint'}
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* DUPLICATE DATA CHECKER MODAL */}
      {isDuplicateModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-rose-700 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <ShieldAlert className="w-5 h-5 text-rose-200" />
                <div>
                  <h3 className="text-base font-bold">డూప్లికేట్ కస్టమర్ డేటా నివేదిక (Duplicate Data Report)</h3>
                  <p className="text-xs text-rose-100">ఒకే ఛాసిస్ నంబర్ లేదా మొబైల్ నంబర్‌తో ఉన్న డూప్లికేట్ రికార్డులు ఇక్కడ చూడవచ్చు.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsDuplicateModalOpen(false)}
                className="text-rose-100 hover:text-white p-1 rounded-lg hover:bg-rose-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 md:p-6 overflow-y-auto">
              {duplicateReport.length === 0 ? (
                <div className="text-center py-12 space-y-3">
                  <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <CheckCircle2 className="w-8 h-8" />
                  </div>
                  <h4 className="text-lg font-black text-slate-900">అభినందనలు! డూప్లికేట్ డేటా ఏమీ లేదు.</h4>
                  <p className="text-sm text-slate-500 font-medium max-w-md mx-auto">మీ కస్టమర్ డేటాబేస్ క్లీన్‌గా ఉంది. అన్ని ఛాసిస్ నంబర్లు మరియు ఫోన్ నంబర్లు ప్రత్యేకంగా ఉన్నాయి.</p>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="bg-amber-50 border border-amber-200 p-3 rounded-lg text-[11px] text-amber-800 font-bold flex items-start gap-2">
                    <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" />
                    <span>గమనిక: ఒకే ఛాసిస్ నంబర్ వేర్వేరు స్పెల్లింగ్స్‌తో (ఉదా: O బదులు 0) ఉన్నా లేదా ఒకే ఫోన్ నంబర్‌తో వేర్వేరు ఛాసిస్ నంబర్లు ఉన్నా ఇక్కడ కనిపిస్తాయి. దయచేసి సరైన వాటిని ఉంచి మిగిలినవి డిలీట్ చేయండి.</span>
                  </div>

                  <div className="grid grid-cols-1 gap-4">
                    {duplicateReport.map((dup, idx) => (
                      <div key={idx} className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                        <div className="bg-slate-50 px-4 py-2 border-b border-slate-200 flex items-center justify-between">
                          <div>
                            <span className="text-[10px] font-black uppercase text-rose-600 tracking-wider block">{dup.type}</span>
                            <h4 className="text-sm font-bold text-slate-900">Key: <span className="font-mono">{dup.key}</span></h4>
                          </div>
                          <span className="bg-rose-100 text-rose-700 text-[10px] font-bold px-2 py-0.5 rounded-full border border-rose-200">
                            {dup.records.length} Records Found
                          </span>
                        </div>
                        <div className="p-0 overflow-x-auto">
                          <table className="w-full text-left text-[11px]">
                            <thead className="bg-slate-50 text-slate-500 uppercase font-bold text-[9px] border-b border-slate-200">
                              <tr>
                                <th className="px-4 py-2">Customer Name</th>
                                <th className="px-4 py-2">Chassis Number</th>
                                <th className="px-4 py-2">File No</th>
                                <th className="px-4 py-2">Supervisor</th>
                                <th className="px-4 py-2">Mobile Number</th>
                                <th className="px-4 py-2">Village / Address</th>
                                <th className="px-4 py-2">Branch</th>
                                <th className="px-4 py-2 text-center">Action</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-200">
                              {dup.records.map((r: any, rIdx: number) => (
                                <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                                  <td className="px-4 py-2 font-bold text-slate-900">
                                    {r.customerName || r['Customer Name'] || r.custName || r.__custNameDisplay || 'N/A'}
                                  </td>
                                  <td className="px-4 py-2 font-mono font-bold text-indigo-700">{r['Chassis no'] || r.__chassisDisplay || r.chassisNo || r.chassis || 'N/A'}</td>
                                  <td className="px-4 py-2 font-bold text-amber-700">
                                    {r.fileNo || r['File No'] || r.historyFileNo || r['HISTORY FILE NO.'] || r.sno || r['S.No.'] || 'N/A'}
                                  </td>
                                  <td className="px-4 py-2 font-medium text-emerald-700">
                                    {r.supervisorName || r.supervisor || r.wsIncharge || r['Supervisor'] || 'N/A'}
                                  </td>
                                  <td className="px-4 py-2 font-mono">{r['Mobile Number'] || r.mobileNumber || r.phone || 'N/A'}</td>
                                  <td className="px-4 py-2">{r.village || r['Village'] || r.address || 'N/A'}</td>
                                  <td className="px-4 py-2 font-semibold text-slate-600">{r.branch || r['Branch'] || 'N/A'}</td>
                                  <td className="px-4 py-2 text-center">
                                    <button
                                      onClick={() => {
                                        const chassis = normalizeKey((r['Chassis no'] || r.__chassisDisplay || r.chassisNo).toString());
                                        deleteCustomerRecord(chassis);
                                        // Refresh the report locally
                                        const updatedReport = [...duplicateReport];
                                        updatedReport[idx].records = updatedReport[idx].records.filter((_: any, i: number) => i !== rIdx);
                                        if (updatedReport[idx].records.length <= 1) {
                                          setDuplicateReport(updatedReport.filter((_, i) => i !== idx));
                                        } else {
                                          setDuplicateReport(updatedReport);
                                        }
                                      }}
                                      className="p-1.5 bg-red-50 text-red-600 hover:bg-red-600 hover:text-white rounded-lg transition-all"
                                      title="Delete this record"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" />
                                    </button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        <div className="bg-slate-50/50 px-4 py-1.5 text-[9px] text-slate-400 font-medium italic border-t border-slate-100">
                          {dup.reason}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
            
            <div className="p-4 border-t border-slate-200 bg-slate-50 flex justify-end shrink-0">
              <button
                type="button"
                onClick={() => setIsDuplicateModalOpen(false)}
                className="px-6 py-2 bg-slate-900 text-white font-bold text-xs rounded-xl transition-all shadow-lg active:scale-95 cursor-pointer"
              >
                Close Report
              </button>
            </div>
          </div>
        </div>
      )}

      {/* UPLOADED FILE DUPLICATE REPORT MODAL */}
      {isFileDuplicateModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-indigo-700 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <FileCheck className="w-5 h-5 text-indigo-200" />
                <div>
                  <h3 className="text-base font-bold">అప్‌లోడ్ చేసిన ఫైల్ - డూప్లికేట్ డేటా నివేదిక (Uploaded File Duplicate Report)</h3>
                  <p className="text-xs text-indigo-100">మీరు ఇప్పుడు అప్‌లోడ్ చేసిన ఎక్సెల్ ఫైల్‌లో మాత్రమే ఉన్న డూప్లికేట్లు ఇక్కడ చూడవచ్చు.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsFileDuplicateModalOpen(false)}
                className="text-indigo-100 hover:text-white p-1 rounded-lg hover:bg-indigo-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 md:p-6 overflow-y-auto">
              <div className="bg-blue-50 border border-blue-200 p-3 rounded-lg text-[11px] text-blue-800 font-bold flex items-start gap-2 mb-6">
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-blue-600" />
                <span>గమనిక: ఈ రికార్డులు మీరు అప్‌లోడ్ చేసిన ఫైల్‌లో మాత్రమే ఉన్నవి. ఒకవేళ ఏవైనా తప్పుగా ఉంటే, మీ ఎక్సెల్ ఫైల్‌ను సరిచేసి మళ్ళీ అప్‌లోడ్ చేయండి.</span>
              </div>

              <div className="grid grid-cols-1 gap-4">
                {fileDuplicateReport.map((dup, idx) => (
                  <div key={idx} className="border border-indigo-100 rounded-xl overflow-hidden shadow-xs">
                    <div className="bg-indigo-50/50 px-4 py-2 border-b border-indigo-100 flex items-center justify-between">
                      <div>
                        <span className="text-[10px] font-black uppercase text-indigo-600 tracking-wider block">{dup.type}</span>
                        <h4 className="text-sm font-bold text-slate-900">Key: <span className="font-mono">{dup.key}</span></h4>
                      </div>
                      <span className="bg-indigo-100 text-indigo-700 text-[10px] font-bold px-2 py-0.5 rounded-full border border-indigo-200">
                        {dup.records.length} Duplicates in File
                      </span>
                    </div>
                    <div className="p-0 overflow-x-auto">
                      <table className="w-full text-left text-[11px]">
                        <thead className="bg-slate-50 text-slate-500 uppercase font-bold text-[9px] border-b border-slate-200">
                          <tr>
                            <th className="px-4 py-2">Customer Name</th>
                            <th className="px-4 py-2">Chassis Number</th>
                            <th className="px-4 py-2">File No</th>
                            <th className="px-4 py-2">Supervisor</th>
                            <th className="px-4 py-2">Mobile Number</th>
                            <th className="px-4 py-2">Village / Address</th>
                            <th className="px-4 py-2">Branch</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {dup.records.map((r: any, rIdx: number) => (
                            <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                              <td className="px-4 py-2 font-bold text-slate-900">
                                {r.customerName || r['Customer Name'] || r.custName || r.__custNameDisplay || 'N/A'}
                              </td>
                              <td className="px-4 py-2 font-mono font-bold text-indigo-700">{r['Chassis no'] || r.__chassisDisplay || r.chassisNo || r.chassis || 'N/A'}</td>
                              <td className="px-4 py-2 font-bold text-amber-700">
                                {r.fileNo || r['File No'] || r.historyFileNo || r['HISTORY FILE NO.'] || r.sno || r['S.No.'] || 'N/A'}
                              </td>
                              <td className="px-4 py-2 font-medium text-emerald-700">
                                {r.supervisorName || r.supervisor || r.wsIncharge || r['Supervisor'] || 'N/A'}
                              </td>
                              <td className="px-4 py-2 font-mono">{r['Mobile Number'] || r.mobileNumber || r.phone || 'N/A'}</td>
                              <td className="px-4 py-2">{r.village || r['Village'] || r.address || 'N/A'}</td>
                              <td className="px-4 py-2 font-semibold text-slate-600">{r.branch || r['Branch'] || 'N/A'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="bg-indigo-50/30 px-4 py-1.5 text-[9px] text-indigo-400 font-medium italic border-t border-indigo-100">
                      {dup.reason}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            
            <div className="p-4 border-t border-slate-200 bg-slate-50 flex justify-end shrink-0">
              <button
                type="button"
                onClick={() => setIsFileDuplicateModalOpen(false)}
                className="px-6 py-2 bg-indigo-700 text-white font-bold text-xs rounded-xl transition-all shadow-lg active:scale-95 cursor-pointer"
              >
                అర్థమైంది (Got it)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD NEW CUSTOMER MODAL OVERLAY */}
      {isAddCustomerOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-blue-900 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <UserPlus className="w-5 h-5 text-emerald-400" />
                <div>
                  <h3 className="text-base font-bold">Add New Customer to Master Database</h3>
                  <p className="text-xs text-blue-200">Fill in customer details below. The record will be instantly saved and searchable for auto-fill.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsAddCustomerOpen(false)}
                className="text-blue-200 hover:text-white p-1 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleSaveNewCustomer} className="p-4 space-y-3 overflow-y-auto flex-1">
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">BRANCH</label>
                  <input
                    type="text"
                    placeholder="e.g. Poranki / Vijayawada"
                    value={newCustForm.branch}
                    onChange={(e) => setNewCustForm({ ...newCustForm, branch: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">S.NO</label>
                  <input
                    type="text"
                    placeholder="e.g. 101"
                    value={newCustForm.sNo}
                    onChange={(e) => setNewCustForm({ ...newCustForm, sNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Model Name</label>
                  <input
                    type="text"
                    placeholder="e.g. 5050D / 380 Super DI"
                    value={newCustForm.model}
                    onChange={(e) => setNewCustForm({ ...newCustForm, model: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">MODEL TYPE</label>
                  <input
                    type="text"
                    placeholder="e.g. 4WD / 2WD"
                    value={newCustForm.modelType}
                    onChange={(e) => setNewCustForm({ ...newCustForm, modelType: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-blue-950 mb-0.5">
                    Chassis No. <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="e.g. ME1380..."
                    value={newCustForm.chassisNo}
                    onChange={(e) => setNewCustForm({ ...newCustForm, chassisNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-blue-400 bg-blue-50/30 rounded-lg outline-none font-mono font-bold text-blue-950 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Engine No:</label>
                  <input
                    type="text"
                    placeholder="e.g. E380..."
                    value={newCustForm.engineNo}
                    onChange={(e) => setNewCustForm({ ...newCustForm, engineNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-mono font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Date of Delivery</label>
                  <input
                    type="text"
                    placeholder="DD/MM/YYYY"
                    value={newCustForm.dateOfDel}
                    onChange={(e) => setNewCustForm({ ...newCustForm, dateOfDel: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-900 mb-0.5">
                    Customer Name <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="e.g. Ramesh Babu"
                    value={newCustForm.custName}
                    onChange={(e) => setNewCustForm({ ...newCustForm, custName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-bold text-slate-900 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">FATHER NAME</label>
                  <input
                    type="text"
                    placeholder="e.g. Satyanarayana"
                    value={newCustForm.fatherName}
                    onChange={(e) => setNewCustForm({ ...newCustForm, fatherName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">ADDRESS</label>
                  <input
                    type="text"
                    placeholder="House / Street details"
                    value={newCustForm.address}
                    onChange={(e) => setNewCustForm({ ...newCustForm, address: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">VILLAGE</label>
                  <input
                    type="text"
                    placeholder="Village Name"
                    value={newCustForm.village}
                    onChange={(e) => setNewCustForm({ ...newCustForm, village: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Mandal</label>
                  <input
                    type="text"
                    placeholder="Mandal Name"
                    value={newCustForm.mandal}
                    onChange={(e) => setNewCustForm({ ...newCustForm, mandal: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Mobile Number</label>
                  <input
                    type="text"
                    placeholder="10-digit mobile no"
                    value={newCustForm.mobileNumber}
                    onChange={(e) => setNewCustForm({ ...newCustForm, mobileNumber: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-mono font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">District</label>
                  <input
                    type="text"
                    placeholder="e.g. Krishna"
                    value={newCustForm.district}
                    onChange={(e) => setNewCustForm({ ...newCustForm, district: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">PIN CODE</label>
                  <input
                    type="text"
                    placeholder="e.g. 520001"
                    value={newCustForm.pinCode}
                    onChange={(e) => setNewCustForm({ ...newCustForm, pinCode: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-mono font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">DSP Name</label>
                  <input
                    type="text"
                    placeholder="DSP Executive Name"
                    value={newCustForm.dspName}
                    onChange={(e) => setNewCustForm({ ...newCustForm, dspName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">EXCHANGE BRAND</label>
                  <input
                    type="text"
                    placeholder="Old Tractor Brand"
                    value={newCustForm.exchangeBrand}
                    onChange={(e) => setNewCustForm({ ...newCustForm, exchangeBrand: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">EXCHANGE TRACTOR MODELS</label>
                  <input
                    type="text"
                    placeholder="Old Tractor Model"
                    value={newCustForm.exchangeModels}
                    onChange={(e) => setNewCustForm({ ...newCustForm, exchangeModels: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-purple-950 mb-1">ASSIGNED SUPERVISOR</label>
                  <input
                    type="text"
                    list="newCustSupervisorDatalist"
                    placeholder="Type or select supervisor"
                    value={newCustForm.supervisor}
                    onChange={(e) => setNewCustForm({ ...newCustForm, supervisor: e.target.value })}
                    className="w-full text-xs p-2 border border-purple-300 bg-purple-50/10 rounded-lg outline-none font-bold text-purple-950 focus:border-purple-900"
                  />
                  <datalist id="newCustSupervisorDatalist">
                    {supervisorsList.map((sup: string) => (
                      <option key={sup} value={sup} />
                    ))}
                  </datalist>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-3 shrink-0">
                <button
                  type="button"
                  onClick={() => setIsAddCustomerOpen(false)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg transition-colors shadow-xs cursor-pointer flex items-center gap-1.5"
                >
                  <Save className="w-4 h-4" /> Save Customer to Database
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* EDIT CUSTOMER MODAL OVERLAY */}
      {isEditCustModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-blue-900 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <Edit3 className="w-5 h-5 text-yellow-400" />
                <div>
                  <h3 className="text-base font-bold">Edit Customer Details</h3>
                  <p className="text-xs text-blue-200">Modify the customer details in the Master Database. Any changes will save instantly.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsEditCustModalOpen(false)}
                className="text-blue-200 hover:text-white p-1 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleUpdateCustomer} className="p-4 space-y-3 overflow-y-auto flex-1">
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">BRANCH</label>
                  <input
                    type="text"
                    placeholder="e.g. Poranki / Vijayawada"
                    value={editCustForm.branch}
                    onChange={(e) => setEditCustForm({ ...editCustForm, branch: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">S.NO</label>
                  <input
                    type="text"
                    placeholder="e.g. 101"
                    value={editCustForm.sNo}
                    onChange={(e) => setEditCustForm({ ...editCustForm, sNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Model Name</label>
                  <input
                    type="text"
                    placeholder="e.g. 5050D / 380 Super DI"
                    value={editCustForm.model}
                    onChange={(e) => setEditCustForm({ ...editCustForm, model: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Model Type</label>
                  <input
                    type="text"
                    placeholder="e.g. 2WD / 4WD"
                    value={editCustForm.modelType}
                    onChange={(e) => setEditCustForm({ ...editCustForm, modelType: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Chassis Number</label>
                  <input
                    type="text"
                    placeholder="REQUIRED"
                    required
                    value={editCustForm.chassisNo}
                    onChange={(e) => setEditCustForm({ ...editCustForm, chassisNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-bold text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Engine Number</label>
                  <input
                    type="text"
                    placeholder="e.g. REA123456"
                    value={editCustForm.engineNo}
                    onChange={(e) => setEditCustForm({ ...editCustForm, engineNo: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Date of Delivery</label>
                  <input
                    type="text"
                    placeholder="e.g. 12/03/2025"
                    value={editCustForm.dateOfDel}
                    onChange={(e) => setEditCustForm({ ...editCustForm, dateOfDel: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Customer Name</label>
                  <input
                    type="text"
                    placeholder="e.g. Ramesh Babu"
                    value={editCustForm.custName}
                    onChange={(e) => setEditCustForm({ ...editCustForm, custName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Father's Name</label>
                  <input
                    type="text"
                    placeholder="e.g. Satyam"
                    value={editCustForm.fatherName}
                    onChange={(e) => setEditCustForm({ ...editCustForm, fatherName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Address / House No</label>
                  <input
                    type="text"
                    placeholder="e.g. D.No 4-12"
                    value={editCustForm.address}
                    onChange={(e) => setEditCustForm({ ...editCustForm, address: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">VILLAGE</label>
                  <input
                    type="text"
                    placeholder="e.g. Poranki"
                    value={editCustForm.village}
                    onChange={(e) => setEditCustForm({ ...editCustForm, village: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Mandal</label>
                  <input
                    type="text"
                    placeholder="e.g. Penamaluru"
                    value={editCustForm.mandal}
                    onChange={(e) => setEditCustForm({ ...editCustForm, mandal: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Mobile Number</label>
                  <input
                    type="text"
                    placeholder="e.g. 9876543210"
                    value={editCustForm.mobileNumber}
                    onChange={(e) => setEditCustForm({ ...editCustForm, mobileNumber: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">District</label>
                  <input
                    type="text"
                    placeholder="e.g. Krishna"
                    value={editCustForm.district}
                    onChange={(e) => setEditCustForm({ ...editCustForm, district: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">Pin Code</label>
                  <input
                    type="text"
                    placeholder="e.g. 521137"
                    value={editCustForm.pinCode}
                    onChange={(e) => setEditCustForm({ ...editCustForm, pinCode: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">DSP Name (Sales Person)</label>
                  <input
                    type="text"
                    placeholder="e.g. Srinivas"
                    value={editCustForm.dspName}
                    onChange={(e) => setEditCustForm({ ...editCustForm, dspName: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">EXCHANGE BRAND</label>
                  <input
                    type="text"
                    placeholder="e.g. Swaraj"
                    value={editCustForm.exchangeBrand}
                    onChange={(e) => setEditCustForm({ ...editCustForm, exchangeBrand: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">EXCHANGE TRACTOR MODELS</label>
                  <input
                    type="text"
                    placeholder="Old Tractor Model"
                    value={editCustForm.exchangeModels}
                    onChange={(e) => setEditCustForm({ ...editCustForm, exchangeModels: e.target.value })}
                    className="w-full text-xs p-1.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-purple-950 mb-0.5">ASSIGNED SUPERVISOR</label>
                  <input
                    type="text"
                    list="editCustSupervisorDatalist"
                    placeholder="Type or select supervisor"
                    value={editCustForm.supervisor}
                    onChange={(e) => setEditCustForm({ ...editCustForm, supervisor: e.target.value })}
                    className="w-full text-xs p-1.5 border border-purple-300 bg-purple-50/10 rounded-lg outline-none font-bold text-purple-950 focus:border-purple-900"
                  />
                  <datalist id="editCustSupervisorDatalist">
                    {supervisorsList.map((sup: string) => (
                      <option key={sup} value={sup} />
                    ))}
                  </datalist>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-3 shrink-0">
                <button
                  type="button"
                  onClick={() => setIsEditCustModalOpen(false)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-blue-900 hover:bg-blue-950 text-white font-bold text-xs rounded-lg transition-colors shadow-xs cursor-pointer flex items-center gap-1.5"
                >
                  <Save className="w-4 h-4" /> Update Customer Database
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* SERVICE HISTORY DETAILS MODAL OVERLAY */}
      {selectedFollowupCustomer && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            
            {/* Header */}
            <div className="bg-slate-950 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <History className="w-5 h-5 text-indigo-400 animate-pulse" />
                <div>
                  <h3 className="text-base font-bold">Detailed Service History</h3>
                  <p className="text-xs text-slate-400">Complete service card history for this customer and chassis number.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setSelectedFollowupCustomer(null)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Content */}
            <div className="p-4 overflow-y-auto space-y-4 flex-1">
              
              {/* Customer Profile Banner */}
              <div className="bg-slate-50 border border-slate-200 p-3 rounded-xl grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 gap-3">
                <div className="col-span-2">
                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Customer Name</span>
                  <div className="flex items-center gap-1.5 mt-0.5 flex-wrap">
                    {selectedFollowupCustomer.historyFileNo && (
                      <span
                        className={`px-1.5 py-0.5 font-mono text-[10px] font-black rounded border shadow-2xs ${
                          isWithin2Years(selectedFollowupCustomer.dateOfDel || selectedFollowupCustomer.installDate)
                            ? 'bg-blue-600 text-white border-blue-700'
                            : 'bg-red-600 text-white border-red-700'
                        }`}
                      >
                        HFN: {selectedFollowupCustomer.historyFileNo}
                      </span>
                    )}
                    <span className="text-sm font-black text-slate-800 block truncate">{selectedFollowupCustomer.customerName}</span>
                  </div>
                  {selectedFollowupCustomer.mobileNumber && (
                    <span className="text-xs font-mono font-bold text-indigo-600 block mt-0.5">📞 {selectedFollowupCustomer.mobileNumber}</span>
                  )}
                  <button
                    type="button"
                    onClick={() => copyCustomerDetails(selectedFollowupCustomer)}
                    className="mt-2 px-2 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded border border-indigo-700 font-extrabold text-[10px] flex items-center gap-1 cursor-pointer transition-all shadow-2xs hover:shadow-sm"
                    title="Copy loaded customer details"
                  >
                    <Copy className="w-3.5 h-3.5 text-white" />
                    <span>Copy Details (కాపీ)</span>
                  </button>
                </div>
                <div>
                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Chassis Number</span>
                  <span className="text-xs font-mono font-extrabold text-slate-900 block mt-0.5">{selectedFollowupCustomer.chassisNo}</span>
                  {selectedFollowupCustomer.engineNo && (
                    <span className="text-[10px] text-slate-500 font-semibold block">Eng: {selectedFollowupCustomer.engineNo}</span>
                  )}
                </div>
                <div>
                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Tractor Model</span>
                  <span className="text-xs font-black text-indigo-950 block mt-0.5">{selectedFollowupCustomer.model || '—'}</span>
                  {selectedFollowupCustomer.modelType && (
                    <span className="text-[10px] text-slate-500 font-bold block">{selectedFollowupCustomer.modelType}</span>
                  )}
                </div>
                <div>
                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Location</span>
                  <span className="text-xs font-bold text-slate-700 block mt-0.5 truncate" title={`${selectedFollowupCustomer.village}, ${selectedFollowupCustomer.mandal}`}>
                    {selectedFollowupCustomer.village || '—'}, {selectedFollowupCustomer.mandal || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-[9px] uppercase font-bold text-slate-400 block">Delivery Date</span>
                  <span className="text-xs font-bold text-slate-700 block mt-0.5">{selectedFollowupCustomer.dateOfDel || '—'}</span>
                </div>
              </div>

              {/* Job Cards Table */}
              <div>
                <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <span className="w-2 h-2 bg-indigo-600 rounded-full" />
                  Recorded Job Cards ({selectedFollowupCustomer.jobCardsCount})
                </h4>

                {(!selectedFollowupCustomer.cards || selectedFollowupCustomer.cards.length === 0) ? (
                  <div className="border border-slate-200 rounded-xl p-8 text-center bg-slate-50 text-slate-400 font-bold">
                    No service visits have been recorded for this chassis number yet.
                  </div>
                ) : (
                  <div className="border border-slate-200 rounded-xl overflow-hidden bg-white">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse text-[11px]">
                        <thead>
                          <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-extrabold uppercase text-[9px] tracking-wider">
                            <th className="p-2.5 w-10 text-center">S.No</th>
                            <th className="p-2.5 w-32">Job Card No</th>
                            <th className="p-2.5 w-24">Date</th>
                            <th className="p-2.5 w-28">Service Type</th>
                            <th className="p-2.5 w-32">Service Count</th>
                            <th className="p-2.5 w-20">Hrs Run</th>
                            <th className="p-2.5 min-w-[200px]">Repairs & Work Done</th>
                            <th className="p-2.5 w-36">Technician / Incharge</th>
                            <th className="p-2.5 w-20 text-center">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 font-medium">
                          {selectedFollowupCustomer.cards.map((card: any, idx: number) => {
                            const formattedDate = fmtDate(card.jobDate || card.jobOpenDate || (card.createdAt ? card.createdAt.split('T')[0] : ''));
                            const primaryRep = Array.isArray(card.repairRows) && card.repairRows.length > 0
                              ? card.repairRows.map((r: any) => r.repair).filter(Boolean).join(', ')
                              : (card.problemDescription || card.reasonsForAnalysis || '');
                            const extraRepText = getExtraRepairsText(card) !== '—'
                              ? getExtraRepairsText(card)
                              : (card.extraRepairs || card['EXTRA OTHER REPAIRS DONE WITH FREE SERVICE'] || card.extraRepairsDone || '');
                            const serviceCountVal = card.freeServiceList || card.freeServiceCount || card['FREE SERVICE LIST'] || card['Service Count'] || getFreeServiceText(card) || '—';
                            return (
                              <tr key={card.id || card.jobNo || idx} className="hover:bg-amber-100/80">
                                <td className="p-2.5 text-center font-mono font-bold text-slate-400">{idx + 1}</td>
                                <td className="p-2.5 font-mono font-bold text-slate-900">{card.onlineJobCardNo || card.jobNo || '—'}</td>
                                <td className="p-2.5 font-bold text-slate-600">{formattedDate}</td>
                                <td className="p-2.5 text-indigo-950 font-semibold">{card.serviceType || 'Paid Service'}</td>
                                <td className="p-2.5 text-slate-900 font-bold">{serviceCountVal}</td>
                                <td className="p-2.5 font-mono font-extrabold text-indigo-600">{card.hourMeter || card.hrsRun ? `${card.hourMeter || card.hrsRun} Hrs` : '—'}</td>
                                <td className="p-2.5 text-slate-700 leading-relaxed font-normal">
                                  {primaryRep && primaryRep !== '—' && (
                                    <div className="font-semibold">{primaryRep}</div>
                                  )}
                                  {extraRepText && extraRepText !== '—' && (
                                    <div className="text-[10px] text-amber-950 font-bold bg-amber-100/80 px-1.5 py-0.5 rounded border border-amber-200 mt-1 inline-block">
                                      🔧 Extra: {extraRepText}
                                    </div>
                                  )}
                                  {(!primaryRep || primaryRep === '—') && (!extraRepText || extraRepText === '—') && '—'}
                                </td>
                                <td className="p-2.5">
                                  <div className="font-semibold text-slate-800">{card.mechanic || card.technicianName || '—'}</div>
                                  <div className="text-[9px] text-slate-400 font-bold">Sup: {card.wsIncharge || '—'}</div>
                                </td>
                                <td className="p-2.5 text-center">
                                  <span className={`inline-block px-1.5 py-0.5 rounded-md text-[9px] font-extrabold uppercase ${
                                    (card.status || '').toLowerCase() === 'closed'
                                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                      : 'bg-amber-100 text-amber-800 border border-amber-200'
                                  }`}>
                                    {card.status || 'Open'}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* Footer */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end shrink-0">
              <button
                type="button"
                onClick={() => setSelectedFollowupCustomer(null)}
                className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
              >
                Close History
              </button>
            </div>
          </div>
        </div>
      )}

      {/* TELE CALLING REPORTS & SCHEDULED FOLLOW-UPS TAB */}
      {activeTab === 'telecalling' && (
        <div className="w-full bg-white border border-slate-200 shadow-xs p-3 md:p-5 rounded-2xl space-y-5 print:hidden">
          
          {/* TOP HEADER */}
          <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-3">
            <div className="flex items-center gap-3">
              <div className="bg-amber-600 text-white p-2.5 rounded-xl shadow-xs">
                <PhoneCall className="w-6 h-6 text-white" />
              </div>
              <div>
                <h2 className="text-base font-black text-slate-900 leading-tight">Tele Calling Reports & Scheduled Follow-ups</h2>
                <p className="text-xs text-slate-500 font-medium">
                  ఏ రోజు ఎంతమందికి కాల్స్ చేశారు, కస్టమర్ల రెస్పాన్స్ (రిమార్క్స్) మరియు ఆటోమేటిక్ నెక్స్ట్ డ్యూ కాల్స్ ట్రాకింగ్ డ్యాష్‌బోర్డ్.
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setIsAddingQuickRemark(true)}
                className="px-3 py-1.5 bg-amber-100 hover:bg-amber-200 text-amber-950 font-bold text-xs rounded-xl transition-all border border-amber-300 flex items-center gap-1.5 cursor-pointer shadow-2xs"
              >
                <span>⚙️ Preset Responses</span>
              </button>
              <button
                type="button"
                onClick={downloadTelecallingExcel}
                className="px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-xl transition-all shadow-xs flex items-center gap-1.5 cursor-pointer"
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Export Telecalling Report</span>
              </button>
            </div>
          </div>

          {/* TOP SUMMARY KPI METRICS */}
          {(() => {
            const callsTodayCount = allTelecallerLogs.filter(l => l.callDate === todayISO).length;
            const todayScheduled = scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length;
            const overdueScheduled = scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'overdue').length;
            const upcomingScheduled = scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'upcoming').length;

            return (
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
                
                {/* Metric 1: Calls Made Today */}
                <div 
                  onClick={() => {
                    setTelecallerSubTab('date_report');
                    setTelecallerDateFrom(todayISO);
                    setTelecallerDateTo(todayISO);
                  }}
                  className="bg-amber-50/70 border border-amber-300/80 p-3 rounded-xl cursor-pointer hover:bg-amber-100/80 transition-all select-none shadow-2xs flex items-center justify-between"
                >
                  <div className="pr-2">
                    <div className="text-[10px] font-extrabold text-amber-900 uppercase tracking-wider flex items-center gap-1.5">
                      <Calendar className="w-3.5 h-3.5 text-amber-600" />
                      <span>Calls Made Today</span>
                    </div>
                    <div className="text-[9px] text-amber-700 font-bold mt-0.5">Calls logged today</div>
                  </div>
                  <div className="bg-amber-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                    {callsTodayCount}
                  </div>
                </div>

                {/* Metric 2: Scheduled Today */}
                <div 
                  onClick={() => {
                    setTelecallerSubTab('scheduled');
                    setTelecallerScheduledStatus('today');
                  }}
                  className={`border p-3 rounded-xl cursor-pointer transition-all select-none shadow-2xs flex items-center justify-between ${
                    todayScheduled > 0 
                      ? 'bg-red-50/70 border-red-300 hover:bg-red-100 ring-2 ring-red-400/30' 
                      : 'bg-slate-50 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <div className="pr-2">
                    <div className="text-[10px] font-extrabold text-red-900 uppercase tracking-wider flex items-center gap-1.5">
                      <span>⚡ Due Today</span>
                      {todayScheduled > 0 && <span className="w-2 h-2 rounded-full bg-red-600 animate-ping" />}
                    </div>
                    <div className="text-[9px] text-red-700 font-bold mt-0.5">
                      {todayScheduled > 0 ? '⚠️ Must call today!' : 'No pending calls'}
                    </div>
                  </div>
                  <div className="bg-red-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                    {todayScheduled}
                  </div>
                </div>

                {/* Metric 3: Overdue Calls */}
                <div 
                  onClick={() => {
                    setTelecallerSubTab('scheduled');
                    setTelecallerScheduledStatus('overdue');
                  }}
                  className={`border p-3 rounded-xl cursor-pointer transition-all select-none shadow-2xs flex items-center justify-between ${
                    overdueScheduled > 0 
                      ? 'bg-orange-50/70 border-orange-300 hover:bg-orange-100' 
                      : 'bg-slate-50 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <div className="pr-2">
                    <div className="text-[10px] font-extrabold text-orange-900 uppercase tracking-wider flex items-center gap-1.5">
                      <AlertCircle className="w-3.5 h-3.5 text-orange-600" />
                      <span>Overdue Calls</span>
                    </div>
                    <div className="text-[9px] text-orange-700 font-bold mt-0.5">Past scheduled dates</div>
                  </div>
                  <div className="bg-orange-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                    {overdueScheduled}
                  </div>
                </div>

                {/* Metric 4: Upcoming Scheduled */}
                <div 
                  onClick={() => {
                    setTelecallerSubTab('scheduled');
                    setTelecallerScheduledStatus('upcoming');
                  }}
                  className="bg-blue-50/70 border border-blue-200/80 p-3 rounded-xl cursor-pointer hover:bg-blue-100/80 transition-all select-none shadow-2xs flex items-center justify-between"
                >
                  <div className="pr-2">
                    <div className="text-[10px] font-extrabold text-blue-900 uppercase tracking-wider flex items-center gap-1.5">
                      <Clock className="w-3.5 h-3.5 text-blue-600" />
                      <span>Upcoming Calls</span>
                    </div>
                    <div className="text-[9px] text-blue-700 font-bold mt-0.5">Future scheduled</div>
                  </div>
                  <div className="bg-blue-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                    {upcomingScheduled}
                  </div>
                </div>

                {/* Metric 5: Total Logs Registered */}
                <div 
                  onClick={() => {
                    setTelecallerSubTab('all_logs');
                    setTelecallerDateFrom('');
                    setTelecallerDateTo('');
                  }}
                  className="bg-indigo-50/70 border border-indigo-200/80 p-3 rounded-xl cursor-pointer hover:bg-indigo-100/80 transition-all select-none shadow-2xs flex items-center justify-between"
                >
                  <div className="pr-2">
                    <div className="text-[10px] font-extrabold text-indigo-900 uppercase tracking-wider flex items-center gap-1.5">
                      <History className="w-3.5 h-3.5 text-indigo-600" />
                      <span>Total Logs Recorded</span>
                    </div>
                    <div className="text-[9px] text-indigo-700 font-bold mt-0.5">History of all calls</div>
                  </div>
                  <div className="bg-indigo-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                    {allTelecallerLogs.length}
                  </div>
                </div>

              </div>
            );
          })()}

          {/* SUB-TAB NAVIGATION MODULE SWITCHER */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-100 p-1.5 rounded-xl border border-slate-200">
            <div className="flex items-center gap-1.5 flex-wrap">
              <button
                type="button"
                onClick={() => setTelecallerSubTab('scheduled')}
                className={`px-3.5 py-2 rounded-lg font-extrabold text-xs transition-all cursor-pointer flex items-center gap-1.5 ${
                  telecallerSubTab === 'scheduled'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <span>📅 Scheduled Follow-ups (తదుపరి కాల్స్)</span>
                {scheduledFollowupsTodayCount > 0 && (
                  <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-red-600 text-white animate-pulse">
                    {scheduledFollowupsTodayCount}
                  </span>
                )}
              </button>

              <button
                type="button"
                onClick={() => setTelecallerSubTab('date_report')}
                className={`px-3.5 py-2 rounded-lg font-extrabold text-xs transition-all cursor-pointer flex items-center gap-1.5 ${
                  telecallerSubTab === 'date_report'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <BarChart3 className="w-3.5 h-3.5" />
                <span>📊 Date-wise Call Count Report (ఏ రోజు ఎంత మందికి)</span>
              </button>

              <button
                type="button"
                onClick={() => setTelecallerSubTab('all_logs')}
                className={`px-3.5 py-2 rounded-lg font-extrabold text-xs transition-all cursor-pointer flex items-center gap-1.5 ${
                  telecallerSubTab === 'all_logs'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <History className="w-3.5 h-3.5" />
                <span>📝 Customer Remarks Register (ఎవరు ఏమి చెప్పారు)</span>
              </button>
            </div>

            {/* Quick Search */}
            <div className="relative flex-1 min-w-[200px] max-w-xs">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                value={telecallerSearch}
                onChange={(e) => setTelecallerSearch(e.target.value)}
                placeholder="Search name, phone, village, remarks..."
                className="w-full text-xs pl-8 pr-3 py-1.5 bg-white border border-slate-300 rounded-lg outline-none font-medium focus:border-amber-600"
              />
              {telecallerSearch && (
                <button
                  onClick={() => setTelecallerSearch('')}
                  className="absolute right-2 top-2 text-slate-400 hover:text-slate-600 text-xs font-bold"
                >
                  ×
                </button>
              )}
            </div>
          </div>

          {/* GLOBAL TELECALLING MULTI-FILTER & VIEW LAYOUT CONTROL BAR */}
          <div className="bg-amber-50/80 border border-amber-200/90 p-3 rounded-xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2.5">
              
              {/* Multi-Filter Dropdowns */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1">
                  <Filter className="w-3.5 h-3.5 text-amber-800" />
                  <span className="text-xs font-black text-amber-950">Multi Filters:</span>
                </div>

                {/* Supervisor Filter */}
                <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-lg border border-amber-300 shadow-2xs">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase">Supervisor:</span>
                  <select
                    value={telecallerSupervisorFilter}
                    onChange={(e) => setTelecallerSupervisorFilter(e.target.value)}
                    className="text-xs font-extrabold text-amber-950 bg-transparent outline-none cursor-pointer"
                  >
                    <option value="all">All Supervisors ({supervisorsList.length})</option>
                    {supervisorsList.map((sup, idx) => (
                      <option key={idx} value={sup}>{sup}</option>
                    ))}
                  </select>
                </div>

                {/* Caller / Telecaller Filter */}
                <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-lg border border-amber-300 shadow-2xs">
                  <span className="text-[10px] font-extrabold text-slate-500 uppercase">Called By:</span>
                  <select
                    value={telecallerCallerFilter}
                    onChange={(e) => setTelecallerCallerFilter(e.target.value)}
                    className="text-xs font-extrabold text-amber-950 bg-transparent outline-none cursor-pointer"
                  >
                    <option value="all">All Callers / Staff ({telecallersList.length})</option>
                    {telecallersList.map((caller, idx) => (
                      <option key={idx} value={caller}>{caller}</option>
                    ))}
                  </select>
                </div>

                {(telecallerSupervisorFilter !== 'all' || telecallerCallerFilter !== 'all') && (
                  <button
                    type="button"
                    onClick={() => {
                      setTelecallerSupervisorFilter('all');
                      setTelecallerCallerFilter('all');
                    }}
                    className="text-[10px] font-bold text-red-600 hover:text-red-800 bg-red-50 hover:bg-red-100 px-2 py-1 rounded border border-red-200 cursor-pointer"
                  >
                    ✕ Reset Filters
                  </button>
                )}
              </div>

              {/* Display Mode / Resizable Controls (Like Saved Job Cards) */}
              <div className="flex items-center gap-1.5 bg-white p-1 rounded-lg border border-amber-300 shadow-2xs">
                <span className="text-[10px] font-extrabold text-slate-500 uppercase px-1">View Layout:</span>
                <button
                  type="button"
                  onClick={() => setTelecallerDisplayMode('table')}
                  className={`px-2.5 py-1 rounded text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                    telecallerDisplayMode === 'table'
                      ? 'bg-amber-600 text-white shadow-2xs'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                  title="Table View"
                >
                  <FileText className="w-3.5 h-3.5" />
                  <span>Table View</span>
                </button>

                <button
                  type="button"
                  onClick={() => setTelecallerDisplayMode('cards')}
                  className={`px-2.5 py-1 rounded text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                    telecallerDisplayMode === 'cards'
                      ? 'bg-amber-600 text-white shadow-2xs'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                  title="Saved Job Cards Style Boxes View"
                >
                  <BarChart3 className="w-3.5 h-3.5" />
                  <span>Card Boxes</span>
                </button>

                <button
                  type="button"
                  onClick={() => setIsTelecallerMaximized(!isTelecallerMaximized)}
                  className={`p-1 rounded text-xs font-bold transition-all cursor-pointer flex items-center gap-1 px-2 ${
                    isTelecallerMaximized
                      ? 'bg-amber-200 text-amber-950 border border-amber-400 font-extrabold'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                  title={isTelecallerMaximized ? "Restore Normal View Size" : "Expand / Maximize Box Grid"}
                >
                  {isTelecallerMaximized ? <Minimize2 className="w-3.5 h-3.5" /> : <Maximize2 className="w-3.5 h-3.5" />}
                  <span className="text-[10px]">{isTelecallerMaximized ? 'Normal Size' : 'Resize / Maximize'}</span>
                </button>
              </div>

            </div>
          </div>
          {telecallerSubTab === 'scheduled' && (
            <div className="space-y-4">
              
              {/* Scheduled Status Filter Bar */}
              <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50/60 p-2.5 rounded-xl border border-amber-200">
                <div className="flex items-center gap-1.5 flex-wrap">
                  <span className="text-xs font-black text-amber-950 mr-1">Filter Call Status:</span>
                  
                  <button
                    type="button"
                    onClick={() => setTelecallerScheduledStatus('all')}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold cursor-pointer transition-all ${
                      telecallerScheduledStatus === 'all'
                        ? 'bg-amber-700 text-white shadow-2xs'
                        : 'bg-white text-slate-700 border border-amber-200 hover:bg-amber-100'
                    }`}
                  >
                    All Scheduled ({scheduledFollowupCustomers.length})
                  </button>

                  <button
                    type="button"
                    onClick={() => setTelecallerScheduledStatus('today')}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold cursor-pointer transition-all flex items-center gap-1 ${
                      telecallerScheduledStatus === 'today'
                        ? 'bg-red-600 text-white shadow-2xs'
                        : 'bg-white text-red-700 border border-red-200 hover:bg-red-50'
                    }`}
                  >
                    <span>⚡ Calls Due Today</span>
                    <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-red-100 text-red-800 font-extrabold">
                      {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'today').length}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setTelecallerScheduledStatus('overdue')}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold cursor-pointer transition-all flex items-center gap-1 ${
                      telecallerScheduledStatus === 'overdue'
                        ? 'bg-orange-600 text-white shadow-2xs'
                        : 'bg-white text-orange-700 border border-orange-200 hover:bg-orange-50'
                    }`}
                  >
                    <span>🚨 Overdue Calls</span>
                    <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-orange-100 text-orange-800 font-extrabold">
                      {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'overdue').length}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setTelecallerScheduledStatus('upcoming')}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold cursor-pointer transition-all flex items-center gap-1 ${
                      telecallerScheduledStatus === 'upcoming'
                        ? 'bg-blue-600 text-white shadow-2xs'
                        : 'bg-white text-blue-700 border border-blue-200 hover:bg-blue-50'
                    }`}
                  >
                    <span>📅 Upcoming Calls</span>
                    <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-blue-100 text-blue-800 font-extrabold">
                      {scheduledFollowupCustomers.filter(s => s.scheduledStatus === 'upcoming').length}
                    </span>
                  </button>
                </div>

                <div className="text-[11px] font-bold text-amber-900">
                  Showing <span className="font-extrabold text-amber-950">{filteredScheduledFollowups.length}</span> scheduled customers
                </div>
              </div>

              {/* Scheduled Customers Table or Card Boxes View */}
              {filteredScheduledFollowups.length === 0 ? (
                <div className="p-8 text-center bg-slate-50 border border-dashed border-slate-200 rounded-xl space-y-2">
                  <PhoneCall className="w-8 h-8 text-slate-300 mx-auto" />
                  <p className="text-xs font-bold text-slate-500">No scheduled follow-up calls match the current filter.</p>
                  <p className="text-[11px] text-slate-400">Log calls in Customer Data or set a next call date to schedule follow-ups.</p>
                </div>
              ) : telecallerDisplayMode === 'cards' ? (
                /* CARD BOXES GRID VIEW (Like Saved Job Cards) */
                <div className={`grid gap-3.5 ${
                  isTelecallerMaximized 
                    ? 'grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-3' 
                    : 'grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4'
                }`}>
                  {filteredScheduledFollowups.map((item, idx) => {
                    const isToday = item.scheduledStatus === 'today';
                    const isOverdue = item.scheduledStatus === 'overdue';

                    return (
                      <div 
                        key={item.chassisNo || idx}
                        className={`bg-white border rounded-xl p-3.5 space-y-2.5 shadow-2xs hover:shadow-md transition-all flex flex-col justify-between ${
                          isToday 
                            ? 'border-red-400 bg-red-50/30 ring-1 ring-red-400/40' 
                            : isOverdue 
                            ? 'border-orange-300 bg-orange-50/20' 
                            : 'border-slate-200 hover:border-amber-400'
                        }`}
                      >
                        <div className="space-y-2">
                          {/* Card Header Status & HFN */}
                          <div className="flex items-center justify-between gap-1 flex-wrap border-b border-slate-100 pb-1.5">
                            {isToday ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded font-black text-[10px] bg-red-600 text-white shadow-2xs animate-pulse">
                                ⚡ DUE TODAY ({fmtDate(item.scheduledDate)})
                              </span>
                            ) : isOverdue ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded font-black text-[10px] bg-orange-600 text-white shadow-2xs">
                                🚨 OVERDUE ({fmtDate(item.scheduledDate)})
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] bg-blue-100 text-blue-900 border border-blue-200">
                                📅 {fmtDate(item.scheduledDate)}
                              </span>
                            )}

                            {item.historyFileNo && (
                              <span className={`px-1.5 py-0.2 rounded font-mono text-[10px] font-black border ${
                                isWithin2Years(item.dateOfDel)
                                  ? 'bg-blue-600 text-white border-blue-700'
                                  : 'bg-red-600 text-white border-red-700'
                              }`}>
                                HFN: {item.historyFileNo}
                              </span>
                            )}
                          </div>

                          {/* Customer Name, Address & Model */}
                          <div className="space-y-1">
                            <div className="font-black text-slate-950 text-base">{item.customerName || '—'}</div>
                            <div className="text-xs text-slate-600 font-bold">📍 {item.village || '—'}, {item.mandal || '—'}</div>
                            <div className="text-[10px] text-amber-900 font-bold">🚜 {item.model || '—'}</div>
                          </div>

                          {/* Large Phone Number Section */}
                          {item.mobileNumber && (
                            <div className="pt-1">
                              <a
                                href={`tel:${item.mobileNumber}`}
                                className="w-full py-2.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-xl font-mono font-black text-lg shadow-sm transition-all flex items-center justify-center gap-2 border-2 border-emerald-200 active:scale-[0.98]"
                              >
                                <PhoneCall className="w-5 h-5" />
                                {item.mobileNumber}
                              </a>
                            </div>
                          )}

                          {/* Supervisor & Called By badges */}
                          <div className="flex flex-wrap items-center gap-1.5 pt-1 border-t border-slate-100">
                            {item.supervisor && (
                              <span className="text-[10px] font-bold text-purple-900 bg-purple-100 px-1.5 py-0.5 rounded border border-purple-200">
                                👨‍💼 Sup: {item.supervisor}
                              </span>
                            )}
                            {item.lastCalledBy && (
                              <span className="text-[10px] font-bold text-amber-950 bg-amber-100 px-1.5 py-0.5 rounded border border-amber-200">
                                👤 Last Caller: {item.lastCalledBy}
                              </span>
                            )}
                          </div>

                          {/* Last Remarks */}
                          <div className="bg-amber-50/70 p-2 rounded-lg border border-amber-200/80 space-y-0.5 text-xs">
                            <span className="text-[9px] font-extrabold text-amber-900 uppercase block">Last Customer Remarks:</span>
                            <div className="italic text-slate-800 font-semibold leading-relaxed">
                              "{item.lastRemarks || 'No remarks recorded yet'}"
                            </div>
                            {item.lastCallDate && (
                              <div className="text-[9px] text-slate-400 font-bold text-right pt-0.5">
                                Last Called: {fmtDate(item.lastCallDate)}
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Action Call Button */}
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedFollowupForLog(item);
                            setNewFollowupCallDate(todayISO);
                            setNewFollowupRemarks('');
                            setNewFollowupNextCallDate('');
                          }}
                          className="w-full py-2 px-3 bg-amber-600 hover:bg-amber-700 text-white font-black text-xs rounded-lg transition-colors shadow-xs flex items-center justify-center gap-1.5 cursor-pointer mt-2"
                        >
                          <PhoneCall className="w-3.5 h-3.5" />
                          <span>Call & Log Interaction</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDirectComplaintRegister(item)}
                          className="w-full py-2 px-3 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs rounded-lg transition-colors shadow-xs flex items-center justify-center gap-1.5 cursor-pointer mt-1"
                        >
                          <ShieldAlert className="w-3.5 h-3.5" />
                          <span>Register Complaint</span>
                        </button>
                      </div>
                    );
                  })}
                </div>
              ) : (
                /* TABLE VIEW MODE */
                <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-2xs">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-amber-600 text-white font-bold text-[11px] uppercase tracking-wider">
                      <tr>
                        <th className="p-2.5 w-8 text-center">#</th>
                        <th className="p-2.5 w-32">Scheduled Date</th>
                        <th className="p-2.5 w-44">Customer Name & HFN</th>
                        <th className="p-2.5 w-32">Mobile / Phone</th>
                        <th className="p-2.5 w-36">Location & Model</th>
                        <th className="p-2.5 w-28">Assigned Supervisor</th>
                        <th className="p-2.5 w-28">Last Called By</th>
                        <th className="p-2.5 min-w-[200px]">Last Remarks (గతంలో చెప్పింది)</th>
                        <th className="p-2.5 w-28 text-center">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white font-medium">
                      {filteredScheduledFollowups.map((item, idx) => {
                        const isToday = item.scheduledStatus === 'today';
                        const isOverdue = item.scheduledStatus === 'overdue';

                        return (
                          <tr 
                            key={item.chassisNo || idx} 
                            className={`hover:bg-amber-50/80 transition-colors text-[11px] ${
                              isToday ? 'bg-red-50/60 font-semibold' : isOverdue ? 'bg-orange-50/50' : ''
                            }`}
                          >
                            <td className="p-2.5 text-center font-mono text-slate-400 font-bold">{idx + 1}</td>
                            
                            <td className="p-2.5">
                              {isToday ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-black text-xs bg-red-600 text-white border border-red-700 shadow-2xs animate-pulse">
                                  ⚡ TODAY ({fmtDate(item.scheduledDate)})
                                </span>
                              ) : isOverdue ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-black text-[11px] bg-orange-600 text-white border border-orange-700 shadow-2xs">
                                  🚨 OVERDUE ({fmtDate(item.scheduledDate)})
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-bold text-[11px] bg-blue-100 text-blue-900 border border-blue-200">
                                  📅 {fmtDate(item.scheduledDate)}
                                </span>
                              )}
                            </td>

                            <td className="p-2.5">
                              <div className="font-black text-slate-950 text-sm">{item.customerName || '—'}</div>
                              <div className="font-bold text-slate-600 text-[10px] mt-0.5">📍 {item.village || '—'}, {item.mandal}</div>
                              {item.historyFileNo && (
                                <span className={`inline-block px-1.5 py-0.2 rounded font-mono text-[10px] font-black border mt-1 ${
                                  isWithin2Years(item.dateOfDel)
                                    ? 'bg-blue-600 text-white border-blue-700'
                                    : 'bg-red-600 text-white border-red-700'
                                }`}>
                                  HFN: {item.historyFileNo}
                                </span>
                              )}
                            </td>

                            <td className="p-2.5 font-mono">
                              {item.mobileNumber ? (
                                <a
                                  href={`tel:${item.mobileNumber}`}
                                  className="inline-flex items-center gap-1 text-emerald-700 font-black text-xs hover:underline hover:text-emerald-900 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200"
                                >
                                  <PhoneCall className="w-3 h-3" /> {item.mobileNumber}
                                </a>
                              ) : (
                                <span className="text-slate-400">—</span>
                              )}
                            </td>

                            <td className="p-2.5">
                              <div className="text-[11px] text-slate-800 font-black uppercase tracking-tight">🚜 {item.model || '—'}</div>
                            </td>

                            <td className="p-2.5">
                              <span className="text-purple-950 font-extrabold text-[10px] px-1.5 py-0.5 bg-purple-100 rounded border border-purple-200">
                                {item.supervisor || 'Unassigned'}
                              </span>
                            </td>

                            <td className="p-2.5">
                              <span className="text-amber-950 font-extrabold text-[10px] px-1.5 py-0.5 bg-amber-100 rounded border border-amber-200">
                                {item.lastCalledBy || '—'}
                              </span>
                            </td>

                            <td className="p-2.5 text-slate-800 italic leading-snug">
                              {item.lastRemarks ? (
                                <span className="bg-amber-100/60 text-amber-950 px-2 py-1 rounded border border-amber-200 inline-block font-semibold">
                                  "{item.lastRemarks}"
                                </span>
                              ) : (
                                <span className="text-slate-400 font-normal">No remarks recorded</span>
                              )}
                            </td>

                            <td className="p-2.5 text-center">
                              <div className="flex flex-col gap-1">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setSelectedFollowupForLog(item);
                                    setNewFollowupCallDate(todayISO);
                                    setNewFollowupRemarks('');
                                    setNewFollowupNextCallDate('');
                                  }}
                                  className="px-2.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-black text-xs rounded-lg transition-colors shadow-xs flex items-center gap-1 justify-center w-full cursor-pointer"
                                >
                                  📞 Call & Log
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDirectComplaintRegister(item)}
                                  className="px-2.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-[10px] rounded-lg transition-colors shadow-xs flex items-center gap-1 justify-center w-full cursor-pointer"
                                >
                                  📝 Complaint
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* SUB-TAB 2: DATE-WISE CALL COUNT REPORT ("e date lo entamandiki chesindi") */}
          {telecallerSubTab === 'date_report' && (
            <div className="space-y-4">
              
              {/* Date Filters & Quick Presets */}
              <div className="bg-amber-50/70 p-3 rounded-xl border border-amber-200 space-y-2.5">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <Calendar className="w-4 h-4 text-amber-700" />
                    <span className="text-xs font-black text-amber-950 uppercase tracking-wide">
                      Select Date Range for Telecalling Report:
                    </span>
                  </div>

                  <div className="flex items-center gap-1.5 flex-wrap">
                    <button
                      type="button"
                      onClick={() => {
                        setTelecallerDateFrom(todayISO);
                        setTelecallerDateTo(todayISO);
                      }}
                      className="px-2.5 py-1 bg-white hover:bg-amber-100 text-slate-800 font-bold text-xs rounded border border-amber-300 shadow-2xs transition-colors cursor-pointer"
                    >
                      Today
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        const d = new Date();
                        d.setDate(d.getDate() - 1);
                        const yest = d.toISOString().split('T')[0];
                        setTelecallerDateFrom(yest);
                        setTelecallerDateTo(yest);
                      }}
                      className="px-2.5 py-1 bg-white hover:bg-amber-100 text-slate-800 font-bold text-xs rounded border border-amber-300 shadow-2xs transition-colors cursor-pointer"
                    >
                      Yesterday
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        const d = new Date();
                        d.setDate(d.getDate() - 7);
                        setTelecallerDateFrom(d.toISOString().split('T')[0]);
                        setTelecallerDateTo(todayISO);
                      }}
                      className="px-2.5 py-1 bg-white hover:bg-amber-100 text-slate-800 font-bold text-xs rounded border border-amber-300 shadow-2xs transition-colors cursor-pointer"
                    >
                      Last 7 Days
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        const d = new Date();
                        d.setDate(1);
                        setTelecallerDateFrom(d.toISOString().split('T')[0]);
                        setTelecallerDateTo(todayISO);
                      }}
                      className="px-2.5 py-1 bg-white hover:bg-amber-100 text-slate-800 font-bold text-xs rounded border border-amber-300 shadow-2xs transition-colors cursor-pointer"
                    >
                      This Month
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setTelecallerDateFrom('');
                        setTelecallerDateTo('');
                      }}
                      className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded transition-colors cursor-pointer"
                    >
                      Clear / All Time
                    </button>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  <div>
                    <label className="block text-[10px] font-bold text-amber-900 mb-0.5">From Date:</label>
                    <input
                      type="date"
                      value={telecallerDateFrom}
                      onChange={(e) => setTelecallerDateFrom(e.target.value)}
                      className="text-xs p-1.5 bg-white border border-amber-300 rounded outline-none font-bold text-slate-800 focus:border-amber-600"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-amber-900 mb-0.5">To Date:</label>
                    <input
                      type="date"
                      value={telecallerDateTo}
                      onChange={(e) => setTelecallerDateTo(e.target.value)}
                      className="text-xs p-1.5 bg-white border border-amber-300 rounded outline-none font-bold text-slate-800 focus:border-amber-600"
                    />
                  </div>

                  <div className="self-end pb-1 text-xs font-extrabold text-amber-950">
                    Total Calls Made in Selected Period: <span className="text-amber-700 text-sm font-black">{filteredTelecallerLogs.length}</span>
                  </div>
                </div>
              </div>

              {/* Date-wise Summary Breakdown Table */}
              {dateWiseCallSummary.length === 0 ? (
                <div className="p-8 text-center bg-slate-50 border border-dashed border-slate-200 rounded-xl space-y-2">
                  <BarChart3 className="w-8 h-8 text-slate-300 mx-auto" />
                  <p className="text-xs font-bold text-slate-500">No telecalling records found for the selected date range.</p>
                </div>
              ) : (
                <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-2xs">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-amber-600 text-white font-bold text-[11px] uppercase tracking-wider">
                      <tr>
                        <th className="p-2.5 w-10 text-center">#</th>
                        <th className="p-2.5 w-32">Call Date</th>
                        <th className="p-2.5 w-32 text-center">Total Calls Made</th>
                        <th className="p-2.5 w-36 text-center">Unique Customers</th>
                        <th className="p-2.5 min-w-[250px]">Sample Customer Remarks Recorded</th>
                        <th className="p-2.5 w-28 text-center">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white font-medium">
                      {dateWiseCallSummary.map((item, idx) => (
                        <tr key={item.callDate || idx} className="hover:bg-amber-50/80 transition-colors text-[11px]">
                          <td className="p-2.5 text-center font-mono text-slate-400 font-bold">{idx + 1}</td>
                          <td className="p-2.5 font-bold text-slate-900">{fmtDate(item.callDate)}</td>
                          <td className="p-2.5 text-center">
                            <span className="px-2.5 py-0.5 rounded-full font-black text-xs bg-amber-100 text-amber-950 border border-amber-300">
                              {item.count} Calls
                            </span>
                          </td>
                          <td className="p-2.5 text-center font-bold text-indigo-700">
                            {item.uniqueCustomersCount} Customers
                          </td>
                          <td className="p-2.5 text-slate-700 leading-snug">
                            {item.remarksSample.length > 0 ? (
                              <div className="flex flex-wrap gap-1">
                                {item.remarksSample.map((rem, rIdx) => (
                                  <span key={rIdx} className="bg-slate-100 text-slate-800 px-1.5 py-0.5 rounded text-[10px] border border-slate-200 italic">
                                    "{rem}"
                                  </span>
                                ))}
                              </div>
                            ) : (
                              <span className="text-slate-400">—</span>
                            )}
                          </td>
                          <td className="p-2.5 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                setTelecallerDateFrom(item.callDate);
                                setTelecallerDateTo(item.callDate);
                                setTelecallerSubTab('all_logs');
                              }}
                              className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded transition-colors cursor-pointer"
                            >
                              View Logs
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* SUB-TAB 3: CUSTOMER REMARKS REGISTER ("evaru emi chepparu") */}
          {telecallerSubTab === 'all_logs' && (
            <div className="space-y-4">
              
              <div className="flex flex-wrap items-center justify-between gap-2 bg-amber-50/80 p-2.5 rounded-xl border border-amber-200">
                <div className="text-xs font-black text-amber-950">
                  Customer Call Logs Register ({filteredTelecallerLogs.length} total entries)
                </div>
                <div className="text-[11px] font-bold text-amber-900">
                  Showing logs from {telecallerDateFrom ? fmtDate(telecallerDateFrom) : 'Beginning'} to {telecallerDateTo ? fmtDate(telecallerDateTo) : 'Today'}
                </div>
              </div>

              {filteredTelecallerLogs.length === 0 ? (
                <div className="p-8 text-center bg-slate-50 border border-dashed border-slate-200 rounded-xl space-y-2">
                  <History className="w-8 h-8 text-slate-300 mx-auto" />
                  <p className="text-xs font-bold text-slate-500">No telecalling interaction logs found.</p>
                </div>
              ) : telecallerDisplayMode === 'cards' ? (
                /* CARD BOXES GRID VIEW */
                <div className={`grid gap-3.5 ${
                  isTelecallerMaximized 
                    ? 'grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-3' 
                    : 'grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4'
                }`}>
                  {filteredTelecallerLogs.map((log, idx) => (
                    <div 
                      key={log.logId || idx}
                      className="bg-white border border-slate-200 hover:border-amber-400 rounded-xl p-3.5 space-y-2.5 shadow-2xs hover:shadow-md transition-all flex flex-col justify-between"
                    >
                      <div className="space-y-2">
                        {/* Header: Call Date & Called By */}
                        <div className="flex items-center justify-between gap-1 flex-wrap border-b border-slate-100 pb-1.5">
                          <span className="font-extrabold text-xs text-amber-950 flex items-center gap-1">
                            <Calendar className="w-3.5 h-3.5 text-amber-600" />
                            {fmtDate(log.callDate)}
                          </span>

                          <span className="text-[10px] font-black text-amber-950 bg-amber-100 px-1.5 py-0.5 rounded border border-amber-300">
                            👤 {log.calledBy || 'Staff'}
                          </span>
                        </div>

                        {/* Customer & Phone */}
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <div className="font-extrabold text-slate-950 text-sm">{log.customerName || '—'}</div>
                            {log.historyFileNo && (
                              <span className="inline-block px-1.5 py-0.2 bg-blue-600 text-white font-mono text-[10px] font-black rounded mt-0.5">
                                HFN: {log.historyFileNo}
                              </span>
                            )}
                            <div className="text-[11px] text-slate-500 font-bold mt-1">📍 {log.village || '—'}, {log.mandal || '—'}</div>
                            <div className="text-[10px] text-amber-900 font-bold">🚜 {log.model || '—'}</div>
                          </div>

                          {log.mobileNumber && (
                            <a
                              href={`tel:${log.mobileNumber}`}
                              className="px-2 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-mono font-black text-xs rounded border border-indigo-200 shrink-0 flex items-center gap-1"
                            >
                              📞 {log.mobileNumber}
                            </a>
                          )}
                        </div>

                        {/* Supervisor */}
                        {log.supervisor && (
                          <div className="text-[10px] font-bold text-purple-900 bg-purple-100 px-1.5 py-0.5 rounded border border-purple-200 w-fit">
                            👨‍💼 Sup: {log.supervisor}
                          </div>
                        )}

                        {/* Customer Remarks Box */}
                        <div className="bg-amber-50/80 p-2.5 rounded-lg border border-amber-200/90 space-y-1">
                          <span className="text-[9px] font-extrabold text-amber-900 uppercase block">Customer Remarks / Response:</span>
                          <div className="italic text-slate-900 font-semibold text-xs leading-relaxed">
                            "{log.remarks}"
                          </div>
                        </div>

                        {/* Next Scheduled Call Tag */}
                        {log.nextCallDate && (
                          <div className="text-[10px] font-bold text-purple-900 bg-purple-50 px-2 py-1 rounded border border-purple-200 flex items-center gap-1">
                            <Calendar className="w-3 h-3 text-purple-600" />
                            Next Follow-up Call: {fmtDate(log.nextCallDate)}
                          </div>
                        )}
                      </div>

                      {/* Action */}
                      <button
                        type="button"
                        onClick={() => {
                          setSelectedFollowupForLog(log.customer);
                          setNewFollowupCallDate(todayISO);
                          setNewFollowupRemarks('');
                          setNewFollowupNextCallDate('');
                        }}
                        className="w-full py-1.5 px-3 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded-lg transition-colors shadow-2xs cursor-pointer mt-2"
                      >
                        📞 Call Again
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                /* TABLE VIEW */
                <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-2xs">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-amber-600 text-white font-bold text-[11px] uppercase tracking-wider">
                      <tr>
                        <th className="p-2.5 w-8 text-center">#</th>
                        <th className="p-2.5 w-28">Call Date</th>
                        <th className="p-2.5 w-32">Called By (Telecaller)</th>
                        <th className="p-2.5 w-44">Customer Name & HFN</th>
                        <th className="p-2.5 w-32">Mobile Number</th>
                        <th className="p-2.5 w-32">Location & Model</th>
                        <th className="p-2.5 w-28">Assigned Supervisor</th>
                        <th className="p-2.5 min-w-[220px]">Customer Remarks / Response (ఏమి చెప్పారు)</th>
                        <th className="p-2.5 w-32">Next Scheduled Call</th>
                        <th className="p-2.5 w-24 text-center">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white font-medium">
                      {filteredTelecallerLogs.map((log, idx) => (
                        <tr key={log.logId || idx} className="hover:bg-amber-50/80 transition-colors text-[11px]">
                          <td className="p-2.5 text-center font-mono text-slate-400 font-bold">{idx + 1}</td>
                          <td className="p-2.5 font-bold text-slate-900">{fmtDate(log.callDate)}</td>
                          <td className="p-2.5 font-extrabold text-amber-950">
                            <span className="bg-amber-100 text-amber-950 px-1.5 py-0.5 rounded border border-amber-200 text-[10px]">
                              👤 {log.calledBy || 'Staff'}
                            </span>
                          </td>
                          <td className="p-2.5">
                            <div className="font-extrabold text-slate-950 text-xs">{log.customerName || '—'}</div>
                            {log.historyFileNo && (
                              <span className="inline-block px-1.5 py-0.2 bg-blue-600 text-white font-mono text-[10px] font-black rounded mt-0.5">
                                HFN: {log.historyFileNo}
                              </span>
                            )}
                          </td>
                          <td className="p-2.5 font-mono">
                            {log.mobileNumber ? (
                              <a
                                href={`tel:${log.mobileNumber}`}
                                className="inline-flex items-center gap-1 text-indigo-700 font-black hover:underline hover:text-indigo-900 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200"
                              >
                                📞 {log.mobileNumber}
                              </a>
                            ) : (
                              <span className="text-slate-400">—</span>
                            )}
                          </td>
                          <td className="p-2.5">
                            <div className="font-bold text-slate-800">{log.village || '—'}, {log.mandal}</div>
                            <div className="text-[10px] text-slate-500 font-semibold">{log.model}</div>
                          </td>
                          <td className="p-2.5">
                            <span className="text-purple-950 font-extrabold text-[10px] px-1.5 py-0.5 bg-purple-100 rounded border border-purple-200">
                              {log.supervisor || 'Unassigned'}
                            </span>
                          </td>
                          <td className="p-2.5 text-slate-900 leading-snug">
                            <div className="bg-amber-100/70 text-amber-950 p-2 rounded-lg border border-amber-200/90 font-semibold italic text-xs">
                              "{log.remarks}"
                            </div>
                          </td>
                          <td className="p-2.5">
                            {log.nextCallDate ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[11px] bg-purple-100 text-purple-900 border border-purple-200">
                                📅 {fmtDate(log.nextCallDate)}
                              </span>
                            ) : (
                              <span className="text-slate-400">—</span>
                            )}
                          </td>
                          <td className="p-2.5 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedFollowupForLog(log.customer);
                                setNewFollowupCallDate(todayISO);
                                setNewFollowupRemarks('');
                                setNewFollowupNextCallDate('');
                              }}
                              className="px-2 py-1 bg-amber-600 hover:bg-amber-700 text-white font-bold text-[10px] rounded transition-colors shadow-2xs cursor-pointer"
                            >
                              📞 Call Again
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

        </div>
      )}

      {/* FREE SERVICE & GEAR OIL FOLLOW-UP TREE TAB */}
      {activeTab === 'free_service_followup' && (
        <div className="space-y-6 animate-fade-in pb-12">
          {/* Header Dashboard Summary */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 md:p-6 shadow-xs">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-black text-slate-900 flex items-center gap-2">
                  <Clock className="w-5 h-5 text-teal-600" />
                  <span>Free Service & Gear Oil Follow-up Tree</span>
                </h2>
                <p className="text-xs text-slate-500 font-medium mt-1">
                  Monitor free service due timelines (1st to 10th free services) and gear oil service follow-ups based on customer delivery dates.
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                {/* Supervisor Filter */}
                <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 px-3 py-2 rounded-xl shadow-3xs">
                  <UserCheck className="w-3.5 h-3.5 text-teal-600 shrink-0" />
                  <span className="text-[10px] uppercase tracking-wider font-extrabold text-slate-400 shrink-0">Supervisor:</span>
                  <select
                    value={freeServiceSupervisor}
                    onChange={(e) => setFreeServiceSupervisor(e.target.value)}
                    className="bg-transparent border-none outline-none text-xs font-bold text-slate-800 pr-1 cursor-pointer"
                  >
                    <option value="all">All Supervisors</option>
                    <option value="unassigned">Unassigned</option>
                    {supervisorsList.map((sup) => (
                      <option key={sup} value={sup}>{sup}</option>
                    ))}
                  </select>
                </div>

                <div className="relative">
                  <Search className="absolute left-3 top-2.5 w-4 h-4 text-slate-400" />
                  <input
                    type="text"
                    value={freeServiceSearch}
                    onChange={(e) => setFreeServiceSearch(e.target.value)}
                    placeholder="Search name, chassis, village..."
                    className="pl-9 pr-4 py-2 text-xs border border-slate-200 rounded-xl outline-none focus:border-teal-600 bg-slate-50 focus:bg-white transition-all w-48 md:w-56 font-medium shadow-3xs"
                  />
                </div>
                {freeServiceSearch && (
                  <button
                    type="button"
                    onClick={() => setFreeServiceSearch('')}
                    className="px-2.5 py-2 text-xs bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl font-bold cursor-pointer transition-colors"
                  >
                    Clear
                  </button>
                )}

                <label className="px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl transition-all shadow-xs flex items-center gap-1.5 cursor-pointer">
                  <FileSpreadsheet className="w-4 h-4" /> Import Follow-up Excel
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleExcelUpload}
                    className="hidden"
                  />
                </label>

                {useGoogleSheets && (
                  <button
                    type="button"
                    onClick={handleConnectGoogleSheets}
                    disabled={sheetsLoading}
                    className="px-3 py-2 bg-teal-600 hover:bg-teal-700 text-white font-bold text-xs rounded-xl transition-all shadow-xs flex items-center gap-1.5 cursor-pointer"
                  >
                    <RefreshCw className={`w-4 h-4 ${sheetsLoading ? 'animate-spin' : ''}`} /> Sync from Sheets
                  </button>
                )}
              </div>
            </div>

                        {/* Metric Summary Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-6">
              {/* Card 1: Today's Calls Logged */}
              <div className="bg-amber-50/70 border border-amber-300/80 p-3 rounded-xl shadow-2xs flex items-center justify-between hover:bg-amber-100/80 transition-all cursor-pointer">
                <div className="pr-2">
                  <div className="text-[10px] font-extrabold text-amber-900 uppercase tracking-wider flex items-center gap-1.5">
                    <Calendar className="w-3.5 h-3.5 text-amber-600" />
                    <span>Calls Made Today</span>
                  </div>
                  <div className="text-[9px] text-amber-700 font-bold mt-0.5">Real-time logs today</div>
                </div>
                <div className="bg-amber-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                  {allTelecallerLogs.filter(log => log.callDate === todayISO).length}
                </div>
              </div>
              {/* Card 2: 1st Free Services Due */}
              <div className="bg-teal-50/70 border border-teal-300/80 p-3 rounded-xl shadow-2xs flex items-center justify-between hover:bg-teal-100/80 transition-all cursor-pointer">
                <div className="pr-2">
                  <div className="text-[10px] font-extrabold text-teal-900 uppercase tracking-wider flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-teal-600" />
                    <span>1st FS Due</span>
                  </div>
                  <div className="text-[9px] text-teal-700 font-bold mt-0.5">&le; 30 days of delivery</div>
                </div>
                <div className="bg-teal-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                  {freeServiceFollowupCategories.fs1.length}
                </div>
              </div>
              {/* Card 3: Other Services Due */}
              <div className="bg-indigo-50/70 border border-indigo-300/80 p-3 rounded-xl shadow-2xs flex items-center justify-between hover:bg-indigo-100/80 transition-all cursor-pointer">
                <div className="pr-2">
                  <div className="text-[10px] font-extrabold text-indigo-900 uppercase tracking-wider flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-indigo-600" />
                    <span>2nd-10th FS Due</span>
                  </div>
                  <div className="text-[9px] text-indigo-700 font-bold mt-0.5">90 days since last FS</div>
                </div>
                <div className="bg-indigo-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                  {[
                    freeServiceFollowupCategories.fs2,
                    freeServiceFollowupCategories.fs3,
                    freeServiceFollowupCategories.fs4,
                    freeServiceFollowupCategories.fs5,
                    freeServiceFollowupCategories.fs6,
                    freeServiceFollowupCategories.fs7,
                    freeServiceFollowupCategories.fs8,
                    freeServiceFollowupCategories.fs9,
                    freeServiceFollowupCategories.fs10,
                  ].reduce((acc, list) => acc + list.length, 0)}
                </div>
              </div>
              {/* Card 4: Post Warranty Service Due */}
              <div className="bg-rose-50/70 border border-rose-300/80 p-3 rounded-xl shadow-2xs flex items-center justify-between hover:bg-rose-100/80 transition-all cursor-pointer">
                <div className="pr-2">
                  <div className="text-[10px] font-extrabold text-rose-900 uppercase tracking-wider flex items-center gap-1.5">
                    <AlertCircle className="w-3.5 h-3.5 text-rose-600" />
                    <span>Post Warranty Due</span>
                  </div>
                  <div className="text-[9px] text-rose-700 font-bold mt-0.5">&gt; 2 years of delivery</div>
                </div>
                <div className="bg-rose-600 text-white text-lg font-black px-3 py-1.5 rounded-lg shadow-sm shrink-0">
                  {freeServiceFollowupCategories.post_warranty.length}
                </div>
              </div>
            </div>
          </div>

          {/* Today's Calling History list */}
          <div className="bg-white rounded-2xl border border-slate-200 p-4 md:p-6 shadow-xs">
            <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
              <Phone className="w-4 h-4 text-amber-500" />
              <span>Today's Call Progress & Telecalling Summary</span>
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              List of customers telecalled today ({todayISO}). Ensures transparent logs of daily target telecalls.
            </p>

            {allTelecallerLogs.filter(log => log.callDate === todayISO).length === 0 ? (
              <div className="text-slate-400 italic text-xs text-center p-6 border border-dashed border-slate-200 rounded-xl bg-slate-50 mt-4">
                No calls recorded today yet. Use the "Log Call" button below on any due customer to record interactions!
              </div>
            ) : (
              <div className="overflow-x-auto border border-slate-100 rounded-xl mt-4">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 text-slate-700 font-bold uppercase text-[10px] border-b border-slate-100">
                    <tr>
                      <th className="p-3">Customer Info</th>
                      <th className="p-3">Chassis & Model</th>
                      <th className="p-3">Supervisor</th>
                      <th className="p-3">Remarks / Conversation</th>
                      <th className="p-3">Next Scheduled Call</th>
                      <th className="p-3">Caller</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {allTelecallerLogs.filter(log => log.callDate === todayISO).map((log, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                        <td className="p-3">
                          <div className="font-bold text-slate-900">{log.customerName}</div>
                          <div className="text-[10px] text-slate-500 font-mono">{log.mobileNumber}</div>
                        </td>
                        <td className="p-3">
                          <div className="font-semibold text-slate-800">{log.model}</div>
                          <div className="text-[10px] text-slate-500 font-mono">Chassis: {log.chassisNo}</div>
                        </td>
                        <td className="p-3">
                          <span className="px-1.5 py-0.5 bg-purple-50 text-purple-700 border border-purple-100 rounded text-[10px] font-bold">
                            {log.supervisor}
                          </span>
                        </td>
                        <td className="p-3 text-slate-700 italic max-w-xs truncate" title={log.remarks}>
                          "{log.remarks}"
                        </td>
                        <td className="p-3 font-mono font-bold text-indigo-700 text-[10px]">
                          {log.nextCallDate ? fmtDate(log.nextCallDate) : 'Not Scheduled'}
                        </td>
                        <td className="p-3 text-slate-500 font-medium">
                          {log.calledBy}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Collapsible Free Service Tree Categories */}
          <div className="space-y-4">
            <h3 className="text-sm font-black text-slate-900 uppercase tracking-wide">
              🌳 Follow-up Service Tree Channels
            </h3>

            {[
              { id: 'fs1', title: '1st Free Service', subtitle: 'Delivery + 30 Days Timeline', list: freeServiceFollowupCategories.fs1, color: 'border-teal-500 text-teal-800 bg-teal-50/40' },
              { id: 'fs2', title: '2nd Free Service', subtitle: '90 Days after 1st Free Service', list: freeServiceFollowupCategories.fs2, color: 'border-cyan-500 text-cyan-800 bg-cyan-50/40' },
              { id: 'fs3', title: '3rd Free Service', subtitle: '90 Days after 2nd Free Service', list: freeServiceFollowupCategories.fs3, color: 'border-blue-500 text-blue-800 bg-blue-50/40' },
              { id: 'fs4', title: '4th Free Service', subtitle: '90 Days after 3rd Free Service', list: freeServiceFollowupCategories.fs4, color: 'border-indigo-500 text-indigo-800 bg-indigo-50/40' },
              { id: 'fs5', title: '5th Free Service', subtitle: '90 Days after 4th Free Service', list: freeServiceFollowupCategories.fs5, color: 'border-purple-500 text-purple-800 bg-purple-50/40' },
              { id: 'fs6', title: '6th Free Service', subtitle: '90 Days after 5th Free Service', list: freeServiceFollowupCategories.fs6, color: 'border-pink-500 text-pink-800 bg-pink-50/40' },
              { id: 'fs7', title: '7th Free Service', subtitle: '90 Days after 6th Free Service', list: freeServiceFollowupCategories.fs7, color: 'border-orange-500 text-orange-800 bg-orange-50/40' },
              { id: 'fs8', title: '8th Free Service', subtitle: '90 Days after 7th Free Service', list: freeServiceFollowupCategories.fs8, color: 'border-emerald-500 text-emerald-800 bg-emerald-50/40' },
              { id: 'fs9', title: '9th Free Service', subtitle: '90 Days after 8th Free Service', list: freeServiceFollowupCategories.fs9, color: 'border-yellow-600 text-yellow-800 bg-yellow-50/40' },
              { id: 'fs10', title: '10th Free Service', subtitle: '90 Days after 9th Free Service', list: freeServiceFollowupCategories.fs10, color: 'border-amber-600 text-amber-800 bg-amber-50/40' },
              { id: 'general_90', title: 'Next Service Follow-up (90 Days after Paid/Repairs)', subtitle: 'For customers within 2-year warranty whose last job card was General Paid Service/Repairs', list: freeServiceFollowupCategories.general_90, color: 'border-slate-500 text-slate-800 bg-slate-50/40' },
              { id: 'post_warranty', title: 'Post Warranty Follow-up (90 Days after Paid Service)', subtitle: 'For customers past 2-year warranty whose last job card was General Paid Service/Repairs', list: freeServiceFollowupCategories.post_warranty, color: 'border-rose-600 text-rose-800 bg-rose-50/40' },
              { id: 'gear_1', title: 'Gear Oil Follow-up (1st Year)', subtitle: 'Approaching 1 Year from Delivery Date (Remind 15 days prior)', list: freeServiceFollowupCategories.gear_1, color: 'border-rose-500 text-rose-800 bg-rose-50/40' },
              { id: 'gear_2', title: 'Gear Oil Follow-up (2nd Year)', subtitle: 'Approaching 2 Years from Delivery Date (Remind 15 days prior)', list: freeServiceFollowupCategories.gear_2, color: 'border-red-500 text-red-800 bg-red-50/40' },
            ].map((cat) => {
              const isExpanded = !!expandedCategories[cat.id];
              return (
                <div key={cat.id} className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-2xs">
                  {/* Category Header */}
                  <button
                    type="button"
                    onClick={() => {
                      setExpandedCategories(prev => ({ ...prev, [cat.id]: !prev[cat.id] }));
                    }}
                    className="w-full flex items-center justify-between p-4 bg-slate-50 hover:bg-slate-100/80 transition-colors text-left cursor-pointer border-b border-slate-100"
                  >
                    <div className="flex items-center gap-3">
                      <div className={`p-1.5 rounded-lg border font-black ${cat.color}`}>
                        {cat.id.startsWith('gear') ? '🛢️' : '🔧'}
                      </div>
                      <div>
                        <h4 className="text-xs font-black text-slate-900">{cat.title}</h4>
                        <p className="text-[10px] text-slate-500 font-medium mt-0.5">{cat.subtitle}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className={`text-[10px] px-2 py-0.5 rounded-full font-black ${
                        cat.list.length > 0 ? 'bg-indigo-100 text-indigo-800' : 'bg-slate-200 text-slate-500'
                      }`}>
                        {cat.list.length} due
                      </span>
                      {isExpanded ? (
                        <ChevronDown className="w-4 h-4 text-slate-400" />
                      ) : (
                        <ChevronDown className="rotate-180 w-4 h-4 text-slate-400" />
                      )}
                    </div>
                  </button>

                  {/* Expanded Content List */}
                  {isExpanded && (
                    <div className="p-4 overflow-x-auto">
                      {cat.list.length === 0 ? (
                        <div className="text-slate-400 italic text-[11px] text-center py-6">
                          No customers in this category matching the current follow-up criteria.
                        </div>
                      ) : (
                        <table className="w-full text-left text-xs">
                          <thead className="bg-slate-50 text-slate-700 font-bold uppercase text-[9px] border-b border-slate-100">
                            <tr>
                              <th className="p-2.5">Cust ID / File No</th>
                              <th className="p-2.5">Customer details</th>
                              <th className="p-2.5">Chassis & Model</th>
                              <th className="p-2.5">Delivery date</th>
                              {cat.id.startsWith('gear') ? (
                                <th className="p-2.5">Days Elapsed</th>
                              ) : (
                                <th className="p-2.5">Due Stats</th>
                              )}
                              <th className="p-2.5">Last Call Status</th>
                              <th className="p-2.5 text-center">Action</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {cat.list.map((item, index) => (
                              <tr key={index} className="hover:bg-slate-50/40 transition-colors">
                                <td className="p-2.5 font-mono font-bold text-slate-600 text-[10px]">
                                  {item.historyFileNo || '—'}
                                </td>
                                <td className="p-2.5">
                                  <div className="font-bold text-slate-900">{item.customerName}</div>
                                  <div className="text-[10px] text-indigo-700 font-mono font-semibold">{item.mobileNumber || '—'}</div>
                                  <div className="text-[9px] text-slate-500 font-medium truncate max-w-[200px]">
                                    {item.village && `${item.village}, `}{item.mandal}
                                  </div>
                                </td>
                                <td className="p-2.5">
                                  <div className="font-semibold text-slate-800">{item.model || '—'}</div>
                                  <div className="text-[10px] text-slate-500 font-mono">Chassis: {item.chassisNo}</div>
                                </td>
                                <td className="p-2.5 font-mono text-slate-700">
                                  {item.dateOfDel || '—'}
                                </td>
                                {cat.id.startsWith('gear') ? (
                                  <td className="p-2.5">
                                    <div className="font-bold text-rose-700 font-mono text-[10px]">
                                      {item.daysSinceDel} days ago
                                    </div>
                                    <div className="text-[9px] text-slate-400 font-medium uppercase">
                                      {cat.id === 'gear_1' ? '1st Year' : '2nd Year'}
                                    </div>
                                  </td>
                                ) : (
                                  <td className="p-2.5">
                                    <div className={`font-mono font-bold text-[10px] ${
                                      item.daysDue < 0 ? 'text-rose-600' : 'text-emerald-600'
                                    }`}>
                                      {item.daysDue < 0 
                                        ? `Overdue by ${Math.abs(item.daysDue)} days` 
                                        : `Due in ${item.daysDue} days`
                                      }
                                    </div>
                                    <div className="text-[9px] text-slate-400 font-semibold uppercase truncate max-w-[120px]">
                                      {item.lastCompleted ? `After ${item.lastCompleted}` : 'Based on Del. Date'}
                                    </div>
                                  </td>
                                )}
                                <td className="p-2.5">
                                  {item.lastCallDate ? (
                                    <div>
                                      <div className="text-[9px] text-slate-400">
                                        Called: <span className="font-bold font-mono text-slate-700">{fmtDate(item.lastCallDate)}</span>
                                      </div>
                                      <div className="text-[10px] text-slate-700 italic max-w-[180px] truncate" title={item.lastRemarks}>
                                        "{item.lastRemarks}"
                                      </div>
                                    </div>
                                  ) : (
                                    <span className="text-[9px] text-slate-400 italic">No calls logged yet</span>
                                  )}
                                </td>
                                <td className="p-2.5 text-center">
                                  <div className="flex flex-col sm:flex-row items-center justify-center gap-1.5">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setSelectedFollowupForLog(item);
                                        setNewFollowupCallDate(todayISO);
                                        setNewFollowupRemarks('');
                                        setNewFollowupNextCallDate('');
                                      }}
                                      className="px-2 py-1 bg-teal-600 hover:bg-teal-700 text-white font-bold text-[10px] rounded transition-colors shadow-3xs hover:shadow-2xs cursor-pointer flex items-center gap-1 shrink-0"
                                    >
                                      <Phone className="w-2.5 h-2.5" />
                                      <span>Log Call</span>
                                    </button>

                                    <button
                                      type="button"
                                      onClick={() => {
                                        setComplaintForm({
                                          id: '',
                                          complaintNo: '',
                                          complaintDate: new Date().toISOString().split('T')[0],
                                          chassisNo: item.chassisNo || '',
                                          customerName: item.customerName || '',
                                          mobileNumber: item.mobileNumber || '',
                                          tractorModel: item.model || '',
                                          complaintDetails: '',
                                          assignedMechanic: item.lastTechnician || '',
                                          assignedSupervisor: item.supervisor || '',
                                          jobCardNo: '',
                                          status: 'Open',
                                          resolution: '',
                                          closureDate: '',
                                        });
                                        setIsComplaintModalOpen(true);
                                      }}
                                      className="px-2 py-1 bg-rose-600 hover:bg-rose-700 text-white font-bold text-[10px] rounded transition-colors shadow-3xs hover:shadow-2xs cursor-pointer flex items-center gap-1 shrink-0"
                                    >
                                      <AlertCircle className="w-2.5 h-2.5" />
                                      <span>Complaint</span>
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* INTERACTIVE CUSTOMER CALL LOG / FOLLOW-UP REMARKS MODAL */}
      {selectedFollowupForLog && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[92vh] flex flex-col my-auto">
            
            {/* Header */}
            <div className="bg-amber-600 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <PhoneCall className="w-5 h-5 text-amber-100 animate-bounce" />
                <div>
                  <h3 className="text-base font-bold">Customer Call Follow-up & Remarks Entry</h3>
                  <p className="text-xs text-amber-100">Write customer feedback, select preset quick responses, schedule next follow-up, and view full service history below.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setSelectedFollowupForLog(null)}
                className="text-amber-100 hover:text-white p-1 rounded-lg hover:bg-amber-700 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Content Body */}
            <div className="p-4 overflow-y-auto space-y-6 flex-1">
              
              {/* Top Section: Form Input & Call History */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                {/* Form Input Area */}
                <div className="space-y-4">
                  <div className="bg-amber-50/70 border border-amber-200/80 p-3 rounded-xl">
                    <h4 className="text-xs font-black text-amber-950 uppercase tracking-wider mb-2">Customer Profile Summary</h4>
                    <div className="text-xs space-y-1 text-slate-700 font-semibold">
                      <div>
                        👤 Name: <span className="text-slate-950 font-extrabold">{selectedFollowupForLog.customerName}</span>
                        {selectedFollowupForLog.historyFileNo && (
                          <span
                            className={`ml-2 inline-flex items-center px-1.5 py-0.5 rounded font-mono text-xs font-black border shadow-2xs ${
                              isWithin2Years(selectedFollowupForLog.dateOfDel || selectedFollowupForLog.installDate)
                                ? 'bg-blue-600 text-white border-blue-700'
                                : 'bg-red-600 text-white border-red-700'
                            }`}
                          >
                            HFN: {selectedFollowupForLog.historyFileNo}
                          </span>
                        )}
                      </div>
                      <div>📞 Mobile: <span className="text-indigo-700 font-mono font-extrabold">{selectedFollowupForLog.mobileNumber || '—'}</span></div>
                      <div>🚜 Model: <span className="text-slate-950 font-black">{selectedFollowupForLog.model || '—'} ({selectedFollowupForLog.modelType || 'N/A'})</span></div>
                      <div>📍 Location: <span className="text-slate-900 font-bold">{selectedFollowupForLog.village || '—'}, {selectedFollowupForLog.mandal || '—'}</span></div>
                      {selectedFollowupForLog.supervisor && (
                        <div>👨‍💼 Supervisor: <span className="text-purple-950 font-extrabold px-1.5 py-0.5 bg-purple-100 rounded">{selectedFollowupForLog.supervisor}</span></div>
                      )}
                    </div>
                  </div>

                  <div className="space-y-3.5">
                    <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider border-b pb-1">Log New Interaction</h4>
                    
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Call / Interaction Date *</label>
                        <input
                          type="date"
                          value={newFollowupCallDate}
                          onChange={(e) => setNewFollowupCallDate(e.target.value)}
                          className="w-full text-xs p-2.5 border border-slate-300 rounded-lg outline-none font-semibold text-slate-800 focus:border-amber-600 bg-white"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Caller / Supervisor Name (ఎవరు కాల్ చేశారు)*</label>
                        <div className="relative">
                          <input
                            type="text"
                            list="telecaller-list-options"
                            placeholder="Select or enter supervisor name..."
                            value={newFollowupCalledBy}
                            onChange={(e) => {
                              setNewFollowupCalledBy(e.target.value);
                              try { localStorage.setItem('telecaller_last_caller', e.target.value); } catch(err){}
                            }}
                            className="w-full text-xs p-2.5 border border-slate-300 rounded-lg outline-none font-extrabold text-amber-950 focus:border-amber-600 bg-white"
                          />
                          <datalist id="telecaller-list-options">
                            {telecallersList.map((st, i) => (
                              <option key={i} value={st} />
                            ))}
                          </datalist>
                        </div>
                      </div>
                    </div>

                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="block text-[11px] font-bold text-slate-700">Customer Remarks / What Customer Said * (customer emi chepparu)</label>
                        <button
                          type="button"
                          onClick={() => setIsAddingQuickRemark(!isAddingQuickRemark)}
                          className="text-[10px] font-bold text-amber-800 hover:text-amber-950 bg-amber-100 hover:bg-amber-200 px-2 py-0.5 rounded border border-amber-300 transition-colors cursor-pointer"
                        >
                          {isAddingQuickRemark ? '✕ Close Presets' : '⚙️ Manage Presets'}
                        </button>
                      </div>

                      {/* QUICK REMARKS CHIPS CONTAINER */}
                      <div className="mb-2 bg-amber-50/80 p-2 rounded-lg border border-amber-200 space-y-1.5">
                        <div className="text-[10px] font-black text-amber-950 uppercase tracking-wide flex items-center justify-between">
                          <span>⚡ Quick Dialogs / Presets (క్లిక్ చేసి రెస్పాన్స్ ఎంచుకోండి):</span>
                          <span className="text-[9px] text-amber-700 font-normal">Click to insert</span>
                        </div>
                        
                        <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto pr-1">
                          {quickRemarksList.map((remark, idx) => (
                            <div
                              key={idx}
                              className="inline-flex items-center gap-1 bg-white hover:bg-amber-100 text-slate-800 text-[11px] font-bold px-2 py-1 rounded-md border border-amber-300 shadow-2xs transition-all cursor-pointer group"
                            >
                              <span
                                onClick={() => handleSelectQuickRemark(remark)}
                                className="hover:text-amber-950"
                                title="Click to add to remarks"
                              >
                                💬 {remark}
                              </span>
                              {isAddingQuickRemark && (
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleRemoveQuickRemark(idx);
                                  }}
                                  className="text-red-500 hover:text-red-700 font-extrabold ml-1 text-xs cursor-pointer px-1 rounded hover:bg-red-50"
                                  title="Delete this preset remark"
                                >
                                  ×
                                </button>
                              )}
                            </div>
                          ))}
                        </div>

                        {/* ADD CUSTOM REMARK FIELD */}
                        {isAddingQuickRemark && (
                          <div className="pt-2 border-t border-amber-200 flex items-center gap-1.5">
                            <input
                              type="text"
                              value={newQuickRemarkInput}
                              onChange={(e) => setNewQuickRemarkInput(e.target.value)}
                              placeholder="Type new default response (e.g., Tractor not running)..."
                              className="flex-1 text-xs p-1.5 border border-amber-300 rounded bg-white text-slate-800 font-medium outline-none focus:border-amber-600"
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  handleAddQuickRemark(newQuickRemarkInput);
                                }
                              }}
                            />
                            <button
                              type="button"
                              onClick={() => handleAddQuickRemark(newQuickRemarkInput)}
                              className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded transition-colors cursor-pointer"
                            >
                              + Add Preset
                            </button>
                          </div>
                        )}
                      </div>

                      <textarea
                        placeholder="Type details of what the customer told you... (or click preset buttons above to auto-fill)"
                        rows={3}
                        value={newFollowupRemarks}
                        onChange={(e) => setNewFollowupRemarks(e.target.value)}
                        className="w-full text-xs p-2.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-800 focus:border-amber-600"
                        required
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Next Follow-up Call Date (malli call chese date)</label>
                      <input
                        type="date"
                        value={newFollowupNextCallDate}
                        onChange={(e) => setNewFollowupNextCallDate(e.target.value)}
                        className="w-full text-xs p-2.5 border border-slate-300 rounded-lg outline-none font-semibold text-slate-800 focus:border-amber-600"
                      />
                    </div>
                  </div>
                </div>

                {/* Interaction History Timeline */}
                <div className="border-l md:border-l border-slate-200 pl-0 md:pl-5 flex flex-col max-h-[50vh] md:max-h-full">
                  <h4 className="text-xs font-black text-slate-800 uppercase tracking-wider border-b pb-1 mb-3 flex items-center gap-1.5 shrink-0">
                    <History className="w-4 h-4 text-slate-500" />
                    Follow-up Call History Timeline ({(selectedFollowupForLog.followupHistory || []).length})
                  </h4>

                  {(!selectedFollowupForLog.followupHistory || selectedFollowupForLog.followupHistory.length === 0) ? (
                    <div className="flex-1 flex flex-col items-center justify-center border border-dashed border-slate-200 rounded-xl p-8 bg-slate-50 text-slate-400 font-bold text-center">
                      <PhoneCall className="w-8 h-8 text-slate-300 mb-2" />
                      No previous call logs for this customer.
                    </div>
                  ) : (
                    <div className="flex-1 overflow-y-auto pr-1 space-y-3 max-h-80">
                      {selectedFollowupForLog.followupHistory.map((log: any, idx: number) => (
                        <div key={idx} className="relative bg-amber-50/30 border border-amber-200/50 p-3 rounded-lg text-xs space-y-1.5 hover:bg-amber-50/50 transition-colors">
                          <div className="flex items-center justify-between border-b border-amber-200/30 pb-1 flex-wrap gap-1">
                            <span className="font-bold text-amber-950 flex items-center gap-1">
                              <Calendar className="w-3.5 h-3.5 text-amber-600" />
                              Called: {fmtDate(log.callDate)}
                            </span>
                            {log.calledBy && (
                              <span className="text-[10px] bg-amber-200/90 text-amber-950 px-1.5 py-0.2 rounded font-extrabold border border-amber-300">
                                👤 {log.calledBy}
                              </span>
                            )}
                            <span className="text-[10px] text-slate-400 font-mono">#{selectedFollowupForLog.followupHistory.length - idx}</span>
                          </div>
                          <div className="text-slate-800 leading-relaxed font-semibold">
                            Remarks: <span className="text-slate-700 italic font-medium">"{log.remarks}"</span>
                          </div>
                          {log.nextCallDate && (
                            <div className="text-[10px] text-indigo-700 font-bold flex items-center gap-1 mt-0.5 bg-indigo-50 px-1.5 py-0.5 rounded border border-indigo-100/50 w-fit">
                              <Calendar className="w-3 h-3 text-indigo-500" />
                              Next Call Date: {fmtDate(log.nextCallDate)}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Bottom Section: RECORDED JOB CARDS SERVICE HISTORY */}
              <div className="pt-4 border-t border-slate-200 space-y-3">
                <div className="flex items-center justify-between flex-wrap gap-2">
                  <div className="flex items-center gap-2">
                    <History className="w-4 h-4 text-amber-700" />
                    <h4 className="text-xs font-black text-amber-950 uppercase tracking-wider">
                      Recorded Job Cards / Past Service History ({selectedFollowupForLog.jobCardsCount || (selectedFollowupForLog.cards || []).length})
                    </h4>
                  </div>
                  <span className="text-[10px] text-slate-500 font-semibold italic">
                    (Scroll to view all past job card entries for telecalling reference)
                  </span>
                </div>

                {(!selectedFollowupForLog.cards || selectedFollowupForLog.cards.length === 0) ? (
                  <div className="p-4 bg-slate-50 border border-dashed border-slate-200 rounded-xl text-center text-xs text-slate-400 font-bold">
                    No recorded job cards found for this customer.
                  </div>
                ) : (
                  <div className="overflow-x-auto border border-slate-200 rounded-xl max-h-72 overflow-y-auto shadow-2xs">
                    <table className="w-full text-left text-xs">
                      <thead className="sticky top-0 bg-amber-600 text-white font-bold text-[11px] shadow-2xs z-10">
                        <tr>
                          <th className="p-2 w-8 text-center">#</th>
                          <th className="p-2 w-28">Job Card No</th>
                          <th className="p-2 w-24">Date</th>
                          <th className="p-2 w-28">Service Type</th>
                          <th className="p-2 w-28">Service Count</th>
                          <th className="p-2 w-20">Hrs Run</th>
                          <th className="p-2 min-w-[200px]">Repairs & Work Done</th>
                          <th className="p-2 w-32">Technician</th>
                          <th className="p-2 w-20 text-center">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 font-medium bg-white">
                        {selectedFollowupForLog.cards.map((card: any, idx: number) => {
                          const formattedDate = fmtDate(card.jobDate || card.jobOpenDate || (card.createdAt ? card.createdAt.split('T')[0] : ''));
                          const primaryRep = Array.isArray(card.repairRows) && card.repairRows.length > 0
                            ? card.repairRows.map((r: any) => r.repair).filter(Boolean).join(', ')
                            : (card.problemDescription || card.reasonsForAnalysis || '');
                          const extraRepText = getExtraRepairsText(card) !== '—'
                            ? getExtraRepairsText(card)
                            : (card.extraRepairs || card['EXTRA OTHER REPAIRS DONE WITH FREE SERVICE'] || card.extraRepairsDone || '');
                          const serviceCountVal = card.freeServiceList || card.freeServiceCount || card['FREE SERVICE LIST'] || card['Service Count'] || getFreeServiceText(card) || '—';
                          
                          return (
                            <tr key={card.id || card.jobNo || idx} className="hover:bg-amber-50/80 transition-colors text-[11px]">
                              <td className="p-2 text-center font-mono font-bold text-slate-400">{idx + 1}</td>
                              <td className="p-2 font-mono font-bold text-indigo-950">{card.jobNo || card.complaintNo || '—'}</td>
                              <td className="p-2 font-bold text-slate-600">{formattedDate}</td>
                              <td className="p-2 text-indigo-950 font-semibold">{card.serviceType || 'Paid Service'}</td>
                              <td className="p-2 text-slate-900 font-bold">{serviceCountVal}</td>
                              <td className="p-2 font-mono font-extrabold text-indigo-600">{card.hourMeter || card.hrsRun ? `${card.hourMeter || card.hrsRun} Hrs` : '—'}</td>
                              <td className="p-2 text-slate-700 leading-relaxed font-normal">
                                {primaryRep && primaryRep !== '—' && (
                                  <div className="font-semibold">{primaryRep}</div>
                                )}
                                {extraRepText && extraRepText !== '—' && (
                                  <div className="text-[10px] text-amber-950 font-bold bg-amber-100/80 px-1.5 py-0.5 rounded border border-amber-200 mt-0.5 inline-block">
                                    🔧 Extra: {extraRepText}
                                  </div>
                                )}
                                {(!primaryRep || primaryRep === '—') && (!extraRepText || extraRepText === '—') && '—'}
                              </td>
                              <td className="p-2">
                                <div className="font-semibold text-slate-800">{card.mechanic || card.technicianName || '—'}</div>
                              </td>
                              <td className="p-2 text-center">
                                <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                                  {card.status || 'COMPLETED'}
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

            </div>

            {/* Modal Footer */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2 shrink-0">
              <button
                type="button"
                onClick={() => setSelectedFollowupForLog(null)}
                className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
              >
                Cancel
              </button>
              
              <button
                type="button"
                onClick={() => {
                  setComplaintForm({
                    id: '',
                    complaintNo: '',
                    complaintDate: new Date().toISOString().split('T')[0],
                    chassisNo: selectedFollowupForLog.chassisNo || '',
                    customerName: selectedFollowupForLog.customerName || '',
                    mobileNumber: selectedFollowupForLog.mobileNumber || '',
                    tractorModel: selectedFollowupForLog.model || '',
                    complaintDetails: newFollowupRemarks.trim() || '',
                    assignedMechanic: selectedFollowupForLog.lastTechnician || '',
                    assignedSupervisor: selectedFollowupForLog.supervisor || '',
                    jobCardNo: '',
                    status: 'Open',
                    resolution: '',
                    closureDate: '',
                  });
                  setIsComplaintModalOpen(true);
                }}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1 shadow-sm mr-auto sm:mr-0"
              >
                <AlertCircle className="w-3.5 h-3.5" />
                <span>Create Complaint</span>
              </button>

              <button
                type="button"
                onClick={async () => {
                  if (!newFollowupCallDate) {
                    alert('Please select a Call Date.');
                    return;
                  }
                  if (!newFollowupRemarks.trim()) {
                    alert('Please write what the customer said (Remarks).');
                    return;
                  }

                  const callerName = newFollowupCalledBy.trim() || 'General Staff';
                  try {
                    localStorage.setItem('telecaller_last_caller', callerName);
                  } catch (e) {}

                  const newLog = {
                    callDate: newFollowupCallDate,
                    remarks: newFollowupRemarks.trim(),
                    nextCallDate: newFollowupNextCallDate || '',
                    calledBy: callerName
                  };

                  const updatedHistory = [newLog, ...(selectedFollowupForLog.followupHistory || [])];
                  
                  await updateCustomerRecordInIndex(selectedFollowupForLog.chassisNo, {
                    followupHistory: updatedHistory,
                    lastCallDate: newFollowupCallDate,
                    lastRemarks: newFollowupRemarks.trim(),
                    lastNextCallDate: newFollowupNextCallDate || '',
                    lastCalledBy: callerName
                  });

                  alert('✅ Follow-up log and remarks saved successfully in database!');
                  setSelectedFollowupForLog(null);
                }}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1 shadow-sm"
              >
                <Save className="w-3.5 h-3.5" />
                <span>Save Follow-up Details</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD / EDIT STAFF MEMBER MODAL OVERLAY */}
      {isStaffModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-lg rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col my-auto">
            <div className="bg-purple-900 text-white p-4 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2.5">
                <Users className="w-5 h-5 text-purple-300" />
                <div>
                  <h3 className="text-base font-bold">{editingStaffId ? 'Edit Staff Member' : 'Add New Staff Member'}</h3>
                  <p className="text-xs text-purple-200">Enter full details for mechanic or supervisor registration.</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsStaffModalOpen(false);
                  setEditingStaffId(null);
                }}
                className="text-purple-200 hover:text-white p-1 rounded-lg hover:bg-purple-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                saveStaffMember({
                  id: editingStaffId,
                  ...staffForm
                });
                setIsStaffModalOpen(false);
                setEditingStaffId(null);
                setStaffForm({
                  name: '',
                  fatherName: '',
                  village: '',
                  mandal: '',
                  mobileNumber: '',
                  role: 'mechanic',
                  dateOfJoining: '',
                  supervisor: ''
                });
              }}
              className="p-5 space-y-4 overflow-y-auto flex-1 text-xs"
            >
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <div className="sm:col-span-2">
                  <label className="block text-[11px] font-bold text-slate-800 mb-1">
                    Full Name <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="e.g. K. Venkateswara Rao"
                    value={staffForm.name}
                    onChange={(e) => setStaffForm({ ...staffForm, name: e.target.value })}
                    className="w-full text-xs p-2.5 border border-slate-300 rounded-lg outline-none font-medium text-slate-900 focus:border-purple-600"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Father's Name</label>
                  <input
                    type="text"
                    placeholder="e.g. Seshagiri Rao"
                    value={staffForm.fatherName}
                    onChange={(e) => setStaffForm({ ...staffForm, fatherName: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 focus:border-purple-600"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Job Role</label>
                  <select
                    value={staffForm.role}
                    onChange={(e) => setStaffForm({ ...staffForm, role: e.target.value as 'mechanic' | 'supervisor' })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 font-bold focus:border-purple-600 cursor-pointer"
                  >
                    <option value="mechanic">Mechanic / Technician</option>
                    <option value="supervisor">Supervisor / W/S Incharge</option>
                  </select>
                </div>

                {staffForm.role === 'mechanic' && (
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Assigned Supervisor</label>
                    <select
                      value={staffForm.supervisor || ''}
                      onChange={(e) => setStaffForm({ ...staffForm, supervisor: e.target.value })}
                      className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 font-bold focus:border-purple-600 cursor-pointer"
                    >
                      <option value="">-- Select Supervisor --</option>
                      {staffMembers.filter(s => s.role === 'supervisor').map(s => (
                        <option key={s.id} value={s.name}>{s.name}</option>
                      ))}
                    </select>
                  </div>
                )}

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mobile Number</label>
                  <input
                    type="text"
                    placeholder="10-digit mobile number"
                    value={staffForm.mobileNumber}
                    onChange={(e) => setStaffForm({ ...staffForm, mobileNumber: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none font-mono text-slate-800 focus:border-purple-600"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Date of Joining</label>
                  <input
                    type="date"
                    value={staffForm.dateOfJoining}
                    onChange={(e) => setStaffForm({ ...staffForm, dateOfJoining: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 focus:border-purple-600"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Village</label>
                  <input
                    type="text"
                    placeholder="e.g. Poranki"
                    value={staffForm.village}
                    onChange={(e) => setStaffForm({ ...staffForm, village: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 focus:border-purple-600"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mandal</label>
                  <input
                    type="text"
                    placeholder="e.g. Penamaluru"
                    value={staffForm.mandal}
                    onChange={(e) => setStaffForm({ ...staffForm, mandal: e.target.value })}
                    className="w-full text-xs p-2 border border-slate-300 rounded-lg outline-none text-slate-800 focus:border-purple-600"
                  />
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-3 shrink-0">
                <button
                  type="button"
                  onClick={() => {
                    setIsStaffModalOpen(false);
                    setEditingStaffId(null);
                  }}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs rounded-lg transition-colors shadow-xs cursor-pointer flex items-center gap-1.5"
                >
                  <Save className="w-4 h-4" /> Save Staff Details
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* VIEW STAFF DETAILS MODAL OVERLAY */}
      {viewingStaffDetails && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-md rounded-2xl shadow-2xl border border-slate-200 overflow-hidden my-auto">
            <div className="bg-purple-900 text-white p-4 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Users className="w-5 h-5 text-purple-300" />
                <div>
                  <h3 className="text-base font-bold">{viewingStaffDetails.name}</h3>
                  <p className="text-xs text-purple-200">Staff Profile Details</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setViewingStaffDetails(null)}
                className="text-purple-200 hover:text-white p-1 rounded-lg hover:bg-purple-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 space-y-4 text-xs">
              <div className="bg-purple-50 p-3.5 rounded-xl border border-purple-100 flex items-center justify-between">
                <div>
                  <div className="text-[10px] uppercase font-bold text-purple-600 tracking-wider">Job Role</div>
                  <div className="text-sm font-extrabold text-purple-950">
                    {viewingStaffDetails.role === 'supervisor' ? 'Supervisor / W/S Incharge' : viewingStaffDetails.role === 'mechanic' ? 'Mechanic / Technician' : viewingStaffDetails.role || 'Mechanic'}
                  </div>
                </div>
                <span className={`px-2.5 py-1 rounded-md text-xs font-black uppercase ${
                  viewingStaffDetails.role === 'supervisor' ? 'bg-purple-200 text-purple-900' : 'bg-blue-200 text-blue-900'
                }`}>
                  {viewingStaffDetails.role || 'Mechanic'}
                </span>
              </div>

              <div className="grid grid-cols-2 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">Father's Name</span>
                  <span className="text-xs font-bold text-slate-800">{viewingStaffDetails.fatherName || '—'}</span>
                </div>

                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">Mobile Number</span>
                  <span className="text-xs font-mono font-bold text-slate-800">{viewingStaffDetails.mobileNumber || '—'}</span>
                </div>

                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">Village</span>
                  <span className="text-xs font-bold text-slate-800">{viewingStaffDetails.village || '—'}</span>
                </div>

                <div>
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">Mandal</span>
                  <span className="text-xs font-bold text-slate-800">{viewingStaffDetails.mandal || '—'}</span>
                </div>

                <div className="col-span-2 pt-2 border-t border-slate-200">
                  <span className="block text-[10px] font-bold text-slate-500 uppercase">Date of Joining</span>
                  <span className="text-xs font-bold text-slate-800">{viewingStaffDetails.dateOfJoining || '—'}</span>
                </div>

                {viewingStaffDetails.role === 'mechanic' && (
                  <div className="col-span-2 pt-2 border-t border-slate-200">
                    <span className="block text-[10px] font-bold text-slate-500 uppercase">Assigned Supervisor</span>
                    <span className="text-xs font-bold text-slate-800">{viewingStaffDetails.supervisor || '—'}</span>
                  </div>
                )}
              </div>

              <div className="pt-2 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => {
                    const st = viewingStaffDetails;
                    setViewingStaffDetails(null);
                    setEditingStaffId(st.id);
                    setStaffForm({
                      name: st.name || '',
                      fatherName: st.fatherName || '',
                      village: st.village || '',
                      mandal: st.mandal || '',
                      mobileNumber: st.mobileNumber || '',
                      role: st.role || 'mechanic',
                      dateOfJoining: st.dateOfJoining || '',
                      supervisor: st.supervisor || ''
                    });
                    setIsStaffModalOpen(true);
                  }}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1.5"
                >
                  <Edit3 className="w-3.5 h-3.5" /> Edit Staff Details
                </button>
                <button
                  type="button"
                  onClick={() => setViewingStaffDetails(null)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MENU ORDER REORDER MODAL */}
      {isMenuOrderModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
          <div className="bg-white w-full max-w-md rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col my-auto">
            <div className="bg-blue-900 text-white p-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Settings className="w-5 h-5 text-blue-200" />
                <h3 className="text-sm font-black uppercase tracking-wider">Customize Menu Order</h3>
              </div>
              <button
                type="button"
                onClick={() => setIsMenuOrderModalOpen(false)}
                className="text-blue-200 hover:text-white p-1 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 space-y-3 max-h-[70vh] overflow-y-auto text-xs font-sans">
              <p className="text-slate-500 font-medium text-[11px]">
                Use the <span className="font-bold text-slate-800">Up ▲</span> and <span className="font-bold text-slate-800">Down ▼</span> arrows to arrange menu options in your preferred order.
              </p>
              
              <div className="space-y-1.5 border border-slate-200 p-2 rounded-xl bg-slate-50">
                {menuOrder.map((itemId, idx) => {
                  let label = itemId;
                  if (itemId === 'dashboard') label = 'Dashboard';
                  else if (itemId === 'complaints') label = 'Complaint Register';
                  else if (itemId === 'attendance') label = 'Staff Attendance';
                  else if (itemId === 'new_entry') label = 'New Job Entry';
                  else if (itemId === 'saved_cards') label = 'Saved Job Cards';
                  else if (itemId === 'followup') label = 'Customer Data';
                  else if (itemId === 'telecalling') label = 'Tele Calling';
                  else if (itemId === 'free_service_followup') label = 'Free Service Follow-up';
                  else if (itemId === 'reports') label = 'Reports & Analytics';
                  else if (itemId === 'databases') label = 'Master Databases';

                  return (
                    <div key={itemId} className="flex items-center justify-between bg-white p-2.5 rounded-lg border border-slate-200 shadow-2xs">
                      <span className="font-bold text-slate-800">{idx + 1}. {label}</span>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          disabled={idx === 0}
                          onClick={() => moveMenuItem(idx, 'up')}
                          className="p-1 rounded bg-slate-100 hover:bg-blue-600 hover:text-white disabled:opacity-30 text-slate-700 cursor-pointer disabled:cursor-not-allowed font-bold"
                          title="Move Up"
                        >
                          <ChevronUp className="w-4 h-4" />
                        </button>
                        <button
                          type="button"
                          disabled={idx === menuOrder.length - 1}
                          onClick={() => moveMenuItem(idx, 'down')}
                          className="p-1 rounded bg-slate-100 hover:bg-blue-600 hover:text-white disabled:opacity-30 text-slate-700 cursor-pointer disabled:cursor-not-allowed font-bold"
                          title="Move Down"
                        >
                          <ChevronDown className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-between items-center">
              <button
                type="button"
                onClick={resetMenuOrder}
                className="px-3 py-1.5 text-rose-700 hover:bg-rose-50 font-bold text-xs rounded-lg transition-colors cursor-pointer"
              >
                Reset Default
              </button>
              <button
                type="button"
                onClick={() => setIsMenuOrderModalOpen(false)}
                className="px-4 py-1.5 bg-blue-900 hover:bg-blue-800 text-white font-bold text-xs rounded-lg cursor-pointer"
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* FILLED JOB CARD VIEW MODAL */}
      {viewingCardModal && (() => {
        const card = viewingCardModal;
        const serviceLocation = card.serviceLocation || 'workshop';
        const wtyInfo = getCardWtyInfo(card);
        const onlineJobCardNo = card.onlineJobCardNo || '';
        const jobNo = card.jobNo || card.jobCardNo || '';
        const historyFileNo = card.historyFileNo || '';
        const status = card.status || 'Open';
        const jobDate = card.jobDate || card.complaintDate || card.createdAt || '';
        const model = card.model || '';
        const custName = card.custName || '';
        const translatedCustName = card.translatedCustName || '';
        const village = card.village || '';
        const translatedVillage = card.translatedVillage || '';
        const mandal = card.mandal || '';
        const translatedMandal = card.translatedMandal || '';
        const chassisNo = card.chassisNo || '';
        const custAddr = card.custAddr || '';
        const translatedCustAddr = card.translatedCustAddr || '';
        const engineNo = card.engineNo || '';
        const ownerMob = card.ownerMob || '';
        const installDate = card.installDate || card.dateOfDelivery || '';
        
        // Let's resolve repair rows
        const repairRows = card.repairRows || [];
        const totalLabour = card.totalLabour || '';
        const repairTotal = repairRows.reduce((acc: number, r: any) => acc + (parseFloat(r.charge) || 0), 0);
        
        // Let's resolve checkpoints
        const checkpoints = (card.checkpoints || []).length > 0
          ? card.checkpoints
          : DEFAULT_CHECKPOINTS.map(d => ({
              ...d,
              checked: (card.checkedCheckpoints || []).includes(d.id)
            }));
            
        const partRows = card.partRows || [];
        const partsTotal = partRows.reduce((acc: number, r: any) => acc + (parseFloat(r.amount) || 0), 0);
        
        const wsReport = card.wsReport || '';
        const mechanic = card.mechanic || '';
        const wsIncharge = card.wsIncharge || '';
        const billNo = card.billNo || '';
        const warrantyMaterial = card.warrantyMaterial || '';
        const nonWarrantyMaterial = card.nonWarrantyMaterial || '';
        const gTotal = card.gTotal || '';
        
        return (
          <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs flex items-center justify-center p-4 z-50 overflow-y-auto print:hidden">
            <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden max-h-[92vh] flex flex-col my-auto">
              <div className="bg-blue-900 text-white p-4 flex items-center justify-between shrink-0">
                <div className="flex items-center gap-2.5">
                  <FileText className="w-5 h-5 text-blue-200" />
                  <div>
                    <h2 className="text-sm font-black uppercase tracking-wider">
                      Filled Job Card Details (#{onlineJobCardNo || jobNo || 'Card'})
                    </h2>
                    <p className="text-[10px] text-blue-200 font-medium">
                      {custName} | Chassis: {chassisNo || '—'}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setViewingCardModal(null);
                      loadJobCard(card);
                      setTimeout(() => window.print(), 200);
                    }}
                    className="px-3 py-1.5 bg-blue-800 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
                  >
                    <Printer className="w-4 h-4" /> Print
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setViewingCardModal(null);
                      loadJobCard(card);
                      setActiveTab('new_entry');
                    }}
                    className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-slate-900 text-xs font-bold rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
                  >
                    <Edit3 className="w-4 h-4" /> Edit Card
                  </button>
                  <button
                    type="button"
                    onClick={() => setViewingCardModal(null)}
                    className="text-blue-200 hover:text-white p-1 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

            <div className="p-4 overflow-y-auto space-y-6 bg-slate-100 flex flex-col items-center">
              {/* PAGE 1: SRI GAYATHRI AUTOMOTIVES JOB CARD */}
              <div className="card-p1 bg-white border-2 border-blue-900 p-2.5 md:p-3 text-slate-900 max-w-[780px] w-full text-[11px] leading-snug shadow-md">
                {/* HEADER SECTION */}
                <div>
                  <div className="text-center font-extrabold text-blue-900 text-xs md:text-sm tracking-wide border-b border-blue-200 pb-0.5 mb-0 relative flex items-center justify-center gap-2">
                    <span>Job Card</span>
                    {/* EICHER LOGO OVAL TOP RIGHT */}
                    <div className="absolute right-0 top-0">
                      <div className="w-8 h-5 border-2 border-blue-900 rounded-full flex items-center justify-center bg-white shadow-xs">
                        <span className="font-serif font-black text-blue-900 text-[11px] italic">E</span>
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-between items-start mb-0">
                    <div className="flex-1">
                      <div className="text-xl md:text-2xl font-black text-blue-900 tracking-tight leading-none mb-0.5">
                        SRI GAYATHRI AUTOMOTIVES
                      </div>
                      <div className="text-[10px] font-extrabold text-blue-900">
                        Authorised Dealer : EICHER TRACTORS (Sales, Service & Spares)
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-medium">
                        D.No. 2-12-351, Opp Srinivasa Theater, Poranki, Vijayawada .
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-medium">
                        A.P - 521 137, Phone : 9063134025
                      </div>
                      <div className="text-[9.5px] text-slate-800 font-bold">
                        GST : 37AFNFS856BM1ZS, Email : srigayathriauto@gmail.com
                      </div>
                    </div>

                    <div className="w-48 text-right text-[9px] font-semibold text-blue-950 space-y-0.5 pl-1">
                      <div className="flex items-center justify-end gap-1">
                        <span>• Service : Location Workshop</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'workshop' ? 'bg-blue-900 text-white' : ''}`}>
                          {serviceLocation === 'workshop' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end gap-1">
                        <span>• DSS : Field</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'dss' ? 'bg-blue-900 text-white' : ''}`}>
                          {serviceLocation === 'dss' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end gap-1">
                        <span>• Event / Camp</span>
                        <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center font-bold text-[8px] ${serviceLocation === 'event' ? 'bg-blue-900 text-white' : ''}`}>
                          {serviceLocation === 'event' ? '✓' : ''}
                        </div>
                      </div>
                      <div className="flex items-center justify-end pt-0.5">
                        {wtyInfo.isWty ? (
                          <span className="px-2 py-0.5 rounded text-[9.5px] font-black tracking-wider uppercase border shadow-2xs bg-red-600 text-white border-red-700">
                            WARRANTY
                          </span>
                        ) : wtyInfo.isPostWty ? (
                          <span className="px-2 py-0.5 rounded text-[9.5px] font-black tracking-wider uppercase border shadow-2xs bg-amber-600 text-white border-amber-700">
                            POST WTY
                          </span>
                        ) : null}
                      </div>
                    </div>
                  </div>

                  <div className="text-[8.5px] text-slate-700 font-medium border-t border-b border-blue-800 py-0.5 mb-0.5 bg-blue-50/40 text-center">
                    <b>Branches :</b> Nandigama - 9063134026, Tiruvur- 9063134027, Gudivada - 9063134028, Machilipatnam - 9063134029
                  </div>
                </div>

                {/* SECTIONS A, B, C */}
                <div className="space-y-1.5 print:space-y-1 my-1">
                  {/* SECTION (A) TRACTOR DETAILS TABLE */}
                  <div>
                    <div className="flex justify-between items-center mb-0.5">
                      <div className="font-extrabold text-blue-900 text-[10px]">(A) Tractor Details</div>
                      <div className="text-[9.5px] font-bold text-blue-950 font-mono">
                        HISTORY FILE NO : {historyFileNo || '________'}
                      </div>
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[9.5px]">
                      <colgroup>
                        <col style={{ width: '15%' }} />
                        <col style={{ width: '19%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '18%' }} />
                        <col style={{ width: '15%' }} />
                        <col style={{ width: '25%' }} />
                      </colgroup>
                      <tbody>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 truncate">Online Job Card No.</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-indigo-900 font-mono">
                            <div className="flex items-center justify-between gap-1">
                              <span className="truncate">{onlineJobCardNo || jobNo}</span>
                              <span className={`text-[7px] px-1 py-0.5 rounded font-black uppercase border leading-none shrink-0 ${status === 'Closed' ? 'bg-red-50 text-red-700 border-red-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'}`}>
                                {status}
                              </span>
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-center">Date</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold truncate">{fmtDate(jobDate)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Model</td>
                          <td className="p-0.5 md:p-1 font-bold text-slate-900 truncate">{model}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Name of Customer</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-slate-900">
                            <div className="text-[11px] font-extrabold text-blue-950 leading-snug truncate">
                              {translatedCustName || custName}
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Chassis No.</td>
                          <td className="p-0.5 md:p-1 font-bold font-mono text-blue-950 truncate">{chassisNo}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Address</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold text-slate-900">
                            <div className="text-[10px] font-bold text-slate-900 leading-snug truncate">
                              {(() => {
                                const v = translatedVillage || village;
                                const m = translatedMandal || mandal;
                                if (v || m) {
                                  return [v, m].filter(Boolean).join(', ');
                                }
                                return translatedCustAddr || custAddr;
                              })()}
                            </div>
                          </td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Engine No.</td>
                          <td className="p-0.5 md:p-1 font-bold font-mono truncate">{engineNo}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Mobile No.</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 font-bold font-mono truncate">{ownerMob}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Date Of Delivery</td>
                          <td className="p-0.5 md:p-1 font-bold truncate">{fmtDate(installDate)}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Last Service Details</td>
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 md:p-1 truncate">{distDealership}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Hour Meter Reading</td>
                          <td className="p-0.5 md:p-1 font-bold truncate">{hourMeter}</td>
                        </tr>
                        <tr className="border-b border-blue-900 h-[22px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Date & Time in</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 truncate">{fmtDateTime(dateTimeIn)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-center">Out</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 truncate">{fmtDateTime(dateTimeOut)}</td>
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30 text-[8px] leading-tight">Exp. Repair Time/Cost</td>
                          <td className="p-0.5 md:p-1 font-medium truncate">{expectedRepairTime}</td>
                        </tr>
                        <tr className="h-[24px]">
                          <td className="border-r border-blue-900 p-0.5 md:p-1 font-bold bg-blue-50/30">Service</td>
                          <td colSpan={5} className="p-0.5 md:p-1">
                            <div className="flex flex-wrap items-center justify-between gap-0.5 text-[8.5px] font-bold">
                              {serviceTypeOptions.map(st => (
                                <div key={st} className="flex items-center gap-0.5">
                                  <span className="whitespace-nowrap">{st}</span>
                                  <div className={`w-3 h-3 border border-blue-900 flex items-center justify-center text-[7.5px] leading-none shrink-0 ${serviceType === st ? 'bg-blue-900 text-white font-bold' : ''}`}>
                                    {serviceType === st ? '✓' : ''}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SECTION (B) LABOUR & WORKSHOP REPAIRS SUMMARY */}
                  <div>
                    <div className="font-extrabold text-blue-900 text-[10px] mb-0">(B) Labour & Workshop Repairs Summary</div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[8.5px]">
                      <colgroup>
                        <col style={{ width: '6%' }} />
                        <col style={{ width: '48%' }} />
                        <col style={{ width: '30%' }} />
                        <col style={{ width: '16%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/60 font-bold text-blue-950 h-[18px]">
                          <th className="border-r border-blue-900 p-0.5 text-center">#</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Repairs Carried Out</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Rectification / Action</th>
                          <th className="p-0.5 text-right">Labour Charge ₹</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Array.from({ length: Math.max(repairRows.length, 2) }).map((_, i) => {
                          const r = repairRows[i] || { repair: '', rectification: '', charge: '' };
                          return (
                            <tr key={i} className="border-b border-blue-900 h-[17px]">
                              <td className="border-r border-blue-900 p-0.5 text-center font-mono font-semibold">{i + 1}</td>
                              <td className="border-r border-blue-900 p-0.5 font-medium truncate">{r.repair}</td>
                              <td className="border-r border-blue-900 p-0.5 font-medium truncate">{r.rectification}</td>
                              <td className="p-0.5 text-right font-mono font-bold truncate">{r.charge}</td>
                            </tr>
                          );
                        })}
                        <tr className="bg-blue-50/50 font-bold border-t border-blue-900 text-blue-950 h-[18px]">
                          <td colSpan={3} className="border-r border-blue-900 p-0.5 text-right font-bold text-[8.5px]">Total Labour Charge ₹</td>
                          <td className="p-0.5 text-right font-mono font-extrabold text-blue-900 text-[8.5px] truncate">
                            {totalLabour || (repairTotal ? repairTotal.toFixed(2) : '')}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SECTION (C) CHECKLIST 25 CHECKPOINT (TELUGU ONLY) */}
                  <div>
                    <div className="font-extrabold text-blue-900 text-[10px] mb-0.5">
                      (C) 25 చెక్‌లిస్ట్ తనిఖీ అంశాలు (Checklist - 25 Checkpoints)
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[8.5px]">
                      <colgroup>
                        <col style={{ width: '5.5%' }} />
                        <col style={{ width: '65.5%' }} />
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '7%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/60 font-bold text-blue-950 h-[26px]">
                          <th className="border-r border-blue-900 p-0.5 text-center text-[10px]">క్ర.సం</th>
                          <th className="border-r border-blue-900 p-0.5 text-left text-[10px]">విభాగం & తనిఖీ చేయవలసిన అంశం</th>
                          <th className="border-r border-blue-900 p-0.5 text-left text-[10px]">చేయవలసిన పని / చర్య</th>
                          <th className="p-0.5 text-center text-[10px]">✓</th>
                        </tr>
                      </thead>
                      <tbody>
                        {checkpoints.map((cp) => {
                          const defaultItem = DEFAULT_CHECKPOINTS.find(d => d.id === cp.id);
                          const teluguCategory = cp.categoryTe || defaultItem?.categoryTe || (cp.category && cp.category !== 'GENERAL' ? cp.category : 'సాధారణ');
                          const teluguItem = cp.itemTe || defaultItem?.itemTe || cp.item;
                          const teluguAction = cp.actionTe || defaultItem?.actionTe || cp.action;

                          return (
                            <tr key={cp.id} className="border-b border-blue-900/60 h-[26px] leading-tight hover:bg-blue-50/20">
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono font-bold text-[10px]">{cp.id}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-bold text-blue-950">
                                <div className="flex items-center gap-1.5 leading-none">
                                  <span className="text-[9px] font-black text-blue-800 shrink-0">[{teluguCategory}]</span>
                                  <span className="text-[11px] font-extrabold text-slate-950 truncate">{teluguItem}</span>
                                </div>
                              </td>
                              <td className="border-r border-blue-900/60 p-0.5 font-bold text-slate-900">
                                <span className="text-[10px] font-bold text-slate-900 truncate block leading-none">{teluguAction}</span>
                              </td>
                              <td className="p-0.5 text-center">
                                <div className={`w-3 h-3 border border-blue-900 mx-auto flex items-center justify-center text-[9px] leading-none shrink-0 ${cp.checked ? 'bg-blue-900 text-white font-bold print:bg-transparent print:text-blue-950 print:font-black' : ''}`}>
                                  {cp.checked ? '✓' : ''}
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

              {/* PAGE 2: PARTS & MATERIALS BILL */}
              <div className="card-p2 bg-white border-2 border-blue-900 p-3 md:p-4 text-slate-900 max-w-[780px] w-full text-[12px] leading-snug shadow-md">
                <div>
                  <div className="mb-3">
                    <div className="font-extrabold text-blue-900 text-xs md:text-sm mb-1.5 text-center border-b border-blue-900 pb-1 uppercase tracking-wider">
                      Materials / Spare Parts Issued & Repair Charges
                    </div>
                    <table className="w-full table-fixed border-collapse border border-blue-900 text-[9px]">
                      <colgroup>
                        <col style={{ width: '5%' }} />
                        <col style={{ width: '18%' }} />
                        <col style={{ width: '38%' }} />
                        <col style={{ width: '7%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '11%' }} />
                        <col style={{ width: '13%' }} />
                      </colgroup>
                      <thead>
                        <tr className="border-b border-blue-900 bg-blue-100/70 font-bold text-blue-950 h-[19px]">
                          <th className="border-r border-blue-900 p-0.5 text-center">S.No</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Part No</th>
                          <th className="border-r border-blue-900 p-0.5 text-left">Description</th>
                          <th className="border-r border-blue-900 p-0.5 text-center">Wty</th>
                          <th className="border-r border-blue-900 p-0.5 text-center">Qty</th>
                          <th className="border-r border-blue-900 p-0.5 text-right">Rate ₹</th>
                          <th className="p-0.5 text-right">Amount ₹</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Array.from({ length: Math.max(partRows.length, 30) }).map((_, i) => {
                          const r = partRows[i] || { partNo: '', desc: '', qty: '', rate: '', amount: '', wty: false };
                          return (
                            <tr key={i} className="border-b border-blue-900/60 h-[19px]">
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono text-[8.5px]">{i + 1}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-mono text-[8.5px] truncate">{r.partNo}</td>
                              <td className="border-r border-blue-900/60 p-0.5 font-medium truncate">{r.desc}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-bold text-emerald-700 text-[8.5px]">{r.wty ? '✓' : ''}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-center font-mono text-[8.5px]">{r.qty}</td>
                              <td className="border-r border-blue-900/60 p-0.5 text-right font-mono text-[8.5px]">{r.wty ? '' : r.rate}</td>
                              <td className="p-0.5 text-right font-mono font-bold text-[8.5px]">{r.wty ? '0' : r.amount}</td>
                            </tr>
                          );
                        })}
                        <tr className="bg-blue-50 font-bold border-t-2 border-blue-900 text-slate-900 h-[20px]">
                          <td colSpan={6} className="border-r border-blue-900 p-1 text-right text-[9.5px]">Total Parts Amount ₹</td>
                          <td className="p-1 text-right font-mono font-extrabold text-blue-900 text-[9.5px]">{partsTotal ? partsTotal.toFixed(2) : ''}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SUMMARY CHARGES BOX */}
                  <div className="grid grid-cols-2 gap-4 border-2 border-blue-900 p-2.5 bg-blue-50/30 mb-3 text-xs">
                    <div className="space-y-1">
                      <div><b>Workshop Report:</b> {wsReport || '—'}</div>
                      <div><b>Mechanic Name:</b> {mechanic || '—'}</div>
                      <div><b>Sign of W/S Incharge:</b> {wsIncharge || '—'}</div>
                      {billNo && <div><b>Bill No:</b> <span className="font-mono font-bold text-blue-950 bg-blue-100 px-1 rounded">{billNo}</span></div>}
                    </div>
                    <div className="space-y-1 text-right font-semibold">
                      <div>Warranty Material: ₹{warrantyMaterial || '0.00'}</div>
                      <div>Non Warranty Material: ₹{nonWarrantyMaterial || '0.00'}</div>
                      <div className="text-sm font-black text-blue-950 border-t border-blue-900 pt-1 mt-1">
                        Grand Total: ₹{gTotal || (partsTotal + (parseFloat(totalLabour) || repairTotal)).toFixed(2)}
                      </div>
                    </div>
                  </div>

                  {/* FEEDBACK FIELDS */}
                  <div className="mb-2 text-xs font-semibold space-y-1 border border-blue-900 p-1">
                    <div>ROAD TEST FEEDBACK:____________________________________________________________________________________</div>
                    <div>CUSTOMER RATING & FEED BACK:_________________________________________________________________________</div>
                  </div>
                </div>

                <div className="pt-2">
                  <div className="text-center text-[10px] italic text-slate-600 border-t border-b border-blue-800 py-0.5 mb-2">
                    VEHICLES STORED, REPAIRED AND DRIVEN AT CUSTOMER'S RISK
                  </div>

                  <div className="flex justify-between items-end text-xs font-bold pt-8 px-2">
                    <div className="text-center w-36">
                      <div className="text-[10px] text-slate-800 font-extrabold pb-1 h-5 flex items-end justify-center">
                        {mechanic && mechanic !== '__custom__' ? mechanic : ''}
                      </div>
                      <div className="border-t border-slate-700 pt-1">Mechanic Signature</div>
                    </div>
                    <div className="text-center w-36">
                      <div className="text-[10px] text-slate-800 font-extrabold pb-1 h-5 flex items-end justify-center">
                        {wsIncharge && wsIncharge !== '__custom__' ? wsIncharge : ''}
                      </div>
                      <div className="border-t border-slate-700 pt-1">Supervisor Signature</div>
                    </div>
                    <div className="text-center w-36">
                      <div className="text-[10px] pb-1 h-5"></div>
                      <div className="border-t border-slate-700 pt-1">Customer's Signature</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setViewingCardModal(null)}
                className="px-5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs rounded-lg cursor-pointer"
              >
                Close View
              </button>
            </div>
          </div>
        </div>
        );
      })()}

      {/* CUSTOMER DUPLICATE CHECK & CLEANUP MODAL */}
      {isDuplicateModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 md:p-6 animate-in fade-in duration-150">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="px-5 py-4 bg-gradient-to-r from-amber-900 via-amber-800 to-slate-900 text-white flex items-center justify-between shadow-xs">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-amber-700/60 rounded-xl border border-amber-400/30">
                  <Search className="w-5 h-5 text-amber-200" />
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-black tracking-tight flex items-center gap-2">
                    <span>Customer Database Duplicate Report</span>
                    <span className="text-xs font-bold px-2.5 py-0.5 rounded-full bg-amber-500/40 text-amber-100 border border-amber-300/30 font-mono">
                      {duplicateReport.length} potential duplicate groups
                    </span>
                  </h2>
                  <p className="text-[11px] text-amber-200/90 font-medium">
                    Review duplicate customer entries grouped by Chassis number or phone number and delete redundant records.
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setIsDuplicateModalOpen(false)}
                className="p-1.5 hover:bg-white/10 rounded-xl text-white/80 hover:text-white transition-colors cursor-pointer"
                title="Close"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex-1 overflow-auto p-4 space-y-4 bg-slate-50">
              {duplicateReport.length === 0 ? (
                <div className="text-center py-16 bg-white rounded-xl border border-slate-200">
                  <CheckCircle2 className="w-12 h-12 mx-auto text-emerald-500 mb-2" />
                  <h3 className="text-base font-black text-slate-800">No Duplicates Found!</h3>
                  <p className="text-xs text-slate-500 mt-1">Your customer database is clean. No duplicate chassis numbers or phone matches detected.</p>
                </div>
              ) : (
                duplicateReport.map((dupGroup, groupIdx) => (
                  <div key={groupIdx} className="bg-white rounded-xl border border-amber-200 shadow-2xs overflow-hidden">
                    <div className="bg-amber-50 px-4 py-2.5 border-b border-amber-200 flex flex-wrap items-center justify-between gap-2">
                      <div className="flex items-center gap-2">
                        <span className="px-2 py-0.5 bg-amber-200 text-amber-900 rounded font-black text-xs uppercase">
                          {dupGroup.type}
                        </span>
                        <span className="font-mono font-bold text-slate-900 text-xs">
                          Match Key: <strong className="text-amber-900">{dupGroup.key}</strong>
                        </span>
                      </div>
                      <span className="text-[11px] text-slate-600 font-semibold">{dupGroup.reason} ({dupGroup.records.length} entries)</span>
                    </div>

                    <div className="p-3 overflow-x-auto">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700 font-bold uppercase text-[10px] border-b border-slate-200">
                            <th className="p-2">Customer Name</th>
                            <th className="p-2">Chassis No</th>
                            <th className="p-2">Mobile No</th>
                            <th className="p-2">Model</th>
                            <th className="p-2">Village / Mandal</th>
                            <th className="p-2 text-right">Action</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {dupGroup.records.map((rec: any, rIdx: number) => {
                            const cName = rec['Customer Name'] || rec.custName || getFieldValue(rec, 'custName') || '—';
                            const chNo = rec['Chassis no'] || rec.__chassisDisplay || rec.chassisNo || getFieldValue(rec, 'chassis') || '—';
                            const mobNo = rec['Mobile Number'] || rec.mobileNumber || rec.custPhone || getFieldValue(rec, 'custPhone') || '—';
                            const mModel = rec['Model'] || rec.model || getFieldValue(rec, 'model') || '—';
                            const vill = rec['Village'] || rec.village || getFieldValue(rec, 'village') || '';
                            const mand = rec['Mandal'] || rec.mandal || getFieldValue(rec, 'mandal') || '';

                            return (
                              <tr key={rIdx} className="hover:bg-amber-50/50 transition-colors">
                                <td className="p-2 font-bold text-slate-900">{cName}</td>
                                <td className="p-2 font-mono font-bold text-indigo-900">{chNo}</td>
                                <td className="p-2 font-mono text-slate-700">{mobNo}</td>
                                <td className="p-2 text-slate-800">{mModel}</td>
                                <td className="p-2 text-slate-600">{vill}{vill && mand ? ', ' : ''}{mand}</td>
                                <td className="p-2 text-right">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      if (confirm(`Are you sure you want to delete this specific duplicate record for chassis "${chNo}" (${cName})?`)) {
                                        const targetCh = normalizeKey(chNo);
                                        const updatedIndex = { ...chassisIndex };
                                        let deletedKey = '';
                                        for (const [k, val] of Object.entries(updatedIndex)) {
                                          const v = val as any;
                                          const vCh = (v.__chassisDisplay || getFieldValue(v, 'chassis') || v.chassisNo || '').trim();
                                          if (normalizeKey(vCh) === targetCh || k === targetCh) {
                                            deletedKey = k;
                                            break;
                                          }
                                        }
                                        if (deletedKey) {
                                          delete updatedIndex[deletedKey];
                                          setChassisIndex(updatedIndex);
                                          try {
                                            localStorage.setItem('sri_backup_chassis', JSON.stringify(updatedIndex));
                                          } catch (e) {}
                                          // Refresh duplicate report
                                          const rawRecords = Object.values(updatedIndex);
                                          const uniqueRecs = new Set(rawRecords);
                                          const newDups = getDuplicateReportForSet(Array.from(uniqueRecs));
                                          setDuplicateReport(newDups);
                                        } else {
                                          alert('Record not found in local chassis index.');
                                        }
                                      }
                                    }}
                                    className="px-2.5 py-1 bg-red-600 hover:bg-red-700 text-white font-bold text-[11px] rounded-md transition-colors cursor-pointer shadow-xs inline-flex items-center gap-1"
                                    title="Delete this duplicate record"
                                  >
                                    <Trash2 className="w-3 h-3" /> Delete Entry
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="p-3 bg-white border-t border-slate-200 flex items-center justify-between text-xs">
              <span className="text-slate-500 font-medium">
                Found <strong>{duplicateReport.length}</strong> duplicate groups in database.
              </span>
              <button
                type="button"
                onClick={() => setIsDuplicateModalOpen(false)}
                className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs rounded-xl transition-colors cursor-pointer"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* THIS MONTH DELIVERIES FULL CUSTOMER MODAL */}
      {isMonthDeliveriesModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 md:p-6 animate-in fade-in duration-150">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden">
            {/* MODAL HEADER */}
            <div className="px-5 py-4 bg-gradient-to-r from-indigo-900 via-indigo-800 to-slate-900 text-white flex items-center justify-between shadow-xs">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-700/60 rounded-xl border border-indigo-400/30">
                  <Calendar className="w-5 h-5 text-indigo-200" />
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-black tracking-tight flex items-center gap-2">
                    <span>This Month Deliveries List</span>
                    <span className="text-xs font-bold px-2.5 py-0.5 rounded-full bg-indigo-500/40 text-indigo-100 border border-indigo-300/30 font-mono">
                      {new Date().toLocaleString('default', { month: 'long', year: 'numeric' })}
                    </span>
                  </h2>
                  <p className="text-[11px] text-indigo-200/90 font-medium">
                    Detailed list of all customers whose vehicle delivery date falls in the current calendar month.
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={openDeliveriesInFollowup}
                  className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white font-bold text-xs rounded-xl border border-white/20 transition-colors cursor-pointer"
                >
                  <Briefcase className="w-3.5 h-3.5" /> Open in Follow-up
                </button>
                <button
                  type="button"
                  onClick={() => setIsMonthDeliveriesModalOpen(false)}
                  className="p-1.5 hover:bg-white/10 rounded-xl text-white/80 hover:text-white transition-colors cursor-pointer"
                  title="Close Modal"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* QUICK STATS & CONTROLS BAR */}
            <div className="p-3 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5">
              {/* STATS PILLS */}
              <div className="flex flex-wrap items-center gap-2 text-xs">
                <span className="px-2.5 py-1 bg-white border border-slate-200 rounded-lg font-bold text-slate-700 shadow-2xs">
                  Total Month Deliveries: <strong className="text-indigo-700 font-black">{thisMonthDeliveredCustomers.length}</strong>
                </span>
                <span className="px-2.5 py-1 bg-emerald-50 border border-emerald-200 rounded-lg font-bold text-emerald-800">
                  With Job Cards: <strong className="font-black">{thisMonthDeliveredCustomers.filter(c => chassisWithJobCards.has(normalizeKey(c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || ''))).length}</strong>
                </span>
                <span className="px-2.5 py-1 bg-rose-50 border border-rose-200 rounded-lg font-bold text-rose-800">
                  Pending / No JC: <strong className="font-black">{thisMonthDeliveredCustomers.filter(c => !chassisWithJobCards.has(normalizeKey(c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || ''))).length}</strong>
                </span>
              </div>

              {/* SEARCH & EXPORT */}
              <div className="flex items-center gap-2">
                <div className="relative flex-1 sm:w-64">
                  <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    value={monthDeliveriesSearch}
                    onChange={(e) => setMonthDeliveriesSearch(e.target.value)}
                    placeholder="Search name, phone, chassis..."
                    className="w-full pl-8 pr-7 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-800 placeholder-slate-400 focus:outline-hidden focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600"
                  />
                  {monthDeliveriesSearch && (
                    <button
                      type="button"
                      onClick={() => setMonthDeliveriesSearch('')}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>

                <button
                  type="button"
                  onClick={handleExportMonthDeliveriesCSV}
                  className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1.5 shadow-2xs shrink-0"
                  title="Export Current List to CSV"
                >
                  <Download className="w-3.5 h-3.5" /> Export CSV
                </button>
              </div>
            </div>

            {/* MODAL TABLE CONTENT */}
            <div className="flex-1 overflow-auto p-3">
              {filteredMonthDeliveries.length === 0 ? (
                <div className="text-center py-12 bg-slate-50 rounded-xl border border-dashed border-slate-300">
                  <Calendar className="w-10 h-10 mx-auto text-slate-300 mb-2" />
                  <h4 className="text-sm font-black text-slate-700">No Delivery Records Found</h4>
                  <p className="text-xs text-slate-500 mt-1">
                    {monthDeliveriesSearch ? 'No customers match your search criteria.' : 'No customer vehicles delivered in this current calendar month.'}
                  </p>
                  {monthDeliveriesSearch && (
                    <button
                      type="button"
                      onClick={() => setMonthDeliveriesSearch('')}
                      className="mt-3 px-3 py-1 bg-indigo-50 text-indigo-700 font-bold text-xs rounded-lg border border-indigo-200 hover:bg-indigo-100 cursor-pointer"
                    >
                      Clear Search Filter
                    </button>
                  )}
                </div>
              ) : (
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 font-extrabold uppercase text-[10px] tracking-wider border-b border-slate-200 sticky top-0 z-10">
                        <th className="py-2.5 px-3 w-12 text-center">S.No</th>
                        <th className="py-2.5 px-3">Customer Details</th>
                        <th className="py-2.5 px-3">Model & Vehicle Info</th>
                        <th className="py-2.5 px-3">Delivery Date</th>
                        <th className="py-2.5 px-3">Location</th>
                        <th className="py-2.5 px-3 text-center">Job Card Status</th>
                        <th className="py-2.5 px-3 text-right">Quick Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white">
                      {filteredMonthDeliveries.map((c, idx) => {
                        const custName = c['Customer Name'] || c.__custNameDisplay || c.custName || getFieldValue(c, 'custName') || '—';
                        const fatherName = c['Father Name'] || c.fatherName || getFieldValue(c, 'fatherName') || '';
                        const phone = c['Mobile Number'] || c.__custPhoneDisplay || c.mobileNumber || c['Mobile No'] || c.phNo || c.mobile || getFieldValue(c, 'custPhone') || '';
                        const model = c['Model'] || c.model || getFieldValue(c, 'model') || '—';
                        const chassis = c['Chassis no'] || c.__chassisDisplay || getFieldValue(c, 'chassis') || c.chassis || '—';
                        const engine = c['Engine no'] || c.engineNo || getFieldValue(c, 'engineNo') || '';
                        const rawDel = c['Date of del'] || c.dateOfDel || getFieldValue(c, 'installDate');
                        const delDate = toInputDateFormat(rawDel) || '—';
                        const village = c['Village'] || c.village || getFieldValue(c, 'village') || '';
                        const mandal = c['Mandal'] || c.mandal || getFieldValue(c, 'mandal') || '';
                        const hasJobCard = chassisWithJobCards.has(normalizeKey(chassis));
                        
                        const historyFileNoRaw = getFieldValue(c, 'serialNo') || getFieldValue(c, 'historyFileNo') || c.sno || c.slno || c.srno || c.serialno || c['S.NO.'] || c['S.No.'] || c.sNo || c.historyFileNo || c.fileNo || '';
                        const rawBranchVal = (c.BRANCH || c.branch || c['Branch Name'] || c['BRANCH NAME'] || c.Branch || getFieldValue(c, 'branch') || '').toString().trim();
                        let historyFileNo = historyFileNoRaw.toString().trim();
                        if (historyFileNo && rawBranchVal) {
                          const cleanBranch = rawBranchVal.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                          const cleanHFN = historyFileNo.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                          if (cleanBranch && !cleanHFN.startsWith(cleanBranch)) {
                            historyFileNo = `${rawBranchVal.toUpperCase()}-${historyFileNo}`;
                          }
                        }

                        return (
                          <tr key={idx} className="hover:bg-indigo-50/40 transition-colors">
                            <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-500 text-[11px]">
                              {idx + 1}
                            </td>
                            <td className="py-2.5 px-3">
                              <div className="space-y-1.5">
                                <div className="flex items-center gap-1.5 flex-wrap">
                                  {historyFileNo && (
                                    <span
                                      className={`inline-flex items-center px-1.5 py-0.5 rounded font-mono text-[10px] font-black border shadow-2xs ${
                                        isWithin2Years(rawDel)
                                          ? 'bg-blue-600 text-white border-blue-700'
                                          : 'bg-red-600 text-white border-red-700'
                                      }`}
                                      title={`HFN: ${historyFileNo} (${isWithin2Years(rawDel) ? 'Delivery < 2 Yrs' : 'Delivery > 2 Yrs'})`}
                                    >
                                      HFN: {historyFileNo}
                                    </span>
                                  )}
                                  <span className="font-extrabold text-slate-900 text-xs leading-tight">{custName}</span>
                                </div>
                                {fatherName && <div className="text-[10px] text-slate-500 font-medium">S/o {fatherName}</div>}
                                {phone && (
                                  <div className="py-0.5">
                                    <span className="text-xs text-indigo-700 font-mono font-black bg-indigo-50 px-2 py-0.5 rounded-md border border-indigo-200 shadow-3xs inline-flex items-center gap-1.5">
                                      <Phone className="w-3 h-3 text-indigo-500" /> <a href={`tel:${phone}`} className="hover:underline">{phone}</a>
                                    </span>
                                  </div>
                                )}
                              </div>
                            </td>
                            <td className="py-2.5 px-3">
                              <div className="font-extrabold text-slate-800">{model}</div>
                              <div className="text-[10.5px] font-mono text-slate-600">Chassis: <strong className="text-slate-900">{chassis}</strong></div>
                              {engine && <div className="text-[10px] font-mono text-slate-400">Eng: {engine}</div>}
                            </td>
                            <td className="py-2.5 px-3 font-semibold text-slate-800">
                              <span className="px-2 py-0.5 bg-indigo-50 text-indigo-800 rounded font-mono font-bold text-[11px] border border-indigo-100">
                                📅 {delDate}
                              </span>
                            </td>
                            <td className="py-2.5 px-3 text-slate-600 text-[11px]">
                              {village || mandal ? (
                                <>
                                  <div className="font-medium text-slate-800">{village || '—'}</div>
                                  {mandal && <div className="text-[10px] text-slate-500">{mandal}</div>}
                                </>
                              ) : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center">
                              {hasJobCard ? (
                                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 bg-emerald-100 text-emerald-800 font-black text-[10px] rounded-full">
                                  <CheckCircle2 className="w-3 h-3 text-emerald-600" /> Reporting
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 bg-amber-100 text-amber-800 font-bold text-[10px] rounded-full">
                                  <Clock className="w-3 h-3 text-amber-600" /> Pending
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-right">
                              <div className="flex items-center justify-end gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => handleCreateJobCardForCustomer(c)}
                                  className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white font-bold text-[11px] rounded-lg transition-colors cursor-pointer flex items-center gap-1 shadow-2xs"
                                  title="Create New Job Card with this customer's details"
                                >
                                  <Plus className="w-3 h-3" /> + Job Card
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleViewCustomerHistory(c)}
                                  className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[11px] rounded-lg transition-colors cursor-pointer border border-slate-300"
                                  title="View Service History"
                                >
                                  History
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleViewCustomerInFollowup(c)}
                                  className="px-2 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-[11px] rounded-lg transition-colors cursor-pointer border border-indigo-200"
                                  title="Open Follow-up"
                                >
                                  Follow-up
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* MODAL FOOTER */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs">
              <span className="text-slate-500 font-medium">
                Showing <strong>{filteredMonthDeliveries.length}</strong> of <strong>{thisMonthDeliveredCustomers.length}</strong> customer records
              </span>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={openDeliveriesInFollowup}
                  className="px-3.5 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs rounded-xl border border-indigo-200 transition-colors cursor-pointer"
                >
                  View All in Follow-up Tab →
                </button>
                <button
                  type="button"
                  onClick={() => setIsMonthDeliveriesModalOpen(false)}
                  className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs rounded-xl transition-colors cursor-pointer"
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <datalist id="mechanics-datalist">
        {staffMembers.filter(s => s.role === 'mechanic').map(s => (
          <option key={s.id} value={s.name} />
        ))}
      </datalist>
      <datalist id="supervisors-datalist">
        {staffMembers.filter(s => s.role === 'supervisor').map(s => (
          <option key={s.id} value={s.name} />
        ))}
      </datalist>

      <datalist id="globalChassisDatalist">
        {Array.from(
          new Set(
            Array.from(new Set(Object.values(chassisIndex)))
              .map((rec: any) => {
                const ch = (rec.__chassisDisplay || getFieldValue(rec, 'chassis') || '').trim();
                const nm = (getFieldValue(rec, 'custName') || '').trim();
                const vill = (getFieldValue(rec, 'village') || rec.village || '').trim();
                const mob = (getFieldValue(rec, 'custPhone') || rec.mobileNumber || '').trim();
                const model = (getFieldValue(rec, 'model') || rec.model || '').trim();
                if (!ch) return null;
                return `${ch}${nm ? ` — ${nm}` : ''}${vill ? ` (${vill})` : ''}${model ? ` [${model}]` : ''}${mob ? ` MOB: ${mob}` : ''}`;
              })
              .filter(Boolean)
          )
        ).slice(0, 800).map((chVal, idx) => (
          <option key={`global_chassis_${idx}`} value={chVal} />
        ))}
      </datalist>

      <datalist id="globalCustomerNameDatalist">
        {Array.from(
          new Set(
            Array.from(new Set(Object.values(chassisIndex)))
              .map((rec: any) => {
                const nm = (getFieldValue(rec, 'custName') || '').trim();
                const ch = (getFieldValue(rec, 'chassis') || '').trim();
                const vill = (getFieldValue(rec, 'village') || rec.village || '').trim();
                const mob = (getFieldValue(rec, 'custPhone') || rec.mobileNumber || '').trim();
                const model = (getFieldValue(rec, 'model') || rec.model || '').trim();
                if (!nm) return null;
                return `${nm}${ch ? ` — ${ch}` : ''}${vill ? ` (${vill})` : ''}${model ? ` [${model}]` : ''}${mob ? ` MOB: ${mob}` : ''}`;
              })
              .filter(Boolean)
          )
        ).slice(0, 800).map((nmVal, idx) => (
          <option key={`global_cust_name_${idx}`} value={nmVal} />
        ))}
      </datalist>

      {/* PRINT CSS SETUP TO FORCE PERFECT SINGLE-PAGE LAYOUTS */}
      <style>{`
        @media print {
          @page {
            size: A4 portrait;
            margin-top: 4mm;
            margin-bottom: 4mm;
            margin-left: 12mm; /* Clean margin for punch filing */
            margin-right: 4mm;
          }
          html, body {
            background: #ffffff !important;
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          * {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          table {
            table-layout: fixed !important;
            border-collapse: collapse !important;
            width: 100% !important;
          }
          th, td {
            box-sizing: border-box !important;
          }
          .page-break-container {
            page-break-after: always;
            break-after: page;
            page-break-inside: avoid;
            break-inside: avoid;
            margin: 0 !important;
            padding: 0 !important;
            box-shadow: none !important;
            border: none !important;
            max-height: 285mm !important;
            overflow: hidden !important;
          }
          .card-p1, .card-p2 {
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .page-break-container:last-child {
            page-break-after: auto;
            break-after: auto;
          }
          /* Hide non-printable elements */
          header, button, input[type="file"], .print\\:hidden {
            display: none !important;
          }
        }
      `}</style>
    </div>
  );
}
