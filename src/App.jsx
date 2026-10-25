import React, { useState, useEffect, useCallback } from 'react';
import { HashRouter as Router } from 'react-router-dom';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import {
  Search, Filter, Plus, Edit, Trash2, Eye, X, ChevronLeft, ChevronRight,
  RefreshCw, CheckCircle, Printer, FileSpreadsheet,
  Hash, User, Globe, FileText, Shield, Calendar, Tag, Gauge, Camera, MapPin,
  Users, Building2, FolderOpen, ClipboardList, Bell, LogOut, Lock, Mail,
  Phone, Award, TrendingUp, AlertCircle, Clock, Database, Columns,
  Smartphone, Download, Upload, Activity, Layers, FileCheck,
  Banknote, Wallet, List, Bug, Rocket, Server, Zap, Ruler, TreePine, FileWarning,
  Construction, ArrowRight, Navigation, Compass, ExternalLink
} from 'lucide-react';

// =============================================
// HELPERS: Google Earth & Maps URL
// =============================================
const getGoogleEarthUrl = (coords) => {
  if (!coords) return '';
  const clean = String(coords).replace(/\s+/g, '');
  return `https://earth.google.com/web/@${clean},0a,1000d,35y`;
};
const getGoogleMapsUrl = (coords) => {
  if (!coords) return '';
  const clean = String(coords).replace(/\s+/g, '');
  return `https://www.google.com/maps?q=${clean}`;
};

// =============================================
// MSKVY CONTRACTORS — ALL 127 records
// KMZ verified: PAHAL (15) + TERAVON (10) + SOLTRIUM (22) + SOLACE (7) + NVVP (5) + RKBG (2) = 61
// =============================================
const MSKVY_CONTRACTORS = [
  // Sr.1 KRIVI (4)
  { id: 1, s_no: 1, contractor: 'KRIVI Energy Pvt Ltd', director: 'Mr. Bhupendra Ji', director_contact: '9326275972', file_name: 'KIRVI 01', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '8380868899', substation: '33/11 KV NIGHOJ', area: '34.05', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 2, s_no: 2, contractor: 'KRIVI Energy Pvt Ltd', director: 'Mr. Bhupendra Ji', director_contact: '9326275972', file_name: 'KIRVI 02', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '', substation: '33/11 KV JAWLA', area: '33.55', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 3, s_no: 3, contractor: 'KRIVI Energy Pvt Ltd', director: 'Mr. Bhupendra Ji', director_contact: '9326275972', file_name: 'KIRVI 03', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '', substation: '33/11 KV SUPA OLD', area: '28.1', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 4, s_no: 4, contractor: 'KRIVI Energy Pvt Ltd', director: 'Mr. Bhupendra Ji', director_contact: '9326275972', file_name: 'KIRVI 04', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '', substation: '33/11 KV LONI MAWALA', area: '25.65', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // Sr.2 KYRA (6)
  { id: 5, s_no: 5, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 01', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '', substation: '33/11 KV WADZIRE', area: '24.45', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  { id: 6, s_no: 6, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 02', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '', substation: '33/11 KV SUPA MIDC', area: '11', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  { id: 7, s_no: 7, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 03', district: 'AHILYANAGAR', taluka: 'Pathardi', supervisor: 'Asif Sayyad', contact: '9970289478', substation: '33/11 KV KASAR PIMPALGAON', area: '27.35', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  { id: 8, s_no: 8, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 04', district: 'AHILYANAGAR', taluka: 'Pathardi', supervisor: 'Asif Sayyad', contact: '', substation: '33/11 KV MANIKDAUNDI', area: '14.05', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  { id: 9, s_no: 9, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 05', district: 'AHILYANAGAR', taluka: 'Rahata', supervisor: 'Shrikant Labade', contact: '8010029784', substation: '33/11 KV KORHALE', area: '26.6', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  { id: 10, s_no: 10, contractor: 'KYRA Photon Tech Pvt Ltd', director: 'Mr. V Balaji', director_contact: '9790923111', file_name: 'KYRA 06', district: 'AHILYANAGAR', taluka: 'Sangamner', supervisor: 'Shrikant Labade', contact: '', substation: '33/11 KV PEMGIRI', area: '34.85', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Site engineer visited. Land in KMZ file was Forest Land.' },
  // Sr.3 RK (2)
  { id: 11, s_no: 11, contractor: 'R K Engineers & Consultants', director: 'Mr. Karan Masuria', director_contact: '9601199047', file_name: 'RK 01', district: 'AHILYANAGAR', taluka: 'Sangamner', supervisor: 'Shrikant Labade', contact: '', substation: '33/11 KV NIMON', area: '30.65', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '6 MW plant already connected. Land in KMZ was private, farmer unaware of lease.' },
  { id: 12, s_no: 12, contractor: 'R K Engineers & Consultants', director: 'Mr. Karan Masuria', director_contact: '9601199047', file_name: 'RK 02', district: 'AHILYANAGAR', taluka: 'Shrigonda', supervisor: 'Yogesh Parkhe', contact: '9890440425', substation: '33/11 KV PIMPALGAON PISA', area: '42.65', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Land private, farmer growing sugarcane, unaware of lease.' },
  // Sr.4 RVP (4)
  { id: 13, s_no: 13, contractor: 'SHREE RVP Construction Pvt Ltd', director: '', director_contact: '', file_name: 'RVP 01', district: 'AHILYANAGAR', taluka: 'Karjat', supervisor: 'Yogesh Parkhe', contact: '', substation: '33/11 KV PATEWADI', area: '25.7', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Private land, farmer unaware of lease.' },
  { id: 14, s_no: 14, contractor: 'SHREE RVP Construction Pvt Ltd', director: '', director_contact: '', file_name: 'RVP 02', district: 'AHILYANAGAR', taluka: 'Nevasa', supervisor: 'Asif Sayyad', contact: '', substation: '33/11 KV MIDC PANDHARIPUL', area: '31.40', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Private land, farmer unaware of lease.' },
  { id: 15, s_no: 15, contractor: 'SHREE RVP Construction Pvt Ltd', director: '', director_contact: '', file_name: 'RVP 03', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Parkhe', contact: '', substation: '33/11 KV TAKALI DHOKESHWAR', area: '43.95', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Private land, farmer unaware of lease.' },
  { id: 16, s_no: 16, contractor: 'SHREE RVP Construction Pvt Ltd', director: '', director_contact: '', file_name: 'RVP 04', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Parkhe', contact: '', substation: '33/11 KV RALEGAN THERPAL', area: '34.45', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Private land, farmer unaware of lease.' },
  // ========== Sr.5 SOLACE (7) — KMZ VERIFIED ==========
  { id: 17, s_no: 17, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 01', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Shubham Yadav', contact: '9359874996', substation: '33/11 kV KURHE', area: '26.00', solar: '', coordinates: '21.0301435, 75.1311775', distance_km: '1.410', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 18, s_no: 18, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 02', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Shubham Yadav', contact: '', substation: '33/11 kV BHARWAS', area: '19.75', solar: '', coordinates: '21.0942480, 74.9453075', distance_km: '1.059', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 19, s_no: 19, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 03', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Shubham Yadav', contact: '', substation: '33/11kV GANDHALI', area: '19.05', solar: '', coordinates: '21.1076344, 75.0983474', distance_km: '1.096', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 20, s_no: 20, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 04', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Shubham Yadav', contact: '', substation: '33/11 kV WAWADE', area: '16.05', solar: '', coordinates: '21.0620957, 74.9297792', distance_km: '1.080', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 21, s_no: 21, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 05', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Shubham Yadav', contact: '', substation: '33/11 kV ANTURLI (DHAR)', area: '15.65', solar: '', coordinates: '21.0916978, 75.0509381', distance_km: '1.375', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 22, s_no: 22, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 06', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Prithviraj Bhoi', contact: '9096472314', substation: '33/11 kV JANVA', area: '11.65', solar: '', coordinates: '20.9694776, 74.9678378', distance_km: '1.051', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 23, s_no: 23, contractor: 'SOLACE Cogen Pvt Ltd', director: '', director_contact: '', file_name: 'SOLACE COGEN 07', district: 'JALGAON', taluka: 'Amalner', supervisor: 'Prithviraj Bhoi', contact: '', substation: 'DAHIWAD', area: '6.55', solar: '', coordinates: '21.0775706, 75.1983137', distance_km: '0.998', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Amalner Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  // ========== Sr.6 NVVP (5) — KMZ VERIFIED ==========
  { id: 24, s_no: 24, contractor: 'NVVP Green Energy Pvt Ltd', director: '', director_contact: '', file_name: 'NVVP 01', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '9922140587', substation: '33/11 kV KURWEL', area: '34.25', solar: '', coordinates: '21.1685696, 75.2783386', distance_km: '1.156', obstruction_clearance_m: '350', substation_number: '124093', sr_no: '220', land_status: 'Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 25, s_no: 25, contractor: 'NVVP Green Energy Pvt Ltd', director: '', director_contact: '', file_name: 'NVVP 02', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '', substation: '33/11 kV VIRWADE', area: '27.30', solar: '', coordinates: '21.3063102, 75.3642453', distance_km: '1.133', obstruction_clearance_m: '350', substation_number: '124076', sr_no: '221', land_status: 'Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 26, s_no: 26, contractor: 'NVVP Green Energy Pvt Ltd', director: '', director_contact: '', file_name: 'NVVP 03', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '', substation: '33/11 kV NIMGAVHAN', area: '15.45', solar: '', coordinates: '21.1644767, 75.2256799', distance_km: '1.377', obstruction_clearance_m: '350', substation_number: '124070', sr_no: '222', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 27, s_no: 27, contractor: 'NVVP Green Energy Pvt Ltd', director: '', director_contact: '', file_name: 'NVVP 04', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '', substation: '33/11 kV BUDHGAON', area: '11.15', solar: '', coordinates: '21.1778513, 75.1439688', distance_km: '1.039', obstruction_clearance_m: '350', substation_number: '124172', sr_no: '223', land_status: 'Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 28, s_no: 28, contractor: 'NVVP Green Energy Pvt Ltd', director: '', director_contact: '', file_name: 'NVVP 05', district: 'JALGAON', taluka: 'Erandol', supervisor: 'Gajanan Sonawane', contact: '9529887256', substation: '33/11 kV KASODA', area: '40.65', solar: '', coordinates: '20.8242600, 75.3090049', distance_km: '1.486', obstruction_clearance_m: '350', substation_number: '124065', sr_no: '224', land_status: 'Replacement Search', location: 'Erandol Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  // ========== Sr.7 RKBG (2) — KMZ VERIFIED ==========
  { id: 29, s_no: 29, contractor: 'RKBG Greentech Pvt Ltd', director: '', director_contact: '', file_name: 'RKBG 01', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '', substation: '33/11 kV LASUR', area: '70.70', solar: '', coordinates: '21.3033087, 75.2087132', distance_km: '1.878', obstruction_clearance_m: '350', substation_number: '124075', sr_no: '218', land_status: 'Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  { id: 30, s_no: 30, contractor: 'RKBG Greentech Pvt Ltd', director: '', director_contact: '', file_name: 'RKBG 02', district: 'JALGAON', taluka: 'Chopda', supervisor: 'Gajanan Sonawane', contact: '', substation: '33/11 kV MANGARUL', area: '40.50', solar: '', coordinates: '21.2235286, 75.4026217', distance_km: '1.182', obstruction_clearance_m: '350', substation_number: '124074', sr_no: '219', land_status: 'Replacement Search', location: 'Chopda Taluka, Jalgaon District, Maharashtra', remarks: 'KMZ verified. Zero intersections. Min clearance 350m.' },
  // Sr.8 WAN (5)
  { id: 31, s_no: 31, contractor: 'WAN Technology (OPC) Pvt Ltd', director: '', director_contact: '', file_name: 'WAN 01', district: 'JALGAON', taluka: 'Bhadgaon', supervisor: 'Hemant Agarwal', contact: '', substation: '33 KV KAJGAON', area: '41.1', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC & design team. Farming land, farmer denied lease.' },
  { id: 32, s_no: 32, contractor: 'WAN Technology (OPC) Pvt Ltd', director: '', director_contact: '', file_name: 'WAN 02', district: 'JALGAON', taluka: 'Bhadgaon', supervisor: 'Hemant Agarwal', contact: '', substation: '33 KV GIRAD', area: '29.7', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Farming land, farmer denied lease.' },
  { id: 33, s_no: 33, contractor: 'WAN Technology (OPC) Pvt Ltd', director: '', director_contact: '', file_name: 'WAN 03', district: 'JALGAON', taluka: 'Bhadgaon', supervisor: 'Hemant Agarwal', contact: '', substation: '33 KV LON PIRACHE', area: '22.2', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Farming land, farmer denied lease.' },
  { id: 34, s_no: 34, contractor: 'WAN Technology (OPC) Pvt Ltd', director: '', director_contact: '', file_name: 'WAN 04', district: 'JALGAON', taluka: 'Bhadgaon', supervisor: 'Hemant Agarwal', contact: '', substation: '33 KV AMADADA', area: '5.4', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Survey with EPC. Farming land, farmer denied lease.' },
  { id: 35, s_no: 35, contractor: 'WAN Technology (OPC) Pvt Ltd', director: '', director_contact: '', file_name: 'WAN 05', district: 'JALGAON', taluka: 'Bhusawal', supervisor: 'Hemant Agarwal', contact: '', substation: '33/11 KV VELHALE', area: '39.75', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: 'Farming land, farmer denied lease. Road too thin for heavy vehicles.' },
  // Sr.9 SWELECT (16)
  { id: 36, s_no: 36, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 1', district: 'JALGAON', taluka: 'Jalgaon', supervisor: '', contact: '', substation: '33/11 KV VASANTWADI WAWADADA', area: '15.55', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 37, s_no: 37, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 2', district: 'JALGAON', taluka: 'Jamner', supervisor: 'Mangesh Kale', contact: '8483981271', substation: '33/11 KV FATTEPUR', area: '35.2', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 38, s_no: 38, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 3', district: 'JALGAON', taluka: 'Jamner', supervisor: 'Mangesh Kale', contact: '', substation: '33/11 KV TONDAPUR', area: '35.15', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 39, s_no: 39, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 4', district: 'JALGAON', taluka: 'Jamner', supervisor: 'Mangesh Kale', contact: '', substation: '33/11 KV GARKHEDA', area: '22.25', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 40, s_no: 40, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 5', district: 'JALGAON', taluka: 'Muktainagar', supervisor: 'Mangesh Kale', contact: '', substation: '33/11KV HARTALA', area: '11.3', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 41, s_no: 41, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 6', district: 'JALGAON', taluka: 'Muktainagar', supervisor: 'Mangesh Kale', contact: '', substation: '33/11 KV VADHODA', area: '7.1', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 42, s_no: 42, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 7', district: 'JALGAON', taluka: 'Pachora', supervisor: '', contact: '', substation: '33/11 KV LASGAON', area: '25.4', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 43, s_no: 43, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 8', district: 'JALGAON', taluka: 'Parola', supervisor: 'Prithviraj Bhoi', contact: '', substation: '33/11 KV SAVKHEDA', area: '43.80', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '9 MW plant already connected. Land in KMZ was private land, farmer unaware of lease.' },
  { id: 44, s_no: 44, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 9', district: 'JALGAON', taluka: 'Parola', supervisor: 'Prithviraj Bhoi', contact: '', substation: '33/11 KV BAHADARPUR', area: '26.35', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '5 MW plant already connected. Land in KMZ was Gram Panchayat Land.' },
  { id: 45, s_no: 45, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 10', district: 'JALGAON', taluka: 'Parola', supervisor: 'Prithviraj Bhoi', contact: '', substation: '33/11 KV RATNAPIMPRI', area: '19.45', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '5 MW plant already connected. Land in KMZ was Gram Panchayat Land.' },
  { id: 46, s_no: 46, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 11', district: 'JALGAON', taluka: 'Raver', supervisor: 'Brijesh Verma', contact: '7355135069', substation: '33/11 KV NIMBOI', area: '58.85', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 47, s_no: 47, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 12', district: 'JALGAON', taluka: 'Raver', supervisor: 'Brijesh Verma', contact: '', substation: '33/11 KV CHINAWAL', area: '57.85', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 48, s_no: 48, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 13', district: 'JALGAON', taluka: 'Raver', supervisor: 'Brijesh Verma', contact: '', substation: '33KV MORGAON', area: '47.8', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 49, s_no: 49, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 14', district: 'JALGAON', taluka: 'Raver', supervisor: 'Brijesh Verma', contact: '', substation: '33/11 KV AINPUR', area: '40.4', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 50, s_no: 50, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 15', district: 'JALGAON', taluka: 'Raver', supervisor: 'Brijesh Verma', contact: '', substation: '33/11 KV MOHAGAN', area: '20.4', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 51, s_no: 51, contractor: 'SWELECT Energy Systems Ltd', director: '', director_contact: '', file_name: 'SWELECT 16', district: 'JALGAON', taluka: 'Yawai', supervisor: '', contact: '', substation: '33/11 KV DAHIGAON', area: '46', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // Sr.10 AMBIT (3)
  { id: 52, s_no: 52, contractor: 'AMBIT Energy Pvt Ltd', director: '', director_contact: '', file_name: 'AMBIT 01', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV SUBSTATION SADE', area: '26', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 53, s_no: 53, contractor: 'AMBIT Energy Pvt Ltd', director: '', director_contact: '', file_name: 'AMBIT 02', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV POTEGAON', area: '26', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 54, s_no: 54, contractor: 'AMBIT Energy Pvt Ltd', director: '', director_contact: '', file_name: 'AMBIT 03', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV VITE', area: '21', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // Sr.11 OMSHAM (2)
  { id: 55, s_no: 55, contractor: 'OMSHAM Solutions LLP', director: '', director_contact: '', file_name: 'OMSHAM 01', district: 'SOLAPUR', taluka: 'Madha', supervisor: '', contact: '', substation: '33/11KV PIMPLENEAR', area: '41', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 56, s_no: 56, contractor: 'OMSHAM Solutions LLP', director: '', director_contact: '', file_name: 'OMSHAM 02', district: 'SOLAPUR', taluka: 'Madha', supervisor: '', contact: '', substation: '33/11 KV ALEGAON', area: '26', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // Sr.12 Prashant Plastic (5)
  { id: 57, s_no: 57, contractor: 'Prashant Plastic Industries LLP', director: '', director_contact: '', file_name: 'PRASHANT PLASTIC 01', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV MANJARGAON', area: '16', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 58, s_no: 58, contractor: 'Prashant Plastic Industries LLP', director: '', director_contact: '', file_name: 'PRASHANT PLASTIC 02', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV JINTI', area: '16', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 59, s_no: 59, contractor: 'Prashant Plastic Industries LLP', director: '', director_contact: '', file_name: 'PRASHANT PLASTIC 03', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV PAREWADES', area: '11', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 60, s_no: 60, contractor: 'Prashant Plastic Industries LLP', director: '', director_contact: '', file_name: 'PRASHANT PLASTIC 04', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: '33/11 KV SALSE', area: '11', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 61, s_no: 61, contractor: 'Prashant Plastic Industries LLP', director: '', director_contact: '', file_name: 'PRASHANT PLASTIC 05', district: 'SOLAPUR', taluka: 'Karmala', supervisor: '', contact: '', substation: 'HINGANI', area: '11', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // ========== Sr.13 PAHAL (15) — KMZ VERIFIED ==========
  { id: 62, s_no: 62, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL BORGAON', district: 'NASHIK', taluka: 'Surgana', supervisor: '', contact: '', substation: '33/11 KV BORGAON', area: '24.80', solar: '4.76', coordinates: '20.5150398, 73.7320758', distance_km: '1.122', obstruction_clearance_m: '0', substation_number: '', sr_no: '', land_status: 'Proposed ground-mounted solar project', location: 'Surgana Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Zero intersections.' },
  { id: 63, s_no: 63, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL CHAS', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '', substation: 'CHAS', area: '12.25', solar: '2.25', coordinates: '19.6728002, 74.0814986', distance_km: '1.021', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 64, s_no: 64, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL CHIKHALAMBE', district: 'NASHIK', taluka: 'Chandvad', supervisor: '', contact: '', substation: '33 kV CHIKHALAMBE', area: '50.15', solar: '9.83', coordinates: '20.2478347, 74.0388140', distance_km: '1.175', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 65, s_no: 65, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL DEOPUR', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '', substation: 'DEOPUR', area: '48.35', solar: '9.47', coordinates: '19.8624194, 74.1596466', distance_km: '1.194', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 66, s_no: 66, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL JAIKHEDA', district: 'NASHIK', taluka: 'Satana', supervisor: 'Swapnil Ahire', contact: '8554868750', substation: '33/11 kV JAIKHEDA', area: '52.55', solar: '10.31', coordinates: '20.7805219, 74.2101976', distance_km: '1.479', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Satana Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 67, s_no: 67, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL KAZI-SANGVI', district: 'NASHIK', taluka: 'Chandvad', supervisor: 'Avinash Nikam', contact: '', substation: 'KAZI-SANGVI', area: '55.60', solar: '10.92', coordinates: '20.2328406, 74.2844871', distance_km: '2.092', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 68, s_no: 68, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL KOTAMGAON', district: 'NASHIK', taluka: 'Yeola', supervisor: '', contact: '', substation: '33/11 kV KOTAMGAON', area: '8.50', solar: '1.50', coordinates: '20.0089096, 74.5135529', distance_km: '1.300', obstruction_clearance_m: '350', substation_number: '184281', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Yeola Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 69, s_no: 69, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL MAHAL PATNE', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Sanket Jadhav', contact: '8007382443', substation: 'MAHAL PATNE', area: '32.90', solar: '6.38', coordinates: '20.5176882, 74.3280848', distance_km: '1.417', obstruction_clearance_m: '350', substation_number: '184259', sr_no: '117', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 70, s_no: 70, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL NIMBOLA', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Sanket Jadhav', contact: '', substation: 'NIMBOLA', area: '21.00', solar: '4.00', coordinates: '20.5113461, 74.3370174', distance_km: '2.279', obstruction_clearance_m: '350', substation_number: '184271', sr_no: '118', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 71, s_no: 71, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL NIRPUR', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Sanket Jadhav', contact: '', substation: 'NIRPUR', area: '34.40', solar: '6.68', coordinates: '20.6109740, 74.1384050', distance_km: '1.147', obstruction_clearance_m: '350', substation_number: '184234', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 72, s_no: 72, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL SONAMBE', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '', substation: '33/11 kV SONAMBE', area: '34.30', solar: '6.66', coordinates: '19.7956497, 73.9622351', distance_km: '2.053', obstruction_clearance_m: '350', substation_number: '184217', sr_no: '154', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 73, s_no: 73, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL THANGAON', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '', substation: '33/11 kV THANGAON', area: '37.65', solar: '7.33', coordinates: '19.7104330, 73.9789673', distance_km: '1.775', obstruction_clearance_m: '350', substation_number: '184043', sr_no: '153', land_status: 'Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 74, s_no: 74, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL WADANGALI', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '', substation: 'WADANGALI', area: '42.35', solar: '8.27', coordinates: '19.9195374, 74.1504808', distance_km: '1.779', obstruction_clearance_m: '350', substation_number: '184057', sr_no: '152', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 75, s_no: 75, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL KHAMBALE', district: 'NASHIK', taluka: 'Trimbakeshwar', supervisor: 'Manish Gupta', contact: '8368550313', substation: '33/11 kV KHAMBALE', area: '28.75', solar: '5.55', coordinates: '19.9512815, 73.6245782', distance_km: '1.412', obstruction_clearance_m: '350', substation_number: '184052', sr_no: '', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Trimbakeshwar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 76, s_no: 76, contractor: 'PAHAL Green Energy Pvt Ltd', director: 'Mukesh Patel', director_contact: '9870037830', file_name: 'PAHAL WADZIRE', district: 'NASHIK', taluka: 'Sinnar', supervisor: 'Kiran Deore', contact: '8788624365', substation: '33/11 kV WADZIRE', area: '18.40', solar: '3.48', coordinates: '19.9179711, 73.9647255', distance_km: '1.083', obstruction_clearance_m: '350', substation_number: '184134', sr_no: '155', land_status: 'Replacement Search', location: 'Sinnar Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  // ========== Sr.14 TERAVON (10) — KMZ VERIFIED ==========
  { id: 77, s_no: 77, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON VIRGAON', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Sanket Jadhav', contact: '', substation: 'VIRGAON', area: '71.10', solar: '14.02', coordinates: '20.6773146, 74.1407232', distance_km: '1.276', obstruction_clearance_m: '350', substation_number: '184037', sr_no: '113', land_status: 'Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 78, s_no: 78, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON JORAN', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Sanket Jadhav', contact: '', substation: 'JORAN', area: '62.00', solar: '12.20', coordinates: '20.6555234, 74.0638627', distance_km: '2.392', obstruction_clearance_m: '350', substation_number: '184071', sr_no: '114', land_status: 'Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 79, s_no: 79, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON AJANG', district: 'NASHIK', taluka: 'Malegaon', supervisor: 'Kalpesh Ahire', contact: '8459834191', substation: '33/11 kV AJANG', area: '48.80', solar: '9.56', coordinates: '20.6209890, 74.4622634', distance_km: '1.211', obstruction_clearance_m: '139.8', substation_number: '184203', sr_no: '135', land_status: 'Replacement Search', location: 'Malegaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 139.8m.' },
  { id: 80, s_no: 80, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON KHAMTANE', district: 'NASHIK', taluka: 'Baglan', supervisor: 'Swapnil Ahire', contact: '', substation: '33/11 kV KHAMTANE', area: '35.35', solar: '6.87', coordinates: '20.7842842, 74.1638789', distance_km: '1.118', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Baglan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 81, s_no: 81, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON MOHEGA', district: 'NASHIK', taluka: 'Nandgaon', supervisor: 'Avinash Nikam', contact: '7666037160', substation: '33/11 kV MOHEGAON', area: '21.00', solar: '4.00', coordinates: '20.2572839, 74.5241561', distance_km: '1.100', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Nandgaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 82, s_no: 82, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON RANWAD', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Manish Gupta', contact: '', substation: '132 kV RANWAD', area: '38.70', solar: '7.54', coordinates: '20.1484048, 74.0958339', distance_km: '1.463', obstruction_clearance_m: '350', substation_number: '187006', sr_no: '144', land_status: 'Replacement Search', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 83, s_no: 83, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON SAMANGAON', district: 'NASHIK', taluka: 'Nashik', supervisor: 'Manish Gupta', contact: '', substation: '33/11 kV SAMANGAON', area: '44.75', solar: '8.75', coordinates: '19.9929572, 73.8818530', distance_km: '2.376', obstruction_clearance_m: '186.8', substation_number: '184124', sr_no: '140', land_status: 'Replacement Search', location: 'Nashik Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 186.8m.' },
  { id: 84, s_no: 84, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON THENGODA', district: 'NASHIK', taluka: 'Satana', supervisor: 'Swapnil Ahire', contact: '', substation: 'THENGODA', area: '86.95', solar: '17.19', coordinates: '20.5047361, 74.1817143', distance_km: '3.119', obstruction_clearance_m: '350', substation_number: '184027', sr_no: '149', land_status: 'Replacement Search', location: 'Satana Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 85, s_no: 85, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON VADNER', district: 'NASHIK', taluka: 'Malegaon', supervisor: 'Kalpesh Ahire', contact: '', substation: '33/11 kV VADNER', area: '68.45', solar: '13.49', coordinates: '20.5367553, 74.3879955', distance_km: '3.591', obstruction_clearance_m: '91.1', substation_number: '184033', sr_no: '134', land_status: 'Adaptive Replacement Search', location: 'Malegaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 91.1m.' },
  { id: 86, s_no: 86, contractor: 'TERAVON Green Energies Ltd', director: 'Parth Patel', director_contact: '9723266543', file_name: 'TERAVON VINCHUR', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Akash Tripathi', contact: '6267082203', substation: '33/11 kV VINCHUR', area: '40.10', solar: '7.82', coordinates: '20.0781559, 74.2223381', distance_km: '1.185', obstruction_clearance_m: '350', substation_number: '', sr_no: '', land_status: 'Replacement Search', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  // ========== Sr.15 SOLTRIUM (23) — KMZ VERIFIED ==========
  { id: 87, s_no: 87, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 01', district: 'NASHIK', taluka: 'Chandvad', supervisor: 'Avinash Nikam', contact: '', substation: '33kV SHIRUR TANGADI', area: '28.40', solar: '5.48', coordinates: '20.3621432, 74.1447237', distance_km: '1.708', obstruction_clearance_m: '350', substation_number: '184219', sr_no: '121', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 88, s_no: 88, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 02', district: 'NASHIK', taluka: 'Chandvad', supervisor: 'Kalpesh Ahire', contact: '', substation: 'DUGAON', area: '24.30', solar: '4.66', coordinates: '20.3293939, 74.3434043', distance_km: '2.314', obstruction_clearance_m: '350', substation_number: '184077', sr_no: '122', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 89, s_no: 89, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 03', district: 'NASHIK', taluka: 'Chandvad', supervisor: '', contact: '', substation: '33/11 kV KHADAK-OZAR', area: '22.25', solar: '4.25', coordinates: '20.2535649, 74.1478250', distance_km: '1.993', obstruction_clearance_m: '350', substation_number: '184144', sr_no: '123', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 90, s_no: 90, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 04', district: 'NASHIK', taluka: 'Chandvad', supervisor: 'Avinash Nikam', contact: '', substation: 'JOPUL', area: '9.15', solar: '1.63', coordinates: '20.2436549, 74.1541206', distance_km: '1.004', obstruction_clearance_m: '350', substation_number: '184089', sr_no: '124', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Chandvad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 91, s_no: 91, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 05', district: 'NASHIK', taluka: 'Dindori', supervisor: 'Akshay Kumar Vashishth', contact: '9119008557', substation: '33/11 kV KOSHIMBE', area: '48.65', solar: '9.53', coordinates: '20.3616544, 73.7776373', distance_km: '1.510', obstruction_clearance_m: '350', substation_number: '184250', sr_no: '125', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Dindori Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 92, s_no: 92, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 06', district: 'NASHIK', taluka: 'Dindori', supervisor: 'Akshay Kumar Vashishth', contact: '', substation: '33/11 kV KORATE', area: '33.60', solar: '6.52', coordinates: '20.2110782, 73.9011236', distance_km: '1.161', obstruction_clearance_m: '350', substation_number: '184147', sr_no: '126', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Dindori Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 93, s_no: 93, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 07', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Akash Tripathi', contact: '', substation: '33/11 kV DEOGAON', area: '37.80', solar: '7.36', coordinates: '20.0171640, 74.2584407', distance_km: '1.423', obstruction_clearance_m: '350', substation_number: '184005', sr_no: '145', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 94, s_no: 94, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 08', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Akash Tripathi', contact: '', substation: '132/11 kV OZAR', area: '24.10', solar: '4.62', coordinates: '20.1251649, 73.9520240', distance_km: '2.411', obstruction_clearance_m: '350', substation_number: '187003', sr_no: '146', land_status: 'Validated Existing', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 95, s_no: 95, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 09', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Akash Tripathi', contact: '', substation: '33/11 kV PIMPALGAON', area: '10.20', solar: '1.84', coordinates: '20.1489559, 73.9972238', distance_km: '2.243', obstruction_clearance_m: '350', substation_number: '184003', sr_no: '147', land_status: 'Replacement Search', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 96, s_no: 96, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 10', district: 'NASHIK', taluka: 'Niphad', supervisor: 'Akash Tripathi', contact: '', substation: '33/11 kV DAWACHWADI', area: '8.50', solar: '1.50', coordinates: '20.1448144, 74.0484589', distance_km: '1.031', obstruction_clearance_m: '350', substation_number: '184288', sr_no: '148', land_status: 'Replacement Search', location: 'Niphad Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 97, s_no: 97, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 11', district: 'NASHIK', taluka: 'Nandgaon', supervisor: 'Avinash Nikam', contact: '', substation: 'JALGAON (BK)', area: '14.75', solar: '2.75', coordinates: '20.2602827, 74.7514601', distance_km: '1.051', obstruction_clearance_m: '350', substation_number: '184258', sr_no: '139', land_status: 'Replacement Search', location: 'Nandgaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 98, s_no: 98, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 12', district: 'NASHIK', taluka: 'Nashik', supervisor: 'Manish Gupta', contact: '', substation: '33 KV MUNGSARA', area: '19.75', solar: '3.75', coordinates: '20.0700192, 73.7146988', distance_km: '1.096', obstruction_clearance_m: '350', substation_number: '184127', sr_no: '141', land_status: 'Alternate Land Option 1 – Replacement Search', location: 'Nashik Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 99, s_no: 99, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 13', district: 'NASHIK', taluka: 'Nashik', supervisor: 'Manish Gupta', contact: '', substation: '33/11KV BHAGUR', area: '19.55', solar: '3.71', coordinates: '19.8503844, 73.8656156', distance_km: '4.991', obstruction_clearance_m: '350', substation_number: '184029', sr_no: '142', land_status: 'Replacement Search', location: 'Nashik Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 100, s_no: 100, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 14', district: 'NASHIK', taluka: 'Dindori', supervisor: 'Akshay Kumar Vashishth', contact: '', substation: 'NANASHI', area: '25.10', solar: '4.82', coordinates: '20.3488713, 73.6145100', distance_km: '1.125', obstruction_clearance_m: '350', substation_number: '184010', sr_no: '127', land_status: 'Replacement Search', location: 'Dindori Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 101, s_no: 101, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 15', district: 'NASHIK', taluka: 'Dindori', supervisor: 'Akshay Kumar Vashishth', contact: '', substation: '33/11 KV WARE (DINDORI)', area: '20.85', solar: '3.97', coordinates: '20.3839907, 73.6954599', distance_km: '1.081', obstruction_clearance_m: '350', substation_number: '184248', sr_no: '128', land_status: 'Replacement Search', location: 'Dindori Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 102, s_no: 102, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 16', district: 'NASHIK', taluka: 'Dindori', supervisor: 'Akshay Kumar Vashishth', contact: '', substation: '33/11 KV JANORI', area: '11.00', solar: '2.00', coordinates: '20.0931223, 73.9047615', distance_km: '1.348', obstruction_clearance_m: '350', substation_number: '184150', sr_no: '129', land_status: 'Replacement Search', location: 'Dindori Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 103, s_no: 103, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 17', district: 'NASHIK', taluka: 'Igatpuri', supervisor: '', contact: '', substation: 'SHENWAD', area: '28.05', solar: '5.41', coordinates: '19.7956397, 73.7050793', distance_km: '2.035', obstruction_clearance_m: '350', substation_number: '184265', sr_no: '130', land_status: 'Replacement Search', location: 'Igatpuri Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 104, s_no: 104, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 18', district: 'NASHIK', taluka: 'Igatpuri', supervisor: '', contact: '', substation: 'PADALI', area: '14.10', solar: '2.62', coordinates: '19.8200917, 73.6857766', distance_km: '2.565', obstruction_clearance_m: '350', substation_number: '184266', sr_no: '131', land_status: 'Replacement Search', location: 'Igatpuri Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 105, s_no: 105, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 19', district: 'NASHIK', taluka: 'Kalwan', supervisor: 'Swapnil Ahire', contact: '', substation: '33/11 kV BEI', area: '23.50', solar: '4.50', coordinates: '20.4997183, 74.0603151', distance_km: '1.994', obstruction_clearance_m: '350', substation_number: '184242', sr_no: '132', land_status: 'Replacement Search', location: 'Kalwan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 106, s_no: 106, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 20', district: 'NASHIK', taluka: 'Kalwan', supervisor: 'Swapnil Ahire', contact: '', substation: 'NANDURI', area: '16.65', solar: '3.13', coordinates: '20.4273550, 73.8867976', distance_km: '2.255', obstruction_clearance_m: '350', substation_number: '184076', sr_no: '133', land_status: 'Replacement Search', location: 'Kalwan Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 107, s_no: 107, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 21', district: 'NASHIK', taluka: 'Malegaon', supervisor: 'Kalpesh Ahire', contact: '', substation: '33/11 KV SAVKARWADI', area: '24.90', solar: '4.78', coordinates: '20.3811258, 74.4385746', distance_km: '2.324', obstruction_clearance_m: '350', substation_number: '184274', sr_no: '136', land_status: 'Replacement Search', location: 'Malegaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 108, s_no: 108, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 22', district: 'NASHIK', taluka: 'Malegaon', supervisor: 'Kalpesh Ahire', contact: '', substation: '33/11 kV VIRANI', area: '11.55', solar: '2.11', coordinates: '20.7568888, 74.4705646', distance_km: '1.352', obstruction_clearance_m: '350', substation_number: '184287', sr_no: '137', land_status: 'Replacement Search', location: 'Malegaon Taluka, Nashik District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  { id: 109, s_no: 109, contractor: 'SOLTRIUM Power Pvt Ltd', director: 'Venkatesh Rajaratnam', director_contact: '8056015332', file_name: 'SOLTRIUM 23', district: 'AHILYANAGAR', taluka: 'Sangamner', supervisor: 'Shrikant Labade', contact: '', substation: '33 kV TALEGAON', area: '61.85', solar: '12.17', coordinates: '19.6986009, 74.3104266', distance_km: '1.251', obstruction_clearance_m: '350', substation_number: '14018', sr_no: '109', land_status: 'Replacement Search', location: 'Sangamner Taluka, Ahilyanagar District, Maharashtra', remarks: 'KMZ verified. Min clearance 350m.' },
  // Sr.16 XYZ (18)
  { id: 110, s_no: 110, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ1', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KV SHIRWAL S/S', area: '46', solar: '10.22', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 111, s_no: 111, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ2', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KV KUDAL', area: '46', solar: '10.22', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 112, s_no: 112, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ3', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KV AKKALKOT', area: '31', solar: '6.89', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 113, s_no: 113, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ4', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KV DUDHANI', area: '26', solar: '5.78', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 114, s_no: 114, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ5', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KV GUDDEWADI', area: '26', solar: '5.78', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 115, s_no: 115, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ6', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 MOTYAL', area: '26', solar: '5.78', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 116, s_no: 116, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ7', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 KORSEGAON', area: '21', solar: '4.67', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 117, s_no: 117, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ8', district: 'SOLAPUR', taluka: 'AKKALKOT', supervisor: '', contact: '', substation: '33/11 JAMGAON', area: '36', solar: '8', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 118, s_no: 118, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ9', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV KASARWADI', area: '31', solar: '6.89', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 119, s_no: 119, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ10', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV RATANJAN', area: '26', solar: '5.78', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 120, s_no: 120, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ11', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV DHAMANGAO', area: '26', solar: '5.78', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 121, s_no: 121, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ12', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV BAVI', area: '21', solar: '4.67', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 122, s_no: 122, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ13', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV PATHARI', area: '16', solar: '3.56', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 123, s_no: 123, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ14', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV VAIRAG', area: '11', solar: '2.44', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 124, s_no: 124, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ15', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV JAVALGAON', area: '11', solar: '2.44', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 125, s_no: 125, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ16', district: 'SOLAPUR', taluka: 'BARSHI', supervisor: '', contact: '', substation: '33/11 KV HINGANI', area: '11', solar: '2.44', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 126, s_no: 126, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ17', district: 'SOLAPUR', taluka: 'KARMALA', supervisor: '', contact: '', substation: '33/11 KV KORTI', area: '36', solar: '8', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  { id: 127, s_no: 127, contractor: 'XYZ', director: '', director_contact: '', file_name: 'XYZ18', district: 'SOLAPUR', taluka: 'KARMALA', supervisor: '', contact: '', substation: '33/11 KV WASHIMBE', area: '31', solar: '6.89', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
];

// =============================================
// MSKVY SUBSTATION SURVEYS
// =============================================
const MSKVY_SUBSTATION_SURVEYS = [
  {
    id: 1, s_no: 1, substation: 'KAJGAON 33/11 KV Substation', subdivision: 'Bhadgaon', surveyor: 'Hemant Agrawal',
    district: 'JALGAON', taluka: 'Bhadgaon', contractor: 'WAN Technology (OPC) Pvt Ltd',
    survey_date: '2026-10-22', status: 'Completed',
    voltage_level: '33/11 kV', discom_division: 'Division - Bhadgaon',
    connected_feeders: '5 Agri, 6 Domestic', feeder_length: '10km, 12km, 7km',
    agri_pumps_load: '', peak_load: '', supply_hours: '8 Hours', outage_history: '',
    transformer_capacity: '5 MVA X 4 Nos', transformer_make: '', transformer_year: '',
    loading_percent: '', spare_capacity: '', tap_changer: '33/11',
    oil_condition: 'Good', bushing_condition: 'Good', earthing_condition: 'Good',
    spare_bay: 'New Bay Needed', bay_space: 'Available', switchgear: 'Good',
    protection_relays: 'Yes', scada_rtu: 'Yes', metering: 'Available',
    evacuation_distance: '1.5 Km', right_of_way: 'Not Found Any',
    received_voltage: '11 KV', fault_level: '', lightning_protection: 'Yes',
    fire_safety: 'Yes', feeder_demand: '15 MW', existing_capacity: '5 MW', pending_capacity: '0.464 MW',
    transformer_capacity_check: '4 X 5 MVA',
    sld_available: 'Available', environmental_clearance: 'Yes', feeder_load_data: '89Amp, 41Amp, 73Amp',
    geotagged_photos: 'Attached', coordinates: '20.571375, 75.150647',
    surveyor_sign: 'Signed', discom_sign: 'Signed', remarks: '1 transformer add come to 8 to 10 days'
  },
];

// =============================================
// MSKVY LAND SURVEYS
// =============================================
const MSKVY_LAND_SURVEYS = [
  {
    id: 1, s_no: 1, village: 'KAJGAON', block: 'Bhadgaon', district: 'JALGAON',
    feeder: 'KAJGAON 33/11 KV', surveyor: 'Hemant Agrawal', survey_date: '2026-10-22',
    contractor: 'WAN Technology (OPC) Pvt Ltd', status: 'Completed',
    land_nearby: 'YES', distance_km: '1.7 KM', ownership: 'Private', area_acres: '',
    title_status: 'Farming Land', litigation: '',
    topography: 'No', waterlogging: 'NA', soil_type: '', approach_road: 'access for heavy vehicles',
    water_availability: 'Yes', environmental_clearance: 'NA',
    geotagged_photos: 'Yes', coordinates: '20.561303, 75.159826', survey_map: 'Yes',
    remarks: 'Village KAJGAON, Block Bhadgaon, Dist - Jalgaon, Maharashtra.'
  },
];

// =============================================
// MASTER DATA
// =============================================
const MASTER_MSKVY_DISTRICT = [
  { id: 1, s_no: 1, code: 'AHI', name: 'AHILYANAGAR', state: 'Maharashtra', talukas: 'Parner, Pathardi, Rahata, Sangamner, Shrigonda, Karjat, Nevasa', total_substations: 16, contractors: 'KRIVI Energy, KYRA Photon, R K Engineers, SHREE RVP', status: 'Active', remarks: '' },
  { id: 2, s_no: 2, code: 'JAL', name: 'JALGAON', state: 'Maharashtra', talukas: 'Amalner, Chopda, Erandol, Bhadgaon, Bhusawal, Jalgaon, Jamner, Muktainagar, Pachora, Parola, Raver, Yawai', total_substations: 17, contractors: 'SOLACE, NVVP, RKBG, WAN, SWELECT', status: 'Active', remarks: '' },
  { id: 3, s_no: 3, code: 'SOL', name: 'SOLAPUR', state: 'Maharashtra', talukas: 'Karmala, Madha, Akkalkot, Barshi', total_substations: 10, contractors: 'AMBIT, OMSHAM, Prashant Plastic, XYZ', status: 'Active', remarks: '' },
  { id: 4, s_no: 4, code: 'NAS', name: 'NASHIK', state: 'Maharashtra', talukas: 'Surgana, Sinnar, Chandvad, Satana, Yeola, Baglan, Trimbakeshwar, Malegaon, Nandgaon, Niphad, Nashik, Dindori, Igatpuri, Kalwan', total_substations: 40, contractors: 'PAHAL, TERAVON, SOLTRIUM', status: 'Active', remarks: '' },
];

const MASTER_MSKVY_SURVEY_STATUS = [
  { id: 1, s_no: 1, code: 'PEND', name: 'Pending', status: 'Active', remarks: 'Not started' },
  { id: 2, s_no: 2, code: 'PROG', name: 'In Progress', status: 'Active', remarks: 'Ongoing' },
  { id: 3, s_no: 3, code: 'COMP', name: 'Completed', status: 'Active', remarks: 'Done' },
];

// =============================================
// GLOBAL STYLES
// =============================================
const S = {
  page: { minHeight: '100vh', background: '#f9fafb', padding: 'clamp(10px, 2.5vw, 20px)', fontFamily: 'system-ui, -apple-system, sans-serif', boxSizing: 'border-box' },
  header: { marginBottom: 'clamp(12px, 2.5vw, 20px)' },
  title: { fontSize: 'clamp(16px, 3.5vw, 22px)', fontWeight: 'bold', color: '#111827', margin: 0 },
  subtitle: { fontSize: 'clamp(10px, 2vw, 12px)', color: '#6b7280', marginTop: '4px' },
  statsGrid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 150px), 1fr))', gap: 'clamp(6px, 1.5vw, 12px)', marginBottom: 'clamp(10px, 2.5vw, 16px)' },
  statCard: { background: '#fff', padding: 'clamp(10px, 2vw, 16px)', borderRadius: '12px', border: '1px solid #f3f4f6', boxShadow: '0 1px 3px rgba(0,0,0,0.04)', position: 'relative', overflow: 'hidden' },
  statAccent: { position: 'absolute', top: 0, left: 0, width: '4px', height: '100%' },
  statLabel: { fontSize: 'clamp(9px, 1.8vw, 11px)', color: '#6b7280', margin: 0, fontWeight: '500' },
  statValue: { fontSize: 'clamp(16px, 3.5vw, 24px)', fontWeight: 'bold', margin: '4px 0 0 0', lineHeight: 1 },
  statIconBox: { width: 'clamp(28px, 5vw, 36px)', height: 'clamp(28px, 5vw, 36px)', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  topToolbar: { background: '#fff', padding: 'clamp(8px, 1.8vw, 12px) clamp(10px, 2.5vw, 16px)', borderRadius: '10px', border: '1px solid #e5e7eb', marginBottom: '10px', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 'clamp(6px, 1.5vw, 10px)' },
  searchBar: { background: '#fff', padding: 'clamp(8px, 1.8vw, 12px) clamp(10px, 2.5vw, 16px)', borderRadius: '10px', border: '1px solid #e5e7eb', marginBottom: 'clamp(10px, 2.5vw, 16px)', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 'clamp(6px, 1.5vw, 10px)' },
  searchInput: { width: '100%', padding: '8px 32px', fontSize: 'clamp(11px, 2vw, 13px)', border: '1px solid #d1d5db', borderRadius: '8px', outline: 'none', background: '#fff', color: '#374151', boxSizing: 'border-box' },
  searchIcon: { position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af', pointerEvents: 'none' },
  clearIcon: { position: 'absolute', right: '10px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af', cursor: 'pointer', background: 'transparent', border: 'none' },
  btn: { display: 'inline-flex', alignItems: 'center', gap: '5px', padding: 'clamp(6px, 1.3vw, 8px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '500', borderRadius: '8px', border: '1px solid #d1d5db', background: '#fff', color: '#374151', cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'inherit' },
  btnPrimary: { display: 'inline-flex', alignItems: 'center', gap: '5px', padding: 'clamp(6px, 1.3vw, 8px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '600', borderRadius: '8px', border: 'none', background: '#2563eb', color: '#fff', cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'inherit' },
  btnDanger: { padding: 'clamp(6px, 1.3vw, 8px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', color: '#dc2626', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: '8px', cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'inherit' },
  filterInput: { width: '100%', padding: '8px 12px', fontSize: 'clamp(11px, 2vw, 13px)', border: '1px solid #d1d5db', borderRadius: '8px', outline: 'none', background: '#fff', color: '#374151', boxSizing: 'border-box' },
  tableCard: { background: '#fff', borderRadius: '12px', border: '1px solid #f3f4f6', overflow: 'hidden' },
  tableHeaderBar: { padding: 'clamp(8px, 1.8vw, 12px) clamp(10px, 2.5vw, 16px)', borderBottom: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '8px' },
  tableTitle: { fontSize: 'clamp(12px, 2.2vw, 14px)', fontWeight: '600', color: '#374151', margin: 0 },
  countBadge: { fontSize: '10px', color: '#6b7280', background: '#e5e7eb', padding: '2px 8px', borderRadius: '12px' },
  tableWrap: { overflowX: 'auto', WebkitOverflowScrolling: 'touch' },
  table: { width: '100%', borderCollapse: 'collapse', minWidth: '700px' },
  thead: { background: '#f9fafb' },
  th: { padding: 'clamp(8px, 1.5vw, 10px) clamp(8px, 1.8vw, 12px)', textAlign: 'left', fontSize: 'clamp(9px, 1.7vw, 10px)', fontWeight: '600', color: '#6b7280', textTransform: 'uppercase', letterSpacing: '0.5px', borderBottom: '1px solid #e5e7eb', whiteSpace: 'nowrap' },
  td: { padding: 'clamp(8px, 1.5vw, 10px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', color: '#374151', borderBottom: '1px solid #f3f4f6', whiteSpace: 'nowrap' },
  tdStrong: { padding: 'clamp(8px, 1.5vw, 10px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', color: '#1f2937', fontWeight: '600', borderBottom: '1px solid #f3f4f6', whiteSpace: 'nowrap' },
  iconBtn: { padding: '5px', borderRadius: '6px', border: 'none', background: 'transparent', cursor: 'pointer', display: 'inline-flex', alignItems: 'center' },
  pagination: { padding: 'clamp(8px, 1.8vw, 12px) clamp(10px, 2.5vw, 16px)', borderTop: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '8px' },
  pageBtn: { display: 'inline-flex', alignItems: 'center', gap: '3px', padding: 'clamp(4px, 1vw, 6px) clamp(6px, 1.5vw, 10px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '500', color: '#374151', background: '#fff', border: '1px solid #d1d5db', borderRadius: '6px', cursor: 'pointer', fontFamily: 'inherit' },
  pageBtnActive: { background: '#2563eb', color: '#fff', borderColor: '#2563eb' },
  pageBtnDisabled: { opacity: 0.5, cursor: 'not-allowed' },
  statusBadge: { display: 'inline-block', padding: '3px 8px', borderRadius: '10px', fontSize: 'clamp(9px, 1.7vw, 10px)', fontWeight: '600', whiteSpace: 'nowrap' },
  modalOverlay: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 'clamp(6px, 1.5vw, 12px)', zIndex: 100 },
  modal: { background: '#fff', borderRadius: '12px', width: '100%', maxWidth: '800px', maxHeight: '92vh', overflowY: 'auto' },
  modalHeader: { padding: 'clamp(10px, 2vw, 16px) clamp(12px, 2.5vw, 20px)', borderBottom: '1px solid #e5e7eb', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '8px' },
  modalTitle: { fontSize: 'clamp(13px, 2.5vw, 16px)', fontWeight: 'bold', color: '#111827', margin: 0 },
  modalBody: { padding: 'clamp(10px, 2.5vw, 20px)' },
  modalFooter: { padding: 'clamp(10px, 2vw, 16px) clamp(12px, 2.5vw, 20px)', borderTop: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'flex-end', gap: '6px' },
  formGrid3: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 220px), 1fr))', gap: 'clamp(10px, 2vw, 16px)' },
  formGroup: { display: 'flex', flexDirection: 'column', gap: '5px', minWidth: 0 },
  formGroupFull: { display: 'flex', flexDirection: 'column', gap: '5px', gridColumn: '1 / -1' },
  formLabel: { fontSize: 'clamp(10px, 1.8vw, 11px)', fontWeight: '600', color: '#374151', display: 'flex', alignItems: 'center', gap: '5px' },
  formInput: { width: '100%', padding: 'clamp(7px, 1.5vw, 9px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(11px, 2vw, 13px)', border: '1px solid #d1d5db', borderRadius: '8px', outline: 'none', color: '#374151', background: '#fff', boxSizing: 'border-box', fontFamily: 'inherit' },
  formTextarea: { width: '100%', padding: 'clamp(7px, 1.5vw, 9px) clamp(8px, 1.8vw, 12px)', fontSize: 'clamp(11px, 2vw, 13px)', border: '1px solid #d1d5db', borderRadius: '8px', outline: 'none', color: '#374151', resize: 'vertical', fontFamily: 'inherit', boxSizing: 'border-box' },
  btnSecondary: { display: 'inline-flex', alignItems: 'center', gap: '5px', padding: 'clamp(7px, 1.5vw, 9px) clamp(10px, 2vw, 16px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '500', color: '#374151', background: '#fff', border: '1px solid #d1d5db', borderRadius: '8px', cursor: 'pointer', fontFamily: 'inherit' },
  btnSubmit: { display: 'inline-flex', alignItems: 'center', gap: '5px', padding: 'clamp(7px, 1.5vw, 9px) clamp(10px, 2vw, 16px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '600', color: '#fff', background: '#2563eb', border: '1px solid #2563eb', borderRadius: '8px', cursor: 'pointer', fontFamily: 'inherit' },
};

// =============================================
// HELPERS
// =============================================
const getStatusStyle = (status) => {
  const map = {
    'Approved': { background: '#dcfce7', color: '#166534' },
    'Completed': { background: '#dcfce7', color: '#166534' },
    'Active': { background: '#dcfce7', color: '#166534' },
    'Pending': { background: '#fef3c7', color: '#854d0e' },
    'In Progress': { background: '#dbeafe', color: '#1e40af' },
    'Inactive': { background: '#f3f4f6', color: '#4b5563' },
  };
  return map[status] || { background: '#f3f4f6', color: '#4b5563' };
};
const formatDate = (d) => d ? dayjs(d).format('DD/MM/YYYY') : '—';

const loadMSKVY = (key, fallback) => {
  try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : fallback; } catch { return fallback; }
};
const saveMSKVY = (key, data) => {
  try { localStorage.setItem(key, JSON.stringify(data)); } catch {}
};

// =============================================
// LOGIN
// =============================================
const Login = ({ onLogin }) => {
  const [email, setEmail] = useState('admin@solar.com');
  const [password, setPassword] = useState('admin123');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const submit = (e) => {
    e.preventDefault();
    setLoading(true); setError('');
    setTimeout(() => {
      if (email === 'admin@solar.com' && password === 'admin123') {
        const u = { id: 1, name: 'Admin User', email, role: 'Super Admin' };
        localStorage.setItem('authToken', 'demo-token-' + Date.now());
        localStorage.setItem('user', JSON.stringify(u));
        onLogin(u);
      } else setError('Invalid credentials');
      setLoading(false);
    }, 500);
  };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #1e3a8a, #3b82f6)', padding: 'clamp(10px, 3vw, 20px)', fontFamily: 'system-ui, sans-serif', boxSizing: 'border-box' }}>
      <div style={{ background: '#fff', padding: 'clamp(20px, 5vw, 36px)', borderRadius: '14px', boxShadow: '0 20px 60px rgba(0,0,0,0.3)', width: '100%', maxWidth: '400px', boxSizing: 'border-box' }}>
        <div style={{ width: 'clamp(44px, 10vw, 56px)', height: 'clamp(44px, 10vw, 56px)', background: '#dbeafe', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 14px', color: '#2563eb' }}><Shield size={26} /></div>
        <h1 style={{ fontSize: 'clamp(16px, 4vw, 22px)', fontWeight: 'bold', textAlign: 'center', color: '#1e3a8a', marginBottom: '4px' }}>solar ERP Portal</h1>
        <p style={{ fontSize: 'clamp(10px, 2vw, 12px)', textAlign: 'center', color: '#6b7280', marginBottom: '20px' }}>Solar Project Management</p>
        {error && <div style={{ background: '#fee2e2', color: '#991b1b', padding: '8px 12px', borderRadius: '6px', fontSize: '12px', marginBottom: '12px' }}>{error}</div>}
        <form onSubmit={submit} style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
          <div style={S.formGroup}><label style={S.formLabel}><Mail size={11} /> Email</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} style={S.formInput} required />
          </div>
          <div style={S.formGroup}><label style={S.formLabel}><Lock size={11} /> Password</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={S.formInput} required />
          </div>
          <button type="submit" disabled={loading} style={{ background: '#2563eb', color: '#fff', padding: 'clamp(9px, 2vw, 11px)', border: 'none', borderRadius: '8px', fontSize: 'clamp(12px, 2.2vw, 14px)', fontWeight: '600', cursor: 'pointer', marginTop: '6px', fontFamily: 'inherit' }}>
            {loading ? 'Logging in...' : 'Secure Login'}
          </button>
        </form>
        <div style={{ marginTop: '16px', textAlign: 'center', fontSize: 'clamp(9px, 1.8vw, 11px)', color: '#6b7280', background: '#f3f4f6', padding: '8px', borderRadius: '6px' }}>
          <strong>Demo:</strong> admin@solar.com / admin123
        </div>
      </div>
    </div>
  );
};

// =============================================
// PRINT
// =============================================
const printReport = (data, columns, title) => {
  try {
    if (!data.length) { toast.error('No data to print'); return; }
    const existing = document.getElementById('print-preview-overlay');
    if (existing) existing.remove();

    let tableRows = '';
    data.forEach((item, idx) => {
      tableRows += `<tr><td style="text-align:center;">${idx + 1}</td>${columns.map(c => `<td>${c.exportValue ? c.exportValue(item) : (item[c.key] ?? '—')}</td>`).join('')}</tr>`;
    });

    const overlay = document.createElement('div');
    overlay.id = 'print-preview-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.7);z-index:99999;display:flex;flex-direction:column;padding:10px;box-sizing:border-box;overflow:hidden;';

    overlay.innerHTML = `
      <div style="background:#fff;border-radius:8px;flex:1;display:flex;flex-direction:column;overflow:hidden;max-width:1400px;margin:0 auto;width:100%;">
        <div style="padding:10px 16px;background:#1e3a8a;color:#fff;display:flex;justify-content:space-between;align-items:center;flex-shrink:0;gap:8px;flex-wrap:wrap;">
          <div><div style="font-size:14px;font-weight:700;">Print Preview — ${title}</div><div style="font-size:11px;opacity:0.8;margin-top:2px;">${data.length} records • Landscape</div></div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button id="print-now-btn" style="padding:8px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;font-weight:600;font-size:12px;cursor:pointer;">🖨️ Print Now</button>
            <button id="print-close-btn" style="padding:8px 16px;background:#dc2626;color:#fff;border:none;border-radius:6px;font-weight:600;font-size:12px;cursor:pointer;">✕ Close</button>
          </div>
        </div>
        <div style="flex:1;overflow:auto;padding:12px;background:#f3f4f6;">
          <div style="background:#fff;padding:12px;border-radius:6px;">
            <div style="text-align:center;border-bottom:2px solid #000;padding-bottom:8px;margin-bottom:10px;">
              <div style="font-size:15px;font-weight:700;">INDIAN OIL CORPORATION LIMITED</div>
              <div style="font-size:12px;font-weight:600;margin-top:3px;">${title.toUpperCase()}</div>
              <div style="font-size:10px;color:#555;margin-top:3px;">Generated: ${dayjs().format('DD/MM/YYYY HH:mm')} | Total: ${data.length}</div>
            </div>
            <table id="print-preview-table" style="width:100%;border-collapse:collapse;font-size:10px;">
              <thead><tr style="background:#e5e7eb;">
                <th style="border:1px solid #000;padding:5px 6px;text-align:left;font-weight:700;font-size:9px;">S/No.</th>
                ${columns.map(c => `<th style="border:1px solid #000;padding:5px 6px;text-align:left;font-weight:700;font-size:9px;">${c.label}</th>`).join('')}
              </tr></thead>
              <tbody>${tableRows}</tbody>
            </table>
            <div style="text-align:center;font-size:9px;color:#666;margin-top:8px;padding-top:6px;border-top:1px solid #ccc;">System generated report</div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const styleEl = document.createElement('style');
    styleEl.id = 'print-preview-style';
    styleEl.innerHTML = `@media print { @page { size: A4 landscape; margin: 8mm; } body * { visibility: hidden !important; } #print-preview-overlay, #print-preview-overlay * { visibility: visible !important; } #print-preview-overlay { position: static !important; background: white !important; padding: 0 !important; display: block !important; overflow: visible !important; } #print-preview-overlay > div { max-width: 100% !important; } #print-preview-overlay button { display: none !important; } #print-preview-table { width: 100% !important; } #print-preview-table th, #print-preview-table td { border: 1px solid #000 !important; padding: 3px 4px !important; font-size: 8px !important; word-break: break-word; white-space: normal !important; } }`;
    document.head.appendChild(styleEl);

    document.getElementById('print-now-btn').addEventListener('click', () => window.print());
    document.getElementById('print-close-btn').addEventListener('click', () => {
      overlay.remove();
      const s = document.getElementById('print-preview-style');
      if (s) s.remove();
    });
    toast.success(`Preview opened for ${data.length} records`);
  } catch (e) { console.error(e); toast.error('Print failed'); }
};

const exportToExcel = (data, columns, filename) => {
  try {
    if (!data.length) { toast.error('No data'); return; }
    const wb = XLSX.utils.book_new();
    const excelData = data.map((item, idx) => {
      const row = { 'S/No.': idx + 1 };
      columns.forEach(c => { row[c.label] = c.exportValue ? c.exportValue(item) : (item[c.key] ?? '-'); });
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(excelData);
    ws['!cols'] = [{ wch: 6 }, ...columns.map(() => ({ wch: 20 }))];
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${filename}_${dayjs().format('YYYY-MM-DD')}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast.success(`Exported ${data.length} records`);
  } catch (e) { toast.error('Failed'); }
};

const Field = ({ label, icon, name, value, onChange, type = 'text', placeholder, required, options }) => (
  <div style={S.formGroup}>
    <label style={S.formLabel}>{icon} {label} {required && <span style={{ color: '#ef4444' }}>*</span>}</label>
    {options ? (
      <select name={name} value={value} onChange={onChange} required={required} style={{ ...S.formInput, cursor: 'pointer' }}>
        <option value="">-- Select {label} --</option>
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    ) : type === 'textarea' ? (
      <textarea name={name} value={value} onChange={onChange} placeholder={placeholder} rows={3} style={S.formTextarea} />
    ) : (
      <input type={type} name={name} value={value} onChange={onChange} placeholder={placeholder} required={required} style={S.formInput} />
    )}
  </div>
);

const TopToolbar = ({ itemsPerPage, onItemsPerPageChange, onExcel, onPrint }) => (
  <div style={S.topToolbar}>
    <button onClick={onExcel} style={S.btn}><FileSpreadsheet size={13} /> Excel</button>
    <button onClick={onPrint} style={S.btn}><Printer size={13} /> Print</button>
    <div style={{ display: 'flex', alignItems: 'center', gap: '5px' }}>
      <span style={{ fontSize: 'clamp(10px, 1.9vw, 12px)', color: '#6b7280' }}>Show:</span>
      <select value={itemsPerPage} onChange={e => onItemsPerPageChange(Number(e.target.value))}
        style={{ fontSize: 'clamp(10px, 1.9vw, 12px)', border: '1px solid #d1d5db', borderRadius: '6px', padding: '4px 6px', background: '#fff', cursor: 'pointer', fontFamily: 'inherit' }}>
        <option value="10">10</option><option value="25">25</option><option value="50">50</option><option value="100">100</option>
      </select>
    </div>
  </div>
);

const BottomPagination = ({ currentPage, totalPages, totalItems, itemsPerPage, onPageChange }) => {
  if (totalPages <= 1) return null;
  const pages = [];
  let start = Math.max(1, currentPage - 1);
  let end = Math.min(totalPages, start + 2);
  if (end - start < 2) start = Math.max(1, end - 2);
  for (let i = start; i <= end; i++) pages.push(i);
  return (
    <div style={S.pagination}>
      <div style={{ fontSize: 'clamp(10px, 1.9vw, 12px)', color: '#4b5563' }}>
        Showing <b>{totalItems === 0 ? 0 : ((currentPage - 1) * itemsPerPage) + 1}</b>–<b>{Math.min(currentPage * itemsPerPage, totalItems)}</b> of <b>{totalItems}</b>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: '3px', flexWrap: 'wrap' }}>
        <button onClick={() => onPageChange(currentPage - 1)} disabled={currentPage === 1} style={{ ...S.pageBtn, ...(currentPage === 1 ? S.pageBtnDisabled : {}) }}><ChevronLeft size={12} /> Prev</button>
        {pages.map(p => <button key={p} onClick={() => onPageChange(p)} style={{ ...S.pageBtn, ...(currentPage === p ? S.pageBtnActive : {}) }}>{p}</button>)}
        <button onClick={() => onPageChange(currentPage + 1)} disabled={currentPage === totalPages} style={{ ...S.pageBtn, ...(currentPage === totalPages ? S.pageBtnDisabled : {}) }}>Next <ChevronRight size={12} /></button>
      </div>
    </div>
  );
};

const StatsRow = ({ stats }) => (
  <div style={S.statsGrid}>
    {stats.map((s, i) => (
      <div key={i} style={S.statCard}>
        <div style={{ ...S.statAccent, background: s.color || '#3b82f6' }} />
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '6px' }}>
          <div style={{ minWidth: 0, flex: 1 }}>
            <p style={S.statLabel}>{s.label}</p>
            <p style={{ ...S.statValue, color: s.color || '#1f2937' }}>{s.value}</p>
          </div>
          <div style={{ ...S.statIconBox, background: (s.color || '#3b82f6') + '15', color: s.color || '#3b82f6' }}>{s.icon}</div>
        </div>
      </div>
    ))}
  </div>
);

const ComingSoonPage = ({ title = 'Coming Soon' }) => (
  <div style={S.page}>
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '60vh', textAlign: 'center' }}>
      <div style={{ width: 'clamp(60px, 15vw, 90px)', height: 'clamp(60px, 15vw, 90px)', background: '#dbeafe', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#2563eb', marginBottom: '20px' }}><Construction size={36} /></div>
      <h1 style={{ fontSize: 'clamp(18px, 4vw, 26px)', fontWeight: 'bold', color: '#1f2937', margin: 0, marginBottom: '8px' }}>Coming Soon</h1>
      <p style={{ fontSize: 'clamp(11px, 2vw, 14px)', color: '#6b7280', margin: 0, maxWidth: '400px' }}>{title} module is under development.</p>
    </div>
  </div>
);

// =============================================
// DASHBOARD PAGE
// =============================================
const DashboardPage = ({ onNavigate }) => {
  const totalSubstations = [...new Set(MSKVY_CONTRACTORS.map(c => c.substation))].length;
  const totalArea = MSKVY_CONTRACTORS.reduce((s, c) => s + (parseFloat(c.area) || 0), 0).toFixed(2);
  const totalDistricts = [...new Set(MSKVY_CONTRACTORS.map(c => c.district))].length;
  const totalContractors = [...new Set(MSKVY_CONTRACTORS.map(c => c.contractor))].length;
  const totalTalukas = [...new Set(MSKVY_CONTRACTORS.map(c => c.taluka))].length;
  const totalSupervisors = [...new Set(MSKVY_CONTRACTORS.map(c => c.supervisor).filter(Boolean))].length;
  const totalSolarMW = MSKVY_CONTRACTORS.reduce((s, c) => s + (parseFloat(c.solar) || 0), 0).toFixed(2);
  const kmzVerified = MSKVY_CONTRACTORS.filter(c => c.coordinates).length;

  const cards = [
    { id: 'contractors', title: 'Contractor List', subtitle: 'All EPC Sub-Contractors', value: MSKVY_CONTRACTORS.length, label: 'Total Records', icon: <Building2 size={26} />, color: '#2563eb', bg: '#eff6ff', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'substations', title: 'Substations', subtitle: 'Unique 33/11 KV feeders', value: totalSubstations, label: 'Active Substations', icon: <Zap size={26} />, color: '#ea580c', bg: '#fff7ed', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'districts', title: 'Districts', subtitle: 'Coverage in Maharashtra', value: totalDistricts, label: 'Districts', icon: <MapPin size={26} />, color: '#7c3aed', bg: '#f3e8ff', onClick: () => onNavigate('mskvy', 'mskvy-district') },
    { id: 'talukas', title: 'Talukas', subtitle: 'Sub-district areas', value: totalTalukas, label: 'Talukas', icon: <Globe size={26} />, color: '#0891b2', bg: '#cffafe', onClick: () => onNavigate('mskvy', 'mskvy-district') },
    { id: 'area', title: 'Total Land Area', subtitle: 'Across all sites', value: `${totalArea} ac`, label: 'Total Acres', icon: <Ruler size={26} />, color: '#16a34a', bg: '#dcfce7', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'solar', title: 'Solar Capacity', subtitle: 'Planned installations', value: `${totalSolarMW} MW`, label: 'Solar MW', icon: <Activity size={26} />, color: '#f59e0b', bg: '#fef3c7', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'kmz', title: 'KMZ Verified', subtitle: 'GIS mapped sites', value: kmzVerified, label: 'Sites Verified', icon: <Compass size={26} />, color: '#0d9488', bg: '#ccfbf1', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'supervisors', title: 'Site Supervisors', subtitle: 'Field engineers', value: totalSupervisors, label: 'Team Members', icon: <Users size={26} />, color: '#dc2626', bg: '#fee2e2', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'contractors-unique', title: 'Unique Contractors', subtitle: 'Different companies', value: totalContractors, label: 'Companies', icon: <FolderOpen size={26} />, color: '#9333ea', bg: '#f3e8ff', onClick: () => onNavigate('mskvy', 'mskvy-contractors') },
    { id: 'substation-surveys', title: 'Substation Surveys', subtitle: 'Completed checklist', value: MSKVY_SUBSTATION_SURVEYS.length, label: 'Surveys', icon: <ClipboardList size={26} />, color: '#1e40af', bg: '#dbeafe', onClick: () => onNavigate('mskvy', 'mskvy-substation-survey') },
    { id: 'land-surveys', title: 'Land Surveys', subtitle: 'Land verification done', value: MSKVY_LAND_SURVEYS.length, label: 'Surveys', icon: <FileCheck size={26} />, color: '#15803d', bg: '#dcfce7', onClick: () => onNavigate('mskvy', 'mskvy-land-survey') },
    { id: 'district-master', title: 'District Master', subtitle: 'Manage districts', value: MASTER_MSKVY_DISTRICT.length, label: 'Districts', icon: <Database size={26} />, color: '#ca8a04', bg: '#fef9c3', onClick: () => onNavigate('mskvy', 'mskvy-district') },
  ];

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Management Dashboard</h1><p style={S.subtitle}>Click any card to open the module</p></div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 220px), 1fr))', gap: 'clamp(10px, 2vw, 16px)', marginBottom: 'clamp(14px, 3vw, 24px)' }}>
        {cards.map(card => (
          <div key={card.id} onClick={card.onClick}
            style={{ background: '#fff', borderRadius: '14px', padding: 'clamp(14px, 2.5vw, 20px)', border: '1px solid #e5e7eb', cursor: 'pointer', transition: 'all 0.2s', boxShadow: '0 1px 3px rgba(0,0,0,0.04)', position: 'relative', overflow: 'hidden', userSelect: 'none' }}
            onMouseEnter={e => { e.currentTarget.style.transform = 'translateY(-4px)'; e.currentTarget.style.boxShadow = '0 12px 24px rgba(0,0,0,0.1)'; e.currentTarget.style.borderColor = card.color; }}
            onMouseLeave={e => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.04)'; e.currentTarget.style.borderColor = '#e5e7eb'; }}>
            <div style={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', background: card.color }} />
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
              <div style={{ width: 'clamp(40px, 8vw, 52px)', height: 'clamp(40px, 8vw, 52px)', background: card.bg, color: card.color, borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{card.icon}</div>
              <div style={{ color: '#9ca3af' }}><ArrowRight size={16} /></div>
            </div>
            <p style={{ fontSize: 'clamp(11px, 2vw, 13px)', color: '#6b7280', margin: 0, fontWeight: '600', textTransform: 'uppercase', letterSpacing: '0.3px' }}>{card.title}</p>
            <p style={{ fontSize: 'clamp(18px, 4vw, 28px)', color: card.color, fontWeight: 'bold', margin: '4px 0 2px 0', lineHeight: 1, wordBreak: 'break-word' }}>{card.value}</p>
            <p style={{ fontSize: 'clamp(10px, 1.8vw, 11px)', color: '#9ca3af', margin: 0 }}>{card.label}</p>
          </div>
        ))}
      </div>
    </div>
  );
};

// =============================================
// MSKVY CONTRACTOR LIST — with Google Earth links
// =============================================
const MSKVYContractorPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_contractors_v6', MSKVY_CONTRACTORS));
  const [searchTerm, setSearchTerm] = useState('');
  const [districtFilter, setDistrictFilter] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_contractors_v6', data); }, [data]);

  const emptyForm = { contractor: '', director: '', director_contact: '', file_name: '', district: 'AHILYANAGAR', taluka: '', supervisor: '', contact: '', substation: '', area: '', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);

  const filtered = data.filter(i =>
    (!searchTerm || ['contractor', 'file_name', 'substation', 'district', 'taluka', 'supervisor', 'coordinates', 'substation_number', 'location', 'land_status'].some(k => String(i[k] || '').toLowerCase().includes(searchTerm.toLowerCase())))
    && (!districtFilter || i.district === districtFilter)
  );

  const totalPages = Math.ceil(filtered.length / itemsPerPage) || 1;
  const start = (page - 1) * itemsPerPage;
  const display = filtered.slice(start, start + itemsPerPage);

  const openNew = () => { setFormData(emptyForm); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm(`Delete "${item.contractor} - ${item.file_name}"?`)) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (!formData.contractor || !formData.file_name) { toast.error('Required'); return; }
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false); setEditingId(null);
  };

  const stats = [
    { label: 'Total Records', value: data.length, color: '#2563eb', icon: <Building2 size={18} /> },
    { label: 'Districts', value: [...new Set(data.map(c => c.district))].length, color: '#7c3aed', icon: <MapPin size={18} /> },
    { label: 'Total Area (acres)', value: data.reduce((s, c) => s + (parseFloat(c.area) || 0), 0).toFixed(2), color: '#16a34a', icon: <Ruler size={18} /> },
    { label: 'KMZ Verified', value: data.filter(c => c.coordinates).length, color: '#0d9488', icon: <Compass size={18} /> },
  ];

  const exportCols = [
    { key: 'contractor', label: 'Contractor' }, { key: 'file_name', label: 'File Name' },
    { key: 'district', label: 'District' }, { key: 'taluka', label: 'Taluka' },
    { key: 'supervisor', label: 'Supervisor' }, { key: 'contact', label: 'Contact' },
    { key: 'substation', label: 'Substation' }, { key: 'area', label: 'Area (acres)' },
    { key: 'solar', label: 'Solar (MW)' },
    { key: 'coordinates', label: 'Coordinates' },
    { key: 'distance_km', label: 'Distance (km)' },
    { key: 'obstruction_clearance_m', label: 'Clearance (m)' },
    { key: 'substation_number', label: 'Substation No.' },
    { key: 'sr_no', label: 'Client Sr.No.' },
    { key: 'land_status', label: 'Land Status' },
    { key: 'remarks', label: 'Remarks' },
  ];

  if (showForm) {
    const onChange = e => setFormData({ ...formData, [e.target.name]: e.target.value });
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: 'clamp(12px, 2.5vw, 20px)', display: 'flex', flexWrap: 'wrap', gap: '8px', justifyContent: 'space-between', alignItems: 'center' }}>
            <h1 style={S.title}>{editingId ? 'Edit Contractor' : 'New Contractor'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', border: '1px solid #f3f4f6', padding: 'clamp(12px, 2.5vw, 24px)' }}>
            <form onSubmit={submit}>
              <h3 style={{ fontSize: 'clamp(11px, 2.2vw, 13px)', fontWeight: '700', color: '#1e3a8a', marginBottom: '12px', paddingBottom: '6px', borderBottom: '2px solid #dbeafe' }}>Contractor Info</h3>
              <div style={S.formGrid3}>
                <Field label="Contractor" icon={<Building2 size={11} />} name="contractor" value={formData.contractor} onChange={onChange} required />
                <Field label="Director" icon={<User size={11} />} name="director" value={formData.director} onChange={onChange} />
                <Field label="Director Contact" icon={<Phone size={11} />} name="director_contact" value={formData.director_contact} onChange={onChange} />
                <Field label="File Name" icon={<FileText size={11} />} name="file_name" value={formData.file_name} onChange={onChange} required />
                <Field label="District" icon={<MapPin size={11} />} name="district" value={formData.district} onChange={onChange} options={['AHILYANAGAR', 'JALGAON', 'SOLAPUR', 'NASHIK']} />
                <Field label="Taluka" icon={<MapPin size={11} />} name="taluka" value={formData.taluka} onChange={onChange} />
                <Field label="Site Supervisor" icon={<User size={11} />} name="supervisor" value={formData.supervisor} onChange={onChange} />
                <Field label="Supervisor Contact" icon={<Phone size={11} />} name="contact" value={formData.contact} onChange={onChange} />
                <Field label="Substation" icon={<Zap size={11} />} name="substation" value={formData.substation} onChange={onChange} />
                <Field label="Area (acres)" icon={<Ruler size={11} />} name="area" type="number" value={formData.area} onChange={onChange} />
                <Field label="Solar (MW)" icon={<Gauge size={11} />} name="solar" type="number" value={formData.solar} onChange={onChange} />
              </div>
              <h3 style={{ fontSize: 'clamp(11px, 2.2vw, 13px)', fontWeight: '700', color: '#0d9488', marginBottom: '12px', marginTop: '20px', paddingBottom: '6px', borderBottom: '2px solid #ccfbf1' }}>KMZ GIS Data</h3>
              <div style={S.formGrid3}>
                <Field label="Coordinates (lat, long)" icon={<Compass size={11} />} name="coordinates" value={formData.coordinates} onChange={onChange} placeholder="20.3621432, 74.1447237" />
                <Field label="Distance to Substation (km)" icon={<Navigation size={11} />} name="distance_km" type="number" value={formData.distance_km} onChange={onChange} />
                <Field label="Obstruction Clearance (m)" icon={<Shield size={11} />} name="obstruction_clearance_m" type="number" value={formData.obstruction_clearance_m} onChange={onChange} />
                <Field label="Substation Number" icon={<Hash size={11} />} name="substation_number" value={formData.substation_number} onChange={onChange} />
                <Field label="Client Sr. No." icon={<Hash size={11} />} name="sr_no" value={formData.sr_no} onChange={onChange} />
                <Field label="Land Status" icon={<FileCheck size={11} />} name="land_status" value={formData.land_status} onChange={onChange} />
                <div style={S.formGroupFull}><label style={S.formLabel}><MapPin size={11} /> Location</label>
                  <input type="text" name="location" value={formData.location} onChange={onChange} style={S.formInput} />
                </div>
              </div>
              <div style={{ marginTop: '20px' }}>
                <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={11} /> Remarks</label>
                  <textarea name="remarks" value={formData.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
                </div>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'flex-end', gap: '8px', paddingTop: 'clamp(12px, 2.5vw, 20px)', marginTop: 'clamp(12px, 2.5vw, 20px)', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={13} /> {editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>MSKVY - Contractor List</h1><p style={S.subtitle}>EPC Sub-Contractors & Site Details (KMZ GIS verified)</p></div>
      <StatsRow stats={stats} />

      <TopToolbar itemsPerPage={itemsPerPage} onItemsPerPageChange={(v) => { setItemsPerPage(v); setPage(1); }}
        onExcel={() => exportToExcel(filtered, exportCols, 'mskvy_contractors')}
        onPrint={() => printReport(filtered, exportCols, 'MSKVY Contractor List')} />

      <div style={S.searchBar}>
        <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 0 }}>
          <Search size={14} style={S.searchIcon} />
          <input type="text" placeholder="Search contractor, file, substation, coordinates..." value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => { setSearchTerm(''); setPage(1); }} style={S.clearIcon}><X size={12} /></button>}
        </div>
        <select value={districtFilter} onChange={e => { setDistrictFilter(e.target.value); setPage(1); }} style={{ ...S.filterInput, width: 'auto', minWidth: '110px', flex: '0 1 auto' }}>
          <option value="">All Districts</option>
          {[...new Set(data.map(c => c.district))].map(d => <option key={d} value={d}>{d}</option>)}
        </select>
        <button onClick={() => { setSearchTerm(''); setDistrictFilter(''); setPage(1); }} style={S.btnDanger}>Clear</button>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={13} /> Add</button>
      </div>

      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            <h2 style={S.tableTitle}>All Contractor Records</h2>
            <span style={S.countBadge}>{filtered.length}</span>
          </div>
        </div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}>
              <tr>
                <th style={S.th}>S/No.</th><th style={S.th}>Contractor</th><th style={S.th}>File Name</th>
                <th style={S.th}>District</th><th style={S.th}>Taluka</th><th style={S.th}>Supervisor</th>
                <th style={S.th}>Substation</th><th style={S.th}>Area</th><th style={S.th}>Solar (MW)</th>
                <th style={S.th}>Coordinates</th><th style={S.th}>Dist (km)</th><th style={S.th}>Clearance (m)</th>
                <th style={S.th}>Sub No.</th><th style={S.th}>Client Sr.</th><th style={S.th}>Land Status</th>
                <th style={S.th}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {display.map((item, idx) => (
                <tr key={item.id}>
                  <td style={S.td}>{start + idx + 1}</td>
                  <td style={S.tdStrong}>{item.contractor}</td>
                  <td style={S.td}>{item.file_name}</td>
                  <td style={S.td}><span style={{ ...S.statusBadge, background: '#dbeafe', color: '#1e40af' }}>{item.district}</span></td>
                  <td style={S.td}>{item.taluka}</td>
                  <td style={S.td}>{item.supervisor || '—'}</td>
                  <td style={S.td}>{item.substation}</td>
                  <td style={S.td}>{item.area}</td>
                  <td style={S.td}>{item.solar || '—'}</td>
                  <td style={S.td}>
                    {item.coordinates ? (
                      <a href={getGoogleEarthUrl(item.coordinates)} target="_blank" rel="noopener noreferrer" title="Click to open in Google Earth"
                        style={{ fontSize: '10px', color: '#0d9488', fontWeight: '600', display: 'inline-flex', alignItems: 'center', gap: '3px', textDecoration: 'none', cursor: 'pointer', padding: '3px 6px', borderRadius: '4px', background: '#ccfbf1', border: '1px solid #5eead4', transition: 'all 0.15s' }}
                        onMouseEnter={e => { e.currentTarget.style.background = '#0d9488'; e.currentTarget.style.color = '#fff'; }}
                        onMouseLeave={e => { e.currentTarget.style.background = '#ccfbf1'; e.currentTarget.style.color = '#0d9488'; }}>
                        <Compass size={11} /> {item.coordinates} <ExternalLink size={9} />
                      </a>
                    ) : '—'}
                  </td>
                  <td style={S.td}>{item.distance_km ? <span style={{ fontSize: '11px', color: '#2563eb', fontWeight: '600' }}>{item.distance_km} km</span> : '—'}</td>
                  <td style={S.td}>{item.obstruction_clearance_m ? <span style={{ fontSize: '11px', color: '#16a34a', fontWeight: '600' }}>{item.obstruction_clearance_m} m</span> : '—'}</td>
                  <td style={S.td}>{item.substation_number ? <span style={{ fontSize: '10px', background: '#fef3c7', color: '#92400e', padding: '2px 6px', borderRadius: '4px', fontWeight: '600' }}>{item.substation_number}</span> : '—'}</td>
                  <td style={S.td}>{item.sr_no ? <span style={{ fontSize: '10px', background: '#f3e8ff', color: '#6b21a8', padding: '2px 6px', borderRadius: '4px', fontWeight: '600' }}>{item.sr_no}</span> : '—'}</td>
                  <td style={S.td}>{item.land_status ? <span style={{ fontSize: '10px', background: '#dcfce7', color: '#166534', padding: '2px 6px', borderRadius: '4px', fontWeight: '600', whiteSpace: 'nowrap' }}>{item.land_status.length > 25 ? item.land_status.substring(0, 25) + '...' : item.land_status}</span> : '—'}</td>
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '2px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }} title="View"><Eye size={14} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }} title="Edit"><Edit size={14} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }} title="Delete"><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <BottomPagination currentPage={page} totalPages={totalPages} totalItems={filtered.length} itemsPerPage={itemsPerPage} onPageChange={setPage} />
      </div>

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={{ ...S.modal, maxWidth: '900px' }} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}><h2 style={S.modalTitle}>Contractor Details</h2><button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button></div>
            <div style={S.modalBody}>
              <h3 style={{ fontSize: '11px', fontWeight: '700', color: '#1e3a8a', marginBottom: '8px', paddingBottom: '5px', borderBottom: '2px solid #dbeafe' }}>CONTRACTOR INFO</h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: 'clamp(8px, 1.5vw, 12px)', marginBottom: '14px' }}>
                {[['Contractor', selected.contractor], ['File Name', selected.file_name], ['Director', selected.director], ['Director Contact', selected.director_contact], ['District', selected.district], ['Taluka', selected.taluka], ['Site Supervisor', selected.supervisor], ['Supervisor Contact', selected.contact], ['Substation', selected.substation], ['Area (acres)', selected.area], ['Solar (MW)', selected.solar || '—']].map(([l, v]) => (
                  <div key={l}>
                    <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>{l}</p>
                    <p style={{ fontSize: 'clamp(11px, 2vw, 13px)', color: '#111827', margin: 0, wordBreak: 'break-word' }}>{v || '—'}</p>
                  </div>
                ))}
              </div>

              {(selected.coordinates || selected.distance_km || selected.obstruction_clearance_m || selected.substation_number || selected.sr_no || selected.land_status || selected.location) && (
                <>
                  <h3 style={{ fontSize: '11px', fontWeight: '700', color: '#0d9488', marginBottom: '8px', marginTop: '14px', paddingBottom: '5px', borderBottom: '2px solid #ccfbf1' }}>KMZ GIS VERIFIED DATA</h3>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: 'clamp(8px, 1.5vw, 12px)' }}>
                    {selected.coordinates && (
                      <div style={{ gridColumn: '1 / -1' }}>
                        <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 6px 0', fontWeight: '600', textTransform: 'uppercase' }}>Coordinates</p>
                        <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '8px' }}>
                          <span style={{ fontSize: 'clamp(12px, 2.2vw, 14px)', color: '#0d9488', fontWeight: '700', display: 'inline-flex', alignItems: 'center', gap: '4px', padding: '6px 10px', background: '#ccfbf1', borderRadius: '6px', border: '1px solid #5eead4' }}>
                            <Compass size={14} /> {selected.coordinates}
                          </span>
                          <a href={getGoogleEarthUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                            style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#0d9488', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600', border: 'none', cursor: 'pointer', transition: 'all 0.15s' }}
                            onMouseEnter={e => { e.currentTarget.style.background = '#0f766e'; }}
                            onMouseLeave={e => { e.currentTarget.style.background = '#0d9488'; }}>
                            <Globe size={14} /> Open in Google Earth
                          </a>
                          <a href={getGoogleMapsUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                            style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#2563eb', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600', border: 'none', cursor: 'pointer', transition: 'all 0.15s' }}
                            onMouseEnter={e => { e.currentTarget.style.background = '#1e40af'; }}
                            onMouseLeave={e => { e.currentTarget.style.background = '#2563eb'; }}>
                            <MapPin size={14} /> Google Maps
                          </a>
                        </div>
                      </div>
                    )}
                    {selected.distance_km && <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Distance to Substation</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', color: '#2563eb', fontWeight: '600', margin: 0 }}>{selected.distance_km} km</p></div>}
                    {selected.obstruction_clearance_m && <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Obstruction Clearance</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', color: '#16a34a', fontWeight: '600', margin: 0 }}>{selected.obstruction_clearance_m} m</p></div>}
                    {selected.substation_number && <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Substation Number</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', margin: 0, fontWeight: '600' }}>{selected.substation_number}</p></div>}
                    {selected.sr_no && <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Client Sr. No.</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', margin: 0, fontWeight: '600' }}>{selected.sr_no}</p></div>}
                    {selected.land_status && <div style={{ gridColumn: '1 / -1' }}><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Land Selection Status</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', margin: 0, color: '#166534', fontWeight: '600' }}>{selected.land_status}</p></div>}
                    {selected.location && <div style={{ gridColumn: '1 / -1' }}><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Location</p><p style={{ fontSize: 'clamp(11px, 2vw, 13px)', margin: 0 }}><MapPin size={12} style={{ display: 'inline', marginRight: '4px' }} />{selected.location}</p></div>}
                  </div>
                </>
              )}

              {selected.remarks && (
                <div style={{ marginTop: '14px', padding: '10px', background: '#fef3c7', borderRadius: '6px' }}>
                  <p style={{ fontSize: '10px', color: '#92400e', fontWeight: '700', margin: '0 0 3px 0' }}>REMARKS</p>
                  <p style={{ fontSize: 'clamp(11px, 2vw, 13px)', color: '#78350f', margin: 0, wordBreak: 'break-word' }}>{selected.remarks}</p>
                </div>
              )}
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={13} /> Edit</button>
              <button onClick={() => del(selected)} style={{ ...S.btnSecondary, color: '#dc2626', borderColor: '#fecaca' }}><Trash2 size={13} /> Delete</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// DISTRICT MASTER
// =============================================
const DistrictMasterPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_district_v5', MASTER_MSKVY_DISTRICT));
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [showView, setShowView] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_district_v5', data); }, [data]);

  const emptyForm = { code: '', name: '', state: 'Maharashtra', talukas: '', total_substations: '', contractors: '', status: 'Active', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);
  const filtered = data.filter(i => !searchTerm || ['name', 'code', 'talukas', 'contractors', 'state'].some(k => String(i[k] || '').toLowerCase().includes(searchTerm.toLowerCase())));
  const totalPages = Math.ceil(filtered.length / itemsPerPage) || 1;
  const start = (page - 1) * itemsPerPage;
  const display = filtered.slice(start, start + itemsPerPage);

  const openNew = () => { setFormData(emptyForm); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm(`Delete "${item.name}"?`)) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (!formData.name) { toast.error('Required'); return; }
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false); setEditingId(null);
  };

  const columns = [
    { key: 'code', label: 'Code' }, { key: 'name', label: 'District Name', strong: true },
    { key: 'state', label: 'State' }, { key: 'talukas', label: 'Talukas' },
    { key: 'total_substations', label: 'Substations' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];

  if (showForm) {
    const onChange = e => setFormData({ ...formData, [e.target.name]: e.target.value });
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: 'clamp(12px, 2.5vw, 20px)', display: 'flex', flexWrap: 'wrap', gap: '8px', justifyContent: 'space-between', alignItems: 'center' }}>
            <h1 style={S.title}>{editingId ? 'Edit District' : 'New District'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', border: '1px solid #f3f4f6', padding: 'clamp(12px, 2.5vw, 24px)' }}>
            <form onSubmit={submit}>
              <div style={S.formGrid3}>
                <Field label="Code" icon={<Hash size={11} />} name="code" value={formData.code} onChange={onChange} />
                <Field label="District Name" icon={<MapPin size={11} />} name="name" value={formData.name} onChange={onChange} required />
                <Field label="State" icon={<Globe size={11} />} name="state" value={formData.state} onChange={onChange} />
                <Field label="Total Substations" icon={<Zap size={11} />} name="total_substations" type="number" value={formData.total_substations} onChange={onChange} />
                <Field label="Status" icon={<Shield size={11} />} name="status" value={formData.status} onChange={onChange} options={['Active', 'Inactive']} />
                <div style={S.formGroupFull}><label style={S.formLabel}><MapPin size={11} /> Talukas</label>
                  <textarea name="talukas" value={formData.talukas} onChange={onChange} rows={2} style={S.formTextarea} />
                </div>
                <div style={S.formGroupFull}><label style={S.formLabel}><Building2 size={11} /> Contractors</label>
                  <textarea name="contractors" value={formData.contractors} onChange={onChange} rows={2} style={S.formTextarea} />
                </div>
                <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={11} /> Remarks</label>
                  <textarea name="remarks" value={formData.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
                </div>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'flex-end', gap: '8px', paddingTop: 'clamp(12px, 2.5vw, 20px)', marginTop: 'clamp(12px, 2.5vw, 20px)', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={13} /> {editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>District Master</h1><p style={S.subtitle}>All districts with talukas, substations & contractors</p></div>

      <TopToolbar itemsPerPage={itemsPerPage} onItemsPerPageChange={(v) => { setItemsPerPage(v); setPage(1); }}
        onExcel={() => exportToExcel(filtered, columns, 'district_master')}
        onPrint={() => printReport(filtered, columns, 'District Master')} />

      <div style={S.searchBar}>
        <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 0 }}>
          <Search size={14} style={S.searchIcon} />
          <input type="text" placeholder="Search district, taluka, contractor..." value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => { setSearchTerm(''); setPage(1); }} style={S.clearIcon}><X size={12} /></button>}
        </div>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={13} /> Add District</button>
      </div>

      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Districts</h2><span style={S.countBadge}>{filtered.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr></thead>
            <tbody>
              {display.map((item, idx) => (
                <tr key={item.id}>
                  <td style={S.td}>{start + idx + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(item) : (item[c.key] ?? '—')}</td>)}
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '2px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={14} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={14} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <BottomPagination currentPage={page} totalPages={totalPages} totalItems={filtered.length} itemsPerPage={itemsPerPage} onPageChange={setPage} />
      </div>

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}><h2 style={S.modalTitle}>District Details</h2><button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button></div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: '12px' }}>
                {[['Code', selected.code], ['District', selected.name], ['State', selected.state], ['Substations', selected.total_substations], ['Status', selected.status]].map(([l, v]) => (
                  <div key={l}>
                    <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>{l}</p>
                    <p style={{ fontSize: '13px', margin: 0 }}>{v || '—'}</p>
                  </div>
                ))}
              </div>
              {selected.talukas && <div style={{ marginTop: '12px', padding: '10px', background: '#eff6ff', borderRadius: '6px' }}><p style={{ fontSize: '10px', fontWeight: '700', color: '#1e40af', margin: '0 0 3px 0' }}>TALUKAS</p><p style={{ fontSize: '12px', color: '#1e3a8a', margin: 0 }}>{selected.talukas}</p></div>}
              {selected.contractors && <div style={{ marginTop: '10px', padding: '10px', background: '#f3e8ff', borderRadius: '6px' }}><p style={{ fontSize: '10px', fontWeight: '700', color: '#6b21a8', margin: '0 0 3px 0' }}>CONTRACTORS</p><p style={{ fontSize: '12px', color: '#581c87', margin: 0 }}>{selected.contractors}</p></div>}
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={13} /> Edit</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// SURVEY STATUS MASTER
// =============================================
const SurveyStatusMasterPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_survey_status_v5', MASTER_MSKVY_SURVEY_STATUS));
  const [showForm, setShowForm] = useState(false);
  const [showView, setShowView] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const emptyForm = { code: '', name: '', status: 'Active', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);
  useEffect(() => { saveMSKVY('mskvy_survey_status_v5', data); }, [data]);

  const openNew = () => { setFormData(emptyForm); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm('Delete?')) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (!formData.name) { toast.error('Required'); return; }
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false); setEditingId(null);
  };

  const columns = [
    { key: 'code', label: 'Code' }, { key: 'name', label: 'Name', strong: true },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
    { key: 'remarks', label: 'Remarks' },
  ];

  if (showForm) {
    const onChange = e => setFormData({ ...formData, [e.target.name]: e.target.value });
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '800px', margin: '0 auto' }}>
          <div style={{ marginBottom: 'clamp(12px, 2.5vw, 20px)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
            <h1 style={S.title}>{editingId ? 'Edit Status' : 'New Status'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', padding: 'clamp(12px, 2.5vw, 24px)', border: '1px solid #f3f4f6' }}>
            <form onSubmit={submit}>
              <div style={S.formGrid3}>
                <Field label="Code" icon={<Hash size={11} />} name="code" value={formData.code} onChange={onChange} />
                <Field label="Name" icon={<Tag size={11} />} name="name" value={formData.name} onChange={onChange} required />
                <Field label="Status" icon={<Shield size={11} />} name="status" value={formData.status} onChange={onChange} options={['Active', 'Inactive']} />
                <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={11} /> Remarks</label>
                  <textarea name="remarks" value={formData.remarks} onChange={onChange} rows={2} style={S.formTextarea} />
                </div>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '20px', paddingTop: '16px', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={13} /> {editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Survey Status Master</h1><p style={S.subtitle}>Manage survey statuses</p></div>

      <TopToolbar itemsPerPage={10} onItemsPerPageChange={() => {}}
        onExcel={() => exportToExcel(data, columns, 'survey_status')}
        onPrint={() => printReport(data, columns, 'Survey Status')} />

      <div style={S.searchBar}>
        <div style={{ flex: 1 }}></div>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={13} /> Add</button>
      </div>

      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Statuses</h2><span style={S.countBadge}>{data.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr></thead>
            <tbody>
              {data.map((item, idx) => (
                <tr key={item.id}>
                  <td style={S.td}>{idx + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(item) : (item[c.key] ?? '—')}</td>)}
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '2px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={14} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={14} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}><h2 style={S.modalTitle}>Status Details</h2><button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button></div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: '12px' }}>
                <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Code</p><p style={{ fontSize: '13px', margin: 0 }}>{selected.code || '—'}</p></div>
                <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Name</p><p style={{ fontSize: '13px', fontWeight: '600', margin: 0 }}>{selected.name}</p></div>
                <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Status</p><p style={{ margin: 0 }}><span style={{ ...S.statusBadge, ...getStatusStyle(selected.status) }}>{selected.status}</span></p></div>
                <div><p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>Remarks</p><p style={{ fontSize: '13px', margin: 0 }}>{selected.remarks || '—'}</p></div>
              </div>
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={13} /> Edit</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// SUBSTATION SURVEY PAGE
// =============================================
const MSKVYSubstationSurveyPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_substation_v5', MSKVY_SUBSTATION_SURVEYS));
  const [searchTerm, setSearchTerm] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_substation_v5', data); }, [data]);

  const emptyForm = { substation: '', subdivision: '', surveyor: '', district: 'JALGAON', taluka: '', contractor: '', survey_date: '', status: 'Pending', voltage_level: '', connected_feeders: '', transformer_capacity: '', tap_changer: '', oil_condition: '', feeder_demand: '', pending_capacity: '', coordinates: '', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);

  const filtered = data.filter(i => !searchTerm || ['substation', 'surveyor', 'district'].some(k => String(i[k] || '').toLowerCase().includes(searchTerm.toLowerCase())));
  const totalPages = Math.ceil(filtered.length / itemsPerPage) || 1;
  const start = (page - 1) * itemsPerPage;
  const display = filtered.slice(start, start + itemsPerPage);

  const openNew = () => { setFormData(emptyForm); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm('Delete?')) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (!formData.substation || !formData.surveyor) { toast.error('Required'); return; }
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false); setEditingId(null);
  };

  const columns = [
    { key: 'substation', label: 'Substation', strong: true }, { key: 'subdivision', label: 'Subdivision' },
    { key: 'surveyor', label: 'Surveyor' }, { key: 'district', label: 'District' },
    { key: 'transformer_capacity', label: 'Transformer' }, { key: 'pending_capacity', label: 'Spare' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];

  if (showForm) {
    const onChange = e => setFormData({ ...formData, [e.target.name]: e.target.value });
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: 'clamp(12px, 2.5vw, 20px)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
            <h1 style={S.title}>{editingId ? 'Edit Survey' : 'New Survey'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', padding: 'clamp(12px, 2.5vw, 24px)', border: '1px solid #f3f4f6' }}>
            <form onSubmit={submit}>
              <div style={S.formGrid3}>
                <Field label="Substation" icon={<Zap size={11} />} name="substation" value={formData.substation} onChange={onChange} required />
                <Field label="Subdivision" icon={<MapPin size={11} />} name="subdivision" value={formData.subdivision} onChange={onChange} />
                <Field label="Surveyor" icon={<User size={11} />} name="surveyor" value={formData.surveyor} onChange={onChange} required />
                <Field label="District" icon={<MapPin size={11} />} name="district" value={formData.district} onChange={onChange} options={['JALGAON', 'AHILYANAGAR', 'SOLAPUR', 'NASHIK']} />
                <Field label="Taluka" icon={<MapPin size={11} />} name="taluka" value={formData.taluka} onChange={onChange} />
                <Field label="Survey Date" icon={<Calendar size={11} />} name="survey_date" type="date" value={formData.survey_date} onChange={onChange} />
                <Field label="Voltage Level" icon={<Zap size={11} />} name="voltage_level" value={formData.voltage_level} onChange={onChange} />
                <Field label="Connected Feeders" icon={<List size={11} />} name="connected_feeders" value={formData.connected_feeders} onChange={onChange} />
                <Field label="Transformer" icon={<Zap size={11} />} name="transformer_capacity" value={formData.transformer_capacity} onChange={onChange} />
                <Field label="Tap Changer" icon={<Zap size={11} />} name="tap_changer" value={formData.tap_changer} onChange={onChange} />
                <Field label="Oil Condition" icon={<Shield size={11} />} name="oil_condition" value={formData.oil_condition} onChange={onChange} />
                <Field label="Feeder Demand" icon={<Gauge size={11} />} name="feeder_demand" value={formData.feeder_demand} onChange={onChange} />
                <Field label="Pending Capacity" icon={<TrendingUp size={11} />} name="pending_capacity" value={formData.pending_capacity} onChange={onChange} />
                <Field label="Coordinates" icon={<Compass size={11} />} name="coordinates" value={formData.coordinates} onChange={onChange} />
                <Field label="Status" icon={<Shield size={11} />} name="status" value={formData.status} onChange={onChange} options={['Pending', 'In Progress', 'Completed']} />
                <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={11} /> Remarks</label>
                  <textarea name="remarks" value={formData.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
                </div>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '20px', paddingTop: '16px', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={13} /> {editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Substation Survey Checklist</h1><p style={S.subtitle}>MSKVY Feeder Solarisation</p></div>

      <TopToolbar itemsPerPage={itemsPerPage} onItemsPerPageChange={(v) => { setItemsPerPage(v); setPage(1); }}
        onExcel={() => exportToExcel(filtered, columns, 'substation_surveys')}
        onPrint={() => printReport(filtered, columns, 'Substation Surveys')} />

      <div style={S.searchBar}>
        <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 0 }}>
          <Search size={14} style={S.searchIcon} />
          <input type="text" placeholder="Search substation, surveyor..." value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => setSearchTerm('')} style={S.clearIcon}><X size={12} /></button>}
        </div>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={13} /> Add</button>
      </div>

      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Surveys</h2><span style={S.countBadge}>{filtered.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr></thead>
            <tbody>
              {display.map((item, idx) => (
                <tr key={item.id}>
                  <td style={S.td}>{start + idx + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(item) : (item[c.key] ?? '—')}</td>)}
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '2px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={14} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={14} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <BottomPagination currentPage={page} totalPages={totalPages} totalItems={filtered.length} itemsPerPage={itemsPerPage} onPageChange={setPage} />
      </div>

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}><h2 style={S.modalTitle}>Substation Survey Details</h2><button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button></div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: '12px' }}>
                {[['Substation', selected.substation], ['Subdivision', selected.subdivision], ['Surveyor', selected.surveyor], ['District', selected.district], ['Voltage', selected.voltage_level], ['Feeders', selected.connected_feeders], ['Transformer', selected.transformer_capacity], ['Spare', selected.pending_capacity], ['Status', selected.status]].map(([l, v]) => (
                  <div key={l}>
                    <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>{l}</p>
                    <p style={{ fontSize: '13px', margin: 0, wordBreak: 'break-word' }}>{v || '—'}</p>
                  </div>
                ))}
              </div>
              {selected.coordinates && (
                <div style={{ marginTop: '12px' }}>
                  <a href={getGoogleEarthUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#0d9488', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <Globe size={14} /> Open {selected.coordinates} in Google Earth
                  </a>
                </div>
              )}
              {selected.remarks && <div style={{ marginTop: '12px', padding: '10px', background: '#fef3c7', borderRadius: '6px' }}><p style={{ fontSize: '10px', fontWeight: '700', color: '#92400e', margin: '0 0 3px 0' }}>REMARKS</p><p style={{ fontSize: '12px', color: '#78350f', margin: 0 }}>{selected.remarks}</p></div>}
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={13} /> Edit</button>
              <button onClick={() => del(selected)} style={{ ...S.btnSecondary, color: '#dc2626', borderColor: '#fecaca' }}><Trash2 size={13} /> Delete</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// LAND SURVEY PAGE
// =============================================
const MSKVYLandSurveyPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_land_v5', MSKVY_LAND_SURVEYS));
  const [searchTerm, setSearchTerm] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_land_v5', data); }, [data]);

  const emptyForm = { village: '', block: '', district: 'JALGAON', feeder: '', surveyor: '', survey_date: '', status: 'Pending', land_nearby: '', distance_km: '', ownership: '', title_status: '', topography: '', approach_road: '', water_availability: '', environmental_clearance: '', geotagged_photos: '', coordinates: '', survey_map: '', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);

  const filtered = data.filter(i => !searchTerm || ['village', 'block', 'feeder', 'surveyor'].some(k => String(i[k] || '').toLowerCase().includes(searchTerm.toLowerCase())));
  const totalPages = Math.ceil(filtered.length / itemsPerPage) || 1;
  const start = (page - 1) * itemsPerPage;
  const display = filtered.slice(start, start + itemsPerPage);

  const openNew = () => { setFormData(emptyForm); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm('Delete?')) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (!formData.village || !formData.surveyor) { toast.error('Required'); return; }
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false); setEditingId(null);
  };

  const columns = [
    { key: 'village', label: 'Village', strong: true }, { key: 'block', label: 'Block' }, { key: 'district', label: 'District' },
    { key: 'feeder', label: 'Feeder' }, { key: 'surveyor', label: 'Surveyor' }, { key: 'ownership', label: 'Ownership' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];

  if (showForm) {
    const onChange = e => setFormData({ ...formData, [e.target.name]: e.target.value });
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: 'clamp(12px, 2.5vw, 20px)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
            <h1 style={S.title}>{editingId ? 'Edit Land Survey' : 'New Land Survey'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', padding: 'clamp(12px, 2.5vw, 24px)', border: '1px solid #f3f4f6' }}>
            <form onSubmit={submit}>
              <div style={S.formGrid3}>
                <Field label="Village" icon={<MapPin size={11} />} name="village" value={formData.village} onChange={onChange} required />
                <Field label="Block" icon={<MapPin size={11} />} name="block" value={formData.block} onChange={onChange} />
                <Field label="District" icon={<MapPin size={11} />} name="district" value={formData.district} onChange={onChange} options={['JALGAON', 'AHILYANAGAR', 'SOLAPUR', 'NASHIK']} />
                <Field label="Feeder" icon={<Zap size={11} />} name="feeder" value={formData.feeder} onChange={onChange} />
                <Field label="Surveyor" icon={<User size={11} />} name="surveyor" value={formData.surveyor} onChange={onChange} required />
                <Field label="Survey Date" icon={<Calendar size={11} />} name="survey_date" type="date" value={formData.survey_date} onChange={onChange} />
                <Field label="Ownership" icon={<Shield size={11} />} name="ownership" value={formData.ownership} onChange={onChange} options={['Private', 'Govt', 'Panchayat']} />
                <Field label="Topography" icon={<Layers size={11} />} name="topography" value={formData.topography} onChange={onChange} />
                <Field label="Coordinates" icon={<Compass size={11} />} name="coordinates" value={formData.coordinates} onChange={onChange} />
                <Field label="Status" icon={<Shield size={11} />} name="status" value={formData.status} onChange={onChange} options={['Pending', 'In Progress', 'Completed']} />
                <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={11} /> Remarks</label>
                  <textarea name="remarks" value={formData.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
                </div>
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '20px', paddingTop: '16px', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={13} /> {editingId ? 'Update' : 'Submit'}</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Land Survey Checklist</h1><p style={S.subtitle}>MSKVY Feeder Solarisation</p></div>

      <TopToolbar itemsPerPage={itemsPerPage} onItemsPerPageChange={(v) => { setItemsPerPage(v); setPage(1); }}
        onExcel={() => exportToExcel(filtered, columns, 'land_surveys')}
        onPrint={() => printReport(filtered, columns, 'Land Surveys')} />

      <div style={S.searchBar}>
        <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 0 }}>
          <Search size={14} style={S.searchIcon} />
          <input type="text" placeholder="Search village, feeder..." value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => setSearchTerm('')} style={S.clearIcon}><X size={12} /></button>}
        </div>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={13} /> Add</button>
      </div>

      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Land Surveys</h2><span style={S.countBadge}>{filtered.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr></thead>
            <tbody>
              {display.map((item, idx) => (
                <tr key={item.id}>
                  <td style={S.td}>{start + idx + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(item) : (item[c.key] ?? '—')}</td>)}
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '2px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={14} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={14} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <BottomPagination currentPage={page} totalPages={totalPages} totalItems={filtered.length} itemsPerPage={itemsPerPage} onPageChange={setPage} />
      </div>

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}><h2 style={S.modalTitle}>Land Survey Details</h2><button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button></div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 180px), 1fr))', gap: '12px' }}>
                {[['Village', selected.village], ['Block', selected.block], ['District', selected.district], ['Feeder', selected.feeder], ['Surveyor', selected.surveyor], ['Ownership', selected.ownership], ['Topography', selected.topography], ['Status', selected.status]].map(([l, v]) => (
                  <div key={l}>
                    <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 3px 0', fontWeight: '600', textTransform: 'uppercase' }}>{l}</p>
                    <p style={{ fontSize: '13px', margin: 0, wordBreak: 'break-word' }}>{v || '—'}</p>
                  </div>
                ))}
              </div>
              {selected.coordinates && (
                <div style={{ marginTop: '12px' }}>
                  <a href={getGoogleEarthUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#0d9488', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <Globe size={14} /> Open {selected.coordinates} in Google Earth
                  </a>
                </div>
              )}
              {selected.remarks && <div style={{ marginTop: '12px', padding: '10px', background: '#dcfce7', borderRadius: '6px' }}><p style={{ fontSize: '10px', fontWeight: '700', color: '#166534', margin: '0 0 3px 0' }}>REMARKS</p><p style={{ fontSize: '12px', color: '#14532d', margin: 0 }}>{selected.remarks}</p></div>}
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={13} /> Edit</button>
              <button onClick={() => del(selected)} style={{ ...S.btnSecondary, color: '#dc2626', borderColor: '#fecaca' }}><Trash2 size={13} /> Delete</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// MENU CONFIG
// =============================================
const MENU_CONFIG = {
  dashboard: { main: [{ id: 'dashboard-main', label: 'Management Dashboard' }, { id: 'activity-log', label: 'Activity Log' }], masters: [] },
  mskvy: {
    main: [
      { id: 'mskvy-contractors', label: 'MSKVY Contractor List' },
      { id: 'mskvy-substation-survey', label: 'Substation Surveys' },
      { id: 'mskvy-land-survey', label: 'Land Surveys' },
    ],
    masters: [
      { id: 'mskvy-district', label: 'District Master' },
      { id: 'mskvy-survey-status', label: 'Survey Status' },
    ],
  },
  projects: { main: [{ id: 'proj-list', label: 'All Projects' }], masters: [] },
  contractors: { main: [{ id: 'contractors-all', label: 'All Contractors' }], masters: [] },
  'field-workers': { main: [{ id: 'fw-all', label: 'All Field Workers' }], masters: [] },
  reports: { main: [{ id: 'reports-all', label: 'All Reports' }], masters: [] },
  media: { main: [{ id: 'photos', label: 'Photo Gallery' }, { id: 'documents', label: 'Documents' }], masters: [] },
  approvals: { main: [{ id: 'approvals-all', label: 'Pending Approvals' }], masters: [] },
  applications: { main: [{ id: 'app-contractor', label: 'Contractor App' }, { id: 'app-worker', label: 'Field Worker App' }], masters: [] },
  masters: { main: [], masters: [{ id: 'mm-status', label: 'General Status' }] },
  users: { main: [{ id: 'users-all', label: 'All Users' }], masters: [] },
};

// =============================================
// MAIN APP
// =============================================
const App = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [user, setUser] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [activeMenu, setActiveMenu] = useState('dashboard');
  const [activeSubMenu, setActiveSubMenu] = useState('dashboard-main');
  const [isMobile, setIsMobile] = useState(false);

  useEffect(() => {
    const checkMobile = () => setIsMobile(window.innerWidth < 768);
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  useEffect(() => {
    const token = localStorage.getItem('authToken');
    const savedUser = localStorage.getItem('user');
    if (token && savedUser) { setIsAuthenticated(true); setUser(JSON.parse(savedUser)); }
  }, []);

  const handleLogin = (u) => { setUser(u); setIsAuthenticated(true); };
  const handleLogout = () => {
    localStorage.removeItem('authToken'); localStorage.removeItem('user');
    setIsAuthenticated(false); setUser(null);
  };

  if (!isAuthenticated) return <Login onLogin={handleLogin} />;

  const topMenus = [
    { id: 'dashboard', label: 'Dashboard' },
    { id: 'mskvy', label: 'MSKVY Solarisation' },
    { id: 'projects', label: 'Projects' },
    { id: 'contractors', label: 'Contractors' },
    { id: 'field-workers', label: 'Field Workers' },
    { id: 'reports', label: 'Reports' },
    { id: 'media', label: 'Media' },
    { id: 'approvals', label: 'Approvals' },
    { id: 'applications', label: 'Applications' },
    { id: 'masters', label: 'Masters' },
    { id: 'users', label: 'Users' },
  ];

  const handleCardNavigate = (menuId, subMenuId) => {
    setActiveMenu(menuId);
    setActiveSubMenu(subMenuId);
    if (isMobile) setMobileMenuOpen(false);
  };

  const renderContent = () => {
    switch (activeSubMenu) {
      case 'dashboard-main': return <DashboardPage onNavigate={handleCardNavigate} />;
      case 'mskvy-contractors': return <MSKVYContractorPage />;
      case 'mskvy-substation-survey': return <MSKVYSubstationSurveyPage />;
      case 'mskvy-land-survey': return <MSKVYLandSurveyPage />;
      case 'mskvy-district': return <DistrictMasterPage />;
      case 'mskvy-survey-status': return <SurveyStatusMasterPage />;
      default: {
        const cm = MENU_CONFIG[activeMenu];
        const subItem = cm && (cm.main.find(s => s.id === activeSubMenu) || cm.masters.find(m => m.id === activeSubMenu));
        return <ComingSoonPage title={subItem?.label || 'Module'} />;
      }
    }
  };

  const cm = MENU_CONFIG[activeMenu] || { main: [], masters: [] };

  const handleMenuClick = (menuId) => {
    setActiveMenu(menuId);
    const cfg = MENU_CONFIG[menuId];
    const first = cfg.main[0] || cfg.masters[0];
    if (first) setActiveSubMenu(first.id);
    if (isMobile) setMobileMenuOpen(false);
  };

  const handleSubMenuClick = (subId) => {
    setActiveSubMenu(subId);
    if (isMobile) setMobileMenuOpen(false);
  };

  return (
    <Router>
      <div style={{ display: 'flex', height: '100vh', background: '#f3f4f6', fontFamily: 'system-ui, sans-serif', overflow: 'hidden' }}>
        {isMobile && mobileMenuOpen && (
          <div onClick={() => setMobileMenuOpen(false)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 40 }} />
        )}
        <div style={{
          background: '#111827', color: '#fff', display: 'flex', flexDirection: 'column',
          width: isMobile ? (mobileMenuOpen ? '260px' : '0') : (sidebarOpen ? '240px' : '60px'),
          transition: 'width 0.3s', flexShrink: 0,
          position: isMobile ? 'fixed' : 'relative', top: 0, bottom: 0, left: 0,
          zIndex: isMobile ? 50 : 'auto', overflow: 'hidden', height: '100vh'
        }}>
          {!isMobile && (
            <div style={{ padding: '14px', borderBottom: '1px solid #1f2937', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              {sidebarOpen && <span style={{ fontWeight: 'bold', fontSize: '14px', letterSpacing: '1px' }}>solar ERP</span>}
              <button onClick={() => setSidebarOpen(!sidebarOpen)} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex' }}>
                {sidebarOpen ? <ChevronLeft size={16} /> : <ChevronRight size={16} />}
              </button>
            </div>
          )}
          {isMobile && (
            <div style={{ padding: '14px', borderBottom: '1px solid #1f2937', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <span style={{ fontWeight: 'bold', fontSize: '14px', letterSpacing: '1px' }}>solar ERP</span>
              <button onClick={() => setMobileMenuOpen(false)} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex' }}><X size={18} /></button>
            </div>
          )}
          <div style={{ flex: 1, overflowY: 'auto', padding: '10px 6px' }}>
            {cm.main.length > 0 && (
              <>
                {(sidebarOpen || isMobile) && <p style={{ fontSize: '9px', color: '#6b7280', margin: '6px 10px 5px', fontWeight: '700', letterSpacing: '1px', textTransform: 'uppercase' }}>Main</p>}
                {cm.main.map(item => (
                  <button key={item.id} onClick={() => handleSubMenuClick(item.id)} style={{
                    width: '100%', display: 'flex', alignItems: 'center', gap: '8px',
                    padding: '9px 10px', borderRadius: '6px', border: 'none',
                    background: activeSubMenu === item.id ? '#2563eb' : 'transparent',
                    color: activeSubMenu === item.id ? '#fff' : '#d1d5db',
                    cursor: 'pointer', fontSize: '12px', textAlign: 'left', marginBottom: '2px', fontFamily: 'inherit'
                  }}>
                    <span style={{ width: '5px', height: '5px', borderRadius: '50%', background: activeSubMenu === item.id ? '#fff' : '#60a5fa', flexShrink: 0 }} />
                    {(sidebarOpen || isMobile) && <span>{item.label}</span>}
                  </button>
                ))}
              </>
            )}
            {cm.masters.length > 0 && (
              <>
                {(sidebarOpen || isMobile) && <p style={{ fontSize: '9px', color: '#fbbf24', margin: '12px 10px 5px', fontWeight: '700', letterSpacing: '1px', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '5px' }}><Database size={9} /> Masters</p>}
                {cm.masters.map(item => (
                  <button key={item.id} onClick={() => handleSubMenuClick(item.id)} style={{
                    width: '100%', display: 'flex', alignItems: 'center', gap: '8px',
                    padding: '8px 10px', borderRadius: '6px', border: 'none',
                    background: activeSubMenu === item.id ? '#7c3aed' : 'transparent',
                    color: activeSubMenu === item.id ? '#fff' : '#9ca3af',
                    cursor: 'pointer', fontSize: '11.5px', textAlign: 'left', marginBottom: '2px', fontFamily: 'inherit'
                  }}>
                    <List size={11} style={{ flexShrink: 0 }} />
                    {(sidebarOpen || isMobile) && <span>{item.label}</span>}
                  </button>
                ))}
              </>
            )}
          </div>
          <div style={{ padding: '10px', borderTop: '1px solid #1f2937' }}>
            {(sidebarOpen || isMobile) ? (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '6px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', minWidth: 0 }}>
                  <div style={{ width: '26px', height: '26px', background: '#374151', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}><User size={12} /></div>
                  <div style={{ minWidth: 0 }}>
                    <p style={{ fontSize: '11px', fontWeight: '600', margin: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{user?.name || 'Admin'}</p>
                    <p style={{ fontSize: '9px', color: '#9ca3af', margin: 0 }}>{user?.role || 'Super Admin'}</p>
                  </div>
                </div>
                <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex', flexShrink: 0 }}><LogOut size={13} /></button>
              </div>
            ) : (
              <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', width: '100%', display: 'flex', justifyContent: 'center' }}><LogOut size={15} /></button>
            )}
          </div>
        </div>

        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minWidth: 0 }}>
          <div style={{ background: '#fff', borderBottom: '1px solid #e5e7eb', flexShrink: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: 'clamp(8px, 1.8vw, 10px) clamp(10px, 2.5vw, 20px)', gap: '8px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', minWidth: 0, flex: 1 }}>
                {isMobile && (
                  <button onClick={() => setMobileMenuOpen(true)} style={{ background: 'transparent', border: '1px solid #d1d5db', borderRadius: '5px', padding: '5px', cursor: 'pointer', display: 'flex', flexShrink: 0 }}>
                    <List size={14} />
                  </button>
                )}
                <h2 style={{ fontSize: 'clamp(10px, 2vw, 13px)', fontWeight: '600', color: '#1f2937', margin: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {topMenus.find(m => m.id === activeMenu)?.label} / {cm.main.find(s => s.id === activeSubMenu)?.label || cm.masters.find(m => m.id === activeSubMenu)?.label || ''}
                </h2>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 'clamp(4px, 1.5vw, 12px)', flexShrink: 0 }}>
                {!isMobile && <span style={{ fontSize: '11px', color: '#4b5563' }}>Welcome, {user?.name || 'Admin'}</span>}
                <button style={{ background: 'transparent', border: 'none', color: '#4b5563', cursor: 'pointer', display: 'flex' }}><Bell size={15} /></button>
                <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#dc2626', fontSize: 'clamp(9px, 1.8vw, 11px)', fontWeight: '600', cursor: 'pointer', fontFamily: 'inherit' }}>Logout</button>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '2px', padding: '0 clamp(8px, 2vw, 20px)', borderTop: '1px solid #f3f4f6', overflowX: 'auto', WebkitOverflowScrolling: 'touch' }}>
              {topMenus.map(menu => (
                <button key={menu.id} onClick={() => handleMenuClick(menu.id)} style={{
                  padding: 'clamp(7px, 1.5vw, 9px) clamp(7px, 1.5vw, 12px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '500', background: 'transparent', border: 'none',
                  borderBottom: activeMenu === menu.id ? '2px solid #2563eb' : '2px solid transparent',
                  color: activeMenu === menu.id ? '#2563eb' : '#4b5563',
                  cursor: 'pointer', whiteSpace: 'nowrap', fontFamily: 'inherit'
                }}>{menu.label}</button>
              ))}
            </div>
          </div>
          <div style={{ flex: 1, overflowY: 'auto' }}>{renderContent()}</div>
        </div>
      </div>
    </Router>
  );
};

export default App;