import React, { useState, useEffect } from 'react';
import { HashRouter as Router } from 'react-router-dom';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import {
  Search, Plus, Edit, Trash2, Eye, X, ChevronLeft, ChevronRight,
  CheckCircle, Printer, FileSpreadsheet,
  Hash, User, Globe, FileText, Shield, Calendar, Tag, Gauge, MapPin,
  Users, Building2, FolderOpen, ClipboardList, Bell, LogOut, Lock, Mail,
  Phone, TrendingUp, Database, List, Zap, Ruler,
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
// MSKVY CONTRACTORS — ALL 127 records (Trimmed for space, keep original)
// =============================================
const MSKVY_CONTRACTORS = [
  { id: 1, s_no: 1, contractor: 'KRIVI Energy Pvt Ltd', director: 'Mr. Bhupendra Ji', director_contact: '9326275972', file_name: 'KIRVI 01', district: 'AHILYANAGAR', taluka: 'Parner', supervisor: 'Yogesh Chitalkar', contact: '8380868899', substation: '33/11 KV NIGHOJ', area: '34.05', solar: '', coordinates: '', distance_km: '', obstruction_clearance_m: '', substation_number: '', sr_no: '', land_status: '', location: '', remarks: '' },
  // ... (baaki 127 records same rahenge, maine yahan short rakha hai)
];

// =============================================
// MSKVY SUBSTATION SURVEYS — KAJGAON full 28 items
// =============================================
const MSKVY_SUBSTATION_SURVEYS = [
  {
    id: 1, s_no: 1,
    substation: 'KAJGAON 33/11 KV Substation',
    subdivision: 'Bhadgaon',
    surveyor: 'Hemant Agrawal',
    district: 'JALGAON',
    taluka: 'Bhadgaon',
    contractor: 'WAN Technology (OPC) Pvt Ltd',
    survey_date: '2026-10-22',
    status: 'Completed',
    voltage_level: '33/11 kV',
    discom_division: 'Division - Bhadgaon',
    connected_feeders: '5 Agri, 6 Domestic',
    feeder_length: '10km,12km,7km',
    agri_pumps_load: '',
    peak_load: '',
    supply_hours: '8 Hours',
    outage_history: '',
    transformer_capacity: '5 MVA X 4 Nos',
    transformer_make: '',
    transformer_year: '',
    loading_percent: '',
    spare_capacity: '',
    tap_changer: '33/11',
    oil_condition: 'Good',
    bushing_condition: 'Good',
    earthing_condition: 'Good',
    spare_bay: 'New Bat Needed',
    bay_space: 'Available',
    switchgear: 'Good',
    protection_relays: 'Yes',
    scada_rtu: 'Yes',
    metering: 'Available',
    evacuation_distance: '1.5 Km',
    right_of_way: 'Not Found Any',
    received_voltage: '11 KV',
    fault_level: '',
    lightning_protection: 'Yes',
    fire_safety: 'Yes',
    feeder_demand: '15 MW',
    existing_capacity: '5 MW',
    pending_capacity: '0.464MW',
    transformer_capacity_check: '4 X 5 MVA',
    sld_available: 'Available',
    environmental_clearance: 'Yes',
    feeder_load_data: '89Amp,41Amp,73Amp,',
    geotagged_photos: 'Attached',
    coordinates: '20.571375 75.150647',
    surveyor_sign: 'Signed',
    discom_sign: 'Signed',
    remarks: '1 transformer add come to 8 to 10 days'
  },
];

// =============================================
// MSKVY LAND SURVEYS — KAJGAON full 9 items
// =============================================
const MSKVY_LAND_SURVEYS = [
  {
    id: 1, s_no: 1,
    village: 'KAJGAON',
    block: 'Bhadgaon',
    district: 'JALGAON',
    feeder: 'KAJGAON 33/11 KV',
    surveyor: 'Hemant Agrawal',
    survey_date: '2026-10-22',
    contractor: 'WAN Technology (OPC) Pvt Ltd',
    status: 'Completed',
    land_nearby: 'YES',
    distance_km: '1.7 KM',
    ownership: 'Private',
    area_acres: '',
    title_status: 'Farming Land',
    litigation: '',
    topography: 'No',
    waterlogging: 'NA',
    soil_type: '',
    approach_road: 'access for heavy vehicles',
    water_availability: 'Yes',
    environmental_clearance: 'NA',
    geotagged_photos: 'Yes',
    coordinates: '20.561303 75.159826',
    survey_map: 'Yes',
    remarks: 'Village KAJGAON ,Block Bhadgaon,Dist - Jalgaon, Maharshtra.'
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
  iconBtn: { padding: '3px', borderRadius: '4px', border: 'none', background: 'transparent', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' },
  pagination: { padding: 'clamp(8px, 1.8vw, 12px) clamp(10px, 2.5vw, 16px)', borderTop: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '8px' },
  pageBtn: { display: 'inline-flex', alignItems: 'center', gap: '3px', padding: 'clamp(4px, 1vw, 6px) clamp(6px, 1.5vw, 10px)', fontSize: 'clamp(10px, 1.9vw, 12px)', fontWeight: '500', color: '#374151', background: '#fff', border: '1px solid #d1d5db', borderRadius: '6px', cursor: 'pointer', fontFamily: 'inherit' },
  pageBtnActive: { background: '#2563eb', color: '#fff', borderColor: '#2563eb' },
  pageBtnDisabled: { opacity: 0.5, cursor: 'not-allowed' },
  statusBadge: { display: 'inline-block', padding: '3px 8px', borderRadius: '10px', fontSize: 'clamp(9px, 1.7vw, 10px)', fontWeight: '600', whiteSpace: 'nowrap' },
  modalOverlay: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 'clamp(6px, 1.5vw, 12px)', zIndex: 100 },
  modal: { background: '#fff', borderRadius: '12px', width: '100%', maxWidth: '900px', maxHeight: '92vh', overflowY: 'auto' },
  modalHeader: { padding: 'clamp(10px, 2vw, 16px) clamp(12px, 2.5vw, 20px)', borderBottom: '1px solid #e5e7eb', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '8px', position: 'sticky', top: 0, background: '#fff', zIndex: 2 },
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
  excelTable: { width: '100%', borderCollapse: 'collapse', border: '1px solid #000', fontSize: 'clamp(10px, 1.8vw, 12px)' },
  excelHeader: { background: '#1e3a8a', color: '#fff', padding: '8px', textAlign: 'center', fontWeight: 'bold', border: '1px solid #000' },
  excelTh: { background: '#e5e7eb', color: '#000', padding: '6px 8px', border: '1px solid #000', textAlign: 'left', fontWeight: 'bold' },
  excelTd: { padding: '6px 8px', border: '1px solid #000', color: '#000', verticalAlign: 'top' },
  excelSection: { background: '#e5e7eb', color: '#000', padding: '6px 8px', fontWeight: 'bold', border: '1px solid #000', borderLeft: '4px solid #1e3a8a' },
  excelInput: { width: '100%', padding: '6px', border: '1px solid #ccc', background: '#fff9e6', fontSize: '11px', fontFamily: 'inherit', boxSizing: 'border-box' }
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

const loadMSKVY = (key, fallback) => {
  try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : fallback; } catch { return fallback; }
};
const saveMSKVY = (key, data) => {
  try { localStorage.setItem(key, JSON.stringify(data)); } catch {}
};

// =============================================
// EXPORT TO EXCEL (for top toolbar)
// =============================================
const exportToExcel = (data, columns, filename) => {
  try {
    if (!data.length) { toast.error('No data to export'); return; }
    const excelData = [
      columns.map(c => c.label),
      ...data.map(item => columns.map(c => c.exportValue ? c.exportValue(item) : (item[c.key] ?? '')))
    ];
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    ws['!cols'] = columns.map(() => ({ wch: 20 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${filename}_${dayjs().format('YYYY-MM-DD')}.xlsx`);
    toast.success('Excel exported');
  } catch (e) { console.error(e); toast.error('Export failed'); }
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
// PRINT LIST REPORT (for top toolbar — multiple records table)
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
              <div style="font-size:15px;font-weight:700;">MSKVY FEEDER SOLARISATION</div>
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

// =============================================
// PRINT SINGLE CHECKLIST (Excel-format, yellow style) — for Actions column
// =============================================
const printSingleChecklist = (item, type) => {
  try {
    const existing = document.getElementById('print-checklist-overlay');
    if (existing) existing.remove();

    let title = '';
    let tableHTML = '';
    const cellStyle = 'border:1px solid #000;padding:6px 8px;vertical-align:top;';
    const headStyle = 'border:1px solid #000;padding:6px 8px;background:#e5e7eb;text-align:left;font-weight:bold;';
    const labelStyle = 'border:1px solid #000;padding:6px 8px;font-weight:bold;background:#f3f4f6;';
    const sectionStyle = 'background:#e5e7eb;border:1px solid #000;padding:6px 8px;font-weight:bold;border-left:4px solid #1e3a8a;';

    if (type === 'substation') {
      title = 'Substation Survey Checklist';
      tableHTML = `
        <tr><th colspan="5" style="background:#1e3a8a;color:#fff;padding:8px;text-align:center;font-weight:bold;border:1px solid #000;font-size:14px;">MSKVY Feeder Solarisation: Substation Survey Checklist</th></tr>
        <tr><td colspan="2" style="${labelStyle}">Substation:</td><td colspan="3" style="${cellStyle}">${item.substation || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">Subdivision Name :</td><td colspan="3" style="${cellStyle}">${item.subdivision || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">Surveyor / Date:</td><td colspan="3" style="${cellStyle}">${item.surveyor || ''} - Date - ${item.survey_date ? dayjs(item.survey_date).format('DD/MM/YYYY') : ''}</td></tr>
        <tr>
          <th style="${headStyle}width:40px;">Sr.</th>
          <th style="${headStyle}">Checklist Item</th>
          <th style="${headStyle}width:80px;">Status</th>
          <th style="${headStyle}">Details / Value</th>
          <th style="${headStyle}">Remarks</th>
        </tr>

        <tr><td colspan="5" style="${sectionStyle}">1. Substation & Feeder Basics</td></tr>
        <tr><td style="${cellStyle}">1.1</td><td style="${cellStyle}">Substation name, voltage level (33/11 kV), DISCOM / circle / division</td><td style="${cellStyle}">${item.status || ''}</td><td style="${cellStyle}">${item.substation}, ${item.voltage_level}, ${item.discom_division}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.2</td><td style="${cellStyle}">Connected Feeder Details</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.connected_feeders || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.3</td><td style="${cellStyle}">Feeder length (km), no. of DTs</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.feeder_length || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.4</td><td style="${cellStyle}">Connected agri pumps and total load (HP / kW)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.agri_pumps_load || '—'}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.5</td><td style="${cellStyle}">Peak and average feeder load (last 12 months)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.peak_load || '—'}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.6</td><td style="${cellStyle}">Supply hours (day / night) and outage history</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.supply_hours || ''} ${item.outage_history ? ', ' + item.outage_history : ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">2. Transformer Details</td></tr>
        <tr><td style="${cellStyle}">2.1</td><td style="${cellStyle}">Power transformer capacity (MVA), make, year of installation</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.transformer_capacity || ''}</td><td style="${cellStyle}">${item.remarks || ''}</td></tr>
        <tr><td style="${cellStyle}">2.2</td><td style="${cellStyle}">Loading % and spare capacity available</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.loading_percent || '—'} / ${item.spare_capacity || '—'}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">2.3</td><td style="${cellStyle}">Tap changer type / range</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.tap_changer || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">2.4</td><td style="${cellStyle}">Condition of oil, bushings and earthing</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.oil_condition || ''}, ${item.bushing_condition || ''}, ${item.earthing_condition || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">3. Interconnection Readiness</td></tr>
        <tr><td style="${cellStyle}">3.1</td><td style="${cellStyle}">Spare 11 kV / 33 kV bay available (or new bay needed)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.spare_bay || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.2</td><td style="${cellStyle}">Space for new bay, metering and protection panels</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.bay_space || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.3</td><td style="${cellStyle}">Switchgear condition (VCB, isolators, CTs, PTs)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.switchgear || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.4</td><td style="${cellStyle}">Protection relays, SCADA / RTU availability</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.protection_relays || ''} / ${item.scada_rtu || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.5</td><td style="${cellStyle}">Metering (ABT / energy meter) provision</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.metering || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.6</td><td style="${cellStyle}">Distance and route of evacuation line to solar plant</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.evacuation_distance || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">3.7</td><td style="${cellStyle}">Right-of-way issues</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.right_of_way || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">4. Grid Quality & Safety</td></tr>
        <tr><td style="${cellStyle}">4.1</td><td style="${cellStyle}">Received Voltage By Solar Plant</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.received_voltage || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">4.2</td><td style="${cellStyle}">Fault level at the bus</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.fault_level || '—'}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">4.3</td><td style="${cellStyle}">Is Earthing system and lightning protection Is in good condition</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.lightning_protection || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">4.4</td><td style="${cellStyle}">Fire safety and clearances</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.fire_safety || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">5. Solar Plant Sizing Inputs</td></tr>
        <tr><td style="${cellStyle}">5.1</td><td style="${cellStyle}">Feeder demand to be met (MW)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.feeder_demand || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">5.2</td><td style="${cellStyle}">Existing Plant Connected Capacity ( MW)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.existing_capacity || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">5.3</td><td style="${cellStyle}">Pending Capacity For Injecting Solar Plant (MW)</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.pending_capacity || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">5.4</td><td style="${cellStyle}">Plant capacity vs substation transformer capacity check</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.transformer_capacity_check || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">6. Regulatory & Documentation</td></tr>
        <tr><td style="${cellStyle}">6.1</td><td style="${cellStyle}">Single line diagram (SLD) of substation</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.sld_available || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">6.2</td><td style="${cellStyle}">Environmental, forest and revenue clearances For Solar Plant To Substation( Transmission line )</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.environmental_clearance || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">6.3</td><td style="${cellStyle}">Feeder-wise load data</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.feeder_load_data || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">7. Site Photos & Records</td></tr>
        <tr><td style="${cellStyle}">7.1</td><td style="${cellStyle}">Geo-tagged photos: land, substation, bays, approach road</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.geotagged_photos || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">7.2</td><td style="${cellStyle}">Coordinates (lat / long) and survey map</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.coordinates || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td colspan="2" style="${labelStyle}">Surveyor signature:</td><td colspan="3" style="${cellStyle}">${item.surveyor_sign || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">DISCOM representative signature:</td><td colspan="3" style="${cellStyle}">${item.discom_sign || ''}</td></tr>
      `;
    } else if (type === 'land') {
      title = 'Land Survey Checklist';
      tableHTML = `
        <tr><th colspan="5" style="background:#1e3a8a;color:#fff;padding:8px;text-align:center;font-weight:bold;border:1px solid #000;font-size:14px;">MSKVY Feeder Solarisation: Land Survey Checklist</th></tr>
        <tr><td colspan="2" style="${labelStyle}">Village Name :</td><td colspan="3" style="${cellStyle}">${item.village || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">Block / District :</td><td colspan="3" style="${cellStyle}">${item.block || ''} / ${item.district || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">Nearest Feeder Name :</td><td colspan="3" style="${cellStyle}">${item.feeder || ''}</td></tr>
        <tr><td colspan="2" style="${labelStyle}">Surveyor / Date:</td><td colspan="3" style="${cellStyle}">${item.surveyor || ''} - Date - ${item.survey_date ? dayjs(item.survey_date).format('DD/MM/YYYY') : ''}</td></tr>
        <tr>
          <th style="${headStyle}width:40px;">Sr.</th>
          <th style="${headStyle}">Checklist Item</th>
          <th style="${headStyle}width:80px;">Status</th>
          <th style="${headStyle}">Details / Value</th>
          <th style="${headStyle}">Remarks</th>
        </tr>

        <tr><td colspan="5" style="${sectionStyle}">1. Land Suitability</td></tr>
        <tr><td style="${cellStyle}">1.1</td><td style="${cellStyle}">Land within / adjacent to substation (ideally within 5 km)</td><td style="${cellStyle}">${item.status || ''}</td><td style="${cellStyle}">${item.land_nearby || ''}</td><td style="${cellStyle}">${item.distance_km || ''}</td></tr>
        <tr><td style="${cellStyle}">1.2</td><td style="${cellStyle}">Ownership (govt / private / panchayat) and area in acres</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.ownership || ''} ${item.area_acres ? '- ' + item.area_acres : ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.3</td><td style="${cellStyle}">Land title, encumbrances and litigation status</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.title_status || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.4</td><td style="${cellStyle}">Topography, shading, flooding / waterlogging risk</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.topography || ''} / ${item.waterlogging || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.5</td><td style="${cellStyle}">Soil type, approach road, access for heavy vehicles</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.approach_road || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.6</td><td style="${cellStyle}">Water availability for panel cleaning</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.water_availability || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">1.7</td><td style="${cellStyle}">Environmental, forest and revenue clearances For Solar Plant To Substation( Transmission line )</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.environmental_clearance || ''}</td><td style="${cellStyle}"></td></tr>

        <tr><td colspan="5" style="${sectionStyle}">2. Site Photos & Records</td></tr>
        <tr><td style="${cellStyle}">2.1</td><td style="${cellStyle}">Geo-tagged photos: land, approach road</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.geotagged_photos || ''}</td><td style="${cellStyle}"></td></tr>
        <tr><td style="${cellStyle}">2.2</td><td style="${cellStyle}">Coordinates (lat / long) and survey map</td><td style="${cellStyle}"></td><td style="${cellStyle}">${item.coordinates || ''}</td><td style="${cellStyle}">${item.survey_map || ''}</td></tr>
      `;
    }

    const overlay = document.createElement('div');
    overlay.id = 'print-checklist-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.7);z-index:99999;display:flex;flex-direction:column;padding:10px;box-sizing:border-box;overflow:hidden;';

    overlay.innerHTML = `
      <div style="background:#fff;border-radius:8px;flex:1;display:flex;flex-direction:column;overflow:hidden;max-width:1400px;margin:0 auto;width:100%;">
        <div style="padding:10px 16px;background:#1e3a8a;color:#fff;display:flex;justify-content:space-between;align-items:center;flex-shrink:0;gap:8px;flex-wrap:wrap;">
          <div><div style="font-size:14px;font-weight:700;">Print Preview — ${title}</div><div style="font-size:11px;opacity:0.8;margin-top:2px;">${item.substation || item.village || ''}</div></div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button id="print-checklist-now-btn" style="padding:8px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;font-weight:600;font-size:12px;cursor:pointer;">🖨️ Print Now</button>
            <button id="print-checklist-close-btn" style="padding:8px 16px;background:#dc2626;color:#fff;border:none;border-radius:6px;font-weight:600;font-size:12px;cursor:pointer;">✕ Close</button>
          </div>
        </div>
        <div style="flex:1;overflow:auto;padding:12px;background:#f3f4f6;">
          <div style="background:#fff;padding:12px;border-radius:6px;">
            <div style="text-align:center;border-bottom:2px solid #000;padding-bottom:8px;margin-bottom:10px;">
              <div style="font-size:15px;font-weight:700;">MSKVY FEEDER SOLARISATION</div>
              <div style="font-size:12px;font-weight:600;margin-top:3px;">${title.toUpperCase()}</div>
              <div style="font-size:10px;color:#555;margin-top:3px;">Generated: ${dayjs().format('DD/MM/YYYY HH:mm')}</div>
            </div>
            <table id="print-checklist-table" style="width:100%;border-collapse:collapse;font-size:11px;border:1px solid #000;">
              <tbody>${tableHTML}</tbody>
            </table>
            <div style="text-align:center;font-size:9px;color:#666;margin-top:8px;padding-top:6px;border-top:1px solid #ccc;">System generated report</div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const styleEl = document.createElement('style');
    styleEl.id = 'print-checklist-style';
    styleEl.innerHTML = `@media print { @page { size: A4 portrait; margin: 8mm; } body * { visibility: hidden !important; } #print-checklist-overlay, #print-checklist-overlay * { visibility: visible !important; } #print-checklist-overlay { position: static !important; background: white !important; padding: 0 !important; display: block !important; overflow: visible !important; } #print-checklist-overlay > div { max-width: 100% !important; } #print-checklist-overlay button { display: none !important; } #print-checklist-table { width: 100% !important; border: 1px solid #000 !important; } #print-checklist-table td, #print-checklist-table th { border: 1px solid #000 !important; padding: 4px 6px !important; font-size: 10px !important; word-break: break-word; white-space: normal !important; } }`;
    document.head.appendChild(styleEl);

    document.getElementById('print-checklist-now-btn').addEventListener('click', () => window.print());
    document.getElementById('print-checklist-close-btn').addEventListener('click', () => {
      overlay.remove();
      const s = document.getElementById('print-checklist-style');
      if (s) s.remove();
    });
    toast.success('Checklist preview opened');
  } catch (e) { console.error(e); toast.error('Print failed'); }
};

// =============================================
// EXCEL EXPORT single checklist record — for Actions column
// =============================================
const exportSingleExcel = (item, type) => {
  try {
    const wb = XLSX.utils.book_new();
    
    if (type === 'substation') {
      const excelData = [
        ['MSKVY Feeder Solarisation: Substation Survey Checklist', '', '', '', ''],
        ['Substation:', '', item.substation, '', ''],
        ['Subdivision Name :', '', item.subdivision, '', ''],
        ['Surveyor / Date:', '', `${item.surveyor} - ${item.survey_date}`, '', ''],
        ['Sr.', 'Checklist Item', 'Status', 'Details / Value', 'Remarks'],
        ['1. Substation & Feeder Basics', '', '', '', ''],
        ['1.1', 'Substation name, voltage level (33/11 kV), DISCOM / circle / division', item.status, `${item.substation}, ${item.voltage_level}, ${item.discom_division}`, ''],
        ['1.2', 'Connected Feeder Details', '', item.connected_feeders, ''],
        ['1.3', 'Feeder length (km), no. of DTs', '', item.feeder_length, ''],
        ['1.4', 'Connected agri pumps and total load (HP / kW)', '', item.agri_pumps_load, ''],
        ['1.5', 'Peak and average feeder load (last 12 months)', '', item.peak_load, ''],
        ['1.6', 'Supply hours (day / night) and outage history', '', `${item.supply_hours}, ${item.outage_history}`, ''],
        ['2. Transformer Details', '', '', '', ''],
        ['2.1', 'Power transformer capacity (MVA), make, year of installation', '', item.transformer_capacity, item.remarks],
        ['2.2', 'Loading % and spare capacity available', '', `${item.loading_percent} / ${item.spare_capacity}`, ''],
        ['2.3', 'Tap changer type / range', '', item.tap_changer, ''],
        ['2.4', 'Condition of oil, bushings and earthing', '', `${item.oil_condition}, ${item.bushing_condition}, ${item.earthing_condition}`, ''],
        ['3. Interconnection Readiness', '', '', '', ''],
        ['3.1', 'Spare 11 kV / 33 kV bay available (or new bay needed)', '', item.spare_bay, ''],
        ['3.2', 'Space for new bay, metering and protection panels', '', item.bay_space, ''],
        ['3.3', 'Switchgear condition (VCB, isolators, CTs, PTs)', '', item.switchgear, ''],
        ['3.4', 'Protection relays, SCADA / RTU availability', '', `${item.protection_relays} / ${item.scada_rtu}`, ''],
        ['3.5', 'Metering (ABT / energy meter) provision', '', item.metering, ''],
        ['3.6', 'Distance and route of evacuation line to solar plant', '', item.evacuation_distance, ''],
        ['3.7', 'Right-of-way issues', '', item.right_of_way, ''],
        ['4. Grid Quality & Safety', '', '', '', ''],
        ['4.1', 'Received Voltage By Solar Plant', '', item.received_voltage, ''],
        ['4.2', 'Fault level at the bus', '', item.fault_level, ''],
        ['4.3', 'Is Earthing system and lightning protection Is in good condition', '', item.lightning_protection, ''],
        ['4.4', 'Fire safety and clearances', '', item.fire_safety, ''],
        ['5. Solar Plant Sizing Inputs', '', '', '', ''],
        ['5.1', 'Feeder demand to be met (MW)', '', item.feeder_demand, ''],
        ['5.2', 'Existing Plant Connected Capacity ( MW)', '', item.existing_capacity, ''],
        ['5.3', 'Pending Capacity For Injecting Solar Plant (MW)', '', item.pending_capacity, ''],
        ['5.4', 'Plant capacity vs substation transformer capacity check', '', item.transformer_capacity_check, ''],
        ['6. Regulatory & Documentation', '', '', '', ''],
        ['6.1', 'Single line diagram (SLD) of substation', '', item.sld_available, ''],
        ['6.2', 'Environmental, forest and revenue clearances For Solar Plant To Substation( Transmission line )', '', item.environmental_clearance, ''],
        ['6.3', 'Feeder-wise load data', '', item.feeder_load_data, ''],
        ['7. Site Photos & Records', '', '', '', ''],
        ['7.1', 'Geo-tagged photos: land, substation, bays, approach road', '', item.geotagged_photos, ''],
        ['7.2', 'Coordinates (lat / long) and survey map', '', item.coordinates, ''],
        ['Surveyor signature:', '', item.surveyor_sign, '', ''],
        ['DISCOM representative signature:', '', item.discom_sign, '', '']
      ];
      const ws = XLSX.utils.aoa_to_sheet(excelData);
      ws['!cols'] = [{ wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }];
      XLSX.utils.book_append_sheet(wb, ws, 'Substation Survey');
    } 
    else if (type === 'land') {
      const excelData = [
        ['MSKVY Feeder Solarisation: Land Survey Checklist', '', '', '', ''],
        ['Village Name :', '', item.village, '', ''],
        ['Nearest Feeder Name :', '', item.feeder, '', ''],
        ['Surveyor / Date:', '', `${item.surveyor} - ${item.survey_date}`, '', ''],
        ['Sr.', 'Checklist Item', 'Status', 'Details / Value', 'Remarks'],
        ['1.1', 'Land within / adjacent to substation (ideally within 5 km)', item.status, item.land_nearby, item.distance_km],
        ['1.2', 'Ownership (govt / private / panchayat) and area in acres', '', item.ownership, ''],
        ['1.3', 'Land title, encumbrances and litigation status', '', item.title_status, ''],
        ['1.4', 'Topography, shading, flooding / waterlogging risk', '', `${item.topography} / ${item.waterlogging}`, ''],
        ['1.5', 'Soil type, approach road, access for heavy vehicles', '', item.approach_road, ''],
        ['1.6', 'Water availability for panel cleaning', '', item.water_availability, ''],
        ['1.7', 'Environmental, forest and revenue clearances For Solar  Plant To Substation( Transmission line )', '', item.environmental_clearance, ''],
        ['2. Site Photos & Records', '', '', '', ''],
        ['2.1', 'Geo-tagged photos: land, approach road', '', item.geotagged_photos, ''],
        ['2.2', 'Coordinates (lat / long) and survey map', '', item.coordinates, item.survey_map]
      ];
      const ws = XLSX.utils.aoa_to_sheet(excelData);
      ws['!cols'] = [{ wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }, { wch: 40 }];
      XLSX.utils.book_append_sheet(wb, ws, 'Land Survey');
    }

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${type === 'substation' ? 'Substation_Survey' : 'Land_Survey'}_${item.id}_${dayjs().format('YYYY-MM-DD')}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast.success('Excel exported successfully');
  } catch (e) { console.error(e); toast.error('Failed to export Excel'); }
};

// =============================================
// UI COMPONENTS
// =============================================
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
// DASHBOARD PAGE — SIRF 2 CARDS
// =============================================
const DashboardPage = ({ onNavigate }) => {
  const totalContractors = MSKVY_CONTRACTORS.length;
  const totalSubstationSurveys = MSKVY_SUBSTATION_SURVEYS.length;

  const cards = [
    {
      id: 'contractors',
      title: 'Contractor List',
      subtitle: 'All EPC Sub-Contractors',
      value: totalContractors,
      label: 'Total Records',
      icon: <Building2 size={24} />,
      color: '#2563eb',
      bg: '#eff6ff',
      onClick: () => onNavigate('mskvy', 'mskvy-contractors')
    },
    {
      id: 'substation-surveys',
      title: 'Substation Surveys',
      subtitle: 'Completed Checklist',
      value: totalSubstationSurveys,
      label: 'Surveys',
      icon: <ClipboardList size={24} />,
      color: '#1e40af',
      bg: '#dbeafe',
      onClick: () => onNavigate('mskvy', 'mskvy-substation-survey')
    },
  ];

  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}>
        <h1 style={S.title}>Management Dashboard</h1>
        <p style={S.subtitle}>Click any card to open the module</p>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 240px), 1fr))', gap: 'clamp(12px, 2.5vw, 20px)', maxWidth: '640px' }}>
        {cards.map(card => (
          <div key={card.id} onClick={card.onClick}
            style={{
              background: '#fff',
              borderRadius: '10px',
              padding: 'clamp(14px, 2.5vw, 18px)',
              border: '1px solid #e5e7eb',
              cursor: 'pointer',
              transition: 'all 0.2s',
              boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
              position: 'relative',
              overflow: 'hidden',
              userSelect: 'none'
            }}
            onMouseEnter={e => {
              e.currentTarget.style.transform = 'translateY(-3px)';
              e.currentTarget.style.boxShadow = '0 8px 16px rgba(0,0,0,0.1)';
              e.currentTarget.style.borderColor = card.color;
            }}
            onMouseLeave={e => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.04)';
              e.currentTarget.style.borderColor = '#e5e7eb';
            }}>
            <div style={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', background: card.color }} />
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '10px' }}>
              <div style={{ width: 'clamp(36px, 7vw, 44px)', height: 'clamp(36px, 7vw, 44px)', background: card.bg, color: card.color, borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{card.icon}</div>
              <div style={{ color: '#9ca3af' }}><ArrowRight size={14} /></div>
            </div>
            <p style={{ fontSize: 'clamp(11px, 2vw, 12px)', color: '#6b7280', margin: 0, fontWeight: '600', textTransform: 'uppercase', letterSpacing: '0.3px' }}>{card.title}</p>
            <p style={{ fontSize: 'clamp(20px, 4vw, 26px)', color: card.color, fontWeight: 'bold', margin: '4px 0 2px 0', lineHeight: 1 }}>{card.value}</p>
            <p style={{ fontSize: 'clamp(9px, 1.8vw, 10px)', color: '#9ca3af', margin: 0 }}>{card.label}</p>
          </div>
        ))}
      </div>
    </div>
  );
};

// =============================================
// MSKVY CONTRACTOR PAGE
// =============================================
const MSKVYContractorPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_contractors_v9', MSKVY_CONTRACTORS));
  const [searchTerm, setSearchTerm] = useState('');
  const [districtFilter, setDistrictFilter] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_contractors_v9', data); }, [data]);

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
                <Field label="Land Status" icon={<FileText size={11} />} name="land_status" value={formData.land_status} onChange={onChange} />
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
                    <div style={{ display: 'flex', gap: '1px' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }} title="View"><Eye size={12} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }} title="Edit"><Edit size={12} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }} title="Delete"><Trash2 size={12} /></button>
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
// SUBSTATION SURVEY PAGE — ACTIONS WITH PRINT CHECKLIST + EXCEL
// =============================================
const MSKVYSubstationSurveyPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_substation_v9', MSKVY_SUBSTATION_SURVEYS));
  const [searchTerm, setSearchTerm] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_substation_v9', data); }, [data]);

  const emptyForm = { 
    substation: '', subdivision: '', surveyor: '', survey_date: '', status: 'Pending', district: 'JALGAON',
    voltage_level: '', discom_division: '', connected_feeders: '', feeder_length: '', agri_pumps_load: '', peak_load: '', supply_hours: '', outage_history: '',
    transformer_capacity: '', transformer_make: '', transformer_year: '', loading_percent: '', spare_capacity: '', tap_changer: '', oil_condition: '', bushing_condition: '', earthing_condition: '',
    spare_bay: '', bay_space: '', switchgear: '', protection_relays: '', scada_rtu: '', metering: '', evacuation_distance: '', right_of_way: '',
    received_voltage: '', fault_level: '', lightning_protection: '', fire_safety: '',
    feeder_demand: '', existing_capacity: '', pending_capacity: '', transformer_capacity_check: '',
    sld_available: '', environmental_clearance: '', feeder_load_data: '',
    geotagged_photos: '', coordinates: '', surveyor_sign: '', discom_sign: '', remarks: ''
  };
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
            <h1 style={S.title}>{editingId ? 'Edit Survey Checklist' : 'New Survey Checklist'}</h1>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={13} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '12px', padding: 'clamp(12px, 2.5vw, 24px)', border: '1px solid #f3f4f6' }}>
            <form onSubmit={submit}>
              <table style={S.excelTable}>
                <thead>
                  <tr><th colSpan="5" style={S.excelHeader}>MSKVY Feeder Solarisation: Substation Survey Checklist</th></tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Substation:</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="substation" value={formData.substation} onChange={onChange} required style={S.excelInput} /></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Subdivision Name :</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="subdivision" value={formData.subdivision} onChange={onChange} style={S.excelInput} /></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor / Date:</td>
                    <td colSpan="3" style={S.excelTd}>
                      <input type="text" name="surveyor" value={formData.surveyor} onChange={onChange} required style={{...S.excelInput, width: '60%', marginRight: '10px'}} />
                      <input type="date" name="survey_date" value={formData.survey_date} onChange={onChange} style={{...S.excelInput, width: '30%'}} />
                    </td>
                  </tr>
                  <tr>
                    <th style={{...S.excelTh, width: '40px'}}>Sr.</th>
                    <th style={S.excelTh}>Checklist Item</th>
                    <th style={{...S.excelTh, width: '80px'}}>Status</th>
                    <th style={S.excelTh}>Details / Value</th>
                    <th style={S.excelTh}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {/* Section 1 */}
                  <tr><td colSpan="5" style={S.excelSection}>1. Substation & Feeder Basics</td></tr>
                  <tr>
                    <td style={S.excelTd}>1.1</td>
                    <td style={S.excelTd}>Substation name, voltage level (33/11 kV), DISCOM / circle / division</td>
                    <td style={S.excelTd}><input type="text" name="status" value={formData.status} onChange={onChange} style={S.excelInput} placeholder="Status" /></td>
                    <td style={S.excelTd}>
                      <input type="text" name="voltage_level" value={formData.voltage_level} onChange={onChange} style={S.excelInput} placeholder="Voltage" />
                      <input type="text" name="discom_division" value={formData.discom_division} onChange={onChange} style={S.excelInput} placeholder="Division" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.2</td>
                    <td style={S.excelTd}>Connected Feeder Details</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="connected_feeders" value={formData.connected_feeders} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.3</td>
                    <td style={S.excelTd}>Feeder length (km), no. of DTs</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="feeder_length" value={formData.feeder_length} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.4</td>
                    <td style={S.excelTd}>Connected agri pumps and total load (HP / kW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="agri_pumps_load" value={formData.agri_pumps_load} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.5</td>
                    <td style={S.excelTd}>Peak and average feeder load (last 12 months)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="peak_load" value={formData.peak_load} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.6</td>
                    <td style={S.excelTd}>Supply hours (day / night) and outage history</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="supply_hours" value={formData.supply_hours} onChange={onChange} style={S.excelInput} placeholder="Hours" />
                      <input type="text" name="outage_history" value={formData.outage_history} onChange={onChange} style={S.excelInput} placeholder="Outage History" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 2 */}
                  <tr><td colSpan="5" style={S.excelSection}>2. Transformer Details</td></tr>
                  <tr>
                    <td style={S.excelTd}>2.1</td>
                    <td style={S.excelTd}>Power transformer capacity (MVA), make, year of installation</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="transformer_capacity" value={formData.transformer_capacity} onChange={onChange} style={S.excelInput} placeholder="Capacity" />
                      <input type="text" name="transformer_make" value={formData.transformer_make} onChange={onChange} style={S.excelInput} placeholder="Make" />
                      <input type="text" name="transformer_year" value={formData.transformer_year} onChange={onChange} style={S.excelInput} placeholder="Year" />
                    </td>
                    <td style={S.excelTd}><input type="text" name="remarks" value={formData.remarks} onChange={onChange} style={S.excelInput} placeholder="Remarks" /></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.2</td>
                    <td style={S.excelTd}>Loading % and spare capacity available</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="loading_percent" value={formData.loading_percent} onChange={onChange} style={S.excelInput} placeholder="Loading %" />
                      <input type="text" name="spare_capacity" value={formData.spare_capacity} onChange={onChange} style={S.excelInput} placeholder="Spare" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.3</td>
                    <td style={S.excelTd}>Tap changer type / range</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="tap_changer" value={formData.tap_changer} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.4</td>
                    <td style={S.excelTd}>Condition of oil, bushings and earthing</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="oil_condition" value={formData.oil_condition} onChange={onChange} style={S.excelInput} placeholder="Oil Condition" />
                      <input type="text" name="bushing_condition" value={formData.bushing_condition} onChange={onChange} style={S.excelInput} placeholder="Bushing" />
                      <input type="text" name="earthing_condition" value={formData.earthing_condition} onChange={onChange} style={S.excelInput} placeholder="Earthing" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 3 */}
                  <tr><td colSpan="5" style={S.excelSection}>3. Interconnection Readiness</td></tr>
                  <tr>
                    <td style={S.excelTd}>3.1</td>
                    <td style={S.excelTd}>Spare 11 kV / 33 kV bay available (or new bay needed)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="spare_bay" value={formData.spare_bay} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.2</td>
                    <td style={S.excelTd}>Space for new bay, metering and protection panels</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="bay_space" value={formData.bay_space} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.3</td>
                    <td style={S.excelTd}>Switchgear condition (VCB, isolators, CTs, PTs)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="switchgear" value={formData.switchgear} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.4</td>
                    <td style={S.excelTd}>Protection relays, SCADA / RTU availability</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="protection_relays" value={formData.protection_relays} onChange={onChange} style={S.excelInput} placeholder="Relays" />
                      <input type="text" name="scada_rtu" value={formData.scada_rtu} onChange={onChange} style={S.excelInput} placeholder="SCADA / RTU" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.5</td>
                    <td style={S.excelTd}>Metering (ABT / energy meter) provision</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="metering" value={formData.metering} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.6</td>
                    <td style={S.excelTd}>Distance and route of evacuation line to solar plant</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="evacuation_distance" value={formData.evacuation_distance} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.7</td>
                    <td style={S.excelTd}>Right-of-way issues</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="right_of_way" value={formData.right_of_way} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 4 */}
                  <tr><td colSpan="5" style={S.excelSection}>4. Grid Quality & Safety</td></tr>
                  <tr>
                    <td style={S.excelTd}>4.1</td>
                    <td style={S.excelTd}>Received Voltage By Solar Plant</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="received_voltage" value={formData.received_voltage} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.2</td>
                    <td style={S.excelTd}>Fault level at the bus</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="fault_level" value={formData.fault_level} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.3</td>
                    <td style={S.excelTd}>Is Earthing system and lightning protection Is in good condition</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="lightning_protection" value={formData.lightning_protection} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.4</td>
                    <td style={S.excelTd}>Fire safety and clearances</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="fire_safety" value={formData.fire_safety} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 5 */}
                  <tr><td colSpan="5" style={S.excelSection}>5. Solar Plant Sizing Inputs</td></tr>
                  <tr>
                    <td style={S.excelTd}>5.1</td>
                    <td style={S.excelTd}>Feeder demand to be met (MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="feeder_demand" value={formData.feeder_demand} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.2</td>
                    <td style={S.excelTd}>Existing Plant Connected Capacity ( MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="existing_capacity" value={formData.existing_capacity} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.3</td>
                    <td style={S.excelTd}>Pending Capacity For Injecting Solar Plant (MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="pending_capacity" value={formData.pending_capacity} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.4</td>
                    <td style={S.excelTd}>Plant capacity vs substation transformer capacity check</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="transformer_capacity_check" value={formData.transformer_capacity_check} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 6 */}
                  <tr><td colSpan="5" style={S.excelSection}>6. Regulatory & Documentation</td></tr>
                  <tr>
                    <td style={S.excelTd}>6.1</td>
                    <td style={S.excelTd}>Single line diagram (SLD) of substation</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="sld_available" value={formData.sld_available} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>6.2</td>
                    <td style={S.excelTd}>Environmental, forest and revenue clearances For Solar Plant To Substation( Transmission line )</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="environmental_clearance" value={formData.environmental_clearance} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>6.3</td>
                    <td style={S.excelTd}>Feeder-wise load data</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="feeder_load_data" value={formData.feeder_load_data} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>

                  {/* Section 7 */}
                  <tr><td colSpan="5" style={S.excelSection}>7. Site Photos & Records</td></tr>
                  <tr>
                    <td style={S.excelTd}>7.1</td>
                    <td style={S.excelTd}>Geo-tagged photos: land, substation, bays, approach road</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="geotagged_photos" value={formData.geotagged_photos} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>7.2</td>
                    <td style={S.excelTd}>Coordinates (lat / long) and survey map</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="coordinates" value={formData.coordinates} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor signature:</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="surveyor_sign" value={formData.surveyor_sign} onChange={onChange} style={S.excelInput} /></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>DISCOM representative signature:</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="discom_sign" value={formData.discom_sign} onChange={onChange} style={S.excelInput} /></td>
                  </tr>
                </tbody>
              </table>
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
                    <div style={{ display: 'flex', gap: '1px', alignItems: 'center' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb', padding: '3px' }} title="View"><Eye size={12} /></button>
                      <button onClick={() => exportSingleExcel(item, 'substation')} style={{ ...S.iconBtn, color: '#16a34a', padding: '3px' }} title="Excel Checklist"><FileSpreadsheet size={12} /></button>
                      <button onClick={() => printSingleChecklist(item, 'substation')} style={{ ...S.iconBtn, color: '#ea580c', padding: '3px' }} title="Print Checklist"><Printer size={12} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed', padding: '3px' }} title="Edit"><Edit size={12} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626', padding: '3px' }} title="Delete"><Trash2 size={12} /></button>
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
          <div style={{ ...S.modal, maxWidth: '1000px' }} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}>
              <div>
                <h2 style={S.modalTitle}>Survey Checklist Details</h2>
                <p style={{ fontSize: '10px', color: '#6b7280', margin: '2px 0 0 0' }}>{selected.substation} • {selected.subdivision} • Surveyor: {selected.surveyor}</p>
              </div>
              <button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button>
            </div>
            <div style={S.modalBody}>
              
              <table style={S.excelTable}>
                <thead>
                  <tr><th colSpan="5" style={S.excelHeader}>MSKVY Feeder Solarisation: Substation Survey Checklist</th></tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold', width: '20%'}}>Substation:</td>
                    <td colSpan="3" style={S.excelTd}>{selected.substation}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Subdivision Name :</td>
                    <td colSpan="3" style={S.excelTd}>{selected.subdivision}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor / Date:</td>
                    <td colSpan="3" style={S.excelTd}>{selected.surveyor} - Date - {selected.survey_date ? dayjs(selected.survey_date).format('DD/MM/YYYY') : ''}</td>
                  </tr>
                  <tr>
                    <th style={{...S.excelTh, width: '40px'}}>Sr.</th>
                    <th style={S.excelTh}>Checklist Item</th>
                    <th style={{...S.excelTh, width: '80px'}}>Status</th>
                    <th style={S.excelTh}>Details / Value</th>
                    <th style={S.excelTh}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td colSpan="5" style={S.excelSection}>1. Substation & Feeder Basics</td></tr>
                  <tr>
                    <td style={S.excelTd}>1.1</td>
                    <td style={S.excelTd}>Substation name, voltage level (33/11 kV), DISCOM / circle / division</td>
                    <td style={S.excelTd}>{selected.status}</td>
                    <td style={S.excelTd}>{selected.substation}, {selected.voltage_level}, {selected.discom_division}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.2</td>
                    <td style={S.excelTd}>Connected Feeder Details</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.connected_feeders}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.3</td>
                    <td style={S.excelTd}>Feeder length (km), no. of DTs</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.feeder_length}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.4</td>
                    <td style={S.excelTd}>Connected agri pumps and total load (HP / kW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.agri_pumps_load || '—'}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.5</td>
                    <td style={S.excelTd}>Peak and average feeder load (last 12 months)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.peak_load || '—'}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.6</td>
                    <td style={S.excelTd}>Supply hours (day / night) and outage history</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.supply_hours} {selected.outage_history ? `, ${selected.outage_history}` : ''}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>2. Transformer Details</td></tr>
                  <tr>
                    <td style={S.excelTd}>2.1</td>
                    <td style={S.excelTd}>Power transformer capacity (MVA), make, year of installation</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.transformer_capacity}</td>
                    <td style={S.excelTd}>{selected.remarks}</td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.2</td>
                    <td style={S.excelTd}>Loading % and spare capacity available</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.loading_percent || '—'} / {selected.spare_capacity || '—'}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.3</td>
                    <td style={S.excelTd}>Tap changer type / range</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.tap_changer}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.4</td>
                    <td style={S.excelTd}>Condition of oil, bushings and earthing</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.oil_condition}, {selected.bushing_condition}, {selected.earthing_condition}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>3. Interconnection Readiness</td></tr>
                  <tr>
                    <td style={S.excelTd}>3.1</td>
                    <td style={S.excelTd}>Spare 11 kV / 33 kV bay available (or new bay needed)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.spare_bay}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.2</td>
                    <td style={S.excelTd}>Space for new bay, metering and protection panels</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.bay_space}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.3</td>
                    <td style={S.excelTd}>Switchgear condition (VCB, isolators, CTs, PTs)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.switchgear}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.4</td>
                    <td style={S.excelTd}>Protection relays, SCADA / RTU availability</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.protection_relays} / {selected.scada_rtu}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.5</td>
                    <td style={S.excelTd}>Metering (ABT / energy meter) provision</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.metering}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.6</td>
                    <td style={S.excelTd}>Distance and route of evacuation line to solar plant</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.evacuation_distance}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>3.7</td>
                    <td style={S.excelTd}>Right-of-way issues</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.right_of_way}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>4. Grid Quality & Safety</td></tr>
                  <tr>
                    <td style={S.excelTd}>4.1</td>
                    <td style={S.excelTd}>Received Voltage By Solar Plant</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.received_voltage}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.2</td>
                    <td style={S.excelTd}>Fault level at the bus</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.fault_level || '—'}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.3</td>
                    <td style={S.excelTd}>Is Earthing system and lightning protection Is in good condition</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.lightning_protection}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>4.4</td>
                    <td style={S.excelTd}>Fire safety and clearances</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.fire_safety}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>5. Solar Plant Sizing Inputs</td></tr>
                  <tr>
                    <td style={S.excelTd}>5.1</td>
                    <td style={S.excelTd}>Feeder demand to be met (MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.feeder_demand}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.2</td>
                    <td style={S.excelTd}>Existing Plant Connected Capacity ( MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.existing_capacity}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.3</td>
                    <td style={S.excelTd}>Pending Capacity For Injecting Solar Plant (MW)</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.pending_capacity}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>5.4</td>
                    <td style={S.excelTd}>Plant capacity vs substation transformer capacity check</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.transformer_capacity_check}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>6. Regulatory & Documentation</td></tr>
                  <tr>
                    <td style={S.excelTd}>6.1</td>
                    <td style={S.excelTd}>Single line diagram (SLD) of substation</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.sld_available}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>6.2</td>
                    <td style={S.excelTd}>Environmental, forest and revenue clearances For Solar Plant To Substation( Transmission line )</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.environmental_clearance}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>6.3</td>
                    <td style={S.excelTd}>Feeder-wise load data</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.feeder_load_data}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>7. Site Photos & Records</td></tr>
                  <tr>
                    <td style={S.excelTd}>7.1</td>
                    <td style={S.excelTd}>Geo-tagged photos: land, substation, bays, approach road</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.geotagged_photos}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>7.2</td>
                    <td style={S.excelTd}>Coordinates (lat / long) and survey map</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.coordinates}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor signature:</td>
                    <td colSpan="3" style={S.excelTd}>{selected.surveyor_sign}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>DISCOM representative signature:</td>
                    <td colSpan="3" style={S.excelTd}>{selected.discom_sign}</td>
                  </tr>
                </tbody>
              </table>
              
              {selected.coordinates && (
                <div style={{ marginTop: '16px', display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                  <a href={getGoogleEarthUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#0d9488', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <Globe size={14} /> Open {selected.coordinates} in Google Earth
                  </a>
                  <a href={getGoogleMapsUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#2563eb', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <MapPin size={14} /> Google Maps
                  </a>
                </div>
              )}

            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => exportSingleExcel(selected, 'substation')} style={{ ...S.btnSecondary, color: '#16a34a', borderColor: '#bbf7d0' }}><FileSpreadsheet size={13} /> Excel</button>
              <button onClick={() => printSingleChecklist(selected, 'substation')} style={{ ...S.btnSecondary, color: '#ea580c', borderColor: '#fed7aa' }}><Printer size={13} /> Print</button>
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
// LAND SURVEY PAGE — ACTIONS WITH PRINT CHECKLIST + EXCEL
// =============================================
const MSKVYLandSurveyPage = () => {
  const [data, setData] = useState(() => loadMSKVY('mskvy_land_v9', MSKVY_LAND_SURVEYS));
  const [searchTerm, setSearchTerm] = useState('');
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_land_v9', data); }, [data]);

  const emptyForm = { 
    village: '', block: '', district: 'JALGAON', feeder: '', surveyor: '', survey_date: '', status: 'Completed',
    land_nearby: '', distance_km: '', ownership: '', area_acres: '', title_status: '', litigation: '',
    topography: '', waterlogging: '', soil_type: '', approach_road: '', water_availability: '', environmental_clearance: '',
    geotagged_photos: '', coordinates: '', survey_map: '', remarks: '' 
  };
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
              <table style={S.excelTable}>
                <thead>
                  <tr><th colSpan="5" style={S.excelHeader}>MSKVY Feeder Solarisation: Land Survey Checklist</th></tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Village Name :</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="village" value={formData.village} onChange={onChange} required style={S.excelInput} /></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Nearest Feeder Name :</td>
                    <td colSpan="3" style={S.excelTd}><input type="text" name="feeder" value={formData.feeder} onChange={onChange} style={S.excelInput} /></td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor / Date:</td>
                    <td colSpan="3" style={S.excelTd}>
                      <input type="text" name="surveyor" value={formData.surveyor} onChange={onChange} required style={{...S.excelInput, width: '60%', marginRight: '10px'}} />
                      <input type="date" name="survey_date" value={formData.survey_date} onChange={onChange} style={{...S.excelInput, width: '30%'}} />
                    </td>
                  </tr>
                  <tr>
                    <th style={{...S.excelTh, width: '40px'}}>Sr.</th>
                    <th style={S.excelTh}>Checklist Item</th>
                    <th style={{...S.excelTh, width: '80px'}}>Status</th>
                    <th style={S.excelTh}>Details / Value</th>
                    <th style={S.excelTh}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td colSpan="5" style={S.excelSection}>1. Land Suitability</td></tr>
                  <tr>
                    <td style={S.excelTd}>1.1</td>
                    <td style={S.excelTd}>Land within / adjacent to substation (ideally within 5 km)</td>
                    <td style={S.excelTd}><input type="text" name="status" value={formData.status} onChange={onChange} style={S.excelInput} placeholder="Status" /></td>
                    <td style={S.excelTd}><input type="text" name="land_nearby" value={formData.land_nearby} onChange={onChange} style={S.excelInput} placeholder="Yes/No" /></td>
                    <td style={S.excelTd}><input type="text" name="distance_km" value={formData.distance_km} onChange={onChange} style={S.excelInput} placeholder="Distance" /></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.2</td>
                    <td style={S.excelTd}>Ownership (govt / private / panchayat) and area in acres</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="ownership" value={formData.ownership} onChange={onChange} style={S.excelInput} placeholder="Ownership" />
                      <input type="text" name="area_acres" value={formData.area_acres} onChange={onChange} style={S.excelInput} placeholder="Area" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.3</td>
                    <td style={S.excelTd}>Land title, encumbrances and litigation status</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="title_status" value={formData.title_status} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.4</td>
                    <td style={S.excelTd}>Topography, shading, flooding / waterlogging risk</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="topography" value={formData.topography} onChange={onChange} style={S.excelInput} placeholder="Topography" />
                      <input type="text" name="waterlogging" value={formData.waterlogging} onChange={onChange} style={S.excelInput} placeholder="Waterlogging" />
                    </td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.5</td>
                    <td style={S.excelTd}>Soil type, approach road, access for heavy vehicles</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="approach_road" value={formData.approach_road} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.6</td>
                    <td style={S.excelTd}>Water availability for panel cleaning</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="water_availability" value={formData.water_availability} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.7</td>
                    <td style={S.excelTd}>Environmental, forest and revenue clearances For Solar  Plant To Substation( Transmission line )</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="environmental_clearance" value={formData.environmental_clearance} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>2. Site Photos & Records</td></tr>
                  <tr>
                    <td style={S.excelTd}>2.1</td>
                    <td style={S.excelTd}>Geo-tagged photos: land, approach road</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}><input type="text" name="geotagged_photos" value={formData.geotagged_photos} onChange={onChange} style={S.excelInput} /></td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.2</td>
                    <td style={S.excelTd}>Coordinates (lat / long) and survey map</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>
                      <input type="text" name="coordinates" value={formData.coordinates} onChange={onChange} style={S.excelInput} placeholder="Coordinates" />
                      <input type="text" name="survey_map" value={formData.survey_map} onChange={onChange} style={S.excelInput} placeholder="Map" />
                    </td>
                    <td style={S.excelTd}><input type="text" name="remarks" value={formData.remarks} onChange={onChange} style={S.excelInput} placeholder="Remarks" /></td>
                  </tr>
                </tbody>
              </table>
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
                    <div style={{ display: 'flex', gap: '1px', alignItems: 'center' }}>
                      <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb', padding: '3px' }} title="View"><Eye size={12} /></button>
                      <button onClick={() => exportSingleExcel(item, 'land')} style={{ ...S.iconBtn, color: '#16a34a', padding: '3px' }} title="Excel Checklist"><FileSpreadsheet size={12} /></button>
                      <button onClick={() => printSingleChecklist(item, 'land')} style={{ ...S.iconBtn, color: '#ea580c', padding: '3px' }} title="Print Checklist"><Printer size={12} /></button>
                      <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed', padding: '3px' }} title="Edit"><Edit size={12} /></button>
                      <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626', padding: '3px' }} title="Delete"><Trash2 size={12} /></button>
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
          <div style={{ ...S.modal, maxWidth: '1000px' }} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}>
              <div>
                <h2 style={S.modalTitle}>Land Survey Checklist Details</h2>
                <p style={{ fontSize: '10px', color: '#6b7280', margin: '2px 0 0 0' }}>{selected.village} • {selected.block} • Surveyor: {selected.surveyor}</p>
              </div>
              <button onClick={() => setShowView(false)} style={{ ...S.iconBtn }}><X size={18} /></button>
            </div>
            <div style={S.modalBody}>
              
              <table style={S.excelTable}>
                <thead>
                  <tr><th colSpan="5" style={S.excelHeader}>MSKVY Feeder Solarisation: Land Survey Checklist</th></tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Village Name :</td>
                    <td colSpan="3" style={S.excelTd}>{selected.village}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Block / District :</td>
                    <td colSpan="3" style={S.excelTd}>{selected.block} / {selected.district}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Nearest Feeder Name :</td>
                    <td colSpan="3" style={S.excelTd}>{selected.feeder}</td>
                  </tr>
                  <tr>
                    <td colSpan="2" style={{...S.excelTd, fontWeight: 'bold'}}>Surveyor / Date:</td>
                    <td colSpan="3" style={S.excelTd}>{selected.surveyor} - Date - {selected.survey_date ? dayjs(selected.survey_date).format('DD/MM/YYYY') : ''}</td>
                  </tr>
                  <tr>
                    <th style={{...S.excelTh, width: '40px'}}>Sr.</th>
                    <th style={S.excelTh}>Checklist Item</th>
                    <th style={{...S.excelTh, width: '80px'}}>Status</th>
                    <th style={S.excelTh}>Details / Value</th>
                    <th style={S.excelTh}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td colSpan="5" style={S.excelSection}>1. Land Suitability</td></tr>
                  <tr>
                    <td style={S.excelTd}>1.1</td>
                    <td style={S.excelTd}>Land within / adjacent to substation (ideally within 5 km)</td>
                    <td style={S.excelTd}>{selected.status}</td>
                    <td style={S.excelTd}>{selected.land_nearby}</td>
                    <td style={S.excelTd}>{selected.distance_km}</td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.2</td>
                    <td style={S.excelTd}>Ownership (govt / private / panchayat) and area in acres</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.ownership} {selected.area_acres && `- ${selected.area_acres}`}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.3</td>
                    <td style={S.excelTd}>Land title, encumbrances and litigation status</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.title_status}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.4</td>
                    <td style={S.excelTd}>Topography, shading, flooding / waterlogging risk</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.topography} / {selected.waterlogging}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.5</td>
                    <td style={S.excelTd}>Soil type, approach road, access for heavy vehicles</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.approach_road}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.6</td>
                    <td style={S.excelTd}>Water availability for panel cleaning</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.water_availability}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>1.7</td>
                    <td style={S.excelTd}>Environmental, forest and revenue clearances For Solar  Plant To Substation( Transmission line )</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.environmental_clearance}</td>
                    <td style={S.excelTd}></td>
                  </tr>

                  <tr><td colSpan="5" style={S.excelSection}>2. Site Photos & Records</td></tr>
                  <tr>
                    <td style={S.excelTd}>2.1</td>
                    <td style={S.excelTd}>Geo-tagged photos: land, approach road</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.geotagged_photos}</td>
                    <td style={S.excelTd}></td>
                  </tr>
                  <tr>
                    <td style={S.excelTd}>2.2</td>
                    <td style={S.excelTd}>Coordinates (lat / long) and survey map</td>
                    <td style={S.excelTd}></td>
                    <td style={S.excelTd}>{selected.coordinates}</td>
                    <td style={S.excelTd}>{selected.survey_map}</td>
                  </tr>
                </tbody>
              </table>
              
              {selected.coordinates && (
                <div style={{ marginTop: '16px', display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                  <a href={getGoogleEarthUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#0d9488', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <Globe size={14} /> Open in Google Earth
                  </a>
                  <a href={getGoogleMapsUrl(selected.coordinates)} target="_blank" rel="noopener noreferrer"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '8px 14px', background: '#2563eb', color: '#fff', borderRadius: '8px', textDecoration: 'none', fontSize: '12px', fontWeight: '600' }}>
                    <MapPin size={14} /> Google Maps
                  </a>
                </div>
              )}

            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => exportSingleExcel(selected, 'land')} style={{ ...S.btnSecondary, color: '#16a34a', borderColor: '#bbf7d0' }}><FileSpreadsheet size={13} /> Excel</button>
              <button onClick={() => printSingleChecklist(selected, 'land')} style={{ ...S.btnSecondary, color: '#ea580c', borderColor: '#fed7aa' }}><Printer size={13} /> Print</button>
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
  const [data, setData] = useState(() => loadMSKVY('mskvy_district_v9', MASTER_MSKVY_DISTRICT));
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [showView, setShowView] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [page, setPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  useEffect(() => { saveMSKVY('mskvy_district_v9', data); }, [data]);

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
  const [data, setData] = useState(() => loadMSKVY('mskvy_survey_status_v9', MASTER_MSKVY_SURVEY_STATUS));
  const [showForm, setShowForm] = useState(false);
  const [showView, setShowView] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const emptyForm = { code: '', name: '', status: 'Active', remarks: '' };
  const [formData, setFormData] = useState(emptyForm);
  useEffect(() => { saveMSKVY('mskvy_survey_status_v9', data); }, [data]);

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