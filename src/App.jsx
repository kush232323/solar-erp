import React, { useState, useEffect, useCallback } from 'react';
import { HashRouter as Router } from 'react-router-dom';
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import dayjs from 'dayjs';
import * as XLSX from 'xlsx';
import {
  Search, Filter, Plus, Edit, Trash2, Eye, X, ChevronLeft, ChevronRight,
  RefreshCw, ChevronDown, CheckCircle, Printer, FileSpreadsheet,
  Hash, User, Globe, FileText, Shield, Calendar, Tag, Gauge, Camera, MapPin,
  Users, Building2, FolderOpen, ClipboardList, Bell, LogOut, Lock, Mail,
  Phone, Award, TrendingUp, AlertCircle, Clock, Database, Columns,
  Smartphone, Download, Upload, Activity, Layers, FileCheck,
  MapPinned, Banknote, Wallet, List, Wifi, Bug, Rocket, Server,
  BellRing, PlayCircle, Box
} from 'lucide-react';

// =============================================
// DEMO DATA
// =============================================
const DEMO_PROJECTS = [
  { id: 1, s_no: 1, site_code: 'SP-MUM-001', name: 'Rooftop Solar Plant - Unit A', location: 'Mumbai Refinery', category: 'Rooftop Solar', capacity: '250 kW', contractor: 'SunPower Ltd', field_worker: 'Rajesh Kumar', status: 'Approved', progress: 75, start_date: '2024-01-10', end_date: '2024-03-15', remarks: 'On track' },
  { id: 2, s_no: 2, site_code: 'SP-KOY-002', name: 'Ground Mount Solar - Plant 2', location: 'Koyali Refinery', category: 'Ground Mount', capacity: '500 kW', contractor: 'GreenTech Solutions', field_worker: 'Suresh Patel', status: 'Pending', progress: 30, start_date: '2024-02-01', end_date: '2024-05-30', remarks: 'Awaiting approval' },
  { id: 3, s_no: 3, site_code: 'SP-PAN-003', name: 'Solar Water Heater - Admin Block', location: 'Panipat Refinery', category: 'Solar Water Heater', capacity: '100 kW', contractor: 'EcoEnergy Pvt Ltd', field_worker: 'Amit Singh', status: 'Correction', progress: 60, start_date: '2024-01-20', end_date: '2024-04-10', remarks: 'Photo re-upload needed' },
  { id: 4, s_no: 4, site_code: 'SP-BAR-004', name: 'Solar Street Light - Township', location: 'Barauni Refinery', category: 'Street Light', capacity: '50 kW', contractor: 'SunPower Ltd', field_worker: 'Vikram Sharma', status: 'Completed', progress: 100, start_date: '2023-11-05', end_date: '2024-01-20', remarks: 'Completed' },
  { id: 5, s_no: 5, site_code: 'SP-MAT-005', name: 'Solar Pump Installation - Site 5', location: 'Mathura Refinery', category: 'Solar Pump', capacity: '75 kW', contractor: 'GreenTech Solutions', field_worker: 'Deepak Verma', status: 'Approved', progress: 45, start_date: '2024-02-15', end_date: '2024-05-01', remarks: 'In progress' },
  { id: 6, s_no: 6, site_code: 'SP-JAM-006', name: 'Solar Panel Cleaning System', location: 'Jamnagar Refinery', category: 'Rooftop Solar', capacity: '200 kW', contractor: 'EcoEnergy Pvt Ltd', field_worker: 'Ravi Nair', status: 'Pending', progress: 15, start_date: '2024-03-01', end_date: '2024-06-30', remarks: 'Just started' },
  { id: 7, s_no: 7, site_code: 'SP-GUW-007', name: 'Solar Rooftop - Guwahati', location: 'Guwahati Refinery', category: 'Rooftop Solar', capacity: '150 kW', contractor: 'SunPower Ltd', field_worker: 'Manoj Das', status: 'Approved', progress: 55, start_date: '2024-01-25', end_date: '2024-04-20', remarks: 'Good progress' },
  { id: 8, s_no: 8, site_code: 'SP-DIG-008', name: 'Ground Solar - Digboi', location: 'Digboi Refinery', category: 'Ground Mount', capacity: '300 kW', contractor: 'GreenTech Solutions', field_worker: 'Pankaj Gogoi', status: 'Pending', progress: 10, start_date: '2024-03-10', end_date: '2024-07-15', remarks: 'Survey ongoing' },
];

const DEMO_CONTRACTORS = [
  { id: 1, s_no: 1, name: 'SunPower Ltd', type: 'Solar Installer', contact: '9876543210', email: 'info@sunpower.com', license: 'LIC-2024-001', license_type: 'Electrical License', projects: 3, status: 'Approved', joined: '2023-06-15', remarks: 'Top performer' },
  { id: 2, s_no: 2, name: 'GreenTech Solutions', type: 'Solar Installer', contact: '9876543211', email: 'contact@greentech.com', license: 'LIC-2024-002', license_type: 'Electrical License', projects: 3, status: 'Approved', joined: '2023-08-20', remarks: 'Good' },
  { id: 3, s_no: 3, name: 'EcoEnergy Pvt Ltd', type: 'Civil Contractor', contact: '9876543212', email: 'sales@ecoenergy.com', license: 'LIC-2024-003', license_type: 'Civil License', projects: 2, status: 'Pending', joined: '2024-01-10', remarks: 'Verification pending' },
  { id: 4, s_no: 4, name: 'Bharat Solar Pvt Ltd', type: 'Electrical Contractor', contact: '9876543213', email: 'bharat@bharatsolar.com', license: 'LIC-2024-004', license_type: 'Electrical License', projects: 1, status: 'Approved', joined: '2023-11-05', remarks: 'New' },
  { id: 5, s_no: 5, name: 'MegaSun Technologies', type: 'Mechanical Contractor', contact: '9876543214', email: 'info@megasun.com', license: 'LIC-2024-005', license_type: 'Safety License', projects: 0, status: 'Pending', joined: '2024-02-20', remarks: 'Under review' },
];

const DEMO_FIELD_WORKERS = [
  { id: 1, s_no: 1, name: 'Rajesh Kumar', phone: '9876500001', role: 'Installer', skill: 'Panel Installation', contractor: 'SunPower Ltd', assigned_sites: 3, status: 'Active', last_activity: '2024-03-10', remarks: 'Active' },
  { id: 2, s_no: 2, name: 'Suresh Patel', phone: '9876500002', role: 'Supervisor', skill: 'Wiring', contractor: 'GreenTech Solutions', assigned_sites: 2, status: 'Active', last_activity: '2024-03-12', remarks: 'Active' },
  { id: 3, s_no: 3, name: 'Amit Singh', phone: '9876500003', role: 'Technician', skill: 'Mounting', contractor: 'EcoEnergy Pvt Ltd', assigned_sites: 2, status: 'Active', last_activity: '2024-03-11', remarks: 'Active' },
  { id: 4, s_no: 4, name: 'Vikram Sharma', phone: '9876500004', role: 'Installer', skill: 'Panel Installation', contractor: 'SunPower Ltd', assigned_sites: 1, status: 'Active', last_activity: '2024-03-08', remarks: 'Active' },
  { id: 5, s_no: 5, name: 'Deepak Verma', phone: '9876500005', role: 'Helper', skill: 'Wiring', contractor: 'GreenTech Solutions', assigned_sites: 2, status: 'Inactive', last_activity: '2024-02-28', remarks: 'On leave' },
  { id: 6, s_no: 6, name: 'Ravi Nair', phone: '9876500006', role: 'Installer', skill: 'Panel Installation', contractor: 'EcoEnergy Pvt Ltd', assigned_sites: 2, status: 'Active', last_activity: '2024-03-09', remarks: 'Active' },
  { id: 7, s_no: 7, name: 'Manoj Das', phone: '9876500007', role: 'Technician', skill: 'Mounting', contractor: 'SunPower Ltd', assigned_sites: 1, status: 'Active', last_activity: '2024-03-10', remarks: 'Active' },
  { id: 8, s_no: 8, name: 'Pankaj Gogoi', phone: '9876500008', role: 'Supervisor', skill: 'Wiring', contractor: 'GreenTech Solutions', assigned_sites: 1, status: 'Active', last_activity: '2024-03-11', remarks: 'Active' },
];

const DEMO_REPORTS = [
  { id: 1, s_no: 1, project: 'Rooftop Solar Plant - Unit A', worker: 'Rajesh Kumar', report_type: 'Daily Report', priority: 'High', submitted: '2024-03-10', status: 'Approved', photos: 8, gps: 'Yes', remarks: 'All panels installed' },
  { id: 2, s_no: 2, project: 'Ground Mount Solar - Plant 2', worker: 'Suresh Patel', report_type: 'Weekly Report', priority: 'Medium', submitted: '2024-03-12', status: 'Pending', photos: 12, gps: 'Yes', remarks: 'Awaiting review' },
  { id: 3, s_no: 3, project: 'Solar Water Heater - Admin Block', worker: 'Amit Singh', report_type: 'Daily Report', priority: 'High', submitted: '2024-03-11', status: 'Correction', photos: 5, gps: 'Yes', remarks: 'Photo quality poor' },
  { id: 4, s_no: 4, project: 'Solar Street Light - Township', worker: 'Vikram Sharma', report_type: 'Final Report', priority: 'High', submitted: '2024-03-08', status: 'Approved', photos: 15, gps: 'Yes', remarks: 'Completed' },
  { id: 5, s_no: 5, project: 'Solar Pump Installation - Site 5', worker: 'Deepak Verma', report_type: 'Daily Report', priority: 'Medium', submitted: '2024-03-09', status: 'Approved', photos: 10, gps: 'Yes', remarks: 'GPS verified' },
  { id: 6, s_no: 6, project: 'Solar Panel Cleaning System', worker: 'Ravi Nair', report_type: 'Daily Report', priority: 'Low', submitted: '2024-03-10', status: 'Pending', photos: 6, gps: 'Yes', remarks: 'Initial report' },
];

const DEMO_USERS = [
  { id: 1, s_no: 1, name: 'Admin User', email: 'admin@iocl.com', role: 'Super Admin', status: 'Active', remarks: 'System admin' },
  { id: 2, s_no: 2, name: 'SunPower Ltd', email: 'info@sunpower.com', role: 'Contractor', status: 'Active', remarks: 'Approved' },
  { id: 3, s_no: 3, name: 'GreenTech Solutions', email: 'contact@greentech.com', role: 'Contractor', status: 'Active', remarks: 'Approved' },
  { id: 4, s_no: 4, name: 'Rajesh Kumar', email: 'rajesh@iocl.com', role: 'Field Worker', status: 'Active', remarks: 'Assigned' },
  { id: 5, s_no: 5, name: 'EcoEnergy Pvt Ltd', email: 'sales@ecoenergy.com', role: 'Contractor', status: 'Active', remarks: 'Approved' },
  { id: 6, s_no: 6, name: 'Suresh Patel', email: 'suresh@iocl.com', role: 'Field Worker', status: 'Active', remarks: 'Assigned' },
];

const DEMO_PHOTOS = Array.from({ length: 12 }, (_, i) => ({
  id: i + 1, s_no: i + 1, name: `Site Photo #${i + 1}`,
  project: ['Rooftop Solar Plant - Unit A', 'Ground Mount Solar - Plant 2', 'Solar Water Heater - Admin Block'][i % 3],
  worker: ['Rajesh Kumar', 'Suresh Patel', 'Amit Singh'][i % 3],
  media_type: 'Photograph', category: ['Site Photo', 'Progress Photo', 'Issue Photo'][i % 3],
  gps: '19.0760, 72.8777', date: '2024-03-10', status: i % 3 === 0 ? 'Approved' : (i % 3 === 1 ? 'Pending' : 'Approved')
}));

const DEMO_DOCUMENTS = [
  { id: 1, s_no: 1, name: 'Project Quotation.pdf', project: 'Rooftop Solar Plant - Unit A', type: 'Quotation', size: '2.4 MB', uploaded: '2024-01-15', status: 'Approved' },
  { id: 2, s_no: 2, name: 'Site Layout.pdf', project: 'Ground Mount Solar - Plant 2', type: 'Layout', size: '1.8 MB', uploaded: '2024-02-01', status: 'Approved' },
  { id: 3, s_no: 3, name: 'License Copy.pdf', project: 'Solar Water Heater - Admin Block', type: 'License', size: '0.9 MB', uploaded: '2024-02-10', status: 'Pending' },
  { id: 4, s_no: 4, name: 'Approval Letter.pdf', project: 'Solar Street Light - Township', type: 'Approval', size: '1.2 MB', uploaded: '2024-01-20', status: 'Approved' },
  { id: 5, s_no: 5, name: 'Technical Spec.pdf', project: 'Solar Pump Installation - Site 5', type: 'Specification', size: '3.1 MB', uploaded: '2024-02-10', status: 'Approved' },
];

const DEMO_ACTIVITIES = [
  { id: 1, s_no: 1, type: 'Report Submission', description: 'Rajesh Kumar submitted field report', user: 'Rajesh Kumar', date: '2024-03-10 14:30', status: 'Completed' },
  { id: 2, s_no: 2, type: 'Project Approval', description: 'Admin approved SunPower project', user: 'Admin User', date: '2024-03-10 11:15', status: 'Completed' },
  { id: 3, s_no: 3, type: 'Correction Request', description: 'Correction for Solar Heater project', user: 'Admin User', date: '2024-03-09 16:45', status: 'Pending' },
  { id: 4, s_no: 4, type: 'Contractor Registration', description: 'New contractor registered', user: 'Bharat Solar Pvt Ltd', date: '2024-03-08 09:20', status: 'Pending' },
  { id: 5, s_no: 5, type: 'Photo Upload', description: '12 photos uploaded', user: 'Suresh Patel', date: '2024-03-08 15:00', status: 'Completed' },
];

// Project sub-pages
const DEMO_PROJECT_SITES = [
  { id: 1, s_no: 1, site_code: 'SITE-MUM-01', project: 'Rooftop Solar Plant - Unit A', category: 'Rooftop Solar', location: 'Mumbai Refinery', area: '2500 sqft', gps: '19.0760, 72.8777', status: 'Active', remarks: 'Prime location' },
  { id: 2, s_no: 2, site_code: 'SITE-KOY-02', project: 'Ground Mount Solar - Plant 2', category: 'Ground Mount', location: 'Koyali Refinery', area: '5000 sqft', gps: '22.3126, 73.1812', status: 'Active', remarks: 'Large ground area' },
  { id: 3, s_no: 3, site_code: 'SITE-PAN-03', project: 'Solar Water Heater - Admin Block', category: 'Solar Water Heater', location: 'Panipat Refinery', area: '800 sqft', gps: '29.3909, 76.9635', status: 'Active', remarks: 'Admin roof' },
  { id: 4, s_no: 4, site_code: 'SITE-BAR-04', project: 'Solar Street Light - Township', category: 'Street Light', location: 'Barauni Refinery', area: '1500 sqft', gps: '25.4667, 85.9667', status: 'Completed', remarks: 'Township area' },
];

const DEMO_PROJECT_MILESTONES = [
  { id: 1, s_no: 1, project: 'Rooftop Solar Plant - Unit A', milestone_type: 'Site Survey', milestone: 'Site Survey Complete', target_date: '2024-01-15', actual_date: '2024-01-14', status: 'Completed', remarks: 'Ahead of schedule' },
  { id: 2, s_no: 2, project: 'Ground Mount Solar - Plant 2', milestone_type: 'Foundation', milestone: 'Foundation Work', target_date: '2024-02-20', actual_date: '2024-02-25', status: 'Completed', remarks: '5 days late' },
  { id: 3, s_no: 3, project: 'Solar Water Heater - Admin Block', milestone_type: 'Installation', milestone: 'Panel Installation', target_date: '2024-03-15', actual_date: '', status: 'In Progress', remarks: 'Ongoing' },
  { id: 4, s_no: 4, project: 'Solar Street Light - Township', milestone_type: 'Testing', milestone: 'Final Testing', target_date: '2024-01-15', actual_date: '2024-01-15', status: 'Completed', remarks: 'On time' },
];

const DEMO_PROJECT_BUDGETS = [
  { id: 1, s_no: 1, project: 'Rooftop Solar Plant - Unit A', category: 'Rooftop Solar', sanctioned: '2500000', spent: '1875000', balance: '625000', utilization: '75%', status: 'Approved', remarks: 'On budget' },
  { id: 2, s_no: 2, project: 'Ground Mount Solar - Plant 2', category: 'Ground Mount', sanctioned: '5000000', spent: '1500000', balance: '3500000', utilization: '30%', status: 'Approved', remarks: 'Early stage' },
  { id: 3, s_no: 3, project: 'Solar Water Heater - Admin Block', category: 'Solar Water Heater', sanctioned: '1000000', spent: '600000', balance: '400000', utilization: '60%', status: 'Correction', remarks: 'Pending approval' },
  { id: 4, s_no: 4, project: 'Solar Street Light - Township', category: 'Street Light', sanctioned: '500000', spent: '500000', balance: '0', utilization: '100%', status: 'Completed', remarks: 'Fully utilized' },
];

const DEMO_PROJECT_DOCUMENTS = [
  { id: 1, s_no: 1, project: 'Rooftop Solar Plant - Unit A', doc_name: 'Project Quotation.pdf', doc_type: 'Quotation', size: '2.4 MB', uploaded: '2024-01-05', status: 'Approved', remarks: 'Final version' },
  { id: 2, s_no: 2, project: 'Ground Mount Solar - Plant 2', doc_name: 'Site Layout.pdf', doc_type: 'Layout', size: '1.8 MB', uploaded: '2024-01-28', status: 'Approved', remarks: 'Approved by admin' },
  { id: 3, s_no: 3, project: 'Solar Water Heater - Admin Block', doc_name: 'License Copy.pdf', doc_type: 'License', size: '0.9 MB', uploaded: '2024-01-18', status: 'Pending', remarks: 'Under review' },
];

const DEMO_PROJECT_ASSIGNMENTS = [
  { id: 1, s_no: 1, project: 'Rooftop Solar Plant - Unit A', contractor: 'SunPower Ltd', worker: 'Rajesh Kumar', role: 'Installer', assigned_date: '2024-01-12', status: 'Active', remarks: 'Lead installer' },
  { id: 2, s_no: 2, project: 'Ground Mount Solar - Plant 2', contractor: 'GreenTech Solutions', worker: 'Suresh Patel', role: 'Supervisor', assigned_date: '2024-02-05', status: 'Active', remarks: 'Field supervisor' },
  { id: 3, s_no: 3, project: 'Solar Water Heater - Admin Block', contractor: 'EcoEnergy Pvt Ltd', worker: 'Amit Singh', role: 'Technician', assigned_date: '2024-01-22', status: 'Active', remarks: 'Technical work' },
];

// Application data
const DEMO_CONTRACTOR_APPS = [
  { id: 1, s_no: 1, contractor: 'SunPower Ltd', app_name: 'Contractor App', version: '1.2.0', device: 'Android 12', os_version: '12.0', last_sync: '2024-03-10 14:30', status: 'Active', projects: 3, remarks: 'Latest version' },
  { id: 2, s_no: 2, contractor: 'GreenTech Solutions', app_name: 'Contractor App', version: '1.1.5', device: 'Android 11', os_version: '11.0', last_sync: '2024-03-12 09:15', status: 'Active', projects: 3, remarks: 'Update available' },
  { id: 3, s_no: 3, contractor: 'EcoEnergy Pvt Ltd', app_name: 'Contractor App', version: '1.0.8', device: 'Android 10', os_version: '10.0', last_sync: '2024-03-08 17:45', status: 'Inactive', projects: 2, remarks: 'Old version' },
  { id: 4, s_no: 4, contractor: 'Bharat Solar Pvt Ltd', app_name: 'Contractor App', version: '1.2.0', device: 'Android 13', os_version: '13.0', last_sync: '2024-03-11 10:20', status: 'Active', projects: 1, remarks: 'Latest' },
];

const DEMO_WORKER_APPS = [
  { id: 1, s_no: 1, worker: 'Rajesh Kumar', app_name: 'Field Worker App', version: '1.2.0', device: 'Android 12', last_sync: '2024-03-10 14:30', gps: 'Enabled', status: 'Active', remarks: 'Active' },
  { id: 2, s_no: 2, worker: 'Suresh Patel', app_name: 'Field Worker App', version: '1.2.0', device: 'Android 12', last_sync: '2024-03-12 09:15', gps: 'Enabled', status: 'Active', remarks: 'Active' },
  { id: 3, s_no: 3, worker: 'Amit Singh', app_name: 'Field Worker App', version: '1.1.5', device: 'Android 11', last_sync: '2024-03-11 16:20', gps: 'Enabled', status: 'Active', remarks: 'Update available' },
  { id: 4, s_no: 4, worker: 'Vikram Sharma', app_name: 'Field Worker App', version: '1.2.0', device: 'Android 12', last_sync: '2024-03-08 12:00', gps: 'Disabled', status: 'Active', remarks: 'GPS off' },
  { id: 5, s_no: 5, worker: 'Deepak Verma', app_name: 'Field Worker App', version: '1.0.8', device: 'Android 10', last_sync: '2024-02-28 10:30', gps: 'Disabled', status: 'Inactive', remarks: 'On leave' },
  { id: 6, s_no: 6, worker: 'Ravi Nair', app_name: 'Field Worker App', version: '1.2.0', device: 'Android 13', last_sync: '2024-03-09 13:00', gps: 'Enabled', status: 'Active', remarks: 'Active' },
];

const DEMO_APP_VERSIONS = [
  { id: 1, s_no: 1, app: 'Contractor App', version: '1.2.0', released: '2024-03-01', size: '24.5 MB', downloads: 45, status: 'Active', remarks: 'Stable release' },
  { id: 2, s_no: 2, app: 'Contractor App', version: '1.1.5', released: '2024-01-15', size: '23.8 MB', downloads: 120, status: 'Deprecated', remarks: 'Old version' },
  { id: 3, s_no: 3, app: 'Contractor App', version: '1.0.8', released: '2023-11-10', size: '22.1 MB', downloads: 210, status: 'Deprecated', remarks: 'Very old' },
  { id: 4, s_no: 4, app: 'Field Worker App', version: '1.2.0', released: '2024-03-01', size: '28.2 MB', downloads: 78, status: 'Active', remarks: 'GPS enhanced' },
  { id: 5, s_no: 5, app: 'Field Worker App', version: '1.1.5', released: '2024-01-15', size: '26.5 MB', downloads: 210, status: 'Deprecated', remarks: 'Old version' },
  { id: 6, s_no: 6, app: 'Field Worker App', version: '1.0.8', released: '2023-11-10', size: '25.0 MB', downloads: 340, status: 'Deprecated', remarks: 'Very old' },
];

const DEMO_APP_CRASH_LOGS = [
  { id: 1, s_no: 1, app: 'Contractor App', version: '1.1.5', user: 'SunPower Ltd', crash_type: 'Network Error', occurred: '2024-03-09 14:30', device: 'Android 11', status: 'Resolved', remarks: 'Fixed in 1.2.0' },
  { id: 2, s_no: 2, app: 'Field Worker App', version: '1.1.5', user: 'Amit Singh', crash_type: 'GPS Timeout', occurred: '2024-03-10 11:15', device: 'Android 11', status: 'Pending', remarks: 'Under investigation' },
  { id: 3, s_no: 3, app: 'Field Worker App', version: '1.2.0', user: 'Vikram Sharma', crash_type: 'Camera Error', occurred: '2024-03-11 09:45', device: 'Android 12', status: 'Resolved', remarks: 'User restarted app' },
];

const DEMO_APP_FEEDBACK = [
  { id: 1, s_no: 1, app: 'Contractor App', user: 'SunPower Ltd', rating: '5', feedback: 'Excellent app, easy to use', submitted: '2024-03-10', status: 'Reviewed', remarks: 'Positive' },
  { id: 2, s_no: 2, app: 'Field Worker App', user: 'Rajesh Kumar', rating: '4', feedback: 'GPS works great, but needs offline mode', submitted: '2024-03-09', status: 'Reviewed', remarks: 'Feature request' },
  { id: 3, s_no: 3, app: 'Contractor App', user: 'GreenTech Solutions', rating: '3', feedback: 'Sometimes slow to load reports', submitted: '2024-03-08', status: 'Pending', remarks: 'Performance' },
];

const DEMO_APP_RELEASES = [
  { id: 1, s_no: 1, app: 'Contractor App', version: '1.2.0', release_type: 'Major', released: '2024-03-01', features: 'New dashboard, bug fixes', status: 'Released', remarks: 'Stable' },
  { id: 2, s_no: 2, app: 'Contractor App', version: '1.2.1', release_type: 'Patch', released: '2024-03-15', features: 'Minor bug fixes', status: 'In Testing', remarks: 'QA in progress' },
  { id: 3, s_no: 3, app: 'Field Worker App', version: '1.3.0', release_type: 'Major', released: '2024-04-01', features: 'Offline mode, GPS enhance', status: 'In Development', remarks: 'Dev stage' },
];

// =============================================
// MASTER DATA - Sab form fields ke liye
// =============================================
const MASTER_PROJ_STATUS = [
  { id: 1, s_no: 1, code: 'PEND', name: 'Pending', color: '#f59e0b', status: 'Active', remarks: 'Awaiting review' },
  { id: 2, s_no: 2, code: 'APPR', name: 'Approved', color: '#16a34a', status: 'Active', remarks: 'Approved' },
  { id: 3, s_no: 3, code: 'CORR', name: 'Correction', color: '#ea580c', status: 'Active', remarks: 'Correction' },
  { id: 4, s_no: 4, code: 'COMP', name: 'Completed', color: '#8b5cf6', status: 'Active', remarks: 'Completed' },
  { id: 5, s_no: 5, code: 'REJ', name: 'Rejected', color: '#dc2626', status: 'Active', remarks: 'Rejected' },
];

const MASTER_PROJ_CATEGORY = [
  { id: 1, s_no: 1, code: 'ROOF', name: 'Rooftop Solar', status: 'Active', remarks: 'Rooftop' },
  { id: 2, s_no: 2, code: 'GRND', name: 'Ground Mount', status: 'Active', remarks: 'Ground' },
  { id: 3, s_no: 3, code: 'PUMP', name: 'Solar Pump', status: 'Active', remarks: 'Pump' },
  { id: 4, s_no: 4, code: 'LIGHT', name: 'Street Light', status: 'Active', remarks: 'Lights' },
  { id: 5, s_no: 5, code: 'HEAT', name: 'Solar Water Heater', status: 'Active', remarks: 'Heater' },
];

const MASTER_CAPACITY = [
  { id: 1, s_no: 1, code: '50KW', name: '50 kW', status: 'Active', remarks: 'Small' },
  { id: 2, s_no: 2, code: '75KW', name: '75 kW', status: 'Active', remarks: 'Small' },
  { id: 3, s_no: 3, code: '100KW', name: '100 kW', status: 'Active', remarks: 'Medium' },
  { id: 4, s_no: 4, code: '150KW', name: '150 kW', status: 'Active', remarks: 'Medium' },
  { id: 5, s_no: 5, code: '200KW', name: '200 kW', status: 'Active', remarks: 'Medium' },
  { id: 6, s_no: 6, code: '250KW', name: '250 kW', status: 'Active', remarks: 'Large' },
  { id: 7, s_no: 7, code: '300KW', name: '300 kW', status: 'Active', remarks: 'Large' },
  { id: 8, s_no: 8, code: '500KW', name: '500 kW', status: 'Active', remarks: 'Extra Large' },
];

const MASTER_MILESTONE_TYPE = [
  { id: 1, s_no: 1, code: 'SURV', name: 'Site Survey', status: 'Active', remarks: 'Survey' },
  { id: 2, s_no: 2, code: 'FOUND', name: 'Foundation', status: 'Active', remarks: 'Foundation' },
  { id: 3, s_no: 3, code: 'INST', name: 'Installation', status: 'Active', remarks: 'Installation' },
  { id: 4, s_no: 4, code: 'TEST', name: 'Testing', status: 'Active', remarks: 'Testing' },
  { id: 5, s_no: 5, code: 'COMM', name: 'Commissioning', status: 'Active', remarks: 'Commissioning' },
];

const MASTER_DOC_TYPE = [
  { id: 1, s_no: 1, code: 'QUOT', name: 'Quotation', status: 'Active', remarks: 'Quotation' },
  { id: 2, s_no: 2, code: 'LAYT', name: 'Layout', status: 'Active', remarks: 'Layout' },
  { id: 3, s_no: 3, code: 'LIC', name: 'License', status: 'Active', remarks: 'License' },
  { id: 4, s_no: 4, code: 'APPR', name: 'Approval', status: 'Active', remarks: 'Approval' },
  { id: 5, s_no: 5, code: 'SPEC', name: 'Specification', status: 'Active', remarks: 'Specs' },
];

const MASTER_CONTRACTOR_TYPE = [
  { id: 1, s_no: 1, code: 'SOLR', name: 'Solar Installer', status: 'Active', remarks: 'Solar' },
  { id: 2, s_no: 2, code: 'ELEC', name: 'Electrical Contractor', status: 'Active', remarks: 'Electrical' },
  { id: 3, s_no: 3, code: 'CIVL', name: 'Civil Contractor', status: 'Active', remarks: 'Civil' },
  { id: 4, s_no: 4, code: 'MECH', name: 'Mechanical Contractor', status: 'Active', remarks: 'Mechanical' },
];

const MASTER_LICENSE_TYPE = [
  { id: 1, s_no: 1, code: 'ELEC', name: 'Electrical License', status: 'Active', remarks: 'Electrical' },
  { id: 2, s_no: 2, code: 'CIVL', name: 'Civil License', status: 'Active', remarks: 'Civil' },
  { id: 3, s_no: 3, code: 'SAFE', name: 'Safety License', status: 'Active', remarks: 'Safety' },
];

const MASTER_WORKER_ROLE = [
  { id: 1, s_no: 1, code: 'INST', name: 'Installer', status: 'Active', remarks: 'Installer' },
  { id: 2, s_no: 2, code: 'TECH', name: 'Technician', status: 'Active', remarks: 'Tech' },
  { id: 3, s_no: 3, code: 'SUPV', name: 'Supervisor', status: 'Active', remarks: 'Supervisor' },
  { id: 4, s_no: 4, code: 'HELP', name: 'Helper', status: 'Active', remarks: 'Helper' },
];

const MASTER_WORKER_SKILL = [
  { id: 1, s_no: 1, code: 'PANEL', name: 'Panel Installation', status: 'Active', remarks: 'Panel' },
  { id: 2, s_no: 2, code: 'WIRING', name: 'Wiring', status: 'Active', remarks: 'Wiring' },
  { id: 3, s_no: 3, code: 'MOUNT', name: 'Mounting', status: 'Active', remarks: 'Mount' },
];

const MASTER_REPORT_TYPE = [
  { id: 1, s_no: 1, code: 'DAILY', name: 'Daily Report', status: 'Active', remarks: 'Daily' },
  { id: 2, s_no: 2, code: 'WEEKLY', name: 'Weekly Report', status: 'Active', remarks: 'Weekly' },
  { id: 3, s_no: 3, code: 'FINAL', name: 'Final Report', status: 'Active', remarks: 'Final' },
];

const MASTER_REPORT_PRIORITY = [
  { id: 1, s_no: 1, code: 'LOW', name: 'Low', status: 'Active', remarks: 'Low' },
  { id: 2, s_no: 2, code: 'MED', name: 'Medium', status: 'Active', remarks: 'Medium' },
  { id: 3, s_no: 3, code: 'HIGH', name: 'High', status: 'Active', remarks: 'High' },
];

const MASTER_MEDIA_TYPE = [
  { id: 1, s_no: 1, code: 'PHOTO', name: 'Photograph', status: 'Active', remarks: 'Photo' },
  { id: 2, s_no: 2, code: 'DOC', name: 'Document', status: 'Active', remarks: 'Document' },
];

const MASTER_MEDIA_CATEGORY = [
  { id: 1, s_no: 1, code: 'SITE', name: 'Site Photo', status: 'Active', remarks: 'Site' },
  { id: 2, s_no: 2, code: 'PROG', name: 'Progress Photo', status: 'Active', remarks: 'Progress' },
  { id: 3, s_no: 3, code: 'ISSUE', name: 'Issue Photo', status: 'Active', remarks: 'Issue' },
];

const MASTER_APPROVAL_STEP = [
  { id: 1, s_no: 1, code: 'REG', name: 'Registration', status: 'Active', remarks: 'Reg' },
  { id: 2, s_no: 2, code: 'PROJ', name: 'Project Submission', status: 'Active', remarks: 'Proj' },
  { id: 3, s_no: 3, code: 'REP', name: 'Report Review', status: 'Active', remarks: 'Rep' },
  { id: 4, s_no: 4, code: 'FINAL', name: 'Final Approval', status: 'Active', remarks: 'Final' },
];

const MASTER_USER_ROLE = [
  { id: 1, s_no: 1, code: 'ADMIN', name: 'Super Admin', status: 'Active', remarks: 'Admin' },
  { id: 2, s_no: 2, code: 'CONT', name: 'Contractor', status: 'Active', remarks: 'Cont' },
  { id: 3, s_no: 3, code: 'WORK', name: 'Field Worker', status: 'Active', remarks: 'Work' },
];

const MASTER_APP_NAME = [
  { id: 1, s_no: 1, code: 'CAPP', name: 'Contractor App', status: 'Active', remarks: 'Contractor app' },
  { id: 2, s_no: 2, code: 'WAPP', name: 'Field Worker App', status: 'Active', remarks: 'Worker app' },
];

const MASTER_GPS_STATUS = [
  { id: 1, s_no: 1, code: 'ENB', name: 'Enabled', status: 'Active', remarks: 'Enabled' },
  { id: 2, s_no: 2, code: 'DIS', name: 'Disabled', status: 'Active', remarks: 'Disabled' },
];

const MASTER_STATUS = [
  { id: 1, s_no: 1, code: 'ACT', name: 'Active', status: 'Active', remarks: 'Active' },
  { id: 2, s_no: 2, code: 'INACT', name: 'Inactive', status: 'Active', remarks: 'Inactive' },
];

const MASTER_APPROVAL_STATUS = [
  { id: 1, s_no: 1, code: 'PEND', name: 'Pending', status: 'Active', remarks: 'Pending' },
  { id: 2, s_no: 2, code: 'APPR', name: 'Approved', status: 'Active', remarks: 'Approved' },
  { id: 3, s_no: 3, code: 'REJ', name: 'Rejected', status: 'Active', remarks: 'Rejected' },
];

const MASTER_OS_VERSION = [
  { id: 1, s_no: 1, code: 'A10', name: 'Android 10', status: 'Active', remarks: 'Android 10' },
  { id: 2, s_no: 2, code: 'A11', name: 'Android 11', status: 'Active', remarks: 'Android 11' },
  { id: 3, s_no: 3, code: 'A12', name: 'Android 12', status: 'Active', remarks: 'Android 12' },
  { id: 4, s_no: 4, code: 'A13', name: 'Android 13', status: 'Active', remarks: 'Android 13' },
];

const MASTER_RELEASE_TYPE = [
  { id: 1, s_no: 1, code: 'MAJ', name: 'Major', status: 'Active', remarks: 'Major release' },
  { id: 2, s_no: 2, code: 'MIN', name: 'Minor', status: 'Active', remarks: 'Minor release' },
  { id: 3, s_no: 3, code: 'PATCH', name: 'Patch', status: 'Active', remarks: 'Patch' },
];

const MASTER_CRASH_TYPE = [
  { id: 1, s_no: 1, code: 'NET', name: 'Network Error', status: 'Active', remarks: 'Network' },
  { id: 2, s_no: 2, code: 'GPS', name: 'GPS Timeout', status: 'Active', remarks: 'GPS' },
  { id: 3, s_no: 3, code: 'CAM', name: 'Camera Error', status: 'Active', remarks: 'Camera' },
  { id: 4, s_no: 4, code: 'CRASH', name: 'App Crash', status: 'Active', remarks: 'Crash' },
];

const MASTER_RATING = [
  { id: 1, s_no: 1, code: '5', name: '5 Star', status: 'Active', remarks: 'Excellent' },
  { id: 2, s_no: 2, code: '4', name: '4 Star', status: 'Active', remarks: 'Good' },
  { id: 3, s_no: 3, code: '3', name: '3 Star', status: 'Active', remarks: 'Average' },
  { id: 4, s_no: 4, code: '2', name: '2 Star', status: 'Active', remarks: 'Poor' },
  { id: 5, s_no: 5, code: '1', name: '1 Star', status: 'Active', remarks: 'Bad' },
];

// =============================================
// GLOBAL STYLES
// =============================================
const S = {
  page: { minHeight: '100vh', background: '#f9fafb', padding: '24px', fontFamily: 'system-ui, -apple-system, sans-serif' },
  header: { marginBottom: '24px' },
  title: { fontSize: '26px', fontWeight: 'bold', color: '#111827', margin: 0 },
  subtitle: { fontSize: '13px', color: '#6b7280', marginTop: '4px' },
  statsGrid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '16px', marginBottom: '24px' },
  statCard: { background: '#fff', padding: '20px', borderRadius: '14px', border: '1px solid #f3f4f6', boxShadow: '0 1px 3px rgba(0,0,0,0.04)', position: 'relative', overflow: 'hidden' },
  statAccent: { position: 'absolute', top: 0, left: 0, width: '4px', height: '100%' },
  statLabel: { fontSize: '12px', color: '#6b7280', margin: 0, fontWeight: '500' },
  statValue: { fontSize: '28px', fontWeight: 'bold', margin: '4px 0 0 0', lineHeight: 1 },
  statIconBox: { width: '42px', height: '42px', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  searchBarTop: { background: '#fff', padding: '16px 20px', borderRadius: '12px', border: '1px solid #e5e7eb', marginBottom: '20px', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '12px' },
  searchInput: { width: '100%', padding: '10px 36px', fontSize: '13px', border: '1px solid #d1d5db', borderRadius: '10px', outline: 'none', background: '#fff', color: '#374151' },
  searchIcon: { position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af', pointerEvents: 'none' },
  clearIcon: { position: 'absolute', right: '12px', top: '50%', transform: 'translateY(-50%)', color: '#9ca3af', cursor: 'pointer', background: 'transparent', border: 'none' },
  btn: { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '10px 16px', fontSize: '13px', fontWeight: '500', borderRadius: '10px', border: '1px solid #d1d5db', background: '#fff', color: '#374151', cursor: 'pointer' },
  btnActive: { background: '#eff6ff', color: '#1d4ed8', borderColor: '#93c5fd' },
  btnPrimary: { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '10px 16px', fontSize: '13px', fontWeight: '600', borderRadius: '10px', border: 'none', background: '#2563eb', color: '#fff', cursor: 'pointer' },
  btnPurple: { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '10px 16px', fontSize: '13px', fontWeight: '600', borderRadius: '10px', border: 'none', background: '#7c3aed', color: '#fff', cursor: 'pointer' },
  btnDanger: { padding: '10px 16px', fontSize: '13px', color: '#dc2626', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: '10px', cursor: 'pointer' },
  filterPanel: { marginBottom: '24px', padding: '20px', border: '1px solid #e5e7eb', borderRadius: '16px', background: '#fff' },
  filterGrid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' },
  filterLabel: { display: 'block', fontSize: '11px', color: '#6b7280', marginBottom: '4px', fontWeight: '600' },
  filterInput: { width: '100%', padding: '10px 14px', fontSize: '13px', border: '1px solid #d1d5db', borderRadius: '10px', outline: 'none', background: '#fff', color: '#374151' },
  tableCard: { background: '#fff', borderRadius: '16px', border: '1px solid #f3f4f6', overflow: 'hidden' },
  tableHeaderBar: { padding: '12px 20px', borderBottom: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '12px' },
  tableTitle: { fontSize: '15px', fontWeight: '600', color: '#374151', margin: 0 },
  countBadge: { fontSize: '11px', color: '#6b7280', background: '#e5e7eb', padding: '2px 8px', borderRadius: '12px' },
  tableWrap: { overflowX: 'auto' },
  table: { width: '100%', borderCollapse: 'collapse', minWidth: '900px' },
  thead: { background: '#f9fafb' },
  th: { padding: '12px 16px', textAlign: 'left', fontSize: '11px', fontWeight: '600', color: '#6b7280', textTransform: 'uppercase', letterSpacing: '0.5px', borderBottom: '1px solid #e5e7eb', whiteSpace: 'nowrap' },
  td: { padding: '12px 16px', fontSize: '13px', color: '#374151', borderBottom: '1px solid #f3f4f6', whiteSpace: 'nowrap' },
  tdStrong: { padding: '12px 16px', fontSize: '13px', color: '#1f2937', fontWeight: '600', borderBottom: '1px solid #f3f4f6', whiteSpace: 'nowrap' },
  iconBtn: { padding: '6px', borderRadius: '8px', border: 'none', background: 'transparent', cursor: 'pointer', display: 'inline-flex', alignItems: 'center' },
  pagination: { padding: '16px 20px', borderTop: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'center', gap: '12px' },
  pageBtn: { display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '6px 12px', fontSize: '13px', fontWeight: '500', color: '#374151', background: '#fff', border: '1px solid #d1d5db', borderRadius: '8px', cursor: 'pointer' },
  pageBtnActive: { background: '#2563eb', color: '#fff', borderColor: '#2563eb' },
  pageBtnDisabled: { opacity: 0.5, cursor: 'not-allowed' },
  emptyState: { textAlign: 'center', padding: '48px 20px', background: '#fff', borderRadius: '16px', border: '1px solid #f3f4f6' },
  statusBadge: { display: 'inline-block', padding: '4px 10px', borderRadius: '12px', fontSize: '11px', fontWeight: '600' },
  modalOverlay: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '16px', zIndex: 50 },
  modal: { background: '#fff', borderRadius: '16px', width: '100%', maxWidth: '800px', maxHeight: '90vh', overflowY: 'auto' },
  modalHeader: { padding: '20px 24px', borderBottom: '1px solid #e5e7eb', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' },
  modalTitle: { fontSize: '18px', fontWeight: 'bold', color: '#111827', margin: 0 },
  modalBody: { padding: '24px' },
  modalFooter: { padding: '20px 24px', borderTop: '1px solid #e5e7eb', background: '#f9fafb', display: 'flex', justifyContent: 'flex-end', gap: '12px' },
  formGrid3: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px 24px' },
  formGroup: { display: 'flex', flexDirection: 'column', gap: '8px' },
  formGroupFull: { display: 'flex', flexDirection: 'column', gap: '8px', gridColumn: '1 / -1' },
  formLabel: { fontSize: '12px', fontWeight: '600', color: '#374151', display: 'flex', alignItems: 'center', gap: '6px' },
  formInput: { width: '100%', padding: '11px 14px', fontSize: '13px', border: '1px solid #d1d5db', borderRadius: '10px', outline: 'none', color: '#374151', background: '#fff', boxSizing: 'border-box' },
  formTextarea: { width: '100%', padding: '11px 14px', fontSize: '13px', border: '1px solid #d1d5db', borderRadius: '10px', outline: 'none', color: '#374151', resize: 'vertical', fontFamily: 'inherit', boxSizing: 'border-box' },
  btnSecondary: { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '10px 20px', fontSize: '13px', fontWeight: '500', color: '#374151', background: '#fff', border: '1px solid #d1d5db', borderRadius: '10px', cursor: 'pointer' },
  btnSubmit: { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '10px 20px', fontSize: '13px', fontWeight: '600', color: '#fff', background: '#2563eb', border: '1px solid #2563eb', borderRadius: '10px', cursor: 'pointer' },
  photoGrid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '16px' },
  photoCard: { background: '#fff', borderRadius: '12px', overflow: 'hidden', border: '1px solid #e5e7eb' },
  photoThumb: { height: '140px', background: 'linear-gradient(135deg, #dbeafe, #bfdbfe)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#2563eb' },
  photoInfo: { padding: '10px 12px' },
  card: { background: '#fff', borderRadius: '16px', border: '1px solid #f3f4f6', padding: '20px', boxShadow: '0 1px 3px rgba(0,0,0,0.04)' },
  cardHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' },
  cardTitle: { fontSize: '15px', fontWeight: '600', color: '#1f2937', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' },
  progressTrack: { width: '100%', height: '8px', background: '#e5e7eb', borderRadius: '4px', overflow: 'hidden', marginTop: '6px' },
  progressFill: { height: '100%', borderRadius: '4px', transition: 'width 0.3s' },
};

// =============================================
// HELPERS
// =============================================
const getStatusStyle = (status) => {
  const map = {
    'Approved': { background: '#dcfce7', color: '#166534' },
    'Completed': { background: '#dcfce7', color: '#166534' },
    'Active': { background: '#dcfce7', color: '#166534' },
    'Enabled': { background: '#dcfce7', color: '#166534' },
    'Resolved': { background: '#dcfce7', color: '#166534' },
    'Released': { background: '#dcfce7', color: '#166534' },
    'Reviewed': { background: '#dcfce7', color: '#166534' },
    'Pending': { background: '#fef3c7', color: '#854d0e' },
    'In Progress': { background: '#dbeafe', color: '#1e40af' },
    'In Testing': { background: '#dbeafe', color: '#1e40af' },
    'In Development': { background: '#dbeafe', color: '#1e40af' },
    'Correction': { background: '#fed7aa', color: '#9a3412' },
    'Rejected': { background: '#fee2e2', color: '#991b1b' },
    'Deprecated': { background: '#fee2e2', color: '#991b1b' },
    'Disabled': { background: '#fee2e2', color: '#991b1b' },
    'Inactive': { background: '#f3f4f6', color: '#4b5563' },
  };
  return map[status] || { background: '#f3f4f6', color: '#4b5563' };
};

const formatDate = (d) => d ? dayjs(d).format('DD/MM/YYYY') : '—';
const getNames = (arr) => arr.filter(a => a.status === 'Active').map(a => a.name);

// =============================================
// LOGIN
// =============================================
const Login = ({ onLogin }) => {
  const [email, setEmail] = useState('admin@iocl.com');
  const [password, setPassword] = useState('admin123');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const submit = (e) => {
    e.preventDefault();
    setLoading(true); setError('');
    setTimeout(() => {
      if (email === 'admin@iocl.com' && password === 'admin123') {
        const u = { id: 1, name: 'Admin User', email, role: 'Super Admin' };
        localStorage.setItem('authToken', 'demo-token-' + Date.now());
        localStorage.setItem('user', JSON.stringify(u));
        onLogin(u);
      } else setError('Invalid credentials');
      setLoading(false);
    }, 500);
  };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #1e3a8a, #3b82f6)', padding: '20px', fontFamily: 'system-ui, sans-serif' }}>
      <div style={{ background: '#fff', padding: '40px', borderRadius: '16px', boxShadow: '0 20px 60px rgba(0,0,0,0.3)', width: '100%', maxWidth: '420px' }}>
        <div style={{ width: '64px', height: '64px', background: '#dbeafe', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px', color: '#2563eb' }}><Shield size={32} /></div>
        <h1 style={{ fontSize: '24px', fontWeight: 'bold', textAlign: 'center', color: '#1e3a8a', marginBottom: '4px' }}>IOCL ERP Portal</h1>
        <p style={{ fontSize: '13px', textAlign: 'center', color: '#6b7280', marginBottom: '24px' }}>Solar Project Management</p>
        {error && <div style={{ background: '#fee2e2', color: '#991b1b', padding: '10px 14px', borderRadius: '8px', fontSize: '13px', marginBottom: '16px' }}>{error}</div>}
        <form onSubmit={submit} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div style={S.formGroup}><label style={S.formLabel}><Mail size={12} /> Email</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} style={S.formInput} required />
          </div>
          <div style={S.formGroup}><label style={S.formLabel}><Lock size={12} /> Password</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={S.formInput} required />
          </div>
          <button type="submit" disabled={loading} style={{ background: '#2563eb', color: '#fff', padding: '12px', border: 'none', borderRadius: '10px', fontSize: '15px', fontWeight: '600', cursor: 'pointer', marginTop: '8px' }}>
            {loading ? 'Logging in...' : 'Secure Login'}
          </button>
        </form>
        <div style={{ marginTop: '20px', textAlign: 'center', fontSize: '12px', color: '#6b7280', background: '#f3f4f6', padding: '10px', borderRadius: '8px' }}>
          <strong>Demo:</strong> admin@iocl.com / admin123
        </div>
      </div>
    </div>
  );
};

// =============================================
// HOOKS + PAGINATION
// =============================================
const useTablePage = (initialData, searchKeys) => {
  const [allData, setAllData] = useState([...initialData]);
  const [displayData, setDisplayData] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const [totalItems, setTotalItems] = useState(0);
  const [itemsPerPage, setItemsPerPage] = useState(5);
  const [filters, setFilters] = useState({ status: '' });

  const filterAndPaginate = useCallback(() => {
    let f = [...allData];
    if (searchTerm) {
      const t = searchTerm.toLowerCase();
      f = f.filter(i => searchKeys.some(k => String(i[k] || '').toLowerCase().includes(t)));
    }
    if (filters.status) f = f.filter(i => i.status === filters.status);
    setTotalItems(f.length);
    const pages = Math.ceil(f.length / itemsPerPage) || 1;
    setTotalPages(pages);
    if (currentPage > pages) setCurrentPage(pages);
    const s = (currentPage - 1) * itemsPerPage;
    setDisplayData(f.slice(s, s + itemsPerPage));
  }, [allData, searchTerm, filters, currentPage, itemsPerPage, searchKeys]);

  useEffect(() => { filterAndPaginate(); }, [filterAndPaginate]);

  return {
    allData, setAllData: (n) => setAllData([...n]), displayData, searchTerm, setSearchTerm,
    currentPage, setCurrentPage, totalPages, totalItems, itemsPerPage, setItemsPerPage,
    filters, setFilters
  };
};

const Pagination = ({ currentPage, totalPages, totalItems, itemsPerPage, onPageChange }) => {
  if (totalPages <= 1) return null;
  const pages = [];
  let start = Math.max(1, currentPage - 2);
  let end = Math.min(totalPages, start + 4);
  if (end - start < 4) start = Math.max(1, end - 4);
  for (let i = start; i <= end; i++) pages.push(i);

  return (
    <div style={S.pagination}>
      <div style={{ fontSize: '13px', color: '#4b5563' }}>
        Showing <b>{totalItems === 0 ? 0 : ((currentPage - 1) * itemsPerPage) + 1}</b> to <b>{Math.min(currentPage * itemsPerPage, totalItems)}</b> of <b>{totalItems}</b>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
        <button onClick={() => onPageChange(currentPage - 1)} disabled={currentPage === 1} style={{ ...S.pageBtn, ...(currentPage === 1 ? S.pageBtnDisabled : {}) }}><ChevronLeft size={14} /> Prev</button>
        {pages.map(p => <button key={p} onClick={() => onPageChange(p)} style={{ ...S.pageBtn, ...(currentPage === p ? S.pageBtnActive : {}) }}>{p}</button>)}
        <button onClick={() => onPageChange(currentPage + 1)} disabled={currentPage === totalPages} style={{ ...S.pageBtn, ...(currentPage === totalPages ? S.pageBtnDisabled : {}) }}>Next <ChevronRight size={14} /></button>
      </div>
    </div>
  );
};

const StatsRow = ({ stats }) => (
  <div style={S.statsGrid}>
    {stats.map((s, i) => (
      <div key={i} style={S.statCard}>
        <div style={{ ...S.statAccent, background: s.color || '#3b82f6' }} />
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
          <div>
            <p style={S.statLabel}>{s.label}</p>
            <p style={{ ...S.statValue, color: s.color || '#1f2937' }}>{s.value}</p>
          </div>
          <div style={{ ...S.statIconBox, background: (s.color || '#3b82f6') + '15', color: s.color || '#3b82f6' }}>{s.icon}</div>
        </div>
      </div>
    ))}
  </div>
);

// =============================================
// EXPORT / PRINT
// =============================================
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

const printReport = (data, columns, title) => {
  try {
    if (!data.length) { toast.error('No data'); return; }
    const w = window.open('', '_blank', 'width=1200,height=800');
    if (!w) { toast.error('Allow popups'); return; }
    let html = `<!DOCTYPE html><html><head><title>${title}</title><style>
      * { margin:0; padding:0; box-sizing:border-box; }
      body { font-family: Arial, sans-serif; padding: 10px; background: #fff; font-size: 11px; }
      .container { max-width: 1100px; margin: 0 auto; border: 2px solid #000; }
      .header { text-align: center; padding: 12px; border-bottom: 2px solid #000; }
      .header h1 { font-size: 16px; font-weight: bold; }
      .header h2 { font-size: 13px; font-weight: bold; margin-top: 4px; }
      .header p { font-size: 11px; margin-top: 6px; text-decoration: underline; }
      .meta { padding: 6px 12px; border-bottom: 1px solid #000; font-size: 10px; display: flex; justify-content: space-between; }
      table { width: 100%; border-collapse: collapse; font-size: 10px; }
      th, td { border: 1px solid #000; padding: 5px 8px; text-align: left; }
      th { background: #e0e0e0; font-weight: bold; text-transform: uppercase; }
      .footer { text-align: center; padding: 8px; font-size: 9px; color: #666; border-top: 1px solid #000; }
      @media print { th { background: #e0e0e0 !important; } }
    </style></head><body>
    <div class="container"><div class="header"><h1>INDIAN OIL CORPORATION LIMITED</h1><h2>${title.toUpperCase()}</h2><p>System Generated</p></div>
      <div class="meta"><span><b>Generated:</b> ${dayjs().format('DD/MM/YYYY HH:mm')}</span><span><b>Total:</b> ${data.length}</span></div>
      <table><thead><tr><th style="width:50px;">S/No.</th>${columns.map(c => `<th>${c.label}</th>`).join('')}</tr></thead><tbody>`;
    data.forEach((item, idx) => {
      html += `<tr><td>${idx + 1}</td>${columns.map(c => `<td>${c.exportValue ? c.exportValue(item) : (item[c.key] ?? '-')}</td>`).join('')}</tr>`;
    });
    html += `</tbody></table><div class="footer">System generated report</div></div>
    <script>window.onload=function(){window.print();window.onafterprint=function(){window.close();}}<\/script></body></html>`;
    w.document.write(html); w.document.close();
    toast.success(`Printing ${data.length}`);
  } catch (e) { toast.error('Failed'); }
};

// =============================================
// FORM FIELD (with dropdown)
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

// =============================================
// TABLE PAGE
// =============================================
const TablePage = ({ title, subtitle, columns, initialData, searchKeys, statusOptions, renderForm, formInitial, validateForm, stats, addLabel }) => {
  const { allData, setAllData, displayData, searchTerm, setSearchTerm, currentPage, setCurrentPage, totalPages, totalItems, itemsPerPage, setItemsPerPage, filters, setFilters } = useTablePage(initialData, searchKeys);
  const [showFilters, setShowFilters] = useState(false);
  const [showView, setShowView] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState(formInitial);

  const activeFilters = !!(filters.status || searchTerm);
  const resetForm = () => { setFormData(formInitial); setEditingId(null); };
  const openNew = () => { resetForm(); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const openView = (item) => { setSelected(item); setShowView(true); };

  const handleDelete = (item) => {
    if (!window.confirm('Delete this record?')) return;
    setAllData(allData.filter(d => d.id !== item.id));
    toast.success('Deleted'); setShowView(false); setSelected(null);
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    const error = validateForm ? validateForm(formData) : null;
    if (error) { toast.error(error); return; }
    if (editingId) {
      setAllData(allData.map(d => d.id === editingId ? { ...d, ...formData } : d));
      toast.success('Updated');
    } else {
      const nextSNo = allData.length > 0 ? Math.max(...allData.map(d => d.s_no || 0)) + 1 : 1;
      setAllData([{ ...formData, id: Date.now(), s_no: nextSNo }, ...allData]);
      toast.success('Added');
    }
    setShowForm(false); resetForm();
  };

  const clearFilters = () => { setFilters({ status: '' }); setSearchTerm(''); setCurrentPage(1); };
  const refresh = () => { setAllData([...initialData]); setCurrentPage(1); toast.success('Refreshed'); };

  if (showForm) {
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div><h1 style={S.title}>{editingId ? `Edit ${title}` : `New ${title}`}</h1><p style={S.subtitle}>Fill the details below</p></div>
            <button onClick={() => { setShowForm(false); resetForm(); }} style={S.btnSecondary}><X size={14} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '16px', border: '1px solid #f3f4f6', padding: '28px' }}>
            <form onSubmit={handleSubmit}>
              <div style={S.formGrid3}>{renderForm(formData, setFormData, statusOptions)}</div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', paddingTop: '24px', marginTop: '28px', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => { setShowForm(false); resetForm(); }} style={S.btnSecondary}><X size={14} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={14} /> {editingId ? 'Update' : 'Submit'}</button>
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
      <div style={S.header}><h1 style={S.title}>{title}</h1><p style={S.subtitle}>{subtitle}</p></div>
      {stats && <StatsRow stats={stats} />}

      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 340px', maxWidth: '520px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder={`Search ${title}...`} value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setCurrentPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => { setSearchTerm(''); setCurrentPage(1); }} style={S.clearIcon}><X size={14} /></button>}
        </div>
        <button onClick={refresh} style={S.btn}><RefreshCw size={14} /> Refresh</button>
        <button onClick={() => setShowFilters(!showFilters)} style={{ ...S.btn, ...(showFilters ? S.btnActive : {}) }}><Filter size={14} /> {showFilters ? 'Hide' : 'Filters'}</button>
        <button onClick={() => exportToExcel(allData, columns, title.toLowerCase().replace(/\s+/g, '_'))} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(allData, columns, title)} style={S.btn}><Printer size={14} /> Print</button>
        {activeFilters && <button onClick={clearFilters} style={S.btnDanger}>Clear</button>}
        <button onClick={openNew} style={S.btnPrimary}><Plus size={14} /> {addLabel || `Add`}</button>
      </div>

      {showFilters && (
        <div style={S.filterPanel}>
          <div style={S.filterGrid}>
            <div><label style={S.filterLabel}>Status</label>
              <select value={filters.status} onChange={e => { setFilters({ ...filters, status: e.target.value }); setCurrentPage(1); }} style={S.filterInput}>
                <option value="">All Status</option>
                {statusOptions.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
        </div>
      )}

      {displayData.length === 0 ? (
        <div style={S.emptyState}>
          <div style={{ width: '64px', height: '64px', background: '#dbeafe', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px', color: '#2563eb' }}><FileText size={28} /></div>
          <h3 style={{ fontSize: '15px', color: '#374151', margin: '0 0 8px 0' }}>No records</h3>
          <button onClick={openNew} style={S.btnPrimary}><Plus size={14} /> Add</button>
        </div>
      ) : (
        <div style={S.tableCard}>
          <div style={S.tableHeaderBar}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <h2 style={S.tableTitle}>All {title}</h2>
              <span style={S.countBadge}>{totalItems}</span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '11px', color: '#6b7280' }}>Show:</span>
              <select value={itemsPerPage} onChange={e => { setItemsPerPage(Number(e.target.value)); setCurrentPage(1); }} style={{ fontSize: '12px', border: '1px solid #d1d5db', borderRadius: '6px', padding: '4px 6px', background: '#fff' }}>
                <option value="5">5</option><option value="10">10</option><option value="25">25</option><option value="50">50</option>
              </select>
            </div>
          </div>
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead style={S.thead}>
                <tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr>
              </thead>
              <tbody>
                {displayData.map((item, idx) => {
                  const seq = ((currentPage - 1) * itemsPerPage) + idx + 1;
                  return (
                    <tr key={item.id}>
                      <td style={S.td}>{item.s_no || seq}</td>
                      {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(item) : (item[c.key] ?? '—')}</td>)}
                      <td style={S.td}>
                        <div style={{ display: 'flex', gap: '2px' }}>
                          <button onClick={() => openView(item)} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={15} /></button>
                          <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={15} /></button>
                          <button onClick={() => handleDelete(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={15} /></button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <Pagination currentPage={currentPage} totalPages={totalPages} totalItems={totalItems} itemsPerPage={itemsPerPage} onPageChange={setCurrentPage} />
        </div>
      )}

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => { setShowView(false); setSelected(null); }}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}>
              <div><h2 style={S.modalTitle}>{title} Details</h2></div>
              <button onClick={() => { setShowView(false); setSelected(null); }} style={{ ...S.iconBtn, color: '#6b7280' }}><X size={18} /></button>
            </div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px' }}>
                {columns.map(c => (
                  <div key={c.key}>
                    <p style={{ fontSize: '11px', color: '#6b7280', margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>{c.label}</p>
                    <p style={{ fontSize: '14px', color: '#111827', margin: 0 }}>{c.render ? c.render(selected) : (selected[c.key] ?? '—')}</p>
                  </div>
                ))}
              </div>
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => { setShowView(false); setSelected(null); }} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={14} /> Edit</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// SIMPLE MASTER PAGE
// =============================================
const SimpleMasterPage = ({ title, subtitle, storageKey, defaultColumns, defaultData }) => {
  const [columns, setColumns] = useState(() => {
    const s = localStorage.getItem(storageKey + '_cols');
    return s ? JSON.parse(s) : defaultColumns;
  });
  const [data, setData] = useState(() => {
    const s = localStorage.getItem(storageKey + '_data');
    return s ? JSON.parse(s) : defaultData;
  });
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [showColForm, setShowColForm] = useState(false);
  const [showView, setShowView] = useState(false);
  const [selected, setSelected] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState({});
  const [newCol, setNewCol] = useState({ key: '', label: '' });
  const [page, setPage] = useState(1);
  const perPage = 8;

  useEffect(() => { localStorage.setItem(storageKey + '_cols', JSON.stringify(columns)); }, [columns, storageKey]);
  useEffect(() => { localStorage.setItem(storageKey + '_data', JSON.stringify(data)); }, [data, storageKey]);

  const filtered = data.filter(item => !searchTerm || Object.values(item).some(v => String(v || '').toLowerCase().includes(searchTerm.toLowerCase())));
  const totalPages = Math.ceil(filtered.length / perPage) || 1;
  const start = (page - 1) * perPage;
  const display = filtered.slice(start, start + perPage);

  const openNew = () => { const e = {}; columns.forEach(c => e[c.key] = ''); setFormData(e); setEditingId(null); setShowForm(true); };
  const openEdit = (item) => { setFormData({ ...item }); setEditingId(item.id); setShowForm(true); setShowView(false); };
  const del = (item) => { if (!window.confirm('Delete?')) return; setData(data.filter(d => d.id !== item.id)); toast.success('Deleted'); setShowView(false); };
  const submit = (e) => {
    e.preventDefault();
    if (editingId) { setData(data.map(d => d.id === editingId ? { ...d, ...formData } : d)); toast.success('Updated'); }
    else { const n = data.length > 0 ? Math.max(...data.map(d => d.s_no || 0)) + 1 : 1; setData([{ ...formData, id: Date.now(), s_no: n }, ...data]); toast.success('Added'); }
    setShowForm(false);
  };
  const addCol = (e) => {
    e.preventDefault();
    if (!newCol.key || !newCol.label) { toast.error('Both required'); return; }
    if (columns.some(c => c.key === newCol.key)) { toast.error('Key exists'); return; }
    setColumns([...columns, { key: newCol.key, label: newCol.label, dynamic: true }]);
    setNewCol({ key: '', label: '' }); setShowColForm(false); toast.success('Column added');
  };
  const rmCol = (key) => { if (!window.confirm('Remove column?')) return; setColumns(columns.filter(c => c.key !== key)); toast.success('Removed'); };

  if (showForm) {
    return (
      <div style={S.page}>
        <ToastContainer position="top-right" autoClose={3000} />
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div><h1 style={S.title}>{editingId ? `Edit ${title}` : `New ${title}`}</h1><p style={S.subtitle}>Fill below</p></div>
            <button onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={14} /> Cancel</button>
          </div>
          <div style={{ background: '#fff', borderRadius: '16px', padding: '28px', border: '1px solid #f3f4f6' }}>
            <form onSubmit={submit}>
              <div style={S.formGrid3}>
                {columns.map(col => <Field key={col.key} label={col.label} name={col.key} value={formData[col.key] || ''} onChange={e => setFormData({ ...formData, [col.key]: e.target.value })} placeholder={`Enter ${col.label}`} />)}
              </div>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', paddingTop: '24px', marginTop: '28px', borderTop: '1px solid #e5e7eb' }}>
                <button type="button" onClick={() => setShowForm(false)} style={S.btnSecondary}><X size={14} /> Cancel</button>
                <button type="submit" style={S.btnSubmit}><CheckCircle size={14} /> {editingId ? 'Update' : 'Submit'}</button>
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
      <div style={S.header}><h1 style={S.title}>{title}</h1><p style={S.subtitle}>{subtitle}</p></div>

      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 300px', maxWidth: '500px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder={`Search ${title}...`} value={searchTerm} onChange={e => { setSearchTerm(e.target.value); setPage(1); }} style={S.searchInput} />
          {searchTerm && <button onClick={() => setSearchTerm('')} style={S.clearIcon}><X size={14} /></button>}
        </div>
        <button onClick={openNew} style={S.btnPrimary}><Plus size={14} /> Add</button>
        <button onClick={() => setShowColForm(true)} style={S.btnPurple}><Columns size={14} /> Add Column</button>
        <button onClick={() => exportToExcel(data, columns, title.toLowerCase().replace(/\s+/g, '_'))} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(data, columns, title)} style={S.btn}><Printer size={14} /> Print</button>
      </div>

      {showColForm && (
        <div style={S.modalOverlay} onClick={() => setShowColForm(false)}>
          <div style={{ ...S.modal, maxWidth: '480px' }} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}>
              <div><h2 style={S.modalTitle}>Add Column</h2><p style={{ fontSize: '12px', color: '#6b7280', marginTop: '4px' }}>Add custom field</p></div>
              <button onClick={() => setShowColForm(false)} style={{ ...S.iconBtn, color: '#6b7280' }}><X size={18} /></button>
            </div>
            <form onSubmit={addCol}>
              <div style={S.modalBody}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                  <Field label="Key" icon={<Hash size={12} />} name="key" value={newCol.key} onChange={e => setNewCol({ ...newCol, key: e.target.value })} placeholder="e.g. manager_name" required />
                  <Field label="Label" icon={<Tag size={12} />} name="label" value={newCol.label} onChange={e => setNewCol({ ...newCol, label: e.target.value })} placeholder="e.g. Manager Name" required />
                </div>
              </div>
              <div style={S.modalFooter}>
                <button type="button" onClick={() => setShowColForm(false)} style={S.btnSecondary}>Cancel</button>
                <button type="submit" style={S.btnSubmit}><Plus size={14} /> Add</button>
              </div>
            </form>
          </div>
        </div>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '16px', marginBottom: '24px' }}>
        {display.map(item => (
          <div key={item.id} style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: '12px', padding: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
              <div>
                <p style={{ fontSize: '12px', color: '#6b7280', margin: '0 0 2px 0' }}>S/No. {item.s_no}</p>
                <p style={{ fontSize: '15px', fontWeight: '600', color: '#111827', margin: 0 }}>{item.name || item[columns[0]?.key] || 'Record'}</p>
              </div>
              <div style={{ display: 'flex', gap: '2px' }}>
                <button onClick={() => { setSelected(item); setShowView(true); }} style={{ ...S.iconBtn, color: '#2563eb' }}><Eye size={15} /></button>
                <button onClick={() => openEdit(item)} style={{ ...S.iconBtn, color: '#7c3aed' }}><Edit size={15} /></button>
                <button onClick={() => del(item)} style={{ ...S.iconBtn, color: '#dc2626' }}><Trash2 size={15} /></button>
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
              {columns.map(col => (
                <div key={col.key} style={{ background: '#f9fafb', padding: '8px 10px', borderRadius: '8px', position: 'relative' }}>
                  <p style={{ fontSize: '10px', color: '#6b7280', margin: '0 0 2px 0', fontWeight: '600', textTransform: 'uppercase' }}>{col.label}</p>
                  <p style={{ fontSize: '12px', color: '#111827', margin: 0, fontWeight: '500' }}>
                    {col.key === 'status' ? <span style={{ ...S.statusBadge, ...getStatusStyle(item[col.key]) }}>{item[col.key] || '—'}</span> : (item[col.key] || '—')}
                  </p>
                  {col.dynamic && <button onClick={() => rmCol(col.key)} style={{ position: 'absolute', top: '2px', right: '2px', background: 'transparent', border: 'none', color: '#dc2626', cursor: 'pointer' }}><X size={10} /></button>}
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>

      {totalPages > 1 && (
        <div style={{ ...S.tableCard, padding: '16px 20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ fontSize: '13px', color: '#4b5563' }}>Showing <b>{start + 1}</b> to <b>{Math.min(start + perPage, filtered.length)}</b> of <b>{filtered.length}</b></div>
            <div style={{ display: 'flex', gap: '6px' }}>
              <button onClick={() => setPage(page - 1)} disabled={page === 1} style={{ ...S.pageBtn, ...(page === 1 ? S.pageBtnDisabled : {}) }}><ChevronLeft size={14} /> Prev</button>
              {Array.from({ length: totalPages }, (_, i) => i + 1).map(p => <button key={p} onClick={() => setPage(p)} style={{ ...S.pageBtn, ...(page === p ? S.pageBtnActive : {}) }}>{p}</button>)}
              <button onClick={() => setPage(page + 1)} disabled={page === totalPages} style={{ ...S.pageBtn, ...(page === totalPages ? S.pageBtnDisabled : {}) }}>Next <ChevronRight size={14} /></button>
            </div>
          </div>
        </div>
      )}

      {filtered.length === 0 && (
        <div style={S.emptyState}>
          <div style={{ width: '64px', height: '64px', background: '#dbeafe', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px', color: '#2563eb' }}><FileText size={28} /></div>
          <h3 style={{ fontSize: '15px', color: '#374151', margin: '0 0 8px 0' }}>No records</h3>
          <button onClick={openNew} style={S.btnPrimary}><Plus size={14} /> Add</button>
        </div>
      )}

      {showView && selected && (
        <div style={S.modalOverlay} onClick={() => setShowView(false)}>
          <div style={S.modal} onClick={e => e.stopPropagation()}>
            <div style={S.modalHeader}>
              <div><h2 style={S.modalTitle}>{title} Details</h2></div>
              <button onClick={() => setShowView(false)} style={{ ...S.iconBtn, color: '#6b7280' }}><X size={18} /></button>
            </div>
            <div style={S.modalBody}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                {columns.map(col => (
                  <div key={col.key}>
                    <p style={{ fontSize: '11px', color: '#6b7280', margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>{col.label}</p>
                    <p style={{ fontSize: '14px', color: '#111827', margin: 0 }}>
                      {col.key === 'status' ? <span style={{ ...S.statusBadge, ...getStatusStyle(selected[col.key]) }}>{selected[col.key] || '—'}</span> : (selected[col.key] || '—')}
                    </p>
                  </div>
                ))}
              </div>
            </div>
            <div style={S.modalFooter}>
              <button onClick={() => setShowView(false)} style={S.btnSecondary}>Close</button>
              <button onClick={() => openEdit(selected)} style={S.btnSubmit}><Edit size={14} /> Edit</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// =============================================
// PAGE WRAPPERS WITH DROPDOWNS
// =============================================

// Projects
const ProjectListPage = () => {
  const columns = [
    { key: 'site_code', label: 'Site Code', strong: true },
    { key: 'name', label: 'Project Name' },
    { key: 'category', label: 'Category' },
    { key: 'capacity', label: 'Capacity' },
    { key: 'location', label: 'Location' },
    { key: 'contractor', label: 'Contractor' },
    { key: 'field_worker', label: 'Field Worker' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', site_code: '', name: '', category: '', capacity: '', location: '', contractor: '', field_worker: '', progress: '', status: 'Pending', start_date: '', end_date: '', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Site Code" icon={<Hash size={12} />} name="site_code" value={fd.site_code} onChange={onChange} placeholder="SP-XXX-001" required />
        <Field label="Project Name" icon={<FileText size={12} />} name="name" value={fd.name} onChange={onChange} required />
        <Field label="Category" icon={<Tag size={12} />} name="category" value={fd.category} onChange={onChange} options={getNames(MASTER_PROJ_CATEGORY)} />
        <Field label="Capacity" icon={<Gauge size={12} />} name="capacity" value={fd.capacity} onChange={onChange} options={getNames(MASTER_CAPACITY)} />
        <Field label="Location" icon={<Globe size={12} />} name="location" value={fd.location} onChange={onChange} />
        <Field label="Contractor" icon={<Building2 size={12} />} name="contractor" value={fd.contractor} onChange={onChange} options={DEMO_CONTRACTORS.map(c => c.name)} />
        <Field label="Field Worker" icon={<User size={12} />} name="field_worker" value={fd.field_worker} onChange={onChange} options={DEMO_FIELD_WORKERS.map(w => w.name)} />
        <Field label="Progress (%)" icon={<TrendingUp size={12} />} name="progress" type="number" value={fd.progress} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={getNames(MASTER_PROJ_STATUS)} />
        <Field label="Start Date" icon={<Calendar size={12} />} name="start_date" type="date" value={fd.start_date} onChange={onChange} />
        <Field label="End Date" icon={<Calendar size={12} />} name="end_date" type="date" value={fd.end_date} onChange={onChange} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="All Projects" subtitle="Manage solar projects" columns={columns} initialData={DEMO_PROJECTS} searchKeys={['name', 'site_code', 'location']} statusOptions={getNames(MASTER_PROJ_STATUS)} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.name || !fd.site_code) ? 'Required' : null} addLabel="Add Project" />;
};

const ProjectSitesPage = () => {
  const columns = [
    { key: 'site_code', label: 'Site Code', strong: true },
    { key: 'project', label: 'Project' },
    { key: 'category', label: 'Category' },
    { key: 'location', label: 'Location' },
    { key: 'area', label: 'Area' },
    { key: 'gps', label: 'GPS' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', site_code: '', project: '', category: '', location: '', area: '', gps: '', status: 'Active', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Site Code" icon={<Hash size={12} />} name="site_code" value={fd.site_code} onChange={onChange} required />
        <Field label="Project" icon={<FolderOpen size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Category" icon={<Tag size={12} />} name="category" value={fd.category} onChange={onChange} options={getNames(MASTER_PROJ_CATEGORY)} />
        <Field label="Location" icon={<Globe size={12} />} name="location" value={fd.location} onChange={onChange} />
        <Field label="Area" icon={<Gauge size={12} />} name="area" value={fd.area} onChange={onChange} />
        <Field label="GPS" icon={<MapPin size={12} />} name="gps" value={fd.gps} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Inactive', 'Completed']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Project Sites" subtitle="Manage sites" columns={columns} initialData={DEMO_PROJECT_SITES} searchKeys={['site_code', 'project', 'location']} statusOptions={['Active', 'Inactive', 'Completed']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.site_code || !fd.project) ? 'Required' : null} addLabel="Add Site" />;
};

const ProjectMilestonesPage = () => {
  const columns = [
    { key: 'project', label: 'Project', strong: true },
    { key: 'milestone_type', label: 'Type' },
    { key: 'milestone', label: 'Milestone' },
    { key: 'target_date', label: 'Target', render: (i) => formatDate(i.target_date) },
    { key: 'actual_date', label: 'Actual', render: (i) => formatDate(i.actual_date) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', project: '', milestone_type: '', milestone: '', target_date: '', actual_date: '', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Project" icon={<FolderOpen size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Milestone Type" icon={<Layers size={12} />} name="milestone_type" value={fd.milestone_type} onChange={onChange} options={getNames(MASTER_MILESTONE_TYPE)} />
        <Field label="Milestone" icon={<FileText size={12} />} name="milestone" value={fd.milestone} onChange={onChange} required />
        <Field label="Target Date" icon={<Calendar size={12} />} name="target_date" type="date" value={fd.target_date} onChange={onChange} />
        <Field label="Actual Date" icon={<Calendar size={12} />} name="actual_date" type="date" value={fd.actual_date} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'In Progress', 'Completed']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Milestones" subtitle="Track milestones" columns={columns} initialData={DEMO_PROJECT_MILESTONES} searchKeys={['project', 'milestone']} statusOptions={['Pending', 'In Progress', 'Completed']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.project || !fd.milestone) ? 'Required' : null} addLabel="Add Milestone" />;
};

const ProjectBudgetPage = () => {
  const columns = [
    { key: 'project', label: 'Project', strong: true },
    { key: 'category', label: 'Category' },
    { key: 'sanctioned', label: 'Sanctioned', render: (i) => `₹${Number(i.sanctioned).toLocaleString('en-IN')}` },
    { key: 'spent', label: 'Spent', render: (i) => `₹${Number(i.spent).toLocaleString('en-IN')}` },
    { key: 'balance', label: 'Balance', render: (i) => `₹${Number(i.balance).toLocaleString('en-IN')}` },
    { key: 'utilization', label: 'Utilization' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', project: '', category: '', sanctioned: '', spent: '', balance: '', utilization: '', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Project" icon={<FolderOpen size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Category" icon={<Tag size={12} />} name="category" value={fd.category} onChange={onChange} options={getNames(MASTER_PROJ_CATEGORY)} />
        <Field label="Sanctioned (₹)" icon={<Banknote size={12} />} name="sanctioned" type="number" value={fd.sanctioned} onChange={onChange} required />
        <Field label="Spent (₹)" icon={<Wallet size={12} />} name="spent" type="number" value={fd.spent} onChange={onChange} />
        <Field label="Balance (₹)" icon={<Banknote size={12} />} name="balance" type="number" value={fd.balance} onChange={onChange} />
        <Field label="Utilization" icon={<TrendingUp size={12} />} name="utilization" value={fd.utilization} onChange={onChange} placeholder="75%" />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Approved', 'Correction', 'Completed']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Project Budgets" subtitle="Track budgets" columns={columns} initialData={DEMO_PROJECT_BUDGETS} searchKeys={['project']} statusOptions={['Pending', 'Approved', 'Correction', 'Completed']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.project || !fd.sanctioned) ? 'Required' : null} addLabel="Add Budget" />;
};

const ProjectDocumentsPage = () => {
  const columns = [
    { key: 'project', label: 'Project', strong: true },
    { key: 'doc_name', label: 'Document' },
    { key: 'doc_type', label: 'Type' },
    { key: 'size', label: 'Size' },
    { key: 'uploaded', label: 'Uploaded', render: (i) => formatDate(i.uploaded) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', project: '', doc_name: '', doc_type: '', size: '', uploaded: '', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Project" icon={<FolderOpen size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Document Name" icon={<FileText size={12} />} name="doc_name" value={fd.doc_name} onChange={onChange} required />
        <Field label="Type" icon={<Tag size={12} />} name="doc_type" value={fd.doc_type} onChange={onChange} options={getNames(MASTER_DOC_TYPE)} />
        <Field label="Size" icon={<Gauge size={12} />} name="size" value={fd.size} onChange={onChange} placeholder="2.4 MB" />
        <Field label="Uploaded" icon={<Calendar size={12} />} name="uploaded" type="date" value={fd.uploaded} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Approved', 'Rejected']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Project Documents" subtitle="Manage documents" columns={columns} initialData={DEMO_PROJECT_DOCUMENTS} searchKeys={['project', 'doc_name']} statusOptions={['Pending', 'Approved', 'Rejected']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.project || !fd.doc_name) ? 'Required' : null} addLabel="Add Document" />;
};

const ProjectAssignmentsPage = () => {
  const columns = [
    { key: 'project', label: 'Project', strong: true },
    { key: 'contractor', label: 'Contractor' },
    { key: 'worker', label: 'Worker' },
    { key: 'role', label: 'Role' },
    { key: 'assigned_date', label: 'Assigned', render: (i) => formatDate(i.assigned_date) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', project: '', contractor: '', worker: '', role: '', assigned_date: '', status: 'Active', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Project" icon={<FolderOpen size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Contractor" icon={<Building2 size={12} />} name="contractor" value={fd.contractor} onChange={onChange} options={DEMO_CONTRACTORS.map(c => c.name)} />
        <Field label="Field Worker" icon={<User size={12} />} name="worker" value={fd.worker} onChange={onChange} options={DEMO_FIELD_WORKERS.map(w => w.name)} />
        <Field label="Role" icon={<Award size={12} />} name="role" value={fd.role} onChange={onChange} options={getNames(MASTER_WORKER_ROLE)} />
        <Field label="Assigned Date" icon={<Calendar size={12} />} name="assigned_date" type="date" value={fd.assigned_date} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Completed', 'Inactive']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Assignments" subtitle="Manage assignments" columns={columns} initialData={DEMO_PROJECT_ASSIGNMENTS} searchKeys={['project', 'contractor', 'worker']} statusOptions={['Active', 'Completed', 'Inactive']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.project ? 'Required' : null} addLabel="Add Assignment" />;
};

// Contractors
const ContractorPage = () => {
  const columns = [
    { key: 'name', label: 'Contractor Name', strong: true }, { key: 'type', label: 'Type' },
    { key: 'contact', label: 'Contact' }, { key: 'email', label: 'Email' },
    { key: 'license_type', label: 'License Type' }, { key: 'projects', label: 'Projects' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', name: '', type: '', contact: '', email: '', license: '', license_type: '', projects: '', status: 'Pending', joined: '', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Contractor Name" icon={<Building2 size={12} />} name="name" value={fd.name} onChange={onChange} required />
        <Field label="Contractor Type" icon={<Tag size={12} />} name="type" value={fd.type} onChange={onChange} options={getNames(MASTER_CONTRACTOR_TYPE)} />
        <Field label="Contact" icon={<Phone size={12} />} name="contact" value={fd.contact} onChange={onChange} />
        <Field label="Email" icon={<Mail size={12} />} name="email" type="email" value={fd.email} onChange={onChange} />
        <Field label="License No" icon={<Award size={12} />} name="license" value={fd.license} onChange={onChange} />
        <Field label="License Type" icon={<Shield size={12} />} name="license_type" value={fd.license_type} onChange={onChange} options={getNames(MASTER_LICENSE_TYPE)} />
        <Field label="Projects" icon={<FolderOpen size={12} />} name="projects" type="number" value={fd.projects} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Approved', 'Rejected']} />
        <Field label="Joined Date" icon={<Calendar size={12} />} name="joined" type="date" value={fd.joined} onChange={onChange} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Contractors" subtitle="Manage contractors" columns={columns} initialData={DEMO_CONTRACTORS} searchKeys={['name', 'email', 'license']} statusOptions={['Pending', 'Approved', 'Rejected']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.name ? 'Name required' : null} addLabel="Add Contractor" />;
};

// Field Workers
const FieldWorkerPage = () => {
  const columns = [
    { key: 'name', label: 'Worker Name', strong: true }, { key: 'role', label: 'Role' },
    { key: 'skill', label: 'Skill' }, { key: 'phone', label: 'Phone' },
    { key: 'contractor', label: 'Contractor' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', name: '', phone: '', role: '', skill: '', contractor: '', assigned_sites: '', status: 'Active', last_activity: '', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Worker Name" icon={<User size={12} />} name="name" value={fd.name} onChange={onChange} required />
        <Field label="Role" icon={<Award size={12} />} name="role" value={fd.role} onChange={onChange} options={getNames(MASTER_WORKER_ROLE)} />
        <Field label="Skill" icon={<Tag size={12} />} name="skill" value={fd.skill} onChange={onChange} options={getNames(MASTER_WORKER_SKILL)} />
        <Field label="Phone" icon={<Phone size={12} />} name="phone" value={fd.phone} onChange={onChange} />
        <Field label="Contractor" icon={<Building2 size={12} />} name="contractor" value={fd.contractor} onChange={onChange} options={DEMO_CONTRACTORS.map(c => c.name)} />
        <Field label="Assigned Sites" icon={<FolderOpen size={12} />} name="assigned_sites" type="number" value={fd.assigned_sites} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Inactive']} />
        <Field label="Last Activity" icon={<Calendar size={12} />} name="last_activity" type="date" value={fd.last_activity} onChange={onChange} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Field Workers" subtitle="Manage workers" columns={columns} initialData={DEMO_FIELD_WORKERS} searchKeys={['name', 'phone', 'contractor']} statusOptions={['Active', 'Inactive']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.name ? 'Required' : null} addLabel="Add Worker" />;
};

// Reports
const ReportsPage = () => {
  const columns = [
    { key: 'project', label: 'Project', strong: true }, { key: 'worker', label: 'Worker' },
    { key: 'report_type', label: 'Type' }, { key: 'priority', label: 'Priority' },
    { key: 'submitted', label: 'Submitted', render: (i) => formatDate(i.submitted) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', project: '', worker: '', report_type: '', priority: '', submitted: '', photos: '', gps: 'Yes', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Project" icon={<FileText size={12} />} name="project" value={fd.project} onChange={onChange} options={DEMO_PROJECTS.map(p => p.name)} required />
        <Field label="Field Worker" icon={<User size={12} />} name="worker" value={fd.worker} onChange={onChange} options={DEMO_FIELD_WORKERS.map(w => w.name)} />
        <Field label="Report Type" icon={<Tag size={12} />} name="report_type" value={fd.report_type} onChange={onChange} options={getNames(MASTER_REPORT_TYPE)} />
        <Field label="Priority" icon={<AlertCircle size={12} />} name="priority" value={fd.priority} onChange={onChange} options={getNames(MASTER_REPORT_PRIORITY)} />
        <Field label="Submitted" icon={<Calendar size={12} />} name="submitted" type="date" value={fd.submitted} onChange={onChange} />
        <Field label="Photos" icon={<Camera size={12} />} name="photos" type="number" value={fd.photos} onChange={onChange} />
        <Field label="GPS" icon={<MapPin size={12} />} name="gps" value={fd.gps} onChange={onChange} options={['Yes', 'No']} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Approved', 'Correction', 'Rejected']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Reports" subtitle="Field reports" columns={columns} initialData={DEMO_REPORTS} searchKeys={['project', 'worker']} statusOptions={['Pending', 'Approved', 'Correction', 'Rejected']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.project ? 'Required' : null} addLabel="Add Report" />;
};

// Users
const UsersPage = () => {
  const columns = [
    { key: 'name', label: 'Name', strong: true }, { key: 'email', label: 'Email' },
    { key: 'role', label: 'Role', render: (i) => <span style={{ ...S.statusBadge, background: '#dbeafe', color: '#1e40af' }}>{i.role}</span> },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
    { key: 'remarks', label: 'Remarks' },
  ];
  const formInitial = { s_no: '', name: '', email: '', role: '', status: 'Active', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Name" icon={<User size={12} />} name="name" value={fd.name} onChange={onChange} required />
        <Field label="Email" icon={<Mail size={12} />} name="email" type="email" value={fd.email} onChange={onChange} />
        <Field label="Role" icon={<Shield size={12} />} name="role" value={fd.role} onChange={onChange} options={getNames(MASTER_USER_ROLE)} />
        <Field label="Status" icon={<CheckCircle size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Inactive']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Users" subtitle="Manage users" columns={columns} initialData={DEMO_USERS} searchKeys={['name', 'email', 'role']} statusOptions={['Active', 'Inactive']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.name || !fd.email) ? 'Required' : null} addLabel="Add User" />;
};

// Applications
const ContractorAppPage = () => {
  const columns = [
    { key: 'contractor', label: 'Contractor', strong: true }, { key: 'app_name', label: 'App' },
    { key: 'version', label: 'Version' }, { key: 'device', label: 'Device' },
    { key: 'last_sync', label: 'Last Sync' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', contractor: '', app_name: 'Contractor App', version: '', device: '', os_version: '', last_sync: '', status: 'Active', projects: '', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Contractor" icon={<Building2 size={12} />} name="contractor" value={fd.contractor} onChange={onChange} options={DEMO_CONTRACTORS.map(c => c.name)} required />
        <Field label="App Name" icon={<Smartphone size={12} />} name="app_name" value={fd.app_name} onChange={onChange} options={getNames(MASTER_APP_NAME)} />
        <Field label="Version" icon={<Hash size={12} />} name="version" value={fd.version} onChange={onChange} placeholder="1.2.0" />
        <Field label="Device" icon={<Smartphone size={12} />} name="device" value={fd.device} onChange={onChange} placeholder="Android 12" />
        <Field label="OS Version" icon={<Server size={12} />} name="os_version" value={fd.os_version} onChange={onChange} options={getNames(MASTER_OS_VERSION)} />
        <Field label="Last Sync" icon={<Clock size={12} />} name="last_sync" value={fd.last_sync} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Inactive']} />
        <Field label="Projects" icon={<FolderOpen size={12} />} name="projects" type="number" value={fd.projects} onChange={onChange} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Contractor App" subtitle="Manage contractor app" columns={columns} initialData={DEMO_CONTRACTOR_APPS} searchKeys={['contractor', 'device']} statusOptions={['Active', 'Inactive']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.contractor ? 'Required' : null} addLabel="Add Installation" />;
};

const WorkerAppPage = () => {
  const columns = [
    { key: 'worker', label: 'Field Worker', strong: true }, { key: 'app_name', label: 'App' },
    { key: 'version', label: 'Version' }, { key: 'device', label: 'Device' },
    { key: 'gps', label: 'GPS' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', worker: '', app_name: 'Field Worker App', version: '', device: '', last_sync: '', gps: 'Enabled', status: 'Active', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Field Worker" icon={<User size={12} />} name="worker" value={fd.worker} onChange={onChange} options={DEMO_FIELD_WORKERS.map(w => w.name)} required />
        <Field label="App Name" icon={<Smartphone size={12} />} name="app_name" value={fd.app_name} onChange={onChange} options={getNames(MASTER_APP_NAME)} />
        <Field label="Version" icon={<Hash size={12} />} name="version" value={fd.version} onChange={onChange} placeholder="1.2.0" />
        <Field label="Device" icon={<Smartphone size={12} />} name="device" value={fd.device} onChange={onChange} />
        <Field label="Last Sync" icon={<Clock size={12} />} name="last_sync" value={fd.last_sync} onChange={onChange} />
        <Field label="GPS" icon={<MapPin size={12} />} name="gps" value={fd.gps} onChange={onChange} options={getNames(MASTER_GPS_STATUS)} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Inactive']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Field Worker App" subtitle="Manage worker app" columns={columns} initialData={DEMO_WORKER_APPS} searchKeys={['worker', 'device']} statusOptions={['Active', 'Inactive']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.worker ? 'Required' : null} addLabel="Add Installation" />;
};

const AppVersionsPage = () => {
  const columns = [
    { key: 'app', label: 'App', strong: true }, { key: 'version', label: 'Version' },
    { key: 'released', label: 'Released', render: (i) => formatDate(i.released) },
    { key: 'size', label: 'Size' }, { key: 'downloads', label: 'Downloads' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', app: '', version: '', released: '', size: '', downloads: '', status: 'Active', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Application" icon={<Smartphone size={12} />} name="app" value={fd.app} onChange={onChange} options={getNames(MASTER_APP_NAME)} required />
        <Field label="Version" icon={<Hash size={12} />} name="version" value={fd.version} onChange={onChange} required />
        <Field label="Released" icon={<Calendar size={12} />} name="released" type="date" value={fd.released} onChange={onChange} />
        <Field label="Size" icon={<Gauge size={12} />} name="size" value={fd.size} onChange={onChange} />
        <Field label="Downloads" icon={<Download size={12} />} name="downloads" type="number" value={fd.downloads} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Active', 'Deprecated']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="App Versions" subtitle="Manage versions" columns={columns} initialData={DEMO_APP_VERSIONS} searchKeys={['app', 'version']} statusOptions={['Active', 'Deprecated']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.app || !fd.version) ? 'Required' : null} addLabel="Add Version" />;
};

const AppReleasesPage = () => {
  const columns = [
    { key: 'app', label: 'App', strong: true }, { key: 'version', label: 'Version' },
    { key: 'release_type', label: 'Type' },
    { key: 'released', label: 'Released', render: (i) => formatDate(i.released) },
    { key: 'features', label: 'Features' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', app: '', version: '', release_type: '', released: '', features: '', status: 'In Development', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Application" icon={<Smartphone size={12} />} name="app" value={fd.app} onChange={onChange} options={getNames(MASTER_APP_NAME)} required />
        <Field label="Version" icon={<Hash size={12} />} name="version" value={fd.version} onChange={onChange} required />
        <Field label="Release Type" icon={<Rocket size={12} />} name="release_type" value={fd.release_type} onChange={onChange} options={getNames(MASTER_RELEASE_TYPE)} />
        <Field label="Released" icon={<Calendar size={12} />} name="released" type="date" value={fd.released} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['In Development', 'In Testing', 'Released']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Features</label>
          <textarea name="features" value={fd.features} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={2} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="App Releases" subtitle="Manage app releases" columns={columns} initialData={DEMO_APP_RELEASES} searchKeys={['app', 'version']} statusOptions={['In Development', 'In Testing', 'Released']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.app || !fd.version) ? 'Required' : null} addLabel="Add Release" />;
};

const AppCrashLogsPage = () => {
  const columns = [
    { key: 'app', label: 'App', strong: true }, { key: 'version', label: 'Version' },
    { key: 'user', label: 'User' }, { key: 'crash_type', label: 'Crash Type' },
    { key: 'occurred', label: 'Occurred' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', app: '', version: '', user: '', crash_type: '', occurred: '', device: '', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Application" icon={<Smartphone size={12} />} name="app" value={fd.app} onChange={onChange} options={getNames(MASTER_APP_NAME)} required />
        <Field label="Version" icon={<Hash size={12} />} name="version" value={fd.version} onChange={onChange} required />
        <Field label="User" icon={<User size={12} />} name="user" value={fd.user} onChange={onChange} />
        <Field label="Crash Type" icon={<Bug size={12} />} name="crash_type" value={fd.crash_type} onChange={onChange} options={getNames(MASTER_CRASH_TYPE)} />
        <Field label="Occurred" icon={<Clock size={12} />} name="occurred" value={fd.occurred} onChange={onChange} />
        <Field label="Device" icon={<Smartphone size={12} />} name="device" value={fd.device} onChange={onChange} options={getNames(MASTER_OS_VERSION)} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Resolved']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Crash Logs" subtitle="App crash logs" columns={columns} initialData={DEMO_APP_CRASH_LOGS} searchKeys={['app', 'user', 'crash_type']} statusOptions={['Pending', 'Resolved']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => (!fd.app || !fd.version) ? 'Required' : null} addLabel="Add Crash Log" />;
};

const AppFeedbackPage = () => {
  const columns = [
    { key: 'app', label: 'App', strong: true }, { key: 'user', label: 'User' },
    { key: 'rating', label: 'Rating' }, { key: 'feedback', label: 'Feedback' },
    { key: 'submitted', label: 'Submitted', render: (i) => formatDate(i.submitted) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  const formInitial = { s_no: '', app: '', user: '', rating: '', feedback: '', submitted: '', status: 'Pending', remarks: '' };
  const renderForm = (fd, setFd) => {
    const onChange = e => setFd({ ...fd, [e.target.name]: e.target.value });
    return (
      <>
        <Field label="Application" icon={<Smartphone size={12} />} name="app" value={fd.app} onChange={onChange} options={getNames(MASTER_APP_NAME)} required />
        <Field label="User" icon={<User size={12} />} name="user" value={fd.user} onChange={onChange} />
        <Field label="Rating" icon={<Award size={12} />} name="rating" value={fd.rating} onChange={onChange} options={getNames(MASTER_RATING)} />
        <Field label="Submitted" icon={<Calendar size={12} />} name="submitted" type="date" value={fd.submitted} onChange={onChange} />
        <Field label="Status" icon={<Shield size={12} />} name="status" value={fd.status} onChange={onChange} options={['Pending', 'Reviewed']} />
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Feedback</label>
          <textarea name="feedback" value={fd.feedback} onChange={onChange} rows={3} style={S.formTextarea} />
        </div>
        <div style={S.formGroupFull}><label style={S.formLabel}><FileText size={12} /> Remarks</label>
          <textarea name="remarks" value={fd.remarks} onChange={onChange} rows={2} style={S.formTextarea} />
        </div>
      </>
    );
  };
  return <TablePage title="Feedback" subtitle="App feedback" columns={columns} initialData={DEMO_APP_FEEDBACK} searchKeys={['app', 'user']} statusOptions={['Pending', 'Reviewed']} formInitial={formInitial} renderForm={renderForm} validateForm={(fd) => !fd.app ? 'Required' : null} addLabel="Add Feedback" />;
};

// Media
const PhotosPage = () => {
  const columns = [
    { key: 'name', label: 'Photo', strong: true }, { key: 'project', label: 'Project' },
    { key: 'media_type', label: 'Media Type' }, { key: 'category', label: 'Category' },
    { key: 'worker', label: 'Uploaded By' }, { key: 'gps', label: 'GPS' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Photo Gallery</h1><p style={S.subtitle}>View photographs</p></div>
      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 300px', maxWidth: '500px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder="Search..." style={S.searchInput} />
        </div>
        <button onClick={() => exportToExcel(DEMO_PHOTOS, columns, 'photos')} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(DEMO_PHOTOS, columns, 'Photos')} style={S.btn}><Printer size={14} /> Print</button>
      </div>
      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Photos</h2><span style={S.countBadge}>{DEMO_PHOTOS.length}</span></div>
        <div style={{ padding: '20px' }}>
          <div style={S.photoGrid}>
            {DEMO_PHOTOS.map(p => (
              <div key={p.id} style={S.photoCard}>
                <div style={S.photoThumb}><Camera size={36} /></div>
                <div style={S.photoInfo}>
                  <p style={{ fontSize: '12px', fontWeight: '600', color: '#374151', margin: '0 0 2px 0' }}>{p.name}</p>
                  <p style={{ fontSize: '11px', color: '#9ca3af', margin: 0 }}>{p.project}</p>
                  <p style={{ fontSize: '10px', color: '#9ca3af', marginTop: '4px' }}>GPS: {p.gps}</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

const DocumentsPage = () => {
  const columns = [
    { key: 'name', label: 'Document', strong: true }, { key: 'project', label: 'Project' },
    { key: 'type', label: 'Type' }, { key: 'size', label: 'Size' },
    { key: 'uploaded', label: 'Uploaded', render: (i) => formatDate(i.uploaded) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Documents</h1><p style={S.subtitle}>Manage documents</p></div>
      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 300px', maxWidth: '500px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder="Search..." style={S.searchInput} />
        </div>
        <button onClick={() => exportToExcel(DEMO_DOCUMENTS, columns, 'documents')} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(DEMO_DOCUMENTS, columns, 'Documents')} style={S.btn}><Printer size={14} /> Print</button>
      </div>
      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Documents</h2><span style={S.countBadge}>{DEMO_DOCUMENTS.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}</tr></thead>
            <tbody>
              {DEMO_DOCUMENTS.map((d, i) => (
                <tr key={d.id}>
                  <td style={S.td}>{i + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(d) : (d[c.key] ?? '—')}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// Approvals
const ApprovalsPage = () => {
  const items = [
    { id: 1, s_no: 1, step: 'Registration', title: 'Contractor Registration', desc: 'EcoEnergy Pvt Ltd', status: 'Pending', date: '2024-03-10' },
    { id: 2, s_no: 2, step: 'Project Submission', title: 'Project Submission', desc: 'Solar Pump Installation', status: 'Pending', date: '2024-03-11' },
    { id: 3, s_no: 3, step: 'Report Review', title: 'Field Report Review', desc: 'Ground Mount Solar', status: 'Pending', date: '2024-03-12' },
    { id: 4, s_no: 4, step: 'Final Approval', title: 'Correction Required', desc: 'Solar Water Heater', status: 'Correction', date: '2024-03-13' },
  ];
  const columns = [
    { key: 'step', label: 'Step', strong: true }, { key: 'title', label: 'Title' }, { key: 'desc', label: 'Description' },
    { key: 'date', label: 'Date', render: (i) => formatDate(i.date) },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Approvals</h1><p style={S.subtitle}>Approval workflow</p></div>
      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 300px', maxWidth: '500px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder="Search..." style={S.searchInput} />
        </div>
        <button onClick={() => exportToExcel(items, columns, 'approvals')} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(items, columns, 'Approvals')} style={S.btn}><Printer size={14} /> Print</button>
      </div>
      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>Pending Approvals</h2><span style={S.countBadge}>{items.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}<th style={S.th}>Actions</th></tr></thead>
            <tbody>
              {items.map(item => (
                <tr key={item.id}>
                  <td style={S.td}>{item.s_no}</td>
                  <td style={S.tdStrong}>{item.step}</td>
                  <td style={S.td}>{item.title}</td>
                  <td style={S.td}>{item.desc}</td>
                  <td style={S.td}>{formatDate(item.date)}</td>
                  <td style={S.td}><span style={{ ...S.statusBadge, ...getStatusStyle(item.status) }}>{item.status}</span></td>
                  <td style={S.td}>
                    <div style={{ display: 'flex', gap: '6px' }}>
                      <button onClick={() => toast.success('Approved')} style={{ ...S.iconBtn, color: '#16a34a' }}><CheckCircle size={15} /></button>
                      <button onClick={() => toast.warning('Correction')} style={{ ...S.iconBtn, color: '#f59e0b' }}><Edit size={15} /></button>
                      <button onClick={() => toast.error('Rejected')} style={{ ...S.iconBtn, color: '#dc2626' }}><X size={15} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// Activity
const ActivityPage = () => {
  const columns = [
    { key: 'type', label: 'Type', strong: true }, { key: 'description', label: 'Description' },
    { key: 'user', label: 'User' }, { key: 'date', label: 'Date' },
    { key: 'status', label: 'Status', render: (i) => <span style={{ ...S.statusBadge, ...getStatusStyle(i.status) }}>{i.status}</span> },
  ];
  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Activity Log</h1><p style={S.subtitle}>Recent activities</p></div>
      <div style={S.searchBarTop}>
        <div style={{ position: 'relative', flex: '1 1 300px', maxWidth: '500px' }}>
          <Search size={16} style={S.searchIcon} />
          <input type="text" placeholder="Search..." style={S.searchInput} />
        </div>
        <button onClick={() => exportToExcel(DEMO_ACTIVITIES, columns, 'activities')} style={S.btn}><FileSpreadsheet size={14} /> Excel</button>
        <button onClick={() => printReport(DEMO_ACTIVITIES, columns, 'Activities')} style={S.btn}><Printer size={14} /> Print</button>
      </div>
      <div style={S.tableCard}>
        <div style={S.tableHeaderBar}><h2 style={S.tableTitle}>All Activities</h2><span style={S.countBadge}>{DEMO_ACTIVITIES.length}</span></div>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead style={S.thead}><tr><th style={S.th}>S/No.</th>{columns.map(c => <th key={c.key} style={S.th}>{c.label}</th>)}</tr></thead>
            <tbody>
              {DEMO_ACTIVITIES.map((a, i) => (
                <tr key={a.id}>
                  <td style={S.td}>{i + 1}</td>
                  {columns.map(c => <td key={c.key} style={c.strong ? S.tdStrong : S.td}>{c.render ? c.render(a) : (a[c.key] ?? '—')}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// Dashboard
const DashboardPage = () => {
  const stats = [
    { label: 'Total Projects', value: DEMO_PROJECTS.length, color: '#3b82f6', icon: <FolderOpen size={20} /> },
    { label: 'Approved', value: DEMO_PROJECTS.filter(p => p.status === 'Approved').length, color: '#16a34a', icon: <CheckCircle size={20} /> },
    { label: 'Pending', value: DEMO_PROJECTS.filter(p => p.status === 'Pending').length, color: '#f59e0b', icon: <Clock size={20} /> },
    { label: 'Completed', value: DEMO_PROJECTS.filter(p => p.status === 'Completed').length, color: '#8b5cf6', icon: <Award size={20} /> },
    { label: 'Contractors', value: DEMO_CONTRACTORS.length, color: '#0891b2', icon: <Building2 size={20} /> },
    { label: 'Workers', value: DEMO_FIELD_WORKERS.length, color: '#7c3aed', icon: <Users size={20} /> },
    { label: 'Reports', value: DEMO_REPORTS.length, color: '#dc2626', icon: <ClipboardList size={20} /> },
    { label: 'Documents', value: DEMO_DOCUMENTS.length, color: '#ea580c', icon: <FileText size={20} /> },
  ];
  return (
    <div style={S.page}>
      <ToastContainer position="top-right" autoClose={3000} />
      <div style={S.header}><h1 style={S.title}>Management Dashboard</h1><p style={S.subtitle}>Complete overview</p></div>
      <StatsRow stats={stats} />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: '20px' }}>
        <div style={S.card}>
          <div style={S.cardHeader}><h2 style={S.cardTitle}><FolderOpen size={18} color="#2563eb" /> Project Progress</h2></div>
          {DEMO_PROJECTS.slice(0, 6).map(p => {
            const c = p.progress === 100 ? '#16a34a' : p.progress >= 50 ? '#3b82f6' : '#f59e0b';
            return (
              <div key={p.id} style={{ marginBottom: '14px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                  <div style={{ minWidth: 0, flex: 1 }}>
                    <p style={{ fontSize: '13px', fontWeight: '600', color: '#111827', margin: 0 }}>{p.name}</p>
                    <p style={{ fontSize: '11px', color: '#9ca3af', margin: '2px 0 0 0' }}>{p.site_code}</p>
                  </div>
                  <span style={{ fontSize: '13px', fontWeight: '700', color: c }}>{p.progress}%</span>
                </div>
                <div style={S.progressTrack}><div style={{ ...S.progressFill, width: `${p.progress}%`, background: c }} /></div>
              </div>
            );
          })}
        </div>
        <div style={S.card}>
          <div style={S.cardHeader}><h2 style={S.cardTitle}><Activity size={18} color="#dc2626" /> Recent Activity</h2></div>
          {DEMO_ACTIVITIES.map((a, i) => (
            <div key={a.id} style={{ display: 'flex', gap: '12px', padding: '12px 0', borderBottom: i < DEMO_ACTIVITIES.length - 1 ? '1px solid #f3f4f6' : 'none' }}>
              <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: '#eff6ff', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}><Activity size={14} /></div>
              <div style={{ flex: 1 }}>
                <p style={{ fontSize: '12px', color: '#111827', margin: 0, fontWeight: '600' }}>{a.type}</p>
                <p style={{ fontSize: '11px', color: '#6b7280', margin: '2px 0 0 0' }}>{a.description}</p>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// =============================================
// MENU CONFIG - MORE SUB-MENUS
// =============================================
const MENU_CONFIG = {
  dashboard: {
    main: [
      { id: 'dashboard-main', label: 'Management Dashboard' },
      { id: 'activity-log', label: 'Activity Log' },
    ],
    masters: [
      { id: 'dm-widget', label: 'Widget Types', table: 'dm_widget', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: [{ id: 1, s_no: 1, code: 'STAT', name: 'Stat Card', status: 'Active', remarks: 'Card' }, { id: 2, s_no: 2, code: 'CHART', name: 'Chart', status: 'Active', remarks: 'Chart' }] },
      { id: 'dm-report', label: 'Report Types', table: 'dm_report', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: [{ id: 1, s_no: 1, code: 'DAILY', name: 'Daily', status: 'Active', remarks: 'Daily' }, { id: 2, s_no: 2, code: 'MONTHLY', name: 'Monthly', status: 'Active', remarks: 'Monthly' }] },
    ],
  },
  projects: {
    main: [
      { id: 'proj-list', label: 'All Projects' },
      { id: 'proj-sites', label: 'Project Sites' },
      { id: 'proj-milestones', label: 'Milestones' },
      { id: 'proj-budgets', label: 'Budgets' },
      { id: 'proj-documents', label: 'Documents' },
      { id: 'proj-assignments', label: 'Assignments' },
    ],
    masters: [
      { id: 'pm-status', label: 'Project Status', table: 'pm_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'color', label: 'Color' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_PROJ_STATUS },
      { id: 'pm-category', label: 'Project Category', table: 'pm_category', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_PROJ_CATEGORY },
      { id: 'pm-capacity', label: 'Capacity', table: 'pm_capacity', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Capacity' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_CAPACITY },
      { id: 'pm-milestone', label: 'Milestone Type', table: 'pm_milestone', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_MILESTONE_TYPE },
      { id: 'pm-doctype', label: 'Document Type', table: 'pm_doctype', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_DOC_TYPE },
    ],
  },
  contractors: {
    main: [{ id: 'contractors-all', label: 'All Contractors' }],
    masters: [
      { id: 'cm-type', label: 'Contractor Type', table: 'cm_type', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_CONTRACTOR_TYPE },
      { id: 'cm-license', label: 'License Type', table: 'cm_license', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_LICENSE_TYPE },
      { id: 'cm-status', label: 'Approval Status', table: 'cm_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APPROVAL_STATUS },
    ],
  },
  'field-workers': {
    main: [{ id: 'fw-all', label: 'All Field Workers' }],
    masters: [
      { id: 'wm-role', label: 'Worker Role', table: 'wm_role', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_WORKER_ROLE },
      { id: 'wm-skill', label: 'Worker Skill', table: 'wm_skill', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_WORKER_SKILL },
      { id: 'wm-status', label: 'Worker Status', table: 'wm_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_STATUS },
    ],
  },
  reports: {
    main: [{ id: 'reports-all', label: 'All Reports' }],
    masters: [
      { id: 'rm-type', label: 'Report Type', table: 'rm_type', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_REPORT_TYPE },
      { id: 'rm-priority', label: 'Report Priority', table: 'rm_priority', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_REPORT_PRIORITY },
      { id: 'rm-status', label: 'Report Status', table: 'rm_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APPROVAL_STATUS },
    ],
  },
  media: {
    main: [{ id: 'photos', label: 'Photo Gallery' }, { id: 'documents', label: 'Documents' }],
    masters: [
      { id: 'mm-type', label: 'Media Type', table: 'mm_type', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_MEDIA_TYPE },
      { id: 'mm-category', label: 'Media Category', table: 'mm_category', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_MEDIA_CATEGORY },
      { id: 'mm-doctype', label: 'Document Type', table: 'mm_doctype', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_DOC_TYPE },
    ],
  },
  approvals: {
    main: [{ id: 'approvals-all', label: 'Pending Approvals' }],
    masters: [
      { id: 'am-step', label: 'Approval Step', table: 'am_step', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APPROVAL_STEP },
      { id: 'am-status', label: 'Approval Status', table: 'am_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APPROVAL_STATUS },
    ],
  },
  applications: {
    main: [
      { id: 'app-contractor', label: 'Contractor App' },
      { id: 'app-worker', label: 'Field Worker App' },
      { id: 'app-versions', label: 'App Versions' },
      { id: 'app-releases', label: 'App Releases' },
      { id: 'app-crashes', label: 'Crash Logs' },
      { id: 'app-feedback', label: 'App Feedback' },
    ],
    masters: [
      { id: 'apm-app', label: 'App Names', table: 'apm_app', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APP_NAME },
      { id: 'apm-gps', label: 'GPS Status', table: 'apm_gps', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_GPS_STATUS },
      { id: 'apm-os', label: 'OS Version', table: 'apm_os', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_OS_VERSION },
      { id: 'apm-release-type', label: 'Release Type', table: 'apm_release_type', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_RELEASE_TYPE },
      { id: 'apm-crash-type', label: 'Crash Type', table: 'apm_crash_type', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_CRASH_TYPE },
      { id: 'apm-rating', label: 'Rating', table: 'apm_rating', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_RATING },
    ],
  },
  masters: {
    main: [],
    masters: [
      { id: 'mm-approval', label: 'Approval Status', table: 'mm_approval', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_APPROVAL_STATUS },
      { id: 'mm-status', label: 'General Status', table: 'mm_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_STATUS },
      { id: 'mm-priority', label: 'Priority', table: 'mm_priority', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_REPORT_PRIORITY },
    ],
  },
  users: {
    main: [{ id: 'users-all', label: 'All Users' }],
    masters: [
      { id: 'um-role', label: 'User Role', table: 'um_role', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_USER_ROLE },
      { id: 'um-status', label: 'User Status', table: 'um_status', columns: [{ key: 'code', label: 'Code' }, { key: 'name', label: 'Name' }, { key: 'status', label: 'Status' }, { key: 'remarks', label: 'Remarks' }], data: MASTER_STATUS },
    ],
  },
};

// =============================================
// MAIN APP
// =============================================
const App = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [user, setUser] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [activeMenu, setActiveMenu] = useState('dashboard');
  const [activeSubMenu, setActiveSubMenu] = useState('dashboard-main');

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

  const renderContent = () => {
    const cm = MENU_CONFIG[activeMenu];
    if (cm) {
      const mi = cm.masters.find(m => m.id === activeSubMenu);
      if (mi) return <SimpleMasterPage title={mi.label} subtitle={`Master list of ${mi.label.toLowerCase()}`} storageKey={mi.table} defaultColumns={mi.columns} defaultData={mi.data} />;
    }
    switch (activeSubMenu) {
      case 'dashboard-main': return <DashboardPage />;
      case 'activity-log': return <ActivityPage />;
      case 'proj-list': return <ProjectListPage />;
      case 'proj-sites': return <ProjectSitesPage />;
      case 'proj-milestones': return <ProjectMilestonesPage />;
      case 'proj-budgets': return <ProjectBudgetPage />;
      case 'proj-documents': return <ProjectDocumentsPage />;
      case 'proj-assignments': return <ProjectAssignmentsPage />;
      case 'contractors-all': return <ContractorPage />;
      case 'fw-all': return <FieldWorkerPage />;
      case 'reports-all': return <ReportsPage />;
      case 'photos': return <PhotosPage />;
      case 'documents': return <DocumentsPage />;
      case 'approvals-all': return <ApprovalsPage />;
      case 'app-contractor': return <ContractorAppPage />;
      case 'app-worker': return <WorkerAppPage />;
      case 'app-versions': return <AppVersionsPage />;
      case 'app-releases': return <AppReleasesPage />;
      case 'app-crashes': return <AppCrashLogsPage />;
      case 'app-feedback': return <AppFeedbackPage />;
      case 'users-all': return <UsersPage />;
      default: return <DashboardPage />;
    }
  };

  const cm = MENU_CONFIG[activeMenu] || { main: [], masters: [] };

  return (
    <Router>
      <div style={{ display: 'flex', height: '100vh', background: '#f3f4f6', fontFamily: 'system-ui, sans-serif', overflow: 'hidden' }}>
        {/* SIDEBAR */}
        <div style={{ background: '#111827', color: '#fff', display: 'flex', flexDirection: 'column', width: sidebarOpen ? '250px' : '60px', transition: 'width 0.3s', flexShrink: 0 }}>
          <div style={{ padding: '16px', borderBottom: '1px solid #1f2937', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            {sidebarOpen && <span style={{ fontWeight: 'bold', fontSize: '15px', letterSpacing: '1px' }}>IOCL ERP</span>}
            <button onClick={() => setSidebarOpen(!sidebarOpen)} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex' }}>
              {sidebarOpen ? <ChevronLeft size={16} /> : <ChevronRight size={16} />}
            </button>
          </div>
          <div style={{ flex: 1, overflowY: 'auto', padding: '12px 8px' }}>
            {cm.main.length > 0 && (
              <>
                {sidebarOpen && <p style={{ fontSize: '10px', color: '#6b7280', margin: '8px 12px 6px', fontWeight: '700', letterSpacing: '1px', textTransform: 'uppercase' }}>Main</p>}
                {cm.main.map(item => (
                  <button key={item.id} onClick={() => setActiveSubMenu(item.id)} style={{
                    width: '100%', display: 'flex', alignItems: 'center', gap: '10px',
                    padding: '10px 12px', borderRadius: '8px', border: 'none',
                    background: activeSubMenu === item.id ? '#2563eb' : 'transparent',
                    color: activeSubMenu === item.id ? '#fff' : '#d1d5db',
                    cursor: 'pointer', fontSize: '13px', textAlign: 'left', marginBottom: '2px', fontFamily: 'inherit'
                  }}>
                    <span style={{ width: '6px', height: '6px', borderRadius: '50%', background: activeSubMenu === item.id ? '#fff' : '#60a5fa', flexShrink: 0 }} />
                    {sidebarOpen && <span>{item.label}</span>}
                  </button>
                ))}
              </>
            )}
            {cm.masters.length > 0 && (
              <>
                {sidebarOpen && <p style={{ fontSize: '10px', color: '#fbbf24', margin: '16px 12px 6px', fontWeight: '700', letterSpacing: '1px', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '6px' }}><Database size={10} /> Masters</p>}
                {!sidebarOpen && <div style={{ height: '1px', background: '#374151', margin: '12px 8px' }} />}
                {cm.masters.map(item => (
                  <button key={item.id} onClick={() => setActiveSubMenu(item.id)} style={{
                    width: '100%', display: 'flex', alignItems: 'center', gap: '10px',
                    padding: '9px 12px', borderRadius: '8px', border: 'none',
                    background: activeSubMenu === item.id ? '#7c3aed' : 'transparent',
                    color: activeSubMenu === item.id ? '#fff' : '#9ca3af',
                    cursor: 'pointer', fontSize: '12.5px', textAlign: 'left', marginBottom: '2px', fontFamily: 'inherit'
                  }}>
                    <List size={12} style={{ flexShrink: 0 }} />
                    {sidebarOpen && <span>{item.label}</span>}
                  </button>
                ))}
              </>
            )}
          </div>
          <div style={{ padding: '12px', borderTop: '1px solid #1f2937' }}>
            {sidebarOpen ? (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <div style={{ width: '30px', height: '30px', background: '#374151', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><User size={14} /></div>
                  <div>
                    <p style={{ fontSize: '12px', fontWeight: '600', margin: 0 }}>{user?.name || 'Admin'}</p>
                    <p style={{ fontSize: '10px', color: '#9ca3af', margin: 0 }}>{user?.role || 'Super Admin'}</p>
                  </div>
                </div>
                <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex' }}><LogOut size={14} /></button>
              </div>
            ) : (
              <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#9ca3af', cursor: 'pointer', width: '100%', display: 'flex', justifyContent: 'center' }}><LogOut size={16} /></button>
            )}
          </div>
        </div>

        {/* MAIN */}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          <div style={{ background: '#fff', borderBottom: '1px solid #e5e7eb', flexShrink: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 24px' }}>
              <h2 style={{ fontSize: '15px', fontWeight: '600', color: '#1f2937', margin: 0 }}>
                {topMenus.find(m => m.id === activeMenu)?.label} / {cm.main.find(s => s.id === activeSubMenu)?.label || cm.masters.find(m => m.id === activeSubMenu)?.label || ''}
              </h2>
              <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
                <span style={{ fontSize: '12px', color: '#4b5563' }}>Welcome, {user?.name || 'Admin'}</span>
                <button style={{ background: 'transparent', border: 'none', color: '#4b5563', cursor: 'pointer', display: 'flex' }}><Bell size={16} /></button>
                <button onClick={handleLogout} style={{ background: 'transparent', border: 'none', color: '#dc2626', fontSize: '12px', fontWeight: '600', cursor: 'pointer' }}>Logout</button>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '4px', padding: '0 24px', borderTop: '1px solid #f3f4f6', overflowX: 'auto' }}>
              {topMenus.map(menu => (
                <button key={menu.id} onClick={() => {
                  setActiveMenu(menu.id);
                  const cfg = MENU_CONFIG[menu.id];
                  const first = cfg.main[0] || cfg.masters[0];
                  if (first) setActiveSubMenu(first.id);
                }} style={{
                  padding: '10px 14px', fontSize: '13px', fontWeight: '500', background: 'transparent', border: 'none',
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