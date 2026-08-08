// @ts-nocheck
import React, { useState, useEffect, useMemo } from 'react';
import { 
  Home, Award, FileText, BookOpen, Mic, Briefcase, Users, 
  Settings, LogOut, Plus, Edit2, Trash2, Link as LinkIcon, 
  Printer, User, UserCheck, Menu, X, Save, Eye, Palette, Lock, Upload, Loader2, ChevronDown, ChevronUp, FileSpreadsheet, AlertCircle, Filter, Monitor,
  BarChart3, CheckCircle2, Database, History, ShieldAlert
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import { 
  getAuth, signInWithCustomToken, signInAnonymously, onAuthStateChanged,
  GoogleAuthProvider, signInWithPopup, signOut,
  signInWithEmailAndPassword, createUserWithEmailAndPassword
} from 'firebase/auth';
import { getFirestore, doc, setDoc, getDoc, onSnapshot, collection, deleteDoc, getDocFromServer, updateDoc, getDocs } from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';
import { toCanvas } from 'html-to-image';
import { jsPDF } from 'jspdf';

// --- Firebase Configuration ---
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
// Use default database if firestoreDatabaseId is not specified or set to "(default)"
const db = !firebaseConfig.firestoreDatabaseId || firebaseConfig.firestoreDatabaseId === "(default)" 
  ? getFirestore(app) 
  : getFirestore(app, firebaseConfig.firestoreDatabaseId);

// Test Connection as per instructions
async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if(error instanceof Error && error.message.includes('the client is offline')) {
      console.error("Please check your Firebase configuration.");
    }
  }
}
testConnection();

// --- System Configuration (Dynamic Schema) ---
const SCHEMA = {
  certificates: {
    id: 'certificates', name: 'เกียรติบัตร', icon: Award,
    fields: [
      { key: 'title', label: 'ชื่อเกียรติบัตร', type: 'text', required: true },
      { key: 'organization', label: 'หน่วยงานที่ออก', type: 'text', required: true },
      { key: 'location', label: 'สถานที่อบรม', type: 'text' },
      { key: 'date', label: 'วันที่ได้รับ', type: 'date', required: true },
      { key: 'description', label: 'รายละเอียด', type: 'textarea' },
      { key: 'fileUrl', label: 'หลักฐาน', type: 'url' },
    ]
  },
  orders: {
    id: 'orders', name: 'คำสั่งปฏิบัติงาน', icon: FileText,
    fields: [
      { key: 'orderNo', label: 'เลขที่คำสั่ง', type: 'text', required: true },
      { key: 'organization', label: 'หน่วยงาน', type: 'text', required: true },
      { key: 'date', label: 'ลงวันที่', type: 'date', required: true },
      { key: 'details', label: 'รายละเอียดหน้าที่', type: 'textarea', required: true },
      { key: 'fileUrl', label: 'ลิงก์ไฟล์คำสั่ง', type: 'url' },
    ]
  },
  trainings: {
    id: 'trainings', name: 'ประวัติการอบรม', icon: BookOpen,
    fields: [
      { key: 'courseName', label: 'ชื่อหลักสูตร', type: 'text', required: true },
      { key: 'organization', label: 'หน่วยงานผู้จัด', type: 'text', required: true },
      { key: 'dateStr', label: 'ระยะเวลา (เช่น 1-2 ส.ค. 66)', type: 'text', required: true },
      { key: 'hours', label: 'จำนวนชั่วโมง', type: 'number', required: true },
      { key: 'fileUrl', label: 'ลิงก์วุฒิบัตร', type: 'url' },
    ]
  },
  speaker: {
    id: 'speaker', name: 'วิทยากร', icon: Mic,
    collection: 'speakerSessions',
    fields: [
      { key: 'topic', label: 'หัวข้อบรรยาย', type: 'text', required: true },
      { key: 'location', label: 'สถานที่', type: 'text', required: true },
      { key: 'dateStr', label: 'วันที่/ระยะเวลา', type: 'text', required: true },
      { key: 'details', label: 'รายละเอียด', type: 'textarea' },
    ]
  },
  teacherWorks: {
    id: 'teacherWorks', name: 'ผลงานครู', icon: Briefcase,
    fields: [
      { key: 'workName', label: 'ชื่อผลงาน/นวัตกรรม', type: 'text', required: true },
      { key: 'category', label: 'ประเภท', type: 'text', required: true },
      { key: 'details', label: 'รายละเอียด/ผลลัพธ์', type: 'textarea' },
      { key: 'fileUrl', label: 'ลิงก์เอกสาร/เว็บไซต์', type: 'url' },
    ]
  },
  studentWorks: {
    id: 'studentWorks', name: 'ผลงานนักเรียน', icon: Users,
    fields: [
      { key: 'studentName', label: 'ชื่อนักเรียน/กลุ่ม', type: 'text', required: true },
      { key: 'workName', label: 'ชื่อผลงาน/รายการประกวด', type: 'text', required: true },
      { key: 'award', label: 'รางวัลที่ได้รับ', type: 'text', required: true },
      { key: 'year', label: 'ปีการศึกษา', type: 'text', required: true },
      { key: 'fileUrl', label: 'ลิงก์รูปภาพ/เอกสาร', type: 'url' },
    ]
  }
};

const DEFAULT_AVATAR = 'https://ui-avatars.com/api/?name=Profile&background=e2e8f0&color=475569&size=150';

const INITIAL_PROFILE = {
  name: 'ชื่อ-นามสกุล',
  position: 'ตำแหน่ง ครู วิทยฐานะ -',
  school: 'โรงเรียน...',
  bio: 'คำอธิบายประวัติย่อ ปรัชญาการสอน หรือข้อมูลอื่นๆ',
  photoUrl: DEFAULT_AVATAR,
  email: '',
  phone: '',
  googleDriveLink: ''
};

// 50 Themes Collection
const THEME_OPTIONS = [
  { id: 'blue', label: 'น้ำเงิน (ค่าเริ่มต้น)', colorCode: '#2563eb' },
  { id: 'sky', label: 'ฟ้าสว่าง', colorCode: '#0284c7' },
  { id: 'cyan', label: 'ฟ้าคราม', colorCode: '#0891b2' },
  { id: 'teal', label: 'ฟ้าน้ำทะเล', colorCode: '#0d9488' },
  { id: 'emerald', label: 'มรกต', colorCode: '#059669' },
  { id: 'green', label: 'เขียว', colorCode: '#16a34a' },
  { id: 'lime', label: 'เขียวมะนาว', colorCode: '#65a30d' },
  { id: 'olive', label: 'เขียวมะกอก', colorCode: '#6c7c34' },
  { id: 'mint', label: 'มิ้นต์', colorCode: '#148169' },
  { id: 'indigo', label: 'คราม', colorCode: '#4f46e5' },
  { id: 'violet', label: 'ม่วงเข้ม', colorCode: '#7c3aed' },
  { id: 'purple', label: 'ม่วง', colorCode: '#9333ea' },
  { id: 'lavender', label: 'ลาเวนเดอร์', colorCode: '#6a3de6' },
  { id: 'fuchsia', label: 'บานเย็น', colorCode: '#c026d3' },
  { id: 'pink', label: 'ชมพูบานเย็น', colorCode: '#db2777' },
  { id: 'rose', label: 'ชมพู', colorCode: '#e11d48' },
  { id: 'red', label: 'แดง', colorCode: '#dc2626' },
  { id: 'orange', label: 'ส้ม', colorCode: '#ea580c' },
  { id: 'peach', label: 'พีช', colorCode: '#e6451e' },
  { id: 'amber', label: 'เหลืองอำพัน', colorCode: '#d97706' },
  { id: 'yellow', label: 'เหลือง', colorCode: '#ca8a04' },
  { id: 'brown', label: 'น้ำตาล', colorCode: '#977669' },
  { id: 'navy', label: 'กรมท่า', colorCode: '#486581' },
  { id: 'slate', label: 'เทาอมฟ้า', colorCode: '#475569' },
  { id: 'gray', label: 'เทามาตรฐาน', colorCode: '#4b5563' },
  { id: 'zinc', label: 'ซิงค์', colorCode: '#52525b' },
  { id: 'neutral', label: 'เทากลาง', colorCode: '#525252' },
  { id: 'stone', label: 'หิน', colorCode: '#57534e' },
  { id: 'pastelPink', label: 'ชมพูพาสเทล', colorCode: '#f25c94' },
  { id: 'pastelBlue', label: 'ฟ้าพาสเทล', colorCode: '#5e93db' },
  { id: 'pastelGreen', label: 'เขียวพาสเทล', colorCode: '#5ab07d' },
  { id: 'pastelYellow', label: 'เหลืองพาสเทล', colorCode: '#e6c033' },
  { id: 'pastelPurple', label: 'ม่วงพาสเทล', colorCode: '#9666f2' },
  { id: 'pastelPeach', label: 'พีชพาสเทล', colorCode: '#ed753e' },
  { id: 'pastelMint', label: 'มิ้นต์พาสเทล', colorCode: '#41a890' },
  { id: 'babyBlue', label: 'ฟ้าเบบี้บลู', colorCode: '#429bf5' },
  { id: 'blush', label: 'บลัช (ชมพูอ่อน)', colorCode: '#ed6161' },
  { id: 'lilac', label: 'ไลแลค', colorCode: '#b054ed' },
  { id: 'sage', label: 'เขียวเซจ', colorCode: '#7a9175' },
  { id: 'matcha', label: 'มัทฉะ', colorCode: '#94ad51' },
  { id: 'latte', label: 'ลาเต้', colorCode: '#a8876e' },
  { id: 'sand', label: 'ทราย', colorCode: '#ad9e84' },
  { id: 'melon', label: 'เมลอน', colorCode: '#eb6e52' },
  { id: 'ice', label: 'ฟ้าหิมะ', colorCode: '#42b8aa' },
  { id: 'coralSoft', label: 'คอรัลอ่อน', colorCode: '#ed6540' },
  { id: 'mauve', label: 'ม่วงกะปิ', colorCode: '#a374a0' },
  { id: 'periwinkle', label: 'เพอริวิงเคิล', colorCode: '#7287b5' },
  { id: 'seafoam', label: 'ซีโฟม', colorCode: '#46a871' },
  { id: 'butter', label: 'เนยเหลือง', colorCode: '#e6bc2e' },
  { id: 'cottonCandy', label: 'คอตตอนแคนดี้', colorCode: '#c45a9a' }
];

const THEME_PALETTES = {
  blue: null,
  sky: { '50': '#f0f9ff', '100': '#e0f2fe', '200': '#bae6fd', '300': '#7dd3fc', '400': '#38bdf8', '500': '#0ea5e9', '600': '#0284c7', '700': '#0369a1', '800': '#075985', '900': '#0c4a6e' },
  cyan: { '50': '#ecfeff', '100': '#cffafe', '200': '#a5f3fc', '300': '#67e8f9', '400': '#22d3ee', '500': '#06b6d4', '600': '#0891b2', '700': '#0e7490', '800': '#155e75', '900': '#164e63' },
  teal: { '50': '#f0fdfa', '100': '#ccfbf1', '200': '#99f6e4', '300': '#5eead4', '400': '#2dd4bf', '500': '#14b8a6', '600': '#0d9488', '700': '#0f766e', '800': '#115e59', '900': '#134e4a' },
  emerald: { '50': '#ecfdf5', '100': '#d1fae5', '200': '#a7f3d0', '300': '#6ee7b7', '400': '#34d399', '500': '#10b981', '600': '#059669', '700': '#047857', '800': '#065f46', '900': '#064e3b' },
  green: { '50': '#f0fdf4', '100': '#dcfce7', '200': '#bbf7d0', '300': '#86efac', '400': '#4ade80', '500': '#22c55e', '600': '#16a34a', '700': '#15803d', '800': '#166534', '900': '#14532d' },
  lime: { '50': '#f7fee7', '100': '#ecfccb', '200': '#d9f99d', '300': '#bef264', '400': '#a3e635', '500': '#84cc16', '600': '#65a30d', '700': '#4d7c0f', '800': '#3f6212', '900': '#365314' },
  olive: { '50': '#f7f9f2', '100': '#eaf0d8', '200': '#d5e0b5', '300': '#bccb8e', '400': '#a3b56a', '500': '#899b49', '600': '#6c7c34', '700': '#525e24', '800': '#3a4418', '900': '#262d0e' },
  mint: { '50': '#f0fcf9', '100': '#cbf4e9', '200': '#9de7d2', '300': '#6dd4b9', '400': '#42bda0', '500': '#22a084', '600': '#148169', '700': '#0d6653', '800': '#094f41', '900': '#053d33' },
  indigo: { '50': '#eef2ff', '100': '#e0e7ff', '200': '#c7d2fe', '300': '#a5b4fc', '400': '#818cf8', '500': '#6366f1', '600': '#4f46e5', '700': '#4338ca', '800': '#3730a3', '900': '#312e81' },
  violet: { '50': '#f5f3ff', '100': '#ede9fe', '200': '#ddd6fe', '300': '#c4b5fd', '400': '#a78bfa', '500': '#8b5cf6', '600': '#7c3aed', '700': '#6d28d9', '800': '#5b21b6', '900': '#4c1d95' },
  purple: { '50': '#faf5ff', '100': '#f3e8ff', '200': '#e9d5ff', '300': '#d8b4fe', '400': '#c084fc', '500': '#a855f7', '600': '#9333ea', '700': '#7e22ce', '800': '#6b21a8', '900': '#581c87' },
  lavender: { '50': '#f8f4ff', '100': '#ebe0ff', '200': '#d6c2ff', '300': '#bda1ff', '400': '#a17eff', '500': '#845aff', '600': '#6a3de6', '700': '#542ac2', '800': '#421d9e', '900': '#341580' },
  fuchsia: { '50': '#fdf4ff', '100': '#fae8ff', '200': '#f5d0fe', '300': '#f0abfc', '400': '#e879f9', '500': '#d946ef', '600': '#c026d3', '700': '#a21caf', '800': '#86198f', '900': '#701a75' },
  pink: { '50': '#fdf2f8', '100': '#fce7f3', '200': '#fbcfe8', '300': '#f9a8d4', '400': '#f472b6', '500': '#ec4899', '600': '#db2777', '700': '#be185d', '800': '#9d174d', '900': '#831843' },
  rose: { '50': '#fff1f2', '100': '#ffe4e6', '200': '#fecdd3', '300': '#fda4af', '400': '#fb7185', '500': '#f43f5e', '600': '#e11d48', '700': '#be123c', '800': '#9f1239', '900': '#881337' },
  red: { '50': '#fef2f2', '100': '#fee2e2', '200': '#fecaca', '300': '#fca5a5', '400': '#f87171', '500': '#ef4444', '600': '#dc2626', '700': '#b91c1c', '800': '#991b1b', '900': '#7f1d1d' },
  orange: { '50': '#fff7ed', '100': '#ffedd5', '200': '#fed7aa', '300': '#fdba74', '400': '#fb923c', '500': '#f97316', '600': '#ea580c', '700': '#c2410c', '800': '#9a3412', '900': '#7c2d12' },
  peach: { '50': '#fff8f6', '100': '#ffe9e0', '200': '#ffd1be', '300': '#ffb094', '400': '#ff8a66', '500': '#ff6339', '600': '#e6451e', '700': '#bf3313', '800': '#99260d', '900': '#7a1d09' },
  amber: { '50': '#fffbeb', '100': '#fef3c7', '200': '#fde68a', '300': '#fcd34d', '400': '#fbbf24', '500': '#f59e0b', '600': '#d97706', '700': '#b45309', '800': '#92400e', '900': '#78350f' },
  yellow: { '50': '#fefce8', '100': '#fef9c3', '200': '#fef08a', '300': '#fde047', '400': '#facc15', '500': '#eab308', '600': '#ca8a04', '700': '#a16207', '800': '#854d0e', '900': '#713f12' },
  brown: { '50': '#fdf8f6', '100': '#f2e8e5', '200': '#eaddd7', '300': '#e0cec7', '400': '#d2bab0', '500': '#a18072', '600': '#977669', '700': '#846358', '800': '#43302b', '900': '#2a1a17' },
  navy: { '50': '#f0f4f8', '100': '#d9e2ec', '200': '#bcccdc', '300': '#9fb3c8', '400': '#829ab1', '500': '#627d98', '600': '#486581', '700': '#334e68', '800': '#243b53', '900': '#102a43' },
  slate: { '50': '#f8fafc', '100': '#f1f5f9', '200': '#e2e8f0', '300': '#cbd5e1', '400': '#94a3b8', '500': '#64748b', '600': '#475569', '700': '#334155', '800': '#1e293b', '900': '#0f172a' },
  gray: { '50': '#f9fafb', '100': '#f3f4f6', '200': '#e5e7eb', '300': '#d1d5db', '400': '#9ca3af', '500': '#6b7280', '600': '#4b5563', '700': '#374151', '800': '#1f2937', '900': '#111827' },
  zinc: { '50': '#fafafa', '100': '#f4f4f5', '200': '#e4e4e7', '300': '#d4d4d8', '400': '#a1a1aa', '500': '#71717a', '600': '#52525b', '700': '#3f3f46', '800': '#27272a', '900': '#18181b' },
  neutral: { '50': '#fafafa', '100': '#f5f5f5', '200': '#e5e5e5', '300': '#d4d4d4', '400': '#a3a3a3', '500': '#737373', '600': '#525252', '700': '#404040', '800': '#262626', '900': '#171717' },
  stone: { '50': '#fafaf9', '100': '#f5f5f4', '200': '#e7e5e4', '300': '#d6d3d1', '400': '#a8a29e', '500': '#78716c', '600': '#57534e', '700': '#44403c', '800': '#292524', '900': '#1c1917' },
  pastelPink: { '50': '#fff5f8', '100': '#ffe6ee', '200': '#ffcce0', '300': '#ffa3c7', '400': '#ff75a8', '500': '#f25c94', '600': '#d64278', '700': '#b52a5c', '800': '#941b44', '900': '#7a1033' },
  pastelBlue: { '50': '#f2f7fc', '100': '#e0effa', '200': '#c2e0f5', '300': '#9ec8f0', '400': '#7aabeb', '500': '#5e93db', '600': '#4377c2', '700': '#2f5ca3', '800': '#204482', '900': '#142e63' },
  pastelGreen: { '50': '#f2fcf5', '100': '#e0f7e8', '200': '#c0ecd0', '300': '#9cdbb4', '400': '#77c796', '500': '#5ab07d', '600': '#409462', '700': '#2c7a4d', '800': '#1e613b', '900': '#124729' },
  pastelYellow: { '50': '#fffdf2', '100': '#fff9d6', '200': '#fff2a8', '300': '#ffe775', '400': '#fcd84e', '500': '#e6c033', '600': '#c7a11c', '700': '#a3810f', '800': '#806307', '900': '#614903' },
  pastelPurple: { '50': '#f9f5ff', '100': '#f0e6ff', '200': '#dfccff', '300': '#c8abff', '400': '#ae85ff', '500': '#9666f2', '600': '#7d4ad6', '700': '#6332b5', '800': '#4c2291', '900': '#381570' },
  pastelPeach: { '50': '#fff8f2', '100': '#ffebe0', '200': '#ffd2ba', '300': '#ffb48f', '400': '#ff9261', '500': '#ed753e', '600': '#cc5621', '700': '#a83e11', '800': '#852d08', '900': '#661f03' },
  pastelMint: { '50': '#f0fcf9', '100': '#d6f7ee', '200': '#b0ebdc', '300': '#85dbc5', '400': '#5dc4ab', '500': '#41a890', '600': '#2b8c75', '700': '#1a705b', '800': '#105745', '900': '#084031' },
  babyBlue: { '50': '#f0f9ff', '100': '#dff0ff', '200': '#c2e3ff', '300': '#99d0ff', '400': '#6bb8ff', '500': '#429bf5', '600': '#287bd6', '700': '#175eb3', '800': '#0e458c', '900': '#073069' },
  blush: { '50': '#fff5f5', '100': '#ffe6e6', '200': '#ffcaca', '300': '#ffaaaa', '400': '#fa8484', '500': '#ed6161', '600': '#cc4343', '700': '#a82d2d', '800': '#851d1d', '900': '#661010' },
  lilac: { '50': '#fcf5ff', '100': '#f6e6ff', '200': '#ebc7ff', '300': '#dda3ff', '400': '#c97aff', '500': '#b054ed', '600': '#9336cc', '700': '#7620a8', '800': '#5a1385', '900': '#430a66' },
  sage: { '50': '#f5f7f4', '100': '#e6ede4', '200': '#ced9ca', '300': '#b0c2ac', '400': '#93a88e', '500': '#7a9175', '600': '#5f785b', '700': '#475e43', '800': '#334530', '900': '#22301f' },
  matcha: { '50': '#f8faf2', '100': '#edf2dd', '200': '#dae6ba', '300': '#c3d692', '400': '#acc46e', '500': '#94ad51', '600': '#789139', '700': '#5d7326', '800': '#455717', '900': '#303d0d' },
  latte: { '50': '#fcfaf9', '100': '#f5efe9', '200': '#e6d8cb', '300': '#d4beaa', '400': '#bfa18a', '500': '#a8876e', '600': '#8c6a51', '700': '#70503a', '800': '#573b28', '900': '#40291a' },
  sand: { '50': '#fcfbfb', '100': '#f5f3ef', '200': '#e8e2d8', '300': '#d6cdbc', '400': '#c2b59f', '500': '#ad9e84', '600': '#918166', '700': '#75654c', '800': '#5c4e38', '900': '#453826' },
  melon: { '50': '#fff7f5', '100': '#ffece6', '200': '#ffd4c9', '300': '#ffb5a3', '400': '#fa9178', '500': '#eb6e52', '600': '#c94f36', '700': '#a63620', '800': '#822312', '900': '#631508' },
  ice: { '50': '#f0fdfc', '100': '#dcfaf6', '200': '#bdf2ec', '300': '#94e6dc', '400': '#65d4c7', '500': '#42b8aa', '600': '#2b998d', '700': '#1c7a6f', '800': '#125e55', '900': '#0a453e' },
  coralSoft: { '50': '#fff5f2', '100': '#ffe7e0', '200': '#ffcbb8', '300': '#ffaa8f', '400': '#fa8664', '500': '#ed6540', '600': '#cc4623', '700': '#a83114', '800': '#85220c', '900': '#661605' },
  mauve: { '50': '#faf7fa', '100': '#f3ebf2', '200': '#e3d1e1', '300': '#cfb2cd', '400': '#ba91b7', '500': '#a374a0', '600': '#875884', '700': '#6b4168', '800': '#522d4f', '900': '#3b1e39' },
  periwinkle: { '50': '#f5f7fa', '100': '#e6ecf5', '200': '#cdd8e8', '300': '#aebedd', '400': '#8ea2ce', '500': '#7287b5', '600': '#586b99', '700': '#42527a', '800': '#2f3c5c', '900': '#1f2840' },
  seafoam: { '50': '#f2fcf7', '100': '#e0f7e9', '200': '#bcebd0', '300': '#92d9b1', '400': '#67c48f', '500': '#46a871', '600': '#2f8c57', '700': '#1f7041', '800': '#145731', '900': '#0a4021' },
  butter: { '50': '#fffef2', '100': '#fffce0', '200': '#fff6ba', '300': '#ffe98a', '400': '#fcd756', '500': '#e6bc2e', '600': '#c79d1a', '700': '#a37c0f', '800': '#805e08', '900': '#614404' },
  cottonCandy: { '50': '#fcf5f9', '100': '#f7e6f0', '200': '#f0c7e0', '300': '#e6a3cc', '400': '#d97cb5', '500': '#c45a9a', '600': '#a83f7f', '700': '#872863', '800': '#69184a', '900': '#4d0d35' }
};

const CustomThemeStyle = ({ theme }) => {
  const colors = THEME_PALETTES[theme];
  if (!colors) return null;

  return (
    <style>{`
      .bg-blue-50 { background-color: ${colors['50']} !important; }
      .bg-blue-100 { background-color: ${colors['100']} !important; }
      .bg-blue-500 { background-color: ${colors['500']} !important; }
      .bg-blue-600 { background-color: ${colors['600']} !important; }
      .bg-blue-800 { background-color: ${colors['800']} !important; }
      .bg-blue-900 { background-color: ${colors['900']} !important; }
      
      .hover\\:bg-blue-50:hover { background-color: ${colors['50']} !important; }
      .hover\\:bg-blue-100:hover { background-color: ${colors['100']} !important; }
      .hover\\:bg-blue-700:hover { background-color: ${colors['700']} !important; }
      .hover\\:bg-blue-800:hover { background-color: ${colors['800']} !important; }
      
      .text-blue-100 { color: ${colors['100']} !important; }
      .text-blue-200 { color: ${colors['200']} !important; }
      .text-blue-300 { color: ${colors['300']} !important; }
      .text-blue-400 { color: ${colors['400']} !important; }
      .text-blue-500 { color: ${colors['500']} !important; }
      .text-blue-600 { color: ${colors['600']} !important; }
      .text-blue-700 { color: ${colors['700']} !important; }
      .text-blue-800 { color: ${colors['800']} !important; }
      .text-blue-900 { color: ${colors['900']} !important; }
      
      .border-blue-50 { border-color: ${colors['50']} !important; }
      .border-blue-100 { border-color: ${colors['100']} !important; }
      .border-blue-500 { border-color: ${colors['500']} !important; }
      .border-blue-800 { border-color: ${colors['800']} !important; }
      .border-blue-900 { border-color: ${colors['900']} !important; }
      
      .focus\\:ring-blue-500:focus { --tw-ring-color: ${colors['500']} !important; }
      .focus\\:border-blue-500:focus { border-color: ${colors['500']} !important; }
      
      .peer:checked ~ .peer-checked\\:bg-blue-600 { background-color: ${colors['600']} !important; }
    `}</style>
  );
};

// --- Helper Functions ---
const getDirectImageUrl = (url) => {
  if (!url) return '';
  try {
    if (url.includes('drive.google.com')) {
      const match = url.match(/(?:file\/d\/|id=)([\w-]+)/);
      if (match && match[1]) {
        // ใช้ Thumbnail API ที่มีความแน่นอนในการดึงภาพมากที่สุดและฝ่าการบล็อคของเบราว์เซอร์ได้ดี
        // ตั้ง sz=w2048-h2048 เพื่อให้ภาพมีความคมชัดสูงที่สุด
        return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w2048-h2048`;
      }
    }
  } catch(e) {}
  return url;
};

const formatThaiDate = (dateString) => {
  if (!dateString) return '-';
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return dateString; 

  const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  const d = date.getDate();
  const m = thaiMonths[date.getMonth()];
  const y = date.getFullYear() + 543;
  return `${d} ${m} ${y}`;
};

// --- New: Year Extraction Helper ---
const getItemYear = (item) => {
  if (item.year) return item.year;
  if (item.date) {
    const d = new Date(item.date);
    if (!isNaN(d.getTime())) {
      return (d.getFullYear() + 543).toString();
    }
  }
  if (item.dateStr) {
    const match = item.dateStr.match(/(\d{2,4}$)/);
    if (match) {
      const y = match[1];
      if (y.length === 2) return "25" + y;
      return y;
    }
  }
  return null;
};

// --- New: Get Unique Years for Category ---
const getUniqueYears = (items) => {
  if (!items || !Array.isArray(items)) return [];
  const years = new Set();
  items.forEach(item => {
    const y = getItemYear(item);
    if (y) years.add(y);
  });
  return Array.from(years).sort((a, b) => b.localeCompare(a));
};

const getSortedData = (data, category, filterYear = 'all') => {
  if (!data || !Array.isArray(data)) return [];
  
  let filteredData = [...data];
  if (filterYear !== 'all') {
    filteredData = filteredData.filter(item => getItemYear(item) === filterYear);
  }

  const dateField = category.fields.find(f => f.type === 'date');
  
  return filteredData.sort((a, b) => {
    if (dateField) {
      const key = dateField.key;
      const dateA = a[key] ? new Date(a[key]).getTime() : 0;
      const dateB = b[key] ? new Date(b[key]).getTime() : 0;
      if (dateA !== dateB) return dateB - dateA; 
    }
    return Number(b.id) - Number(a.id); 
  });
};

export default function App() {
  const [isAdminAuthenticated, setIsAdminAuthenticated] = useState(false); 
  const [passwordInput, setPasswordInput] = useState('');
  const [emailInput, setEmailInput] = useState('');
  const [isSignUpMode, setIsSignUpMode] = useState(false);
  const [loginError, setLoginError] = useState('');

  const [user, setUser] = useState(null);
  const [authErrorCode, setAuthErrorCode] = useState(null);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('profile'); // Default to profile for new users
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [viewMode, setViewMode] = useState('admin'); 
  const [showLanding, setShowLanding] = useState(true); // เริ่มต้นที่หน้าเข้าสู่ระบบ
  const [publicUserUid, setPublicUserUid] = useState(null); // สำหรับดู Portfolio ของคนอื่นผ่านลิงก์
  const [systemStats, setSystemStats] = useState({ totalUsers: 0, recentUsers: [] });
  const [isStatsLoading, setIsStatsLoading] = useState(false);
  
  const [profile, setProfile] = useState(INITIAL_PROFILE);
  const [portfolioData, setPortfolioData] = useState({
    certificates: [], orders: [], trainings: [], speaker: [], teacherWorks: [], studentWorks: []
  });
  const [settings, setSettings] = useState({
    showCertificates: true, showOrders: true, showTrainings: true, 
    showSpeaker: true, showTeacherWorks: true, showStudentWorks: true,
    themeColor: 'blue'
  });

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [currentCategory, setCurrentCategory] = useState(null);
  const [formData, setFormData] = useState({});
  const [editingId, setEditingId] = useState(null);
  const [saveStatus, setSaveStatus] = useState('');
  const [isPrinting, setIsPrinting] = useState(false);
  const [isPrintingCategory, setIsPrintingCategory] = useState(false);
  const [isUploadingImage, setIsUploadingImage] = useState(false);
  const [showImageTools, setShowImageTools] = useState(false);
  const [printModalState, setPrintModalState] = useState(null);
  const [yearFilter, setYearFilter] = useState('all'); 
  const [deleteConfirm, setDeleteConfirm] = useState(null); // { categoryKey, id, title }

  useEffect(() => {
    // ตรวจสอบพารามิเตอร์ u ใน URL เพื่อแสดงผลหน้าสาธารณะของผู้อื่น
    const params = new URLSearchParams(window.location.search);
    const u = params.get('u');
    if (u) {
      setPublicUserUid(u);
      setViewMode('public');
      setShowLanding(false);
    }

    const savedAuth = sessionStorage.getItem('isAdminAuthenticated');
    if (savedAuth === 'true') {
      setIsAdminAuthenticated(true);
      setShowLanding(false);
      setUser({ uid: 'admin-master', email: 'admin@eportfolio.local', isRealFirebaseUser: false });
    }

    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          // สั่ง Login แบบ Anonymous
          await signInAnonymously(auth);
        }
      } catch (error) {
        console.error("Auth error details:", error);
        setAuthErrorCode(error.code);
        if (error.code === 'auth/admin-restricted-operation') {
          console.warn("Anonymous Auth is disabled in Firebase Console.");
          setUser(prev => prev?.isRealFirebaseUser ? prev : { uid: 'guest-user', isGuest: true, authError: error.code });
        } else if (error.code === 'auth/unauthorized-domain') {
          console.error("This domain is not authorized in Firebase Console.");
          setUser(prev => prev?.isRealFirebaseUser ? prev : { uid: 'guest-user', isGuest: true, authError: error.code });
        } else {
          console.error("Auth error message:", error.message);
          setUser(prev => prev?.isRealFirebaseUser ? prev : { uid: 'guest-user', isGuest: true, authError: error.code });
        }
        setLoading(false);
      }
    };
    initAuth();

    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      const savedAuth = sessionStorage.getItem('isAdminAuthenticated');
      const isMasterMode = savedAuth === 'true';

      if (currentUser && !currentUser.isAnonymous) {
        // 1. กรณีเป็นผู้ใช้จริง (Google Account)
        console.log("Logged in as Google User:", currentUser.email);
        setUser(currentUser);
        setIsAdminAuthenticated(true);
        setShowLanding(false);
        setViewMode('admin');
        // เคลียร์ค่า Master Admin ออกถ้ามี เพื่อไม่ให้สับสน
        if (isMasterMode) sessionStorage.removeItem('isAdminAuthenticated');
      } 
      else if (isMasterMode) {
        // 2. กรณีเป็น Master Admin (ล็อคอินด้วยรหัสผ่าน)
        console.log("Logged in as Master Admin");
        setUser({ 
          uid: 'admin-master', 
          email: 'admin@eportfolio.local',
          isMasterAdmin: true,
          isGuest: false,
          displayName: 'Admin'
        });
        setIsAdminAuthenticated(true);
        setShowLanding(false);
        setViewMode('admin');
      } 
      else {
        // 3. กรณีไม่ได้เข้าสู่ระบบ หรือเป็น Guest (Anonymous)
        console.log("User state:", currentUser ? "Guest" : "Not logged in");
        setUser(currentUser);
        setIsAdminAuthenticated(false);
        setShowLanding(true);
      }
      
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Fetch system stats for Master Admin
  useEffect(() => {
    if (user?.isMasterAdmin && activeTab === 'system-stats') {
      const fetchStats = async () => {
        setIsStatsLoading(true);
        try {
          const db = getFirestore();
          const usersCol = collection(db, 'users');
          const snapshot = await getDocs(usersCol);
          const users = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
          
          setSystemStats({
            totalUsers: users.length,
            recentUsers: users.slice(-10).reverse() // Show last 10 users
          });
        } catch (error) {
          console.error("Error fetching stats:", error);
        } finally {
          setIsStatsLoading(false);
        }
      };
      fetchStats();
    }
  }, [user, activeTab]);

  const handleGoogleLogin = async () => {
    const provider = new GoogleAuthProvider();
    // บังคับให้เลือก Account ใหม่เสมอเพื่อป้องกันการค้าง
    provider.setCustomParameters({ prompt: 'select_account' });
    
    try {
      setLoading(true);
      // เคลียร์สถานะเก่าก่อนเริ่มใหม่
      setIsAdminAuthenticated(false);
      sessionStorage.removeItem('isAdminAuthenticated');
      
      const result = await signInWithPopup(auth, provider);
      if (result.user) {
        // Listener (onAuthStateChanged) จะทำหน้าที่เปลี่ยน State ให้เอง 
        // แต่ใส่ไว้ตรงนี้เพื่อความรวดเร็วในการตอบสนอง (UI Optimistic Update)
        setIsAdminAuthenticated(true);
        setShowLanding(false);
        setViewMode('admin');
      }
    } catch (error) {
      console.error("Google Login Error:", error);
      setAuthErrorCode(error.code);
      if (error.code === 'auth/unauthorized-domain') {
        alert("โดเมนนี้ยังไม่ได้รับอนุญาตใน Firebase! กรุณาเพิ่มโดเมนใน Authorized Domains");
      } else if (error.code === 'auth/operation-not-allowed') {
        alert("ข้อผิดพลาด: ยังไม่ได้เปิดใช้งาน Google Sign-in ใน Firebase Console\n\nวิธีแก้:\n1. ไปที่ Firebase Console > Authentication > Sign-in method\n2. เพิ่ม Google และกด Enable");
      } else if (error.code === 'auth/popup-blocked') {
        alert("ป๊อปอัพถูกบล็อก! กรุณาอนุญาตให้เปิดป๊อปอัพสำหรับเว็บนี้");
      } else if (error.code === 'auth/popup-closed-by-user') {
        // ผู้ใช้ปิดหน้าต่างเอง ไม่ต้องทำอะไร
      } else {
        alert("เข้าสู่ระบบไม่สำเร็จ: " + error.message);
      }
    } finally {
      // อย่าเพิ่งปิด Loading ทันทีเพื่อให้ Listener ทำงานเสร็จก่อน
      setTimeout(() => setLoading(false), 500);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
      setIsAdminAuthenticated(false);
      sessionStorage.removeItem('isAdminAuthenticated');
      setShowLanding(true);
      setUser(null);
      setPasswordInput('');
      setProfile(INITIAL_PROFILE);
      setPortfolioData({
        certificates: [], orders: [], trainings: [], speaker: [], teacherWorks: [], studentWorks: []
      });
    } catch (error) {
      console.error("Logout Error:", error);
    }
  };

  useEffect(() => {
    // ถ้ามี publicUserUid และอยู่ในโหมด public ให้ดึงข้อมูลของคนนั้น
    // ถ้าไม่มี ให้ดึงข้อมูลของผู้ใช้ที่ล็อคอินอยู่
    const targetUid = (viewMode === 'public' && publicUserUid) ? publicUserUid : user?.uid;
    if (!targetUid) return;

    // Fetch Profile & Settings
    const userRef = doc(db, 'users', targetUid);
    const unsubProfile = onSnapshot(userRef, (docSnap) => {
      // ป้องกันการคืนค่าเดิมขณะที่คุณกำลังลาก Slider (ถ้าข้อมูลมาจากเครื่องเราเอง ไม่ต้องอัปเดตซ้ำ)
      if (docSnap.metadata.hasPendingWrites) return;

      if (docSnap.exists()) {
        const data = docSnap.data();
        setProfile(prev => ({
          ...INITIAL_PROFILE,
          ...prev, // เก็บค่าปัจจุบันไว้
          ...data  // ทับด้วยค่าจาก Server
        }));
        if (data.settings) setSettings(prev => ({ ...prev, ...data.settings }));
      } else if (user && user.uid === targetUid && !user.isGuest && !user.isMasterAdmin) {
        // กรณีเป็นผู้ใช้ใหม่ (ไม่มีข้อมูลใน Firestore) ให้ใช้ข้อมูลจาก Google Profile เบื้องต้น
        setProfile(prev => ({
          ...INITIAL_PROFILE,
          name: user.displayName || prev.name,
          email: user.email || prev.email,
          photoUrl: user.photoURL || prev.photoUrl
        }));
      }
    }, (error) => {
      console.error("Error fetching profile:", error);
    });

    // Fetch Items for each category in SCHEMA
    const unsubs = Object.keys(SCHEMA).map(key => {
      const colName = SCHEMA[key].collection || key;
      const colRef = collection(db, 'users', targetUid, colName);
      return onSnapshot(colRef, (snapshot) => {
        const items = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setPortfolioData(prev => ({ ...prev, [key]: items }));
      }, (error) => {
        console.error(`Error fetching ${key}:`, error);
      });
    });

    return () => {
      unsubProfile();
      unsubs.forEach(unsub => unsub());
    };
  }, [user]);

  useEffect(() => {
    if (!document.getElementById('kanit-font-style')) {
      const style = document.createElement('style');
      style.id = 'kanit-font-style';
      style.innerHTML = `
        @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@200;300;400;500;600;700&display=swap');
        body, button, input, textarea, select, span, p, h1, h2, h3, h4, h5, h6, th, td {
          font-family: 'Kanit', sans-serif !important;
        }
      `;
      document.head.appendChild(style);
    }
  }, []);

  const handleFirestoreError = (error, operationType, path = null) => {
    const errorInfo = {
      error: error.message,
      operationType,
      path,
      authInfo: {
        userId: auth.currentUser?.uid || 'guest-user',
        email: auth.currentUser?.email || 'N/A',
        emailVerified: auth.currentUser?.emailVerified || false,
        isAnonymous: auth.currentUser?.isAnonymous || false,
        providerInfo: auth.currentUser?.providerData?.map(p => ({
          providerId: p.providerId,
          displayName: p.displayName,
          email: p.email
        })) || []
      }
    };
    console.error("Firestore Error Detailed:", errorInfo);
    return JSON.stringify(errorInfo, null, 2);
  };

  const saveProfileToFirebase = async (newProfile, newSettings) => {
    if (!user || !isAdminAuthenticated) return;
    if (user.isGuest) {
      alert("ไม่สามารถบันทึกได้ในโหมดผู้เยี่ยมชม กรุณารีเฟรชหน้าเว็บหรือตรวจสอบสถานะการเชื่อมต่อ");
      return;
    }
    setSaveStatus('กำลังบันทึก...');
    try {
      const docRef = doc(db, 'users', user.uid);
      await setDoc(docRef, {
        ...(newProfile || profile),
        settings: newSettings || settings,
        updatedAt: new Date().toISOString()
      }, { merge: true });
      setSaveStatus('บันทึกสำเร็จ');
      setTimeout(() => setSaveStatus(''), 2000);
    } catch (error) {
      console.error("Save error:", error);
      if (error.code === 'permission-denied') {
        handleFirestoreError(error, 'update', `users/${user.uid}`);
        setSaveStatus('ไม่มีสิทธิ์บันทึก');
      } else {
        setSaveStatus('เกิดข้อผิดพลาด');
      }
    }
  };

  const handleProfileChange = (e) => {
    if (!isAdminAuthenticated) return;
    const { name, value, type } = e.target;
    
    let finalValue = value;
    if (type === 'range' || (['imageSize', 'imageZoom', 'imagePanX', 'imagePanY', 'imageBorderWidth'].includes(name))) {
      finalValue = parseFloat(value);
    }
    
    const newProfile = { ...profile, [name]: finalValue };
    setProfile(newProfile);
    
    // ใช้ Debounce สำหรับการบันทึกลง Firebase เพื่อป้องกัน Error: resource-exhausted
    // หากเป็นการเปลี่ยนข้อความทั่วไปบันทึกช้าหน่อยได้ หากเป็น Slider ให้รอจนกว่าจะหยุดเลื่อน
    if (window.profileSaveTimeout) clearTimeout(window.profileSaveTimeout);
    window.profileSaveTimeout = setTimeout(() => {
      saveProfileToFirebase(newProfile, null);
    }, 1000); // รอ 1 วินาทีก่อนบันทึกจริง
  };

  const handleImageUpload = (e) => {
    if (!isAdminAuthenticated) return;
    const file = e.target.files[0];
    if (!file) return;

    setIsUploadingImage(true);
    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const MAX_SIZE = 1000;
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > MAX_SIZE) {
            height *= MAX_SIZE / width;
            width = MAX_SIZE;
          }
        } else {
          if (height > MAX_SIZE) {
            width *= MAX_SIZE / height;
            height = MAX_SIZE;
          }
        }

        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);

        // Convert to base64 with high quality JPEG to preserve crispness but reduce size
        const dataUrl = canvas.toDataURL('image/jpeg', 0.9);

        const newProfile = { ...profile, photoUrl: dataUrl };
        setProfile(newProfile);
        saveProfileToFirebase(newProfile, null);
        setIsUploadingImage(false);
      };
      img.src = event.target.result;
    };
    reader.readAsDataURL(file);
  };

  const handleSettingToggle = (settingKey) => {
    if (!isAdminAuthenticated) return;
    const currentValue = settings[settingKey] !== false; 
    const newSettings = { ...settings, [settingKey]: !currentValue };
    setSettings(newSettings);
    saveProfileToFirebase(null, newSettings);
  };

  const handleThemeChange = (themeId) => {
    if (!isAdminAuthenticated) return;
    const newSettings = { ...settings, themeColor: themeId };
    setSettings(newSettings);
    saveProfileToFirebase(null, newSettings);
  };

  const openModal = (categoryKey, item = null) => {
    if (!isAdminAuthenticated) return;
    setCurrentCategory(categoryKey);
    if (item) {
      setFormData(item);
      setEditingId(item.id);
    } else {
      setFormData({});
      setEditingId(null);
    }
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
    setFormData({});
    setEditingId(null);
    setCurrentCategory(null);
  };

  const handleFormChange = (e, key) => {
    setFormData({ ...formData, [key]: e.target.value });
  };

  const handleFormSubmit = async (e) => {
    e.preventDefault();
    if (!user) return;

    // ถ้าเป็นโหมด Admin ไม่บล็อคการบันทึก
    if (user.isAnonymous && !user.isMasterAdmin) {
      alert("คุณกำลังใช้งานในโหมดทดลอง ข้อมูลจะไม่ถูกบันทึกอย่างถาวรจนกว่าจะล็อคอินด้วย Google");
      return;
    }
    
    setSaveStatus('กำลังบันทึก...');
    try {
      if (!user?.uid) throw new Error("User ID is missing");
      console.log("Saving UID:", user.uid);
      console.log("Current Auth UID:", auth.currentUser?.uid);
      console.log("Saving to path:", `users/${user.uid}/${currentCategory}/${editingId || 'new'}`);
      
      const colName = SCHEMA[currentCategory].collection || currentCategory;
      const itemId = editingId || Date.now().toString();
      const docRef = doc(db, 'users', user.uid, colName, itemId);
      
      const payload = { ...formData };
      delete payload.addToCertificates;
      payload.updatedAt = new Date().toISOString();
      if (!editingId) payload.createdAt = new Date().toISOString();

      await setDoc(docRef, {
        ...payload,
        updatedAt: new Date().toISOString()
      }, { merge: true });

      // Auto-copy to certificates if checked
      if (formData.addToCertificates && currentCategory !== 'certificates') {
        const certId = (Date.now() + 1).toString();
        const certRef = doc(db, 'users', user.uid, 'certificates', certId);
        
        const title = formData.courseName || formData.workName || formData.topic || formData.orderNo || formData.title || `เกียรติบัตรจาก ${SCHEMA[currentCategory].name}`;
        const org = formData.organization || formData.location || '-';
        
        let dateValue = new Date().toISOString().split('T')[0];
        if (formData.date && !formData.dateStr) dateValue = formData.date;
        
        let desc = formData.details || formData.description || `เพิ่มอัตโนมัติจากแฟ้ม ${SCHEMA[currentCategory].name}`;
        if (currentCategory === 'studentWorks') desc = `รางวัล: ${formData.award || '-'} ปีการศึกษา: ${formData.year || '-'}`;
        if (currentCategory === 'trainings') desc = `ระยะเวลา: ${formData.dateStr || '-'} (${formData.hours || '0'} ชั่วโมง)`;

        await setDoc(certRef, {
          title: title,
          organization: org,
          location: formData.location || '',
          date: dateValue,
          description: desc,
          fileUrl: formData.fileUrl || '',
          createdAt: payload.createdAt || new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });
      }

      setSaveStatus('บันทึกสำเร็จ');
      setTimeout(() => setSaveStatus(''), 2000);
      closeModal();
    } catch (error) {
      console.error("Submit error:", error);
      if (error.code === 'permission-denied') {
        handleFirestoreError(error, 'write', `users/${user.uid}/${currentCategory}`);
        setSaveStatus('ไม่มีสิทธิ์บันทึก');
      } else {
        setSaveStatus('เกิดข้อผิดพลาด');
      }
    }
  };

  const handleDelete = async (categoryKey, id, title = 'ข้อมูลนี้') => {
    if (!user || user.isGuest) {
      alert('คุณอยู่ในโหมดผู้เยี่ยมชม ไม่สามารถลบข้อมูลได้ กรุณาเข้าสู่ระบบจริง');
      return;
    }
    if (!isAdminAuthenticated) return;
    
    setDeleteConfirm({ categoryKey, id, title });
  };

  const confirmDelete = async () => {
    if (!deleteConfirm) return;
    const { categoryKey, id } = deleteConfirm;
    
    setSaveStatus('กำลังลบ...');
    setDeleteConfirm(null);
    
    try {
      const colName = SCHEMA[categoryKey].collection || categoryKey;
      const docRef = doc(db, 'users', user.uid, colName, id);
      await deleteDoc(docRef);
      setSaveStatus('ลบสำเร็จ');
      setTimeout(() => setSaveStatus(''), 2000);
    } catch (error) {
      console.error("Delete error:", error);
      let errorMsg = 'เกิดข้อผิดพลาดในการลบ';
      if (error.code === 'permission-denied') {
        errorMsg = 'ไม่มีสิทธิ์ในการลบ (Permission Denied)';
      }
      setSaveStatus(errorMsg);
      alert(errorMsg);
      setTimeout(() => setSaveStatus(''), 4000);
    }
  };

  const handleEmailAuth = async (e) => {
    e.preventDefault();
    setLoginError('');
    setLoading(true);
    try {
      if (isSignUpMode) {
        await createUserWithEmailAndPassword(auth, emailInput, passwordInput);
      } else {
        await signInWithEmailAndPassword(auth, emailInput, passwordInput);
      }
      setShowLanding(false);
    } catch (error) {
      console.error("Email Auth Error:", error);
      let msg = 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ';
      if (error.code === 'auth/email-already-in-use') msg = 'อีเมลนี้ถูกใช้งานแล้ว';
      if (error.code === 'auth/invalid-email') msg = 'รูปแบบอีเมลไม่ถูกต้อง';
      if (error.code === 'auth/user-not-found' || error.code === 'auth/wrong-password') msg = 'อีเมลหรือรหัสผ่านไม่ถูกต้อง';
      if (error.code === 'auth/weak-password') msg = 'รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร';
      setLoginError(msg);
    } finally {
      setLoading(false);
    }
  };

  const handleLogin = async (e) => {
    e.preventDefault();
    if (passwordInput === 'admin0000') {
        // Sign out any firebase user to ensure clean admin mode
        if (auth.currentUser) await signOut(auth);
        
        setIsAdminAuthenticated(true);
        sessionStorage.setItem('isAdminAuthenticated', 'true');
        setShowLanding(false);
        setViewMode('admin');
        setLoginError('');
        
        setUser({ 
          uid: 'admin-master', 
          email: 'admin@eportfolio.local', 
          isRealFirebaseUser: true, 
          isMasterAdmin: true,
          isGuest: false 
        });
      } else {
        setLoginError('รหัสผ่าน Admin ไม่ถูกต้อง');
      }
  };

  // ----------------------------------------------------
  // ระบบสร้าง PDF (หน้าภาพรวม)
  // ----------------------------------------------------
  const handlePrintPDF = async (orientation = 'portrait') => {
    setIsPrinting(true);
    
    // เลื่อนขึ้นบนสุดเพื่อกันปัญหาภาพแหว่ง
    window.scrollTo(0, 0); 
    
    let element = document.getElementById('portfolio-content');
    
    // หากไม่พบ และเราไม่ได้อยู่ในหน้า Preview ให้ลองสลับไปหน้า Preview ก่อน
    if (!element && viewMode !== 'public') {
      setViewMode('public');
      // รอให้ React Render สักพักแล้วลองใหม่
      setTimeout(() => handlePrintPDF(orientation), 500);
      return;
    }
    
    try {
      if (!element) throw new Error("ไม่พบเนื้อหาที่ต้องการพิมพ์");
      
      const elWidth = element.offsetWidth || element.clientWidth || 800;
      const elHeight = element.offsetHeight || element.clientHeight || 1200;

      // ใช้ html-to-image จับภาพเป็น JPEG ตรงๆ
      const { toJpeg } = await import('html-to-image');
      const imgData = await toJpeg(element, { 
        quality: 0.98, 
        backgroundColor: '#ffffff',
        pixelRatio: 2,
        width: elWidth,
        height: elHeight
      });
      
      // สร้างหน้ากระดาษ A4
      const pdf = new jsPDF({
        orientation: orientation,
        unit: 'mm',
        format: 'a4'
      });
      
      const margin = 10;
      const pdfWidth = pdf.internal.pageSize.getWidth() - (margin * 2);
      const pdfHeight = pdf.internal.pageSize.getHeight() - (margin * 2);

      // คำนวณความสูงเมื่อย่อขยายเข้ากับหน้า A4
      const ratio = pdfWidth / elWidth;
      const scaledHeight = elHeight * ratio;

      if (!isFinite(scaledHeight) || scaledHeight <= 0) {
          throw new Error(`การคำนวณขนาดผิดพลาด (elWidth: ${elWidth}, elHeight: ${elHeight})`);
      }

      let heightLeft = scaledHeight;
      let position = 0;

      // เพิ่มภาพลงใน PDF (หน้าแรก)
      pdf.addImage(imgData, 'JPEG', margin, margin, pdfWidth, Math.max(1, scaledHeight));
      heightLeft -= pdfHeight;

      // หากข้อมูลยาวเกิน 1 หน้า ให้เพิ่มหน้าและตัดครอบต่อไป
      while (heightLeft > 0) {
        position -= pdfHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'JPEG', margin, margin + position, pdfWidth, Math.max(1, scaledHeight));
        heightLeft -= pdfHeight;
      }

      pdf.save(`E-Portfolio_${profile.name}.pdf`);
    } catch (err) {
      console.error('PDF error:', err);
      alert("เกิดข้อผิดพลาดในการสร้างเอกสาร PDF: " + (err.message || ''));
    } finally {
      setIsPrinting(false);
    }
  };

  // ----------------------------------------------------
  // ระบบสร้าง Excel (แฟ้มข้อมูลรายหมวดหมู่) - แอดมิน
  // ----------------------------------------------------
  const handleExportCategoryExcel = (categoryId) => {
    try {
      const category = SCHEMA[categoryId];
      const data = getSortedData(portfolioData[categoryId] || [], category);
      
      if (!data || data.length === 0) {
        alert("ไม่มีข้อมูลสำหรับส่งออก");
        return;
      }

      // Format data for Excel
      const excelData = data.map((item, index) => {
        const rowData = { 'ลำดับที่': index + 1 };
        category.fields.forEach(field => {
          rowData[field.label] = item[field.key] || '-';
        });
        return rowData;
      });

      const worksheet = XLSX.utils.json_to_sheet(excelData);
      
      // Auto-size columns (basic logic)
      const columnWidths = [ { wpx: 50 } ]; // First column (ลำดับที่)
      category.fields.forEach(field => {
         columnWidths.push({ wpx: Math.max(120, field.label.length * 10) }); // basic width mapping
      });
      worksheet['!cols'] = columnWidths;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, category.name.substring(0, 31)); // sheet names max 31 chars
      
      XLSX.writeFile(workbook, `แฟ้มข้อมูล_${category.name}_${profile.name}.xlsx`);
    } catch (err) {
      console.error('Excel error:', err);
      alert("เกิดข้อผิดพลาดในการสร้างเอกสาร Excel: " + (err.message || ''));
    }
  };

  // ----------------------------------------------------
  // ระบบสร้าง PDF (แฟ้มข้อมูลรายหมวดหมู่) - แอดมิน
  // ----------------------------------------------------
  const handlePrintCategoryPDF = async (categoryId, orientation = 'portrait', year = 'all') => {
    setIsPrintingCategory(true);
    const category = SCHEMA[categoryId];
    let rawData = portfolioData[categoryId] || [];
    if (year !== 'all') {
      rawData = rawData.filter(item => getItemYear(item) === year);
    }
    const data = getSortedData(rawData, category);
    
    // ตั้งค่าความกว้างเพื่อให้พอดีกับแนวกระดาษ
    const isLandscape = orientation === 'landscape';
    const containerWidthStyle = isLandscape ? '277mm' : '190mm';
    const containerHeightStyle = isLandscape ? '190mm' : '277mm';

    const tempContainer = document.createElement('div');
    tempContainer.style.position = 'fixed';
    tempContainer.style.left = '200vw'; // ดันออกนอกจอ
    tempContainer.style.top = '0';
    document.body.appendChild(tempContainer);

    // Get expected real pixel height of A4 printable area
    const measurer = document.createElement('div');
    measurer.style.width = containerWidthStyle;
    measurer.style.height = containerHeightStyle;
    tempContainer.appendChild(measurer);
    const maxPageHeightPx = measurer.clientHeight;
    tempContainer.removeChild(measurer);

    let pagesHTML = [];
    const imageUrlsToPreload = [];

    const getTableHeader = () => `
      <div style="text-align: center; margin-bottom: 20px;">
        <h2 style="font-size: 24px; font-weight: bold; color: #1e293b; margin-bottom: 5px;">แฟ้มข้อมูล: ${category.name} ${year !== 'all' ? `(ประจำปี ${year})` : ''}</h2>
        <p style="font-size: 16px; color: #475569;">${profile.name} | ${profile.position} | ${profile.school}</p>
      </div>
      <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 14px; table-layout: fixed;">
        <thead>
          <tr style="background-color: #f1f5f9; border-bottom: 2px solid #cbd5e1;">
            <th style="padding: 10px; border: 1px solid #e2e8f0; font-weight: 600; width: 6%; text-align: center;">#</th>
            ${category.fields.map(f => {
              let widthStyle = '';
              if (f.type === 'url') widthStyle = 'width: 10%; text-align: center;';
              else if (f.type === 'date' || f.key === 'dateStr' || String(f.label).includes('วัน')) widthStyle = 'width: 18%;'; // ขยายช่องวันที่ให้พอดีกับข้อความ
              else if (f.type === 'textarea') widthStyle = 'width: 22%;';
              else if (['title', 'courseName', 'topic', 'workName', 'organization'].includes(f.key)) widthStyle = 'width: 28%;';
              else widthStyle = 'width: 18%;';
              return `<th style="padding: 10px; border: 1px solid #e2e8f0; font-weight: 600; ${widthStyle}">${f.label}</th>`;
            }).join('')}
          </tr>
        </thead>
        <tbody>
    `;

    const getTableFooter = () => `</tbody></table>`;

    const createPageElement = () => {
      const pageEl = document.createElement('div');
      pageEl.className = 'pdf-page-render';
      pageEl.style.fontFamily = "'Kanit', sans-serif";
      pageEl.style.color = "#334155";
      pageEl.style.padding = "20px";
      pageEl.style.width = containerWidthStyle;
      pageEl.style.backgroundColor = "#ffffff";
      pageEl.style.boxSizing = "border-box";
      // Ensure specific global styles for rendering
      pageEl.innerHTML = `
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600&display=swap');
          .pdf-page-render * { font-family: 'Kanit', sans-serif !important; letter-spacing: 0px !important; }
        </style>
        ${getTableHeader()}
      `;
      return pageEl;
    };

    let currentPageEl = createPageElement();
    tempContainer.appendChild(currentPageEl);
    let tbody = currentPageEl.querySelector('tbody');

    if (data.length > 0) {
      for (let i = 0; i < data.length; i++) {
        const item = data[i];
        let trHtml = `<tr style="border-bottom: 1px solid #e2e8f0; page-break-inside: avoid;">`;
        trHtml += `<td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; vertical-align: top;"><div style="line-height: 1.5;">${i + 1}</div></td>`;

        category.fields.forEach(f => {
          if (f.type === 'url') {
            if (item[f.key]) {
              const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=${encodeURIComponent(item[f.key])}`;
              imageUrlsToPreload.push(qrUrl);
              trHtml += `<td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; vertical-align: middle;"><img src="${qrUrl}" alt="QR" style="width: 60px; height: 60px; display: block; margin: 0 auto;" crossorigin="anonymous" /></td>`;
            } else {
              trHtml += `<td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; vertical-align: middle;">-</td>`;
            }
          } else {
            let displayValue = item[f.key] ? String(item[f.key]).replace(/\n/g, '<br/>') : '-';
            if (f.type === 'date' && item[f.key]) {
              displayValue = formatThaiDate(item[f.key]);
            }
            trHtml += `<td style="padding: 10px; border: 1px solid #e2e8f0; vertical-align: top;"><div style="word-break: break-word; overflow-wrap: break-word; white-space: normal; line-height: 1.5;">${displayValue}</div></td>`;
          }
        });
        trHtml += '</tr>';

        // Add to DOM temporarily to measure height
        tbody.insertAdjacentHTML('beforeend', trHtml);

        // Check if exceeded max height (leaving 40px buffer)
        const SAFE_BUFFER = 40;
        if (currentPageEl.offsetHeight > (maxPageHeightPx - SAFE_BUFFER)) {
          // If there's more than one row on this page, bump the last one to the next page
          if (tbody.querySelectorAll('tr').length > 1) {
            tbody.lastElementChild.remove();
            
            // Push current page HTML
            pagesHTML.push(currentPageEl.innerHTML + getTableFooter());
            tempContainer.removeChild(currentPageEl);
            
            // Create new page
            currentPageEl = createPageElement();
            tempContainer.appendChild(currentPageEl);
            tbody = currentPageEl.querySelector('tbody');
            
            // Add the item to the new page
            tbody.insertAdjacentHTML('beforeend', trHtml);
          }
        }
      }
    } else {
      tbody.insertAdjacentHTML('beforeend', `<tr><td colspan="${category.fields.length + 1}" style="text-align: center; padding: 20px; border: 1px solid #e2e8f0;">ไม่มีข้อมูล</td></tr>`);
    }
    
    // Push the final page
    pagesHTML.push(currentPageEl.innerHTML + getTableFooter());
    tempContainer.removeChild(currentPageEl);

    try {
      // Preload ALL images needed
      await Promise.all([...new Set(imageUrlsToPreload)].map(url => {
        return new Promise((resolve) => {
          const img = new Image();
          img.crossOrigin = 'anonymous';
          img.onload = resolve;
          img.onerror = resolve;
          img.src = url;
        });
      }));

      const { toJpeg } = await import('html-to-image');
      const pdf = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
      const margin = 10;
      const pdfWidth = pdf.internal.pageSize.getWidth() - (margin * 2);
      const pdfHeight = pdf.internal.pageSize.getHeight() - (margin * 2);

      for (let p = 0; p < pagesHTML.length; p++) {
        const renderContainer = document.createElement('div');
        renderContainer.className = 'pdf-page-render';
        renderContainer.style.position = 'relative';
        renderContainer.style.width = containerWidthStyle;
        renderContainer.style.height = containerHeightStyle; // Force strict size proportion
        renderContainer.style.backgroundColor = "#ffffff";
        renderContainer.innerHTML = pagesHTML[p];
        tempContainer.appendChild(renderContainer);

        // Wait to render
        await new Promise(r => setTimeout(r, 400)); 

        const imgData = await toJpeg(renderContainer, { 
          quality: 0.98, 
          backgroundColor: '#ffffff',
          pixelRatio: 2,
          width: renderContainer.offsetWidth,
          height: renderContainer.offsetHeight
        });

        if (p > 0) pdf.addPage();
        pdf.addImage(imgData, 'JPEG', margin, margin, pdfWidth, pdfHeight);
        tempContainer.removeChild(renderContainer);
      }

      pdf.save(`${category.name}_${profile.name}.pdf`);
    } catch (err) {
      console.error('PDF error:', err);
      alert("เกิดข้อผิดพลาดในการสร้างเอกสาร PDF: " + (err.message || ''));
    } finally {
      setIsPrintingCategory(false);
      if (document.body.contains(tempContainer)) {
        document.body.removeChild(tempContainer);
      }
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
        <div className="bg-white p-8 rounded-3xl shadow-xl flex flex-col items-center">
          <div className="w-16 h-16 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mb-4"></div>
          <p className="text-slate-600 font-medium animate-pulse">กำลังเตรียมห้องทำงานครู...</p>
        </div>
      </div>
    );
  }

  // --- Simple Login Page ---
  if (showLanding && (!user || user.isAnonymous)) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
        <CustomThemeStyle theme={settings.themeColor || 'blue'} />
        <div className="w-full max-w-md bg-white rounded-2xl shadow-xl overflow-hidden animate-in fade-in zoom-in duration-300 border border-slate-200">
          <div className="bg-slate-800 p-8 text-white text-center relative overflow-hidden">
            <div className="absolute top-0 left-0 w-full h-full bg-blue-600/10 pointer-events-none"></div>
            <div className="w-16 h-16 bg-white/10 rounded-2xl flex items-center justify-center mx-auto mb-4 backdrop-blur-sm border border-white/20">
              <UserCheck size={32} className="text-white" />
            </div>
            <h1 className="text-2xl font-bold tracking-tight">E-Portfolio</h1>
            <p className="text-slate-400 mt-2 text-sm font-medium">จัดการผลงานระดับมืออาชีพ</p>
          </div>
          
          <div className="p-8">
            <div className="space-y-6">
              <div className="text-center space-y-2">
                <h3 className="font-bold text-slate-800 text-lg">เริ่มต้นสร้างแฟ้มผลงาน</h3>
                <p className="text-sm text-slate-500 leading-relaxed">
                  ใช้บัญชี Google ของคุณเพื่อจัดการข้อมูลส่วนตัว 
                  ข้อมูลจะถูกเก็บเป็นความลับและเข้าถึงได้เฉพาะคุณเท่านั้น
                </p>
              </div>

              <button 
                type="button"
                onClick={handleGoogleLogin}
                className="w-full flex items-center justify-center gap-3 bg-white text-slate-700 hover:text-slate-900 transition-all text-base font-bold border-2 border-slate-100 p-4 rounded-2xl hover:border-blue-500 hover:bg-blue-50 shadow-sm active:scale-95"
              >
                <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-6 h-6" />
                เข้าใช้งานด้วย Google Account
              </button>

              <div className="relative py-4">
                <div className="absolute inset-0 flex items-center">
                  <span className="w-full border-t border-slate-100"></span>
                </div>
                <div className="relative flex justify-center text-[10px] uppercase tracking-widest font-bold">
                  <span className="bg-white px-3 text-slate-400">สำหรับผู้ดูแลระบบ</span>
                </div>
              </div>

              <form onSubmit={handleLogin} className="space-y-3">
                <div className="relative">
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                    <Lock size={14} />
                  </div>
                  <input 
                    type="password" 
                    value={passwordInput} 
                    onChange={(e) => setPasswordInput(e.target.value)}
                    placeholder="รหัสผ่าน Admin" 
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-10 pr-4 py-2.5 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-slate-800 focus:bg-white transition-all"
                  />
                </div>

                {loginError && (
                  <div className="bg-red-50 text-red-600 text-[10px] p-2 rounded-lg border border-red-100 flex items-center gap-2 animate-shake">
                    <AlertCircle size={12} className="shrink-0" />
                    <span className="font-medium">{loginError}</span>
                  </div>
                )}

                <button 
                  type="submit"
                  className="w-full bg-slate-700 text-white py-2.5 rounded-xl text-xs font-bold hover:bg-slate-800 transition-all active:scale-95 flex items-center justify-center gap-2"
                >
                  เข้าสู่ระบบ Admin
                </button>
              </form>
            </div>


            {authErrorCode === 'auth/admin-restricted-operation' && (
              <div className="mt-4 bg-amber-50 text-amber-700 text-[10px] p-3 rounded-xl border border-amber-200">
                <p className="font-bold flex items-center gap-1 mb-1">
                  <Lock size={12} /> ยังไม่ได้เปิดใช้งาน Guest Access
                </p>
                <p>กรุณาไปที่ Firebase Console &gt; Auth &gt; Sign-in method และเปิดใช้งาน <b>Anonymous</b> ครับ</p>
              </div>
            )}

            {authErrorCode === 'auth/unauthorized-domain' && (
              <div className="mt-4 bg-orange-50 text-orange-700 text-[10px] p-3 rounded-xl border border-orange-200">
                <p className="font-bold flex items-center gap-1 mb-1">
                  <AlertCircle size={12} /> โดเมนยังไม่ได้รับอนุญาต
                </p>
                <p>กรุณาเพิ่ม <b>{window.location.hostname}</b> ในรายการ Authorized Domains ของ Firebase ครับ</p>
              </div>
            )}
          </div>
          
          <div className="bg-slate-50 p-5 border-t border-slate-100 text-center">
             <p className="text-[10px] text-slate-500 font-medium">ระบบแฟ้มสะสมผลงานครู - พัฒนาเพื่อการศึกษาไทย</p>
          </div>
        </div>
      </div>
    );
  }

  const isPublicView = viewMode === 'public';

  const NavItem = ({ id, icon: Icon, label }) => (
    <button
      onClick={() => { setActiveTab(id); setIsSidebarOpen(false); }}
      className={`w-full flex items-center px-3 py-2 rounded-md mb-0.5 transition-colors text-sm ${
        activeTab === id ? 'bg-blue-800 text-white' : 'text-blue-100 hover:bg-blue-800 hover:text-white'
      }`}
    >
      <Icon size={18} className="mr-2.5" />
      <span>{label}</span>
      {['profile', 'settings', 'theme'].includes(id) && !isAdminAuthenticated && <Lock size={14} className="ml-auto opacity-60" />}
    </button>
  );

  if (isPublicView) {
    return (
      <>
        <CustomThemeStyle theme={settings.themeColor || 'blue'} />
        <div className="bg-slate-50 min-h-screen text-slate-800 print:bg-white print:m-0 print:p-0">
          <nav className="bg-white/80 backdrop-blur-md shadow-sm sticky top-0 z-50 print:hidden border-b border-slate-200">
            <div className="max-w-5xl mx-auto px-4 md:px-8 h-16 flex items-center justify-between">
              <div className="flex items-center">
                <div className="w-8 h-8 bg-blue-600 text-white rounded-lg flex items-center justify-center font-bold mr-3 shadow-sm">EP</div>
                <span className="font-bold text-slate-800 text-lg hidden sm:block tracking-wide">E-Portfolio Preview</span>
              </div>
              <div className="flex items-center gap-2 md:gap-3">
                <button onClick={() => setPrintModalState({ target: 'overview' })} disabled={isPrinting} className={`text-sm bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 md:px-4 py-2 rounded-lg transition flex items-center font-medium ${isPrinting ? 'opacity-50 cursor-wait' : ''}`}>
                  <Printer size={16} className="md:mr-2" /> <span className="hidden md:inline">{isPrinting ? 'กำลังสร้าง...' : 'ดาวน์โหลด PDF'}</span>
                </button>
                <button onClick={() => setViewMode('admin')} className="text-sm bg-slate-800 hover:bg-slate-900 text-white px-3 md:px-4 py-2 rounded-lg transition flex items-center shadow-sm font-medium">
                  <Edit2 size={16} className="md:mr-2" /> <span className="hidden md:inline">กลับหน้าระบบหลัก</span>
                </button>
              </div>
            </div>
          </nav>

          <div id="portfolio-content" className="max-w-4xl mx-auto p-4 md:p-8 print:p-0 mt-4 md:mt-8 bg-slate-50">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 md:p-8 mb-8 print:shadow-none print:border-none print:p-0 print:mb-8">
              <div className="flex flex-col md:flex-row items-center md:items-start gap-8">
                <div 
                  className="rounded-full border-4 border-blue-50 shadow-md overflow-hidden shrink-0 transition-all duration-300 relative" 
                  style={{ 
                    width: profile.imageSize ? `${profile.imageSize}px` : '160px', 
                    height: profile.imageSize ? `${profile.imageSize}px` : '160px'
                  }}
                >
                  <img 
                    src={getDirectImageUrl(profile.photoUrl) || DEFAULT_AVATAR} 
                    alt="Profile" 
                    className="absolute inset-0 w-full h-full object-cover transition-transform duration-300" 
                    style={{ 
                      objectPosition: `${profile.imagePanX ?? 50}% ${profile.imagePanY ?? 50}%`,
                      transform: `scale(${profile.imageZoom ?? 1})`
                    }}
                    onError={(e) => { e.target.onerror = null; e.target.src = DEFAULT_AVATAR; }} 
                  />
                </div>
                <div className="text-center md:text-left flex-1 w-full">
                  <h1 className="text-2xl md:text-3xl font-bold text-slate-900 mb-2">{profile.name}</h1>
                  <p className="text-lg md:text-xl text-blue-600 font-medium mb-2">{profile.position}</p>
                  <p className="text-sm md:text-base text-slate-500 mb-4 flex items-center justify-center md:justify-start">
                    <Home size={16} className="mr-2"/> {profile.school}
                  </p>
                  <div className="bg-slate-50 p-4 md:p-5 rounded-xl border border-slate-100 text-slate-700 leading-relaxed text-sm md:text-base text-left">
                    {profile.bio}
                  </div>
                  <div className="mt-5 flex flex-wrap gap-3 justify-center md:justify-start text-sm text-slate-600">
                    {profile.email && <span className="flex items-center bg-white px-3 py-1.5 rounded-lg border border-slate-200 shadow-sm"><span className="font-medium mr-2">Email:</span> {profile.email}</span>}
                    {profile.phone && <span className="flex items-center bg-white px-3 py-1.5 rounded-lg border border-slate-200 shadow-sm"><span className="font-medium mr-2">โทร:</span> {profile.phone}</span>}
                  </div>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-8 print:gap-6">
            {Object.values(SCHEMA).map(category => {
              const settingKey = `show${category.id.charAt(0).toUpperCase() + category.id.slice(1)}`;
              const isVisible = settings[settingKey] !== false;
              const allData = portfolioData[category.id] || [];
              const availableYears = getUniqueYears(allData);
              const data = getSortedData(allData, category, yearFilter === 'all' ? 'all' : yearFilter);
              
              if (!isVisible || allData.length === 0) return null;

              const Icon = category.icon;

              return (
                <div key={category.id} className="page-break-inside-avoid shadow-sm rounded-2xl overflow-hidden bg-white border border-slate-100 p-6 md:p-8">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6 border-b border-slate-100 pb-4">
                    <h2 className="text-xl md:text-2xl font-bold text-slate-800 flex items-center">
                      <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-lg flex items-center justify-center mr-3 border border-blue-100">
                        <Icon size={20} />
                      </div>
                      {category.name}
                    </h2>
                    
                    {/* Year Filter for Public View */}
                    {availableYears.length > 0 && (
                      <div className="flex flex-wrap gap-2 print:hidden">
                        <button 
                          onClick={() => setYearFilter('all')}
                          className={`px-3 py-1 rounded-full text-xs font-semibold transition ${yearFilter === 'all' ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                        >
                          ทั้งหมด
                        </button>
                        {availableYears.map(year => (
                          <button 
                            key={year}
                            onClick={() => setYearFilter(year)}
                            className={`px-3 py-1 rounded-full text-xs font-semibold transition ${yearFilter === year ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                          >
                            ปี {year}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  {data.length > 0 ? (
                    <div className="grid gap-4">
                      {data.map((item) => (
                        <div key={item.id} className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm relative hover:shadow-md transition duration-200">
                           <div className="absolute top-0 left-0 w-1.5 h-full bg-blue-500 rounded-l-xl"></div>
                           <div className="ml-3">
                            {category.fields.map((field, fIdx) => {
                              if (!item[field.key]) return null;
                              if (field.type === 'url') return null; 
                              
                              const isMain = fIdx === 0;
                              const displayValue = field.type === 'date' ? formatThaiDate(item[field.key]) : item[field.key];

                              return (
                                <div key={field.key} className={isMain ? "mb-2" : "mb-1.5"}>
                                  {isMain ? (
                                    <h3 className="text-lg font-bold text-slate-800">{displayValue}</h3>
                                  ) : (
                                    <p className="text-sm text-slate-600 flex flex-col sm:flex-row sm:items-start">
                                      <span className="font-medium text-slate-500 sm:w-36 inline-block shrink-0 mb-0.5 sm:mb-0">{field.label}:</span> 
                                      <span className="text-slate-700">{displayValue}</span>
                                    </p>
                                  )}
                                </div>
                              );
                            })}
                            {category.fields.filter(f => f.type === 'url').map(field => {
                              if (item[field.key]) {
                                return (
                                  <div key={field.key} className="mt-4 pt-3 border-t border-slate-100">
                                    <a href={item[field.key]} target="_blank" rel="noopener noreferrer" className="inline-flex items-center px-3 py-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg text-sm font-medium transition">
                                      <LinkIcon size={14} className="mr-1.5" /> เปิดดูเอกสาร/ไฟล์แนบ
                                    </a>
                                  </div>
                                )
                              }
                              return null;
                            })}
                           </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="py-12 text-center text-slate-400 italic bg-slate-50/50 rounded-xl border border-dashed border-slate-200">
                      ไม่พบข้อมูลในปี {yearFilter}
                    </div>
                  )}
                </div>
              );
            })}
            </div>
          </div>
        </div>

        {/* UI Selection Modal for Print Orientation (Shared for Public View) */}
        {printModalState && (
          <div className="fixed inset-0 bg-slate-900/50 z-[100] flex items-center justify-center p-4">
            <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden">
              <div className="bg-blue-600 p-4 text-white">
                <h3 className="text-lg font-bold flex items-center"><Printer className="mr-2" size={20} /> รูปแบบการดาวน์โหลด PDF</h3>
              </div>
              
              <div className="p-6">
                {printModalState.target === 'category' && (
                  <div className="mb-6">
                    <label className="block text-sm font-bold text-slate-700 mb-2 flex items-center">
                      <Filter size={14} className="mr-2" /> เลือกช่วงเวลา (ปี พ.ศ.)
                    </label>
                    <div className="grid grid-cols-3 gap-2">
                      <button 
                        onClick={() => setYearFilter('all')}
                        className={`py-2 px-3 rounded-lg text-sm font-medium border transition ${yearFilter === 'all' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
                      >
                        ทั้งหมด
                      </button>
                      {getUniqueYears(portfolioData[printModalState.categoryId]).map(y => (
                        <button 
                          key={y}
                          onClick={() => setYearFilter(y)}
                          className={`py-2 px-3 rounded-lg text-sm font-medium border transition ${yearFilter === y ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
                        >
                          ปี {y}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                <label className="block text-sm font-bold text-slate-700 mb-3 flex items-center">
                  <Monitor size={14} className="mr-2" /> การวางแนวหน้ากระดาษ
                </label>
                <div className="grid grid-cols-2 gap-4 mb-6">
                  <button 
                    onClick={() => {
                       const { target, categoryId } = printModalState;
                       setPrintModalState(null);
                       if (target === 'overview') handlePrintPDF('portrait');
                       else handlePrintCategoryPDF(categoryId, 'portrait', yearFilter);
                    }}
                    className="flex flex-col items-center justify-center p-4 border-2 border-slate-200 rounded-xl hover:bg-blue-50 hover:border-blue-500 transition group"
                  >
                    <div className="border-2 border-slate-300 w-10 h-14 rounded-sm mb-3 transition group-hover:border-blue-500 flex items-center justify-center">
                      <div className="w-6 h-0.5 bg-slate-200 group-hover:bg-blue-200 mb-1"></div>
                      <div className="w-6 h-0.5 bg-slate-200 group-hover:bg-blue-200"></div>
                    </div>
                    <span className="text-sm font-bold text-slate-700 group-hover:text-blue-700">แนวตั้ง (Portrait)</span>
                  </button>
                  <button 
                    onClick={() => {
                       const { target, categoryId } = printModalState;
                       setPrintModalState(null);
                       if (target === 'overview') handlePrintPDF('landscape');
                       else handlePrintCategoryPDF(categoryId, 'landscape', yearFilter);
                    }}
                    className="flex flex-col items-center justify-center p-4 border-2 border-slate-200 rounded-xl hover:bg-blue-50 hover:border-blue-500 transition group"
                  >
                    <div className="border-2 border-slate-300 w-14 h-10 rounded-sm mb-3 transition group-hover:border-blue-500 flex flex-col items-center justify-center">
                      <div className="w-10 h-0.5 bg-slate-200 group-hover:bg-blue-200 mb-1"></div>
                      <div className="w-10 h-0.5 bg-slate-200 group-hover:bg-blue-200"></div>
                    </div>
                    <span className="text-sm font-bold text-slate-700 group-hover:text-blue-700">แนวนอน (Landscape)</span>
                  </button>
                </div>

                <button 
                  onClick={() => setPrintModalState(null)}
                  className="w-full py-2.5 text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-lg transition font-bold text-sm"
                >
                  ยกเลิก
                </button>
              </div>
            </div>
          </div>
        )}
      </>
    );
  }

  return (
    <>
      <CustomThemeStyle theme={settings.themeColor || 'blue'} />
      <div className="flex h-screen bg-slate-50 text-slate-800">
        
        {isSidebarOpen && (
          <div className="fixed inset-0 bg-black bg-opacity-50 z-40 md:hidden" onClick={() => setIsSidebarOpen(false)}></div>
        )}

        <aside className={`fixed md:static inset-y-0 left-0 w-60 bg-blue-900 text-white z-50 transform transition-transform duration-300 ease-in-out ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full md:translate-x-0'} flex flex-col print:hidden`}>
          <div className="p-4 border-b border-blue-800 flex justify-between items-center">
            <div className="flex items-center gap-3 overflow-hidden">
              <div className="w-11 h-11 rounded-full border-2 border-blue-400/50 shadow-sm shrink-0 overflow-hidden relative">
                <img src={getDirectImageUrl(profile.photoUrl) || DEFAULT_AVATAR} alt="Profile" className="absolute inset-0 w-full h-full object-cover" style={{ objectPosition: `${profile.imagePanX ?? 50}% ${profile.imagePanY ?? 50}%`, transform: `scale(${profile.imageZoom ?? 1})` }} onError={(e) => { e.target.onerror = null; e.target.src = DEFAULT_AVATAR; }} />
              </div>
              <div className="flex flex-col overflow-hidden">
                <h1 className="text-sm font-bold tracking-wider text-white truncate">E-Portfolio</h1>
                <p className="text-xs text-blue-200 truncate mt-0.5">{profile.name}</p>
              </div>
            </div>
            <button className="md:hidden text-blue-200 shrink-0 ml-2" onClick={() => setIsSidebarOpen(false)}><X size={20} /></button>
          </div>
          <div className="flex-1 overflow-y-auto p-3 custom-scrollbar">
            <div className="text-[11px] font-semibold text-blue-300 uppercase tracking-wider mb-1.5 mt-2 pl-3">ระบบจัดการ</div>
            <NavItem id="dashboard" icon={Home} label="แดชบอร์ดภาพรวม" />
            {user?.isMasterAdmin && (
              <button
                onClick={() => setActiveTab('system-stats')}
                className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-xl transition-all duration-200 group mb-1 ${
                  activeTab === 'system-stats' 
                    ? 'bg-blue-600 text-white shadow-md shadow-blue-200' 
                    : 'text-slate-400 hover:bg-blue-50 hover:text-blue-600'
                }`}
              >
                <BarChart3 size={18} className={activeTab === 'system-stats' ? 'text-white' : 'group-hover:text-blue-600'} />
                <span className="text-sm font-medium">สถิติระบบ</span>
              </button>
            )}
            
            <div className="text-[11px] font-semibold text-blue-300 uppercase tracking-wider mb-1.5 mt-4 pl-3">แฟ้มข้อมูล</div>
            {Object.values(SCHEMA).map(cat => {
              const settingKey = `show${cat.id.charAt(0).toUpperCase() + cat.id.slice(1)}`;
              const isVisible = settings[settingKey] !== false;
              if (!isVisible) return null;
              return <NavItem key={cat.id} id={cat.id} icon={cat.icon} label={cat.name} />
            })}

            {isAdminAuthenticated ? (
              <>
                <div className="text-[11px] font-semibold text-blue-300 uppercase tracking-wider mb-1.5 mt-4 pl-3">ระบบจัดการส่วนตัว</div>
                <NavItem id="profile" icon={User} label="โปรไฟล์" />
                <NavItem id="settings" icon={Settings} label="ตั้งค่าระบบ" />
                <NavItem id="theme" icon={Palette} label="สีสันของระบบ" />
              </>
            ) : (
              <>
                <div className="text-[11px] font-semibold text-blue-300 uppercase tracking-wider mb-1.5 mt-4 pl-3">สำหรับผู้ดูแลระบบ</div>
                <NavItem id="settings" icon={Settings} label="เข้าสู่ระบบ" />
              </>
            )}
          </div>
          {user && (
            <div className="p-3 border-t border-blue-800">
              <button onClick={handleLogout} className="w-full flex items-center justify-center px-3 py-2 bg-blue-800 hover:bg-red-600 text-blue-100 hover:text-white rounded-md text-sm transition font-medium">
                <LogOut size={16} className="mr-2" /> ออกจากระบบ
              </button>
            </div>
          )}
        </aside>

        <main className="flex-1 flex flex-col h-screen print:h-auto overflow-hidden print:overflow-visible">
          <header className="bg-white shadow-sm h-16 flex items-center justify-between px-4 md:px-8 border-b border-slate-200 print:hidden">
            <div className="flex items-center">
              <button className="md:hidden mr-4 text-slate-600" onClick={() => setIsSidebarOpen(true)}>
                <Menu size={24} />
              </button>
              <h2 className="text-xl font-semibold text-slate-800">
                {activeTab === 'dashboard' ? 'แดชบอร์ดภาพรวม' : 
                 activeTab === 'profile' ? 'จัดการโปรไฟล์' :
                 activeTab === 'settings' ? (!isAdminAuthenticated ? 'เข้าสู่ระบบผู้ดูแล' : 'ตั้งค่าระบบ') : 
                 activeTab === 'theme' ? 'สีสันของระบบ' :
                 SCHEMA[activeTab]?.name}
              </h2>
            </div>
            <div className="flex items-center gap-2 md:gap-3">
              {saveStatus && <span className="hidden sm:inline-block text-sm text-green-600 bg-green-50 px-3 py-1 rounded-full animate-pulse">{saveStatus}</span>}
              
              <button onClick={() => setViewMode('public')} className="flex items-center text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg transition shadow-sm font-medium">
                <Eye size={16} className="mr-1.5" /> <span className="hidden sm:inline">พรีวิว / พิมพ์ PDF</span><span className="sm:hidden">พรีวิว</span>
              </button>
              
              <div className="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center text-blue-700 ml-1 md:ml-2 shadow-sm overflow-hidden shrink-0 relative">
                <img src={getDirectImageUrl(profile.photoUrl) || DEFAULT_AVATAR} alt="User" className="absolute inset-0 w-full h-full object-cover" style={{ objectPosition: `${profile.imagePanX ?? 50}% ${profile.imagePanY ?? 50}%`, transform: `scale(${profile.imageZoom ?? 1})` }} onError={(e) => { e.target.onerror = null; e.target.src = DEFAULT_AVATAR; }} />
              </div>
            </div>
          </header>

          {user?.isGuest && isAdminAuthenticated && (
            <div className="bg-red-50 border-b border-red-200 px-4 py-3 text-red-900 print:hidden">
              <div className="max-w-4xl mx-auto flex items-start gap-4">
                <AlertCircle size={20} className="shrink-0 mt-0.5 text-red-500" />
                <div className="text-sm">
                  <p className="font-bold mb-1">แจ้งเตือน: ไม่สามารถบันทึกข้อมูลได้ (Permissions Error)</p>
                  <p className="mb-2">ระบบไม่สามารถยืนยันตัวตนกับ Firebase ได้สำเร็จ ทำให้คุณอยู่ในโหมด "อ่านอย่างเดียว"</p>
                  
                  {authErrorCode === 'auth/unauthorized-domain' ? (
                    <div className="bg-white/60 p-3 rounded-lg border border-red-100 mt-2 shadow-sm">
                      <p className="font-bold text-red-700 text-xs mb-1 uppercase tracking-wider">สาเหตุ: โดเมนยังไม่ได้รับอนุญาต (Unauthorized Domain)</p>
                      <p className="text-xs text-red-600 mb-2 leading-relaxed">กรุณานำชื่อโดเมนด้านล่างนี้ไปเพิ่มที่ <b>Firebase Console &gt; Authentication &gt; Settings &gt; Authorized Domains</b></p>
                      <div className="flex gap-2 items-center">
                        <code className="flex-1 bg-red-100/50 p-2 rounded text-[11px] font-mono select-all border border-red-200 overflow-x-auto whitespace-nowrap">
                          {window.location.hostname}
                        </code>
                        <button onClick={() => navigator.clipboard.writeText(window.location.hostname)} className="text-[10px] bg-red-200 hover:bg-red-300 px-2 py-2 rounded shrink-0 transition-colors">คัดลอก</button>
                      </div>
                      <p className="text-[10px] text-gray-500 mt-2 italic font-medium">* สำคัญมาก: หลังจากเพิ่มแล้ว คุณต้อง "รีเฟรชหน้าเว็บ" นี้อีกครั้ง</p>
                    </div>
                  ) : (
                    <p className="text-xs opacity-75 italic mt-1 font-medium bg-white/40 p-2 rounded">
                      Error ที่พบ: {authErrorCode || 'Anonymous Auth อาจยังไม่ได้เปิดใช้งาน หรือมีปัญหาการเชื่อมต่อ'}
                    </p>
                  )}
                </div>
              </div>
            </div>
          )}

          <div className="flex-1 overflow-y-auto print:overflow-visible p-4 md:p-8 bg-slate-50 print:bg-white print:p-0">
            
            {activeTab === 'dashboard' && (
              <div className="max-w-4xl mx-auto space-y-6">
                {/* Share Link Card - Compact Version */}
                {user && !user.isGuest && (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                    <div className="bg-gradient-to-r from-blue-700 to-indigo-800 rounded-2xl shadow-lg overflow-hidden text-white border border-white/10 relative">
                      <div className="absolute top-0 right-0 p-2 opacity-5 pointer-events-none">
                        <LinkIcon size={80} />
                      </div>
                      
                      <div className="p-5 md:p-6 relative z-10">
                        <div className="flex items-center gap-4 mb-4">
                          <div className="w-10 h-10 bg-white/20 backdrop-blur-md rounded-xl flex items-center justify-center shrink-0 border border-white/30">
                            <LinkIcon size={20} className="text-white" />
                          </div>
                          <div>
                            <h4 className="text-lg font-bold tracking-tight">แชร์แฟ้มผลงาน</h4>
                            <p className="text-blue-100/70 text-[11px] font-medium">คัดลอกลิงก์เพื่อส่งให้คณะกรรมการประเมิน</p>
                          </div>
                        </div>

                        <div className="flex items-stretch gap-2 bg-black/10 p-1.5 rounded-xl border border-white/5">
                          <div className="flex-1 px-3 py-2 text-[11px] font-mono truncate select-all text-blue-50/80 flex items-center">
                            {`${window.location.origin}${window.location.pathname}?u=${user.uid === 'admin-master' ? 'admin-master' : user.uid}`}
                          </div>
                          <button 
                            onClick={() => {
                              const shareUrl = `${window.location.origin}${window.location.pathname}?u=${user.uid === 'admin-master' ? 'admin-master' : user.uid}`;
                              navigator.clipboard.writeText(shareUrl);
                              setSaveStatus('คัดลอกลิงก์สำเร็จ!');
                              setTimeout(() => setSaveStatus(''), 2000);
                            }}
                            className="bg-white text-blue-700 hover:bg-blue-50 px-4 py-2 rounded-lg font-bold text-[11px] transition-all shadow-sm active:scale-95 whitespace-nowrap flex items-center gap-1.5"
                          >
                            <Save size={14} />
                            คัดลอก
                          </button>
                        </div>
                        
                        {saveStatus === 'คัดลอกลิงก์สำเร็จ!' && (
                          <div className="absolute top-4 right-4 bg-green-500 text-white text-[9px] font-bold px-2 py-0.5 rounded-full animate-fade-in shadow-sm">
                            คัดลอกแล้ว!
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Google Drive Shortcut Card */}
                    <div className="bg-white rounded-2xl shadow-lg overflow-hidden border border-slate-200 relative group transition-all hover:shadow-xl">
                      <div className="absolute top-0 right-0 p-2 opacity-5 pointer-events-none text-blue-600">
                        <Monitor size={80} />
                      </div>
                      
                      <div className="p-5 md:p-6 relative z-10 flex flex-col h-full justify-between">
                        <div className="flex items-center gap-4 mb-4">
                          <div className="w-10 h-10 bg-blue-50 rounded-xl flex items-center justify-center shrink-0 border border-blue-100 text-blue-600">
                            <Monitor size={20} />
                          </div>
                          <div>
                            <h4 className="text-lg font-bold tracking-tight text-slate-800">คลังเอกสาร Drive</h4>
                            <p className="text-slate-500 text-[11px] font-medium">เข้าถึงไฟล์หลักฐานทั้งหมดของคุณอย่างรวดเร็ว</p>
                          </div>
                        </div>

                        {profile.googleDriveLink ? (
                          <a 
                            href={profile.googleDriveLink} 
                            target="_blank" 
                            rel="noopener noreferrer"
                            className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-xl font-bold text-sm transition-all shadow-md active:scale-95 flex items-center justify-center gap-2"
                          >
                            <Monitor size={18} />
                            เปิดคลังเอกสาร (Drive)
                          </a>
                        ) : (
                          <button 
                            onClick={() => setActiveTab('profile')}
                            className="w-full bg-slate-100 hover:bg-slate-200 text-slate-600 py-3 rounded-xl font-bold text-sm transition-all border border-slate-200 border-dashed flex items-center justify-center gap-2"
                          >
                            <Plus size={18} />
                            ตั้งค่าลิงก์ Google Drive
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                )}

                <h3 className="text-lg font-bold text-slate-800 mb-4">ภาพรวมข้อมูลผลงาน</h3>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
                  {Object.values(SCHEMA).map(cat => {
                    const settingKey = `show${cat.id.charAt(0).toUpperCase() + cat.id.slice(1)}`;
                    const isVisible = settings[settingKey] !== false;
                    if (!isVisible) return null;
                    
                    const CategoryIcon = cat.icon;
                    return (
                      <div key={cat.id} onClick={() => setActiveTab(cat.id)} className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex items-center cursor-pointer hover:shadow-md hover:border-blue-300 transition-all group">
                        <div className="w-14 h-14 bg-blue-50 group-hover:bg-blue-100 rounded-xl flex items-center justify-center text-blue-600 mr-4 transition-colors shrink-0">
                          <CategoryIcon size={28} />
                        </div>
                        <div className="flex-1 overflow-hidden">
                          <p className="text-slate-500 font-medium text-sm truncate">{cat.name}</p>
                          <p className="text-3xl font-bold text-slate-800 tracking-tight mt-0.5">{portfolioData[cat.id]?.length || 0} <span className="text-sm font-normal text-slate-400">รายการ</span></p>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {['profile', 'settings', 'theme'].includes(activeTab) && (
              !isAdminAuthenticated ? (
                <div className="flex items-center justify-center p-4 pt-10">
                  <div className="bg-white p-8 rounded-2xl shadow-sm border border-slate-200 w-full max-w-sm">
                    <div className="text-center mb-6">
                      <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mx-auto mb-4">
                        <Settings size={32} />
                      </div>
                      <h2 className="text-2xl font-bold text-slate-800">ผู้ดูแลระบบ</h2>
                      <p className="text-sm text-slate-500 mt-1">กรุณาใส่รหัสผ่านเพื่อตั้งค่าและจัดการข้อมูล</p>
                    </div>
                    <form onSubmit={handleLogin}>
                      <div className="mb-4">
                        <input
                          type="password"
                          placeholder="รหัสผ่าน"
                          value={passwordInput}
                          onChange={(e) => setPasswordInput(e.target.value)}
                          className="w-full p-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition"
                        />
                        {loginError && <p className="text-red-500 text-sm mt-2">{loginError}</p>}
                      </div>
                      <button type="submit" className="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 rounded-lg shadow transition">
                        เข้าสู่ระบบ
                      </button>
                    </form>
                  </div>
                </div>
              ) : (
                <div className="max-w-4xl mx-auto space-y-6">
                  
                  {activeTab === 'system-stats' && (
                    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-5">
                          <div className="w-14 h-14 bg-blue-100 rounded-2xl flex items-center justify-center text-blue-600">
                            <Users size={28} />
                          </div>
                          <div>
                            <p className="text-sm font-medium text-slate-500">ผู้ใช้งานทั้งหมด</p>
                            <h4 className="text-2xl font-bold text-slate-900">{isStatsLoading ? '...' : systemStats.totalUsers} ท่าน</h4>
                          </div>
                        </div>
                        
                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-5">
                          <div className="w-14 h-14 bg-green-100 rounded-2xl flex items-center justify-center text-green-600">
                            <CheckCircle2 size={28} />
                          </div>
                          <div>
                            <p className="text-sm font-medium text-slate-500">สถานะระบบ</p>
                            <h4 className="text-2xl font-bold text-slate-900">ปกติ</h4>
                          </div>
                        </div>

                        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-5">
                          <div className="w-14 h-14 bg-amber-100 rounded-2xl flex items-center justify-center text-amber-600">
                            <Database size={28} />
                          </div>
                          <div>
                            <p className="text-sm font-medium text-slate-500">ฐานข้อมูล</p>
                            <h4 className="text-2xl font-bold text-slate-900">Cloud Firestore</h4>
                          </div>
                        </div>
                      </div>

                      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                        <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                          <h3 className="font-bold text-slate-800 flex items-center gap-2">
                            <History size={18} className="text-blue-600" /> รายชื่อผู้สมัครใช้งานล่าสุด
                          </h3>
                          <span className="text-[10px] font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full uppercase tracking-wider">Top 10 Latest</span>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-50 text-slate-500 text-[11px] uppercase tracking-wider font-bold">
                                <th className="px-6 py-4 border-b">ลำดับ</th>
                                <th className="px-6 py-4 border-b">ชื่อ-นามสกุล</th>
                                <th className="px-6 py-4 border-b">อีเมล</th>
                                <th className="px-6 py-4 border-b">UID</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {isStatsLoading ? (
                                <tr>
                                  <td colSpan="4" className="px-6 py-10 text-center text-slate-400 italic">กำลังโหลดข้อมูล...</td>
                                </tr>
                              ) : systemStats.recentUsers.length === 0 ? (
                                <tr>
                                  <td colSpan="4" className="px-6 py-10 text-center text-slate-400 italic">ยังไม่มีข้อมูลผู้ใช้งาน</td>
                                </tr>
                              ) : (
                                systemStats.recentUsers.map((u, idx) => (
                                  <tr key={u.id} className="hover:bg-blue-50/30 transition-colors">
                                    <td className="px-6 py-4 text-sm text-slate-500 font-mono">{idx + 1}</td>
                                    <td className="px-6 py-4">
                                      <div className="font-bold text-slate-800 text-sm">{u.profile?.name || 'ไม่ระบุชื่อ'}</div>
                                      <div className="text-[10px] text-slate-500">{u.profile?.position || '-'}</div>
                                    </td>
                                    <td className="px-6 py-4 text-sm text-slate-600 font-medium">{u.email || '-'}</td>
                                    <td className="px-6 py-4 text-[10px] font-mono text-slate-400">{u.id}</td>
                                  </tr>
                                ))
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>

                      <div className="bg-amber-50 border border-amber-200 p-4 rounded-xl flex gap-4 items-start">
                        <ShieldAlert size={20} className="text-amber-600 shrink-0 mt-0.5" />
                        <div className="text-xs text-amber-800 leading-relaxed">
                          <p className="font-bold mb-1">นโยบายความเป็นส่วนตัวและการจัดการข้อมูล:</p>
                          <p>ข้อมูลที่แสดงด้านบนเป็นเพียงข้อมูลพื้นฐานเพื่อการตรวจสอบจำนวนผู้ใช้งานในระบบเท่านั้น ผู้ดูแลระบบควรระมัดระวังการเผยแพร่ข้อมูลเหล่านี้สู่สาธารณะตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล (PDPA)</p>
                        </div>
                      </div>
                    </div>
                  )}

                  {activeTab === 'profile' && (
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
                      <h3 className="text-lg font-bold text-slate-800 mb-4 border-b pb-2 flex items-center">
                        <User className="mr-2" size={20} /> ข้อมูลส่วนตัว
                      </h3>
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 items-start">
                        <div className="col-span-1 flex flex-col items-center">
                          <div 
                            className="rounded-full overflow-hidden shrink-0 transition-all duration-300 relative group"
                            style={{ 
                              width: profile.imageSize ? `${profile.imageSize}px` : '160px', 
                              height: profile.imageSize ? `${profile.imageSize}px` : '160px',
                              border: `${profile.imageBorderWidth ?? 4}px solid ${profile.imageBorderColor ?? '#f1f5f9'}`,
                              boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1)'
                            }}
                          >
                            {isUploadingImage && (
                              <div className="absolute inset-0 bg-slate-900/50 flex flex-col items-center justify-center text-white z-10 transition-all">
                                <Loader2 className="animate-spin mb-1" size={24} />
                                <span className="text-[10px] font-medium">กำลังโหลด...</span>
                              </div>
                            )}
                            <img 
                              src={getDirectImageUrl(profile.photoUrl) || DEFAULT_AVATAR} 
                              alt="Profile" 
                              className={`absolute inset-0 w-full h-full object-cover transition-transform duration-300 ${isUploadingImage ? 'opacity-50' : 'opacity-100'}`}
                              style={{ 
                                objectPosition: `${profile.imagePanX ?? 50}% ${profile.imagePanY ?? 50}%`,
                                transform: `scale(${profile.imageZoom ?? 1})`
                              }}
                              onError={(e) => { e.target.onerror = null; e.target.src = DEFAULT_AVATAR; }} 
                            />
                          </div>

                          <div className="w-full mb-4">
                            <input 
                              type="file" 
                              accept="image/*" 
                              id="profile-upload" 
                              className="hidden" 
                              onChange={handleImageUpload} 
                            />
                            <label 
                              htmlFor="profile-upload" 
                              className="w-full flex items-center justify-center bg-blue-100 hover:bg-blue-200 text-blue-700 py-2.5 rounded-lg cursor-pointer text-sm font-semibold transition shadow-sm border border-blue-200"
                            >
                              <Upload size={16} className="mr-2" /> อัปโหลดรูปภาพ (แนะนำ)
                            </label>
                          </div>

                          <div className="w-full flex items-center mb-3">
                            <div className="h-px bg-slate-200 flex-1"></div>
                            <span className="px-2 text-[10px] text-slate-400 font-medium tracking-wide border border-slate-200 rounded-full py-0.5 bg-slate-50">หรือใช้ลิงก์ภาพเดิม</span>
                            <div className="h-px bg-slate-200 flex-1"></div>
                          </div>

                          <input 
                            type="text" name="photoUrl" value={profile.photoUrl || ''} onChange={handleProfileChange}
                            placeholder="URL รูปภาพ (เช่น Image Link)"
                            className="w-full text-xs p-2.5 border border-slate-300 rounded focus:ring-2 focus:ring-blue-500 mb-4 bg-slate-50"
                          />

                          <div className="w-full bg-slate-50 rounded-xl border border-slate-200 mb-1 overflow-hidden transition-all duration-300">
                            <button 
                              type="button"
                              onClick={() => setShowImageTools(!showImageTools)}
                              className="w-full flex items-center justify-between p-3.5 text-sm font-bold text-slate-700 bg-slate-100/50 hover:bg-slate-100 transition-colors"
                            >
                              <span className="flex items-center"><Palette size={16} className="mr-2 text-blue-500" /> แต่งกรอบรูปภาพ (ซูม/เลื่อน)</span>
                              {showImageTools ? <ChevronUp size={16} className="text-slate-400" /> : <ChevronDown size={16} className="text-slate-400" />}
                            </button>
                            
                            {showImageTools && (
                              <div className="p-4 space-y-4 border-t border-slate-200 bg-white">
                                <div>
                                  <label className="flex justify-between text-xs font-medium text-slate-700 mb-1">
                                    <span>ขนาดวงกลม (กว้าง)</span>
                                    <span className="text-blue-600 font-bold">{profile.imageSize || 160}px</span>
                                  </label>
                                  <input type="range" name="imageSize" min="80" max="300" value={profile.imageSize || 160} onChange={handleProfileChange} className="w-full accent-blue-600 cursor-pointer" />
                                </div>

                                <div className="grid grid-cols-2 gap-4">
                                  <div>
                                    <label className="flex justify-between text-xs font-medium text-slate-700 mb-1">
                                      <span>ขนาดขอบ (px)</span>
                                      <span className="text-blue-600 font-bold">{profile.imageBorderWidth ?? 4}px</span>
                                    </label>
                                    <input type="range" name="imageBorderWidth" min="0" max="25" value={profile.imageBorderWidth ?? 4} onChange={handleProfileChange} className="w-full accent-blue-600 cursor-pointer" />
                                  </div>
                                  <div>
                                    <label className="block text-xs font-medium text-slate-700 mb-1">สีขอบรูป</label>
                                    <input type="color" name="imageBorderColor" value={profile.imageBorderColor || '#f1f5f9'} onChange={handleProfileChange} className="w-full h-8 p-0 border border-slate-200 rounded cursor-pointer" />
                                  </div>
                                </div>

                                <div>
                                  <label className="flex justify-between text-xs font-medium text-slate-700 mb-1">
                                    <span>ซูมภาพ (Zoom)</span>
                                    <span className="text-blue-600 font-bold">{parseFloat(profile.imageZoom || 1).toFixed(1)}x</span>
                                  </label>
                                  <input type="range" name="imageZoom" min="1" max="4" step="0.1" value={profile.imageZoom || 1} onChange={handleProfileChange} className="w-full accent-blue-600 cursor-pointer" />
                                </div>

                                <div>
                                  <label className="flex justify-between text-xs font-medium text-slate-700 mb-1">
                                    <span>เลื่อน ซ้าย-ขวา</span>
                                  </label>
                                  <input type="range" name="imagePanX" min="0" max="100" value={profile.imagePanX ?? 50} onChange={handleProfileChange} className="w-full accent-blue-600 cursor-pointer" />
                                </div>

                                <div>
                                  <label className="flex justify-between text-xs font-medium text-slate-700 mb-1">
                                    <span>เลื่อน ขึ้น-ลง</span>
                                  </label>
                                  <input type="range" name="imagePanY" min="0" max="100" value={profile.imagePanY ?? 50} onChange={handleProfileChange} className="w-full accent-blue-600 cursor-pointer" />
                                </div>
                                
                                <button 
                                  type="button" 
                                  onClick={() => {
                                    const newProfile = { 
                                      ...profile, 
                                      imageSize: 160,
                                      imageZoom: 1, 
                                      imagePanX: 50, 
                                      imagePanY: 50,
                                      imageBorderWidth: 4,
                                      imageBorderColor: '#f1f5f9'
                                    };
                                    setProfile(newProfile);
                                    saveProfileToFirebase(newProfile, settings);
                                  }}
                                  className="w-full text-xs bg-slate-200 hover:bg-slate-300 text-slate-700 py-1.5 rounded-lg transition font-medium"
                                >
                                  คืนค่าเริ่มต้น (Reset Crop)
                                </button>
                              </div>
                            )}
                          </div>
                          <p className="text-[11px] text-slate-400 mt-2 text-center">หากใช้ลิงก์ Google Drive ระบบจะดึงภาพความละเอียดสูงให้อัตโนมัติ</p>
                        </div>
                        <div className="col-span-1 md:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div className="md:col-span-2">
                            <label className="block text-sm font-medium text-slate-700 mb-1">ชื่อ-นามสกุล</label>
                            <input type="text" name="name" value={profile.name} onChange={handleProfileChange} className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500" />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-700 mb-1">ตำแหน่ง / วิทยฐานะ</label>
                            <input type="text" name="position" value={profile.position} onChange={handleProfileChange} className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500" />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-700 mb-1">โรงเรียน / หน่วยงาน</label>
                            <input type="text" name="school" value={profile.school} onChange={handleProfileChange} className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500" />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-700 mb-1">อีเมล</label>
                            <input type="email" name="email" value={profile.email || ''} onChange={handleProfileChange} className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500" />
                          </div>
                          <div>
                            <label className="block text-sm font-medium text-slate-700 mb-1">เบอร์โทรศัพท์</label>
                            <input type="text" name="phone" value={profile.phone || ''} onChange={handleProfileChange} className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500" />
                          </div>
                          <div className="md:col-span-2">
                            <label className="block text-sm font-medium text-slate-700 mb-1 flex items-center text-blue-700">
                              <Monitor size={16} className="mr-2" /> ลิงก์ Google Drive สำหรับเก็บเอกสาร (หลักฐานรวม)
                            </label>
                            <input 
                              type="text" 
                              name="googleDriveLink" 
                              value={profile.googleDriveLink || ''} 
                              onChange={handleProfileChange} 
                              placeholder="เช่น https://drive.google.com/drive/folders/..."
                              className="w-full p-2 border border-blue-200 bg-blue-50/30 rounded-md focus:ring-2 focus:ring-blue-500 text-blue-800 font-medium" 
                            />
                            <p className="text-[10px] text-slate-400 mt-1 italic">วางลิงก์โฟลเดอร์ Google Drive ที่คุณครูใช้เก็บไฟล์หลักฐานทั้งหมด เพื่อความสะดวกในการเข้าถึง</p>
                          </div>
                          <div className="md:col-span-2">
                            <label className="block text-sm font-medium text-slate-700 mb-1">ประวัติย่อ / คติพจน์</label>
                            <textarea name="bio" value={profile.bio} onChange={handleProfileChange} rows="3" className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500"></textarea>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {activeTab === 'settings' && (
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
                       <h3 className="text-lg font-bold text-slate-800 mb-6 border-b pb-2 flex items-center">
                          <Eye className="mr-2" size={20} /> เปิด/ปิด หมวดหมู่ระบบ
                       </h3>
                       <div className="space-y-4">
                          {Object.values(SCHEMA).map(cat => {
                            const settingKey = `show${cat.id.charAt(0).toUpperCase() + cat.id.slice(1)}`;
                            const isChecked = settings[settingKey] !== false;
                            const SettingIcon = cat.icon;
                            return (
                              <div key={cat.id} className="flex items-center justify-between p-3 hover:bg-slate-50 rounded-lg">
                                <div className="flex items-center">
                                  <SettingIcon className="text-slate-400 mr-3" size={20} />
                                  <span className="font-medium text-slate-700">หมวดหมู่: {cat.name}</span>
                                </div>
                                <label className="relative inline-flex items-center cursor-pointer">
                                  <input 
                                    type="checkbox" 
                                    className="sr-only peer" 
                                    checked={isChecked} 
                                    onChange={() => {
                                      handleSettingToggle(settingKey);
                                      if (isChecked && activeTab === cat.id) {
                                        setActiveTab('settings');
                                      }
                                    }} 
                                  />
                                  <div className="w-11 h-6 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
                                  <span className="ml-3 text-sm font-medium text-slate-500">{isChecked ? 'แสดง' : 'ซ่อน'}</span>
                                </label>
                              </div>
                            )
                          })}
                       </div>
                       <p className="mt-5 text-sm text-slate-500 flex items-start">
                          <Eye size={16} className="mr-2 mt-0.5 shrink-0" /> การปิดการแสดงผล จะซ่อนหมวดหมู่นั้นจากเมนูด้านซ้ายและหน้าพรีวิว แต่ข้อมูลที่เคยบันทึกไว้จะยังคงอยู่ในระบบ ไม่สูญหาย
                       </p>
                    </div>
                  )}

                  {activeTab === 'theme' && (
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
                       <h3 className="text-lg font-bold text-slate-800 mb-6 border-b pb-2 flex items-center">
                          <Palette className="mr-2" size={20} /> สีสันของระบบ (Theme)
                       </h3>
                       <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
                          {THEME_OPTIONS.map(theme => {
                             const isActive = (settings.themeColor || 'blue') === theme.id;
                             return (
                               <button
                                 key={theme.id}
                                 onClick={() => handleThemeChange(theme.id)}
                                 className={`flex items-center p-2 rounded-lg border-2 transition text-sm ${isActive ? 'border-blue-500 bg-blue-50' : 'border-slate-200 hover:border-slate-300'}`}
                               >
                                 <div className="w-5 h-5 rounded-full mr-2 shadow-sm border border-slate-200 shrink-0" style={{ backgroundColor: theme.colorCode }}></div>
                                 <span className={`font-medium truncate ${isActive ? 'text-blue-700' : 'text-slate-700'}`}>{theme.label}</span>
                               </button>
                             )
                          })}
                       </div>
                       <p className="mt-5 text-sm text-slate-500 flex items-center border-t border-slate-100 pt-4">
                          <Eye size={16} className="mr-2 text-slate-400" /> สีที่เลือกจะถูกนำไปใช้เป็นแถบเมนู และปุ่มต่างๆ ทั้งระบบโดยอัตโนมัติ
                       </p>
                    </div>
                  )}
                </div>
              )
            )}

            {Object.keys(SCHEMA).includes(activeTab) && (
              <div className="max-w-6xl mx-auto">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
                  <div>
                    <h3 className="text-lg font-bold text-slate-800">จัดการข้อมูล{SCHEMA[activeTab].name}</h3>
                    <p className="text-slate-500 text-xs">รายการทั้งหมด {portfolioData[activeTab]?.length || 0} รายการ</p>
                  </div>
                  
                  <div className="flex flex-wrap items-center gap-2">
                    {/* Year Filter for Admin View */}
                    {getUniqueYears(portfolioData[activeTab]).length > 0 && (
                      <div className="flex items-center bg-white border border-slate-200 rounded-lg px-2 py-1 shadow-sm">
                        <Filter size={14} className="text-slate-400 mr-2" />
                        <select 
                          className="text-sm bg-transparent border-none focus:ring-0 text-slate-700 font-medium"
                          value={yearFilter}
                          onChange={(e) => setYearFilter(e.target.value)}
                        >
                          <option value="all">ทุกปี พ.ศ.</option>
                          {getUniqueYears(portfolioData[activeTab]).map(y => (
                            <option key={y} value={y}>ปี {y}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <button 
                      onClick={() => handleExportCategoryExcel(activeTab)}
                      disabled={!portfolioData[activeTab]?.length}
                      className={`bg-white border border-slate-200 hover:bg-green-50 text-slate-700 px-3 py-2 rounded-lg flex items-center text-sm shadow-sm transition ${!portfolioData[activeTab]?.length ? 'opacity-50 cursor-not-allowed' : ''}`}
                    >
                      <FileSpreadsheet size={16} className="mr-2 text-green-600" /> 
                      <span className="hidden sm:inline">Excel</span>
                    </button>
                    <button 
                      onClick={() => setPrintModalState({ target: 'category', categoryId: activeTab })}
                      disabled={isPrintingCategory || !portfolioData[activeTab]?.length}
                      className={`bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 px-3 py-2 rounded-lg flex items-center text-sm shadow-sm transition ${isPrintingCategory || !portfolioData[activeTab]?.length ? 'opacity-50 cursor-not-allowed' : ''}`}
                    >
                      <Printer size={16} className="mr-2" /> 
                      <span className="hidden sm:inline">{isPrintingCategory ? 'กำลังสร้าง...' : 'ดาวน์โหลด PDF'}</span>
                    </button>
                    {isAdminAuthenticated && (
                      <button 
                        onClick={() => openModal(activeTab)}
                        className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg flex items-center text-sm shadow-sm transition"
                      >
                        <Plus size={16} className="mr-2" /> เพิ่มข้อมูล
                      </button>
                    )}
                  </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                  {portfolioData[activeTab]?.length > 0 ? (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200">
                            <th className="p-4 text-sm font-semibold text-slate-600 w-12">#</th>
                            {SCHEMA[activeTab].fields.slice(0, 3).map(field => (
                              <th key={field.key} className="p-4 text-sm font-semibold text-slate-600">{field.label}</th>
                            ))}
                            {SCHEMA[activeTab].fields.some(f => f.key === 'fileUrl') && (
                              <th className="p-4 text-sm font-semibold text-slate-600 text-center w-28">หลักฐาน</th>
                            )}
                            {isAdminAuthenticated && (
                              <th className="p-4 text-sm font-semibold text-slate-600 text-center w-32">จัดการ</th>
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {getSortedData(portfolioData[activeTab], SCHEMA[activeTab], yearFilter).map((item, index) => (
                            <tr key={item.id} className="border-b border-slate-100 hover:bg-slate-50">
                              <td className="p-4 text-sm text-slate-500">{index + 1}</td>
                              {SCHEMA[activeTab].fields.slice(0, 3).map(field => {
                                const displayValue = field.type === 'date' && item[field.key] ? formatThaiDate(item[field.key]) : (item[field.key] || '-');
                                return (
                                  <td key={field.key} className="p-4 text-sm text-slate-800 font-medium">
                                    {displayValue}
                                  </td>
                                );
                              })}
                              {SCHEMA[activeTab].fields.some(f => f.key === 'fileUrl') && (
                                <td className="p-4 text-center">
                                  {item.fileUrl ? (
                                    <a href={item.fileUrl} target="_blank" rel="noopener noreferrer" className="inline-flex items-center text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-3 py-1.5 rounded-full text-xs font-medium transition-colors">
                                      <LinkIcon size={14} className="mr-1" />
                                      เปิดดู
                                    </a>
                                  ) : (
                                    <span className="text-slate-300 text-xs">-</span>
                                  )}
                                </td>
                              )}
                              {isAdminAuthenticated && (
                                <td className="p-4 text-center">
                                  <div className="flex justify-center gap-2">
                                    <button onClick={() => openModal(activeTab, item)} className="p-1.5 text-blue-600 hover:bg-blue-50 rounded" title="แก้ไข">
                                      <Edit2 size={16} />
                                    </button>
                                    <button onClick={() => handleDelete(activeTab, item.id, item.title || item.workName || item.courseName || item.topic || 'ข้อมูลนี้')} className="p-1.5 text-red-500 hover:bg-red-50 rounded" title="ลบ">
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="text-center py-16 px-4">
                      <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mx-auto mb-4 text-slate-400">
                        {(() => {
                          const ActiveIcon = SCHEMA[activeTab].icon;
                          return <ActiveIcon size={32} />;
                        })()}
                      </div>
                      <p className="text-slate-500 mb-4">ยังไม่มีข้อมูลในหมวดหมู่นี้</p>
                      {isAdminAuthenticated && (
                        <button onClick={() => openModal(activeTab)} className="text-blue-600 font-medium hover:underline">
                          + เพิ่มข้อมูลแรกของคุณ
                        </button>
                      )}
                    </div>
                  )}
                </div>
              </div>
            )}

          </div>
        </main>

        {isModalOpen && currentCategory && isAdminAuthenticated && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black bg-opacity-50 p-4">
            <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]">
              <div className="p-5 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <h3 className="text-lg font-bold text-slate-800 flex items-center">
                  {editingId ? <Edit2 className="mr-2" size={20} /> : <Plus className="mr-2" size={20} />}
                  {editingId ? 'แก้ไขข้อมูล' : 'เพิ่มข้อมูลใหม่'} - {SCHEMA[currentCategory].name}
                </h3>
                <button onClick={closeModal} className="text-slate-400 hover:text-slate-700">
                  <X size={24} />
                </button>
              </div>
              <div className="p-6 overflow-y-auto flex-1">
                <form id="dynamic-form" onSubmit={handleFormSubmit} className="space-y-4">
                  {SCHEMA[currentCategory].fields.map((field) => (
                    <div key={field.key}>
                      <label className="block text-sm font-medium text-slate-700 mb-1">
                        {field.label} {field.required && <span className="text-red-500">*</span>}
                      </label>
                      {field.type === 'textarea' ? (
                        <textarea
                          required={field.required}
                          value={formData[field.key] || ''}
                          onChange={(e) => handleFormChange(e, field.key)}
                          rows="3"
                          className="w-full p-2.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                          placeholder={`กรอก${field.label}`}
                        ></textarea>
                      ) : (
                        <input
                          type={field.type}
                          required={field.required}
                          value={formData[field.key] || ''}
                          onChange={(e) => handleFormChange(e, field.key)}
                          className="w-full p-2.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                          placeholder={field.type === 'url' ? 'https://...' : `กรอก${field.label}`}
                        />
                      )}
                      {field.type === 'url' && (
                        <p className="text-xs text-slate-400 mt-1">อัปโหลดไฟล์ลง Google Drive แล้วนำลิงก์ (เปิดแชร์) มาวางที่นี่</p>
                      )}
                    </div>
                  ))}

                  {currentCategory !== 'certificates' && (
                    <div className="mt-6 p-4 bg-blue-50 border border-blue-100 rounded-lg flex items-start">
                      <div className="flex-shrink-0 mt-0.5">
                        <input
                          id="addToCertificates"
                          type="checkbox"
                          checked={formData.addToCertificates || false}
                          onChange={(e) => setFormData({ ...formData, addToCertificates: e.target.checked })}
                          className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-slate-300 rounded cursor-pointer"
                        />
                      </div>
                      <div className="ml-3">
                        <label htmlFor="addToCertificates" className="text-sm font-medium text-slate-700 cursor-pointer">
                          มีการได้รับเกียรติบัตรในรายการนี้ด้วย (เพิ่มลงแฟ้มเกียรติบัตรอัตโนมัติ)
                        </label>
                        <p className="text-xs text-slate-500 mt-1">
                          ระบบจะนำข้อมูลที่คุณกรอกทั้งหมด (รวมลิงก์ไฟล์) ไปสร้างข้อมูลใหม่ในหมวดหมู่ "เกียรติบัตร" ให้อีกหนึ่งรายการ เพื่อลดการกรอกข้อมูลซ้ำซ้อน
                        </p>
                      </div>
                    </div>
                  )}
                </form>
              </div>
              <div className="p-4 border-t border-slate-200 flex justify-end gap-3 bg-slate-50">
                <button onClick={closeModal} className="px-4 py-2 text-slate-600 hover:bg-slate-200 rounded-lg transition">
                  ยกเลิก
                </button>
                <button type="submit" form="dynamic-form" className="px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg shadow transition flex items-center">
                  <Save size={18} className="mr-2" /> บันทึกข้อมูล
                </button>
              </div>
            </div>
          </div>
        )}

      </div>
      
      {/* UI Selection Modal for Print Orientation */}
      {printModalState && (
        <div className="fixed inset-0 bg-slate-900/50 z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden">
            <div className="bg-blue-600 p-4 text-white">
              <h3 className="text-lg font-bold flex items-center"><Printer className="mr-2" size={20} /> รูปแบบการดาวน์โหลด PDF</h3>
            </div>
            
            <div className="p-6">
              {printModalState.target === 'category' && (
                <div className="mb-6">
                  <label className="block text-sm font-bold text-slate-700 mb-2 flex items-center">
                    <Filter size={14} className="mr-2" /> เลือกช่วงเวลา (ปี พ.ศ.)
                  </label>
                  <div className="grid grid-cols-3 gap-2">
                    <button 
                      onClick={() => setYearFilter('all')}
                      className={`py-2 px-3 rounded-lg text-sm font-medium border transition ${yearFilter === 'all' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
                    >
                      ทั้งหมด
                    </button>
                    {getUniqueYears(portfolioData[printModalState.categoryId]).map(y => (
                      <button 
                        key={y}
                        onClick={() => setYearFilter(y)}
                        className={`py-2 px-3 rounded-lg text-sm font-medium border transition ${yearFilter === y ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
                      >
                        ปี {y}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              <label className="block text-sm font-bold text-slate-700 mb-3 flex items-center">
                <Monitor size={14} className="mr-2" /> การวางแนวหน้ากระดาษ
              </label>
              <div className="grid grid-cols-2 gap-4 mb-6">
                <button 
                  onClick={() => {
                     const { target, categoryId } = printModalState;
                     setPrintModalState(null);
                     if (target === 'overview') handlePrintPDF('portrait');
                     else handlePrintCategoryPDF(categoryId, 'portrait', yearFilter);
                  }}
                  className="flex flex-col items-center justify-center p-4 border-2 border-slate-200 rounded-xl hover:bg-blue-50 hover:border-blue-500 transition group"
                >
                  <div className="border-2 border-slate-300 w-10 h-14 rounded-sm mb-3 transition group-hover:border-blue-500 flex items-center justify-center">
                    <div className="w-6 h-0.5 bg-slate-200 group-hover:bg-blue-200 mb-1"></div>
                    <div className="w-6 h-0.5 bg-slate-200 group-hover:bg-blue-200"></div>
                  </div>
                  <span className="text-sm font-bold text-slate-700 group-hover:text-blue-700">แนวตั้ง (Portrait)</span>
                </button>
                <button 
                  onClick={() => {
                     const { target, categoryId } = printModalState;
                     setPrintModalState(null);
                     if (target === 'overview') handlePrintPDF('landscape');
                     else handlePrintCategoryPDF(categoryId, 'landscape', yearFilter);
                  }}
                  className="flex flex-col items-center justify-center p-4 border-2 border-slate-200 rounded-xl hover:bg-blue-50 hover:border-blue-500 transition group"
                >
                  <div className="border-2 border-slate-300 w-14 h-10 rounded-sm mb-3 transition group-hover:border-blue-500 flex flex-col items-center justify-center">
                    <div className="w-10 h-0.5 bg-slate-200 group-hover:bg-blue-200 mb-1"></div>
                    <div className="w-10 h-0.5 bg-slate-200 group-hover:bg-blue-200"></div>
                  </div>
                  <span className="text-sm font-bold text-slate-700 group-hover:text-blue-700">แนวนอน (Landscape)</span>
                </button>
              </div>

              <button 
                onClick={() => setPrintModalState(null)}
                className="w-full py-2.5 text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-lg transition font-bold text-sm"
              >
                ยกเลิก
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Delete Confirmation Modal */}
      {deleteConfirm && (
        <div className="fixed inset-0 bg-slate-900/50 z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="bg-red-50 p-6 text-center">
              <div className="w-16 h-16 bg-red-100 text-red-600 rounded-full flex items-center justify-center mx-auto mb-4">
                <Trash2 size={32} />
              </div>
              <h3 className="text-xl font-bold text-slate-800 mb-2">ยืนยันการลบข้อมูล</h3>
              <p className="text-slate-600 text-sm">
                คุณแน่ใจหรือไม่ว่าต้องการลบ <br/>
                <span className="font-bold text-red-600">"{deleteConfirm.title}"</span>? <br/>
                การกระทำนี้ไม่สามารถย้อนกลับได้
              </p>
            </div>
            <div className="p-4 bg-slate-50 border-t border-slate-100 flex gap-3">
              <button 
                onClick={() => setDeleteConfirm(null)}
                className="flex-1 py-2.5 bg-white border border-slate-200 text-slate-700 rounded-lg font-bold hover:bg-slate-50 transition"
              >
                ยกเลิก
              </button>
              <button 
                onClick={confirmDelete}
                className="flex-1 py-2.5 bg-red-600 text-white rounded-lg font-bold hover:bg-red-700 transition shadow-sm"
              >
                ลบข้อมูล
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
