import React, { useState, useEffect } from 'react';
import { 
  Printer, ArrowLeft, GraduationCap, User, FileText, School, 
  CheckCircle2, Loader2, ChevronRight, Edit2, Search, Download, Lock, Eye, RefreshCw, AlertTriangle, Trash2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { auth, db } from './lib/firebase';
import { signInAnonymously, onAuthStateChanged, User as FirebaseUser } from 'firebase/auth';
import { collection, doc, setDoc, serverTimestamp, query, where, getDocs, deleteDoc } from 'firebase/firestore';
import { OperationType, handleFirestoreError } from './lib/firebase-utils';
import { cn } from './lib/utils';
import * as XLSX from 'xlsx';

type Achievement = 'EXCELLENT' | 'GOOD' | 'COMPLETED' | '';

interface FormData {
  fullName: string;
  dob: string;
  pobProvince: string;
  gender: string;
  idCard: string;
  school: string;
  studentClass: string;
  address: string;
  parentPhone: string;
  fatherName: string;
  fatherJob: string;
  fatherPhone: string;
  motherName: string;
  motherJob: string;
  motherPhone: string;
  grades: {
    1: Achievement;
    2: Achievement;
    3: Achievement;
    4: Achievement;
    5: Achievement;
  };
}

const initialFormData: FormData = {
  fullName: '',
  dob: '',
  pobProvince: '',
  gender: 'Nam',
  idCard: '',
  school: '',
  studentClass: '',
  address: '',
  parentPhone: '',
  fatherName: '',
  fatherJob: '',
  fatherPhone: '',
  motherName: '',
  motherJob: '',
  motherPhone: '',
  grades: { 1: '', 2: '', 3: '', 4: '', 5: '' },
};

const achievementOptions = [
  { value: 'EXCELLENT', label: 'Hoàn thành xuất sắc', score: 10.0 },
  { value: 'GOOD', label: 'Hoàn thành tốt', score: 9.0 },
  { value: 'COMPLETED', label: 'Hoàn thành', score: 8.0 },
];

export default function App() {
  const [formData, setFormData] = useState<FormData>(initialFormData);
  const [isSubmitted, setIsSubmitted] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [viewState, setViewState] = useState<'home' | 'form' | 'print' | 'search' | 'info' | 'admin'>('home');
  const [searchQuery, setSearchQuery] = useState({ name: '', dob: '' });
  const [isSearching, setIsSearching] = useState(false);
  const [searchError, setSearchError] = useState('');
  const [formValidationError, setFormValidationError] = useState('');

  // Admin and CSV Export States
  const [isAdminAuthenticated, setIsAdminAuthenticated] = useState(false);
  const [adminPasswordInput, setAdminPasswordInput] = useState('');
  const [adminError, setAdminError] = useState('');
  const [submissionsList, setSubmissionsList] = useState<any[]>([]);
  const [isLoadingSubmissions, setIsLoadingSubmissions] = useState(false);
  const [adminSearchQuery, setAdminSearchQuery] = useState('');

  // Deletion States
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [deletingName, setDeletingName] = useState<string>('');
  const [isDeleting, setIsDeleting] = useState(false);

  // Get filtered student list shown in the dashboard table and matched with Excel/CSV export
  const filteredSubmissions = submissionsList.filter(sub => {
    const q = adminSearchQuery.toLowerCase().trim();
    if (!q) return true;
    return (
      (sub.fullName || '').toLowerCase().includes(q) ||
      (sub.idCard || '').toLowerCase().includes(q) ||
      (sub.parentPhone || '').toLowerCase().includes(q) ||
      (sub.school || '').toLowerCase().includes(q) ||
      (sub.studentClass || '').toLowerCase().includes(q)
    );
  });

  const fetchSubmissions = async () => {
    setIsLoadingSubmissions(true);
    try {
      const querySnapshot = await getDocs(collection(db, 'submissions'));
      const list: any[] = [];
      querySnapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() });
      });
      // Sort: newest first
      list.sort((a, b) => {
        const dateA = a.createdAt && typeof a.createdAt.toDate === 'function' ? a.createdAt.toDate().getTime() : 0;
        const dateB = b.createdAt && typeof b.createdAt.toDate === 'function' ? b.createdAt.toDate().getTime() : 0;
        return dateB - dateA;
      });
      setSubmissionsList(list);
    } catch (err) {
      console.error("Lỗi khi tải danh sách: ", err);
    } finally {
      setIsLoadingSubmissions(false);
    }
  };

  useEffect(() => {
    if (viewState === 'admin' && isAdminAuthenticated) {
      fetchSubmissions();
    }
  }, [viewState, isAdminAuthenticated]);

  const handleAdminLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (adminPasswordInput === 'admin123' || adminPasswordInput.toLowerCase() === 'thcsnguyenkhuyen') {
      setIsAdminAuthenticated(true);
      setAdminError('');
    } else {
      setAdminError('Mật khẩu quản trị không chính xác.');
    }
  };

  const exportToExcel = () => {
    if (filteredSubmissions.length === 0) return;
    
    const headers = [
      'STT',
      'Họ và tên học sinh',
      'Ngày sinh',
      'Giới tính',
      'Nơi sinh (Tỉnh/Thành phố)',
      'Số CCCD / Định danh',
      'Trường Tiểu học',
      'Lớp',
      'Địa chỉ thường trú',
      'Số điện thoại liên hệ',
      'Họ tên cha',
      'Nghề nghiệp cha',
      'Số điện thoại cha',
      'Họ tên mẹ',
      'Nghề nghiệp mẹ',
      'Số điện thoại mẹ',
      'Mức Đạt Lớp 1 (Quy đổi)',
      'Mức Đạt Lớp 2 (Quy đổi)',
      'Mức Đạt Lớp 3 (Quy đổi)',
      'Mức Đạt Lớp 4 (Quy đổi)',
      'Mức Đạt Lớp 5 (Quy đổi)',
      'Tổng Điểm Xét Tuyển',
      'Ngày đăng ký'
    ];

    const data = [
      headers,
      ...filteredSubmissions.map((sub, index) => {
        const formatDob = sub.dob ? sub.dob.split('-').reverse().join('/') : '';
        const registrationDate = sub.createdAt && typeof sub.createdAt.toDate === 'function' 
          ? sub.createdAt.toDate().toLocaleDateString('vi-VN') 
          : sub.createdAt ? new Date(sub.createdAt).toLocaleDateString('vi-VN') : '';

        const contactPhone = sub.parentPhone || sub.fatherPhone || sub.motherPhone || '';

        return [
          index + 1,
          sub.fullName ? sub.fullName.toUpperCase() : '',
          formatDob,
          sub.gender || '',
          sub.pobProvince || '',
          sub.idCard || '',
          sub.school || '',
          sub.studentClass || '',
          sub.address || '',
          contactPhone,
          sub.fatherName || '',
          sub.fatherJob || '',
          sub.fatherPhone || '',
          sub.motherName || '',
          sub.motherJob || '',
          sub.motherPhone || '',
          sub.grades?.[1] ? `${getLabel(sub.grades[1])} (${getScore(sub.grades[1])})` : '',
          sub.grades?.[2] ? `${getLabel(sub.grades[2])} (${getScore(sub.grades[2])})` : '',
          sub.grades?.[3] ? `${getLabel(sub.grades[3])} (${getScore(sub.grades[3])})` : '',
          sub.grades?.[4] ? `${getLabel(sub.grades[4])} (${getScore(sub.grades[4])})` : '',
          sub.grades?.[5] ? `${getLabel(sub.grades[5])} (${getScore(sub.grades[5])})` : '',
          sub.totalScore || 0,
          registrationDate
        ];
      })
    ];

    // Create a new workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Set auto col widths or pre-defined comfortable widths to fit all data
    const wscols = [
      { wch: 6 },   // STT
      { wch: 25 },  // Họ và tên học sinh
      { wch: 12 },  // Ngày sinh
      { wch: 10 },  // Giới tính
      { wch: 22 },  // Nơi sinh - Tỉnh
      { wch: 20 },  // Số CCCD / Định danh
      { wch: 30 },  // Trường Tiểu học
      { wch: 10 },  // Lớp
      { wch: 35 },  // Địa chỉ thường trú
      { wch: 18 },  // Số điện thoại liên hệ
      { wch: 20 },  // Họ tên cha
      { wch: 20 },  // Nghề nghiệp cha
      { wch: 18 },  // Số điện thoại cha
      { wch: 20 },  // Họ tên mẹ
      { wch: 20 },  // Nghề nghiệp mẹ
      { wch: 18 },  // Số điện thoại mẹ
      { wch: 22 },  // Mức Đạt Lớp 1
      { wch: 22 },  // Mức Đạt Lớp 2
      { wch: 22 },  // Mức Đạt Lớp 3
      { wch: 22 },  // Mức Đạt Lớp 4
      { wch: 22 },  // Mức Đạt Lớp 5
      { wch: 18 },  // Tổng Điểm Xét Tuyển
      { wch: 15 }   // Ngày đăng ký
    ];
    ws['!cols'] = wscols;

    // Append worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Danh sách đăng ký');

    // Generate buffer & download standard .xlsx file
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    XLSX.writeFile(wb, `DSTuyenSinh_THCS_NguyenKhuyen_${today}.xlsx`);
  };


  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const handleGradeChange = (grade: 1 | 2 | 3 | 4 | 5, value: Achievement) => {
    setFormData((prev) => ({
      ...prev,
      grades: { ...prev.grades, [grade]: value },
    }));
  };

  const getScore = (ach: Achievement) => achievementOptions.find((opt) => opt.value === ach)?.score || 0;
  const getLabel = (ach: Achievement) => achievementOptions.find((opt) => opt.value === ach)?.label || '';

  const totalScore = [1, 2, 3, 4, 5].reduce(
    (acc, grade) => acc + getScore(formData.grades[grade as 1 | 2 | 3 | 4 | 5]),
    0
  );

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setFormValidationError('');

    const emptyFields: string[] = [];
    if (!formData.fullName.trim()) emptyFields.push("Họ tên học sinh");
    if (!formData.dob.trim()) emptyFields.push("Ngày sinh");
    if (!formData.pobProvince.trim()) emptyFields.push("Nơi sinh (Tỉnh)");
    if (!formData.idCard.trim()) emptyFields.push("Số CCCD/Mã định danh");
    if (!formData.school.trim()) emptyFields.push("Trường Tiểu học");
    if (!formData.studentClass.trim()) emptyFields.push("Lớp");
    if (!formData.address.trim()) emptyFields.push("Địa chỉ thường trú");
    if (!formData.fatherName.trim()) emptyFields.push("Họ tên cha");
    if (!formData.fatherJob.trim()) emptyFields.push("Nghề nghiệp cha");
    if (!formData.fatherPhone.trim()) emptyFields.push("Số điện thoại cha");
    if (!formData.motherName.trim()) emptyFields.push("Họ tên mẹ");
    if (!formData.motherJob.trim()) emptyFields.push("Nghề nghiệp mẹ");
    if (!formData.motherPhone.trim()) emptyFields.push("Số điện thoại mẹ");

    const missingGrades = [1, 2, 3, 4, 5].filter(g => !formData.grades[g as 1 | 2 | 3 | 4 | 5]);
    if (missingGrades.length > 0) {
      emptyFields.push("Kết quả học tập các năm học (Lớp " + missingGrades.join(", Lớp ") + ")");
    }

    if (emptyFields.length > 0) {
      setFormValidationError(`Vui lòng nhập đầy đủ tất cả thông tin bắt buộc: ${emptyFields.join(", ")}.`);
      return;
    }

    setIsSaving(true);
    try {
      // 1. Kiểm tra trùng Số định danh/CCCD
      const idCardToCheck = formData.idCard.trim();
      const idCardQuery = query(
        collection(db, 'submissions'),
        where('idCard', '==', idCardToCheck)
      );
      const idCardSnapshot = await getDocs(idCardQuery);
      if (!idCardSnapshot.empty) {
        setFormValidationError(`Số căn cước công dân / định danh "${idCardToCheck}" đã được đăng ký xét tuyển trên hệ thống trước đó! Vui lòng kiểm tra lại.`);
        setIsSaving(false);
        return;
      }

      // 2. Kiểm tra trùng Họ tên và Ngày sinh (Chống trùng tên)
      const fullNameUpper = formData.fullName.trim().toUpperCase();
      const dobToCheck = formData.dob.trim();
      const duplicateNameQuery = query(
        collection(db, 'submissions'),
        where('fullName', '==', fullNameUpper),
        where('dob', '==', dobToCheck)
      );
      const nameSnapshot = await getDocs(duplicateNameQuery);
      if (!nameSnapshot.empty) {
        const readableDob = dobToCheck.split('-').reverse().join('/');
        setFormValidationError(`Học sinh "${fullNameUpper}" sinh ngày ${readableDob} đã được đăng ký xét tuyển trên hệ thống! Vui lòng không gửi hồ sơ trùng lặp.`);
        setIsSaving(false);
        return;
      }

      const subId = crypto.randomUUID();
      const payload = {
        ...formData,
        fullName: fullNameUpper,
        parentPhone: formData.fatherPhone || formData.motherPhone || '',
        totalScore,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };
      
      await setDoc(doc(db, 'submissions', subId), payload);
      setViewState('print');
      setIsSubmitted(true);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, 'submissions', auth);
    } finally {
      setIsSaving(false);
    }
  };

  const startNew = () => {
    setFormData(initialFormData);
    setViewState('home');
    setIsSubmitted(false);
    setSearchQuery({ name: '', dob: '' });
    setSearchError('');
    setFormValidationError('');
  };

  const handleSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!searchQuery.name || !searchQuery.dob) {
      setSearchError('Vui lòng nhập đầy đủ thông tin');
      return;
    }

    setIsSearching(true);
    setSearchError('');
    try {
      const q = query(
        collection(db, 'submissions'),
        where('fullName', '==', searchQuery.name.trim().toUpperCase()),
        where('dob', '==', searchQuery.dob)
      );
      
      const querySnapshot = await getDocs(q);
      if (querySnapshot.empty) {
        setSearchError('Không tìm thấy hồ sơ nào khớp với thông tin cung cấp.');
      } else {
        // Lấy hồ sơ đầu tiên (nếu có nhiều cái thì lấy cái mới nhất, nhưng thôi lấy cái đầu)
        const docData = querySnapshot.docs[0].data() as FormData;
        setFormData(docData);
        setViewState('print');
        setIsSubmitted(true);
      }
    } catch (err) {
      console.error(err);
      setSearchError('Đã xảy ra lỗi khi tìm kiếm.');
    } finally {
      setIsSearching(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans flex flex-col">
      
      {/* Top Navbar */}
      {viewState !== 'home' && (
        <nav className="p-4 bg-white sticky top-0 z-10 border-b border-gray-200 shadow-sm no-print flex justify-between items-center px-4 sm:px-8">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg text-white">
              <School className="w-5 h-5" />
            </div>
            <span className="font-bold text-xl text-blue-800 hidden sm:block">
              Hệ Thống Tuyển Sinh
            </span>
          </div>
          <div className="flex items-center gap-4">
            {viewState !== 'admin' && (
              <button 
                onClick={() => setViewState('admin')} 
                className="hidden md:flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-blue-600 hover:bg-blue-50 font-semibold transition text-sm cursor-pointer border-none bg-transparent"
              >
                <Lock className="w-4 h-4" /> Dành cho nhà trường
              </button>
            )}
            <button onClick={() => setViewState('home')} className="text-gray-500 hover:text-gray-900 font-medium transition flex items-center border-none bg-transparent cursor-pointer">
              <ArrowLeft className="w-5 h-5 mr-1" /> Trang chủ
            </button>
          </div>
        </nav>
      )}

      {viewState === 'home' ? (
        <div className="relative flex-1 w-full flex flex-col items-center justify-center bg-gray-900 bg-[url('https://i.postimg.cc/yNGwbLmr/DSC08251.jpg')] bg-cover bg-center">
          <div className="absolute inset-0 bg-black/40"></div>
          
          <div className="relative z-10 w-full px-4 flex flex-col items-center justify-center py-20">
            <div className="text-center mb-16 text-white drop-shadow-md">
              <GraduationCap className="w-24 h-24 mx-auto mb-4 text-white drop-shadow-md" strokeWidth={1.5} />
              <h1 className="text-2xl font-medium mb-2 drop-shadow-md text-white">
                Hệ thống tuyển sinh trực tuyến
              </h1>
              <h2 className="text-4xl md:text-5xl lg:text-6xl font-bold drop-shadow-lg max-w-4xl mx-auto leading-tight mt-2 text-white">
                Trường Trung Học Cơ Sở<br />
                Nguyễn Khuyến, Xã Ea Kar
              </h2>
            </div>

            <div className="w-full max-w-5xl grid grid-cols-1 md:grid-cols-3 gap-[1px] bg-white/20">
              <button onClick={() => setViewState('info')} className="flex flex-col items-center justify-center py-16 px-4 bg-[#2C6E91]/90 hover:bg-[#2C6E91] transition-colors text-white min-h-[200px] border-none outline-none cursor-pointer">
                <GraduationCap className="w-10 h-10 mb-4" />
                <span className="text-xl font-medium">Thông tin tuyển sinh</span>
              </button>

              <button onClick={() => setViewState('form')} className="flex flex-col items-center justify-center py-16 px-4 bg-[#2C6E91]/90 hover:bg-[#2C6E91] transition-colors text-white min-h-[200px] border-none outline-none cursor-pointer">
                <Edit2 className="w-9 h-9 mb-4" />
                <span className="text-xl font-medium">Đăng ký tuyển sinh</span>
              </button>

              <button onClick={() => setViewState('search')} className="flex flex-col items-center justify-center py-16 px-4 bg-[#2C6E91]/90 hover:bg-[#2C6E91] transition-colors text-white min-h-[200px] border-none outline-none cursor-pointer">
                <Search className="w-9 h-9 mb-4" strokeWidth={2.5} />
                <span className="text-xl font-medium">Tra cứu</span>
              </button>
            </div>

            <div className="mt-12 text-center text-white/80 text-xs sm:text-sm drop-shadow-md space-y-2">
              <p>© 2026 Trường THCS Nguyễn Khuyến, Xã Ea Kar, Tỉnh Đắk Lắk.</p>
              <button 
                onClick={() => setViewState('admin')} 
                className="inline-flex items-center gap-1.5 px-4 py-2 mt-2 rounded-full bg-white/10 hover:bg-white/25 text-white font-medium transition cursor-pointer text-xs border-none"
              >
                <Lock className="w-3.5 h-3.5" /> Dành cho nhà trường
              </button>
            </div>
          </div>
        </div>
      ) : (
        <main className="flex-1 py-10 px-4 sm:px-6 lg:px-8 w-full max-w-5xl mx-auto">
          <AnimatePresence mode="wait">
            
            {/* INFO VIEW */}
            {viewState === 'info' && (
              <motion.div key="info" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="max-w-4xl mx-auto">
                <div className="bg-white p-2 shadow-xl border border-gray-200 rounded-2xl overflow-hidden text-center mb-10">
                  <h2 className="text-2xl font-bold text-gray-800 my-6 bg-clip-text text-transparent bg-gradient-to-r from-blue-600 to-indigo-600 uppercase">
                    THÔNG TIN TUYỂN SINH NĂM HỌC
                  </h2>
                  <img src="https://i.postimg.cc/g0yjdVjT/708191303-1015773824251817-8708733752349593381-n.jpg" alt="Thông tin tuyển sinh Nguyễn Khuyến" className="w-full h-auto rounded-xl shadow-sm" />
                </div>
              </motion.div>
            )}

            {/* SEARCH VIEW */}
            {viewState === 'search' && (
              <motion.div key="search" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="max-w-xl mx-auto mt-10">
                <div className="bg-white p-8 shadow-lg border border-gray-200 rounded-xl">
                  <div className="text-center mb-8">
                    <div className="bg-blue-100 text-blue-600 p-3 rounded-full inline-block mb-4">
                      <Search className="w-8 h-8" />
                    </div>
                    <h2 className="text-2xl font-bold text-gray-900">Tra Cứu Hồ Sơ</h2>
                    <p className="text-gray-600 mt-2">Nhập chính xác họ tên và ngày sinh để tra cứu kết quả hồ sơ của bạn.</p>
                  </div>

                  {searchError && (
                    <div className="mb-6 p-4 bg-red-50 text-red-600 border border-red-100 rounded-lg text-sm text-center">
                      {searchError}
                    </div>
                  )}

                  <form onSubmit={handleSearch} className="space-y-6">
                    <div>
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Họ và tên học sinh</label>
                      <input 
                        type="text" 
                        required 
                        placeholder="VD: NGUYỄN VĂN A" 
                        value={searchQuery.name} 
                        onChange={(e) => setSearchQuery({...searchQuery, name: e.target.value})} 
                        className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all uppercase" 
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Ngày sinh</label>
                      <input 
                        type="date" 
                        required 
                        value={searchQuery.dob} 
                        onChange={(e) => setSearchQuery({...searchQuery, dob: e.target.value})} 
                        className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" 
                      />
                    </div>
                    <div className="pt-4">
                      <button 
                        type="submit" 
                        disabled={isSearching} 
                        className="w-full flex justify-center items-center py-3.5 px-4 text-base font-bold rounded-lg text-white bg-blue-600 hover:bg-blue-700 transition-all focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:opacity-70 disabled:cursor-not-allowed gap-2"
                      >
                        {isSearching ? <Loader2 className="w-5 h-5 animate-spin" /> : <Search className="w-5 h-5" />}
                        {isSearching ? 'Đang tìm kiếm...' : 'Tra Cứu Kết Quả'}
                      </button>
                    </div>
                  </form>
                </div>
              </motion.div>
            )}

            {/* ADMIN DASHBOARD VIEW */}
            {viewState === 'admin' && (
              <motion.div key="admin" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} className="w-full max-w-6xl mx-auto">
                {!isAdminAuthenticated ? (
                  /* Passcode Entry Form */
                  <div className="max-w-md mx-auto mt-10">
                    <div className="bg-white p-8 shadow-lg border border-gray-200 rounded-xl">
                      <div className="text-center mb-8">
                        <div className="bg-blue-100 text-blue-600 p-3 rounded-full inline-block mb-4">
                          <Lock className="w-8 h-8" />
                        </div>
                        <h2 className="text-2xl font-bold text-gray-900">Ban Tuyển Sinh Nhà Trường</h2>
                        <p className="text-gray-600 mt-2">Vui lòng nhập mật khẩu quản trị viên để tiếp tục.</p>
                      </div>

                      {adminError && (
                        <div className="mb-6 p-4 bg-red-50 text-red-600 border border-red-100 rounded-lg text-sm text-center">
                          {adminError}
                        </div>
                      )}

                      <form onSubmit={handleAdminLogin} className="space-y-6">
                        <div>
                          <label className="block text-sm font-semibold text-gray-700 mb-2">Mật khẩu bảo mật</label>
                          <input 
                            type="password" 
                            required 
                            placeholder="Nhập mật khẩu..." 
                            value={adminPasswordInput} 
                            onChange={(e) => setAdminPasswordInput(e.target.value)} 
                            className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" 
                          />
                        </div>
                        <div className="pt-2">
                          <button 
                            type="submit" 
                            className="w-full flex justify-center items-center py-3.5 px-4 text-base font-bold rounded-lg text-white bg-blue-600 hover:bg-blue-700 transition-all focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 border-none cursor-pointer"
                          >
                            Đăng Nhập Quản Trị
                          </button>
                        </div>
                      </form>
                    </div>
                  </div>
                ) : (
                  /* Main Dashboard Area */
                  <div className="space-y-6">
                    {/* Welcome Banner */}
                    <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                      <div>
                        <h2 className="text-2xl font-bold text-gray-900">Màn Hình Quản Lý Tuyển Sinh</h2>
                        <p className="text-gray-600 mt-1">Trường THCS Nguyễn Khuyến - Tra cứu thống kê và xuất dữ liệu ban tuyển sinh.</p>
                      </div>
                      <div className="flex gap-2 w-full sm:w-auto">
                        <button 
                          onClick={fetchSubmissions} 
                          disabled={isLoadingSubmissions}
                          className="flex items-center justify-center gap-1.5 px-4 py-2 border border-gray-300 rounded-lg text-gray-700 hover:bg-gray-50 transition font-medium text-sm disabled:opacity-50 cursor-pointer bg-white"
                        >
                          <RefreshCw className={cn("w-4 h-4", isLoadingSubmissions && "animate-spin")} />
                          Tải lại dữ liệu
                        </button>
                        <button 
                          onClick={() => {
                            setIsAdminAuthenticated(false);
                            setAdminPasswordInput('');
                          }} 
                          className="px-4 py-2 rounded-lg text-red-600 hover:bg-red-50 bg-white border border-red-200 transition font-medium text-sm cursor-pointer"
                        >
                          Đăng xuất
                        </button>
                      </div>
                    </div>

                    {/* Stats Panels */}
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                      <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                        <p className="text-gray-500 font-medium text-xs uppercase tracking-wider">Tổng số hồ sơ xét tuyển</p>
                        <p className="text-4xl font-bold text-blue-600 mt-2">{submissionsList.length}</p>
                        <p className="text-xs text-gray-400 mt-1">Hồ sơ đã lưu trên hệ thống</p>
                      </div>
                      <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                        <p className="text-gray-500 font-medium text-xs uppercase tracking-wider font-semibold">Điểm quy đổi TB</p>
                        <p className="text-4xl font-bold text-green-600 mt-2">
                          {(submissionsList.reduce((acc, crr) => acc + (crr.totalScore || 0), 0) / (submissionsList.length || 1)).toFixed(1)}
                        </p>
                        <p className="text-xs text-gray-400 mt-1">Học bạ Lớp 1 - Lớp 5</p>
                      </div>
                      <div className="bg-white p-6 rounded-xl border border-gray-200 shadow-sm">
                        <p className="text-gray-500 font-medium text-xs uppercase tracking-wider font-semibold">Đạt điểm tuyệt đối (50đ)</p>
                        <p className="text-4xl font-bold text-yellow-600 mt-2">
                          {submissionsList.filter(s => s.totalScore === 50).length}
                        </p>
                        <p className="text-xs text-gray-400 mt-1">Gồm 5 năm xuất sắc</p>
                      </div>
                    </div>

                    {/* Action Panel & Table */}
                    <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
                      <div className="p-6 border-b border-gray-100 flex flex-col md:flex-row justify-between items-stretch md:items-center gap-4">
                        <div className="relative flex-1">
                          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-gray-400">
                            <Search className="w-4 h-4" />
                          </span>
                          <input 
                            type="text" 
                            placeholder="Tìm kiếm theo họ tên, CCCD, SĐT học sinh..." 
                            value={adminSearchQuery}
                            onChange={(e) => setAdminSearchQuery(e.target.value)}
                            className="w-full pl-9 pr-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                          />
                        </div>

                        <button 
                          onClick={exportToExcel}
                          disabled={filteredSubmissions.length === 0}
                          className="flex items-center justify-center gap-2 px-5 py-2.5 bg-green-600 hover:bg-green-700 disabled:bg-gray-100 disabled:text-gray-400 disabled:cursor-not-allowed text-white text-sm font-bold rounded-lg transition-all shadow-sm cursor-pointer border-none"
                        >
                          <Download className="w-4 h-4" />
                          Xuất danh sách hiển thị (Excel / CSV)
                        </button>
                      </div>

                      {/* Submissions List Table */}
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-gray-50 border-b border-gray-200 text-gray-500 text-xs uppercase font-medium">
                              <th className="py-3 px-4 text-center">STT</th>
                              <th className="py-3 px-4">Học Sinh</th>
                              <th className="py-3 px-4">Ngày Sinh / Giới</th>
                              <th className="py-3 px-4">Số định danh/CCCD</th>
                              <th className="py-3 px-4">Trường Tiểu Học</th>
                              <th className="py-3 px-4">SĐT Liên Hệ</th>
                              <th className="py-3 px-4 text-center">Tổng Điểm</th>
                              <th className="py-3 px-4 text-center">Hành Động</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100 text-sm">
                            {isLoadingSubmissions ? (
                              <tr>
                                <td colSpan={8} className="py-12 text-center text-gray-500">
                                  <div className="flex flex-col items-center justify-center gap-2">
                                    <Loader2 className="w-8 h-8 animate-spin text-blue-600" />
                                    <span>Đang tải danh sách học sinh đã tuyển sinh...</span>
                                  </div>
                                </td>
                              </tr>
                            ) : submissionsList.length === 0 ? (
                              <tr>
                                <td colSpan={8} className="py-12 text-center text-gray-500">
                                  Chưa có học sinh nào đăng ký xét tuyển.
                                </td>
                              </tr>
                            ) : filteredSubmissions.length === 0 ? (
                              <tr>
                                <td colSpan={8} className="py-12 text-center text-gray-500">
                                  Không tìm thấy hồ sơ học sinh khớp với từ khóa.
                                </td>
                              </tr>
                            ) : filteredSubmissions.map((sub, index) => (
                              <tr key={sub.id || index} className="hover:bg-gray-50/50 transition whitespace-nowrap">
                                <td className="py-4 px-4 text-center font-medium text-gray-500">{index + 1}</td>
                                <td className="py-4 px-4 font-bold text-gray-900 uppercase">{sub.fullName}</td>
                                <td className="py-4 px-4">
                                  <div className="text-gray-700">{sub.dob ? sub.dob.split('-').reverse().join('/') : ''}</div>
                                  <div className="text-xs text-gray-400">{sub.gender}</div>
                                </td>
                                <td className="py-4 px-4 font-mono text-xs font-semibold text-gray-600">{sub.idCard}</td>
                                <td className="py-4 px-4 text-gray-600 max-w-[240px] truncate">
                                  <div>{sub.school}</div>
                                  {sub.studentClass && (
                                    <div className="text-xs text-blue-600 font-semibold mt-0.5">Lớp: {sub.studentClass}</div>
                                  )}
                                </td>
                                <td className="py-4 px-4 text-gray-600">{sub.parentPhone || sub.fatherPhone || sub.motherPhone || 'Chưa cung cấp'}</td>
                                <td className="py-4 px-4 text-center">
                                  <span className="inline-block px-2.5 py-1 text-sm font-bold bg-blue-50 text-blue-700 rounded-full">
                                    {(sub.totalScore || 0).toFixed(1)}đ
                                  </span>
                                </td>
                                <td className="py-4 px-4 text-center">
                                  <div className="flex items-center justify-center gap-1.5">
                                    <button 
                                      onClick={() => {
                                        setFormData(sub);
                                        setViewState('print');
                                        setIsSubmitted(true);
                                      }}
                                      className="inline-flex items-center gap-1 px-3 py-1.5 border border-blue-100 rounded-lg text-blue-600 bg-blue-50/50 hover:bg-blue-50 transition font-medium text-xs cursor-pointer"
                                    >
                                      <Eye className="w-3.5 h-3.5" /> Xem phiếu
                                    </button>
                                    <button 
                                      onClick={() => {
                                        setDeletingId(sub.id);
                                        setDeletingName(sub.fullName);
                                      }}
                                      className="inline-flex items-center gap-1 px-3 py-1.5 border border-red-100 rounded-lg text-red-600 bg-red-50/50 hover:bg-red-50 transition font-medium text-xs cursor-pointer"
                                    >
                                      <Trash2 className="w-3.5 h-3.5" /> Xóa
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {/* FORM VIEW */}
            {(viewState === 'form' && !isSubmitted) && (
            <motion.div key="form" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
              <div className="mb-10 text-center max-w-3xl mx-auto">
                <h1 className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
                  Đăng Ký Hồ Sơ Xét Tuyển
                </h1>
                <p className="text-lg text-gray-600">
                  Điền thông tin cá nhân và kết quả học tập để hệ thống tự động quy đổi điểm.
                </p>
              </div>

              <form onSubmit={handleSubmit} className="bg-white p-6 sm:p-10 shadow-lg border border-gray-200 rounded-xl relative">
                <div className="mb-10">
                  <div className="flex items-center gap-3 mb-6 border-b border-gray-100 pb-4">
                    <div className="bg-blue-100 text-blue-600 p-2 rounded-lg">
                      <User className="w-5 h-5" />
                    </div>
                    <h2 className="text-xl font-bold text-gray-800">
                      I. Thông tin học sinh
                    </h2>
                  </div>
                  
                  <div className="grid grid-cols-1 gap-6 sm:grid-cols-2">
                    <div className="sm:col-span-2">
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Họ và tên học sinh <span className="text-red-500">*</span></label>
                      <input type="text" name="fullName" required placeholder="VD: NGUYỄN VĂN A" value={formData.fullName} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all uppercase" />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Ngày sinh <span className="text-red-500">*</span></label>
                      <input type="date" name="dob" required value={formData.dob} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                    </div>
                    <div>
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Giới tính</label>
                      <div className="relative">
                        <select name="gender" value={formData.gender} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all appearance-none cursor-pointer">
                          <option value="Nam">Nam</option>
                          <option value="Nữ">Nữ</option>
                        </select>
                        <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-4 text-gray-500">
                          <ChevronRight className="w-4 h-4 rotate-90" />
                        </div>
                      </div>
                    </div>
                    <div className="sm:col-span-2">
                       <label className="block text-sm font-semibold text-gray-700 mb-2">Nơi sinh (Tỉnh) <span className="text-red-500">*</span></label>
                      <input type="text" name="pobProvince" required placeholder="VD: Tỉnh Đắk Lắk" value={formData.pobProvince} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Số căn cước công dân (Mã định danh) <span className="text-red-500">*</span></label>
                      <input type="text" name="idCard" required placeholder="" value={formData.idCard} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                    </div>
                    <div className="sm:col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-6">
                      <div className="sm:col-span-2">
                        <label className="block text-sm font-semibold text-gray-700 mb-2">Trường Tiểu học <span className="text-red-500">*</span></label>
                        <input type="text" name="school" required placeholder="VD: Trường Tiểu học Nguyễn Đức Cảnh, xã Ea Kar" value={formData.school} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                      </div>
                      <div>
                        <label className="block text-sm font-semibold text-gray-700 mb-2">Lớp <span className="text-red-500">*</span></label>
                        <input type="text" name="studentClass" required placeholder="VD: 5A" value={formData.studentClass || ''} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                        <div className="flex gap-1.5 mt-2 flex-wrap">
                          {['5A', '5B', '5C', '5D', '5E', '5G'].map((cls) => (
                            <button
                              key={cls}
                              type="button"
                              onClick={() => {
                                setFormData(prev => ({ ...prev, studentClass: cls }));
                              }}
                              className={cn(
                                "text-xs px-2.5 py-1 rounded border transition-all cursor-pointer font-medium",
                                formData.studentClass === cls
                                  ? "bg-blue-600 border-blue-600 text-white shadow-sm"
                                  : "bg-gray-50 hover:bg-gray-100 text-gray-600 border-gray-200"
                              )}
                            >
                              {cls}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block text-sm font-semibold text-gray-700 mb-2">Địa chỉ thường trú <span className="text-red-500">*</span></label>
                      <input type="text" name="address" required placeholder="VD: Thôn 3, Xã Ea Kar, Tỉnh Đắk Lắk" value={formData.address} onChange={handleInputChange} className="w-full p-3 border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                    </div>
                  </div>
                </div>

                {/* II. Thông tin cha, mẹ học sinh */}
                <div className="mb-10">
                  <div className="flex items-center gap-3 mb-6 border-b border-gray-100 pb-4">
                    <div className="bg-blue-100 text-blue-600 p-2 rounded-lg">
                      <User className="w-5 h-5" />
                    </div>
                    <h2 className="text-xl font-bold text-gray-800">
                      II. Thông tin cha, mẹ học sinh
                    </h2>
                  </div>
                  
                  <div className="grid grid-cols-1 gap-6 md:grid-cols-2">
                    {/* Bố */}
                    <div className="bg-gray-50/50 p-5 rounded-lg border border-gray-100 space-y-4">
                      <h3 className="font-bold text-blue-800 text-sm uppercase tracking-wide border-b border-blue-100 pb-2">Thông tin cha (bố)</h3>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Họ tên cha <span className="text-red-500">*</span>:</label>
                        <input type="text" name="fatherName" required placeholder="VD: NGUYỄN VĂN BÌNH" value={formData.fatherName} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all uppercase" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Nghề nghiệp <span className="text-red-500">*</span>:</label>
                        <input type="text" name="fatherJob" required placeholder="VD: Làm nông, Kinh doanh, Giáo viên..." value={formData.fatherJob} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Số điện thoại <span className="text-red-500">*</span>:</label>
                        <input type="tel" name="fatherPhone" required placeholder="VD: 0912345678" value={formData.fatherPhone} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                      </div>
                    </div>

                    {/* Mẹ */}
                    <div className="bg-gray-50/50 p-5 rounded-lg border border-gray-100 space-y-4">
                      <h3 className="font-bold text-blue-800 text-sm uppercase tracking-wide border-b border-blue-100 pb-2">Thông tin mẹ</h3>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Họ tên mẹ <span className="text-red-500">*</span>:</label>
                        <input type="text" name="motherName" required placeholder="VD: TRẦN THỊ MAI" value={formData.motherName} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all uppercase" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Nghề nghiệp <span className="text-red-500">*</span>:</label>
                        <input type="text" name="motherJob" required placeholder="VD: Làm nông, Giáo viên, Nội trợ..." value={formData.motherJob} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-gray-600 mb-1.5 font-sans">Số điện thoại <span className="text-red-500">*</span>:</label>
                        <input type="tel" name="motherPhone" required placeholder="VD: 0987654321" value={formData.motherPhone} onChange={handleInputChange} className="w-full p-2.5 text-sm bg-white border border-gray-300 rounded-md focus:bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all" />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="mb-10">
                  <div className="flex items-center gap-3 mb-6 border-b border-gray-100 pb-4">
                    <div className="bg-blue-100 text-blue-600 p-2 rounded-lg">
                      <FileText className="w-5 h-5" />
                    </div>
                    <h2 className="text-xl font-bold text-gray-800">
                      III. Kết quả học tập
                    </h2>
                  </div>
                  <div className="space-y-4">
                    {[1, 2, 3, 4, 5].map((grade) => (
                      <div key={grade} className="flex flex-col sm:flex-row sm:items-center p-4 border border-gray-200 rounded-lg hover:border-blue-300 transition-colors">
                        <label className="block text-sm font-bold text-gray-800 sm:w-1/4 mb-3 sm:mb-0">
                          Lớp {grade} <span className="text-red-500">*</span>
                        </label>
                        <div className="flex-1 relative">
                          <select
                            required
                            value={formData.grades[grade as 1 | 2 | 3 | 4 | 5]}
                            onChange={(e) => handleGradeChange(grade as 1 | 2 | 3 | 4 | 5, e.target.value as Achievement)}
                            className="w-full p-3 bg-white border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all cursor-pointer appearance-none"
                          >
                            <option value="" disabled>-- Chọn mức đạt được --</option>
                            {achievementOptions.map((opt) => (
                              <option key={opt.value} value={opt.value}>
                                {opt.label} ({opt.score} điểm)
                              </option>
                            ))}
                          </select>
                          <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-4 text-gray-400">
                            <ChevronRight className="w-4 h-4 rotate-90" />
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {formValidationError && (
                  <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm flex items-start gap-2.5 shadow-sm">
                    <AlertTriangle className="w-5 h-5 shrink-0 text-red-500 mt-0.5" />
                    <div>
                      <strong className="font-bold block mb-1">Chưa hoàn tất hồ sơ:</strong>
                      <p className="leading-relaxed">{formValidationError}</p>
                    </div>
                  </div>
                )}

                <div className="pt-6 border-t border-gray-200">
                  <button type="submit" disabled={isSaving} className="w-full flex justify-center items-center py-4 px-4 text-lg font-bold rounded-lg text-white bg-blue-600 hover:bg-blue-700 transition-all focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:opacity-70 disabled:cursor-not-allowed gap-2">
                    {isSaving ? <Loader2 className="w-6 h-6 animate-spin" /> : <CheckCircle2 className="w-6 h-6" />}
                    {isSaving ? 'Đang lưu hệ thống...' : 'Hoàn Tất & In Phiếu'}
                  </button>
                </div>
              </form>
            </motion.div>
          )}

          {/* PRINT VIEW */}
          {(viewState === 'print' || isSubmitted) && (
            <motion.div key="print" initial={{ opacity: 0, scale: 0.98 }} animate={{ opacity: 1, scale: 1 }}>
              <div className="max-w-4xl mx-auto mb-6 flex justify-between items-center bg-white p-4 rounded-lg shadow-sm border border-gray-200 no-print">
                <button
                  onClick={() => window.print()}
                  className="flex items-center px-6 py-2.5 rounded-lg text-white bg-blue-600 hover:bg-blue-700 transition font-medium shadow-sm gap-2 border-none cursor-pointer"
                >
                  <Printer className="w-5 h-5" /> In Phiếu
                </button>
                <div className="flex gap-2">
                  {isAdminAuthenticated && (
                    <button
                      onClick={() => {
                        setViewState('admin');
                        setIsSubmitted(false);
                      }}
                      className="flex items-center px-4 py-2.5 rounded-lg text-gray-700 hover:bg-gray-50 transition font-medium gap-2 border border-gray-300 bg-white cursor-pointer"
                    >
                      <ArrowLeft className="w-4 h-4" /> Quay lại Dashboard
                    </button>
                  )}
                  <button
                    onClick={startNew}
                    className="flex items-center px-4 py-2.5 rounded-lg text-blue-600 hover:bg-blue-50 transition font-medium gap-2 border border-blue-200 bg-white cursor-pointer"
                  >
                    Đăng ký mới
                  </button>
                </div>
              </div>

              {/* Printable Area (Always classic/clean styling for printing) */}
              <div className="max-w-4xl mx-auto bg-white p-12 shadow-lg print-area border border-gray-200 rounded-sm print:border-none print:shadow-none">
                {/* Header */}
                <div className="flex justify-between items-start border-b-2 border-black pb-6 mb-8 text-black">
                  <div className="text-center w-[45%]">
                    <p className="font-bold text-[13px] uppercase tracking-tighter sm:tracking-normal">UỶ BAN NHÂN DÂN XÃ EA KAR</p>
                    <p className="font-bold text-[13px] uppercase underline pb-1">TRƯỜNG THCS NGUYỄN KHUYẾN</p>
                    <p className="text-[11px] italic mt-2">Mã HS: NK-{Math.floor(Math.random() * 900000) + 100000}</p>
                  </div>
                  <div className="text-center w-[55%]">
                    <p className="font-bold text-base uppercase">Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam</p>
                    <p className="font-bold text-base underline pb-2">Độc lập - Tự do - Hạnh phúc</p>
                    <p className="text-sm italic mt-2">Ngày {new Date().getDate()} tháng {new Date().getMonth() + 1} năm {new Date().getFullYear()}</p>
                  </div>
                </div>

                <div className="text-center mb-10 text-black">
                  <h1 className="text-2xl font-bold uppercase tracking-wide mb-2">Phiếu Đăng Ký Xét Tuyển</h1>
                  <p className="text-gray-700 italic">Dựa trên kết quả học bạ Tiểu học (Lớp 1 - Lớp 5)</p>
                </div>

                <div className="mb-8 text-black">
                  <h2 className="text-lg font-bold mb-4 border-b border-black pb-2">I. Thông tin học sinh</h2>
                  <div className="grid grid-cols-2 gap-y-4 gap-x-8 text-base">
                    <div className="col-span-2 sm:col-span-1">
                      <span className="font-medium mr-2">Họ và tên:</span>
                      <span className="uppercase font-bold">{formData.fullName}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Ngày sinh:</span>
                      <span>{formData.dob ? formData.dob.split('-').reverse().join('/') : ''}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Giới tính:</span>
                      <span>{formData.gender}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Nơi sinh (Tỉnh):</span>
                      <span>{formData.pobProvince || 'Chưa nhập'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Số CCCD/Định danh:</span>
                      <span>{formData.idCard}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Số điện thoại liên hệ:</span>
                      <span>{formData.parentPhone || formData.fatherPhone || formData.motherPhone || 'Chưa cung cấp'}</span>
                    </div>
                    <div className="col-span-2 sm:col-span-1">
                      <span className="font-medium mr-2">Trường TH đã học:</span>
                      <span>{formData.school}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Lớp:</span>
                      <span>{formData.studentClass || 'Chưa nhập'}</span>
                    </div>
                    <div className="col-span-2">
                      <span className="font-medium mr-2">Địa chỉ thường trú:</span>
                      <span>{formData.address}</span>
                    </div>
                  </div>
                </div>

                <div className="mb-8 text-black">
                  <h2 className="text-lg font-bold mb-4 border-b border-black pb-2">II. Thông tin cha, mẹ học sinh</h2>
                  <div className="grid grid-cols-2 gap-y-4 gap-x-8 text-base">
                    <div>
                      <span className="font-medium mr-2">Họ và tên cha:</span>
                      <span className="uppercase font-semibold">{formData.fatherName || 'Không khai báo'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Họ và tên mẹ:</span>
                      <span className="uppercase font-semibold">{formData.motherName || 'Không khai báo'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Nghề nghiệp cha:</span>
                      <span>{formData.fatherJob || 'Không khai báo'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Nghề nghiệp mẹ:</span>
                      <span>{formData.motherJob || 'Không khai báo'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Số điện thoại cha:</span>
                      <span>{formData.fatherPhone || 'Không khai báo'}</span>
                    </div>
                    <div>
                      <span className="font-medium mr-2">Số điện thoại mẹ:</span>
                      <span>{formData.motherPhone || 'Không khai báo'}</span>
                    </div>
                  </div>
                </div>

                <div className="mb-8 text-black">
                  <h2 className="text-lg font-bold mb-4 border-b border-black pb-2">III. Kết quả học tập</h2>
                  <table className="w-full text-left border-collapse border border-black text-black">
                    <thead>
                      <tr>
                        <th className="border border-black py-2 px-4 text-center font-bold bg-gray-50 print:bg-transparent">Năm học</th>
                        <th className="border border-black py-2 px-4 font-bold bg-gray-50 print:bg-transparent">Mức đạt được (Kết quả giáo dục)</th>
                        <th className="border border-black py-2 px-4 text-center font-bold bg-gray-50 print:bg-transparent">Điểm quy đổi</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[1, 2, 3, 4, 5].map((grade) => (
                        <tr key={grade}>
                          <td className="border border-black py-3 px-4 text-center font-medium">Lớp {grade}</td>
                          <td className="border border-black py-3 px-4">{getLabel(formData.grades[grade as 1 | 2 | 3 | 4 | 5])}</td>
                          <td className="border border-black py-3 px-4 text-center font-bold">{getScore(formData.grades[grade as 1 | 2 | 3 | 4 | 5]).toFixed(1)}</td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td colSpan={2} className="border border-black py-4 px-4 text-right font-bold uppercase">Tổng điểm xét tuyển:</td>
                        <td className="border border-black py-4 px-4 text-center font-bold text-xl">{totalScore.toFixed(1)}</td>
                      </tr>
                    </tfoot>
                  </table>
                  
                  <div className="mt-4 text-sm text-gray-700 italic">
                    Ghi chú: Hoàn thành xuất sắc (10.0), Hoàn thành tốt (9.0), Hoàn thành (8.0).
                  </div>
                </div>

                <div className="flex justify-between mt-16 pt-8 px-8 text-black">
                  <div className="text-center">
                    <p className="font-bold mb-16">Phụ huynh / Người giám hộ</p>
                    <p className="italic text-gray-600">(Ký và ghi rõ họ tên)</p>
                  </div>
                  <div className="text-center">
                    <p className="font-bold mb-16">Người nhận hồ sơ</p>
                    <p className="italic text-gray-600">(Ký và ghi rõ họ tên)</p>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

        </AnimatePresence>
        </main>
      )}

      {/* Deletion Confirmation Modal */}
      <AnimatePresence>
        {deletingId && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full border border-gray-100"
            >
              <div className="flex items-center gap-3 text-red-600 mb-4">
                <div className="bg-red-50 p-2 rounded-full">
                  <AlertTriangle className="w-6 h-6" />
                </div>
                <h3 className="text-lg font-bold">Xác nhận xóa hồ sơ</h3>
              </div>
              
              <p className="text-gray-600 text-sm mb-6">
                Bạn có chắc chắn muốn xóa trực tiếp hồ sơ của học sinh <strong className="text-gray-900 uppercase font-bold">{deletingName}</strong> khỏi hệ thống? 
                <span className="block mt-2 text-red-600 font-semibold">⚠️ Cảnh báo: Hành động này sẽ xóa vĩnh viễn dữ liệu và không thể khôi phục lại!</span>
              </p>
              
              <div className="flex gap-3 justify-end">
                <button
                  type="button"
                  onClick={() => {
                    setDeletingId(null);
                    setDeletingName('');
                  }}
                  disabled={isDeleting}
                  className="px-4 py-2 border border-gray-300 rounded-lg text-gray-700 hover:bg-gray-50 transition font-medium text-sm disabled:opacity-50 cursor-pointer bg-white"
                >
                  Hủy bỏ
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    if (!deletingId) return;
                    setIsDeleting(true);
                    try {
                      await deleteDoc(doc(db, 'submissions', deletingId));
                      await fetchSubmissions();
                      setDeletingId(null);
                      setDeletingName('');
                    } catch (err) {
                      console.error("Lỗi khi xóa hồ sơ:", err);
                    } finally {
                      setIsDeleting(false);
                    }
                  }}
                  disabled={isDeleting}
                  className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition font-bold text-sm disabled:opacity-50 cursor-pointer flex items-center gap-1.5 border-none"
                >
                  {isDeleting ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" /> Đang xóa...
                    </>
                  ) : (
                    <>
                      <Trash2 className="w-4 h-4" /> Xác nhận xóa
                    </>
                  )}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}


