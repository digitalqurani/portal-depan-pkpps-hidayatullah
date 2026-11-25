/**
 * APLIKASI ABSENSI KBM PKPPS HIDAYATULLAH KARIMUN
 * Versi 3.3 + Portal Terpadu V1.5 - Perbaikan tampilan Manajemen Pengguna.
 *
 * CARA PEMASANGAN:
 * 1. Buat salinan/cadangan Google Spreadsheet database sebelum mengganti kode.
 * 2. Ganti isi Code.gs dengan file ini.
 * 3. Ganti isi Index.html dengan file Index.html versi 2.0.
 * 4. Pastikan nama file HTML di Apps Script adalah "Index".
 * 5. Jalankan setupAplikasi() satu kali dari editor Apps Script.
 *    Fungsi ini mempertahankan data lama dan hanya menambahkan sheet PENGESAHAN.
 * 6. Deploy > Manage deployments > Edit > New version > Deploy.
 *
 * CATATAN:
 * - Rekap "Belum absen" dihitung dari jadwal mingguan aktif pada setiap tanggal
 *   kalender bulan tersebut dikurangi absensi yang tercatat.
 * - Hari libur, kegiatan khusus, dan perubahan jadwal harus disesuaikan pada
 *   sheet JADWAL atau ditangani melalui kebijakan lembaga sebelum pengesahan.
 * - Ekspor membuat file di Google Drive pemilik aplikasi. File tidak dibagikan
 *   otomatis kepada publik.
 */

const APP = {
  name: 'Absensi KBM PKPPS Hidayatullah Karimun',
  year: '2026/2027',
  timezone: 'Asia/Jakarta',
  autoAbsenceGraceMinutes: 10,
  sheets: {
    guru: 'GURU',
    jadwal: 'JADWAL',
    absensi: 'ABSENSI',
    config: 'PENGATURAN',
    pengesahan: 'PENGESAHAN'
  }
};

function doGet(e) {
  const requested = e && e.parameter ? String(e.parameter.page || '') : '';
  if (requested === 'generator') {
    const template = HtmlService.createTemplateFromFile('GeneratorSoal');
    template.INITIAL_TOKEN = String(e && e.parameter ? e.parameter.token || '' : '');
    return template.evaluate()
      .setTitle('Generator Soal AI · PKPPS Hidayatullah Karimun')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const page = requested === 'portal' ? 'Portal' : (requested === 'ppdb' ? 'PPDBPublic' : 'Index');
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Portal' ? 'Portal Terpadu PKPPS Hidayatullah Karimun' : page === 'PPDBPublic' ? 'Pendaftaran Santri Baru 2027/2028' : APP.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** URL deployment Web App aktif, digunakan Portal untuk membuka modul terpisah. */
function getPortalWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('URL Web App belum tersedia. Deploy aplikasi sebagai Web App terlebih dahulu.');
  return url;
}

/** Jalankan satu kali. Data pada sheet yang sudah ada tidak dihapus. */
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet terlebih dahulu.');

  const guru = getOrCreateSheet_(ss, APP.sheets.guru,
    ['Kode Guru','Nama Guru','Username','PIN Hash','Peran','Aktif']);
  const jadwal = getOrCreateSheet_(ss, APP.sheets.jadwal,
    ['ID Jadwal','Hari','Jam Ke','Mulai','Selesai','Kelas','Mata Pelajaran','Kode Guru','Nama Guru','Aktif']);
  const absensi = getOrCreateSheet_(ss, APP.sheets.absensi,
    ['ID Jadwal','Tanggal','Hari','Waktu Presensi','Kode Guru','Nama Guru','Kelas','Jam Ke','Mulai','Selesai','Mata Pelajaran','Status','Catatan']);
  const config = getOrCreateSheet_(ss, APP.sheets.config, ['Kunci','Nilai']);
  const approval = getOrCreateSheet_(ss, APP.sheets.pengesahan,
    ['Bulan','Tahun','Status','Catatan','Disahkan Oleh','Waktu Pengesahan']);

  if (guru.getLastRow() < 2) {
    const teachers = [
      ['1','Ust. Saefuddin, M.Pd'],['2','Ust. Zainul Arifin, S.H.I'],
      ['3','Ust. Abd. Wahab S, S.Sos.I., M.Pd'],['4','Ust. Sahrin Yusuf, S.E., M.Pd'],
      ['5','Ust. Firman Firdaus, S.Pd.I'],['6','Ust. Ahmad Fauzan, S.Pd'],
      ['7','Ust. Mustofa, S.Pd'],['8','Ust. Joko Hadi Purnomo, S.Pd'],
      ['9','Ust. Anuar Ramadhan, S.Pd'],['10','Ust. Urip Al Islam, S.Pd'],
      ['11','Ustzh. Nurul Hidayah, M.Pd'],['12','Ustzh. Ridea Natasya, S.Pd'],
      ['13','Ustzh. Wardah S, S.Pd'],['14','Ustzh. Ulfa Maria'],
      ['15','Ustzh. Juliantika, S.Pd'],['16','Ustzh. Syahira Purnama Silaban'],
      ['17','Ustzh. Nur Syahida'],['18','Ustzh. Sabila Nisa Ayu Azhara'],
      ['19','Ustdzh. Suhaila, S.Pd'],['20','Ustzh. Sulita Asmawati, S.Pd'],
      ['21','Ust. M. Arsyil Azim']
    ];
    const rows = teachers.map(([id,n]) => [id,n,'kg'+id,hashPin_('123456'),id==='1'?'ADMIN':'GURU','YA']);
    guru.getRange(2,1,rows.length,6).setValues(rows);
  }

  if (config.getLastRow() < 2) {
    config.getRange(2,1,4,2).setValues([
      ['NAMA_LEMBAGA','PKPPS Hidayatullah Karimun'],
      ['TAHUN_AJARAN',APP.year],
      ['PIN_AWAL','123456'],
      ['CATATAN_JADWAL','Isi jadwal berdasarkan PDF yang telah diverifikasi oleh Waka Kurikulum.']
    ]);
  }

  [guru,jadwal,absensi,config,approval].forEach(sh => {
    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) {
      sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold')
        .setBackground('#173b63').setFontColor('#ffffff');
      sh.autoResizeColumns(1, sh.getLastColumn());
    }
  });
  installAutoAbsenceTrigger_();

  SpreadsheetApp.flush();
  return 'Pemeriksaan selesai. Data lama dipertahankan; pemantauan absensi otomatis setiap 5 menit telah dipasang.';
}

function login(username, pin) {
  username = String(username || '').trim().toLowerCase();
  pin = String(pin || '').trim();
  if (!username || !pin) throw new Error('Username dan PIN wajib diisi.');
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sheets.guru);
  const values = sh.getDataRange().getValues();
  const row = values.slice(1).find(r =>
    String(r[2]).toLowerCase() === username &&
    r[3] === hashPin_(pin) &&
    String(r[5]).toUpperCase() === 'YA'
  );
  if (!row) throw new Error('Username/PIN salah atau akun tidak aktif.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sess_'+token, JSON.stringify({
    kode:String(row[0]), nama:String(row[1]), username:String(row[2]), peran:String(row[4])
  }), 21600);
  return {token, user:{kode:String(row[0]), nama:String(row[1]), peran:String(row[4])}};
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_'+token);
  return true;
}

function getDashboard(token) {
  const user = getSession_(token);
  const today = new Date();
  const day = getDayName_(today);
  const dateKey = Utilities.formatDate(today, APP.timezone, 'yyyy-MM-dd');
  const allSchedules = readObjects_(APP.sheets.jadwal);
  const teacherSchedules = allSchedules.filter(s => normalizeCode_(s['Kode Guru']) === normalizeCode_(user.kode));
  const daySchedules = teacherSchedules.filter(s => normalizeDay_(s['Hari']) === normalizeDay_(day));
  const schedules = daySchedules.filter(s => isActive_(s['Aktif']))
    .sort((a,b)=>timeToMinutes_(a['Mulai']) - timeToMinutes_(b['Mulai']));
  const attendance = readObjects_(APP.sheets.absensi).filter(a =>
    normalizeDate_(a['Tanggal']) === dateKey && normalizeCode_(a['Kode Guru']) === normalizeCode_(user.kode)
  );
  const list = schedules.map(s => {
    const record = attendance.find(a => String(a['ID Jadwal']) === String(s['ID Jadwal']));
    const note = record ? String(record['Catatan'] || '') : '';
    return {
      ...s,
      sudahAbsen: !!record,
      status: record ? record['Status'] : '',
      catatan: note,
      otomatis: note.indexOf('OTOMATIS:') === 0
    };
  });
  return {
    user, day, date: Utilities.formatDate(today, APP.timezone, 'dd MMMM yyyy'),
    schedules:list, total:list.length, done:list.filter(x=>x.sudahAbsen).length,
    scheduleDiagnostic: {
      teacherSchedules: teacherSchedules.length,
      matchingDay: daySchedules.length,
      active: schedules.length
    }
  };
}

/**
 * Jadwal harian seluruh guru dapat dilihat oleh semua akun yang sudah login.
 * Hanya menampilkan jadwal aktif pada tanggal yang dipilih.
 */
function getAllDailySchedule(token, selectedDate) {
  getSession_(token); // autentikasi: semua guru boleh melihat, tidak boleh mengubah
  const raw = String(selectedDate || '').trim();
  const dateKey = raw || Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Tanggal jadwal tidak valid.');
  const parts = dateKey.split('-').map(Number);
  const dateObj = new Date(Date.UTC(parts[0], parts[1]-1, parts[2], 12));
  if (dateObj.getUTCFullYear() !== parts[0] || dateObj.getUTCMonth() !== parts[1]-1 || dateObj.getUTCDate() !== parts[2]) {
    throw new Error('Tanggal jadwal tidak valid.');
  }
  const dayNames = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const day = dayNames[dateObj.getUTCDay()];
  const teachers = readObjects_(APP.sheets.guru);
  const teacherByCode = new Map(teachers.map(t => [normalizeCode_(t['Kode Guru']), String(t['Nama Guru'] || '')]));
  const schedules = readObjects_(APP.sheets.jadwal)
    .filter(s => isActive_(s['Aktif']) && normalizeDay_(s['Hari']) === normalizeDay_(day))
    .map(s => ({
      id: String(s['ID Jadwal'] || ''),
      hari: day,
      jamKe: String(s['Jam Ke'] || ''),
      mulai: displayTime_(s['Mulai']),
      selesai: displayTime_(s['Selesai']),
      kelas: String(s['Kelas'] || ''),
      mapel: String(s['Mata Pelajaran'] || ''),
      kodeGuru: String(s['Kode Guru'] || ''),
      namaGuru: String(s['Nama Guru'] || teacherByCode.get(normalizeCode_(s['Kode Guru'])) || '')
    }))
    .sort((a,b) => timeToMinutes_(a.mulai) - timeToMinutes_(b.mulai) ||
      String(a.kelas).localeCompare(String(b.kelas), 'id') ||
      String(a.namaGuru).localeCompare(String(b.namaGuru), 'id'));
  return {date:dateKey, day, schedules, total:schedules.length};
}


/**
 * Monitoring keterisian kelas VII-XII untuk seluruh guru yang sudah login.
 * Terisi = absensi berstatus Hadir; Kosong = status selain Hadir atau
 * belum absen setelah 10 menit dari jam selesai; Menunggu = belum melewati batas.
 */
function getClassMonitoring(token) {
  getPortalSession_(token);
  const now = new Date();
  const dateKey = Utilities.formatDate(now, APP.timezone, 'yyyy-MM-dd');
  const todayName = getDayName_(now);
  const currentMinutes = timeToMinutes_(Utilities.formatDate(now, APP.timezone, 'HH:mm'));
  const grades = [
    {key:'VII', label:'Kelas VII'}, {key:'VIII', label:'Kelas VIII'},
    {key:'IX', label:'Kelas IX'}, {key:'X', label:'Kelas X'},
    {key:'XI', label:'Kelas XI'}, {key:'XII', label:'Kelas XII'}
  ];
  const result = grades.map(g => ({...g, total:0, terisi:0, kosong:0, menunggu:0}));
  const byGrade = new Map(result.map(g => [g.key, g]));
  const attendance = readObjects_(APP.sheets.absensi).filter(a =>
    normalizeDate_(a['Tanggal']) === dateKey
  );
  const attendanceById = new Map();
  attendance.forEach(a => {
    const key = String(a['ID Jadwal'] || '').trim();
    if (key) attendanceById.set(key, a);
  });
  const schedules = readObjects_(APP.sheets.jadwal).filter(s =>
    isActive_(s['Aktif']) && normalizeDay_(s['Hari']) === normalizeDay_(todayName)
  );

  schedules.forEach(s => {
    const rawClass = String(s['Kelas'] || '').toUpperCase().trim();
    const normalized = rawClass.replace(/[^A-Z0-9]/g, ' ');
    let gradeKey = '';
    // Support labels such as 7A, Kelas VII, VII A, Kelas 7A.
    const roman = normalized.match(/(^| )(XII|XI|IX|VIII|VII|VI|IV|V|III|II|I)(?= |$)/);
    const numeric = normalized.match(/(^| )(12|11|10|9|8|7)(?= |[A-Z]|$)/);
    if (roman) gradeKey = roman[2];
    else if (numeric) gradeKey = ({'7':'VII','8':'VIII','9':'IX','10':'X','11':'XI','12':'XII'})[numeric[2]] || '';
    const bucket = byGrade.get(gradeKey);
    if (!bucket) return;
    bucket.total++;
    const record = attendanceById.get(String(s['ID Jadwal'] || '').trim());
    if (record) {
      if (String(record['Status'] || '').trim() === 'Hadir') bucket.terisi++;
      else bucket.kosong++;
      return;
    }
    const endMinutes = timeToMinutes_(s['Selesai']);
    if (endMinutes !== null && currentMinutes >= endMinutes + Number(APP.autoAbsenceGraceMinutes || 10)) {
      bucket.kosong++;
    } else {
      bucket.menunggu++;
    }
  });
  return {
    date: dateKey,
    day: todayName,
    updatedAt: Utilities.formatDate(now, APP.timezone, 'HH:mm:ss'),
    grades: result,
    total: result.reduce((n,g)=>n+g.total,0),
    terisi: result.reduce((n,g)=>n+g.terisi,0),
    kosong: result.reduce((n,g)=>n+g.kosong,0),
    menunggu: result.reduce((n,g)=>n+g.menunggu,0)
  };
}

function submitAttendance(token, payload) {
  const user = getSession_(token);
  if (!payload || !payload.idJadwal || !payload.status) throw new Error('Data absensi tidak lengkap.');
  const valid = ['Hadir','Izin','Sakit','Tidak Hadir'];
  if (!valid.includes(payload.status)) throw new Error('Status absensi tidak valid.');
  const schedules = readObjects_(APP.sheets.jadwal);
  const s = schedules.find(x => String(x['ID Jadwal']) === String(payload.idJadwal));
  if (!s || normalizeCode_(s['Kode Guru']) !== normalizeCode_(user.kode) || !isActive_(s['Aktif'])) {
    throw new Error('Jadwal tidak ditemukan untuk akun ini.');
  }
  const now = new Date();
  const dateKey = Utilities.formatDate(now, APP.timezone, 'yyyy-MM-dd');
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sheets.absensi);
  const rows = sh.getDataRange().getValues();
  const exists = rows.slice(1).some(r =>
    normalizeDate_(r[1])===dateKey && String(r[4])===user.kode && String(r[0])===String(payload.idJadwal)
  );
  if (exists) throw new Error('Absensi untuk sesi ini sudah tercatat.');
  const day = getDayName_(now);
  sh.appendRow([
    String(s['ID Jadwal']), dateKey, day, Utilities.formatDate(now,APP.timezone,'HH:mm:ss'),
    user.kode,user.nama,String(s['Kelas']),String(s['Jam Ke']),String(s['Mulai']),
    String(s['Selesai']),String(s['Mata Pelajaran']),payload.status,String(payload.catatan||'')
  ]);
  return {success:true, message:'Absensi berhasil disimpan.'};
}

function getAdminData(token) {
  requireAdmin_(token);
  const teachers = readObjects_(APP.sheets.guru);
  const schedules = readObjects_(APP.sheets.jadwal);
  const teacherByCode = new Map(teachers.map(t => [normalizeCode_(t['Kode Guru']), t]));
  const scheduleById = new Map(schedules.map(s => [String(s['ID Jadwal'] || '').trim(), s]));
  const allAttendance = readObjects_(APP.sheets.absensi);
  const today = new Date();
  const todayKey = Utilities.formatDate(today, APP.timezone, 'yyyy-MM-dd');
  const todayName = getDayName_(today);
  const activeSchedules = schedules.filter(s => isActive_(s['Aktif']));
  const todaySchedules = activeSchedules.filter(s => normalizeDay_(s['Hari']) === normalizeDay_(todayName));
  const todayAttendanceCount = allAttendance.filter(a => normalizeDate_(a['Tanggal']) === todayKey).length;
  const attendance = allAttendance.slice(-100).reverse().map(a => {
    const schedule = scheduleById.get(String(a['ID Jadwal'] || '').trim()) || {};
    const teacher = teacherByCode.get(normalizeCode_(a['Kode Guru'])) || {};
    const result = Object.assign({}, a);
    // Lengkapi kolom yang kosong dari master jadwal/guru. Tanggal tidak ditebak.
    if (!result['Nama Guru']) result['Nama Guru'] = schedule['Nama Guru'] || teacher['Nama Guru'] || '';
    if (!result['Kelas']) result['Kelas'] = schedule['Kelas'] || '';
    if (!result['Jam Ke']) result['Jam Ke'] = schedule['Jam Ke'] || '';
    if (!result['Mata Pelajaran']) result['Mata Pelajaran'] = schedule['Mata Pelajaran'] || '';
    if (!result['Mulai']) result['Mulai'] = schedule['Mulai'] || '';
    if (!result['Selesai']) result['Selesai'] = schedule['Selesai'] || '';
    if (!result['Tanggal']) result['Tanggal'] = pickValue_(a, ['Tgl','Tanggal Absen','Tanggal Absensi','Date']) || '';
    if (!result['Nama Guru']) result['Nama Guru'] = pickValue_(a, ['Guru','Nama']) || '';
    if (!result['Kelas']) result['Kelas'] = pickValue_(a, ['Rombel','Kelas']) || '';
    return result;
  });
  return {
    guru: teachers.map(r=>({
      kode:r['Kode Guru'], nama:r['Nama Guru'], username:r['Username'], peran:r['Peran'], aktif:r['Aktif']
    })),
    jadwal: schedules,
    absensi: attendance,
    ringkasan: {
      guruAktif: teachers.filter(t => isActive_(t['Aktif'])).length,
      jadwalAktif: activeSchedules.length,
      absensiHariIni: todayAttendanceCount,
      jadwalHariIni: todaySchedules.length,
      tanggal: todayKey,
      hari: todayName
    }
  };
}

function saveSchedule(token, item) {
  requireAdmin_(token);
  const required = ['hari','jamKe','mulai','selesai','kelas','mapel','kodeGuru'];
  required.forEach(k=>{ if(!String(item[k]||'').trim()) throw new Error('Kolom '+k+' wajib diisi.'); });
  const teachers = readObjects_(APP.sheets.guru);
  const teacher = teachers.find(t=>String(t['Kode Guru'])===String(item.kodeGuru));
  if (!teacher) throw new Error('Kode guru tidak ditemukan.');
  const id = Utilities.getUuid();
  SpreadsheetApp.getActive().getSheetByName(APP.sheets.jadwal).appendRow([
    id,item.hari,item.jamKe,item.mulai,item.selesai,item.kelas,item.mapel,
    item.kodeGuru,teacher['Nama Guru'],'YA'
  ]);
  return {success:true};
}

function changePin(token, oldPin, newPin) {
  const user = getSession_(token);
  if (!/^\d{6,12}$/.test(String(newPin))) throw new Error('PIN baru harus 6–12 digit angka.');
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sheets.guru);
  const vals = sh.getDataRange().getValues();
  const idx = vals.findIndex((r,i)=>i>0 && String(r[0])===user.kode);
  if (idx<1 || vals[idx][3]!==hashPin_(oldPin)) throw new Error('PIN lama tidak sesuai.');
  sh.getRange(idx+1,4).setValue(hashPin_(newPin));
  return {success:true};
}


/** Rekap mingguan berdasarkan pekan Senin-Minggu dari tanggal pilihan. */
function getWeeklyReport(token, selectedDate) {
  requireAdmin_(token);
  const raw = String(selectedDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) throw new Error('Tanggal rekap mingguan tidak valid.');
  const parts = raw.split('-').map(Number);
  const chosen = new Date(Date.UTC(parts[0], parts[1]-1, parts[2], 12));
  if (chosen.getUTCFullYear() !== parts[0] || chosen.getUTCMonth() !== parts[1]-1 || chosen.getUTCDate() !== parts[2]) {
    throw new Error('Tanggal rekap mingguan tidak valid.');
  }
  const mondayOffset = (chosen.getUTCDay() + 6) % 7;
  const start = new Date(chosen.getTime() - mondayOffset * 86400000);
  const end = new Date(start.getTime() + 6 * 86400000);
  const toKey = d => Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
  const startKey = toKey(start), endKey = toKey(end);
  const dayNames = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const teachers = readObjects_(APP.sheets.guru).filter(t => isActive_(t['Aktif']));
  const schedules = readObjects_(APP.sheets.jadwal).filter(s => isActive_(s['Aktif']));
  const attendanceRows = readObjects_(APP.sheets.absensi).filter(a => {
    const d = normalizeDate_(a['Tanggal']);
    return d && d >= startKey && d <= endKey;
  });
  const stats = {};
  teachers.forEach(t => {
    const code = normalizeCode_(t['Kode Guru']);
    stats[code] = {kode:code, nama:String(t['Nama Guru'] || ''), expected:0, recorded:0,
      hadir:0, izin:0, sakit:0, tidakHadir:0, missing:0, percent:0};
  });
  const expectedKeys = new Set();
  const daily = {};
  for (let offset=0; offset<7; offset++) {
    const d = new Date(start.getTime() + offset*86400000);
    const keyDate = toKey(d);
    const dayName = dayNames[d.getUTCDay()];
    daily[keyDate] = {tanggal:keyDate, hari:dayName, jadwal:0, tercatat:0, hadir:0, izin:0, sakit:0, tidakHadir:0, belum:0};
    schedules.forEach(s => {
      if (normalizeDay_(s['Hari']) !== normalizeDay_(dayName)) return;
      const code = normalizeCode_(s['Kode Guru']);
      const key = keyDate + '|' + String(s['ID Jadwal'] || '').trim();
      if (expectedKeys.has(key)) return;
      expectedKeys.add(key);
      daily[keyDate].jadwal++;
      if (stats[code]) stats[code].expected++;
    });
  }
  const attendanceByKey = new Map();
  attendanceRows.forEach(a => {
    const key = normalizeDate_(a['Tanggal']) + '|' + String(a['ID Jadwal'] || '').trim();
    attendanceByKey.set(key, a);
  });
  expectedKeys.forEach(key => {
    const [dateKey, scheduleId] = key.split('|');
    const schedule = schedules.find(s => String(s['ID Jadwal'] || '').trim() === scheduleId);
    if (!schedule) return;
    const a = attendanceByKey.get(key);
    const day = daily[dateKey];
    if (!a) {
      if (stats[normalizeCode_(schedule['Kode Guru'])]) stats[normalizeCode_(schedule['Kode Guru'])].missing++;
      day.belum++;
      return;
    }
    day.tercatat++;
    const status = String(a['Status'] || '').trim();
    if (status === 'Hadir') { day.hadir++; if (stats[normalizeCode_(schedule['Kode Guru'])]) { stats[normalizeCode_(schedule['Kode Guru'])].recorded++; stats[normalizeCode_(schedule['Kode Guru'])].hadir++; } }
    else if (status === 'Izin') { day.izin++; if (stats[normalizeCode_(schedule['Kode Guru'])]) { stats[normalizeCode_(schedule['Kode Guru'])].recorded++; stats[normalizeCode_(schedule['Kode Guru'])].izin++; } }
    else if (status === 'Sakit') { day.sakit++; if (stats[normalizeCode_(schedule['Kode Guru'])]) { stats[normalizeCode_(schedule['Kode Guru'])].recorded++; stats[normalizeCode_(schedule['Kode Guru'])].sakit++; } }
    else if (status === 'Tidak Hadir') { day.tidakHadir++; if (stats[normalizeCode_(schedule['Kode Guru'])]) { stats[normalizeCode_(schedule['Kode Guru'])].recorded++; stats[normalizeCode_(schedule['Kode Guru'])].tidakHadir++; } }
    else if (stats[normalizeCode_(schedule['Kode Guru'])]) stats[normalizeCode_(schedule['Kode Guru'])].recorded++;
  });
  const list = Object.values(stats).map(s => {
    s.percent = s.expected ? Math.round((s.hadir/s.expected)*1000)/10 : 0;
    return s;
  }).sort((a,b) => a.nama.localeCompare(b.nama, 'id'));
  const totals = list.reduce((o,s) => {
    ['expected','recorded','hadir','izin','sakit','tidakHadir','missing'].forEach(k => o[k] += s[k]);
    return o;
  }, {expected:0,recorded:0,hadir:0,izin:0,sakit:0,tidakHadir:0,missing:0,percent:0});
  totals.percent = totals.expected ? Math.round((totals.hadir/totals.expected)*1000)/10 : 0;
  return {start:startKey, end:endKey, teachers:list, totals, daily:Object.values(daily)};
}

/** Menghasilkan rekap bulanan. Bulan menggunakan angka 1-12. */
function getMonthlyReport(token, month, year) {
  requireAdmin_(token);
  month = Number(month); year = Number(year);
  validateMonthYear_(month, year);
  const result = buildMonthlyReport_(month, year);
  const approval = getApproval_(month, year);
  return {...result, approval};
}

/** Persetujuan bulanan hanya dapat dilakukan oleh akun ADMIN. */
function saveMonthlyApproval(token, month, year, note) {
  const user = requireAdmin_(token);
  month = Number(month); year = Number(year);
  validateMonthYear_(month, year);
  const sh = SpreadsheetApp.getActive().getSheetByName(APP.sheets.pengesahan);
  const rows = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i=1; i<rows.length; i++) {
    if (Number(rows[i][0])===month && Number(rows[i][1])===year) { rowIndex=i+1; break; }
  }
  const values = [month,year,'DISETUJUI',String(note||''),user.nama,
    Utilities.formatDate(new Date(),APP.timezone,'yyyy-MM-dd HH:mm:ss')];
  if (rowIndex > 0) sh.getRange(rowIndex,1,1,values.length).setValues([values]);
  else sh.appendRow(values);
  return {success:true, approval:getApproval_(month,year)};
}

/** Membuat file PDF atau XLSX dan menyimpannya di Drive pemilik aplikasi. */
function exportMonthlyReport(token, month, year, format) {
  requireAdmin_(token);
  month = Number(month); year = Number(year);
  validateMonthYear_(month, year);
  format = String(format||'').toLowerCase();
  if (!['pdf','xlsx'].includes(format)) throw new Error('Format ekspor harus PDF atau XLSX.');

  const report = buildMonthlyReport_(month, year);
  const approval = getApproval_(month, year);
  const monthName = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][month-1];
  const title = 'Rekap Absensi Guru '+monthName+' '+year;
  const temp = SpreadsheetApp.create('TEMP - '+title);
  const sh = temp.getSheets()[0];
  sh.setName('Rekap Bulanan');
  const rows = [
    ['PKPPS HIDAYATULLAH KARIMUN'],
    [title],
    ['Tahun Ajaran', APP.year],
    ['Status Pengesahan', approval ? approval.status : 'BELUM DISAHKAN'],
    ['Disahkan Oleh', approval ? approval.by : '-'],
    ['Waktu Pengesahan', approval ? approval.at : '-'],
    [],
    ['Kode Guru','Nama Guru','Jadwal Terencana','Absensi Tercatat','Hadir','Izin','Sakit','Tidak Hadir','Belum Absen','Kehadiran (%)']
  ];
  report.teachers.forEach(t => rows.push([
    t.kode,t.nama,t.expected,t.recorded,t.hadir,t.izin,t.sakit,t.tidakHadir,t.missing,t.percent
  ]));
  rows.push([]);
  rows.push(['TOTAL','',report.totals.expected,report.totals.recorded,report.totals.hadir,
    report.totals.izin,report.totals.sakit,report.totals.tidakHadir,report.totals.missing,report.totals.percent]);
  rows.push([]);
  rows.push(['Catatan pengesahan', approval ? approval.note : '-']);
  sh.getRange(1,1,rows.length,10).setValues(rows.map(r=>{
    const a=r.slice(); while(a.length<10) a.push(''); return a;
  }));
  sh.getRange(1,1,1,10).merge().setFontWeight('bold').setFontSize(14);
  sh.getRange(2,1,1,10).merge().setFontWeight('bold').setFontSize(12);
  sh.getRange(8,1,1,10).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');
  sh.getRange(rows.length-2,1,1,10).setFontWeight('bold').setBackground('#eaf1f8');
  sh.setFrozenRows(8);
  sh.autoResizeColumns(1,10);
  sh.getRange(9,10,Math.max(report.teachers.length,1),1).setNumberFormat('0.0"%"');
  SpreadsheetApp.flush();

  const mime = format==='pdf' ? 'application/pdf' :
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const ext = format==='pdf' ? 'pdf' : 'xlsx';
  const url = 'https://www.googleapis.com/drive/v3/files/'+temp.getId()+'/export?mimeType='+encodeURIComponent(mime);
  const response = UrlFetchApp.fetch(url,{
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
    muteHttpExceptions:true
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
    throw new Error('Ekspor gagal. Jalankan kembali dan setujui izin Google Drive jika diminta.');
  }
  const blob = response.getBlob().setName(title+'.'+ext);
  const file = DriveApp.createFile(blob);
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  return {success:true, name:file.getName(), url:file.getUrl()};
}

/* -------------------- FUNGSI REKAP INTERNAL -------------------- */

function buildMonthlyReport_(month, year) {
  const teachers = readObjects_(APP.sheets.guru).filter(t=>String(t['Aktif']).toUpperCase()==='YA');
  const schedules = readObjects_(APP.sheets.jadwal).filter(s=>isActive_(s['Aktif']));
  const allAttendance = readObjects_(APP.sheets.absensi).filter(a=>{
    const d = normalizeDate_(a['Tanggal']);
    return d && Number(d.slice(5,7))===month && Number(d.slice(0,4))===year;
  });
  const dayNames = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const expectedByTeacher = {};
  teachers.forEach(t=>expectedByTeacher[String(t['Kode Guru'])]=0);

  const expectedKeys = new Set();
  const daysInMonth = new Date(year,month,0).getDate();
  for (let day=1; day<=daysInMonth; day++) {
    const date = new Date(year,month-1,day);
    const dayName = dayNames[date.getDay()];
    schedules.forEach(s=>{
      if (String(s['Hari']).toLowerCase()!==dayName.toLowerCase()) return;
      const code = String(s['Kode Guru']);
      const key = Utilities.formatDate(date,APP.timezone,'yyyy-MM-dd')+'|'+String(s['ID Jadwal']);
      if (!expectedKeys.has(key)) {
        expectedKeys.add(key);
        expectedByTeacher[code] = (expectedByTeacher[code]||0)+1;
      }
    });
  }

  const attendanceByKey = new Map();
  allAttendance.forEach(a=>{
    const key = normalizeDate_(a['Tanggal'])+'|'+String(a['ID Jadwal']);
    attendanceByKey.set(key,a);
  });

  const stats = {};
  teachers.forEach(t=>{
    const code=String(t['Kode Guru']);
    stats[code]={kode:code,nama:String(t['Nama Guru']),expected:expectedByTeacher[code]||0,
      recorded:0,hadir:0,izin:0,sakit:0,tidakHadir:0,missing:0,percent:0};
  });

  expectedKeys.forEach(key=>{
    const parts=key.split('|');
    const scheduleId=parts[1];
    const schedule=schedules.find(s=>String(s['ID Jadwal'])===scheduleId);
    if (!schedule) return;
    const code=String(schedule['Kode Guru']);
    const st=stats[code];
    if (!st) return;
    const a=attendanceByKey.get(key);
    if (!a) { st.missing++; return; }
    st.recorded++;
    const status=String(a['Status']);
    if (status==='Hadir') st.hadir++;
    else if (status==='Izin') st.izin++;
    else if (status==='Sakit') st.sakit++;
    else if (status==='Tidak Hadir') st.tidakHadir++;
  });

  const list=Object.keys(stats).map(k=>{
    const s=stats[k];
    s.percent=s.expected ? Math.round((s.hadir/s.expected)*1000)/10 : 0;
    return s;
  }).sort((a,b)=>a.nama.localeCompare(b.nama,'id'));
  const totals=list.reduce((o,s)=>{
    ['expected','recorded','hadir','izin','sakit','tidakHadir','missing'].forEach(k=>o[k]+=s[k]);
    return o;
  },{expected:0,recorded:0,hadir:0,izin:0,sakit:0,tidakHadir:0,missing:0,percent:0});
  totals.percent=totals.expected ? Math.round((totals.hadir/totals.expected)*1000)/10 : 0;
  return {month,year,monthName:['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][month-1],teachers:list,totals};
}

function getApproval_(month,year) {
  const rows=readObjects_(APP.sheets.pengesahan);
  const r=rows.find(x=>Number(x['Bulan'])===month && Number(x['Tahun'])===year);
  return r ? {status:String(r['Status']),note:String(r['Catatan']||''),by:String(r['Disahkan Oleh']||''),at:String(r['Waktu Pengesahan']||'')} : null;
}

function requireAdmin_(token) {
  const user=getSession_(token);
  if (user.peran!=='ADMIN') throw new Error('Akses khusus admin.');
  return user;
}
function validateMonthYear_(month,year) {
  if (!Number.isInteger(month)||month<1||month>12) throw new Error('Bulan tidak valid.');
  if (!Number.isInteger(year)||year<2020||year>2100) throw new Error('Tahun tidak valid.');
}
function isActive_(value) {
  const v=String(value===undefined?'YA':value).toUpperCase();
  return v!=='TIDAK' && v!=='FALSE' && v!=='0' && v!=='NONAKTIF';
}
function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value,APP.timezone,'yyyy-MM-dd');
  const s=String(value||'').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
    const p=s.split('/'); return p[2]+'-'+('0'+p[1]).slice(-2)+'-'+('0'+p[0]).slice(-2);
  }
  return s;
}
function getDayName_(date) {
  const n = Number(Utilities.formatDate(date, APP.timezone, 'u'));
  return ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'][n - 1];
}

function normalizeDay_(value) {
  const v = String(value === null || value === undefined ? '' : value)
    .trim().toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ');
  const aliases = {
    'senin':'senin','mon':'senin','monday':'senin',
    'selasa':'selasa','tue':'selasa','tuesday':'selasa',
    'rabu':'rabu','wed':'rabu','wednesday':'rabu',
    'kamis':'kamis','thu':'kamis','thursday':'kamis',
    'jumat':'jumat',"jum'at":'jumat','fri':'jumat','friday':'jumat',
    'sabtu':'sabtu','sat':'sabtu','saturday':'sabtu',
    'minggu':'minggu','ahad':'minggu','sun':'minggu','sunday':'minggu'
  };
  return aliases[v] || v;
}

function normalizeCode_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function pickValue_(obj, keys) {
  for (const key of keys) {
    const value = obj[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}

function getSession_(token) {
  const raw = CacheService.getScriptCache().get('sess_'+String(token||''));
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  return JSON.parse(raw);
}
function hashPin_(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin), Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function getOrCreateSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  return sh;
}
function readObjects_(sheetName) {
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2) return [];
  const values=sh.getDataRange().getValues();
  // Rapikan judul kolom: database lama memiliki spasi di awal/akhir header,
  // misalnya ' Tanggal', 'Nama Guru ', 'Kelas ', dan 'Username '.
  const headers=values[0].map(h=>String(h == null ? '' : h).trim());
  return values.slice(1).filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>{
    const o={};
    headers.forEach((h,i)=>{
      if (!h) return;
      o[h]=r[i] instanceof Date
        ? Utilities.formatDate(r[i],APP.timezone,'yyyy-MM-dd')
        : r[i];
    });
    return o;
  });
}


/* ---------------- ABSENSI OTOMATIS V2.1 ---------------- */

/**
 * Memasang satu trigger otomatis setiap 5 menit.
 * Trigger lama untuk fungsi yang sama dihapus agar tidak ganda.
 */
function installAutoAbsenceTrigger_() {
  const handler = 'autoMarkUnattendedClasses';

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();
}

/**
 * Menandai sesi yang belum diabsen sebagai Tidak Hadir setelah
 * 10 menit dari jam selesai pelajaran. Hanya jadwal aktif hari ini.
 */
function autoMarkUnattendedClasses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const now = new Date();
    const dateKey = Utilities.formatDate(now, APP.timezone, 'yyyy-MM-dd');
    const currentMinutes = timeToMinutes_(
      Utilities.formatDate(now, APP.timezone, 'HH:mm')
    );

    // Utilities format 'u': Senin=1 sampai Minggu=7.
    const dayNumber = Number(Utilities.formatDate(now, APP.timezone, 'u'));
    const days = ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];
    const todayName = days[dayNumber - 1];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scheduleSheet = ss.getSheetByName(APP.sheets.jadwal);
    const attendanceSheet = ss.getSheetByName(APP.sheets.absensi);

    if (!scheduleSheet || !attendanceSheet) {
      throw new Error('Sheet JADWAL atau ABSENSI tidak ditemukan.');
    }

    const schedules = getRawObjects_(scheduleSheet);
    const attendance = readObjects_(APP.sheets.absensi);
    const existing = new Set(
      attendance
        .filter(a => normalizeDate_(a['Tanggal']) === dateKey)
        .map(a => [
          dateKey,
          normalizeCode_(a['Kode Guru']),
          String(a['ID Jadwal']).trim()
        ].join('|'))
    );

    const grace = Number(APP.autoAbsenceGraceMinutes || 10);
    const rowsToInsert = [];

    schedules.forEach(s => {
      if (normalizeDay_(s['Hari']) !== normalizeDay_(todayName)) return;
      if (!isActive_(s['Aktif'])) return;

      const endMinutes = timeToMinutes_(s['Selesai']);
      if (endMinutes === null) return;
      if (currentMinutes < endMinutes + grace) return;

      const key = [
        dateKey,
        normalizeCode_(s['Kode Guru']),
        String(s['ID Jadwal']).trim()
      ].join('|');

      if (existing.has(key)) return;

      rowsToInsert.push([
        String(s['ID Jadwal']),
        dateKey,
        todayName,
        'OTOMATIS ' + Utilities.formatDate(now, APP.timezone, 'HH:mm:ss'),
        String(s['Kode Guru']),
        String(s['Nama Guru']),
        String(s['Kelas']),
        String(s['Jam Ke']),
        displayTime_(s['Mulai']),
        displayTime_(s['Selesai']),
        String(s['Mata Pelajaran']),
        'Tidak Hadir',
        'OTOMATIS: guru tidak melakukan absensi hingga ' + grace +
          ' menit setelah jam selesai; KELAS KOSONG.'
      ]);

      // Cegah sesi yang sama masuk dua kali dalam satu eksekusi.
      existing.add(key);
    });

    if (rowsToInsert.length) {
      attendanceSheet
        .getRange(attendanceSheet.getLastRow() + 1, 1, rowsToInsert.length, 13)
        .setValues(rowsToInsert);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

/** Membaca jadwal mentah agar nilai waktu tidak berubah menjadi tanggal. */
function getRawObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1)
    .filter(row => row.some(value => value !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => obj[header] = row[i]);
      return obj;
    });
}

/** Konversi nilai jam Sheets (Date, pecahan angka, atau HH:mm) ke menit. */
function timeToMinutes_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date) {
    const text = Utilities.formatDate(value, APP.timezone, 'HH:mm');
    return timeToMinutes_(text);
  }

  if (typeof value === 'number') {
    const fraction = ((value % 1) + 1) % 1;
    return Math.round(fraction * 1440);
  }

  const match = String(value).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!match) return null;

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > 23 || minute > 59) return null;

  return hour * 60 + minute;
}

/** Mengubah nilai jam menjadi teks HH:mm untuk disimpan di ABSENSI. */
function displayTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP.timezone, 'HH:mm');
  }
  if (typeof value === 'number') {
    const minutes = timeToMinutes_(value);
    if (minutes === null) return String(value);
    return ('0' + Math.floor(minutes / 60)).slice(-2) + ':' +
      ('0' + (minutes % 60)).slice(-2);
  }
  return String(value || '');
}

/**
 * Koreksi catatan absensi otomatis oleh admin.
 * dateKey harus yyyy-MM-dd dan scheduleId adalah ID Jadwal.
 */
function correctAutoAbsence(token, dateKey, scheduleId, newStatus, note) {
  requireAdmin_(token);

  dateKey = String(dateKey || '').trim();
  scheduleId = String(scheduleId || '').trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
    throw new Error('Tanggal harus berformat yyyy-MM-dd.');
  }
  if (!scheduleId) throw new Error('ID Jadwal wajib diisi.');

  const allowed = ['Hadir', 'Izin', 'Sakit', 'Tidak Hadir'];
  if (!allowed.includes(newStatus)) {
    throw new Error('Status koreksi tidak valid.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('Sistem sedang memproses. Coba lagi.');

  try {
    const sheet = SpreadsheetApp.getActive()
      .getSheetByName(APP.sheets.absensi);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const rowDate = normalizeDate_(values[i][1]);
      const rowSchedule = String(values[i][0]);
      const oldNote = String(values[i][12] || '');

      if (
        rowDate === dateKey &&
        rowSchedule === scheduleId &&
        oldNote.indexOf('OTOMATIS:') === 0
      ) {
        sheet.getRange(i + 1, 12).setValue(newStatus);
        sheet.getRange(i + 1, 13).setValue(
          'DIKOREKSI ADMIN: ' + String(note || '-') +
          ' | Catatan awal: ' + oldNote
        );

        return {
          success: true,
          message: 'Catatan absensi otomatis berhasil dikoreksi.'
        };
      }
    }
  } finally {
    lock.releaseLock();
  }

  throw new Error('Catatan absensi otomatis tidak ditemukan.');
}


/* ===== PORTAL TERPADU PKPPS - TAHAP LOGIN DAN HAK AKSES ===== */
const PORTAL_ROLES_ = ['ADMIN','TU','BENDAHARA','GURU','WALI','SISWA'];
const PORTAL_USER_SHEET_ = 'PENGGUNA';

function setupPortalTerpadu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Spreadsheet database terlebih dahulu.');
  const headers = ['ID Pengguna','Username','PIN Hash','Nama Lengkap','Peran','Kode Guru','ID Wali','ID Siswa','Aktif','Dibuat Pada'];
  let sh = ss.getSheetByName(PORTAL_USER_SHEET_);
  if (!sh) sh = ss.insertSheet(PORTAL_USER_SHEET_);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  const existing = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues() : [];
  const usernames = new Set(existing.map(r=>String(r[1]||'').trim().toLowerCase()));
  const guru = ss.getSheetByName(APP.sheets.guru);
  if (guru && guru.getLastRow() > 1) {
    const values = guru.getDataRange().getValues();
    const h = values[0].map(x=>String(x).trim());
    const idx = key => h.indexOf(key);
    values.slice(1).forEach(r=>{
      const username=String(r[idx('Username')]||'').trim();
      if (!username || usernames.has(username.toLowerCase())) return;
      const role=String(r[idx('Peran')]||'GURU').trim().toUpperCase()==='ADMIN'?'ADMIN':'GURU';
      sh.appendRow(['USR-'+Utilities.getUuid().slice(0,8),username,r[idx('PIN Hash')],r[idx('Nama Guru')],role,String(r[idx('Kode Guru')]),'','','YA',new Date()]);
      usernames.add(username.toLowerCase());
    });
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');
  return 'PENGGUNA siap. Akun ADMIN dan GURU yang ada pada sheet GURU disalin dengan hash PIN yang sama. Data lama tidak dihapus.';
}

function loginPortal(username, pin) {
  username=String(username||'').trim().toLowerCase();
  pin=String(pin||'').trim();
  if(!username||!pin) throw new Error('Username dan PIN wajib diisi.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if(!sh || sh.getLastRow()<2) throw new Error('Admin perlu menjalankan setupPortalTerpadu() terlebih dahulu.');
  const vals=sh.getDataRange().getValues();
  const heads=vals[0].map(x=>String(x).trim());
  const at=k=>heads.indexOf(k);
  const row=vals.slice(1).find(r=>String(r[at('Username')]||'').trim().toLowerCase()===username &&
    String(r[at('PIN Hash')]||'')===hashPin_(pin) && String(r[at('Aktif')]||'').toUpperCase()==='YA');
  if(!row) throw new Error('Username/PIN salah atau akun tidak aktif.');
  const role=String(row[at('Peran')]||'').toUpperCase();
  if(!PORTAL_ROLES_.includes(role)) throw new Error('Peran akun belum dikenali.');
  const roleText=at('Daftar Peran')>=0?String(row[at('Daftar Peran')]||''):'';
  const roles=[...new Set([role,...roleText.split(',').map(x=>x.trim().toUpperCase()).filter(x=>PORTAL_ROLES_.includes(x))])];
  const user={id:String(row[at('ID Pengguna')]),username,nama:String(row[at('Nama Lengkap')]),peran:role,roles,
    kodeGuru:String(row[at('Kode Guru')]||''),idWali:String(row[at('ID Wali')]||''),idSiswa:String(row[at('ID Siswa')]||'')};
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('portal_'+token,JSON.stringify(user),21600);
  return {token,user,roles};
}

function activatePortalRole(token, role) {
  const user=getPortalSession_(token); role=String(role||'').trim().toUpperCase();
  const roles=Array.isArray(user.roles)?user.roles:[user.peran];
  if(!roles.includes(role)) throw new Error('Peran tersebut tidak terdaftar pada akun ini.');
  user.peran=role; CacheService.getScriptCache().put('portal_'+String(token),JSON.stringify(user),21600);
  return {ok:true,peran:role};
}

function getPortalSession_(token) {
  const raw=CacheService.getScriptCache().get('portal_'+String(token||''));
  if(!raw) throw new Error('Sesi portal berakhir. Silakan login kembali.');
  return JSON.parse(raw);
}
function logoutPortal(token) {
  if(token) CacheService.getScriptCache().remove('portal_'+token);
  return true;
}

function isPortalTeacherHomeroom_(user) {
  const code=normalizeCode_(user && (user.kodeGuru || user.kode) || '');
  if(!code) return false;
  const sh=SpreadsheetApp.getActive().getSheetByName('KELAS');
  if(!sh || sh.getLastRow()<2) return false;
  const values=sh.getDataRange().getDisplayValues();
  const headers=values[0].map(x=>String(x||'').trim());
  const codeCol=headers.indexOf('Kode Wali Kelas');
  if(codeCol<0) return false;
  const statusCol=headers.indexOf('Status');
  const activeCol=headers.indexOf('Aktif');
  return values.slice(1).some(r=>{
    const assigned=normalizeCode_(r[codeCol]);
    const status=String(r[statusCol>=0?statusCol:activeCol]||'AKTIF').trim().toUpperCase();
    return assigned===code && !['NONAKTIF','TIDAK AKTIF','NON AKTIF','FALSE','0'].includes(status);
  });
}


function portalAdminOrTU_(token){
  const u=getPortalSession_(token);const role=String(u.peran||'').trim().toUpperCase();
  if(!['ADMIN','TU'].includes(role))throw new Error('Akses khusus Admin/TU.');
  return u;
}
function portalAdmin_(token){const u=getPortalSession_(token);if(String(u.peran||'').trim().toUpperCase()!=='ADMIN')throw new Error('Akses khusus Admin.');return u;}

function getPortalDashboard(token) {
  const user=getPortalSession_(token);
  const ss=SpreadsheetApp.getActive();
  const count=name=>{const sh=ss.getSheetByName(name);return sh?Math.max(0,sh.getLastRow()-1):0;};
  const menus={
    ADMIN:['Ringkasan','Honor Guru','Data Pengguna','Data Guru','Data Santri','Kelas & Wali Kelas','Jadwal KBM','Absensi KBM','Absensi Santri','Nilai & Rapor','Jurnal KBM','Tahfidz','Keuangan','PPDB','CBT','Pengaturan'],
    TU:['Ringkasan','Data Santri','Data Guru','Jadwal KBM','Monitoring KBM','Absensi Santri','Nilai & Rapor','Tahfidz','PPDB'],
    BENDAHARA:['Ringkasan','Honor Guru','Tagihan','Pembayaran','Konfirmasi Pembayaran','Kuitansi','Tabungan','Laporan Keuangan'],
    GURU:['Ringkasan','Jadwal KBM','Absensi KBM','Absensi Santri','Jurnal Mengajar','Input Nilai','Tahfidz'],
    WALI:['Ringkasan','Profil Anak','Nilai & Rapor','Kehadiran','Tahfidz','Tagihan & Pembayaran','Tabungan','Pengumuman'],
    SISWA:['Ringkasan','Jadwal Pelajaran','Nilai','Kehadiran','Tahfidz','Pengumuman']
  };
  const isWaliKelas=user.peran==='GURU' ? isPortalTeacherHomeroom_(user) : false;
  const visibleMenus=(menus[user.peran]||[]).filter(m=>!(user.peran==='GURU' && m==='Rapor Wali Kelas' && !isWaliKelas));
  return {user:Object.assign({},user,{isWaliKelas}),menus:visibleMenus,counts:{guru:count('GURU'),jadwal:count('JADWAL'),absensi:count('ABSENSI'),pengguna:count('PENGGUNA'),santri:count(ss.getSheetByName('SANTRI')?'SANTRI':'SISWA')}};
}

/** Admin dapat membuat akun untuk peran selain akun admin yang sedang login. */
function createPortalUser(token, data) {
  const admin=getPortalSession_(token);
  if(admin.peran!=='ADMIN') throw new Error('Hanya admin yang boleh membuat akun.');
  data=data||{};
  const username=String(data.username||'').trim().toLowerCase();
  const pin=String(data.pin||'').trim();
  const nama=String(data.nama||'').trim();
  const role=String(data.peran||'').trim().toUpperCase();
  if(!username||!nama||pin.length<6) throw new Error('Nama, username, dan PIN minimal 6 karakter wajib diisi.');
  if(!['TU','BENDAHARA','GURU','WALI','SISWA'].includes(role)) throw new Error('ADMIN hanya dapat membuat akun TU, BENDAHARA, GURU, WALI, atau SISWA melalui formulir ini.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  const vals=sh.getDataRange().getValues(), heads=vals[0].map(x=>String(x).trim()), at=k=>heads.indexOf(k);
  if(vals.slice(1).some(r=>String(r[at('Username')]||'').trim().toLowerCase()===username)) throw new Error('Username sudah digunakan.');
  sh.appendRow(['USR-'+Utilities.getUuid().slice(0,8),username,hashPin_(pin),nama,role,String(data.kodeGuru||''),String(data.idWali||''),String(data.idSiswa||''),'YA',new Date()]);
  return 'Akun '+role+' berhasil dibuat.';
}


/* ===== MANAJEMEN PENGGUNA ADMIN - V1.3 ===== */
function getPortalUsers(token) {
  const admin=getPortalSession_(token);
  if(admin.peran!=='ADMIN') throw new Error('Akses ditolak. Hanya ADMIN yang dapat melihat daftar pengguna.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if(!sh || sh.getLastRow()<2) return [];
  const vals=sh.getDataRange().getValues();
  const heads=vals[0].map(x=>String(x).trim());
  const at=k=>heads.indexOf(k);
  return vals.slice(1).filter(r=>r[at('Username')]).map(r=>({
    id:String(r[at('ID Pengguna')]||''),
    username:String(r[at('Username')]||''),
    nama:String(r[at('Nama Lengkap')]||''),
    peran:String(r[at('Peran')]||''),
    roles:[...new Set([String(r[at('Peran')]||'').toUpperCase(),...(at('Daftar Peran')>=0?String(r[at('Daftar Peran')]||''):'').split(',').map(x=>x.trim().toUpperCase())].filter(x=>PORTAL_ROLES_.includes(x)))],
    kodeGuru:String(r[at('Kode Guru')]||''),
    aktif:String(r[at('Aktif')]||'').toUpperCase()==='YA'
  })).sort((a,b)=>a.peran.localeCompare(b.peran)||a.nama.localeCompare(b.nama,'id'));
}

function setPortalUserRoles(token, userId, roles) {
  const admin=getPortalSession_(token);
  if(String(admin.peran||'').toUpperCase()!=='ADMIN') throw new Error('Hanya ADMIN yang dapat mengatur peran.');
  const allowed=['ADMIN','TU','BENDAHARA','GURU','WALI','SISWA'];
  roles=[...new Set((Array.isArray(roles)?roles:[]).map(x=>String(x).trim().toUpperCase()).filter(x=>allowed.includes(x)))];
  if(!roles.length) throw new Error('Pilih minimal satu peran.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if(!sh) throw new Error('Sheet PENGGUNA tidak ditemukan.');
  const heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
  let roleListCol=heads.indexOf('Daftar Peran')+1;
  if(!roleListCol){roleListCol=sh.getLastColumn()+1;sh.getRange(1,roleListCol).setValue('Daftar Peran');}
  const idCol=heads.indexOf('ID Pengguna')+1, primaryCol=heads.indexOf('Peran')+1;
  const vals=sh.getDataRange().getValues();
  const idx=vals.findIndex((r,i)=>i>0&&String(r[idCol-1])===String(userId));
  if(idx<1) throw new Error('Akun pengguna tidak ditemukan.');
  const primary=roles.includes(String(vals[idx][primaryCol-1]||'').toUpperCase())?String(vals[idx][primaryCol-1]).toUpperCase():roles[0];
  sh.getRange(idx+1,primaryCol).setValue(primary);
  sh.getRange(idx+1,roleListCol).setValue(roles.join(', '));
  return 'Peran pengguna berhasil diperbarui: '+roles.join(', ')+'. Pengguna perlu login kembali.';
}

function setPortalUserActive(token, userId, active) {
  const admin=getPortalSession_(token);
  if(admin.peran!=='ADMIN') throw new Error('Hanya ADMIN yang dapat mengubah status akun.');
  if(String(userId)===admin.id && !active) throw new Error('Akun ADMIN yang sedang digunakan tidak dapat dinonaktifkan.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if(!sh) throw new Error('Sheet PENGGUNA tidak ditemukan.');
  const vals=sh.getDataRange().getValues(), heads=vals[0].map(x=>String(x).trim());
  const idCol=heads.indexOf('ID Pengguna'), activeCol=heads.indexOf('Aktif');
  const row=vals.findIndex((r,i)=>i>0 && String(r[idCol])===String(userId));
  if(row<1) throw new Error('Akun pengguna tidak ditemukan.');
  sh.getRange(row+1,activeCol+1).setValue(active?'YA':'TIDAK');
  return 'Status akun berhasil diperbarui.';
}

function resetPortalUserPin(token, userId, newPin) {
  const admin=getPortalSession_(token);
  if(admin.peran!=='ADMIN') throw new Error('Hanya ADMIN yang dapat mereset PIN.');
  newPin=String(newPin||'').trim();
  if(!/^\d{6,12}$/.test(newPin)) throw new Error('PIN baru harus 6–12 digit angka.');
  const sh=SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if(!sh) throw new Error('Sheet PENGGUNA tidak ditemukan.');
  const vals=sh.getDataRange().getValues(), heads=vals[0].map(x=>String(x).trim());
  const idCol=heads.indexOf('ID Pengguna'), pinCol=heads.indexOf('PIN Hash');
  const row=vals.findIndex((r,i)=>i>0 && String(r[idCol])===String(userId));
  if(row<1) throw new Error('Akun pengguna tidak ditemukan.');
  sh.getRange(row+1,pinCol+1).setValue(hashPin_(newPin));
  return 'PIN akun berhasil direset. Sampaikan PIN baru kepada pemilik akun secara pribadi.';
}


/**
 * Menghapus satu akun dari sheet PENGGUNA.
 * Hanya ADMIN; tidak menghapus data guru/jadwal/absensi terkait.
 * Melindungi akun yang sedang dipakai dan ADMIN aktif terakhir.
 */
function deletePortalUser(token, userId) {
  const admin = getPortalSession_(token);
  if (String(admin.peran || '').trim().toUpperCase() !== 'ADMIN') {
    throw new Error('Akses ditolak. Hanya ADMIN yang dapat menghapus akun.');
  }
  userId = String(userId || '').trim();
  if (!userId) throw new Error('ID pengguna tidak valid.');
  if (userId === String(admin.id)) {
    throw new Error('Akun ADMIN yang sedang digunakan tidak dapat dihapus.');
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(PORTAL_USER_SHEET_);
  if (!sh) throw new Error('Sheet PENGGUNA tidak ditemukan.');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('Data pengguna kosong.');

  const heads = vals[0].map(x => String(x).trim());
  const idCol = heads.indexOf('ID Pengguna');
  const roleCol = heads.indexOf('Peran');
  const activeCol = heads.indexOf('Aktif');
  if (idCol < 0 || roleCol < 0 || activeCol < 0) {
    throw new Error('Struktur kolom sheet PENGGUNA tidak sesuai.');
  }

  const rowIndex = vals.findIndex((r, i) => i > 0 && String(r[idCol]).trim() === userId);
  if (rowIndex < 1) throw new Error('Akun pengguna tidak ditemukan.');

  const targetRole = String(vals[rowIndex][roleCol] || '').trim().toUpperCase();
  const targetActive = String(vals[rowIndex][activeCol] || '').trim().toUpperCase() === 'YA';

  if (targetRole === 'ADMIN' && targetActive) {
    const activeAdmins = vals.slice(1).filter(r =>
      String(r[roleCol] || '').trim().toUpperCase() === 'ADMIN' &&
      String(r[activeCol] || '').trim().toUpperCase() === 'YA'
    ).length;
    if (activeAdmins <= 1) {
      throw new Error('ADMIN aktif terakhir tidak dapat dihapus. Tambahkan ADMIN pengganti terlebih dahulu.');
    }
  }

  sh.deleteRow(rowIndex + 1);
  return 'Akun berhasil dihapus dari PENGGUNA. Data guru, jadwal, dan riwayat absensi tidak ikut dihapus.';
}



/* ===== DIREKTORI DATA SANTRI V1.24 ===== */
function getSantriDirectory(token){
  const u=getPortalSession_(token);
  if(!['ADMIN','TU'].includes(String(u.peran||'').trim().toUpperCase())) throw new Error('Akses Data Santri hanya untuk ADMIN/TU.');
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('SANTRI')||ss.getSheetByName('SISWA');
  if(!sh||sh.getLastRow()<2) return [];
  const values=sh.getDataRange().getDisplayValues();
  const headers=values[0].map(x=>String(x||'').trim());
  const ix=(...names)=>{for(const n of names){const i=headers.indexOf(n);if(i>=0)return i;}return -1;};
  const c={nis:ix('NIS','NISN','ID Santri','ID Siswa'),nama:ix('Nama Santri','Nama Siswa','Nama','Nama Lengkap'),gender:ix('Jenis Kelamin','Gender'),kelas:ix('Kelas','Nama Kelas'),jenjang:ix('Jenjang'),status:ix('Status')};
  if(c.nama<0) throw new Error('Kolom Nama Santri tidak ditemukan pada sheet '+sh.getName()+'.');
  return values.slice(1).filter(r=>r[c.nama]).map(r=>({nis:c.nis>=0?String(r[c.nis]||''):'',nama:String(r[c.nama]||''),gender:c.gender>=0?String(r[c.gender]||''):'',kelas:c.kelas>=0?String(r[c.kelas]||''):'',jenjang:c.jenjang>=0?String(r[c.jenjang]||''):'',status:c.status>=0?String(r[c.status]||''):'Aktif'}));
}

/* ===== MODUL HONOR GURU ===== */
const HONOR_SHEET_='HONOR_GURU';
const HONOR_HEADERS_=['ID','Bulan','Nama Guru','Program','Kelas','Mata Pelajaran','JP Terlaksana','Tarif per JP','Total Honor','Verifikasi','Keterangan','Dicatat Oleh','Waktu Catat'];
function ensureHonorSheet_(){
 const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(HONOR_SHEET_);
 if(!sh){sh=ss.insertSheet(HONOR_SHEET_);sh.appendRow(HONOR_HEADERS_);sh.setFrozenRows(1);}
 return sh;
}
function honorUser_(token){
 const u=getPortalSession_(token);
 if(!['ADMIN','BENDAHARA'].includes(String(u.peran||'').trim().toUpperCase()))throw new Error('Akses honor hanya untuk ADMIN dan BENDAHARA.');
 return u;
}

function getHonorSetup(token){
 honorUser_(token);
 const teachers=readObjects_('GURU').filter(t=>isActive_(t['Aktif'])).map(t=>({kode:String(t['Kode Guru']||''),nama:String(t['Nama Guru']||'')})).filter(t=>t.kode&&t.nama).sort((a,b)=>a.nama.localeCompare(b.nama,'id'));
 const assignments=[];
 const add=(kode,kelas,mapel,program)=>{if(!kode||!kelas||!mapel)return;const key=[kode,kelas,mapel].join('|');if(!assignments.some(x=>x.key===key))assignments.push({key,kode:String(kode),kelas:String(kelas),mapel:String(mapel),program:String(program||'')});};
 readObjects_('JADWAL').filter(x=>isActive_(x['Aktif'])).forEach(x=>add(normalizeCode_(x['Kode Guru']),String(x['Kelas']||''),String(x['Mata Pelajaran']||''),''));
 if(SpreadsheetApp.getActive().getSheetByName('GURU_MAPEL'))readObjects_('GURU_MAPEL').forEach(x=>add(normalizeCode_(x['Kode Guru']),String(x['Kelas']||''),String(x['Nama Mapel']||''),''));
 return {teachers,assignments};
}

function getHonorRecords(token,month){
 honorUser_(token);const sh=ensureHonorSheet_();if(sh.getLastRow()<2)return {rows:[],summary:{jp:0,total:0,verified:0,pending:0}};
 const v=sh.getDataRange().getValues(),rows=v.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>({id:String(r[0]),bulan:String(r[1]),namaGuru:String(r[2]),program:String(r[3]),kelas:String(r[4]),mapel:String(r[5]),jp:Number(r[6])||0,tarif:Number(r[7])||0,total:Number(r[8])||0,verifikasi:String(r[9]),keterangan:String(r[10]||'')})).filter(x=>!month||x.bulan===month);
 const sum=a=>a.reduce((n,x)=>n+x,0);
 return {rows,summary:{jp:sum(rows.map(x=>x.jp)),total:sum(rows.map(x=>x.total)),verified:sum(rows.filter(x=>x.verifikasi==='Diverifikasi').map(x=>x.total)),pending:sum(rows.filter(x=>x.verifikasi!=='Diverifikasi').map(x=>x.total))}};
}
function saveHonorRecord(token,data){
 const u=honorUser_(token);data=data||{};
 const bulan=String(data.bulan||'').trim(),nama=String(data.namaGuru||'').trim(),jp=Number(data.jp),tarif=Number(data.tarif);
 if(!/^\d{4}-\d{2}$/.test(bulan))throw new Error('Pilih bulan rekap.');
 if(!nama)throw new Error('Nama guru wajib diisi.');
 if(!Number.isFinite(jp)||jp<0||!Number.isFinite(tarif)||tarif<0)throw new Error('Jumlah JP dan tarif harus berupa angka positif.');
 const sh=ensureHonorSheet_(),id='HR-'+Utilities.getUuid().slice(0,8);
 sh.appendRow([id,bulan,nama,String(data.program||''),String(data.kelas||''),String(data.mapel||''),jp,tarif,jp*tarif,String(data.verifikasi||'Belum'),String(data.keterangan||''),u.nama,new Date()]);
 return 'Rekap honor berhasil disimpan.';
}

/* ===== MODUL KEUANGAN BENDAHARA V1.7 ===== */
const FINANCE_SHEETS_ = {
  TAGIHAN: ['ID Tagihan','Tanggal','Nama Santri','Kelas','Jenis Tagihan','Keterangan','Nominal','Jatuh Tempo','Status','Dibuat Oleh','Dibuat Pada'],
  PEMBAYARAN: ['ID Pembayaran','Tanggal','Nama Santri','Kelas','Jenis Pembayaran','Keterangan','Nominal','Metode','Referensi','Dicatat Oleh','Dicatat Pada','ID Tagihan'],
  PENGELUARAN: ['ID Pengeluaran','Tanggal','Kategori','Keterangan','Nominal','Metode','Bukti/Referensi','Dicatat Oleh','Dicatat Pada']
};
function ensureFinanceSheets_(){
  const ss=SpreadsheetApp.getActive();
  Object.keys(FINANCE_SHEETS_).forEach(name=>{
    let sh=ss.getSheetByName(name);
    if(!sh){sh=ss.insertSheet(name);sh.appendRow(FINANCE_SHEETS_[name]);sh.setFrozenRows(1);}
    else if(name==='PEMBAYARAN'){const h=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(x=>String(x).trim());if(!h.includes('ID Tagihan'))sh.getRange(1,sh.getLastColumn()+1).setValue('ID Tagihan');}
  });
}
function financeUser_(token){
  const u=getPortalSession_(token);
  if(!['ADMIN','BENDAHARA'].includes(String(u.peran||'').trim().toUpperCase())) throw new Error('Akses keuangan hanya untuk ADMIN dan BENDAHARA.');
  ensureFinanceSheets_();
  return u;
}
function financeRead_(name){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(!sh||sh.getLastRow()<2)return [];
  const v=sh.getDataRange().getValues(),h=v[0].map(x=>String(x).trim());
  return v.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i] instanceof Date?Utilities.formatDate(r[i],Session.getScriptTimeZone(),'yyyy-MM-dd'):r[i]])));
}
function getFinanceData(token){
  financeUser_(token);
  const tagihan=financeRead_('TAGIHAN'),pembayaran=financeRead_('PEMBAYARAN'),pengeluaran=financeRead_('PENGELUARAN');
  const sum=a=>a.reduce((s,x)=>s+(Number(x.Nominal)||0),0);
  const paidByInvoice={};
  pembayaran.forEach(p=>{const id=String(p['ID Tagihan']||'').trim();if(id)paidByInvoice[id]=(paidByInvoice[id]||0)+(Number(p.Nominal)||0);});
  tagihan.forEach(t=>{const paid=paidByInvoice[String(t['ID Tagihan'])]||0;t['Total Dibayar']=paid;t['Sisa Tagihan']=Math.max(0,(Number(t.Nominal)||0)-paid);t.Status=t['Sisa Tagihan']<=0?'LUNAS':(paid>0?'CICIL':'BELUM LUNAS');});
  const totalTagihan=sum(tagihan),totalBayar=sum(pembayaran),totalKeluar=sum(pengeluaran);
  const tunggakan=tagihan.reduce((s,t)=>s+(Number(t['Sisa Tagihan'])||0),0);
  return {tagihan,pembayaran,pengeluaran,summary:{tagihan:totalTagihan,pembayaran:totalBayar,pengeluaran:totalKeluar,saldo:totalBayar-totalKeluar,tunggakan}};
}
function saveFinanceRecord(token,kind,data){
  const u=financeUser_(token);data=data||{};
  const map={TAGIHAN:'TAGIHAN',PEMBAYARAN:'PEMBAYARAN',PENGELUARAN:'PENGELUARAN'};
  const sheet=map[String(kind||'').toUpperCase()];
  if(!sheet)throw new Error('Jenis pencatatan tidak valid.');
  const amount=Number(data.nominal);
  if(!Number.isFinite(amount)||amount<=0)throw new Error('Nominal harus berupa angka lebih dari nol.');
  const date=String(data.tanggal||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Tanggal wajib diisi.');
  let row;
  if(sheet==='TAGIHAN'){
    if(!String(data.nama||'').trim()||!String(data.jenis||'').trim())throw new Error('Nama santri dan jenis tagihan wajib diisi.');
    row=['INV-'+Utilities.getUuid().slice(0,8),date,data.nama,data.kelas||'',data.jenis,data.keterangan||'',amount,data.jatuhTempo||'','BELUM LUNAS',u.nama,new Date()];
  }else if(sheet==='PEMBAYARAN'){
    if(!String(data.nama||'').trim()||!String(data.jenis||'').trim())throw new Error('Nama santri dan jenis pembayaran wajib diisi.');
    row=['BYR-'+Utilities.getUuid().slice(0,8),date,data.nama,data.kelas||'',data.jenis,data.keterangan||'',amount,data.metode||'Tunai',data.referensi||'',u.nama,new Date()];
  }else{
    if(!String(data.kategori||'').trim()||!String(data.keterangan||'').trim())throw new Error('Kategori dan keterangan pengeluaran wajib diisi.');
    row=['KLR-'+Utilities.getUuid().slice(0,8),date,data.kategori,data.keterangan,amount,data.metode||'Tunai',data.referensi||'',u.nama,new Date()];
  }
  SpreadsheetApp.getActive().getSheetByName(sheet).appendRow(row);
  return 'Data '+sheet+' berhasil disimpan.';
}


function recordInvoicePayment(token,data){
  const u=financeUser_(token);data=data||{};
  const invoiceId=String(data.invoiceId||'').trim();
  if(!invoiceId)throw new Error('Pilih tagihan yang akan dibayar.');
  const amount=Number(data.nominal);
  if(!Number.isFinite(amount)||amount<=0)throw new Error('Nominal pembayaran harus lebih dari nol.');
  const date=String(data.tanggal||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Tanggal pembayaran wajib diisi.');
  ensureFinanceSheets_();
  const inv=financeRead_('TAGIHAN').find(x=>String(x['ID Tagihan'])===invoiceId);
  if(!inv)throw new Error('Tagihan tidak ditemukan.');
  const payments=financeRead_('PEMBAYARAN');
  const already=payments.filter(x=>String(x['ID Tagihan']||'')===invoiceId).reduce((s,x)=>s+(Number(x.Nominal)||0),0);
  const remaining=Math.max(0,Number(inv.Nominal||0)-already);
  if(amount>remaining)throw new Error('Nominal melebihi sisa tagihan '+remaining+'.');
  const sh=SpreadsheetApp.getActive().getSheetByName('PEMBAYARAN');
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const record={
    'ID Pembayaran':'BYR-'+Utilities.getUuid().slice(0,8),'Tanggal':date,
    'Nama Santri':inv['Nama Santri'],'Kelas':inv.Kelas||'',
    'Jenis Pembayaran':inv['Jenis Tagihan'],'Keterangan':data.keterangan||'Pembayaran tagihan '+invoiceId,
    'Nominal':amount,'Metode':data.metode||'Tunai','Referensi':data.referensi||'',
    'Dicatat Oleh':u.nama,'Dicatat Pada':new Date(),'ID Tagihan':invoiceId
  };
  sh.appendRow(headers.map(h=>record[h]===undefined?'':record[h]));
  const totalPaid=already+amount,remainingAfter=Math.max(0,Number(inv.Nominal||0)-totalPaid);
  return {message:'Pembayaran berhasil dicatat.',receipt:{id:record['ID Pembayaran'],tanggal:date,nama:inv['Nama Santri'],kelas:inv.Kelas||'',jenis:inv['Jenis Tagihan'],nominal:amount,metode:record.Metode,referensi:record.Referensi,invoiceId,totalTagihan:Number(inv.Nominal||0),totalDibayar:totalPaid,sisa:remainingAfter,status:remainingAfter<=0?'LUNAS':'CICIL',petugas:u.nama}};
}


/* ===== CETAK ULANG KUITANSI PEMBAYARAN V1.11 ===== */
function getPaymentReceipt(token, paymentId) {
  financeUser_(token);

  paymentId = String(paymentId || '').trim();
  if (!paymentId) throw new Error('ID pembayaran tidak valid.');

  const payments = financeRead_('PEMBAYARAN');
  const index = payments.findIndex(
    p => String(p['ID Pembayaran'] || '').trim() === paymentId
  );
  if (index < 0) throw new Error('Transaksi pembayaran tidak ditemukan.');

  const p = payments[index];
  const invoiceId = String(p['ID Tagihan'] || '').trim();
  if (!invoiceId) {
    throw new Error('Transaksi ini belum terhubung dengan ID Tagihan sehingga kuitansi rinci tidak dapat dibuat.');
  }

  const invoice = financeRead_('TAGIHAN').find(
    t => String(t['ID Tagihan'] || '').trim() === invoiceId
  );
  if (!invoice) throw new Error('Data tagihan tidak ditemukan.');

  // Akumulasi pembayaran untuk invoice ini sampai transaksi yang dipilih.
  const totalDibayar = payments
    .slice(0, index + 1)
    .filter(x => String(x['ID Tagihan'] || '').trim() === invoiceId)
    .reduce((sum, x) => sum + (Number(x.Nominal) || 0), 0);

  const totalTagihan = Number(invoice.Nominal) || 0;
  const sisa = Math.max(0, totalTagihan - totalDibayar);

  return {
    id: p['ID Pembayaran'],
    tanggal: p.Tanggal,
    nama: p['Nama Santri'],
    kelas: p.Kelas || '',
    jenis: p['Jenis Pembayaran'],
    nominal: Number(p.Nominal) || 0,
    metode: p.Metode || '',
    referensi: p.Referensi || '',
    invoiceId: invoiceId,
    totalTagihan: totalTagihan,
    totalDibayar: totalDibayar,
    sisa: sisa,
    status: sisa <= 0 ? 'LUNAS' : (totalDibayar > 0 ? 'CICIL' : 'BELUM LUNAS'),
    petugas: p['Dicatat Oleh'] || ''
  };
}


/* ===== MODUL ABSENSI KBM + JURNAL MENGAJAR PORTAL V1.13 ===== */
const JOURNAL_SHEET_ = 'JURNAL_KBM';
const JOURNAL_HEADERS_ = ['ID Jurnal','Tanggal','Hari','Waktu Input','ID Jadwal','Kode Guru','Nama Guru','Kelas','Jam Ke','Mulai','Selesai','Mata Pelajaran','Status Kehadiran','Materi Pembelajaran','Metode Pembelajaran','Catatan'];

// Mendukung dua variasi judul kolom pada sheet JADWAL: 'Jam Mulai'/'Selesai'
// maupun 'Mulai'/'Selesai'. Nilai waktu dinormalisasi ke teks HH:mm.
function getScheduleStart_(row) {
  return row['Jam Mulai'] !== undefined && row['Jam Mulai'] !== '' ? row['Jam Mulai'] : row['Mulai'];
}
function getScheduleEnd_(row) {
  return row['Jam Selesai'] !== undefined && row['Jam Selesai'] !== '' ? row['Jam Selesai'] : row['Selesai'];
}

function getTeacherKbmData(token, selectedDate) {
  const user = getPortalSession_(token);
  if (!['ADMIN','GURU'].includes(String(user.peran||'').toUpperCase())) throw new Error('Menu Absensi KBM hanya untuk ADMIN dan GURU.');
  const dateKey = String(selectedDate || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Tanggal tidak valid.');
  const parts=dateKey.split('-').map(Number);
  const d=new Date(Date.UTC(parts[0],parts[1]-1,parts[2],12));
  if(d.getUTCFullYear()!==parts[0]||d.getUTCMonth()!==parts[1]-1||d.getUTCDate()!==parts[2]) throw new Error('Tanggal tidak valid.');
  const days=['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const day=days[d.getUTCDay()];
  // Baca jadwal menggunakan getDisplayValues agar kolom waktu tetap berupa teks HH:mm,
  // bukan objek Date Sheets (yang dapat tampil sebagai 1899-12-30 di browser).
  const jadwalSheet = SpreadsheetApp.getActive().getSheetByName('JADWAL');
  if (!jadwalSheet || jadwalSheet.getLastRow() < 2) return {date:dateKey,day,schedules:[],user:{nama:user.nama,peran:user.peran},total:0,done:0};
  const jadwalValues = jadwalSheet.getDataRange().getDisplayValues();
  const jadwalHeaders = jadwalValues[0].map(h => String(h).trim());
  const all = jadwalValues.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {}; jadwalHeaders.forEach((h,i) => obj[h] = row[i]); return obj;
  }).filter(x=>isActive_(x['Aktif']) && normalizeDay_(x['Hari'])===normalizeDay_(day));
  const isAdmin=String(user.peran).toUpperCase()==='ADMIN';
  const schedules=all.filter(x=>isAdmin || normalizeCode_(x['Kode Guru'])===normalizeCode_(user.kodeGuru)).map(x=>({
    id:String(x['ID Jadwal']||''),hari:day,jamKe:String(x['Jam Ke']||''),mulai:displayTime_(getScheduleStart_(x)),selesai:displayTime_(getScheduleEnd_(x)),
    kelas:String(x['Kelas']||''),mapel:String(x['Mata Pelajaran']||''),kodeGuru:String(x['Kode Guru']||''),namaGuru:String(x['Nama Guru']||'')
  })).sort((a,b)=>timeToMinutes_(a.mulai)-timeToMinutes_(b.mulai)||a.kelas.localeCompare(b.kelas,'id'));
  const attendance=readObjects_('ABSENSI').filter(x=>normalizeDate_(x['Tanggal'])===dateKey);
  const journalSheet=SpreadsheetApp.getActive().getSheetByName(JOURNAL_SHEET_);
  const journals=journalSheet&&journalSheet.getLastRow()>1?readObjects_(JOURNAL_SHEET_).filter(x=>normalizeDate_(x['Tanggal'])===dateKey):[];
  schedules.forEach(x=>{
    const a=attendance.find(r=>String(r['ID Jadwal'])===x.id && normalizeCode_(r['Kode Guru'])===normalizeCode_(x.kodeGuru));
    const j=journals.find(r=>String(r['ID Jadwal'])===x.id && normalizeCode_(r['Kode Guru'])===normalizeCode_(x.kodeGuru));
    x.status=a?String(a['Status']||''):''; x.sudahAbsen=!!a;
    x.materi=j?String(j['Materi Pembelajaran']||''):'';x.metode=j?String(j['Metode Pembelajaran']||''):'';x.catatan=j?String(j['Catatan']||''):'';x.sudahJurnal=!!j;
  });
  return {date:dateKey,day,schedules,user:{nama:user.nama,peran:user.peran},total:schedules.length,done:schedules.filter(x=>x.sudahAbsen&&x.sudahJurnal).length};
}

function saveTeacherKbm(token, payload) {
  const user=getPortalSession_(token);
  if(!['ADMIN','GURU'].includes(String(user.peran||'').toUpperCase())) throw new Error('Akses ditolak.');
  payload=payload||{};
  const id=String(payload.idJadwal||'').trim();
  const dateKey=String(payload.tanggal||'');
  const status=String(payload.status||'');
  const materi=String(payload.materi||'').trim();
  const metode=String(payload.metode||'').trim();
  const catatan=String(payload.catatan||'').trim();
  if(!id||!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Jadwal dan tanggal wajib diisi.');
  if(!['Hadir','Izin','Sakit','Tidak Hadir'].includes(status)) throw new Error('Pilih status kehadiran.');
  if(status==='Hadir' && !materi) throw new Error('Materi pembelajaran wajib diisi untuk KBM yang dilaksanakan.');
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const schedule=readObjects_('JADWAL').find(x=>String(x['ID Jadwal'])===id && isActive_(x['Aktif']));
    if(!schedule) throw new Error('Jadwal tidak ditemukan atau tidak aktif.');
    const teacherCode=String(schedule['Kode Guru']||'');
    if(String(user.peran).toUpperCase()!=='ADMIN' && normalizeCode_(teacherCode)!==normalizeCode_(user.kodeGuru)) throw new Error('Jadwal ini bukan milik akun guru Anda.');
    const ss=SpreadsheetApp.getActive();
    const abs=ss.getSheetByName('ABSENSI');if(!abs) throw new Error('Sheet ABSENSI tidak ditemukan. Hubungi admin.');
    const absVals=abs.getDataRange().getValues(), ah=absVals[0].map(x=>String(x).trim());
    const aCol=k=>ah.indexOf(k);
    ['ID Jadwal','Tanggal','Hari','Waktu Presensi','Kode Guru','Nama Guru','Kelas','Jam Ke','Mulai','Selesai','Mata Pelajaran','Status','Catatan'].forEach(k=>{if(aCol(k)<0)throw new Error('Kolom '+k+' tidak ditemukan pada sheet ABSENSI.');});
    const exists=absVals.slice(1).some(r=>normalizeDate_(r[aCol('Tanggal')])===dateKey&&String(r[aCol('ID Jadwal')])===id&&normalizeCode_(r[aCol('Kode Guru')])===normalizeCode_(teacherCode));
    if(exists) throw new Error('Absensi sesi ini sudah tersimpan. Hubungi admin jika perlu koreksi.');
    let jsh=ss.getSheetByName(JOURNAL_SHEET_);
    if(!jsh){jsh=ss.insertSheet(JOURNAL_SHEET_);jsh.appendRow(JOURNAL_HEADERS_);jsh.setFrozenRows(1);jsh.getRange(1,1,1,JOURNAL_HEADERS_.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');}
    const jvals=jsh.getDataRange().getValues(), jh=jvals[0].map(x=>String(x).trim()), jc=k=>jh.indexOf(k);
    if(jvals.slice(1).some(r=>normalizeDate_(r[jc('Tanggal')])===dateKey&&String(r[jc('ID Jadwal')])===id&&normalizeCode_(r[jc('Kode Guru')])===normalizeCode_(teacherCode))) throw new Error('Jurnal sesi ini sudah tersimpan. Hubungi admin jika perlu koreksi.');
    const day=String(schedule['Hari']||'');const now=new Date();
    const row=new Array(ah.length).fill('');
    const values={'ID Jadwal':id,'Tanggal':dateKey,'Hari':day,'Waktu Presensi':now,'Kode Guru':teacherCode,'Nama Guru':String(schedule['Nama Guru']||user.nama),'Kelas':String(schedule['Kelas']||''),'Jam Ke':String(schedule['Jam Ke']||''),'Mulai':displayTime_(getScheduleStart_(schedule)),'Selesai':displayTime_(getScheduleEnd_(schedule)),'Mata Pelajaran':String(schedule['Mata Pelajaran']||''),'Status':status,'Catatan':catatan};
    Object.keys(values).forEach(k=>row[aCol(k)]=values[k]);
    abs.appendRow(row);
    const jr={'ID Jurnal':'JRN-'+Utilities.getUuid().slice(0,12),'Tanggal':dateKey,'Hari':day,'Waktu Input':now,'ID Jadwal':id,'Kode Guru':teacherCode,'Nama Guru':String(schedule['Nama Guru']||user.nama),'Kelas':String(schedule['Kelas']||''),'Jam Ke':String(schedule['Jam Ke']||''),'Mulai':displayTime_(getScheduleStart_(schedule)),'Selesai':displayTime_(getScheduleEnd_(schedule)),'Mata Pelajaran':String(schedule['Mata Pelajaran']||''),'Status Kehadiran':status,'Materi Pembelajaran':materi,'Metode Pembelajaran':metode,'Catatan':catatan};
    jsh.appendRow(JOURNAL_HEADERS_.map(k=>jr[k]??''));
    return 'Absensi dan jurnal KBM berhasil disimpan.';
  } finally {lock.releaseLock();}
}

function getKbmJournalReport(token, dateKey) {
  const user=getPortalSession_(token);
  if(!['ADMIN','GURU'].includes(String(user.peran||'').toUpperCase())) throw new Error('Akses ditolak.');
  const d=String(dateKey||Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'));
  const sh=SpreadsheetApp.getActive().getSheetByName(JOURNAL_SHEET_);if(!sh||sh.getLastRow()<2)return [];
  return readObjects_(JOURNAL_SHEET_).filter(r=>normalizeDate_(r['Tanggal'])===d&&(user.peran==='ADMIN'||normalizeCode_(r['Kode Guru'])===normalizeCode_(user.kodeGuru))).map(r=>({tanggal:d,guru:r['Nama Guru'],kelas:r['Kelas'],jam:r['Jam Ke'],mapel:r['Mata Pelajaran'],status:r['Status Kehadiran'],materi:r['Materi Pembelajaran'],metode:r['Metode Pembelajaran'],catatan:r['Catatan']}));
}


/* ===== LAPORAN ABSENSI KBM HARIAN / MINGGUAN / BULANAN V1.14 ===== */
function getKbmPeriodReport(token, period, selectedDate) {
  const user = getPortalSession_(token);
  const role = String(user.peran || '').toUpperCase();
  if (!['ADMIN','GURU'].includes(role)) throw new Error('Akses laporan KBM hanya untuk ADMIN dan GURU.');
  const key = String(selectedDate || Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) throw new Error('Tanggal laporan tidak valid.');
  const parts = key.split('-').map(Number);
  const selected = new Date(Date.UTC(parts[0], parts[1]-1, parts[2], 12));
  if (selected.getUTCFullYear()!==parts[0] || selected.getUTCMonth()!==parts[1]-1 || selected.getUTCDate()!==parts[2]) throw new Error('Tanggal laporan tidak valid.');
  const mode = ['harian','mingguan','bulanan'].includes(String(period||'').toLowerCase()) ? String(period).toLowerCase() : 'harian';
  let start = new Date(selected), end = new Date(selected);
  if (mode === 'mingguan') {
    const offset = (selected.getUTCDay()+6)%7;
    start.setUTCDate(start.getUTCDate()-offset);
    end = new Date(start); end.setUTCDate(start.getUTCDate()+6);
  } else if (mode === 'bulanan') {
    start = new Date(Date.UTC(parts[0], parts[1]-1, 1, 12));
    end = new Date(Date.UTC(parts[0], parts[1], 0, 12));
  }
  const fmt = d => Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
  const startKey=fmt(start), endKey=fmt(end);
  const inRange = value => { const d=normalizeDate_(value); return d>=startKey && d<=endKey; };
  const own = row => role==='ADMIN' || normalizeCode_(row['Kode Guru'])===normalizeCode_(user.kodeGuru);
  const attendance = readObjects_('ABSENSI').filter(r=>inRange(r['Tanggal']) && own(r));
  const journals = readObjects_(JOURNAL_SHEET_).filter(r=>inRange(r['Tanggal']) && own(r));

  // Ambil jam langsung dari tampilan sheet JADWAL. getValues() dapat
  // mengembalikan objek Date dengan tanggal dasar 1899-12-30 untuk sel waktu.
  // getDisplayValues() mempertahankan tampilan jam seperti 08:00.
  const scheduleTimeMap = {};
  const scheduleSheet = SpreadsheetApp.getActive().getSheetByName('JADWAL');
  if (scheduleSheet && scheduleSheet.getLastRow() > 1) {
    const displayRows = scheduleSheet.getDataRange().getDisplayValues();
    const headers = displayRows[0].map(h => String(h).trim());
    const idIndex = headers.indexOf('ID Jadwal');
    const startIndex = headers.indexOf('Jam Mulai') >= 0 ? headers.indexOf('Jam Mulai') : headers.indexOf('Mulai');
    const endIndex = headers.indexOf('Jam Selesai') >= 0 ? headers.indexOf('Jam Selesai') : headers.indexOf('Selesai');
    if (idIndex >= 0) displayRows.slice(1).forEach(row => {
      const id = String(row[idIndex] || '').trim();
      if (id) scheduleTimeMap[id] = {
        mulai: startIndex >= 0 ? row[startIndex] : '',
        selesai: endIndex >= 0 ? row[endIndex] : ''
      };
    });
  }
  const reportTime = value => {
    if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, APP.timezone, 'HH:mm');
    if (typeof value === 'number' && isFinite(value)) {
      const mins = Math.round((value % 1) * 1440);
      return String(Math.floor(mins / 60) % 24).padStart(2, '0') + ':' + String(mins % 60).padStart(2, '0');
    }
    const text = String(value == null ? '' : value).trim();
    const match = text.match(/(?:T|\s)(\d{1,2}:\d{2})(?::\d{2})?/i) || text.match(/^(\d{1,2}:\d{2})(?::\d{2})?$/);
    if (match) {
      const hm = match[1].split(':');
      return String(Number(hm[0])).padStart(2, '0') + ':' + hm[1];
    }
    return text;
  };
  const journalMap = {};
  journals.forEach(j=>{journalMap[normalizeDate_(j['Tanggal'])+'|'+String(j['ID Jadwal']||'')+'|'+normalizeCode_(j['Kode Guru'])]=j;});
  const days = [];
  for (let d=new Date(start); d<=end; d.setUTCDate(d.getUTCDate()+1)) {
    const date=fmt(d), list=attendance.filter(r=>normalizeDate_(r['Tanggal'])===date);
    const jr=journals.filter(r=>normalizeDate_(r['Tanggal'])===date);
    const countStatus = status => list.filter(r=>String(r['Status']||'').trim().toLowerCase()===status.toLowerCase()).length;
    days.push({date, label:Utilities.formatDate(d, 'UTC', 'dd/MM'), hadir:countStatus('Hadir'), izin:countStatus('Izin'), sakit:countStatus('Sakit'), tidakHadir:countStatus('Tidak Hadir'), total:list.length, jurnal:jr.filter(r=>String(r['Materi Pembelajaran']||'').trim()!=='').length});
  }
  const countStatus = status => attendance.filter(r=>String(r['Status']||'').trim().toLowerCase()===status.toLowerCase()).length;
  const detail = attendance.map(a=>{
    const date=normalizeDate_(a['Tanggal']);
    const j=journalMap[date+'|'+String(a['ID Jadwal']||'')+'|'+normalizeCode_(a['Kode Guru'])] || {};
    const scheduled = scheduleTimeMap[String(a['ID Jadwal'] || '').trim()] || {};
    const startTime = reportTime(scheduled.mulai || a['Mulai']);
    const endTime = reportTime(scheduled.selesai || a['Selesai']);
    return {tanggal:date,guru:a['Nama Guru']||'',kelas:a['Kelas']||'',jam:(startTime||'--:--')+'–'+(endTime||'--:--'),mapel:a['Mata Pelajaran']||'',status:a['Status']||'',materi:j['Materi Pembelajaran']||'',metode:j['Metode Pembelajaran']||'',jurnal:!!String(j['Materi Pembelajaran']||'').trim()};
  }).sort((a,b)=>a.tanggal.localeCompare(b.tanggal)||String(a.jam).localeCompare(String(b.jam)));
  const totalJadwal = attendance.length;
  const totalJurnal = journals.filter(r=>String(r['Materi Pembelajaran']||'').trim()!=='').length;
  return {period:mode,start:startKey,end:endKey,userRole:role,userName:user.nama,summary:{total:totalJadwal,hadir:countStatus('Hadir'),izin:countStatus('Izin'),sakit:countStatus('Sakit'),tidakHadir:countStatus('Tidak Hadir'),jurnal:totalJurnal,belumJurnal:Math.max(0,totalJadwal-totalJurnal)},days,detail};
}




/* ===== DASHBOARD MONITORING KBM BACKEND V1.15 ===== */
function getKbmMonitoringDashboard(token, selectedDate) {
  const user = getPortalSession_(token);
  if (!['ADMIN','TU'].includes(String(user.peran || '').toUpperCase())) throw new Error('Dashboard monitoring hanya untuk ADMIN/TU.');
  const key = String(selectedDate || Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) throw new Error('Tanggal pemantauan tidak valid.');
  const parts = key.split('-').map(Number);
  const selected = new Date(Date.UTC(parts[0], parts[1]-1, parts[2], 12));
  if (selected.getUTCFullYear() !== parts[0] || selected.getUTCMonth() !== parts[1]-1 || selected.getUTCDate() !== parts[2]) throw new Error('Tanggal pemantauan tidak valid.');
  const days = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const dayName = days[selected.getUTCDay()];
  const schedules = readObjects_('JADWAL').filter(x => isActive_(x['Aktif']) && normalizeDay_(x['Hari']) === normalizeDay_(dayName));
  const attendanceAll = readObjects_('ABSENSI');
  const journalsAll = readObjects_(JOURNAL_SHEET_);
  const attendance = attendanceAll.filter(x => normalizeDate_(x['Tanggal']) === key);
  const journals = journalsAll.filter(x => normalizeDate_(x['Tanggal']) === key);
  const keyOf = x => String(x['ID Jadwal'] || '').trim() + '|' + normalizeCode_(x['Kode Guru']);
  const attendanceMap = new Map(attendance.map(x => [keyOf(x), x]));
  const journalMap = new Map(journals.map(x => [keyOf(x), x]));
  const teachersRaw = readObjects_('GURU');
  const teacherNames = new Map(teachersRaw.map(t => [normalizeCode_(t['Kode Guru']), String(t['Nama Guru'] || '')]));
  const groups = new Map();
  schedules.forEach(sc => {
    const code = normalizeCode_(sc['Kode Guru']);
    if (!groups.has(code)) groups.set(code, {nama: String(sc['Nama Guru'] || teacherNames.get(code) || code), jadwal:0, tercatat:0, hadir:0, izin:0, sakit:0, tidakHadir:0, belumAbsen:0, jurnal:0, belumJurnal:0});
    const g = groups.get(code); g.jadwal++;
    const k = String(sc['ID Jadwal'] || '').trim() + '|' + code;
    const a = attendanceMap.get(k), j = journalMap.get(k);
    if (a) {
      g.tercatat++;
      const st = String(a['Status'] || '').trim().toLowerCase();
      if (st === 'hadir') g.hadir++;
      else if (st === 'izin') g.izin++;
      else if (st === 'sakit') g.sakit++;
      else if (st === 'tidak hadir') g.tidakHadir++;
    } else g.belumAbsen++;
    if (j && String(j['Materi Pembelajaran'] || '').trim()) g.jurnal++; else g.belumJurnal++;
  });
  const teachers = Array.from(groups.values()).sort((a,b) => a.nama.localeCompare(b.nama,'id'));
  const sum = field => teachers.reduce((n,t) => n + t[field], 0);
  const summary = {jadwal:sum('jadwal'),tercatat:sum('tercatat'),hadir:sum('hadir'),izin:sum('izin'),sakit:sum('sakit'),tidakHadir:sum('tidakHadir'),belumAbsen:sum('belumAbsen'),jurnal:sum('jurnal'),belumJurnal:sum('belumJurnal')};
  summary.persenHadir = summary.jadwal ? Math.round(summary.hadir / summary.jadwal * 100) : 0;
  const trend = [];
  for (let offset=6; offset>=0; offset--) {
    const d = new Date(selected); d.setUTCDate(d.getUTCDate()-offset);
    const date = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
    const weekday = days[d.getUTCDay()];
    const daySchedules = readObjects_('JADWAL').filter(x => isActive_(x['Aktif']) && normalizeDay_(x['Hari']) === normalizeDay_(weekday));
    const dayAttendance = attendanceAll.filter(x => normalizeDate_(x['Tanggal']) === date);
    trend.push({tanggal:date,label:Utilities.formatDate(d,'UTC','dd/MM'),jadwal:daySchedules.length,hadir:dayAttendance.filter(x=>String(x['Status']||'').trim().toLowerCase()==='hadir').length});
  }
  return {tanggal:key,hari:dayName,summary,teachers,trend};
}

/* ===== PENCARIAN RIWAYAT ABSENSI GURU V1.16 ===== */
function searchTeacherAttendanceHistory(token, filters) {
  const user = getPortalSession_(token);
  if (String(user.peran || '').toUpperCase() !== 'ADMIN') {
    throw new Error('Pencarian riwayat absensi seluruh guru hanya dapat diakses ADMIN.');
  }
  filters = filters || {};
  const today = Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
  const start = String(filters.startDate || today).trim();
  const end = String(filters.endDate || start).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end) || start > end) {
    throw new Error('Rentang tanggal tidak valid. Pastikan tanggal awal tidak melewati tanggal akhir.');
  }
  const teacherCode = String(filters.teacherCode || '').trim();
  const dayFilter = normalizeDay_(filters.day || '');
  const statusFilter = String(filters.status || '').trim().toLowerCase();

  const teachers = readObjects_('GURU')
    .filter(t => isActive_(t['Aktif']))
    .map(t => ({ kode: String(t['Kode Guru'] || ''), nama: String(t['Nama Guru'] || '') }))
    .filter(t => t.kode && t.nama)
    .sort((a,b) => a.nama.localeCompare(b.nama, 'id'));

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ABSENSI');
  if (!sh || sh.getLastRow() < 2) return { teachers, rows: [], total: 0, start, end };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const idx = key => headers.indexOf(key);
  ['ID Jadwal','Tanggal','Hari','Waktu Presensi','Kode Guru','Nama Guru','Kelas','Jam Ke','Mulai','Selesai','Mata Pelajaran','Status']
    .forEach(key => { if (idx(key) < 0) throw new Error('Kolom ' + key + ' tidak ditemukan pada sheet ABSENSI.'); });

  const scheduleSheet = ss.getSheetByName('JADWAL');
  const scheduleMap = {};
  if (scheduleSheet && scheduleSheet.getLastRow() > 1) {
    const display = scheduleSheet.getDataRange().getDisplayValues();
    const h = display[0].map(x => String(x || '').trim());
    const idCol = h.indexOf('ID Jadwal');
    const startCol = h.indexOf('Jam Mulai') >= 0 ? h.indexOf('Jam Mulai') : h.indexOf('Mulai');
    const endCol = h.indexOf('Jam Selesai') >= 0 ? h.indexOf('Jam Selesai') : h.indexOf('Selesai');
    if (idCol >= 0) display.slice(1).forEach(r => {
      const id = String(r[idCol] || '').trim();
      if (id) scheduleMap[id] = { mulai: startCol >= 0 ? r[startCol] : '', selesai: endCol >= 0 ? r[endCol] : '' };
    });
  }
  const clock = value => {
    if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, APP.timezone, 'HH:mm:ss');
    if (typeof value === 'number' && isFinite(value)) {
      const mins = Math.round((value % 1) * 1440);
      return String(Math.floor(mins / 60) % 24).padStart(2,'0') + ':' + String(mins % 60).padStart(2,'0');
    }
    const text = String(value == null ? '' : value).trim();
    const m = text.match(/(?:T|\s)(\d{1,2}:\d{2}(?::\d{2})?)/) || text.match(/^(\d{1,2}:\d{2}(?::\d{2})?)$/);
    if (!m) return text;
    const parts = m[1].split(':');
    return String(Number(parts[0])).padStart(2,'0') + ':' + parts[1] + (parts[2] ? ':' + parts[2] : '');
  };
  const rows = values.slice(1).map(r => {
    const obj = {}; headers.forEach((h,i) => { if (h) obj[h] = r[i]; }); return obj;
  }).filter(a => {
    const date = normalizeDate_(a['Tanggal']);
    const code = String(a['Kode Guru'] || '').trim();
    const day = String(a['Hari'] || '');
    return date >= start && date <= end && (!teacherCode || code === teacherCode) &&
      (!dayFilter || normalizeDay_(day) === dayFilter) &&
      (!statusFilter || String(a['Status'] || '').trim().toLowerCase() === statusFilter);
  }).map(a => {
    const date = normalizeDate_(a['Tanggal']);
    const sched = scheduleMap[String(a['ID Jadwal'] || '').trim()] || {};
    return {
      tanggal: date,
      hari: String(a['Hari'] || ''),
      guru: String(a['Nama Guru'] || ''),
      kodeGuru: String(a['Kode Guru'] || ''),
      kelas: String(a['Kelas'] || ''),
      jamKe: String(a['Jam Ke'] || ''),
      jamPelajaran: (clock(sched.mulai || a['Mulai']) || '--:--') + '–' + (clock(sched.selesai || a['Selesai']) || '--:--'),
      waktuAbsen: clock(a['Waktu Presensi']) || '-',
      mapel: String(a['Mata Pelajaran'] || ''),
      status: String(a['Status'] || ''),
      catatan: String(a['Catatan'] || '')
    };
  }).sort((a,b) => b.tanggal.localeCompare(a.tanggal) || b.waktuAbsen.localeCompare(a.waktuAbsen) || a.guru.localeCompare(b.guru,'id'));

  return { teachers, rows: rows.slice(0, 1000), total: rows.length, start, end, truncated: rows.length > 1000 };
}


/* ===== MODUL KELOLA JADWAL KBM PORTAL V1.17 ===== */
function portalScheduleAdmin_(token){
  const u=getPortalSession_(token);
  if(!['ADMIN','TU'].includes(String(u.peran||'').toUpperCase())) throw new Error('Hanya ADMIN/TU yang dapat mengelola jadwal.');
  return u;
}
function getPortalScheduleData(token){
  const u=getPortalSession_(token); const admin=['ADMIN','TU'].includes(String(u.peran||'').toUpperCase());
  const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName('JADWAL');
  if(!sh||sh.getLastRow()<1) return {schedules:[],teachers:[],admin};
  const values=sh.getDataRange().getDisplayValues(); const headers=values[0].map(x=>String(x).trim());
  const obj=values.slice(1).filter(r=>r.some(x=>x!=='')).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
  const teachers=readObjects_('GURU').filter(t=>isActive_(t['Aktif'])).map(t=>({kode:String(t['Kode Guru']||''),nama:String(t['Nama Guru']||'')}));
  const schedules=obj.filter(x=>isActive_(x['Aktif'])&&(admin||normalizeCode_(x['Kode Guru'])===normalizeCode_(u.kodeGuru||u.kode))).map(x=>({id:String(x['ID Jadwal']||''),hari:String(x['Hari']||''),jamKe:String(x['Jam Ke']||''),mulai:String(x['Jam Mulai']||x['Mulai']||''),selesai:String(x['Jam Selesai']||x['Selesai']||''),kelas:String(x['Kelas']||''),mapel:String(x['Mata Pelajaran']||''),kodeGuru:String(x['Kode Guru']||''),namaGuru:String(x['Nama Guru']||''),aktif:String(x['Aktif']||'YA')}));
  schedules.sort((a,b)=>['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'].indexOf(a.hari)-['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'].indexOf(b.hari)||String(a.mulai).localeCompare(String(b.mulai))||a.kelas.localeCompare(b.kelas,'id'));
  return {schedules,teachers:admin?teachers:[],admin};
}
function savePortalSchedule(token,item){
  portalScheduleAdmin_(token); item=item||{};
  ['hari','jamKe','mulai','selesai','kelas','mapel','kodeGuru'].forEach(k=>{if(!String(item[k]||'').trim())throw new Error('Kolom '+k+' wajib diisi.');});
  const days=['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu']; if(!days.includes(item.hari))throw new Error('Hari tidak valid.');
  const time=/^([01]\d|2[0-3]):[0-5]\d$/; if(!time.test(item.mulai)||!time.test(item.selesai)||item.selesai<=item.mulai)throw new Error('Format jam harus HH:mm dan jam selesai setelah jam mulai.');
  const teacher=readObjects_('GURU').find(t=>normalizeCode_(t['Kode Guru'])===normalizeCode_(item.kodeGuru)&&isActive_(t['Aktif'])); if(!teacher)throw new Error('Guru aktif tidak ditemukan.');
  const sh=SpreadsheetApp.getActive().getSheetByName('JADWAL'); if(!sh)throw new Error('Sheet JADWAL tidak ditemukan.');
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
  const rowMap={'ID Jadwal':item.id||Utilities.getUuid(),'Hari':item.hari,'Jam Ke':item.jamKe,'Mulai':item.mulai,'Jam Mulai':item.mulai,'Selesai':item.selesai,'Jam Selesai':item.selesai,'Kelas':item.kelas,'Mata Pelajaran':item.mapel,'Kode Guru':item.kodeGuru,'Nama Guru':teacher['Nama Guru'],'Aktif':'YA'};
  const row=headers.map(h=>rowMap[h]!==undefined?rowMap[h]:'');
  if(item.id){const vals=sh.getDataRange().getValues();const idCol=headers.indexOf('ID Jadwal');const idx=vals.findIndex((r,i)=>i>0&&String(r[idCol])===String(item.id));if(idx<1)throw new Error('Jadwal yang akan diubah tidak ditemukan.');const existing=vals[idx].slice();headers.forEach((h,i)=>{if(rowMap[h]!==undefined)existing[i]=rowMap[h];});sh.getRange(idx+1,1,1,headers.length).setValues([existing]);}
  else sh.appendRow(row);
  return {success:true,message:item.id?'Jadwal berhasil diperbarui.':'Jadwal berhasil ditambahkan.'};
}
function deactivatePortalSchedule(token,id){portalScheduleAdmin_(token);const sh=SpreadsheetApp.getActive().getSheetByName('JADWAL');if(!sh)throw new Error('Sheet JADWAL tidak ditemukan.');const v=sh.getDataRange().getDisplayValues();const h=v[0].map(x=>String(x).trim());const ic=h.indexOf('ID Jadwal'),ac=h.indexOf('Aktif');if(ic<0||ac<0)throw new Error('Kolom ID Jadwal/Aktif tidak ditemukan.');const i=v.findIndex((r,n)=>n>0&&r[ic]===String(id));if(i<1)throw new Error('Jadwal tidak ditemukan.');sh.getRange(i+1,ac+1).setValue('TIDAK');return {success:true};}


/* ===== NOTIFIKASI JADWAL WHATSAPP GROUP V1.18 =====
 * Karena grup WhatsApp lama dan akun pengirim masih WhatsApp Business App,
 * GAS hanya menyiapkan draft otomatis pukul 06.00 WIB. Admin menyalin dan
 * mengirimkannya sendiri ke grup. Tidak ada pengiriman WhatsApp otomatis.
 */
function getDailyWhatsAppDraft(token) {
  const u = getPortalSession_(token);
  if (String(u.peran || '').toUpperCase() !== 'ADMIN') throw new Error('Hanya ADMIN yang dapat mengakses notifikasi grup.');
  return buildDailyWhatsAppDraft_();
}

function installDailyWhatsAppDraftTrigger(token) {
  const u = getPortalSession_(token);
  if (String(u.peran || '').toUpperCase() !== 'ADMIN') throw new Error('Hanya ADMIN yang dapat mengatur notifikasi.');
  const triggers = ScriptApp.getProjectTriggers();
  triggers.filter(t => t.getHandlerFunction() === 'prepareDailyWhatsAppDraft').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('prepareDailyWhatsAppDraft').timeBased().everyDays(1).atHour(6).inTimezone('Asia/Jakarta').create();
  // Buat draft awal sekarang agar menu langsung bisa diuji.
  prepareDailyWhatsAppDraft();
  return {success:true, message:'Pemicu harian dipasang. Google Apps Script akan menyiapkan draft sekitar pukul 06.00 WIB setiap hari. Admin tetap perlu menyalin dan mengirim pesan ke grup WhatsApp.'};
}

function prepareDailyWhatsAppDraft() {
  const draft = buildDailyWhatsAppDraft_();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('NOTIFIKASI_JADWAL');
  if (!sh) sh = ss.insertSheet('NOTIFIKASI_JADWAL');
  if (sh.getLastRow() === 0) sh.appendRow(['Tanggal','Hari','Disiapkan Pada','Pesan','Jumlah Jadwal']);
  const row = [draft.date, draft.day, new Date(), draft.message, draft.total];
  const last = sh.getLastRow();
  if (last >= 2 && String(sh.getRange(last,1).getDisplayValue()) === draft.date) {
    sh.getRange(last,1,1,row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');
  sh.getRange('C:C').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return {success:true, date:draft.date, total:draft.total};
}

function buildDailyWhatsAppDraft_() {
  const now = new Date();
  const dateKey = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
  const dayNames = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const day = dayNames[Number(Utilities.formatDate(now, 'Asia/Jakarta', 'u')) % 7];
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('JADWAL');
  if (!sh || sh.getLastRow() < 2) throw new Error('Sheet JADWAL belum berisi jadwal.');
  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0].map(x => String(x || '').trim());
  const idx = key => headers.indexOf(key);
  const col = {
    day: idx('Hari'), start: idx('Jam Mulai') >= 0 ? idx('Jam Mulai') : idx('Mulai'),
    end: idx('Jam Selesai') >= 0 ? idx('Jam Selesai') : idx('Selesai'),
    period: idx('Jam Ke'), className: idx('Kelas'), subject: idx('Mata Pelajaran'),
    teacher: idx('Nama Guru'), code: idx('Kode Guru'), active: idx('Aktif')
  };
  if (col.day < 0 || col.start < 0 || col.end < 0 || col.className < 0 || col.subject < 0) {
    throw new Error('Kolom wajib pada sheet JADWAL tidak lengkap.');
  }
  const isActive = v => ['YA','AKTIF','TRUE','1','YES'].includes(String(v || '').trim().toUpperCase());
  const rows = values.slice(1).filter(r => r.some(v => v !== '') && String(r[col.day] || '').trim().toLowerCase() === day.toLowerCase() && (col.active < 0 || isActive(r[col.active])));
  const time = v => {
    const s = String(v || '').trim();
    const m = s.match(/(?:1899-12-30\s*)?(\d{1,2}):(\d{2})/);
    return m ? ('0'+m[1]).slice(-2)+':'+m[2] : s;
  };
  rows.sort((a,b) => time(a[col.start]).localeCompare(time(b[col.start])) || String(a[col.className]).localeCompare(String(b[col.className]),'id'));
  const dateDisplay = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy');
  const lines = [
    '📢 *JADWAL KBM HARI INI*',
    '*PKPPS HIDAYATULLAH KARIMUN*',
    '🗓️ '+day+', '+dateDisplay,
    '',
    rows.length ? '*Jadwal mengajar:*' : 'Tidak ada jadwal KBM aktif untuk hari ini.'
  ];
  rows.forEach((r,i) => {
    const teacher = col.teacher >= 0 ? String(r[col.teacher] || '').trim() : '';
    const period = col.period >= 0 && r[col.period] ? ' · Jam ke-'+r[col.period] : '';
    lines.push((i+1)+'. '+(teacher || 'Guru belum diisi')+' — '+String(r[col.subject]||'')+' — Kelas '+String(r[col.className]||'')+' — '+time(r[col.start])+'–'+time(r[col.end])+period);
  });
  lines.push('', 'Mohon seluruh guru memperhatikan jadwal mengajar masing-masing serta mengisi Absensi KBM dan Jurnal Mengajar setelah pembelajaran.', '', 'Terima kasih.', '*Admin PKPPS Hidayatullah Karimun*');
  return {date:dateKey, day, total:rows.length, message:lines.join('\n')};
}


/* ===== MODUL KELAS & WALI KELAS ===== */
function classAdmin_(token) {
  const u = getPortalSession_(token);
  if (String(u.peran || '').toUpperCase() !== 'ADMIN') throw new Error('Hanya ADMIN yang dapat mengelola data kelas dan wali kelas.');
  return u;
}
function classHeaderMap_(sh) {
  if (!sh || sh.getLastRow() < 1) throw new Error('Sheet KELAS belum tersedia atau belum memiliki header.');
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x||'').trim());
  const map = {}; h.forEach((v,i)=>map[v]=i);
  return {headers:h,map};
}
function getClassManagementData(token) {
  classAdmin_(token);
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('KELAS');
  if(!sh) throw new Error('Sheet KELAS tidak ditemukan. Pastikan data kelas sudah dibuat.');
  const hm=classHeaderMap_(sh), rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues():[];
  const col=k=>hm.map[k]===undefined?-1:hm.map[k];
  ['ID Kelas','Nama Kelas','Jenjang'].forEach(k=>{if(col(k)<0)throw new Error('Kolom '+k+' tidak ditemukan pada sheet KELAS.');});
  const studentSheet=ss.getSheetByName('SANTRI')||ss.getSheetByName('SISWA');
  let studentCounts={};
  if(studentSheet && studentSheet.getLastRow()>1){
    const sv=studentSheet.getDataRange().getDisplayValues(), shd=sv[0].map(x=>String(x||'').trim());
    let ci=shd.indexOf('Kelas'); if(ci<0)ci=shd.indexOf('Nama Kelas'); if(ci<0)ci=8;
    let si=shd.indexOf('Status');
    sv.slice(1).forEach(r=>{const k=String(r[ci]||'').trim();if(k){if(si<0||!['NONAKTIF','KELUAR','LULUS'].includes(String(r[si]||'').trim().toUpperCase()))studentCounts[k]=(studentCounts[k]||0)+1;}});
  }
  const guruSh=ss.getSheetByName('GURU'); let teachers=[];
  if(guruSh&&guruSh.getLastRow()>1){
    const gv=guruSh.getDataRange().getDisplayValues(), gh=gv[0].map(x=>String(x||'').trim()), ix=k=>gh.indexOf(k);
    const gc=ix('Kode Guru'), gn=ix('Nama Guru'), ga=ix('Aktif');
    if(gc>=0&&gn>=0) teachers=gv.slice(1).filter(r=>r[gc]&&r[gn]&&(ga<0||['YA','AKTIF','TRUE','1'].includes(String(r[ga]).trim().toUpperCase()))).map(r=>({kode:String(r[gc]).trim(),nama:String(r[gn]).trim()}));
  }
  const classes=rows.map(r=>({id:r[col('ID Kelas')],nama:r[col('Nama Kelas')],jenjang:r[col('Jenjang')],kodeWali:col('Kode Wali Kelas')>=0?r[col('Kode Wali Kelas')]: '',namaWali:col('Nama Wali Kelas')>=0?r[col('Nama Wali Kelas')]: (col('Wali Kelas')>=0?r[col('Wali Kelas')]:''),status:col('Status')>=0?r[col('Status')]: (col('Aktif')>=0?r[col('Aktif')]: 'Aktif'),jumlahSantri:studentCounts[String(r[col('Nama Kelas')]||'').trim()]||studentCounts[String(r[col('ID Kelas')]||'').trim()]||0}));
  return {classes,teachers,studentSheet:studentSheet?studentSheet.getName():'Belum tersedia'};
}
function saveClassWali(token, idKelas, kodeGuru) {
  classAdmin_(token);
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('KELAS');
  if(!sh)throw new Error('Sheet KELAS tidak ditemukan.');
  const hm=classHeaderMap_(sh), m=hm.map;
  ['ID Kelas','Kode Wali Kelas'].forEach(k=>{if(m[k]===undefined)throw new Error('Kolom '+k+' tidak ditemukan.');});
  const values=sh.getDataRange().getDisplayValues();
  const ri=values.findIndex((r,i)=>i>0&&String(r[m['ID Kelas']]).trim()===String(idKelas).trim());
  if(ri<1)throw new Error('Kelas tidak ditemukan.');
  const guru=ss.getSheetByName('GURU'); if(!guru)throw new Error('Sheet GURU tidak ditemukan.');
  const gv=guru.getDataRange().getDisplayValues(), gh=gv[0].map(x=>String(x||'').trim()), gc=gh.indexOf('Kode Guru'), gn=gh.indexOf('Nama Guru'), ga=gh.indexOf('Aktif');
  if(gc<0||gn<0)throw new Error('Kolom Kode Guru/Nama Guru tidak ditemukan pada GURU.');
  const gr=gv.slice(1).find(r=>String(r[gc]).trim()===String(kodeGuru).trim()&&(ga<0||['YA','AKTIF','TRUE','1'].includes(String(r[ga]).trim().toUpperCase())));
  if(!gr)throw new Error('Guru tidak ditemukan atau tidak aktif.');
  sh.getRange(ri+1,m['Kode Wali Kelas']+1).setValue(String(gr[gc]).trim());
  const nameCol=m['Nama Wali Kelas']!==undefined?m['Nama Wali Kelas']:(m['Wali Kelas']!==undefined?m['Wali Kelas']:-1);
  if(nameCol>=0)sh.getRange(ri+1,nameCol+1).setValue(String(gr[gn]).trim());
  return {success:true,message:'Wali kelas '+gr[gn]+' berhasil ditetapkan untuk '+values[ri][m['Nama Kelas']]+'.'};
}
function addPortalClass(token, data) {
  classAdmin_(token); data=data||{};
  const id=String(data.id||'').trim(), nama=String(data.nama||'').trim(), jenjang=String(data.jenjang||'').trim();
  if(!id||!nama||!jenjang)throw new Error('ID kelas, nama kelas, dan jenjang wajib diisi.');
  const sh=SpreadsheetApp.getActive().getSheetByName('KELAS'); if(!sh)throw new Error('Sheet KELAS tidak ditemukan.');
  const hm=classHeaderMap_(sh), m=hm.map;
  ['ID Kelas','Nama Kelas','Jenjang'].forEach(k=>{if(m[k]===undefined)throw new Error('Kolom '+k+' tidak ditemukan.');});
  const vals=sh.getDataRange().getDisplayValues();
  if(vals.slice(1).some(r=>String(r[m['ID Kelas']]).trim().toLowerCase()===id.toLowerCase()||String(r[m['Nama Kelas']]).trim().toLowerCase()===nama.toLowerCase()))throw new Error('ID atau nama kelas sudah terdaftar.');
  const row=hm.headers.map(h=>h==='ID Kelas'?id:h==='Nama Kelas'?nama:h==='Jenjang'?jenjang:(h==='Status'||h==='Aktif')?'Aktif':'');
  sh.appendRow(row); return {success:true,message:'Kelas '+nama+' berhasil ditambahkan.'};
}


/* ===== MODUL NILAI & RAPOR ===== */
const GRADE_HEADERS_ = ['ID Nilai','Tanggal Input','NIS','Nama Santri','Kelas','ID Mapel','Nama Mapel','Kode Guru','Nama Guru','Semester','Tahun Pelajaran','PH','PTS','PAS','Nilai Akhir','Catatan'];
const GRADE_DETAIL_HEADERS_ = ['ID Detail','NIS','Nama Santri','Kelas','ID Mapel','Nama Mapel','Semester','Tahun Pelajaran','Urutan','Nilai Tugas','Diperbarui Oleh','Diperbarui Pada'];
function gradeDetailSheet_(){const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName('NILAI_HARIAN');if(!sh){sh=ss.insertSheet('NILAI_HARIAN');sh.appendRow(GRADE_DETAIL_HEADERS_);sh.setFrozenRows(1);sh.getRange(1,1,1,GRADE_DETAIL_HEADERS_.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#fff');}return sh;}
function gradeWeightedFinal_(t,pts,pas){if(t===null||pts===null||pas===null)return '';return Math.round((t*.30+pts*.30+pas*.40)*100)/100;}
function gradeSession_(token){
  const u=getPortalSession_(token); const role=String(u.peran||'').toUpperCase();
  if(!['ADMIN','GURU','WALI','TU'].includes(role)) throw new Error('Akses modul nilai tidak tersedia untuk akun ini.');
  return u;
}
function gradeSheet_(){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName('NILAI');
  if(!sh){sh=ss.insertSheet('NILAI'); sh.appendRow(GRADE_HEADERS_); sh.setFrozenRows(1); sh.getRange(1,1,1,GRADE_HEADERS_.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');}
  return sh;
}
function gradeFindCol_(headers, names){for(const n of names){const i=headers.indexOf(n);if(i>=0)return i;}return -1;}
function gradeStudents_(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('SANTRI')||ss.getSheetByName('SISWA');
  if(!sh||sh.getLastRow()<2)return [];
  const v=sh.getDataRange().getDisplayValues(), h=v[0].map(x=>String(x||'').trim());
  const ix=n=>gradeFindCol_(h,n);
  const iNis=ix(['NIS','NISN','ID Santri','ID Siswa']), iName=ix(['Nama Santri','Nama Siswa','Nama','Nama Lengkap']), iClass=ix(['Kelas','Nama Kelas']), iStatus=ix(['Status']);
  if(iName<0||iClass<0)throw new Error('Kolom Nama Santri dan Kelas tidak ditemukan pada sheet '+sh.getName()+'.');
  return v.slice(1).filter(r=>r[iName]&&r[iClass]&&!(iStatus>=0&&['NONAKTIF','KELUAR','LULUS'].includes(String(r[iStatus]).trim().toUpperCase()))).map((r,i)=>({nis:iNis>=0?String(r[iNis]||'').trim():String(i+1),nama:String(r[iName]).trim(),kelas:String(r[iClass]).trim()}));
}
function gradeClassKey_(value){
  return String(value == null ? '' : value).trim().toUpperCase()
    .replace(/^KELAS\s+/, '')
    .replace(/[\s._-]+/g, '');
}
function getGradeSetup(token){
  const u=gradeSession_(token), ss=SpreadsheetApp.getActive();
  // Login Portal menyimpan kode guru pada properti `kodeGuru`, bukan `kode`.
  // `kode` tetap didukung untuk sesi lama agar kompatibel.
  const userCode=normalizeCode_(u.kodeGuru || u.kode || '');
  const waliCode=normalizeCode_(u.kodeGuru || u.kode || u.idWali || '');
  const classes=readObjects_('KELAS').map(x=>({
    id:String(x['ID Kelas']||'').trim(),
    nama:String(x['Nama Kelas']||'').trim(),
    jenjang:String(x['Jenjang']||'').trim(),
    kodeWali:String(x['Kode Wali Kelas']||'').trim(),
    namaWali:String(x['Nama Wali Kelas']||'').trim()
  })).filter(x=>x.nama || x.id);
  const students=gradeStudents_();
  const maps=readObjects_('MAPEL').map(x=>({
    id:String(x['ID Mapel']||'').trim(),
    nama:String(x['Nama Mapel Indonesia']||x['Nama Mapel']||'').trim(),
    arab:String(x['Nama Mapel Arab']||'').trim()
  })).filter(x=>x.id||x.nama);
  const assignments=readObjects_('GURU_MAPEL').map(x=>({
    kode:normalizeCode_(x['Kode Guru']),
    namaGuru:String(x['Nama Guru']||'').trim(),
    idMapel:String(x['ID Mapel']||'').trim(),
    namaMapel:String(x['Nama Mapel']||'').trim(),
    kelas:String(x['Kelas']||'').trim()
  })).filter(x=>x.kode || x.idMapel || x.namaMapel || x.kelas);
  let allowedClasses=classes, allowedAssignments=assignments;
  if(String(u.peran||'').toUpperCase()==='GURU'){
    allowedAssignments=assignments.filter(a=>a.kode===userCode);
    const keys=new Set(allowedAssignments.map(a=>gradeClassKey_(a.kelas)).filter(Boolean));
    allowedClasses=classes.filter(c=>
      keys.has(gradeClassKey_(c.nama)) || keys.has(gradeClassKey_(c.id))
    );
  } else if(String(u.peran||'').toUpperCase()==='WALI'){
    allowedClasses=classes.filter(c=>normalizeCode_(c.kodeWali)===waliCode);
    allowedAssignments=assignments.filter(a=>
      allowedClasses.some(c=>
        gradeClassKey_(c.nama)===gradeClassKey_(a.kelas) ||
        gradeClassKey_(c.id)===gradeClassKey_(a.kelas)
      )
    );
  }
  return {
    user:Object.assign({},u,{kodeGuru:userCode}),
    classes:allowedClasses,
    students:students.filter(st=>allowedClasses.some(c=>
      gradeClassKey_(c.nama)===gradeClassKey_(st.kelas) ||
      gradeClassKey_(c.id)===gradeClassKey_(st.kelas)
    )),
    subjects:maps,
    assignments:allowedAssignments,
    hasNilai:!!ss.getSheetByName('NILAI')
  };
}
function getGradeRoster(token, kelas, idMapel, semester){
  const u=gradeSession_(token), setup=getGradeSetup(token); kelas=String(kelas||''); idMapel=String(idMapel||''); semester=String(semester||'1');
  const c=setup.classes.find(x=>x.nama===kelas||x.id===kelas); if(!c)throw new Error('Kelas tidak ditemukan atau bukan wewenang akun ini.');
  if(u.peran==='GURU'&&!setup.assignments.some(a=>(a.kelas===c.nama||a.kelas===c.id)&&a.idMapel===idMapel))throw new Error('Anda tidak ditugaskan mengajar mapel ini di kelas tersebut.');
  const a=setup.assignments.find(x=>(x.kelas===c.nama||x.kelas===c.id)&&x.idMapel===idMapel);
  const subject=setup.subjects.find(x=>x.id===idMapel);
  if(!subject&&!a)throw new Error('Mata pelajaran tidak ditemukan pada penugasan kelas.');
  const students=setup.students.filter(s=>s.kelas===c.nama||s.kelas===c.id);
  const grades=readObjects_('NILAI').filter(x=>String(x['Kelas'])===c.nama&&String(x['ID Mapel'])===idMapel&&String(x['Semester'])===semester&&String(x['Tahun Pelajaran']||APP.year)===APP.year);
  const byNis={}; grades.forEach(g=>byNis[String(g['NIS'])]=g);
  const detailSh=gradeDetailSheet_(), detailRows=detailSh.getDataRange().getValues(), dh=detailRows[0].map(String);
  const dcol=n=>dh.indexOf(n), dailyByNis={};
  detailRows.slice(1).forEach(r=>{if(String(r[dcol('Kelas')])===c.nama&&String(r[dcol('ID Mapel')])===idMapel&&String(r[dcol('Semester')])===semester&&String(r[dcol('Tahun Pelajaran')])===APP.year){const key=String(r[dcol('NIS')]);(dailyByNis[key]||(dailyByNis[key]=[])).push({urut:Number(r[dcol('Urutan')])||1,nilai:r[dcol('Nilai Tugas')]});}});
  return {kelas:c.nama,subject:subject||{id:idMapel,nama:a.namaMapel},teacher:a||{kode:u.kode,namaGuru:u.nama},students:students.map(s=>{const g=byNis[s.nis]||{};let daily=(dailyByNis[s.nis]||[]).sort((x,y)=>x.urut-y.urut).map(x=>x.nilai);if(!daily.length&&g['PH']!==''&&g['PH']!==undefined)daily=[g['PH']];return {...s,ph:g['PH']??'',tugasList:daily.join(', '),tugasRata:daily.length?Math.round(daily.reduce((z,n)=>z+Number(n),0)/daily.length*100)/100:'',pts:g['PTS']??'',pas:g['PAS']??'',nilaiAkhir:g['Nilai Akhir']??'',catatan:g['Catatan']??''};})};
}
function saveGradeBatch(token, payload){
  const u=gradeSession_(token); payload=payload||{};
  if(String(u.peran||'').toUpperCase()==='TU') throw new Error('TU hanya dapat melihat, merekap, dan mencetak nilai/rapor. TU tidak dapat mengubah nilai guru.');
  const kelas=String(payload.kelas||''), idMapel=String(payload.idMapel||''), semester=String(payload.semester||'1');
  if(!['1','2'].includes(semester))throw new Error('Semester harus 1 atau 2.');
  const roster=getGradeRoster(token,kelas,idMapel,semester), sh=gradeSheet_(), detailSh=gradeDetailSheet_();
  const vals=sh.getDataRange().getValues(), headers=vals[0].map(x=>String(x||'').trim()), idx=n=>headers.indexOf(n);
  const input=Array.isArray(payload.nilai)?payload.nilai:[]; let saved=0;
  input.forEach(item=>{
    const student=roster.students.find(s=>String(s.nis)===String(item.nis)); if(!student)return;
    const dailyRaw=String(item.tugas||'').trim();
    const daily=dailyRaw?dailyRaw.split(/[,;\n]+/).map(x=>Number(String(x).trim())).filter(x=>String(x).trim()!==''):[];
    const pts=item.pts===''||item.pts===null||item.pts===undefined?null:Number(item.pts), pas=item.pas===''||item.pas===null||item.pas===undefined?null:Number(item.pas);
    if(daily.some(n=>!Number.isFinite(n)||n<0||n>100)||[pts,pas].some(n=>n!==null&&(!Number.isFinite(n)||n<0||n>100)))throw new Error('Nilai harus berupa angka 0 sampai 100. Pisahkan nilai tugas dengan koma.');
    const tugas=daily.length?Math.round(daily.reduce((a,b)=>a+b,0)/daily.length*100)/100:'';
    const final=gradeWeightedFinal_(daily.length?tugas:null,pts,pas);
    // Simpan rincian nilai harian sebagai satu baris per tugas, tanpa menghapus data santri/mapel lain.
    const dvals=detailSh.getDataRange().getValues(), dh=dvals[0].map(String), dc=n=>dh.indexOf(n);
    for(let j=dvals.length-1;j>=1;j--){const r=dvals[j];if(String(r[dc('NIS')])===String(student.nis)&&String(r[dc('Kelas')])===roster.kelas&&String(r[dc('ID Mapel')])===idMapel&&String(r[dc('Semester')])===semester&&String(r[dc('Tahun Pelajaran')])===APP.year)detailSh.deleteRow(j+1);}
    if(daily.length){const rows=daily.map((n,i)=>[Utilities.getUuid(),student.nis,student.nama,roster.kelas,idMapel,roster.subject.nama,semester,APP.year,i+1,n,u.nama,new Date()]);detailSh.getRange(detailSh.getLastRow()+1,1,rows.length,GRADE_DETAIL_HEADERS_.length).setValues(rows);}
    const existing=vals.findIndex((r,i)=>i>0&&String(r[idx('NIS')])===String(student.nis)&&String(r[idx('Kelas')])===roster.kelas&&String(r[idx('ID Mapel')])===idMapel&&String(r[idx('Semester')])===semester&&String(r[idx('Tahun Pelajaran')])===APP.year);
    const obj={'ID Nilai':existing>0?vals[existing][idx('ID Nilai')]:Utilities.getUuid(),'Tanggal Input':new Date(),'NIS':student.nis,'Nama Santri':student.nama,'Kelas':roster.kelas,'ID Mapel':idMapel,'Nama Mapel':roster.subject.nama,'Kode Guru':u.kode,'Nama Guru':u.nama,'Semester':semester,'Tahun Pelajaran':APP.year,'PH':tugas,'PTS':pts??'','PAS':pas??'','Nilai Akhir':final,'Catatan':String(item.catatan||'')};
    const row=headers.map(h=>obj[h]!==undefined?obj[h]:''); if(existing>0)sh.getRange(existing+1,1,1,row.length).setValues([row]);else sh.appendRow(row); saved++;
  });
  return {success:true,saved,message:saved+' nilai berhasil disimpan. Nilai akhir memakai bobot Tugas 30%, PTS 30%, PAS 40%; nilai akhir kosong jika ada komponen belum lengkap.'};
}
function getReportData(token, kelas, semester){
  const u=gradeSession_(token);
  if(String(u.peran||'').toUpperCase()==='GURU') throw new Error('Akses rapor untuk akun Guru dinonaktifkan. Guru hanya dapat menginput nilai; Leger dan Rapor dikelola Admin/TU.');
  const setup=getGradeSetup(token); kelas=String(kelas||''); semester=String(semester||'1');
  const c=setup.classes.find(x=>x.nama===kelas||x.id===kelas); if(!c)throw new Error('Kelas tidak ditemukan atau bukan wewenang akun ini.');
  const students=setup.students.filter(s=>s.kelas===c.nama||s.kelas===c.id);
  const assignments=setup.assignments.filter(a=>a.kelas===c.nama||a.kelas===c.id);
  const grades=readObjects_('NILAI').filter(g=>String(g['Kelas'])===c.nama&&String(g['Semester'])===semester&&String(g['Tahun Pelajaran']||APP.year)===APP.year);
  const subjects=[...new Map(assignments.map(a=>{const m=setup.subjects.find(z=>z.id===a.idMapel)||{};return [a.idMapel,{id:a.idMapel,nama:m.nama||a.namaMapel||'',arab:m.arab||''}]})).values()];
  const out=students.map(s=>{
    const sub=subjects.map(m=>{const g=grades.find(x=>String(x['NIS'])===String(s.nis)&&String(x['ID Mapel'])===m.id);const t=g?g['PH']:'';const p=g?g['PTS']:'';const a=g?g['PAS']:'';const weighted=(t!==''&&p!==''&&a!=='')?gradeWeightedFinal_(Number(t),Number(p),Number(a)):'';return {id:m.id,nama:m.nama,arab:m.arab,ph:t,pts:p,pas:a,akhir:weighted ,catatan:g?g['Catatan']||'':''};});
    // Hanya mata pelajaran dengan nilai akhir yang benar-benar terisi dihitung.
    // Hindari Number('') yang menghasilkan 0 dan membuat semua mapel kosong
    // keliru terhitung sebagai mapel bernilai.
    const nums=sub
      .map(x=>x.akhir)
      .filter(v=>v!==''&&v!==null&&v!==undefined&&String(v).trim()!==''&&Number.isFinite(Number(v)))
      .map(Number);
    const total=nums.reduce((a,b)=>a+b,0), average=nums.length?Math.round(total/nums.length*100)/100:'';
    return {...s,subjects:sub,total:nums.length?Math.round(total*100)/100:'',average,filledSubjects:nums.length};
  });
  const ranked=out.filter(s=>s.average!=='').sort((a,b)=>Number(b.average)-Number(a.average));
  ranked.forEach((s,i)=>{s.peringkat=i+1;}); out.forEach(s=>{if(s.peringkat===undefined)s.peringkat='';});
  return {kelas:c,semester,tahun:APP.year,wali:c.namaWali||'',students:out,subjects};
}


/* ===== MODUL PPDB PKPPS HIDAYATULLAH KARIMUN · TA 2027/2028 ===== */
const PPDB_YEAR_ = '2027/2028';
const PPDB_HEADERS_ = ['ID','Nomor Pendaftaran','Tahun Pelajaran','Tanggal Daftar','Sumber','Nama Lengkap','NIK','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Jenjang','Asal Sekolah','Nama Ayah','Nama Ibu','WhatsApp Orang Tua','Alamat','Status','Catatan Seleksi','Dibuat Oleh','Dibuat Pada','Token Upload Hash'];
function ppdbSheets_(){
 const ss=SpreadsheetApp.getActive();
 let sh=ss.getSheetByName('PPDB');
 if(!sh){sh=ss.insertSheet('PPDB');sh.appendRow(PPDB_HEADERS_);sh.setFrozenRows(1);}
 else { const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim()); if(!h.includes('Token Upload Hash')) sh.getRange(1,sh.getLastColumn()+1).setValue('Token Upload Hash'); }
 let docs=ss.getSheetByName('PPDB_BERKAS');
 if(!docs){docs=ss.insertSheet('PPDB_BERKAS');docs.appendRow(['ID Berkas','ID PPDB','Jenis Berkas','Nama File','URL File','Diupload Pada']);docs.setFrozenRows(1);}
 return {sh,docs};
}
function ppdbHashToken_(token){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(token||'')));}
function ppdbNextNumber_(){
 const sh=ppdbSheets_().sh;
 const vals=sh.getLastRow()>1?sh.getRange(2,2,sh.getLastRow()-1,1).getDisplayValues().flat():[];
 let max=0; vals.forEach(v=>{const m=String(v).match(/^PPDB-2027-(\d{4,})$/);if(m)max=Math.max(max,Number(m[1]));});
 return 'PPDB-2027-'+String(max+1).padStart(4,'0');
}
function ppdbSave_(d,source,actor){
 d=d||{};const nama=String(d.nama||'').trim(),jenjang=String(d.jenjang||'').trim();
 if(!nama)throw new Error('Nama calon santri wajib diisi.');
 if(!['Wustho','Ulya'].includes(jenjang))throw new Error('Jenjang harus Wustho atau Ulya.');
 const nik=String(d.nik||'').replace(/\D/g,'');
 if(nik && nik.length!==16)throw new Error('NIK harus 16 digit.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const {sh}=ppdbSheets_(),id='PPDB-'+Utilities.getUuid().slice(0,10).toUpperCase(),nomor=ppdbNextNumber_();
  const uploadToken=Utilities.getUuid()+Utilities.getUuid();
  const obj={'ID':id,'Nomor Pendaftaran':nomor,'Tahun Pelajaran':PPDB_YEAR_,'Tanggal Daftar':Utilities.formatDate(new Date(),APP.timezone,'yyyy-MM-dd'),'Sumber':source,'Nama Lengkap':nama,'NIK':nik,'Jenis Kelamin':String(d.gender||''),'Tempat Lahir':String(d.tempatLahir||''),'Tanggal Lahir':String(d.tanggalLahir||''),'Jenjang':jenjang,'Asal Sekolah':String(d.asalSekolah||''),'Nama Ayah':String(d.ayah||''),'Nama Ibu':String(d.ibu||''),'WhatsApp Orang Tua':String(d.wa||''),'Alamat':String(d.alamat||''),'Status':'MENUNGGU VERIFIKASI','Catatan Seleksi':'','Dibuat Oleh':actor||source,'Dibuat Pada':new Date(),'Token Upload Hash':ppdbHashToken_(uploadToken)};
  const heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
  sh.appendRow(heads.map(h=>obj[h]===undefined?'':obj[h]));
  return {id,nomor,status:obj.Status,tahun:PPDB_YEAR_,uploadToken:source==='ONLINE'?uploadToken:''};
 }finally{lock.releaseLock();}
}
function submitPpdbPublic(data){return ppdbSave_(data,'ONLINE','Pendaftar Online');}
function createPpdbAdminEntry(token,data){const u=getPortalSession_(token);if(!['ADMIN','TU'].includes(String(u.peran||'').toUpperCase()))throw new Error('Hanya ADMIN yang dapat menginput pendaftaran.');return ppdbSave_(data,'ADMIN/TU',u.nama);}
function getPpdbAdminData(token){
 const u=getPortalSession_(token);if(!['ADMIN','TU'].includes(String(u.peran||'').toUpperCase()))throw new Error('Akses PPDB hanya untuk ADMIN.');
 const {sh}=ppdbSheets_(),rows=sh.getLastRow()>1?sh.getDataRange().getDisplayValues():[PPDB_HEADERS_],h=rows[0].map(x=>String(x).trim()),ix=k=>h.indexOf(k);
 const data=rows.slice(1).filter(r=>r[ix('ID')]).map(r=>({id:r[ix('ID')],nomor:r[ix('Nomor Pendaftaran')],nama:r[ix('Nama Lengkap')],nik:r[ix('NIK')],jenjang:r[ix('Jenjang')],asal:r[ix('Asal Sekolah')],tanggal:r[ix('Tanggal Daftar')],status:r[ix('Status')]}));
 const count=s=>data.filter(x=>x.status===s).length;
 return {rows:data.reverse(),summary:{total:data.length,menunggu:count('MENUNGGU VERIFIKASI'),seleksi:count('DALAM SELEKSI'),diterima:count('DITERIMA'),daftarUlang:count('DAFTAR ULANG')}};
}
function ppdbFind_(id){
 const {sh}=ppdbSheets_();if(sh.getLastRow()<2)throw new Error('Data PPDB belum tersedia.');
 const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ix=k=>h.indexOf(k),i=v.findIndex((r,n)=>n>0&&r[ix('ID')]===String(id));
 if(i<1)throw new Error('Pendaftar tidak ditemukan.');const r=v[i],o={};h.forEach((k,n)=>o[k]=r[n]);return {sh,row:i+1,obj:o};
}
function getPpdbDetail(token,id){
 const u=getPortalSession_(token);if(!['ADMIN','TU'].includes(String(u.peran||'').toUpperCase()))throw new Error('Akses ditolak.');const o=ppdbFind_(id).obj;
 return {id:o.ID,nomor:o['Nomor Pendaftaran'],nama:o['Nama Lengkap'],nik:o.NIK,gender:o['Jenis Kelamin'],tempatLahir:o['Tempat Lahir'],tanggalLahir:o['Tanggal Lahir'],jenjang:o.Jenjang,asal:o['Asal Sekolah'],ayah:o['Nama Ayah'],ibu:o['Nama Ibu'],wa:o['WhatsApp Orang Tua'],alamat:o.Alamat,status:o.Status,catatan:o['Catatan Seleksi'],sumber:o.Sumber,tanggal:o['Tanggal Daftar']};
}
function updatePpdbStatus(token,id,status,catatan){
 const u=getPortalSession_(token);if(!['ADMIN','TU'].includes(String(u.peran||'').toUpperCase()))throw new Error('Hanya ADMIN yang dapat mengubah status PPDB.');
 const allowed=['MENUNGGU VERIFIKASI','BERKAS LENGKAP','DALAM SELEKSI','DITERIMA','TIDAK DITERIMA','DAFTAR ULANG'];if(!allowed.includes(String(status)))throw new Error('Status PPDB tidak valid.');
 const {sh,row}=ppdbFind_(id),h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
 sh.getRange(row,h.indexOf('Status')+1).setValue(status);sh.getRange(row,h.indexOf('Catatan Seleksi')+1).setValue(String(catatan||''));return 'Status pendaftaran berhasil diperbarui.';
}

/* Upload berkas: akses publik hanya dengan token acak yang diterima browser pendaftar saat registrasi. */
const PPDB_UPLOAD_MAX_BYTES_=2*1024*1024;
const PPDB_UPLOAD_MIMES_={'application/pdf':'.pdf','image/jpeg':'.jpg','image/png':'.png'};
function ppdbPrivateFolder_(){const props=PropertiesService.getScriptProperties();let id=props.getProperty('PPDB_PRIVATE_FOLDER_ID');if(id){try{return DriveApp.getFolderById(id);}catch(e){}}const folder=DriveApp.createFolder('PPDB Berkas Calon Santri 2027-2028');props.setProperty('PPDB_PRIVATE_FOLDER_ID',folder.getId());return folder;}
function uploadPpdbDocument(data){
 data=data||{};const id=String(data.id||'').trim(),token=String(data.uploadToken||''),jenis=String(data.jenis||'').trim().toUpperCase(),mime=String(data.mimeType||'').trim().toLowerCase(),namaFile=String(data.namaFile||'berkas').replace(/[\\/:*?"<>|]/g,'_'),base64=String(data.base64||'').replace(/^data:[^;]+;base64,/,'');
 if(!id||!token||!base64)throw new Error('Data pendaftaran, token, atau file tidak lengkap.');if(!['KK','AKTA','IJAZAH','FOTO'].includes(jenis))throw new Error('Jenis berkas tidak dikenal.');if(!Object.prototype.hasOwnProperty.call(PPDB_UPLOAD_MIMES_,mime))throw new Error('Format berkas harus PDF, JPG, atau PNG.');if(jenis==='FOTO'&&mime==='application/pdf')throw new Error('Pas foto harus JPG atau PNG.');
 const found=ppdbFind_(id);if(ppdbHashToken_(token)!==String(found.obj['Token Upload Hash']||''))throw new Error('Token upload tidak valid. Gunakan tautan/sesi pendaftaran yang sama.');
 const bytes=Utilities.base64Decode(base64);if(bytes.length>PPDB_UPLOAD_MAX_BYTES_)throw new Error('Ukuran setiap berkas maksimal 2 MB.');if(!bytes.length)throw new Error('File kosong atau gagal dibaca.');
 const lock=LockService.getScriptLock();if(!lock.tryLock(20000))throw new Error('Sistem sedang sibuk. Coba kembali.');
 try{const folder=ppdbPrivateFolder_(),safeName=id+'_'+jenis+'_'+Date.now()+PPDB_UPLOAD_MIMES_[mime],file=folder.createFile(Utilities.newBlob(bytes,mime,safeName));const sh=ppdbSheets_().docs,headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim()),record={'ID Berkas':'BRK-'+Utilities.getUuid().slice(0,10).toUpperCase(),'ID PPDB':id,'Jenis Berkas':jenis,'Nama File':namaFile,'URL File':file.getUrl(),'Diupload Pada':new Date()};sh.appendRow(headers.map(h=>record[h]===undefined?'':record[h]));return {success:true,jenis,message:'Berkas '+jenis+' berhasil diterima.'};}finally{lock.releaseLock();}
}
function getPpdbDocuments(token,id){const u=getPortalSession_(token);if(u.peran!=='ADMIN')throw new Error('Akses ditolak.');ppdbFind_(id);const sh=ppdbSheets_().docs;if(sh.getLastRow()<2)return [];const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ix=k=>h.indexOf(k);return v.slice(1).filter(r=>String(r[ix('ID PPDB')])===String(id)).map(r=>({jenis:r[ix('Jenis Berkas')],nama:r[ix('Nama File')],url:r[ix('URL File')],tanggal:r[ix('Diupload Pada')]}));}


/* ===== KOREKSI ABSENSI ADMIN PORTAL V1.17 ===== */
const ABSENCE_AUDIT_SHEET_ = 'LOG_KOREKSI_ABSENSI';
function setupAuditAbsensi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Spreadsheet database terlebih dahulu.');
  let sh = ss.getSheetByName(ABSENCE_AUDIT_SHEET_);
  const headers = ['Waktu','Username Admin','Nama Admin','Tanggal','ID Jadwal','Kode Guru','Nama Guru','Kelas','Jam Ke','Mata Pelajaran','Status Lama','Status Baru','Catatan Lama','Catatan Baru','Alasan Koreksi'];
  if (!sh) sh = ss.insertSheet(ABSENCE_AUDIT_SHEET_);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(v=>String(v).trim());
    if (headers.some((h,i)=>existing[i]!==h)) throw new Error('Header LOG_KOREKSI_ABSENSI tidak sesuai. Jangan ubah manual; periksa sheet terlebih dahulu.');
  }
  sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');
  return 'Sheet audit koreksi absensi siap. Data absensi lama tidak diubah.';
}
function getAbsensiUntukKoreksiPortal(token, dateKey, query) {
  const user = getPortalSession_(token);
  if (String(user.peran||'').toUpperCase() !== 'ADMIN') throw new Error('Akses khusus ADMIN.');
  dateKey=String(dateKey||'').trim(); query=String(query||'').trim().toLowerCase();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Pilih tanggal yang valid.');
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.sheets.absensi);
  if(!sh || sh.getLastRow()<2) return [];
  const values=sh.getDataRange().getValues(), headers=values[0].map(v=>String(v).trim());
  const ix=k=>headers.indexOf(k);
  ['ID Jadwal','Tanggal','Kode Guru','Nama Guru','Kelas','Jam Ke','Mata Pelajaran','Status','Catatan'].forEach(k=>{if(ix(k)<0)throw new Error('Kolom ABSENSI tidak ditemukan: '+k);});
  const out=[];
  for(let i=1;i<values.length;i++){
    const r=values[i]; if(normalizeDate_(r[ix('Tanggal')])!==dateKey) continue;
    const obj={row:i+1,idJadwal:String(r[ix('ID Jadwal')]||''),tanggal:dateKey,kodeGuru:String(r[ix('Kode Guru')]||''),namaGuru:String(r[ix('Nama Guru')]||''),kelas:String(r[ix('Kelas')]||''),jamKe:String(r[ix('Jam Ke')]||''),mapel:String(r[ix('Mata Pelajaran')]||''),status:String(r[ix('Status')]||''),catatan:String(r[ix('Catatan')]||'')};
    const hay=[obj.idJadwal,obj.kodeGuru,obj.namaGuru,obj.kelas,obj.mapel,obj.status,obj.catatan].join(' ').toLowerCase();
    if(!query || hay.includes(query)) out.push(obj);
  }
  return out.slice(0,500);
}
function koreksiAbsensiAdminPortal(token, rowNumber, dateKey, scheduleId, newStatus, newNote, reason) {
  const user=getPortalSession_(token);
  if(String(user.peran||'').toUpperCase()!=='ADMIN') throw new Error('Akses khusus ADMIN.');
  rowNumber=Number(rowNumber); dateKey=String(dateKey||'').trim(); scheduleId=String(scheduleId||'').trim();
  newStatus=String(newStatus||'').trim(); newNote=String(newNote||'').trim(); reason=String(reason||'').trim();
  if(!Number.isInteger(rowNumber)||rowNumber<2) throw new Error('Baris absensi tidak valid.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)||!scheduleId) throw new Error('Identitas absensi tidak valid.');
  if(!['Hadir','Izin','Sakit','Tidak Hadir'].includes(newStatus)) throw new Error('Status koreksi tidak valid.');
  if(!reason) throw new Error('Alasan koreksi wajib diisi.');
  if(reason.length>500 || newNote.length>500) throw new Error('Alasan/catatan maksimal 500 karakter.');
  const lock=LockService.getScriptLock(); if(!lock.tryLock(25000)) throw new Error('Sistem sedang memproses. Coba lagi.');
  try {
    const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName(APP.sheets.absensi);
    if(!sh || rowNumber>sh.getLastRow()) throw new Error('Data absensi tidak ditemukan. Muat ulang daftar.');
    const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v).trim());
    const ix=k=>headers.indexOf(k);
    ['ID Jadwal','Tanggal','Kode Guru','Nama Guru','Kelas','Jam Ke','Mata Pelajaran','Status','Catatan'].forEach(k=>{if(ix(k)<0)throw new Error('Kolom ABSENSI tidak ditemukan: '+k);});
    const r=sh.getRange(rowNumber,1,1,sh.getLastColumn()).getValues()[0];
    if(normalizeDate_(r[ix('Tanggal')])!==dateKey || String(r[ix('ID Jadwal')]||'')!==scheduleId) throw new Error('Data berubah atau baris tidak cocok. Muat ulang sebelum mengoreksi.');
    const oldStatus=String(r[ix('Status')]||''), oldNote=String(r[ix('Catatan')]||'');
    const audit=ss.getSheetByName(ABSENCE_AUDIT_SHEET_); if(!audit) throw new Error('Sheet audit belum dibuat. Jalankan setupAuditAbsensi() satu kali.');
    sh.getRange(rowNumber,ix('Status')+1).setValue(newStatus);
    sh.getRange(rowNumber,ix('Catatan')+1).setValue(newNote);
    audit.appendRow([new Date(),String(user.username||''),String(user.nama||user.namaLengkap||''),dateKey,scheduleId,String(r[ix('Kode Guru')]||''),String(r[ix('Nama Guru')]||''),String(r[ix('Kelas')]||''),String(r[ix('Jam Ke')]||''),String(r[ix('Mata Pelajaran')]||''),oldStatus,newStatus,oldNote,newNote,reason]);
    return {success:true,message:'Koreksi absensi tersimpan dan tercatat pada log audit.'};
  } finally { lock.releaseLock(); }
}

/* ===== MODUL GENERATOR SOAL AI ===== */
/* =====================================================================
   GENERATOR SOAL AI — PKPPS HIDAYATULLAH KARIMUN
   Add-on backend. Uses the existing Portal session and account database.
   ===================================================================== */
const SOAL_AI_CFG_ = {
  SHEET: 'BANK_SOAL_AI',
  MODEL_DEFAULT: 'openai/gpt-oss-20b',
  MAX_MATERI: 18000,
  MAX_TOTAL: 60
};

function setupGeneratorSoalAI() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SOAL_AI_CFG_.SHEET);
  if (!sh) sh = ss.insertSheet(SOAL_AI_CFG_.SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID','Waktu','ID Pengguna','Nama Guru','Mata Pelajaran','Program',
      'Kelas','Jenis Ujian','Judul','Jumlah PG','Jumlah Benar Salah',
      'Jumlah Isian','Jumlah Uraian','Data Soal JSON'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,14).setFontWeight('bold')
      .setBackground('#145c46').setFontColor('#ffffff');
  }
  return 'Sheet BANK_SOAL_AI siap. Data lama tidak dihapus.';
}

function getGeneratorSoalData(token) {
  const user = getPortalSession_(token);
  _assertSoalTeacher_(user);
  const sh = SpreadsheetApp.getActive().getSheetByName(SOAL_AI_CFG_.SHEET);
  let rows = [];
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getDataRange().getValues();
    rows = vals.slice(1).filter(r =>
      String(r[2]) === String(user.id) || String(user.peran).toUpperCase() === 'ADMIN'
    ).slice(-50).reverse().map(r => ({
      id:String(r[0]),
      waktu:r[1] instanceof Date
        ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : String(r[1]),
      mapel:String(r[4]), program:String(r[5]), kelas:String(r[6]),
      jenis:String(r[7]), judul:String(r[8])
    }));
  }
  return {user:{id:user.id,nama:user.nama,peran:user.peran},saved:rows};
}

function generateSoalAI(token, payload) {
  const user = getPortalSession_(token);
  _assertSoalTeacher_(user);
  payload = payload || {};
  const mode = String(payload.mode || 'materi').toLowerCase() === 'kisi-kisi' ? 'kisi-kisi' : 'materi';
  const materi = String(payload.materi || '').trim();
  if (materi.length < 30) throw new Error('Materi terlalu singkat. Masukkan minimal 30 karakter.');
  if (materi.length > SOAL_AI_CFG_.MAX_MATERI) throw new Error('Materi terlalu panjang. Maksimal 18.000 karakter.');

  const counts = {
    PG: _numSoal_(payload.pilihanGanda),
    BS: _numSoal_(payload.benarSalah),
    ISIAN: _numSoal_(payload.isian),
    URAIAN: _numSoal_(payload.uraian)
  };
  const total = Object.keys(counts).reduce((n,k)=>n+counts[k],0);
  if (!total || total > SOAL_AI_CFG_.MAX_TOTAL) throw new Error('Jumlah soal harus 1–60 butir.');

  const kisiKisi = Array.isArray(payload.kisiKisi) ? payload.kisiKisi.map((x,i)=>({
    indikator:String(x.indikator||'').trim(),
    level:String(x.level||'C2').trim().toUpperCase(),
    bentuk:String(x.bentuk||'PG').trim().toUpperCase(),
    jumlah:_numSoal_(x.jumlah)
  })).filter(x=>x.indikator && x.jumlah>0) : [];
  if (mode === 'kisi-kisi') {
    if (!kisiKisi.length) throw new Error('Isi minimal satu indikator pada kisi-kisi.');
    const blueprintTotal = kisiKisi.reduce((n,x)=>n+x.jumlah,0);
    if (blueprintTotal !== total) throw new Error('Jumlah butir pada kisi-kisi ('+blueprintTotal+') harus sama dengan total jumlah soal ('+total+').');
    const allowedLevels=['C1','C2','C3','C4','C5','C6'];
    const allowedTypes=['PG','BS','ISIAN','URAIAN'];
    kisiKisi.forEach((x,i)=>{
      if(!allowedLevels.includes(x.level)) throw new Error('Level kognitif baris '+(i+1)+' harus C1 sampai C6.');
      if(!allowedTypes.includes(x.bentuk)) throw new Error('Bentuk soal baris '+(i+1)+' tidak valid.');
    });
    const byType={PG:0,BS:0,ISIAN:0,URAIAN:0};
    kisiKisi.forEach(x=>byType[x.bentuk]+=x.jumlah);
    if(Object.keys(counts).some(k=>counts[k]!==byType[k])) throw new Error('Jumlah per bentuk soal tidak sama dengan rekap jumlah soal. Sesuaikan kisi-kisi atau jumlah soal.');
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (!apiKey) throw new Error('GROQ_API_KEY belum dipasang oleh administrator di Script Properties.');
  const model = PropertiesService.getScriptProperties().getProperty('GROQ_MODEL') || SOAL_AI_CFG_.MODEL_DEFAULT;
  const meta = {
    mode:mode, mapel:String(payload.mapel||'').trim(), program:String(payload.program||'').trim(),
    kelas:String(payload.kelas||'').trim(), ujian:String(payload.ujian||'Ujian Semester').trim(),
    tingkat:String(payload.tingkat||'Sedang').trim(), bahasa:String(payload.bahasa||'Indonesia').trim(),
    durasi:Number(payload.durasi||90), instruksi:String(payload.instruksi||'').trim(), kisiKisi:kisiKisi
  };
  if (!meta.mapel || !meta.kelas) throw new Error('Mata pelajaran dan kelas wajib diisi.');

  const instructions = [
    'Anda adalah penyusun asesmen pendidikan untuk PKPPS di Indonesia.',
    'Buat soal akurat, jelas, sesuai materi sumber, kelas, bahasa, dan tingkat kesulitan.',
    'Jangan mengarang fakta yang tidak didukung materi. Hindari detail di luar materi sumber.',
    mode === 'materi'
      ? 'MODE MATERI: susun kisi-kisi secara otomatis dari materi. Tentukan indikator terukur, level kognitif C1-C6, bentuk soal, dan jumlah butir. Pastikan total dan jumlah per bentuk persis sesuai permintaan.'
      : 'MODE KISI-KISI: kisi-kisi dari guru adalah ketentuan wajib. Setiap soal harus sesuai indikator, level kognitif, bentuk, dan jumlah pada baris kisi-kisi. Jangan mengubah atau mengabaikan kisi-kisi.',
    'Pilihan ganda wajib tepat 4 opsi A-D dan satu jawaban benar.',
    'Benar-salah: jawaban hanya Benar atau Salah.',
    'Isian dan uraian wajib memiliki jawaban acuan serta pedoman penskoran singkat.',
    'Jangan menampilkan kunci jawaban pada teks soal siswa.',
    'Kembalikan hanya JSON valid dengan struktur {"judul":string,"petunjuk":string,"kisiKisi":[{"indikator":string,"level":"C1-C6","bentuk":"PG|BS|ISIAN|URAIAN","jumlah":number}],"soal":[{"nomor":number,"jenis":"PG|BS|ISIAN|URAIAN","indikator":string,"level":"C1-C6","pertanyaan":string,"opsi":{"A":string,"B":string,"C":string,"D":string},"jawaban":string,"pembahasan":string,"skor":number}]}',
    'Untuk jenis selain PG, opsi harus objek kosong. Jumlah soal per jenis wajib persis sesuai permintaan. Nomor berurutan dari 1.'
  ].join('\n');

  const input = [
    'IDENTITAS UJIAN:', JSON.stringify(meta),
    'JUMLAH SOAL PER JENIS:', JSON.stringify(counts),
    'TOTAL BUTIR:', String(total),
    'KISI-KISI YANG WAJIB DIIKUTI (MODE KISI-KISI):', mode==='kisi-kisi'?JSON.stringify(kisiKisi):'Buat otomatis berdasarkan materi.',
    'CATATAN GURU:', meta.instruksi || '-',
    'MATERI SUMBER:', materi,
    'FORMAT: Kembalikan hanya JSON valid, tanpa markdown atau teks tambahan.',
    'Buat soal dan kunci jawaban lengkap dengan pembahasan. Pastikan distribusi kisi-kisi dan soal konsisten.'
  ].join('\n\n');

  const response = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+apiKey},
    payload:JSON.stringify({model:model,messages:[{role:'system',content:instructions},{role:'user',content:input}],response_format:{type:'json_object'},temperature:0.3,max_completion_tokens:9000}),
    muteHttpExceptions:true
  });
  const status=response.getResponseCode(), raw=response.getContentText();
  let api; try{api=JSON.parse(raw);}catch(e){throw new Error('Respons API tidak dapat dibaca. HTTP '+status);}
  if(status<200||status>=300){const msg=api&&api.error&&api.error.message?api.error.message:raw.slice(0,500);throw new Error('Groq API ('+status+'): '+msg);}
  const outputText=api&&api.choices&&api.choices[0]&&api.choices[0].message?String(api.choices[0].message.content||''):'';
  if(!outputText)throw new Error('Groq tidak mengembalikan teks keluaran.');
  let result;try{result=JSON.parse(outputText);}catch(e){throw new Error('AI mengembalikan format soal yang tidak valid. Silakan buat ulang.');}
  if(!result||!Array.isArray(result.soal)||result.soal.length!==total)throw new Error('Jumlah soal dari AI tidak sesuai. Silakan buat ulang.');

  const actual={PG:0,BS:0,ISIAN:0,URAIAN:0};
  result.soal.forEach((q,i)=>{
    q.jenis=String(q.jenis||'').toUpperCase();
    if(!Object.prototype.hasOwnProperty.call(actual,q.jenis))throw new Error('Jenis soal nomor '+(i+1)+' tidak valid.');
    actual[q.jenis]++;q.nomor=i+1;q.pertanyaan=String(q.pertanyaan||'').trim();q.jawaban=String(q.jawaban||'').trim();q.pembahasan=String(q.pembahasan||'').trim();q.indikator=String(q.indikator||'').trim();q.level=String(q.level||'').toUpperCase();q.skor=Number(q.skor||1);
    if(!q.pertanyaan||!q.jawaban)throw new Error('Soal atau kunci nomor '+(i+1)+' kosong.');
    if(q.jenis==='PG'){q.opsi=q.opsi||{};['A','B','C','D'].forEach(k=>q.opsi[k]=String(q.opsi[k]||'').trim());if(['A','B','C','D'].some(k=>!q.opsi[k]))throw new Error('Opsi nomor '+(i+1)+' belum lengkap.');if(!['A','B','C','D'].includes(q.jawaban.toUpperCase()))throw new Error('Kunci pilihan ganda nomor '+(i+1)+' harus A, B, C, atau D.');}else{q.opsi={};}
  });
  if(Object.keys(counts).some(k=>actual[k]!==counts[k]))throw new Error('Komposisi jenis soal dari AI belum sesuai permintaan. Silakan buat ulang.');
  let finalBlueprint=Array.isArray(result.kisiKisi)?result.kisiKisi:kisiKisi;
  if(mode==='materi'&&!finalBlueprint.length)throw new Error('AI belum menghasilkan kisi-kisi otomatis. Silakan buat ulang.');
  return {meta:meta,judul:String(result.judul||meta.ujian+' - '+meta.mapel),petunjuk:String(result.petunjuk||'Bacalah setiap soal dengan teliti.'),kisiKisi:finalBlueprint,soal:result.soal,model:model};
}

function saveSoalAIBank(token, payload) {
  const user = getPortalSession_(token);
  _assertSoalTeacher_(user);
  payload = payload || {};
  const soal = payload.result;
  if (!soal || !Array.isArray(soal.soal) || !soal.soal.length) throw new Error('Tidak ada soal untuk disimpan.');
  const meta = soal.meta || {};
  const sh = _getSoalSheet_();
  const id = 'SOAL-'+Utilities.getUuid().slice(0,8).toUpperCase();
  const json = JSON.stringify(soal);
  if (json.length > 45000) throw new Error('Dokumen soal terlalu besar untuk disimpan. Kurangi jumlah soal atau ringkas pembahasan.');
  sh.appendRow([
    id,new Date(),String(user.id),String(user.nama),String(meta.mapel||''),
    String(meta.program||''),String(meta.kelas||''),String(meta.ujian||''),
    String(soal.judul||''),soal.soal.filter(x=>x.jenis==='PG').length,
    soal.soal.filter(x=>x.jenis==='BS').length,
    soal.soal.filter(x=>x.jenis==='ISIAN').length,
    soal.soal.filter(x=>x.jenis==='URAIAN').length,
    json
  ]);
  return {id:id,message:'Soal berhasil disimpan ke Bank Soal.'};
}

function getSavedSoalAI(token, id) {
  const user = getPortalSession_(token);
  _assertSoalTeacher_(user);
  const sh = _getSoalSheet_();
  if (sh.getLastRow() < 2) throw new Error('Bank Soal masih kosong.');
  const vals = sh.getDataRange().getValues();
  const row = vals.slice(1).find(r=>String(r[0])===String(id) &&
    (String(r[2])===String(user.id) || String(user.peran).toUpperCase()==='ADMIN'));
  if (!row) throw new Error('Soal tidak ditemukan atau bukan milik akun ini.');
  return JSON.parse(String(row[13]));
}

function _assertSoalTeacher_(user) {
  const role = String(user && user.peran || '').toUpperCase();
  if (!['GURU','ADMIN'].includes(role)) throw new Error('Modul Generator Soal hanya untuk guru dan admin.');
}
function _numSoal_(v) {
  const n = Number(v || 0);
  if (!Number.isInteger(n) || n < 0 || n > 60) throw new Error('Jumlah soal harus bilangan bulat 0–60.');
  return n;
}
function _getSoalSheet_() {
  let sh = SpreadsheetApp.getActive().getSheetByName(SOAL_AI_CFG_.SHEET);
  if (!sh) {
    setupGeneratorSoalAI();
    sh = SpreadsheetApp.getActive().getSheetByName(SOAL_AI_CFG_.SHEET);
  }
  return sh;
}



/* ===== MODUL ABSENSI SANTRI V1.0 =====
 * Menyimpan ke sheet ABSENSI_SANTRI. Data SANTRI/SISWA tetap menjadi sumber roster.
 * Fungsi dipanggil dari Portal.html melalui google.script.run.
 */
const SANTRI_ATTENDANCE_SHEET_ = 'ABSENSI_SANTRI';
const SANTRI_ATTENDANCE_HEADERS_ = ['ID Absensi','Tanggal','NIS','Nama Santri','Kelas','Status','Keterangan','Dicatat Oleh','Waktu Catat'];

function santriAttendanceUser_(token) {
  const u = getPortalSession_(token);
  if (!['ADMIN','GURU','WALI','TU'].includes(String(u.peran||'').toUpperCase())) {
    throw new Error('Akses Absensi Santri hanya untuk ADMIN, TU, GURU, dan WALI KELAS.');
  }
  return u;
}
function santriAttendanceSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SANTRI_ATTENDANCE_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(SANTRI_ATTENDANCE_SHEET_);
    sh.appendRow(SANTRI_ATTENDANCE_HEADERS_);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,SANTRI_ATTENDANCE_HEADERS_.length)
      .setFontWeight('bold').setBackground('#173b63').setFontColor('#ffffff');
  }
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x||'').trim());
  SANTRI_ATTENDANCE_HEADERS_.forEach(h=>{ if (!headers.includes(h)) { sh.getRange(1,sh.getLastColumn()+1).setValue(h); headers.push(h); } });
  return {sh, headers};
}
function santriAttendanceDate_(value) {
  const s = String(value||'').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Format tanggal harus YYYY-MM-DD.');
  const d = new Date(s+'T00:00:00+07:00');
  if (isNaN(d.getTime())) throw new Error('Tanggal tidak valid.');
  return s;
}
function getSantriAttendanceSetup(token, selectedDate) {
  const user = santriAttendanceUser_(token);
  const date = selectedDate ? santriAttendanceDate_(selectedDate) : Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd');
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName('SANTRI') || ss.getSheetByName('SISWA');
  if (!source || source.getLastRow()<2) throw new Error('Data santri belum tersedia pada sheet SANTRI/SISWA.');
  const values = source.getDataRange().getDisplayValues();
  const h = values[0].map(x=>String(x||'').trim());
  const find = (...names)=>{ for (const n of names) { const i=h.indexOf(n); if(i>=0)return i; } return -1; };
  const cNis=find('NIS','ID Santri','NISN','ID Siswa'), cName=find('Nama Santri','Nama Siswa','Nama','Nama Lengkap');
  const cClass=find('Kelas','Nama Kelas'), cStatus=find('Status'), cJenjang=find('Jenjang');
  if(cName<0) throw new Error('Kolom Nama Santri tidak ditemukan.');
  const students = values.slice(1).filter(r=>r[cName] && (cStatus<0 || !['NONAKTIF','KELUAR','LULUS'].includes(String(r[cStatus]||'').trim().toUpperCase())))
    .map(r=>({nis:cNis>=0?String(r[cNis]||'').trim():'',nama:String(r[cName]).trim(),kelas:cClass>=0?String(r[cClass]||'').trim():'',jenjang:cJenjang>=0?String(r[cJenjang]||'').trim():''}));
  const {sh,headers}=santriAttendanceSheet_(), rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues():[];
  const ix=k=>headers.indexOf(k);
  const existing={};
  rows.forEach(r=>{if(String(r[ix('Tanggal')])===date) existing[String(r[ix('NIS')])]={status:String(r[ix('Status')]||''),keterangan:String(r[ix('Keterangan')]||'')};});
  const classes=[...new Set(students.map(x=>x.kelas).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id'));
  return {date, user:{nama:user.nama,peran:user.peran},classes,students:students.map(x=>({...x,absensi:existing[x.nis]||null})),statuses:['Hadir','Izin','Sakit','Alfa']};
}
function saveSantriAttendance(token, payload) {
  const user=santriAttendanceUser_(token); payload=payload||{};
  const date=santriAttendanceDate_(payload.tanggal), kelas=String(payload.kelas||'').trim();
  const items=Array.isArray(payload.items)?payload.items:[];
  if(!kelas) throw new Error('Pilih kelas terlebih dahulu.');
  if(!items.length) throw new Error('Tidak ada data absensi untuk disimpan.');
  const ss=SpreadsheetApp.getActive(), source=ss.getSheetByName('SANTRI')||ss.getSheetByName('SISWA');
  if(!source||source.getLastRow()<2) throw new Error('Data santri belum tersedia.');
  const sv=source.getDataRange().getDisplayValues(), sh=sv[0].map(x=>String(x||'').trim());
  const col=(...names)=>{for(const n of names){const i=sh.indexOf(n);if(i>=0)return i;}return -1;};
  const ni=col('NIS','ID Santri','NISN','ID Siswa'), nn=col('Nama Santri','Nama Siswa','Nama','Nama Lengkap'), nc=col('Kelas','Nama Kelas');
  if(nn<0||ni<0) throw new Error('Kolom NIS dan Nama Santri wajib tersedia pada data santri.');
  const roster=new Map(sv.slice(1).filter(r=>r[nn]).map(r=>[String(r[ni]||'').trim(),{nama:String(r[nn]).trim(),kelas:nc>=0?String(r[nc]||'').trim():''}]));
  const {sh:dest,headers}=santriAttendanceSheet_(), ix=k=>headers.indexOf(k);
  const last=dest.getLastRow(), old=last>1?dest.getRange(2,1,last-1,dest.getLastColumn()).getDisplayValues():[];
  const rowByKey=new Map();
  old.forEach((r,i)=>rowByKey.set(date+'|'+String(r[ix('NIS')]).trim(),i+2));
  const now=new Date(); let saved=0;
  items.forEach(item=>{
    const nis=String(item.nis||'').trim(), status=String(item.status||'').trim(), student=roster.get(nis);
    if(!student||student.kelas!==kelas) throw new Error('Data santri tidak sesuai dengan kelas yang dipilih: '+nis);
    if(!['Hadir','Izin','Sakit','Alfa'].includes(status)) throw new Error('Status absensi tidak valid untuk '+student.nama+'.');
    const record={ 'ID Absensi':'AS-'+Utilities.getUuid().slice(0,12), 'Tanggal':date, 'NIS':nis, 'Nama Santri':student.nama, 'Kelas':kelas, 'Status':status, 'Keterangan':String(item.keterangan||'').trim(), 'Dicatat Oleh':String(user.nama||user.username||''), 'Waktu Catat':now };
    const row=headers.map(k=>record[k]!==undefined?record[k]:'');
    const existingRow=rowByKey.get(date+'|'+nis);
    if(existingRow) dest.getRange(existingRow,1,1,headers.length).setValues([row]); else {dest.appendRow(row);rowByKey.set(date+'|'+nis,dest.getLastRow());}
    saved++;
  });
  return {success:true,saved,message:saved+' data absensi santri kelas '+kelas+' tanggal '+date+' berhasil disimpan.'};
}
function getSantriAttendanceReport(token, filters) {
  santriAttendanceUser_(token); filters=filters||{};
  const date=santriAttendanceDate_(filters.tanggal), kelas=String(filters.kelas||'').trim();
  const {sh,headers}=santriAttendanceSheet_(), ix=k=>headers.indexOf(k);
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues():[];
  const data=rows.filter(r=>String(r[ix('Tanggal')])===date&&(!kelas||String(r[ix('Kelas')])===kelas)).map(r=>({nis:r[ix('NIS')],nama:r[ix('Nama Santri')],kelas:r[ix('Kelas')],status:r[ix('Status')],keterangan:r[ix('Keterangan')]}));
  const recap={Hadir:0,Izin:0,Sakit:0,Alfa:0}; data.forEach(x=>{if(recap[x.status]!==undefined)recap[x.status]++;});
  return {date,kelas,data,recap,total:data.length};
}

/**
 * API JSON untuk Generator Soal AI yang tampil di Cloudflare Worker.
 * Tambahkan fungsi ini di bagian paling bawah Code.gs yang sudah ada.
 * Jangan menghapus doGet atau fungsi Portal yang sudah berjalan.
 */
function doPost(e) {
  let result;
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const request = JSON.parse(raw);
    const action = String(request.action || '').trim();
    const args = Array.isArray(request.args) ? request.args : [];

    // Hanya fungsi API publik yang dipakai halaman Vercel yang boleh dipanggil.
    const allowedActions = new Set([
    'activatePortalRole',
    'addPortalClass',
    'changePin',
    'createPortalUser',
    'createPpdbAdminEntry',
    'deactivatePortalSchedule',
    'deletePortalUser',
    'exportMonthlyReport',
    'generateSoalAI',
    'getAbsensiUntukKoreksiPortal',
    'getAdminData',
    'getAllDailySchedule',
    'getClassManagementData',
    'getClassMonitoring',
    'getDailyWhatsAppDraft',
    'getDashboard',
    'getFinanceData',
    'getGeneratorSoalData',
    'getGradeRoster',
    'getGradeSetup',
    'getHonorRecords',
    'getHonorSetup',
    'getKbmMonitoringDashboard',
    'getKbmPeriodReport',
    'getMonthlyReport',
    'getPaymentReceipt',
    'getPortalDashboard',
    'getPortalScheduleData',
    'getPortalSettings',
    'getPortalTeacherDirectory',
    'getPortalUsers',
    'getPortalWebAppUrl',
    'getPpdbAdminData',
    'getPpdbDetail',
    'getPpdbDocuments',
    'getReportData',
    'getRingkasanKehadiranPortal',
    'getSantriAttendanceReport',
    'getSantriAttendanceSetup',
    'getSantriDirectory',
    'getSantriSessionAttendanceRoster',
    'getSantriSessionAttendanceSetup',
    'getSavedSoalAI',
    'getTahfidzPortalData',
    'getTeacherKbmData',
    'getVideoStudioIdentity',
    'getWeeklyReport',
    'installDailyWhatsAppDraftTrigger',
    'koreksiAbsensiAdminPortal',
    'login',
    'loginPortal',
    'logout',
    'logoutPortal',
    'recordInvoicePayment',
    'resetPortalUserPin',
    'saveClassWali',
    'saveFinanceRecord',
    'saveGradeBatch',
    'saveHonorRecord',
    'saveMonthlyApproval',
    'savePortalSchedule',
    'savePortalSettings',
    'savePortalTeacherDirectory',
    'saveSantriAttendance',
    'saveSantriSessionAttendance',
    'saveSchedule',
    'saveSoalAIBank',
    'saveTahfidzPortalRecord',
    'saveTeacherKbm',
    'searchTeacherAttendanceHistory',
    'setPortalTeacherStatus',
    'setPortalUserActive',
    'setPortalUserRoles',
    'submitAttendance',
    'submitPpdbPublic',
    'updatePpdbStatus',
    'uploadPpdbDocument'
    ]);
    if (!allowedActions.has(action)) throw new Error('Aksi API tidak dikenal atau tidak diizinkan: ' + action);

    const handler = globalThis[action];
    if (typeof handler !== 'function') throw new Error('Fungsi layanan tidak tersedia: ' + action);
    result = { ok: true, data: handler.apply(null, args) };
  } catch (err) {
    result = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}




/* ===== RINGKASAN KEHADIRAN GURU & SANTRI PORTAL V1.1 =====
 * Kompatibel dengan ABSENSI_SANTRI (absensi harian) dan
 * ABSENSI_SANTRI_SESI (absensi santri yang ditautkan ke ID Jadwal).
 * Fungsi ini dipanggil dari Portal.html melalui google.script.run.
 */
function getRingkasanKehadiranPortal(token, selectedDate) {
  const user = getPortalSession_(token);
  const role = String(user.peran || '').toUpperCase();
  if (!['ADMIN','GURU','TU'].includes(role)) {
    throw new Error('Ringkasan kehadiran hanya tersedia untuk ADMIN, TU, dan GURU.');
  }

  const dateKey = String(selectedDate || Utilities.formatDate(new Date(), APP.timezone, 'yyyy-MM-dd')).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) throw new Error('Tanggal ringkasan tidak valid.');
  const parts = dateKey.split('-').map(Number);
  const dateObj = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12));
  if (dateObj.getUTCFullYear() !== parts[0] || dateObj.getUTCMonth() !== parts[1] - 1 || dateObj.getUTCDate() !== parts[2]) {
    throw new Error('Tanggal ringkasan tidak valid.');
  }
  const dayNames = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  const dayName = dayNames[dateObj.getUTCDay()];

  let schedules = readObjects_(APP.sheets.jadwal).filter(s =>
    isActive_(s['Aktif']) && normalizeDay_(s['Hari']) === normalizeDay_(dayName)
  );
  if (role === 'GURU') {
    const code = normalizeCode_(user.kodeGuru || user.kode || '');
    schedules = schedules.filter(s => normalizeCode_(s['Kode Guru']) === code);
  }

  const classes = new Map();
  schedules.forEach(s => {
    const cls = String(s['Kelas'] || 'Tidak diketahui').trim() || 'Tidak diketahui';
    if (!classes.has(cls)) classes.set(cls, {
      kelas: cls, totalSesi: 0, guruHadir: 0, guruIzin: 0, guruSakit: 0,
      guruTidakHadir: 0, guruBelum: 0, santriHadir: 0, santriIzin: 0,
      santriSakit: 0, santriAlpa: 0, santriTercatat: 0
    });
    classes.get(cls).totalSesi++;
  });

  // Rekap guru berdasarkan jadwal aktif pada tanggal yang dipilih.
  const teacherRows = readObjects_(APP.sheets.absensi).filter(a => normalizeDate_(a['Tanggal']) === dateKey);
  const teacherById = new Map();
  teacherRows.forEach(a => {
    const id = String(a['ID Jadwal'] || '').trim();
    if (id) teacherById.set(id, a);
  });
  const teacher = { total: schedules.length, hadir: 0, izin: 0, sakit: 0, tidakHadir: 0, belum: 0 };
  schedules.forEach(s => {
    const cls = String(s['Kelas'] || 'Tidak diketahui').trim() || 'Tidak diketahui';
    const c = classes.get(cls);
    const rec = teacherById.get(String(s['ID Jadwal'] || '').trim());
    const status = String(rec && rec['Status'] || '').trim().toLowerCase();
    if (!rec) { teacher.belum++; c.guruBelum++; }
    else if (status === 'hadir') { teacher.hadir++; c.guruHadir++; }
    else if (status === 'izin') { teacher.izin++; c.guruIzin++; }
    else if (status === 'sakit') { teacher.sakit++; c.guruSakit++; }
    else { teacher.tidakHadir++; c.guruTidakHadir++; }
  });

  // Prioritaskan catatan per sesi jika sheet tersebut memang berisi data tanggal ini.
  // Jika belum digunakan, baca sheet ABSENSI_SANTRI dari menu Absensi Santri Portal.
  const ss = SpreadsheetApp.getActive();
  const sessionSheet = ss.getSheetByName('ABSENSI_SANTRI_SESI');
  const dailySheet = ss.getSheetByName('ABSENSI_SANTRI');
  const santri = { total: 0, hadir: 0, izin: 0, sakit: 0, alpa: 0 };
  let sessionRecords = [];

  if (sessionSheet && sessionSheet.getLastRow() > 1) {
    const values = sessionSheet.getDataRange().getValues();
    const headers = values[0].map(v => String(v || '').trim());
    const ix = {}; headers.forEach((h, i) => { ix[h] = i; });
    const required = ['Tanggal','ID Jadwal','Kelas','Status'];
    const hasSchema = required.every(k => ix[k] !== undefined);
    if (hasSchema) {
      const allowedIds = new Set(schedules.map(x => String(x['ID Jadwal'] || '').trim()));
      sessionRecords = values.slice(1).filter(r =>
        normalizeDate_(r[ix['Tanggal']]) === dateKey &&
        allowedIds.has(String(r[ix['ID Jadwal']] || '').trim())
      ).map(r => ({ kelas: String(r[ix['Kelas']] || '').trim(), status: String(r[ix['Status']] || '').trim() }));
    }
  }

  if (sessionRecords.length) {
    sessionRecords.forEach(rec => {
      const c = classes.get(rec.kelas);
      const st = rec.status.toLowerCase();
      santri.total++;
      if (c) c.santriTercatat++;
      if (st === 'hadir') { santri.hadir++; if (c) c.santriHadir++; }
      else if (st === 'izin') { santri.izin++; if (c) c.santriIzin++; }
      else if (st === 'sakit') { santri.sakit++; if (c) c.santriSakit++; }
      else { santri.alpa++; if (c) c.santriAlpa++; }
    });
  } else if (dailySheet && dailySheet.getLastRow() > 1) {
    const values = dailySheet.getDataRange().getDisplayValues();
    const headers = values[0].map(v => String(v || '').trim());
    const ix = {}; headers.forEach((h, i) => { ix[h] = i; });
    const col = (...names) => { for (const n of names) if (ix[n] !== undefined) return ix[n]; return -1; };
    const iDate = col('Tanggal'), iClass = col('Kelas','Nama Kelas'), iStatus = col('Status');
    if (iDate < 0 || iClass < 0 || iStatus < 0) {
      throw new Error('Header ABSENSI_SANTRI harus memuat Tanggal, Kelas, dan Status.');
    }
    const allowedClasses = new Set(Array.from(classes.keys()).map(x => x.toLowerCase()));
    values.slice(1).forEach(r => {
      if (String(r[iDate] || '').trim() !== dateKey) return;
      const cls = String(r[iClass] || '').trim();
      if (role === 'GURU' && !allowedClasses.has(cls.toLowerCase())) return;
      const c = classes.get(cls);
      const st = String(r[iStatus] || '').trim().toLowerCase();
      santri.total++;
      if (c) c.santriTercatat++;
      if (st === 'hadir') { santri.hadir++; if (c) c.santriHadir++; }
      else if (st === 'izin') { santri.izin++; if (c) c.santriIzin++; }
      else if (st === 'sakit') { santri.sakit++; if (c) c.santriSakit++; }
      else if (st === 'alfa' || st === 'alpa' || st === 'tidak hadir') { santri.alpa++; if (c) c.santriAlpa++; }
    });
  }

  return {
    date: dateKey,
    day: dayName,
    role: role,
    teacher: teacher,
    santri: santri,
    classes: Array.from(classes.values()).sort((a, b) => a.kelas.localeCompare(b.kelas, 'id', { numeric: true }))
  };
}


/* ===== MODUL TAHFIDZ PORTAL V1.0 ===== */
const TAHFIDZ_SHEET_='TAHFIDZ_SETORAN';
const TAHFIDZ_HEADERS_=['ID','Tanggal','NIS','Nama Santri','Kelas','Jenis Setoran','Surah','Ayat Awal','Ayat Akhir','Kelancaran','Tajwid & Makhraj','Catatan','Kode Pembimbing','Nama Pembimbing','Waktu Input'];
function ensureTahfidzSheet_(){const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(TAHFIDZ_SHEET_);if(!sh){sh=ss.insertSheet(TAHFIDZ_SHEET_);sh.appendRow(TAHFIDZ_HEADERS_);sh.setFrozenRows(1);sh.getRange(1,1,1,TAHFIDZ_HEADERS_.length).setFontWeight('bold').setBackground('#173b63').setFontColor('#fff');}return sh;}
function getTahfidzSantri_(){const ss=SpreadsheetApp.getActive();const sh=ss.getSheetByName('SANTRI')||ss.getSheetByName('SISWA');if(!sh||sh.getLastRow()<2)return [];const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim());const ix=(...a)=>{for(const n of a){const i=h.indexOf(n);if(i>=0)return i;}return -1;};const n=ix('Nama Santri','Nama Siswa','Nama','Nama Lengkap'),nis=ix('NIS','NISN','ID Santri','ID Siswa'),k=ix('Kelas','Nama Kelas'),st=ix('Status');if(n<0)throw new Error('Kolom nama santri tidak ditemukan pada sheet '+sh.getName()+'.');return v.slice(1).filter(r=>r[n]&&(!st||!r[st]||/aktif|ya|lulus/i.test(r[st]))).map(r=>({nis:nis>=0?String(r[nis]||''):String(r[n]),nama:String(r[n]),kelas:k>=0?String(r[k]||''):''}));}
function getTahfidzPortalData(token){const u=getPortalSession_(token),role=String(u.peran||'').toUpperCase();if(!['ADMIN','GURU','WALI','SISWA','TU'].includes(role))throw new Error('Akses menu Tahfidz ditolak.');const roster=getTahfidzSantri_(),sh=ensureTahfidzSheet_();let records=[];if(sh.getLastRow()>1){const v=sh.getDataRange().getDisplayValues();records=v.slice(1).filter(r=>r[0]).map(r=>({id:r[0],tanggal:r[1],nis:r[2],nama:r[3],kelas:r[4],jenis:r[5],surah:r[6],ayatAwal:r[7],ayatAkhir:r[8],kelancaran:r[9],tajwid:r[10],catatan:r[11],kodePembimbing:r[12],pembimbing:r[13]}));}if(role==='WALI'||role==='SISWA'){const name=String(u.nama||'').trim().toLowerCase();records=records.filter(x=>String(x.nama).trim().toLowerCase()===name);}return {santri:['ADMIN','GURU'].includes(role)?roster:[],records};}
function saveTahfidzPortalRecord(token,d){const u=getPortalSession_(token);if(!['ADMIN','GURU'].includes(String(u.peran||'').toUpperCase()))throw new Error('Hanya ADMIN dan GURU yang dapat menginput setoran.');d=d||{};const roster=getTahfidzSantri_(),s=roster.find(x=>x.nis===String(d.nis));if(!s)throw new Error('Santri tidak ditemukan atau tidak aktif.');const date=String(d.tanggal||'');if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Tanggal setoran tidak valid.');const start=Number(d.ayatAwal),end=Number(d.ayatAkhir);if(!d.surah||!Number.isInteger(start)||!Number.isInteger(end)||start<1||end<start)throw new Error('Surah dan rentang ayat tidak valid.');const jenis=['Hafalan Baru','Murajaah'].includes(d.jenis)?d.jenis:'Hafalan Baru';const sh=ensureTahfidzSheet_();sh.appendRow([Utilities.getUuid(),date,s.nis,s.nama,s.kelas,jenis,String(d.surah).trim(),start,end,String(d.kelancaran||''),String(d.tajwid||''),String(d.catatan||''),String(u.kode||''),String(u.nama||''),new Date()]);return {ok:true,message:'Setoran '+jenis+' untuk '+s.nama+' berhasil disimpan.'};}



/* ===== DATA GURU & PENGATURAN PORTAL V1.36 ===== */
function getPortalTeacherDirectory(token){
  const u=getPortalSession_(token);if(!['ADMIN','TU'].includes(String(u.peran).toUpperCase()))throw new Error('Hanya Admin/TU yang dapat melihat Data Guru.');
  const sh=SpreadsheetApp.getActive().getSheetByName('GURU');if(!sh||sh.getLastRow()<2)return {rows:[]};
  const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ix=k=>h.indexOf(k);
  const iCode=ix('Kode Guru'),iName=ix('Nama Guru'),iUser=ix('Username'),iActive=ix('Aktif');
  if(iCode<0||iName<0)throw new Error('Header sheet GURU harus memuat Kode Guru dan Nama Guru.');
  return {rows:v.slice(1).filter(r=>r[iCode]||r[iName]).map(r=>({kode:String(r[iCode]||''),nama:String(r[iName]||''),username:iUser>=0?String(r[iUser]||''):'',aktif:iActive<0||/^(ya|aktif|1|true)$/i.test(String(r[iActive]||''))}))};
}
function savePortalTeacherDirectory(token,d){
  const u=getPortalSession_(token);if(String(u.peran).toUpperCase()!=='ADMIN')throw new Error('Hanya Admin yang dapat mengubah Data Guru.');d=d||{};
  const kode=String(d.kode||'').trim(),nama=String(d.nama||'').trim();if(!kode||!nama)throw new Error('Kode dan nama guru wajib diisi.');
  const sh=SpreadsheetApp.getActive().getSheetByName('GURU');if(!sh)throw new Error('Sheet GURU tidak ditemukan.');
  const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ic=h.indexOf('Kode Guru'),inn=h.indexOf('Nama Guru'),iu=h.indexOf('Username'),ia=h.indexOf('Aktif');if(ic<0||inn<0)throw new Error('Header sheet GURU tidak sesuai.');
  let row=-1;for(let i=1;i<v.length;i++)if(String(v[i][ic]||'').trim()===kode){row=i+1;break;}
  if(row<0){row=sh.getLastRow()+1;sh.getRange(row,ic+1).setValue(kode);}
  sh.getRange(row,inn+1).setValue(nama);if(iu>=0)sh.getRange(row,iu+1).setValue(String(d.username||''));if(ia>=0)sh.getRange(row,ia+1).setValue(String(d.aktif||'YA'));
  return {ok:true,message:'Data guru '+nama+' berhasil disimpan.'};
}
function setPortalTeacherStatus(token,kode,aktif){
  const u=getPortalSession_(token);if(String(u.peran).toUpperCase()!=='ADMIN')throw new Error('Hanya Admin yang dapat mengubah status guru.');
  const sh=SpreadsheetApp.getActive().getSheetByName('GURU');if(!sh)throw new Error('Sheet GURU tidak ditemukan.');const v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ic=h.indexOf('Kode Guru'),ia=h.indexOf('Aktif');if(ic<0||ia<0)throw new Error('Kolom Kode Guru/Aktif tidak ditemukan.');
  for(let i=1;i<v.length;i++)if(String(v[i][ic]||'').trim()===String(kode)){sh.getRange(i+1,ia+1).setValue(aktif?'YA':'TIDAK');return {ok:true};}throw new Error('Kode guru tidak ditemukan.');
}
function getPortalSettings(token){
  const u=getPortalSession_(token);if(String(u.peran).toUpperCase()!=='ADMIN')throw new Error('Hanya Admin yang dapat membuka Pengaturan.');
  const sh=SpreadsheetApp.getActive().getSheetByName('PENGATURAN');if(!sh)return {settings:{}};const v=sh.getDataRange().getDisplayValues(),out={};v.slice(1).forEach(r=>{if(r[0])out[String(r[0]).trim()]=String(r[1]||'');});return {settings:out};
}
function savePortalSettings(token,settings){
  const u=getPortalSession_(token);if(String(u.peran).toUpperCase()!=='ADMIN')throw new Error('Hanya Admin yang dapat mengubah Pengaturan.');
  const allowed=['NAMA_LEMBAGA','TAHUN_AJARAN','LOGO_URL','KEPALA_PKPPS','KEPALA_WUSTHO','KEPALA_ULYA'];const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName('PENGATURAN');if(!sh){sh=ss.insertSheet('PENGATURAN');sh.appendRow(['Kunci','Nilai']);}
  const v=sh.getDataRange().getDisplayValues(),map={};v.slice(1).forEach((r,i)=>{if(r[0])map[String(r[0]).trim()]=i+2;});settings=settings||{};
  allowed.forEach(k=>{if(settings[k]===undefined)return;const row=map[k]||sh.getLastRow()+1;if(!map[k])sh.getRange(row,1).setValue(k);sh.getRange(row,2).setValue(String(settings[k]||''));});return {ok:true,message:'Pengaturan lembaga berhasil disimpan.'};
}


/* ===== AI VIDEO STUDIO: identity bridge =====
 * Dipanggil oleh backend video untuk memvalidasi sesi Portal yang sudah ada.
 * Tidak membuat sistem login baru dan tidak mengirim PIN/hash ke layanan lain.
 */
function getVideoStudioIdentity(token) {
  const user = getPortalSession_(token);
  const role = String(user.peran || '').trim().toUpperCase();
  if (!['ADMIN', 'TU', 'GURU'].includes(role)) {
    throw new Error('Akses AI Video Studio hanya tersedia untuk ADMIN, TU, dan GURU.');
  }
  return {
    id: String(user.id || user.username || ''),
    username: String(user.username || ''),
    name: String(user.nama || ''),
    role: role,
    organizationId: 'pkpps-hidayatullah-karimun'
  };
}
