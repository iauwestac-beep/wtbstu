// lib/excel_service.dart
//
// ورود/خروجی اکسل برای برنامهٔ «ثبت تخلف پوشش»
// ─ ورود: فایل دانشجویان دانشگاه (نگاشت نرم سرصفحه‌ها)
// ─ خروجی: یک شیت «تخلفات» با همان ستون‌های فایل ورودی + ستون‌های تخلف
import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';
import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:shamsi_date/shamsi_date.dart';
/// ───────── کلیدهای استاندارد ستون‌ها ─────────
class F {
  // کلیدهای پایه (ستون‌های اصلی دانشجو)
  static const studentNo = 'studentNo';
  static const nationalId = 'nationalId';
  static const firstName = 'firstName';
  static const lastName = 'lastName';
  static const major = 'major';
  static const college = 'college';
  static const fatherName = 'fatherName';
  static const birthCertNo = 'birthCertNo';
  static const birthDate = 'birthDate';
  static const birthPlace = 'birthPlace';
  static const issuePlace = 'issuePlace';
  static const calendarCode = 'calendarCode';
  static const calendarName = 'calendarName';
  static const gender = 'gender';
  static const level = 'level';
  static const systemName = 'systemName';
  static const admissionType = 'admissionType';
  static const faLangCertificate = 'faLangCertificate';
  static const faLangCertificateOk = 'faLangCertificateOk';
  static const startDate = 'startDate';
  static const startDateCal = 'startDateCal';
  static const endDate = 'endDate';
  static const studyStatus = 'studyStatus';
  static const active = 'active';
  static const semesterCredits = 'semesterCredits';
  static const termsLeft = 'termsLeft';
  static const termsPassed = 'termsPassed';
  static const statusCode = 'statusCode';
  static const statusName = 'statusName';
  static const statusName2 = 'statusName2';
  static const studyStart = 'studyStart';
  static const studyStart2 = 'studyStart2';
  static const studyEnd = 'studyEnd';
  static const studyEnd2 = 'studyEnd2';
  static const daysPassed = 'daysPassed';
  static const personId = 'personId';
  static const graduationDate = 'graduationDate';
  static const erstDate = 'erstDate';
  static const yearsExtension = 'yearsExtension';
  static const archiveNo = 'archiveNo';
  static const extensionDate = 'extensionDate';
  static const mobile = 'mobile';
  static const entryTerm = 'entryTerm';
}
/// ───────── مدل‌ها ─────────
class Student {
  final String studentNo;
  final String nationalId;
  final String firstName;
  final String lastName;
  final String major;
  final int sourceRow;
  /// ستون‌های دیگر فایل دانشگاه: کلید استاندارد ← مقدار خام
  final Map<String, String> extra;
  const Student({
    required this.studentNo,
    this.nationalId = '',
    this.firstName = '',
    this.lastName = '',
    this.major = '',
    this.sourceRow = 0,
    this.extra = const <String, String>{},
  });
  String get key => '$studentNo|$major';
  String get fullName => ('$firstName $lastName').trim();
  String get initial => firstName.isNotEmpty ? firstName.substring(0, 1) : '؟';
}
class ViolationRow {
  final Student student;
  final DateTime occurredAt; // UTC
  final List<String> checkedItems;
  final String note;
  final String recordedBy;
  const ViolationRow({
    required this.student,
    required this.occurredAt,
    required this.checkedItems,
    required this.note,
    required this.recordedBy,
  });
}
/// یک ستون فایل دانشگاه (سرصفحهٔ اصلی + کلید استاندارد)
class ImportCol {
  final String header;
  final String key;
  const ImportCol(this.header, this.key);
  Map<String, dynamic> toJson() => {'h': header, 'k': key};
  static ImportCol? fromJson(dynamic m) {
    if (m is! Map) return null;
    final h = '${m['h'] ?? ''}';
    final k = '${m['k'] ?? ''}';
    if (h.isEmpty || k.isEmpty) return null;
    return ImportCol(h, k);
  }
}
/// اگر فایلی آپلود نشده باشد (یا نسخهٔ قدیمی)، خروجی با این ستون‌ها ساخته می‌شود
const List<ImportCol> kDefaultColumns = [
  ImportCol('نام دانشکده', F.college),
  ImportCol('نام شخص', F.firstName),
  ImportCol('نام خانوادگي', F.lastName),
  ImportCol('نام پدر', F.fatherName),
  ImportCol('شماره شناسنامه', F.birthCertNo),
  ImportCol('كد ملي', F.nationalId),
  ImportCol('تاريخ تولد', F.birthDate),
  ImportCol('شماره دانشجویی', F.studentNo),
  ImportCol('کد تقويم', F.calendarCode),
  ImportCol('نام رشته خام', F.major),
  ImportCol('نام مقطع', F.level),
  ImportCol('نوع پذيرش', F.admissionType),
  ImportCol('فعال بودن', F.active),
  ImportCol('تعداد واحد اخذ شده نيمسال', F.semesterCredits),
  ImportCol('نام وضعيت', F.statusName),
  ImportCol('شماره موبایل', F.mobile),
  ImportCol('نیمسال ورود', F.entryTerm),
];
/// ستون‌هایی که برنامه به انتهای فایل اضافه می‌کند
const List<String> _violationHeaders = [
  'تاريخ و ساعت تخلف',
  'موارد تخلف',
  'توضيحات',
  'ثبت كننده',
];
class ImportResult {
  final List<Student> toImport;
  /// فقط وقتی ساخته می‌شود که ردیف تکراری وجود داشته باشد
  final File? reportFile;
  final int totalRows;
  final int duplicateRows;
  final List<List<String>> duplicates;
  /// ستون‌های فایل ورودی، به همان ترتیب (برای ساخت خروجی هم‌شکل)
  final List<ImportCol> columns;
  const ImportResult({
    required this.toImport,
    required this.reportFile,
    required this.totalRows,
    required this.duplicateRows,
    required this.duplicates,
    required this.columns,
  });
}
/// ───────── ابزارهای متنی ─────────
const String _faDigits = '۰۱۲۳۴۵۶۷۸۹';
const String _arDigits = '٠١٢٣٤٥٦٧٨٩';
/// یکسان‌سازی حروف عربی/فارسی و حذف فاصله‌های اضافی (برای نام ستون‌ها)
String normKey(String s) {
  var t = s
      .replaceAll('\u200c', ' ')
      .replaceAll('ي', 'ی')
      .replaceAll('ى', 'ی')
      .replaceAll('ك', 'ک')
      .replaceAll('ة', 'ه')
      .replaceAll('ۀ', 'ه')
      .replaceAll('أ', 'ا')
      .replaceAll('إ', 'ا')
      .replaceAll('آ', 'ا')
      .replaceAll('"', ' ')
      .replaceAll('«', ' ')
      .replaceAll('»', ' ')
      .replaceAll('،', ' ')
      .replaceAll(',', ' ')
      .replaceAll('_', ' ')
      .replaceAll('(', ' ')
      .replaceAll(')', ' ');
  t = t.replaceAll(RegExp(r'[\u064B-\u0652\u200e\u200f]'), '');
  t = t.replaceAll(RegExp(r'[0-9]'), ' ');
  return t.replaceAll(RegExp(r'\s+'), ' ').trim();
}
/// ارقام فارسی/عربی → لاتین، حذف فاصله‌ها (برای کدها و شماره‌ها)
String fixDigits(String s) {
  var out = s;
  for (var i = 0; i < 10; i++) {
    out = out.replaceAll(_faDigits[i], '$i').replaceAll(_arDigits[i], '$i');
  }
  return out.replaceAll(RegExp(r'\s+'), '').trim();
}
/// نگاشت سرصفحهٔ فایل دانشگاه به کلید استاندارد
const Map<String, String> _aliases = {
  'نام دانشکده': F.college,
  'دانشکده': F.college,
  'نام شخص': F.firstName,
  'نام خانوادگی': F.lastName,
  'نام پدر': F.fatherName,
  'شماره شناسنامه': F.birthCertNo,
  'کد ملی': F.nationalId,
  'شماره ملی': F.nationalId,
  'تاریخ تولد': F.birthDate,
  'محل تولد': F.birthPlace,
  'محل صدور': F.issuePlace,
  'شماره دانشجویی': F.studentNo,
  'کد دانشجویی': F.studentNo,
  'کد تقویم': F.calendarCode,
  'نام تقویم': F.calendarName,
  'نام رشته خام': F.major,
  'نام رشته': F.major,
  'رشته': F.major,
  'جنسیت': F.gender,
  'نام مقطع': F.level,
  'مقطع': F.level,
  'نام نظام': F.systemName,
  'نوع پذیرش': F.admissionType,
  'مدرک زبان فارسی': F.faLangCertificate,
  'تاییدیه مدرک زبان فارسی': F.faLangCertificateOk,
  'تاریخ شروع': F.startDate,
  'تاریخ شروع تقویم': F.startDateCal,
  'تاریخ پایان': F.endDate,
  'وضعیت تحصیلی': F.studyStatus,
  'فعال بودن': F.active,
  'تعداد واحد اخذ شده نیمسال': F.semesterCredits,
  'تعداد ترم های باقی مانده': F.termsLeft,
  'تعداد ترم های گذرانده شده': F.termsPassed,
  'کد وضعیت': F.statusCode,
  'نام وضعیت': F.statusName,
  'تاریخ شروع تحصیل': F.studyStart,
  'تاریخ پایان تحصیل': F.studyEnd,
  'تعداد روزهای گذشته': F.daysPassed,
  'شناسه شخص': F.personId,
  'تاریخ فارغ التحصیلی': F.graduationDate,
  'تاریخ ERST': F.erstDate,
  'تمدید سنوات': F.yearsExtension,
  'شماره بایگانی': F.archiveNo,
  'تاریخ تمدید': F.extensionDate,
  'شماره موبایل': F.mobile,
  'موبایل': F.mobile,
  'نیمسال ورود': F.entryTerm,
};
/// نام نمایشی هر کلید (برای سرصفحهٔ خروجی؛ مستقل از آشفتگی فایل مبدأ)
const Map<String, String> _display = {
  F.college: 'نام دانشکده',
  F.firstName: 'نام شخص',
  F.lastName: 'نام خانوادگي',
  F.fatherName: 'نام پدر',
  F.birthCertNo: 'شماره شناسنامه',
  F.nationalId: 'كد ملي',
  F.birthDate: 'تاريخ تولد',
  F.birthPlace: 'محل تولد',
  F.issuePlace: 'محل صدور',
  F.studentNo: 'شماره دانشجویی',
  F.calendarCode: 'کد تقويم',
  F.calendarName: 'نام تقويم',
  F.major: 'نام رشته خام',
  F.gender: 'جنسيت',
  F.level: 'نام مقطع',
  F.systemName: 'نام نظام',
  F.admissionType: 'نوع پذيرش',
  F.faLangCertificate: 'مدرک زبان فارسي',
  F.faLangCertificateOk: 'تاييديه مدرک زبان فارسي',
  F.startDate: 'تاريخ شروع',
  F.startDateCal: 'تاريخ شروع تقويم',
  F.endDate: 'تاريخ پايان',
  F.studyStatus: 'وضعيت تحصيلي',
  F.active: 'فعال بودن',
  F.semesterCredits: 'تعداد واحد اخذ شده نيمسال',
  F.termsLeft: 'تعداد ترم هاي باقيمانده',
  F.termsPassed: 'تعداد ترم هاي گذرانده شده',
  F.statusCode: 'کد وضعيت',
  F.statusName: 'نام وضعيت',
  F.statusName2: 'نام وضعيت "_1',
  F.studyStart: 'تاريخ شروع تحصيل',
  F.studyStart2: 'تاريخ شروع تحصيل (۲)',
  F.studyEnd: 'تاريخ پايان تحصيل',
  F.studyEnd2: 'تاريخ پايان تحصيل (۲)',
  F.daysPassed: 'تعداد روزهاي گذشته',
  F.personId: 'شناسه شخص',
  F.graduationDate: 'تاريخ فارغ التحصيلي',
  F.erstDate: 'تاريخ ERST',
  F.yearsExtension: 'تمديد سنوات',
  F.archiveNo: 'شماره بايگاني',
  F.extensionDate: 'تاريخ تمديد',
  F.mobile: 'شماره موبایل',
  F.entryTerm: 'نیمسال ورود',
};
/// ترتیب جست‌وجوی پیشوندی (بلندترین‌ها اول، تا «تاریخ شروع تحصیل» جلوتر از «تاریخ شروع» بیاید)
const List<String> _prefixOrder = [
  'تاریخ شروع تحصیل',
  'تاریخ پایان تحصیل',
  'تاریخ شروع تقویم',
  'تاریخ فارغ التحصیلی',
  'تعداد واحد اخذ شده نیمسال',
  'تعداد ترم های گذرانده شده',
  'تعداد ترم های باقی مانده',
  'تعداد روزهای گذشته',
  'تاییدیه مدرک زبان فارسی',
  'نام خانوادگی',
  'نام دانشکده',
  'نام رشته خام',
  'شماره دانشجویی',
  'شماره شناسنامه',
  'شماره موبایل',
  'شماره بایگانی',
  'تاریخ شروع',
  'تاریخ پایان',
  'تاریخ تولد',
  'تاریخ تمدید',
  'تاریخ ERST',
  'نام وضعیت',
  'نام مقطع',
  'نام نظام',
  'نام تقویم',
  'نام پدر',
  'نام شخص',
  'نوع پذیرش',
  'مدرک زبان فارسی',
  'وضعیت تحصیلی',
  'کد تقویم',
  'کد وضعیت',
  'کد ملی',
  'فعال بودن',
  'شناسه شخص',
  'نیمسال ورود',
  'محل تولد',
  'محل صدور',
  'تمدید سنوات',
  'جنسیت',
];
String? keyForHeader(String header) {
  final h = normKey(header);
  if (h.isEmpty) return null;
  final exact = _aliases[h];
  if (exact != null) return exact;
  for (final token in _prefixOrder) {
    if (h.startsWith(token)) return _aliases[token];
  }
  return null;
}
String jalali(DateTime when) {
  final l = when.toLocal();
  final j = Jalali.fromDateTime(l);
  String p(int n) => n.toString().padLeft(2, '0');
  return '${j.year}/${p(j.month)}/${p(j.day)}  ${p(l.hour)}:${p(l.minute)}';
}
String _jStamp() {
  final j = Jalali.now();
  String p(int n) => n.toString().padLeft(2, '0');
  return '${j.year}${p(j.month)}${p(j.day)}';
}
/// مقدار یک ستون برای یک دانشجو (اول از extras، بعد از فیلدهای پایه)
String valueFor(Student s, String key) {
  final v = s.extra[key];
  if (v != null && v.isNotEmpty) return v;
  switch (key) {
    case F.studentNo:
      return s.studentNo;
    case F.nationalId:
      return s.nationalId;
    case F.firstName:
      return s.firstName;
    case F.lastName:
      return s.lastName;
    case F.major:
      return s.major;
  }
  return '';
}
/// ───────── نگاشت یک ردیف سرصفحه ─────────
class _Mapped {
  final int index;
  final String header;
  final String key;
  const _Mapped(this.index, this.header, this.key);
}
/// ───────── خواندن سبک xlsx (جریانی، بدون DOM) ─────────
final RegExp _reCellRef = RegExp(r'r="([A-Z]+)\d+"');
final RegExp _reCellType = RegExp(r't="([^"]+)"');
ArchiveFile? _findEntry(Archive archive, String name) {
  for (final f in archive.files) {
    if (f.name == name) return f;
  }
  return null;
}
/// فایل‌های شیت داخل بستهٔ xlsx (sheet1.xml اول از همه)
List<ArchiveFile> _worksheetFiles(Archive archive) {
  final list = <ArchiveFile>[];
  for (final f in archive.files) {
    if (f.name.startsWith('xl/worksheets/') && f.name.endsWith('.xml')) {
      list.add(f);
    }
  }
  list.sort((a, b) => a.name.compareTo(b.name));
  final first = _findEntry(archive, 'xl/worksheets/sheet1.xml');
  if (first != null) {
    list.remove(first);
    list.insert(0, first);
  }
  return list;
}
String _entryText(ArchiveFile file) {
  final content = file.content;
  final bytes = content is Uint8List
      ? content
      : Uint8List.fromList((content as List).cast<int>());
  // allowMalformed: فایل خراب یا نویسهٔ نامعتبر باعث خطا نشود
  return utf8.decode(bytes, allowMalformed: true);
}
String _xmlUnescape(String s) {
  if (!s.contains('&')) return s;
  var out = s;
  if (out.contains('&#')) {
    out = out.replaceAllMapped(RegExp(r'&#x([0-9a-fA-F]+);'),
        (m) => String.fromCharCode(int.parse(m[1]!, radix: 16)));
    out = out.replaceAllMapped(RegExp(r'&#(\d+);'),
        (m) => String.fromCharCode(int.parse(m[1]!)));
  }
  return out
      .replaceAll('&lt;', '<')
      .replaceAll('&gt;', '>')
      .replaceAll('&quot;', '"')
      .replaceAll('&apos;', "'")
      .replaceAll('&amp;', '&');
}
/// همهٔ رشته‌های مشترک (xl/sharedStrings.xml) را به لیست تبدیل می‌کند
List<String> _parseSharedStrings(String xml) {
  final out = <String>[];
  var pos = 0;
  while (true) {
    var start = xml.indexOf('<si>', pos);
    final startAttr = xml.indexOf('<si ', pos);
    if (start < 0 || (startAttr >= 0 && startAttr < start)) start = startAttr;
    if (start < 0) break;
    final end = xml.indexOf('</si>', start);
    if (end < 0) break;
    final chunk = xml.substring(start, end);
    final buf = StringBuffer();
    var j = 0;
    while (true) {
      final t1 = chunk.indexOf('<t', j);
      if (t1 < 0) break;
      final gt = chunk.indexOf('>', t1);
      if (gt < 0) break;
      if (chunk[gt - 1] == '/') {
        j = gt + 1;
        continue;
      }
      final t2 = chunk.indexOf('</t>', gt);
      if (t2 < 0) break;
      buf.write(chunk.substring(gt + 1, t2));
      j = t2 + 4;
    }
    out.add(_xmlUnescape(buf.toString()));
    pos = end + 5;
  }
  return out;
}
/// شمارهٔ آخرین ردیف از <dimension ref="A1:Q3000"/>
int _lastRowNumber(String xml) {
  final i = xml.indexOf('<dimension ref="');
  if (i < 0) return 0;
  final start = i + 16;
  final end = xml.indexOf('"', start);
  if (end < 0) return 0;
  final ref = xml.substring(start, end);
  final digits = RegExp(r'(\d+)').allMatches(ref).toList();
  if (digits.isEmpty) return 0;
  return int.tryParse(digits.last[1]!) ?? 0;
}
/// ردیف‌ها یکی‌یکی تولید می‌شوند تا کل جدول در حافظه نماند
Iterable<List<String>> _rowsFromSheet(String xml, List<String> shared) sync* {
  var pos = 0;
  while (true) {
    final rs = xml.indexOf('<row', pos);
    if (rs < 0) break;
    final gt = xml.indexOf('>', rs);
    if (gt < 0) break;
    if (xml[gt - 1] == '/') {
      yield <String>[];
      pos = gt + 1;
      continue;
    }
    final re = xml.indexOf('</row>', gt);
    if (re < 0) break;
    yield _parseRow(xml.substring(gt + 1, re), shared);
    pos = re + 6;
  }
}
List<String> _parseRow(String rowXml, List<String> shared) {
  final cells = <String>[];
  var pos = 0;
  while (true) {
    final cs = rowXml.indexOf('<c', pos);
    if (cs < 0) break;
    final gt = rowXml.indexOf('>', cs);
    if (gt < 0) break;
    final attrs = rowXml.substring(cs + 2, gt);
    final selfClose = attrs.endsWith('/');
    final refMatch = _reCellRef.firstMatch(attrs);
    final col = refMatch == null ? cells.length : _colIndex(refMatch[1]!);
    var value = '';
    if (selfClose) {
      pos = gt + 1;
    } else {
      final ce = rowXml.indexOf('</c>', gt);
      if (ce < 0) break;
      value = _cellValue(rowXml.substring(gt + 1, ce), attrs, shared);
      pos = ce + 4;
    }
    while (cells.length <= col) {
      cells.add('');
    }
    cells[col] = value;
  }
  return cells;
}
String _cellValue(String inner, String attrs, List<String> shared) {
  final type = _reCellType.firstMatch(attrs)?[1];
  if (type == 'inlineStr') {
    final tStart = inner.indexOf('<t');
    if (tStart < 0) return '';
    final gt = inner.indexOf('>', tStart);
    if (gt < 0) return '';
    final te = inner.indexOf('</t>', gt);
    if (te < 0) return '';
    return _xmlUnescape(inner.substring(gt + 1, te));
  }
  final vStart = inner.indexOf('<v>');
  if (vStart < 0) return '';
  final vEnd = inner.indexOf('</v>', vStart);
  if (vEnd < 0) return '';
  final raw = inner.substring(vStart + 3, vEnd);
  if (type == 's') {
    final idx = int.tryParse(raw.trim());
    if (idx == null || idx < 0 || idx >= shared.length) return '';
    return shared[idx];
  }
  if (type == 'b') return raw.trim() == '1' ? 'TRUE' : 'FALSE';
  return _xmlUnescape(raw).trim();
}
/// ستون حرفی (A, B, ..., AA) → شمارهٔ صفر-مبنا
int _colIndex(String letters) {
  var n = 0;
  for (var i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.codeUnitAt(i) - 64);
  }
  return n - 1;
}
/// ───────── سرویس اکسل ─────────
class ExcelService {
  /// ۱) خواندن فایل دانشجویان (xlsx یا csv)، حذف تکراری‌ها، گزارش تکراری‌ها
  /// [onProgress] برای نمایش پیشرفت در فایل‌های بزرگ (مثل ۳۰ هزار ردیف)
  Future<ImportResult> importStudents(
    String path, {
    void Function(int done, int total)? onProgress,
  }) async {
    final lower = path.toLowerCase();
    if (lower.endsWith('.csv') || lower.endsWith('.txt')) {
      return _importDelimited(path, onProgress);
    }
    return _importXlsx(path, onProgress);
  }
  /// مسیر اکسل: فایل به‌صورت جریانی (بدون ساخت DOM) خوانده می‌شود.
  /// با این روش ۳۰ هزار ردیف هم چند ده مگابایت حافظه می‌خواهد، نه چند صد.
  Future<ImportResult> _importXlsx(
    String path,
    void Function(int done, int total)? onProgress,
  ) async {
    final bytes = await File(path).readAsBytes();
    final archive = ZipDecoder().decodeBytes(bytes);
    final ssFile = _findEntry(archive, 'xl/sharedStrings.xml');
    final shared = ssFile == null
        ? const <String>[]
        : _parseSharedStrings(_entryText(ssFile));
    final sheets = _worksheetFiles(archive);
    if (sheets.isEmpty) throw Exception('شیت اکسل در فایل پیدا نشد');
    var firstSample = '';
    for (final sheetFile in sheets) {
      final xml = _entryText(sheetFile);
      if (!xml.contains('<row')) continue;
      final it = _rowsFromSheet(xml, shared).iterator;
      // سرصفحه در یکی از ۸ ردیف اول جست‌وجو می‌شود
      final head = <List<String>>[];
      while (head.length < 8 && it.moveNext()) {
        head.add(it.current);
      }
      if (head.isEmpty) continue;
      if (firstSample.isEmpty) firstSample = head.first.take(12).join(' | ');
      var headerRow = 0;
      var best = <_Mapped>[];
      for (var i = 0; i < head.length; i++) {
        final cols = _mapHeaders(head[i]);
        if (cols.length > best.length) {
          best = cols;
          headerRow = i;
        }
      }
      // این شیت باید هم سرصفحه‌های شناخته‌شده داشته باشد، هم شماره دانشجویی
      if (best.length < 2) continue;
      if (!best.any((c) => c.key == F.studentNo)) continue;
      Iterable<List<String>> dataRows() sync* {
        for (var i = headerRow + 1; i < head.length; i++) {
          yield head[i];
        }
        while (it.moveNext()) {
          yield it.current;
        }
      }
      final total = _lastRowNumber(xml);
      final count = total > headerRow ? total - headerRow - 1 : 0;
      return _assemble(head[headerRow], dataRows(), count, onProgress);
    }
    throw Exception('سرصفحه‌های فایل شناخته نشد. ردیف اول: $firstSample');
  }
  /// مسیر CSV/TSV: مصرف حافظه بسیار کمتر، مناسب فایل‌های چندهزار ردیفی
  Future<ImportResult> _importDelimited(
    String path,
    void Function(int done, int total)? onProgress,
  ) async {
    final bytes = await File(path).readAsBytes();
    String text;
    try {
      text = utf8.decode(bytes);
    } catch (_) {
      throw Exception('فایل CSV با UTF-8 ذخیره نشده است. '
          'در اکسل گزینهٔ «CSV UTF-8» را انتخاب کنید.');
    }
    if (text.startsWith('\uFEFF')) text = text.substring(1);
    final rows = _parseDelimited(text);
    if (rows.isEmpty) throw Exception('فایل خالی است');
    var headerRow = 0;
    var best = <_Mapped>[];
    final limit = rows.length < 8 ? rows.length : 8;
    for (var i = 0; i < limit; i++) {
      final cols = _mapHeaders(rows[i]);
      if (cols.length > best.length) {
        best = cols;
        headerRow = i;
      }
    }
    return _assemble(rows[headerRow], rows.skip(headerRow + 1),
        rows.length - headerRow - 1, onProgress);
  }
  /// منطق مشترک هر دو مسیر: ساخت دانشجوها، حذف تکراری‌ها، گزارش تکراری‌ها
  Future<ImportResult> _assemble(
    List<String> headerCells,
    Iterable<List<String>> dataRows,
    int total,
    void Function(int done, int total)? onProgress,
  ) async {
    final cols = _mapHeaders(headerCells);
    if (cols.length < 2) {
      throw Exception('سرصفحه‌های فایل شناخته نشد. '
          'ردیف اول: ${headerCells.take(12).join(' | ')}');
    }
    final colOf = <String, int>{}; // key → column index
    for (final c in cols) {
      colOf.putIfAbsent(c.key, () => c.index);
    }
    if (!colOf.containsKey(F.studentNo)) {
      throw Exception('ستون «شماره دانشجویی» پیدا نشد. '
          'سرصفحه‌ها: ${headerCells.take(12).join(' | ')}');
    }
    final unique = <String, Student>{};
    final dupRows = <List<String>>[];
    var scanned = 0;
    var withNo = 0;
    for (final cells in dataRows) {
      scanned++;
      final no = fixDigits(_cellAt(cells, colOf[F.studentNo]));
      if (no.isEmpty) {
        if (onProgress != null && scanned % 500 == 0) onProgress(scanned, total);
        continue;
      }
      withNo++;
      final extra = <String, String>{};
      colOf.forEach((key, index) {
        if (key == F.studentNo) return;
        final v = _cellAt(cells, index);
        if (v.isNotEmpty) extra[key] = v;
      });
      final s = Student(
        studentNo: no,
        nationalId: fixDigits(extra[F.nationalId] ?? ''),
        firstName: extra[F.firstName] ?? '',
        lastName: extra[F.lastName] ?? '',
        major: extra[F.major] ?? '',
        sourceRow: scanned,
        extra: extra,
      );
      // تکراری‌ها: همان شماره دانشجویی با همان رشته → رد می‌شود
      if (unique.containsKey(s.key)) {
        dupRows.add([
          s.studentNo,
          s.nationalId,
          s.firstName,
          s.lastName,
          s.major,
          '${s.sourceRow}',
          'ردیف تکراری (وارد نشد)',
        ]);
      } else {
        unique[s.key] = s;
      }
      if (onProgress != null && scanned % 500 == 0) onProgress(scanned, total);
    }
    final imported = unique.values.toList();
    // گزارش فقط برای ردیف‌های تکراری ساخته می‌شود (نه برای همهٔ ۳۰ هزار ردیف)
    final report = dupRows.isEmpty ? null : await _writeReport(dupRows);
    return ImportResult(
      toImport: imported,
      reportFile: report,
      totalRows: withNo,
      duplicateRows: dupRows.length,
      duplicates: dupRows,
      columns: cols.map((c) => ImportCol(c.header, c.key)).toList(),
    );
  }
  /// خواندن CSV/TSV؛ جداکننده (کاما، نقطه‌ویرگول یا تب) خودکار تشخیص داده می‌شود
  List<List<String>> _parseDelimited(String text) {
    final firstLine = text.split('\n').first;
    final commas = ','.allMatches(firstLine).length;
    final semis = ';'.allMatches(firstLine).length;
    final tabs = '\t'.allMatches(firstLine).length;
    final delim = (tabs >= commas && tabs >= semis)
        ? '\t'
        : (semis > commas ? ';' : ',');
    final rows = <List<String>>[];
    var row = <String>[];
    final buf = StringBuffer();
    var inQuotes = false;
    for (var i = 0; i < text.length; i++) {
      final ch = text[i];
      if (inQuotes) {
        if (ch == '"') {
          if (i + 1 < text.length && text[i + 1] == '"') {
            buf.write('"');
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          buf.write(ch);
        }
        continue;
      }
      if (ch == '"') {
        inQuotes = true;
      } else if (ch == delim) {
        row.add(buf.toString().trim());
        buf.clear();
      } else if (ch == '\n') {
        row.add(buf.toString().trim());
        buf.clear();
        rows.add(row);
        row = <String>[];
      } else if (ch != '\r') {
        buf.write(ch);
      }
    }
    if (buf.isNotEmpty || row.isNotEmpty) {
      row.add(buf.toString().trim());
      rows.add(row);
    }
    return rows;
  }
  List<_Mapped> _mapHeaders(List<String> cells) {
    final list = <_Mapped>[];
    final used = <String>{};
    for (var i = 0; i < cells.length; i++) {
      final raw = cells[i].trim();
      if (raw.isEmpty) continue;
      var key = keyForHeader(raw);
      if (key == null) continue;
      if (used.contains(key)) {
        // ستون تکراری در فایل (مثل «نام وضعیت» و «نام وضعیت _1»)
        final alt = _duplicateKeyOf(key);
        if (alt == null || used.contains(alt)) continue;
        key = alt;
      }
      used.add(key);
      list.add(_Mapped(i, _display[key] ?? raw, key));
    }
    return list;
  }
  String? _duplicateKeyOf(String key) {
    switch (key) {
      case F.statusName:
        return F.statusName2;
      case F.studyStart:
        return F.studyStart2;
      case F.studyEnd:
        return F.studyEnd2;
    }
    return null;
  }
  String _cellAt(List<String> cells, int? index) {
    if (index == null || index < 0 || index >= cells.length) return '';
    return cells[index].trim();
  }
  /// ۲) خروجی تخلفات در یک شیت، هم‌شکل فایل ورودی + ستون‌های تخلف
  Future<File> exportViolations(
    List<ViolationRow> rows, {
    List<ImportCol> columns = const [],
  }) async {
    final cols = columns.isEmpty ? kDefaultColumns : columns;
    final ex = Excel.createExcel();
    final sheet = ex['تخلفات'];
    sheet.isRTL = true;
    sheet.appendRow([
      ...cols.map((c) => TextCellValue(c.header)),
      ..._violationHeaders.map((h) => TextCellValue(h)),
    ]);
    for (final v in rows) {
      sheet.appendRow([
        ...cols.map((c) => TextCellValue(valueFor(v.student, c.key))),
        TextCellValue(jalali(v.occurredAt)),
        TextCellValue(v.checkedItems.join(' ، ')),
        TextCellValue(v.note),
        TextCellValue(v.recordedBy),
      ]);
    }
    ex.delete('Sheet1');
    return _save(ex, 'گزارش_تخلفات');
  }
  /// گزارش کوچک و سریع: فقط ردیف‌های تکراری
  Future<File> _writeReport(List<List<String>> dups) async {
    final ex = Excel.createExcel();
    final sheet = ex['ردیف‌های تکراری'];
    sheet.isRTL = true;
    sheet.appendRow([
      TextCellValue('شماره دانشجویی'),
      TextCellValue('کد ملی'),
      TextCellValue('نام'),
      TextCellValue('نام خانوادگی'),
      TextCellValue('رشته'),
      TextCellValue('ردیف در فایل'),
      TextCellValue('وضعیت'),
    ]);
    for (final d in dups) {
      sheet.appendRow([
        TextCellValue(d[0]),
        TextCellValue(d[1]),
        TextCellValue(d[2]),
        TextCellValue(d[3]),
        TextCellValue(d[4]),
        TextCellValue(d[5]),
        TextCellValue(d[6]),
      ]);
    }
    ex.delete('Sheet1');
    return _save(ex, 'گزارش_آپلود');
  }
  Future<File> _save(Excel ex, String name) async {
    final dir = await getApplicationDocumentsDirectory();
    final file = File('${dir.path}/${name}_${_jStamp()}.xlsx');
    await file.writeAsBytes(ex.encode()!);
    return file;
  }
  /// اشتراک‌گذاری فایل (ذخیره، واتساپ، ایمیل، ...)
  Future<void> share(File f) => Share.shareXFiles([XFile(f.path)]);
}
