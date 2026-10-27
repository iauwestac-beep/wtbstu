// lib/excel_service.dart
//
// ورود/خروجی اکسل برای برنامهٔ «ثبت تخلف پوشش»
// ─ ورود: فایل دانشجویان دانشگاه (۱۸ ستون، با ت Matching نرم سرصفحه‌ها)
// ─ خروجی: یک شیت واحد «تخلفات» با ستون‌های استاندارد دانشگاه
import 'dart:io';

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
  String get initial =>
      firstName.isNotEmpty ? firstName.substring(0, 1) : '؟';
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

class ImportResult {
  final List<Student> toImport;
  final File reportFile;
  final int totalRows;
  final int duplicateRows;
  final List<List<String>> duplicates;

  const ImportResult({
    required this.toImport,
    required this.reportFile,
    required this.totalRows,
    required this.duplicateRows,
    required this.duplicates,
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
  t = t.replaceAll('  ', ' ');
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

String _text(Data? d) {
  final v = d?.value;
  if (v == null) return '';
  if (v is TextCellValue) return v.value.toString().trim();
  if (v is IntCellValue) return v.value.toString();
  if (v is DoubleCellValue) {
    final x = v.value;
    return x == x.truncateToDouble() ? x.toInt().toString() : x.toString();
  }
  return v.toString().trim();
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

/// ───────── ستون‌های شیت خروجی ─────────
class _Col {
  final String header;
  final String? core; // first | last | no | nid | major
  final String? key; // کلید در Student.extra
  const _Col(this.header, {this.core, this.key});
}

const List<_Col> _exportColumns = [
  _Col('نام شخص', core: 'first'),
  _Col('نام خانوادگي', core: 'last'),
  _Col('نام پدر', key: F.fatherName),
  _Col('شماره شناسنامه', key: F.birthCertNo),
  _Col('كد ملي', core: 'nid'),
  _Col('محل تولد', key: F.birthPlace),
  _Col('محل صدور', key: F.issuePlace),
  _Col('تاريخ تولد', key: F.birthDate),
  _Col('شماره دانشجویی', core: 'no'),
  _Col('کد تقويم', key: F.calendarCode),
  _Col('نام تقويم', key: F.calendarName),
  _Col('خالی'),
  _Col('نام رشته خام', core: 'major'),
  _Col('جنسيت', key: F.gender),
  _Col('نام مقطع', key: F.level),
  _Col('نام نظام', key: F.systemName),
  _Col('نوع پذيرش', key: F.admissionType),
  _Col('مدرک زبان فارسي', key: F.faLangCertificate),
  _Col('تاييديه مدرک زبان فارسي', key: F.faLangCertificateOk),
  _Col('تاريخ شروع', key: F.startDate),
  _Col('تاريخ شروع تقويم', key: F.startDateCal),
  _Col('تاريخ پايان', key: F.endDate),
  _Col('وضعيت تحصيلي', key: F.studyStatus),
  _Col('فعال بودن', key: F.active),
  _Col('تعداد واحد اخذ شده نيمسال', key: F.semesterCredits),
  _Col('تعداد ترم هاي باقيمانده', key: F.termsLeft),
  _Col('تعداد ترم هاي گذرانده شده', key: F.termsPassed),
  _Col('کد وضعيت', key: F.statusCode),
  _Col('نام وضعيت', key: F.statusName),
  _Col('نام وضعيت "_1', key: F.statusName2),
  _Col('تاريخ شروع تحصيل', key: F.studyStart),
  _Col('تاريخ شروع تحصيل ', key: F.studyStart2),
  _Col('تاريخ پايان تحصيل', key: F.studyEnd),
  _Col('تاريخ پايان تحصيل ', key: F.studyEnd2),
  _Col('تعداد روزهاي گذشته', key: F.daysPassed),
  _Col('شناسه شخص', key: F.personId),
  _Col('تاريخ فارغ التحصيلي', key: F.graduationDate),
  _Col('تاريخ ERST', key: F.erstDate),
  _Col('تمديد سنوات', key: F.yearsExtension),
  _Col('شماره بايگاني', key: F.archiveNo),
  _Col('تاريخ تمديد', key: F.extensionDate),
  _Col('شماره موبایل', key: F.mobile),
];

const List<String> _violationHeaders = [
  'تاريخ و ساعت تخلف',
  'موارد تخلف',
  'توضيحات',
  'ثبت كننده',
];

String _valueOf(Student s, _Col c) {
  switch (c.core) {
    case 'first':
      return s.firstName;
    case 'last':
      return s.lastName;
    case 'no':
      return s.studentNo;
    case 'nid':
      return s.nationalId;
    case 'major':
      return s.major;
  }
  final k = c.key;
  if (k == null) return '';
  return (s.extra[k] ?? '').trim();
}

/// ───────── سرویس اکسل ─────────
class ExcelService {
  /// ۱) خواندن فایل دانشجویان، حذف تکراری‌ها، ساخت گزارش آپلود (تک‌شیت)
  Future<ImportResult> importStudents(String path) async {
    final excel = Excel.decodeBytes(await File(path).readAsBytes());

    // بهترین (شیت، ردیف سرصفحه) را پیدا می‌کند؛ تحمل ردیف عنوان بالای جدول
    Sheet? sheet;
    var headerRow = 0;
    Map<int, String> bestMap = {};
    for (final candidate in excel.tables.values) {
      final limit = candidate.rows.length < 8 ? candidate.rows.length : 8;
      for (var r = 0; r < limit; r++) {
        final map = _mapHeaders(candidate.rows[r]);
        if (map.length > bestMap.length) {
          bestMap = map;
          sheet = candidate;
          headerRow = r;
        }
      }
    }
    if (sheet == null || bestMap.length < 2) {
      // پیام خطا شامل سرصفحه‌های واقعی فایل است تا عیب‌یابی ساده باشد
      final sample = excel.tables.values
          .where((s) => s.rows.isNotEmpty)
          .map((s) => s.rows.first.take(12).map(_text).join(' | '))
          .join('  ||  ');
      throw Exception('سرصفحه‌های فایل شناخته نشد. ردیف اول: $sample');
    }
    final colOf = <String, int>{}; // key → column index
    bestMap.forEach((index, key) => colOf.putIfAbsent(key, () => index));
    if (!colOf.containsKey(F.studentNo)) {
      final sample = sheet.rows[headerRow].take(12).map(_text).join(' | ');
      throw Exception('ستون «شماره دانشجویی» پیدا نشد. سرصفحه‌ها: $sample');
    }

    final all = <Student>[];
    for (var r = headerRow + 1; r < sheet.rows.length; r++) {
      final row = sheet.rows[r];
      final no = fixDigits(_text(_at(row, colOf[F.studentNo])));
      if (no.isEmpty) continue;

      final extra = <String, String>{};
      colOf.forEach((key, index) {
        if (key == F.studentNo) return;
        final v = _text(_at(row, index));
        if (v.isNotEmpty) extra[key] = v;
      });

      all.add(Student(
        studentNo: no,
        nationalId: fixDigits(extra[F.nationalId] ?? ''),
        firstName: extra[F.firstName] ?? '',
        lastName: extra[F.lastName] ?? '',
        major: extra[F.major] ?? '',
        sourceRow: r + 1,
        extra: extra,
      ));
    }

    // تکراری‌ها: همان شماره دانشجویی با همان رشته → رد می‌شود
    final unique = <String, Student>{};
    final dupRows = <List<String>>[];
    for (final s in all) {
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
    }

    final report = await _writeReport(unique.values.toList(), dupRows);
    return ImportResult(
      toImport: unique.values.toList(),
      reportFile: report,
      totalRows: all.length,
      duplicateRows: dupRows.length,
      duplicates: dupRows,
    );
  }

  Map<int, String> _mapHeaders(List<Data?> headerRow) {
    final map = <int, String>{};
    final used = <String>{};
    for (var i = 0; i < headerRow.length; i++) {
      final raw = _text(headerRow[i]);
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
      map[i] = key;
    }
    return map;
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

  Data? _at(List<Data?> row, int? index) {
    if (index == null || index < 0 || index >= row.length) return null;
    return row[index];
  }

  /// ۲) خروجی تخلفات در یک شیت
  Future<File> exportViolations(List<ViolationRow> rows) async {
    final ex = Excel.createExcel();
    final sheet = ex['تخلفات'];
    sheet.isRTL = true;

    sheet.appendRow([
      ..._exportColumns.map((c) => TextCellValue(c.header)),
      ..._violationHeaders.map((h) => TextCellValue(h)),
    ]);

    for (final v in rows) {
      sheet.appendRow([
        ..._exportColumns.map((c) => TextCellValue(_valueOf(v.student, c))),
        TextCellValue(jalali(v.occurredAt)),
        TextCellValue(v.checkedItems.join(' ، ')),
        TextCellValue(v.note),
        TextCellValue(v.recordedBy),
      ]);
    }

    ex.delete('Sheet1');
    return _save(ex, 'گزارش_تخلفات');
  }

  Future<File> _writeReport(List<Student> imported, List<List<String>> dups) async {
    final ex = Excel.createExcel();
    final sheet = ex['گزارش آپلود'];
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
    for (final s in imported) {
      sheet.appendRow([
        TextCellValue(s.studentNo),
        TextCellValue(s.nationalId),
        TextCellValue(s.firstName),
        TextCellValue(s.lastName),
        TextCellValue(s.major),
        IntCellValue(s.sourceRow),
        TextCellValue('وارد شد'),
      ]);
    }
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
