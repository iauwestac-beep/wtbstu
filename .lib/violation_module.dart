
// lib/violation_module.dart
//
// دیتابیس محلی رمزنگاری‌شده + صفحهٔ جستجو + فرم ثبت تخلف
import 'dart:convert';

import 'package:flutter/material.dart';
import 'package:path/path.dart' as p;
import 'package:sqflite_sqlcipher/sqflite.dart';
import 'package:uuid/uuid.dart';

import 'excel_service.dart';

/// ۶ مورد چک‌باکس: عنوان‌ها را با موارد واقعی عوض کنید
const kViolationItems = [
  'لباس کوتاه یا چسبان',
  'لباس نازک یا بدن‌نما',
  'شلوار پاره یا کوتاه',
  'نوشته یا تصویر نامناسب روی لباس',
  'دمپایی یا صندل',
  'آرایش یا اکسسوری نامتعارف',
];

/// یکسان‌سازی متن برای جست‌وجو (حروف عربی → فارسی، ارقام → لاتین)
String searchNorm(String s) {
  var t = s
      .replaceAll('\u200c', ' ')
      .replaceAll('ي', 'ی')
      .replaceAll('ك', 'ک')
      .replaceAll('ى', 'ی');
  const fa = '۰۱۲۳۴۵۶۷۸۹';
  const ar = '٠١٢٣٤٥٦٧٨٩';
  for (var i = 0; i < 10; i++) {
    t = t.replaceAll(fa[i], '$i').replaceAll(ar[i], '$i');
  }
  return t.replaceAll(RegExp(r'\s+'), ' ').trim();
}

Map<String, String> _decodeExtras(Object? raw) {
  if (raw is! String || raw.isEmpty) return const {};
  try {
    final m = jsonDecode(raw);
    if (m is Map) {
      return m.map((k, v) => MapEntry('$k', '$v'));
    }
  } catch (_) {}
  return const {};
}

/// نتیجهٔ جست‌وجو: دانشجو + تعداد تخلف‌های ثبت‌شده‌اش
class StudentHit {
  final Student student;
  final int violations;
  const StudentHit(this.student, this.violations);
}

/// ───────── دیتابیس محلی (رمزنگاری‌شده) ─────────
class LocalDb {
  static Database? _db;
  static const _version = 3;

  static Future<Database> open(String key) async {
    if (_db != null) return _db!;
    final path = p.join(await getDatabasesPath(), 'dress.db');
    _db = await openDatabase(
      path,
      password: key,
      version: _version,
      onCreate: (d, _) async {
        await d.execute('''
          CREATE TABLE students(
            key TEXT PRIMARY KEY,
            student_no TEXT,
            national_id TEXT,
            first_name TEXT,
            last_name TEXT,
            major TEXT,
            extras TEXT,
            search_text TEXT,
            source_row INTEGER DEFAULT 0)''');
        await d.execute('CREATE INDEX idx_search ON students(search_text)');
        await d.execute('''
          CREATE TABLE violations(
            id TEXT PRIMARY KEY,
            student_key TEXT,
            occurred_at TEXT,
            items TEXT,
            note TEXT,
            recorded_by TEXT,
            synced INTEGER DEFAULT 0)''');
        await d.execute('CREATE INDEX idx_v_student ON violations(student_key)');
        await d.execute('CREATE TABLE meta(k TEXT PRIMARY KEY, v TEXT)');
      },
      onUpgrade: (d, oldVersion, newVersion) async {
        // افزودن ستون‌های دانشگاه (نسخهٔ ۱ → ۲)
        final existing = (await d.rawQuery('PRAGMA table_info(students)'))
            .map((r) => '${r['name']}')
            .toSet();
        if (!existing.contains('extras')) {
          await d.execute('ALTER TABLE students ADD COLUMN extras TEXT');
        }
        if (!existing.contains('source_row')) {
          await d.execute(
              'ALTER TABLE students ADD COLUMN source_row INTEGER DEFAULT 0');
        }
        // نگهداری ترتیب ستون‌های فایل ورودی (نسخهٔ ۲ → ۳)
        await d.execute(
            'CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)');
      },
    );
    return _db!;
  }

  static Database get _d => _db!;

  /// ذخیره/به‌روزرسانی دانشجوها (سابقهٔ تخلف دست نمی‌خورد)
  /// در بسته‌های ۴۰۰ تایی نوشته می‌شود تا با ۳۰ هزار دانشجو هم روان بماند
  static Future<void> upsertStudents(
    List<Student> list, {
    void Function(int done, int total)? onProgress,
  }) async {
    const chunk = 400;
    for (var i = 0; i < list.length; i += chunk) {
      final end = (i + chunk) > list.length ? list.length : i + chunk;
      final b = _d.batch();
      for (var j = i; j < end; j++) {
        final s = list[j];
        b.insert(
          'students',
          {
            'key': s.key,
            'student_no': s.studentNo,
            'national_id': s.nationalId,
            'first_name': s.firstName,
            'last_name': s.lastName,
            'major': s.major,
            'extras': jsonEncode(s.extra),
            'search_text': searchNorm(
                '${s.studentNo} ${s.nationalId} ${s.firstName} ${s.lastName}'),
            'source_row': s.sourceRow,
          },
          conflictAlgorithm: ConflictAlgorithm.replace,
        );
      }
      await b.commit(noResult: true);
      onProgress?.call(end, list.length);
    }
  }

  /// ذخیرهٔ ترتیب ستون‌های فایل ورودی (تا خروجی هم‌شکل همان فایل باشد)
  static Future<void> saveColumns(List<ImportCol> cols) async {
    await _d.insert(
      'meta',
      {'k': 'columns', 'v': jsonEncode(cols.map((c) => c.toJson()).toList())},
      conflictAlgorithm: ConflictAlgorithm.replace,
    );
  }

  /// ستون‌های ذخیره‌شده؛ اگر نبود، فهرست خالی (و خروجی با ستون‌های پیش‌فرض ساخته می‌شود)
  static Future<List<ImportCol>> loadColumns() async {
    final rows = await _d.query('meta', where: 'k=?', whereArgs: ['columns']);
    if (rows.isEmpty) return const [];
    try {
      final list = jsonDecode('${rows.first['v']}') as List;
      return list.map(ImportCol.fromJson).whereType<ImportCol>().toList();
    } catch (_) {
      return const [];
    }
  }

  static Student _student(Map<String, Object?> r) => Student(
        studentNo: '${r['student_no'] ?? ''}',
        nationalId: '${r['national_id'] ?? ''}',
        firstName: '${r['first_name'] ?? ''}',
        lastName: '${r['last_name'] ?? ''}',
        major: '${r['major'] ?? ''}',
        sourceRow: (r['source_row'] as int?) ?? 0,
        extra: _decodeExtras(r['extras']),
      );

  /// جست‌وجو: هر کلمه باید در کد دانشجویی، کد ملی یا نام وجود داشته باشد
  /// همراه با تعداد تخلف‌های ثبت‌شدهٔ هر دانشجو (یک کوئری، بدون کندی)
  static Future<List<StudentHit>> search(String q) async {
    final tokens = searchNorm(q).split(' ').where((t) => t.isNotEmpty).toList();
    if (tokens.isEmpty) return [];
    final where =
        tokens.map((_) => 's.search_text LIKE ?').join(' AND ');
    final rows = await _d.rawQuery('''
      SELECT s.*,
             (SELECT COUNT(*) FROM violations v
               WHERE v.student_key = s.key) AS vcount
      FROM students s
      WHERE $where
      ORDER BY s.last_name, s.first_name
      LIMIT 50''', tokens.map((t) => '%$t%').toList());
    return rows
        .map((r) => StudentHit(_student(r), (r['vcount'] as int?) ?? 0))
        .toList();
  }

  static Future<int> violationCount(String studentKey) async =>
      Sqflite.firstIntValue(await _d.rawQuery(
          'SELECT COUNT(*) FROM violations WHERE student_key=?',
          [studentKey])) ??
      0;

  static Future<void> addViolation(Student s, DateTime when, List<String> items,
      String note, String by) async {
    await _d.insert('violations', {
      'id': const Uuid().v4(),
      'student_key': s.key,
      'occurred_at': when.toUtc().toIso8601String(),
      'items': jsonEncode(items),
      'note': note,
      'recorded_by': by,
      'synced': 0,
    });
  }

  /// برای خروجی اکسل (هر تخلف یک ردیف)
  static Future<List<ViolationRow>> allViolations() async {
    final rows = await _d.rawQuery('''
      SELECT v.*, s.student_no, s.national_id, s.first_name, s.last_name,
             s.major, s.extras, s.source_row
      FROM violations v JOIN students s ON s.key = v.student_key
      ORDER BY v.occurred_at DESC''');
    return rows.map((r) {
      return ViolationRow(
        student: _student(r),
        occurredAt: DateTime.parse(r['occurred_at'] as String),
        checkedItems: List<String>.from(
            jsonDecode('${r['items'] ?? '[]'}') as List),
        note: '${r['note'] ?? ''}',
        recordedBy: '${r['recorded_by'] ?? ''}',
      );
    }).toList();
  }

  /// سوابق تخلف یک دانشجو (تازه‌ترین اول)
  static Future<List<ViolationRow>> violationsOf(Student s) async {
    final rows = await _d.query(
      'violations',
      where: 'student_key=?',
      whereArgs: [s.key],
      orderBy: 'occurred_at DESC',
    );
    return rows.map((r) {
      return ViolationRow(
        student: s,
        occurredAt: DateTime.parse('${r['occurred_at']}'),
        checkedItems: List<String>.from(
            jsonDecode('${r['items'] ?? '[]'}') as List),
        note: '${r['note'] ?? ''}',
        recordedBy: '${r['recorded_by'] ?? ''}',
      );
    }).toList();
  }

  static Future<int> studentCount() async =>
      Sqflite.firstIntValue(
          await _d.rawQuery('SELECT COUNT(*) FROM students')) ??
      0;

  static Future<int> violationTotal() async =>
      Sqflite.firstIntValue(
          await _d.rawQuery('SELECT COUNT(*) FROM violations')) ??
      0;

  /// پشتیبان‌گیری دستی
  static Future<List<Map<String, Object?>>> rawViolations() =>
      _d.query('violations');

  /// ادغام تخلف‌های گوشی دیگر؛ شناسهٔ (UUID) تکراری نادیده گرفته می‌شود
  static Future<int> importViolations(List<dynamic> rows) async {
    final have = (await _d.query('violations', columns: ['id']))
        .map((r) => '${r['id']}')
        .toSet();
    final b = _d.batch();
    var added = 0;
    for (final x in rows) {
      final id = '${x['id']}';
      if (have.contains(id)) continue;
      have.add(id);
      added++;
      b.insert('violations', {
        'id': id,
        'student_key': '${x['student_key']}',
        'occurred_at': '${x['occurred_at']}',
        'items': '${x['items']}',
        'note': '${x['note'] ?? ''}',
        'recorded_by': '${x['recorded_by'] ?? ''}',
        'synced': 0,
      });
    }
    await b.commit(noResult: true);
    return added;
  }

  static Future<List<Map<String, Object?>>> unsynced() =>
      _d.query('violations', where: 'synced=0');

  static Future<void> markSynced(String id) =>
      _d.update('violations', {'synced': 1}, where: 'id=?', whereArgs: [id]);
}

/// ───────── اجزای مشترک ظاهری ─────────
class SectionCard extends StatelessWidget {
  final Widget child;
  final EdgeInsets padding;
  const SectionCard({
    super.key,
    required this.child,
    this.padding = const EdgeInsets.all(16),
  });

  @override
  Widget build(BuildContext context) => Card(
        elevation: 0,
        margin: EdgeInsets.zero,
        color: Theme.of(context).colorScheme.surfaceContainerHighest.withValues(alpha: 0.45),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
        child: Padding(padding: padding, child: child),
      );
}

/// ───────── صفحهٔ سوابق یک دانشجو ─────────
class StudentDetailScreen extends StatefulWidget {
  final Student student;
  final String currentUser;
  const StudentDetailScreen({
    super.key,
    required this.student,
    required this.currentUser,
  });

  @override
  State<StudentDetailScreen> createState() => _StudentDetailScreenState();
}

class _StudentDetailScreenState extends State<StudentDetailScreen> {
  List<ViolationRow> _rows = [];
  bool _loading = true;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final rows = await LocalDb.violationsOf(widget.student);
    if (!mounted) return;
    setState(() {
      _rows = rows;
      _loading = false;
    });
  }

  Future<void> _add() async {
    await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => ViolationFormScreen(
            student: widget.student, user: widget.currentUser),
      ),
    );
    await _load();
  }

  @override
  Widget build(BuildContext context) {
    final s = widget.student;
    final scheme = Theme.of(context).colorScheme;
    return Scaffold(
      appBar: AppBar(
        title: Text(s.fullName),
        actions: [
          if (!_loading)
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 12),
              child: Center(
                child: Text('${_rows.length} تخلف',
                    style: TextStyle(
                        fontSize: 12, color: scheme.onSurfaceVariant)),
              ),
            ),
        ],
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: _add,
        icon: const Icon(Icons.add),
        label: const Text('ثبت تخلف'),
      ),
      body: _loading
          ? const Center(child: CircularProgressIndicator())
          : ListView(
              padding: const EdgeInsets.fromLTRB(16, 8, 16, 100),
              children: [
                _header(s, scheme),
                const SizedBox(height: 18),
                Row(
                  children: [
                    Text('سوابق تخلف',
                        style: Theme.of(context).textTheme.titleSmall),
                    const SizedBox(width: 8),
                    Container(
                      padding: const EdgeInsets.symmetric(
                          horizontal: 8, vertical: 2),
                      decoration: BoxDecoration(
                        color: _rows.isEmpty
                            ? scheme.surfaceContainerHighest
                            : scheme.errorContainer,
                        borderRadius: BorderRadius.circular(20),
                      ),
                      child: Text(
                        _rows.isEmpty ? 'ندارد' : '${_rows.length}',
                        style: TextStyle(
                          fontSize: 11,
                          fontWeight: FontWeight.w700,
                          color: _rows.isEmpty
                              ? scheme.onSurfaceVariant
                              : scheme.onErrorContainer,
                        ),
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 8),
                if (_rows.isEmpty)
                  SectionCard(
                    child: Column(
                      children: [
                        Icon(Icons.inbox_outlined,
                            size: 42,
                            color:
                                scheme.onSurfaceVariant.withValues(alpha: 0.5)),
                        const SizedBox(height: 8),
                        Text('هنوز تخلفی برای این دانشجو ثبت نشده است',
                            style: TextStyle(color: scheme.onSurfaceVariant)),
                      ],
                    ),
                  )
                else
                  for (final r in _rows) ...[
                    _card(r, scheme),
                    const SizedBox(height: 10),
                  ],
              ],
            ),
    );
  }

  Widget _header(Student s, ColorScheme scheme) => SectionCard(
        child: Row(
          children: [
            CircleAvatar(
              radius: 26,
              backgroundColor: scheme.primaryContainer,
              foregroundColor: scheme.onPrimaryContainer,
              child: Text(s.initial),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(s.fullName,
                      style: const TextStyle(
                          fontSize: 16, fontWeight: FontWeight.w700)),
                  const SizedBox(height: 4),
                  Text('${s.studentNo}  •  ${s.major}',
                      style: TextStyle(color: scheme.onSurfaceVariant)),
                  if (s.nationalId.isNotEmpty)
                    Text('کد ملی: ${s.nationalId}',
                        style: TextStyle(
                            fontSize: 12, color: scheme.onSurfaceVariant)),
                ],
              ),
            ),
          ],
        ),
      );

  Widget _card(ViolationRow r, ColorScheme scheme) => SectionCard(
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Icon(Icons.event_note, size: 18, color: scheme.primary),
                const SizedBox(width: 6),
                Expanded(
                  child: Text(jalali(r.occurredAt),
                      style: TextStyle(
                          fontWeight: FontWeight.w600,
                          color: scheme.primary,
                          fontSize: 13)),
                ),
                if (r.recordedBy.isNotEmpty)
                  Text(r.recordedBy,
                      style: TextStyle(
                          fontSize: 11, color: scheme.onSurfaceVariant)),
              ],
            ),
            if (r.checkedItems.isNotEmpty) ...[
              const SizedBox(height: 8),
              Wrap(
                spacing: 6,
                runSpacing: 6,
                children: r.checkedItems
                    .map((t) => Container(
                          padding: const EdgeInsets.symmetric(
                              horizontal: 9, vertical: 4),
                          decoration: BoxDecoration(
                            color: scheme.surface,
                            borderRadius: BorderRadius.circular(10),
                            border: Border.all(
                                color: scheme.outlineVariant
                                    .withValues(alpha: 0.7)),
                          ),
                          child: Text(t,
                              style: const TextStyle(fontSize: 11)),
                        ))
                    .toList(),
              ),
            ],
            if (r.note.isNotEmpty) ...[
              const SizedBox(height: 8),
              Row(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Icon(Icons.notes,
                      size: 16, color: scheme.onSurfaceVariant),
                  const SizedBox(width: 6),
                  Expanded(child: Text(r.note)),
                ],
              ),
            ],
          ],
        ),
      );
}

/// ───────── صفحهٔ جستجو ─────────
class SearchScreen extends StatefulWidget {
  final String currentUser;
  const SearchScreen({super.key, required this.currentUser});

  @override
  State<SearchScreen> createState() => _SearchScreenState();
}

class _SearchScreenState extends State<SearchScreen> {
  final _ctrl = TextEditingController();
  List<StudentHit> _results = [];
  bool _searched = false;
  bool _busy = false;

  @override
  void dispose() {
    _ctrl.dispose();
    super.dispose();
  }

  Future<void> _run(String q) async {
    final r = await LocalDb.search(q);
    if (!mounted) return;
    setState(() {
      _results = r;
      _searched = q.trim().isNotEmpty;
    });
  }

  Future<void> _export() async {
    if (_busy) return;
    setState(() => _busy = true);
    try {
      final rows = await LocalDb.allViolations();
      if (rows.isEmpty) {
        _snack('هنوز تخلفی ثبت نشده است');
        return;
      }
      final file = await ExcelService()
          .exportViolations(rows, columns: await LocalDb.loadColumns());
      await ExcelService().share(file);
    } catch (e) {
      _snack('خطا: $e');
    } finally {
      if (mounted) setState(() => _busy = false);
    }
  }

  void _snack(String msg) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(msg)));
  }

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Scaffold(
      appBar: AppBar(
        title: const Text('جستجوی دانشجو'),
        actions: [
          IconButton(
            tooltip: 'خروجی اکسل تخلفات',
            onPressed: _busy ? null : _export,
            icon: const Icon(Icons.file_download_outlined),
          ),
        ],
      ),
      body: Column(
        children: [
          Padding(
            padding: const EdgeInsets.fromLTRB(16, 8, 16, 8),
            child: TextField(
              controller: _ctrl,
              onChanged: _run,
              textInputAction: TextInputAction.search,
              decoration: InputDecoration(
                hintText: 'کد دانشجویی، کد ملی یا نام',
                prefixIcon: const Icon(Icons.search),
                suffixIcon: _ctrl.text.isEmpty
                    ? null
                    : IconButton(
                        icon: const Icon(Icons.clear),
                        onPressed: () {
                          _ctrl.clear();
                          _run('');
                        },
                      ),
                filled: true,
                fillColor: scheme.surfaceContainerHighest.withValues(alpha: 0.5),
                border: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(16),
                  borderSide: BorderSide.none,
                ),
              ),
            ),
          ),
          if (_busy) const LinearProgressIndicator(minHeight: 2),
          Expanded(
            child: _results.isEmpty
                ? _empty(scheme)
                : ListView.separated(
                    padding: const EdgeInsets.fromLTRB(16, 8, 16, 24),
                    itemCount: _results.length,
                    separatorBuilder: (_, __) => const SizedBox(height: 10),
                    itemBuilder: (_, i) {
                      final hit = _results[i];
                      final s = hit.student;
                      final hasHistory = hit.violations > 0;
                      return SectionCard(
                        padding: EdgeInsets.zero,
                        child: ListTile(
                          isThreeLine: true,
                          contentPadding:
                              const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                          leading: CircleAvatar(
                            backgroundColor: scheme.primaryContainer,
                            foregroundColor: scheme.onPrimaryContainer,
                            child: Text(s.initial),
                          ),
                          title: Text(s.fullName,
                              style: const TextStyle(fontWeight: FontWeight.w600)),
                          subtitle: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              const SizedBox(height: 2),
                              Text('${s.studentNo}  •  ${s.major}'),
                              const SizedBox(height: 6),
                              Row(
                                children: [
                                  Container(
                                    padding: const EdgeInsets.symmetric(
                                        horizontal: 8, vertical: 2),
                                    decoration: BoxDecoration(
                                      color: hasHistory
                                          ? scheme.errorContainer
                                          : scheme.surfaceContainerHighest,
                                      borderRadius: BorderRadius.circular(20),
                                    ),
                                    child: Text(
                                      hasHistory
                                          ? '${hit.violations} تخلف ثبت‌شده'
                                          : 'بدون سابقه',
                                      style: TextStyle(
                                        fontSize: 11,
                                        fontWeight: FontWeight.w600,
                                        color: hasHistory
                                            ? scheme.onErrorContainer
                                            : scheme.onSurfaceVariant,
                                      ),
                                    ),
                                  ),
                                ],
                              ),
                            ],
                          ),
                          trailing: const Icon(Icons.chevron_left),
                          onTap: () async {
                            await Navigator.push(
                              context,
                              MaterialPageRoute(
                                builder: (_) => StudentDetailScreen(
                                    student: s,
                                    currentUser: widget.currentUser),
                              ),
                            );
                            _run(_ctrl.text);
                          },
                        ),
                      );
                    },
                  ),
          ),
        ],
      ),
    );
  }

  Widget _empty(ColorScheme scheme) => Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            Icon(_searched ? Icons.search_off : Icons.badge_outlined,
                size: 56,
                color: scheme.onSurfaceVariant.withValues(alpha: 0.5)),
            const SizedBox(height: 12),
            Text(_searched ? 'نتیجه‌ای پیدا نشد' : 'برای جستجو تایپ کنید',
                style: TextStyle(color: scheme.onSurfaceVariant)),
          ],
        ),
      );
}

/// ───────── فرم ثبت تخلف ─────────
class ViolationFormScreen extends StatefulWidget {
  final Student student;
  final String user;
  const ViolationFormScreen(
      {super.key, required this.student, required this.user});

  @override
  State<ViolationFormScreen> createState() => _ViolationFormScreenState();
}

class _ViolationFormScreenState extends State<ViolationFormScreen> {
  final _checked = <String>{};
  final _note = TextEditingController();
  DateTime _when = DateTime.now();
  int _previous = 0;
  bool _saving = false;

  @override
  void initState() {
    super.initState();
    LocalDb.violationCount(widget.student.key).then((c) {
      if (mounted) setState(() => _previous = c);
    });
  }

  @override
  void dispose() {
    _note.dispose();
    super.dispose();
  }

  Future<void> _pickWhen() async {
    final d = await showDatePicker(
      context: context,
      initialDate: _when,
      firstDate: DateTime(2020),
      lastDate: DateTime.now(),
    );
    if (d == null || !mounted) return;
    final t = await showTimePicker(
        context: context, initialTime: TimeOfDay.fromDateTime(_when));
    if (t == null) return;
    setState(() =>
        _when = DateTime(d.year, d.month, d.day, t.hour, t.minute));
  }

  Future<void> _save() async {
    if (_saving) return;
    setState(() => _saving = true);
    try {
      await LocalDb.addViolation(
          widget.student, _when, _checked.toList(), _note.text.trim(), widget.user);
      if (!mounted) return;
      ScaffoldMessenger.of(context)
          .showSnackBar(const SnackBar(content: Text('تخلف ثبت شد')));
      Navigator.pop(context);
    } catch (e) {
      if (!mounted) return;
      setState(() => _saving = false);
      ScaffoldMessenger.of(context)
          .showSnackBar(SnackBar(content: Text('خطا: $e')));
    }
  }

  @override
  Widget build(BuildContext context) {
    final s = widget.student;
    final scheme = Theme.of(context).colorScheme;
    return Scaffold(
      appBar: AppBar(title: const Text('ثبت تخلف')),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(16, 8, 16, 24),
        children: [
          SectionCard(
            child: Row(
              children: [
                CircleAvatar(
                  radius: 24,
                  backgroundColor: scheme.primaryContainer,
                  foregroundColor: scheme.onPrimaryContainer,
                  child: Text(s.initial),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(s.fullName,
                          style: const TextStyle(
                              fontSize: 16, fontWeight: FontWeight.w700)),
                      const SizedBox(height: 4),
                      Text('${s.studentNo}  •  ${s.major}',
                          style: TextStyle(color: scheme.onSurfaceVariant)),
                    ],
                  ),
                ),
                if (_previous > 0)
                  Chip(
                    label: Text('$_previous سابقه'),
                    visualDensity: VisualDensity.compact,
                  ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          Text('موارد تخلف',
              style: Theme.of(context).textTheme.titleSmall),
          const SizedBox(height: 8),
          SectionCard(
            padding: const EdgeInsets.symmetric(vertical: 4),
            child: Column(
              children: [
                for (final item in kViolationItems)
                  CheckboxListTile(
                    value: _checked.contains(item),
                    onChanged: (v) => setState(() =>
                        v == true ? _checked.add(item) : _checked.remove(item)),
                    title: Text(item),
                    dense: true,
                    controlAffinity: ListTileControlAffinity.leading,
                  ),
              ],
            ),
          ),
          const SizedBox(height: 16),
          OutlinedButton.icon(
            icon: const Icon(Icons.schedule),
            label: Text('تاریخ و ساعت:  ${jalali(_when)}'),
            style: OutlinedButton.styleFrom(
              minimumSize: const Size.fromHeight(48),
              shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(14)),
            ),
            onPressed: _pickWhen,
          ),
          const SizedBox(height: 12),
          TextField(
            controller: _note,
            maxLines: 2,
            decoration: InputDecoration(
              labelText: 'توضیحات (اختیاری)',
              filled: true,
              fillColor: scheme.surfaceContainerHighest.withValues(alpha: 0.5),
              border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(14),
                borderSide: BorderSide.none,
              ),
            ),
          ),
          const SizedBox(height: 20),
          FilledButton.icon(
            icon: _saving
                ? SizedBox(
                    height: 16,
                    width: 16,
                    child: CircularProgressIndicator(strokeWidth: 2))
                : const Icon(Icons.check),
            label: const Text('ثبت تخلف'),
            style: FilledButton.styleFrom(
              minimumSize: const Size.fromHeight(52),
              shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(16)),
            ),
            onPressed: _checked.isEmpty || _saving ? null : _save,
          ),
        ],
      ),
    );
  }
}
