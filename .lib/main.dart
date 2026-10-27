// lib/main.dart
//
// نسخهٔ بدون سرور — «ثبت تخلف پوشش»
// ساختار: main.dart + excel_service.dart + violation_module.dart
import 'dart:convert';
import 'dart:io';
import 'dart:math';

import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:flutter_secure_storage/flutter_secure_storage.dart';
import 'package:path_provider/path_provider.dart';
import 'package:shamsi_date/shamsi_date.dart';
import 'package:share_plus/share_plus.dart';
import 'package:shared_preferences/shared_preferences.dart';

import 'excel_service.dart';
import 'violation_module.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  final prefs = await SharedPreferences.getInstance();
  const sec = FlutterSecureStorage();
  var key = await sec.read(key: 'db_key');
  if (key == null) {
    final r = Random.secure();
    key = base64Url.encode(List.generate(32, (_) => r.nextInt(256)));
    await sec.write(key: 'db_key', value: key);
  }
  await LocalDb.open(key);
  runApp(StandaloneApp(prefs: prefs));
}

class StandaloneApp extends StatelessWidget {
  final SharedPreferences prefs;
  const StandaloneApp({super.key, required this.prefs});

  @override
  Widget build(BuildContext context) {
    const seed = Color(0xFF0F766E);
    return MaterialApp(
      title: 'ثبت تخلف پوشش',
      debugShowCheckedModeBanner: false,
      locale: const Locale('fa'),
      supportedLocales: const [Locale('fa')],
      localizationsDelegates: const [
        GlobalMaterialLocalizations.delegate,
        GlobalWidgetsLocalizations.delegate,
        GlobalCupertinoLocalizations.delegate,
      ],
      builder: (c, child) =>
          Directionality(textDirection: TextDirection.rtl, child: child!),
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(
          seedColor: seed,
          brightness: Brightness.light,
        ),
        scaffoldBackgroundColor: const Color(0xFFF6F7F9),
        snackBarTheme: const SnackBarThemeData(
          behavior: SnackBarBehavior.floating,
        ),
      ),
      home: StandaloneHome(prefs: prefs),
    );
  }
}

class StandaloneHome extends StatefulWidget {
  final SharedPreferences prefs;
  const StandaloneHome({super.key, required this.prefs});

  @override
  State<StandaloneHome> createState() => _StandaloneHomeState();
}

class _StandaloneHomeState extends State<StandaloneHome> {
  int _students = 0;
  int _violations = 0;
  String _msg = '';
  bool _busy = false;

  String get _name => widget.prefs.getString('recorder') ?? '';

  @override
  void initState() {
    super.initState();
    _refresh();
    WidgetsBinding.instance.addPostFrameCallback((_) {
      if (_name.isEmpty) _askName();
    });
  }

  Future<void> _refresh() async {
    final s = await LocalDb.studentCount();
    final v = await LocalDb.violationTotal();
    if (!mounted) return;
    setState(() {
      _students = s;
      _violations = v;
    });
  }

  Future<void> _askName() async {
    final c = TextEditingController(text: _name);
    await showDialog<void>(
      context: context,
      barrierDismissible: _name.isNotEmpty,
      builder: (_) => AlertDialog(
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
        title: const Text('نام ثبت‌کننده'),
        content: TextField(
          controller: c,
          autofocus: true,
          decoration: const InputDecoration(
            hintText: 'نام و نام خانوادگی',
            border: OutlineInputBorder(),
          ),
        ),
        actions: [
          FilledButton(
            onPressed: () {
              final v = c.text.trim();
              if (v.isEmpty) return;
              widget.prefs.setString('recorder', v);
              Navigator.pop(context);
            },
            child: const Text('ذخیره'),
          ),
        ],
      ),
    );
    if (mounted) setState(() {});
  }

  void _snack(String msg) {
    if (!mounted) return;
    ScaffoldMessenger.of(context)
      ..hideCurrentSnackBar()
      ..showSnackBar(SnackBar(content: Text(msg)));
  }

  Future<void> _run(Future<void> Function() job) async {
    if (_busy) return;
    setState(() {
      _busy = true;
      _msg = 'در حال انجام...';
    });
    try {
      await job();
    } catch (e) {
      setState(() => _msg = 'خطا: $e');
    } finally {
      if (mounted) setState(() => _busy = false);
      _refresh();
    }
  }

  /// آپلود اکسل دانشجوها (سابقهٔ تخلف‌ها حفظ می‌شود)
  Future<void> _importExcel() async {
    final pick = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );
    if (pick == null) return;
    final path = pick.files.single.path;
    if (path == null) {
      _snack('خواندن فایل ممکن نشد');
      return;
    }
    await _run(() async {
      final svc = ExcelService();
      final res = await svc.importStudents(path);
      await LocalDb.upsertStudents(res.toImport);
      // ترتیب ستون‌های همین فایل ذخیره می‌شود تا خروجی هم‌شکل فایل ورودی باشد
      await LocalDb.saveColumns(res.columns);
      setState(() => _msg = 'کل ردیف‌ها: ${res.totalRows}\n'
          'دانشجوی وارد‌شده: ${res.toImport.length}\n'
          'ردیف تکراری: ${res.duplicateRows}');
      await svc.share(res.reportFile);
    });
  }

  Future<void> _exportExcel() => _run(() async {
        final rows = await LocalDb.allViolations();
        if (rows.isEmpty) {
          setState(() => _msg = 'هنوز تخلفی ثبت نشده است');
          return;
        }
        final svc = ExcelService();
        final cols = await LocalDb.loadColumns();
        await svc.share(await svc.exportViolations(rows, columns: cols));
        setState(() => _msg = 'خروجی اکسل آماده شد (${rows.length} ردیف)');
      });

  /// پشتیبان JSON برای جمع‌کردن داده چند گوشی
  Future<void> _exportBackup() => _run(() async {
        final data = jsonEncode(
            {'v': 1, 'violations': await LocalDb.rawViolations()});
        final dir = await getApplicationDocumentsDirectory();
        final j = Jalali.now();
        final stamp = '${j.year}'
            '${j.month.toString().padLeft(2, '0')}'
            '${j.day.toString().padLeft(2, '0')}';
        final f = File('${dir.path}/پشتیبان_تخلف_$stamp.json');
        await f.writeAsString(data);
        await Share.shareXFiles([XFile(f.path)]);
        setState(() => _msg = 'فایل پشتیبان ساخته شد');
      });

  Future<void> _importBackup() async {
    final pick = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['json'],
    );
    if (pick == null) return;
    final path = pick.files.single.path;
    if (path == null) return;
    await _run(() async {
      final m = jsonDecode(await File(path).readAsString()) as Map;
      final added = await LocalDb.importViolations(m['violations'] as List);
      setState(() => _msg = '$added تخلف جدید اضافه شد '
          '(موارد تکراری نادیده گرفته شدند)');
    });
  }

  Future<void> _openSearch() async {
    if (_students == 0) {
      _snack('ابتدا فایل اکسل دانشجویان را وارد کنید');
      return;
    }
    if (_name.isEmpty) {
      await _askName();
      if (_name.isEmpty) return;
    }
    await Navigator.push(
      context,
      MaterialPageRoute(builder: (_) => SearchScreen(currentUser: _name)),
    );
    _refresh();
  }

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Scaffold(
      body: ListView(
        padding: EdgeInsets.zero,
        children: [
          _Header(
            name: _name,
            students: _students,
            violations: _violations,
            onEditName: _askName,
          ),
          Padding(
            padding: const EdgeInsets.fromLTRB(16, 20, 16, 32),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                FilledButton.icon(
                  onPressed: _openSearch,
                  icon: const Icon(Icons.search),
                  label: const Text('جستجو و ثبت تخلف'),
                  style: FilledButton.styleFrom(
                    minimumSize: const Size.fromHeight(54),
                    shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(16)),
                    textStyle: const TextStyle(
                        fontSize: 16, fontWeight: FontWeight.w600),
                  ),
                ),
                const SizedBox(height: 20),
                Text('فایل‌ها',
                    style: Theme.of(context).textTheme.titleSmall),
                const SizedBox(height: 8),
                _ActionTile(
                  icon: Icons.upload_file,
                  title: 'وارد کردن اکسل دانشجویان',
                  subtitle: 'فایل خروجی سامانهٔ دانشگاه (xlsx)',
                  onTap: _busy ? null : _importExcel,
                ),
                const SizedBox(height: 10),
                _ActionTile(
                  icon: Icons.file_download_outlined,
                  title: 'خروجی اکسل تخلفات',
                  subtitle: 'یک شیت، با ستون‌های دانشجو و تخلف',
                  onTap: _busy ? null : _exportExcel,
                ),
                const SizedBox(height: 20),
                Text('انتقال داده بین گوشی‌ها',
                    style: Theme.of(context).textTheme.titleSmall),
                const SizedBox(height: 8),
                Row(
                  children: [
                    Expanded(
                      child: _MiniTile(
                        icon: Icons.save_alt,
                        label: 'ساخت پشتیبان',
                        onTap: _busy ? null : _exportBackup,
                      ),
                    ),
                    const SizedBox(width: 10),
                    Expanded(
                      child: _MiniTile(
                        icon: Icons.restore_page,
                        label: 'دریافت پشتیبان',
                        onTap: _busy ? null : _importBackup,
                      ),
                    ),
                  ],
                ),
                if (_busy) ...[
                  const SizedBox(height: 20),
                  const LinearProgressIndicator(minHeight: 3),
                ],
                if (_msg.isNotEmpty) ...[
                  const SizedBox(height: 16),
                  Container(
                    padding: const EdgeInsets.all(14),
                    decoration: BoxDecoration(
                      color: scheme.primaryContainer.withValues(alpha: 0.6),
                      borderRadius: BorderRadius.circular(16),
                    ),
                    child: Text(_msg,
                        style: TextStyle(color: scheme.onPrimaryContainer)),
                  ),
                ],
                if (_students == 0 && !_busy) ...[
                  const SizedBox(height: 20),
                  Text(
                    'برای شروع، فایل اکسل دانشجویان را وارد کنید.',
                    textAlign: TextAlign.center,
                    style: TextStyle(color: scheme.onSurfaceVariant),
                  ),
                ],
              ],
            ),
          ),
        ],
      ),
    );
  }
}

/// هدر رنگی بالای صفحه + آمار
class _Header extends StatelessWidget {
  final String name;
  final int students;
  final int violations;
  final VoidCallback onEditName;

  const _Header({
    required this.name,
    required this.students,
    required this.violations,
    required this.onEditName,
  });

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Container(
      padding: const EdgeInsets.fromLTRB(20, 24, 20, 22),
      decoration: BoxDecoration(
        gradient: LinearGradient(
          begin: Alignment.topRight,
          end: Alignment.bottomLeft,
          colors: [scheme.primary, scheme.primary.withValues(alpha: 0.72)],
        ),
        borderRadius:
            const BorderRadius.vertical(bottom: Radius.circular(28)),
      ),
      child: SafeArea(
        bottom: false,
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                CircleAvatar(
                  radius: 22,
                  backgroundColor: Colors.white.withValues(alpha: 0.2),
                  child: const Icon(Icons.verified_user_outlined,
                      color: Colors.white),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        name.isEmpty ? 'ثبت تخلف پوشش' : 'سلام $name',
                        style: const TextStyle(
                            color: Colors.white,
                            fontSize: 18,
                            fontWeight: FontWeight.w700),
                      ),
                      const SizedBox(height: 2),
                      Text('سامانهٔ ثبت و گزارش تخلف پوشش',
                          style: TextStyle(
                              color: Colors.white.withValues(alpha: 0.85),
                              fontSize: 12)),
                    ],
                  ),
                ),
                IconButton(
                  tooltip: 'تغییر نام ثبت‌کننده',
                  onPressed: onEditName,
                  icon: const Icon(Icons.edit_outlined, color: Colors.white),
                ),
              ],
            ),
            const SizedBox(height: 20),
            Row(
              children: [
                Expanded(
                  child: _Stat(
                      icon: Icons.people_alt_outlined,
                      value: '$students',
                      label: 'دانشجو'),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _Stat(
                      icon: Icons.report_gmailerrorred_outlined,
                      value: '$violations',
                      label: 'تخلف ثبت‌شده'),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}

class _Stat extends StatelessWidget {
  final IconData icon;
  final String value;
  final String label;
  const _Stat({required this.icon, required this.value, required this.label});

  @override
  Widget build(BuildContext context) => Container(
        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
        decoration: BoxDecoration(
          color: Colors.white.withValues(alpha: 0.16),
          borderRadius: BorderRadius.circular(18),
        ),
        child: Row(
          children: [
            Icon(icon, color: Colors.white, size: 22),
            const SizedBox(width: 10),
            Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(value,
                    style: const TextStyle(
                        color: Colors.white,
                        fontSize: 18,
                        fontWeight: FontWeight.w700)),
                Text(label,
                    style: TextStyle(
                        color: Colors.white.withValues(alpha: 0.85),
                        fontSize: 11)),
              ],
            ),
          ],
        ),
      );
}

class _ActionTile extends StatelessWidget {
  final IconData icon;
  final String title;
  final String subtitle;
  final VoidCallback? onTap;
  const _ActionTile({
    required this.icon,
    required this.title,
    required this.subtitle,
    this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Material(
      color: Colors.white,
      borderRadius: BorderRadius.circular(18),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(18),
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Row(
            children: [
              Container(
                width: 42,
                height: 42,
                decoration: BoxDecoration(
                  color: scheme.primaryContainer.withValues(alpha: 0.7),
                  borderRadius: BorderRadius.circular(12),
                ),
                child: Icon(icon, color: scheme.onPrimaryContainer, size: 22),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(title,
                        style: const TextStyle(fontWeight: FontWeight.w600)),
                    const SizedBox(height: 2),
                    Text(subtitle,
                        style: TextStyle(
                            fontSize: 12, color: scheme.onSurfaceVariant)),
                  ],
                ),
              ),
              Icon(Icons.chevron_left,
                  color: scheme.onSurfaceVariant.withValues(alpha: 0.6)),
            ],
          ),
        ),
      ),
    );
  }
}

class _MiniTile extends StatelessWidget {
  final IconData icon;
  final String label;
  final VoidCallback? onTap;
  const _MiniTile({required this.icon, required this.label, this.onTap});

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Material(
      color: Colors.white,
      borderRadius: BorderRadius.circular(16),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(16),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 14),
          child: Row(
            children: [
              Icon(icon, size: 20, color: scheme.primary),
              const SizedBox(width: 8),
              Expanded(
                child: Text(label,
                    style: const TextStyle(fontSize: 13),
                    overflow: TextOverflow.ellipsis),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
