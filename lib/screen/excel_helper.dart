import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:open_file/open_file.dart';
import 'package:excel/excel.dart';

class ExcelHelper {
  static const String _fileName = 'contacts.xlsx';
  static const String _sheetName = 'Contacts';

  static Future<bool> _checkAndRequestPermission(BuildContext context) async {
    if (!Platform.isAndroid) return true;

    // On Android 13+ storage permission is no longer applicable
    if (Platform.isAndroid && Platform.version.contains('13') || Platform.version.contains('14') || Platform.version.contains('15')) {
      return true;
    }

    try {
      var status = await Permission.storage.status;
      if (!status.isGranted) {
        status = await Permission.storage.request();
        if (!status.isGranted) {
          if (status.isPermanentlyDenied) {
            await _showPermissionDialog(context);
          }
          return false;
        }
      }
      return true;
    } catch (e) {
      debugPrint('Permission error: $e');
      return false;
    }
  }

  static Future<void> _showPermissionDialog(BuildContext context) async {
    await showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Permission Required'),
        content: const Text(
          'This app needs storage permission to save and export contacts.',
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('Cancel'),
          ),
          TextButton(
            onPressed: () async {
              Navigator.pop(context);
              await openAppSettings();
            },
            child: const Text('Open Settings'),
          ),
        ],
      ),
    );
  }

  static Future<File> get _localFile async {
    final directory = await getApplicationDocumentsDirectory();
    return File('${directory.path}/$_fileName');
  }

  static Future<void> saveContact({
    required BuildContext context,
    required String name,
    required String phone,
  }) async {
    try {
      if (!await _checkAndRequestPermission(context)) {
        throw Exception('Storage permission denied');
      }

      final file = await _localFile;
      Excel excel;

      if (await file.exists()) {
        try {
          final bytes = await file.readAsBytes();
          excel = Excel.decodeBytes(bytes);
          if (!excel.tables.containsKey(_sheetName)) {
            excel.rename(excel.tables.keys.first, _sheetName);
          }
        } catch (e) {
          excel = _createNewExcel();
        }
      } else {
        excel = _createNewExcel();
      }

      final sheet = excel[_sheetName];
      if (sheet.rows.isEmpty || sheet.rows[0][0]?.value != 'Name') {
        sheet.appendRow(['Name', 'Phone Number']);
      }
      sheet.appendRow([name, phone]);

      await file.writeAsBytes(excel.encode()!, flush: true);

      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Contact saved successfully!')),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Error saving contact: ${e.toString()}')),
      );
      rethrow;
    }
  }

  static Excel _createNewExcel() {
    final excel = Excel.createExcel();
    final sheet = excel[_sheetName];
    sheet.appendRow(['Name', 'Phone Number']);
    return excel;
  }

  static Future<void> exportContacts(BuildContext context) async {
    try {
      if (!await _checkAndRequestPermission(context)) {
        throw Exception('Storage permission denied');
      }

      final file = await _localFile;
      if (!await file.exists()) {
        throw Exception('No contacts found to export');
      }

      if (Platform.isAndroid) {
        final downloadsDir = await _getDownloadsDirectory();
        final timestamp = DateTime.now().millisecondsSinceEpoch;
        final newFile = File('${downloadsDir.path}/contacts_$timestamp.xlsx');
        await newFile.writeAsBytes(await file.readAsBytes());

        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Exported to Downloads folder')),
        );
      } else {
        await OpenFile.open(file.path);
      }
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Export failed: ${e.toString()}')),
      );
      rethrow;
    }
  }

  static Future<Directory> _getDownloadsDirectory() async {
    try {
      if (Platform.isAndroid) {
        final directory = Directory('/storage/emulated/0/Download');
        if (await directory.exists()) {
          return directory;
        }
        final fallback = await getExternalStorageDirectory();
        if (fallback != null) {
          final custom = Directory('${fallback.path}/Download');
          if (!await custom.exists()) {
            await custom.create(recursive: true);
          }
          return custom;
        }
      }
      return await getApplicationDocumentsDirectory();
    } catch (e) {
      return await getApplicationDocumentsDirectory();
    }
  }
}