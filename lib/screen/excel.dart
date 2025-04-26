import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:open_file/open_file.dart';
import 'package:excel/excel.dart';

class ExcelHelper {
  static const String _fileName = 'contacts.xlsx';
  static const String _sheetName = 'Contacts';

  /// Checks and requests necessary permissions
  static Future<bool> _checkPermissions() async {
    if (Platform.isAndroid) {
      try {
        final status = await Permission.storage.request();
        if (status.isGranted) return true;
        if (status.isPermanentlyDenied) {
          await openAppSettings();
        }
        return false;
      } catch (e) {
        debugPrint('Permission error: $e');
        return false;
      }
    }
    return true; // For iOS/macOS/Windows, permissions aren't required
  }

  /// Gets the local application documents path
  static Future<String> get _localPath async {
    try {
      final directory = await getApplicationDocumentsDirectory();
      return directory.path;
    } catch (e) {
      debugPrint('Error getting local path: $e');
      rethrow;
    }
  }

  /// Gets the local Excel file reference
  static Future<File> get _localFile async {
    try {
      final path = await _localPath;
      return File('$path/$_fileName');
    } catch (e) {
      debugPrint('Error getting local file: $e');
      rethrow;
    }
  }

  /// Saves contact data to Excel file
  static Future<void> saveToExcel(String name, String mobileNumber) async {
    try {
      if (!await _checkPermissions()) {
        throw Exception('Storage permission not granted');
      }

      final file = await _localFile;
      Excel excel;

      // Try to read existing file or create new one
      if (await file.exists()) {
        try {
          final bytes = await file.readAsBytes();
          excel = Excel.decodeBytes(bytes);

          // Ensure our sheet exists
          if (!excel.tables.containsKey(_sheetName)) {
            excel.rename(excel.tables.keys.first, _sheetName);
          }
        } catch (e) {
          // File exists but is corrupted, create new
          excel = _createNewExcel();
        }
      } else {
        excel = _createNewExcel();
      }

      // Add data to sheet
      final sheet = excel[_sheetName];

      // Check if headers exist, add if missing
      if (sheet.rows.isEmpty ||
          sheet.rows.first[0]?.value != 'Name' ||
          sheet.rows.first[1]?.value != 'Mobile Number') {
        sheet.appendRow(['Name', 'Mobile Number']);
      }

      sheet.appendRow([name, mobileNumber]);

      // Save file
      await file.writeAsBytes(excel.encode()!, flush: true);
    } catch (e) {
      debugPrint('Error saving to Excel: $e');
      rethrow;
    }
  }

  /// Creates a new Excel file with proper structure
  static Excel _createNewExcel() {
    final excel = Excel.createExcel();
    final sheet = excel[_sheetName];
    sheet.appendRow(['Name', 'Mobile Number']); // Header
    return excel;
  }

  /// Exports/downloads the Excel file
  static Future<void> exportExcel(BuildContext context) async {
    try {
      if (!await _checkPermissions()) {
        throw Exception('Storage permission not granted');
      }

      final file = await _localFile;

      if (!await file.exists()) {
        throw Exception('No contacts data found');
      }

      if (Platform.isAndroid) {
        // For Android, copy to Downloads folder
        final downloadsDir = await _getDownloadsDirectory();
        final timestamp = DateTime.now().millisecondsSinceEpoch;
        final newFile = File('${downloadsDir.path}/contacts_$timestamp.xlsx');
        await newFile.writeAsBytes(await file.readAsBytes());

        // Show success message
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('File saved to Downloads folder')),
        );
      } else {
        // For other platforms, just open the file
        await OpenFile.open(file.path);
      }
    } catch (e) {
      debugPrint('Error exporting Excel: $e');
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Error exporting file: ${e.toString()}')),
      );
      rethrow;
    }
  }

  /// Gets the appropriate downloads directory with fallback
  static Future<Directory> _getDownloadsDirectory() async {
    try {
      if (Platform.isAndroid) {
        // Try primary downloads location
        final downloadsDir = Directory('/storage/emulated/0/Download');
        if (await downloadsDir.exists()) return downloadsDir;

        // Try secondary location
        final altDir = Directory('/storage/emulated/0/Downloads');
        if (await altDir.exists()) return altDir;

        // Create if doesn't exist
        await downloadsDir.create(recursive: true);
        return downloadsDir;
      } else {
        // For non-Android, use documents directory
        final dir = await getApplicationDocumentsDirectory();
        return Directory('${dir.path}/Downloads');
      }
    } catch (e) {
      // Fallback to app documents directory
      final dir = await getApplicationDocumentsDirectory();
      return dir;
    }
  }

  /// Reads all contacts from Excel file
  static Future<List<Map<String, String>>> readContacts() async {
    try {
      final file = await _localFile;
      if (!await file.exists()) return [];

      final bytes = await file.readAsBytes();
      final excel = Excel.decodeBytes(bytes);

      if (!excel.tables.containsKey(_sheetName)) return [];

      final sheet = excel[_sheetName];
      final contacts = <Map<String, String>>[];

      // Skip header row if exists
      final startRow = (sheet.rows.isNotEmpty &&
          sheet.rows.first[0]?.value == 'Name') ? 1 : 0;

      for (var i = startRow; i < sheet.rows.length; i++) {
        final row = sheet.rows[i];
        if (row.length >= 2) {
          contacts.add({
            'name': row[0]?.value?.toString() ?? '',
            'number': row[1]?.value?.toString() ?? ''
          });
        }
      }

      return contacts;
    } catch (e) {
      debugPrint('Error reading contacts: $e');
      return [];
    }
  }
}