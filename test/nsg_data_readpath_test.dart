// Functional check of the exact API surface nsg_data uses for xlsx IMPORT.
// Mirrors nsg_data/lib/nsg_excel_import.dart (Excel.decodeBytes -> sheets ->
// sheet.row(i) -> CellValue subtype dispatch) and nsg_excel_aliases.dart.
// Purpose: prove the archive 4 / xml 7 migration did not break reading.
import 'dart:io';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';

/// Verbatim copy of NsgExcelImport.safeCell — the cell -> String dispatch
/// nsg_data relies on for every imported field.
String safeCell(dynamic data) {
  if (data == null) return "";
  if (data is Data) {
    return data.value?.toString() ?? "";
  } else if (data is CellValue) {
    if (data is TextCellValue) return data.value.toString();
    if (data is FormulaCellValue) return data.formula;
    if (data is IntCellValue) return data.value.toString();
    if (data is BoolCellValue) return data.value.toString();
    if (data is DoubleCellValue) return data.value.toString();
    if (data is DateCellValue) return "${data.year}-${data.month}-${data.day}";
    if (data is TimeCellValue) {
      return "${data.hour}:${data.minute}:${data.second}";
    }
    if (data is DateTimeCellValue) {
      return DateTime(data.year, data.month, data.day, data.hour, data.minute,
              data.second)
          .toString();
    }
  }
  return "";
}

void main() {
  group('nsg_data import path', () {
    test('decodeBytes + sheets + row() over a real-world xlsx', () {
      // nsg_data reads via file.readAsBytesSync() -> Excel.decodeBytes.
      final bytes = File('test/test_resources/example.xlsx').readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);

      expect(excel.sheets.keys, isNotEmpty);

      var nonEmptyCells = 0;
      for (final name in excel.sheets.keys) {
        final sheet = excel[name];
        expect(sheet.maxRows, greaterThanOrEqualTo(0));
        for (var row = 0; row < sheet.maxRows; row++) {
          for (final cell in sheet.row(row)) {
            final s = safeCell(cell?.value);
            if (s.isNotEmpty) nonEmptyCells++;
          }
        }
      }
      // Real data must come back, not just an empty grid.
      expect(nonEmptyCells, greaterThan(0));
    });

    test('every CellValue subtype survives decode (all producers)', () {
      // These three files are written by MS Excel 365, Google Sheets and
      // LibreOffice respectively — different zip/compression producers.
      for (final f in [
        'dataTypesUsingMsExcel365Desktop.xlsx',
        'dataTypesUsingGoogleSpreadsheet.xlsx',
        'dataTypesUsingLibreoffice.xlsx',
      ]) {
        final excel =
            Excel.decodeBytes(File('test/test_resources/$f').readAsBytesSync());
        final seen = <String>{};
        for (final name in excel.sheets.keys) {
          final sheet = excel[name];
          for (var row = 0; row < sheet.maxRows; row++) {
            for (final cell in sheet.row(row)) {
              final v = cell?.value;
              if (v == null) continue;
              seen.add(v.runtimeType.toString());
              // safeCell must not throw and must not silently return "" for
              // a known type — that is how nsg_data loses data.
              expect(safeCell(v), isNotEmpty, reason: '$f: $v (${v.runtimeType})');
            }
          }
        }
        expect(seen, isNotEmpty, reason: '$f produced no typed cells');
        print('  $f -> ${seen.toList()..sort()}');
      }
    });

    test('decodeBuffer(InputMemoryStream) — archive 4 stream entry point', () {
      // Excel.decodeBuffer now takes archive 4's InputStream; the concrete
      // impl is InputMemoryStream (was InputStream in archive 3).
      final bytes = File('test/test_resources/example.xlsx').readAsBytesSync();
      final excel = Excel.decodeBuffer(InputMemoryStream(bytes));
      expect(excel.sheets.keys, isNotEmpty);
    });

    test('round-trip: save with archive 4 then re-read (compression change)',
        () {
      // Exercises _cloneArchive's compression = deflate/none rewrite and
      // ZipEncoder under archive 4, then proves the bytes are a valid zip
      // that both excel AND a raw ZipDecoder can open.
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(
          CellIndex.indexByString('A1'), TextCellValue('Фамилия'));
      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Табель'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('Иванов'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(4242));

      final encoded = excel.encode();
      expect(encoded, isNotNull);
      final out = Uint8List.fromList(encoded!);

      // Raw zip integrity — catches a malformed compression header.
      final rawZip = ZipDecoder().decodeBytes(out);
      expect(rawZip.findFile('xl/workbook.xml'), isNotNull);
      for (final f in rawZip.files.where((f) => f.isFile)) {
        expect(f.content, isNotNull, reason: 'undecodable entry ${f.name}');
      }

      // Full read-back through the nsg_data path.
      final reread = Excel.decodeBytes(out);
      final rows = reread['Sheet1'];
      expect(safeCell(rows.row(0)[0]?.value), 'Фамилия');
      expect(safeCell(rows.row(0)[1]?.value), 'Табель');
      expect(safeCell(rows.row(1)[0]?.value), 'Иванов');
      expect(safeCell(rows.row(1)[1]?.value), '4242');
    });
  });
}
