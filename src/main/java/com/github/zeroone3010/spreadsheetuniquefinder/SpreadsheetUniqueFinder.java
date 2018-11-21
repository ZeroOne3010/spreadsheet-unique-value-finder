package com.github.zeroone3010.spreadsheetuniquefinder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.logging.Logger;
import java.util.stream.Collectors;

public class SpreadsheetUniqueFinder {
  private static final Logger logger = Logger.getLogger("SpreadsheetUniqueFinder");
  private final Workbook workbook;
  private final DataFormatter dataFormatter;

  public SpreadsheetUniqueFinder(final Workbook workbook) {
    this.workbook = workbook;
    this.dataFormatter = new DataFormatter();
  }

  public LinkedHashMap<String, SortedMap<String, Integer>> findUniqueColumnValues() {
    final Sheet sheet = workbook.getSheetAt(0);
    final Row headerRow = sheet.getRow(0);
    final int rows = sheet.getLastRowNum();
    final int columns = headerRow.getLastCellNum();

    final LinkedHashMap<String, SortedMap<String, Integer>> allValueCounts = new LinkedHashMap<>();
    for (int column = 0; column < columns; column++) {
      final Cell headerCell = sheet.getRow(0).getCell(column);
      final String header = dataFormatter.formatCellValue(headerCell);

      final SortedMap<String, Integer> valueCounts = new TreeMap<>();
      for (int row = 1; row <= rows; row++) {
        final Cell cell = sheet.getRow(row).getCell(column);
        final String cellValue = dataFormatter.formatCellValue(cell);
        valueCounts.merge(cellValue, 1, (a, b) -> a + b);
      }

      allValueCounts.put(header, valueCounts);
    }
    return allValueCounts;
  }

  public static void main(final String... args) {
    try (final FileInputStream fileInputStream = new FileInputStream(new File(args[0]))) {
      final Workbook workbook = new HSSFWorkbook(fileInputStream);
      final SpreadsheetUniqueFinder uniqueFinder = new SpreadsheetUniqueFinder(workbook);
      final LinkedHashMap<String, SortedMap<String, Integer>> result = uniqueFinder.findUniqueColumnValues();
      result.entrySet().forEach(entry -> System.out.println(entry.getKey() + "," +
          entry.getValue().keySet().stream().collect(Collectors.joining(","))));
    } catch (final FileNotFoundException e) {
      logger.warning("File not found!");
    } catch (final IOException e) {
      logger.warning("Error when reading the input file!");
      e.printStackTrace();
    }
  }
}
