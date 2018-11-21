package com.github.zeroone3010.spreadsheetuniquefinder;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.SortedMap;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

class SpreadsheetUniqueFinderTest {
  @Test
  void testXlsCreatedInLibreOffice() throws IOException {
    final InputStream inputStream = getClass().getClassLoader().getResourceAsStream("test_xls_created_in_libreoffice.xls");
    final Workbook workbook = new HSSFWorkbook(inputStream);

    final SpreadsheetUniqueFinder uniqueFinder = new SpreadsheetUniqueFinder(workbook);
    final LinkedHashMap<String, SortedMap<String, Integer>> result = uniqueFinder.findUniqueColumnValues();

    assertNull(result.get("Foo").get("Foo")); // Should ignore headers
    assertEquals(Integer.valueOf(3), result.get("Foo").get("1"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get("2"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get("3"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get("0"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get("-"));
    assertEquals(Integer.valueOf(4), result.get("Foo").get("banana"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get("terracotta"));
    assertEquals(Integer.valueOf(1), result.get("Foo").get(""));

    assertNull(result.get("Bar").get("Bar"));
    assertEquals(Integer.valueOf(2), result.get("Bar").get("true"));
    assertEquals(Integer.valueOf(1), result.get("Bar").get("false"));
    assertEquals(Integer.valueOf(1), result.get("Bar").get("FALSE"));
    assertEquals(Integer.valueOf(1), result.get("Bar").get("FOO"));
    assertEquals(Integer.valueOf(1), result.get("Bar").get("Foo"));
    assertEquals(Integer.valueOf(1), result.get("Bar").get("fOO"));
    assertEquals(Integer.valueOf(5), result.get("Bar").get(""));

    assertNull(result.get("Baz").get("Baz"));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("    "));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("  "));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("the first row has four spaces"));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("the second row has two"));
    assertEquals(Integer.valueOf(3), result.get("Baz").get(""));
    assertEquals(Integer.valueOf(2), result.get("Baz").get("."));
    assertEquals(Integer.valueOf(2), result.get("Baz").get("9"));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("9999"));
    assertEquals(Integer.valueOf(1), result.get("Baz").get("1"));
  }
}