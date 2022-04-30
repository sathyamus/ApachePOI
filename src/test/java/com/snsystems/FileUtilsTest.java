package com.snsystems;

import org.junit.Test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.Assert.assertTrue;

public class FileUtilsTest {

    @Test
    public void readExcelSheetUsingFile() {
        boolean isFileExists = new File("src/main/resources/ExcelSheet.xls").exists();
        assertTrue(isFileExists);
    }

    @Test
    public void readExcelSheetUsingPaths() {
        boolean isFileExists = Paths.get("src/main/resources/ExcelSheet.xls").toFile().exists();
        assertTrue(isFileExists);
    }

    @Test
    public void readExcelSheetUsingURI() throws Exception {
        Path path = Paths.get(getClass().getClassLoader()
                .getResource("ExcelSheet.xls").toURI());
        boolean isFileExists = path.toFile().exists();
        boolean isReadable = Files.isReadable(path);
        assertTrue(isFileExists);
        assertTrue(isReadable);
    }
}
