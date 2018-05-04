package xyz.hotchpotch.hogandiff.poi;

import static org.hamcrest.CoreMatchers.*;
import static org.hamcrest.MatcherAssert.*;
import static org.junit.jupiter.api.Assertions.*;
import static xyz.hotchpotch.hogandiff.poi.TestFiles.*;

import org.junit.jupiter.api.Test;

import xyz.hotchpotch.hogandiff.ApplicationException;

class XSSFSheetListerWithEventApiTest {
    
    // [static members] ********************************************************
    
    private static final XSSFSheetListerWithEventApi lister = XSSFSheetListerWithEventApi.of();
    
    // [instance members] ******************************************************
    
    @Test
    void testIsSupported_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> XSSFSheetListerWithEventApi.isSupported(null));
    }
    
    @Test
    void testIsSupported() {
        assertTrue(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsx));
        assertTrue(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsm));
        
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsb));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xls));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_csv));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLister1_normal));
    }
    
    @Test
    void testOf() {
        assertThat(
                XSSFSheetListerWithEventApi.of(),
                instanceOf(XSSFSheetListerWithEventApi.class));
    }
    
    @Test
    void testGetSheetNames_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> lister.getSheetNames(null));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> lister.getSheetNames(SheetLister1_normal_xls));
    }
    
    @Test
    void testGetSheetNames_通常ケース() throws ApplicationException {
        assertEquals(
                sheetNames1,
                lister.getSheetNames(SheetLister1_normal_xlsx));
        
        assertEquals(
                sheetNames1,
                lister.getSheetNames(SheetLister1_normal_xlsm));
    }
    
    @Test
    void testGetSheetNames_シート名バリエーション() throws ApplicationException {
        assertEquals(
                sheetNames2,
                lister.getSheetNames(SheetLister2_namevariations_xlsx));
        
        assertEquals(
                sheetNames2,
                lister.getSheetNames(SheetLister2_namevariations_xlsm));
    }
    
    @Test
    void testGetSheetNames_拡張子偽装の不正な内容() throws ApplicationException {
        assertThrows(
                ApplicationException.class,
                () -> lister.getSheetNames(SheetLister3_dummy_xlsx));
    }
    
    @Test
    void testGetSheetsId_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> lister.getSheetsId(null));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> lister.getSheetsId(SheetLister1_normal_xls));
    }
    
    @Test
    void testGetSheetsId_通常ケース() throws ApplicationException {
        assertEquals(
                sheetsId1,
                lister.getSheetsId(SheetLister1_normal_xlsx));
        
        assertEquals(
                sheetsId1,
                lister.getSheetsId(SheetLister1_normal_xlsm));
    }
    
    @Test
    void testGetSheetsId_シート名バリエーション() throws ApplicationException {
        assertEquals(
                sheetsId2,
                lister.getSheetsId(SheetLister2_namevariations_xlsx));
        
        assertEquals(
                sheetsId2,
                lister.getSheetsId(SheetLister2_namevariations_xlsm));
    }
    
    @Test
    void testGetSheetsId_拡張子偽装の不正な内容() throws ApplicationException {
        assertThrows(
                ApplicationException.class,
                () -> lister.getSheetsId(SheetLister3_dummy_xlsx));
    }
}
