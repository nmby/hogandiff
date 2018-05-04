package xyz.hotchpotch.hogandiff.poi;

import static org.hamcrest.CoreMatchers.*;
import static org.hamcrest.MatcherAssert.*;
import static org.junit.jupiter.api.Assertions.*;
import static xyz.hotchpotch.hogandiff.poi.TestFiles.*;

import org.junit.jupiter.api.Test;

import xyz.hotchpotch.hogandiff.ApplicationException;

class HSSFSheetListerWithEventApiTest {
    
    // [static members] ********************************************************
    
    private static final HSSFSheetListerWithEventApi lister = HSSFSheetListerWithEventApi.of();
    
    // [instance members] ******************************************************
    
    @Test
    void testIsSupported_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> HSSFSheetListerWithEventApi.isSupported(null));
    }
    
    @Test
    void testIsSupported() {
        assertTrue(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xls));
        
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsx));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsm));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_xlsb));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal_csv));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLister1_normal));
    }
    
    @Test
    void testOf() {
        assertThat(
                HSSFSheetListerWithEventApi.of(),
                instanceOf(HSSFSheetListerWithEventApi.class));
    }
    
    @Test
    void testGetSheetNames_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> lister.getSheetNames(null));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> lister.getSheetNames(SheetLister1_normal_xlsx));
    }
    
    @Test
    void testGetSheetNames_通常ケース() throws ApplicationException {
        assertEquals(
                sheetNames1,
                lister.getSheetNames(SheetLister1_normal_xls));
    }
    
    @Test
    void testGetSheetNames_シート名バリエーション() throws ApplicationException {
        assertEquals(
                sheetNames2,
                lister.getSheetNames(SheetLister2_namevariations_xls));
    }
    
    @Test
    void testGetSheetNames_拡張子偽装の不正な内容() throws ApplicationException {
        assertThrows(
                ApplicationException.class,
                () -> lister.getSheetNames(SheetLister3_dummy_xls));
    }
}
