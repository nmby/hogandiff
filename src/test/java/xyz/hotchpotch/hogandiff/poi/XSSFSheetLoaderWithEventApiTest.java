package xyz.hotchpotch.hogandiff.poi;

import static org.hamcrest.CoreMatchers.*;
import static org.junit.Assert.*;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static xyz.hotchpotch.hogandiff.poi.TestFiles.*;

import java.util.NoSuchElementException;

import org.junit.jupiter.api.Test;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.poi.XSSFSheetListerWithEventApi;
import xyz.hotchpotch.hogandiff.poi.XSSFSheetLoaderWithEventApi;

class XSSFSheetLoaderWithEventApiTest {
    
    // [static members] ********************************************************
    
    private static final XSSFSheetLoaderWithEventApi valueLoader = XSSFSheetLoaderWithEventApi.of(true);
    private static final XSSFSheetLoaderWithEventApi formulaLoader = XSSFSheetLoaderWithEventApi.of(false);
    
    // [instance members] ******************************************************
    
    @Test
    void testIsSupported_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> XSSFSheetLoaderWithEventApi.isSupported(null));
    }
    
    @Test
    void testIsSupported() {
        assertTrue(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsx));
        assertTrue(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsm));
        
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsb));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xls));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_csv));
        assertFalse(XSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal));
    }
    
    @Test
    void testOf() {
        assertThat(
                valueLoader,
                instanceOf(XSSFSheetLoaderWithEventApi.class));
        
        assertThat(
                formulaLoader,
                instanceOf(XSSFSheetLoaderWithEventApi.class));
    }
    
    @Test
    void testLoadSheetById_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheetById(SheetLoader1_normal_xlsx, null));
        
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheetById(null, "rId1"));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> valueLoader.loadSheetById(SheetLoader1_normal_xls, "rId1"));
        
        assertThrows(
                NoSuchElementException.class,
                () -> valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId999"));
    }
    
    @Test
    void testLoadSheetById_通常ケース_xlsx() throws ApplicationException {
        assertEquals(
                cellsA,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId2"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId2"));
        
        assertEquals(
                cellsB_value,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId3"));
        assertEquals(
                cellsB_formula,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId3"));
        
        assertEquals(
                cellsC,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId4"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId4"));
        
        assertEquals(
                cellsE,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId6"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId6"));
    }
    
    @Test
    void testLoadSheetById_通常ケース_xlsm() throws ApplicationException {
        assertEquals(
                cellsA,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId2"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId2"));
        
        assertEquals(
                cellsB_value,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId3"));
        assertEquals(
                cellsB_formula,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId3"));
        
        assertEquals(
                cellsC,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId4"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId4"));
        
        assertEquals(
                cellsE,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId6"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId6"));
    }
    
    @Test
    void testLoadSheetById_グラフシート() throws ApplicationException {
        // 現状の実装では、グラフシートからは空のセルデータセットが返される。
        // この挙動を仕様として追認することにする。
        assertEquals(
                cellsA,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId5"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsx, "rId5"));
        
        assertEquals(
                cellsA,
                valueLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId5"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheetById(SheetLoader1_normal_xlsm, "rId5"));
    }
    
    @Test
    void testLoadSheet_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xlsx, null));
        
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheet(null, "あああ"));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xls, "あああ"));
        
        assertThrows(
                NoSuchElementException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xlsx, "ををを"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xlsx() throws ApplicationException {
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"));
        
        assertEquals(
                cellsB_value,
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"));
        assertEquals(
                cellsB_formula,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"));
        
        assertEquals(
                cellsC,
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"));
        
        assertEquals(
                cellsE,
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xlsm() throws ApplicationException {
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"));
        
        assertEquals(
                cellsB_value,
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"));
        assertEquals(
                cellsB_formula,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"));
        
        assertEquals(
                cellsC,
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"));
        
        assertEquals(
                cellsE,
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"));
    }
    
    @Test
    void testLoadSheet_グラフシート() throws ApplicationException {
        // 現状の実装では、グラフシートからは空のセルデータセットが返される。
        // この挙動を仕様として追認することにする。
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "えええ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "えええ"));
        
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "えええ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "えええ"));
    }
}
