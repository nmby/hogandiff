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
import xyz.hotchpotch.hogandiff.poi.HSSFSheetListerWithEventApi;
import xyz.hotchpotch.hogandiff.poi.HSSFSheetLoaderWithEventApi;

class HSSFSheetLoaderWithEventApiTest {
    
    // [static members] ********************************************************
    
    private static final HSSFSheetLoaderWithEventApi valueLoader = HSSFSheetLoaderWithEventApi.of(true);
    private static final HSSFSheetLoaderWithEventApi formulaLoader = HSSFSheetLoaderWithEventApi.of(false);
    
    // [instance members] ******************************************************
    
    @Test
    void testIsSupported_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> HSSFSheetLoaderWithEventApi.isSupported(null));
    }
    
    @Test
    void testIsSupported() {
        assertTrue(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xls));
        
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsx));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsm));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_xlsb));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal_csv));
        assertFalse(HSSFSheetListerWithEventApi.isSupported(SheetLoader1_normal));
    }
    
    @Test
    void testOf() {
        assertThat(
                valueLoader,
                instanceOf(HSSFSheetLoaderWithEventApi.class));
        
        assertThat(
                formulaLoader,
                instanceOf(HSSFSheetLoaderWithEventApi.class));
    }
    
    @Test
    void testLoadSheet_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xls, null));
        
        assertThrows(
                NullPointerException.class,
                () -> valueLoader.loadSheet(null, "あああ"));
        
        assertThrows(
                IllegalArgumentException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"));
        
        assertThrows(
                NoSuchElementException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xls, "ををを"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xls() throws ApplicationException {
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xls, "あああ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "あああ"));
        
        assertEquals(
                cellsB_value,
                valueLoader.loadSheet(SheetLoader1_normal_xls, "いいい"));
        assertEquals(
                // TODO: never give up!
                cellsB_formula_giveup,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "いいい"));
        
        assertEquals(
                cellsC,
                valueLoader.loadSheet(SheetLoader1_normal_xls, "ううう"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "ううう"));
        
        assertEquals(
                cellsE,
                valueLoader.loadSheet(SheetLoader1_normal_xls, "おおお"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "おおお"));
    }
    
    @Test
    void testLoadSheet_グラフシート() throws ApplicationException {
        // 現状の実装では、グラフシートからは空のセルデータセットが返される。
        // この挙動を仕様として追認することにする。
        assertEquals(
                cellsA,
                valueLoader.loadSheet(SheetLoader1_normal_xls, "えええ"));
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "えええ"));
    }
}
