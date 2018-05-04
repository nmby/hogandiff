package xyz.hotchpotch.hogandiff.poi;

import static org.hamcrest.CoreMatchers.*;
import static org.junit.Assert.*;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.jupiter.api.Assertions.*;
import static xyz.hotchpotch.hogandiff.poi.TestFiles.*;

import java.util.NoSuchElementException;

import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import xyz.hotchpotch.hogandiff.ApplicationException;

class SheetLoaderWithUserApiTest {
    
    @BeforeAll
    static void setUpBeforeClass() throws Exception {
    }
    
    @AfterAll
    static void tearDownAfterClass() throws Exception {
    }
    
    @BeforeEach
    void setUp() throws Exception {
    }
    
    @AfterEach
    void tearDown() throws Exception {
    }
    
    // [static members] ********************************************************
    
    private static final SheetLoaderWithUserApi valueLoader = SheetLoaderWithUserApi.of(true);
    private static final SheetLoaderWithUserApi formulaLoader = SheetLoaderWithUserApi.of(false);
    
    private static final HSSFSheetLoaderWithEventApi hssfValueLoader = HSSFSheetLoaderWithEventApi.of(true);
    private static final XSSFSheetLoaderWithEventApi xssfValueLoader = XSSFSheetLoaderWithEventApi.of(true);
    
    private static final XSSFSheetLoaderWithEventApi xssfFormulaLoader = XSSFSheetLoaderWithEventApi.of(false);
    
    // [instance members] ******************************************************
    
    @Test
    void testIsSupported_パラメータ不正() {
        assertThrows(
                NullPointerException.class,
                () -> SheetLoaderWithUserApi.isSupported(null));
    }
    
    @Test
    void testIsSupported() {
        assertTrue(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal_xls));
        assertTrue(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal_xlsx));
        assertTrue(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal_xlsm));
        
        assertFalse(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal_xlsb));
        assertFalse(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal_csv));
        assertFalse(SheetLoaderWithUserApi.isSupported(SheetLoader1_normal));
    }
    
    @Test
    void testOf() {
        assertThat(
                valueLoader,
                instanceOf(SheetLoaderWithUserApi.class));
        
        assertThat(
                formulaLoader,
                instanceOf(SheetLoaderWithUserApi.class));
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
                () -> valueLoader.loadSheet(SheetLoader1_normal_xlsb, "あああ"));
        
        assertThrows(
                NoSuchElementException.class,
                () -> valueLoader.loadSheet(SheetLoader1_normal_xls, "ををを"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xls() throws ApplicationException {
        assertEquals(
                hssfValueLoader.loadSheet(SheetLoader1_normal_xls, "あああ"),
                valueLoader.loadSheet(SheetLoader1_normal_xls, "あああ"));
        assertEquals(
                hssfValueLoader.loadSheet(SheetLoader1_normal_xls, "いいい"),
                valueLoader.loadSheet(SheetLoader1_normal_xls, "いいい"));
        assertEquals(
                hssfValueLoader.loadSheet(SheetLoader1_normal_xls, "ううう"),
                valueLoader.loadSheet(SheetLoader1_normal_xls, "ううう"));
        assertEquals(
                hssfValueLoader.loadSheet(SheetLoader1_normal_xls, "おおお"),
                valueLoader.loadSheet(SheetLoader1_normal_xls, "おおお"));
        
        assertEquals(
                cellsA,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "あああ"));
        assertEquals(
                cellsB_formula,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "いいい"));
        assertEquals(
                cellsC,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "ううう"));
        assertEquals(
                cellsE,
                formulaLoader.loadSheet(SheetLoader1_normal_xls, "おおお"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xlsx() throws ApplicationException {
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"));
        
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "あああ"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "いいい"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "ううう"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsx, "おおお"));
    }
    
    @Test
    void testLoadSheet_通常ケース_xlsm() throws ApplicationException {
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"));
        assertEquals(
                xssfValueLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"),
                valueLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"));
        
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "あああ"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "いいい"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "ううう"));
        assertEquals(
                xssfFormulaLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"),
                formulaLoader.loadSheet(SheetLoader1_normal_xlsm, "おおお"));
    }
}
