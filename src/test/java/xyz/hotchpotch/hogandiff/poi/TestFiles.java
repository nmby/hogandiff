package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import xyz.hotchpotch.hogandiff.common.CellReplica;

@SuppressWarnings("javadoc")
public class TestFiles {
    
    // [static members] ********************************************************
    
    public static final File SheetLister1_normal_xlsx = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal.xlsx").getFile());
    
    public static final File SheetLister1_normal_xlsm = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal.xlsm").getFile());
    
    public static final File SheetLister1_normal_xlsb = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal.xlsb").getFile());
    
    public static final File SheetLister1_normal_xls = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal.xls").getFile());
    
    public static final File SheetLister1_normal_csv = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal.csv").getFile());
    
    public static final File SheetLister1_normal = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister1_normal").getFile());
    
    public static final File SheetLister2_namevariations_xlsx = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister2_namevariations.xlsx").getFile());
    
    public static final File SheetLister2_namevariations_xlsm = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister2_namevariations.xlsm").getFile());
    
    public static final File SheetLister2_namevariations_xlsb = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister2_namevariations.xlsb").getFile());
    
    public static final File SheetLister2_namevariations_xls = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister2_namevariations.xls").getFile());
    
    public static final File SheetLister3_dummy_xlsx = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister3_dummy.xlsx").getFile());
    
    public static final File SheetLister3_dummy_xls = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLister3_dummy.xls").getFile());
    
    public static final List<String> sheetNames1 = Arrays.asList(
            "あああ", "いいい", "ううう", "あ");
    
    public static final List<String> sheetNames2 = Arrays.asList(
            "全角　ﾊﾝｶｸ 123 abc ①②③ 1⃣2⃣3⃣　高髙 ★", "!\"#$%&'()-=^~|@`{;+},<.>_ 　");
    
    public static final Map<String, String> sheetsId1 = new HashMap<>();
    static {
        sheetsId1.put("あああ", "rId1");
        sheetsId1.put("いいい", "rId2");
        sheetsId1.put("ううう", "rId3");
        sheetsId1.put("あ", "rId5");
    }
    
    public static final Map<String, String> sheetsId2 = new HashMap<>();
    static {
        sheetsId2.put("全角　ﾊﾝｶｸ 123 abc ①②③ 1⃣2⃣3⃣　高髙 ★", "rId1");
        sheetsId2.put("!\"#$%&'()-=^~|@`{;+},<.>_ 　", "rId2");
    }
    
    public static final File SheetLoader1_normal_xlsx = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal.xlsx").getFile());
    
    public static final File SheetLoader1_normal_xlsm = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal.xlsm").getFile());
    
    public static final File SheetLoader1_normal_xlsb = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal.xlsb").getFile());
    
    public static final File SheetLoader1_normal_xls = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal.xls").getFile());
    
    public static final File SheetLoader1_normal_csv = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal.csv").getFile());
    
    public static final File SheetLoader1_normal = new File(XSSFSheetListerWithEventApiTest.class.getResource(
            "TestSheetLoader1_normal").getFile());
    
    public static final Set<CellReplica> cellsA = Collections.emptySet();
    
    public static final Set<CellReplica> cellsB_base = new HashSet<>(Arrays.asList(
            CellReplica.of("B2", "[b]"),
            CellReplica.of("C2", "true"),
            CellReplica.of("D2", "false"),
            
            CellReplica.of("B3", "[d]"),
            
            CellReplica.of("B5", "[inlineStr]"),
            
            CellReplica.of("B6", "[n]"),
            CellReplica.of("C6", "123"),
            
            CellReplica.of("B7", "[s]"),
            CellReplica.of("C7", " a b c "),
            CellReplica.of("D7", "　あ　い　う　")));
    
    public static final Set<CellReplica> cellsB_value;
    static {
        cellsB_value = new HashSet<>(cellsB_base);
        cellsB_value.addAll(Arrays.asList(
                CellReplica.of("B4", "[e]"),
                CellReplica.of("C4", "#REF!"),
                CellReplica.of("D4", "#REF!"),
                CellReplica.of("E4", "#DIV/0!"),
                CellReplica.of("F4", "#DIV/0!"),
                CellReplica.of("G4", "#NAME?"),
                
                CellReplica.of("B8", "[str]"),
                CellReplica.of("C8", "123abc"),
                CellReplica.of("D8", "579"),
                CellReplica.of("E8", "123")));
    }
    
    public static final Set<CellReplica> cellsB_formula;
    static {
        cellsB_formula = new HashSet<>(cellsB_base);
        cellsB_formula.addAll(Arrays.asList(
                CellReplica.of("B4", "[e]"),
                CellReplica.of("C4", "#REF!"),
                CellReplica.of("D4", "#REF!"),
                CellReplica.of("E4", "1/0"),
                CellReplica.of("F4", "#DIV/0!"),
                CellReplica.of("G4", "#NAME?"),
                
                CellReplica.of("B8", "[str]"),
                CellReplica.of("C8", "\"123\"&\"abc\""),
                CellReplica.of("D8", "123+456"),
                CellReplica.of("E8", "SUM(C6:D6)")));
    }
    
    public static final Set<CellReplica> cellsB_formula_giveup;
    static {
        cellsB_formula_giveup = new HashSet<>(cellsB_base);
        cellsB_formula_giveup.addAll(Arrays.asList(
                CellReplica.of("B4", "[e]"),
                CellReplica.of("C4", "[formula]"),
                CellReplica.of("D4", "#REF!"),
                CellReplica.of("E4", "[formula]"),
                CellReplica.of("F4", "#DIV/0!"),
                CellReplica.of("G4", "#NAME?"),
                
                CellReplica.of("B8", "[str]"),
                CellReplica.of("C8", "[formula]"),
                CellReplica.of("D8", "[formula]"),
                CellReplica.of("E8", "[formula]")));
    }
    
    public static final Set<CellReplica> cellsC = new HashSet<>(Arrays.asList(
            CellReplica.of("A1", "1"),
            CellReplica.of("B2", "2"),
            CellReplica.of("C3", "3")));
    
    public static final Set<CellReplica> cellsE = new HashSet<>(Arrays.asList(
            CellReplica.of("B2", "改行を含む\n文字列"),
            CellReplica.of("B3", ""
                    + "長い文字列（１０００文字）４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●"
                    + "１２３４５６７８９①１２３４５６７８９②１２３４５６７８９③１２３４５６７８９④１２３４５６７８９⑤"
                    + "１２３４５６７８９⑥１２３４５６７８９⑦１２３４５６７８９⑧１２３４５６７８９⑨１２３４５６７８９●")));
    
    // [instance members] ******************************************************
    
    private TestFiles() {
    }
}
