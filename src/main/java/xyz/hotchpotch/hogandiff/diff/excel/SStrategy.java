package xyz.hotchpotch.hogandiff.diff.excel;

import java.util.List;
import java.util.Set;

import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.excel.CellReplica;

/**
 * 2つのExcelシートの行同士または列同士を対応づける戦略を表します。<br>
 * これは、{@link #pairing(Set, Set)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
@FunctionalInterface
/*package*/ interface SStrategy {
    
    // [static members] ********************************************************
    
    // [instance members] ******************************************************
    
    /**
     * 比較対象Excelシートの行同士または列同士の対応関係を返します。<br>
     * 
     * @param cellsA 比較対象ExcelシートAのセルセット
     * @param cellsB 比較対象ExcelシートBのセルセット
     * @return 行同士または列同士の対応関係を表す、インデックスのペアのリスト
     */
    List<Pair<Integer>> pairing(Set<CellReplica> cellsA, Set<CellReplica> cellsB);
}
