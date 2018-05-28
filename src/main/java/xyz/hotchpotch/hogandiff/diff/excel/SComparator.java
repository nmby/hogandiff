package xyz.hotchpotch.hogandiff.diff.excel;

import java.util.Objects;
import java.util.Set;

import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.excel.CellReplica;

/**
 * 2つのExcelシートを比較して結果を返す関数を表します。<br>
 * これは、{@link #compare(Set, Set)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
@FunctionalInterface
public interface SComparator {
    
    // [static members] ********************************************************
    
    /**
     * 指定された条件（コンテキスト）に適した {@link SComparator} オブジェクトを返します。<br>
     * 
     * @param context コンテキスト
     * @return {@link SComparator} オブジェクト
     * @throws NullPointerException {@code context} が {@code null} の場合
     */
    public static SComparator of(Context context) {
        Objects.requireNonNull(context, "context");
        return new SComparatorImpl1(context);
    }
    
    // [instance members] ******************************************************
    
    /**
     * 2つのExcelシートを比較して結果を返します。<br>
     * 
     * @param cellsA 比較対象シートAのセルセット
     * @param cellsB 比較対象シートBのセルセット
     * @return 比較結果
     */
    SResult compare(Set<CellReplica> cellsA, Set<CellReplica> cellsB);
}
