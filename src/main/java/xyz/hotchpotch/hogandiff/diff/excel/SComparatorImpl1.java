package xyz.hotchpotch.hogandiff.diff.excel;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.excel.CellReplica;

/**
 * {@link SComparator} の標準的な実装です。<br>
 * 2つの比較対象Excelシートの行方向、列方向の対応関係をそれぞれ別々に求めたうえでセルの比較を行います。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class SComparatorImpl1 implements SComparator {
    
    // [static members] ********************************************************
    
    // [instance members] ******************************************************
    
    private final boolean considerRowGaps;
    private final boolean considerColumnGaps;
    private final SStrategy rowStrategy;
    private final SStrategy columnStrategy;
    
    /*package*/ SComparatorImpl1(Context context) {
        assert context != null;
        
        considerRowGaps = context.get(Props.APP_CONSIDER_ROW_GAPS);
        considerColumnGaps = context.get(Props.APP_CONSIDER_COLUMN_GAPS);
        
        if (considerRowGaps && considerColumnGaps) {
            rowStrategy = SStrategies.rowStrategy2;
            columnStrategy = SStrategies.columnStrategy2;
        } else if (considerRowGaps) {
            rowStrategy = SStrategies.rowStrategy1;
            columnStrategy = SStrategies.columnStrategy0;
        } else if (considerColumnGaps) {
            rowStrategy = SStrategies.rowStrategy0;
            columnStrategy = SStrategies.columnStrategy1;
        } else {
            rowStrategy = SStrategies.rowStrategy0;
            columnStrategy = SStrategies.columnStrategy0;
        }
    }
    
    /**
     * {@inheritDoc}
     * この実装は、2つの比較対象Excelシートの行方向、列方向の対応関係をそれぞれ別々に求めたうえでセルの比較を行います。<br>
     * 
     * @throws NullPointerException {@code cellsA}, {@code cellsB} のいずれかが {@code null} の場合
     */
    @Override
    public SResult compare(Set<CellReplica> cellsA, Set<CellReplica> cellsB) {
        Objects.requireNonNull(cellsA, "cellsA");
        Objects.requireNonNull(cellsB, "cellsB");
        
        // シート1とシート2の行同士、列同士の対応関係を求める。
        List<Pair<Integer>> rowPairs = rowStrategy.pairing(cellsA, cellsB);
        List<Pair<Integer>> columnPairs = columnStrategy.pairing(cellsA, cellsB);
        
        // 余剰行を収集する。
        List<Integer> redundantRowsA = rowPairs.stream()
                .filter(Pair::isOnlyA)
                .map(Pair::a)
                .collect(Collectors.toList());
        List<Integer> redundantRowsB = rowPairs.stream()
                .filter(Pair::isOnlyB)
                .map(Pair::b)
                .collect(Collectors.toList());
        
        // 余剰列を収集する。
        List<Integer> redundantColumnsA = columnPairs.stream()
                .filter(Pair::isOnlyA)
                .map(Pair::a)
                .collect(Collectors.toList());
        List<Integer> redundantColumnsB = columnPairs.stream()
                .filter(Pair::isOnlyB)
                .map(Pair::b)
                .collect(Collectors.toList());
        
        // 差分セルを収集する。
        List<Pair<CellReplica>> diffCells = compareCells(
                cellsA, cellsB, rowPairs, columnPairs);
        
        return SResult.of(
                considerRowGaps,
                considerColumnGaps,
                redundantRowsA,
                redundantRowsB,
                redundantColumnsA,
                redundantColumnsB,
                diffCells);
    }
    
    /**
     * 指定された行同士、列同士の対応関係に従ってシートAとシートBのセルを比較し、差分セルのペアのリストを返す。<br>
     * 
     * @param cells1 比較対象ExcelシートAのセルセット
     * @param cells2 比較対象ExcelシートBのセルセット
     * @param rowPairs 行同士の対応関係
     * @param columnPairs 列同士の対応関係
     * @return 差分セルのペアのリスト
     */
    private List<Pair<CellReplica>> compareCells(
            Set<CellReplica> cellsA,
            Set<CellReplica> cellsB,
            List<Pair<Integer>> rowPairs,
            List<Pair<Integer>> columnPairs) {
        
        assert cellsA != null;
        assert cellsB != null;
        assert rowPairs != null;
        assert columnPairs != null;
        
        Map<String, CellReplica> mapA = cellsA.stream()
                .collect(Collectors.toMap(CellReplica::address, Function.identity()));
        Map<String, CellReplica> mapB = cellsB.stream()
                .collect(Collectors.toMap(CellReplica::address, Function.identity()));
        
        return rowPairs.parallelStream().filter(Pair::isPaired).flatMap(rp -> {
            int rowA = rp.a();
            int rowB = rp.b();
            
            return columnPairs.stream().filter(Pair::isPaired).map(cp -> {
                int columnA = cp.a();
                int columnB = cp.b();
                String addrA = CellReplica.getAddress(rowA, columnA);
                String addrB = CellReplica.getAddress(rowB, columnB);
                CellReplica cellA = mapA.get(addrA);
                CellReplica cellB = mapB.get(addrB);
                String valueA = (cellA == null ? "" : cellA.value());
                String valueB = (cellB == null ? "" : cellB.value());
                
                return valueA.equals(valueB)
                        ? null
                        : Pair.of(
                                Optional.ofNullable(cellA).orElseGet(() -> CellReplica.of(addrA, "")),
                                Optional.ofNullable(cellB).orElseGet(() -> CellReplica.of(addrB, "")));
            });
        }).filter(p -> p != null).collect(Collectors.toList());
    }
}
