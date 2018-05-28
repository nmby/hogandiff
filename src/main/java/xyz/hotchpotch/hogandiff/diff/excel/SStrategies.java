package xyz.hotchpotch.hogandiff.diff.excel;

import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.function.ToIntBiFunction;
import java.util.function.ToIntFunction;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.diff.list.Correlator;
import xyz.hotchpotch.hogandiff.excel.CellReplica;

/**
 * {@link SStrategy} の標準的な実装を提供するユーティリティクラスです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
/*package*/ class SStrategies {
    
    // [static members] ********************************************************
    
    /** 行の余剰/欠損を考慮しない場合の行関連付け戦略 */
    public static final SStrategy rowStrategy0 = strategy0(CellReplica::row);
    
    /** 列の余剰/欠損を考慮しない場合の列関連付け戦略 */
    public static final SStrategy columnStrategy0 = strategy0(CellReplica::column);
    
    /** 行の余剰/欠損は考慮するが列の余剰/欠損は考慮しない場合の行関連付け戦略 */
    public static final SStrategy rowStrategy1 = strategy12(CellReplica::row, CellReplica::column);
    
    /** 列の余剰/欠損は考慮するが行の余剰/欠損は考慮しない場合の列関連付け戦略 */
    public static final SStrategy columnStrategy1 = strategy12(CellReplica::column, CellReplica::row);
    
    /** 行の余剰/欠損と列の余剰/欠損をともに考慮する場合の行関連付け戦略 */
    public static final SStrategy rowStrategy2 = strategy12(CellReplica::row, CellReplica::value);
    
    /** 行の余剰/欠損と列の余剰/欠損をともに考慮する場合の列関連付け戦略 */
    public static final SStrategy columnStrategy2 = strategy12(CellReplica::column, CellReplica::value);
    
    private static SStrategy strategy0(ToIntFunction<CellReplica> verticality) {
        return (cellsA, cellsB) -> {
            assert cellsA != null;
            assert cellsB != null;
            
            Pair<Integer> range = range(cellsA, cellsB, verticality);
            
            return IntStream.rangeClosed(range.a(), range.b())
                    .mapToObj(n -> Pair.of(n, n))
                    .collect(Collectors.toList());
        };
    }
    
    private static <H extends Comparable<H>> SStrategy strategy12(
            ToIntFunction<CellReplica> verticality,
            Function<CellReplica, H> horizontality) {
        
        return (cellsA, cellsB) -> {
            assert cellsA != null;
            assert cellsB != null;
            
            int start = range(cellsA, cellsB, verticality).a();
            
            Function<Set<CellReplica>, List<List<CellReplica>>> converter = cells -> {
                int end = range(cells, verticality).b();
                Map<Integer, List<CellReplica>> map = cells.stream()
                        .filter(c -> !"".equals(c.value()))
                        .collect(Collectors.groupingBy(verticality::applyAsInt));
                
                return IntStream.rangeClosed(start, end).parallel()
                        .mapToObj(i -> {
                            if (map.containsKey(i)) {
                                List<CellReplica> list = map.get(i);
                                list.sort(Comparator.comparing(horizontality));
                                return list;
                            } else {
                                return Collections.<CellReplica>emptyList();
                            }
                        })
                        .collect(Collectors.toList());
            };
            
            List<List<CellReplica>> listA = converter.apply(cellsA);
            List<List<CellReplica>> listB = converter.apply(cellsB);
            
            Correlator<List<CellReplica>> correlator = Correlator.consideringGaps(
                    gapEvaluator(),
                    diffEvaluator(horizontality));
            
            return correlator.correlate(listA, listB).stream()
                    .map(p -> p.map(i -> i + start))
                    .collect(Collectors.toList());
        };
    }
    
    private static ToIntFunction<List<CellReplica>> gapEvaluator() {
        return List::size;
    }
    
    private static <K extends Comparable<K>> ToIntBiFunction<List<CellReplica>, List<CellReplica>> diffEvaluator(
            Function<CellReplica, K> keyExtractor) {
        
        return (cellsA, cellsB) -> {
            Iterator<CellReplica> itrA = cellsA.iterator();
            Iterator<CellReplica> itrB = cellsB.iterator();
            
            int diff = 0;
            int c = 0;
            K keyA = null;
            K keyB = null;
            String valueA = null;
            String valueB = null;
            
            while (itrA.hasNext() && itrB.hasNext()) {
                if (c <= 0) {
                    CellReplica cellA = itrA.next();
                    keyA = keyExtractor.apply(cellA);
                    valueA = cellA.value();
                }
                if (0 <= c) {
                    CellReplica cellB = itrB.next();
                    keyB = keyExtractor.apply(cellB);
                    valueB = cellB.value();
                }
                c = keyA.compareTo(keyB);
                if (c == 0 && !valueA.equals(valueB)) {
                    diff += 2;
                } else if (c != 0) {
                    diff++;
                }
            }
            while (itrA.hasNext()) {
                diff++;
                itrA.next();
            }
            while (itrB.hasNext()) {
                diff++;
                itrB.next();
            }
            
            return diff;
        };
    }
    
    private static Pair<Integer> range(
            Set<CellReplica> cells,
            ToIntFunction<CellReplica> extractor) {
        
        assert cells != null;
        assert extractor != null;
        
        int min = cells.stream()
                .mapToInt(extractor)
                .min().orElse(0);
        int max = cells.stream()
                .mapToInt(extractor)
                .max().orElse(0);
        
        return Pair.of(min, max);
    }
    
    private static Pair<Integer> range(
            Set<CellReplica> cellsA,
            Set<CellReplica> cellsB,
            ToIntFunction<CellReplica> extractor) {
        
        assert cellsA != null;
        assert cellsB != null;
        assert extractor != null;
        
        Pair<Integer> rangeA = range(cellsA, extractor);
        Pair<Integer> rangeB = range(cellsB, extractor);
        
        return Pair.of(
                Math.min(rangeA.a(), rangeB.a()),
                Math.max(rangeA.b(), rangeB.b()));
    }
    
    // [instance members] ******************************************************
    
    private SStrategies() {
    }
}
