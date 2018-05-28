package xyz.hotchpotch.hogandiff.diff.list;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import java.util.function.ToIntBiFunction;
import java.util.function.ToIntFunction;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import xyz.hotchpotch.hogandiff.common.Pair;

// TODO: 最小費用流問題のアルゴリズムで書き直す
/**
 * 比較対象リストの要素順に関わりなく、最も一致度の高い要素同士からペアリングしていく {@link Correlator} の実装です。<br>
 * 
 * @param <T> 比較対象要素の型
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class ShufflingCorrelator<T> implements Correlator<T> {
    
    // [static members] ********************************************************
    
    /**
     * 要素同士の差分コストもしくは要素単独の余剰コストを保持する、内部計算用のクラスです。<br>
     * 
     * @author nmby
     * @since 0.1.0
     */
    private static class Cost implements Comparable<Cost> {
        
        /** 比較対象リストAにおける比較対象要素Aのインデックス（欠損の場合はnull） */
        private final Integer idxA;
        
        /** 比較対象リストBにおける比較対象要素Bのインデックス（欠損の場合はnull） */
        private final Integer idxB;
        
        /** 差分コストもしくは余剰コスト */
        private final int cost;
        
        private Cost(Integer idxA, Integer idxB, int cost) {
            assert idxA != null || idxB != null;
            assert idxA == null || 0 <= idxA;
            assert idxB == null || 0 <= idxB;
            
            this.idxA = idxA;
            this.idxB = idxB;
            this.cost = cost;
        }
        
        private boolean isPaired() {
            return idxA != null && idxB != null;
        }
        
        /**
         * 次の優先順位で大小を判断します。<br>
         * <ol>
         *   <li>コストが異なる場合は、コストの小さい方を「小さい」と判断</li>
         *   <li>一方のみがペアリング済みの場合は、ペアリング済みの方を「小さい」と判断</li>
         *   <li>双方ともにペアリング済みでペア同士の距離が異なる場合は、距離の小さい方を「小さい」と判断</li>
         *   <li>双方ともにペアリング済みでインデックス値の合計が異なる場合は、合計の小さい方を「小さい」と判断</li>
         *   <li>双方ともにペアリング済みでここまでで大小が決まらない場合は、idxA の小さい方を「小さい」と判断</li>
         *   <li>双方ともに単独で要素Aの有無が異なる場合は、要素Aの存在する方を「小さい」と判断</li>
         *   <li>双方ともに単独でここまでで大小が決まらない場合は、存在する idx の小さい方を「小さい」と判断</li>
         * </ol>
         */
        @Override
        public int compareTo(Cost other) {
            Objects.requireNonNull(other);
            
            if (cost != other.cost) {
                // コストそのものが異なる場合は、それに基づいて比較する。
                return cost < other.cost ? -1 : 1;
            }
            if (isPaired() && other.isPaired()) {
                // コストが同じでともにペアリング済みの場合
                int iA = idxA;
                int iB = idxB;
                int oA = other.idxA;
                int oB = other.idxB;
                
                if (Math.abs(iA - iB) != Math.abs(oA - oB)) {
                    // ペア同士の距離が異なる場合はそれに基づいて比較する。
                    return Math.abs(iA - iB) < Math.abs(oA - oB) ? -1 : 1;
                }
                if (iA + iB != oA + oB) {
                    // 原点からの距離の和が異なる場合はそれに基づいて比較する。
                    return iA + iB < oA + oB ? -1 : 1;
                }
                if (iA != oA) {
                    // 最後は idxA に基づいて比較する。
                    return iA < oA ? -1 : 1;
                }
            } else if (isPaired() != other.isPaired()) {
                // コストが同じで片方のみがペアリング済みの場合
                return isPaired() ? -1 : 1;
                
            } else {
                // コストが同じでともに単独の場合
                if ((idxA == null) != (other.idxA == null)) {
                    // idxA と idxB では idxA を優先する。
                    return idxA != null ? -1 : 1;
                }
                int i = idxA != null ? idxA : idxB;
                int o = other.idxA != null ? other.idxA : other.idxB;
                if (i != o) {
                    // 原点からの距離が異なる場合はそれに基づいて比較する。
                    return i < o ? -1 : 1;
                }
            }
            // ここまでは辿り着かないはず。
            throw new AssertionError(String.format("this:%s, other:%s", this, other));
        }
        
        @Override
        public String toString() {
            return String.format("(%s, %s, %d)", idxA, idxB, cost);
        }
    }
    
    // [instance members] ******************************************************
    
    private final ToIntFunction<? super T> gapEvaluator;
    private final ToIntBiFunction<? super T, ? super T> diffEvaluator;
    
    /*package*/ ShufflingCorrelator(
            ToIntFunction<? super T> gapEvaluator,
            ToIntBiFunction<? super T, ? super T> diffEvaluator) {
        
        Objects.requireNonNull(gapEvaluator, "gapEvaluator");
        Objects.requireNonNull(diffEvaluator, "diffEvaluator");
        
        this.gapEvaluator = gapEvaluator;
        this.diffEvaluator = diffEvaluator;
    }
    
    /**
     * {@inheritDoc}
     * この実装は、比較対象リストの要素順に関わりなく、最も一致度の高い要素から順にペアリングさせます。<br>
     * 
     * @throws NullPointerException {@code listA}, {@code listB} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code listA}, {@code listB} が同一インスタンスの場合
     */
    @Override
    public List<Pair<Integer>> correlate(List<? extends T> listA, List<? extends T> listB) {
        Objects.requireNonNull(listA, "listA");
        Objects.requireNonNull(listB, "listB");
        if (listA == listB) {
            throw new IllegalArgumentException("listA == listB.");
        }
        
        // まず、すべての組み合わせのコストを計算する。
        Stream<Cost> gapCostsA = IntStream.range(0, listA.size()).parallel()
                .mapToObj(i -> new Cost(i, null, gapEvaluator.applyAsInt(listA.get(i))));
        
        Stream<Cost> gapCostsB = IntStream.range(0, listB.size()).parallel()
                .mapToObj(j -> new Cost(null, j, gapEvaluator.applyAsInt(listB.get(j))));
        
        Stream<Cost> diffCosts = IntStream.range(0, listA.size()).parallel().mapToObj(Integer::valueOf)
                .flatMap(i -> IntStream.range(0, listB.size()).parallel()
                        .mapToObj(j -> new Cost(i, j, diffEvaluator.applyAsInt(listA.get(i), listB.get(j)))));
        
        // これらを統合し、小さい順にソートする。
        LinkedList<Cost> costs = Stream.concat(Stream.concat(gapCostsA, gapCostsB), diffCosts)
                .sorted()
                .collect(Collectors.toCollection(LinkedList::new));
        
        List<Pair<Integer>> pairs = new ArrayList<>();
        while (0 < costs.size()) {
            Cost cost = costs.removeFirst();
            
            // 小さいものから結果として採用する。
            pairs.add(Pair.ofNullable(cost.idxA, cost.idxB));
            
            // すでに結果として採用された要素が含まれるものは除去する。
            if (cost.idxA != null) {
                costs.removeIf(c -> cost.idxA.equals(c.idxA));
            }
            if (cost.idxB != null) {
                costs.removeIf(c -> cost.idxB.equals(c.idxB));
            }
        }
        
        costs = null;
        System.gc();
        
        return pairs;
    }
}
