package xyz.hotchpotch.hogandiff.list;

import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import java.util.function.ToIntBiFunction;
import java.util.function.ToIntFunction;
import java.util.stream.IntStream;

import xyz.hotchpotch.hogandiff.common.Pair;

/**
 * 2つのリストをそれぞれの要素の並び順を保ったまま、要素の余剰（欠損）を考慮して
 * 最も一致度が高くなるようにペアリングさせる {@link Correlator} の実装です。<br>
 * 
 * @param <T> 比較対象要素の型
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class SequentialCorrelator<T> implements Correlator<T> {
    
    // [static members] ********************************************************
    
    /**
     * 内部処理用の列挙型です。
     * リストA、リストBの二次元比較マップ上の各格子点における最適遷移方向を表します。<br>
     * 
     * @since 0.1.0
     * @author nmby
     */
    private static enum ComeFrom {
        
        /**
         * 比較マップを左上から右下に遷移すること、すなわち、
         * リストAの要素とリストBの要素が対応することを表します。<br>
         */
        UPPERLEFT,
        
        /**
         * 比較マップを上から下に遷移すること、すなわち、
         * リストAの要素が余剰でありリストBの要素が欠損していることを表します。<br>
         */
        UPPER,
        
        /**
         * 比較マップを左から右に遷移すること、すなわち、
         * リストAの要素が欠損しておりリストBの要素が余剰であることを表します。<br>
         */
        LEFT;
    }
    
    // [instance members] ******************************************************
    
    private final ToIntFunction<? super T> gapEvaluator;
    private final ToIntBiFunction<? super T, ? super T> diffEvaluator;
    
    /*package*/ SequentialCorrelator(
            ToIntFunction<? super T> gapEvaluator,
            ToIntBiFunction<? super T, ? super T> diffEvaluator) {
        
        Objects.requireNonNull(gapEvaluator, "gapEvaluator");
        Objects.requireNonNull(diffEvaluator, "diffEvaluator");
        
        this.gapEvaluator = gapEvaluator;
        this.diffEvaluator = diffEvaluator;
    }
    
    /**
     * {@inheritDoc}
     * この実装は、2つのリストそれぞれの要素の並び順を保ったまま、要素の余剰（欠損）を考慮して
     * 最も一致度が高くなるようにペアリングさせます。<br>
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
        
        // 遷移コストの計算
        int[][] costs = calcCosts(listA, listB);
        
        // 最適遷移方向の計算
        ComeFrom[][] bestDirections = calcBestDirections(costs);
        costs = null;
        
        // 最適ルートの収穫
        List<Pair<Integer>> bestRoute = harvestBestRoute(bestDirections);
        bestDirections = null;
        
        System.gc();
        
        return bestRoute;
    }
    
    /**
     * 比較対象のリストそれぞれの要素の余剰（欠損）コスト、要素同士の差分コストを計算して二次元配列で返します。<br>
     * 返される二次元配列の {@code [i + 1][0]} はリストAの要素 {@code i} の余剰（欠損）コストを、
     * {@code [0][j + 1]} はリストBの要素 {@code j} の余剰（欠損）コストを、
     * {@code [i + 1][j + 1]} はリストAの要素 {@code i} とリストBの要素 {@code j} の差分コストを表します。<br>
     * 返される二次元配列の {@code [0][0]} の意味は定義されず、値は不定です。<br>
     * 
     * @param listA 比較対象リストA
     * @param listB 比較対象リストB
     * @return 余剰（欠損）コスト、差分コストが格納された二次元配列
     */
    private int[][] calcCosts(List<? extends T> listA, List<? extends T> listB) {
        assert listA != null;
        assert listB != null;
        
        int[][] costs = new int[listA.size() + 1][listB.size() + 1];
        
        IntStream.range(0, listA.size()).parallel().forEach(
                i -> costs[i + 1][0] = gapEvaluator.applyAsInt(listA.get(i)));
        IntStream.range(0, listB.size()).parallel().forEach(
                j -> costs[0][j + 1] = gapEvaluator.applyAsInt(listB.get(j)));
        IntStream.range(0, listA.size()).parallel().forEach(
                i -> IntStream.range(0, listB.size()).parallel().forEach(
                        j -> costs[i + 1][j + 1] = diffEvaluator.applyAsInt(listA.get(i), listB.get(j))));
        
        return costs;
    }
    
    /**
     * 余剰（欠損）コスト、差分コストが格納された二次元配列を受け取り、
     * それぞれの格子点における最適遷移方向を計算して二次元配列で返します。<br>
     * 
     * @param costs 余剰（欠損）コスト、差分コストが格納された二次元配列
     * @return それぞれの格子点における最適遷移方向が格納された二次元配列
     */
    private ComeFrom[][] calcBestDirections(int[][] costs) {
        assert costs != null;
        
        long[][] accumulatedCosts = new long[costs.length][costs[0].length];
        ComeFrom[][] bestDirections = new ComeFrom[costs.length][costs[0].length];
        
        for (int i = 1; i < costs.length; i++) {
            accumulatedCosts[i][0] = accumulatedCosts[i - 1][0] + costs[i][0];
            bestDirections[i][0] = ComeFrom.UPPER;
        }
        for (int j = 1; j < costs[0].length; j++) {
            accumulatedCosts[0][j] = accumulatedCosts[0][j - 1] + costs[0][j];
            bestDirections[0][j] = ComeFrom.LEFT;
        }
        // 比較対象リストが長くなるほど、すなわち二次元比較マップ（探索平面）が広くなるほど
        // 処理の並列化が効果を発揮すると信じて、処理を並列化する。
        // 縦方向、横方向には並列化できないため、探索平面を斜めにスライスして処理を並列化する。
        for (int n = 2; n < costs.length + costs[0].length - 1; n++) {
            final int nn = n;
            IntStream.rangeClosed(Math.max(1, nn - costs[0].length + 1), Math.min(nn - 1, costs.length - 1))
                    .parallel().forEach(i -> {
                        int j = nn - i;
                        long minCost = accumulatedCosts[i - 1][j - 1] + costs[i][j];
                        ComeFrom minDirection = ComeFrom.UPPERLEFT;
                        long tmpCost = accumulatedCosts[i][j - 1] + costs[0][j];
                        if (tmpCost < minCost) {
                            minCost = tmpCost;
                            minDirection = ComeFrom.LEFT;
                        }
                        tmpCost = accumulatedCosts[i - 1][j] + costs[i][0];
                        if (tmpCost < minCost) {
                            minCost = tmpCost;
                            minDirection = ComeFrom.UPPER;
                        }
                        accumulatedCosts[i][j] = minCost;
                        bestDirections[i][j] = minDirection;
                    });
        }
        return bestDirections;
    }
    
    /**
     * それぞれの格子点における最適遷移方向が格納された二次元配列を受け取り、
     * 始点 {@code (0, 0)} から終点 {@code (i, j)} までの最適ルートを表すリストを返します。<br>
     * 
     * @param bestDirections それぞれの格子点における最適遷移方向が格納された二次元配列
     * @return 始点 {@code (0, 0)} から終点 {@code (i, j)} までの最適ルートを表すリスト
     */
    private List<Pair<Integer>> harvestBestRoute(ComeFrom[][] bestDirections) {
        assert bestDirections != null;
        
        LinkedList<Pair<Integer>> bestRoute = new LinkedList<>();
        int i = bestDirections.length - 1;
        int j = bestDirections[0].length - 1;
        
        while (0 < i || 0 < j) {
            switch (bestDirections[i][j]) {
            case UPPERLEFT:
                i--;
                j--;
                bestRoute.addFirst(Pair.of(i, j));
                break;
            case UPPER:
                i--;
                bestRoute.addFirst(Pair.onlyA(i));
                break;
            case LEFT:
                j--;
                bestRoute.addFirst(Pair.onlyB(j));
                break;
            default:
                throw new AssertionError(bestDirections[i][j]);
            }
        }
        return bestRoute;
    }
}
