package xyz.hotchpotch.hogandiff.list;

import java.util.List;
import java.util.Objects;
import java.util.function.ToIntBiFunction;
import java.util.function.ToIntFunction;

import xyz.hotchpotch.hogandiff.common.Pair;

/**
 * 2つのリストを比較して最適な対応関係を返す関数を表します。
 * 「最適な」の定義は指定される比較条件により異なります。<br>
 * これは、{@link #correlate(List, List)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @param <T> 比較対象リストの要素の型
 * @since 0.1.0
 * @author nmby
 */
@FunctionalInterface
public interface Correlator<T> {
    
    // [static members] ********************************************************
    
    /**
     * 要素の並び順の入れ替えを伴いながら対応関係を求める {@link Correlator} オブジェクトを返します。<br>
     * 
     * @param <T> 比較対象リストの要素の型
     * @param gapEvaluator 余剰（欠損）コスト計算関数
     * @param diffEvaluator 差分コスト計算関数
     * @return 要素の並び順の入れ替えを伴いながら対応関係を求める {@link Correlator} オブジェクト
     * @throws NullPointerException {@code gapEvaluator}, {@code diffEvaluator} のいずれかが {@code null} の場合
     */
    public static <T> Correlator<T> withShuffling(
            ToIntFunction<? super T> gapEvaluator,
            ToIntBiFunction<? super T, ? super T> diffEvaluator) {
        
        Objects.requireNonNull(gapEvaluator, "gapEvaluator");
        Objects.requireNonNull(diffEvaluator, "diffEvaluator");
        
        return new ShufflingCorrelator<>(gapEvaluator, diffEvaluator);
    }
    
    /**
     * 要素の余剰（欠損）を考慮し、要素順を保ったまま対応関係を求める {@link Correlator} オブジェクトを返します。<br>
     * 
     * @param <T> 比較対象リストの要素の型
     * @param gapEvaluator 余剰（欠損）コスト計算関数
     * @param diffEvaluator 差分コスト計算関数
     * @return 要素の余剰（欠損）を考慮し、要素順を保ったまま対応関係を求める {@link Correlator} オブジェクト
     * @throws NullPointerException {@code gapEvaluator}, {@code diffEvaluator} のいずれかが {@code null} の場合
     */
    public static <T> Correlator<T> consideringGaps(
            ToIntFunction<? super T> gapEvaluator,
            ToIntBiFunction<? super T, ? super T> diffEvaluator) {
        
        Objects.requireNonNull(gapEvaluator, "gapEvaluator");
        Objects.requireNonNull(diffEvaluator, "diffEvaluator");
        
        return new SequentialCorrelator<>(gapEvaluator, diffEvaluator);
    }
    
    // [instance members] ******************************************************
    
    /**
     * 比較対象の2つのリストを受け取り、最適な対応関係を表すリストを返します。<br>
     * 
     * @param listA 比較対象リストA
     * @param listB 比較対象リストB
     * @return リストAとリストBの要素の最適な対応関係を表すインデクスのペアのリスト
     */
    List<Pair<Integer>> correlate(List<? extends T> listA, List<? extends T> listB);
}
