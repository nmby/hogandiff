package xyz.hotchpotch.hogandiff.diff.list;

import java.util.List;
import java.util.Objects;
import java.util.function.ToIntBiFunction;
import java.util.function.ToIntFunction;
import java.util.stream.Collectors;

/**
 * {@link Correlator} と組み合わせて利用するための典型的な各種評価関数を集めたユーティリティクラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class Evaluators {
    
    // [static members] ********************************************************
    
    /** 文字列の余剰（欠損）コストを返す関数です。文字列の長さをコストとして評価します。  */
    public static final ToIntFunction<String> stringGapEvaluator = String::length;
    
    /** 文字列同士の差分コストを返す関数です。一致しない文字の数をコストとして評価します。 */
    public static final ToIntBiFunction<String, String> stringDiffEvaluator = (str1, str2) -> {
        Objects.requireNonNull(str1, "str1");
        Objects.requireNonNull(str2, "str2");
        
        List<Integer> list1 = str1.codePoints().boxed().collect(Collectors.toList());
        List<Integer> list2 = str2.codePoints().boxed().collect(Collectors.toList());
        
        Correlator<Integer> correlator = Correlator.consideringGaps(
                i -> 1,
                (i1, i2) -> Objects.equals(i1, i2) ? 0 : Integer.MAX_VALUE);
        
        return (int) correlator.correlate(list1, list2).stream()
                .filter(p -> !p.isPaired())
                .count();
    };
    
    // [instance members] ******************************************************
    
    private Evaluators() {
    }
}
