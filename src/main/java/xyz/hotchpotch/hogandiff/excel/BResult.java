package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

import xyz.hotchpotch.hogandiff.common.Pair;

/**
 * Excelブック同士の比較結果を表す不変クラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class BResult {
    
    // [static members] ********************************************************
    
    private static final String BR = System.lineSeparator();
    
    /**
     * {@link BResult} オブジェクトを生成して返します。<br>
     * 
     * @param file1 比較対象のExcelファイル1
     * @param file2 比較対象のExcelファイル2
     * @param sheetNamePairs 比較対象シート名のペアのリスト
     * @param sResults シート同士の比較結果。キー：比較対象シート名のペア、値：当該ペアの比較結果
     * @return 新しい {@link BResult} オブジェクト
     * @throws NullPointerException
     *      {@code file1}, {@code file2}, {@code sheetNamePairs}, {@code sResults}
     *      のいずれかが {@code null} の場合
     */
    public static BResult of(
            File file1,
            File file2,
            List<Pair<String>> sheetNamePairs,
            Map<Pair<String>, SResult> sResults) {
        
        Objects.requireNonNull(file1, "file1");
        Objects.requireNonNull(file2, "file2");
        Objects.requireNonNull(sheetNamePairs, "sheetNamePairs");
        Objects.requireNonNull(sResults, "sResults");
        
        return new BResult(file1, file2, sheetNamePairs, sResults);
    }
    
    // [instance members] ******************************************************
    
    // 不変なフィールドは getter を設けずに直接公開してしまう。
    // https://www.ibm.com/developerworks/jp/java/library/j-ft4/index.html
    
    /** 比較対象のExcelファイルのペア */
    public final Pair<File> files;
    
    /** 比較対象シート名のペアのリスト */
    public final List<Pair<String>> sheetNamePairs;
    
    /** シート同士の比較結果。キー：比較対象シート名のペア、値：当該ペアの比較結果 */
    public final Map<Pair<String>, SResult> sResults;
    
    private BResult(
            File file1,
            File file2,
            List<Pair<String>> sheetNamePairs,
            Map<Pair<String>, SResult> sResults) {
        
        assert file1 != null;
        assert file2 != null;
        assert sheetNamePairs != null;
        assert sResults != null;
        
        this.files = Pair.of(file1, file2);
        // 不変にするため、防御的コピーしたうえで変更不可コレクションでラップする。
        this.sheetNamePairs = Collections.unmodifiableList(new ArrayList<>(sheetNamePairs));
        this.sResults = Collections.unmodifiableMap(new HashMap<>(sResults));
    }
    
    @Override
    public String toString() {
        StringBuilder str = new StringBuilder();
        
        str.append("ブックA : ").append(files.a().getPath()).append(BR);
        str.append("ブックB : ").append(files.b().getPath()).append(BR);
        str.append(BR);
        
        str.append("■サマリ========================================================================").append(BR);
        str.append(getSummary());
        str.append(BR);
        
        str.append("■詳細==========================================================================").append(BR);
        str.append(getDetail());
        
        return str.toString();
    }
    
    /**
     * 比較結果のサマリを返します。<br>
     * 
     * @return 比較結果のサマリ
     */
    private String getSummary() {
        return getResult(SResult::getSummary);
    }
    
    /**
     * 比較結果の詳細を返します。<br>
     * 
     * @return 比較結果の詳細
     */
    private String getDetail() {
        return getResult(SResult::getDetail);
    }
    
    private String getResult(Function<SResult, String> recorder) {
        return sheetNamePairs.stream().map(p -> {
            StringBuilder str = new StringBuilder();
            str.append("シートA : ").append(p.aOrElse("（なし）")).append(BR);
            str.append("シートB : ").append(p.bOrElse("（なし）")).append(BR);
            
            if (p.isPaired()) {
                str.append(recorder.apply(sResults.get(p)));
            }
            return str;
            
        }).collect(Collectors.joining(BR));
    }
}
