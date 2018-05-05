package xyz.hotchpotch.hogandiff;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.list.Correlator;
import xyz.hotchpotch.hogandiff.list.Evaluators;
import xyz.hotchpotch.hogandiff.poi.POIUtils;

/**
 * このアプリケーションの実行メニューを表す列挙型です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public enum Menu {
    
    // [static members] ********************************************************
    
    /** 2つのExcelブックを比較します。 */
    COMPARE_BOOKS {
        
        @Override
        protected boolean isValidTargets(Context context) {
            assert context != null;
            
            String pathStr1 = context.get(Props.CURR_FILE1).getPath();
            String pathStr2 = context.get(Props.CURR_FILE2).getPath();
            
            return !pathStr1.equals(pathStr2);
        }
        
        @Override
        protected List<Pair<String>> getPairsOfSheetNames(Context context) throws ApplicationException {
            assert context != null;
            
            File file1 = context.get(Props.CURR_FILE1);
            File file2 = context.get(Props.CURR_FILE2);
            List<String> sheetNames1 = POIUtils.getSheetNames(file1);
            List<String> sheetNames2 = POIUtils.getSheetNames(file2);
            
            Correlator<String> correlator = Correlator.withShuffling(
                    Evaluators.stringGapEvaluator,
                    Evaluators.stringDiffEvaluator);
            
            List<Pair<Integer>> pairs = correlator.correlate(sheetNames1, sheetNames2);
            
            return pairs.stream()
                    .map(p -> Pair.ofNullable(
                            p.a2().map(sheetNames1::get).orElse(null),
                            p.b2().map(sheetNames2::get).orElse(null)))
                    .collect(Collectors.toList());
        }
    },
    
    /** 2つのExcelシートを比較します。 */
    COMPARE_SHEETS {
        
        @Override
        protected boolean isValidTargets(Context context) {
            assert context != null;
            
            String pathStr1 = context.get(Props.CURR_FILE1).getPath();
            String pathStr2 = context.get(Props.CURR_FILE2).getPath();
            String sheetName1 = context.get(Props.CURR_SHEET_NAME1);
            String sheetName2 = context.get(Props.CURR_SHEET_NAME2);
            
            return !pathStr1.equals(pathStr2) || !sheetName1.equals(sheetName2);
        }
        
        @Override
        protected List<Pair<String>> getPairsOfSheetNames(Context context) throws ApplicationException {
            assert context != null;
            
            String sheetName1 = context.get(Props.CURR_SHEET_NAME1);
            String sheetName2 = context.get(Props.CURR_SHEET_NAME2);
            
            return Arrays.asList(Pair.of(sheetName1, sheetName2));
        }
    };
    
    // [instance members] ******************************************************
    
    /**
     * 処理対象の妥当性を確認します。<br>
     * 
     * @param context コンテキスト
     * @return 処理対象として妥当な場合は {@code true}
     */
    protected abstract boolean isValidTargets(Context context);
    
    /**
     * 比較すべきシート名の組み合わせのリストを返します。<br>
     * 
     * @param context コンテキスト
     * @return 比較すべきシート名の組み合わせのリスト
     * @throws ApplicationException 処理に失敗した場合
     */
    protected abstract List<Pair<String>> getPairsOfSheetNames(Context context) throws ApplicationException;
}
