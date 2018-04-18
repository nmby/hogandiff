package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.StreamSupport;

import org.apache.poi.hssf.usermodel.HSSFSheetConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.common.CellReplica;

/**
 * Excelブックに対する各種操作をExcelブックの形式に関わりなく透過的に提供するユーティリティクラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class POIUtils {
    
    // [static members] ********************************************************
    
    /**
     * 指定されたExcelブックに含まれるシート名の一覧を返します。<br>
     * 
     * @param book 対象のExcelブック
     * @return シート名の一覧
     * @throws ApplicationException 処理に失敗した場合
     */
    public static List<String> getSheetNames(File book) throws ApplicationException {
        SheetLister lister = SheetLister.of(book);
        return lister.getSheetNames(book);
    }
    
    /**
     * 指定されたExcelブックからシートデータを読み込んでセルデータのセットとして返します。<br>
     * 
     * @param book 対象のExcelブック
     * @param sheetName 対象のシート名
     * @param extractCachedValue 数式セルから数式文字列ではなくキャッシュされた計算値を取得する場合は {@code true}
     * @return シートに含まれるセルデータのセット
     * @throws ApplicationException 処理に失敗した場合
     */
    public static Set<CellReplica> loadSheet(
            File book,
            String sheetName,
            boolean extractCachedValue)
            throws ApplicationException {
        
        SheetLoader loader = SheetLoader.of(book, extractCachedValue);
        return loader.loadSheet(book, sheetName);
    }
    
    /**
     * 指定されたExcelブックの以下の色をクリアします。<br>
     * <ul>
     *   <li>罫線の色</li>
     *   <li>セル背景色</li>
     *   <li>フォント色</li>
     * </ul>
     * @param book 対象のExcelブック
     * @throws NullPointerException {@code book} が {@code null} の場合
     */
    public static void clearColors(Workbook book) {
        Objects.requireNonNull(book, "book");
        
        if (book instanceof XSSFWorkbook) {
            clearColors((XSSFWorkbook) book);
        } else if (book instanceof HSSFWorkbook) {
            clearColors((HSSFWorkbook) book);
        } else {
            throw new AssertionError("bookType == " + book.getClass().getName());
        }
    }
    
    private static void clearColors(XSSFWorkbook book) {
        assert book != null;
        
        final short automatic = IndexedColors.AUTOMATIC.getIndex();
        
        // ブックが保持する全スタイルの色をとにかく全てリセットする。
        IntStream.range(0, book.getNumCellStyles()).mapToObj(book::getCellStyleAt).forEach(style -> {
            // 罫線の色
            style.setTopBorderColor(automatic);
            style.setBottomBorderColor(automatic);
            style.setLeftBorderColor(automatic);
            style.setRightBorderColor(automatic);
            
            // セルの背景色とパターン
            style.setFillForegroundColor(null);
            style.setFillBackgroundColor(null);
            if (style.getFillPatternEnum() == FillPatternType.SOLID_FOREGROUND) {
                style.setFillPattern(FillPatternType.NO_FILL);
            }
            
            // セルの塗りつぶし効果
            // -> 設定方法が分からず... orz
            // TODO: セルの塗りつぶし効果をリセットする処理を実装する
        });
        
        // ブックが保持する全フォントの色をとにかく全てリセットする。
        IntStream.range(0, book.getNumberOfFonts()).mapToObj(i -> book.getFontAt((short) i)).forEach(font -> {
            font.setColor(null);
        });
        
        // シート見出しの色をリセットする。
        IntStream.range(0, book.getNumberOfSheets()).mapToObj(book::getSheetAt).forEach(sheet -> {
            // sheet.setTabColor(new XSSFColor(java.awt.Color.WHITE));
            // 白に設定する方法はあるのだが、色をリセット方法が分からん... orz
            // TODO: シート見出しの色をリセットする処理を実装する
        });
        
        // 条件付き書式を全て削除する。
        // （条件付き書式の色を削除するのは非常に面倒なため、条件付き書式自体を削除してしまうことにする。）
        // TODO: 条件付き書式自体ではなく条件付き書式の色を削除するようにそのうち改修する。
        IntStream.range(0, book.getNumberOfSheets()).mapToObj(book::getSheetAt)
                .forEach(sheet -> {
                    XSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
                    while (0 < scf.getNumConditionalFormattings()) {
                        // もしかしたら後ろから消してった方が速いのかもしれないけれど、気にしない。
                        scf.removeConditionalFormatting(0);
                    }
                });
    }
    
    private static void clearColors(HSSFWorkbook book) {
        assert book != null;
        
        final short automatic = IndexedColors.AUTOMATIC.getIndex();
        
        // ブックが保持する全スタイルの色をとにかく全てリセットする。
        IntStream.range(0, book.getNumCellStyles()).mapToObj(book::getCellStyleAt).forEach(style -> {
            // 罫線の色
            style.setTopBorderColor(automatic);
            style.setBottomBorderColor(automatic);
            style.setLeftBorderColor(automatic);
            style.setRightBorderColor(automatic);
            
            // セルの背景色とパターン
            style.setFillForegroundColor(automatic);
            style.setFillBackgroundColor(automatic);
            if (style.getFillPatternEnum() == FillPatternType.SOLID_FOREGROUND) {
                style.setFillPattern(FillPatternType.NO_FILL);
            }
        });
        
        // ブックが保持する全フォントの色をとにかく全てリセットする。
        IntStream.range(0, book.getNumberOfFonts()).mapToObj(i -> book.getFontAt((short) i)).forEach(font -> {
            font.setColor(Font.COLOR_NORMAL);
            // カスタムカラーだと↑これじゃダメ。それらしいAPIが見つからず方法が分からん... orz
            // TODO: フォントのカスタムカラーをリセットする処理を実装する。
        });
        
        // シート見出しの色をリセットする。
        IntStream.range(0, book.getNumberOfSheets()).mapToObj(book::getSheetAt).forEach(sheet -> {
            // -> それらしいAPIが見つからず設定方法が分からん... orz
            // TODO: シート見出しの色をリセットする処理を実装する
        });
        
        // 条件付き書式を全て削除する。
        // （条件付き書式の色を削除するのは非常に面倒なため、条件付き書式自体を削除してしまうことにする。）
        // TODO: 条件付き書式自体ではなく条件付き書式の色を削除するようにそのうち改修する。
        IntStream.range(0, book.getNumberOfSheets()).mapToObj(book::getSheetAt)
                .forEach(sheet -> {
                    HSSFSheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
                    while (0 < scf.getNumConditionalFormattings()) {
                        // もしかしたら後ろから消してった方が速いのかもしれないけれど、気にしない。
                        scf.removeConditionalFormatting(0);
                    }
                });
    }
    
    /**
     * 指定された行に指定された色を付けます。<br>
     * 
     * @param sheet 対象のExcelシート
     * @param rows 対象行インデックス（0開始）のセット
     * @param color {@link IndexedColors} で定義された色インデックス
     * @throws NullPointerException {@code sheet}, {@code rows} のいずれかが {@code null} の場合
     */
    public static void paintRows(Sheet sheet, Collection<Integer> rows, short color) {
        Objects.requireNonNull(sheet, "sheet");
        Objects.requireNonNull(rows, "rows");
        
        CellStyle newStyle = sheet.getWorkbook().createCellStyle();
        newStyle.setFillForegroundColor(color);
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        rows.forEach(i -> {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
                // おまじない。これをしないと、空行の高さがデフォルト値に変更されてしまう。
                row.setHeight(row.getHeight());
            }
            row.setRowStyle(newStyle);
        });
        
        Set<CellAddress> cellAddresses = rows.stream()
                .map(sheet::getRow)
                .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
                .map(CellAddress::new)
                .collect(Collectors.toSet());
        paintCells(sheet, cellAddresses, color);
    }
    
    /**
     * 指定された列に指定された色を付けます。<br>
     * 
     * @param sheet 対象のExcelシート
     * @param columns 対象列インデックス（0開始）のセット
     * @param color {@link IndexedColors} で定義された色インデックス
     * @throws NullPointerException {@code sheet}, {@code columns} のいずれかが {@code null} の場合
     */
    public static void paintColumns(Sheet sheet, Collection<Integer> columns, short color) {
        Objects.requireNonNull(sheet, "sheet");
        Objects.requireNonNull(columns, "columns");
        
        CellStyle newStyle = sheet.getWorkbook().createCellStyle();
        newStyle.setFillForegroundColor(color);
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // XSSFSheetでstyleの設定されていないカラムの場合、
        // 次のコードにより列の幅がデフォルト幅に変更されてしまうという問題がある。
        // しかし変更前の元の列幅を取得する方法が不明なため、今のところ解消策なし。
        // TODO: XSSFSheetでstyleの設定されていない列の幅が変更される問題を解消する。
        columns.forEach(j -> sheet.setDefaultColumnStyle(j, newStyle));
        
        Set<CellAddress> cellAddresses = StreamSupport.stream(sheet.spliterator(), false)
                .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
                .filter(cell -> columns.contains(cell.getColumnIndex()))
                .map(CellAddress::new)
                .collect(Collectors.toSet());
        paintCells(sheet, cellAddresses, color);
    }
    
    /**
     * 指定されたセルに指定された色を付けます。<br>
     * 
     * @param sheet 対象のExcelシート
     * @param addresses 対象セルアドレスのセット
     * @param color {@link IndexedColors} で定義された色インデックス
     * @throws NullPointerException {@code sheet}, {@code addresses} のいずれかが {@code null} の場合
     */
    public static void paintCells(Sheet sheet, Collection<CellAddress> addresses, short color) {
        Objects.requireNonNull(sheet, "sheet");
        Objects.requireNonNull(addresses, "addresses");
        
        // 指定された位置の行・セルが存在しない場合は、色を付けるために行・セルを作成する。
        addresses.forEach(address -> {
            Row row = sheet.getRow(address.getRow());
            if (row == null) {
                row = sheet.createRow(address.getRow());
            }
            if (row.getCell(address.getColumn()) == null) {
                row.createCell(address.getColumn());
            }
        });
        
        // 同じスタイルを持つセルごとにグルーピングする。
        Map<CellStyle, List<Cell>> map = addresses.stream()
                .map(address -> sheet.getRow(address.getRow()).getCell(address.getColumn()))
                .collect(Collectors.groupingBy(Cell::getCellStyle));
        
        Map<String, Object> styleModifier = new HashMap<>();
        styleModifier.put(CellUtil.FILL_FOREGROUND_COLOR, color);
        styleModifier.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
        
        // 同じスタイルごとに新スタイルへの付け替えを行う。
        map.forEach((style, cells) -> {
            Iterator<Cell> itr = cells.iterator();
            
            assert itr.hasNext();
            Cell first = itr.next();
            CellUtil.setCellStyleProperties(first, styleModifier);
            CellStyle newStyle = first.getCellStyle();
            
            itr.forEachRemaining(cell -> cell.setCellStyle(newStyle));
        });
    }
    
    // [instance members] ******************************************************
    
    private POIUtils() {
    }
}
