package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Collection;
import java.util.Date;
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
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
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

/**
 * ExcelブックやExcelシートに対する各種操作をExcelブックの形式に関わりなく透過的に提供するユーティリティクラスです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
public class ExcelUtils {
    
    // [static members] ********************************************************
    
    private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter
            .ofPattern("yyyy/MM/dd HH:mm:ss.SSS");
    
    /**
     * セルの形式が何であれ、セルの格納値を表す文字列を返します。<br>
     * セルの形式が数式であり {@code returnCachedValue} が{@code true} の場合は、計算結果の値の文字列表現を返します。
     * そうでない場合は、数式の文字列（例えば {@code "SUM(A4:B4)"}）を返します。<br>
     * セルが空の場合は空文字列（{@code ""}）を返します。このメソッドが {@code null} を返すことはありません。<br>
     * <br>
     * このメソッドは、セルの表示形式は無視します。
     * 例えば、セルの保持する値が {@code 3.14159} でありセルの表示形式が {@code "0.00"} のとき、
     * このメソッドは {@code "3.14"} ではなく {@code "3.141592"} をクライアントに返します。<br>
     * <br>
     * 日付の扱いに関する補足：<br>
     * このメソッドは、日付または時刻のフォーマットが指定されているセルの値を
     * {@code "yyyy/MM/dd HH:mm:ss.SSS"} 形式の文字列に変換して返します。<br>
     * Excelにおける日付・時刻の扱いは独特です。<br>
     * 例えばセルに「{@code 10:27}」と入力したとき、Excel内部ではセルの値は「{@code 0.435417}」として管理されます。
     * 一方で、Excelでは {@code 1900/1/1 00:00} が「{@code 1}」で表されます。<br>
     * 従って、「{@code 0.435417}」というセル値を {@code "yyyy/MM/dd HH:mm:ss.SSS"} という形式で評価すると、
     * {@code "1899/12/31 10:27:00.000"} という文字列になります。<br>
     * 以上の理由により、このメソッドは「{@code 10:27}」と入力されたセルの値を
     * {@code "1899/12/31 10:27:00.000"} という文字列で返します。<br>
     * 
     * @param cell 対象のセル
     * @param returnCachedValue 対象のセルの形式が数式の場合に、数式ではなくキャッシュされた算出値を返す場合は {@code true}
     * @return セルの格納値を表す文字列
     * @throws NullPointerException {@code cell} が {@code null} の場合
     * @throws IllegalStateException {@code cell} が未初期化状態の場合
     * 
     * @since 0.2.0
     */
    public static String getValue(Cell cell, boolean returnCachedValue) {
        Objects.requireNonNull(cell, "cell");
        
        CellType type = returnCachedValue && cell.getCellTypeEnum() == CellType.FORMULA
                ? cell.getCachedFormulaResultTypeEnum()
                : cell.getCellTypeEnum();
        
        switch (type) {
        case STRING:
            return cell.getStringCellValue();
        
        case FORMULA:
            return normalizeFormula(cell.getCellFormula());
        
        case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
        
        case NUMERIC:
            // 日付セルや独自書式セルの値の扱いは甚だ不完全なものの、
            // diffツールとしては内容の比較を行えればよいのだと割り切り、
            // これ以上に凝ったコーディングは行わないこととする。
            if (DateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                LocalDateTime localDateTime = LocalDateTime
                        .ofInstant(date.toInstant(), ZoneId.systemDefault());
                return dateTimeFormatter.format(localDateTime);
            } else {
                String val = String.valueOf(cell.getNumericCellValue());
                if (val.endsWith(".0")) {
                    val = val.substring(0, val.length() - 2);
                }
                return val;
            }
            
        case ERROR:
            return ErrorEval.getText(cell.getErrorCellValue());
        
        case BLANK:
            return "";
        
        case _NONE:
            throw new IllegalStateException("cell type is _NONE.");
            
        default:
            throw new AssertionError("unexpected cell type. cellTypeEnum: " + cell.getCellTypeEnum());
        }
    }
    
    /**
     * 数式文字列を標準化して返します。<br>
     * 例えばセルに「= 1 + 2」と入力されている場合、POIの各種APIによって
     * " 1 + 2" と取得されたり "1+2" と取得されたりするため、
     * 不要なスペースを取り除くことでこれらを標準化します。<br>
     * 
     * @param original 元の数式文字列
     * @return 標準化された数式文字列
     */
    public static String normalizeFormula(String original) {
        StringBuilder str = new StringBuilder(original.trim());
        boolean inString = false;
        int n = 0;
        
        while (n < str.length()) {
            if (str.charAt(n) == '"') {
                inString = !inString;
            } else if (str.charAt(n) == ' ' && !inString) {
                str.deleteCharAt(n);
                n--;
            }
            n++;
        }
        assert !inString;
        
        return str.toString();
    }
    
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
        
        if (rows.isEmpty()) {
            return;
        }
        
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
        
        if (columns.isEmpty()) {
            return;
        }
        
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
        
        if (addresses.isEmpty()) {
            return;
        }
        
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
    
    private ExcelUtils() {
    }
}
