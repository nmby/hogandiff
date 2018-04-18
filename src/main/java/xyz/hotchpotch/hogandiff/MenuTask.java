package xyz.hotchpotch.hogandiff;

import java.awt.Desktop;
import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

import javafx.concurrent.Task;
import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.common.CellReplica;
import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.excel.BResult;
import xyz.hotchpotch.hogandiff.excel.SComparator;
import xyz.hotchpotch.hogandiff.excel.SResult;
import xyz.hotchpotch.hogandiff.poi.POIUtils;

/**
 * {@link Menu} を実行するためのタスクです。<br>
 * このクラスのインスタンスは、いわゆる "ワンショット" です。ひとつのインスタンスでタスクを複数回実行することはできません。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class MenuTask extends Task<Path> {
    
    // [static members] ********************************************************
    
    private static final int PROGRESS_MAX = 100;
    private static final String BR = System.lineSeparator();
    
    /**
     * 新しいタスクを返します。<br>
     * 
     * @param context コンテキスト
     * @return 新しいタスク
     * @throws NullPointerException {@code context} が {@code null} の場合
     */
    public static MenuTask of(Context context) {
        Objects.requireNonNull(context, "context");
        return new MenuTask(context);
    }
    
    // [instance members] ******************************************************
    
    private final Context context;
    private final Menu menu;
    
    private StringBuilder str;
    
    private MenuTask(Context context) {
        assert context != null;
        this.context = context;
        this.menu = context.get(Props.CURR_MENU);
    }
    
    @Override
    protected Path call() throws Exception {
        str = new StringBuilder();
        try {
            // 0.入力チェック
            menu.validateTargets(context);
            
            // 1.作業用フォルダの作成
            Path workDir = createWorkDirectory(0, 2);
            
            // 2.比較するシートの組み合わせの決定
            List<Pair<String>> pairs = this.pairingSheets(2, 5);
            
            // 3.シート同士の比較
            BResult bResult = this.compareSheets(pairs, 5, 70);
            
            // 4. 比較結果の表示（テキスト）
            if (context.get(Props.APP_SHOW_RESULT_TEXT)) {
                showResultText(workDir, bResult, 70, 75);
            }
            
            // 5. 比較結果の表示（Excel）
            if (context.get(Props.APP_SHOW_PAINTED_SHEETS)) {
                showResultBooks(workDir, bResult, 75, 98);
            }
            
            str.append("処理が完了しました。").append(BR);
            updateMessage(str.toString());
            updateProgress(PROGRESS_MAX, PROGRESS_MAX);
            return workDir;
            
        } catch (ApplicationException e) {
            e.printStackTrace();
            throw e;
        } catch (Exception e) {
            e.printStackTrace();
            String msg = "予期せぬエラーが発生しました。\n" + e.getMessage();
            str.append(msg).append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException(msg, e);
        }
    }
    
    /**
     * 1. 作業用フォルダを作成します。<br>
     * 
     * @param progressBefore 処理前進捗率
     * @param progressAfter 処理後進捗率
     * @return 作成した作業用フォルダのパス
     * @throws ApplicationException 処理に失敗した場合
     */
    private Path createWorkDirectory(int progressBefore, int progressAfter)
            throws ApplicationException {
        
        Path path = null;
        try {
            updateProgress(progressBefore, PROGRESS_MAX);
            path = context.get(Props.CURR_WORK_DIR);
            str.append(String.format("作業用フォルダを作成しています...\n  - %s\n\n", path));
            updateMessage(str.toString());
            
            Path workDir = Files.createDirectories(path);
            updateProgress(progressAfter, PROGRESS_MAX);
            return workDir;
            
        } catch (Exception e) {
            str.append("作業用フォルダの作成に失敗しました。").append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException("作業用フォルダの作成に失敗しました。\n" + path);
        }
    }
    
    /**
     * 2. 比較するシートの組み合わせを決定します。<br>
     * 
     * @param progressBefore 処理前進捗率
     * @param progressAfter 処理後進捗率
     * @return 比較するシートの組み合わせ
     * @throws ApplicationException 処理に失敗した場合
     */
    private List<Pair<String>> pairingSheets(int progressBefore, int progressAfter)
            throws ApplicationException {
        
        try {
            updateProgress(progressBefore, PROGRESS_MAX);
            if (menu == Menu.COMPARE_BOOKS) {
                str.append("比較するシートの組み合わせを決定しています...").append(BR);
                updateMessage(str.toString());
            }
            
            List<Pair<String>> pairs = menu.getPairsOfSheetNames(context);
            if (menu == Menu.COMPARE_BOOKS) {
                pairs.forEach(p -> str.append(String.format("  - [%s]  vs  [%s]\n",
                        p.a().orElse("（なし）"),
                        p.b().orElse("（なし）"))));
                str.append(BR);
                updateMessage(str.toString());
            }
            updateProgress(progressAfter, PROGRESS_MAX);
            return pairs;
            
        } catch (ApplicationException e) {
            str.append(e.getMessage()).append(BR).append(BR);
            updateMessage(str.toString());
            throw e;
        } catch (Exception e) {
            str.append("シートの組み合わせ決定に失敗しました。").append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException("シートの組み合わせ決定に失敗しました。", e);
        }
    }
    
    /**
     * 3. シート同士を比較します。<br>
     * 
     * @param pairs 比較するシートの組み合わせ
     * @param progressBefore 処理前進捗率
     * @param progressAfter 処理後進捗率
     * @return Excelブック同士の比較結果
     * @throws ApplicationException 処理に失敗した場合
     */
    private BResult compareSheets(List<Pair<String>> pairs, int progressBefore, int progressAfter)
            throws ApplicationException {
        
        try {
            updateProgress(progressBefore, PROGRESS_MAX);
            
            int total = progressAfter - progressBefore;
            boolean extractCachedValue = context.get(Props.APP_COMPARE_ON_VALUE);
            File file1 = context.get(Props.CURR_FILE1);
            File file2 = context.get(Props.CURR_FILE2);
            SComparator comparator = SComparator.of(context);
            Map<Pair<String>, SResult> sResults = new HashMap<>();
            List<Pair<String>> pairedPairs = pairs.stream()
                    .filter(Pair::isPaired)
                    .collect(Collectors.toList());
            
            int i = 0;
            for (Pair<String> pair : pairedPairs) {
                i++;
                String sheetName1 = pair.a().get();
                String sheetName2 = pair.b().get();
                
                str.append(String.format(
                        "シートを比較しています(%d/%d)...\n  - A : %s\n  - B : %s\n",
                        i, pairedPairs.size(), sheetName1, sheetName2));
                updateMessage(str.toString());
                
                Set<CellReplica> cells1 = POIUtils.loadSheet(file1, sheetName1, extractCachedValue);
                Set<CellReplica> cells2 = POIUtils.loadSheet(file2, sheetName2, extractCachedValue);
                SResult sResult = comparator.compare(cells1, cells2);
                sResults.put(pair, sResult);
                
                str.append(sResult.getSummary()).append(BR);
                updateMessage(str.toString());
                updateProgress(progressBefore + total * i / pairedPairs.size(), PROGRESS_MAX);
            }
            
            BResult bResult = BResult.of(file1, file2, pairs, sResults);
            updateProgress(progressAfter, PROGRESS_MAX);
            return bResult;
            
        } catch (ApplicationException e) {
            str.append(e.getMessage()).append(BR).append(BR);
            updateMessage(str.toString());
            throw e;
        } catch (Exception e) {
            str.append("シートの比較に失敗しました。").append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException("シートの比較に失敗しました。", e);
        }
    }
    
    /**
     * 4. 比較結果をテキストファイルとして保存して表示します。<br>
     * 
     * @param workDir 作業用フォルダ
     * @param bResult Excelブック同士の比較結果
     * @param progressBefore 処理前進捗率
     * @param progressAfter 処理後進捗率
     * @throws ApplicationException 処理に失敗した場合
     */
    private void showResultText(Path workDir, BResult bResult, int progressBefore, int progressAfter)
            throws ApplicationException {
        
        Path filePath = null;
        try {
            updateProgress(progressBefore, PROGRESS_MAX);
            
            filePath = workDir.resolve("result.txt");
            
            str.append(String.format(
                    "比較結果のテキストファイルを保存して表示しています...\n  - %s\n\n", filePath.toString()));
            updateMessage(str.toString());
            
            try (BufferedWriter writer = Files.newBufferedWriter(filePath, Charset.forName("UTF-8"))) {
                writer.write(bResult.toString());
            }
            Desktop.getDesktop().open(filePath.toFile());
            updateProgress(progressAfter, PROGRESS_MAX);
            
        } catch (Exception e) {
            str.append("比較結果テキストの保存と表示に失敗しました。").append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException("比較結果テキストの保存と表示に失敗しました。\n" + filePath);
        }
    }
    
    /**
     * 5. 比較結果をExcelブックとして保存して表示します。<br>
     * 
     * @param workDir 作業用フォルダ
     * @param bResult Excelブック同士の比較結果
     * @param progressBefore 処理前進捗率
     * @param progressAfter 処理後進捗率
     * @throws ApplicationException 処理に失敗した場合
     */
    private void showResultBooks(Path workDir, BResult bResult, int progressBefore, int progressAfter)
            throws ApplicationException {
        
        try {
            updateProgress(progressBefore, PROGRESS_MAX);
            int total = progressAfter - progressBefore;
            
            str.append("Excelブックに比較結果の色を付けて保存しています(1/2)...").append(BR);
            updateMessage(str.toString());
            File file1 = context.get(Props.CURR_FILE1);
            Path storedBook1 = copyLoadPaintAndStoreBook(workDir, file1, bResult, Pair.Side.A);
            str.append(String.format("  - %s\n\n", storedBook1.toString()));
            updateMessage(str.toString());
            updateProgress(progressBefore + total * 2 / 5, PROGRESS_MAX);
            
            str.append("Excelブックに比較結果の色を付けて保存しています(2/2)...").append(BR);
            updateMessage(str.toString());
            File file2 = context.get(Props.CURR_FILE2);
            Path storedBook2 = copyLoadPaintAndStoreBook(workDir, file2, bResult, Pair.Side.B);
            str.append(String.format("  - %s\n\n", storedBook1.toString()));
            updateMessage(str.toString());
            updateProgress(progressBefore + total * 4 / 5, PROGRESS_MAX);
            
            str.append("比較結果のExcelブックを表示しています...").append(BR).append(BR);
            updateMessage(str.toString());
            Desktop.getDesktop().open(storedBook1.toFile());
            Desktop.getDesktop().open(storedBook2.toFile());
            updateProgress(progressAfter, PROGRESS_MAX);
            
        } catch (ApplicationException e) {
            str.append(e.getMessage()).append(BR).append(BR);
            updateMessage(str.toString());
            throw e;
        } catch (Exception e) {
            str.append("比較結果Excelブックの保存と表示に失敗しました。").append(BR).append(BR);
            updateMessage(str.toString());
            throw new ApplicationException("比較結果Excelブックの保存と表示に失敗しました。", e);
        }
    }
    
    private Path copyLoadPaintAndStoreBook(
            Path workDir,
            File file,
            BResult bResult,
            Pair.Side side)
            throws ApplicationException {
        
        assert workDir != null;
        assert file != null;
        assert bResult != null;
        assert side != null;
        
        Path copy = copyFile(file, workDir, String.format("【%s】", side));
        try (Workbook book = loadBook(copy)) {
            paintBook(book, bResult, side);
            storeBook(book, copy);
        } catch (IOException e) {
            // nop : failed to auto-close book.
        } finally {
            System.gc();
        }
        return copy;
    }
    
    private Path copyFile(File src, Path dir, String prefix) throws ApplicationException {
        assert src != null;
        assert dir != null;
        assert prefix != null;
        
        Path target = dir.resolve(prefix + src.getName());
        try {
            Files.copy(src.toPath(), target);
            return target;
            
        } catch (Exception e) {
            throw new ApplicationException("Excelファイルのコピーに失敗しました。\n" + target.toString(), e);
        }
    }
    
    /**
     * Excelブックをファイルからロードして返します。<br>
     * 
     * @param bookPath Excelブックのパス
     * @return Excelブック
     * @throws ApplicationException 処理に失敗した場合
     */
    private Workbook loadBook(Path bookPath) throws ApplicationException {
        assert bookPath != null;
        
        if (!Files.exists(bookPath)) {
            throw new ApplicationException("ファイルが存在しません。\n" + bookPath);
        }
        if (!Files.isReadable(bookPath)) {
            throw new ApplicationException("ファイルが読取不可です。\n" + bookPath);
        }
        
        try (InputStream is = Files.newInputStream(bookPath)) {
            return WorkbookFactory.create(is);
        } catch (EncryptedDocumentException e) {
            throw new ApplicationException("ファイルが暗号化されています。\n" + bookPath, e);
        } catch (InvalidFormatException e) {
            throw new ApplicationException("ファイルフォーマットが不正です。\n" + bookPath, e);
        } catch (Exception e) {
            throw new ApplicationException("Excelブックの読込みに失敗しました。\n" + bookPath, e);
        }
    }
    
    private void paintBook(Workbook book, BResult bResult, Pair.Side side) throws ApplicationException {
        assert book != null;
        assert bResult != null;
        assert side != null;
        
        try {
            POIUtils.clearColors(book);
            
            bResult.sheetNamePairs.stream()
                    .filter(p -> p.get(side).isPresent())
                    .forEach(p -> {
                        Sheet sheet = book.getSheet(p.get(side).get());
                        SResult sResult = bResult.sResults.get(p);
                        paintSheet(sheet, sResult, side);
                    });
            
        } catch (Exception e) {
            throw new ApplicationException("差分結果の着色に失敗しました。", e);
        }
    }
    
    private void paintSheet(Sheet sheet, SResult sResult, Pair.Side side) {
        assert sheet != null;
        assert sResult != null;
        assert side != null;
        
        short color1 = context.get(Props.APP_REDUNDANT_COLOR);
        short color2 = context.get(Props.APP_DIFF_COLOR);
        
        sResult.redundantRows.ifPresent(p -> {
            List<Integer> rows = p.get(side).get();
            POIUtils.paintRows(sheet, rows, color1);
        });
        
        sResult.redundantColumns.ifPresent(p -> {
            List<Integer> columns = p.get(side).get();
            POIUtils.paintColumns(sheet, columns, color1);
        });
        
        Set<CellAddress> addresses = sResult.diffCells.stream()
                .map(p -> p.get(side).get())
                .map(cell -> new CellAddress(cell.row(), cell.column()))
                .collect(Collectors.toSet());
        POIUtils.paintCells(sheet, addresses, color2);
    }
    
    private void storeBook(Workbook book, Path target) throws ApplicationException {
        assert book != null;
        assert target != null;
        
        try (OutputStream os = Files.newOutputStream(target)) {
            book.write(os);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ApplicationException("Excelブックの保存に失敗しました。\n" + target, e);
        }
    }
}
