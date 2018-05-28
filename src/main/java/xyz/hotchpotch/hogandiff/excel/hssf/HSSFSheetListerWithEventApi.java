package xyz.hotchpotch.hogandiff.excel.hssf;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import java.util.Queue;
import java.util.Set;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.excel.BookType;
import xyz.hotchpotch.hogandiff.excel.SheetLister;

/**
 * HSSF（.xls）形式のExcelブックからPOIのイベントモデルAPIを使用してシート名の一覧を取得するための
 * {@link SheetLister} の実装です。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
public class HSSFSheetListerWithEventApi implements SheetLister {
    
    // [static members] ********************************************************
    
    private static class HSSFSheetListingListener implements HSSFListener {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private Queue<String> queue = null;
        private List<String> sheetNames = null;
        
        private HSSFSheetListingListener() {
        }
        
        @Override
        public void processRecord(Record record) {
            switch (record.getSid()) {
            case BOFRecord.sid:
                BOFRecord bofRecord = (BOFRecord) record;
                
                switch (bofRecord.getType()) {
                case BOFRecord.TYPE_WORKBOOK:
                    queue = new LinkedList<>();
                    sheetNames = new ArrayList<>();
                    break;
                    
                case BOFRecord.TYPE_WORKSHEET:
                    sheetNames.add(queue.remove());
                    break;
                    
                case BOFRecord.TYPE_CHART:
                    // グラフシートは処理対象外とする。
                    queue.remove();
                    break;
                }
                
                break;
            
            case BoundSheetRecord.sid:
                BoundSheetRecord bSheetRecord = (BoundSheetRecord) record;
                queue.add(bSheetRecord.getSheetname());
                break;
            }
        }
    }
    
    /** このクラスがサポートするブック形式 */
    private static final Set<BookType> supported = EnumSet.of(BookType.XLS);
    
    /**
     * このクラスが指定されたファイルの形式をサポートするかを返します。<br>
     * 
     * @param file 検査対象のファイル
     * @return このクラスが指定されたファイルをサポートする場合は {@code true}
     * @throws NullPointerException {@code file} が {@code null} の場合
     */
    public static boolean isSupported(File file) {
        Objects.requireNonNull(file);
        return supported.stream()
                .map(BookType::extension)
                .anyMatch(file.getName()::endsWith);
    }
    
    /**
     * リスターオブジェクトを生成して返します。<br>
     * 
     * @return 新しいリスター
     */
    public static HSSFSheetListerWithEventApi of() {
        return new HSSFSheetListerWithEventApi();
    }
    
    // [instance members] ******************************************************
    
    private HSSFSheetListerWithEventApi() {
    }
    
    /**
     * {@inheritDoc}
     * 
     * @throws NullPointerException {@code book} が {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     */
    @Override
    public List<String> getSheetNames(File book) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        try (FileInputStream fin = new FileInputStream(book);
                POIFSFileSystem poifs = new POIFSFileSystem(fin)) {
            
            HSSFRequest req = new HSSFRequest();
            HSSFSheetListingListener listener = new HSSFSheetListingListener();
            req.addListenerForAllRecords(listener);
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.abortableProcessWorkbookEvents(req, poifs);
            return listener.sheetNames;
            
        } catch (Exception e) {
            throw new ApplicationException("Excelブックの読み込みに失敗しました。book:" + book.getPath(), e);
        }
    }
}
