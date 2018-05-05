package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.io.InputStream;
import java.util.AbstractMap.SimpleEntry;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.common.BookType;

/**
 * XSSF（.xlsx, .xlsm）形式のExcelブックからPOIのイベントモデルAPIを使用してシート名の一覧を取得するための
 * {@link SheetLister} の実装です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class XSSFSheetListerWithEventApi implements SheetLister {
    
    // [static members] ********************************************************
    
    private static class XSSFSheetListingHandler extends DefaultHandler {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private List<Map.Entry<String, String>> sheetsInfo;
        
        @Override
        public void startDocument() {
            sheetsInfo = new ArrayList<>();
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("sheet".equals(qName)) {
                sheetsInfo.add(new SimpleEntry<>(
                        attributes.getValue("name"),
                        attributes.getValue("r:id")));
            }
        }
    }
    
    private static class XSSFSheetJudgeHandler extends DefaultHandler {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private Optional<Boolean> isChartSheet;
        private boolean isFirstElement;
        
        @Override
        public void startDocument() {
            isChartSheet = Optional.empty();
            isFirstElement = true;
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (isFirstElement) {
                if ("chartsheet".equals(qName)) {
                    isChartSheet = Optional.of(true);
                } else if ("worksheet".equals(qName)) {
                    isChartSheet = Optional.of(false);
                }
                isFirstElement = false;
            }
        }
    }
    
    /** このクラスがサポートするブック形式 */
    private static final Set<BookType> supported = EnumSet.of(BookType.XLSX, BookType.XLSM);
    
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
    public static XSSFSheetListerWithEventApi of() {
        return new XSSFSheetListerWithEventApi();
    }
    
    // [instance members] ******************************************************
    
    private XSSFSheetListerWithEventApi() {
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
        
        List<Map.Entry<String, String>> sheetsInfo = parse(book);
        return sheetsInfo.stream()
                .map(Map.Entry::getKey)
                .collect(Collectors.toList());
    }
    
    /**
     * Excelブックに含まれるワークシートのシート名とrelIdを取得し、
     * キー：シート名、値：relId のマップとして返します。<br>
     * 
     * @param book Excelブック
     * @return キー：シート名、値：relId のマップ
     * @throws NullPointerException {@code book} が {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     * @throws ApplicationException 処理に失敗した場合
     */
    public Map<String, String> getSheetsId(File book) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        List<Map.Entry<String, String>> sheetsInfo = parse(book);
        return sheetsInfo.stream()
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));
    }
    
    /**
     * Excelブックを読み込み、ワークシート情報（シート名＋relId）のリストを返します。
     * 返されるリストにグラフシートの情報は含まれません。<br>
     * 
     * @param book Excelブック
     * @return ワークシート情報（シート名＋relId）のリスト
     * @throws ApplicationException 処理に失敗した場合
     */
    private List<Map.Entry<String, String>> parse(File book) throws ApplicationException {
        assert book != null;
        assert isSupported(book);
        
        try (OPCPackage pkg = OPCPackage.open(book, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            
            // まず、グラフシートも含むシート情報（シート名＋relId）の一覧を取得する。
            List<Map.Entry<String, String>> sheetsInfo;
            try (InputStream bookData = reader.getWorkbookData()) {
                InputSource bookSource = new InputSource(bookData);
                XMLReader bookParser = XMLReaderFactory.createXMLReader();
                XSSFSheetListingHandler bookHandler = new XSSFSheetListingHandler();
                bookParser.setContentHandler(bookHandler);
                
                bookParser.parse(bookSource);
                sheetsInfo = bookHandler.sheetsInfo;
            }
            
            // 次に、得られた一覧のそれぞれについて、ワークシートかグラフシートかを調べ、
            // グラフシートの場合は一覧から除外する。
            XMLReader sheetParser = XMLReaderFactory.createXMLReader();
            XSSFSheetJudgeHandler sheetHandler = new XSSFSheetJudgeHandler();
            sheetParser.setContentHandler(sheetHandler);
            Iterator<Map.Entry<String, String>> itr = sheetsInfo.iterator();
            while (itr.hasNext()) {
                Map.Entry<String, String> key = itr.next();
                
                try (InputStream sheetData = reader.getSheet(key.getValue())) {
                    InputSource sheetSource = new InputSource(sheetData);
                    
                    sheetParser.parse(sheetSource);
                    if (sheetHandler.isChartSheet.orElseThrow(IllegalStateException::new)) {
                        itr.remove();
                    }
                }
            }
            
            // ワークシートだけが含まれるシート情報の一覧を返す。
            return sheetsInfo;
            
        } catch (Exception e) {
            throw new ApplicationException("Excelブックの読み込みに失敗しました。book:" + book.getPath(), e);
        }
    }
}
