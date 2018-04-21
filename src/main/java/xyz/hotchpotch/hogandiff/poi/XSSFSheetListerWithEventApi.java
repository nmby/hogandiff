package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

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
        
        private List<String> sheetNames;
        private Map<String, String> sheetsData;
        
        @Override
        public void startDocument() {
            sheetNames = new ArrayList<>();
            sheetsData = new HashMap<>();
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("sheet".equals(qName)) {
                sheetNames.add(attributes.getValue("name"));
                sheetsData.put(attributes.getValue("name"), attributes.getValue("r:id"));
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
        
        XSSFSheetListingHandler handler = parse(book);
        return handler.sheetNames;
    }
    
    /**
     * Excelブックに含まれるシート名とrelIdを取得し、キー：シート名、値：relId のマップとして返します。<br>
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
        
        XSSFSheetListingHandler handler = parse(book);
        return handler.sheetsData;
    }
    
    /**
     * Excelブックを読み込み、パースを行った {@link XSSFSheetListingHandler} オブジェクトを返します。<br>
     * 
     * @param book Excelブック
     * @return パースを行った {@link XSSFSheetListingHandler} オブジェクト
     * @throws ApplicationException 処理に失敗した場合
     */
    private XSSFSheetListingHandler parse(File book) throws ApplicationException {
        assert book != null;
        assert isSupported(book);
        
        try {
            XMLReader parser = XMLReaderFactory.createXMLReader();
            XSSFSheetListingHandler handler = new XSSFSheetListingHandler();
            parser.setContentHandler(handler);
            
            try (OPCPackage pkg = OPCPackage.open(book, PackageAccess.READ)) {
                XSSFReader reader = new XSSFReader(pkg);
                
                try (InputStream bookData = reader.getWorkbookData()) {
                    InputSource bookSource = new InputSource(bookData);
                    parser.parse(bookSource);
                    return handler;
                }
            }
        } catch (Exception e) {
            throw new ApplicationException("Excelブックの読み込みに失敗しました。book:" + book.getPath(), e);
        }
    }
}
