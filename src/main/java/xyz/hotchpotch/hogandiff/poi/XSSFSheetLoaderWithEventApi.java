package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.io.InputStream;
import java.util.EnumSet;
import java.util.HashSet;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Set;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.common.BookType;
import xyz.hotchpotch.hogandiff.common.CellReplica;

/**
 * XSSF（.xlsx, .xlsm）形式のExcelブックからPOIのイベントモデルAPIを使用してシートデータを読み込むための
 * {@link SheetLoader} の実装です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class XSSFSheetLoaderWithEventApi implements SheetLoader {
    
    // [static members] ********************************************************
    
    /**
     * セルタイプを表す列挙型です。<br>
     * 
     * @author nmby
     * @since 0.1.0
     * @see <a href="http://officeopenxml.com/SScontentOverview.php">http://officeopenxml.com/SScontentOverview.php</a>
     */
    private static enum CellType {
        
        // [static members] ----------------------------------------------------
        
        /** boolean */
        b,
        
        /** date */
        d,
        
        /** error */
        e,
        
        /** inline string */
        inlineStr,
        
        /** number */
        n,
        
        /** shared string */
        s,
        
        /** formula */
        str;
        
        /**
         * sheet#.xml における c 要素の t 属性に対応する {@link CellType} を返します。
         * t 属性が省略されているとき（{@code null} が指定された場合）は {@link CellType#n} を返します。<br>
         * 
         * @param type sheet#.xml における c 要素の t 属性の値（省略されているときは {@code null}）
         * @return 対応する {@link CellType}
         */
        private static CellType of(String type) {
            return type == null ? n : valueOf(type);
        }
        
        // [instance members] --------------------------------------------------
        
    }
    
    private static class XSSFSheetLoadingHandler extends DefaultHandler {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private final boolean extractCachedValue;
        private final SharedStringsTable table;
        
        private Set<CellReplica> result;
        private String address;
        private CellType type;
        private String value;
        private boolean getNextString;
        
        /**
         * コンストラクタ。<br>
         * 
         * @param extractCachedValue 数式セルから数式文字列ではなくキャッシュされた計算値を取得する場合は {@code true}
         * @param table 読込対象のExcelブックの {@link SharedStringsTable} オブジェクト
         */
        private XSSFSheetLoadingHandler(boolean extractCachedValue, SharedStringsTable table) {
            super();
            
            assert table != null;
            
            this.extractCachedValue = extractCachedValue;
            this.table = table;
        }
        
        @Override
        public void startDocument() {
            result = new HashSet<>();
            
            address = null;
            type = null;
            value = null;
            getNextString = false;
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            assert uri != null;
            assert localName != null;
            assert qName != null;
            assert attributes != null;
            
            switch (qName) {
            case "c":
                address = attributes.getValue("r");
                assert address != null;
                type = CellType.of(attributes.getValue("t"));
                getNextString = false;
                break;
            
            case "f":
                getNextString = !extractCachedValue;
                break;
            
            case "v":
                getNextString = (value == null);
                break;
            
            default:
                getNextString = false;
                break;
            }
        }
        
        @Override
        public void characters(char ch[], int start, int length) {
            assert ch != null;
            assert start + length <= ch.length;
            
            if (getNextString) {
                String v = new String(ch, start, length);
                
                assert type != null;
                switch (type) {
                case s: // shared string
                    int idx = Integer.parseInt(v);
                    value = new XSSFRichTextString(table.getEntryAt(idx)).toString();
                    break;
                
                case b: // boolean
                    value = Boolean.valueOf("1".equals(v)).toString();
                    break;
                
                case d: // date
                case e: // error
                case inlineStr: // inline string
                case n: // number
                case str: // formula
                    value = (value == null ? v : value + v);
                    break;
                
                default:
                    throw new AssertionError(type);
                }
            }
        }
        
        @Override
        public void endElement(String uri, String localName, String qName) {
            assert uri != null;
            assert localName != null;
            assert qName != null;
            
            if ("c".equals(qName)) {
                if (value != null) {
                    if (type == CellType.e || type == CellType.n || type == CellType.str) {
                        value = POIUtils.normalizeFormula(value);
                    }
                    result.add(CellReplica.of(address, value));
                }
                
                address = null;
                type = null;
                value = null;
            }
            if (getNextString) {
                getNextString = false;
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
     * ローダーオブジェクトを生成して返します。<br>
     * 
     * @param extractCachedValue 数式セルから数式文字列ではなくキャッシュされた計算値を取得する場合は {@code true}
     * @return 新しいローダー
     */
    public static XSSFSheetLoaderWithEventApi of(boolean extractCachedValue) {
        return new XSSFSheetLoaderWithEventApi(extractCachedValue);
    }
    
    // [instance members] ******************************************************
    
    private final boolean extractCachedValue;
    
    private XSSFSheetLoaderWithEventApi(boolean extractCachedValue) {
        this.extractCachedValue = extractCachedValue;
    }
    
    /**
     * 指定されたExcelブックから relId で指定されたシートのデータを読み込みます。<br>
     * 
     * @param book 対象のExcelブック
     * @param relId 対象シートのrelId
     * @return シートに含まれるセルデータのセット
     * @throws ApplicationException 処理に失敗した場合
     * @throws NullPointerException {@code book}, {code relId} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     * @throws NoSuchElementException {@code relId} に該当するシートが存在しない場合
     */
    public Set<CellReplica> loadSheetById(File book, String relId) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(relId, "relId");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        try (OPCPackage pkg = OPCPackage.open(book, PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable table = reader.getSharedStringsTable();
            XSSFSheetLoadingHandler handler = new XSSFSheetLoadingHandler(extractCachedValue, table);
            XMLReader parser = XMLReaderFactory.createXMLReader();
            parser.setContentHandler(handler);
            
            try (InputStream sheetData = reader.getSheet(relId)) {
                InputSource sheetSource = new InputSource(sheetData);
                parser.parse(sheetSource);
                return handler.result;
            }
        } catch (IllegalArgumentException e) {
            throw new NoSuchElementException(
                    String.format("book: %s, relId: %s", book.getPath(), relId));
        } catch (Exception e) {
            String msg = String.format("シートの読み込みに失敗しました。book:%s, relId:%s",
                    book.getPath(), relId);
            throw new ApplicationException(msg, e);
        }
    }
    
    /**
     * {@inheritDoc}
     * 
     * @throws NullPointerException {@code book}, {code sheetName} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     * @throws NoSuchElementException {@code sheetName} に該当するシートが存在しない場合
     */
    @Override
    public Set<CellReplica> loadSheet(File book, String sheetName) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(sheetName, "sheetName");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        XSSFSheetListerWithEventApi lister = XSSFSheetListerWithEventApi.of();
        Map<String, String> sheetsMap = lister.getSheetsId(book);
        String relId = sheetsMap.get(sheetName);
        if (relId == null) {
            throw new NoSuchElementException(
                    String.format("book: %s, relId: %s", book.getPath(), relId));
        }
        
        return loadSheetById(book, relId);
    }
}
