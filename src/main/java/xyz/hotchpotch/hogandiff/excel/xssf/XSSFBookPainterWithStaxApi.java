package xyz.hotchpotch.hogandiff.excel.xssf;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.EnumSet;
import java.util.List;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.events.Attribute;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.diff.excel.SResult.Piece;
import xyz.hotchpotch.hogandiff.excel.BookPainter;
import xyz.hotchpotch.hogandiff.excel.BookType;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFUtils.NONS_QNAME;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFUtils.QNAME;
import xyz.hotchpotch.hogandiff.excel.xssf.readers.FilteringReader;
import xyz.hotchpotch.hogandiff.excel.xssf.readers.SheetReader;
import xyz.hotchpotch.hogandiff.excel.xssf.readers.SheetReader.StylesManager;

/**
 * XSSF（.xlsx, .xlsm）形式のExcelブックにStAX APIを使用して色を付けるための
 * {@link BookPainter} の実装です。<br>
 * 
 * @author nmby
 * @since 0.4.0
 */
public class XSSFBookPainterWithStaxApi implements BookPainter {
    
    // [static members] ********************************************************
    
    private static final XMLInputFactory inFactory = XMLInputFactory.newInstance();
    private static final XMLOutputFactory outFactory = XMLOutputFactory.newInstance();
    private static final DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
    private static final TransformerFactory transformerFactory = TransformerFactory.newInstance();
    
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
     * ペインターオブジェクトを生成して返します。<br>
     * 
     * @param context コンテキスト
     * @return 新しいペインター
     * @throws NullPointerException {@code context} が {@code null} の場合
     */
    public static XSSFBookPainterWithStaxApi of(Context context) {
        Objects.requireNonNull(context, "context");
        return new XSSFBookPainterWithStaxApi(context);
    }
    
    // [instance members] ******************************************************
    
    private final Context context;
    
    private SheetReader.StylesManager stylesManager;
    
    private XSSFBookPainterWithStaxApi(Context context) {
        assert context != null;
        this.context = context;
    }
    
    @Override
    public void paintAndSave(
            File book,
            Path copy,
            List<Entry<String, Piece>> results)
            throws ApplicationException {
        
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(copy, "copy");
        Objects.requireNonNull(results, "results");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getPath());
        }
        
        // 1. 処理対象のExcelファイルをコピーする。
        copyFile(book.toPath(), copy);
        
        // 2. 対象のExcelファイルをZipファイルとして扱い各種処理を行う。
        try (FileSystem inFs = FileSystems.newFileSystem(book.toPath(), null);
                FileSystem outFs = FileSystems.newFileSystem(copy, null)) {
            
            // 2-1. xl/sharedStrings.xml エントリに対する処理
            processSharedStringsEntry(inFs, outFs);
            
            // 2-2. xl/styles.xml エントリに対する処理
            processStylesEntry(inFs, outFs);
            
            // 2-3. xl/worksheets/sheet?.xml エントリに対する処理
            processWorksheetEntries(inFs, outFs, book.toPath(), results);
            
        } catch (ApplicationException e) {
            e.printStackTrace();
            throw e;
        } catch (Exception e) {
            e.printStackTrace();
            throw new ApplicationException("比較結果Excelブックの作成に失敗しました。\n" + copy.toString(), e);
        }
    }
    
    // 1. 処理対象のExcelファイルをコピーする。
    private void copyFile(Path src, Path dst) throws ApplicationException {
        try {
            Files.copy(src, dst);
            File f = dst.toFile();
            f.setReadable(true, false);
            f.setWritable(true, false);
        } catch (Exception e) {
            throw new ApplicationException("Excelファイルのコピーに失敗しました。\n" + dst.toString(), e);
        }
    }
    
    // 2-1. xl/sharedStrings.xml エントリに対する処理
    private void processSharedStringsEntry(FileSystem inFs, FileSystem outFs)
            throws ApplicationException {
        
        try (InputStream is = Files.newInputStream(inFs.getPath("xl/sharedStrings.xml"));
                OutputStream os = Files.newOutputStream(outFs.getPath("xl/sharedStrings.xml"),
                        StandardOpenOption.TRUNCATE_EXISTING)) {
            
            XMLEventReader reader = inFactory.createXMLEventReader(is, "UTF-8");
            XMLEventWriter writer = outFactory.createXMLEventWriter(os, "UTF-8");
            
            reader = FilteringReader.builder(reader)
                    .addFilter(QNAME.COLOR)
                    .build();
            
            writer.add(reader);
            
        } catch (Exception e) {
            throw new ApplicationException("xl/sharedStrings.xml エントリの処理に失敗しました。", e);
        }
    }
    
    // 2-2. xl/styles.xml エントリに対する処理
    private void processStylesEntry(FileSystem inFs, FileSystem outFs)
            throws ApplicationException {
        
        try (InputStream is = Files.newInputStream(inFs.getPath("xl/styles.xml"));
                OutputStream os = Files.newOutputStream(outFs.getPath("xl/styles.xml"),
                        StandardOpenOption.TRUNCATE_EXISTING)) {
            
            XMLEventReader reader = inFactory.createXMLEventReader(is, "UTF-8");
            XMLEventWriter writer = outFactory.createXMLEventWriter(os, "UTF-8");
            
            reader = FilteringReader.builder(reader)
                    .addFilter(QNAME.FONTS, QNAME.FONT, QNAME.COLOR)
                    .addFilter(QNAME.FILLS, QNAME.FILL, QNAME.PATTERN_FILL, QNAME.FG_COLOR)
                    .addFilter(QNAME.FILLS, QNAME.FILL, QNAME.PATTERN_FILL, QNAME.BG_COLOR)
                    .addFilter(QNAME.FILLS, QNAME.FILL, QNAME.GRADIENT_FILL)
                    .addFilter(QNAME.BORDERS, QNAME.BORDER, QNAME.TOP, QNAME.COLOR)
                    .addFilter(QNAME.BORDERS, QNAME.BORDER, QNAME.BOTTOM, QNAME.COLOR)
                    .addFilter(QNAME.BORDERS, QNAME.BORDER, QNAME.LEFT, QNAME.COLOR)
                    .addFilter(QNAME.BORDERS, QNAME.BORDER, QNAME.RIGHT, QNAME.COLOR)
                    .addFilter(QNAME.BORDERS, QNAME.BORDER, QNAME.DIAGONAL, QNAME.COLOR)
                    .addFilter(start -> QNAME.PATTERN_FILL.equals(start.getName())
                            && Optional.ofNullable(start.getAttributeByName(NONS_QNAME.PATTERN_TYPE))
                                    .map(Attribute::getValue)
                                    .map("solid"::equals)
                                    .orElse(false))
                    .build();
            
            writer.add(reader);
            
        } catch (Exception e) {
            throw new ApplicationException("xl/styles.xml エントリの処理に失敗しました。", e);
        }
    }
    
    // 2-3. xl/worksheets/sheet?.xml エントリに対する処理
    private void processWorksheetEntries(
            FileSystem inFs,
            FileSystem outFs,
            Path book,
            List<Entry<String, Piece>> results)
            throws ApplicationException {
        
        Document styles;
        try (// 注意：コピー先から読み込むため、outFsを使うので正しい。
                InputStream is = Files.newInputStream(outFs.getPath("xl/styles.xml"))) {
            DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
            styles = docBuilder.parse(is);
            //InputSource is2 = new InputSource(is);
            //is2.setEncoding("Shift_JIS");
            //styles = docBuilder.parse(is2);
            stylesManager = StylesManager.of(styles);
        } catch (Exception e) {
            throw new ApplicationException("xl/styles.xml エントリの読み込みに失敗しました。", e);
        }
        
        XSSFSheetEntryManager sheetManager = XSSFSheetEntryManager.generate(book);
        for (Entry<String, Piece> result : results) {
            String sheetName = result.getKey();
            String source = sheetManager.getSourceByName(sheetName);
            Piece piece = result.getValue();
            
            processWorksheetEntry(inFs, outFs, source, piece);
        }
        
        try (OutputStream os = Files.newOutputStream(outFs.getPath("xl/styles.xml"),
                StandardOpenOption.TRUNCATE_EXISTING)) {
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(styles);
            StreamResult result = new StreamResult(os);
            transformer.transform(source, result);
        } catch (Exception e) {
            throw new ApplicationException("xl/styles.xml エントリの保存に失敗しました。", e);
        }
    }
    
    private void processWorksheetEntry(
            FileSystem inFs,
            FileSystem outFs,
            String source,
            Piece piece)
            throws ApplicationException {
        
        try (InputStream is = Files.newInputStream(inFs.getPath(source));
                OutputStream os = Files.newOutputStream(outFs.getPath(source),
                        StandardOpenOption.TRUNCATE_EXISTING)) {
            
            XMLEventReader reader = inFactory.createXMLEventReader(is, "UTF-8");
            XMLEventWriter writer = outFactory.createXMLEventWriter(os, "UTF-8");
            
            reader = FilteringReader.builder(reader)
                    .addFilter(QNAME.CONDITIONAL_FORMATTING)
                    .build();
            
            reader = SheetReader.of(reader, stylesManager, piece, context);
            
            writer.add(reader);
            
        } catch (Exception e) {
            throw new ApplicationException(source + " エントリの処理に失敗しました。", e);
        }
    }
}
