package xyz.hotchpotch.hogandiff.excel.xssf.readers;

import java.util.ArrayDeque;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Optional;
import java.util.Queue;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.diff.excel.SResult.Piece;
import xyz.hotchpotch.hogandiff.excel.CellReplica;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFUtils.NONS_QNAME;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFUtils.QNAME;

/**
 * XSSF（.xlsx/.xlsm）形式のExcelファイルに含まれる xl/worksheets/sheet?.xml エントリを変換して
 * 差分箇所に色を付けるための {@link XMLEventReader} の実装です。<br>
 * 
 * @author nmby
 * @since 0.4.0
 */
public class SheetReader extends AbstractCachingReader {
    
    // [static members] ********************************************************
    
    private static interface Processor {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        void process() throws XMLStreamException;
    }
    
    /**
     * XSSF（.xlsx/.xlsm）形式のExcelファイルに含まれる xl/styles.xml エントリのラッパーです。<br>
     * 
     * @author nmby
     * @since 0.4.0
     */
    public static class StylesManager {
        
        // [static members] ----------------------------------------------------
        
        /**
         * {@link StylesManager} オブジェクトを生成して返します。<br>
         * 
         * @param styles xl/styles.xml エントリから生成した {@link Document}
         * @return 新しい {@link StylesManager} オブジェクト
         * @throws NullPointerException {@code styles} が {@code null} の場合
         */
        public static StylesManager of(Document styles) {
            Objects.requireNonNull(styles, "styles");
            return new StylesManager(styles);
        }
        
        // [instance members] --------------------------------------------------
        
        private final Document styles;
        private final Element elemCellXfs;
        private final Element elemFills;
        private final Map<Pair<Integer>, Integer> xfsMap = new HashMap<>();
        private final Map<Short, Integer> fillsMap = new HashMap<>();
        
        private int cellXfsCount;
        private int fillsCount;
        
        private StylesManager(Document styles) {
            assert styles != null;
            this.styles = styles;
            
            elemCellXfs = (Element) styles.getElementsByTagName("cellXfs").item(0);
            elemFills = (Element) styles.getElementsByTagName("fills").item(0);
            cellXfsCount = Integer.parseInt(elemCellXfs.getAttribute("count"));
            fillsCount = Integer.parseInt(elemFills.getAttribute("count"));
        }
        
        private int createFill(short newColor) {
            fillsMap.put(newColor, fillsCount);
            fillsCount++;
            elemFills.setAttribute("count", Integer.toString(fillsCount));
            
            Element newFill = styles.createElement("fill");
            elemFills.appendChild(newFill);
            
            Element patternFill = styles.createElement("patternFill");
            patternFill.setAttribute("patternType", "solid");
            newFill.appendChild(patternFill);
            
            Element fgColor = styles.createElement("fgColor");
            fgColor.setAttribute("indexed", Integer.toString(newColor));
            patternFill.appendChild(fgColor);
            
            return fillsCount - 1;
        }
        
        private int copyXf(int idx, short newColor) {
            xfsMap.put(Pair.of(idx, (int) newColor), cellXfsCount);
            cellXfsCount++;
            elemCellXfs.setAttribute("count", Integer.toString(cellXfsCount));
            
            Element orgXf = (Element) elemCellXfs.getElementsByTagName("xf").item(idx);
            Element newXf = (Element) orgXf.cloneNode(true);
            elemCellXfs.appendChild(newXf);
            
            int newFillId = Optional.ofNullable(fillsMap.get(newColor))
                    .orElseGet(() -> createFill(newColor));
            newXf.setAttribute("fillId", Integer.toString(newFillId));
            newXf.setAttribute("applyFill", "1");
            
            return cellXfsCount - 1;
        }
        
        /**
         * 指定されたスタイルに指定された色を適用したスタイルを返します。<br>
         * 該当するスタイルが既に存在すればそれを、存在しなければ新たに作成して返します。<br>
         * 
         * @param idx 元のスタイルのインデックス
         * @param newColor 新たなスタイルの色
         * @return 新たなスタイルのインデックス
         */
        public synchronized int getNewStyle(int idx, short newColor) {
            Pair<Integer> key = Pair.of(idx, (int) newColor);
            if (xfsMap.containsKey(key)) {
                return xfsMap.get(key);
            } else {
                return copyXf(idx, newColor);
            }
        }
    }
    
    private static final XMLEventFactory eventFactory = XMLEventFactory.newFactory();
    
    private static void copyAttributes(
            Iterator<Attribute> original,
            Set<Attribute> dstAttrs) {
        
        while (original.hasNext()) {
            Attribute attr = original.next();
            if (dstAttrs.stream().map(Attribute::getName).noneMatch(attr.getName()::equals)) {
                dstAttrs.add(attr);
            }
        }
    }
    
    /**
     * {@link SheetReader} オブジェクトを生成して返します。<br>
     * 
     * @param source ソースリーダー
     * @param stylesManager 対象Excelブックの xl/styles.xml エントリから構成した {@link StylesManager} オブジェクト
     * @param piece 対象シートの差分箇所
     * @param context コンテキスト
     * @return 新しい {@link SheetReader} オブジェクト
     * @throws NullPointerException {@code source}, {@code stylesManager}, {@code piece}, {@code context}
     *                              のいずれかが {@code null} の場合
     */
    public static SheetReader of(
            XMLEventReader source,
            StylesManager stylesManager,
            Piece piece,
            Context context) {
        
        Objects.requireNonNull(source, "source");
        Objects.requireNonNull(stylesManager, "stylesManager");
        Objects.requireNonNull(piece, "piece");
        Objects.requireNonNull(context, "context");
        
        return new SheetReader(source, stylesManager, piece, context);
    }
    
    // [instance members] ******************************************************
    
    private final XMLEventReader source;
    private final Piece piece;
    private final Queue<XMLEvent> nexts = new ArrayDeque<>();
    private final StylesManager stylesManager;
    private final short redundantColor;
    private final short diffColor;
    
    private Processor processor;
    
    private SheetReader(
            XMLEventReader source,
            StylesManager stylesManager,
            Piece piece,
            Context context) {
        
        super();
        
        assert source != null;
        assert stylesManager != null;
        assert piece != null;
        assert context != null;
        
        this.source = source;
        this.stylesManager = stylesManager;
        this.piece = piece;
        this.redundantColor = context.get(Props.APP_REDUNDANT_COLOR);
        this.diffColor = context.get(Props.APP_DIFF_COLOR);
        
        processor = piece.redundantColumns.isEmpty()
                ? new WaitingRow()
                : new WaitingCol();
    }
    
    @Override
    protected boolean hasNext2() {
        if (nexts.isEmpty()) {
            try {
                processor.process();
            } catch (XMLStreamException e) {
                throw new RuntimeException(e);
            }
        }
        return !nexts.isEmpty();
    }
    
    @Override
    protected XMLEvent peek2() throws XMLStreamException {
        return nexts.peek();
    }
    
    @Override
    protected XMLEvent nextEvent2() throws XMLStreamException {
        return nexts.poll();
    }
    
    @Override
    public void close() throws XMLStreamException {
        source.close();
    }
    
    private class WaitingCol implements Processor {
        
        @Override
        public void process() throws XMLStreamException {
            XMLEvent event = source.peek();
            if (!event.isStartElement()) {
                nexts.add(source.nextEvent());
                return;
            }
            StartElement start = event.asStartElement();
            if (QNAME.COLS.equals(start.getName())) {
                nexts.add(source.nextEvent());
                processor = new ProcessingCol();
                return;
            }
            if (QNAME.SHEET_DATA.equals(start.getName())) {
                nexts.add(eventFactory.createStartElement(QNAME.COLS, Collections.emptyIterator(), null));
                processor = new ProcessingCol();
                return;
            }
            nexts.add(source.nextEvent());
        }
    }
    
    private class ProcessingCol implements Processor {
        
        private final Queue<Pair<Integer>> redundantRanges = new ArrayDeque<>();
        private Pair<Integer> range1;
        private Pair<Integer> range2;
        private StartElement colStart;
        private Queue<XMLEvent> colEvents;
        private String colNewStyle;
        
        private ProcessingCol() {
            int start = -1;
            int end = -1;
            for (int i : piece.redundantColumns) {
                if (start == -1) {
                    start = i;
                    end = i;
                } else if (end + 1 < i) {
                    redundantRanges.add(Pair.of(start, end));
                    start = i;
                    end = i;
                } else if (end + 1 == i) {
                    end = i;
                } else {
                    throw new AssertionError();
                }
            }
            redundantRanges.add(Pair.of(start, end));
        }
        
        @Override
        public void process() throws XMLStreamException {
            if (range1 == null) {
                arrangeRange1();
            }
            if (range2 == null) {
                range2 = redundantRanges.poll();
            }
            
            if (range1 != null && range2 != null) {
                if (range1.a() < range2.a()) {
                    copyCol(range1.a(), Math.min(range1.b(), range2.a() - 1), false);
                    if (range1.b() < range2.a()) {
                        range1 = null;
                    } else {
                        reArrangeRange1(range2.a(), range1.b());
                    }
                    return;
                    
                } else if (range2.a() < range1.a()) {
                    createCol(range2.a(), Math.min(range2.b(), range1.a() - 1));
                    range2 = range2.b() < range1.a()
                            ? redundantRanges.poll()
                            : Pair.of(range1.a(), range2.b());
                    return;
                    
                } else {
                    copyCol(range1.a(), Math.min(range1.b(), range2.b()), true);
                    Pair<Integer> prev1 = range1;
                    Pair<Integer> prev2 = range2;
                    
                    if (prev1.b() < prev2.b()) {
                        range1 = null;
                    } else {
                        reArrangeRange1(prev2.b() + 1, prev1.b());
                    }
                    range2 = prev2.b() < prev1.b()
                            ? redundantRanges.poll()
                            : Pair.of(prev1.b() + 1, prev2.b());
                    return;
                }
                
            } else if (range1 != null) {
                while (range1 != null) {
                    nexts.add(colStart);
                    nexts.addAll(colEvents);
                    range1 = null;
                    arrangeRange1();
                }
                return;
                
            } else if (range2 != null) {
                while (range2 != null) {
                    createCol(range2.a(), range2.b());
                    range2 = redundantRanges.poll();
                }
                return;
                
            } else {
                XMLEvent event = source.peek();
                if (event.isStartElement()
                        && QNAME.SHEET_DATA.equals(event.asStartElement().getName())) {
                    nexts.add(eventFactory.createEndElement(QNAME.COLS, null));
                    
                } else if (event.isEndElement()
                        && QNAME.COLS.equals(event.asEndElement().getName())) {
                    nexts.add(source.nextEvent());
                    
                } else {
                    throw new AssertionError();
                }
                
                processor = new WaitingRow();
            }
        }
        
        private void arrangeRange1() throws XMLStreamException {
            if (findNextCol()) {
                colStart = source.nextEvent().asStartElement();
                
                range1 = Pair.of(
                        Integer.parseInt(colStart.getAttributeByName(NONS_QNAME.MIN).getValue()) - 1,
                        Integer.parseInt(colStart.getAttributeByName(NONS_QNAME.MAX).getValue()) - 1);
                
                colEvents = new ArrayDeque<>();
                while (!source.peek().isEndElement()
                        || !QNAME.COL.equals(source.peek().asEndElement().getName())) {
                    colEvents.add(source.nextEvent());
                }
                colEvents.add(source.nextEvent());
                
                int currStyle = Optional.ofNullable(colStart.getAttributeByName(NONS_QNAME.STYLE))
                        .map(Attribute::getValue)
                        .map(Integer::parseInt)
                        .orElse(0);
                colNewStyle = Integer.toString(stylesManager.getNewStyle(currStyle, redundantColor));
            }
        }
        
        private void reArrangeRange1(int min, int max) {
            range1 = Pair.of(min, max);
            
            Set<Attribute> attrs = new HashSet<>();
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MIN, Integer.toString(min + 1)));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MAX, Integer.toString(max + 1)));
            @SuppressWarnings("unchecked")
            Iterator<Attribute> itr = colStart.getAttributes();
            copyAttributes(itr, attrs);
            
            colStart = eventFactory.createStartElement(QNAME.COL, attrs.iterator(), null);
        }
        
        private boolean findNextCol() throws XMLStreamException {
            while (source.hasNext()) {
                XMLEvent event = source.peek();
                if (event.isStartElement()) {
                    QName name = event.asStartElement().getName();
                    if (QNAME.COL.equals(name)) {
                        return true;
                    } else if (QNAME.SHEET_DATA.equals(name)) {
                        return false;
                    }
                } else if (event.isEndElement()) {
                    QName name = event.asEndElement().getName();
                    if (QNAME.COLS.equals(name)) {
                        return false;
                    }
                }
                source.nextEvent();
            }
            throw new XMLStreamException();
        }
        
        private void copyCol(int start, int end, boolean redundant) {
            Set<Attribute> attrs = new HashSet<>();
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MIN, Integer.toString(start + 1)));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MAX, Integer.toString(end + 1)));
            if (redundant) {
                attrs.add(eventFactory.createAttribute(NONS_QNAME.STYLE, colNewStyle));
            }
            
            @SuppressWarnings("unchecked")
            Iterator<Attribute> itr = colStart.getAttributes();
            SheetReader.copyAttributes(itr, attrs);
            
            nexts.add(eventFactory.createStartElement(QNAME.COL, attrs.iterator(), null));
            nexts.addAll(colEvents);
        }
        
        private void createCol(int start, int end) {
            int newStyle = stylesManager.getNewStyle(0, redundantColor);
            
            Set<Attribute> attrs = new HashSet<>();
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MIN, Integer.toString(start + 1)));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.MAX, Integer.toString(end + 1)));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.STYLE, Integer.toString(newStyle)));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.WIDTH, "8.6640625"));
            
            nexts.add(eventFactory.createStartElement(QNAME.COL, attrs.iterator(), null));
            nexts.add(eventFactory.createEndElement(QNAME.COL, null));
        }
    }
    
    private class WaitingRow implements Processor {
        private WaitingRow() {
        }
        
        @Override
        public void process() throws XMLStreamException {
            while (source.hasNext()) {
                XMLEvent event = source.nextEvent();
                nexts.add(event);
                if (event.isStartElement() && QNAME.SHEET_DATA.equals(event.asStartElement().getName())) {
                    break;
                }
            }
            processor = new ProcessingRow();
        }
    }
    
    private class ProcessingRow implements Processor {
        
        private final Queue<Entry<Integer, Queue<Pair<Integer>>>> targets;
        private final Set<Integer> redundantRows;
        private final Set<Integer> redundantColumns;
        
        private ProcessingRow() {
            List<Integer> rows = Stream.concat(
                    piece.redundantRows.stream(),
                    piece.diffCells.stream().map(CellReplica::row))
                    .sorted()
                    .distinct()
                    .collect(Collectors.toList());
            
            List<Pair<Integer>> cells = Stream.concat(
                    piece.diffCells.stream().map(c -> Pair.of(c.row(), c.column())),
                    rows.stream().flatMap(r -> piece.redundantColumns.stream().map(c -> Pair.of(r, c))))
                    .sorted((c1, c2) -> !c1.a().equals(c2.a())
                            ? Integer.compare(c1.a(), c2.a())
                            : Integer.compare(c1.b(), c2.b()))
                    .distinct()
                    .collect(Collectors.toList());
            
            Map<Integer, Queue<Pair<Integer>>> tmp = cells.stream()
                    .collect(Collectors.groupingBy(
                            Pair::a,
                            Collectors.toCollection(ArrayDeque::new)));
            
            piece.redundantRows.forEach(r -> tmp.putIfAbsent(r, new ArrayDeque<>()));
            
            targets = tmp.entrySet().stream()
                    .sorted(Comparator.comparing(Entry::getKey))
                    .collect(Collectors.toCollection(ArrayDeque::new));
            
            redundantRows = new HashSet<>(piece.redundantRows);
            redundantColumns = new HashSet<>(piece.redundantColumns);
        }
        
        @Override
        public void process() throws XMLStreamException {
            Optional<Integer> eventRowIdx = Optional.ofNullable(source.peek())
                    .filter(XMLEvent::isStartElement)
                    .map(XMLEvent::asStartElement)
                    .filter(start -> QNAME.ROW.equals(start.getName()))
                    .map(start -> start.getAttributeByName(NONS_QNAME.R))
                    .map(Attribute::getValue)
                    .map(str -> Integer.parseInt(str) - 1);
            
            Optional<Integer> diffRowIdx = Optional.ofNullable(targets.peek())
                    .map(Entry::getKey);
            
            if (eventRowIdx.isPresent() && diffRowIdx.isPresent()) {
                int c = Integer.compare(eventRowIdx.get(), diffRowIdx.get());
                if (c < 0) {
                    addModifiedRowEvents(redundantColumns.stream()
                            .map(i -> Pair.of(eventRowIdx.get(), i))
                            .collect(Collectors.toCollection(ArrayDeque::new)));
                    
                } else if (0 < c) {
                    addNewRowEvents(targets.poll());
                    
                } else {
                    addModifiedRowEvents(targets.poll().getValue());
                }
                
            } else if (eventRowIdx.isPresent()) {
                addModifiedRowEvents(redundantColumns.stream()
                        .map(i -> Pair.of(eventRowIdx.get(), i))
                        .collect(Collectors.toCollection(ArrayDeque::new)));
                
            } else if (diffRowIdx.isPresent()) {
                addNewRowEvents(targets.poll());
                
            } else {
                nexts.add(source.nextEvent());
                processor = new ProcessingRemaining();
            }
        }
        
        private void addNewRowEvents(Entry<Integer, Queue<Pair<Integer>>> target) {
            nexts.add(createRowStart(target.getKey()));
            
            target.getValue().forEach(idx -> {
                nexts.add(createCStart(idx));
                nexts.add(eventFactory.createEndElement(QNAME.C, null));
            });
            
            nexts.add(eventFactory.createEndElement(QNAME.ROW, null));
        }
        
        private void addModifiedRowEvents(Queue<Pair<Integer>> cells) throws XMLStreamException {
            StartElement rowStart = source.nextEvent().asStartElement();
            nexts.add(modifyRowStart(rowStart));
            
            while (source.hasNext()) {
                XMLEvent event = source.nextEvent();
                
                if (event.isEndElement() && QNAME.ROW.equals(event.asEndElement().getName())) {
                    while (!cells.isEmpty()) {
                        nexts.add(createCStart(cells.poll()));
                        nexts.add(eventFactory.createEndElement(QNAME.C, null));
                    }
                    
                    nexts.add(event);
                    return;
                    
                } else if (event.isStartElement() && QNAME.C.equals(event.asStartElement().getName())) {
                    StartElement cStart = event.asStartElement();
                    Pair<Integer> eventIdx = CellReplica.getIndex(
                            cStart.getAttributeByName(NONS_QNAME.R).getValue());
                    
                    while (!cells.isEmpty() && cells.peek().b() < eventIdx.b()) {
                        nexts.add(createCStart(cells.poll()));
                        nexts.add(eventFactory.createEndElement(QNAME.C, null));
                    }
                    
                    if (!cells.isEmpty() && cells.peek().b().equals(eventIdx.b())) {
                        cells.poll();
                        nexts.add(modifyCStart(cStart, true));
                    } else {
                        nexts.add(modifyCStart(cStart, false));
                    }
                    
                } else {
                    nexts.add(event);
                }
            }
            throw new XMLStreamException();
        }
        
        private StartElement createRowStart(int rowIdx) {
            Set<Attribute> attrs = new HashSet<>();
            attrs.add(eventFactory.createAttribute(NONS_QNAME.R, Integer.toString(rowIdx + 1)));
            
            if (redundantRows.contains(rowIdx)) {
                attrs.add(eventFactory.createAttribute(NONS_QNAME.CUSTOM_FORMAT, "1"));
                attrs.add(eventFactory.createAttribute(NONS_QNAME.S,
                        Integer.toString(stylesManager.getNewStyle(0, redundantColor))));
            }
            
            return eventFactory.createStartElement(QNAME.ROW, attrs.iterator(), null);
        }
        
        private StartElement createCStart(Pair<Integer> idx) {
            short color = redundantRows.contains(idx.a()) || redundantColumns.contains(idx.b())
                    ? redundantColor
                    : diffColor;
            
            Set<Attribute> attrs = new HashSet<>();
            attrs.add(eventFactory.createAttribute(NONS_QNAME.R,
                    CellReplica.getAddress(idx.a(), idx.b())));
            attrs.add(eventFactory.createAttribute(NONS_QNAME.S,
                    Integer.toString(stylesManager.getNewStyle(0, color))));
            
            return eventFactory.createStartElement(QNAME.C, attrs.iterator(), null);
        }
        
        private StartElement modifyRowStart(StartElement original) {
            int rowIdx = Integer.parseInt(original.getAttributeByName(NONS_QNAME.R).getValue()) - 1;
            
            if (redundantRows.contains(rowIdx)) {
                int currStyle = Optional.ofNullable(original.getAttributeByName(NONS_QNAME.S))
                        .map(Attribute::getValue)
                        .map(Integer::parseInt)
                        .orElse(0);
                int newStyle = stylesManager.getNewStyle(currStyle, redundantColor);
                
                Set<Attribute> attrs = new HashSet<>();
                attrs.add(eventFactory.createAttribute(NONS_QNAME.CUSTOM_FORMAT, "1"));
                attrs.add(eventFactory.createAttribute(NONS_QNAME.S, Integer.toString(newStyle)));
                @SuppressWarnings("unchecked")
                Iterator<Attribute> currAttrs = original.getAttributes();
                copyAttributes(currAttrs, attrs);
                return eventFactory.createStartElement(QNAME.ROW, attrs.iterator(), null);
                
            } else {
                return original;
            }
        }
        
        private StartElement modifyCStart(StartElement original, boolean maybeDiff) {
            String address = original.getAttributeByName(NONS_QNAME.R).getValue();
            Pair<Integer> idx = CellReplica.getIndex(address);
            
            boolean redundant = redundantRows.contains(idx.a()) || redundantColumns.contains(idx.b());
            if (maybeDiff || redundant) {
                int currStyle = Optional.ofNullable(original.getAttributeByName(NONS_QNAME.S))
                        .map(Attribute::getValue)
                        .map(Integer::parseInt)
                        .orElse(0);
                int newStyle = stylesManager.getNewStyle(currStyle, redundant ? redundantColor : diffColor);
                
                Set<Attribute> attrs = new HashSet<>();
                attrs.add(eventFactory.createAttribute(NONS_QNAME.S, Integer.toString(newStyle)));
                @SuppressWarnings("unchecked")
                Iterator<Attribute> currAttrs = original.getAttributes();
                copyAttributes(currAttrs, attrs);
                return eventFactory.createStartElement(QNAME.C, attrs.iterator(), null);
                
            } else {
                return original;
            }
        }
    }
    
    private class ProcessingRemaining implements Processor {
        
        @Override
        public void process() throws XMLStreamException {
            if (source.hasNext()) {
                nexts.add(source.nextEvent());
            }
        }
    }
}
