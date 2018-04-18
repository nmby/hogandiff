package xyz.hotchpotch.hogandiff;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Properties;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.stream.Stream;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.collections.FXCollections;
import javafx.collections.MapChangeListener;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.RadioButton;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.stage.FileChooser;
import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.poi.POIUtils;

/**
 * このアプリケーションのメインコントローラです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class MainAppController {
    
    // [static members] ********************************************************
    
    // [instance members] ******************************************************
    
    @FXML
    private Pane settingsPane;
    
    @FXML
    private RadioButton compareBooks;
    
    @FXML
    private RadioButton compareSheets;
    
    @FXML
    private TextField bookPath1;
    
    @FXML
    private TextField bookPath2;
    
    @FXML
    private Button selectBook1;
    
    @FXML
    private Button selectBook2;
    
    @FXML
    private HBox sheetPane1;
    
    @FXML
    private HBox sheetPane2;
    
    @FXML
    private ChoiceBox<String> sheetName1;
    
    @FXML
    private ChoiceBox<String> sheetName2;
    
    @FXML
    private CheckBox considerRowGaps;
    
    @FXML
    private CheckBox considerColumnGaps;
    
    @FXML
    private RadioButton compareOnFormula;
    
    @FXML
    private RadioButton compareOnValue;
    
    @FXML
    private CheckBox showPaintedSheets;
    
    @FXML
    private CheckBox showResultText;
    
    @FXML
    private CheckBox exitWhenFinished;
    
    @FXML
    private Button saveSettings;
    
    @FXML
    private Button execute;
    
    @FXML
    private ProgressBar progress;
    
    @FXML
    private ScrollPane messageScroll;
    
    @FXML
    private Label messages;
    
    @FXML
    private Pane workDirPane;
    
    @FXML
    private Button showWorkDir;
    
    @FXML
    private Button deleteOldWorkDirs;
    
    private BooleanProperty isRunning = new SimpleBooleanProperty(false);
    private File prevSelected;
    private BooleanProperty settingsChanged = new SimpleBooleanProperty(false);
    private Properties appProps;
    
    @FXML
    private void initialize() {
        appProps = Main.loadProperties();
        Context context = Context.Builder.of(appProps).build();
        
        // バインディングの設定
        settingsPane.disableProperty().bind(isRunning);
        bookPath1.textProperty().bind(Bindings.createStringBinding(() -> {
            File file = (File) bookPath1.getUserData();
            return file == null ? "" : file.getPath();
        }, bookPath1.getProperties()));
        bookPath2.textProperty().bind(Bindings.createStringBinding(() -> {
            File file = (File) bookPath2.getUserData();
            return file == null ? "" : file.getPath();
        }, bookPath2.getProperties()));
        sheetPane1.disableProperty().bind(compareSheets.selectedProperty().not());
        sheetPane2.disableProperty().bind(compareSheets.selectedProperty().not());
        sheetName1.itemsProperty().bind(Bindings.createObjectBinding(
                () -> getSheetNames((File) bookPath1.getUserData()),
                bookPath1.getProperties()));
        sheetName2.itemsProperty().bind(Bindings.createObjectBinding(
                () -> getSheetNames((File) bookPath2.getUserData()),
                bookPath2.getProperties()));
        saveSettings.disableProperty().bind(settingsChanged.not());
        execute.disableProperty().bind(Bindings.createBooleanBinding(
                () -> bookPath1.getUserData() == null || bookPath2.getUserData() == null
                        || (compareSheets.isSelected()
                                && (sheetName1.getValue() == null || sheetName2.getValue() == null)),
                bookPath1.getProperties(), bookPath2.getProperties(),
                compareSheets.selectedProperty(),
                sheetName1.valueProperty(), sheetName2.valueProperty()));
        progress.disableProperty().bind(isRunning.not());
        messageScroll.vmaxProperty().bind(messages.heightProperty());
        messageScroll.vvalueProperty().bind(messages.heightProperty());
        workDirPane.disableProperty().bind(isRunning);
        
        // イベントハンドラの登録
        selectBook1.setOnAction(event -> bookPath1.setUserData(
                getSelectedFile((File) bookPath1.getUserData())));
        selectBook2.setOnAction(event -> bookPath2.setUserData(
                getSelectedFile((File) bookPath2.getUserData())));
        bookPath1.setOnDragOver(this::onDragOver);
        bookPath2.setOnDragOver(this::onDragOver);
        bookPath1.setOnDragDropped(event -> onDragDropped(event, bookPath1));
        bookPath2.setOnDragDropped(event -> onDragDropped(event, bookPath2));
        bookPath1.getProperties().addListener((MapChangeListener.Change<?, ?> change) -> {
            File file = (File) bookPath1.getUserData();
            prevSelected = (file == null ? prevSelected : file);
        });
        bookPath2.getProperties().addListener((MapChangeListener.Change<?, ?> change) -> {
            File file = (File) bookPath2.getUserData();
            prevSelected = (file == null ? prevSelected : file);
        });
        
        considerRowGaps.setOnAction(event -> settingsChanged.set(true));
        considerColumnGaps.setOnAction(event -> settingsChanged.set(true));
        compareOnFormula.setOnAction(event -> settingsChanged.set(true));
        compareOnValue.setOnAction(event -> settingsChanged.set(true));
        showPaintedSheets.setOnAction(event -> settingsChanged.set(true));
        showResultText.setOnAction(event -> settingsChanged.set(true));
        exitWhenFinished.setOnAction(event -> settingsChanged.set(true));
        
        saveSettings.setOnAction(event -> {
            Context ctx = arrangeContext();
            appProps = ctx.extractPropertiesToStore();
            Main.storeProperties(appProps);
            settingsChanged.set(false);
        });
        execute.setOnAction(this::execute);
        
        showWorkDir.setOnAction(event -> openDir(context.get(Props.SYS_WORK_DIR_BASE)));
        deleteOldWorkDirs.setOnAction(event -> {
            Path parent = context.get(Props.SYS_WORK_DIR_BASE);
            String msg = String.format("次のフォルダの内容を削除します。よろしいですか？\n%s", parent);
            new Alert(AlertType.CONFIRMATION, msg)
                    .showAndWait()
                    .filter(response -> response == ButtonType.OK)
                    .ifPresent(response -> deleteChildren(parent));
        });
        
        // 初期値の設定
        considerRowGaps.setSelected(context.get(Props.APP_CONSIDER_ROW_GAPS));
        considerColumnGaps.setSelected(context.get(Props.APP_CONSIDER_COLUMN_GAPS));
        compareOnValue.setSelected(context.get(Props.APP_COMPARE_ON_VALUE));
        showPaintedSheets.setSelected(context.get(Props.APP_SHOW_PAINTED_SHEETS));
        showResultText.setSelected(context.get(Props.APP_SHOW_RESULT_TEXT));
        exitWhenFinished.setSelected(context.get(Props.APP_EXIT_WHEN_FINISHED));
    }
    
    private File getSelectedFile(File current) {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("比較対象ブックの選択");
        if (current != null) {
            chooser.setInitialDirectory(current.getParentFile());
            chooser.setInitialFileName(current.getName());
        } else if (prevSelected != null) {
            chooser.setInitialDirectory(prevSelected.getParentFile());
        }
        chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                "Excel ブック", "*.xls", "*.xlsx", "*.xlsm"));
        
        File selected = chooser.showOpenDialog(settingsPane.getScene().getWindow());
        return selected != null ? selected : current;
    }
    
    private ObservableList<String> getSheetNames(File file) {
        if (file == null || !file.canRead()) {
            return FXCollections.emptyObservableList();
        }
        try {
            return FXCollections.observableList(POIUtils.getSheetNames(file));
        } catch (Exception e) {
            return FXCollections.emptyObservableList();
        }
    }
    
    private void onDragDropped(DragEvent event, Node target) {
        Dragboard db = event.getDragboard();
        if (db.hasFiles()) {
            target.setUserData(db.getFiles().get(0));
        }
        event.setDropCompleted(db.hasFiles());
        event.consume();
    }
    
    private void onDragOver(DragEvent event) {
        if (event.getDragboard().hasFiles()) {
            event.acceptTransferModes(TransferMode.LINK);
        }
        event.consume();
    }
    
    private void execute(ActionEvent event) {
        isRunning.set(true);
        Context context = arrangeContext();
        Task<Path> task = MenuTask.of(context);
        
        progress.progressProperty().bind(task.progressProperty());
        messages.textProperty().bind(task.messageProperty());
        
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.submit(task);
        
        task.addEventHandler(WorkerStateEvent.WORKER_STATE_SUCCEEDED, wse -> {
            executor.shutdown();
            progress.progressProperty().unbind();
            messages.textProperty().unbind();
            if (context.get(Props.APP_EXIT_WHEN_FINISHED)) {
                Platform.exit();
            } else {
                isRunning.set(false);
            }
        });
        task.addEventHandler(WorkerStateEvent.WORKER_STATE_FAILED, wse -> {
            executor.shutdown();
            progress.progressProperty().unbind();
            messages.textProperty().unbind();
            isRunning.set(false);
            new Alert(AlertType.WARNING, task.getException().getMessage(), ButtonType.OK)
                    .showAndWait();
        });
    }
    
    private Context arrangeContext() {
        return Context.Builder.of(appProps)
                .set(Props.CURR_MENU, compareBooks.isSelected() ? Menu.COMPARE_BOOKS : Menu.COMPARE_SHEETS)
                .set(Props.CURR_FILE1, (File) bookPath1.getUserData())
                .set(Props.CURR_FILE2, (File) bookPath2.getUserData())
                .set(Props.CURR_SHEET_NAME1, sheetName1.getValue())
                .set(Props.CURR_SHEET_NAME2, sheetName2.getValue())
                .set(Props.APP_CONSIDER_ROW_GAPS, considerRowGaps.isSelected())
                .set(Props.APP_CONSIDER_COLUMN_GAPS, considerColumnGaps.isSelected())
                .set(Props.APP_COMPARE_ON_VALUE, compareOnValue.isSelected())
                .set(Props.APP_SHOW_PAINTED_SHEETS, showPaintedSheets.isSelected())
                .set(Props.APP_SHOW_RESULT_TEXT, showResultText.isSelected())
                .set(Props.APP_EXIT_WHEN_FINISHED, exitWhenFinished.isSelected())
                .build();
    }
    
    private void openDir(Path target) {
        try {
            if (!Files.isDirectory(target)) {
                Files.createDirectories(target);
            }
            Desktop.getDesktop().open(target.toFile());
        } catch (IOException e) {
            String msg = String.format("作業用フォルダの表示に失敗しました。\n%s", target.toString());
            new Alert(AlertType.WARNING, msg, ButtonType.OK).showAndWait();
        }
    }
    
    private void deleteChildren(Path parent) {
        assert parent != null;
        assert Files.isDirectory(parent);
        
        try (Stream<Path> children = Files.walk(parent)) {
            children.filter(path -> !path.equals(parent))
                    .sorted(Comparator.reverseOrder())
                    .forEach(path -> {
                        try {
                            Files.deleteIfExists(path);
                        } catch (IOException e) {
                            // failed to delete : nop
                        }
                    });
        } catch (IOException e) {
            // failed to walk : nop
        }
    }
}
