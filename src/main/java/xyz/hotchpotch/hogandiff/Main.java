package xyz.hotchpotch.hogandiff;

import java.io.File;
import java.io.IOException;
import java.io.Reader;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Objects;
import java.util.Properties;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

/**
 * このアプリケーションのエントリポイントを含む起動クラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class Main extends Application {
    
    // [static members] ********************************************************
    
    private static final String version = "v0.1.3";
    private static final String APP_PROP_PATH = "hogandiff.properties";
    
    /**
     * このアプリケーションのエントリポイントです。<br>
     * 
     * @param args アプリケーション実行時引数
     */
    public static void main(String[] args) {
        launch(args);
    }
    
    /**
     * アプリケーションプロパティファイルを読み込み、{@link Properties} オブジェクトとして返します。<br>
     * アプリケーションプロパティファイルの読み込みに失敗した場合は、空の {@link Properties} オブジェクトを返します。<br>
     * 
     * @return アプリケーションプロパティセット
     */
    // TODO: このメソッドはこのクラスのメンバで妥当か？？
    public static Properties loadProperties() {
        Properties appProps = new Properties();
        try (Reader r = Files.newBufferedReader(Paths.get(APP_PROP_PATH))) {
            appProps.load(r);
        } catch (IOException e) {
            // nop
        }
        return appProps;
    }
    
    /**
     * アプリケーションプロパティセットの内容をアプリケーションプロパティファイルに保存します。<br>
     * 
     * @param appProps アプリケーションプロパティセット
     */
    // TODO: このメソッドはこのクラスのメンバで妥当か？？
    public static void storeProperties(Properties appProps) {
        Objects.requireNonNull(appProps, "appProps");
        
        try (Writer w = Files.newBufferedWriter(Paths.get(APP_PROP_PATH))) {
            appProps.store(w, null);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    // [instance members] ******************************************************
    
    @Override
    public void start(Stage primaryStage) {
        assert primaryStage != null;
        
        try {
            VBox rootPane = FXMLLoader.load(getClass().getResource("MainApp.fxml"));
            Scene scene = new Scene(rootPane);
            scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
            Image icon = new Image(getClass().getResourceAsStream("favicon.png"));
            primaryStage.getIcons().add(icon);
            primaryStage.setScene(scene);
            primaryStage.setTitle("方眼Diff  -  " + version);
            primaryStage.setMinWidth(505);
            primaryStage.setMinHeight(465);
            primaryStage.show();
            
            List<String> args = getParameters().getRaw();
            if (args.size() == 2) {
                TextField bookPath1 = (TextField) rootPane.lookup("#bookPath1");
                TextField bookPath2 = (TextField) rootPane.lookup("#bookPath2");
                bookPath1.setUserData(new File(args.get(0)));
                bookPath2.setUserData(new File(args.get(1)));
                Button execute = (Button) rootPane.lookup("#execute");
                execute.fire();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
