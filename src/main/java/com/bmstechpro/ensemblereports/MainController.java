package com.bmstechpro.ensemblereports;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.stage.FileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

public class MainController implements Initializable {
    @FXML
    private Label msgLbl;
    @FXML
    private ProgressBar progressBar;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        progressBar.setVisible(false);
        msgLbl.setText("");
    }

    @FXML
    private void onLoad() {
//                ReportLoader loader = new ReportLoader();
//        loader.load();

        msgLbl.setText("");
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
        File file = fileChooser.showOpenDialog(msgLbl.getScene().getWindow());

        if (file == null) return;
        progressBar.setVisible(true);
        Thread thread = new Thread(() -> {
            try {
                new ReportLoader().load(file);

            } catch (IOException | InvalidFormatException e) {
                Platform.runLater(() -> {
                    progressBar.setVisible(false);
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setContentText("Unable to process file");
                    alert.show();
                });
            }

            Platform.runLater(() -> {
                progressBar.setVisible(false);
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setContentText("Report file was generated successfully.\n" +
                        "File saved in the original folder");
                alert.show();

            });
        });
        thread.start();


    }


}