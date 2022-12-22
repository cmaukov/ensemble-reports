package com.bmstechpro.ensemblereports;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.stage.FileChooser;

import java.io.File;
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

        msgLbl.textProperty().unbind();
        msgLbl.setText("");
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
        File file = fileChooser.showOpenDialog(msgLbl.getScene().getWindow());

        if (file == null) return;

        ReportLoader reportLoader = new ReportLoader(file);
        App.service.submit(reportLoader);
        msgLbl.textProperty().bind(reportLoader.messageProperty());
        reportLoader.setOnSucceeded(event -> {
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setContentText("Report file was generated successfully.\n" +
                    "File saved in the original folder");
            alert.show();
        });
        reportLoader.setOnFailed(event -> {
            msgLbl.textProperty().unbind();
            msgLbl.setText("");
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setContentText("Unable to process file");
            alert.show();
        });


    }


}