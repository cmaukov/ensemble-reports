module com.bmstechpro.ensemblereports {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;


    opens com.bmstechpro.ensemblereports to javafx.fxml;
    exports com.bmstechpro.ensemblereports;
}