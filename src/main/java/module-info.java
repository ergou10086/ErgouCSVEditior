module hbnu.project.ergoucsveditior {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;
    requires transitive javafx.base;
    requires java.desktop;

    requires org.controlsfx.controls;
    requires org.kordamp.bootstrapfx.core;
    requires eu.hansolo.tilesfx;
    
    // Apache Commons CSV for CSV parsing
    requires org.apache.commons.csv;
    requires itextpdf;
    requires org.apache.poi.ooxml;
    
    // MySQL JDBC for database persistence
    requires java.sql;

    opens hbnu.project.ergoucsveditior to javafx.fxml;
    opens hbnu.project.ergoucsveditior.controller to javafx.fxml;
    
    exports hbnu.project.ergoucsveditior;
    exports hbnu.project.ergoucsveditior.controller;
    exports hbnu.project.ergoucsveditior.model;
    exports hbnu.project.ergoucsveditior.service;
}