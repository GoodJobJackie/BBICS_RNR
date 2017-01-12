package application;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.stage.Stage;

public class Main extends Application {
	@Override
	public void start(Stage primaryStage) {
		try {
			primaryStage.setTitle("Administrative Documents");
			GridPane grid = new GridPane();
			grid.setAlignment(Pos.CENTER);
			grid.setHgap(20);
			grid.setVgap(10);
			grid.setPadding(new Insets(25, 25, 25, 25));

			Label lblFileDocs = new Label("File Documents");
			grid.add(lblFileDocs, 0, 1, 2, 1);

			Button btnClientTOS = new Button("Client File Table of Contents");
			grid.add(btnClientTOS, 1, 2);
			btnClientTOS.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File("C:/Users/jackie/Documents/Client Files/CLIENT RECORDS TOC.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnEmployeeTOS = new Button("Employee File Table of Contents");
			grid.add(btnEmployeeTOS, 1, 3);
			btnEmployeeTOS.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:/Users/jackie/Documents/Employee Files/Employee File TOC and Section Sheets.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnMgmtGuideline = new Button("Management Guidelines");
			grid.add(btnMgmtGuideline, 1, 4);
			btnMgmtGuideline.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File("C:\\Users\\jackie\\Documents\\Management Guidelines.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});

			Label lblEmployeeInfo = new Label("Employee Information");
			grid.add(lblEmployeeInfo, 0, 5, 2, 1);

			Button btnEmpContact = new Button("Employee Contact Information");
			grid.add(btnEmpContact, 1, 6);
			btnEmpContact.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Employee Contact Information.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnAgreementTech = new Button("Employee Agreement - Behavior Technician");
			grid.add(btnAgreementTech, 1, 7);
			btnAgreementTech.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Employment Agreement - Tutor_Behavioral Tech.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnAgreementLead = new Button("Employee Agreement - Lead Tutor");
			grid.add(btnAgreementLead, 1, 8);
			btnAgreementLead.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Employment Agreement - Lead Tutor.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnAgreementAsst = new Button("Employee Agreement - Assistant Consultant");
			grid.add(btnAgreementAsst, 1, 9);
			btnAgreementAsst.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Employment Agreement - Assistant Consultant_Behavioral Tech.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});

			Label lblPoliciesProceedures = new Label("Policies & Proceedures");
			grid.add(lblPoliciesProceedures, 2, 1, 2, 1);

			Button btnConfidentiality = new Button("Confidentiality Agreement");
			grid.add(btnConfidentiality, 3, 2);
			btnConfidentiality.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Employee Confidentiality Agreement.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnConduct = new Button("Employee Conduct & Expectations");
			grid.add(btnConduct, 3, 3);
			btnConduct.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\ProfessionalConductandExpectations.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnDrugPolicy = new Button("Drug-free Workplace Policy");
			grid.add(btnDrugPolicy, 3, 4);
			btnDrugPolicy.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\Drug Free Workplace Policy.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});

			Label lblPayroll = new Label("Payroll");
			grid.add(lblPayroll, 2, 5, 2, 1);

			Button btnI9 = new Button("I-9");
			grid.add(btnI9, 3, 6);
			btnI9.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File("C:\\Users\\jackie\\Documents\\Employee Files\\i-9.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});
			Button btnW4 = new Button("W-4");
			grid.add(btnW4, 3, 7);
			btnW4.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File("C:\\Users\\jackie\\Documents\\Employee Files\\fw4.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});

			Label lblMisc = new Label("Miscellaneous");
			grid.add(lblMisc, 2, 8, 2, 1);

			Button btnTimeOff = new Button("Time-off Request");
			grid.add(btnTimeOff, 3, 9, 2, 1);
			btnTimeOff.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					if (Desktop.isDesktopSupported()) {
						try {
							File myFile = new File(
									"C:\\Users\\jackie\\Documents\\Employee Files\\BBICS Time Off Request Form.pdf");
							Desktop.getDesktop().open(myFile);
						} catch (IOException ex) {
							// no application registered for PDFs
						}
					}
				}
			});

			Button btnClose = new Button("Close");
			grid.add(btnClose, 4, 10);
			btnClose.setOnAction(new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent e) {
					primaryStage.close();
				}
			});

			Scene scene = new Scene(grid, 700, 400);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.show();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		launch(args);
	}
}
