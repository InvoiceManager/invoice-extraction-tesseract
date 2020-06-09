#!/usr/bin/python3
# -*- coding: utf-8 -*-

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap, QPalette, QPainter
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from PyQt5.QtWidgets import *

from PyQt5.QtGui import *
from PyQt5.QtCore import *
import cv2
import sys
import os

from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2
jsonFile = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\final\body_tabel.json'

file = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\open.png'
cr2 = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\src\crop2.jpg'

header_final = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\final\fact.txt'

class QImageViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.printer = QPrinter()
        self.scaleFactor = 0.0

        self.imageLabel = QLabel()
        self.imageLabel.setBackgroundRole(QPalette.Base)
        self.imageLabel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.imageLabel.setScaledContents(True)

        self.scrollArea = QScrollArea()
        self.scrollArea.setBackgroundRole(QPalette.Dark)
        self.scrollArea.setWidget(self.imageLabel)
        self.scrollArea.setVisible(False)

        self.setCentralWidget(self.scrollArea)

        self.createActions()
        self.createMenus()

        self.setWindowTitle("Invoice Recognition")
        self.resize(800, 600)


    def open(self):
        options = QFileDialog.Options()
        # fileName = QFileDialog.getOpenFileName(self, "Open File", QDir.currentPath())
        fileName, _ = QFileDialog.getOpenFileName(self, 'QFileDialog.getOpenFileName()', '',
                                                  'Images (*.png *.jpeg *.jpg *.bmp *.gif)', options=options)
        if fileName:
            image = QImage(fileName)
            if image.isNull():
                QMessageBox.information(self, "Image Viewer", "Cannot load %s." % fileName)
                return
            image.save('open.png')
            self.imageLabel.setPixmap(QPixmap.fromImage(image))
            self.scaleFactor = 1.0

            self.scrollArea.setVisible(True)
            self.printAct.setEnabled(True)
            self.fitToWindowAct.setEnabled(True)
            self.updateActions()

            if not self.fitToWindowAct.isChecked():
                self.imageLabel.adjustSize()

    def print_(self):
        dialog = QPrintDialog(self.printer, self)
        if dialog.exec_():
            painter = QPainter(self.printer)
            rect = painter.viewport()
            size = self.imageLabel.pixmap().size()
            size.scale(rect.size(), Qt.KeepAspectRatio)
            painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
            painter.setWindow(self.imageLabel.pixmap().rect())
            painter.drawPixmap(0, 0, self.imageLabel.pixmap())

    def zoomIn(self):
        self.scaleImage(1.25)

    def zoomOut(self):
        self.scaleImage(0.8)

    def normalSize(self):
        self.imageLabel.adjustSize()
        self.scaleFactor = 1.0

    def fitToWindow(self):
        fitToWindow = self.fitToWindowAct.isChecked()
        self.scrollArea.setWidgetResizable(fitToWindow)
        if not fitToWindow:
            self.normalSize()

        self.updateActions()

    def slotRotate(self):
        self.rotate(90)

    def slotRotateAnti(self):
        self.rotate(90)

    def extract(self):
        super().__init__()
        os.system('click_and_crop.py --image open.png')

        #header
        type = recognitionTypeHeader.getType(file)
        detectTableWithCv2.detect(cr2)
        header_data = recognitionTypeHeader.getContent(file, header_final, type)

        # box
        # setting title
        self.setWindowTitle("Extract Data")
        # setting geometry panel
        self.setGeometry(100, 100, 600, 1000)

        # create label with text for combo text
        self.tip_factura_label = QLabel(self)
        self.nr_factura_label = QLabel(self)
        self.seria_fact_label = QLabel(self)
        self.data_emiterii_label = QLabel(self)
        self.data_scad_label = QLabel(self)
        self.nume_furn_label = QLabel(self)
        self.adresa_furn_label = QLabel(self)
        self.cont_furn_label = QLabel(self)
        self.banca_furn_label = QLabel(self)
        self.nume_cump_label = QLabel(self)
        self.adresa_cump_label = QLabel(self)
        self.cont_cump_label = QLabel(self)
        self.banca_cump_label = QLabel(self)
        self.total_fact_label = QLabel(self)

        self.tip_factura_label.setText("Tip Fact:")
        self.nr_factura_label.setText("Nr Fact:")
        self.seria_fact_label.setText("Seria Fact:")
        self.data_emiterii_label.setText("Data emiterii:")
        self.data_scad_label.setText("Data scad:")
        self.nume_furn_label.setText("Nume Furn:")
        self.adresa_furn_label.setText("Addresa furnizor:")
        self.cont_furn_label.setText("Cont furn:")
        self.banca_furn_label.setText("Banca Furn:")
        self.nume_cump_label.setText("Nume Cump:")
        self.adresa_cump_label.setText("Adresa Cump:")
        self.cont_cump_label.setText("Cont Cump:")
        self.banca_cump_label.setText("Banca Cump:")
        self.total_fact_label.setText("Total factura:")

        # creating a combo box widget
        self.tip_factura_box = QComboBox(self)
        self.nr_factura_box = QComboBox(self)
        self.seria_fact_box = QComboBox(self)
        self.data_emiterii_box = QComboBox(self)
        self.data_scad_box = QComboBox(self)
        self.nume_furn_box = QComboBox(self)
        self.adresa_furn_box = QComboBox(self)
        self.cont_furn_box = QComboBox(self)
        self.banca_furn_box = QComboBox(self)
        self.nume_cump_box = QComboBox(self)
        self.adresa_cump_box = QComboBox(self)
        self.cont_cump_box = QComboBox(self)
        self.banca_cump_box = QComboBox(self)
        self.total_fact_box = QComboBox(self)

        # setting geometry of combo box
        self.tip_factura_box.setGeometry(200, 150, 300, 30)
        self.nr_factura_box.setGeometry(200, 150, 300, 30)
        self.seria_fact_box.setGeometry(200, 150, 300, 30)
        self.data_emiterii_box.setGeometry(200, 150, 300, 30)
        self.data_scad_box.setGeometry(200, 150, 300, 30)
        self.nume_furn_box.setGeometry(200, 150, 300, 30)
        self.adresa_furn_box.setGeometry(200, 150, 300, 30)
        self.cont_furn_box.setGeometry(200, 150, 300, 30)
        self.banca_furn_box.setGeometry(200, 150, 300, 30)
        self.nume_cump_box.setGeometry(200, 150, 300, 30)
        self.adresa_cump_box.setGeometry(200, 150, 300, 30)
        self.cont_cump_box.setGeometry(200, 150, 300, 30)
        self.banca_cump_box.setGeometry(200, 150, 300, 30)
        self.total_fact_box.setGeometry(200, 150, 300, 30)

        # geek list
        geek_list = ["eon", "cubus", "other"]
        # adding list of items to combo box
        self.tip_factura_box.addItems([type])
        self.nr_factura_box.addItems([header_data[0]])
        self.seria_fact_box.addItems([header_data[1]])
        self.data_emiterii_box.addItems([header_data[2]])
        self.data_scad_box.addItems([header_data[3]])
        self.nume_furn_box.addItems([header_data[4]])
        self.adresa_furn_box.addItems([header_data[5]])
        self.cont_furn_box.addItems([header_data[6]])
        self.banca_furn_box.addItems([header_data[7]])
        self.nume_cump_box.addItems([header_data[8]])
        self.adresa_cump_box.addItems([header_data[9]])
        self.cont_cump_box.addItems([header_data[10]])
        self.banca_cump_box.addItems([header_data[11]])
        self.total_fact_box.addItems([header_data[12]])

        print("Nr Fact:", header_data[0])
        print("Seria Fact:", header_data[1])
        print("Data emiterii:", header_data[2])
        print("Data scad: ", header_data[3])
        print("Nume Furn:", header_data[4])
        print("Addresa furnizor: ", header_data[5])
        print("Cont furn: ", header_data[6])
        print("Banca Furn:", header_data[7])
        print("Nume Cump:", header_data[8])
        print("Adresa Cump: ", header_data[9])
        print("Cont Cump:", header_data[10])
        print("Banca Cump:", header_data[11])
        print("Total factura:", header_data[12])

        # creating a editable combo box
        self.tip_factura_box.setEditable(True)
        self.nr_factura_box.setEditable(True)
        self.seria_fact_box.setEditable(True)
        self.data_emiterii_box.setEditable(True)
        self.data_scad_box.setEditable(True)
        self.nume_furn_box.setEditable(True)
        self.adresa_furn_box.setEditable(True)
        self.cont_furn_box.setEditable(True)
        self.banca_furn_box.setEditable(True)
        self.nume_cump_box.setEditable(True)
        self.adresa_cump_box.setEditable(True)
        self.cont_cump_box.setEditable(True)
        self.banca_cump_box.setEditable(True)
        self.total_fact_box.setEditable(True)

        #position
        self.tip_factura_box.move(100, 100)
        self.nr_factura_box.move(100, 150)
        self.seria_fact_box.move(100, 200)
        self.data_emiterii_box.move(100, 250)
        self.data_scad_box.move(100, 300)
        self.nume_furn_box.move(100, 350)
        self.adresa_furn_box.move(100, 400)
        self.cont_furn_box.move(100, 450)
        self.banca_furn_box.move(100, 500)
        self.nume_cump_box.move(100, 550)
        self.adresa_cump_box.move(100, 600)
        self.cont_cump_box.move(100, 650)
        self.banca_cump_box.move(100, 700)
        self.total_fact_box.move(100, 750)

        self.tip_factura_label.move(10, 100)
        self.nr_factura_label.move(10, 150)
        self.seria_fact_label.move(10, 200)
        self.data_emiterii_label.move(10, 250)
        self.data_scad_label.move(10, 300)
        self.nume_furn_label.move(10, 350)
        self.adresa_furn_label.move(10, 400)
        self.cont_furn_label.move(10, 450)
        self.banca_furn_label.move(10, 500)
        self.nume_cump_label.move(10, 550)
        self.adresa_cump_label.move(10, 600)
        self.cont_cump_label.move(10, 650)
        self.banca_cump_label.move(10, 700)
        self.total_fact_label.move(10, 750)

        # showing all the widgets
        self.show()


    def export(self):
        type = self.tip_factura_box.currentText()
        header_data = recognitionTypeHeader.getContent(file, header_final, type)

        header_data[0] = self.nr_factura_box.currentText()
        header_data[1] = self.seria_fact_box.currentText()
        header_data[2] = self.data_emiterii_box.currentText()
        header_data[3] = self.data_scad_box.currentText()
        header_data[4] = self.nume_furn_box.currentText()
        header_data[5] = self.adresa_furn_box.currentText()
        header_data[6] = self.cont_furn_box.currentText()
        header_data[7] = self.banca_furn_box.currentText()
        header_data[8] = self.nume_cump_box.currentText()
        header_data[9] = self.adresa_cump_box.currentText()
        header_data[10] = self.cont_cump_box.currentText()
        header_data[11] = self.banca_cump_box.currentText()
        header_data[12] = self.total_fact_box.currentText()
        sheet_name = header_data[4].replace(" ", "_") + "-" + header_data[0].replace(" ", "_")
        recognitionTypeHeader.writeExcelHeader(header_data, 'Output_Invoices.xlsx', sheet_name)
        recognitionTypeHeader.writeExcelTable(jsonFile, 'Output_Invoices.xlsx', sheet_name)

    def about(self):
        QMessageBox.about(self, "About Image Recognition",
                            "<p>For an image with an invoice in any format, the aim is "
                            "to extract important and relevant data from it, so that it can be "
                            "easily observed in an excel file by a company's accounting.</p>")

    def createActions(self):
        self.openAct = QAction("&Open...", self, shortcut="Ctrl+O", triggered=self.open)
        self.printAct = QAction("&Print...", self, shortcut="Ctrl+P", enabled=False, triggered=self.print_)
        self.exitAct = QAction("E&xit", self, shortcut="Ctrl+Q", triggered=self.close)
        self.slotRotateAct = QAction("rotate left", self, enabled=False, triggered=self.slotRotate)
        self.slotRotateAntiAct = QAction("rotate right", self, enabled=False, triggered=self.slotRotateAnti)
        self.zoomInAct = QAction("Zoom &In (25%)", self, shortcut="Ctrl++", enabled=False, triggered=self.zoomIn)
        self.zoomOutAct = QAction("Zoom &Out (25%)", self, shortcut="Ctrl+-", enabled=False, triggered=self.zoomOut)
        self.normalSizeAct = QAction("&Normal Size", self, shortcut="Ctrl+S", enabled=False, triggered=self.normalSize)
        self.fitToWindowAct = QAction("&Fit to Window", self, enabled=False, checkable=True, shortcut="Ctrl+F",
                                      triggered=self.fitToWindow)
        self.extractAct = QAction("&Extract data", self, triggered=self.extract)
        self.exportAct = QAction("&Export data", self, triggered=self.export)
        self.aboutAct = QAction("&About", self, triggered=self.about)
        self.aboutQtAct = QAction("About &Qt", self, triggered=qApp.aboutQt)

    def createMenus(self):
        self.fileMenu = QMenu("&File", self)
        self.fileMenu.addAction(self.openAct)
        self.fileMenu.addAction(self.printAct)
        self.fileMenu.addSeparator()
        self.fileMenu.addAction(self.exitAct)

        self.viewMenu = QMenu("&View", self)
        self.viewMenu.addAction(self.slotRotateAct)
        self.viewMenu.addAction(self.slotRotateAntiAct)
        self.viewMenu.addAction(self.zoomInAct)
        self.viewMenu.addAction(self.zoomOutAct)
        self.viewMenu.addAction(self.normalSizeAct)
        self.viewMenu.addSeparator()
        self.viewMenu.addAction(self.fitToWindowAct)

        self.extractMenu = QMenu("&Extract/Export", self)
        self.extractMenu.addAction(self.extractAct)
        self.extractMenu.addAction(self.exportAct)

        self.helpMenu = QMenu("&Help", self)
        self.helpMenu.addAction(self.aboutAct)
        self.helpMenu.addAction(self.aboutQtAct)

        self.menuBar().addMenu(self.fileMenu)
        self.menuBar().addMenu(self.viewMenu)
        self.menuBar().addMenu(self.extractMenu)
        self.menuBar().addMenu(self.helpMenu)

    def updateActions(self):
        self.slotRotateAct.setEnabled(not self.fitToWindowAct.isChecked())
        self.slotRotateAntiAct.setEnabled(not self.fitToWindowAct.isChecked())
        self.zoomInAct.setEnabled(not self.fitToWindowAct.isChecked())
        self.zoomOutAct.setEnabled(not self.fitToWindowAct.isChecked())
        self.normalSizeAct.setEnabled(not self.fitToWindowAct.isChecked())

    def scaleImage(self, factor):
        self.scaleFactor *= factor
        self.imageLabel.resize(self.scaleFactor * self.imageLabel.pixmap().size())

        self.adjustScrollBar(self.scrollArea.horizontalScrollBar(), factor)
        self.adjustScrollBar(self.scrollArea.verticalScrollBar(), factor)

        self.zoomInAct.setEnabled(self.scaleFactor < 3.0)
        self.zoomOutAct.setEnabled(self.scaleFactor > 0.333)

    def adjustScrollBar(self, scrollBar, factor):
        scrollBar.setValue(int(factor * scrollBar.value()
                               + ((factor - 1) * scrollBar.pageStep() / 2)))


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    imageViewer = QImageViewer()
    imageViewer.show()
    sys.exit(app.exec_())