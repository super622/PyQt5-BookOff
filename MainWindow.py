import time
import xlwt
from PyQt5 import QtCore, QtWidgets, Qt, QtGui
from PyQt5.QtCore import QThread, pyqtSignal, QSettings, QRect, QSize
from PyQt5.QtGui import QColor, QIcon
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QProgressBar
from pyqtspinner import WaitingSpinner
import action


class RequestThread(QThread):
	request_completed = pyqtSignal(str)

	def __init__(self, handler):
		super().__init__()
		self.ui_handler = handler

	def run(self):
		self.request_completed.emit("start")

		if not self.ui_handler.main_window.isStop:
			# result = self.ui_handler.product_list_download_from_amazon()
			# print(result)
			time.sleep(5)
			self.request_completed.emit("reading")
			# self.ui_handler.read_product_list_from_file()
			time.sleep(5)
			key_code = '4580128895383'
			self.ui_handler.get_product_url(key_code)
					
		# val = 100 / len(self.keyword_arr) * len(self.ui_handler.products)
		# self.request_completed.emit(str(val))
		
		self.request_completed.emit("stop")
		self.quit()

class Ui_MainWindow(object):
	keyword_arr = []

	def __init__(self):
		super().__init__()
		self.spinner = None
		self.progressBar = None
		self.statusLabel = None
		self.tbl_dataview = None
		self.btn_start = None
		self.horizontalLayout = None
		self.horizontalLayout_2 = None
		self.horizontalLayout_3 = None
		self.btn_export = None
		self.gridLayout = None
		self.centralwidget = None
		self.verticalLayout = None
		self.isStop = True
		self.settings = None
		self.cmb_export_type = None
		self.request_thread = None
		self.ui_handler = action.ActionManagement(self)

	def setupUi(self, MainWindow):
		MainWindow.setObjectName("MainWindow")
		MainWindow.setWindowModality(QtCore.Qt.ApplicationModal)
		MainWindow.setEnabled(True)
		MainWindow.resize(820, 600)
		sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
		sizePolicy.setHorizontalStretch(0)
		sizePolicy.setVerticalStretch(0)
		sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
		MainWindow.setSizePolicy(sizePolicy)
		MainWindow.setMinimumSize(QtCore.QSize(820, 600))
		MainWindow.setAcceptDrops(False)
		MainWindow.setLayoutDirection(QtCore.Qt.LeftToRight)
		MainWindow.setAutoFillBackground(True)
		MainWindow.setAnimated(False)
		MainWindow.setDocumentMode(False)
		MainWindow.setTabShape(QtWidgets.QTabWidget.Triangular)
		MainWindow.setDockNestingEnabled(False)
		MainWindow.setUnifiedTitleAndToolBarOnMac(False)

		self.centralwidget = QtWidgets.QWidget(MainWindow)
		self.centralwidget.setEnabled(True)

		sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
		sizePolicy.setHorizontalStretch(0)
		sizePolicy.setVerticalStretch(0)
		sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())

		self.centralwidget.setSizePolicy(sizePolicy)
		self.centralwidget.setLayoutDirection(QtCore.Qt.LeftToRight)
		self.centralwidget.setObjectName("centralwidget")
		self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
		self.gridLayout.setObjectName("gridLayout")

		self.verticalLayout = QtWidgets.QVBoxLayout()
		self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetMinAndMaxSize)
		self.verticalLayout.setContentsMargins(0, 0, 0, 0)
		self.verticalLayout.setObjectName("verticalLayout")

		self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
		self.horizontalLayout_2.setSpacing(12)
		self.horizontalLayout_2.setObjectName("horizontalLayout_2")
		self.btn_start = QtWidgets.QPushButton(self.centralwidget)
		self.btn_start.setMinimumSize(QtCore.QSize(16777215, 30))
		self.btn_start.setMaximumSize(QtCore.QSize(16777215, 30))
		self.btn_start.setObjectName("btn_start")
		self.btn_start.clicked.connect(self.handle_btn_start_clicked)
		self.horizontalLayout_2.addWidget(self.btn_start)

		self.btn_export = QtWidgets.QPushButton(self.centralwidget)
		self.btn_export.setMinimumSize(QtCore.QSize(16777215, 30))
		self.btn_export.setMaximumSize(QtCore.QSize(16777215, 30))
		self.btn_export.setObjectName("btn_export")
		self.btn_export.setText("Export")
		self.btn_export.setEnabled(False)
		self.btn_export.clicked.connect(self.savefile)
		self.horizontalLayout_2.addWidget(self.btn_export)
		self.verticalLayout.addLayout(self.horizontalLayout_2)
		
		self.horizontalLayout = QtWidgets.QHBoxLayout()
		self.horizontalLayout.setObjectName("horizontalLayout")
		self.tbl_dataview = QtWidgets.QTableView(self.centralwidget)
		self.tbl_dataview.setObjectName("tbl_dataview")
		self.tbl_dataview.doubleClicked.connect(self.handle_cell_click)
		header_labels = ["Result"]
		model = QtGui.QStandardItemModel(0, 0)
		model.setHorizontalHeaderLabels(header_labels)
		self.tbl_dataview.setModel(model)
		self.horizontalLayout.addWidget(self.tbl_dataview)
		self.horizontalLayout.setStretch(0, 6)
		self.verticalLayout.addLayout(self.horizontalLayout)

		self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
		self.horizontalLayout_3.setObjectName('horizontalLayout_3')
		self.statusLabel = QtWidgets.QLabel(self.centralwidget)
		self.statusLabel.setMinimumSize(QtCore.QSize(16777215, 20))
		self.statusLabel.setMaximumSize(QtCore.QSize(16777215, 50))
		self.statusLabel.setVisible(True)
		self.horizontalLayout_3.addWidget(self.statusLabel)
		self.verticalLayout.addLayout(self.horizontalLayout_3)

		self.progressBar = QProgressBar(self)
		self.progressBar.setVisible(False)
		self.verticalLayout.addWidget(self.progressBar)
		self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
		MainWindow.setCentralWidget(self.centralwidget)
		self.spinner = WaitingSpinner(
				parent = self.tbl_dataview,
				disable_parent_when_spinning = False,
				roundness = 100,
				fade = 88,
				radius = 29,
				lines = 58,
				line_length = 11,
				line_width = 14,
				speed = 1,
				color = QColor(0, 0, 0, 255)
		)
		self.spinner.setObjectName("spinner")
		self.retranslateUi(MainWindow)
		self.retranslateUi(MainWindow)
		MainWindow.setWindowIcon(QtGui.QIcon('delivery.ico'))
		QtCore.QMetaObject.connectSlotsByName(MainWindow)
		self.loadSettings()

	# self.loading_thread = threading.Thread(target=self.simulateLoading)
	# self.loading_thread.start()

	def closeEvent(self, event):
		self.saveSettings()
		event.accept()

	def loadSettings(self):
		# Load the previous settings using QSettings
		settings = QSettings("YourCompany", "YourApp")
		self.resize(settings.value("WindowSize", QSize(800, 600)))
		for col in range(self.tbl_dataview.model().columnCount()):
			width = settings.value(f"columnWidth_{col}", type = int, defaultValue = 100)  # Set a default width if not found
			self.tbl_dataview.setColumnWidth(col, width)

	def saveSettings(self):
		# Save the current window size and column state using QSettings
		settings = QSettings("YourCompany", "YourApp")
		settings.setValue("WindowSize", self.size())
		for col in range(self.tbl_dataview.model().columnCount()):
			width = self.tbl_dataview.columnWidth(col)
			settings.setValue(f"columnWidth_{col}", width)

	# Define the custom slot to handle cell clicks
	def handle_cell_click(self, index):
		row = index.row()
		col = index.column()
		if col == 9 and self.ui_handler.products_list[row] != "":
			import webbrowser
			webbrowser.open(self.ui_handler.products_list[row])
	
	def handle_btn_start_clicked(self):
		print(self.isStop)
		if self.isStop:
			self.btn_start.setText("Stop")
			self.ui_handler.products_list = []
			self.request_thread = RequestThread(self.ui_handler)
			self.request_thread.request_completed.connect(self.handle_request_completed)
			self.request_thread.start()
			self.isStop = False
		else:
			self.isStop = True
			self.statusLabel.setVisible(False)
			self.progressBar.setVisible(True)
			self.btn_start.setEnabled(False)
			self.request_thread.exit()

	def savefile(self):
		filename, _ = QFileDialog.getSaveFileName(self, 'Save File', '', ".xls(*.xls)")
		if filename:
			wbk = xlwt.Workbook()
			sheet = wbk.add_sheet("sheet", cell_overwrite_ok = True)
			style = xlwt.XFStyle()
			font = xlwt.Font()
			font.bold = True
			style.font = font
			model = self.tbl_dataview.model()
			if self.cmb_export_type.currentText() == "Disco'd":
				sheet.write(0, 0, 'Variant SKU [ID]', style = style)
				sheet.write(0, 1, 'COMMAND', style = style)
				sheet.write(0, 2, 'STATUS', style = style)
				sheet.write(0, 3, 'PUBLISHED', style = style)
				sheet.write(0, 4, 'PUBLISHED SCOPE', style = style)
				sheet.write(0, 5, 'Variant Inventory Tracker', style = style)
				sheet.write(0, 6, 'Variant Inventory Policy', style = style)
				sheet.write(0, 7, 'Variant Fulfillment Service', style = style)
				sheet.write(0, 8, 'Variant Inventory Qty', style = style)
				row_count = 1
				for r in range(model.rowCount()):
					if model.data(model.index(r, 3)) == self.cmb_export_type.currentText():
						sheet.write(row_count, 0, model.data(model.index(r, 0)), style = style)
						sheet.write(row_count, 1, "MERGE", style = style)
						sheet.write(row_count, 2, "Disco'd", style = style)
						sheet.write(row_count, 3, "FALSE", style = style)
						sheet.write(row_count, 4, "global", style = style)
						sheet.write(row_count, 5, "shopify", style = style)
						sheet.write(row_count, 6, "deny", style = style)
						sheet.write(row_count, 7, "manual", style = style)
						sheet.write(row_count, 8, "0", style = style)
						row_count += 1
				wbk.save(filename)
				return
			row_count = 0
			col_count = 1
			
			for c in range(model.columnCount()):
				if c != 2:  # Skip column 3
					text = model.headerData(c, QtCore.Qt.Horizontal)
					sheet.write(0, c + col_count, text, style = style)
				else:
					col_count = 0
			col_count = 1
			for r in range(model.rowCount()):
				if model.data(model.index(r, 3)) == self.cmb_export_type.currentText() or self.cmb_export_type.currentText() == "All":
					row_count += 1
					sheet.write(row_count, 0, row_count, style = style)
					for c in range(model.columnCount()):
						if c != 2:
							text = model.data(model.index(r, c))
							if c == 0:
								col_count = 1
							sheet.write(row_count, c + col_count, text)
						else:
							col_count = 0
			
			wbk.save(filename)
	
	def handle_request_completed(self, response_text):
		print(response_text)
		if response_text == "start":
			self.statusLabel.setText("Downloading ... ")
			self.btn_export.setEnabled(False)
			self.spinner.start()
		elif response_text == "stop":
			self.spinner.stop()
			self.btn_start.setText("Search")
			self.btn_export.setEnabled(True)
			self.btn_start.setEnabled(True)
			self.isStop = True
		elif response_text == "reading":
			self.progressBar.setValue(0)
			self.statusLabel.setText("Reading ... ")
		else:
			self.progressBar.setVisible(True)
			self.progressBar.setValue(round(float(response_text)))
	
	def retranslateUi(self, MainWindow):
		_translate = QtCore.QCoreApplication.translate
		MainWindow.setWindowTitle(_translate("MainWindow", "Amazon - BookOff"))
		self.btn_start.setText(_translate("MainWindow", "Start"))
