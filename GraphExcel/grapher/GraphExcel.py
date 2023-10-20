import sys
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from Graphs import ExcelManager

class ExcelGraphApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.excelManager = ExcelManager()
        self.fig, self.ax = plt.subplots()
        self.file = ""
        
    def initUI(self):
        self.setWindowTitle('GraphExcel')
        
        self.main_layout = self.createInitialLayout()

        # Set the layout of the main widget
        self.setLayout(self.main_layout)

    def createTitle(self):
        # Calculate the height for the title container (25% of the body height)
        title_container_height = self.height() * 0.25
        
        # Add a QFrame as a container for the title with a darker background
        title_container = QFrame(self)
        title_container.setStyleSheet("background-color: #dae9b6;")  # Darker background color
        title_container.setFrameShape(QFrame.StyledPanel)
        title_container.setFrameShadow(QFrame.Raised)
        
        # Set the height of the title container
        title_container.setFixedHeight(int(title_container_height))
        
        # Create a QVBoxLayout within the title container to center the title horizontally
        title_layout = QVBoxLayout(title_container)
        title_layout.setAlignment(QtCore.Qt.AlignCenter)
        
        # Add a QLabel to display the title within the container
        title_label = QLabel('Excel Data to Bar Graph', title_container)
        title_label.setAlignment(QtCore.Qt.AlignCenter) 
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)
        
        # Add the title label to the layout within the title container
        title_layout.addWidget(title_label)
        
        return title_container
        
    def createButtons(self):
        buttons_layout = QHBoxLayout()
        
        # Button to prompt user to select Excel file
        btn_select_excel = QPushButton('Select Excel File', self)
        btn_select_excel.clicked.connect(self.showDialog)
        btn_select_excel.setStyleSheet("background-color: #dae9b6;")  # Fall color: orange
        buttons_layout.addWidget(btn_select_excel)
        
        # Add "Generate Graph" button
        btn_generate_graph = QPushButton('Generate Graph', self)
        btn_generate_graph.clicked.connect(self.generateGraphLayout)
        btn_generate_graph.setStyleSheet("background-color: #dae9b6;")  # Fall color: orange
        buttons_layout.addWidget(btn_generate_graph)
        
        # Button to exit the application
        btn_exit = QPushButton('Exit', self)
        btn_exit.clicked.connect(self.close_application)
        btn_exit.setStyleSheet("background-color: #dae9b6;")
        buttons_layout.addWidget(btn_exit)
        
        return buttons_layout
        
    def display_file_path(self):
        # Create a container for the file path
        path_container = QFrame(self)
        path_container.setStyleSheet("background-color: #dae9b6;")  # Lighter background color (tan)
        path_container.setFrameShape(QFrame.StyledPanel)
        path_container.setFrameShadow(QFrame.Raised)

        # Create a layout for the file path
        path_layout = QVBoxLayout(path_container)
        path_layout.setAlignment(QtCore.Qt.AlignCenter)

        # Add a QLabel to display the file path within the container
        self.path_label = QLabel(f'Selected File: None', path_container)
        font = self.path_label.font()
        font.setPointSize(12)  # Adjust the font size as needed
        self.path_label.setFont(font)

        # Add the path label to the layout within the container
        path_layout.addWidget(self.path_label)
        
        return path_container

    def createInitialLayout(self):
        # Set an initial fixed size for the window
        self.resize(600, 300)
        
        # Set the background color of the main widget
        self.setStyleSheet("background-color: #ffdcbc;")

        # Create main layout for window
        layout = QVBoxLayout()
        layout.setAlignment(QtCore.Qt.AlignTop)
        
        #Add Title to App
        layout.addWidget(self.createTitle())
        
        layout.addStretch(1)
        
        # Create a label to display the selected file path  
        layout.addWidget(self.display_file_path())
        
        # Add a stretch to push the buttons to the bottom
        layout.addStretch(1)
             
        # Create and Add the buttons layout to the main layout
        layout.addLayout(self.createButtons())
        
        return layout

    def createGraphButtons(self):
        button_names = ["Diabetes", "Cholesterol", "Blood Pressure"]
        
        # Create four new buttons and add them to the layout
        button_layout = QHBoxLayout()
        
        for name in button_names:
            button = QPushButton(name, self)
            button.setStyleSheet("background-color: #FF8C42;")  # Fall color: orange
            button.clicked.connect(lambda _, graph_type=name: self.generate_graph(graph_type))
            button_layout.addWidget(button)
            
        restart_button = QPushButton('Restart', self)
        restart_button.setStyleSheet("background-color: #FF8C42;")  # Fall color: orange
        restart_button.clicked.connect(self.reset_app)
        button_layout.addWidget(restart_button)
        
        return button_layout

    def generateGraphLayout(self):
        if(len(self.file) == 0 ):
            return
        
        # Clear the existing layout and set the new layout with the plot
        self.clearLayout(self.main_layout)
                
        # Create a new layout
        new_layout = QVBoxLayout()

        # Add the buttons layout to the main layout
        new_layout.addLayout(self.createGraphButtons())
        self.main_layout.addLayout(new_layout)

        # Resize the window after plotting the graph
        self.resize(1200, 600)
        
        # Center the window on the screen
        self.center()
        
        # Display initial graph
        self.generate_graph()
        
    def clearLayout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clearLayout(item.layout())
                    
    def showDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)", options=options)

        if self.file:
            print(self.file)
            self.path_label.setText(f'Selected File: {self.file}')
            
    def generate_graph(self, graph_type = "Diabetes"):
        if graph_type is None:
            print("Invalid Graph Type")
            return
        
        # Clear previous graph
        self.clearGraph()
        
        # Based on the button pressed, analyze and generate graph
        if graph_type == "Diabetes":
            self.fig = self.excelManager.analyzeDiabetes(self.file)
        elif graph_type == "Cholesterol":
            self.fig = self.excelManager.analyzeCholesterol(self.file)
        elif graph_type == "Blood Pressure":
            self.fig = self.excelManager.analyzeBloodPressure(self.file)
        else:
            return  # Return if an invalid graph type is specified
        
        # Create a canvas for the plot
        canvas = FigureCanvas(self.fig)
        
        # Create a horizontal layout
        horizontal_layout = QHBoxLayout()
        horizontal_layout.setAlignment(QtCore.Qt.AlignCenter)
        horizontal_layout.addWidget(canvas)
        
        # Add the horizontal layout to the main layout
        self.main_layout.addWidget(canvas)
        
    def clearGraph(self):
        # Clear any graphs displayed in the layout
        for i in reversed(range(self.main_layout.count())):
            item = self.main_layout.itemAt(i)
            widget = item.widget()
            if widget is not None:
                # Check if the widget is a FigureCanvasQTAgg instance (Matplotlib canvas)
                if isinstance(widget, FigureCanvas):
                    widget.deleteLater()
        
        self.ax.cla()
        
    def reset_app(self):
        # Clear the existing layout and set the new layout with the plot
        self.clearLayout(self.main_layout)
        
        # Get initial layout
        layout = self.createInitialLayout()
        
        # Add the new layout to the main layout
        self.main_layout.addLayout(layout)

    def center(self):
        # Center the window on the screen
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move(int((screen.width() - size.width()) / 2), int((screen.height() - size.height()) / 2))

    def close_application(self):
        # Close the application
        sys.exit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelGraphApp()
    ex.show()
    sys.exit(app.exec_())