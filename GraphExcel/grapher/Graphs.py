import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class ExcelManager():
    def __init__(self, file_path):
        self.filePath = file_path
        
    def readExcelFile(self, columnsToRead):
        try:
            data = pd.read_excel(self.filePath, usecols=columnsToRead)
        except FileNotFoundError:
            print("File Not Found")
            return

        return data

    def analyzeBloodPressure(self):
        # Get required data from excel file
        columns_to_read = ["Do you have a primary care physician", "Do you have Medical Coverage", "Diastolic", "Systolic"]
        data = self.readExcelFile(columns_to_read)
        
        # Convert 'Systolic' and 'Diastolic' columns to numeric
        data['Systolic'] = pd.to_numeric(data['Systolic'], errors='coerce')
        data['Diastolic'] = pd.to_numeric(data['Diastolic'], errors='coerce')      
        
        # Filter for patients who have abnormal Systolic or Diastolic values
        data = data[(data['Systolic'] > 120) | (data['Diastolic'] > 80)]
        
        # analyze the retrieved data
        conditionColumns = ['Do you have a primary care physician', 'Do you have Medical Coverage']        
        combinationsToCheck = [
            ['Yes, I do have a primary care physician', 'Yes, I do have medical coverage'],
            ['Yes, I do have a primary care physician', 'No, I do not have medical coverage'],
            ['No, I do not have a primary care physician', 'Yes, I do have medical coverage'],
            ['No, I do not have a primary care physician', 'No, I do not have medical coverage']
        ]
        
        # Count entries for each combination
        count = []
        for combination in combinationsToCheck:
            count.append(self._countData(data, conditionColumns, combination))
            
        print("Total HB Patients: " + str(len(data)))
        dataCount = [count[0], count[1], count[2], count[3], len(data)]
        
        return dataCount
        
    def analyzeDiabetes(self):
        # Get required data from excel file
        columns_to_read = ["Glucose Levels", "Do you have a primary care physician", "Do you have Medical Coverage"]
        data = self.readExcelFile(columns_to_read)
        
        # Filter 
        data = data[data["Glucose Levels"] > 126]
                
        # analyze the retrieved data
        conditionColumns = ['Do you have a primary care physician', 'Do you have Medical Coverage']        
        combinationsToCheck = [
            ['Yes, I do have a primary care physician', 'Yes, I do have medical coverage'],
            ['Yes, I do have a primary care physician', 'No, I do not have medical coverage'],
            ['No, I do not have a primary care physician', 'Yes, I do have medical coverage'],
            ['No, I do not have a primary care physician', 'No, I do not have medical coverage']
        ]
        
        # Count entries for each combination
        count = []
        for combination in combinationsToCheck:
            count.append(self._countData(data, conditionColumns, combination))
        
        print("Total Diabetes Patients: " + str(len(data)))                
        dataCount = [count[0], count[1], count[2], count[3], len(data)]
               
        return dataCount
          
    def analyzeCholesterol(self):
        data = {'Category': ['X', 'Y', 'Z'], 'People with Cholesterol': [15, 25, 35]}
        return self.generateBarGraph(data, "Cholesterol Data")
        
    def _countData(self, data, conditionColumns, conditionValues):        
        if len(conditionColumns) != len(conditionValues):
            raise ValueError("The number of condition columns must match the number of condition values.")

        condition = True
        counter = 0
        for col, val in zip(conditionColumns, conditionValues):
            condition &= (data[col] == val)

        count = len(data[condition])
        
        return count       
        
    def generateComboGraph(self):
        diabetes = self.analyzeDiabetes()
        bloodPressure = self.analyzeBloodPressure()
        
        # X-Axis Labels
        categories = ["Medical Coverage & a Physician", 
                      "Physician Only", 
                      "Medical Coverage Only", 
                      "Have Neither"
                      ]
        
        # Define colors for each data set
        color1 = 'blue'
        color2 = 'red'
        
        # Set the width of the bars
        bar_width = 0.2
        x = range(len(categories))
        
        # Generate graph
        fig, ax = plt.subplots()
        ax.bar(x, diabetes[0:4], width = bar_width, label = "Diabetes", color = color1)
        ax.bar([i + bar_width for i in x], bloodPressure[0:4], width=bar_width, label='Blood Pressure', color=color2)
            
        ax.set_ylabel("High Risk Patients")
        ax.set(xticks=[i + bar_width for i in x], xticklabels=categories)
        fig.legend()
        
        return fig
     
    def generateBarGraph(self, data, title):
        # Get Key values from dict to use as axis names
        categories = list(data.keys())
        counts = list(data.values())
        
        # Generate a graph
        fig, ax = plt.subplots()
        ax.set_ylabel(categories[-1])
        ax.bar(categories[0:4], counts[0:4])
        ax.set_title(title)
        
        # Set the x-axis range from 0 to the count of people who have diabetes
        diabetes_count = counts[-1]
        ax.set_ylim(-0.5, diabetes_count - 0.5)
    
        return fig
    
    # def generateTableData(self, data):
    
        
if __name__ == '__main__':
    file_path = "C:/Users/willi/Documents/Python/GraphExcel/grapher/FINAL 2023 Fiesta de Salud excel Report.xlsx"
    manager = ExcelManager(file_path)
    manager.generateComboGraph()