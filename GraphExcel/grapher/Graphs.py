import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class ExcelManager():
    def readExcelFile(self,file_path, columnsToRead):
        try:
            data = pd.read_excel(file_path, usecols=columnsToRead)
        except FileNotFoundError:
            print("File Not Found")

        return data

    def analyzeBloodPressure(self, file_path):
        # Get required data from excel file
        columns_to_read = ["Have you ever been diagnosed with high or low blood pressure", "Do you have a primary care physician", "Do you have Medical Coverage"]
        data = self.readExcelFile(file_path, columns_to_read)
        
        # Analyze the retrieved data
        conditionColumns = ['Do you have a primary care physician', 'Do you have Medical Coverage', 'Are you Diabetic']        
        combinationsToCheck = [
            ['Yes, I do have a primary care physician', 'Yes, I do have medical coverage', 'Yes, I have been previously diagnosed with High or Low Blood Pressure'],
            ['Yes, I do have a primary care physician', 'No, I do not have medical coverage', 'Yes, I have been previously diagnosed with High or Low Blood Pressure'],
            ['No, I do not have a primary care physician', 'Yes, I do have medical coverage', 'Yes, I have been previously diagnosed with High or Low Blood Pressure'],
            ['No, I do not have a primary care physician', 'No, I do not have medical coverage', 'Yes, I have been previously diagnosed with High or Low Blood Pressure']
        ]
        
        # Count entries for each combination
        count = []
        for combination in combinationsToCheck:
            count.append(self.countData(data, conditionColumns, combination))
                        
        dataCount = {
            "Medical Coverage & a Physician": count[0],
            "Physician Only": count[1],
            "Medical Coverage Only": count[2],
            "Have Neither": count[3],
            "People who have diabetes": len(data[data['Have you ever been diagnosed with high or low blood pressure'] == 'Yes, I have been previously diagnosed with High or Low Blood Pressure'])
        }
        
        return self.generateBarGraph(data, "Blood Pressure Data")
        
    def analyzeDiabetes(self, file_path):
        # Get required data from excel file
        columns_to_read = ["Glucose Levels", "Do you have a primary care physician", "Do you have Medical Coverage"]
        data = self.readExcelFile(file_path, columns_to_read)
        
        # # analyze the retrieved data
        conditionColumns = ['Do you have a primary care physician', 'Do you have Medical Coverage', 'Glucose Levels']        
        combinationsToCheck = [
            ['Yes, I do have a primary care physician', 'Yes, I do have medical coverage', 'Yes, I am a Diabetic'],
            ['Yes, I do have a primary care physician', 'No, I do not have medical coverage', 'Yes, I am a Diabetic'],
            ['No, I do not have a primary care physician', 'Yes, I do have medical coverage', 'Yes, I am a Diabetic'],
            ['No, I do not have a primary care physician', 'No, I do not have medical coverage', 'Yes, I am a Diabetic']
        ]
        
        # Count entries for each combination
        count = []
        for combination in combinationsToCheck:
            count.append(self.countData(data, conditionColumns, combination))
                        
        dataCount = {
            "Medical Coverage & a Physician": count[0],
            "Physician Only": count[1],
            "Medical Coverage Only": count[2],
            "Have Neither": count[3],
            "People who have diabetes": len(data[data['Are you Diabetic'] == 'Yes, I am a Diabetic'])
        }
               
        return self.generateBarGraph(dataCount, "Diabetes Data")
          
    def analyzeCholesterol(self, file_path):
        data = {'Category': ['X', 'Y', 'Z'], 'People with Cholesterol': [15, 25, 35]}
        return self.generateBarGraph(data, "Cholesterol Data")
        
    def countData(self, data, conditionColumns, conditionValues):        
        if len(conditionColumns) != len(conditionValues):
            raise ValueError("The number of condition columns must match the number of condition values.")

        condition = True
        for col, val in zip(conditionColumns, conditionValues):
            condition &= (data[col] == val)

        count = len(data[condition])
        
        return count
        
    def generateBarGraph(self, data, title):
        # Get Key values from dict to use as axis names
        categories = list(data.keys())
        counts = list(data.values())
        
        # Dummy function to generate a cholesterol graph
        fig, ax = plt.subplots()
        ax.set_ylabel(categories[-1])
        ax.bar(categories[0:4], counts[0:4])
        ax.set_title(title)
        
        # Set the x-axis range from 0 to the count of people who have diabetes
        diabetes_count = counts[-1]
        ax.set_ylim(-0.5, diabetes_count - 0.5)
    
        return fig
    
if __name__ == '__main__':
    file_path = "C:/Users/willi/OneDrive/Documents/Cathy_Project/2023 Fiesta de Salud excel Report.xlsx"
    manager = ExcelManager()
    manager.analyzeDiabetes(file_path)