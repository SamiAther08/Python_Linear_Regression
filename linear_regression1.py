import docx
import numpy as np
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt

def extract_data_from_word(file_path):
    doc = docx.Document(file_path)

    data = []
    
    for table in doc.tables:
        for row in table.rows:
            cols = [cell.text for cell in row.cells]
            if len(cols) == 2:  
                try:
                    x_value = float(cols[0])
                    y_value = float(cols[1])
                    data.append((x_value, y_value))
                except ValueError:
                    pass 

    return data

word_file_path = "F:\python\data.docx"

data = extract_data_from_word(word_file_path)

X = np.array([item[0] for item in data]).reshape(-1, 1)
Y = np.array([item[1] for item in data])

model = LinearRegression()
model.fit(X, Y)

slope = model.coef_[0]
intercept = model.intercept_
print("Slope (Coefficient):", slope)
print("Intercept:", intercept)

Y_pred = model.predict(X)

plt.scatter(X, Y, label='Data')
plt.plot(X, Y_pred, color='red', linewidth=2, label='Linear Regression')
plt.xlabel('X')
plt.ylabel('Y')
plt.legend()
plt.show()
