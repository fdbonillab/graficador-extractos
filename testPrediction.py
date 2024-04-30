import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
X, y = np.arange(10).reshape((5, 2)), range(5)
print('X ')
print(X)
print(' y ')
print(y)
print(' np.arange ')
print(np.arange(10))
print(' np reshape')
print(np.arange(10).reshape((5, 2)))
#print(np.arange(10).reshape((4, 3)))
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.33, random_state=42)

print('x train')
print(X_train)
print('y train')
print(y_train)
train_test_split(y, shuffle=False)
arr1 = [[1,2,3],[4,5,6],[7,8,9]]
column_i = np.take(arr1, 1, axis=1)
print(column_i)
#print(arr1[:,1])
arr1 = ['a','b','c','d','e','f','g']
arr2 = [4,5,6,7,8,9,10]
arr3 = [arr1,arr2]
#arr3 = np.concatenate((arr1, arr2))
print(arr3)
columns = ["fecha", "valor"]
rows = ["D", "E", "F"]
fechas = arr1
attrs = arr2
print(' len attrs '+str(len(attrs)))
matrix = [[fechas],[attrs]]
print(matrix)
#matrix = np.reshape(matrix, (-1,2))
#matrix.shape = (2,55)
print(matrix)
df = pd.DataFrame({'fecha': arr1, 'valor':arr2}) 
#df = pd.DataFrame(data=matrix, index=range(0,len(attrs)), columns=columns)
print(df)