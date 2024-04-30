import numpy as np
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
print(arr1[:,1])