
#semeion.py

from numpy import *
from matplotlib.pyplot import *
import pandas as pd
from sklearn.svm import SVC

df = pd.read_csv('semeion.data', sep=' ', header=None).as_matrix()

X = df[:, 0:256]
y = df[:, 256:266].argmax(1)

index = arange(len(y))
random.shuffle(index)

N = 1000
train_index = index[:N]
test_index  = index[N:]

clf = SVC()
clf.fit(X[train_index], y[train_index])

print clf.score(X[test_index], y[test_index])