from numpy import *
import pandas as pd
from sklearn.svm import SVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import Imputer

df = pd.read_csv('http://archive.ics.uci.edu/ml/machine-learning-databases/glass/glass.data', header=None, index_col=0)

index = arange(0, 214)
random.shuffle(index)

N = 114
train_index = index[:N]
test_index  = index[N:]

X = df.ix[:, 1:9]
y = df[10]

imp = Imputer(strategy='mean', axis=0)
X = imp.fit_transform(X)

print 'Training..'
#clf = SVC()
clf = RandomForestClassifier()
clf.fit(X[train_index, :], y[train_index])

print 'Test'
print clf.predict(X[test_index, :])
print y[test_index]
