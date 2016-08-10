# coding=utf-8
import pandas as pd


df = pd.DataFrame({'a': [-1,2,3,4,5,6,7], u'עדי': [8,9,10,-11,12,13,-14], 'c': [15,16,17,18,19,20,21]})
cond = df[u'עדי'] < 0
print df[(cond)]
# df = df[(df.a >= 0)]
