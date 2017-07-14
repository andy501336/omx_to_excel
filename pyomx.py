import openmatrix as omx
import pandas as pd
import numpy as np
import operator

def pyomx(f):
#	f='ODTravDistPass.omx'
	f=omx.open_file(f)
	m=f.list_matrices()
	l = []
	labels = f.mapping(f.listMappings()[0])
	labels= sorted(labels.items(), key=operator.itemgetter(0))
	labels = [i[0] for i in labels]
	labels_y = [0]+labels
	for i in m:
		arr = f[i].read()
		arr=np.insert(arr,0,labels,axis=0)
		arr=np.insert(arr,0,labels_y,axis=1)
		l.append(arr)
	for k,i in enumerate(l):
		print("%s / %s" %(k+1,len(l)-1))
		df = pd.DataFrame(i)
		df.to_excel(m[k]+'.xlsx', index=False)



if __name__='__main__':
	file_name = input_raw('Give the name of omx file:\n')
	pyomx(file_name)
