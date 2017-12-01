import xlrd
import xlwt
import math

def loadData(news,l1,p1,c1,e1,r1,l2,p2,c2,e2,r2):
	wb = xlrd.open_workbook("Dataset Tugas 3 AI 1718.xlsx")
	ws1 = wb.sheet_by_index(0)
	ws2 = wb.sheet_by_index(1)

	for i in range(1,ws1.nrows):
		l1.append(ws1.cell_value(i,1))
		p1.append(ws1.cell_value(i,2))
		c1.append(ws1.cell_value(i,3))
		e1.append(ws1.cell_value(i,4))
		r1.append(ws1.cell_value(i,5))

	for i in range(1,ws2.nrows):
		news.append(ws2.cell_value(i,0))
		l2.append(ws2.cell_value(i,1))
		p2.append(ws2.cell_value(i,2))
		c2.append(ws2.cell_value(i,3))
		e2.append(ws2.cell_value(i,4))
		r2.append(ws2.cell_value(i,5))

	return news,l1,p1,c1,e1,r1,l2,p2,c2,e2,r2

def distance(like1,like2,prov1,prov2,coms1,coms2,emo1,emo2):
	root = (like1-like2)**2 + (prov1-prov2)**2 + (coms1-coms2)**2 + (emo1-emo2)**2
	return math.sqrt(root)

def save(news,result):
	book = xlwt.Workbook()
	ws = book.add_sheet("Hasil(1301154298)")
	ws.write(0, 0,"Berita")
	ws.write(0, 1,"Keterangan")
	for i in range(len(result)):
			ws.write(i+1, 0,news[i])
			ws.write(i+1, 1,result[i])

	book.save('hasil.xls')

def main():
	#l1 =like pada sheet 1
	#p1 =provokasi pada sheet 1
	#c1 =comment pada sheet 1
	#e1 =emosi pada sheet 1
	#r1 =result pada sheet 1
	news,l1,p1,c1,e1,r1,l2,p2,c2,e2,r2 = [],[],[],[],[],[],[],[],[],[],[]
	news,l1,p1,c1,e1,r1,l2,p2,c2,e2,r2 = loadData(news,l1,p1,c1,e1,r1,l2,p2,c2,e2,r2)

	for i in range(len(r2)):
		tmp= []
		hox,no = 0,0

		for ix in range(len(r1)):
			a=(distance(l1[ix],l2[i],p1[ix],p2[i],c1[ix],c2[i],e1[ix],e2[i]))
			tmp.append(a)
		
		best = sorted(tmp)[0:17]

		for isi in best:
			if r1[tmp.index(isi)]==1.0:
				hox+=1
			else:
				no+=1

		if hox>no:
			r2[i]="Hoax"
		else:
			r2[i]="Not Hoax"

	save(news,r2)

main()