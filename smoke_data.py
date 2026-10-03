import csv, sys
d = sys.argv[1]
with open(f"{d}/data.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["地区","产品","金额","数量"])
    w.writerow(["华东","A",100,5])
    w.writerow(["华东","B",200,3])
    w.writerow(["华北","A",150,4])
    w.writerow(["华南","B",250,6])
