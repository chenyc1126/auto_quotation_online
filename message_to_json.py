import json
import io
import typing

quotation = []
lines = ""
file1 = open("src/quotation.txt", "r",encoding='utf-8')

#print("hhih")
for line in file1.readlines():
    #print(line)
    if not ('分隔線' in line):
        lines += line
    else:
        buf = io.StringIO(lines)
        lines = ""
        key = []
        value = []

        for line in buf.readlines():  
            line = line.strip("\n")
            index = line.index('：') + 1 # spacing is weird in the data file
            key.append(line[:index-1])
            value.append(line[index:])

        quotation.append(dict(zip(key,value))) 
        continue

with open("quotation.json", "w") as f:
        f.write(json.dumps(quotation, indent=4, ensure_ascii=False))

file1.close()