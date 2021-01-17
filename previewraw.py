#preview the raw data scrapped from the Government Website wriiten in Chinese
#to resolve the BOM(Byte Order Mark) issues in UTF-8

import codecs
f= codecs.open(u'法拍.csv','r','utf-8')
print(f.read())
f.close()

