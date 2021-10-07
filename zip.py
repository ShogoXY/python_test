import os, zipfile
newZip=zipfile.ZipFile('new.zip', 'w')
newZip.write('test.py', compress_type=zipfile.ZIP_DEFLATED)
newZip.close()


#test = 'C:\\Users\\Dariusz\\Desktop\\Śmieci\\python'
#for folderName, subfolders, filenames in os.walk(test):
#    print('Katalog ' + folderName)
#for subfolder in subfolders:
#        print('')
#        print ('podkatalog ' + folderName + ': ' + subfolder)
#
#    for filename in filenames:
#        print ('plik ' + folderName + ': ' + filename)
#        #print ('')
