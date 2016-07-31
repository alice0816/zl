#! /user/bin/env python
#coding: utf-8
'''
Created on 2016/7/19

@author: fog
'''

import os
import sys
import zipfile
 
path = os.getcwd()
zipfilename = os.path.join(path, 'myzip.zip')

print zipfilename
 
# path is a direactory or not.
if not os.path.isdir(path):
    print path + ' No such a direactory'
    exit()
 
if os.path.exists(zipfilename):
    # zipfilename is exist.Append.
    print 'Add files into ' + zipfilename
    zipfp = zipfile.ZipFile(zipfilename, 'a' ,zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk(path, True):
        for filaname in filenames:
            direactory = os.path.join(dirpath,filaname)
            print 'Add... ' + direactory
            zipfp.write(direactory)
else:
    # zipfilename is not exist.Create.
    print 'Create new file ' + zipfilename
    zipfp = zipfile.ZipFile(zipfilename, 'w' ,zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk(path, True):
        for filaname in filenames:
            direactory = os.path.join(dirpath,filaname)
            print 'Compress... ' + direactory
            zipfp.write(direactory)
# Flush and Close zipfilename at last. 
zipfp.close()