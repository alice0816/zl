
import os
import sys
import shutil
import StringIO
from subprocess import call

EXE_FILE = os.path.join(os.path.dirname(__file__), '7-Zip\\7z.exe')
OS_FILE = os.path.join(os.path.dirname(__file__), 'lib\\libz.so')


PASSWORD = '123456'

cur_path = os.getcwd()
basename = os.path.basename(cur_path)
archive_name = basename + '.zip'
out_file = basename + '.a'
full_archive_name = os.path.join(cur_path, archive_name)
a_file = os.path.join(cur_path, out_file)
length = len(sys.argv) 

if length != 2:
    raise Exception, 'The length of argvs is improper.'

elif sys.argv[1]=='-a':
    zipcommand =EXE_FILE + ' a ' + '-y ' + '-sdel '+ '-p{0} '.format(PASSWORD)  + archive_name 
    if os.path.exists(full_archive_name):
        raise Exception, 'The zipfile already exists.'
    call(zipcommand)
    if not os.path.exists(full_archive_name):
        raise Exception, 'Compress the folder of {0} failed.'.format(basename)
    print 'Compress the folder successfully.'
                
    s = StringIO.StringIO()   
    sofile = open(OS_FILE, 'rb')
    zipfile = open(full_archive_name, 'rb')   
    try:
        s.write(sofile.read(4*1024))
    finally:
        sofile.close()
    try:
        s.write(zipfile.read())
    finally:
        zipfile.close()
    outfile = open(a_file, 'wb')
    outfile.write(s.getvalue())
    outfile.close()
    try:
        os.remove(full_archive_name)
    except Exception, e:
        print e
    
elif sys.argv[1]=='-x':
    s = StringIO.StringIO()   
    afile = open(a_file, 'rb')   
    try:
        s.write(afile.read())
    finally:
        afile.close() 
    outfile = open(full_archive_name, 'wb')
    outfile.write(s.getvalue()[4096:])
    outfile.close()

    if not os.path.exists(full_archive_name):
        raise Exception, 'The file of {0}  is not exists, pls check...'.format(archive_name)
    unzipcommand = EXE_FILE + ' x '  + '-y ' + full_archive_name + ' -p{0}'.format(PASSWORD) 
    call(unzipcommand)
    print 'Unccmpress the folder successfully.'
    try:
        os.remove(full_archive_name)
    except Exception, e:
        print e
    try:
        os.remove(a_file)
    except Exception, e:
        print e

else:
    print 'Usage: zz <command>' + '\n'
    print '<command>' + '\n'
    print '  -a : Compress and encrypt files.' + '\n'
    print '  -x : eXtract files.' + '\n'
    
    