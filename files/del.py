for dirpath, dirnames, filenames in os.walk(cur_path, True):
    for dirs in dirnames:
        try:
            shutil.rmtree(dirs,True) 
        except Exception, e:
            print e
    for filaname in filenames:
        if filaname == archive_name:
            continue
        try:
            os.remove(filaname)
        except Exception, e:
            print e