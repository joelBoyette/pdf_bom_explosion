#https://stackoverflow.com/questions/10973362/python-logging-function-name-file-name-line-number-using-a-single-file
#https://realpython.com/python-logging/
#https://fangpenlin.com/posts/2012/08/26/good-logging-practice-in-python/

log_location=r'C:\Users\JBoyette.BRLEE\Documents\PANDAS\work_scripts' + '\\'
log_filemode='a'
log_format = " %(asctime)s: %(filename)s %(funcName)s() ln %(lineno)s [%(levelname)s] %(message)s"
log_datefmt='%d-%b-%y %H:%M:%S'
