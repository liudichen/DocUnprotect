# -*- coding:utf-8 -*-

##@柳涤尘2020-11-29
import os,time,shutil,sys
import zipfile
import tkinter
import windnd
from tkinter.messagebox import showinfo


def covert2docx(infile,outfile):
	from win32com import client
	wordapp=client.Dispatch('Word.Application')
	file=wordapp.Documents.Open(infile)
	# print(infile,outfile)
	file.SaveAs(outfile,12)
	file.Close()
	wordapp.Quit()
	return None

# def unprotectdocx(docxfile):
# 	tmppath=os.environ['TMP'].replace('\\','/')
# 	print(type(tmppath))
# 	print(tmppath)
# 	return None
def get_dragged_files_name(files):
	arr=[]
	for file in files:
		file=file.decode('gbk')
		if file.upper().endswith('.DOC') or file.upper().endswith('.DOCX') or file.upper().endswith('.WPS'):
			arr.append(file)
	# print(arr) 
	return arr
# 压缩docx 
def zip_docx(startdir, file_news):
    z = zipfile.ZipFile(file_news, 'w', zipfile.ZIP_DEFLATED)
    for dirpath, dirnames, filenames in os.walk(startdir):
        fpath = dirpath.replace(startdir, '')
        fpath = fpath and fpath + os.sep or ''
        for filename in filenames:
            z.write(os.path.join(dirpath, filename), fpath+filename)
    z.close()
def just_do_it(files):
	tmppath=os.environ['TMP'].replace('\\','/')
	filearr=get_dragged_files_name(files)
	
	resarr=[]
	# print(tmppath)
	if len(filearr)<1:
		showinfo('ERROR',"文件格式不匹配！")
		exit()
	else:
		for file in filearr:
			fpath=os.path.dirname(file).replace('\\','/')
			fname=file.split('\\')[-1]
			fext=fname.split('.')[-1]
			fname=fname[:-len(fext)-1]
			# print(fpath,fname,fext)
			fdocx=tmppath+'/'+fname+'.docx'
			zdocx=tmppath+'/iimm/解除保护de-'+fname+'.docx'
			zippath=tmppath+'/iimm/zip'
			setfpath=zippath+'/word/settings.xml'
			# print(file)
			try:
				print(fdocx)
				if os.path.exists(fdocx):
					os.remove(fdocx)
				covert2docx(file,fdocx)
				if os.path.exists(tmppath+'/iimm'):
					shutil.rmtree(tmppath+'/iimm')			
				
				z=zipfile.ZipFile(fdocx,'r')
				z.extractall(path=zippath)
				z.close()
				with open(setfpath,'r') as f:
					xml=f.read()
					startpos=xml.find('<w:documentProtection')
					endpos=xml[startpos:].find('/>')+startpos
					# print(startpos,endpos)
					if startpos!=-1:
						xml=xml[:startpos]+xml[endpos+1:]
					# strs=xml[startpos:endpos+2]
					# strs=strs+'x'
					# print(strs)
				with open(setfpath,'w') as f:
					f.write(xml)			
				zip_docx(zippath,zdocx)
				if os.path.exists(fpath+'/解除保护de-'+fname+'.docx'):
					os.remove(fpath+'/解除保护de-'+fname+'.docx')
				shutil.move(zdocx,fpath)
				if os.path.exists(fdocx):
					os.remove(fdocx)
				if os.path.exists(tmppath+'/iimm'):
					shutil.rmtree(tmppath+'/iimm')
				# print('xxx')
				resarr.append(fname)
			except:
				showinfo('ERROR','未能成功解除保护')
				if os.path.exists(fdocx):
					os.remove(fdocx)
				if os.path.exists(tmppath+'/iimm'):
					shutil.rmtree(tmppath+'/iimm')
				exit()
		if len(resarr)>0:
			showinfo('@柳涤尘','处理成功：\n'+'\n'.join(resarr))
			exit()
n_ts=time.time()
dead_time=4093456027  #2099-09-19
if n_ts>dead_time:	
	showinfo('超出使用期限：2099-09-19')
	exit()
else:
	# print(sys.version)
	tk=tkinter.Tk()
	tk.title('Word解除保护@柳涤尘')
	tk.geometry('300x150')
	tk.resizable(0,0)
	l=tkinter.Label(tk,text="将word文档拖拽到此窗口解除保护\n（python写的，运行较慢）")
	l.pack()
	#打包资源时，打包的资源路径需要这样获取：
	if getattr(sys, 'frozen', None):
		basedir = sys._MEIPASS
	else:
		basedir = os.path.dirname(__file__)
	imagepath=os.path.join(basedir, 'qflyg.gif')
	canvas=tkinter.Canvas(tk,height=150,width=300)
	image_file=tkinter.PhotoImage(file=imagepath)
	image=canvas.create_image(250,40,image=image_file)
	canvas.pack()
	windnd.hook_dropfiles(tk,func=just_do_it)
	tk.mainloop()