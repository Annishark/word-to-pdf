import os, sys
from win32com.client import Dispatch, constants, gencache 
    
def wordtopdf(word,pdf):#主体转换程序
    gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
    # 开始转换
    w = Dispatch("Word.Application")
    try:
        doc = w.Documents.Open(word, ReadOnly=1)
        doc.ExportAsFixedFormat(pdf, constants.wdExportFormatPDF, \
           Item=constants.wdExportDocumentWithMarkup,
           CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    except:
        print('exception')
    finally:
        w.Quit(constants.wdDoNotSaveChanges)
        if os.path.isfile(pdf):
            print ('转换完成')
        else:
            print ('转换失败')

if __name__ == '__main__':
    roottee = os.getcwd()
    for root, dirs, files in os.walk(roottee):
        for name in files:
            word = os.path.join(root, name)
            file_name,suffix=word.split('.')
            if suffix == 'doc' or suffix == 'docx':  
                print(word)
                print("正在转换，请稍等。。。")
                pdf = file_name+'.pdf'
                wordtopdf(word,pdf)
        
    print("转换完成，请退出！")