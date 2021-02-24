# -*- coding: utf-8 -*-
"""
@author: Ganing

@Date:  2021/02/19
"""

import os
import logging
import sys
from reportlab.lib.pagesizes import A4, landscape,portrait 
from reportlab.platypus import SimpleDocTemplate,Image
from reportlab.pdfgen import canvas
import win32com.client
from PIL import Image as pilImage

logger = logging.getLogger('Sun')
logging.basicConfig(level=20,
                    # format="[%(name)s][%(levelname)s][%(asctime)s] %(message)s",
                    format="[%(levelname)s][%(asctime)s] %(message)s",
                    datefmt='%Y-%m-%d %H:%M:%S'  # 注意月份和天数不要搞乱了，这里的格式化符与time模块相同
                    )


def getFiles(dir, suffix, ifsubDir=True):  # 查找根目录，文件后缀
    res = []
    for root, directory, files in os.walk(dir):  # =>当前根,根下目录,目录下的文件
        for filename in files:
            name, suf = os.path.splitext(filename)  # =>文件名,文件后缀
            if suf.upper() == suffix.upper():
                res.append(os.path.join(root, filename))  # =>吧一串字符串组合成路径
        if False is ifsubDir:
            break
    return res


class pptTrans:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.filePath = filePath
        self.powerpoint = None

        self.init_powerpoint()
        self.convert_files_in_folder(self.filePath)
        self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype

        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFileName(infoDict['name'], inputFileName)

        if '' == outputFileName:
            return
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return
        if None is self.powerpoint:
            return
        powerpoint = self.powerpoint
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = powerpoint.Presentations.Open(inputFileName)

        try:
            deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
        deck.Close()

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            dirPath = filePath
            files = os.listdir(dirPath)
            pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        elif True is os.path.isfile(filePath):
            pptfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for pptfile in pptfiles:
            fullpath = os.path.join(filePath, pptfile)
            self.ppt_trans(fullpath)

    def getNewFileName(self, newType, filePath):
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]
            suffix = baseName.rsplit('.', 1)[1]
            if newType == suffix:
                logger.warning('文档[{filePath}]类型和需要转换的类型[{newType}]相同'.format(filePath=filePath, newType=newType))
                return ''
            newFileName = '{dir}/{fileName}.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            if os.path.exists(newFileName):
                newFileName = '{dir}/{fileName}_new.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''


class pngstoPdf:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.powerpoint = None

        self.init_powerpoint()
        self.convert_files_in_folder(filePath)
        # self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.ActivePresentation.Close()
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            # self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFolderName(inputFileName)

        if '' == outputFileName:
            return ''
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return ''
        if None is self.powerpoint:
            return ''
        powerpoint = self.powerpoint
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = powerpoint.Presentations.Open(inputFileName,ReadOnly=True,WithWindow=False)

        try:
            deck.SaveAs(outputFileName, formatType)
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
            return ''
        deck.Close()
        return outputFileName

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            dirPath = filePath
            files = os.listdir(dirPath)
            pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        elif True is os.path.isfile(filePath):
            pptfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for pptfile in pptfiles:
            fullpath = os.path.join(filePath, pptfile)
            folderName = self.ppt_trans(fullpath)
            try:
                self.png_to_pdf(folderName)
            except Exception as e:
                logger.error(str(e))
            for file in os.listdir(folderName):
                os.remove('{0}\\{1}'.format(folderName, file))
            os.rmdir(folderName)

    def png_to_pdf(self, folderName):
        picFiles = getFiles(folderName, '.png')
        pdfName = self.getFileName(folderName)

        '''多个图片合成一个pdf文件'''
        (w, h) = landscape(A4)  #
        cv = canvas.Canvas(pdfName, pagesize=landscape(A4))
        for imagePath in picFiles:
            cv.drawImage(imagePath, 0, 0, w, h)
            cv.showPage()
        cv.save()

    def getFileName(self, folderName):
        dirName = os.path.dirname(folderName)
        folder = os.path.basename(folderName)
        return '{0}\\{1}.pdf'.format(dirName, folder)

    def getNewFolderName(self, filePath):
        index = 0
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]

            newFileName = '{dir}/{fileName}'.format(dir=dirPath, fileName=fileName)
            while True:
                if os.path.exists(newFileName):
                    newFileName = '{dir}/{fileName}({index})'.format(dir=dirPath, fileName=fileName, index=index)
                    index = index + 1
                else:
                    break
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''

class pngs6toPdf:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.powerpoint = None

        self.init_powerpoint()
        self.convert_files_in_folder(filePath)
        # self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.powerpoint:
            self.powerpoint.ActivePresentation.Close()
            self.powerpoint.Quit()

    def init_powerpoint(self):
        try:
            self.powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            # self.powerpoint.Visible = 2
        except Exception as e:
            logger.error(str(e))

    def ppt_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFolderName(inputFileName)

        if '' == outputFileName:
            return ''
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return ''
        if None is self.powerpoint:
            return ''
        powerpoint = self.powerpoint
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = powerpoint.Presentations.Open(inputFileName,ReadOnly=True,WithWindow=False)

        try:
            deck.SaveAs(outputFileName, formatType)
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
            return ''
        deck.Close()
        return outputFileName

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            dirPath = filePath
            files = os.listdir(dirPath)
            pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
        elif True is os.path.isfile(filePath):
            pptfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for pptfile in pptfiles:
            fullpath = os.path.join(filePath, pptfile)
            folderName = self.ppt_trans(fullpath)
            try:
                self.png_to_pdf(folderName)
            except Exception as e:
                logger.error(str(e))
            for file in os.listdir(folderName):
                os.remove('{0}\\{1}'.format(folderName, file))
            os.rmdir(folderName)

    def png_to_pdf(self, folderName):
        logger.info("开始生成PDF文件...请稍候")
        picFiles = getFiles(folderName, '.png')
        pdfName = self.getFileName(folderName)

        '''多个图片合成一个pdf文件'''
        (w, h) = portrait(A4)  #
        w = w - 3
        h = h - 10
        # print(w,h)
        cv = canvas.Canvas(pdfName, pagesize=portrait(A4))
        i = 0
        page_num = 1
        for imagePath in picFiles:
            x1 = (i%2)*(w/2) +1*(i%2+1)
            y1 = (int((5-i)/2))*(h/3)+10
            __a4_w1 = w/2
            __a4_h1 = h/3
            # print(i,(i%2),(int((5-i)/2)),int(x1),int(y1),int(__a4_w1),int(__a4_h1))
            img_w, img_h = ImageTools().getImageSize(imagePath)

            # img_w = img.imageWidth
            # img_h = img.imageHeight

            if __a4_w1 / img_w < __a4_h1 / img_h:
                ratio = __a4_w1 / img_w
                y1 = y1 + (__a4_h1 - img_h * ratio)/2
            else:
                ratio = __a4_h1 / img_h
                x1 = x1 + (__a4_w1 - img_w * ratio)/2


            cv.drawImage(imagePath, x1,y1 ,width=img_w * ratio,height=img_h * ratio)
            i=(i+1)%6
            if i == 0:
                """
                Add the page number
                """
                page = "Page %s" % (page_num)
                cv.setFont("Helvetica", 9)
                cv.drawRightString(w-5, 5, page)
                cv.showPage()
                page_num = page_num +1
        if i != 0:
            """
            Add the page number
            """
            page = "Page %s" % (page_num)
            cv.setFont("Helvetica", 9)
            cv.drawRightString(w-5, 5, page)
        cv.save()
        logger.info("PDF输出完成")

    def getFileName(self, folderName):
        dirName = os.path.dirname(folderName)
        folder = os.path.basename(folderName)
        return '{0}\\{1}.pdf'.format(dirName, folder)

    def getNewFolderName(self, filePath):
        index = 0
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]

            newFileName = '{dir}/{fileName}'.format(dir=dirPath, fileName=fileName)
            while True:
                if os.path.exists(newFileName):
                    newFileName = '{dir}/{fileName}({index})'.format(dir=dirPath, fileName=fileName, index=index)
                    index = index + 1
                else:
                    break
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''

class ImageTools:
    def getImageSize(self, imagePath):
        img = pilImage.open(imagePath)
        return img.size

if __name__ == "__main__":
    if len(sys.argv) == 2:
        pngs6toPdf({'name': 'pdf(6页)', 'formatType': 18}, os.path.abspath(sys.argv[1]))
        exit(0)
    transDict = {}
    transDict.update({1: {'name': 'pptx', 'formatType': 11}})
    transDict.update({2: {'name': 'ppt', 'formatType': 1}})
    transDict.update({3: {'name': 'pdf', 'formatType': 32}})
    transDict.update({4: {'name': 'png', 'formatType': 18}})
    transDict.update({5: {'name': 'pdf(不可编辑)', 'formatType': 18}})
    transDict.update({6: {'name': 'pdf(6页)', 'formatType': 18}})

    hintStr = ''
    for key in transDict:
        hintStr = '{src}{key}:->{type}\n'.format(src=hintStr, key=key, type=transDict[key]['name'])

    while True:
        print(hintStr)
        transFerType = int(input("转换类型:"))
        if None is transDict.get(transFerType):
            logger.error('未知类型')
        else:
            infoDict = transDict[transFerType]
            path = input("文件路径:")
            if 5 == transFerType:
                pngstoPdf(infoDict, path)
            elif 6 == transFerType:
                pngs6toPdf(infoDict, path)
            else:
                op = pptTrans(infoDict, path)
