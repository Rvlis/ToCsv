"""
将给定目录下的文件的内容提取出来,保存到csv文件
jpg/jpeg/png/bmp/pdf/doc/docx/html

python ToCsv.py dir_path [--save xx/xx/xx.csv] [--mod 1]
instance.to_csv()实现转化
"""
# conda install -c conda-forge tesserocr
# pip install pdfminer.six
# pip install docx
# pip install pypiwin32
# pip install bs4
# pip install tencentcloud-sdk-python

from tesserocr import PyTessBaseAPI
from pdfminer.high_level import extract_text
from win32com import client as wc
from bs4 import BeautifulSoup
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models
import base64
import os
import csv
import time
import docx
import argparse
# import codecs

class ToCsv(object):
    """将给定目录下的文件的内容提取出来,保存到csv文件

    instance.to_csv()实现转化
    python xxx.py dir xx/xx/xx [--save xx/xx/xx.csv] [--mod 1] 
    """

    def __init__(self, dirname, csv_save_path="csv/a.csv", ocr_mod=1):
        self.dirname = dirname
        self.ocr_mod = ocr_mod
        self.csv_save_path = csv_save_path


    def to_csv(self):
        """给定目录dirname,将其中的文件的内容保存到csv文件

        Args:
            dirname: 给定目录
            ocr_mod: ocr模式
            csv_save_path: 保存csv文件路径,默认为当前目录下./csv/a.csv
        
        Returns: None
        """
        imgs_path, pdfs_path, docs_path, docxs_path, htmls_path = self.doc_type_class(self.dirname)
        # print(imgs_path)
        # print(pdfs_path)
        # print(docxs_path)
        # print(htmls_path)

        # 不同类型文件转换后的txt可以选择不同目录存放
        img_txt_path = self.img_to_txt(imgs_path, self.ocr_mod)
        pdf_txt_path = self.pdf_to_txt(pdfs_path)
        docx_txt_path = self.docx_to_txt(docxs_path)
        html_txt_path = self.html_to_txt(htmls_path)

        # 这里选择存放到同一目录
        txt_dir = img_txt_path
        self.txt_to_csv(txt_dir, self.csv_save_path)


    def doc_type_class(self, dirname):
        """对给定文件夹下的文件按类型分类 png/jpg/jpeg、pdf、doc、docx、html

        Args:
            dir_name: 目录路径，含要处理的文件
        
        Returns:
            [imgs_path, pdfs_path, docs_path, docxs_path, htmls_path], 每个元素类型是list,存放各自类型文件的路径
        """
        file_types = ["png","jpg","jpeg","pdf","doc","docx","html"]

        imgs_path = []
        pdfs_path = []
        docs_path = []
        docxs_path = []
        htmls_path = []

        for f in os.listdir(dirname):

            # 目录
            if os.path.isdir(os.sep.join([dirname,f])):
                # 逐层目录遍历
                _imgs_path, _pdfs_path, _docs_path, _docxs_path, _htmls_path = self.doc_type_class(os.sep.join([dirname,f]))
                imgs_path.extend(_imgs_path)
                pdfs_path.extend(_pdfs_path)
                docs_path.extend(_docs_path)
                docxs_path.extend(_docxs_path)
                htmls_path.extend(_htmls_path)
            # 文件
            else:
                file_type = os.path.splitext(f)[-1][1:]
                
                if not file_type in file_types:
                    continue
                elif file_type in file_types[0:3]:
                    img_path = os.sep.join([dirname,f])
                    imgs_path.append(img_path)
                elif file_type == "pdf":
                    pdf_path = os.sep.join([dirname,f])
                    pdfs_path.append(pdf_path)
                elif file_type == "doc":
                    doc_path = os.sep.join([dirname,f])
                    docs_path.append(doc_path)
                elif file_type == "docx":
                    docx_path = os.sep.join([dirname,f])
                    docxs_path.append(docx_path)
                elif file_type == "html":
                    html_path = os.sep.join([dirname,f])
                    htmls_path.append(html_path)
        
        # 将doc另存为docx类型，一块处理
        if len(docs_path):
            # 前处理,解决修改类型后可能出现的命名冲突
            for doc_path in docs_path:
                tmp_docx_path = os.path.splitext(doc_path)[0] + ".docx"
                # 检查是否和docxs_path存在命名冲突
                while(docxs_path.count(tmp_docx_path)):
                    tmp_docx_path = self.eliminate_dup_name(tmp_docx_path,dirname)
                # 解决命名冲突后要将新文件名重新存放会docs_path
                new_doc_path = os.path.splitext(tmp_docx_path)[0] + ".doc"
                # 从docs_path删除原文件名
                docs_path.remove(doc_path)
                # 本地文件重命名
                os.rename(doc_path, new_doc_path)
                # 将新文件名追加回docs_path
                docs_path.append(new_doc_path)
            # doc另存为docx
            extra_docxs_path = self.doc_save_as_docx(docs_path,dirname)
            docxs_path.extend(extra_docxs_path)

            # extra_docxs_path = doc_save_as_docx(docs_path, dirname)
            # for docx_path in extra_docxs_path:
            #     while docxs_path.count(docx_path):
            #         docx_path = eliminate_dup_name(docx_path, dirname)
            #     docxs_path.append(docx_path)
            
        return [imgs_path, pdfs_path, docs_path, docxs_path, htmls_path]
                

    def generate_save_path(self, files_path, save_dir="txts"):
        """给定文件路径，得到 [原文件路径，转化的txt文件要保存的路径]

        Args:
            files_path: list, 原文件路径
            save_dir: 保存txt的目录，默认为./txts/
        
        Returns:
            [[file_path, txt_save_path],...]
        """
        if not os.path.exists(save_dir):
            os.mkdir(save_dir)

        file_and_save_path = []

        for file_path in files_path:
            # 原文件目录
            sf = os.path.split(file_path)[1]
            # 原文件名
            file_name = os.path.splitext(sf)[0]
            # 保存txt的路径
            txt_save_path = os.sep.join([save_dir, file_name+".txt"])

            # if os.path.exists(txt_save_path):
            #     txt_save_path = eliminate_dup_name(txt_save_path,save_dir)

            # 循环消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)

            file_and_save_path.append([file_path, txt_save_path])
        
        return file_and_save_path


    # 弃用
    def load_imgs(self, imgs_path):
        """从图片文件夹中读取图片并返回每一个图片的地址和识别成txt后要存放的地址
        
        Args:
            imgs_path:
        
        Returns:
            img_and_txt_paths: [[img_path, txt_save_path],...]
        """
        img_and_txt_paths = []
        for path in os.listdir(imgs_path):
            img_path = os.path.join(imgs_path, path)
            img_name = os.path.splitext(path)[0]

            save_dir = "txts"
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)

            txt_save_path = os.sep.join([save_dir,img_name+".txt"])
            
            img_and_txt_paths.append([img_path, txt_save_path])
        return img_and_txt_paths


    def img_to_txt(self, imgs_path, choice=1):
        """图片转txt

        Args:
            imgs_path: list,图片路径
            choice: OCR模式 0:tesseract 1:tencentcloud

        Returns:
            txt_dirname: 存放txt的目录名
        
        """
        img_and_txt_paths = self.generate_save_path(imgs_path)
        
        
        if choice == 0:     # tesseract
            # 存放txt目录路径
            txt_dirname = self.tesseract(img_and_txt_paths)
        else:       #tencentcloud 
            txt_dirname = self.tencentcloud(img_and_txt_paths)

        return txt_dirname


    def tesseract(self, img_and_txt_paths):
        """tesseract
        
        Args:
            img_and_txt_paths: [[img_path,txt_save_path],...]
        
        Returns:
            txt_dirname: 存放txt的目录名
        """

        for path in img_and_txt_paths:
            img_path = path[0]
            txt_save_path = path[1]
            save_dir = os.path.split(txt_save_path)[0]
            # 循环消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)
            
            with PyTessBaseAPI() as api:
                api.Init(lang="chi_sim+eng+osd")
                api.SetImageFile(img_path)
                # content = ((api.GetUTF8Text()).encode("gbk","ignore")).decode("gbk","ignore")
                content = api.GetUTF8Text()
                with open(txt_save_path, "wt", encoding="gbk") as wf:
                    wf.write(content)

        return os.path.dirname(txt_save_path)

    def tencentcloud(self, img_and_txt_paths):
        """tencentcloud

        Args:
            img_and_txt_paths:
        
        Returns:
            txt_dirname:
        """
        for path in img_and_txt_paths:
            img_path = path[0]
            txt_save_path = path[1]
            save_dir = os.path.split(txt_save_path)[0]
            # 循环消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)
            try:
                cred = credential.Credential(
                    os.environ.get("TENCENTCLOUD_SECRET_ID"),
                    os.environ.get("TENCENTCLOUD_SECRET_KEY"))
                httpProfile = HttpProfile()
                httpProfile.endpoint = "ocr.tencentcloudapi.com"
                
                clientProfile = ClientProfile()
                clientProfile.httpProfile = httpProfile
                client = ocr_client.OcrClient(cred, "ap-shanghai", clientProfile)

                # 识别模型
                req = models.GeneralAccurateOCRRequest()
                
                # 将图片转化为base64编码格式
                with open(img_path, "rb") as rf:
                    base64_data = base64.b64encode(rf.read())
                    req.ImageBase64 = str(base64_data, 'utf-8')

                resp = client.GeneralAccurateOCR(req)
                
                results = resp.TextDetections

                if results != None:
                    with open(txt_save_path, "wt", encoding="gbk") as wf:
                        for line in results:
                            wf.write(str(line.DetectedText))
            except TencentCloudSDKException as err:
                print(err.get_code(),err.get_message())
        
        return os.path.dirname(txt_save_path)


    def pdf_to_txt(self, pdfs_path):
        """pdf转txt

        Args:
            pdf_and_txt_paths:
        
        Return:
            txt_dirname: 

        """
        pdf_and_txt_paths = self.generate_save_path(pdfs_path)

        for path in pdf_and_txt_paths:
            pdf_path = path[0]
            txt_save_path = path[1]
            save_dir = os.path.split(txt_save_path)[0]
            # 循环消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)

            # 提取pdf内容
            text = (extract_text(pdf_path).encode("gbk","ignore")).decode("gbk","ignore")
            with open(txt_save_path, "wt", encoding="gbk") as wf:
                wf.write(text)
        
        txt_dirname = os.path.dirname(txt_save_path)

        return txt_dirname


    def docx_to_txt(self, docxs_path):
        """docx转txt

        Args:
            docx_and_txt_paths:

        Returns:
            txt_dirname:
        """
        docx_and_txt_paths = self.generate_save_path(docxs_path)

        for path in docx_and_txt_paths:
            docx_path = path[0]
            txt_save_path = path[1]
            save_dir = os.path.split(txt_save_path)[0]
            # 循环消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)
            
            # text格式是list,docx库安装标点符号将原文拆分后装到链表中
            text = docx.getdocumenttext(docx.opendocx(docx_path))
            with open(txt_save_path, "wt", encoding="gbk") as wf:
                for line in text:
                    line = str((line.encode("gbk","ignore")).decode("gbk","ignore"))
                    wf.write(line)

        txt_dirname = os.path.dirname(txt_save_path)

        return txt_dirname

    def html_to_txt(self, htmls_path):
        """Html转txt

        Args:
            htmls_path:
        
        Returns:
            txt_dirname:
        """
        html_and_txt_paths = self.generate_save_path(htmls_path)

        for path in html_and_txt_paths:
            html_path = path[0]
            txt_save_path = path[1]
            save_dir = os.path.split(txt_save_path)[0]
            # 循坏消除重名txt文件
            while(os.path.exists(txt_save_path)):
                txt_save_path = self.eliminate_dup_name(txt_save_path,save_dir)

            with open(html_path,"rb") as rf:
                text = ((BeautifulSoup(rf, features="lxml").get_text()).encode("gbk","ignore")).decode("gbk","ignore")
                with open(txt_save_path,"wt",encoding="gbk") as wf:
                    wf.write(text.replace("\n",""))

        txt_dirname = os.path.dirname(txt_save_path)

        return txt_dirname


    def remove_chas(self, txt_dirname):
        """去除txt文档中空格、回车和tab字符

        Args:
            txt_dirname: 存放txt文件的目录名 
        
        Returns: None
        """

        for txt_name in os.listdir(txt_dirname):
                txt_path = os.sep.join([txt_dirname,txt_name])
                # 定义一个临时文件，将原文件修改后的内容存放到临时文件中
                with open(txt_path, "rt", encoding="gbk") as of, open(txt_path+".swp","wt",encoding="gbk") as nf:
                    for line in of:
                        res = line.replace("\n","")
                        # res = res.replace(" ","")
                        res = res.replace("\t","")
                        nf.write(res)
                # 删除原文件，并将临时文件更名为原文件名
                os.remove(txt_path)
                os.renames(txt_path+".swp", txt_path)


    def txt_to_csv(self, txt_dirname, csv_save_path="csv/a.csv"):
        """txt文档逐行存放到csv
        
        Args:
            txt_dir_name:
            csv_save_dir: csv保存路径,默认为当前路径下./csv/a.csv

        Returns:None

        """
        # 对txt文件进行预处理,删除字符
        self.remove_chas(txt_dirname)

        # 保存csv文件目录名
        csv_save_dir = os.path.split(csv_save_path)[0]
        # csv文件名
        csv_save_name = os.path.split(csv_save_path)[1]

        if not os.path.exists(csv_save_dir):
            os.makedirs(csv_save_dir)
        
        # if not csv_save_name:
        #     csv_save_name = "a.csv"

        csv_path = os.sep.join([csv_save_dir,csv_save_name])

        # for txt_name in os.listdir(txt_dirname):
        #     txt_path = os.sep.join([txt_dirname,txt_name])
        #     with open(csv_path, "a", newline="") as csvf:
        #         spamwriter = csv.writer(csvf, dialect="excel")
        #         with open(txt_path, "r", encoding="gbk") as rf:
        #             spamwriter.writerow(rf)

        with open(csv_path, "a", newline="") as csvf:
            spamwriter = csv.writer(csvf, dialect="excel")
            for txt_name in os.listdir(txt_dirname):
                txt_path = os.sep.join([txt_dirname,txt_name])
                # 跳过短文本
                txt_size = os.path.getsize(txt_path)
                if txt_size < 50:
                    continue
                else:
                    with open(txt_path, "r", encoding="gbk") as rf:
                        spamwriter.writerow(rf)


    def eliminate_dup_name(self, save_path, save_dir):
        """消除命名冲突文件

        Args:
            save_dir: 存在命名冲突的目录
            save_path: 要修改命名的文件的路径

        Return:
            save_path: 修改命名后的文件的路径
        """

        file_name = os.path.splitext(os.path.split(save_path)[1])[0]
        file_type = os.path.splitext(os.path.split(save_path)[1])[1]
        # 在原重名文件基础上添加'#'
        file_name += "#"

        save_path = os.sep.join([save_dir, file_name+file_type])
        
        return save_path


    def doc_save_as_docx(self, docs_path, dirname):
        """将给定路径下的doc文件转为docx

        Args:
            docs_path: list,要转换的doc文件路径
        
        Returns:
            docxs_path: list,转换后docx路径
        """

        word = wc.Dispatch("Word.Application")
        docxs_path = []

        for path in docs_path:
            # win32com,使用绝对路径
            abs_doc_path = os.path.abspath(path)
            file_name = os.path.splitext(abs_doc_path)[0]
            docx_path = file_name + ".docx"

            doc = word.Documents.Open(abs_doc_path)
            # 12对应要保存的文件类型docx
            doc.SaveAs(docx_path,12)
            doc.Close()

            docx_path = os.sep.join([dirname,os.path.split(docx_path)[1]])
            docxs_path.append(docx_path)

            # 删除原doc文件
            # os.remove(abs_doc_path)
            
        return docxs_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("dir", help="xx/xxx The directory containing the file will be converted.")
    parser.add_argument("-s", "--save", default="csv/a.csv", help="xx/xxx/x.csv The path the csv file will be saved at.")
    parser.add_argument("-m", "--mod", type=int, default=1, choices=[0,1], help="The Ocr model: 0:tesseract,1:tencentcloud")
    args = parser.parse_args()

    start = time.time()

    # 定义ToCsv对象
    demo = ToCsv(args.dir, args.save, args.mod)
    # 调用to_csv方法
    demo.to_csv()
    
    print("time:",time.time()-start)