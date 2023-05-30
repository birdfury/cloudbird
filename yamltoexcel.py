#! _*_ coding:utf-8 _*_


import os
import yaml
import re
import xlwt
import xlwings as xw
class yml2Excel(object):
    def __init__(self):
        self.second_dir_list = []
        self.fullpath_yamlfile = []
        self.save_excel_list = []

    #获取子文件夹下的sheet文件名
    def print_all_path(self,init_file_path,keyword):

        for cur_dir, sub_dir, include_file in os.walk(init_file_path):
            for dirname in sub_dir:
            #保存二级目录‘h:\workspace\nuclei-template-main\****’
                self.second_dir_list.append(os.path.join(cur_dir,dirname))
            
            for file in include_file:
                if re.search(keyword, file):
                    self.fullpath_yamlfile.append(os.path.join(cur_dir,file))

    #统计Yaml文件，把文件信息保存到一行，作为execl的行数
    #根据子文件夹名字命名sheet文件名
    #根据子文件夹中的yaml文件来获取数据
    #获取yaml文件中的id name serivery word
    def open_yml_file(self):
        excelbook = xlwt.Workbook(encoding = 'utf-8')
        sheet = locals()
        sheettag_list = []
        id_data = []
        name_data = []
        severity_data = []
        requestdata = []
        typetag_data = []
        wordtag_data = []
        dsltag_data = []
        count = 0 
        
     #遍历子文件夹每个yml文件
        for sheetname in self.second_dir_list:
            dirnamelist=sheetname.split('\\')
            print(dirnamelist)
            if len(dirnamelist)>=4:
                sheettag = dirnamelist[3]
                if sheettag not in sheettag_list:
                    sheettag_list.append(sheettag)
                    print(sheettag_list)
                    sheet[sheettag] = excelbook.add_sheet(sheettag,cell_overwrite_ok=True)
                    sheet[sheettag].col(0).width = 30 * 256
                    sheet[sheettag].col(1).width = 70 * 256
                    sheet[sheettag].col(2).width = 15 * 256
                    sheet[sheettag].col(3).width = 90 * 256
                    sheet[sheettag].col(4).width = 130 * 256
                    alignment = xlwt.Alignment()
                    alignment.horz = xlwt.Alignment.HORZ_LEFT
                    alignment.vert = xlwt.Alignment.VERT_CENTER
                    style = xlwt.XFStyle()
                    style.alignment = alignment
                    font = xlwt.Font() # 为样式创建字体
                    font.name = 'Times New Roman' 
                    font.bold = True # 黑体
                    font.underline = True # 下划线
                    font.italic = True # 斜体字
                    style.font = font


                    sheet[sheettag].write(0,0,"id",style)
                    sheet[sheettag].write(0,1,"name",style)
                    sheet[sheettag].write(0,2,"severity",style)
                    sheet[sheettag].write(0,3,"word",style)
                    sheet[sheettag].write(0,4,"dsl",style)
                    print(sheettag_list)
        for filenameyaml in self.fullpath_yamlfile:
            print(filenameyaml)
            filenamelist= filenameyaml.split('\\')
            print(filenamelist)
            strfile = filenamelist[-1]
            fileextend = strfile.split('.')[-1]
            if fileextend == "yaml":
                print("this file is a yaml file")
                file_save_path = filenamelist[3]
                if file_save_path  in sheettag_list:
                    
                    print("filepathis:"+filenameyaml)
                    count = count+1
                    with open(filenameyaml,'r',encoding='utf-8') as f:
                        p = f.read()
                        ymlcfg = yaml.load(p,Loader=yaml.FullLoader)
                #获取yaml文件中的key保存到excel中的列中
                        for key in ymlcfg:
                            if key == 'id':
                                id_data=(ymlcfg[key]) 
                                print(id_data)
                                sheet[file_save_path].write(count,0,id_data)
                            if key == 'info':
                                name_data=(ymlcfg[key].get('name'))
                                print(name_data)
                                sheet[file_save_path].write(count,1,name_data)
                                severity_data=(ymlcfg[key].get('severity'))
                                print(severity_data)
                                sheet[file_save_path].write(count,2,severity_data)
                            if key == 'requests':
                                requestdata = ymlcfg[key][0].get('matchers')
                                print(ymlcfg[key][0])
                                print(requestdata)
                                if requestdata != None:
                                    for key1 in requestdata:
                                        for k,v in key1.items():
                                            print(f"{k}:{v}")
                                            if k == 'type':
                                                typetag_data= key1[k] 
                                                print(typetag_data)
                                                if typetag_data == 'word':
                                                    wordtag_data=key1['words'] 
                                                    sheet[file_save_path].write(count,3,wordtag_data)
                                                    print(wordtag_data)
                                                elif typetag_data == 'dsl':
                                                    dsltag_data = key1['dsl']
                                                    print(dsltag_data)
                                                    sheet[file_save_path].write(count,4,dsltag_data)
                                                else:
                                                    break

                            
                    f.close()
                    print(count)
        excelbook.save('nuclei-template.xls')

    def del_blankrow(self):
        xls_path=os.path.abspath('H:\\workspace\\nuclei-templates-main\\nuclei-template.xls')
        app = xw.App(visible=True,add_book=False)
        xt = app.books.open(xls_path)

        for shname in xt.sheets:
            rows= shname.used_range.last_cell.row
            print(shname.name)
            if rows == 1:
                shname.delete()
        xt.save()
        xt.close()
        wb = app.books.open(xls_path)
        count_rows = 0
        for sn in wb.sheets:
            rows= sn.used_range.last_cell.row
            print(rows)
            count_rows += rows
            for i in range(rows,1,-1):
                print(i)
                if(sn.range('A'+str(i)).value == None):
                    print("delete "+ str(i) )
                    sn.range('A'+str(i)).api.EntireRow.Delete()
                i=i+1
        print(count_rows)
        wb.save()
        wb.close()


if __name__== "__main__":
     yamrun = yml2Excel()
     yamrun.print_all_path("H:\\workspace\\nuclei-templates-main", "yaml")
     yamrun.open_yml_file()
     yamrun.del_blankrow()