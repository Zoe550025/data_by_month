import os
import xlsxwriter
rootdir = 'E:/tzuwen/tree_segmentation2/data_by_month/2'                                             #TODO:計算數量的camera source
wb = xlsxwriter.Workbook("E:/tzuwen/tree_segmentation2/data_by_month/data統計_20220623_2.xlsx")      #TODO:要存的excel名稱
camera_source = str(2) + "-"                                                                        #TODO:2是camera的來源

ws2021 = wb.add_worksheet("2021")
ws2022 = wb.add_worksheet("2022")

bold = wb.add_format({
        'bold':  True,  # 字型加粗
        'border': 1,  # 單元格邊框寬度
        'align': 'center',  # 水平對齊方式
        'valign': 'vcenter',  # 垂直對齊方式
        'fg_color': '#F4B084',  # 單元格背景顏色
        'text_wrap': True,  # 是否自動換行
    })

num_camera = len(os.listdir(rootdir))
camera_dir = []
for i in range(num_camera):
    dirname = camera_source + str(i)
    camera_dir.append(dirname)

ws2021.write_row('B1',[1,2,3,4,5,6,7,8,9,10,11,12],bold)
ws2021.write_column('A2',camera_dir,bold)
ws2022.write_row('B1',[1,2,3,4,5,6,7,8,9,10,11,12],bold)
ws2022.write_column('A2',camera_dir,bold)

for root, subFolders, files in os.walk(rootdir):
    if len(files)!= 0:
        no_dir_root = root.find(camera_source)
        tmp = root[no_dir_root:]        #0-0\2022\1
        dir,year,month = tmp.split("\\")
        dir_index = int(dir.replace(camera_source,""))
        month_index = int(month)
        #print(dir_index,month_index)
        num_pic = len(files)
        if year == "2021":
            print(dir_index,month_index,num_pic)
            ws2021.write(dir_index+1,month_index,num_pic)
        else:
            print(dir_index, month_index, num_pic)
            ws2022.write(dir_index+1,month_index,num_pic)

wb.close()