# encoding:utf-8
import xlwt
import os


def save_data(savepath, filepath):
    allfile = os.listdir(filepath)  # 要統計的資料夾
    d = []
    for file in allfile:
        if 'cluster' == file[-7:]:
            d.append(file)
    # 創建workbook對象
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    for i in range(len(d)):
        print("第%d檔案: " % (i + 1) + d[i])
        # 讀取資料內容有幾筆行數量
        with open(filepath + "/" + d[i], 'r') as f:
            rows = f.readlines()
            print('總共有' + str(len(rows)) + '筆資料')
        with open(filepath + "/" + d[i], 'r') as ff:
            # database = ff.read().split(',')
            database = []
            for x in range(len(rows)):
                ddt = ff.readline().split(',')
                database.append(ddt)
        # print(database[0])
        print("開始執行save.....")
        # # 創建workbook對象, 單一回圈內
        # book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        # 建立sheet頁面名稱(可覆蓋)
        sheet = book.add_sheet(d[i], cell_overwrite_ok=True)
        # 設立第一行固定名稱
        col = ("Label", "frame_num", "cluster_size", "cen_x",
               "cen_y", "cen_z", "cent_diff_x", "cent_diff_y",
               "cent_diff_z", "max_x", "max_y", "max_z", "min_x",
               "min_y", "min_z", "ran_x", "ran_y", "ran_z", "std_x",
               "std_y", "std_z", "max_doppler", "time")
        for i in range(0, len(col)):
            sheet.write(0, i, col[i])
        # 帶入讀取資料放入資後行列
        for i in range(0, len(rows)):
            # print("第%d條" % (i + 1))
            data = database[i]
            for j in range(0, len(col)):
                sheet.write(i + 1, j, data[j])
                # print(data[j])
    book.save(savepath + filename + ".xls")


filelistpath = './lying/'

filelist = os.listdir(filelistpath)
for filename in filelist:
    # print("list", list)
    print(filelistpath + filename)
    savepath = './'
    filepath = filelistpath + filename
    save_data(savepath=savepath, filepath=filepath)
    print("save完成")
