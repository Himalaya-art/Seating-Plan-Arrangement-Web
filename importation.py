import os
import pandas as pd
import tabulate

class importation:
    def __init__(self, filePath):
        self.filePath =filePath
    
    def read(self):

        if os.path.exists(self.filePath):
            fileName = os.path.split(self.filePath)[1]
            if fileName.endswith((".xlsx", ".xls")):
                return self.excel()
            elif fileName.endswith(".csv"):
                return self.csv()
            elif fileName.endswith(".txt"):
                return self.txt()
            else:
                return ["不受支持的文件类型"]
        else:
            return ["文件不存在，请检查路径是否正确"]

    def excel(self):
        return pd.read_excel(self.filePath)

    def csv(self):
        return pd.read_csv(self.filePath)

    def txt(self):
        return pd.read_csv(self.filePath, sep=" ")
    
if __name__ == "__main__":
    filePath = input("请输入文件路径: ")
    importer = importation(filePath)
    data = importer.read()
    # raise
    print(tabulate.tabulate(data, headers="firstrow"))  # 输出读取的数据