import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from tabulate import tabulate

class Exportation:
    def __init__(self):
        # 设置中文字体支持（更健壮的字体处理）
        self.font = None
        try:
            # Windows系统字体路径
            windows_fonts = [
                'C:/Windows/Fonts/simhei.ttf',
                'C:/Windows/Fonts/msyh.ttc',  # 微软雅黑
                'C:/Windows/Fonts/simkai.ttf' # 楷体
            ]
            
            # Linux系统字体路径
            linux_fonts = [
                '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'
            ]
            
            # 检查可用字体
            font_candidates = windows_fonts if os.name == 'nt' else linux_fonts
            for font_path in font_candidates:
                if os.path.exists(font_path):
                    self.font = FontProperties(fname=font_path, size=10)
                    break
        except Exception as e:
            print(f"字体加载失败: {str(e)}")
            self.font = None
    
    def export(self, data, formats, filename):
        """导出数据到指定格式
        
        Args:
            data (pd.DataFrame): 要导出的数据
            formats (list): 导出格式列表 (['excel', 'csv', 'png'])
            filename (str): 输出文件名(不含扩展名)
        """
        for format_type in formats:
            if format_type == 'excel':
                self._export_excel(data, filename)
            elif format_type == 'csv':
                self._export_csv(data, filename)
            elif format_type == 'png':
                self._export_png(data, filename)
            else:
                print(f"警告: 跳过不支持的导出格式 '{format_type}'")
    
    def _export_excel(self, data, filename):
        """导出到Excel文件"""
        filepath = f"{filename}.xlsx"
        if os.path.exists(filepath):
            os.remove(filepath)
        data.to_excel(filepath, index=False, header=False)
        print(f"已导出到Excel: {filepath}")
    
    def _export_csv(self, data, filename):
        """导出到CSV文件"""
        filepath = f"{filename}.csv"
        if os.path.exists(filepath):
            os.remove(filepath)
        data.to_csv(filepath, index=False, header=False, encoding='utf-8-sig')
        print(f"已导出到CSV: {filepath}")
    
    def _export_png(self, data, filename):
        """导出到PNG图片"""
        filepath = f"{filename}.png"
        if os.path.exists(filepath):
            os.remove(filepath)
        
        # 创建表格图像
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.axis('off')
        ax.axis('tight')
        
        # 创建表格
        table = ax.table(
            cellText=data.values,
            loc='center',
            cellLoc='center'
        )
        
        # 设置字体
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        
        # 应用中文字体（使用正确的API）
        if self.font:
            for key, cell in table.get_celld().items():
                text = cell.get_text()
                text.set_fontproperties(self.font)
        
        # 调整表格样式
        table.scale(1.2, 1.2)
        plt.savefig(filepath, bbox_inches='tight', dpi=150)
        plt.close()
        print(f"已导出到PNG: {filepath}")

if __name__ == "__main__":
    # 测试导出功能
    test_data = pd.DataFrame({
        '姓名': ['张三', '李四', '王五'],
        '性别': ['男', '女', '男']
    })
    
    exporter = Exportation()
    exporter.export(test_data, ['excel', 'csv', 'png'], 'test_export')
