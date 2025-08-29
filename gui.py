import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QTableWidgetItem
from PyQt5.QtCore import Qt
from qfluentwidgets import (FluentWindow, NavigationItemPosition, setTheme, Theme,
                            PrimaryPushButton, ComboBox, SpinBox, LineEdit, 
                            CheckBox, BodyLabel, TableWidget, 
                            MessageBox, InfoBar, InfoBarPosition)
from qfluentwidgets import FluentIcon as FIF
import importation
import arrangement
import pandas as pd
import exportation

class SeatingPlanGUI(FluentWindow):
    def __init__(self):
        super().__init__()
        setTheme(Theme.DARK)
        
        # Initialize data first
        self.parameters = {
            'file_path': "name_list.xlsx",
            'col': 11,
            'row': 7,
            'deskLeftSit': False,
            'deskRightSit': True,
            'arrSex': True,
            'sexNum': 3,
            'corridor': [4, 8],
            'export_formats': ['excel', 'csv', 'png']
        }
        
        # Create interface
        self.setWindowTitle("座位表生成器")
        self.resize(1000, 700)
        
        # Add navigation items
        self.addSubInterface(self.createInputInterface(), FIF.EDIT, "参数设置")
        self.addSubInterface(self.createPreviewInterface(), FIF.VIEW, "预览结果", NavigationItemPosition.BOTTOM)
        self.addSubInterface(self.createExportInterface(), FIF.SAVE, "导出结果", NavigationItemPosition.BOTTOM)
    
    def createInputInterface(self):
        widget = QWidget()
        widget.setObjectName("InputInterface")
        layout = QVBoxLayout(widget)
        
        # File path input
        file_layout = QHBoxLayout()
        file_layout.addWidget(BodyLabel("名单文件路径:"))
        self.file_edit = LineEdit()
        self.file_edit.setText(self.parameters['file_path'])
        file_layout.addWidget(self.file_edit)
        layout.addLayout(file_layout)
        
        # Column and row inputs
        grid_layout = QGridLayout()
        grid_layout.addWidget(BodyLabel("列数:"), 0, 0)
        self.col_spin = SpinBox()
        self.col_spin.setValue(self.parameters['col'])
        grid_layout.addWidget(self.col_spin, 0, 1)
        
        grid_layout.addWidget(BodyLabel("行数:"), 0, 2)
        self.row_spin = SpinBox()
        self.row_spin.setValue(self.parameters['row'])
        grid_layout.addWidget(self.row_spin, 0, 3)
        layout.addLayout(grid_layout)
        
        # Checkboxes
        self.left_check = CheckBox("左护法")
        self.left_check.setChecked(self.parameters['deskLeftSit'])
        layout.addWidget(self.left_check)
        
        self.right_check = CheckBox("右护法")
        self.right_check.setChecked(self.parameters['deskRightSit'])
        layout.addWidget(self.right_check)
        
        # 性别和走廊分配选项（可以同时启用）
        gender_layout = QVBoxLayout()
        
        # 按性别安排
        gender_group = QHBoxLayout()
        self.gender_check = CheckBox("按性别安排")
        self.gender_check.setChecked(self.parameters['arrSex'])
        gender_group.addWidget(self.gender_check)
        
        # 连续性别数量
        gender_group.addWidget(BodyLabel("连续性别数量:"))
        self.gender_spin = SpinBox()
        self.gender_spin.setValue(self.parameters['sexNum'])
        gender_group.addWidget(self.gender_spin)
        gender_layout.addLayout(gender_group)
        
        # 按走廊安排
        self.corridor_check = CheckBox("按走廊安排")
        self.corridor_check.setChecked(self.parameters.get('arrCorridor', False))
        gender_layout.addWidget(self.corridor_check)
        
        layout.addLayout(gender_layout)
        
        # Corridor positions (replaced with checkboxes)
        corridor_layout = QVBoxLayout()
        corridor_layout.addWidget(BodyLabel("走廊位置 (勾选为走廊，列号):"))
        
        # 使用可滚动区域包裹复选框
        from qfluentwidgets import ScrollArea
        scroll_area = ScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedHeight(60)  # 固定高度避免占用过多空间
        
        # 走廊复选框容器
        self.corridor_container = QWidget()
        self.corridor_check_layout = QHBoxLayout()
        self.corridor_check_layout.setContentsMargins(5, 5, 5, 5)
        self.corridor_container.setLayout(self.corridor_check_layout)
        
        scroll_area.setWidget(self.corridor_container)
        
        corridor_layout.addWidget(scroll_area)
        layout.addLayout(corridor_layout)
        
        # 初始化走廊复选框
        self.corridor_checks = []
        self.updateCorridorCheckboxes(self.parameters['col'])
        
        # 连接列数变化信号
        self.col_spin.valueChanged.connect(self.updateCorridorCheckboxes)
        
        # Export formats
        export_layout = QHBoxLayout()
        export_layout.addWidget(BodyLabel("导出格式:"))
        self.export_combo = ComboBox()
        self.export_combo.addItems(["excel", "csv", "png"])
        self.export_combo.setCurrentText(",".join(self.parameters['export_formats']))
        export_layout.addWidget(self.export_combo)
        layout.addLayout(export_layout)
        
        # Generate button
        self.generate_btn = PrimaryPushButton("生成座位表")
        self.generate_btn.clicked.connect(self.generateSeatingPlan)
        layout.addWidget(self.generate_btn)
        
        layout.addStretch(1)
        return widget
    
    def createPreviewInterface(self):
        widget = QWidget()
        widget.setObjectName("PreviewInterface")
        layout = QVBoxLayout(widget)
        
        # 创建表格
        self.table = TableWidget()
        self.table.setColumnCount(self.parameters['col'])
        self.table.setRowCount(self.parameters['row'])
        self.table.itemChanged.connect(self.on_table_item_changed)  # 连接单元格修改信号
        layout.addWidget(self.table)
        
        # 添加保存按钮
        self.save_btn = PrimaryPushButton("保存修改")
        self.save_btn.clicked.connect(self.saveChanges)
        layout.addWidget(self.save_btn)
        
        return widget
    
    def createExportInterface(self):
        widget = QWidget()
        widget.setObjectName("ExportInterface")
        layout = QVBoxLayout(widget)
        
        self.export_btn = PrimaryPushButton("导出结果")
        self.export_btn.clicked.connect(self.exportResults)
        layout.addWidget(self.export_btn)
        
        return widget
    
    def updateCorridorCheckboxes(self, col_count):
        """根据列数更新走廊复选框（使用列号1-based）"""
        # 清除现有复选框
        for i in reversed(range(self.corridor_check_layout.count())):
            item = self.corridor_check_layout.itemAt(i)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()
        self.corridor_checks = []
        
        # 创建新的复选框（显示列号1-based）
        for i in range(col_count):
            check = CheckBox(f"{i+1}")
            check.setObjectName(f"corridor_check_{i}")
            self.corridor_checks.append(check)
            self.corridor_check_layout.addWidget(check)
        
        # 设置当前选中的走廊位置（使用列号1-based）
        if hasattr(self, 'current_corridor_positions'):
            for pos in self.current_corridor_positions:
                if pos-1 < len(self.corridor_checks):  # 列号转索引
                    self.corridor_checks[pos-1].setChecked(True)
    
    def generateSeatingPlan(self):
        # 获取走廊位置（使用列号1-based）
        corridor_positions = []
        for i, check in enumerate(self.corridor_checks):
            if check.isChecked():
                corridor_positions.append(i+1)  # 保存列号而不是索引
        self.current_corridor_positions = corridor_positions  # 保存当前选择
        
        # Get parameters from UI
        params = {
            'file_path': self.file_edit.text(),
            'col': self.col_spin.value(),
            'row': self.row_spin.value(),
            'deskLeftSit': self.left_check.isChecked(),
            'deskRightSit': self.right_check.isChecked(),
            'arrSex': self.gender_check.isChecked(),
            'sexNum': self.gender_spin.value(),
            'arrCorridor': self.corridor_check.isChecked(),  # 新增按走廊分配选项
            'corridor': corridor_positions,  # 使用复选框数据
            'export_formats': self.export_combo.currentText().split(",")
        }
        
        # 清除表格内容（解决护法残留问题）
        self.table.clearContents()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)
        
        # 更新表格尺寸
        self.table.setColumnCount(params['col'])
        self.table.setRowCount(params['row'])
        
        # Read name list
        importer = importation.importation(params['file_path'])
        data = importer.read()
        
        # Check if data is valid
        if not isinstance(data, pd.DataFrame):
            MessageBox("错误", f"读取文件失败: {data[0] if isinstance(data, list) else '未知错误'}", self).exec()
            return
        
        # Prepare data
        lists = {}
        sex = {}
        for i in range(len(data)):
            lists[i+1] = data.iloc[i, 0]
            sex[i+1] = data.iloc[i, 1]
        
        # 初始化走廊复选框状态
        if not hasattr(self, 'current_corridor_positions'):
            self.current_corridor_positions = self.parameters['corridor']
            self.updateCorridorCheckboxes(params['col'])
        
        # Create arrangement
        arr = arrangement.Arrangement(
            lists, 
            params['row'], 
            params['col'], 
            params['deskLeftSit'], 
            params['deskRightSit'], 
            params['arrSex'], 
            params['sexNum'], 
            params['corridor'], 
            sex,
            params['arrCorridor']  # 添加按走廊分配选项
        )
        
        # Generate arrangement
        arr_result = arr.arr()
        
        if not arr_result:
            MessageBox("错误", "座位数不足，无法安排", self).exec()
            # 清除预览表格
            self.table.clearContents()
            return
            
        # 创建结果DataFrame（与展示界面布局一致）
        # 计算总行数 = 讲台行(如果有) + 普通行
        total_rows = params['row'] + (1 if (params['deskLeftSit'] or params['deskRightSit']) else 0)
        res_df = pd.DataFrame(index=range(total_rows), columns=range(params['col']))
        sex_df = pd.DataFrame(index=range(total_rows), columns=range(params['col']))
        
        # 处理讲台座位（位于第一行中间）
        if params['deskLeftSit'] or params['deskRightSit']:
            row_idx = 0  # 第一行
            middle_col = params['col'] // 2  # 中间列
            
            # 讲台标签
            res_df.iloc[row_idx, middle_col] = "讲台"
            sex_df.iloc[row_idx, middle_col] = "讲台"
            
            # 左侧座位
            if "left" in arr_result and middle_col > 0:
                res_df.iloc[row_idx, middle_col - 1] = arr_result["left"]
                sex_df.iloc[row_idx, middle_col - 1] = sex.get(arr_result["left"], "")
            
            # 右侧座位
            if params['deskRightSit'] and "right" in arr_result and middle_col < params['col'] - 1:
                res_df.iloc[row_idx, middle_col + 1] = arr_result["right"]
                sex_df.iloc[row_idx, middle_col + 1] = sex.get(arr_result["right"], "")
        
        # 处理普通座位（从第二行开始）
        start_row = 1 if (params['deskLeftSit'] or params['deskRightSit']) else 0
        for i in range(params['row']):
            for j in range(params['col']):
                if (i, j) in arr_result:
                    seat_val = arr_result[(i, j)]
                    row_idx = start_row + i
                    
                    if seat_val == 0:
                        res_df.iloc[row_idx, j] = ""  # 图片导出时空座位留空
                        sex_df.iloc[row_idx, j] = ""
                    elif seat_val == -1:
                        res_df.iloc[row_idx, j] = "走廊"
                        sex_df.iloc[row_idx, j] = "走廊"
                    else:
                        res_df.iloc[row_idx, j] = lists.get(seat_val, "")
                        sex_df.iloc[row_idx, j] = sex.get(seat_val, "")
        
        # 更新表格尺寸（考虑讲台行）
        total_rows = params['row'] + (1 if (params['deskLeftSit'] or params['deskRightSit']) else 0)
        self.table.setRowCount(total_rows)
        self.table.setColumnCount(params['col'])
        
        # 更新表格内容
        row_index = 0
        
        # 添加讲台行（位于第一行中间）
        if params['deskLeftSit'] or params['deskRightSit']:
            # 计算讲台居中位置
            middle_col = params['col'] // 2
            
            # 添加讲台标签（居中位置）
            self.table.setItem(row_index, middle_col, QTableWidgetItem("讲台"))
            
            # 添加左侧座位（讲台左侧）
            if "left" in arr_result and middle_col > 0:
                self.table.setItem(row_index, middle_col - 1, QTableWidgetItem(arr_result["left"]))
            
            # 添加右侧座位（讲台右侧）
            if params['deskRightSit'] and "right" in arr_result and middle_col < params['col'] - 1:
                self.table.setItem(row_index, middle_col + 1, QTableWidgetItem(arr_result["right"]))
            
            row_index += 1
        
        # 添加普通座位
        for i in range(params['row']):
            for j in range(params['col']):
                if (i, j) in arr_result:
                    seat_val = arr_result[(i, j)]
                    if seat_val == 0:
                        self.table.setItem(row_index + i, j, QTableWidgetItem("空"))  # 表格中显示"空"
                    elif seat_val == -1:
                        self.table.setItem(row_index + i, j, QTableWidgetItem("走廊"))
                    else:
                        self.table.setItem(row_index + i, j, QTableWidgetItem(lists.get(seat_val, "")))
        
        # Store results for export (替换NaN值为空字符串)
        self.res_df = res_df.fillna('')
        self.sex_df = sex_df.fillna('')
        
        # Show success message
        InfoBar.success(
            title='生成成功',
            content='座位表已生成',
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=2000,
            parent=self
        )
    
    def on_table_item_changed(self, item):
        """当用户修改座位名称时更新表格数据"""
        # 获取修改的行列
        row = item.row()
        column = item.column()
        
        # 更新座位名称
        if hasattr(self, 'res_df'):
            self.res_df.iloc[row, column] = item.text()
        
    def saveChanges(self):
        """保存用户对座位表的修改"""
        if not hasattr(self, 'table'):
            return
            
        # 显示保存成功消息
        InfoBar.success(
            title='保存成功',
            content='座位表修改已保存',
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=2000,
            parent=self
        )
        
    def exportResults(self):
        if not hasattr(self, 'res_df') or not hasattr(self, 'sex_df'):
            MessageBox("错误", "请先生成座位表", self).exec()
            return
            
        try:
            # Get export formats from UI
            export_formats = self.export_combo.currentText().split(",")
            
            # 从表格获取最新数据（覆盖原始数据）
            for row in range(self.table.rowCount()):
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item:
                        self.res_df.iloc[row, col] = item.text()
            
            # Export results with enhanced error handling
            exporter = exportation.Exportation()
            exporter.export(self.res_df, export_formats, "arrangement_result")
            exporter.export(self.sex_df, export_formats, "arrangement_sex_result")
            
            # Show success message
            InfoBar.success(
                title='导出成功',
                content='结果已导出',
                orient=Qt.Orientation.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=2000,
                parent=self
            )
        except Exception as e:
            # Show error message
            MessageBox("导出错误", f"导出失败: {str(e)}", self).exec()
            # Log the error for debugging
            print(f"导出失败: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SeatingPlanGUI()
    window.show()
    sys.exit(app.exec_())
