import random
import math
import copy

class Arrangement:
    def __init__(self, lists, row, col, deskLeftSit, deskRightSit, arrSex, sexNum, corridor, sex, arrCorridor=False):        #row = 行 col = 列
        self.lists = lists
        self.row = row
        self.col = col
        self.deskLeftSit = deskLeftSit
        self.deskRightSit = deskRightSit
        self.arrSex = arrSex
        self.sexNum = sexNum
        self.corridor = corridor
        self.sex = sex
        self.arrCorridor = arrCorridor  # 新增按走廊分配选项

    def arr(self):
        """核心座位分配方法"""
        # 计算可用座位数（减去走廊）
        available_seats = (self.col - len(self.corridor)) * self.row
        # 计算需要分配的学生数（减去讲台座位）
        student_count = len(self.lists) 
        required_seats = student_count - (1 if self.deskLeftSit else 0) - (1 if self.deskRightSit else 0)
        
        # 检查座位是否足够
        if available_seats < required_seats:
            print("座位数不足")
            return False
        
        print("座位数足够")
        
        # 创建座位映射表 {座位编号: (行, 列)}
        seat_map = {}
        seat_index = 0
        for row in range(self.row):
            for col in range(self.col):
                seat_map[seat_index] = (row, col)
                seat_index += 1
        
        # 深拷贝学生列表
        students = copy.deepcopy(self.lists)
        result = {}
        
        # 处理讲台座位
        if self.deskLeftSit:
            left_student = random.choice(list(students.keys()))
            result["left"] = students[left_student]
            del students[left_student]
        
        if self.deskRightSit:
            right_student = random.choice(list(students.keys()))
            result["right"] = students[right_student]
            del students[right_student]
        
        # 根据分配策略选择算法
        if self.arrCorridor:
            # 使用按走廊分配方法
            assignments = self.arrange_by_corridor(students, self.sex)
            # 将分配结果转换为座位索引形式
            seat_assignments = {}
            seat_index = 0
            for i in range(self.row):
                for j in range(self.col):
                    if (j + 1) in self.corridor:
                        continue  # 跳过走廊
                    if (i, j) in assignments:
                        seat_assignments[seat_index] = assignments[(i, j)]
                        seat_index += 1
        elif self.arrSex:
            seat_assignments = self.arrange_by_gender(available_seats, students, self.sexNum, self.sex)
        else:
            seat_assignments = self.arrange_randomly(students, seat_map)
        
        # 构建最终座位表
        seat_index = 0
        for i in range(self.row * self.col):
            row, col = seat_map[i]
            
            # 处理走廊位置
            if (col + 1) in self.corridor:
                result[(row, col)] = -1  # -1 表示走廊
            # 分配学生或标记空位
            elif seat_index in seat_assignments:
                result[(row, col)] = seat_assignments[seat_index]
                seat_index += 1
            else:
                result[(row, col)] = 0  # 0 表示空位
                seat_index += 1
        
        return result
    
    def arrange_by_gender(self, available_seats, students, group_size, gender_info):
        """按性别分组分配座位"""
        assignments = {}
        # 按性别分组学生
        males = [sid for sid in students.keys() if gender_info.get(sid) == "男"]
        females = [sid for sid in students.keys() if gender_info.get(sid) == "女"]
        
        # 计算需要分配的学生数
        remaining_students = len(students)
        seat_index = 0
        
        while remaining_students > 0 and seat_index < available_seats:
            # 选择当前组的性别（优先选择学生数较多的性别）
            current_group = []
            if len(males) >= len(females) and males:
                current_group = males
            elif females:
                current_group = females
            
            # 如果两个性别都有学生，随机选择性别
            if males and females:
                current_group = random.choice([males, females])
            
            if not current_group:
                break  # 没有学生可分配
                
            # 从当前组中取出最多group_size个学生
            group_count = min(group_size, len(current_group), available_seats - seat_index)
            for _ in range(group_count):
                student = random.choice(current_group)
                assignments[seat_index] = student
                current_group.remove(student)
                seat_index += 1
                remaining_students -= 1
        
        return assignments

    def arrange_by_corridor(self, students, gender_info):
        """按走廊分配座位（改进版）"""
        assignments = {}
        # 按性别分组学生
        males = [sid for sid in students.keys() if gender_info.get(sid) == "男"]
        females = [sid for sid in students.keys() if gender_info.get(sid) == "女"]
        
        # 初始化当前性别（随机开始）
        current_sex = random.choice(["男", "女"])
        
        for i in range(self.row):
            for j in range(self.col):
                # 检查是否是走廊位置
                if (j + 1) in self.corridor:
                    assignments[(i, j)] = -1  # 标记为走廊
                    # 遇到走廊切换性别
                    current_sex = random.choice(["男", "女"])
                else:
                    # 根据当前性别分配学生
                    if current_sex == "男" and males:
                        student_id = males.pop(0)
                        assignments[(i, j)] = student_id
                    elif current_sex == "女" and females:
                        student_id = females.pop(0)
                        assignments[(i, j)] = student_id
                    else:
                        # 如果当前性别组没有学生，尝试从另一组分配
                        if males:
                            student_id = males.pop(0)
                            assignments[(i, j)] = student_id
                        elif females:
                            student_id = females.pop(0)
                            assignments[(i, j)] = student_id
                        else:
                            assignments[(i, j)] = 0  # 没有学生可分配
        
        return assignments

    def arrange_randomly(self, students, seat_map):
        """随机分配座位"""
        assignments = {}
        student_ids = list(students.keys())
        random.shuffle(student_ids)
        
        for idx, student_id in enumerate(student_ids):
            if idx < len(seat_map):
                assignments[idx] = student_id
        
        return assignments

if __name__ == "__main__":
    dict = {1: '皇甫清琪', 2: '东方君华', 3: '杜鹏雁', 4: '童嘉武', 5: '应筠影', 6: '欧阳飘榕', 7: '广芬馨', 8: '米贞利', 9: '路凝梁', 10: '万国波', 11: '管荔梅', 12: ' 别凤娣', 13: '公孙邦才', 14: '华睿枫', 15: '令狐蕊苇', 16: '郎波安', 17: '左辰淑', 18: '荀顺阳', 19: '利枫梦', 20: '司徒妍秀', 21: '孔光琛', 22: '柳言坚', 23: '梅豪娜', 24: '支新旭', 25: '许贝贝', 26: '林烟霭', 27: '韦生蝶', 28: '步玛芳', 29: '向腾琬', 30: '管泽广', 31: '龚良震', 32: '纪桦成', 33: '华菊忠', 34: '皇甫发菡', 35: '邢彪钧', 36: '文剑杰', 37: '季初力', 38: '郝睿贤', 39: '葛子融', 40: '季利功', 41: '葛薇竹', 42: '霍翠莺', 43: '房程韦', 44: '甄可绍', 45: '令狐希 行', 46: '秦眉华', 47: '常伟顺', 48: '金园卿', 49: '司功士', 50: '葛承悦', 51: '司马以腾', 52: '华国芳', 53: '诸固', 54: '邵晨妹', 55: '鲍苑妹', 56: '荆策丹', 57: '仲辰欢', 58: '习宁希', 59: '储琼娴', 60: '贺姣绍', 61: '章春强', 62: '朗雁', 63: '徐离先紫', 64: '司空红元'}
    sex = {1: 'female', 2: 'male', 3: 'female', 4: 'male', 5: 'female', 6: 'female', 7: 'female', 8: 'female', 9: 'female', 10: 'female', 11: 'female', 12: 'female', 13: 'male', 14: 'female', 15: 'male', 16: 'female', 17: 'female', 18: 'male', 19: 'male', 20: 'male', 21: 'female', 22: 'female', 23: 'female', 24: 'female', 25: 'male', 26: 'male', 27: 'male', 28: 'female', 29: 'male', 30: 'female', 31: 'male', 32: 'male', 33: 'male', 34: 'male', 35: 'female', 36: 'male', 37: 'male', 38: 'male', 39: 'female', 40: 'female', 41: 'female', 42: 'male', 43: 'female', 44: 'female', 45: 'male', 46: 'female', 47: 'male', 48: 'female', 49: 'male', 50: 'male', 51: 'male', 52: 'female', 53: 'male', 54: 'female', 55: 'female', 56: 'male', 57: 'female', 58: 'female', 59: 'female', 60: 'male', 61: 'male', 62: 'female', 63: 'female', 64: 'male'}
    row = 7
    col = 11
    corridor = [4,8]
    deskLeftSit = True
    deskRightSit = False
    arrSex = False
    sexNum = 3
    arrangement = Arrangement(dict, row, col, deskLeftSit, deskRightSit, arrSex, sexNum, corridor, sex)
    arr = arrangement.arr()
    print(arr)

    if arr is False:
        print("座位数不足，无法安排座位")
    else:
        # 打印讲台座位
        if deskLeftSit:
            print("讲台左侧:", arr.get("left", "无"))
        if deskRightSit:
            print("讲台右侧:", arr.get("right", "无"))
            
        # 打印座位表
        for i in range(col + 1):
            print(i, end="\t")
        print()
        for i in range(row):
            print(i+1, end="\t")
            for j in range(col):
                # 获取当前座位值
                seat_val = arr.get((i, j), 0)
                
                if seat_val == -1:
                    print("|", end="\t")  # 走廊
                elif seat_val == 0:
                    print("∅", end="\t")  # 空座位
                else:
                    print(dict[seat_val], end="\t")  # 学生姓名
            print()

'''
sex -> {1: "male"}  键值为UUID
lists -> {1: "张三", 2: "李四", 3: "王五", 4: "赵六"}    键值为UUID
corridor -> [0, 1]
'''
