// 座位安排系统核心逻辑
class SeatingArrangement {
    constructor() {
        this.students = new Map(); // {id: {name, gender}}
        this.seatingResult = null;
        this.isEditMode = false;
        
        this.initializeEventListeners();
        this.updateCorridorCheckboxes();
    }

    // 初始化事件监听器
    initializeEventListeners() {
        // 文件上传
        document.getElementById('fileInput').addEventListener('change', (e) => {
            this.handleFileUpload(e);
        });

        // 列数变化时更新走廊选项
        document.getElementById('cols').addEventListener('input', () => {
            this.updateCorridorCheckboxes();
        });

        // 排列方式选择
        document.querySelectorAll('input[name="arrangeType"]').forEach(radio => {
            radio.addEventListener('change', () => {
                this.handleArrangeTypeChange();
            });
        });

        // 生成座位表
        document.getElementById('generateBtn').addEventListener('click', () => {
            this.generateSeatingPlan();
        });

        // 清空
        document.getElementById('clearBtn').addEventListener('click', () => {
            this.clearAll();
        });

        // 编辑模式
        document.getElementById('editModeBtn').addEventListener('click', () => {
            this.toggleEditMode();
        });

        // 导出
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.showExportModal();
        });

        // Modal相关
        document.getElementById('modalClose').addEventListener('click', () => {
            this.hideExportModal();
        });

        document.getElementById('cancelExport').addEventListener('click', () => {
            this.hideExportModal();
        });

        document.getElementById('confirmExport').addEventListener('click', () => {
            this.exportSeatingPlan();
        });

        // 点击modal外部关闭
        document.getElementById('exportModal').addEventListener('click', (e) => {
            if (e.target.id === 'exportModal') {
                this.hideExportModal();
            }
        });
    }

    // 处理文件上传
    async handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        const fileInfo = document.getElementById('fileInfo');
        fileInfo.textContent = `已选择: ${file.name}`;

        try {
            this.showLoading();
            const data = await this.parseFile(file);
            this.students = data;
            this.updateStatistics();
            this.showToast('文件上传成功', 'success');
        } catch (error) {
            this.showToast(`文件解析失败: ${error.message}`, 'error');
        } finally {
            this.hideLoading();
        }
    }

    // 解析文件
    async parseFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    let students = new Map();

                    if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                        // Excel文件
                        const workbook = XLSX.read(data, { type: 'array' });
                        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        jsonData.forEach((row, index) => {
                            if (row.length >= 2 && row[0] && row[1]) {
                                students.set(index + 1, {
                                    name: row[0].toString(),
                                    gender: this.normalizeGender(row[1].toString())
                                });
                            }
                        });
                    } else if (file.name.endsWith('.csv')) {
                        // CSV文件
                        const text = new TextDecoder('utf-8').decode(data);
                        const lines = text.split('\n').filter(line => line.trim());
                        
                        lines.forEach((line, index) => {
                            const parts = line.split(',').map(part => part.trim());
                            if (parts.length >= 2) {
                                students.set(index + 1, {
                                    name: parts[0],
                                    gender: this.normalizeGender(parts[1])
                                });
                            }
                        });
                    } else if (file.name.endsWith('.txt')) {
                        // TXT文件
                        const text = new TextDecoder('utf-8').decode(data);
                        const lines = text.split('\n').filter(line => line.trim());
                        
                        lines.forEach((line, index) => {
                            const parts = line.split(/\s+/);
                            if (parts.length >= 2) {
                                students.set(index + 1, {
                                    name: parts[0],
                                    gender: this.normalizeGender(parts[1])
                                });
                            }
                        });
                    }

                    resolve(students);
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    // 标准化性别格式
    normalizeGender(gender) {
        const genderStr = gender.toLowerCase();
        if (genderStr.includes('男') || genderStr.includes('male') || genderStr === 'm') {
            return '男';
        } else if (genderStr.includes('女') || genderStr.includes('female') || genderStr === 'f') {
            return '女';
        }
        return '未知';
    }

    // 更新走廊复选框
    updateCorridorCheckboxes() {
        const cols = parseInt(document.getElementById('cols').value);
        const container = document.getElementById('corridorCheckboxes');
        container.innerHTML = '';

        for (let i = 1; i <= cols; i++) {
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `corridor_${i}`;
            checkbox.value = i;
            
            // 默认选择第4列和第8列作为走廊
            if (i === 4 || i === 8) {
                checkbox.checked = true;
            }

            const label = document.createElement('label');
            label.htmlFor = `corridor_${i}`;
            label.textContent = i;
            label.style.fontSize = '0.85rem';

            const wrapper = document.createElement('div');
            wrapper.className = 'checkbox-group';
            wrapper.appendChild(checkbox);
            wrapper.appendChild(label);

            container.appendChild(wrapper);
        }
    }

    // 处理排列方式变化
    handleArrangeTypeChange() {
        const selectedType = document.querySelector('input[name="arrangeType"]:checked').value;
        const genderSettings = document.querySelector('.gender-settings');
        
        if (selectedType === 'gender') {
            genderSettings.style.display = 'block';
        } else {
            genderSettings.style.display = 'none';
        }
    }

    // 生成座位表 - 核心算法
    generateSeatingPlan() {
        if (this.students.size === 0) {
            this.showToast('请先上传学生名单', 'warning');
            return;
        }

        try {
            this.showLoading();
            
            // 获取参数
            const params = this.getParameters();
            
            // 验证座位数量
            const availableSeats = (params.cols - params.corridor.length) * params.rows;
            const requiredSeats = this.students.size - (params.deskLeftSit ? 1 : 0) - (params.deskRightSit ? 1 : 0);
            
            if (availableSeats < requiredSeats) {
                throw new Error('座位数不足，无法安排所有学生');
            }

            // 执行排列算法
            this.seatingResult = this.executeArrangement(params);
            
            // 渲染座位表
            this.renderSeatingGrid(params);
            this.updateStatistics();
            
            // 显示导出按钮和统计面板
            document.getElementById('exportBtn').style.display = 'block';
            document.getElementById('statsPanel').style.display = 'block';
            
            this.showToast('座位表生成成功', 'success');
        } catch (error) {
            this.showToast(error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    // 获取参数
    getParameters() {
        const arrangeType = document.querySelector('input[name="arrangeType"]:checked').value;
        const corridorCheckboxes = document.querySelectorAll('#corridorCheckboxes input[type="checkbox"]:checked');
        const corridor = Array.from(corridorCheckboxes).map(cb => parseInt(cb.value));

        return {
            rows: parseInt(document.getElementById('rows').value),
            cols: parseInt(document.getElementById('cols').value),
            deskLeftSit: document.getElementById('leftDesk').checked,
            deskRightSit: document.getElementById('rightDesk').checked,
            arrangeType: arrangeType,
            genderGroupSize: parseInt(document.getElementById('genderGroup').value),
            corridor: corridor
        };
    }

    // 执行排列算法 - 移植自Python代码
    executeArrangement(params) {
        const result = {};
        const students = new Map(this.students);

        // 处理讲台座位
        if (params.deskLeftSit) {
            const leftStudent = this.getRandomStudent(students);
            result.left = leftStudent;
            students.delete(leftStudent.id);
        }

        if (params.deskRightSit) {
            const rightStudent = this.getRandomStudent(students);
            result.right = rightStudent;
            students.delete(rightStudent.id);
        }

        // 创建座位映射
        const seatMap = {};
        let seatIndex = 0;
        for (let row = 0; row < params.rows; row++) {
            for (let col = 0; col < params.cols; col++) {
                seatMap[seatIndex] = { row, col };
                seatIndex++;
            }
        }

        // 根据排列方式执行算法
        let seatAssignments = {};
        
        switch (params.arrangeType) {
            case 'gender':
                seatAssignments = this.arrangeByGender(params, students);
                break;
            case 'corridor':
                seatAssignments = this.arrangeByCorridor(params, students);
                break;
            default:
                seatAssignments = this.arrangeRandomly(students);
                break;
        }

        // 构建最终座位表
        seatIndex = 0;
        for (let i = 0; i < params.rows * params.cols; i++) {
            const { row, col } = seatMap[i];
            
            // 处理走廊位置
            if (params.corridor.includes(col + 1)) {
                result[`${row},${col}`] = { type: 'corridor' };
            } else if (seatIndex in seatAssignments) {
                result[`${row},${col}`] = seatAssignments[seatIndex];
                seatIndex++;
            } else {
                result[`${row},${col}`] = { type: 'empty' };
                seatIndex++;
            }
        }

        return result;
    }

    // 按性别排列
    arrangeByGender(params, students) {
        const assignments = {};
        const males = [];
        const females = [];

        // 分组
        for (const [id, student] of students) {
            if (student.gender === '男') {
                males.push({ id, ...student });
            } else {
                females.push({ id, ...student });
            }
        }

        // 随机打乱
        this.shuffleArray(males);
        this.shuffleArray(females);

        let seatIndex = 0;
        const availableSeats = (params.cols - params.corridor.length) * params.rows;
        
        while ((males.length > 0 || females.length > 0) && seatIndex < availableSeats) {
            // 选择当前组
            let currentGroup = [];
            if (males.length >= females.length && males.length > 0) {
                currentGroup = males;
            } else if (females.length > 0) {
                currentGroup = females;
            }

            if (currentGroup.length === 0) break;

            // 分配一组学生
            const groupSize = Math.min(params.genderGroupSize, currentGroup.length, availableSeats - seatIndex);
            for (let i = 0; i < groupSize; i++) {
                const student = currentGroup.shift();
                assignments[seatIndex] = student;
                seatIndex++;
            }
        }

        return assignments;
    }

    // 按走廊排列
    arrangeByCorridor(params, students) {
        const assignments = {};
        const males = [];
        const females = [];

        // 分组
        for (const [id, student] of students) {
            if (student.gender === '男') {
                males.push({ id, ...student });
            } else {
                females.push({ id, ...student });
            }
        }

        // 随机打乱
        this.shuffleArray(males);
        this.shuffleArray(females);

        let currentGender = Math.random() > 0.5 ? '男' : '女';
        
        for (let row = 0; row < params.rows; row++) {
            for (let col = 0; col < params.cols; col++) {
                // 检查是否是走廊
                if (params.corridor.includes(col + 1)) {
                    // 遇到走廊切换性别
                    currentGender = Math.random() > 0.5 ? '男' : '女';
                    continue;
                }

                // 分配学生
                let student = null;
                if (currentGender === '男' && males.length > 0) {
                    student = males.shift();
                } else if (currentGender === '女' && females.length > 0) {
                    student = females.shift();
                } else {
                    // 当前性别组没有学生，从另一组分配
                    if (males.length > 0) {
                        student = males.shift();
                    } else if (females.length > 0) {
                        student = females.shift();
                    }
                }

                if (student) {
                    const seatKey = this.getSeatKey(row, col, params);
                    if (seatKey !== -1) {
                        assignments[seatKey] = student;
                    }
                }
            }
        }

        return assignments;
    }

    // 随机排列
    arrangeRandomly(students) {
        const assignments = {};
        const studentArray = Array.from(students.values()).map((student, index) => ({
            id: Array.from(students.keys())[index],
            ...student
        }));
        
        this.shuffleArray(studentArray);
        
        studentArray.forEach((student, index) => {
            assignments[index] = student;
        });

        return assignments;
    }

    // 获取座位键
    getSeatKey(row, col, params) {
        let seatIndex = 0;
        for (let r = 0; r < params.rows; r++) {
            for (let c = 0; c < params.cols; c++) {
                if (params.corridor.includes(c + 1)) {
                    continue;
                }
                if (r === row && c === col) {
                    return seatIndex;
                }
                seatIndex++;
            }
        }
        return -1;
    }

    // 获取随机学生
    getRandomStudent(students) {
        const studentIds = Array.from(students.keys());
        const randomId = studentIds[Math.floor(Math.random() * studentIds.length)];
        const student = students.get(randomId);
        return { id: randomId, ...student };
    }

    // 数组随机打乱
    shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [array[i], array[j]] = [array[j], array[i]];
        }
    }

    // 渲染整合的座位布局
    renderSeatingGrid(params) {
        const seatingArea = document.getElementById('seatingArea');
        seatingArea.innerHTML = '';

        // 创建整合布局容器
        const layout = document.createElement('div');
        layout.className = 'integrated-layout';
        layout.style.gridTemplateColumns = `repeat(${params.cols}, 1fr)`;

        // 添加黑板
        const blackboard = document.createElement('div');
        blackboard.className = 'blackboard-row';
        blackboard.textContent = '黑板';
        layout.appendChild(blackboard);

        // 添加讲台区域（如果有护法座位）
        if (params.deskLeftSit || params.deskRightSit) {
            const podiumRow = this.createPodiumRow(params);
            layout.appendChild(podiumRow);
        }

        // 添加普通座位
        for (let row = 0; row < params.rows; row++) {
            for (let col = 0; col < params.cols; col++) {
                const seat = this.createSeatElement(row, col, params);
                layout.appendChild(seat);
            }
        }

        seatingArea.appendChild(layout);
    }

    // 创建讲台行
    createPodiumRow(params) {
        const podiumRow = document.createElement('div');
        podiumRow.className = 'podium-row';
        podiumRow.style.gridTemplateColumns = `repeat(${params.cols}, 1fr)`;

        const middleCol = Math.floor(params.cols / 2);

        for (let col = 0; col < params.cols; col++) {
            const element = document.createElement('div');
            
            if (col === middleCol - 1 && params.deskLeftSit) {
                // 左护法
                element.className = 'teacher-seat left';
                element.dataset.position = 'left';
                if (this.seatingResult.left) {
                    element.textContent = this.seatingResult.left.name;
                    element.classList.add('occupied');
                } else {
                    element.textContent = '左护法';
                }
                
                // 添加编辑功能
                element.addEventListener('click', () => {
                    if (this.isEditMode) {
                        this.editTeacherSeat('left', element);
                    }
                });
            } else if (col === middleCol) {
                // 讲台
                element.className = 'podium';
                element.textContent = '讲台';
            } else if (col === middleCol + 1 && params.deskRightSit) {
                // 右护法
                element.className = 'teacher-seat right';
                element.dataset.position = 'right';
                if (this.seatingResult.right) {
                    element.textContent = this.seatingResult.right.name;
                    element.classList.add('occupied');
                } else {
                    element.textContent = '右护法';
                }
                
                // 添加编辑功能
                element.addEventListener('click', () => {
                    if (this.isEditMode) {
                        this.editTeacherSeat('right', element);
                    }
                });
            } else {
                // 空位置
                element.style.width = '50px';
                element.style.height = '40px';
            }
            
            podiumRow.appendChild(element);
        }

        return podiumRow;
    }

    // 编辑讲台座位
    editTeacherSeat(position, element) {
        const currentName = element.classList.contains('occupied') ? element.textContent : '';
        const newName = prompt(`请输入${position === 'left' ? '左护法' : '右护法'}姓名（留空清除）:`, currentName);
        
        if (newName !== null) {
            if (newName.trim() === '') {
                // 清除座位
                delete this.seatingResult[position];
                element.textContent = position === 'left' ? '左护法' : '右护法';
                element.classList.remove('occupied');
            } else {
                // 设置学生
                this.seatingResult[position] = {
                    name: newName.trim(),
                    gender: '未知'
                };
                element.textContent = newName.trim();
                element.classList.add('occupied');
            }
            
            this.updateStatistics();
        }
    }

    // 创建座位元素
    createSeatElement(row, col, params) {
        const seat = document.createElement('div');
        seat.className = 'seat';
        seat.dataset.row = row;
        seat.dataset.col = col;

        const key = `${row},${col}`;
        const seatData = this.seatingResult[key];

        if (seatData) {
            if (seatData.type === 'corridor') {
                seat.classList.add('corridor');
                seat.textContent = '走廊';
            } else if (seatData.type === 'empty') {
                seat.classList.add('empty');
                seat.textContent = '空';
            } else {
                seat.classList.add('occupied');
                seat.textContent = seatData.name;
                
                if (seatData.gender === '男') {
                    seat.classList.add('male');
                } else if (seatData.gender === '女') {
                    seat.classList.add('female');
                }
            }
        } else {
            seat.classList.add('empty');
            seat.textContent = '空';
        }

        // 添加编辑功能
        if (!seat.classList.contains('corridor')) {
            seat.addEventListener('click', () => {
                if (this.isEditMode) {
                    this.editSeat(seat);
                }
            });
        }

        return seat;
    }

    // 编辑座位
    editSeat(seatElement) {
        const currentName = seatElement.textContent;
        const newName = prompt('请输入学生姓名（留空清除座位）:', currentName === '空' ? '' : currentName);
        
        if (newName !== null) {
            const row = parseInt(seatElement.dataset.row);
            const col = parseInt(seatElement.dataset.col);
            const key = `${row},${col}`;

            if (newName.trim() === '') {
                // 清除座位
                this.seatingResult[key] = { type: 'empty' };
                seatElement.className = 'seat empty';
                seatElement.textContent = '空';
            } else {
                // 设置学生
                this.seatingResult[key] = {
                    name: newName.trim(),
                    gender: '未知' // 编辑时默认性别
                };
                seatElement.className = 'seat occupied';
                seatElement.textContent = newName.trim();
            }
            
            this.updateStatistics();
        }
    }

    // 切换编辑模式
    toggleEditMode() {
        this.isEditMode = !this.isEditMode;
        const btn = document.getElementById('editModeBtn');
        const seats = document.querySelectorAll('.seat:not(.corridor)');
        const teacherSeats = document.querySelectorAll('.teacher-seat');

        if (this.isEditMode) {
            btn.innerHTML = '<i class="fas fa-check"></i> 完成编辑';
            btn.classList.add('btn-primary');
            btn.classList.remove('btn-secondary');
            seats.forEach(seat => seat.classList.add('editable'));
            teacherSeats.forEach(seat => seat.classList.add('editable'));
            this.showToast('已进入编辑模式，点击座位可修改', 'success');
        } else {
            btn.innerHTML = '<i class="fas fa-edit"></i> 编辑模式';
            btn.classList.add('btn-secondary');
            btn.classList.remove('btn-primary');
            seats.forEach(seat => seat.classList.remove('editable'));
            teacherSeats.forEach(seat => seat.classList.remove('editable'));
            this.showToast('已退出编辑模式', 'success');
        }
    }

    // 更新统计信息
    updateStatistics() {
        const maleCount = Array.from(this.students.values()).filter(s => s.gender === '男').length;
        const femaleCount = Array.from(this.students.values()).filter(s => s.gender === '女').length;
        
        document.getElementById('totalStudents').textContent = this.students.size;
        document.getElementById('maleCount').textContent = maleCount;
        document.getElementById('femaleCount').textContent = femaleCount;

        if (this.seatingResult) {
            const params = this.getParameters();
            
            // 计算已安排的学生座位（包括左右护法）
            let assignedStudentSeats = 0;
            let regularStudentSeats = 0;
            let teacherSeats = 0;
            
            // 统计普通座位中的学生（排除走廊和空座位）
            Object.entries(this.seatingResult).forEach(([key, seat]) => {
                // 跳过left和right护法座位，这些会单独统计
                if (key === 'left' || key === 'right') return;
                
                // 检查是否是有效的学生座位
                if (seat && seat.name && seat.type !== 'empty' && seat.type !== 'corridor') {
                    regularStudentSeats += 1;
                }
            });
            
            // 统计左右护法
            if (this.seatingResult.left && this.seatingResult.left.name) {
                teacherSeats += 1;
            }
            if (this.seatingResult.right && this.seatingResult.right.name) {
                teacherSeats += 1;
            }
            
            assignedStudentSeats = regularStudentSeats + teacherSeats;
            
            // 总可用座位数 = 普通座位（扣除走廊）+ 护法座位
            const regularSeatsTotal = (params.cols - params.corridor.length) * params.rows;
            const teacherSeatsTotal = (params.deskLeftSit ? 1 : 0) + (params.deskRightSit ? 1 : 0);
            const totalAvailableSeats = regularSeatsTotal + teacherSeatsTotal;
            
            // 调试信息
            console.log('统计调试信息:', {
                regularStudentSeats,
                teacherSeats,
                assignedStudentSeats,
                regularSeatsTotal,
                teacherSeatsTotal,
                totalAvailableSeats,
                params,
                leftSeat: this.seatingResult.left,
                rightSeat: this.seatingResult.right
            });
            
            document.getElementById('assignedSeats').textContent = assignedStudentSeats;
            document.getElementById('emptySeats').textContent = totalAvailableSeats - assignedStudentSeats;
        } else {
            document.getElementById('assignedSeats').textContent = 0;
            document.getElementById('emptySeats').textContent = 0;
        }
    }

    // 清空所有数据
    clearAll() {
        if (confirm('确定要清空所有数据吗？')) {
            this.students.clear();
            this.seatingResult = null;
            document.getElementById('seatingArea').innerHTML = '';
            document.getElementById('fileInput').value = '';
            document.getElementById('fileInfo').textContent = '';
            document.getElementById('exportBtn').style.display = 'none';
            document.getElementById('statsPanel').style.display = 'none';
            this.updateStatistics();
            this.showToast('已清空所有数据', 'success');
        }
    }

    // 显示导出模态框
    showExportModal() {
        if (!this.seatingResult) {
            this.showToast('请先生成座位表', 'warning');
            return;
        }
        document.getElementById('exportModal').classList.add('active');
    }

    // 隐藏导出模态框
    hideExportModal() {
        document.getElementById('exportModal').classList.remove('active');
    }

    // 导出座位表
    exportSeatingPlan() {
        const exportExcel = document.getElementById('exportExcel').checked;
        const exportCSV = document.getElementById('exportCSV').checked;
        const exportPNG = document.getElementById('exportPNG').checked;

        if (!exportExcel && !exportCSV && !exportPNG) {
            this.showToast('请选择至少一种导出格式', 'warning');
            return;
        }

        try {
            if (exportExcel) this.exportToExcel();
            if (exportCSV) this.exportToCSV();
            if (exportPNG) this.exportToPNG();
            
            this.hideExportModal();
            this.showToast('导出成功', 'success');
        } catch (error) {
            this.showToast(`导出失败: ${error.message}`, 'error');
        }
    }

    // 导出到Excel
    exportToExcel() {
        const data = this.prepareExportData();
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '座位表');
        XLSX.writeFile(wb, '座位安排表.xlsx');
    }

    // 导出到CSV
    exportToCSV() {
        const data = this.prepareExportData();
        const csvContent = data.map(row => 
            row.map(cell => `"${cell || ''}"`).join(',')
        ).join('\n');
        
        const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = '座位安排表.csv';
        link.click();
    }

    // 导出到PNG
    exportToPNG() {
        const seatingArea = document.getElementById('seatingArea');
        if (!seatingArea) {
            this.showToast('没有找到座位布局', 'error');
            return;
        }

        // 使用html2canvas库导出整个座位区域
        // 如果没有html2canvas，则使用fallback方法
        if (typeof html2canvas !== 'undefined') {
            html2canvas(seatingArea, {
                backgroundColor: '#f7fafc',
                scale: 2,
                useCORS: true
            }).then(canvas => {
                canvas.toBlob(blob => {
                    const link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = '座位安排表.png';
                    link.click();
                }, 'image/png');
            });
        } else {
            // Fallback: 使用Canvas绘制
            this.exportToPNGCanvas();
        }
    }

    // Canvas绘制PNG
    exportToPNGCanvas() {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const params = this.getParameters();
        
        // 计算画布尺寸
        const cellWidth = 80;
        const cellHeight = 50;
        const padding = 20;
        const blackboardHeight = 60;
        const podiumHeight = params.deskLeftSit || params.deskRightSit ? 50 : 0;
        
        canvas.width = params.cols * cellWidth + padding * 2;
        canvas.height = blackboardHeight + podiumHeight + params.rows * cellHeight + padding * 2;
        
        // 设置背景
        ctx.fillStyle = '#f7fafc';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        let currentY = padding;
        
        // 绘制黑板
        ctx.fillStyle = '#2d3748';
        ctx.fillRect(padding, currentY, params.cols * cellWidth, blackboardHeight);
        ctx.fillStyle = 'white';
        ctx.font = 'bold 18px Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('黑板', canvas.width / 2, currentY + blackboardHeight / 2);
        
        currentY += blackboardHeight + 10;
        
        // 绘制讲台区域
        if (params.deskLeftSit || params.deskRightSit) {
            const middleCol = Math.floor(params.cols / 2);
            
            for (let col = 0; col < params.cols; col++) {
                const x = padding + col * cellWidth;
                
                if (col === middleCol - 1 && params.deskLeftSit) {
                    // 左护法
                    ctx.fillStyle = this.seatingResult.left ? '#667eea' : '#4a5568';
                    ctx.fillRect(x, currentY, cellWidth, podiumHeight);
                    ctx.strokeStyle = '#667eea';
                    ctx.strokeRect(x, currentY, cellWidth, podiumHeight);
                    
                    ctx.fillStyle = 'white';
                    ctx.font = '12px Arial, sans-serif';
                    ctx.textAlign = 'center';
                    const text = this.seatingResult.left ? this.seatingResult.left.name : '左护法';
                    ctx.fillText(text, x + cellWidth/2, currentY + podiumHeight/2);
                } else if (col === middleCol) {
                    // 讲台
                    ctx.fillStyle = '#667eea';
                    ctx.fillRect(x, currentY, cellWidth, podiumHeight);
                    ctx.strokeStyle = '#667eea';
                    ctx.strokeRect(x, currentY, cellWidth, podiumHeight);
                    
                    ctx.fillStyle = 'white';
                    ctx.font = 'bold 14px Arial, sans-serif';
                    ctx.fillText('讲台', x + cellWidth/2, currentY + podiumHeight/2);
                } else if (col === middleCol + 1 && params.deskRightSit) {
                    // 右护法
                    ctx.fillStyle = this.seatingResult.right ? '#667eea' : '#4a5568';
                    ctx.fillRect(x, currentY, cellWidth, podiumHeight);
                    ctx.strokeStyle = '#667eea';
                    ctx.strokeRect(x, currentY, cellWidth, podiumHeight);
                    
                    ctx.fillStyle = 'white';
                    ctx.font = '12px Arial, sans-serif';
                    const text = this.seatingResult.right ? this.seatingResult.right.name : '右护法';
                    ctx.fillText(text, x + cellWidth/2, currentY + podiumHeight/2);
                }
            }
            
            currentY += podiumHeight + 10;
        }
        
        // 绘制学生座位
        ctx.font = '12px Arial, sans-serif';
        for (let row = 0; row < params.rows; row++) {
            for (let col = 0; col < params.cols; col++) {
                const x = padding + col * cellWidth;
                const y = currentY + row * cellHeight;
                const key = `${row},${col}`;
                const seatData = this.seatingResult[key];
                
                // 设置颜色
                if (seatData) {
                    if (seatData.type === 'corridor') {
                        ctx.fillStyle = '#2d3748';
                        ctx.fillRect(x, y, cellWidth, cellHeight);
                        ctx.fillStyle = 'white';
                        ctx.fillText('走廊', x + cellWidth/2, y + cellHeight/2);
                    } else if (seatData.type === 'empty') {
                        ctx.fillStyle = '#e2e8f0';
                        ctx.fillRect(x, y, cellWidth, cellHeight);
                        ctx.strokeStyle = '#cbd5e0';
                        ctx.setLineDash([5, 5]);
                        ctx.strokeRect(x, y, cellWidth, cellHeight);
                        ctx.setLineDash([]);
                    } else {
                        // 学生座位
                        if (seatData.gender === '男') {
                            ctx.fillStyle = '#c6f6d5';
                            ctx.strokeStyle = '#38a169';
                        } else if (seatData.gender === '女') {
                            ctx.fillStyle = '#fbb6ce';
                            ctx.strokeStyle = '#d53f8c';
                        } else {
                            ctx.fillStyle = '#bee3f8';
                            ctx.strokeStyle = '#3182ce';
                        }
                        
                        ctx.fillRect(x, y, cellWidth, cellHeight);
                        ctx.strokeRect(x, y, cellWidth, cellHeight);
                        
                        ctx.fillStyle = '#333';
                        ctx.fillText(seatData.name, x + cellWidth/2, y + cellHeight/2);
                    }
                } else {
                    // 空座位
                    ctx.fillStyle = '#e2e8f0';
                    ctx.fillRect(x, y, cellWidth, cellHeight);
                    ctx.strokeStyle = '#cbd5e0';
                    ctx.setLineDash([5, 5]);
                    ctx.strokeRect(x, y, cellWidth, cellHeight);
                    ctx.setLineDash([]);
                }
            }
        }
        
        // 下载图片
        canvas.toBlob(blob => {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = '座位安排表.png';
            link.click();
        }, 'image/png');
    }

    // 准备导出数据
    prepareExportData() {
        const params = this.getParameters();
        const data = [];
        
        // 添加讲台行
        if (params.deskLeftSit || params.deskRightSit) {
            const teacherRow = new Array(params.cols).fill('');
            const middle = Math.floor(params.cols / 2);
            
            teacherRow[middle] = '讲台';
            if (this.seatingResult.left && middle > 0) {
                teacherRow[middle - 1] = this.seatingResult.left.name;
            }
            if (this.seatingResult.right && middle < params.cols - 1) {
                teacherRow[middle + 1] = this.seatingResult.right.name;
            }
            
            data.push(teacherRow);
        }
        
        // 添加普通座位行
        for (let row = 0; row < params.rows; row++) {
            const rowData = [];
            for (let col = 0; col < params.cols; col++) {
                const key = `${row},${col}`;
                const seatData = this.seatingResult[key];
                
                if (seatData) {
                    if (seatData.type === 'corridor') {
                        rowData.push('走廊');
                    } else if (seatData.type === 'empty') {
                        rowData.push('');
                    } else {
                        rowData.push(seatData.name);
                    }
                } else {
                    rowData.push('');
                }
            }
            data.push(rowData);
        }
        
        return data;
    }

    // 显示加载动画
    showLoading() {
        document.getElementById('loadingOverlay').style.display = 'flex';
    }

    // 隐藏加载动画
    hideLoading() {
        document.getElementById('loadingOverlay').style.display = 'none';
    }

    // 显示Toast消息
    showToast(message, type = 'success') {
        const container = document.getElementById('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        
        container.appendChild(toast);
        
        // 3秒后自动移除
        setTimeout(() => {
            if (toast.parentNode) {
                toast.parentNode.removeChild(toast);
            }
        }, 3000);
    }
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    new SeatingArrangement();
});