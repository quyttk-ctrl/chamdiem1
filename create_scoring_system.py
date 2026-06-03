<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Chấm Điểm Bài Thi</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
        }

        .header p {
            font-size: 14px;
            opacity: 0.9;
        }

        .tabs {
            display: flex;
            border-bottom: 2px solid #e0e0e0;
            background: #f5f5f5;
        }

        .tab-button {
            flex: 1;
            padding: 15px;
            border: none;
            background: none;
            cursor: pointer;
            font-size: 16px;
            font-weight: 500;
            color: #666;
            transition: all 0.3s;
        }

        .tab-button.active {
            color: #667eea;
            border-bottom: 3px solid #667eea;
            background: white;
        }

        .tab-button:hover {
            background: white;
            color: #667eea;
        }

        .tab-content {
            display: none;
            padding: 30px;
            animation: fadeIn 0.3s;
        }

        .tab-content.active {
            display: block;
        }

        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }

        .form-group {
            margin-bottom: 20px;
        }

        .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
        }

        .form-group input,
        .form-group textarea {
            width: 100%;
            padding: 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 14px;
            transition: border-color 0.3s;
        }

        .form-group input:focus,
        .form-group textarea:focus {
            outline: none;
            border-color: #667eea;
        }

        .row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }

        .button-group {
            display: flex;
            gap: 10px;
            margin-top: 20px;
        }

        button {
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }

        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            flex: 1;
        }

        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }

        .btn-secondary {
            background: #f0f0f0;
            color: #333;
            flex: 1;
        }

        .btn-secondary:hover {
            background: #e0e0e0;
        }

        .table-container {
            overflow-x: auto;
            margin-top: 20px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
        }

        table thead {
            background: #667eea;
            color: white;
        }

        table th {
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }

        table td {
            padding: 12px;
            border-bottom: 1px solid #e0e0e0;
        }

        table tr:hover {
            background: #f5f5f5;
        }

        table input {
            width: 100%;
            padding: 6px;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
        }

        .info-box {
            background: #e3f2fd;
            border-left: 4px solid #667eea;
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 20px;
        }

        .info-box h3 {
            color: #1565c0;
            margin-bottom: 8px;
        }

        .info-box p {
            color: #0d47a1;
            font-size: 14px;
            line-height: 1.6;
        }

        .config-section {
            background: #f9f9f9;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        .config-section h3 {
            color: #333;
            margin-bottom: 15px;
            font-size: 16px;
        }

        .answer-input-row {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(50px, 1fr));
            gap: 8px;
            margin-bottom: 15px;
        }

        .answer-input-row input {
            padding: 8px;
            text-align: center;
            font-size: 12px;
        }

        .success {
            background: #c8e6c9;
            color: #2e7d32;
            padding: 15px;
            border-radius: 6px;
            margin-top: 20px;
        }

        .error {
            background: #ffcdd2;
            color: #c62828;
            padding: 15px;
            border-radius: 6px;
            margin-top: 20px;
        }

        .score-result {
            background: linear-gradient(135deg, #fff9c4 0%, #fff176 100%);
            padding: 20px;
            border-radius: 8px;
            margin-top: 20px;
            border: 2px solid #fdd835;
        }

        .score-result h3 {
            color: #f57f17;
            margin-bottom: 15px;
        }

        .score-item {
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid #fdd835;
        }

        .score-item:last-child {
            border-bottom: none;
        }

        .score-item strong {
            color: #333;
        }

        .score-value {
            background: white;
            padding: 5px 15px;
            border-radius: 4px;
            font-weight: bold;
            color: #f57f17;
        }

        @media (max-width: 768px) {
            .row {
                grid-template-columns: 1fr;
            }

            .tabs {
                flex-wrap: wrap;
            }

            .tab-button {
                flex: 1 1 auto;
                font-size: 14px;
                padding: 12px;
            }

            .answer-input-row {
                grid-template-columns: repeat(auto-fill, minmax(40px, 1fr));
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Hệ Thống Chấm Điểm Bài Thi</h1>
            <p>CLO 1: Câu 1-40 (0.25 điểm/câu) | CLO 2: Câu 41-55 (0.5 điểm/câu)</p>
        </div>

        <div class="tabs">
            <button class="tab-button active" onclick="switchTab('config')">⚙️ Cấu Hình</button>
            <button class="tab-button" onclick="switchTab('score')">✏️ Chấm Điểm</button>
            <button class="tab-button" onclick="switchTab('guide')">📖 Hướng Dẫn</button>
        </div>

        <!-- TAB 1: CẤU HÌNH -->
        <div id="config" class="tab-content active">
            <div class="info-box">
                <h3>ℹ️ Cấu Hình Đáp Án Theo Mã Đề</h3>
                <p>Nhập đáp án cho từng mã đề. Mỗi mã đề có 55 câu từ A, B, C, D</p>
            </div>

            <div class="config-section">
                <h3>Mã Đề 1</h3>
                <div class="form-group">
                    <label>Tên mã đề:</label>
                    <input type="text" id="codeA" placeholder="VD: ĐHXD-2024-A" value="ĐHXD-2024-A">
                </div>
                <label style="margin-bottom: 10px; font-weight: 600;">Đáp án 55 câu:</label>
                <div class="answer-input-row" id="answerA"></div>
            </div>

            <div class="config-section">
                <h3>Mã Đề 2</h3>
                <div class="form-group">
                    <label>Tên mã đề:</label>
                    <input type="text" id="codeB" placeholder="VD: ĐHXD-2024-B" value="ĐHXD-2024-B">
                </div>
                <label style="margin-bottom: 10px; font-weight: 600;">Đáp án 55 câu:</label>
                <div class="answer-input-row" id="answerB"></div>
            </div>

            <div class="button-group">
                <button class="btn-secondary" onclick="addMoreCodes()">+ Thêm Mã Đề Khác</button>
                <button class="btn-primary" onclick="saveConfig()">💾 Lưu Cấu Hình</button>
            </div>
            <div id="configMessage"></div>
        </div>

        <!-- TAB 2: CHẤM ĐIỂM -->
        <div id="score" class="tab-content">
            <div class="info-box">
                <h3>✏️ Nhập Thông Tin & Đáp Án Sinh Viên</h3>
                <p>Điền thông tin sinh viên, chọn mã đề, và dán đáp án để tính điểm tự động</p>
            </div>

            <div class="row">
                <div class="form-group">
                    <label>Mã sinh viên:</label>
                    <input type="text" id="studentId" placeholder="VD: 21010101">
                </div>
                <div class="form-group">
                    <label>Họ tên sinh viên:</label>
                    <input type="text" id="studentName" placeholder="VD: Nguyễn Văn A">
                </div>
            </div>

            <div class="form-group">
                <label>Chọn mã đề:</label>
                <select id="codeSelect" style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px;">
                    <option value="">-- Chọn mã đề --</option>
                    <option value="ĐHXD-2024-A">ĐHXD-2024-A</option>
                    <option value="ĐHXD-2024-B">ĐHXD-2024-B</option>
                </select>
            </div>

            <div class="form-group">
                <label>Dán đáp án 55 câu (cách nhau bằng dấu phẩy hoặc khoảng trắng):</label>
                <textarea id="studentAnswers" placeholder="VD: A B C D A B C D ... (55 câu)" rows="6" style="font-family: monospace;"></textarea>
            </div>

            <div class="button-group">
                <button class="btn-secondary" onclick="clearScoreForm()">🔄 Xóa</button>
                <button class="btn-primary" onclick="calculateScore()">🧮 Tính Điểm</button>
            </div>

            <div id="scoreResult"></div>
        </div>

        <!-- TAB 3: HƯỚNG DẪN -->
        <div id="guide" class="tab-content">
            <h2 style="margin-bottom: 20px; color: #333;">📖 Hướng Dẫn Sử Dụng</h2>

            <div class="config-section">
                <h3>1️⃣ Bước 1: Cấu Hình Đáp Án</h3>
                <p style="line-height: 1.8; color: #555;">
                    • Vào tab <strong>"⚙️ Cấu Hình"</strong><br>
                    • Nhập tên mã đề (VD: ĐHXD-2024-A)<br>
                    • Nhập 55 đáp án, mỗi câu 1 ô (A, B, C hoặc D)<br>
                    • Bấm <strong>"💾 Lưu Cấu Hình"</strong><br>
                    • Cấu hình được lưu trong trình duyệt (không mất khi reload)
                </p>
            </div>

            <div class="config-section">
                <h3>2️⃣ Bước 2: Chấm Điểm Sinh Viên</h3>
                <p style="line-height: 1.8; color: #555;">
                    • Vào tab <strong>"✏️ Chấm Điểm"</strong><br>
                    • Điền: Mã sinh viên, Họ tên<br>
                    • Chọn mã đề tương ứng<br>
                    • Dán đáp án sinh viên (55 câu)<br>
                    • Bấm <strong>"🧮 Tính Điểm"</strong><br>
                    • Hệ thống hiển thị kết quả ngay
                </p>
            </div>

            <div class="config-section">
                <h3>3️⃣ Bước 3: Xuất File Excel</h3>
                <p style="line-height: 1.8; color: #555;">
                    • Sau khi tính điểm, bấm nút <strong>"📥 Xuất File Excel"</strong><br>
                    • File Excel sẽ tự động download<br>
                    • File chứa tất cả thông tin chấm điểm
                </p>
            </div>

            <div class="config-section">
                <h3>📊 Tiêu Chí Chấm Điểm</h3>
                <table>
                    <tr style="background: #667eea; color: white;">
                        <td style="color: white;"><strong>CLO</strong></td>
                        <td style="color: white;"><strong>Câu</strong></td>
                        <td style="color: white;"><strong>Điểm/Câu</strong></td>
                        <td style="color: white;"><strong>Tổng Điểm</strong></td>
                    </tr>
                    <tr>
                        <td><strong>CLO 1</strong></td>
                        <td>Câu 1-40</td>
                        <td>0.25 điểm</td>
                        <td style="background: #fff9c4;"><strong>10 điểm</strong></td>
                    </tr>
                    <tr>
                        <td><strong>CLO 2</strong></td>
                        <td>Câu 41-55</td>
                        <td>0.5 điểm</td>
                        <td style="background: #fff9c4;"><strong>7.5 điểm</strong></td>
                    </tr>
                    <tr style="background: #e3f2fd;">
                        <td colspan="3" style="text-align: right;"><strong>Tổng Cộng:</strong></td>
                        <td style="background: #fff9c4;"><strong>17.5 điểm</strong></td>
                    </tr>
                </table>
            </div>

            <div class="config-section">
                <h3>💡 Mẹo Sử Dụng</h3>
                <p style="line-height: 1.8; color: #555;">
                    ✓ Dữ liệu cấu hình lưu trong trình duyệt → không mất dữ liệu khi reload<br>
                    ✓ Có thể chấm nhiều sinh viên liên tiếp<br>
                    ✓ Xuất file Excel để lưu trữ hoặc chia sẻ kết quả<br>
                    ✓ Nếu cần đổi mã đề, làm lại từ Bước 1
                </p>
            </div>
        </div>
    </div>

    <script>
        // Khởi tạo các ô nhập đáp án
        function initAnswerInputs() {
            const containerA = document.getElementById('answerA');
            const containerB = document.getElementById('answerB');
            
            containerA.innerHTML = '';
            containerB.innerHTML = '';
            
            for (let i = 1; i <= 55; i++) {
                containerA.innerHTML += `<input type="text" id="ans-a-${i}" maxlength="1" placeholder="${i}" style="text-transform: uppercase;">`;
                containerB.innerHTML += `<input type="text" id="ans-b-${i}" maxlength="1" placeholder="${i}" style="text-transform: uppercase;">`;
            }
        }

        // Chuyển tab
        function switchTab(tabName) {
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
        }

        // Lưu cấu hình
        function saveConfig() {
            const codeA = document.getElementById('codeA').value.trim();
            const codeB = document.getElementById('codeB').value.trim();
            
            if (!codeA || !codeB) {
                showMessage('configMessage', 'Vui lòng nhập tên mã đề!', 'error');
                return;
            }

            const answersA = [];
            const answersB = [];
            
            for (let i = 1; i <= 55; i++) {
                const a = document.getElementById(`ans-a-${i}`).value.toUpperCase();
                const b = document.getElementById(`ans-b-${i}`).value.toUpperCase();
                
                if (!a || !b) {
                    showMessage('configMessage', `Vui lòng điền đầy đủ 55 câu cho cả 2 mã đề!`, 'error');
                    return;
                }
                if (!'ABCD'.includes(a) || !'ABCD'.includes(b)) {
                    showMessage('configMessage', `Đáp án phải là A, B, C hoặc D!`, 'error');
                    return;
                }
                
                answersA.push(a);
                answersB.push(b);
            }

            const config = {
                [codeA]: answersA,
                [codeB]: answersB
            };

            localStorage.setItem('examConfig', JSON.stringify(config));
            localStorage.setItem('examCodes', JSON.stringify([codeA, codeB]));
            
            updateCodeSelect();
            showMessage('configMessage', '✅ Cấu hình đã được lưu thành công!', 'success');
        }

        // Cập nhật dropdown mã đề
        function updateCodeSelect() {
            const codes = JSON.parse(localStorage.getItem('examCodes') || '[]');
            const select = document.getElementById('codeSelect');
            
            select.innerHTML = '<option value="">-- Chọn mã đề --</option>';
            codes.forEach(code => {
                select.innerHTML += `<option value="${code}">${code}</option>`;
            });
        }

        // Tính điểm
        function calculateScore() {
            const studentId = document.getElementById('studentId').value.trim();
            const studentName = document.getElementById('studentName').value.trim();
            const selectedCode = document.getElementById('codeSelect').value;
            const answersText = document.getElementById('studentAnswers').value.trim();

            if (!studentId || !studentName || !selectedCode || !answersText) {
                showMessage('scoreResult', 'Vui lòng điền đầy đủ thông tin!', 'error');
                return;
            }

            const config = JSON.parse(localStorage.getItem('examConfig') || '{}');
            const correctAnswers = config[selectedCode];

            if (!correctAnswers) {
                showMessage('scoreResult', 'Mã đề không hợp lệ. Vui lòng kiểm tra cấu hình!', 'error');
                return;
            }

            // Parse đáp án sinh viên
            const studentAnswersArray = answersText.toUpperCase().split(/[\s,]+/).filter(a => a);

            if (studentAnswersArray.length !== 55) {
                showMessage('scoreResult', `Bạn nhập ${studentAnswersArray.length} câu, cần đúng 55 câu!`, 'error');
                return;
            }

            // Kiểm tra các ký tự hợp lệ
            for (let ans of studentAnswersArray) {
                if (!'ABCD'.includes(ans)) {
                    showMessage('scoreResult', `Đáp án "${ans}" không hợp lệ (chỉ nhận A, B, C, D)!`, 'error');
                    return;
                }
            }

            // Tính điểm
            let score1 = 0, score2 = 0;
            for (let i = 0; i < 40; i++) {
                if (studentAnswersArray[i] === correctAnswers[i]) score1 += 0.25;
            }
            for (let i = 40; i < 55; i++) {
                if (studentAnswersArray[i] === correctAnswers[i]) score2 += 0.5;
            }

            const totalScore = score1 + score2;

            // Hiển thị kết quả
            const result = `
                <div class="score-result">
                    <h3>📋 Kết Quả Chấm Điểm</h3>
                    <div style="margin-bottom: 15px; padding-bottom: 15px; border-bottom: 2px solid #fdd835;">
                        <p><strong>Mã SV:</strong> ${studentId}</p>
                        <p><strong>Họ tên:</strong> ${studentName}</p>
                        <p><strong>Mã đề:</strong> ${selectedCode}</p>
                    </div>
                    <div class="score-item">
                        <strong>CLO 1 (Câu 1-40):</strong>
                        <span class="score-value">${score1.toFixed(2)} / 10</span>
                    </div>
                    <div class="score-item">
                        <strong>CLO 2 (Câu 41-55):</strong>
                        <span class="score-value">${score2.toFixed(2)} / 7.5</span>
                    </div>
                    <div class="score-item" style="border-bottom: none; padding-top: 15px; border-top: 2px solid #fdd835; margin-top: 15px;">
                        <strong style="font-size: 18px;">TỔNG ĐIỂM:</strong>
                        <span class="score-value" style="font-size: 18px;">${totalScore.toFixed(2)} / 17.5</span>
                    </div>
                    <button class="btn-primary" onclick="exportToExcel('${studentId}', '${studentName}', '${selectedCode}', ${score1.toFixed(2)}, ${score2.toFixed(2)}, ${totalScore.toFixed(2)})" style="margin-top: 20px; width: 100%;">📥 Xuất File Excel</button>
                </div>
            `;

            document.getElementById('scoreResult').innerHTML = result;
        }

        // Xuất file Excel
        function exportToExcel(studentId, studentName, code, clo1, clo2, total) {
            const ws = XLSX.utils.aoa_to_sheet([
                ['HỆ THỐNG CHẤM ĐIỂM BÀI THI'],
                [],
                ['Mã sinh viên:', studentId],
                ['Họ tên:', studentName],
                ['Mã đề:', code],
                [],
                ['CLO', 'Điểm', 'Tổng'],
                ['CLO 1 (Câu 1-40)', clo1, '10'],
                ['CLO 2 (Câu 41-55)', clo2, '7.5'],
                ['TỔNG ĐIỂM', total, '17.5']
            ]);

            ws['A1'].font = { bold: true, size: 14 };
            ws['A7'].font = { bold: true };
            ws['A10'].font = { bold: true };

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Kết quả');
            XLSX.writeFile(wb, `ChamDiem_${studentId}_${Date.now()}.xlsx`);
        }

        // Xóa form
        function clearScoreForm() {
            document.getElementById('studentId').value = '';
            document.getElementById('studentName').value = '';
            document.getElementById('codeSelect').value = '';
            document.getElementById('studentAnswers').value = '';
            document.getElementById('scoreResult').innerHTML = '';
        }

        // Thêm mã đề khác
        function addMoreCodes() {
            alert('Tính năng này sẽ được cập nhật. Hiện tại chỉ hỗ trợ 2 mã đề.');
        }

        // Hiển thị thông báo
        function showMessage(elementId, message, type) {
            const element = document.getElementById(elementId);
            element.innerHTML = `<div class="${type}">${message}</div>`;
            setTimeout(() => {
                element.innerHTML = '';
            }, 5000);
        }

        // Khởi tạo khi load trang
        window.addEventListener('load', () => {
            initAnswerInputs();
            updateCodeSelect();
        });
    </script>
</body>
</html>
