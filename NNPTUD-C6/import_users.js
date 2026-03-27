const nodemailer = require('nodemailer');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

// Cấu hình nodemailer transporter sử dụng Mailtrap
const transporter = nodemailer.createTransport({
    host: "sandbox.smtp.mailtrap.io",
    port: 2525,
    auth: {
        user: "6a8e2e9654b287",
        pass: "1993b84301587f"
    }
});

const generatePassword = () => crypto.randomBytes(8).toString('hex');

const sendPasswordEmail = async (email, password) => {
    const mailOptions = {
        from: '"Hệ thống Admin" <admin@yourdomain.com>',
        to: email,
        subject: 'Thông tin tài khoản của bạn',
        text: `Email: ${email}\nPassword: ${password}`
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`[Thành công] Đã gửi thông tin cho: ${email}`);
    } catch (error) {
        console.error(`[Thất bại] Lỗi gửi email cho: ${email} - ${error.message}`);
    }
};

const processUsersFromExcel = async () => {
    const filePath = 'c:\\WORK\\NNPTDUM-27-3\\user.xlsx';
    console.log(`Đang phân tích dữ liệu từ Excel: ${filePath}`);

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(filePath);
    } catch (err) {
        console.error("Không thể đọc file Excel. Lỗi:", err.message);
        return;
    }

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
        console.error("File Excel không có sheet nào.");
        return;
    }

    // Duyệt từng dòng bắt đầu từ dòng 1
    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        
        let username = row.getCell(1).text;
        let email = row.getCell(2).text;

        // Xử lý an toàn nếu ô trống
        username = username ? username.trim() : '';
        email = email ? email.trim() : '';

        // Bỏ qua dòng trống hoặc dòng không có email
        if (!email) continue;
        
        // Nhận diện dòng tiêu đề (header) để bỏ qua
        if (email.toLowerCase() === 'email') continue;

        const password = generatePassword();
        const role = 'user'; // Đánh dấu cứng role theo yêu cầu

        console.log(`Đang chạy gửi email cho: ${username} (${email})...`);
        await sendPasswordEmail(email, password);

        // Chặn 0.5s giữa mỗi lần gửi để không bị Mailtrap chặn rate-limit
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    console.log("Hoàn thành quá trình khởi tạo cấu hình và gửi email qua Mailtrap từ file Excel.");
};

processUsersFromExcel().catch(err => {
    console.error("Có lỗi xảy ra:", err);
});
