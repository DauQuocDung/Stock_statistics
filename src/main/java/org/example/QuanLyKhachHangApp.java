package org.example;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.time.LocalDateTime;

import java.awt.Desktop;
import java.io.File;
import java.time.format.DateTimeFormatter;



class StartFrame extends JFrame {

    private JLabel timeLabel;
    private Image bgImage;
    private int width = 800;
    private int height = 500;

    private String getAppDirectory() {
        try {
            String jarPath = StartFrame.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI().getPath();
            File jarFile = new File(jarPath);
            return jarFile.getParentFile().getAbsolutePath();
        } catch (Exception e) {
            return System.getProperty("user.dir");
        }
    }

    private String imagePath = getAppDirectory() + File.separator + "start.png";

    public StartFrame() {

        // ===== LOAD ẢNH AN TOÀN =====
        ImageIcon icon = new ImageIcon(imagePath);

        if (icon.getIconWidth() > 0) {
            bgImage = icon.getImage();
            width = icon.getIconWidth();
            height = icon.getIconHeight();
        }

        setTitle("Quản Lý Chứng Khoán");
        setSize(width, height);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setResizable(false);

        JPanel panel = new JPanel() {
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);

                if (bgImage != null) {
                    g.drawImage(bgImage, 0, 0, null);
                }

                // overlay nhẹ
                g.setColor(new Color(0, 0, 0, 60));
                g.fillRect(0, 0, getWidth(), getHeight());
            }
        };

        panel.setLayout(null);

        // ===== LABEL TIME (FIX VẼ CHUẨN) =====
        timeLabel = new JLabel() {
            protected void paintComponent(Graphics g) {
                super.paintComponent(g); // QUAN TRỌNG

                Graphics2D g2 = (Graphics2D) g;


                // text chính
                g2.setColor(Color.WHITE);
                g2.drawString(getText(), 0, 16);
            }
        };

        timeLabel.setFont(new Font("Segoe UI", Font.BOLD, 16));
        timeLabel.setBounds(20, 30, 450, 30);

        panel.add(timeLabel);
        startClock();

        // ===== BUTTON =====
        JButton startBtn = new JButton("BẮT ĐẦU") {
            protected void paintComponent(Graphics g) {

                Graphics2D g2 = (Graphics2D) g;
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                        RenderingHints.VALUE_ANTIALIAS_ON);

                GradientPaint gp = new GradientPaint(
                        0, 0, new Color(46, 204, 113),
                        0, getHeight(), new Color(39, 174, 96)
                );

                g2.setPaint(gp);
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 30, 30);

                super.paintComponent(g);
            }
        };

        startBtn.setFont(new Font("Segoe UI", Font.BOLD, 16));
        startBtn.setForeground(Color.WHITE);
        startBtn.setContentAreaFilled(false);
        startBtn.setBorderPainted(false);
        startBtn.setFocusPainted(false);
        startBtn.setCursor(new Cursor(Cursor.HAND_CURSOR));

        startBtn.setBounds(width - 150, 25, 120, 40);

        startBtn.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent e) {
                startBtn.setLocation(width - 152, 23);
            }

            public void mouseExited(MouseEvent e) {
                startBtn.setLocation(width - 150, 25);
            }
        });

        // ===== FIX QUAN TRỌNG NHẤT =====
        startBtn.addActionListener(e -> {
            try {
                new QuanLyKhachHangApp().setVisible(true);
                dispose();
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(null,
                        "Lỗi mở màn hình chính:\n" + ex.getMessage());
            }
        });

        panel.add(startBtn);

        add(panel);
        setVisible(true);
    }

    private void startClock() {
        DateTimeFormatter formatter =
                DateTimeFormatter.ofPattern("dd/MM - HH:mm:ss");

        Timer timer = new Timer(1000, e -> {
            LocalDateTime now = LocalDateTime.now();
            timeLabel.setText("Xin chào anh trai, hôm nay " + now.format(formatter));
        });

        timer.start();
    }

}
public class QuanLyKhachHangApp extends JFrame {


        private String getAppDirectory() {
            try {
                String jarPath = QuanLyKhachHangApp.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
                File jarFile = new File(jarPath);
                return jarFile.getParentFile().getAbsolutePath();
            } catch (Exception e) {
                e.printStackTrace();
                return System.getProperty("user.dir"); // fallback
            }
        }

        private final String basePath = System.getProperty("user.dir");

        private final String fileExcel = basePath + File.separator + "QuanLyChungKhoan.xlsx";
        private final String templatePath = basePath + File.separator + "template_giaodich.xlsx";

        private Workbook openWorkbookSafely(InputStream is) throws IOException {
            ZipSecureFile.setMinInflateRatio(0.001);
            return WorkbookFactory.create(is);
        }

        private JTable tableKhachHang;
        private DefaultTableModel model;

        private void moFileExcel() {
            try {
                File file = new File(fileExcel);
                if (file.exists()) {
                    Desktop.getDesktop().open(file);
                } else {
                    JOptionPane.showMessageDialog(this, "File Excel chưa tồn tại: " + fileExcel);
                }
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Lỗi mở file Excel: " + e.getMessage());
            }
        }

        public QuanLyKhachHangApp() {
            setTitle("Quản Lý Giao Dịch Chứng Khoán");
            setSize(900, 600);
            setDefaultCloseOperation(EXIT_ON_CLOSE);
            setLocationRelativeTo(null);


            // ================= DANH SÁCH KHÁCH HÀNG =================
            String[] columns = {"Tên Khách Hàng", "SĐT", "Tài Khoản"};
            model = new DefaultTableModel(columns, 0);
            tableKhachHang = new JTable(model);
            tableKhachHang.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

            DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
            centerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
            for (int i = 0; i < tableKhachHang.getColumnCount(); i++) {
                tableKhachHang.getColumnModel().getColumn(i).setCellRenderer(centerRenderer);
            }

            JScrollPane scroll = new JScrollPane(tableKhachHang);
            add(scroll, BorderLayout.CENTER);

            // ================= BUTTON =================
            JPanel panelButton = new JPanel();
            JButton btnThem = new JButton("Thêm Khách Hàng Mới");
            JButton btnGiaoDich = new JButton("Giao Dịch");
            JButton btnSua = new JButton("Sửa Thông Tin KH");
            JButton btnMoFile = new JButton("Mở File Excel");
            JButton xoaKhachHangBtn = new JButton("Xóa Khách Hàng");
            JButton thoatBtn = new JButton("Đóng");

            panelButton.add(btnThem);
            panelButton.add(btnSua);
            panelButton.add(xoaKhachHangBtn);
            panelButton.add(btnGiaoDich);
            panelButton.add(btnMoFile);
            panelButton.add(thoatBtn);
            add(panelButton, BorderLayout.SOUTH);

            // Load dữ liệu
            loadDanhSachKhachHang();

            // ================= EVENT =================
            btnThem.addActionListener(e -> themKhachHang());
            btnGiaoDich.addActionListener(e -> moGiaoDich());
            btnSua.addActionListener(e -> suaThongTinKH());
            btnMoFile.addActionListener(e -> moFileExcel());
//        xoaKhachHangBtn.addActionListener(e -> xoaKhachHang());
            thoatBtn.addActionListener(e -> {
                int result = JOptionPane.showConfirmDialog(this, "Anh chắc chắn muốn đóng chứ?", "Xác nhận", JOptionPane.YES_NO_OPTION);
                if (result == JOptionPane.YES_OPTION) {
                    System.exit(0);
                }
            });

            setVisible(true);
        }

        private String getStringCell(Cell cell) {
            if (cell == null) return "";
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    return String.valueOf((long) cell.getNumericCellValue());
                default:
                    return "";
            }
        }

        // ================= LOAD DANH SÁCH KHÁCH HÀNG =================
        private void loadDanhSachKhachHang() {
            model.setRowCount(0);
            File file = new File(fileExcel);
            if (!file.exists()) {
                JOptionPane.showMessageDialog(this, "File Excel chưa tồn tại: " + fileExcel + "\nĐang tạo mới...");
                taoFileMoi();
            }

            try (FileInputStream fis = new FileInputStream(file);
                 Workbook wb = openWorkbookSafely(fis)) {

                Sheet sheet = wb.getSheet("DanhSachKhachHang");
                if (sheet == null) {
                    JOptionPane.showMessageDialog(this, "Không tìm thấy sheet DanhSachKhachHang trong file!");
                    return;
                }

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row r = sheet.getRow(i);
                    if (r == null) continue;

                    String ten = getStringCell(r.getCell(0));
                    String sdt = getStringCell(r.getCell(1));
                    String taiKhoan = getStringCell(r.getCell(2));

                    model.addRow(new Object[]{ten, sdt, taiKhoan});
                }
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Lỗi khi load danh sách: " + e.getMessage());
            }
        }

        private void taoFileMoi() {
            try (Workbook wb = new XSSFWorkbook()) {
                Sheet ds = wb.createSheet("DanhSachKhachHang");
                Row header = ds.createRow(0);
//                header.createCell(0).setCellValue("Mã KH");
                header.createCell(0).setCellValue("Tên Khách Hàng");
//                header.createCell(1).setCellValue("Tên Khách Hàng");
                header.createCell(1).setCellValue("SĐT");
//                header.createCell(3).setCellValue("Trạng Thái");
                header.createCell(2).setCellValue("Tài Khoản");

                try (FileOutputStream fos = new FileOutputStream(fileExcel)) {
                    wb.write(fos);
                }
                JOptionPane.showMessageDialog(this, "Đã tạo file Excel mới tại: " + fileExcel);
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Lỗi tạo file Excel: " + e.getMessage());
            }
        }

        // ================= THÊM KHÁCH HÀNG =================
        // ================= SUA TT KHÁCH HÀNG =================
        private void suaThongTinKH() {
            int row = tableKhachHang.getSelectedRow();
            if (row == -1) {
                JOptionPane.showMessageDialog(this, "Chọn một khách hàng trước!");
                return;
            }

            String tenCu = model.getValueAt(row, 0).toString(); // Tên cũ để so sánh

            // Dialog sửa thông tin
            String tenMoi = JOptionPane.showInputDialog(this, "Nhập Tên Khách Hàng mới:", tenCu);
            while (tenMoi == null || tenMoi.trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "Tên khách hàng không được bỏ trống !!!", "Warning", JOptionPane.ERROR_MESSAGE);
                tenMoi = JOptionPane.showInputDialog(this, "Nhập Tên Khách Hàng mới:", tenCu);
            }

            String sdtMoi = JOptionPane.showInputDialog(this, "Nhập SĐT mới:", model.getValueAt(row, 1).toString());
            if (sdtMoi == null) sdtMoi = ""; // Nếu hủy thì giữ cũ hoặc để trống

            String taiKhoanMoi = JOptionPane.showInputDialog(this, "Nhập tài khoản mới của khách hàng:", model.getValueAt(row, 2).toString());
            if (taiKhoanMoi == null) taiKhoanMoi = "";
            try (FileInputStream fis = new FileInputStream(fileExcel);
                 Workbook wb = WorkbookFactory.create(fis)) {

                // 1. Cập nhật sheet DanhSachKhachHang
                Sheet ds = wb.getSheet("DanhSachKhachHang");
                if (ds != null) {
                    Row rowDS = ds.getRow(row + 1); // +1 vì row 0 là header
                    if (rowDS != null) {
                        rowDS.getCell(0).setCellValue(tenMoi.trim());
                        rowDS.getCell(1).setCellValue(sdtMoi.trim());
                        rowDS.getCell(2).setCellValue(taiKhoanMoi.trim());
                    }
                }

                // 2. Đổi tên sheet giao dịch cũ
                String tenSheetCu = tenCu;
                String tenSheetMoi =tenMoi.trim();

                Sheet sheetGiaoDich = wb.getSheet(tenSheetCu);
                if (sheetGiaoDich != null) {
                    wb.setSheetName(wb.getSheetIndex(sheetGiaoDich), tenSheetMoi);
                } else {
                    JOptionPane.showMessageDialog(this, "Không tìm thấy sheet giao dịch cũ: " + tenSheetCu);
                }

                // 3. Lưu file
                try (FileOutputStream fos = new FileOutputStream(fileExcel)) {
                    wb.write(fos);
                }

                // 4. Cập nhật lại bảng danh sách
                model.setValueAt(tenMoi.trim(), row, 0);
                model.setValueAt(sdtMoi.trim(), row, 1);
                model.setValueAt(taiKhoanMoi.trim(), row, 2);
                JOptionPane.showMessageDialog(this, "Cập nhật thông tin khách hàng thành công!");

            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Lỗi khi sửa thông tin: " + e.getMessage());
            }
        }

        private void themKhachHang() {
//            String maKH = JOptionPane.showInputDialog("Nhập Mã KH:");
//            if (maKH == null || maKH.trim().isEmpty()) return;

            String ten = JOptionPane.showInputDialog("Nhập Tên Khách Hàng:");
            String sdt = JOptionPane.showInputDialog("Nhập SĐT:");
//            String trangThai = JOptionPane.showInputDialog("Nhập Trạng Thái Khách Hàng:");
            String taiKhoan = JOptionPane.showInputDialog("Nhập Tài Khoản Khách Hàng:");
            try (FileInputStream fis = new FileInputStream(fileExcel);
                 Workbook wb = openWorkbookSafely(fis)) {

                Sheet ds = wb.getSheet("DanhSachKhachHang");
                int last = ds.getLastRowNum() + 1;
                Row row = ds.createRow(last);
//                row.createCell(0).setCellValue(maKH.trim());
                row.createCell(0).setCellValue(ten);
                row.createCell(1).setCellValue(sdt);
//                row.createCell(3).setCellValue(trangThai);
                row.createCell(2).setCellValue(taiKhoan);

                // Tạo sheet giao dịch mới từ template
                taoSheetGiaoDich(wb, ten);

                try (FileOutputStream fos = new FileOutputStream(fileExcel)) {
                    wb.write(fos);
                }
                loadDanhSachKhachHang();
                JOptionPane.showMessageDialog(this, "Thêm khách hàng thành công!");
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Lỗi khi thêm khách hàng!");
            }
        }

        private void taoSheetGiaoDich(Workbook wb, String sheetName) {
            Sheet newSheet = wb.createSheet(sheetName);

            try (FileInputStream fis = new FileInputStream(templatePath);
                 Workbook templateWb = openWorkbookSafely(fis)) {

                Sheet templateSheet = templateWb.getSheetAt(0);
                for (int i = 0; i <= templateSheet.getLastRowNum(); i++) {
                    Row srcRow = templateSheet.getRow(i);
                    if (srcRow == null) continue;
                    Row destRow = newSheet.createRow(i);

                    for (int j = 0; j < srcRow.getLastCellNum(); j++) {
                        Cell srcCell = srcRow.getCell(j);
                        if (srcCell == null) continue;
                        Cell destCell = destRow.createCell(j);

                        switch (srcCell.getCellType()) {
                            case STRING:
                                destCell.setCellValue(srcCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                destCell.setCellValue(srcCell.getNumericCellValue());
                                break;
                        }

                        CellStyle newStyle = wb.createCellStyle();
                        newStyle.cloneStyleFrom(srcCell.getCellStyle());
                        destCell.setCellStyle(newStyle);
                    }
                }
            } catch (Exception ex) {
                Row header = newSheet.createRow(0);
                String[] headers = {
                        "Thời Gian", "Loại", "Mã CP", "Giá Mua", "SL Mua",
                        "Tổng Mua", "Phí Mua", "Total Mua",
                        "SL Còn Lại", "Giá TB",
                        "SL Bán", "Giá Bán", "Tổng Bán", "Phí Bán",
                        "Thuế", "Total Bán", "Lợi Nhuận",
                        "Tổng phí" // 👈 thêm cột này (index = 17)
                };
                for (int i = 0; i < headers.length; i++) {
                    header.createCell(i).setCellValue(headers[i]);
                }
                JOptionPane.showMessageDialog(this, "Không tìm thấy template, tạo sheet mặc định!");
            }
        }

        // ================= MỞ DIALOG GIAO DỊCH =================
        private void moGiaoDich() {
            int row = tableKhachHang.getSelectedRow();
            if (row == -1) {
                JOptionPane.showMessageDialog(this, "Chọn một khách hàng trước!");
                return;
            }

//            String maKH = model.getValueAt(row, 0).toString();
            String tenKH = model.getValueAt(row, 0).toString();
//            String sheetName = maKH + " - " + tenKH;
            String sheetName = tenKH;

            new GiaoDichDialog(this, sheetName, fileExcel).setVisible(true);
        }

        // ================= CLASS DIALOG GIAO DỊCH (gộp vào 1 file) =================
        private class GiaoDichDialog extends JDialog {

            private JTextField maMuaField = new JTextField(8);
            private JTextField giaMuaField = new JTextField(8);
            private JTextField slMuaField = new JTextField(8);

            private JTextField maBanField = new JTextField(8);
            private JTextField giaBanField = new JTextField(8);
            private JTextField slBanField = new JTextField(8);

            private final String sheetName;
            private final String excelFile;


            public GiaoDichDialog(Frame owner, String sheetName, String excelFile) {
                super(owner, "Giao Dịch - " + sheetName, true);
                this.sheetName = sheetName;
                this.excelFile = excelFile;

                setSize(600, 300);
                setLayout(new GridLayout(3, 1));
                setLocationRelativeTo(owner);

                JPanel muaPanel = new JPanel();
                muaPanel.setBorder(BorderFactory.createTitledBorder("MUA"));
                muaPanel.add(new JLabel("Mã:"));
                muaPanel.add(maMuaField);
                muaPanel.add(new JLabel("Giá:"));
                muaPanel.add(giaMuaField);
                muaPanel.add(new JLabel("SL:"));
                muaPanel.add(slMuaField);

                JButton muaBtn = new JButton("Mua");
                muaPanel.add(muaBtn);

                JPanel banPanel = new JPanel();
                banPanel.setBorder(BorderFactory.createTitledBorder("BÁN"));
                banPanel.add(new JLabel("Mã:"));
                banPanel.add(maBanField);
                banPanel.add(new JLabel("Giá:"));
                banPanel.add(giaBanField);
                banPanel.add(new JLabel("SL:"));
                banPanel.add(slBanField);

                JButton banBtn = new JButton("Bán");
                banPanel.add(banBtn);


                JPanel menuPanel = new JPanel();
                menuPanel.setBorder(BorderFactory.createTitledBorder("Menu"));

                JButton moFileBtn = new JButton("Mở File Excel");
                menuPanel.add(moFileBtn);

                JButton xongBtn = new JButton("Xong");
                menuPanel.add(xongBtn);


                add(muaPanel);
                add(banPanel);
                add(menuPanel);

                // Nếu không có template, tạo header thủ công

                muaBtn.addActionListener(e -> {
                    try {
                        ghiGiaoDich("MUA",
                                maMuaField.getText().trim(),
                                Double.parseDouble(giaMuaField.getText()),
                                Integer.parseInt(slMuaField.getText()));
                    } catch (NumberFormatException ex) {
                        JOptionPane.showMessageDialog(this, "Vui lòng nhập số hợp lệ!");
                    }
                });

                banBtn.addActionListener(e -> {
                    try {
                        ghiGiaoDich("BÁN",
                                maBanField.getText().trim(),
                                Double.parseDouble(giaBanField.getText()),
                                Integer.parseInt(slBanField.getText()));
                    } catch (NumberFormatException ex) {
                        JOptionPane.showMessageDialog(this, "Vui lòng nhập số hợp lệ!");
                    }
                });

                moFileBtn.addActionListener(e -> {
                    moFileExcel();
                });

                xongBtn.addActionListener(e -> {
                    this.dispose();
                });
            }

            // ================= GHI GIAO DỊCH =================
            private void ghiGiaoDich(String loai, String ma, double gia, int sl) {
                try {
                    Result result = tinhToan(ma, loai, gia, sl);
                    if (result == null) return;

                    File file = new File(excelFile);
                    if (!file.exists()) {
                        JOptionPane.showMessageDialog(this, "Không tìm thấy file Excel!");
                        return;
                    }

                    try (FileInputStream fis = new FileInputStream(file);
                         Workbook workbook = openWorkbookSafely(fis)) {

                        Sheet sheet = workbook.getSheet(sheetName);
                        if (sheet == null) {
                            JOptionPane.showMessageDialog(this, "Không tìm thấy sheet " + sheetName + "!");
                            return;
                        }

                        XSSFWorkbook xssf = (XSSFWorkbook) workbook;

                        XSSFCellStyle crossStyle = createCrossStyle(xssf);
                        XSSFCellStyle highlightStyle = createHighlightStyle(xssf);

                        XSSFCellStyle whiteStyle = xssf.createCellStyle();
                        whiteStyle.setFillPattern(FillPatternType.NO_FILL);
                        whiteStyle.setBorderTop(BorderStyle.THIN);
                        whiteStyle.setBorderBottom(BorderStyle.THIN);
                        whiteStyle.setBorderLeft(BorderStyle.THIN);
                        whiteStyle.setBorderRight(BorderStyle.THIN);
                        whiteStyle.setAlignment(HorizontalAlignment.CENTER);
                        whiteStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                        XSSFCellStyle moneyStyle = xssf.createCellStyle();
                        moneyStyle.cloneStyleFrom(whiteStyle);
                        moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));

                        XSSFCellStyle dateStyle = xssf.createCellStyle();
                        dateStyle.cloneStyleFrom(whiteStyle);
                        dateStyle.setDataFormat(workbook.createDataFormat().getFormat("dd/MM/yyyy"));

                        XSSFCellStyle muaStyle = xssf.createCellStyle();
                        muaStyle.cloneStyleFrom(whiteStyle);
                        org.apache.poi.ss.usermodel.Font muaFont = xssf.createFont();
                        muaFont.setColor(IndexedColors.GREEN.getIndex());
                        muaFont.setBold(true);
                        muaStyle.setFont(muaFont);

                        XSSFCellStyle banStyle = xssf.createCellStyle();
                        banStyle.cloneStyleFrom(whiteStyle);
                        org.apache.poi.ss.usermodel.Font banFont = xssf.createFont();
                        banFont.setColor(IndexedColors.RED.getIndex());
                        banFont.setBold(true);
                        banStyle.setFont(banFont);

                        // ================= TÍNH lastDataRow TRƯỚC KHI TẠO DÒNG MỚI =================
                        int lastDataRow = 0;
                        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                            Row r = sheet.getRow(i);
                            if (r != null && !getStringCell(r.getCell(2)).trim().isEmpty()) {
                                lastDataRow = i;
                            }
                        }

                        Row newRow = sheet.createRow(lastDataRow + 1);

                        // Set style mặc định cho dòng mới
                        for (int col = 0; col <= 17; col++) {
                            Cell cell = newRow.createCell(col);
                            cell.setCellStyle(whiteStyle);
                            if (col == 0) {
                                cell.setCellStyle(dateStyle);
                            } else if (col == 3 || col == 5 || col == 6 || col == 7 || col == 9 ||
                                    col == 11 || col == 12 || col == 13 || col == 14 || col == 15 || col == 16) {
                                cell.setCellStyle(moneyStyle);
                            }
                        }

                        // ================= GHI DỮ LIỆU CHÍNH =================
                        LocalDateTime now = LocalDateTime.now();
                        newRow.getCell(0).setCellValue(now.toLocalDate().atStartOfDay());

                        Cell loaiCell = newRow.getCell(1);
                        loaiCell.setCellValue(loai);
                        if ("MUA".equals(loai)) {
                            loaiCell.setCellStyle(muaStyle);
                        } else {
                            loaiCell.setCellStyle(banStyle);
                        }

                        newRow.getCell(2).setCellValue(ma.trim().toUpperCase());

                        if ("MUA".equals(loai)) {
                            //Get tong phi moi


                            newRow.getCell(3).setCellValue(gia);
                            newRow.getCell(4).setCellValue(sl);
                            newRow.getCell(5).setCellValue(result.tongTienMua);
                            newRow.getCell(6).setCellValue(result.phiMua);
                            newRow.getCell(7).setCellValue(result.totalMua);

                            for (int i = 10; i <= 16; i++) {
                                crossCell(newRow, i, crossStyle);
                            }

                            newRow.getCell(8).setCellValue(result.slConLai);   // SL Còn Lại
                            newRow.getCell(9).setCellValue(result.giaTB);

                        } else {  // BÁN
                            newRow.getCell(10).setCellValue(sl);
                            newRow.getCell(11).setCellValue(gia);
                            newRow.getCell(12).setCellValue(result.P_tongBan);
                            newRow.getCell(13).setCellValue(result.phiBan);
                            newRow.getCell(14).setCellValue(result.thueBan);
                            newRow.getCell(15).setCellValue(result.totalBan);
                            newRow.getCell(16).setCellValue(result.loiNhuan);

                            for (int i = 3; i <= 7; i++) {
                                crossCell(newRow, i, crossStyle);
                            }

                            newRow.getCell(8).setCellValue(result.slConLai);   // SL Còn Lại
                            newRow.getCell(9).setCellValue(result.giaTB);
                        }

                        // ================= XỬ LÝ TỔNG PHÍ - KHÔNG BỊ TRỄ NỮA =================


                        // Xóa Tổng Phí ở tất cả hàng cũ của mã CP này
                        for (int i = 1; i <= lastDataRow; i++) {
                            Row r = sheet.getRow(i);
                            if (r == null) continue;
                            if (getStringCell(r.getCell(2)).equals(ma.trim().toUpperCase())) {
                                crossCell(r, 17, crossStyle);
                            }
                        }

                        // Ghi Tổng Phí vào dòng mới nhất
                        Cell tongPhiCell = newRow.getCell(17);
                        if (tongPhiCell == null) {
                            tongPhiCell = newRow.createCell(17);
                        }
                        Cell giaTbCell = newRow.getCell(9);

                        tongPhiCell.setCellValue(result.tongPhi);
                        tongPhiCell.setCellStyle(moneyStyle);
// ===== RESET highlight cũ của mã CP =====
                        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                            Row r = sheet.getRow(i);
                            if (r == null) continue;

                            String currentMa = getStringCell(r.getCell(2));
                            if (currentMa.equals(ma.trim().toUpperCase())) {

                                Cell oldGiaTb = r.getCell(9);
                                if (oldGiaTb != null) {
                                    oldGiaTb.setCellStyle(moneyStyle); // trả về style bình thường
                                }

                                Cell oldTongPhi = r.getCell(17);
                                if (oldTongPhi != null) {
                                    oldTongPhi.setCellStyle(moneyStyle);
                                }
                            }
                        }
                        if (result.slConLai > 0) {
                            XSSFCellStyle hl = (XSSFCellStyle) workbook.createCellStyle();
                            hl.cloneStyleFrom(moneyStyle);
                            hl.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            hl.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            giaTbCell.setCellStyle(hl);
                        }
                        XSSFCellStyle hlp = (XSSFCellStyle) workbook.createCellStyle();
                        hlp.cloneStyleFrom(moneyStyle);
                        hlp.setFillForegroundColor(new XSSFColor(
                                new java.awt.Color(205, 144, 111), null
                        ));
                        hlp.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        hlp.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        tongPhiCell.setCellStyle(hlp);


                        // ================= LƯU FILE =================
                        try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                            workbook.write(fos);
                        }

                        JOptionPane.showMessageDialog(this, "Ghi giao dịch thành công!");
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(this, "Lỗi khi ghi giao dịch: " + e.getMessage());
                }
            }
            private double getTongPhi(String ma) {
                double tongPhiCu = 0.0;
                try (FileInputStream fis = new FileInputStream(excelFile);
                     Workbook wb = openWorkbookSafely(fis)) {

                    Sheet sheet = wb.getSheet(sheetName);
                    if (sheet == null) return 0.0;
                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {   // lấy hết, bao gồm dòng mới
                        Row r = sheet.getRow(i);
                        if (r == null) continue;

                        if (getStringCell(r.getCell(2)).equals(ma.trim().toUpperCase())) {
                            tongPhiCu = getNumericCell(r.getCell(17));  // Phí Mua
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return tongPhiCu;
            }
            // ================= CÁC HÀM PHỤ =================
            private double getNumericCell(Cell cell) {
                if (cell == null) return 0;
                switch (cell.getCellType()) {
                    case NUMERIC:
                        return cell.getNumericCellValue();
                    case STRING:
                        try {
                            return Double.parseDouble(cell.getStringCellValue().trim());
                        } catch (Exception e) {
                            return 0;
                        }
                    case FORMULA:
                        return cell.getNumericCellValue();
                    default:
                        return 0;
                }
            }

            private String getStringCell(Cell cell) {
                if (cell == null) return "";
                switch (cell.getCellType()) {
                    case STRING:
                        return cell.getStringCellValue().trim();
                    case NUMERIC:
                        return String.valueOf((long) cell.getNumericCellValue());
                    default:
                        return "";
                }
            }

            private XSSFCellStyle createCrossStyle(XSSFWorkbook workbook) {
                XSSFCellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setAlignment(HorizontalAlignment.CENTER);
                return style;
            }

            private XSSFCellStyle createHighlightStyle(XSSFWorkbook workbook) {
                XSSFCellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                return style;
            }

            private void crossCell(Row row, int colIndex, CellStyle style) {
                Cell cell = row.createCell(colIndex);
                cell.setCellStyle(style);
            }

            class Result {
                double tongTienMua, phiMua, totalMua;
                double P_tongBan, phiBan, thueBan, totalBan;
                double slConLai, giaTB, loiNhuan;

                double tongPhi;
            }

            private double getSoLuong(String ma) {
                double slConLai = 0;
                File file = new File(excelFile);
                if (!file.exists()) return 0;

                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = WorkbookFactory.create(fis)) {

                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet == null) return 0;

                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row r = sheet.getRow(i);
                        if (r == null) continue;

                        String currentMa = getStringCell(r.getCell(2));
                        if (currentMa.equals(ma)) {
                            String loai = getStringCell(r.getCell(1));
                            double slMua = getNumericCell(r.getCell(4));
                            double slBan = getNumericCell(r.getCell(10));

                            if ("MUA".equals(loai)) {
                                slConLai += slMua;
                            } else if ("BÁN".equals(loai)) {
                                slConLai -= slBan;
                            }
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

                return Math.max(slConLai, 0);
            }

            private double getGiaTB(String ma) {

                File file = new File(excelFile);
                if (!file.exists()) return 0;

                double currentGiaTB = 0;
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = WorkbookFactory.create(fis)) {

                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet == null) return 0;
                    for (int i = sheet.getLastRowNum(); i >= 1; i--) {
                        Row r = sheet.getRow(i);
                        if (r == null) continue;
                        String currentMa = getStringCell(r.getCell(2));
                        if (currentMa.equals(ma)) {
                            String loai = getStringCell(r.getCell(1));
                            if ("MUA".equals(loai)) {
                                currentGiaTB = getNumericCell(r.getCell(9));
                                return currentGiaTB;
                            }
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

                return currentGiaTB;
            }

            private Result tinhToan(String ma, String loai, double gia, int sl) {
                Result result = new Result();
                double slHienTai = getSoLuong(ma);
                double giaTBHienTai = getGiaTB(ma);
                double tongPhiCu = getTongPhi(ma);

                if (loai.equals("MUA")) {
                    double tongTienMua = gia * sl;
                    double phiMua = tongTienMua * 0.0015;
                    double tongPhiMoi = tongPhiCu + phiMua;
                    double totalMua = tongTienMua + phiMua;

                    double slConLaiMoi = slHienTai + sl;
                    double giaTBMoi = (slHienTai <= 0) ? totalMua / sl :
                            (slHienTai * giaTBHienTai + totalMua) / slConLaiMoi;
                    result.tongPhi = tongPhiMoi;
                    result.tongTienMua = tongTienMua;
                    result.phiMua = phiMua;
                    result.totalMua = totalMua;
                    result.slConLai = slConLaiMoi;
                    result.giaTB = giaTBMoi;
                } else {
                    if (sl > slHienTai && slHienTai > 0) {
                        JOptionPane.showMessageDialog(this, "Không đủ số lượng để bán! Còn lại: " + slHienTai, "Infor", JOptionPane.WARNING_MESSAGE);
                        return null;
                    } else if (slHienTai == 0) {
                        JOptionPane.showMessageDialog(this, "Bạn chưa sở hữu mã cổ phiếu này.", "Infor", JOptionPane.WARNING_MESSAGE);
                        return null;
                    } else {
                        double P_tongBan = gia * sl;
                        double phiBan = P_tongBan * 0.0015;
                        double tongPhiMoi = tongPhiCu +phiBan;
                        double thueBan = P_tongBan * 0.001;
                        double totalBan = P_tongBan - phiBan - thueBan;
                        double loiNhuan = totalBan - (sl * giaTBHienTai);

                        result.P_tongBan = P_tongBan;
                        result.phiBan = phiBan;
                        result.thueBan = thueBan;
                        result.tongPhi = tongPhiMoi;
                        result.totalBan = totalBan;
                        result.loiNhuan = loiNhuan;
                        result.slConLai = slHienTai - sl;
                        result.giaTB = giaTBHienTai;
                    }
                }

                return result;
            }

        }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            new StartFrame();
        });
    }
    }