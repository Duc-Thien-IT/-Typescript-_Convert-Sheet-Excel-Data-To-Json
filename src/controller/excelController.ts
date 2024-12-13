import { Request, Response, NextFunction } from 'express';
import XLSX from 'xlsx';

// Hàm xử lý đọc sheet đầu tiên
export const processFirstSheet = async (req: Request, res: Response, next: NextFunction): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ message: 'No file uploaded' });
      return;
    }

    // Đọc file Excel
    const filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Lấy sheet đầu tiên
    const sheet = workbook.Sheets[sheetName];

    // Chuyển đổi dữ liệu từ sheet đầu tiên thành JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    res.json({
      message: 'File processed successfully',
      data: jsonData,
    });
  } catch (error) {
    next(error);
  }
};

// Hàm xử lý đọc tất cả các sheet
export const processAllSheets = async (req: Request, res: Response, next: NextFunction): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ message: 'No file uploaded' });
      return;
    }

    // Đọc file Excel
    const filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames; // Lấy danh sách tất cả các sheet

    // Tạo kết quả JSON cho tất cả các sheet
    const jsonData: Record<string, any[]> = {};
    sheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      jsonData[sheetName] = XLSX.utils.sheet_to_json(sheet);
    });

    res.json({
      message: 'File processed successfully',
      data: jsonData,
    });
  } catch (error) {
    next(error);
  }
};
