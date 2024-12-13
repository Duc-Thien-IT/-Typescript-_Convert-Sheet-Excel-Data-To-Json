import express, { Request, Response, NextFunction } from 'express';
import multer from 'multer';
import XLSX from 'xlsx';

const router = express.Router();

// Cấu hình Multer để xử lý file upload
const upload = multer({ dest: 'uploads/' });

// Hàm xử lý logic chính
const processExcelFile = async (req: Request, res: Response, next: NextFunction): Promise<void> => {
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

    // Chuyển đổi dữ liệu từ sheet thành JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    res.json({
      message: 'File processed successfully',
      data: jsonData,
    });
  } catch (error) {
    next(error); // Truyền lỗi cho middleware xử lý lỗi
  }
};

/**
 * @swagger
 * /api/excel/upload:
 *   post:
 *     summary: Upload an Excel file and convert it to JSON
 *     tags:
 *       - Excel
 *     requestBody:
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               file:
 *                 type: string
 *                 format: binary
 *                 description: The Excel file to upload
 *     responses:
 *       200:
 *         description: File processed successfully
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *                 data:
 *                   type: array
 *                   items:
 *                     type: object
 *       400:
 *         description: No file uploaded
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *       500:
 *         description: Error processing file
 *         content:
 *           application/json:
 *             schema:
 *               type: object
 *               properties:
 *                 message:
 *                   type: string
 *                 error:
 *                   type: object
 */
router.post('/upload', upload.single('file'), processExcelFile);

export default router;
