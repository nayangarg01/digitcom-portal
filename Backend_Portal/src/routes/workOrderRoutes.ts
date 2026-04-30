import { Router } from 'express';
import { extractWorkOrder, downloadWorkOrderFile } from '../controllers/workOrderController';
import multer from 'multer';
import { authenticateToken } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// Route to extract Work Order from PDF
router.post('/extract', authenticateToken, upload.fields([{ name: 'woFile', maxCount: 1 }]), extractWorkOrder);

// Route to download generated Excel
router.get('/download/:fileName', authenticateToken, downloadWorkOrderFile);

export default router;
