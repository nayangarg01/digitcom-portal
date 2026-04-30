import { Router } from 'express';
import { extractWorkOrder, downloadWorkOrderFile } from '../controllers/workOrderController';
import multer from 'multer';
import { authenticateJWT } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// Route to extract Work Order from PDF
router.post('/extract', authenticateJWT, upload.fields([{ name: 'woFile', maxCount: 1 }]), extractWorkOrder);

// Route to download generated Excel
router.get('/download/:fileName', authenticateJWT, downloadWorkOrderFile);

export default router;
