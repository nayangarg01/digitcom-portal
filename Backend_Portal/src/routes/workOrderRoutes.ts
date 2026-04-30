import { Router } from 'express';
import { extractWorkOrder, downloadWorkOrderFile, getJobStatus } from '../controllers/workOrderController';
import multer from 'multer';
import { authenticateJWT } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// Route to extract Work Order from PDF
router.post('/extract', authenticateJWT, upload.fields([{ name: 'woFile', maxCount: 1 }]), extractWorkOrder);

// Route to check status of extraction job
router.get('/status/:jobId', authenticateJWT, getJobStatus);

// Route to download generated Excel
router.get('/download/:fileName', authenticateJWT, downloadWorkOrderFile);

export default router;
