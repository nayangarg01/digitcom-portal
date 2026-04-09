import { Router } from 'express';
import multer from 'multer';
import { generateFullBilling, downloadBillingFile } from '../controllers/billingController';
import { authenticateJWT, isAdmin } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// Unified Billing Generation Route
router.post('/generate-file', authenticateJWT, isAdmin, upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'mindumpFile', maxCount: 1 }
]), generateFullBilling);

// Download Route
router.get('/download/:fileName', downloadBillingFile);

export default router;
