import { Router } from 'express';
import multer from 'multer';
import { generateFullBilling, downloadBillingFile } from '../controllers/billingController';
import { generatePerformaInvoice } from '../controllers/performaInvoiceController';
import { authenticateJWT, isAdmin } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// Unified Billing Generation Route
router.post('/generate-file', authenticateJWT, isAdmin, upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'mindumpFile', maxCount: 1 }
]), generateFullBilling);

// Performa Invoice Generation (Clubbing)
router.post('/generate-performa', authenticateJWT, isAdmin, upload.fields([
    { name: 'dcFiles', maxCount: 10 },
    { name: 'mindumpFile', maxCount: 1 }
]), generatePerformaInvoice);

// Download Route
router.get('/download/:fileName', downloadBillingFile);

export default router;
