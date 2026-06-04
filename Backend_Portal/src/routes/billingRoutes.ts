import { Router } from 'express';
import multer from 'multer';
import { generateFullBilling, downloadBillingFile } from '../controllers/billingController';
import { generatePerformaInvoice } from '../controllers/performaInvoiceController';
import { syncOOPDb, generateOOPBilling } from '../controllers/oopBillingController';
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

// OOP Database Sync Route
router.post('/oop-sync-db', authenticateJWT, isAdmin, upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'mindumpFile', maxCount: 1 }
]), syncOOPDb);

// OOP Billing Generation Route
router.post('/oop-generate', authenticateJWT, isAdmin, upload.none(), generateOOPBilling);

// Download Route
router.get('/download/:fileName', downloadBillingFile);

export default router;
