import { Router } from 'express';
import multer from 'multer';
import { generateWCC, downloadBillingFile } from '../controllers/billingController';
import { authenticateJWT, isAdmin } from '../middleware/auth';

const router = Router();
const upload = multer({ dest: 'uploads/' });

// WCC Generation Route
router.post('/generate-wcc', authenticateJWT, isAdmin, upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'mindumpFile', maxCount: 1 }
]), generateWCC);

// Download Route
router.get('/download/:fileName', downloadBillingFile);

export default router;
