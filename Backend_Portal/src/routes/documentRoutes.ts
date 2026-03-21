import express from 'express';
import { uploadDocument, getSiteDocuments } from '../controllers/documentController';
import { upload } from '../middleware/uploadMiddleware';
import { authenticateJWT } from '../middleware/auth';

const router = express.Router();

router.post('/upload', authenticateJWT, upload.single('file'), uploadDocument);
router.get('/site/:site_id', authenticateJWT, getSiteDocuments);

export default router;
