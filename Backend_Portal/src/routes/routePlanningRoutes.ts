import { Router } from 'express';
import { generateRoutes, exportRoutePlan, downloadOptimized, calculateManualDistances } from '../controllers/routePlanningController';
import { authenticateJWT, isAdmin } from '../middleware/auth';
import { upload } from '../middleware/uploadMiddleware';

const router = Router();

// Route Planning Endpoints
router.post('/generate', authenticateJWT, isAdmin, upload.single('file'), generateRoutes);
router.post('/calculate-manual-distances', authenticateJWT, isAdmin, upload.single('file'), calculateManualDistances);
router.post('/export', authenticateJWT, isAdmin, exportRoutePlan);
router.get('/download-optimized', authenticateJWT, isAdmin, downloadOptimized);

export default router;
