import { Router } from 'express';
import { generateRoutes, exportRoutePlan } from '../controllers/routePlanningController';
import { authenticateJWT, isAdmin } from '../middleware/auth';
import { upload } from '../middleware/uploadMiddleware';

const router = Router();

// Route Planning Endpoints
router.post('/generate', authenticateJWT, isAdmin, upload.single('file'), generateRoutes);
router.post('/export', authenticateJWT, isAdmin, exportRoutePlan);

export default router;
