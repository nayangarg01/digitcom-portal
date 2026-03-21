import { Router } from 'express';
import { getSites, createSite, getMaterials, createMaterial, updateMaterial, kitMaterials, reconcileMaterials, updateSiteStage, bulkImportSites } from '../controllers/logisticsController';
import { authenticateJWT, authorizeRole, isAdmin } from '../middleware/auth';
import { upload } from '../middleware/uploadMiddleware';

const router = Router();

// Sites - Admins can create, everyone authenticated can view
router.get('/sites', authenticateJWT, getSites);
router.post('/sites', authenticateJWT, authorizeRole(['ADMIN']), createSite);
router.post('/bulk-import', authenticateJWT, isAdmin, upload.single('file'), bulkImportSites);

// Materials - Everyone authenticated can view and manage
router.get('/materials', authenticateJWT, getMaterials);
router.post('/materials', authenticateJWT, createMaterial);
router.patch('/materials/:id', authenticateJWT, updateMaterial);

// Phase 4 - Kitting & Reconciliation
router.post('/kit', authenticateJWT, isAdmin, kitMaterials);
router.post('/reconcile', authenticateJWT, isAdmin, reconcileMaterials);
router.patch('/sites/:id/stage', authenticateJWT, isAdmin, updateSiteStage);

export default router;
