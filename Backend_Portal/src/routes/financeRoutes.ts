import express from 'express';
import { logExpense, logAdvance, getFinancialSummary } from '../controllers/financeController';
import { authenticateJWT, isAdmin } from '../middleware/auth';

const router = express.Router();

router.post('/expenses', authenticateJWT, logExpense);
router.post('/advances', authenticateJWT, isAdmin, logAdvance);
router.get('/summary', authenticateJWT, isAdmin, getFinancialSummary);

export default router;
