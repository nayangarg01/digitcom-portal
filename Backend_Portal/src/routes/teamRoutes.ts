import { Router } from 'express';
import { createTeam, getTeams, getTeamById } from '../controllers/teamController';
import { authenticateJWT, isAdmin } from '../middleware/auth';

const router = Router();

router.post('/', authenticateJWT, isAdmin, createTeam);
router.get('/', authenticateJWT, getTeams);
router.get('/:id', authenticateJWT, getTeamById);

export default router;
