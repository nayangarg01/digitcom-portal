import { Request, Response } from 'express';
import prisma from '../utils/prisma';

export const createTeam = async (req: Request, res: Response) => {
    try {
        const { name, members } = req.body;
        const team = await (prisma.team as any).create({
            data: { name, members }
        });
        res.status(201).json(team);
    } catch (error) {
        res.status(400).json({ error: 'Team name must be unique' });
    }
};

export const getTeams = async (_req: Request, res: Response) => {
    const teams = await (prisma.team as any).findMany({
        include: { _count: { select: { sites: true, materials: true } } }
    });
    res.json(teams);
};

export const getTeamById = async (req: Request, res: Response) => {
    const { id } = req.params as { id: string };
    const team = await (prisma.team as any).findUnique({
        where: { id },
        include: { sites: true, materials: true }
    });
    if (!team) return res.status(404).json({ error: 'Team not found' });
    res.json(team);
};
