import { Request, Response } from 'express';
import prisma from '../utils/prisma';

export const logExpense = async (req: Request, res: Response): Promise<void> => {
    try {
        const { site_id, category, amount, description, date } = req.body;
        const expense = await (prisma as any).expense.create({
            data: {
                site_id,
                category,
                amount: parseFloat(amount),
                description,
                date: date ? new Date(date) : new Date()
            }
        });
        res.status(201).json(expense);
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};

export const logAdvance = async (req: Request, res: Response): Promise<void> => {
    try {
        const { team_id, amount, description, date } = req.body;
        const advance = await (prisma as any).advance.create({
            data: {
                team_id,
                amount: parseFloat(amount),
                description,
                date: date ? new Date(date) : new Date()
            }
        });
        res.status(201).json(advance);
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};

export const getFinancialSummary = async (req: Request, res: Response): Promise<void> => {
    try {
        const expenses = await (prisma as any).expense.findMany({
            include: { site: true }
        });
        const advances = await (prisma as any).advance.findMany({
            include: { team: true }
        });

        res.json({ expenses, advances });
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};
