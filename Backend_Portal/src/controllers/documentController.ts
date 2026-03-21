import { Request, Response } from 'express';
import prisma from '../utils/prisma';

export const uploadDocument = async (req: Request, res: Response): Promise<void> => {
    try {
        const { site_id, type } = req.body;
        const file = req.file;

        if (!file) {
            res.status(400).json({ error: 'No file uploaded' });
            return;
        }

        const document = await (prisma as any).document.create({
            data: {
                site_id,
                type, // PERMIT, ATP_11A, ATP_11C, PHOTO
                url: `/uploads/${file.filename}`
            }
        });

        // Optionally update Site model direct fields if needed
        if (type === 'PERMIT') {
            await prisma.site.update({
                where: { id: site_id as any },
                data: { work_permit_url: `/uploads/${file.filename}` } as any
            });
        } else if (type === 'ATP_11A') {
            await prisma.site.update({
                where: { id: site_id as any },
                data: { atp_11a_url: `/uploads/${file.filename}` } as any
            });
        } else if (type === 'ATP_11C') {
            await prisma.site.update({
                where: { id: site_id as any },
                data: { atp_11c_url: `/uploads/${file.filename}` } as any
            });
        }

        res.status(201).json(document);
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};

export const getSiteDocuments = async (req: Request, res: Response): Promise<void> => {
    try {
        const { site_id } = req.params;
        const documents = await (prisma as any).document.findMany({
            where: { site_id }
        });
        res.json(documents);
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};
