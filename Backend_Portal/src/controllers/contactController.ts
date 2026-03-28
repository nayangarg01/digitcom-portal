import { Request, Response } from 'express';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export const submitContactForm = async (req: Request, res: Response) => {
  try {
    const { name, email, company, message } = req.body;

    if (!name || !email || !message) {
      return res.status(400).json({ error: 'Name, email, and message are required.' });
    }

    const lead = await (prisma as any).lead.create({
      data: {
        name,
        email,
        company,
        message,
      },
    });

    return res.status(201).json({
      message: 'Thank you! Your message has been received.',
      leadId: lead.id,
    });
  } catch (error) {
    console.error('Contact Form Submission Error:', error);
    return res.status(500).json({ error: 'Internal server error. Please try again later.' });
  }
};
