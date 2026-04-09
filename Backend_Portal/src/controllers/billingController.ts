import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for handling Billing related operations like WCC generation.
 */
export const generateWCC = async (req: Request, res: Response) => {
    try {
        const { billingTarget } = req.body;
        const file = req.file;

        if (!file) {
            return res.status(400).json({ error: 'No Master DPR file uploaded' });
        }

        if (!billingTarget || billingTarget.trim() === '') {
            return res.status(400).json({ error: 'Billing Target (e.g. DC0105) is required' });
        }

        // Ensure temporary billing outputs directory exists
        const outputDir = path.join(__dirname, '../../uploads/billing_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const scriptPath = path.join(__dirname, '../../../Billing/generate_billing.py');
        const templatePath = path.join(__dirname, '../../../Billing/MASTER_JMS_TEMPLATE.xlsx');
        const outputFileName = `${billingTarget.toUpperCase()}_Unified_Billing_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);

        console.log(`Billing: Starting Unified Portfolio Generation for ${billingTarget}`);

        // Spawn Python process for Unified Billing generation
        const pythonProcess = spawn('python3', [
            scriptPath,
            file.path,
            billingTarget,
            '--template', templatePath,
            '--output', outputPath
        ]);

        let pythonOutput = '';
        let pythonError = '';

        pythonProcess.stdout.on('data', (data: any) => {
            pythonOutput += data.toString();
        });

        pythonProcess.stderr.on('data', (data: any) => {
            pythonError += data.toString();
        });

        pythonProcess.on('close', (code: number) => {
            if (code !== 0) {
                console.error('Billing Python Error:', pythonError);
                return res.status(500).json({ 
                    error: 'Billing engine failed to generate WCC.',
                    details: pythonError
                });
            }

            if (!fs.existsSync(outputPath)) {
                console.error('Billing Output Error: File not generated at', outputPath);
                return res.status(500).json({ error: 'Failed to generate output file' });
            }

            console.log(`Billing: Successfully generated WCC for ${billingTarget}`);
            
            // Return the download link
            res.json({
                success: true,
                message: `WCC for ${billingTarget} generated successfully.`,
                downloadUrl: `/billing/download/${outputFileName}`
            });
        });

    } catch (error: any) {
        console.error('WCC Generation Route Error:', error);
        res.status(500).json({ error: error.message || 'Internal Server Error' });
    }
};

export const downloadBillingFile = (req: Request, res: Response) => {
    const fileName = req.params.fileName;
    if (typeof fileName !== 'string') {
        return res.status(400).json({ error: 'Invalid file name' });
    }
    const filePath = path.join(__dirname, '../../uploads/billing_outputs', fileName);
    
    if (fs.existsSync(filePath)) {
        res.download(filePath, fileName);
    } else {
        res.status(404).json({ error: 'Billing file not found or expired.' });
    }
};
