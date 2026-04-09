import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for handling Billing related operations like WCC generation.
 */
export const generateFullBilling = async (req: Request, res: Response) => {
    try {
        const { billingTarget } = req.body;
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };

        if (!files || !files['masterFile'] || files['masterFile'].length === 0) {
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

        // Robust absolute path resolution
        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/generate_billing_FULL.py');
        const templatePath = path.join(backendRoot, 'scripts/MASTER_JMS_TEMPLATE.xlsx');
        const outputFileName = `${billingTarget.toUpperCase()}_Unified_Billing_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        const absoluteMasterPath = path.resolve(backendRoot, files['masterFile'][0].path);
        
        // Construct arguments for Python
        const pythonArgs = [
            scriptPath,
            absoluteMasterPath,
            billingTarget,
            '--template', templatePath,
            '--output', outputPath
        ];

        // Add mindump if provided
        if (files['mindumpFile'] && files['mindumpFile'].length > 0) {
            const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
            pythonArgs.push('--mindump', absoluteMindumpPath);
        }

        console.log(`Billing: Starting Unified Generation for ${billingTarget}`);

        // Spawn Python process
        const pythonProcess = spawn('python3', pythonArgs);

        let pythonOutput = '';
        let pythonError = '';

        pythonProcess.stdout.on('data', (data: any) => {
            pythonOutput += data.toString();
        });

        pythonProcess.stderr.on('data', (data: any) => {
            pythonError += data.toString();
        });

        pythonProcess.on('close', (code: number) => {
            // Cleanup uploaded temp files
            try {
                if (fs.existsSync(absoluteMasterPath)) fs.unlinkSync(absoluteMasterPath);
                if (files['mindumpFile'] && files['mindumpFile'].length > 0) {
                    const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
                    if (fs.existsSync(absoluteMindumpPath)) fs.unlinkSync(absoluteMindumpPath);
                }
            } catch (err) {
                console.warn('Billing: Temp file cleanup failed', err);
            }

            if (code !== 0) {
                console.error('Billing Python Error:', pythonError);
                return res.status(500).json({ 
                    error: 'Billing engine failed to generate file.',
                    details: pythonError
                });
            }

            if (!fs.existsSync(outputPath)) {
                console.error('Billing Output Error: File not generated at', outputPath);
                return res.status(500).json({ error: 'Failed to generate output file' });
            }

            console.log(`Billing: Successfully generated Billing File for ${billingTarget}`);
            
            // Return the download link (no leading slash to avoid double-slash with API_URL)
            res.json({
                success: true,
                message: `Billing file for ${billingTarget} generated successfully.`,
                downloadUrl: `billing/download/${outputFileName}`
            });
        });

    } catch (error: any) {
        console.error('Billing Generation Route Error:', error);
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
