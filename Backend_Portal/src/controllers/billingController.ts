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
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };
        const masterFile = files['masterFile'] ? files['masterFile'][0] : null;
        const mindumpFile = files['mindumpFile'] ? files['mindumpFile'][0] : null;

        if (!masterFile) {
            return res.status(400).json({ error: 'No Master DPR file uploaded' });
        }
        if (!mindumpFile) {
            return res.status(400).json({ error: 'No MINDUMP file uploaded' });
        }

        if (!billingTarget || billingTarget.trim() === '') {
            return res.status(400).json({ error: 'Billing Target (e.g. DC0105) is required' });
        }

        // Use __dirname to resolve paths relative to this file's location for maximum reliability
        const projectRoot = path.resolve(__dirname, '../../..');
        const scriptPath = path.join(projectRoot, 'Billing/generate_billing.py');
        const templatePath = path.join(projectRoot, 'Billing/MASTER_JMS_TEMPLATE.xlsx');
        
        // Ensure temporary billing outputs directory exists in Backend_Portal/uploads/billing_outputs
        const outputDir = path.resolve(__dirname, '../../uploads/billing_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const outputFileName = `${billingTarget.toUpperCase()}_Unified_Billing_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);

        // Set headers for streaming response
        res.setHeader('Content-Type', 'text/plain; charset=utf-8');
        res.setHeader('Transfer-Encoding', 'chunked');

        res.write(`--- Launching Unified Precision Billing Engine for ${billingTarget} ---\n`);

        // Spawn Python process with absolute paths for all file inputs
        const pythonProcess = spawn('python3', [
            scriptPath,
            path.resolve(projectRoot, 'Backend_Portal', masterFile.path),
            billingTarget,
            '--template', templatePath,
            '--output', outputPath,
            '--mindump', path.resolve(projectRoot, 'Backend_Portal', mindumpFile.path)
        ], {
            cwd: projectRoot // Run Python from the project root so it can find secondary files if needed
        });

        let pythonOutput = '';
        let pythonError = '';

        pythonProcess.stdout.on('data', (data) => {
            const chunk = data.toString();
            pythonOutput += chunk;
            // Support direct streaming to frontend
            res.write(chunk);
            console.log(`Python STDOUT: ${chunk.trim()}`);
        });

        pythonProcess.stderr.on('data', (data) => {
            const chunk = data.toString();
            pythonError += chunk;
            res.write(`ERROR: ${chunk}`);
            console.error(`Python STDERR: ${chunk.trim()}`);
        });

        pythonProcess.on('close', (code: number) => {
            if (code !== 0) {
                res.write(`\nBUILD FAILED (Code ${code})\n`);
                return res.end();
            }

            if (!fs.existsSync(outputPath)) {
                res.write(`\nFAILED: Output file not found at ${outputPath}\n`);
                return res.end();
            }

            // Provide final JSON-like delimiter for the frontend to parse if needed, 
            // or just the download URL as a special final chunk
            const relativeDownloadPath = `/billing/download/${outputFileName}`;
            res.write(`\nCOMPLETE_PATH:${relativeDownloadPath}\n`);
            res.end();
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
