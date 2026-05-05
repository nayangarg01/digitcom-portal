import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for handling Billing related operations like WCC generation.
 */
export const generateFullBilling = async (req: Request, res: Response) => {
    try {
        const { billingTarget, activity } = req.body;
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
        const scriptPath = path.join(backendRoot, 'scripts/generate_clean_billing.py');
        const templatePath = path.join(backendRoot, 'templates/billing_template.xlsx');
        const outputFileName = `${billingTarget.toUpperCase()}_Clean_Billing_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        const absoluteMasterPath = path.resolve(backendRoot, files['masterFile'][0].path);
        const masterWithExt = absoluteMasterPath + '.xlsx';
        fs.renameSync(absoluteMasterPath, masterWithExt);
        
        // Construct arguments for Python
        const pythonArgs = [
            scriptPath,
            masterWithExt,
            billingTarget,
            '--template', templatePath,
            '--output', outputPath,
            '--activity', activity || 'A6'
        ];

        // Add mindump if provided
        if (files['mindumpFile'] && files['mindumpFile'].length > 0) {
            const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
            const mindumpWithExt = absoluteMindumpPath + '.xlsx';
            fs.renameSync(absoluteMindumpPath, mindumpWithExt);
            pythonArgs.push('--mindump', mindumpWithExt);
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
                if (fs.existsSync(masterWithExt)) fs.unlinkSync(masterWithExt);
                if (files['mindumpFile'] && files['mindumpFile'].length > 0) {
                    const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
                    const mindumpWithExt = absoluteMindumpPath + '.xlsx';
                    if (fs.existsSync(mindumpWithExt)) fs.unlinkSync(mindumpWithExt);
                }
                if (files['dcFiles'] && files['dcFiles'].length > 0) {
                    files['dcFiles'].forEach(f => {
                        const p = path.resolve(backendRoot, f.path) + '.xlsx';
                        if (fs.existsSync(p)) fs.unlinkSync(p);
                    });
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
                downloadUrl: `/billing/download/${outputFileName}`,
                filename: outputFileName,
                logs: pythonOutput.split('\n')
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
