import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for Object-Oriented Billing and Database operations.
 */
export const syncOOPDb = async (req: Request, res: Response) => {
    try {
        console.log("OOP Controller: Received sync request");
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };

        if (!files || !files['masterFile'] || files['masterFile'].length === 0) {
            return res.status(400).json({ error: 'No Master DPR file uploaded' });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/sync_oop_db.py');

        // Rename Master File
        const absoluteMasterPath = path.resolve(backendRoot, files['masterFile'][0].path);
        const masterWithExt = absoluteMasterPath + '.xlsx';
        fs.renameSync(absoluteMasterPath, masterWithExt);

        const pythonArgs = [scriptPath, masterWithExt];

        // Rename Mindump File if provided
        let mindumpWithExt = '';
        if (files['mindumpFile'] && files['mindumpFile'].length > 0) {
            const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
            mindumpWithExt = absoluteMindumpPath + '.xlsx';
            fs.renameSync(absoluteMindumpPath, mindumpWithExt);
            pythonArgs.push('--mindump', mindumpWithExt);
        }

        console.log(`OOP Controller: Spawning Python script to sync database...`);
        // We use python3 (which is standard and stable on their hosting platform)
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
            // Clean up uploaded files
            try {
                if (fs.existsSync(masterWithExt)) fs.unlinkSync(masterWithExt);
                if (mindumpWithExt && fs.existsSync(mindumpWithExt)) fs.unlinkSync(mindumpWithExt);
            } catch (err) {
                console.warn('OOP Controller: Sync temp file cleanup failed', err);
            }

            if (code !== 0) {
                console.error('OOP Controller Sync Python Error:', pythonError);
                return res.status(500).json({
                    success: false,
                    error: 'OOP sync process failed.',
                    details: pythonError,
                    logs: (pythonOutput + "\n" + pythonError).split('\n')
                });
            }

            console.log(`OOP Controller: Database sync completed successfully.`);
            res.json({
                success: true,
                message: 'OOP database synchronized successfully.',
                logs: pythonOutput.split('\n')
            });
        });

    } catch (error: any) {
        console.error('OOP Controller Sync Route Error:', error);
        res.status(500).json({ success: false, error: error.message || 'Internal Server Error' });
    }
};

export const generateOOPBilling = async (req: Request, res: Response) => {
    try {
        const { billingTarget, activity } = req.body;
        console.log(`OOP Controller: Received billing generation request for ${billingTarget}`);

        if (!billingTarget || billingTarget.trim() === '') {
            return res.status(400).json({ error: 'Billing Target (e.g. DC0122) is required' });
        }

        const outputDir = path.join(__dirname, '../../uploads/billing_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/generate_oop_billing.py');
        const outputFileName = `${billingTarget.toUpperCase()}_OOP_Billing_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);

        const pythonArgs = [
            scriptPath,
            billingTarget.trim(),
            '--output', outputPath
        ];

        if (activity && activity.trim() !== '' && activity !== 'AUTO') {
            pythonArgs.push('--activity', activity);
        }

        console.log(`OOP Controller: Spawning Python script to generate billing...`);
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
            if (code !== 0) {
                console.error('OOP Controller Generation Python Error:', pythonError);
                return res.status(500).json({
                    success: false,
                    error: 'OOP generation process failed.',
                    details: pythonError,
                    logs: (pythonOutput + "\n" + pythonError).split('\n')
                });
            }

            if (!fs.existsSync(outputPath)) {
                console.error('OOP Controller Output Error: File not generated at', outputPath);
                return res.status(500).json({ success: false, error: 'Failed to generate output file' });
            }

            console.log(`OOP Controller: Successfully generated OOP Billing File for ${billingTarget}`);

            res.json({
                success: true,
                message: `OOP Billing file for ${billingTarget} generated successfully.`,
                downloadUrl: `/billing/download/${outputFileName}`,
                filename: outputFileName,
                logs: pythonOutput.split('\n')
            });
        });

    } catch (error: any) {
        console.error('OOP Controller Generation Route Error:', error);
        res.status(500).json({ success: false, error: error.message || 'Internal Server Error' });
    }
};

export const generateOOPPerforma = async (req: Request, res: Response) => {
    try {
        const { ivNumber, dcNumbers, activity } = req.body;
        console.log(`OOP Controller: Received performa invoice generation request for IV ${ivNumber}`);

        if (!ivNumber || ivNumber.trim() === '') {
            return res.status(400).json({ error: 'Invoice Number is required' });
        }
        if (!dcNumbers || dcNumbers.length === 0) {
            return res.status(400).json({ error: 'DC Numbers are required' });
        }

        const targetDcs = Array.isArray(dcNumbers)
            ? dcNumbers
            : String(dcNumbers).split(',').map(s => s.trim()).filter(Boolean);

        if (targetDcs.length === 0) {
            return res.status(400).json({ error: 'No valid DC Numbers provided' });
        }

        const outputDir = path.join(__dirname, '../../uploads/billing_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/generate_oop_performa.py');
        const outputFileName = `Performa_Invoice_${ivNumber}_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);

        const pythonArgs = [
            scriptPath,
            ivNumber.trim(),
            ...targetDcs,
            '--output', outputPath,
            '--activity', activity || 'A6'
        ];

        console.log(`OOP Controller: Spawning Python script to generate performa invoice...`);
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
            if (code !== 0) {
                console.error('OOP Controller Performa Python Error:', pythonError);
                return res.status(500).json({
                    success: false,
                    error: 'OOP Performa invoice generation failed.',
                    details: pythonError,
                    logs: (pythonOutput + "\n" + pythonError).split('\n')
                });
            }

            if (!fs.existsSync(outputPath)) {
                return res.status(500).json({ success: false, error: 'Failed to generate performa output file' });
            }

            res.json({
                success: true,
                message: `OOP Performa invoice for IV ${ivNumber} generated successfully.`,
                downloadUrl: `/billing/download/${outputFileName}`,
                filename: outputFileName,
                logs: pythonOutput.split('\n')
            });
        });

    } catch (error: any) {
        console.error('OOP Controller Performa Route Error:', error);
        res.status(500).json({ success: false, error: error.message || 'Internal Server Error' });
    }
};

export const generateOOPRoutes = async (req: Request, res: Response) => {
    try {
        const { dcNumbers, minDates } = req.body;
        console.log(`OOP Controller: Received routing generation request`);

        const apiKey = process.env.Maps_API_KEY;
        if (!apiKey) {
            return res.status(500).json({ error: 'Maps API key not configured' });
        }

        const outputDir = path.join(__dirname, '../../uploads');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/generate_oop_routes.py');
        const outputFileName = `Routing_Result_OOP_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);

        const pythonArgs = [
            scriptPath,
            '--api_key', apiKey,
            '--output', outputPath
        ];

        if (dcNumbers && String(dcNumbers).trim() !== '') {
            pythonArgs.push('--dc_numbers', String(dcNumbers).trim());
        } else if (minDates && String(minDates).trim() !== '') {
            pythonArgs.push('--dates', String(minDates).trim());
        }

        console.log(`OOP Controller: Spawning Python script to generate optimized routes...`);
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
            if (code !== 0) {
                console.error('OOP Controller Routing Python Error:', pythonError);
                return res.status(500).json({
                    success: false,
                    error: 'OOP Route planning failed.',
                    details: pythonError,
                    logs: (pythonOutput + "\n" + pythonError).split('\n')
                });
            }

            try {
                const result = JSON.parse(pythonOutput.trim());
                if (result.error) {
                    return res.status(400).json({ success: false, error: result.error, logs: [result.error] });
                }

                res.json({
                    success: true,
                    num_routes: result.num_routes,
                    routes: result.routes,
                    downloadUrl: '/api/route-planning/download-optimized', // Use existing download route
                    filename: outputFileName
                });
            } catch (e) {
                console.error('OOP Controller Routing JSON Parse Error:', pythonOutput);
                res.status(500).json({
                    success: false,
                    error: 'Failed to parse routing results',
                    details: pythonOutput,
                    logs: pythonOutput.split('\n')
                });
            }
        });

    } catch (error: any) {
        console.error('OOP Controller Routing Route Error:', error);
        res.status(500).json({ success: false, error: error.message || 'Internal Server Error' });
    }
};
