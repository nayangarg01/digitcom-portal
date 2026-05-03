import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for handling Performa Invoice generation by clubbing multiple DC files.
 */
export const generatePerformaInvoice = async (req: Request, res: Response) => {
    try {
        const { ivNumber, activity } = req.body;
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };

        if (!files || !files['dcFiles'] || files['dcFiles'].length === 0) {
            return res.status(400).json({ error: 'No DC billing files uploaded' });
        }

        if (!files['mindumpFile'] || files['mindumpFile'].length === 0) {
            return res.status(400).json({ error: 'MINDUMP file is required' });
        }

        if (!ivNumber) {
            return res.status(400).json({ error: 'Performa Invoice Number is required' });
        }

        const outputDir = path.join(__dirname, '../../uploads/billing_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/generate_performa_invoice.py');
        const outputFileName = `Performa_Invoice_${ivNumber}_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        const absoluteMindumpPath = path.resolve(backendRoot, files['mindumpFile'][0].path);
        
        // Collect all DC file paths
        const dcFilePaths = files['dcFiles'].map(f => path.resolve(backendRoot, f.path));

        const pythonArgs = [
            scriptPath,
            '--files', ...dcFilePaths,
            '--mindump', absoluteMindumpPath,
            '--iv_number', ivNumber,
            '--activity', activity || 'A6',
            '--output', outputPath
        ];

        console.log(`Performa: Starting generation for IV ${ivNumber} with ${dcFilePaths.length} files`);

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
                dcFilePaths.forEach(p => { if (fs.existsSync(p)) fs.unlinkSync(p); });
                if (fs.existsSync(absoluteMindumpPath)) fs.unlinkSync(absoluteMindumpPath);
            } catch (err) {
                console.warn('Performa: Temp file cleanup failed', err);
            }

            if (code !== 0) {
                console.error('Performa Python Error:', pythonError);
                return res.status(500).json({ 
                    error: 'Performa engine failed to generate file.',
                    details: pythonError
                });
            }

            if (!fs.existsSync(outputPath)) {
                return res.status(500).json({ error: 'Failed to generate output file' });
            }

            res.json({
                success: true,
                message: `Performa Invoice ${ivNumber} generated successfully.`,
                downloadUrl: `/billing/download/${outputFileName}`,
                filename: outputFileName
            });
        });

    } catch (error: any) {
        console.error('Performa Generation Error:', error);
        res.status(500).json({ error: error.message || 'Internal Server Error' });
    }
};
