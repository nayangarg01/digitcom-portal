import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';

/**
 * Controller for handling Work Order Extraction from PDF.
 */
export const extractWorkOrder = async (req: Request, res: Response) => {
    try {
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };

        if (!files || !files['woFile'] || files['woFile'].length === 0) {
            return res.status(400).json({ error: 'No Work Order PDF file uploaded' });
        }

        // Ensure temporary work order outputs directory exists
        const outputDir = path.join(__dirname, '../../uploads/wo_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // Absolute path resolution
        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/parse_work_order.py');
        const outputFileName = `WorkOrder_Extract_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        
        const absolutePdfPath = path.resolve(backendRoot, files['woFile'][0].path);
        
        // Construct arguments for Python
        const pythonArgs = [
            scriptPath,
            absolutePdfPath,
            '--output', outputPath
        ];

        console.log(`WorkOrder: Starting Extraction for ${files['woFile'][0].originalname}`);

        // Spawn Python process
        // Note: Using 'python3' as per the user's environment
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
            // Cleanup uploaded temp PDF
            try {
                if (fs.existsSync(absolutePdfPath)) fs.unlinkSync(absolutePdfPath);
            } catch (err) {
                console.warn('WorkOrder: Temp file cleanup failed', err);
            }

            if (code !== 0) {
                console.error('WorkOrder Python Error:', pythonError);
                return res.status(500).json({ 
                    error: 'Work Order extraction engine failed.',
                    details: pythonError
                });
            }

            if (!fs.existsSync(outputPath)) {
                console.error('WorkOrder Output Error: File not generated at', outputPath);
                return res.status(500).json({ error: 'Failed to generate output file' });
            }

            console.log(`WorkOrder: Successfully generated Excel for ${files['woFile'][0].originalname}`);
            
            res.json({
                success: true,
                message: `Work Order extracted successfully.`,
                downloadUrl: `/work-order/download/${outputFileName}`,
                filename: outputFileName
            });
        });

    } catch (error: any) {
        console.error('WorkOrder Extraction Route Error:', error);
        res.status(500).json({ error: error.message || 'Internal Server Error' });
    }
};

export const downloadWorkOrderFile = (req: Request, res: Response) => {
    const fileName = req.params.fileName;
    if (typeof fileName !== 'string') {
        return res.status(400).json({ error: 'Invalid file name' });
    }
    const filePath = path.join(__dirname, '../../uploads/wo_outputs', fileName);
    
    if (fs.existsSync(filePath)) {
        res.download(filePath, fileName);
    } else {
        res.status(404).json({ error: 'Work Order file not found or expired.' });
    }
};
