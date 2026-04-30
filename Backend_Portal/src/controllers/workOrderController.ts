import { Request, Response } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';
import { v4 as uuidv4 } from 'uuid';

// In-memory job store
interface JobStatus {
    id: string;
    status: 'pending' | 'processing' | 'completed' | 'failed';
    error?: string;
    downloadUrl?: string;
    filename?: string;
}

const jobs = new Map<string, JobStatus>();

/**
 * Controller for initiating Work Order Extraction.
 * Returns a jobId immediately.
 */
export const extractWorkOrder = async (req: Request, res: Response) => {
    try {
        const files = req.files as { [fieldname: string]: Express.Multer.File[] };

        if (!files || !files['woFile'] || files['woFile'].length === 0) {
            return res.status(400).json({ error: 'No Work Order PDF file uploaded' });
        }

        const jobId = uuidv4();
        const outputDir = path.join(__dirname, '../../uploads/wo_outputs');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        const backendRoot = path.resolve(__dirname, '../..');
        const scriptPath = path.join(backendRoot, 'scripts/parse_work_order.py');
        const outputFileName = `WorkOrder_Extract_${Date.now()}.xlsx`;
        const outputPath = path.join(outputDir, outputFileName);
        const absolutePdfPath = path.resolve(backendRoot, files['woFile'][0].path);

        // Initialize Job
        jobs.set(jobId, { id: jobId, status: 'processing' });

        // Start background process
        const pythonArgs = [scriptPath, absolutePdfPath, '--output', outputPath];
        const pythonProcess = spawn('python3', pythonArgs);

        let pythonError = '';

        pythonProcess.stderr.on('data', (data: any) => {
            pythonError += data.toString();
        });

        pythonProcess.on('close', (code: number) => {
            // Cleanup uploaded temp PDF
            try {
                if (fs.existsSync(absolutePdfPath)) fs.unlinkSync(absolutePdfPath);
            } catch (err) {
                console.warn(`Job ${jobId}: Temp file cleanup failed`, err);
            }

            const job = jobs.get(jobId);
            if (!job) return;

            if (code !== 0) {
                console.error(`Job ${jobId} failed:`, pythonError);
                job.status = 'failed';
                job.error = pythonError;
            } else if (!fs.existsSync(outputPath)) {
                job.status = 'failed';
                job.error = 'Output file not generated.';
            } else {
                job.status = 'completed';
                job.downloadUrl = `/work-order/download/${outputFileName}`;
                job.filename = outputFileName;
            }
            jobs.set(jobId, job);
        });

        // Return jobId immediately
        res.json({
            success: true,
            jobId: jobId,
            message: 'Extraction started in the background.'
        });

    } catch (error: any) {
        console.error('WorkOrder Extraction Error:', error);
        res.status(500).json({ error: error.message || 'Internal Server Error' });
    }
};

/**
 * Endpoint to poll for job status
 */
export const getJobStatus = (req: Request, res: Response) => {
    const { jobId } = req.params;
    const job = jobs.get(jobId);

    if (!job) {
        return res.status(404).json({ error: 'Job not found' });
    }

    res.json(job);
};

export const downloadWorkOrderFile = (req: Request, res: Response) => {
    const fileName = req.params.fileName;
    const filePath = path.join(__dirname, '../../uploads/wo_outputs', fileName);
    
    if (fs.existsSync(filePath)) {
        res.download(filePath, fileName);
    } else {
        res.status(404).json({ error: 'Work Order file not found or expired.' });
    }
};
