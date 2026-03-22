import { Request, Response } from 'express';
import * as xlsx from 'xlsx';
const { kmeans } = require('ml-kmeans');
import axios from 'axios';
import { createObjectCsvWriter } from 'csv-writer';
import path from 'path';
import fs from 'fs';

interface Site {
  id: string;
  lat: number;
  lng: number;
  originalIndex: number;
}

interface Leg {
    routeLabel: string;
    stopSequence: number;
    distanceKm: number;
    site: Site;
}

interface Route {
    routeNumber: number;
    label: string;
    legs: Leg[];
}

interface RouteLeg {
  routeNumber: number;
  stopSequence: number;
  fromLocation: string;
  toSiteId: string;
  distanceKm: number;
  cumulativeDistanceKm: number;
}

// Jaipur - Bagru (Lat: 26.8139, Lon: 75.5450)
// Jodhpur - Mogra Khurd (Lat: 26.1245, Lon: 73.0543)
// Lucknow - Safedabad (Lat: 26.8906, Lon: 81.0558)

const WAREHOUSES: Record<string, { lat: number; lng: number }> = {
  'Jaipur - Bagru': { lat: 26.8139, lng: 75.5450 },
  'Jodhpur - Mogra Khurd': { lat: 26.1245, lng: 73.0543 },
  'Lucknow - Safedabad': { lat: 26.8906, lng: 81.0558 }
};

export const generateRoutes = async (req: Request, res: Response) => {
  try {
    const { originName } = req.body;
    const file = req.file;

    // Ensure uploads directory exists
    const uploadsDir = path.join(__dirname, '../../uploads');
    if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
    }

    if (!file) {
      console.error('Route Generation: No file uploaded');
      return res.status(400).json({ error: 'No file uploaded' });
    }

    if (!originName || !WAREHOUSES[originName]) {
      console.error('Route Generation: Invalid origin warehouse:', originName);
      return res.status(400).json({ error: 'Invalid origin warehouse selected' });
    }

    const warehouse = WAREHOUSES[originName];

    // 1. File Parsing (Preserve all columns)
    const workbook = xlsx.readFile(file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const originalRows: any[] = xlsx.utils.sheet_to_json(sheet);

    const sites_count = originalRows.length;
    console.log(`Route Generation: Parsing ${sites_count} rows from ${file.originalname}`);

    // Safety Valve: Check for > 40 sites
    if (sites_count > 40) {
        console.error(`Route Generation: Upload Limit Exceeded (${sites_count} sites)`);
        return res.status(400).json({ 
            error: `Upload Limit Exceeded: Please upload a maximum of 40 sites per batch to prevent API rate limits. (Found ${sites_count} sites)` 
        });
    }

    const apiKey = process.env.Maps_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: 'Maps API key not configured' });
    }

    // Prepare paths for Python script
    const scriptPath = path.join(__dirname, '../../scripts/route_optimizer.py');
    const outputPath = path.join(__dirname, '../../uploads/optimized_route_plan.xlsx');
    
    // Spawn Python process
    const { spawn } = require('child_process');
    const pythonProcess = spawn('python3', [
        scriptPath,
        file.path,
        warehouse.lat.toString(),
        warehouse.lng.toString(),
        apiKey,
        outputPath
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
            console.error('Python Script Error:', pythonError);
            return res.status(500).json({ error: 'Routing engine failed. Ensure OR-Tools and Pandas are installed.' });
        }

        try {
            const result = JSON.parse(pythonOutput);
            if (result.error) {
                return res.status(400).json({ error: result.error });
            }

            res.json({
                success: true,
                num_routes: result.num_routes,
                downloadUrl: '/api/route-planning/download-optimized'
            });
        } catch (e) {
            console.error('JSON Parse Error:', pythonOutput);
            res.status(500).json({ error: 'Failed to parse routing results' });
        }
    });

  } catch (error: any) {
    console.error('Route Generation Error:', error);
    res.status(500).json({ error: error.message || 'Internal Server Error' });
  }
};

export const downloadOptimized = (req: Request, res: Response) => {
    const filePath = path.join(__dirname, '../../uploads/optimized_route_plan.xlsx');
    if (fs.existsSync(filePath)) {
        res.download(filePath, 'Optimized_Dispatch_Plan.xlsx');
    } else {
        res.status(404).json({ error: 'File not found. Please generate the route first.' });
    }
};

export const exportRoutePlan = async (req: Request, res: Response) => {
    // Legacy CSV export (kept for compatibility or can be removed)
    res.status(405).json({ error: 'Please use /download-optimized for the new Excel format' });
};
