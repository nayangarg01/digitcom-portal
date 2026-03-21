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

    console.log(`Route Generation: Parsing ${originalRows.length} rows from ${file.originalname}`);

    const sites: Site[] = originalRows.map((row: any, index: number) => {
      // Find keys ignoring case and spaces
      const findValue = (possibleKeys: string[]) => {
          const key = Object.keys(row).find(k => 
            possibleKeys.some(pk => k.trim().toLowerCase() === pk.toLowerCase())
          );
          return key ? row[key] : null;
      };

      const latVal = findValue(['latitude', 'lat', 'lat ']);
      const lngVal = findValue(['longitude', 'lng', 'lon', 'long']);
      const idVal = findValue(['site id', 'site_id', 'siteid', 'enbsiteid']);

      return {
        id: String(idVal || index),
        lat: parseFloat(String(latVal || '')),
        lng: parseFloat(String(lngVal || '')),
        originalIndex: index
      };
    }).filter(s => !isNaN(s.lat) && !isNaN(s.lng));

    console.log(`Route Generation: Found ${sites.length} valid sites with coordinates`);

    if (sites.length === 0) {
      console.error('Route Generation: No valid coordinates found in file');
      return res.status(400).json({ error: 'No valid coordinates found in the file. Check if Latitude/Longitude columns exist.' });
    }

    // 2. Greedy Grouping (Max Size 3)
    // Heuristic: Start from the site furthest from warehouse, and pick 2 nearest neighbors
    const clusters: Site[][] = [];
    let remainingSites = [...sites];

    while (remainingSites.length > 0) {
        // Find site furthest from warehouse
        let furthestIdx = 0;
        let maxDist = -1;
        remainingSites.forEach((s, idx) => {
            const d = Math.sqrt(Math.pow(s.lat - warehouse.lat, 2) + Math.pow(s.lng - warehouse.lng, 2));
            if (d > maxDist) {
                maxDist = d;
                furthestIdx = idx;
            }
        });

        const seedSite = remainingSites[furthestIdx];
        remainingSites.splice(furthestIdx, 1);
        const currentCluster = [seedSite];

        // Find up to 2 nearest neighbors to the seed site
        for (let i = 0; i < 2; i++) {
            if (remainingSites.length === 0) break;
            let nearestIdx = 0;
            let minDist = Infinity;
            remainingSites.forEach((s, idx) => {
                const d = Math.sqrt(Math.pow(s.lat - seedSite.lat, 2) + Math.pow(s.lng - seedSite.lng, 2));
                if (d < minDist) {
                    minDist = d;
                    nearestIdx = idx;
                }
            });
            currentCluster.push(remainingSites[nearestIdx]);
            remainingSites.splice(nearestIdx, 1);
        }
        clusters.push(currentCluster);
    }

    const routes: Route[] = [];
    const apiKey = process.env.Maps_API_KEY;

    if (!apiKey) {
      return res.status(500).json({ error: 'Maps API key not configured' });
    }

    // 3. Routing for each cluster
    for (let i = 0; i < clusters.length; i++) {
        const clusterSites = clusters[i];
        let currentLocation = { lat: warehouse.lat, lng: warehouse.lng };
        let clusterRemaining = [...clusterSites];

        const legs = [];
        const orderedSitesInRoute = [];
        let stopSequence = 1;

        while (clusterRemaining.length > 0) {
            const originStr = `${currentLocation.lat},${currentLocation.lng}`;
            const destinations = clusterRemaining.map(s => `${s.lat},${s.lng}`).join('|');
            
            try {
                const response = await axios.get(`https://maps.googleapis.com/maps/api/distancematrix/json?origins=${originStr}&destinations=${destinations}&key=${apiKey}`);
                
                if (response.data.status !== 'OK') {
                    throw new Error(`Google Maps API error: ${response.data.status}`);
                }

                const results = response.data.rows[0].elements;
                let bestIdx = 0;
                let minLegDist = Infinity;

                results.forEach((el: any, idx: number) => {
                    if (el.status === 'OK' && el.distance.value < minLegDist) {
                        minLegDist = el.distance.value;
                        bestIdx = idx;
                    }
                });

                const nextSite = clusterRemaining[bestIdx];
                let distanceKm = minLegDist / 1000;

                // SPECIAL RULE: Subtract 50km for each A1, B1, C1 entry
                if (stopSequence === 1) {
                    distanceKm = Math.max(0, distanceKm - 50);
                }

                legs.push({
                    routeLabel: String.fromCharCode(65 + i), // A, B, C...
                    stopSequence: stopSequence++,
                    distanceKm: parseFloat(distanceKm.toFixed(2)),
                    site: nextSite
                });

                orderedSitesInRoute.push(nextSite);
                currentLocation = { lat: nextSite.lat, lng: nextSite.lng };
                clusterRemaining.splice(bestIdx, 1);

            } catch (err: any) {
                console.error('Routing error:', err.message);
                return res.status(500).json({ error: 'Failed to optimize route legs' });
            }
        }

        routes.push({
            routeNumber: i + 1,
            label: String.fromCharCode(65 + i),
            legs
        });
    }

    // 4. Update Original Data for Export
    let exportData = originalRows.map((row, index) => {
        // Find if this row is part of any route
        let legInfo = null;
        for (const r of routes) {
            const leg = r.legs.find((l: any) => l.site.originalIndex === index);
            if (leg) {
                legInfo = leg;
                break;
            }
        }

        if (legInfo) {
            return {
                ...row,
                CLUBBING: `${legInfo.routeLabel}${legInfo.stopSequence}`,
                AKTBC: legInfo.distanceKm
            };
        }
        return row;
    });

    // Sort by CLUBBING (A1, A2, A3... B1, B2... Z3)
    exportData.sort((a, b) => {
        if (!a.CLUBBING) return 1;
        if (!b.CLUBBING) return -1;
        // Natural sort (supports A1, A2, A10)
        return a.CLUBBING.localeCompare(b.CLUBBING, undefined, { numeric: true, sensitivity: 'base' });
    });

    // 5. Generate Excel File
    const newWB = xlsx.utils.book_new();
    const newWS = xlsx.utils.json_to_sheet(exportData);
    xlsx.utils.book_append_sheet(newWB, newWS, 'Route Plan');
    
    const exportPath = path.join(__dirname, '../../uploads/optimized_route_plan.xlsx');
    xlsx.writeFile(newWB, exportPath);

    res.json({
        success: true,
        routes,
        downloadUrl: '/api/route-planning/download-optimized'
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
