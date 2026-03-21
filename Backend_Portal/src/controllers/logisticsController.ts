import { Response } from 'express';
import { AuthRequest } from '../middleware/auth';
import prisma from '../utils/prisma';
import * as xlsx from 'xlsx';

// SITE CONTROLLERS
export const bulkImportSites = async (req: AuthRequest, res: Response) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet) as any[];

    const results = {
      total: rows.length,
      success: 0,
      errors: [] as string[]
    };

    // Master Column Template (from RJ-13 March Tracker)
    const MASTER_COLUMNS = [
      'PMP ID', 'Sector', 'NO OF\n SECTOR', 'BAND', 'JC NAME', 'CMP', 'Team', 'Plan Date', 
      'Current Status', 'LATITUDE', 'LONGITUDE', 'SITE TYPE', 'SITE IP', 'ACTIVITY', 
      'ACTIVITY CHANGES IF ANY', 'EXTRA MIN IF ANY', 'ALLOTMENT \nLOT', 'ALLOTMENT DATE', 
      'MATERIAL \nDISPATCHED  \nDATE', 'INSTALLATION \nDATE', 'INTEGRATION\nDATE', 
      'ATP 11A DATE', 'ATP11C', 'ATP11C DONE BY', 'HOTO DATE', '11A SUBMISSION DATE', 
      '11BC SUBMISSION DATE', 'HOTO SUBMISSION DATE', 'Material Need to MRN', 
      'EXTRA MIN (YES/NO)', 'MRN \nREQUIRED (YES/NO)', 'MRN FINALIZATION STATUS', 
      'REASON \nIF MRN PENDING', 'RECO\nSTATUS', 'READY FOR BILLING OR NOT', 
      'REMARKS IF ANY', 'Digitcom Remark', 'POWER CABLE LENGTH', 'CPRI LENGTH', 
      'HARD CONDUIT LENGTH ', 'SITE\nNAME', 'SITE \nADDRESS', 'JIO Technician Name', 
      'Jio Technician Contact', 'Rigger Name', 'Rigger Contact', 'Superviser Name', 
      'Su[perviser Contact', 'MIN Location', 'MRN Location'
    ];

    for (const row of rows) {
      try {
        const site_id = (row['SITE ID'] || row['site_id'])?.toString().trim();
        if (!site_id) {
          results.errors.push(`Row ${rows.indexOf(row) + 2}: Missing Site ID`);
          continue;
        }

        const latitude = parseFloat(row['LATITUDE'] || row['latitude']);
        const longitude = parseFloat(row['LONGITUDE'] || row['longitude']);
        const work_type = row['ACTIVITY'] || row['BAND'] || 'UNKNOWN';

        // 1. Padding: Merge uploaded row with Master Column template
        const padded_row = { ...row };
        MASTER_COLUMNS.forEach(col => {
          if (!(col in padded_row)) {
            padded_row[col] = ''; // Initialize missing columns as empty/null/0
          }
        });

        // 2. Upsert using SITE ID as the primary key
        await prisma.site.upsert({
          where: { site_id },
          update: {
            work_type,
            status: row['Current Status']?.toUpperCase() || 'ALLOTTED',
            latitude: isNaN(latitude) ? null : latitude,
            longitude: isNaN(longitude) ? null : longitude,
            dpr_data: JSON.stringify(padded_row)
          },
          create: {
            site_id,
            work_type,
            status: row['Current Status']?.toUpperCase() || 'ALLOTTED',
            latitude: isNaN(latitude) ? null : latitude,
            longitude: isNaN(longitude) ? null : longitude,
            dpr_data: JSON.stringify(padded_row)
          }
        });
        results.success++;
      } catch (err: any) {
        results.errors.push(`Row ${rows.indexOf(row) + 2}: ${err.message}`);
      }
    }

    res.json(results);
  } catch (error: any) {
    res.status(500).json({ error: 'Bulk import failed.', details: error.message });
  }
};
export const getSites = async (req: AuthRequest, res: Response) => {
  try {
    const sites = await prisma.site.findMany({
      include: { materials: true }
    });
    res.json(sites);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to fetch sites.', details: error.message });
  }
};

export const createSite = async (req: AuthRequest, res: Response) => {
  try {
    const { site_id, work_type, status } = req.body;
    const site = await prisma.site.create({
      data: { site_id, work_type, status }
    });
    res.status(201).json(site);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to create site.', details: error.message });
  }
};

// MATERIAL CONTROLLERS
export const getMaterials = async (req: AuthRequest, res: Response) => {
  try {
    const materials = await prisma.material.findMany({
      include: { site: true, user: true }
    });
    res.json(materials);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to fetch materials.', details: error.message });
  }
};

export const createMaterial = async (req: AuthRequest, res: Response) => {
  try {
    const { site_id, equipment_name, serial_number, jio_min, status } = req.body;
    
    if (!req.user) return res.status(401).json({ error: 'Unauthorized' });

    const material = await prisma.material.create({
      data: {
        site_id,
        equipment_name,
        serial_number,
        jio_min,
        status: status || 'RECEIVED',
        logged_by: req.user.userId
      }
    });
    res.status(201).json(material);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to log material.', details: error.message });
  }
};

export const updateMaterial = async (req: AuthRequest, res: Response) => {
  try {
    const { id } = req.params as { id: string };
    const { status, jio_mrn } = req.body;
    
    const material = await prisma.material.update({
      where: { id },
      data: { status, jio_mrn }
    });
    res.json(material);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to update material.', details: error.message });
  }
};

export const kitMaterials = async (req: AuthRequest, res: Response) => {
  try {
    const { site_id, team_id, material_ids } = req.body;
    
    // Multi-update materials to KITTED and assign to site/team
    await prisma.material.updateMany({
      where: { id: { in: material_ids } },
      data: {
        site_id,
        team_id: team_id as any,
        status: 'KITTED',
        kitted_at: new Date()
      } as any
    });

    // Update site status to KITTED
    await prisma.site.update({
      where: { id: site_id },
      data: { teamId: team_id as any, status: 'KITTED' }
    });

    res.json({ message: 'Kitting completed successfully.' });
  } catch (error: any) {
    res.status(500).json({ error: 'Kitting failed.', details: error.message });
  }
};

export const reconcileMaterials = async (req: AuthRequest, res: Response) => {
  try {
    const { material_ids } = req.body;
    
    await prisma.material.updateMany({
      where: { id: { in: material_ids } },
      data: {
        status: 'RECONCILED',
        reconciled_at: new Date()
      } as any
    });

    res.json({ message: 'Materials reconciled.' });
  } catch (error: any) {
    res.status(500).json({ error: 'Reconciliation failed.', details: error.message });
  }
};

export const updateSiteStage = async (req: AuthRequest, res: Response) => {
  try {
    const { id } = req.params as { id: string };
    const { status, mac_address } = req.body;
    
    const site = await prisma.site.update({
      where: { id },
      data: { 
        status: status as any, 
        mac_address: mac_address as string | undefined 
      }
    });
    res.json(site);
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to update site stage.', details: error.message });
  }
};
