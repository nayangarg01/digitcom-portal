import dotenv from 'dotenv';
dotenv.config();

import express, { Request, Response } from 'express';
import cors from 'cors';
import authRoutes from './routes/authRoutes';
import logisticsRoutes from './routes/logisticsRoutes';
import teamRoutes from './routes/teamRoutes';
import routePlanningRoutes from './routes/routePlanningRoutes';
import financeRoutes from './routes/financeRoutes';
import documentRoutes from './routes/documentRoutes';
import path from 'path';

const app = express();
const PORT = process.env.PORT || 10000; // Use 10000 as default for Render if PORT is missing

app.use(cors());
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, '../uploads')));

// Main API Routes
app.use('/api/auth', authRoutes);
app.use('/api/logistics', logisticsRoutes);
app.use('/api/teams', teamRoutes);
app.use('/api/route-planning', routePlanningRoutes);
app.use('/api/finance', financeRoutes);
app.use('/api/documents', documentRoutes);

// Basic health check route
app.get('/api/health', (req: Request, res: Response) => {
  res.json({ status: 'success', message: 'Backend portal is running smoothly.' });
});

// Error Handling Middleware
app.use((err: any, req: Request, res: Response, next: any) => {
  console.error('API Error:', err);
  res.status(err.status || 500).json({
    error: err.message || 'Internal Server Error',
  });
});

// Start the server
app.listen(Number(PORT), '0.0.0.0', () => {
  console.log(`Server is running on port ${PORT}`);
  console.log(`Checking PORT: ${process.env.PORT || 'undefined (using default 3000)'}`);
});
