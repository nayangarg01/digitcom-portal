import { Request, Response } from 'express';
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';
import prisma from '../utils/prisma';

const JWT_SECRET = process.env.JWT_SECRET || 'super_secret_jwt_key_for_telecom_logistics';

export const signup = async (req: Request, res: Response) => {
  try {
    const { username, password, role } = req.body;

    // Check if user already exists
    const existingUser = await prisma.user.findUnique({ where: { username } });
    if (existingUser) {
      return res.status(400).json({ error: 'Username already taken.' });
    }

    const passwordHash = await bcrypt.hash(password, 10);

    const user = await prisma.user.create({
      data: {
        username,
        password_hash: passwordHash,
        role: role || 'WAREHOUSE_MANAGER',
      },
    });

    res.status(201).json({ message: 'User created successfully.', userId: user.id });
  } catch (error: any) {
    res.status(500).json({ error: 'Failed to create user.', details: error.message });
  }
};

export const login = async (req: Request, res: Response) => {
  try {
    const { username, password } = req.body;
    console.log(`Auth: Login attempt for user: ${username}`);

    const user = await prisma.user.findUnique({ where: { username } });
    if (!user) {
      console.warn(`Auth: User not found: ${username}`);
      return res.status(401).json({ error: 'Invalid username or password.' });
    }

    const isMatch = await bcrypt.compare(password, user.password_hash);
    if (!isMatch) {
      console.warn(`Auth: Password mismatch for user: ${username}`);
      return res.status(401).json({ error: 'Invalid username or password.' });
    }

    const token = jwt.sign(
      { userId: user.id, role: user.role },
      JWT_SECRET,
      { expiresIn: '24h' }
    );

    console.log(`Auth: Successful login for user: ${username}`);
    res.json({
      status: 'success',
      token,
      user: {
        id: user.id,
        username: user.username,
        role: user.role,
      },
    });
  } catch (error: any) {
    console.error(`Auth: Login exception:`, error);
    res.status(500).json({ error: 'Login failed.', details: error.message });
  }
};
