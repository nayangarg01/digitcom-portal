import bcrypt from 'bcrypt';
import prisma from './prisma';

export async function seedAdmin() {
  const username = 'admin';
  const password = 'adminPassword123'; // Recommended to change this immediately after login

  const existingUser = await prisma.user.findUnique({ where: { username } });
  if (existingUser) {
    console.log('Auth: Admin user already exists.');
    return;
  }

  const passwordHash = await bcrypt.hash(password, 10);
  
  await prisma.user.create({
    data: {
      username,
      password_hash: passwordHash,
      role: 'ADMIN',
    },
  });

  console.log('Auth: Initial Admin user seeded successfully.');
  console.log('Auth: Username: admin');
  console.log('Auth: Password: adminPassword123');
}
