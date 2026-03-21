import bcrypt from 'bcrypt';
import prisma from './prisma';

async function main() {
  const username = 'admin';
  const password = 'adminPassword123'; // Recommended to change this immediately after login

  const existingUser = await prisma.user.findUnique({ where: { username } });
  if (existingUser) {
    console.log('Admin user already exists.');
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

  console.log('Initial Admin user seeded successfully.');
  console.log('Username: admin');
  console.log('Password: adminPassword123');
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
