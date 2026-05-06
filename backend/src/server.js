const express = require('express');
const cors = require('cors');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const dotenv = require('dotenv');
const fs = require('fs/promises');
const path = require('path');

dotenv.config();

const app = express();

const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET;
const FRONTEND_URL = process.env.FRONTEND_URL;
const USERS_FILE = path.join(__dirname, '..', 'data', 'users.json');

app.use(cors({
  origin: FRONTEND_URL,
  credentials: true
}));

app.use(express.json());

async function readUsers() {
  const data = await fs.readFile(USERS_FILE, 'utf-8');
  return JSON.parse(data);
}

async function writeUsers(users) {
  await fs.writeFile(USERS_FILE, JSON.stringify(users, null, 2));
}

function createToken(user) {
  return jwt.sign(
    {
      id: user.id,
      name: user.name,
      email: user.email
    },
    JWT_SECRET,
    {
      expiresIn: '8h'
    }
  );
}

function authMiddleware(req, res, next) {
  const authHeader = req.headers.authorization;

  if (!authHeader) {
    return res.status(401).json({ message: 'Token não informado.' });
  }

  const [, token] = authHeader.split(' ');

  if (!token) {
    return res.status(401).json({ message: 'Token inválido.' });
  }

  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded;
    next();
  } catch {
    return res.status(401).json({ message: 'Token expirado ou inválido.' });
  }
}

app.post('/auth/register', async (req, res) => {
  const { name, email, password } = req.body;

  if (!name || !email || !password) {
    return res.status(400).json({ message: 'Nome, email e senha são obrigatórios.' });
  }

  const users = await readUsers();

  const emailAlreadyExists = users.some(user => user.email === email);

  if (emailAlreadyExists) {
    return res.status(409).json({ message: 'Este email já está cadastrado.' });
  }

  const passwordHash = await bcrypt.hash(password, 10);

  const newUser = {
    id: crypto.randomUUID(),
    name,
    email,
    passwordHash,
    active: true,
    createdAt: new Date().toISOString()
  };

  users.push(newUser);

  await writeUsers(users);

  return res.status(201).json({
    message: 'Usuário criado com sucesso.',
    user: {
      id: newUser.id,
      name: newUser.name,
      email: newUser.email,
      active: newUser.active
    }
  });
});

app.post('/auth/login', async (req, res) => {
  const { email, password } = req.body;

  if (!email || !password) {
    return res.status(400).json({ message: 'Email e senha são obrigatórios.' });
  }

  const users = await readUsers();

  const user = users.find(item => item.email === email);

  if (!user) {
    return res.status(401).json({ message: 'Email ou senha inválidos.' });
  }

  if (!user.active) {
    return res.status(403).json({ message: 'Usuário inativo.' });
  }

  const passwordIsValid = await bcrypt.compare(password, user.passwordHash);

  if (!passwordIsValid) {
    return res.status(401).json({ message: 'Email ou senha inválidos.' });
  }

  const token = createToken(user);

  return res.json({
    message: 'Login realizado com sucesso.',
    token,
    user: {
      id: user.id,
      name: user.name,
      email: user.email
    }
  });
});

app.get('/auth/me', authMiddleware, async (req, res) => {
  const users = await readUsers();

  const user = users.find(item => item.id === req.user.id);

  if (!user) {
    return res.status(404).json({ message: 'Usuário não encontrado.' });
  }

  return res.json({
    id: user.id,
    name: user.name,
    email: user.email,
    active: user.active
  });
});

app.get('/users', authMiddleware, async (req, res) => {
  const users = await readUsers();

  const safeUsers = users.map(user => ({
    id: user.id,
    name: user.name,
    email: user.email,
    active: user.active,
    createdAt: user.createdAt
  }));

  return res.json(safeUsers);
});

app.patch('/users/:id/status', authMiddleware, async (req, res) => {
  const { id } = req.params;
  const { active } = req.body;

  if (typeof active !== 'boolean') {
    return res.status(400).json({ message: 'Status inválido.' });
  }

  const users = await readUsers();

  const userIndex = users.findIndex(user => user.id === id);

  if (userIndex === -1) {
    return res.status(404).json({ message: 'Usuário não encontrado.' });
  }

  users[userIndex].active = active;

  await writeUsers(users);

  return res.json({
    message: 'Status atualizado com sucesso.'
  });
});

app.get('/', (req, res) => {
  res.json({ message: 'API funcionando' });
});

app.listen(PORT, () => {
  console.log(`API rodando na porta ${PORT}`);
});