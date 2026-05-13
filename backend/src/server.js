const express = require('express');
const cors = require('cors');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const dotenv = require('dotenv');

dotenv.config();

const app = express();

const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET;
const FRONTEND_URL = process.env.FRONTEND_URL;

// Usuários fixos — senhas gerenciadas via hash bcrypt
const USERS = [
  {
    id: '8740c0de-ab12-4d23-9fec-69cc231a8944',
    name: 'Administrador',
    email: 'admin@empresa.com',
    passwordHash: '$2b$10$5.ITh6n7ND0EdDxt.SbopeOYfjTMoLpQa79hh9P9O67eWyQHfRluW',
    active: true
  },
  {
    id: 'bfaf2743-0366-44dd-a96f-29df42a1743f',
    name: 'Ana Paula',
    email: 'anapaula@santoamarolimpeza.com.br',
    passwordHash: '$2b$10$bWDqZfGNHTVPS8OkCU3t6.DJVVx.G1sok3iFa2uhYUFfYnflZ3KvW',
    active: true
  },
  {
    id: 'ab5cb4a3-ce30-4316-ba95-e249a7931823',
    name: 'Vendas',
    email: 'vendas@santoamarolimpeza.com.br',
    passwordHash: '$2b$10$hDTWkXvEmm07PhFUp8AmgOJoWRtFX1azX1kMdtK.SZtVKNDpur.G6',
    active: true
  },
  {
    id: '2b278497-10e8-4277-8046-82d0bb8a16ca',
    name: 'Loja Virtual',
    email: 'lojavirtual@santoamarolimpeza.com.br',
    passwordHash: '$2b$10$be32TVltcd8jL6UJVbvlWOFRtqVtlUyNvFBnNTMGPlR9VMrH4j3Nq',
    active: true
  }
];

app.use(cors({
  origin: FRONTEND_URL,
  credentials: true
}));

app.use(express.json());

function createToken(user) {
  return jwt.sign(
    {
      id: user.id,
      name: user.name,
      email: user.email
    },
    JWT_SECRET,
    { expiresIn: '8h' }
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

app.post('/auth/login', async (req, res) => {
  const { email, password } = req.body;

  if (!email || !password) {
    return res.status(400).json({ message: 'Email e senha são obrigatórios.' });
  }

  const user = USERS.find(u => u.email.toLowerCase() === email.toLowerCase());

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

app.get('/auth/me', authMiddleware, (req, res) => {
  const user = USERS.find(u => u.id === req.user.id);

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

app.get('/users', authMiddleware, (req, res) => {
  const safeUsers = USERS.map(({ id, name, email, active }) => ({
    id, name, email, active
  }));

  return res.json(safeUsers);
});

app.get('/', (req, res) => {
  res.json({ message: 'API funcionando' });
});

app.listen(PORT, () => {
  console.log(`API rodando na porta ${PORT}`);
});