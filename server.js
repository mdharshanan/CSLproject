require('dotenv').config();
const path = require('path');
const express = require('express');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const bcrypt = require('bcrypt');
const cors = require('cors');
const session = require('express-session');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'lms-session-secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    secure: false,
    sameSite: 'lax',
    maxAge: 24 * 60 * 60 * 1000,
  },
}));

const protectedPages = new Set([
  '/admin.html',
  '/ai.html',
  '/assignments.html',
  '/business.html',
  '/cloud-computing.html',
  '/courses.html',
  '/css.html',
  '/cyber-security.html',
  '/data-analytics.html',
  '/data-science.html',
  '/finance.html',
  '/html.html',
  '/javascript.html',
  '/leadership.html',
  '/machine-learning.html',
  '/management.html',
  '/marketing.html',
  '/node.html',
  '/power-bi.html',
  '/programming.html',
  '/python.html',
  '/react.html',
  '/sql.html',
  '/study-materials.html',
  '/web-development.html',
]);

function requireLogin(req, res, next) {
  if (req.session && req.session.user) {
    return next();
  }
  return res.redirect('/login.html?next=' + encodeURIComponent(req.originalUrl));
}

function requireAdmin(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 'admin') {
    return next();
  }
  return res.redirect('/login.html?role=admin&next=' + encodeURIComponent(req.originalUrl));
}

app.use((req, res, next) => {
  if (req.path === '/admin.html') {
    return requireAdmin(req, res, next);
  }

  if (protectedPages.has(req.path)) {
    return requireLogin(req, res, next);
  }

  return next();
});

app.use(express.static(path.join(__dirname, '.')));

const mongoUri = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/lms';
console.log("Mongo URI =", process.env.MONGO_URI);

mongoose.connect(mongoUri, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
}).then(async () => {
  console.log('Connected to MongoDB');
  await ensureAdminUser();
  const port = parseInt(process.env.PORT, 10) || 5000;
  app.listen(port, () => {
    console.log(`LMS backend running at http://localhost:${port}`);
  });
}).catch((error) => {
  console.error('MongoDB connection error:', error);
});

const userSchema = new mongoose.Schema({
  name: { type: String, default: 'LMS User' },
  email: { type: String, required: true, unique: true },
  mobile: { type: String, default: '' },
  passwordHash: { type: String, required: true },
  role: { type: String, enum: ['admin', 'user'], default: 'user' },
  createdAt: { type: Date, default: Date.now },
});

const loginOtpSchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
  emailOtpHash: { type: String, required: true },
  mobileOtpHash: { type: String, required: true },
  expiresAt: { type: Date, required: true, index: { expires: 0 } },
  createdAt: { type: Date, default: Date.now },
});

const courseSchema = new mongoose.Schema({
  title: { type: String, required: true },
  category: { type: String, required: true },
  teacher: { type: String, required: true },
  duration: { type: String, required: true },
  startDate: { type: Date, required: true },
  certificateEnabled: { type: Boolean, default: false },
  certificateRequirement: { type: String },
  description: { type: String },
  createdAt: { type: Date, default: Date.now },
});

const contactSchema = new mongoose.Schema({
  name: String,
  email: String,
  message: String,
  createdAt: { type: Date, default: Date.now },
});

const User = mongoose.model('User', userSchema);
const LoginOtp = mongoose.model('LoginOtp', loginOtpSchema);
const Course = mongoose.model('Course', courseSchema);
const Contact = mongoose.model('Contact', contactSchema);

const emailConfigured = Boolean(process.env.EMAIL_USER && process.env.EMAIL_PASS);
const transporter = emailConfigured
  ? nodemailer.createTransport({
      host: process.env.EMAIL_HOST || 'smtp.gmail.com',
      port: parseInt(process.env.EMAIL_PORT, 10) || 587,
      secure: false,
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    })
  : null;

async function sendEmail({ to, subject, html, text }) {
  if (!emailConfigured) {
    console.log('Email not sent because EMAIL_USER and EMAIL_PASS are missing.');
    console.log(`To: ${to}`);
    console.log(`Subject: ${subject}`);
    console.log(text || html || '');
    return { skipped: true };
  }

  const fromAddress = process.env.EMAIL_USER || 'no-reply@lms.local';
  return transporter.sendMail({
    from: fromAddress,
    to,
    subject,
    text,
    html,
  });
}

function generateOtp() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

async function sendMobileOtp({ mobile, otp }) {
  if (!mobile) {
    return;
  }

  console.log(`Mobile OTP for ${mobile}: ${otp}`);
}

async function ensureAdminUser() {
  const email = (process.env.ADMIN_EMAIL || 'admin@lms.edu').trim().toLowerCase();
  const password = process.env.ADMIN_PASSWORD || 'Admin@123';
  const existing = await User.findOne({ email });
  if (!existing) {
    const passwordHash = await bcrypt.hash(password, 10);
    await User.create({
      name: 'Admin',
      email,
      mobile: process.env.ADMIN_MOBILE || '',
      passwordHash,
      role: 'admin',
    });
    console.log('Default admin user created:', email);
  }
}

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.post('/api/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required.' });
  }

  const normalizedEmail = email.trim().toLowerCase();
  const user = await User.findOne({ email: normalizedEmail });
  if (!user) {
    return res.status(401).json({ error: 'Invalid email or password.' });
  }

  const valid = await bcrypt.compare(password, user.passwordHash);
  if (!valid) {
    return res.status(401).json({ error: 'Invalid email or password.' });
  }

  // Set session
  req.session.user = { email: user.email, role: user.role };
  return res.json({ role: user.role, email: user.email });
});

app.post('/api/verify-login-otp', async (req, res) => {
  const { email, emailOtp, mobileOtp } = req.body;
  if (!email || !emailOtp || !mobileOtp) {
    return res.status(400).json({ error: 'Email OTP and mobile OTP are required.' });
  }

  const normalizedEmail = email.trim().toLowerCase();
  const user = await User.findOne({ email: normalizedEmail });
  if (!user) {
    return res.status(401).json({ error: 'Invalid OTP request.' });
  }

  const otpRecord = await LoginOtp.findOne({ userId: user._id });
  if (!otpRecord || otpRecord.expiresAt < new Date()) {
    return res.status(401).json({ error: 'OTP has expired. Please sign in again.' });
  }

  const [emailOtpValid, mobileOtpValid] = await Promise.all([
    bcrypt.compare(emailOtp.trim(), otpRecord.emailOtpHash),
    bcrypt.compare(mobileOtp.trim(), otpRecord.mobileOtpHash),
  ]);

  if (!emailOtpValid || !mobileOtpValid) {
    return res.status(401).json({ error: 'Invalid OTP. Please check both codes.' });
  }

  await LoginOtp.deleteMany({ userId: user._id });
  req.session.user = { email: user.email, role: user.role };
  return res.json({ role: user.role, email: user.email });
});

app.post('/api/register', async (req, res) => {
  const { name, email, password } = req.body;
  if (!name || !email || !password) {
    return res.status(400).json({ error: 'Name, email, and password are required.' });
  }

  const normalizedEmail = email.trim().toLowerCase();
  const existing = await User.findOne({ email: normalizedEmail });
  if (existing) {
    return res.status(409).json({ error: 'A user already exists with that email.' });
  }

  const passwordHash = await bcrypt.hash(password, 10);
  const user = await User.create({
    name: name.trim(),
    email: normalizedEmail,
    password: '',
    passwordHash,
  });

  req.session.user = { email: user.email, role: user.role };

  try {
    await sendEmail({
      to: user.email,
      subject: 'Welcome to LearnHub',
      text: `Hello ${user.name},\n\nYour account has been created successfully.`,
      html: `<p>Hello <strong>${user.name}</strong>,</p><p>Your account has been created successfully.</p>`,
    });
  } catch (error) {
    console.warn('Failed to send welcome email:', error.message);
  }

  return res.status(201).json({ message: 'Registration successful.', role: user.role });
});

app.post('/api/contact', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !email || !message) {
    return res.status(400).json({ error: 'Name, email, and message are required.' });
  }

  await Contact.create({ name, email, message });

  const adminEmail = process.env.ADMIN_NOTIFICATION_EMAIL || process.env.ADMIN_EMAIL || 'admin@lms.edu';

  try {
    await sendEmail({
      to: adminEmail,
      subject: 'New LMS Contact Request',
      text: `New message from ${name} (${email}):\n\n${message}`,
      html: `<p><strong>New message from ${name} (${email}):</strong></p><p>${message}</p>`,
    });
  } catch (error) {
    console.warn('Contact email failed:', error.message);
  }

  return res.json({ message: 'Contact message received.' });
});

app.post('/api/courses', requireAdmin, async (req, res) => {
  const {
    title,
    category,
    teacher,
    duration,
    startDate,
    certificateEnabled,
    certificateRequirement,
    description,
  } = req.body;

  if (!title || !category || !teacher || !duration || !startDate) {
    return res.status(400).json({ error: 'Required fields are missing.' });
  }

  const course = await Course.create({
    title: title.trim(),
    category: category.trim(),
    teacher: teacher.trim(),
    duration: duration.trim(),
    startDate: new Date(startDate),
    certificateEnabled: certificateEnabled === true || certificateEnabled === 'true',
    certificateRequirement: certificateRequirement ? certificateRequirement.trim() : '',
    description: description ? description.trim() : '',
  });

  try {
    await sendEmail({
      to: process.env.ADMIN_NOTIFICATION_EMAIL || process.env.ADMIN_EMAIL || 'admin@lms.edu',
      subject: 'New Course Created',
      text: `A new course was created: ${course.title} (${course.category})`,
      html: `<p>A new course was created:</p><ul><li><strong>${course.title}</strong></li><li>${course.category}</li><li>${course.teacher}</li></ul>`,
    });
  } catch (error) {
    console.warn('Course notification email failed:', error.message);
  }

  return res.status(201).json({ message: 'Course created successfully.', course });
});

app.post('/api/teacher/courses', async (req, res) => {
  const {
    title,
    category,
    teacher,
    duration,
    startDate,
    description,
  } = req.body;

  if (!title || !category || !teacher || !duration || !startDate) {
    return res.status(400).json({ error: 'Course title, category, teacher, duration, and start date are required.' });
  }

  const course = await Course.create({
    title: title.trim(),
    category: category.trim(),
    teacher: teacher.trim(),
    duration: duration.trim(),
    startDate: new Date(startDate),
    certificateEnabled: false,
    certificateRequirement: '',
    description: description ? description.trim() : '',
  });

  return res.status(201).json({ message: 'Course created successfully.', course });
});

app.get('/api/courses', requireLogin, async (req, res) => {
  const courses = await Course.find().sort({ createdAt: -1 }).limit(50);
  res.json(courses);
});

app.post('/api/logout', (req, res) => {
  req.session.destroy((error) => {
    if (error) {
      console.error('Logout error:', error);
      return res.status(500).json({ error: 'Logout failed.' });
    }
    res.clearCookie('connect.sid');
    return res.json({ message: 'Logged out.' });
  });
});

app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ error: 'Unexpected server error.' });
});
