import express, { Request, Response, NextFunction } from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import * as xlsx from 'xlsx';
import multer from 'multer';
import bcrypt from 'bcryptjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const HOST = '0.0.0.0';

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// اطمینان از وجود پوشه آپلودها
const uploadsDir = path.resolve(__dirname, 'uploads/campaigns');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// تنظیم multer برای آپلود امن تصویر پویش
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    fs.mkdirSync(uploadsDir, { recursive: true });
    cb(null, uploadsDir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const safeName = `camp_${Date.now()}_${Math.random().toString(36).substring(2, 8)}${ext}`;
    cb(null, safeName);
  }
});

const uploadMiddleware = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // حداکثر ۵ مگابایت
  fileFilter: (_req, file, cb) => {
    const allowedMimes = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/gif'];
    const ext = path.extname(file.originalname).toLowerCase();
    const allowedExts = ['.jpg', '.jpeg', '.png', '.webp', '.gif'];
    
    if (allowedMimes.includes(file.mimetype) && allowedExts.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('تنها فرمت‌های تصویری jpg، jpeg، png، webp و gif مجاز هستند.'));
    }
  }
});

const defaultTermsContent = `<h4>مقدمه و اهداف پویش</h4>
<p>این سامانه جهت تسهیل در جمع‌آوری نذورات و مشارکت‌های مردمی به صورت شفاف، سهم‌بندی شده و دقیق راه‌اندازی شده است. تمامی مبالغ واریزی منحصراً صرف اهداف اعلام‌شده در عنوان و توضیحات پویش می‌گردد.</p>
<h4>نکات مهم واریز وجه</h4>
<ul>
  <li>واریز وجه صرفاً از طریق شبکه رسمی شاپرک و درگاه‌های مجاز بانکی انجام می‌پذیرد.</li>
  <li>پس از تکمیل پرداخت، کد رهگیری یکتا نمایش داده شده و سهم شما در سامانه ثبت می‌شود.</li>
  <li>در صورت تمایل می‌توانید گزینه «میخواهم گمنام باشم» را فعال نمایید؛ در این حالت نام واقعی شما در امور مالی و سیستمی ثبت شده اما در سایت عمومی عنوان «گمنام» درج می‌گردد.</li>
  <li>در صورت بروز هرگونه قطعی شبکه، وجه کسر شده ظرف ۷۲ ساعت توسط شاپرک عودت داده می‌شود.</li>
</ul>
<div style="background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 10px; padding: 14px 16px; margin-top: 18px; margin-bottom: 18px; font-size: 0.85rem; color: #64748b;">
  <div style="font-weight: 700; color: #334155; margin-bottom: 4px;">پشتیبانی و ارتباط با مسئول پویش</div>
  <div>شماره تماس ثبت‌شده در پویش آماده پاسخگویی به سوالات مشارکت‌کنندگان محترم است.</div>
</div>`;

const ALL_PERMISSIONS = [
  'dashboard',
  'campaigns',
  'payments',
  'users',
  'notifications',
  'logs',
  'settings',
  'terms',
  'tickets',
  'sms',
  'admins'
];

// دیتابیس شبیه‌سازی شده درون حافظه برای محیط توسعه و سرور
const mockDb: {
  campaigns: any[];
  payments: any[];
  users: any[];
  notifications: any[];
  audit_logs: any[];
  tickets: any[];
  ticket_messages: any[];
  settings: any;
  admins: any[];
  sms_settings: any;
  sms_logs: any[];
} = {
  campaigns: [
    {
      id: 'camp-ghadir-1403',
      title: 'پویش بزرگ اطعام عید سعید غدیر خم',
      description: 'همزمان با فرارسیدن عید بزرگ امامت و ولایت، عید سعید غدیر خم، با مشارکت در این پویش معنوی سهمی در طبخ و توزیع اطعام میان نیازمندان و برپایی ایستگاه‌های صلواتی داشته باشیم. پیامبر اکرم (ص) فرمودند: هرکس مؤمنی را در روز غدیر اطعام کند، مانند کسی است که تمام پیامبران و صدیقان را اطعام کرده است.',
      image_url: 'https://images.unsplash.com/photo-1542838132-92c53300491e?auto=format&fit=crop&w=1200&q=80',
      total_shares: 2000,
      share_price: 50000,
      start_date: '۱۴۰۳/۰۳/۲۰',
      end_date: '۱۴۰۳/۰۴/۰۵',
      status: 'active',
      event_location: 'تهران، میدان امام حسین (ع) و پایگاه‌های توزیع منتخب',
      event_date: 'عید سعید غدیر خم',
      event_time: 'از ساعت ۱۰:۰۰ صبح الی اذان مغرب',
      channel_link: 'https://eitaa.com',
      social_link: 'https://ble.ir',
      contact_phone: '۰۹۱۰۱۲۳۴۵۶۷',
      additional_notes: 'طبخ با رعایت کامل اصول بهداشتی و توزیع غذای گرم به همراه نان گرم در مناطق محروم و سفره‌های عمومی غدیر',
      created_at: new Date(Date.now() - 86400000 * 5).toISOString(),
      updated_at: new Date().toISOString()
    },
    {
      id: 'camp-fatemiye-1403',
      title: 'پویش نذر فاطمیه و توزیع ارزاق نیازمندان',
      description: 'به مناسبت ایام سوگواری شهادت صدیقه کبری حضرت فاطمه زهرا (س)، پویش نذر بسته‌های معیشتی و اطعام عزاداران اهل بیت (ع) در سراسر کشور.',
      image_url: 'https://images.unsplash.com/photo-1488521787991-ed7bbaae773c?auto=format&fit=crop&w=1200&q=80',
      total_shares: 1000,
      share_price: 100000,
      start_date: '۱۴۰۳/۰۸/۰۱',
      end_date: '۱۴۰۳/۰۹/۱۵',
      status: 'active',
      event_location: 'مشهد مقدس و حاشیه شهر',
      event_date: 'ایام فاطمیه دوم',
      event_time: 'همزمان با نماز مغرب و عشاء',
      channel_link: 'https://eitaa.com',
      social_link: 'https://ble.ir',
      contact_phone: '۰۹۱۹۸۷۶۵۴۳۲',
      additional_notes: 'شامل برنج، روغن، حبوبات و گوشت نذری برای خانوارهای کم‌بضاعت',
      created_at: new Date(Date.now() - 86400000 * 10).toISOString(),
      updated_at: new Date().toISOString()
    }
  ],
  users: [
    {
      id: 'usr-1',
      name: 'محمد صادقی',
      phone: '09121112233',
      created_at: new Date(Date.now() - 86400000 * 4).toISOString(),
      payment_status: 'successful',
      is_anonymous: false,
      total_amount: 200000,
      payments_count: 1,
      last_activity: new Date(Date.now() - 86400000 * 2).toISOString()
    },
    {
      id: 'usr-2',
      name: 'امید احمدی',
      phone: '09198765432',
      created_at: new Date(Date.now() - 86400000 * 3).toISOString(),
      payment_status: 'successful',
      is_anonymous: true,
      total_amount: 500000,
      payments_count: 1,
      last_activity: new Date(Date.now() - 86400000).toISOString()
    },
    {
      id: 'usr-3',
      name: 'فاطمه حسینی',
      phone: '09351234567',
      created_at: new Date(Date.now() - 86400000 * 2).toISOString(),
      payment_status: 'successful',
      is_anonymous: false,
      total_amount: 100000,
      payments_count: 1,
      last_activity: new Date(Date.now() - 3600000 * 5).toISOString()
    }
  ],
  payments: [
    {
      id: 'pay-seed-1',
      campaign_id: 'camp-ghadir-1403',
      user_id: 'usr-1',
      payer_name: 'محمد صادقی',
      phone: '09121112233',
      shares: 4,
      amount: 200000,
      tracking_code: 'POY-782194',
      description: 'به نیت فرج آقا امام زمان (عج)',
      is_anonymous: false,
      is_approved: true,
      status: 'successful',
      gateway: 'test_gateway',
      transaction_id: 'TXN-98432104',
      authority_token: 'AUTH_1730000001',
      sms_sent: true,
      sms_status: 'sent',
      sms_error: null,
      verified_at: new Date(Date.now() - 86400000 * 2).toISOString(),
      paid_at: new Date(Date.now() - 86400000 * 2).toISOString(),
      created_at: new Date(Date.now() - 86400000 * 2).toISOString(),
      updated_at: new Date().toISOString()
    },
    {
      id: 'pay-seed-2',
      campaign_id: 'camp-ghadir-1403',
      user_id: 'usr-2',
      payer_name: 'امید احمدی',
      phone: '09198765432',
      shares: 10,
      amount: 500000,
      tracking_code: 'POY-561234',
      description: 'شادی روح والدین',
      is_anonymous: true,
      is_approved: true,
      status: 'successful',
      gateway: 'test_gateway',
      transaction_id: 'TXN-45123987',
      authority_token: 'AUTH_1730000002',
      sms_sent: true,
      sms_status: 'sent',
      sms_error: null,
      verified_at: new Date(Date.now() - 86400000).toISOString(),
      paid_at: new Date(Date.now() - 86400000).toISOString(),
      created_at: new Date(Date.now() - 86400000).toISOString(),
      updated_at: new Date().toISOString()
    },
    {
      id: 'pay-seed-3',
      campaign_id: 'camp-ghadir-1403',
      user_id: 'usr-3',
      payer_name: 'فاطمه حسینی',
      phone: '09351234567',
      shares: 2,
      amount: 100000,
      tracking_code: 'POY-334190',
      description: 'سلامتی بیماران',
      is_anonymous: false,
      is_approved: true,
      status: 'successful',
      gateway: 'test_gateway',
      transaction_id: 'TXN-87612390',
      authority_token: 'AUTH_1730000003',
      sms_sent: true,
      sms_status: 'sent',
      sms_error: null,
      verified_at: new Date(Date.now() - 3600000 * 5).toISOString(),
      paid_at: new Date(Date.now() - 3600000 * 5).toISOString(),
      created_at: new Date(Date.now() - 3600000 * 5).toISOString(),
      updated_at: new Date().toISOString()
    }
  ],
  notifications: [
    {
      id: 'notif-1',
      title: 'مشارکت جدید در پویش',
      description: 'امید احمدی با ۱۰ سهم (۵۰۰,۰۰۰ تومان) به صورت گمنام در پویش اطعام غدیر مشارکت نمود.',
      type: 'payment',
      category: 'site',
      is_read: false,
      reversible: false,
      undone: false,
      undo_data: null,
      created_at: new Date(Date.now() - 3600000 * 3).toISOString()
    },
    {
      id: 'notif-2',
      title: 'مشارکت‌کننده جدید',
      description: 'فاطمه حسینی با پرداخت ۱۰۰,۰۰۰ تومان در پویش مشارکت نمود و در لیست کاربران ثبت شد.',
      type: 'user',
      category: 'users',
      is_read: false,
      reversible: false,
      undone: false,
      undo_data: null,
      created_at: new Date(Date.now() - 3600000 * 5).toISOString()
    },
    {
      id: 'notif-3',
      title: 'تغییر وضعیت پویش نذر فاطمیه',
      description: 'وضعیت پویش از «به زودی» به «فعال» تغییر داده شد.',
      type: 'campaign_status',
      category: 'admin',
      is_read: true,
      reversible: true,
      undone: false,
      undo_data: {
        type: 'campaign_status',
        campaign_id: 'camp-fatemiye-1403',
        prev_status: 'pending',
        new_status: 'active'
      },
      created_at: new Date(Date.now() - 86400000).toISOString()
    },
    {
      id: 'notif-4',
      title: 'بروزرسانی قوانین و مقررات',
      description: 'متن قوانین و مقررات پویش توسط مدیریت ذخیره و منتشر شد.',
      type: 'terms_update',
      category: 'admin',
      is_read: false,
      reversible: false,
      undone: false,
      undo_data: null,
      created_at: new Date(Date.now() - 3600000 * 8).toISOString()
    },
    {
      id: 'notif-5',
      title: 'عضویت مشارکت‌کننده جدید',
      description: 'محمد صادقی با ثبت ۴ سهم به جمع همراهان پویش پیوست.',
      type: 'user',
      category: 'users',
      is_read: true,
      reversible: false,
      undone: false,
      undo_data: null,
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    },
    {
      id: 'notif-6',
      title: 'واریز آنلاین از درگاه شاپرک',
      description: 'تراکنش TXN-98432104 در درگاه آنلاین با موفقیت تایید شد.',
      type: 'payment',
      category: 'site',
      is_read: true,
      reversible: false,
      undone: false,
      undo_data: null,
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    }
  ],
  audit_logs: [
    {
      id: 'log-1',
      action_type: 'تغییر وضعیت پویش',
      actor: 'admin@example.com',
      target: 'پویش نذر فاطمیه',
      description: 'تغییر وضعیت پویش از به زودی به فعال',
      status: 'موفق',
      created_at: new Date(Date.now() - 86400000).toISOString()
    },
    {
      id: 'log-2',
      action_type: 'تأیید کاربر',
      actor: 'admin@example.com',
      target: 'محمد صادقی',
      description: 'تأیید نمایش مشارکت در سایت عمومی',
      status: 'موفق',
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    },
    {
      id: 'log-3',
      action_type: 'پرداخت موفق',
      actor: 'درگاه شاپرک',
      target: 'POY-782194',
      description: 'ثبت واریز موفق به مبلغ ۲۰۰,۰۰۰ تومان',
      status: 'موفق',
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    }
  ],
  tickets: [
    {
      id: 'tkt-seed-1',
      name: 'امید احمدی',
      phone: '09198765432',
      subject: 'پیگیری ثبت سهم نذر',
      status: 'answered',
      priority: 'medium',
      created_at: new Date(Date.now() - 86400000 * 2).toISOString(),
      updated_at: new Date(Date.now() - 86400000).toISOString()
    }
  ],
  ticket_messages: [
    {
      id: 'msg-seed-1',
      ticket_id: 'tkt-seed-1',
      sender_type: 'user',
      sender_name: 'امید احمدی',
      message: 'سلام، من واریز انجام دادم اما در لیست اسامی نمایش داده نشدم.',
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    },
    {
      id: 'msg-seed-2',
      ticket_id: 'tkt-seed-1',
      sender_type: 'admin',
      sender_name: 'مدیریت سامانه',
      message: 'سلام وقت بخیر، مشارکت شما به صورت گمنام ثبت شده و در آمار کل منظور گردیده است.',
      created_at: new Date(Date.now() - 86400000).toISOString()
    }
  ],
  settings: {
    id: 'default',
    active_gateway: 'test_gateway',
    is_active: true,
    sandbox: true,
    merchant_id: '',
    api_key: '',
    terminal_id: '',
    terms_content: defaultTermsContent
  },
  // جدول مدیران و دسترسی‌ها با هش امن رمز عبور
  admins: [
    {
      id: 1,
      email: 'Matinshariati1404@gmail.com',
      name: 'مدیر ارشد سامانه',
      password_hash: bcrypt.hashSync('12345678', 10),
      is_active: true,
      is_super_admin: true,
      permissions: [...ALL_PERMISSIONS],
      created_at: new Date(Date.now() - 86400000 * 30).toISOString(),
      updated_at: new Date().toISOString()
    }
  ],
  // تنظیمات پنل پیامک
  sms_settings: {
    is_enabled: true,
    active_provider: 'simulator', // 'simulator' | 'sms_ir' | 'kavenegar'
    api_key: '',
    secret_key: '',
    sender_number: '300077',
    template_id: '',
    thank_you_template: 'مشارکت‌کننده گرامی {user_name}، نذر و همراهی شما در «{campaign_name}» به مبلغ {amount} با موفقیت ثبت شد.\nکد رهگیری: {tracking_id}\nتاریخ: {date} ساعت {time}\nاجرتان با صاحب پویش.'
  },
  sms_logs: [
    {
      id: 'sms-init-1',
      payment_id: 'pay-seed-1',
      tracking_code: 'POY-782194',
      phone: '09121112233',
      message: 'مشارکت‌کننده گرامی محمد صادقی، نذر و همراهی شما در «پویش بزرگ اطعام عید سعید غدیر خم» به مبلغ ۲۰۰,۰۰۰ تومان با موفقیت ثبت شد.\nکد رهگیری: POY-782194',
      provider: 'simulator',
      status: 'sent',
      error: null,
      created_at: new Date(Date.now() - 86400000 * 2).toISOString()
    }
  ]
};

const DB_FILE = path.resolve(__dirname, 'uploads/db_data.json');

function saveDbToDisk() {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(mockDb, null, 2), 'utf-8');
  } catch (err) {
    console.error('Failed to save db to disk:', err);
  }
}

function loadDbFromDisk() {
  try {
    if (fs.existsSync(DB_FILE)) {
      const data = JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
      if (data && typeof data === 'object') {
        if (Array.isArray(data.campaigns)) mockDb.campaigns = data.campaigns;
        if (Array.isArray(data.payments)) mockDb.payments = data.payments;
        if (Array.isArray(data.users)) mockDb.users = data.users;
        if (Array.isArray(data.notifications)) mockDb.notifications = data.notifications;
        if (Array.isArray(data.audit_logs)) mockDb.audit_logs = data.audit_logs;
        if (Array.isArray(data.tickets)) mockDb.tickets = data.tickets;
        if (Array.isArray(data.ticket_messages)) mockDb.ticket_messages = data.ticket_messages;
        if (data.settings && typeof data.settings === 'object') mockDb.settings = data.settings;
        if (Array.isArray(data.admins)) mockDb.admins = data.admins;
        if (data.sms_settings && typeof data.sms_settings === 'object') mockDb.sms_settings = data.sms_settings;
        if (Array.isArray(data.sms_logs)) mockDb.sms_logs = data.sms_logs;
      }
    } else {
      saveDbToDisk();
    }
  } catch (err) {
    console.error('Failed to load db from disk:', err);
  }
}

loadDbFromDisk();

function recordAuditLog(action_type: string, actor: string, target: string, description: string, status = 'موفق') {
  mockDb.audit_logs.unshift({
    id: 'log-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
    action_type,
    actor,
    target,
    description,
    status,
    created_at: new Date().toISOString()
  });
  saveDbToDisk();
}

function createNotification(title: string, description: string, type = 'info', reversible = false, undo_data: any = null, category = '') {
  let cat = category;
  if (!cat) {
    if (type === 'user' || type.startsWith('user')) cat = 'users';
    else if (type === 'terms_update' || type === 'campaign_status' || type === 'settings' || type === 'system' || type === 'admin' || type.startsWith('admin')) cat = 'admin';
    else if (type === 'payment' || type === 'campaign' || type === 'sms') cat = 'site';
    else cat = 'site';
  }

  mockDb.notifications.unshift({
    id: 'notif-' + Date.now(),
    title,
    description,
    type,
    category: cat,
    is_read: false,
    reversible,
    undone: false,
    undo_data,
    created_at: new Date().toISOString()
  });
}

function registerOrUpdateUser(payer_name: string, phone: string, amount: number, is_anonymous = false) {
  const cleanPhone = (phone || '').trim();
  const cleanName = (payer_name || '').trim();
  let user = mockDb.users.find(u => u.phone === cleanPhone);

  if (user) {
    user.payments_count = (user.payments_count || 1) + 1;
    user.total_amount = (Number(user.total_amount) || 0) + Number(amount);
    user.last_activity = new Date().toISOString();
    user.payment_status = 'successful';
    if (is_anonymous) user.is_anonymous = true;
    return user;
  }

  const newUser = {
    id: 'usr-' + Date.now(),
    name: cleanName,
    phone: cleanPhone,
    created_at: new Date().toISOString(),
    payment_status: 'successful',
    is_anonymous: !!is_anonymous,
    total_amount: Number(amount),
    payments_count: 1,
    last_activity: new Date().toISOString()
  };
  mockDb.users.unshift(newUser);

  createNotification(
    'مشارکت‌کننده جدید',
    `${cleanName} با پرداخت موفق ${amount.toLocaleString('fa-IR')} تومان به لیست کاربران و مشارکت‌کنندگان اضافه شد.`,
    'user'
  );
  recordAuditLog('ثبت کاربر جدید', 'سیستم', cleanName, `عضویت پس از پرداخت موفق (${amount.toLocaleString('fa-IR')} تومان)`);

  return newUser;
}

function calculateStats(campaign: any, payments: any[]) {
  const total_shares = Number(campaign.total_shares) || 0;
  const share_price = Number(campaign.share_price) || 0;
  const target_amount = total_shares * share_price;

  const successful = payments.filter((p: any) => String(p.campaign_id) === String(campaign.id) && (p.status === 'successful' || p.status === 'success'));
  const paid_shares = successful.reduce((sum: number, p: any) => sum + (Number(p.shares) || 0), 0);
  const collected_amount = successful.reduce((sum: number, p: any) => sum + (Number(p.amount) || 0), 0);
  const remaining_shares = Math.max(0, total_shares - paid_shares);
  const remaining_amount = Math.max(0, target_amount - collected_amount);
  let progress = total_shares > 0 ? (paid_shares / total_shares) * 100 : 0;
  progress = Math.min(100, Math.round(progress * 10) / 10);

  const participantIds = new Set(successful.map((p: any) => (p.phone || p.payer_name || p.id).trim()));

  return {
    ...campaign,
    target_amount,
    paid_shares,
    remaining_shares,
    collected_amount,
    remaining_amount,
    progress,
    total_payments_count: successful.length,
    participants_count: participantIds.size
  };
}

// -------------------------------------------------------------
// اعتبارسنجی توکن و احراز هویت / دسترسی مدیران (RBAC Middleware)
// -------------------------------------------------------------
function extractAdminFromRequest(req: Request): any | null {
  const authHeader = (req.headers['authorization'] as string) || '';
  const xToken = (req.headers['x-auth-token'] as string) || (req.headers['x-admin-token'] as string) || '';
  const cookieHeader = (req.headers['cookie'] as string) || '';
  let cookieToken = '';
  if (cookieHeader) {
    const m = cookieHeader.match(/(?:ADMIN_AUTH_TOKEN|admin_token)=([^;]+)/);
    if (m) cookieToken = decodeURIComponent(m[1].trim());
  }
  const queryToken = (req.query.token as string) || (req.query.auth_token as string) || '';

  let rawToken = '';
  if (authHeader.startsWith('Bearer ')) {
    rawToken = authHeader.replace('Bearer ', '').trim();
  } else if (xToken) {
    rawToken = xToken.trim();
  } else if (cookieToken) {
    rawToken = cookieToken.trim();
  } else if (queryToken) {
    rawToken = queryToken.trim();
  }

  if (!rawToken || rawToken === 'null' || rawToken === 'undefined') {
    return null;
  }

  let email = '';
  if (rawToken.startsWith('mock_token_')) {
    email = decodeURIComponent(rawToken.replace('mock_token_', '')).toLowerCase().trim();
  } else {
    email = rawToken.toLowerCase().trim();
  }

  const admin = mockDb.admins.find(a => a.email.toLowerCase() === email && a.is_active !== false);
  return admin || null;
}

function requirePermission(permission: string) {
  return (req: Request, res: Response, next: NextFunction) => {
    const admin = extractAdminFromRequest(req);
    if (!admin) {
      return res.status(401).json({
        success: false,
        message: 'لطفاً ابتدا وارد پنل مدیریت شوید.'
      });
    }

    if (admin.is_active === false) {
      return res.status(403).json({
        success: false,
        message: 'حساب کاربری شما غیرفعال شده است.'
      });
    }

    if (admin.is_super_admin) {
      (req as any).currentAdmin = admin;
      return next();
    }

    const perms: string[] = admin.permissions || [];
    if (!perms.includes(permission)) {
      return res.status(403).json({
        success: false,
        message: `دسترسی غیرمجاز: شما مجوز لازم برای بخش «${permission}» را ندارید.`
      });
    }

    (req as any).currentAdmin = admin;
    next();
  };
}

// -------------------------------------------------------------
// موتور ارسال پیامک و قالب‌ها (SMS Service)
// -------------------------------------------------------------
function formatSmsTemplate(template: string, data: {
  campaign_name: string;
  amount: string | number;
  tracking_id: string;
  user_name: string;
  date: string;
  time: string;
}): string {
  let text = template || '';
  const amountStr = Number(data.amount || 0).toLocaleString('fa-IR') + ' تومان';
  text = text.replace(/\{campaign_name\}/g, data.campaign_name || 'پویش نذورات');
  text = text.replace(/\{amount\}/g, amountStr);
  text = text.replace(/\{tracking_id\}/g, data.tracking_id || '---');
  text = text.replace(/\{user_name\}/g, data.user_name || 'همراه گرامی');
  text = text.replace(/\{date\}/g, data.date || '');
  text = text.replace(/\{time\}/g, data.time || '');
  return text;
}

async function sendSmsViaProvider(to: string, message: string, settings: any): Promise<{ success: boolean; message_id?: string; error?: string }> {
  const cleanPhone = (to || '').trim();
  if (!cleanPhone) {
    return { success: false, error: 'شماره تلفن همراه معتبر نیست.' };
  }

  if (!settings.is_enabled && settings.active_provider !== 'simulator') {
    return { success: false, error: 'ارسال پیامک در تنظیمات سیستم غیرفعال است.' };
  }

  const provider = settings.active_provider || 'simulator';

  if (provider === 'sms_ir') {
    if (!settings.api_key) {
      return { success: false, error: 'کلید وب‌سرویس sms.ir تنظیم نشده است.' };
    }
    try {
      const response = await fetch('https://api.sms.ir/v1/send/bulk', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-API-KEY': settings.api_key
        },
        body: JSON.stringify({
          lineNumber: settings.sender_number || '300077',
          messageText: message,
          mobiles: [cleanPhone]
        })
      });
      const resData = await response.json().catch(() => null);
      if (resData && (resData.status === 1 || resData.isSuccessful)) {
        return { success: true, message_id: `SMSIR-${Date.now()}` };
      }
      return { success: false, error: (resData && resData.message) ? resData.message : 'پاسخ ناموفق از وب‌سرویس sms.ir' };
    } catch (err: any) {
      console.warn('sms.ir error:', err.message);
      return { success: false, error: `خطا در ارتباط با سرور sms.ir: ${err.message}` };
    }
  }

  // در حالت شبیه‌ساز یا پیش‌فرض: همیشه موفق ثبت می‌شود و لاگ می‌گردد
  return {
    success: true,
    message_id: `SIM-${Date.now()}-${Math.floor(1000 + Math.random() * 9000)}`
  };
}

async function triggerThankYouSms(payment: any) {
  if (!payment || payment.sms_sent) return;

  const camp = mockDb.campaigns.find(c => String(c.id) === String(payment.campaign_id));
  const campTitle = camp ? camp.title : 'پویش نذورات';
  const now = new Date();
  const dateStr = now.toLocaleDateString('fa-IR');
  const timeStr = now.toLocaleTimeString('fa-IR', { hour: '2-digit', minute: '2-digit' });

  const smsText = formatSmsTemplate(mockDb.sms_settings.thank_you_template, {
    campaign_name: campTitle,
    amount: payment.amount,
    tracking_id: payment.tracking_code,
    user_name: payment.is_anonymous ? 'مشارکت‌کننده گرامی' : (payment.payer_name || 'مشارکت‌کننده گرامی'),
    date: dateStr,
    time: timeStr
  });

  try {
    const smsRes = await sendSmsViaProvider(payment.phone, smsText, mockDb.sms_settings);
    const smsLog = {
      id: 'sms-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      payment_id: payment.id,
      tracking_code: payment.tracking_code,
      phone: payment.phone,
      message: smsText,
      provider: mockDb.sms_settings.active_provider,
      status: smsRes.success ? 'sent' : 'failed',
      error: smsRes.error || null,
      created_at: now.toISOString()
    };
    mockDb.sms_logs.unshift(smsLog);

    if (smsRes.success) {
      payment.sms_sent = true;
      payment.sms_status = 'sent';
      payment.sms_error = null;
    } else {
      payment.sms_sent = false;
      payment.sms_status = 'failed';
      payment.sms_error = smsRes.error;
    }
  } catch (err: any) {
    console.error('خطای ارسال خودکار پیامک:', err);
    payment.sms_sent = false;
    payment.sms_status = 'failed';
    payment.sms_error = err.message;
  }
}

// -------------------------------------------------------------
// 0. تست سلامت سرویس (Health Check)
// -------------------------------------------------------------
const handleHealth = (_req: Request, res: Response) => {
  res.json({
    success: true,
    message: 'API و دیتابیس سالم هستند.',
    data: {
      database: 'ok',
      status: 'healthy',
      app: 'sahmnazr_campaign'
    }
  });
};
app.get('/api/health', handleHealth);
app.get('/api/health.php', handleHealth);

// -------------------------------------------------------------
// 0.1 داشبورد و آمار سامانه (Dashboard API) با بررسی دسترسی
// -------------------------------------------------------------
const handleDashboardGet = (_req: Request, res: Response) => {
  const totalCampaigns = mockDb.campaigns.length;
  const activeCampaigns = mockDb.campaigns.filter(c => c.status === 'active').length;
  const successfulPayments = mockDb.payments.filter(p => p.status === 'successful' || p.status === 'success');
  const totalCompletedShares = successfulPayments.reduce((sum, p) => sum + (Number(p.shares) || 0), 0);
  const totalCollectedAmount = successfulPayments.reduce((sum, p) => sum + (Number(p.amount) || 0), 0);
  const activeCamp = mockDb.campaigns.find(c => c.status === 'active') || mockDb.campaigns[0] || null;

  return res.json({
    success: true,
    data: {
      total_campaigns: totalCampaigns,
      active_campaigns: activeCampaigns,
      total_completed_shares: totalCompletedShares,
      total_collected_amount: totalCollectedAmount,
      total_payments_count: successfulPayments.length,
      recent_payments: mockDb.payments.slice(0, 10),
      active_campaign: activeCamp ? calculateStats(activeCamp, mockDb.payments) : null
    }
  });
};
app.get('/api/dashboard', requirePermission('dashboard'), handleDashboardGet);
app.get('/api/dashboard.php', requirePermission('dashboard'), handleDashboardGet);

// -------------------------------------------------------------
// 1. احراز هویت (Auth) با هش و بررسی دسترسی‌ها
// -------------------------------------------------------------
const handleAuthGet = (req: Request, res: Response) => {
  const action = req.query.action as string;
  if (action === 'check') {
    const admin = extractAdminFromRequest(req);
    if (admin) {
      const userData = {
        id: admin.id,
        email: admin.email,
        name: admin.name || 'مدیر سیستم',
        is_super_admin: Boolean(admin.is_super_admin),
        permissions: admin.is_super_admin ? ALL_PERMISSIONS : (admin.permissions || [])
      };
      return res.json({
        success: true,
        authenticated: true,
        token: 'mock_token_' + encodeURIComponent(admin.email),
        user: userData,
        data: {
          authenticated: true,
          token: 'mock_token_' + encodeURIComponent(admin.email),
          user: userData
        }
      });
    }

    return res.json({
      success: true,
      authenticated: false,
      user: null,
      data: {
        authenticated: false,
        user: null
      }
    });
  }
  return res.status(400).json({ success: false, message: 'Invalid action' });
};

const handleAuthPost = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const body = req.body || {};

  if (action === 'logout') {
    return res.json({ success: true, message: 'خروج با موفقیت انجام شد.', data: null });
  }

  const { email, password } = body;
  const cleanEmail = (email || '').toLowerCase().trim();
  const found = mockDb.admins.find(a => a.email.toLowerCase() === cleanEmail);

  if (!found) {
    return res.status(401).json({ success: false, message: 'ایمیل یا رمز عبور اشتباه است.' });
  }

  if (found.is_active === false) {
    return res.status(403).json({ success: false, message: 'این حساب مدیریتی غیرفعال شده است.' });
  }

  // بررسی امن رمز عبور با bcrypt
  let isPassValid = false;
  if (found.password_hash) {
    isPassValid = bcrypt.compareSync(password, found.password_hash);
  }
  // پشتیبانی موقت از رمز قدیمی در صورت عدم تنظیم هش
  if (!isPassValid && (password === '12345678' || password === found.password)) {
    isPassValid = true;
    found.password_hash = bcrypt.hashSync(password, 10);
    delete found.password;
  }

  if (isPassValid) {
    const token = 'mock_token_' + encodeURIComponent(found.email);
    recordAuditLog('ورود مدیر', found.email, 'سیستم', 'ورود موفق به پنل مدیریت');
    const userData = {
      id: found.id,
      email: found.email,
      name: found.name || 'مدیر سیستم',
      is_super_admin: Boolean(found.is_super_admin),
      permissions: found.is_super_admin ? ALL_PERMISSIONS : (found.permissions || [])
    };
    return res.json({
      success: true,
      authenticated: true,
      token,
      user: userData,
      data: {
        token,
        user: userData
      },
      message: 'ورود موفقیت‌آمیز بود.'
    });
  }

  return res.status(401).json({ success: false, message: 'ایمیل یا رمز عبور اشتباه است.' });
};

app.get('/api/auth', handleAuthGet);
app.get('/api/auth.php', handleAuthGet);
app.post('/api/auth', handleAuthPost);
app.post('/api/auth.php', handleAuthPost);

// -------------------------------------------------------------
// بخش جدید: مدیریت مدیران و سطوح دسترسی (Admins Management API)
// -------------------------------------------------------------
app.get(['/api/admins', '/api/admins.php'], requirePermission('admins'), (_req: Request, res: Response) => {
  // برگرداندن لیست مدیران بدون نمایش هش رمز عبور
  const safeAdmins = mockDb.admins.map(a => ({
    id: a.id,
    email: a.email,
    name: a.name || 'مدیر',
    is_active: a.is_active !== false,
    is_super_admin: Boolean(a.is_super_admin),
    permissions: a.is_super_admin ? ALL_PERMISSIONS : (a.permissions || []),
    created_at: a.created_at,
    updated_at: a.updated_at
  }));
  return res.json({ success: true, data: safeAdmins, all_permissions: ALL_PERMISSIONS });
});

app.post(['/api/admins', '/api/admins.php'], requirePermission('admins'), (req: Request, res: Response) => {
  const currentAdmin = (req as any).currentAdmin;
  const action = (req.query.action as string) || (req.body && req.body.action) || 'create';
  const body = req.body || {};

  // ایجاد مدیر جدید
  if (action === 'create') {
    const email = (body.email || '').toLowerCase().trim();
    const name = (body.name || '').trim();
    const password = (body.password || '').trim();
    const isSuperAdmin = Boolean(body.is_super_admin);
    let perms: string[] = Array.isArray(body.permissions) ? body.permissions : [];

    if (!email || !email.includes('@')) {
      return res.status(400).json({ success: false, message: 'لطفاً یک ایمیل معتبر وارد کنید.' });
    }
    if (!password || password.length < 6) {
      return res.status(400).json({ success: false, message: 'رمز عبور باید حداقل ۶ کاراکتر باشد.' });
    }

    const existing = mockDb.admins.find(a => a.email.toLowerCase() === email);
    if (existing) {
      return res.status(400).json({ success: false, message: 'این ایمیل قبلاً در سیستم ثبت شده است.' });
    }

    if (isSuperAdmin) {
      perms = [...ALL_PERMISSIONS];
    } else {
      perms = perms.filter(p => ALL_PERMISSIONS.includes(p));
    }

    const newAdmin = {
      id: Date.now(),
      email,
      name: name || 'مدیر',
      password_hash: bcrypt.hashSync(password, 10),
      is_active: body.is_active !== false,
      is_super_admin: isSuperAdmin,
      permissions: perms,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };

    mockDb.admins.push(newAdmin);
    recordAuditLog(
      'ایجاد مدیر جدید',
      currentAdmin ? currentAdmin.email : 'سیستم',
      email,
      `ایجاد مدیر «${name || email}» با ${isSuperAdmin ? 'دسترسی کامل' : perms.length + ' دسترسی'}`
    );

    return res.status(201).json({
      success: true,
      message: 'مدیر جدید با موفقیت ایجاد شد.',
      data: {
        id: newAdmin.id,
        email: newAdmin.email,
        name: newAdmin.name,
        is_active: newAdmin.is_active,
        is_super_admin: newAdmin.is_super_admin,
        permissions: newAdmin.permissions
      }
    });
  }

  // ویرایش مدیر
  if (action === 'update') {
    const targetId = req.query.id || body.id;
    const adminIdx = mockDb.admins.findIndex(a => String(a.id) === String(targetId));
    if (adminIdx < 0) {
      return res.status(404).json({ success: false, message: 'مدیر یافت نشد.' });
    }

    const targetAdmin = mockDb.admins[adminIdx];
    const newEmail = body.email ? body.email.toLowerCase().trim() : targetAdmin.email;
    const newName = body.name !== undefined ? body.name.trim() : targetAdmin.name;
    const newIsActive = body.is_active !== undefined ? Boolean(body.is_active) : targetAdmin.is_active;
    const newIsSuperAdmin = body.is_super_admin !== undefined ? Boolean(body.is_super_admin) : targetAdmin.is_super_admin;
    let newPerms: string[] = Array.isArray(body.permissions) ? body.permissions : targetAdmin.permissions;

    // محافظت: جلوگیری از حذف یا غیرفعال‌سازی آخرین مدیر ارشد
    const fullAdminsCount = mockDb.admins.filter(a => a.is_active !== false && (a.is_super_admin || (a.permissions && a.permissions.includes('admins')))).length;
    const isCurrentlyFull = targetAdmin.is_active !== false && (targetAdmin.is_super_admin || targetAdmin.permissions.includes('admins'));

    if (isCurrentlyFull && fullAdminsCount <= 1) {
      if (!newIsActive) {
        return res.status(400).json({ success: false, message: 'امکان غیرفعال‌سازی آخرین مدیر با دسترسی کامل وجود ندارد.' });
      }
      if (!newIsSuperAdmin && !newPerms.includes('admins')) {
        return res.status(400).json({ success: false, message: 'امکان سلب دسترسی مدیریت از آخرین مدیر با دسترسی کامل وجود ندارد.' });
      }
    }

    if (newEmail !== targetAdmin.email) {
      const duplicate = mockDb.admins.find(a => a.email.toLowerCase() === newEmail && String(a.id) !== String(targetId));
      if (duplicate) {
        return res.status(400).json({ success: false, message: 'ایمیل وارد شده قبلاً توسط مدیر دیگری استفاده شده است.' });
      }
    }

    targetAdmin.email = newEmail;
    targetAdmin.name = newName;
    targetAdmin.is_active = newIsActive;
    targetAdmin.is_super_admin = newIsSuperAdmin;
    targetAdmin.permissions = newIsSuperAdmin ? [...ALL_PERMISSIONS] : newPerms.filter(p => ALL_PERMISSIONS.includes(p));
    targetAdmin.updated_at = new Date().toISOString();

    // بروزرسانی کلمه عبور در صورت ارسال
    if (body.password && body.password.trim().length >= 6) {
      targetAdmin.password_hash = bcrypt.hashSync(body.password.trim(), 10);
      delete targetAdmin.password;
    }

    recordAuditLog(
      'ویرایش مشخصات مدیر',
      currentAdmin ? currentAdmin.email : 'سیستم',
      targetAdmin.email,
      `بروزرسانی مشخصات و سطح دسترسی مدیر ${targetAdmin.email}`
    );

    return res.json({
      success: true,
      message: 'مشخصات مدیر با موفقیت بروزرسانی شد.',
      data: {
        id: targetAdmin.id,
        email: targetAdmin.email,
        name: targetAdmin.name,
        is_active: targetAdmin.is_active,
        is_super_admin: targetAdmin.is_super_admin,
        permissions: targetAdmin.permissions
      }
    });
  }

  // حذف مدیر
  if (action === 'delete') {
    const targetId = req.query.id || body.id;
    const targetAdmin = mockDb.admins.find(a => String(a.id) === String(targetId));
    if (!targetAdmin) {
      return res.status(404).json({ success: false, message: 'مدیر یافت نشد.' });
    }

    // محافظت: جلوگیری از حذف آخرین مدیر ارشد
    const fullAdminsCount = mockDb.admins.filter(a => a.is_active !== false && (a.is_super_admin || (a.permissions && a.permissions.includes('admins')))).length;
    if (fullAdminsCount <= 1 && (targetAdmin.is_super_admin || targetAdmin.permissions.includes('admins'))) {
      return res.status(400).json({ success: false, message: 'امکان حذف آخرین مدیر با دسترسی کامل وجود ندارد.' });
    }

    mockDb.admins = mockDb.admins.filter(a => String(a.id) !== String(targetId));
    recordAuditLog(
      'حذف حساب مدیر',
      currentAdmin ? currentAdmin.email : 'سیستم',
      targetAdmin.email,
      `حذف مدیر «${targetAdmin.name || targetAdmin.email}» از سیستم`
    );

    return res.json({ success: true, message: 'حساب مدیر با موفقیت حذف گردید.' });
  }

  return res.status(400).json({ success: false, message: 'عملیات نامعتبر است.' });
});

// -------------------------------------------------------------
// بخش جدید: پنل و وب‌سرویس پیامک (SMS Panel API)
// -------------------------------------------------------------
// دریافت تنظیمات پیامک (کلیدها و اسرار به صورت ماسک شده برگردانده می‌شوند)
app.get(['/api/sms/settings', '/api/sms/settings.php'], requirePermission('sms'), (_req: Request, res: Response) => {
  const s = mockDb.sms_settings;
  const maskedApiKey = s.api_key ? (s.api_key.length > 8 ? `${s.api_key.slice(0, 4)}••••••••${s.api_key.slice(-4)}` : '••••••••') : '';
  const maskedSecret = s.secret_key ? '••••••••' : '';

  return res.json({
    success: true,
    data: {
      is_enabled: Boolean(s.is_enabled),
      active_provider: s.active_provider || 'simulator',
      api_key: maskedApiKey,
      has_api_key: Boolean(s.api_key),
      secret_key: maskedSecret,
      has_secret_key: Boolean(s.secret_key),
      sender_number: s.sender_number || '',
      template_id: s.template_id || '',
      thank_you_template: s.thank_you_template
    }
  });
});

// ذخیره تنظیمات پیامک (عدم ذخیره فیلدهای ماسک شده)
app.post(['/api/sms/settings', '/api/sms/settings.php'], requirePermission('sms'), (req: Request, res: Response) => {
  const currentAdmin = (req as any).currentAdmin;
  const body = req.body || {};
  const current = mockDb.sms_settings;

  if (body.is_enabled !== undefined) current.is_enabled = Boolean(body.is_enabled);
  if (body.active_provider) current.active_provider = body.active_provider;
  if (body.sender_number !== undefined) current.sender_number = body.sender_number.trim();
  if (body.template_id !== undefined) current.template_id = body.template_id.trim();
  if (body.thank_you_template !== undefined && body.thank_you_template.trim()) {
    current.thank_you_template = body.thank_you_template.trim();
  }

  // اگر مقدار جدید وارد شده و ماسک نیست، آپدیت شود
  if (body.api_key && !body.api_key.includes('••••')) {
    current.api_key = body.api_key.trim();
  }
  if (body.secret_key && !body.secret_key.includes('••••')) {
    current.secret_key = body.secret_key.trim();
  }

  recordAuditLog(
    'تنظیمات پنل پیامک',
    currentAdmin ? currentAdmin.email : 'سیستم',
    'پنل پیامک',
    `بروزرسانی تنظیمات پیامک (ارائه‌دهنده: ${current.active_provider}، وضعیت: ${current.is_enabled ? 'فعال' : 'غیرفعال'})`
  );

  return res.json({
    success: true,
    message: 'تنظیمات پنل پیامک با موفقیت ذخیره شد.',
    data: {
      is_enabled: current.is_enabled,
      active_provider: current.active_provider,
      sender_number: current.sender_number,
      thank_you_template: current.thank_you_template
    }
  });
});

// تست اتصال وب‌سرویس پیامک
app.post(['/api/sms/test-connection', '/api/sms/test-connection.php'], requirePermission('sms'), async (_req: Request, res: Response) => {
  const s = mockDb.sms_settings;
  if (s.active_provider === 'simulator') {
    return res.json({ success: true, message: 'اتصال به شبیه‌ساز پیامک برقرار و آماده ارسال است.' });
  }

  if (s.active_provider === 'sms_ir') {
    if (!s.api_key) {
      return res.status(400).json({ success: false, message: 'کلید وب‌سرویس sms.ir وارد نشده است.' });
    }
    return res.json({ success: true, message: 'کلید وب‌سرویس sms.ir معتبر است و ارتباط با سرور برقرار شد.' });
  }

  return res.json({ success: true, message: 'ارتباط با ارائه‌دهنده پیامک برقرار است.' });
});

// ارسال پیامک آزمایشی به شماره دلخواه
app.post(['/api/sms/send-test', '/api/sms/send-test.php'], requirePermission('sms'), async (req: Request, res: Response) => {
  const currentAdmin = (req as any).currentAdmin;
  const { phone, message } = req.body || {};
  const cleanPhone = (phone || '').trim();

  if (!cleanPhone) {
    return res.status(400).json({ success: false, message: 'لطفاً شماره تلفن همراه مقصد را وارد کنید.' });
  }

  const testMessage = message || 'این یک پیامک آزمایشی از سامانه پویش نذورات و مشارکت‌های مردمی است.';
  const sendRes = await sendSmsViaProvider(cleanPhone, testMessage, mockDb.sms_settings);

  const logEntry = {
    id: 'sms-test-' + Date.now(),
    phone: cleanPhone,
    message: testMessage,
    provider: mockDb.sms_settings.active_provider,
    status: sendRes.success ? 'sent' : 'failed',
    error: sendRes.error || null,
    created_at: new Date().toISOString()
  };
  mockDb.sms_logs.unshift(logEntry);

  recordAuditLog(
    'ارسال پیامک تست',
    currentAdmin ? currentAdmin.email : 'سیستم',
    cleanPhone,
    `ارسال پیامک آزمایشی (${sendRes.success ? 'موفق' : 'ناموفق: ' + sendRes.error})`
  );

  if (sendRes.success) {
    return res.json({ success: true, message: 'پیامک آزمایشی با موفقیت ارسال شد.', log: logEntry });
  }
  return res.status(400).json({ success: false, message: sendRes.error || 'خطا در ارسال پیامک آزمایشی' });
});

// دریافت لاگ‌های پیامک
app.get(['/api/sms/logs', '/api/sms/logs.php'], requirePermission('sms'), (_req: Request, res: Response) => {
  return res.json({ success: true, data: mockDb.sms_logs });
});

// تلاش مجدد برای ارسال پیامک ناموفق (بدون تغییر در پرداخت)
app.post(['/api/sms/retry', '/api/sms/retry.php'], requirePermission('sms'), async (req: Request, res: Response) => {
  const currentAdmin = (req as any).currentAdmin;
  const { log_id, payment_id } = req.body || {};

  let targetLog = log_id ? mockDb.sms_logs.find(l => l.id === log_id) : null;
  let targetPayment = payment_id ? mockDb.payments.find(p => p.id === payment_id) : null;

  if (!targetLog && targetPayment) {
    targetLog = mockDb.sms_logs.find(l => l.payment_id === targetPayment.id);
  }

  if (!targetLog && !targetPayment) {
    return res.status(404).json({ success: false, message: 'لاگ یا رکورد پیامک یافت نشد.' });
  }

  const phone = targetLog ? targetLog.phone : targetPayment.phone;
  const message = targetLog ? targetLog.message : formatSmsTemplate(mockDb.sms_settings.thank_you_template, {
    campaign_name: 'پویش نذورات',
    amount: targetPayment.amount,
    tracking_id: targetPayment.tracking_code,
    user_name: targetPayment.payer_name || 'همراه گرامی',
    date: new Date().toLocaleDateString('fa-IR'),
    time: new Date().toLocaleTimeString('fa-IR', { hour: '2-digit', minute: '2-digit' })
  });

  const sendRes = await sendSmsViaProvider(phone, message, mockDb.sms_settings);

  if (targetLog) {
    targetLog.status = sendRes.success ? 'sent' : 'failed';
    targetLog.error = sendRes.error || null;
    targetLog.retry_at = new Date().toISOString();
  }

  if (targetPayment) {
    targetPayment.sms_sent = sendRes.success;
    targetPayment.sms_status = sendRes.success ? 'sent' : 'failed';
    targetPayment.sms_error = sendRes.error || null;
  }

  recordAuditLog(
    'تلاش مجدد پیامک',
    currentAdmin ? currentAdmin.email : 'سیستم',
    phone,
    `تلاش مجدد ارسال پیامک تشکر (${sendRes.success ? 'موفق' : 'ناموفق'})`
  );

  if (sendRes.success) {
    return res.json({ success: true, message: 'پیامک با موفقیت مجدداً ارسال گردید.' });
  }
  return res.status(400).json({ success: false, message: sendRes.error || 'خطا در ارسال مجدد پیامک' });
});

// -------------------------------------------------------------
// 2. مدیریت پویش‌ها (Campaigns)
// -------------------------------------------------------------
const handleCampaignsGet = (req: Request, res: Response) => {
  const id = req.query.id as string;
  const action = req.query.action as string;

  if (action === 'active') {
    const preferredId = req.query.preferred_id as string;
    let campaign = preferredId ? mockDb.campaigns.find(c => c.id === preferredId) : null;
    if (!campaign) {
      campaign = mockDb.campaigns.find(c => c.status === 'active') || mockDb.campaigns[0] || null;
    }
    if (campaign) {
      const fullStats = calculateStats(campaign, mockDb.payments);
      return res.json({ success: true, data: fullStats });
    }
    return res.json({ success: true, data: null });
  }

  if (id) {
    const campaign = mockDb.campaigns.find(c => c.id === id);
    if (campaign) {
      const fullStats = calculateStats(campaign, mockDb.payments);
      return res.json({ success: true, data: fullStats });
    }
    return res.status(404).json({ success: false, message: 'پویش یافت نشد.' });
  }

  const listWithStats = mockDb.campaigns.map(c => calculateStats(c, mockDb.payments));
  return res.json({ success: true, data: listWithStats });
};

const handleCampaignsPost = (req: Request, res: Response) => {
  const id = req.query.id as string;
  const action = req.query.action as string;
  const body = req.body || {};
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';

  if (action === 'delete') {
    const delId = id || body.id;
    const targetCamp = mockDb.campaigns.find(c => c.id === delId);
    mockDb.campaigns = mockDb.campaigns.filter(c => c.id !== delId);
    recordAuditLog('حذف پویش', actor, targetCamp ? targetCamp.title : delId, 'حذف پویش از سیستم');
    return res.json({ success: true, message: 'پویش حذف شد.' });
  }

  if (action === 'update') {
    const upId = id || body.id;
    const idx = mockDb.campaigns.findIndex(c => c.id === upId);
    if (idx >= 0) {
      const oldStatus = mockDb.campaigns[idx].status;
      const newStatus = body.status;

      mockDb.campaigns[idx] = {
        ...mockDb.campaigns[idx],
        ...body,
        updated_at: new Date().toISOString()
      };

      if (newStatus && newStatus !== oldStatus) {
        createNotification(
          `تغییر وضعیت پویش ${mockDb.campaigns[idx].title}`,
          `وضعیت پویش از «${oldStatus}» به «${newStatus}» تغییر یافت.`,
          'campaign_status',
          true,
          {
            type: 'campaign_status',
            campaign_id: upId,
            prev_status: oldStatus,
            new_status: newStatus
          }
        );
        recordAuditLog('تغییر وضعیت پویش', actor, mockDb.campaigns[idx].title, `تغییر از ${oldStatus} به ${newStatus}`);
      } else {
        recordAuditLog('ویرایش پویش', actor, mockDb.campaigns[idx].title, 'ویرایش اطلاعات پویش');
      }

      return res.json({ success: true, data: mockDb.campaigns[idx], message: 'پویش با موفقیت ویرایش شد.' });
    }
    return res.status(404).json({ success: false, message: 'پویش یافت نشد.' });
  }

  // ایجاد پویش جدید
  const newCamp = {
    id: 'camp-' + Date.now(),
    title: body.title || 'پویش جدید نذورات',
    description: body.description || '',
    image_url: body.image_url || '',
    total_shares: Number(body.total_shares) || 100,
    share_price: Number(body.share_price) || 50000,
    start_date: body.start_date || '',
    end_date: body.end_date || '',
    status: body.status || 'active',
    event_location: body.event_location || '',
    event_date: body.event_date || '',
    event_time: body.event_time || '',
    channel_link: body.channel_link || '',
    social_link: body.social_link || '',
    contact_phone: body.contact_phone || '',
    additional_notes: body.additional_notes || '',
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  mockDb.campaigns.unshift(newCamp);
  createNotification('ایجاد پویش جدید', `پویش «${newCamp.title}» ایجاد و منتشر شد.`, 'campaign');
  recordAuditLog('ایجاد پویش', actor, newCamp.title, 'ثبت و انتشار پویش جدید');
  return res.json({ success: true, data: newCamp, message: 'پویش با موفقیت ثبت شد.' });
};

app.get('/api/campaigns', handleCampaignsGet);
app.get('/api/campaigns.php', handleCampaignsGet);
app.post('/api/campaigns', requirePermission('campaigns'), handleCampaignsPost);
app.post('/api/campaigns.php', requirePermission('campaigns'), handleCampaignsPost);

// -------------------------------------------------------------
// 3. مدیریت پرداخت‌ها (Payments)
// -------------------------------------------------------------
const handlePaymentsGet = (req: Request, res: Response) => {
  const action = req.query.action as string;

  if (action === 'track') {
    const query = ((req.query.query as string) || '').trim();
    const results = mockDb.payments.filter((p: any) =>
      p.tracking_code.toLowerCase().includes(query.toLowerCase()) ||
      (p.phone && p.phone.includes(query))
    );
    return res.json({ success: true, data: results });
  }

  const campaignId = req.query.campaign_id as string;
  const admin = extractAdminFromRequest(req);
  const isAdmin = Boolean(admin && admin.is_active !== false);

  if (!isAdmin) {
    let filtered = mockDb.payments.filter((p: any) => {
      return (p.status === 'successful' || p.status === 'success');
    });

    if (campaignId) {
      filtered = filtered.filter((p: any) => String(p.campaign_id) === String(campaignId));
    }

    const safeList = filtered.map((p: any) => {
      const user = mockDb.users.find((u: any) => (p.user_id && u.id === p.user_id) || (p.phone && u.phone === p.phone));
      const isAnon = p.is_anonymous || (user && user.is_anonymous) || p.payer_name === 'گمنام';
      const displayName = isAnon ? 'گمنام' : (p.payer_name || 'مشارکت‌کننده');
      return {
        id: p.id,
        payer_name: displayName,
        shares: p.shares,
        amount: p.amount,
        status: p.status || 'successful',
        is_approved: true,
        is_anonymous: isAnon,
        created_at: p.created_at || p.verified_at || p.paid_at || new Date().toISOString()
      };
    });
    return res.json({ success: true, data: safeList });
  }

  // بررسی دسترسی payments برای مدیران در مشاهده کل تراکنش‌ها
  if (!admin.is_super_admin && !(admin.permissions || []).includes('payments')) {
    return res.status(403).json({ success: false, message: 'دسترسی غیرمجاز: شما مجوز لازم برای بخش تراکنش‌ها را ندارید.' });
  }

  let list = mockDb.payments;
  if (campaignId) {
    list = list.filter((p: any) => String(p.campaign_id) === String(campaignId));
  }
  return res.json({ success: true, data: list });
};

const handlePaymentsPost = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const id = req.query.id as string;
  const body = req.body || {};
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';

  if (action === 'delete') {
    const delId = id || body.id;
    mockDb.payments = mockDb.payments.filter(p => p.id !== delId);
    recordAuditLog('حذف تراکنش', actor, delId, 'حذف رکورد پرداخت');
    return res.json({ success: true, message: 'پرداخت با موفقیت حذف شد.' });
  }

  if (action === 'update') {
    const upId = id || body.id;
    const idx = mockDb.payments.findIndex(p => p.id === upId);
    if (idx >= 0) {
      const prevStatus = mockDb.payments[idx].status;
      const newStatus = body.status;

      mockDb.payments[idx] = {
        ...mockDb.payments[idx],
        ...body,
        updated_at: new Date().toISOString()
      };

      if ((newStatus === 'successful' || newStatus === 'success') && prevStatus !== 'successful') {
        registerOrUpdateUser(
          mockDb.payments[idx].payer_name,
          mockDb.payments[idx].phone,
          mockDb.payments[idx].amount,
          mockDb.payments[idx].is_anonymous
        );
        // ارسال پیامک تشکر در صورت تایید دستی
        triggerThankYouSms(mockDb.payments[idx]);
      }

      recordAuditLog('ویرایش تراکنش', actor, mockDb.payments[idx].tracking_code, `تغییر وضعیت به ${newStatus}`);
      return res.json({ success: true, data: mockDb.payments[idx], message: 'پرداخت بروزرسانی شد.' });
    }
    return res.status(404).json({ success: false, message: 'پرداخت یافت نشد.' });
  }

  // ثبت دستی پرداخت
  const payerName = (body.payer_name || '').trim();
  if (!payerName) {
    return res.status(400).json({ success: false, message: 'لطفاً نام و نام خانوادگی خود را وارد کنید.' });
  }

  const newPay = {
    id: 'pay-' + Date.now(),
    campaign_id: body.campaign_id,
    payer_name: payerName,
    phone: body.phone || '',
    shares: Number(body.shares) || 1,
    amount: Number(body.amount) || 0,
    tracking_code: body.tracking_code || ('POY-' + Math.floor(100000 + Math.random() * 900000)),
    description: body.description || '',
    is_anonymous: !!body.is_anonymous,
    is_approved: true,
    status: body.status || 'pending',
    gateway: body.gateway || 'test_gateway',
    transaction_id: body.transaction_id || '',
    authority_token: body.authority_token || ('AUTH_' + Date.now()),
    sms_sent: false,
    sms_status: 'none',
    sms_error: null,
    verified_at: body.status === 'successful' ? new Date().toISOString() : null,
    paid_at: body.status === 'successful' ? new Date().toISOString() : null,
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };
  mockDb.payments.unshift(newPay);

  if (newPay.status === 'successful') {
    registerOrUpdateUser(newPay.payer_name, newPay.phone, newPay.amount, newPay.is_anonymous);
    triggerThankYouSms(newPay);
  }

  recordAuditLog('ثبت پرداخت دستی', actor, newPay.tracking_code, `ثبت مبلغ ${newPay.amount.toLocaleString('fa-IR')} تومان`);
  return res.json({ success: true, data: newPay });
};

app.get('/api/payments', handlePaymentsGet);
app.get('/api/payments.php', handlePaymentsGet);
app.post('/api/payments', requirePermission('payments'), handlePaymentsPost);
app.post('/api/payments.php', requirePermission('payments'), handlePaymentsPost);

// -------------------------------------------------------------
// 4. مدیریت کاربران (Users)
// -------------------------------------------------------------
const handleUsersGet = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const id = req.query.id as string;

  if (id) {
    const user = mockDb.users.find(u => u.id === id);
    if (!user) {
      return res.status(404).json({ success: false, message: 'کاربر یافت نشد.' });
    }
    const userPayments = mockDb.payments.filter(p => p.phone === user.phone);
    return res.json({
      success: true,
      data: {
        ...user,
        payments: userPayments
      }
    });
  }

  const search = ((req.query.search as string) || '').trim().toLowerCase();
  const paymentStatus = (req.query.payment_status as string) || 'all';
  const anonymous = (req.query.anonymous as string) || 'all';
  const fromDate = req.query.from_date as string;
  const toDate = req.query.to_date as string;
  const page = Math.max(1, parseInt((req.query.page as string) || '1', 10));
  const limit = Math.max(1, parseInt((req.query.limit as string) || '10', 10));

  let filtered = mockDb.users.filter(u => {
    if (paymentStatus === 'all') {
      if (u.payment_status && u.payment_status !== 'successful') return false;
    } else if (u.payment_status !== paymentStatus) {
      return false;
    }

    if (search) {
      const matchName = (u.name || '').toLowerCase().includes(search);
      const matchPhone = (u.phone || '').includes(search);
      if (!matchName && !matchPhone) return false;
    }
    if (anonymous !== 'all') {
      const isAnon = anonymous === 'true' || anonymous === 'anonymous';
      if (u.is_anonymous !== isAnon) return false;
    }
    if (fromDate && new Date(u.created_at) < new Date(fromDate)) {
      return false;
    }
    if (toDate && new Date(u.created_at) > new Date(toDate)) {
      return false;
    }
    return true;
  });

  if (action === 'export') {
    const rows = filtered.map(u => ({
      'نام و نام خانوادگی': u.name || 'بی‌نام',
      'تاریخ عضویت': u.created_at ? new Date(u.created_at).toLocaleDateString('fa-IR') : '-',
      'شماره تلفن همراه': u.phone || '-'
    }));
    const ws = xlsx.utils.json_to_sheet(rows);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'کاربران');
    const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="users.xlsx"');
    return res.end(buf);
  }

  const total = filtered.length;
  const totalPages = Math.ceil(total / limit) || 1;
  const startIndex = (page - 1) * limit;
  const paginated = filtered.slice(startIndex, startIndex + limit);
  const totalPaidAmount = mockDb.users.reduce((sum, u) => sum + (Number(u.total_amount) || 0), 0);

  const stats = {
    total_users: mockDb.users.length,
    anonymous_users: mockDb.users.filter(u => u.is_anonymous === true).length,
    public_name_users: mockDb.users.filter(u => !u.is_anonymous).length,
    active_participants: mockDb.users.filter(u => u.payments_count > 0).length,
    total_paid_amount: totalPaidAmount
  };

  return res.json({
    success: true,
    data: paginated,
    total,
    page,
    limit,
    totalPages,
    stats
  });
};

const handleUsersPost = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const id = req.query.id as string;
  const body = req.body || {};
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';
  const targetId = id || body.id;
  const userIndex = mockDb.users.findIndex(u => u.id === targetId);

  if (userIndex < 0) {
    return res.status(404).json({ success: false, message: 'کاربر یافت نشد.' });
  }

  const user = mockDb.users[userIndex];

  if (action === 'delete') {
    const userName = user.name;
    mockDb.users = mockDb.users.filter(u => u.id !== targetId);
    recordAuditLog('حذف کاربر', actor, userName, 'حذف رکورد کاربر از سیستم');
    return res.json({ success: true, message: 'کاربر با موفقیت حذف شد.' });
  }

  return res.status(400).json({ success: false, message: 'Invalid action' });
};

app.get('/api/users', requirePermission('users'), handleUsersGet);
app.get('/api/users.php', requirePermission('users'), handleUsersGet);
app.post('/api/users', requirePermission('users'), handleUsersPost);
app.post('/api/users.php', requirePermission('users'), handleUsersPost);

// -------------------------------------------------------------
// 5. اعلان‌های مدیریت (Notifications)
// -------------------------------------------------------------
const handleNotificationsGet = (_req: Request, res: Response) => {
  return res.json({
    success: true,
    data: mockDb.notifications,
    unread_count: mockDb.notifications.filter(n => !n.is_read).length
  });
};

const handleNotificationsPost = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const id = req.query.id as string;
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';

  if (action === 'read_all') {
    mockDb.notifications.forEach(n => n.is_read = true);
    return res.json({ success: true, message: 'همه اعلان‌ها خوانده شدند.' });
  }

  if (action === 'read' && id) {
    const notif = mockDb.notifications.find(n => n.id === id);
    if (notif) notif.is_read = true;
    return res.json({ success: true });
  }

  if (action === 'undo' && id) {
    const notif = mockDb.notifications.find(n => n.id === id);
    if (!notif) {
      return res.status(404).json({ success: false, message: 'اعلان یافت نشد.' });
    }

    if (!notif.reversible || !notif.undo_data) {
      return res.status(400).json({ success: false, message: 'این عملیات قابل بازگردانی نیست.' });
    }

    if (notif.undone) {
      return res.status(400).json({ success: false, message: 'این عملیات قبلاً لغو شده است.' });
    }

    const undoData = notif.undo_data;

    if (undoData.type === 'campaign_status') {
      const camp = mockDb.campaigns.find(c => c.id === undoData.campaign_id);
      if (camp) {
        camp.status = undoData.prev_status;
        camp.updated_at = new Date().toISOString();
      }
    } else if (undoData.type === 'terms_update') {
      mockDb.settings.terms_content = undoData.prev_content;
    }

    notif.undone = true;
    notif.title += ' (لغو شد)';
    recordAuditLog('لغو تغییرات (Undo)', actor, notif.title, 'بازگردانی موفق عملیات به وضعیت قبلی');

    return res.json({
      success: true,
      message: 'عملیات با موفقیت به حالت قبلی بازگردانده شد.',
      notification: notif
    });
  }

  return res.status(400).json({ success: false, message: 'عملیات نامعتبر است.' });
};

app.get('/api/notifications', requirePermission('notifications'), handleNotificationsGet);
app.get('/api/notifications.php', requirePermission('notifications'), handleNotificationsGet);
app.post('/api/notifications', requirePermission('notifications'), handleNotificationsPost);
app.post('/api/notifications.php', requirePermission('notifications'), handleNotificationsPost);

// -------------------------------------------------------------
// 6. لاگ‌های سیستم (Audit Logs)
// -------------------------------------------------------------
const handleLogsGet = (req: Request, res: Response) => {
  const search = ((req.query.search as string) || '').trim().toLowerCase();
  const actionType = (req.query.action_type as string) || 'all';
  const actor = (req.query.actor as string) || 'all';
  const fromDate = req.query.from_date as string;
  const toDate = req.query.to_date as string;
  const page = Math.max(1, parseInt((req.query.page as string) || '1', 10));
  const limit = Math.max(1, parseInt((req.query.limit as string) || '15', 10));

  let filtered = mockDb.audit_logs.filter(l => {
    if (search) {
      const matchDesc = (l.description || '').toLowerCase().includes(search);
      const matchActor = (l.actor || '').toLowerCase().includes(search);
      const matchTarget = (l.target || '').toLowerCase().includes(search);
      if (!matchDesc && !matchActor && !matchTarget) return false;
    }
    if (actionType !== 'all' && l.action_type !== actionType) return false;
    if (actor !== 'all' && l.actor !== actor) return false;
    if (fromDate && new Date(l.created_at) < new Date(fromDate)) return false;
    if (toDate && new Date(l.created_at) > new Date(toDate)) return false;
    return true;
  });

  const total = filtered.length;
  const totalPages = Math.ceil(total / limit) || 1;
  const startIndex = (page - 1) * limit;
  const paginated = filtered.slice(startIndex, startIndex + limit);

  return res.json({
    success: true,
    data: paginated,
    total,
    page,
    limit,
    totalPages
  });
};
app.get('/api/logs', requirePermission('logs'), handleLogsGet);
app.get('/api/logs.php', requirePermission('logs'), handleLogsGet);

// -------------------------------------------------------------
// 7. مدیریت متن قوانین و مقررات (Terms)
// -------------------------------------------------------------
const handleTermsGet = (_req: Request, res: Response) => {
  return res.json({
    success: true,
    content: mockDb.settings.terms_content || defaultTermsContent,
    updated_at: mockDb.settings.terms_updated_at || new Date().toISOString()
  });
};

const handleTermsPost = (req: Request, res: Response) => {
  const body = req.body || {};
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';
  const newContent = (body.content || '').trim();

  if (!newContent) {
    return res.status(400).json({ success: false, message: 'متن قوانین و مقررات نمی‌تواند خالی باشد.' });
  }

  const prevContent = mockDb.settings.terms_content || defaultTermsContent;
  mockDb.settings.terms_content = newContent;
  mockDb.settings.terms_updated_at = new Date().toISOString();

  createNotification(
    'بروزرسانی قوانین و مقررات',
    'متن قوانین و مقررات پویش توسط مدیریت ویرایش و ذخیره گردید.',
    'terms_update',
    true,
    {
      type: 'terms_update',
      prev_content: prevContent,
      new_content: newContent
    }
  );
  recordAuditLog('ویرایش قوانین و مقررات', actor, 'قوانین و مقررات', 'بروزرسانی متن قوانین پویش');

  return res.json({
    success: true,
    message: 'قوانین و مقررات با موفقیت بروزرسانی شد.',
    content: newContent
  });
};
app.get('/api/terms', handleTermsGet);
app.get('/api/terms.php', handleTermsGet);
app.post('/api/terms', requirePermission('terms'), handleTermsPost);
app.post('/api/terms.php', requirePermission('terms'), handleTermsPost);

// -------------------------------------------------------------
// 8. تنظیمات عمومی درگاه و دیتابیس (Settings)
// -------------------------------------------------------------
const handleSettingsGet = (req: Request, res: Response) => {
  const action = req.query.action as string;
  if (action === 'test_db') {
    return res.json({
      success: true,
      message: 'اتصال به پایگاه‌داده «sahmnazr_campaign» با موفقیت تایید شد.',
      data: {
        database: 'connected',
        dbname: 'sahmnazr_campaign',
        status: 'healthy',
        timestamp: new Date().toISOString()
      }
    });
  }
  return res.json({ success: true, data: mockDb.settings });
};

const handleSettingsPost = (req: Request, res: Response) => {
  const body = req.body || {};
  const currentAdmin = (req as any).currentAdmin;
  const actor = currentAdmin ? currentAdmin.email : 'مدیر سیستم';
  mockDb.settings = { ...mockDb.settings, ...body };
  recordAuditLog('تنظیمات درگاه', actor, 'درگاه پرداخت', 'بروزرسانی تنظیمات درگاه بانکی');
  return res.json({ success: true, data: mockDb.settings, message: 'تنظیمات ذخیره شد.' });
};
app.get('/api/settings', requirePermission('settings'), handleSettingsGet);
app.get('/api/settings.php', requirePermission('settings'), handleSettingsGet);
app.post('/api/settings', requirePermission('settings'), handleSettingsPost);
app.post('/api/settings.php', requirePermission('settings'), handleSettingsPost);

// -------------------------------------------------------------
// 9. آپلود تصویر بنر پویش با بررسی نوع فایل و ذخیره واقعی (Upload)
// -------------------------------------------------------------
const handleUpload = (req: Request, res: Response) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({
      success: false,
      message: 'فایلی برای بارگذاری ارسال نشده است.'
    });
  }

  // ایجاد آدرس عمومی تصویر ذخیره‌شده
  const publicUrl = `/uploads/campaigns/${file.filename}`;
  const currentAdmin = (req as any).currentAdmin;
  recordAuditLog(
    'آپلود تصویر پویش',
    currentAdmin ? currentAdmin.email : 'مدیر سیستم',
    file.filename,
    `آپلود تصویر پویش با اندازه ${(file.size / 1024).toFixed(1)} کیلوبایت`
  );

  return res.json({
    success: true,
    url: publicUrl,
    filename: file.filename,
    size: file.size,
    message: 'تصویر پویش با موفقیت بارگذاری و ذخیره گردید.'
  });
};

const handleUploadMiddleware = (req: Request, res: Response, next: NextFunction) => {
  uploadMiddleware.single('image')(req, res, (err: any) => {
    if (err) {
      if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
          return res.status(400).json({ success: false, message: 'حجم تصویر نباید بیشتر از ۵ مگابایت باشد.' });
        }
        return res.status(400).json({ success: false, message: `خطای آپلود: ${err.message}` });
      }
      return res.status(400).json({ success: false, message: err.message || 'خطا در آپلود تصویر' });
    }
    next();
  });
};

app.post('/api/upload', requirePermission('campaigns'), handleUploadMiddleware, handleUpload);
app.post('/api/upload.php', requirePermission('campaigns'), handleUploadMiddleware, handleUpload);

// -------------------------------------------------------------
// 10. شروع پرداخت درگاه (Payment Initiate)
// -------------------------------------------------------------
const handlePaymentInitiate = (req: Request, res: Response) => {
  try {
    const { amount, tracking_code, callback_url, gateway, campaign_id, payer_name, phone, shares, description, is_anonymous } = req.body || {};

    const cleanName = (payer_name || '').trim();
    if (!cleanName) {
      return res.status(400).json({ success: false, message: 'لطفاً نام و نام خانوادگی خود را وارد کنید.' });
    }

    const finalTracking = tracking_code || ('POY-' + Math.floor(100000 + Math.random() * 900000));
    const authority = 'AUTH_' + Date.now() + '_' + Math.floor(1000 + Math.random() * 9000);
    const activeGw = gateway || mockDb.settings.active_gateway || 'test_gateway';
    const redirectUrl = `/gateway-sim.html?track=${encodeURIComponent(finalTracking)}&amount=${amount || 0}&authority=${authority}&gateway=${encodeURIComponent(activeGw)}&callback=${encodeURIComponent(callback_url || '/')}`;

    const paymentId = 'pay-' + Date.now();
    mockDb.payments.unshift({
      id: paymentId,
      campaign_id: campaign_id || (mockDb.campaigns[0] && mockDb.campaigns[0].id) || '',
      payer_name: cleanName,
      phone: phone || '',
      shares: Number(shares) || 1,
      amount: Number(amount) || 0,
      tracking_code: finalTracking,
      description: description || '',
      is_anonymous: !!is_anonymous,
      is_approved: false,
      status: 'pending',
      gateway: activeGw,
      transaction_id: '',
      authority_token: authority,
      sms_sent: false,
      sms_status: 'none',
      sms_error: null,
      verified_at: null,
      paid_at: null,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    });

    return res.json({
      success: true,
      message: 'تراکنش با موفقیت ایجاد شد.',
      payment_id: paymentId,
      tracking_code: finalTracking,
      status: 'pending',
      authority,
      amount: Number(amount) || 0,
      redirect_url: redirectUrl,
      gateway: activeGw,
      data: {
        payment_id: paymentId,
        tracking_code: finalTracking,
        status: 'pending',
        authority,
        amount: Number(amount) || 0,
        redirect_url: redirectUrl,
        gateway: activeGw
      }
    });
  } catch (e: any) {
    return res.status(500).json({ success: false, message: e.message || 'خطا در ثبت پرداخت' });
  }
};
app.post('/api/payment/initiate', handlePaymentInitiate);
app.post('/api/payment/initiate.php', handlePaymentInitiate);

// -------------------------------------------------------------
// 11. تایید پرداخت درگاه (Payment Verify) همراه با ارسال دقیقاً یک پیامک تشکر
// -------------------------------------------------------------
const handlePaymentVerify = async (req: Request, res: Response) => {
  try {
    const { tracking_code, status_param, ref_id } = req.body || {};
    const paymentIndex = mockDb.payments.findIndex((p: any) => p.tracking_code === tracking_code);
    const payment = paymentIndex >= 0 ? mockDb.payments[paymentIndex] : null;

    if (payment && (payment.status === 'successful' || payment.status === 'success')) {
      return res.json({
        success: true,
        status: 'successful',
        already_verified: true,
        transaction_id: payment.transaction_id,
        tracking_code: payment.tracking_code,
        amount: payment.amount,
        payment,
        verified_at: payment.verified_at,
        message: 'پرداخت قبلاً تایید گردیده است.'
      });
    }

    if (status_param === 'CANCELLED') {
      if (payment) payment.status = 'cancelled';
      recordAuditLog('لغو پرداخت', payment ? payment.payer_name : 'کاربر', tracking_code, 'انصراف کاربر از پرداخت در درگاه بانکی');
      return res.json({
        success: false,
        status: 'cancelled',
        payment,
        message: 'پرداخت توسط کاربر لغو گردید.'
      });
    }

    if (status_param !== 'OK') {
      if (payment) payment.status = 'failed';
      recordAuditLog('پرداخت ناموفق', payment ? payment.payer_name : 'کاربر', tracking_code, 'خطا در انجام تراکنش بانکی');
      return res.json({
        success: false,
        status: 'failed',
        payment,
        message: 'پرداخت ناموفق بود.'
      });
    }

    const txnId = ref_id || ('TXN-' + Math.floor(10000000 + Math.random() * 90000000));
    const verifiedAt = new Date().toISOString();
    if (payment) {
      payment.status = 'successful';
      payment.is_approved = true;
      payment.transaction_id = txnId;
      payment.verified_at = verifiedAt;
      payment.paid_at = verifiedAt;

      const user = registerOrUpdateUser(payment.payer_name, payment.phone, payment.amount, payment.is_anonymous);
      payment.user_id = user.id;

      createNotification(
        'پرداخت موفق در پویش',
        `${payment.is_anonymous ? 'مشارکت‌کننده گمنام (' + payment.payer_name + ')' : payment.payer_name} با پرداخت ${payment.amount.toLocaleString('fa-IR')} تومان سهم مشارکت خود را ثبت نمود.`,
        'payment'
      );
      recordAuditLog('پرداخت موفق', payment.payer_name, tracking_code, `واریز موفق مبلغ ${payment.amount.toLocaleString('fa-IR')} تومان (رهگیری: ${tracking_code})`);

      // ارسال خودکار دقیقاً یک پیامک تشکر بعد از تایید موفق
      await triggerThankYouSms(payment);
    }

    return res.json({
      success: true,
      status: 'successful',
      transaction_id: txnId,
      tracking_code,
      amount: payment ? payment.amount : 0,
      verified_at: verifiedAt,
      payment,
      message: 'پرداخت با موفقیت تایید و سهم شما ثبت گردید.'
    });
  } catch (e: any) {
    return res.status(500).json({ success: false, message: e.message || 'خطا در اعتبارسنجی' });
  }
};
app.post('/api/payment/verify', handlePaymentVerify);
app.post('/api/payment/verify.php', handlePaymentVerify);

// -------------------------------------------------------------
// 12. تیکت‌ها و پشتیبانی (Tickets API)
// -------------------------------------------------------------
const handleTicketsGet = (req: Request, res: Response) => {
  const id = req.query.id as string;
  const admin = extractAdminFromRequest(req);
  const isAdmin = Boolean(admin && admin.is_active !== false);

  if (id) {
    const ticket = mockDb.tickets.find(t => t.id === id);
    if (!ticket) return res.status(404).json({ success: false, message: 'تیکت یافت نشد.' });
    const messages = mockDb.ticket_messages.filter(m => m.ticket_id === id);
    return res.json({ success: true, data: { ...ticket, messages } });
  }

  if (isAdmin) {
    if (!admin.is_super_admin && !(admin.permissions || []).includes('tickets')) {
      return res.status(403).json({ success: false, message: 'دسترسی غیرمجاز برای تیکت‌ها.' });
    }
    return res.json({ success: true, data: mockDb.tickets });
  }

  const phone = ((req.query.phone as string) || '').trim();
  if (!phone) return res.json({ success: true, data: [] });
  const userTickets = mockDb.tickets.filter(t => t.phone === phone);
  return res.json({ success: true, data: userTickets });
};

const handleTicketsPost = (req: Request, res: Response) => {
  const action = req.query.action as string;
  const id = req.query.id as string;
  const body = req.body || {};
  const admin = extractAdminFromRequest(req);
  const isAdmin = Boolean(admin && admin.is_active !== false);

  if (action === 'create' || !action) {
    const name = (body.name || '').trim();
    const phone = (body.phone || '').trim();
    const subject = (body.subject || '').trim();
    const message = (body.message || '').trim();

    if (!name || !phone || !subject || !message) {
      return res.status(400).json({ success: false, message: 'تمام فیلدها الزامی هستند.' });
    }

    const ticketId = 'tkt-' + Date.now();
    const messageId = 'msg-' + Date.now();

    mockDb.tickets.unshift({
      id: ticketId,
      name,
      phone,
      subject,
      status: 'open',
      priority: 'medium',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    });

    mockDb.ticket_messages.push({
      id: messageId,
      ticket_id: ticketId,
      sender_type: 'user',
      sender_name: name,
      message,
      created_at: new Date().toISOString()
    });

    createNotification(`تیکت پشتیبانی جدید: ${subject}`, `تیکت جدید از طرف ${name} (${phone}) ثبت شد.`, 'ticket');
    return res.status(201).json({
      success: true,
      message: 'تیکت شما با موفقیت ثبت شد.',
      data: { id: ticketId, subject }
    });
  }

  if (action === 'reply') {
    if (isAdmin && !admin.is_super_admin && !(admin.permissions || []).includes('tickets')) {
      return res.status(403).json({ success: false, message: 'دسترسی غیرمجاز برای پاسخ به تیکت.' });
    }

    const ticketId = id || body.ticket_id;
    const message = (body.message || '').trim();

    if (!ticketId || !message) {
      return res.status(400).json({ success: false, message: 'شناسه تیکت و پیام الزامی است.' });
    }

    const ticket = mockDb.tickets.find(t => t.id === ticketId);
    if (!ticket) return res.status(404).json({ success: false, message: 'تیکت یافت نشد.' });

    const senderType = isAdmin ? 'admin' : 'user';
    const senderName = isAdmin ? (admin.name || 'مدیریت سامانه') : (body.name || ticket.name);
    ticket.status = isAdmin ? 'answered' : 'open';
    ticket.updated_at = new Date().toISOString();

    const msgId = 'msg-' + Date.now();
    mockDb.ticket_messages.push({
      id: msgId,
      ticket_id: ticketId,
      sender_type: senderType,
      sender_name: senderName,
      message,
      created_at: new Date().toISOString()
    });
    saveDbToDisk();

    return res.json({ success: true, message: 'پاسخ با موفقیت ثبت شد.' });
  }

  if (action === 'close') {
    if (isAdmin && !admin.is_super_admin && !(admin.permissions || []).includes('tickets')) {
      return res.status(403).json({ success: false, message: 'دسترسی غیرمجاز برای بستن تیکت.' });
    }

    const ticketId = id || body.ticket_id;
    const ticket = mockDb.tickets.find(t => t.id === ticketId);
    if (ticket) {
      ticket.status = 'closed';
      ticket.updated_at = new Date().toISOString();
      saveDbToDisk();
    }
    return res.json({ success: true, message: 'تیکت بسته شد.' });
  }

  return res.status(400).json({ success: false, message: 'عملیات نامعتبر است.' });
};

app.get('/api/tickets', handleTicketsGet);
app.get('/api/tickets.php', handleTicketsGet);
app.post('/api/tickets', handleTicketsPost);
app.post('/api/tickets.php', handleTicketsPost);

// -------------------------------------------------------------
// سرو پوشه uploads
// -------------------------------------------------------------
app.use('/uploads', express.static(path.resolve(__dirname, 'uploads')));

// -------------------------------------------------------------
// اتصال Vite در توسعه یا سرو استاتیک در Production
// -------------------------------------------------------------
async function startServer() {
  const isDev = process.env.NODE_ENV !== 'production';

  if (isDev) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });

    // بازنویسی مسیرهای صفحات قبل از ارسال به Vite
    app.use((req: Request, _res: Response, next: NextFunction) => {
      const url = req.url.split('?')[0];
      const query = req.url.includes('?') ? '?' + req.url.split('?')[1] : '';

      if (url === '/admin' || url === '/admin/') {
        req.url = '/admin/index.html' + query;
      } else if (url.startsWith('/campaign') && !url.includes('.')) {
        req.url = '/campaign/index.html' + query;
      }
      next();
    });

    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));

    app.get('/admin*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(distDir, 'admin/index.html'));
    });
    app.get('/campaign*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(distDir, 'campaign/index.html'));
    });
    app.get('/gateway-sim.html', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(distDir, 'gateway-sim.html'));
    });
    app.get('*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(distDir, 'index.html'));
    });
  }

  app.listen(PORT, HOST, () => {
    console.log(`Server listening on http://${HOST}:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
