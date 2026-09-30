# react-richtexteditor-webmail
This repository demonstrates the integration of the Syncfusion EJ2 Rich Text Editor within a webmail application built using React. It serves as a practical example showcasing rich text editing capabilities, email composition workflows, and seamless integration with modern React applications.


# Mail Application - Complete Setup & Implementation Guide

A production-ready webmail application built with **React (Frontend)** and **Fastify + PostgreSQL (Backend)**, integrated with **Microsoft Graph API** for email synchronization and real-time updates.

## 📋 Table of Contents

- [Project Overview](#project-overview)
- [Architecture](#architecture)
- [Prerequisites](#prerequisites)
- [Backend Setup](#backend-setup)
- [Frontend Setup](#frontend-setup)
- [Database Configuration](#database-configuration)
- [Environment Variables](#environment-variables)
- [Running the Application](#running-the-application)
- [Key Features](#key-features)
- [API Endpoints](#api-endpoints)
- [Troubleshooting](#troubleshooting)

---

## 🎯 Project Overview

This mail application provides a seamless email experience by:
- Integrating with **Microsoft Graph API** for email management
- Using **Azure AD/MSAL** for secure authentication
- Implementing **delta queries** for efficient sync (Item 8)
- Using **real-time webhooks** for instant notifications (Item 9)
- Processing jobs with **message queues** for reliability (Item 10)
- Updating the **React frontend in real-time** with WebSocket (Item 11)

### Tech Stack

| Component | Technology |
|-----------|------------|
| **Frontend** | React 19 + TypeScript + Vite + Tailwind CSS |
| **Backend** | Fastify + TypeScript |
| **Database** | PostgreSQL 15 |
| **Authentication** | Azure AD / MSAL |
| **Email API** | Microsoft Graph API |
| **Real-time** | WebSocket / Socket.IO |
| **Queue** | RabbitMQ (optional for advanced features) |
| **ORM** | Prisma |

---

## 🏗️ Architecture

### System Flow

```
User Browser (React)
    ↓
    └─→ [Frontend - Port 5173]
        ├─→ Authentication (MSAL)
        ├─→ REST API Calls
        └─→ WebSocket Connection
        
[Backend - Port 3001] (Fastify)
    ├─→ Authentication Routes (/auth/*)
    ├─→ Mail Routes (/mail/*)
    ├─→ Webhook Routes (/webhook/*)
    ├─→ Attachment Routes (/attachments/*)
    └─→ WebSocket Server (/ws)
    
[Database - Port 5432] (PostgreSQL)
    ├─→ Users Table
    ├─→ CachedEmails Table
    └─→ SyncStates Table
    
[Microsoft Graph API]
    ├─→ Authentication (OAuth 2.0)
    ├─→ Mail API
    ├─→ Delta Queries
    └─→ Webhooks/Subscriptions
```

### Key Features Implementation

#### 1. Delta Queries (Item 8) ⚡
- **Purpose**: Sync only changed emails instead of all emails
- **Benefit**: 98% fewer API calls
- **Files**: `src/services/deltaService.ts`
- **How It Works**:
  - First sync: Download all emails
  - Store delta link in database
  - Next syncs: Only get new/updated/deleted emails

#### 2. Real-Time Webhooks (Item 9) 🔔
- **Purpose**: Instant notifications when emails arrive
- **Files**: `src/services/subscriptionService.ts`, `src/routes/webhook.ts`
- **How It Works**:
  - Subscribe to Microsoft Graph change notifications
  - Graph posts to your webhook endpoint
  - Backend queues the job
  - Frontend receives update via WebSocket

#### 3. Message Queue (Item 10) 📦
- **Purpose**: Reliable background job processing
- **Infrastructure**: RabbitMQ with 3 queues
- **Files**: `src/services/queueService.ts`, `src/services/workerService.ts`
- **How It Works**:
  - Webhook arrives → Queue job
  - Background worker processes
  - Database updated → Frontend notified

#### 4. Frontend Real-Time Updates (Item 11) 💫
- **Purpose**: Instant UI updates without polling
- **Technology**: WebSocket / Socket.IO
- **Files**: `src/services/realtimeClient.ts`, `src/hooks/useRealtime.ts`
- **How It Works**:
  - Connect to WebSocket on app load
  - Backend broadcasts changes
  - React components update instantly

---

## 📦 Prerequisites

Before you begin, ensure you have installed:

- **Node.js**: v18 or higher
  - [Download Node.js](https://nodejs.org/)
  - Verify: `node --version` and `npm --version`

- **PostgreSQL**: v15 or higher
  - **Windows**: [Download PostgreSQL](https://www.postgresql.org/download/windows/)
  - **Mac**: `brew install postgresql`
  - **Linux**: `sudo apt-get install postgresql-15`

- **Docker & Docker Compose** (for PostgreSQL + RabbitMQ)
  - [Download Docker Desktop](https://www.docker.com/products/docker-desktop)
  - Verify: `docker --version` and `docker-compose --version`

- **Git** (for version control)
  - Verify: `git --version`

---

## 🔧 Backend Setup

### Step 1: Navigate to Backend Directory

```bash
cd my-app-backend
```

### Step 2: Install Dependencies

```bash
npm install
```

**Key Dependencies**:
- `fastify` - Web framework
- `@prisma/client` - Database ORM
- `@azure/msal-node` - Azure AD authentication
- `@microsoft/microsoft-graph-client` - Microsoft Graph API
- `@fastify/cors` - CORS support
- `@fastify/jwt` - JWT authentication
- `dotenv` - Environment variables

### Step 3: Set Up PostgreSQL Database

#### Option A: Using Docker Compose (Recommended)

```bash
# In my-app-backend directory
docker-compose up -d

# Verify PostgreSQL is running
docker-compose ps
# Output should show:
# NAME                STATUS           PORTS
# mail-app-postgres   Up (healthy)     0.0.0.0:5432->5432/tcp
```

#### Option B: Using Local PostgreSQL Installation

```bash
# Windows (if installed)
net start postgresql-x64-15

# Or manually in pgAdmin:
# 1. Open pgAdmin
# 2. Right-click "Servers" → Create → Server
# 3. Name: mail-app
# 4. Connection tab:
#    - Host: localhost
#    - Port: 5432
#    - Username: postgres
#    - Password: (your password)
# 5. Click Save
```

### Step 4: Create Database

```bash
# Using psql (command line)
psql -U postgres -h localhost

# In psql console:
CREATE DATABASE maildb;
CREATE USER mailuser WITH PASSWORD 'mailpassword';
ALTER ROLE mailuser SET client_encoding TO 'utf8';
ALTER ROLE mailuser SET default_transaction_isolation TO 'read committed';
ALTER ROLE mailuser SET default_transaction_deferrable TO on;
ALTER ROLE mailuser SET timezone TO 'UTC';
GRANT ALL PRIVILEGES ON DATABASE maildb TO mailuser;
\q
```

**Or using pgAdmin GUI**:
1. Right-click "Databases" → Create → Database
2. Name: `maildb`
3. Owner: (your user)
4. Click Save

### Step 5: Generate Prisma Client

```bash
npm run db:generate
```

### Step 6: Run Database Migrations

```bash
npm run db:migrate

# Follow prompts:
# ? What would you like to do? › Create a new migration
# ? Enter a name for the new migration: › init
```

This creates the following tables:
- `User` - Stores user authentication info
- `CachedEmail` - Stores email data
- `SyncState` - Stores delta links for sync

---

## 🗄️ Database Configuration

### PostgreSQL Connection String

The database is configured via the `DATABASE_URL` environment variable in `.env`:

```env
DATABASE_URL=postgres://username:password@localhost:5432/databasename
```

**Format Breakdown**:
```
postgres://[user]:[password]@[host]:[port]/[database]
          │       │           │       │    │
          user    password    host   port database
```

### Common Connection Strings

```bash
# Docker Compose (default credentials)
DATABASE_URL=postgres://postgres:postgres123@localhost:5432/maildb

# Local PostgreSQL (standard install)
DATABASE_URL=postgres://postgres:password@localhost:5432/mail_app

# Production (Azure/AWS example)
DATABASE_URL=postgres://prod_user:secure_password@db.example.com:5432/maildb_prod
```

### Database Schema

#### User Table
```sql
CREATE TABLE "User" (
  id            TEXT PRIMARY KEY DEFAULT cuid(),
  email         TEXT UNIQUE NOT NULL,
  displayName   TEXT,
  accessToken   TEXT NOT NULL,
  refreshToken  TEXT,
  tokenExpiry   TIMESTAMP,
  createdAt     TIMESTAMP DEFAULT NOW(),
  updatedAt     TIMESTAMP,
  
  INDEX idx_user_email (email)
);
```

#### CachedEmail Table
```sql
CREATE TABLE "CachedEmail" (
  id              TEXT PRIMARY KEY DEFAULT cuid(),
  graphId         TEXT UNIQUE NOT NULL,
  userEmail       TEXT NOT NULL,
  subject         TEXT NOT NULL,
  bodyPreview     TEXT,
  senderEmail     TEXT NOT NULL,
  senderName      TEXT,
  receivedAt      TIMESTAMP NOT NULL,
  isRead          BOOLEAN DEFAULT FALSE,
  hasAttachments  BOOLEAN DEFAULT FALSE,
  rawData         JSON NOT NULL,
  createdAt       TIMESTAMP DEFAULT NOW(),
  updatedAt       TIMESTAMP,
  
  INDEX idx_cached_email_user (userEmail),
  INDEX idx_cached_email_received (receivedAt)
);
```

#### SyncState Table (Item 8)
```sql
CREATE TABLE "SyncState" (
  id              TEXT PRIMARY KEY DEFAULT cuid(),
  userEmail       TEXT NOT NULL UNIQUE,
  deltaLink       TEXT NOT NULL,
  syncToken       TEXT,
  lastSyncAt      TIMESTAMP,
  createdAt       TIMESTAMP DEFAULT NOW(),
  updatedAt       TIMESTAMP,
  
  INDEX idx_sync_state_user (userEmail)
);
```

---

## 🔐 Environment Variables

### Backend Configuration

Create or update `my-app-backend/.env`:

```env
# =====================================
# Azure AD / MSAL Configuration
# =====================================
# Register your app at: https://portal.azure.com
# App registrations → New registration
AZURE_CLIENT_ID= 
AZURE_CLIENT_SECRET=
AZURE_TENANT_ID=
REDIRECT_URI=

# =====================================
# Database Configuration
# =====================================
# PostgreSQL connection string
# Format: postgres://user:password@host:port/database
DATABASE_URL=

# Encryption key for sensitive data (min 16 chars)
ENCRYPTION_SECRET=mysecret123

# =====================================
# JWT Configuration
# =====================================
# Secret key for signing JWT tokens (use strong random string in production)
JWT_SECRET=your-super-secret-key-min-32-characters-long

# =====================================
# Server Configuration
# =====================================
# Server port
PORT=3001

# Environment (development, staging, production)
NODE_ENV=development

# =====================================
# Frontend Configuration
# =====================================
# Frontend URL for CORS
FRONTEND_URL=http://localhost:5173

# =====================================
# Webhook Configuration
# =====================================
# Your public webhook URL for Microsoft Graph subscriptions
# Use ngrok for local development: https://ngrok.com/
# Example: https://gangrene-distill-suds.ngrok-free.dev/webhook
WEBHOOK_URL=

# =====================================
# RabbitMQ Configuration (Optional)
# =====================================
# RabbitMQ connection URL (for message queue)
RABBITMQ_URL=amqp://guest:guest@localhost:5672

# =====================================
# WebSocket Configuration
# =====================================
# WebSocket URL for frontend
VITE_API_WS_URL=ws://localhost:3001/ws
```

### Generating Secure Secrets

**Generate JWT_SECRET**:
```bash
# Using OpenSSL (Windows/Mac/Linux)
openssl rand -base64 32

# Using Node.js
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"

# Example output:
# a7f9c3e2b1d0f8a4c6e5b2d9f0a1c3e4f6b8d0e2a4c6e8f0a2b4d6e8f0a1c
```

**Generate ENCRYPTION_SECRET**:
```bash
node -e "console.log(require('crypto').randomBytes(16).toString('hex'))"

# Example output:
# f3d5b9e1c7a2f8d0b4e6c1a9f5d7b3e1
```

### Azure AD Configuration

#### 1. Register Your Application

```
Go to: https://portal.azure.com
Login with your Microsoft account
Navigate to: Azure Active Directory → App registrations
Click: "+ New registration"
```

#### 2. Set Application Details

```
Name: Mail Application
Supported account types: Accounts in any organizational directory (Multi-tenant)
Redirect URI:
  - Platform: Web
  - URI: http://localhost:3001/auth/callback
Click: "Register"
```

#### 3. Create Client Secret

```
In app page, go to: Certificates & secrets
Click: "+ New client secret"
Description: Backend Application Secret
Expires: 24 months
Click: "Add"
Copy the secret VALUE (not ID)
→ Set as AZURE_CLIENT_SECRET in .env
```

#### 4. Configure API Permissions

```
In app page, go to: API permissions
Click: "+ Add a permission"
Select: Microsoft Graph
Select: Delegated permissions
Search and select:
  ☐ Mail.Read
  ☐ Mail.ReadWrite
  ☐ MailboxSettings.Read
  ☐ Calendar.Read
  ☐ Calendar.ReadWrite
  ☐ offline_access
Click: "Add permissions"
Click: "Grant admin consent for [Organization]"
```

#### 5. Get Configuration Values

```
In app page, overview tab:
- Application (client) ID → AZURE_CLIENT_ID
- Directory (tenant) ID → AZURE_TENANT_ID
```

### Microsoft Graph Configuration

#### Webhook Setup (for Item 9)

1. **Set Up ngrok for Local Testing**:

```bash
# Download ngrok: https://ngrok.com/download

# Run ngrok to expose local server
ngrok http 3001

# Output:
# Forwarding     https://abc123.ngrok-free.dev → http://localhost:3001

# Set in .env:
WEBHOOK_URL=https://abc123.ngrok-free.dev/webhook
```

2. **Configure Webhook Endpoint**:

The backend automatically handles Microsoft Graph subscriptions. The webhook is registered at:
```
POST /webhook
```

3. **Create Subscription**:

Use the API to create a subscription:
```bash
curl -X POST http://localhost:3001/mail/subscribe \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "changeType": "created",
    "resource": "/me/mailFolders/inbox/messages",
    "notificationUrl": "https://your-ngrok-url.ngrok-free.dev/webhook",
    "expirationDateTime": "2026-10-15T00:00:00Z",
    "clientState": "secret-string"
  }'
```

---

## 🚀 Running the Application

### Step 1: Start Backend

```bash
cd my-app-backend

# Install dependencies (if not done)
npm install

# Start development server
npm run dev

# Output should show:
# 🚀 Server running at http://0.0.0.0:3001
# 📨 WebSocket server ready at ws://localhost:3001/ws
```

### Step 2: Start Frontend

In a new terminal:

```bash
cd my-app

# Install dependencies (if not done)
npm install

# Start development server
npm run dev

# Output should show:
# VITE v8.0.12  ready in 312 ms
# ➜  Local:   http://localhost:5173/
# ➜  press h to show help
```

### Step 3: Open Application

Open your browser:
```
http://localhost:5173
```

You should see:
1. Login button
2. Click to authenticate with Microsoft
3. Redirected to Azure AD login
4. After success, emails load automatically
5. Real-time updates as emails arrive

---

## 🎯 Key Features

### 1. Authentication Flow

```
User clicks "Login"
    ↓
MSAL redirects to Azure AD
    ↓
User enters credentials
    ↓
Azure AD redirects back with auth code
    ↓
Backend exchanges code for tokens
    ↓
JWT token sent to frontend
    ↓
Frontend stores JWT in localStorage
    ↓
All API requests include JWT
```

### 2. Email Synchronization

#### First Sync (Full)
```bash
GET /me/mailFolders/inbox/messages
  ← Returns: 1000 emails
Store delta link in database
```

#### Subsequent Syncs (Delta)
```bash
GET /me/mailFolders/inbox/messages/delta?$deltatoken=TOKEN
  ← Returns: Only new/updated/deleted emails (typically 0-50)
Update delta link
```

**Benefits**:
- 95%+ fewer API calls
- Support for 100k+ emails
- Faster sync times

### 3. Real-Time Updates

When a new email arrives:

```
1. Microsoft Graph detects change
2. Posts to your webhook: POST /webhook
3. Backend queues change notification
4. Background worker processes
5. Database updated
6. WebSocket broadcast to all connected clients
7. React components auto-update (real-time!)
```

### 4. Attachment Handling

Files are uploaded directly to Microsoft Graph:

```
Frontend
  ↓ (file selected)
Upload directly to Microsoft Graph
  ↓ (backend proxies chunk info)
Graph returns attachment ID
  ↓
Message sent with attachment
```

**Benefits**:
- Backend not bottleneck for large files
- Supports files up to 25MB
- Chunked uploads for resume capability

---

## 📡 API Endpoints

### Authentication Routes

```bash
# Start login flow
GET /auth/login
Response: { authCodeUrl: "https://..." }

# Callback (automatic redirect)
GET /auth/callback?code=...&state=...

# Logout
POST /auth/logout
Headers: Authorization: Bearer JWT_TOKEN

# Get current user
GET /auth/me
Headers: Authorization: Bearer JWT_TOKEN
Response: { id, email, displayName, profile }
```

### Mail Routes

```bash
# Get inbox messages
GET /mail/inbox?skip=0&take=50
Headers: Authorization: Bearer JWT_TOKEN
Response: { value: [...messages], '@odata.nextLink': '...' }

# Get message details with attachments
GET /mail/messages/:messageId
Headers: Authorization: Bearer JWT_TOKEN
Response: { id, subject, body, attachments: [...] }

# Mark as read
PATCH /mail/messages/:messageId
Headers: Authorization: Bearer JWT_TOKEN
Body: { isRead: true }

# Delete message
DELETE /mail/messages/:messageId
Headers: Authorization: Bearer JWT_TOKEN
```

### Webhook Routes

```bash
# Webhook validation (Microsoft Graph)
POST /webhook
Body: { value: [{ changeType, resourceData, ... }] }
Response: 202 Accepted (queued)

# Create subscription (Item 9)
POST /mail/subscribe
Headers: Authorization: Bearer JWT_TOKEN
Body: { changeType, resource, expirationDateTime }

# List subscriptions
GET /mail/subscriptions
Headers: Authorization: Bearer JWT_TOKEN
Response: { value: [...subscriptions] }
```

### Attachment Routes

```bash
# Upload small file
POST /attachments/upload
Headers: 
  Authorization: Bearer JWT_TOKEN
  Content-Type: multipart/form-data
Body: FormData { file: File }
Response: { attachmentId, fileName, size }

# Upload large file (chunked)
POST /attachments/upload-chunk
Headers: Authorization: Bearer JWT_TOKEN
Body: { chunkId, chunkIndex, totalChunks, data: base64 }

# Download attachment
GET /attachments/:attachmentId/download
Headers: Authorization: Bearer JWT_TOKEN
Response: File binary data
```

---

## 🐛 Troubleshooting

### Backend Issues

#### 1. Port Already in Use

```
Error: listen EADDRINUSE: address already in use :::3001

Solution:
# Find process using port 3001
lsof -i :3001
# Kill the process
kill -9 PID

# Or use different port
PORT=3002 npm run dev
```

#### 2. Database Connection Failed

```
Error: connect ECONNREFUSED 127.0.0.1:5432

Solutions:
1. Check PostgreSQL is running:
   docker-compose ps
   
2. Verify DATABASE_URL in .env:
   DATABASE_URL=postgres://user:password@localhost:5432/dbname
   
3. Test connection:
   psql -U postgres -h localhost -d maildb
   
4. Check PostgreSQL service:
   # Windows:
   net start postgresql-x64-15
   
   # Mac:
   brew services start postgresql
   
   # Linux:
   sudo systemctl start postgresql
```

#### 3. Missing Environment Variables

```
Error: ⚠️ Missing environment variable: AZURE_CLIENT_ID

Solution:
1. Copy template: cp .env.example .env
2. Fill all values from Azure Portal
3. Verify no typos or empty values
4. Restart backend: npm run dev
```

#### 4. Authentication Fails

```
Error: OAuth 2.0 token exchange failed

Solutions:
1. Verify AZURE_CLIENT_SECRET is correct (copy from Portal)
2. Check AZURE_TENANT_ID matches your Azure tenant
3. Ensure REDIRECT_URI matches app registration:
   http://localhost:3001/auth/callback
4. Check app has Mail.Read permission granted
5. Try clearing browser cookies and login again
```

### Frontend Issues

#### 1. Cannot Connect to Backend

```
Error: Failed to fetch from http://localhost:3001

Solutions:
1. Verify backend is running: npm run dev
2. Check VITE_API_URL in .env or .env.local
3. Verify CORS is enabled in backend
4. Check firewall allows port 3001
```

#### 2. Login Redirects Loop

```
Issue: Infinite redirect loop after login

Solutions:
1. Clear browser localStorage:
   localStorage.clear()
   
2. Clear cookies and cache
3. Ensure token is stored correctly:
   localStorage.getItem('authToken')
   
4. Check AuthProvider in App.tsx handles redirect
```

#### 3. Emails Not Loading

```
Issue: Inbox shows empty after login

Solutions:
1. Check network tab for API errors
2. Verify backend logs for errors:
   npm run dev (look for console output)
   
3. Check user has Mail.Read permission in Azure
4. Try manual API call:
   curl -H "Authorization: Bearer TOKEN" \
        http://localhost:3001/mail/inbox
        
5. Check database has CachedEmail records:
   psql -U postgres -d maildb
   SELECT * FROM "CachedEmail" LIMIT 5;
```

### Database Issues

#### 1. Migrations Failed

```
Error: Migration failed

Solution:
# Reset migrations (⚠️ clears all data)
npm run db:migrate:reset

# Or manually run migrations:
npm run db:migrate

# Check migration status:
npm run db:status
```

#### 2. Prisma Client Out of Sync

```
Error: Cannot find @prisma/client

Solution:
npm run db:generate
npm install
```

---

## 📝 Development Workflow

### Adding a New Email Field

1. **Update Prisma Schema**:
```prisma
model CachedEmail {
  // ... existing fields
  category String? // NEW
}
```

2. **Create Migration**:
```bash
npm run db:migrate
# Enter: add_category_field
```

3. **Update Backend API**:
```typescript
export async function getInbox(graphClient: any) {
  return graphClient
    .api("/me/mailFolders/inbox/messages")
    .select([
      // ... existing fields
      "categories", // NEW
    ])
    .get();
}
```

4. **Update Frontend Types**:
```typescript
interface Email {
  // ... existing fields
  category?: string; // NEW
}
```

### Testing Changes

```bash
# Backend
npm run dev # Terminal 1

# Frontend
npm run dev # Terminal 2

# Test in browser
http://localhost:5173

# Build for production
npm run build
npm run preview
```

---

## 🚢 Production Deployment

### Backend Deployment (Example: Azure App Service)

```bash
# 1. Build
npm run build

# 2. Create .env for production
DATABASE_URL=<production-db-url>
AZURE_CLIENT_ID=<production-client-id>
# ... other production env vars

# 3. Deploy
az webapp up --name mail-app-backend --resource-group mygroup

# 4. Monitor
az webapp log tail --name mail-app-backend --resource-group mygroup
```

### Frontend Deployment (Example: Azure Static Web Apps)

```bash
# 1. Build
npm run build

# 2. Output in dist/ folder
# 3. Deploy
az staticwebapp up --app-name mail-app-frontend --output-location dist
```

---

## 📚 Additional Resources

- [Microsoft Graph API Docs](https://docs.microsoft.com/en-us/graph/api/)
- [Azure AD Authentication](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [Fastify Documentation](https://www.fastify.io/docs/latest/)
- [React Documentation](https://react.dev/)
- [PostgreSQL Documentation](https://www.postgresql.org/docs/)
- [Prisma ORM Docs](https://www.prisma.io/docs/)

---

## ✅ Checklist Before Deployment

- [ ] All environment variables configured
- [ ] Database backed up
- [ ] SSL certificates installed (HTTPS)
- [ ] CORS configured for production domain
- [ ] Webhook URL updated to production
- [ ] JWT secret is strong (32+ characters)
- [ ] Encryption secret is strong (16+ characters)
- [ ] Rate limiting configured
- [ ] Logging and monitoring enabled
- [ ] Tests passing
- [ ] No console errors in browser
- [ ] Backend logs show no errors

---

**Last Updated**: September 2026
**Version**: 1.0.0
