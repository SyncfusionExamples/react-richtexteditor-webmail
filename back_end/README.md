# Mail App Backend - Phase 1 Setup Guide

## Overview
This backend provides authentication via MSAL, Microsoft Graph integration, and database persistence using Prisma and PostgreSQL.

## Quick Start

### 1. Install Dependencies
```bash
cd my-app-backend
npm install
```

### 2. Set Up Environment Variables
```bash
cp .env.example .env
```

Edit `.env` with your credentials:
- `AZURE_CLIENT_ID` - from Azure Portal
- `AZURE_CLIENT_SECRET` - from Azure Portal  
- `AZURE_TENANT_ID` - from Azure Portal
- `DATABASE_URL` - PostgreSQL connection string
- `JWT_SECRET` - any secure random string

### 3. Azure AD Setup (Required)

#### Create Azure App Registration:
1. Go to [Azure Portal](https://portal.azure.com)
2. Search for "App registrations"
3. Click "New registration"
4. Enter name: "Mail App Backend"
5. Select "Accounts in this organizational directory only"
6. Click Register

#### Get Credentials:
1. In the app, go to "Overview"
2. Copy **Application (client) ID** → `AZURE_CLIENT_ID`
3. Go to "Certificates & secrets"
4. Click "New client secret"
5. Copy the **Value** → `AZURE_CLIENT_SECRET`
6. Copy **Directory (tenant) ID** from Overview → `AZURE_TENANT_ID`

#### Add API Permissions:
1. Go to "API permissions"
2. Click "Add a permission"
3. Select "Microsoft Graph"
4. Select "Delegated permissions"
5. Search and add:
   - `Mail.Read`
   - `Mail.Send`
6. Click "Grant admin consent"

### 4. Set Up Database

#### Option A: Docker (Recommended)
```bash
docker run --name postgres-mail \
  -e POSTGRES_DB=mail_app \
  -e POSTGRES_PASSWORD=password \
  -p 5432:5432 \
  -d postgres:15
```

Update `.env`:
```
DATABASE_URL="postgresql://postgres:password@localhost:5432/mail_app"
```

#### Option B: Local PostgreSQL
Create a database:
```sql
CREATE DATABASE mail_app;
```

Update `.env` with your PostgreSQL connection string.

### 5. Initialize Database
```bash
npm run db:generate
npm run db:migrate
```

### 6. Run Development Server
```bash
npm run dev
```

Server will start at: `http://localhost:3001`

## API Endpoints

### Authentication
- `GET /auth/login` - Get authorization URL
- `GET /auth/callback?code=...` - Handle OAuth callback
- `POST /auth/refresh` - Refresh access token

### Mail
- `GET /mail/messages?skip=0&take=50` - List inbox
- `GET /mail/messages/:id` - Get specific email
- `POST /mail/send` - Send email
- `GET /mail/folders` - List all folders

## Next Steps
- ✅ Phase 1: Auth + basic endpoints (current)
- 📋 Phase 2: Database caching + background sync
- ⏳ Phase 3: Real-time updates + advanced features
