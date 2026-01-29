# PDF Normalizer

A Vue 3 + Vuetify application that processes and normalizes PDFs from Google Drive folders.

## Features

- **Google Drive Integration**: Connect to public Google Drive folders and process all PDFs
- **OCR Support**: Extract text from scanned PDFs (using Tesseract.js)
- **Page Standardization**: Normalize page sizes to Letter, A4, or Legal
- **Compression**: Optimize PDF file sizes while maintaining quality
- **Upload Back to Drive**: Processed files are automatically uploaded back to your Google Drive

## Quick Start

### 1. Clone and Install

```bash
git clone https://github.com/Mohammad-Alavi/pdf-normalizer.git
cd pdf-normalizer
npm install
```

### 2. Set Up Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Drive API**
4. Go to **Credentials** → **Create Credentials** → **OAuth 2.0 Client ID**
5. Configure the OAuth consent screen
6. Create a Web Application credential
7. Add authorized JavaScript origins:
   - `http://localhost:3000` (development)
   - `https://your-vercel-domain.vercel.app` (production)
8. Add authorized redirect URIs (same as above)

### 3. Configure Environment Variables

Create a `.env` file:

```bash
cp .env.example .env
```

Edit `.env` with your credentials:

```
VITE_GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
VITE_GOOGLE_API_KEY=your-api-key
```

### 4. Run Development Server

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000)

## Deployment to Vercel

### Option 1: Using Vercel CLI

```bash
npm i -g vercel
vercel
```

### Option 2: GitHub Integration

1. Push your code to GitHub
2. Import the repository in [Vercel Dashboard](https://vercel.com/new)
3. Add environment variables in Vercel project settings:
   - `VITE_GOOGLE_CLIENT_ID`
   - `VITE_GOOGLE_API_KEY`
4. Deploy!

### Important: Update Google OAuth Settings

After deploying, add your Vercel domain to the authorized origins in Google Cloud Console.

## Usage

1. **Sign in with Google**: Click the "Sign In" button to authenticate
2. **Paste Drive Link**: Enter a public Google Drive folder URL
3. **Configure Options**: Choose processing options (OCR, standardization, compression)
4. **Process**: Click "Process PDFs" to start
5. **Results**: Processed files are uploaded back to the same Drive folder with `normalized_` prefix

## Tech Stack

- **Vue 3** - Progressive JavaScript Framework
- **Vuetify 3** - Material Design Component Framework
- **pdf-lib** - PDF manipulation library
- **Tesseract.js** - OCR engine
- **Vite** - Next Generation Frontend Tooling

## License

MIT
