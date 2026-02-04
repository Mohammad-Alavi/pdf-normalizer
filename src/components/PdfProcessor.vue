<template>
  <v-card class="mx-auto" max-width="900">
    <v-card-title class="text-h5 pa-6 pb-2 d-flex align-center">
      <v-icon color="primary" class="mr-2">mdi-file-document-edit</v-icon>
      PDF Invoice Processor
      <v-spacer></v-spacer>
      <v-btn
        icon
        variant="text"
        color="info"
        @click="showHelpDialog = true"
        title="راهنمای استفاده / How to Use"
      >
        <v-icon>mdi-help-circle</v-icon>
      </v-btn>
    </v-card-title>

    <!-- Help Dialog -->
    <v-dialog v-model="showHelpDialog" max-width="700">
      <v-card>
        <v-card-title class="d-flex align-center">
          <v-icon class="mr-2" color="info">mdi-help-circle</v-icon>
          راهنمای استفاده / How to Use
          <v-spacer></v-spacer>
          <v-btn icon variant="text" @click="showHelpDialog = false">
            <v-icon>mdi-close</v-icon>
          </v-btn>
        </v-card-title>
        <v-card-text>
          <div class="text-body-2" style="direction: rtl; text-align: right;">
            <h4 class="mb-2">۱. آماده‌سازی پوشه Google Drive:</h4>
            <ul class="mb-3">
              <li>یک پوشه در Google Drive ایجاد کنید</li>
              <li>فایل‌های PDF فاکتور را در این پوشه قرار دهید</li>
              <li>نام فایل‌های PDF باید به این فرمت باشد: <code style="direction: ltr;">شماره فاکتور - نام خریدار.pdf</code></li>
              <li>مثال: <code style="direction: ltr;">3427 - آموزشگاه زبان کوشش.pdf</code></li>
            </ul>

            <h4 class="mb-2">۲. آماده‌سازی فایل اکسل (اختیاری):</h4>
            <ul class="mb-3">
              <li>فایل اکسل الگو (<code>فرم جدید اطلاعات قراردادها و فاکتورها.xlsx</code>) را در همان پوشه قرار دهید</li>
              <li>این فایل باید شیت <code>فرم جدید اطلاعات فاکتورها</code> را داشته باشد</li>
              <li>اگر فایل اکسل موجود نباشد، فقط Master Sheet ایجاد می‌شود</li>
            </ul>

            <h4 class="mb-2">۳. اجرای برنامه:</h4>
            <ul class="mb-3">
              <li>لینک پوشه Google Drive را در کادر بالا وارد کنید</li>
              <li>با حساب Google خود وارد شوید</li>
              <li>تنظیمات را بررسی کنید و دکمه "Process PDFs" را بزنید</li>
            </ul>

            <h4 class="mb-2">۴. خروجی‌ها:</h4>
            <ul class="mb-3">
              <li>پوشه <code>Processed</code> ایجاد می‌شود با ساختار <code>شناسه/سال/Mostanadat</code></li>
              <li>فایل‌های PDF پردازش‌شده در پوشه Mostanadat قرار می‌گیرند</li>
              <li>Master Sheet (Invoice Data) در پوشه Processed ایجاد می‌شود</li>
              <li>کپی فایل اکسل (به فرمت Google Sheets) با داده‌های پر شده در پوشه Processed قرار می‌گیرد</li>
            </ul>

            <v-divider class="my-3"></v-divider>

            <h4 class="mb-2" style="direction: ltr; text-align: left;">English Instructions:</h4>
            <div style="direction: ltr; text-align: left;">
              <p><strong>1. Prepare Google Drive folder:</strong> Create a folder and add PDF invoices with filename format: <code>InvoiceNumber - BuyerName.pdf</code></p>
              <p><strong>2. Optional Excel template:</strong> Place the Excel template file in the same folder</p>
              <p><strong>3. Run:</strong> Paste folder link, sign in with Google, click "Process PDFs"</p>
              <p><strong>4. Output:</strong> Processed folder with PDFs in Mostanadat, master sheet, and populated Excel (as Google Sheet)</p>
            </div>
          </div>
        </v-card-text>
      </v-card>
    </v-dialog>

    <v-card-subtitle class="px-6">
      Enter a public Google Drive folder link to process all PDFs
    </v-card-subtitle>

    <v-card-text class="pa-6">
      <!-- Drive Link Input -->
      <v-text-field
        v-model="driveLink"
        label="Google Drive Folder Link"
        placeholder="https://drive.google.com/drive/folders/..."
        variant="outlined"
        prepend-inner-icon="mdi-link"
        :error-messages="linkError"
        :disabled="isProcessing"
        clearable
        @update:model-value="validateLink"
      />

      <!-- Processing Options -->
      <v-expansion-panels class="mb-4">
        <v-expansion-panel>
          <v-expansion-panel-title>
            <v-icon class="mr-2">mdi-cog</v-icon>
            Processing Options
          </v-expansion-panel-title>
          <v-expansion-panel-text>
            <v-row>
              <v-col cols="12" md="6">
                <v-checkbox
                  v-model="options.ocr"
                  label="Extract Text"
                  hint="Extract text from PDFs (uses OCR for scanned docs). Saves as TXT and Excel."
                  persistent-hint
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="6">
                <v-checkbox
                  v-model="options.standardize"
                  label="Standardize Format"
                  hint="Normalize page sizes and orientation"
                  persistent-hint
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="6">
                <v-checkbox
                  v-model="options.compress"
                  label="Compress & Optimize"
                  hint="Reduce file size while maintaining quality"
                  persistent-hint
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="6">
                <v-select
                  v-model="options.pageSize"
                  :items="pageSizes"
                  label="Target Page Size"
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing || !options.standardize"
                />
              </v-col>
            </v-row>
          </v-expansion-panel-text>
        </v-expansion-panel>
      </v-expansion-panels>

      <!-- Excel Mapping Options -->
      <v-expansion-panels class="mb-4">
        <v-expansion-panel>
          <v-expansion-panel-title>
            <v-icon class="mr-2">mdi-microsoft-excel</v-icon>
            Excel Mapping Settings / تنظیمات اکسل
          </v-expansion-panel-title>
          <v-expansion-panel-text>
            <v-row>
              <v-col cols="12" md="6">
                <v-text-field
                  v-model="excelSettings.filename"
                  label="Excel File Name / نام فایل اکسل"
                  hint="Name of Excel file in root folder (without .xlsx)"
                  persistent-hint
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="6">
                <v-text-field
                  v-model="excelSettings.sheetName"
                  label="Sheet Name / نام شیت"
                  hint="Target sheet name in Excel file"
                  persistent-hint
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="4">
                <v-text-field
                  v-model="excelSettings.year"
                  label="سال / Year"
                  hint="Default year for column E"
                  persistent-hint
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="4">
                <v-text-field
                  v-model="excelSettings.companyId"
                  label="شناسه / Company ID"
                  hint="Value for column C"
                  persistent-hint
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing"
                />
              </v-col>
              <v-col cols="12" md="4">
                <v-text-field
                  v-model="excelSettings.companyName"
                  label="شرکت / Company Name"
                  hint="Value for column D"
                  persistent-hint
                  variant="outlined"
                  density="compact"
                  :disabled="isProcessing"
                />
              </v-col>
            </v-row>
          </v-expansion-panel-text>
        </v-expansion-panel>
      </v-expansion-panels>

      <!-- Google OAuth -->
      <v-alert
        v-if="!isAuthenticated"
        type="info"
        variant="tonal"
        class="mb-4"
      >
        <div class="d-flex align-center">
          <div class="flex-grow-1">
            <strong>Sign in with Google</strong> to upload processed files back to your Drive
          </div>
          <v-btn
            color="primary"
            variant="flat"
            @click="authenticateGoogle"
            :loading="isAuthenticating"
          >
            <v-icon class="mr-2">mdi-google</v-icon>
            Sign In
          </v-btn>
        </div>
      </v-alert>

      <v-alert
        v-else
        type="success"
        variant="tonal"
        class="mb-4"
        icon="mdi-check-circle"
      >
        <div class="d-flex align-center">
          <div class="flex-grow-1">
            Signed in as <strong>{{ userEmail || 'Google User' }}</strong>
          </div>
          <v-btn
            variant="text"
            size="small"
            @click="signOut"
          >
            Sign Out
          </v-btn>
        </div>
      </v-alert>

      <!-- Action Button -->
      <div class="d-flex mb-4">
        <v-btn
          color="primary"
          size="large"
          :loading="isProcessing"
          :disabled="!canProcess"
          @click="processFiles"
        >
          <v-icon class="mr-2">mdi-play</v-icon>
          Process PDFs
        </v-btn>
      </div>

      <!-- Progress Section -->
      <v-card v-if="isProcessing || files.length > 0" variant="outlined" class="mb-4">
        <v-card-text>
          <div class="d-flex align-center mb-2">
            <v-icon class="mr-2" :color="isProcessing ? 'primary' : 'success'">
              {{ isProcessing ? 'mdi-loading mdi-spin' : 'mdi-check-circle' }}
            </v-icon>
            <span class="font-weight-medium">
              {{ progressText }}
            </span>
          </div>

          <v-progress-linear
            v-if="isProcessing"
            :model-value="progress"
            color="primary"
            height="8"
            rounded
            class="mb-4"
          />

          <!-- File List -->
          <v-list density="compact" class="bg-transparent">
            <v-list-item
              v-for="file in files"
              :key="file.id"
              :title="file.name"
            >
              <template v-slot:prepend>
                <v-icon :color="getFileStatusColor(file.status)">
                  {{ getFileStatusIcon(file.status) }}
                </v-icon>
              </template>
              <template v-slot:subtitle>
                <span :class="getFileStatusColor(file.status) + '--text'">
                  {{ file.statusText }}
                </span>
              </template>
              <template v-slot:append>
                <v-chip
                  v-if="file.originalSize && file.processedSize"
                  size="x-small"
                  :color="getSizeChange(file.originalSize, file.processedSize).color"
                  variant="flat"
                >
                  {{ getSizeChange(file.originalSize, file.processedSize).text }}
                </v-chip>
              </template>
            </v-list-item>
          </v-list>
        </v-card-text>
      </v-card>

      <!-- Link to Processed folder -->
      <v-alert
        v-if="processedFolderLink"
        type="info"
        variant="tonal"
        class="mb-4"
      >
        <div class="d-flex align-center">
          <v-icon class="mr-2">mdi-folder-open</v-icon>
          <span>فایل‌های پردازش‌شده: </span>
          <a :href="processedFolderLink" target="_blank" class="ml-1">
            باز کردن پوشه Processed در Google Drive
          </a>
        </div>
      </v-alert>

      <!-- Results -->
      <v-alert
        v-if="error"
        type="error"
        variant="tonal"
        class="mb-4"
        closable
        @click:close="error = ''"
      >
        {{ error }}
      </v-alert>

      <v-alert
        v-if="successMessage"
        type="success"
        variant="tonal"
        class="mb-4"
        closable
        @click:close="successMessage = ''"
      >
        {{ successMessage }}
      </v-alert>

      <!-- Processing Logs -->
      <v-expansion-panels v-if="processingLogs.length > 0" class="mb-4">
        <v-expansion-panel>
          <v-expansion-panel-title>
            <v-icon class="mr-2">mdi-text-box-outline</v-icon>
            Processing Logs ({{ processingLogs.length }} entries)
          </v-expansion-panel-title>
          <v-expansion-panel-text>
            <div class="logs-container" style="max-height: 300px; overflow-y: auto; font-family: monospace; font-size: 12px;">
              <div
                v-for="log in processingLogs"
                :key="log.id"
                class="log-entry py-1"
                :class="{
                  'text-success': log.type === 'success',
                  'text-error': log.type === 'error',
                  'text-warning': log.type === 'warning',
                  'text-info': log.type === 'info'
                }"
              >
                <span class="text-grey-darken-1">[{{ log.timestamp }}]</span>
                <v-icon
                  size="x-small"
                  class="mx-1"
                  :color="log.type === 'success' ? 'success' : log.type === 'error' ? 'error' : log.type === 'warning' ? 'warning' : 'info'"
                >
                  {{ log.type === 'success' ? 'mdi-check-circle' : log.type === 'error' ? 'mdi-alert-circle' : log.type === 'warning' ? 'mdi-alert' : 'mdi-information' }}
                </v-icon>
                <span>{{ log.message }}</span>
                <span v-if="log.details" class="text-grey-darken-1 ml-2">— {{ log.details }}</span>
              </div>
            </div>
            <div class="d-flex mt-2" style="gap: 8px;">
              <v-btn
                variant="text"
                size="small"
                color="primary"
                @click="copyLogs"
              >
                <v-icon size="small" class="mr-1">mdi-content-copy</v-icon>
                Copy Logs
              </v-btn>
              <v-btn
                variant="text"
                size="small"
                color="error"
                @click="clearLogs"
              >
                <v-icon size="small" class="mr-1">mdi-delete</v-icon>
                Clear Logs
              </v-btn>
            </div>
          </v-expansion-panel-text>
        </v-expansion-panel>
      </v-expansion-panels>
    </v-card-text>
  </v-card>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { PDFDocument } from 'pdf-lib'
import Tesseract from 'tesseract.js'
import * as pdfjsLib from 'pdfjs-dist'

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`

// State
const driveLink = ref('')
const linkError = ref('')
const isProcessing = ref(false)
const isAuthenticating = ref(false)
const isAuthenticated = ref(false)
const userEmail = ref('')
const accessToken = ref('')
const progress = ref(0)
const files = ref([])
const error = ref('')
const successMessage = ref('')
const processingLogs = ref([])
const processedFolderLink = ref('')
const showHelpDialog = ref(false)

// Options
const options = ref({
  ocr: true,
  standardize: true,
  compress: true,
  pageSize: 'ORIGINAL'
})

// Excel Mapping Settings
const excelSettings = ref({
  filename: 'فرم جدید اطلاعات قراردادها و فاکتورها',
  sheetName: 'فرم جدید اطلاعات فاکتورها',
  year: '1403',
  companyId: '10260538003',
  companyName: 'شرکت راهبرد رایانه رستاک'
})

const pageSizes = [
  { title: 'Letter (8.5" x 11")', value: 'LETTER' },
  { title: 'A4 (210mm x 297mm)', value: 'A4' },
  { title: 'Legal (8.5" x 14")', value: 'LEGAL' },
  { title: 'Keep Original', value: 'ORIGINAL' }
]

// Google OAuth Config
const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || ''
const SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile'

// Computed
const canProcess = computed(() => {
  return driveLink.value && !linkError.value && extractFolderId(driveLink.value) && isAuthenticated.value
})

const progressText = computed(() => {
  if (!isProcessing.value && files.value.length === 0) return ''
  if (isProcessing.value) {
    const completed = files.value.filter(f => f.status === 'done' || f.status === 'error').length
    return `Processing ${completed} of ${files.value.length} files...`
  }
  const successful = files.value.filter(f => f.status === 'done').length
  return `Completed: ${successful} of ${files.value.length} files processed successfully`
})

// Helper functions
function extractFolderId(url) {
  if (!url) return null
  const patterns = [/\/folders\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/, /^([a-zA-Z0-9_-]{25,})/]
  for (const pattern of patterns) {
    const match = url.match(pattern)
    if (match) return match[1]
  }
  return null
}

function validateLink() {
  if (!driveLink.value) {
    linkError.value = ''
    return
  }
  linkError.value = extractFolderId(driveLink.value) ? '' : 'Invalid Google Drive folder link'
}

function getFileStatusIcon(status) {
  const icons = {
    pending: 'mdi-clock-outline',
    downloading: 'mdi-download',
    processing: 'mdi-cog mdi-spin',
    uploading: 'mdi-upload',
    done: 'mdi-check-circle',
    error: 'mdi-alert-circle'
  }
  return icons[status] || 'mdi-file'
}

function getFileStatusColor(status) {
  const colors = {
    pending: 'grey',
    downloading: 'info',
    processing: 'warning',
    uploading: 'info',
    done: 'success',
    error: 'error'
  }
  return colors[status] || 'grey'
}

function getSizeChange(original, processed) {
  const change = ((original - processed) / original * 100).toFixed(0)
  if (change > 0) return { text: `Saved ${change}%`, color: 'success' }
  if (change < 0) return { text: `+${Math.abs(change)}%`, color: 'warning' }
  return { text: 'No change', color: 'grey' }
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

function addLog(type, message, details = null) {
  processingLogs.value.push({
    id: Date.now() + Math.random(),
    timestamp: new Date().toLocaleTimeString(),
    type,
    message,
    details
  })
}

function clearLogs() {
  processingLogs.value = []
}

async function copyLogs() {
  const logText = processingLogs.value.map(log => {
    const prefix = { success: '✓', error: '✗', warning: '⚠', info: 'ℹ' }[log.type] || 'ℹ'
    return `[${log.timestamp}] ${prefix} ${log.message}${log.details ? ` — ${log.details}` : ''}`
  }).join('\n')

  try {
    await navigator.clipboard.writeText(logText)
    addLog('success', 'Logs copied to clipboard!')
  } catch (err) {
    addLog('error', 'Failed to copy logs to clipboard')
  }
}

// Google OAuth
async function authenticateGoogle() {
  isAuthenticating.value = true
  error.value = ''

  try {
    if (!window.google?.accounts?.oauth2) {
      await loadGoogleScript()
    }

    const tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: async (response) => {
        if (response.error) {
          error.value = 'Authentication failed: ' + response.error
          isAuthenticating.value = false
          return
        }

        accessToken.value = response.access_token
        localStorage.setItem('google_access_token', response.access_token)
        localStorage.setItem('google_token_expiry', String(Date.now() + (response.expires_in * 1000)))
        isAuthenticated.value = true

        try {
          await fetchUserInfo()
          if (userEmail.value) localStorage.setItem('google_user_email', userEmail.value)
        } catch (err) {
          console.error('Failed to fetch user info:', err)
        }

        isAuthenticating.value = false
      }
    })

    tokenClient.requestAccessToken({ prompt: 'consent' })
  } catch (err) {
    error.value = 'Failed to initialize Google authentication: ' + err.message
    isAuthenticating.value = false
  }
}

function loadGoogleScript() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts?.oauth2) {
      resolve()
      return
    }
    const script = document.createElement('script')
    script.src = 'https://accounts.google.com/gsi/client'
    script.async = true
    script.defer = true
    script.onload = resolve
    script.onerror = () => reject(new Error('Failed to load Google script'))
    document.head.appendChild(script)
  })
}

async function fetchUserInfo() {
  const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: `Bearer ${accessToken.value}` }
  })
  if (response.ok) {
    const data = await response.json()
    userEmail.value = data.email || ''
  }
}

function signOut() {
  const token = accessToken.value
  clearAuthData()
  if (window.google?.accounts?.oauth2 && token) {
    google.accounts.oauth2.revoke(token)
  }
}

function clearAuthData() {
  accessToken.value = ''
  isAuthenticated.value = false
  userEmail.value = ''
  localStorage.removeItem('google_access_token')
  localStorage.removeItem('google_token_expiry')
  localStorage.removeItem('google_user_email')
}

// Main processing function
async function processFiles() {
  if (!canProcess.value) return

  isProcessing.value = true
  error.value = ''
  successMessage.value = ''
  files.value = []
  progress.value = 0
  processedFolderLink.value = ''
  clearLogs()

  try {
    const folderId = extractFolderId(driveLink.value)
    addLog('info', 'Starting PDF processing', `Folder ID: ${folderId}`)

    // Create Processed folder structure first
    addLog('info', 'Creating Processed folder structure...')
    const processedStructure = await getOrCreateProcessedStructure(folderId)
    addLog('success', `Created: Processed/${excelSettings.value.companyId}/${excelSettings.value.year}/Mostanadat`)
    processedFolderLink.value = `https://drive.google.com/drive/folders/${processedStructure.processedFolderId}`

    // Get or create master Google Sheet in Processed folder
    let masterSheetId = null
    try {
      addLog('info', 'Checking for master Invoice Data sheet...')
      const { spreadsheetId, created: sheetCreated } = await getOrCreateMasterSheet(processedStructure.processedFolderId)
      masterSheetId = spreadsheetId
      addLog('success', sheetCreated ? 'Created new master Invoice Data sheet' : 'Found existing master Invoice Data sheet')
    } catch (sheetErr) {
      if (sheetErr.message?.includes('403') || sheetErr.message?.includes('permission')) {
        addLog('error', 'Sheets API permission denied', sheetErr.message)
        addLog('warning', 'To fix: Enable Google Sheets API and re-sign in')
      } else {
        addLog('warning', 'Could not create/find master sheet', sheetErr.message)
      }
    }

    // Fetch PDF files from source folder
    addLog('info', 'Fetching PDF files from Google Drive...')
    const pdfFiles = await fetchDriveFiles(folderId)
    addLog('info', `Found ${pdfFiles.length} PDF file(s) to process`)

    if (pdfFiles.length === 0) {
      error.value = 'No PDF files found in the specified folder'
      addLog('warning', 'No PDF files found in the folder')
      isProcessing.value = false
      return
    }

    files.value = pdfFiles.map(f => ({
      id: f.id,
      name: f.name,
      status: 'pending',
      statusText: 'Waiting...',
      originalSize: null,
      processedSize: null,
      blob: null
    }))

    // Process each file
    for (let i = 0; i < files.value.length; i++) {
      const file = files.value[i]
      addLog('info', `Processing file ${i + 1}/${files.value.length}: ${file.name}`)

      try {
        file.status = 'downloading'
        file.statusText = 'Downloading...'
        const pdfBytes = await downloadFile(file.id)
        file.originalSize = pdfBytes.byteLength
        addLog('success', `Downloaded: ${file.name}`, formatBytes(file.originalSize))

        file.status = 'processing'
        file.statusText = 'Processing...'
        const { pdfBytes: processedBytes, extractedText } = await processPdf(pdfBytes, file.name, status => { file.statusText = status })
        file.processedSize = processedBytes.byteLength
        file.blob = new Blob([processedBytes], { type: 'application/pdf' })

        file.status = 'uploading'
        file.statusText = 'Uploading to Mostanadat...'

        // Extract invoice number from filename and rename PDF (e.g., "3427 - آموزشگاه.pdf" → "3427.pdf")
        const invoiceNumber = parseFilename(file.name).invoiceNumber
        const newPdfName = invoiceNumber ? `${invoiceNumber}.pdf` : file.name

        // Upload PDF directly to Mostanadat folder with renamed filename
        await uploadToDrive(file.blob, newPdfName, processedStructure.mostanadatFolderId)
        addLog('success', `Uploaded PDF to Mostanadat: ${newPdfName}`)

        // Handle extracted text - append to master sheet
        if (options.value.ocr && extractedText?.trim() && masterSheetId) {
          try {
            const invoiceData = parseInvoiceText(extractedText, file.name)
            if (invoiceData?.items?.length > 0) {
              await appendToMasterSheet(masterSheetId, invoiceData)
              addLog('success', `Appended ${invoiceData.items.length} row(s) to master sheet`)
            }
          } catch (err) {
            addLog('warning', `Failed to append to master sheet`, err.message)
          }
        }

        file.status = 'done'
        file.statusText = `Complete (${formatBytes(file.originalSize)} → ${formatBytes(file.processedSize)})`

      } catch (err) {
        file.status = 'error'
        file.statusText = err.message || 'Processing failed'
        addLog('error', `Failed to process: ${file.name}`, err.message)
      }

      progress.value = ((i + 1) / files.value.length) * 100
    }

    const successCount = files.value.filter(f => f.status === 'done').length

    // Process Excel template - copy as Google Sheet and populate
    if (masterSheetId && successCount > 0) {
      try {
        addLog('info', 'Starting Excel template processing...')
        await processExcelTemplate(folderId, masterSheetId, processedStructure.processedFolderId)
        addLog('success', 'Excel template processing complete')

        // Delete the temporary Invoice Data sheet after processing
        try {
          addLog('info', 'Cleaning up: Deleting Invoice Data sheet...')
          await deleteFileFromDrive(masterSheetId)
          addLog('success', 'Deleted Invoice Data sheet')
        } catch (deleteErr) {
          addLog('warning', 'Failed to delete Invoice Data sheet', deleteErr.message)
        }
      } catch (err) {
        addLog('warning', 'Failed to process Excel template', err.message)
      }
    }

    if (successCount > 0) {
      successMessage.value = `Successfully processed ${successCount} files!`
      addLog('success', `Processing complete: ${successCount} succeeded`)
    }

  } catch (err) {
    error.value = err.message || 'An error occurred'
    addLog('error', 'Processing failed', err.message)
  } finally {
    isProcessing.value = false
  }
}

// Google Drive API functions
async function fetchDriveFiles(folderId) {
  const query = `'${folderId}' in parents and mimeType='application/pdf' and trashed=false`
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,size)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to fetch files: ${errorData.error?.message || response.status}`)
  }

  const data = await response.json()
  const seenIds = new Set()
  const seenNames = new Set()

  return (data.files || []).filter(file => {
    if (seenIds.has(file.id)) return false
    seenIds.add(file.id)

    const nameLower = file.name.toLowerCase()
    if (seenNames.has(nameLower)) return false
    seenNames.add(nameLower)

    if (nameLower.startsWith('temp_') || nameLower.startsWith('normalized_')) return false
    if (!nameLower.endsWith('.pdf')) return false

    return true
  })
}

async function getOrCreateFolder(parentFolderId, folderName) {
  const query = `'${parentFolderId}' in parents and name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`
  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,parents)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  if (searchResponse.ok) {
    const data = await searchResponse.json()
    if (data.files?.[0]?.parents?.includes(parentFolderId)) {
      return { folderId: data.files[0].id, created: false }
    }
  }

  const createResponse = await fetch(
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true',
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: folderName, mimeType: 'application/vnd.google-apps.folder', parents: [parentFolderId] })
    }
  )

  if (!createResponse.ok) {
    const errorData = await createResponse.json().catch(() => ({}))
    throw new Error(`Failed to create folder: ${errorData.error?.message || 'Unknown error'}`)
  }

  const createData = await createResponse.json()
  return { folderId: createData.id, created: true }
}

async function getOrCreateProcessedStructure(parentFolderId) {
  const processed = await getOrCreateFolder(parentFolderId, 'Processed')
  const company = await getOrCreateFolder(processed.folderId, excelSettings.value.companyId)
  const year = await getOrCreateFolder(company.folderId, excelSettings.value.year)
  const mostanadat = await getOrCreateFolder(year.folderId, 'Mostanadat')
  return {
    processedFolderId: processed.folderId,
    companyFolderId: company.folderId,
    yearFolderId: year.folderId,
    mostanadatFolderId: mostanadat.folderId,
    created: processed.created
  }
}

async function getOrCreateMasterSheet(sheetsFolderId) {
  const sheetName = 'Invoice Data'
  const query = `'${sheetsFolderId}' in parents and name='${sheetName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  if (searchResponse.ok) {
    const data = await searchResponse.json()
    if (data.files?.[0]) return { spreadsheetId: data.files[0].id, created: false }
  }

  const createResponse = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      properties: { title: sheetName },
      sheets: [{ properties: { title: 'Invoice Data', rightToLeft: true } }]
    })
  })

  if (!createResponse.ok) {
    const errorData = await createResponse.json().catch(() => ({}))
    throw new Error(`Failed to create master sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  const sheetData = await createResponse.json()
  const spreadsheetId = sheetData.spreadsheetId

  // Move to target folder
  const fileResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${spreadsheetId}?fields=parents`, {
    headers: { Authorization: `Bearer ${accessToken.value}` }
  })
  if (fileResponse.ok) {
    const fileData = await fileResponse.json()
    await fetch(
      `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?addParents=${sheetsFolderId}&removeParents=${fileData.parents?.join(',') || ''}`,
      { method: 'PATCH', headers: { Authorization: `Bearer ${accessToken.value}` } }
    )
  }

  // Add headers
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A1:G1?valueInputOption=RAW`, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [['شماره', 'تاریخ', 'نام شخص حقیقی / حقوقی', 'ردیف', 'شرح کالا یا خدمات', 'مبلغ کل پس از تخفیف', 'جمع مالیات و عوارض']] })
  })

  return { spreadsheetId, created: true }
}

async function appendToMasterSheet(spreadsheetId, invoiceData) {
  if (!invoiceData?.items?.length) return { updatedRows: 0 }

  const rows = invoiceData.items.map(item => [
    invoiceData.invoiceNumber || '',
    invoiceData.date || '',
    invoiceData.buyerName || '',
    item.rowNumber || '',
    item.description || '',
    item.totalAfterDiscount || '',
    item.taxAndDuties || ''
  ])

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A:G:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: rows })
    }
  )

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to append: ${errorData.error?.message || 'Unknown error'}`)
  }

  return { updatedRows: rows.length }
}

async function findExcelTemplateFile(folderId, filename) {
  const xlsxFilename = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`

  // First, list ALL xlsx files in the folder for debugging
  const listQuery = `'${folderId}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false`
  const listResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(listQuery)}&fields=files(id,name,mimeType)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  if (listResponse.ok) {
    const listData = await listResponse.json()
    const xlsxFiles = listData.files || []
    console.log(`[findExcelTemplateFile] Found ${xlsxFiles.length} xlsx file(s) in folder:`, xlsxFiles.map(f => f.name))
    addLog('info', `Found ${xlsxFiles.length} xlsx file(s) in source folder`)
    xlsxFiles.forEach(f => addLog('info', `  - "${f.name}"`))

    // Try exact match first
    let match = xlsxFiles.find(f => f.name === xlsxFilename)
    if (match) {
      addLog('info', `Exact match found: "${match.name}"`)
      return match
    }

    // Try case-insensitive match
    match = xlsxFiles.find(f => f.name.toLowerCase() === xlsxFilename.toLowerCase())
    if (match) {
      addLog('info', `Case-insensitive match found: "${match.name}"`)
      return match
    }

    // Try partial match (contains the filename without extension)
    const baseFilename = filename.replace(/\.xlsx$/i, '')
    match = xlsxFiles.find(f => f.name.includes(baseFilename) || baseFilename.includes(f.name.replace(/\.xlsx$/i, '')))
    if (match) {
      addLog('info', `Partial match found: "${match.name}"`)
      return match
    }

    addLog('warning', `No match for "${xlsxFilename}" among: ${xlsxFiles.map(f => f.name).join(', ') || '(none)'}`)
  } else {
    addLog('warning', 'Failed to list xlsx files in folder')
  }

  return null
}

async function readMasterSheetData(spreadsheetId) {
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A:G`, {
    headers: { Authorization: `Bearer ${accessToken.value}` }
  })

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to read master sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  const data = await response.json()
  return data.values || []
}

// Duplicate xlsx file AS-IS to destination folder (keeps xlsx format, does NOT convert)
async function duplicateXlsxFile(fileId, newName, destinationFolderId) {
  const xlsxName = newName.endsWith('.xlsx') ? newName : `${newName}.xlsx`
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/copy?supportsAllDrives=true`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: xlsxName,
      parents: [destinationFolderId]
      // No mimeType specified = keeps original xlsx format
    })
  })

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to duplicate xlsx file: ${errorData.error?.message || 'Unknown error'}`)
  }

  return await response.json()
}

// Convert an xlsx file to Google Sheets format (for API editing)
async function convertXlsxToGoogleSheet(fileId, newName, destinationFolderId) {
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/copy?supportsAllDrives=true`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: newName,
      parents: [destinationFolderId],
      // This mimeType converts xlsx to native Google Sheets format
      mimeType: 'application/vnd.google-apps.spreadsheet'
    })
  })

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to convert to Google Sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  return await response.json()
}

// Export a Google Sheet back to xlsx format
async function exportSheetToXlsx(sheetId, filename, destinationFolderId) {
  // Download as xlsx
  const xlsxName = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  const exportResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files/${sheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  if (!exportResponse.ok) {
    const errorData = await exportResponse.json().catch(() => ({}))
    throw new Error(`Failed to export sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  const xlsxBlob = await exportResponse.blob()

  // Upload the xlsx to destination folder
  return await uploadToDrive(xlsxBlob, xlsxName, destinationFolderId, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

// Update cells in Google Sheets (triggers formula recalculation automatically)
async function updateSheetCells(spreadsheetId, sheetName, cellUpdates) {
  const data = cellUpdates.map(update => ({
    range: `'${sheetName}'!${update.cell}`,
    values: [[update.value]]
  }))

  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken.value}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data })
  })

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to update cells: ${errorData.error?.message || 'Unknown error'}`)
  }

  return await response.json()
}

// Process Excel template: duplicate to Processed folder, edit via Sheets API, export back to xlsx
// Flow: 1) Duplicate original xlsx AS-IS to Processed folder
//       2) Convert the DUPLICATE (not original!) to temp Google Sheet for editing
//       3) Edit temp sheet, export back to xlsx
// This ensures: Original xlsx is NEVER touched
// NOTE: Temp files are kept for debugging - user can manually delete them
async function processExcelTemplate(sourceFolderId, masterSheetId, processedFolderId) {
  const filename = excelSettings.value.filename
  const sheetName = excelSettings.value.sheetName

  addLog('info', `Looking for Excel template: ${filename}.xlsx`)
  const templateFile = await findExcelTemplateFile(sourceFolderId, filename)

  if (!templateFile) {
    addLog('warning', `Excel template not found`, 'Skipping Excel mapping')
    return null
  }

  addLog('success', `Found Excel template: ${templateFile.name} (ID: ${templateFile.id})`)

  // Read master sheet data first
  addLog('info', 'Reading data from master sheet...')
  const masterSheetData = await readMasterSheetData(masterSheetId)
  const dataRowCount = masterSheetData.length - 1

  if (dataRowCount === 0) {
    addLog('warning', 'No data rows in master sheet (only header)')
    return null
  }

  addLog('info', `Read ${dataRowCount} data row(s) from master sheet`)

  // Log the actual data for debugging
  const dataRows = masterSheetData.slice(1) // Skip header
  dataRows.forEach((row, idx) => {
    const [invoiceNumber, date, buyerName, rowNum, description, totalAfterDiscount, taxAndDuties] = row
    addLog('info', `Row ${idx + 1}: Invoice=${invoiceNumber}, Desc="${description}", Total=${totalAfterDiscount}`)
  })

  // Step 1: DUPLICATE the original xlsx to Processed folder with a TEMP name
  // Use temp name to avoid conflict when we upload the final edited version
  // The ORIGINAL file is NEVER modified!
  const tempDuplicateName = `_TEMP_DUPLICATE_${filename}_${Date.now()}`
  addLog('info', `Step 1: Duplicating original xlsx to Processed folder (temp name)...`)
  let duplicatedXlsx
  try {
    duplicatedXlsx = await duplicateXlsxFile(templateFile.id, tempDuplicateName, processedFolderId)
    addLog('success', `Created temp duplicate: ${tempDuplicateName}.xlsx (ID: ${duplicatedXlsx.id})`)
  } catch (dupErr) {
    addLog('error', `Failed to duplicate xlsx: ${dupErr.message}`)
    throw dupErr
  }

  // Step 2: Convert the temp DUPLICATE to temp Google Sheet for editing
  // We use Google Sheets format temporarily because Sheets API can update cells and recalculate formulas
  addLog('info', `Step 2: Converting temp duplicate to Google Sheet for editing...`)
  const tempSheetName = `_TEMP_SHEET_${filename}_${Date.now()}`
  let tempSheet
  try {
    tempSheet = await convertXlsxToGoogleSheet(duplicatedXlsx.id, tempSheetName, processedFolderId)
    addLog('success', `Created Google Sheet: ${tempSheetName} (ID: ${tempSheet.id})`)
  } catch (convErr) {
    addLog('error', `Failed to convert to Google Sheet: ${convErr.message}`)
    throw convErr
  }

  // Build cell updates from master sheet data
  const cellUpdates = []
  const startRow = 8 // Excel data starts at row 8

  dataRows.forEach((row, index) => {
    const excelRow = startRow + index
    const [invoiceNumber, date, buyerName, rowNum, description, totalAfterDiscount, taxAndDuties] = row

    cellUpdates.push(
      { cell: `C${excelRow}`, value: excelSettings.value.companyId },
      { cell: `D${excelRow}`, value: excelSettings.value.companyName },
      { cell: `E${excelRow}`, value: excelSettings.value.year },
      { cell: `F${excelRow}`, value: rowNum || '' },
      { cell: `G${excelRow}`, value: description || '' },
      { cell: `H${excelRow}`, value: invoiceNumber || '' },
      { cell: `I${excelRow}`, value: date || '' },
      { cell: `J${excelRow}`, value: totalAfterDiscount || '' },
      { cell: `K${excelRow}`, value: taxAndDuties || '' },
      { cell: `L${excelRow}`, value: buyerName || '' }
    )
  })

  // Step 3: Update cells in the Google Sheet (formulas auto-recalculate)
  addLog('info', `Step 3: Updating ${cellUpdates.length} cells in sheet "${sheetName}"...`)
  try {
    await updateSheetCells(tempSheet.id, sheetName, cellUpdates)
    addLog('success', 'Data populated successfully - formulas recalculated')
  } catch (updateErr) {
    addLog('error', `Failed to update cells: ${updateErr.message}`)
    throw updateErr
  }

  // Step 4: Export the edited Google Sheet back to xlsx format
  addLog('info', `Step 4: Exporting to xlsx format...`)
  let exportedFile
  try {
    exportedFile = await exportSheetToXlsx(tempSheet.id, filename, processedFolderId)
    addLog('success', `Exported xlsx file: ${filename}.xlsx (ID: ${exportedFile.id})`)
  } catch (exportErr) {
    addLog('error', `Failed to export xlsx: ${exportErr.message}`)
    throw exportErr
  }

  // Step 5: Delete the temp duplicate xlsx (we now have the final edited version with the correct name)
  try {
    addLog('info', 'Cleaning up: Deleting temp duplicate xlsx...')
    await deleteFileFromDrive(duplicatedXlsx.id)
    addLog('success', 'Deleted temp duplicate')
  } catch (deleteErr) {
    addLog('warning', `Could not delete duplicate (ID: ${duplicatedXlsx.id}): ${deleteErr.message}`)
  }

  // Step 6: Delete the temporary Google Sheet
  try {
    addLog('info', 'Cleaning up: Deleting temp Google Sheet...')
    await deleteFileFromDrive(tempSheet.id)
    addLog('success', 'Deleted temp Google Sheet')
  } catch (deleteErr) {
    addLog('warning', `Could not delete temp sheet (ID: ${tempSheet.id}): ${deleteErr.message}`)
  }

  addLog('success', `Excel processing complete! File: ${filename}.xlsx in Processed folder`)
  return exportedFile
}

async function downloadFile(fileId) {
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
    headers: { Authorization: `Bearer ${accessToken.value}` }
  })
  if (!response.ok) throw new Error('Failed to download file')
  return await response.arrayBuffer()
}

// Delete a file from Google Drive
async function deleteFileFromDrive(fileId) {
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?supportsAllDrives=true`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${accessToken.value}` }
  })
  if (!response.ok && response.status !== 204) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to delete file: ${errorData.error?.message || 'Unknown error'}`)
  }
  return true
}

async function processPdf(pdfBytes, filename, onProgress = null) {
  let extractedText = ''

  try {
    const pdfDoc = await PDFDocument.load(pdfBytes, { ignoreEncryption: true, updateMetadata: false })
    const pages = pdfDoc.getPages()

    if (options.value.ocr) {
      try {
        const loadingTask = pdfjsLib.getDocument({ data: pdfBytes })
        const pdfJsDoc = await loadingTask.promise
        const textParts = []
        let totalDirectTextLength = 0

        // Try direct text extraction first
        for (let i = 1; i <= pdfJsDoc.numPages; i++) {
          onProgress?.(`Extracting text from page ${i}/${pdfJsDoc.numPages}`)
          const page = await pdfJsDoc.getPage(i)
          const textContent = await page.getTextContent()
          const pageText = textContent.items.map(item => item.str).join(' ').trim()
          if (pageText) {
            textParts.push(`--- Page ${i} ---\n${pageText}`)
            totalDirectTextLength += pageText.length
          }
        }

        if (totalDirectTextLength / pdfJsDoc.numPages > 50) {
          extractedText = normalizeArabicText(textParts.join('\n\n'))
          addLog('success', `Direct text extraction complete`, `${extractedText.length} chars`)
        } else {
          // Fall back to OCR
          textParts.length = 0
          for (let i = 1; i <= pdfJsDoc.numPages; i++) {
            onProgress?.(`OCR page ${i}/${pdfJsDoc.numPages}`)
            const page = await pdfJsDoc.getPage(i)
            const viewport = page.getViewport({ scale: 2.0 })
            const canvas = document.createElement('canvas')
            const context = canvas.getContext('2d')
            canvas.height = viewport.height
            canvas.width = viewport.width
            await page.render({ canvasContext: context, viewport }).promise
            const result = await Tesseract.recognize(canvas.toDataURL('image/png'), 'eng+fas+ara')
            if (result.data.text.trim()) textParts.push(`--- Page ${i} ---\n${result.data.text.trim()}`)
          }
          extractedText = normalizeArabicText(textParts.join('\n\n'))
          if (extractedText) addLog('success', `OCR complete`, `${extractedText.length} chars`)
        }
      } catch (err) {
        addLog('warning', `Text extraction failed`, err.message)
      }
    }

    // Standardize page sizes if enabled
    if (options.value.standardize && options.value.pageSize !== 'ORIGINAL') {
      const targetSize = getPageDimensions(options.value.pageSize)
      for (const page of pages) {
        const { width, height } = page.getSize()
        if (Math.abs(width - targetSize.width) > 1 || Math.abs(height - targetSize.height) > 1) {
          const scale = Math.min(targetSize.width / width, targetSize.height / height)
          page.setSize(targetSize.width, targetSize.height)
          page.scaleContent(scale, scale)
          page.translateContent((targetSize.width - width * scale) / 2, (targetSize.height - height * scale) / 2)
        }
      }
    }

    pdfDoc.setTitle(filename.replace('.pdf', ''))
    pdfDoc.setProducer('PDF Processor')
    pdfDoc.setModificationDate(new Date())

    const processedPdfBytes = await pdfDoc.save(options.value.compress ? { useObjectStreams: true } : {})
    return { pdfBytes: processedPdfBytes, extractedText }

  } catch (err) {
    return { pdfBytes, extractedText }
  }
}

function getPageDimensions(size) {
  const dimensions = {
    LETTER: { width: 612, height: 792 },
    A4: { width: 595.28, height: 841.89 },
    LEGAL: { width: 612, height: 1008 }
  }
  return dimensions[size] || dimensions.LETTER
}

function normalizeArabicText(text) {
  if (!text) return text
  return text.normalize('NFKC').replace(/[\u200B-\u200F\u202A-\u202E\uFEFF]/g, '').replace(/  +/g, ' ')
}

function parseFilename(filename) {
  if (!filename) return { invoiceNumber: '', buyerName: '' }
  const nameWithoutExt = filename.replace(/\.pdf$/i, '')
  const parts = nameWithoutExt.split(' - ')
  if (parts.length >= 2) {
    return { invoiceNumber: parts[0].trim(), buyerName: parts.slice(1).join(' - ').trim() }
  }
  return { invoiceNumber: '', buyerName: '' }
}

function parseInvoiceText(text, filename = '') {
  if (!text) return null
  const normalizedText = normalizeArabicText(text)
  const result = { invoiceNumber: '', date: '', buyerName: '', items: [], rawText: text }

  // Extract from filename first
  const filenameData = parseFilename(filename)
  if (filenameData.invoiceNumber) result.invoiceNumber = filenameData.invoiceNumber
  if (filenameData.buyerName) result.buyerName = filenameData.buyerName

  // Fallback to text extraction
  if (!result.invoiceNumber) {
    const match = normalizedText.match(/شماره[:\s]*(\d+)/)
    if (match) result.invoiceNumber = match[1]
  }

  const dateMatch = normalizedText.match(/(\d{4}\/\d{2}\/\d{2})\s*:?\s*تاریخ|تاریخ\s*:?\s*(\d{4}\/\d{2}\/\d{2})/)
  if (dateMatch) result.date = dateMatch[1] || dateMatch[2]

  if (!result.buyerName) {
    const label = 'نام شخص حقیقی / حقوقی'
    const firstIndex = normalizedText.indexOf(label)
    if (firstIndex !== -1) {
      const secondIndex = normalizedText.indexOf(label, firstIndex + label.length)
      if (secondIndex !== -1) {
        const match = normalizedText.substring(secondIndex + label.length).match(/^[:\s]*([^\n:]+)/)
        if (match?.[1]?.length > 2 && !match[1].includes('شناسه')) result.buyerName = match[1].trim()
      }
    }
  }

  // Parse table rows
  const headerMatch = normalizedText.match(/تعداد\s+شرح کالا یا خدمات\s+ردیف/)
  if (headerMatch) {
    let dataSection = normalizedText.substring(normalizedText.indexOf(headerMatch[0]) + headerMatch[0].length)
    const totalsIndex = dataSection.indexOf('جمع کل')
    if (totalsIndex !== -1) dataSection = dataSection.substring(0, totalsIndex)

    // Pattern for table rows - description (.+?) uses NON-GREEDY match
    // This correctly handles descriptions with numbers like "پلن 2 از خدمات میزبانی وب"
    // The lookahead ensures row number is followed by next row's total ([\d,]{3,}) or "جمع" (totals section)
    const rowPattern = /([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+(\d+)\s+(.+?)\s+(\d{1,2})(?=\s+[\d,]{3,}|\s+جمع|\s*$)/g
    let match
    while ((match = rowPattern.exec(dataSection)) !== null) {
      const description = match[8].trim()
      if (!description.includes('شرح') && !description.includes('ردیف') && !description.includes('جمع')) {
        result.items.push({
          rowNumber: match[9],
          description,
          quantity: match[7],
          unitPrice: match[6],
          total: match[5],
          discount: match[4],
          totalAfterDiscount: match[3],
          taxAndDuties: match[2],
          grandTotal: match[1]
        })
      }
    }
  }

  return result
}

async function uploadToDrive(blob, filename, parentFolderId, mimeType = null) {
  const fileMimeType = mimeType || blob.type || 'application/pdf'
  const escapedFilename = filename.replace(/'/g, "\\'")
  const searchQuery = `'${parentFolderId}' in parents and name='${escapedFilename}' and trashed=false`

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(searchQuery)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: { Authorization: `Bearer ${accessToken.value}` } }
  )

  let existingFileId = null
  if (searchResponse.ok) {
    const data = await searchResponse.json()
    existingFileId = data.files?.[0]?.id
  }

  const form = new FormData()
  if (existingFileId) {
    form.append('metadata', new Blob([JSON.stringify({ mimeType: fileMimeType })], { type: 'application/json' }))
    form.append('file', blob)

    const response = await fetch(
      `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=multipart&supportsAllDrives=true`,
      { method: 'PATCH', headers: { Authorization: `Bearer ${accessToken.value}` }, body: form }
    )
    if (!response.ok) throw new Error('Failed to update file')
    return await response.json()
  } else {
    form.append('metadata', new Blob([JSON.stringify({ name: filename, parents: [parentFolderId], mimeType: fileMimeType })], { type: 'application/json' }))
    form.append('file', blob)

    const response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
      { method: 'POST', headers: { Authorization: `Bearer ${accessToken.value}` }, body: form }
    )
    if (!response.ok) throw new Error('Failed to upload file')
    return await response.json()
  }
}

// Initialize
onMounted(async () => {
  const savedToken = localStorage.getItem('google_access_token')
  const tokenExpiry = localStorage.getItem('google_token_expiry')
  const savedEmail = localStorage.getItem('google_user_email')

  if (!savedToken || (tokenExpiry && Date.now() > parseInt(tokenExpiry))) {
    clearAuthData()
    return
  }

  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${savedToken}` }
    })
    if (response.ok) {
      const data = await response.json()
      accessToken.value = savedToken
      isAuthenticated.value = true
      userEmail.value = data.email || savedEmail || ''
    } else {
      clearAuthData()
    }
  } catch (err) {
    if (savedEmail) {
      accessToken.value = savedToken
      isAuthenticated.value = true
      userEmail.value = savedEmail
    } else {
      clearAuthData()
    }
  }
})
</script>

<style scoped>
.mdi-spin {
  animation: spin 1s linear infinite;
}
@keyframes spin {
  from { transform: rotate(0deg); }
  to { transform: rotate(360deg); }
}
</style>
