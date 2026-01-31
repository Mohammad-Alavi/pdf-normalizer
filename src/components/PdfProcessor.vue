<template>
  <v-card class="mx-auto" max-width="900">
    <v-card-title class="text-h5 pa-6 pb-2 d-flex align-center">
      <v-icon color="primary" class="mr-2">mdi-google-drive</v-icon>
      Google Drive PDF Processor
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
              <li>پوشه <code>Normalized</code> ایجاد می‌شود با زیرپوشه‌های Texts و Sheets</li>
              <li>فایل‌های PDF نرمال‌شده در پوشه Normalized قرار می‌گیرند</li>
              <li>فایل‌های متنی در پوشه Texts ذخیره می‌شوند</li>
              <li>Master Sheet (Invoice Data) در پوشه Sheets ایجاد می‌شود</li>
              <li>پوشه <code>Processed</code> ایجاد می‌شود با ساختار <code>شناسه/سال/Mostanadat</code></li>
              <li>کپی فایل اکسل با داده‌های پر شده در پوشه Processed قرار می‌گیرد</li>
            </ul>

            <v-divider class="my-3"></v-divider>

            <h4 class="mb-2" style="direction: ltr; text-align: left;">English Instructions:</h4>
            <div style="direction: ltr; text-align: left;">
              <p><strong>1. Prepare Google Drive folder:</strong> Create a folder and add PDF invoices with filename format: <code>InvoiceNumber - BuyerName.pdf</code></p>
              <p><strong>2. Optional Excel template:</strong> Place the Excel template file in the same folder</p>
              <p><strong>3. Run:</strong> Paste folder link, sign in with Google, click "Process PDFs"</p>
              <p><strong>4. Output:</strong> Normalized folder with PDFs/texts/sheets, Processed folder with Excel and company/year/Mostanadat structure</p>
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

      <!-- Google OAuth (needed for uploading back) -->
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

      <!-- Action Buttons -->
      <div class="d-flex mb-4" style="gap: 16px;">
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
        <v-btn
          v-if="processedFiles.length > 0"
          color="secondary"
          size="large"
          variant="outlined"
          @click="downloadAll"
        >
          <v-icon class="mr-2">mdi-download</v-icon>
          Download All (ZIP)
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

      <!-- Link to Normalized folder -->
      <v-alert
        v-if="normalizedFolderLink"
        type="info"
        variant="tonal"
        class="mb-4"
      >
        <div class="d-flex align-center">
          <v-icon class="mr-2">mdi-folder-open</v-icon>
          <span>Normalized files uploaded to: </span>
          <a :href="normalizedFolderLink" target="_blank" class="ml-1">
            Open Normalized folder in Google Drive
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
import JSZip from 'jszip'
import Tesseract from 'tesseract.js'
import * as pdfjsLib from 'pdfjs-dist'
import * as XLSX from 'xlsx'

// Configure PDF.js worker - use jsdelivr CDN
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
const processedFiles = ref([])
const error = ref('')
const successMessage = ref('')
const processingLogs = ref([])
const normalizedFolderLink = ref('')
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
  return driveLink.value &&
         !linkError.value &&
         extractFolderId(driveLink.value) &&
         isAuthenticated.value
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

// Methods
function extractFolderId(url) {
  if (!url) return null
  // Handle various Google Drive URL formats
  const patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{25,})/
  ]
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
  const folderId = extractFolderId(driveLink.value)
  if (!folderId) {
    linkError.value = 'Invalid Google Drive folder link'
  } else {
    linkError.value = ''
  }
}

function getFileStatusIcon(status) {
  switch (status) {
    case 'pending': return 'mdi-clock-outline'
    case 'downloading': return 'mdi-download'
    case 'processing': return 'mdi-cog mdi-spin'
    case 'uploading': return 'mdi-upload'
    case 'done': return 'mdi-check-circle'
    case 'error': return 'mdi-alert-circle'
    default: return 'mdi-file'
  }
}

function getFileStatusColor(status) {
  switch (status) {
    case 'pending': return 'grey'
    case 'downloading': return 'info'
    case 'processing': return 'warning'
    case 'uploading': return 'info'
    case 'done': return 'success'
    case 'error': return 'error'
    default: return 'grey'
  }
}

function getSizeChange(original, processed) {
  const change = ((original - processed) / original * 100).toFixed(0)
  if (change > 0) {
    return { text: `Saved ${change}%`, color: 'success' }
  } else if (change < 0) {
    return { text: `+${Math.abs(change)}%`, color: 'warning' }
  } else {
    return { text: 'No change', color: 'grey' }
  }
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

function addLog(type, message, details = null) {
  const timestamp = new Date().toLocaleTimeString()
  processingLogs.value.push({
    id: Date.now(),
    timestamp,
    type, // 'info', 'success', 'warning', 'error'
    message,
    details
  })
}

function clearLogs() {
  processingLogs.value = []
}

async function copyLogs() {
  const logText = processingLogs.value.map(log => {
    const prefix = log.type === 'success' ? '✓' : log.type === 'error' ? '✗' : log.type === 'warning' ? '⚠' : 'ℹ'
    const details = log.details ? ` — ${log.details}` : ''
    return `[${log.timestamp}] ${prefix} ${log.message}${details}`
  }).join('\n')

  try {
    await navigator.clipboard.writeText(logText)
    // Show a brief success message
    const originalLogs = [...processingLogs.value]
    addLog('success', 'Logs copied to clipboard!')
    // Remove the "copied" message after 2 seconds
    setTimeout(() => {
      processingLogs.value = originalLogs
    }, 2000)
  } catch (err) {
    console.error('Failed to copy logs:', err)
    addLog('error', 'Failed to copy logs to clipboard')
  }
}

// Google OAuth
async function authenticateGoogle() {
  isAuthenticating.value = true
  error.value = ''

  try {
    // Initialize Google Identity Services
    if (!window.google?.accounts?.oauth2) {
      // Load Google Identity Services script
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
        const expiryTime = Date.now() + (response.expires_in * 1000)
        localStorage.setItem('google_token_expiry', String(expiryTime))

        isAuthenticated.value = true

        // Fetch and save user info
        try {
          await fetchUserInfo()
          if (userEmail.value) {
            localStorage.setItem('google_user_email', userEmail.value)
          }
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
  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    })

    if (!response.ok) {
      console.error('Failed to fetch user info:', response.status)
      return null
    }

    const data = await response.json()
    userEmail.value = data.email || ''
    return data
  } catch (err) {
    console.error('Failed to fetch user info:', err)
    return null
  }
}

function signOut() {
  const token = accessToken.value
  clearAuthData()
  if (window.google?.accounts?.oauth2 && token) {
    google.accounts.oauth2.revoke(token)
  }
}

// File Processing
async function processFiles() {
  if (!canProcess.value) return

  isProcessing.value = true
  error.value = ''
  successMessage.value = ''
  files.value = []
  processedFiles.value = []
  progress.value = 0
  normalizedFolderLink.value = ''
  clearLogs()

  try {
    const folderId = extractFolderId(driveLink.value)
    addLog('info', 'Starting PDF processing', `Folder ID: ${folderId}`)

    // Get or create the "Normalized" subfolder (for PDFs, Texts, Sheets)
    addLog('info', 'Checking for Normalized folder...')
    const { folderId: normalizedFolderId, textsFolderId, sheetsFolderId, created } = await getOrCreateNormalizedFolder(folderId)

    if (created) {
      addLog('success', 'Created new Normalized folder structure', `Main: ${normalizedFolderId}, Texts: ${textsFolderId}, Sheets: ${sheetsFolderId}`)
    } else {
      addLog('success', 'Found existing Normalized folder structure', `Main: ${normalizedFolderId}, Texts: ${textsFolderId}, Sheets: ${sheetsFolderId}`)
    }

    // Get or create the master Google Sheet for all invoice data
    addLog('info', 'Checking for master Invoice Data sheet...')
    let masterSheetId = null
    try {
      const { spreadsheetId, created: sheetCreated } = await getOrCreateMasterSheet(sheetsFolderId)
      masterSheetId = spreadsheetId
      if (sheetCreated) {
        addLog('success', 'Created new master Invoice Data sheet', `Sheet ID: ${spreadsheetId}`)
      } else {
        addLog('success', 'Found existing master Invoice Data sheet', `Sheet ID: ${spreadsheetId}`)
      }
    } catch (sheetErr) {
      // Check if this is a permission/scope error
      const isAuthError = sheetErr.message?.includes('403') ||
                          sheetErr.message?.includes('permission') ||
                          sheetErr.message?.includes('scope') ||
                          sheetErr.message?.includes('access') ||
                          sheetErr.message?.includes('PERMISSION_DENIED')
      if (isAuthError) {
        addLog('error', 'Sheets API permission denied', sheetErr.message)
        addLog('warning', 'To fix: 1) Go to Google Cloud Console and enable Google Sheets API', 'https://console.cloud.google.com/apis/library/sheets.googleapis.com')
        addLog('warning', 'To fix: 2) Go to myaccount.google.com/permissions and revoke this app, then sign in again')
      } else {
        addLog('warning', 'Could not create/find master sheet, will skip data aggregation', sheetErr.message)
      }
    }

    // Set the link to open the Normalized folder
    normalizedFolderLink.value = `https://drive.google.com/drive/folders/${normalizedFolderId}`

    // Fetch files from Google Drive (excluding files already in Normalized folder)
    addLog('info', 'Fetching PDF files from Google Drive...')
    const pdfFiles = await fetchDriveFiles(folderId, normalizedFolderId)
    addLog('info', `Found ${pdfFiles.length} PDF file(s) to process`, pdfFiles.map(f => f.name).join(', '))

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
        // Download
        file.status = 'downloading'
        file.statusText = 'Downloading...'
        addLog('info', `Downloading: ${file.name}`)
        const pdfBytes = await downloadFile(file.id)
        file.originalSize = pdfBytes.byteLength
        addLog('success', `Downloaded: ${file.name}`, `Size: ${formatBytes(file.originalSize)}`)

        // Process
        file.status = 'processing'
        file.statusText = 'Processing...'
        addLog('info', `Processing PDF: ${file.name}`, `Options: OCR=${options.value.ocr}, Standardize=${options.value.standardize}, Compress=${options.value.compress}, PageSize=${options.value.pageSize}`)

        const processResult = await processPdf(pdfBytes, file.name, (status) => {
          file.statusText = status
        })
        const processedBytes = processResult.pdfBytes
        const extractedText = processResult.extractedText

        file.processedSize = processedBytes.byteLength
        file.blob = new Blob([processedBytes], { type: 'application/pdf' })

        const sizeChange = file.originalSize - file.processedSize
        const sizeChangeText = sizeChange > 0 ? `Reduced by ${formatBytes(sizeChange)}` : sizeChange < 0 ? `Increased by ${formatBytes(Math.abs(sizeChange))}` : 'No size change'
        addLog('success', `Processed: ${file.name}`, `${formatBytes(file.originalSize)} → ${formatBytes(file.processedSize)} (${sizeChangeText})`)

        // Upload back to Drive (to Normalized folder)
        file.status = 'uploading'
        file.statusText = 'Uploading to Drive...'
        addLog('info', `Uploading PDF to Drive: ${file.name}`)
        const uploadResult = await uploadToDrive(file.blob, file.name, normalizedFolderId)
        addLog('success', `Uploaded PDF to Drive: ${file.name}`, `File ID: ${uploadResult.id}`)

        // If text extraction was enabled and text was extracted, upload both TXT and XLSX files
        if (options.value.ocr && extractedText && extractedText.trim()) {
          const baseFilename = file.name.replace('.pdf', '').replace('.PDF', '')

          // Upload TXT file to Texts subfolder
          const textFilename = `${baseFilename}.txt`
          const textBlob = new Blob([extractedText], { type: 'text/plain;charset=utf-8' })
          addLog('info', `Uploading text file: ${textFilename}`, `${extractedText.length} characters`)
          try {
            const textUploadResult = await uploadToDrive(textBlob, textFilename, textsFolderId, 'text/plain')
            addLog('success', `Uploaded text file to Texts folder: ${textFilename}`, `File ID: ${textUploadResult.id}`)
          } catch (textErr) {
            addLog('warning', `Failed to upload text file for ${file.name}`, textErr.message)
          }

          // Upload XLSX file to Sheets subfolder
          const xlsxFilename = `${baseFilename}.xlsx`
          try {
            const xlsxBlob = createExcelFromText(extractedText, file.name)
            addLog('info', `Uploading Excel file: ${xlsxFilename}`)
            const xlsxUploadResult = await uploadToDrive(xlsxBlob, xlsxFilename, sheetsFolderId, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            addLog('success', `Uploaded Excel file to Sheets folder: ${xlsxFilename}`, `File ID: ${xlsxUploadResult.id}`)
          } catch (xlsxErr) {
            addLog('warning', `Failed to upload Excel file for ${file.name}`, xlsxErr.message)
          }

          // Append invoice data to master Google Sheet
          if (masterSheetId) {
            try {
              const invoiceData = parseInvoiceText(extractedText, file.name)

              // Log extracted data for debugging
              addLog('info', `Parsed invoice data for ${file.name}:`,
                `Invoice #: ${invoiceData?.invoiceNumber || '(empty)'}, ` +
                `Date: ${invoiceData?.date || '(empty)'}, ` +
                `Buyer: ${invoiceData?.buyerName || '(empty)'}, ` +
                `Items: ${invoiceData?.items?.length || 0}`)

              if (invoiceData && invoiceData.items && invoiceData.items.length > 0) {
                addLog('info', `Appending ${invoiceData.items.length} row(s) to master sheet for ${file.name}`)
                const appendResult = await appendToMasterSheet(masterSheetId, invoiceData)
                addLog('success', `Appended ${appendResult.updatedRows} row(s) to master Invoice Data sheet`)
              } else {
                addLog('info', `No invoice items found in ${file.name}, skipping master sheet append`)
              }
            } catch (appendErr) {
              addLog('warning', `Failed to append data to master sheet for ${file.name}`, appendErr.message)
            }
          }
        }

        file.status = 'done'
        file.statusText = `Complete (${formatBytes(file.originalSize)} → ${formatBytes(file.processedSize)})`
        processedFiles.value.push(file)

      } catch (err) {
        file.status = 'error'
        file.statusText = err.message || 'Processing failed'
        addLog('error', `Failed to process: ${file.name}`, err.message || 'Unknown error')
        console.error(`Error processing ${file.name}:`, err)
      }

      progress.value = ((i + 1) / files.value.length) * 100
    }

    const successCount = files.value.filter(f => f.status === 'done').length
    const errorCount = files.value.filter(f => f.status === 'error').length

    // Process Excel template if master sheet exists and has data
    if (masterSheetId && successCount > 0) {
      try {
        addLog('info', 'Starting Excel template processing...')
        const excelResult = await processExcelTemplate(folderId, masterSheetId)
        if (excelResult) {
          addLog('success', 'Excel template processing complete')
        }
      } catch (excelErr) {
        addLog('warning', 'Failed to process Excel template', excelErr.message)
        console.error('Excel template processing error:', excelErr)
      }
    }

    if (successCount > 0) {
      successMessage.value = `Successfully processed ${successCount} files and uploaded to Google Drive!`
      addLog('success', `Processing complete: ${successCount} succeeded, ${errorCount} failed`)
    } else {
      addLog('warning', 'No files were successfully processed')
    }

  } catch (err) {
    error.value = err.message || 'An error occurred while processing files'
    addLog('error', 'Processing failed', err.message || 'Unknown error')
  } finally {
    isProcessing.value = false
  }
}

async function fetchDriveFiles(folderId, normalizedFolderId = null) {
  // Build query to get PDFs from the main folder only (not from Normalized subfolder)
  // Use spaces in query (they get encoded properly), not + signs
  const query = `'${folderId}' in parents and mimeType='application/pdf' and trashed=false`

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,size,parents)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    const errorMessage = errorData.error?.message || `API Error ${response.status}`
    throw new Error(`Failed to fetch files: ${errorMessage}`)
  }

  const data = await response.json()
  const allFiles = data.files || []

  // Filter out temp files, already normalized files, and files without .pdf extension
  // Also de-duplicate by file ID
  const seenIds = new Set()
  const seenNames = new Set()

  return allFiles.filter(file => {
    // Skip if we've already seen this file ID
    if (seenIds.has(file.id)) return false
    seenIds.add(file.id)

    // Skip if we've already seen this filename (to avoid processing duplicates)
    const nameLower = file.name.toLowerCase()
    if (seenNames.has(nameLower)) return false
    seenNames.add(nameLower)

    // Skip temp files
    if (nameLower.startsWith('temp_') || nameLower.startsWith('temp-')) return false
    // Skip already normalized files (with old prefix)
    if (nameLower.startsWith('normalized_')) return false
    // Only include files with .pdf extension
    if (!nameLower.endsWith('.pdf')) return false

    // Skip files that are in the Normalized subfolder
    if (normalizedFolderId && file.parents && file.parents.includes(normalizedFolderId)) {
      return false
    }

    return true
  })
}

// Helper function to get or create a folder by name within a parent folder
async function getOrCreateFolder(parentFolderId, folderName) {
  const folderQuery = `'${parentFolderId}' in parents and name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(folderQuery)}&fields=files(id,name,parents)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (searchResponse.ok) {
    const searchData = await searchResponse.json()
    if (searchData.files && searchData.files.length > 0) {
      const foundFolder = searchData.files[0]
      if (foundFolder.parents && foundFolder.parents.includes(parentFolderId)) {
        return { folderId: foundFolder.id, created: false }
      }
    }
  }

  // Folder doesn't exist, create it
  const createResponse = await fetch(
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken.value}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentFolderId]
      })
    }
  )

  if (!createResponse.ok) {
    const errorData = await createResponse.json().catch(() => ({}))
    throw new Error(`Failed to create ${folderName} folder: ${errorData.error?.message || 'Unknown error'}`)
  }

  const createData = await createResponse.json()
  return { folderId: createData.id, created: true }
}

// Get or create "Normalized" folder with Texts and Sheets subfolders
async function getOrCreateNormalizedFolder(parentFolderId) {
  // Get or create main Normalized folder
  const normalizedResult = await getOrCreateFolder(parentFolderId, 'Normalized')
  const normalizedFolderId = normalizedResult.folderId

  // Get or create Texts subfolder
  const textsResult = await getOrCreateFolder(normalizedFolderId, 'Texts')

  // Get or create Sheets subfolder
  const sheetsResult = await getOrCreateFolder(normalizedFolderId, 'Sheets')

  return {
    folderId: normalizedFolderId,
    textsFolderId: textsResult.folderId,
    sheetsFolderId: sheetsResult.folderId,
    created: normalizedResult.created
  }
}

// Get or create "Processed" folder structure: Processed/{companyId}/{year}/Mostanadat
async function getOrCreateProcessedStructure(parentFolderId) {
  const companyId = excelSettings.value.companyId
  const year = excelSettings.value.year

  // Get or create main Processed folder
  const processedResult = await getOrCreateFolder(parentFolderId, 'Processed')
  const processedFolderId = processedResult.folderId

  // Get or create company ID subfolder (e.g., "10260538003")
  const companyResult = await getOrCreateFolder(processedFolderId, companyId)
  const companyFolderId = companyResult.folderId

  // Get or create year subfolder (e.g., "1403")
  const yearResult = await getOrCreateFolder(companyFolderId, year)
  const yearFolderId = yearResult.folderId

  // Get or create Mostanadat subfolder
  const mostanadatResult = await getOrCreateFolder(yearFolderId, 'Mostanadat')
  const mostanadatFolderId = mostanadatResult.folderId

  return {
    processedFolderId: processedFolderId,  // Root Processed folder (for Excel file)
    companyFolderId: companyFolderId,
    yearFolderId: yearFolderId,
    mostanadatFolderId: mostanadatFolderId,
    created: processedResult.created
  }
}

// Get or create the master "Invoice Data" Google Sheet in the Sheets folder
async function getOrCreateMasterSheet(sheetsFolderId) {
  const sheetName = 'Invoice Data'

  // Search for existing Google Sheet with this name in the Sheets folder
  const query = `'${sheetsFolderId}' in parents and name='${sheetName}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (searchResponse.ok) {
    const searchData = await searchResponse.json()
    if (searchData.files && searchData.files.length > 0) {
      // Found existing master sheet
      return { spreadsheetId: searchData.files[0].id, created: false }
    }
  }

  // Create new Google Sheet
  const createResponse = await fetch(
    'https://sheets.googleapis.com/v4/spreadsheets',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken.value}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        properties: {
          title: sheetName
        },
        sheets: [{
          properties: {
            title: 'Invoice Data',
            rightToLeft: true
          }
        }]
      })
    }
  )

  if (!createResponse.ok) {
    const errorData = await createResponse.json().catch(() => ({}))
    const errorMsg = errorData.error?.message || JSON.stringify(errorData) || 'Unknown error'
    const statusCode = createResponse.status
    throw new Error(`Failed to create master sheet (HTTP ${statusCode}): ${errorMsg}`)
  }

  const sheetData = await createResponse.json()
  const spreadsheetId = sheetData.spreadsheetId

  // Move the sheet to the Sheets folder
  // First, get the current parent
  const fileResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?fields=parents`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (fileResponse.ok) {
    const fileData = await fileResponse.json()
    const previousParents = fileData.parents ? fileData.parents.join(',') : ''

    // Move to Sheets folder
    await fetch(
      `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?addParents=${sheetsFolderId}&removeParents=${previousParents}`,
      {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${accessToken.value}` }
      }
    )
  }

  // Add header row to the new sheet
  const headers = [
    ['شماره', 'تاریخ', 'نام شخص حقیقی / حقوقی', 'ردیف', 'شرح کالا یا خدمات', 'مبلغ کل پس از تخفیف', 'جمع مالیات و عوارض']
  ]

  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A1:G1?valueInputOption=RAW`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${accessToken.value}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ values: headers })
    }
  )

  return { spreadsheetId, created: true }
}

// Append invoice data rows to the master Google Sheet
async function appendToMasterSheet(spreadsheetId, invoiceData) {
  if (!invoiceData || !invoiceData.items || invoiceData.items.length === 0) {
    return { updatedRows: 0 }
  }

  // Prepare rows to append
  // Columns: شماره | تاریخ | نام شخص حقیقی / حقوقی | ردیف | شرح کالا یا خدمات | مبلغ کل پس از تخفیف | جمع مالیات و عوارض
  const rows = invoiceData.items.map(item => [
    invoiceData.invoiceNumber || '',
    invoiceData.date || '',
    invoiceData.buyerName || '',
    item.rowNumber || '',
    item.description || '',
    item.totalAfterDiscount || '',
    item.taxAndDuties || ''
  ])

  // Append to the sheet
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A:G:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken.value}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ values: rows })
    }
  )

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to append to master sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  const result = await response.json()
  return { updatedRows: rows.length, updates: result.updates }
}

// Find Excel template file in the root folder
async function findExcelTemplateFile(folderId, filename) {
  const xlsxFilename = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`
  const query = `'${folderId}' in parents and name='${xlsxFilename}' and trashed=false`

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,mimeType)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (!response.ok) {
    return null
  }

  const data = await response.json()
  if (data.files && data.files.length > 0) {
    return data.files[0]
  }
  return null
}

// Read all data from master Google Sheet
async function readMasterSheetData(spreadsheetId) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/A:G`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}))
    throw new Error(`Failed to read master sheet: ${errorData.error?.message || 'Unknown error'}`)
  }

  const data = await response.json()
  return data.values || []
}

// Helper function to set cell value while preserving existing cell properties
function setCellValue(worksheet, cellAddress, value) {
  const existingCell = worksheet[cellAddress]
  if (existingCell) {
    // Preserve existing cell properties (formatting, style, etc.) and only update the value
    existingCell.v = value
    // Update type based on value
    if (typeof value === 'number') {
      existingCell.t = 'n'
    } else {
      existingCell.t = 's'
    }
    // Clear any cached formatted value
    delete existingCell.w
  } else {
    // Create new cell with just value and type (no other modifications)
    worksheet[cellAddress] = { t: 's', v: value }
  }
}

// Populate Excel template with mapped data from master sheet
// IMPORTANT: This function only fills in the mapped data cells without modifying anything else
function populateExcelTemplate(xlsxArrayBuffer, masterSheetData, targetSheetName) {
  // Read the Excel file with all options to preserve formatting
  const workbook = XLSX.read(xlsxArrayBuffer, {
    type: 'array',
    cellStyles: true,    // Preserve cell styles
    cellFormula: true,   // Preserve formulas
    cellDates: true,     // Preserve date formatting
    cellNF: true,        // Preserve number formats
    sheetStubs: true     // Include empty cells to preserve structure
  })

  // Find the target sheet
  if (!workbook.SheetNames.includes(targetSheetName)) {
    throw new Error(`Sheet "${targetSheetName}" not found in Excel file. Available sheets: ${workbook.SheetNames.join(', ')}`)
  }

  const worksheet = workbook.Sheets[targetSheetName]

  // Skip header row (row 1 in masterSheetData is headers)
  // Master sheet columns: A=شماره, B=تاریخ, C=نام شخص, D=ردیف, E=شرح کالا, F=مبلغ کل پس از تخفیف, G=جمع مالیات
  // Excel columns (data starts at row 8):
  // B=فایل (skip), C=شناسه, D=شرکت, E=سال, F=ردیف, G=عنوان, H=شماره, I=تاریخ, J=مبلغ بدون مالیات, K=مبلغ مالیات, L=کارفرما/خریدار

  const dataRows = masterSheetData.slice(1) // Skip header row
  const startRow = 8 // Excel data starts at row 8

  dataRows.forEach((row, index) => {
    const excelRow = startRow + index
    const [invoiceNumber, date, buyerName, rowNum, description, totalAfterDiscount, taxAndDuties] = row

    // Only set cell values - preserve all existing formatting and structure
    // Column C (شناسه) - Company ID from settings
    setCellValue(worksheet, `C${excelRow}`, excelSettings.value.companyId)
    // Column D (شرکت) - Company Name from settings
    setCellValue(worksheet, `D${excelRow}`, excelSettings.value.companyName)
    // Column E (سال) - Year from settings
    setCellValue(worksheet, `E${excelRow}`, excelSettings.value.year)
    // Column F (ردیف) - Row number from master sheet
    setCellValue(worksheet, `F${excelRow}`, rowNum || '')
    // Column G (عنوان) - Description from master sheet
    setCellValue(worksheet, `G${excelRow}`, description || '')
    // Column H (شماره) - Invoice number from master sheet
    setCellValue(worksheet, `H${excelRow}`, invoiceNumber || '')
    // Column I (تاریخ) - Date from master sheet
    setCellValue(worksheet, `I${excelRow}`, date || '')
    // Column J (مبلغ بدون مالیات) - Total after discount from master sheet
    setCellValue(worksheet, `J${excelRow}`, totalAfterDiscount || '')
    // Column K (مبلغ مالیات) - Tax from master sheet
    setCellValue(worksheet, `K${excelRow}`, taxAndDuties || '')
    // Column L (کارفرما/خریدار) - Buyer name from master sheet
    setCellValue(worksheet, `L${excelRow}`, buyerName || '')
  })

  // DO NOT modify worksheet['!ref'] - preserve the original structure exactly

  // Write back to buffer, preserving all properties
  const xlsxBuffer = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
    cellStyles: true  // Preserve styles in output
  })
  return new Blob([xlsxBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

// Process Excel template: find, download, populate, and upload to Processed folder
async function processExcelTemplate(sourceFolderId, masterSheetId) {
  const filename = excelSettings.value.filename
  const sheetName = excelSettings.value.sheetName

  // Find the Excel template file in the source folder
  addLog('info', `Looking for Excel template: ${filename}.xlsx`)
  const templateFile = await findExcelTemplateFile(sourceFolderId, filename)

  if (!templateFile) {
    addLog('warning', `Excel template "${filename}.xlsx" not found in source folder`, 'Skipping Excel mapping')
    return null
  }

  addLog('success', `Found Excel template: ${templateFile.name}`, `File ID: ${templateFile.id}`)

  // Download the Excel template
  addLog('info', 'Downloading Excel template...')
  const xlsxArrayBuffer = await downloadFile(templateFile.id)
  addLog('success', 'Downloaded Excel template', `Size: ${formatBytes(xlsxArrayBuffer.byteLength)}`)

  // Read master sheet data
  addLog('info', 'Reading data from master Invoice Data sheet...')
  const masterSheetData = await readMasterSheetData(masterSheetId)
  const dataRowCount = masterSheetData.length - 1 // Exclude header
  addLog('success', `Read ${dataRowCount} data row(s) from master sheet`)

  if (dataRowCount === 0) {
    addLog('warning', 'No data in master sheet to map to Excel template')
    return null
  }

  // Populate the Excel template
  addLog('info', `Populating Excel template sheet "${sheetName}" with ${dataRowCount} rows...`)
  const populatedExcelBlob = populateExcelTemplate(xlsxArrayBuffer, masterSheetData, sheetName)
  addLog('success', 'Excel template populated successfully')

  // Create the Processed folder structure: Processed/{companyId}/{year}/Mostanadat
  addLog('info', 'Creating Processed folder structure...')
  const processedStructure = await getOrCreateProcessedStructure(sourceFolderId)
  addLog('success', `Created Processed folder structure: Processed/${excelSettings.value.companyId}/${excelSettings.value.year}/Mostanadat`)

  // Upload the populated Excel directly to Processed folder (not in subfolders)
  const outputFilename = `${filename}.xlsx`
  addLog('info', `Uploading populated Excel to Processed folder: ${outputFilename}`)
  const uploadResult = await uploadToDrive(populatedExcelBlob, outputFilename, processedStructure.processedFolderId, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  addLog('success', `Uploaded populated Excel: ${outputFilename}`, `File ID: ${uploadResult.id}`)

  return uploadResult
}

async function downloadFile(fileId) {
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (!response.ok) {
    throw new Error('Failed to download file')
  }

  return await response.arrayBuffer()
}

async function processPdf(pdfBytes, filename, onOcrProgress = null) {
  let extractedText = ''

  try {
    // Load the PDF with pdf-lib for manipulation
    const pdfDoc = await PDFDocument.load(pdfBytes, {
      ignoreEncryption: true,
      updateMetadata: false
    })

    const pages = pdfDoc.getPages()

    // Text extraction: Try direct extraction first, fall back to OCR if needed
    if (options.value.ocr) {
      try {
        addLog('info', `Extracting text from ${filename}...`, `${pages.length} page(s)`)

        // Load PDF with pdf.js
        const loadingTask = pdfjsLib.getDocument({ data: pdfBytes })
        const pdfJsDoc = await loadingTask.promise

        const textParts = []
        let totalDirectTextLength = 0

        // STEP 1: Try direct text extraction first (for text-based PDFs)
        addLog('info', 'Attempting direct text extraction...', filename)

        for (let i = 1; i <= pdfJsDoc.numPages; i++) {
          if (onOcrProgress) {
            onOcrProgress(`Extracting text from page ${i}/${pdfJsDoc.numPages}`)
          }

          const page = await pdfJsDoc.getPage(i)
          const textContent = await page.getTextContent()

          // Extract text from the text content items
          const pageText = textContent.items
            .map(item => item.str)
            .join(' ')
            .trim()

          if (pageText) {
            textParts.push(`--- Page ${i} ---\n${pageText}`)
            totalDirectTextLength += pageText.length
          }
        }

        // Check if we got meaningful text (at least 50 chars per page on average)
        const avgCharsPerPage = totalDirectTextLength / pdfJsDoc.numPages
        const hasDirectText = avgCharsPerPage > 50

        if (hasDirectText) {
          // Direct extraction worked! No OCR needed
          extractedText = normalizeArabicText(textParts.join('\n\n'))
          addLog('success', `Direct text extraction complete for ${filename}`, `Extracted ${extractedText.length} characters (no OCR needed)`)
        } else {
          // STEP 2: Fall back to OCR for scanned/image-based PDFs
          addLog('info', `Little/no text found directly. Running OCR on ${filename}...`, `Found only ${totalDirectTextLength} chars`)

          textParts.length = 0 // Clear the array for OCR results

          for (let i = 1; i <= pdfJsDoc.numPages; i++) {
            if (onOcrProgress) {
              onOcrProgress(`OCR page ${i}/${pdfJsDoc.numPages}`)
            }
            addLog('info', `OCR processing page ${i}/${pdfJsDoc.numPages}`, filename)

            const page = await pdfJsDoc.getPage(i)
            const viewport = page.getViewport({ scale: 2.0 }) // Higher scale = better OCR quality

            // Create canvas to render PDF page
            const canvas = document.createElement('canvas')
            const context = canvas.getContext('2d')
            canvas.height = viewport.height
            canvas.width = viewport.width

            await page.render({
              canvasContext: context,
              viewport: viewport
            }).promise

            // Convert canvas to image data URL
            const imageDataUrl = canvas.toDataURL('image/png')

            // Run Tesseract OCR on the image
            // Support multiple languages: English + Persian (Farsi) + Arabic
            const result = await Tesseract.recognize(
              imageDataUrl,
              'eng+fas+ara', // English + Persian (Farsi) + Arabic
              {
                logger: m => {
                  if (m.status === 'recognizing text' && m.progress) {
                    // Could add progress updates here
                  }
                }
              }
            )

            if (result.data.text.trim()) {
              textParts.push(`--- Page ${i} ---\n${result.data.text.trim()}`)
            }
          }

          extractedText = normalizeArabicText(textParts.join('\n\n'))

          if (extractedText) {
            addLog('success', `OCR complete for ${filename}`, `Extracted ${extractedText.length} characters`)
          } else {
            addLog('warning', `OCR complete but no text found in ${filename}`)
          }
        }

      } catch (extractErr) {
        console.error('Text extraction error:', extractErr)
        addLog('warning', `Text extraction failed for ${filename}`, extractErr.message)
        // Continue without extracted text
      }
    }

    // Standardize page sizes if enabled
    if (options.value.standardize && options.value.pageSize !== 'ORIGINAL') {
      const targetSize = getPageDimensions(options.value.pageSize)

      for (const page of pages) {
        const { width, height } = page.getSize()

        // Only resize if significantly different
        if (Math.abs(width - targetSize.width) > 1 || Math.abs(height - targetSize.height) > 1) {
          // Scale content to fit new size
          const scaleX = targetSize.width / width
          const scaleY = targetSize.height / height
          const scale = Math.min(scaleX, scaleY)

          page.setSize(targetSize.width, targetSize.height)
          page.scaleContent(scale, scale)

          // Center content
          const offsetX = (targetSize.width - width * scale) / 2
          const offsetY = (targetSize.height - height * scale) / 2
          page.translateContent(offsetX, offsetY)
        }
      }
    }

    // Update metadata
    pdfDoc.setTitle(filename.replace('.pdf', ''))
    pdfDoc.setProducer('PDF Normalizer')
    pdfDoc.setModificationDate(new Date())

    // Save with compression if enabled
    const saveOptions = {}
    if (options.value.compress) {
      saveOptions.useObjectStreams = true
    }

    const processedPdfBytes = await pdfDoc.save(saveOptions)

    return {
      pdfBytes: processedPdfBytes,
      extractedText: extractedText
    }

  } catch (err) {
    console.error('PDF processing error:', err)
    // If processing fails, return original
    return {
      pdfBytes: pdfBytes,
      extractedText: extractedText
    }
  }
}

function getPageDimensions(size) {
  switch (size) {
    case 'LETTER':
      return { width: 612, height: 792 } // 8.5 x 11 inches
    case 'A4':
      return { width: 595.28, height: 841.89 } // 210 x 297 mm
    case 'LEGAL':
      return { width: 612, height: 1008 } // 8.5 x 14 inches
    default:
      return { width: 612, height: 792 }
  }
}

// Normalize Arabic/Persian text by converting presentation forms to base characters
function normalizeArabicText(text) {
  if (!text) return text

  // Apply NFKC normalization to convert Arabic Presentation Forms to base characters
  let normalized = text.normalize('NFKC')

  // Additional cleanup for common issues
  // Remove zero-width characters that can cause display issues
  normalized = normalized.replace(/[\u200B-\u200F\u202A-\u202E\uFEFF]/g, '')

  // Normalize multiple spaces to single space
  normalized = normalized.replace(/  +/g, ' ')

  return normalized
}

// Parse filename to extract invoice number and buyer name
// Format: "3427 - آموزشگاه زبان کوشش" or "3427 - آموزشگاه زبان کوشش.pdf"
function parseFilename(filename) {
  if (!filename) return { invoiceNumber: '', buyerName: '' }

  // Remove .pdf extension if present
  const nameWithoutExt = filename.replace(/\.pdf$/i, '')

  // Split by " - " separator
  const parts = nameWithoutExt.split(' - ')

  if (parts.length >= 2) {
    const invoiceNumber = parts[0].trim()
    const buyerName = parts.slice(1).join(' - ').trim() // Handle case where buyer name contains " - "
    return { invoiceNumber, buyerName }
  }

  return { invoiceNumber: '', buyerName: '' }
}

// Parse Persian invoice text and extract structured data
// Specialized for Rastakhiz invoice format
function parseInvoiceText(text, filename = '') {
  if (!text) return null

  // Normalize text first
  const normalizedText = normalizeArabicText(text)

  const result = {
    // Header fields (non-table)
    invoiceNumber: '',  // شماره
    date: '',           // تاریخ
    buyerName: '',      // نام شخص حقیقی / حقوقی (from buyer section)

    // Table items (can have multiple rows)
    items: [],

    // Raw text for reference
    rawText: text
  }

  // PRIMARY: Extract invoice number and buyer name from filename
  // Format: "3427 - آموزشگاه زبان کوشش.pdf"
  const filenameData = parseFilename(filename)
  if (filenameData.invoiceNumber) {
    result.invoiceNumber = filenameData.invoiceNumber
  }
  if (filenameData.buyerName) {
    result.buyerName = filenameData.buyerName
  }

  // FALLBACK: Extract invoice number from text if not found in filename
  // Pattern: "شماره: 3427" or "شماره:3427"
  if (!result.invoiceNumber) {
    const invoiceNumMatch = normalizedText.match(/شماره[:\s]*(\d+)/)
    if (invoiceNumMatch) {
      result.invoiceNumber = invoiceNumMatch[1]
    }
  }

  // Extract تاریخ (Date) from text
  // Pattern: "1403/04/02 :تاریخ" (RTL format) or "تاریخ: 1403/04/02"
  const dateMatch = normalizedText.match(/(\d{4}\/\d{2}\/\d{2})\s*:?\s*تاریخ|تاریخ\s*:?\s*(\d{4}\/\d{2}\/\d{2})/)
  if (dateMatch) {
    result.date = dateMatch[1] || dateMatch[2]
  }

  // FALLBACK: Extract buyer name from text if not found in filename
  // Find the SECOND instance of "نام شخص حقیقی / حقوقی" (first is seller, second is buyer)
  if (!result.buyerName) {
    const label = 'نام شخص حقیقی / حقوقی'
    const firstIndex = normalizedText.indexOf(label)

    if (firstIndex !== -1) {
      // Find second instance starting after the first one
      const secondIndex = normalizedText.indexOf(label, firstIndex + label.length)

      if (secondIndex !== -1) {
        // Get text after the second instance of the label
        const afterLabel = normalizedText.substring(secondIndex + label.length)
        // Extract value until newline or colon (next field)
        const match = afterLabel.match(/^[:\s]*([^\n:]+)/)
        if (match && match[1]) {
          const value = match[1].trim()
          // Make sure it's not another field label
          if (value.length > 2 && !value.includes('شناسه') && !value.includes('کد')) {
            result.buyerName = value
          }
        }
      }
    }
  }

  // 4. Parse table rows
  // The table has these columns (RTL order in headers, but data appears LTR in extracted text):
  // Headers: ردیف | شرح کالا یا خدمات | تعداد | مبلغ واحد | مبلغ کل | مبلغ تخفیف | مبلغ کل پس از تخفیف | جمع مالیات و عوارض | جمع مبلغ کل بعد از مالیات
  //
  // In extracted text, after "ردیف" header, data appears as:
  // grandTotal tax totalAfterDiscount discount total unitPrice quantity description rowNumber
  // Example: 11,550,000 1,050,000 10,500,000 0 10,500,000 150,000 70 نرم افزار کلاس آنلاین 1

  // Find the position after "ردیف" header to start looking for data
  const headerEndMatch = normalizedText.match(/تعداد\s+شرح کالا یا خدمات\s+ردیف/)
  if (headerEndMatch) {
    const dataStartIndex = normalizedText.indexOf(headerEndMatch[0]) + headerEndMatch[0].length

    // Get the text after headers until "جمع کل" (totals row)
    let dataSection = normalizedText.substring(dataStartIndex)
    const totalsIndex = dataSection.indexOf('جمع کل')
    if (totalsIndex !== -1) {
      dataSection = dataSection.substring(0, totalsIndex)
    }

    // Pattern to match each row:
    // Numbers (8 of them) followed by Persian text (description) followed by row number
    // grandTotal tax totalAfterDiscount discount total unitPrice quantity description rowNum
    const rowPattern = /([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+(\d+)\s+([^\d]+?)\s+(\d+)(?=\s|$)/g

    let match
    while ((match = rowPattern.exec(dataSection)) !== null) {
      const description = match[8].trim()
      // Skip if this looks like a header or total row
      if (description.includes('شرح') || description.includes('ردیف') || description.includes('جمع')) {
        continue
      }

      result.items.push({
        rowNumber: match[9],                    // ردیف
        description: description,               // شرح کالا یا خدمات
        quantity: match[7],                     // تعداد
        unitPrice: match[6],                    // مبلغ واحد
        total: match[5],                        // مبلغ کل
        discount: match[4],                     // مبلغ تخفیف
        totalAfterDiscount: match[3],           // مبلغ کل پس از تخفیف
        taxAndDuties: match[2],                 // جمع مالیات و عوارض
        grandTotal: match[1]                    // جمع مبلغ کل بعد از مالیات و عوارض
      })
    }
  }

  return result
}

function createExcelFromText(extractedText, originalFilename) {
  // Create workbook
  const wb = XLSX.utils.book_new()

  // Try to parse as invoice first
  const invoiceData = parseInvoiceText(extractedText, originalFilename)

  // Sheet 1: Invoice Data (main sheet with all essential fields)
  // Columns: شماره | تاریخ | نام شخص حقیقی / حقوقی | ردیف | شرح کالا یا خدمات | مبلغ کل پس از تخفیف | جمع مالیات و عوارض
  const invoiceSheetData = [
    // Header row with exact column names
    [
      'شماره',
      'تاریخ',
      'نام شخص حقیقی / حقوقی',
      'ردیف',
      'شرح کالا یا خدمات',
      'مبلغ کل پس از تخفیف',
      'جمع مالیات و عوارض'
    ]
  ]

  // Add data rows - one row per item in the invoice
  if (invoiceData?.items && invoiceData.items.length > 0) {
    invoiceData.items.forEach((item) => {
      invoiceSheetData.push([
        invoiceData.invoiceNumber || '',        // شماره
        invoiceData.date || '',                 // تاریخ
        invoiceData.buyerName || '',            // نام شخص حقیقی / حقوقی
        item.rowNumber || '',                   // ردیف
        item.description || '',                 // شرح کالا یا خدمات
        item.totalAfterDiscount || '',          // مبلغ کل پس از تخفیف
        item.taxAndDuties || ''                 // جمع مالیات و عوارض
      ])
    })
  } else {
    // If no items parsed, add a row with just header info
    invoiceSheetData.push([
      invoiceData?.invoiceNumber || '',
      invoiceData?.date || '',
      invoiceData?.buyerName || '',
      '',
      '',
      '',
      ''
    ])
  }

  const wsInvoice = XLSX.utils.aoa_to_sheet(invoiceSheetData)
  wsInvoice['!cols'] = [
    { wch: 12 },  // شماره
    { wch: 12 },  // تاریخ
    { wch: 40 },  // نام شخص حقیقی / حقوقی
    { wch: 8 },   // ردیف
    { wch: 35 },  // شرح کالا یا خدمات
    { wch: 20 },  // مبلغ کل پس از تخفیف
    { wch: 20 },  // جمع مالیات و عوارض
  ]
  wsInvoice['!dir'] = 'rtl'
  XLSX.utils.book_append_sheet(wb, wsInvoice, 'Invoice Data')

  // Sheet 2: All Item Details (full table data for reference)
  if (invoiceData?.items && invoiceData.items.length > 0) {
    const fullItemsData = [
      [
        'ردیف',
        'شرح کالا یا خدمات',
        'تعداد',
        'مبلغ واحد',
        'مبلغ کل',
        'مبلغ تخفیف',
        'مبلغ کل پس از تخفیف',
        'جمع مالیات و عوارض',
        'جمع مبلغ کل بعد از مالیات'
      ]
    ]

    invoiceData.items.forEach((item) => {
      fullItemsData.push([
        item.rowNumber || '',
        item.description || '',
        item.quantity || '',
        item.unitPrice || '',
        item.total || '',
        item.discount || '',
        item.totalAfterDiscount || '',
        item.taxAndDuties || '',
        item.grandTotal || ''
      ])
    })

    const wsFullItems = XLSX.utils.aoa_to_sheet(fullItemsData)
    wsFullItems['!cols'] = [
      { wch: 8 },   // ردیف
      { wch: 35 },  // شرح کالا
      { wch: 10 },  // تعداد
      { wch: 15 },  // مبلغ واحد
      { wch: 15 },  // مبلغ کل
      { wch: 15 },  // مبلغ تخفیف
      { wch: 20 },  // مبلغ کل پس از تخفیف
      { wch: 20 },  // جمع مالیات
      { wch: 22 },  // جمع مبلغ کل بعد از مالیات
    ]
    wsFullItems['!dir'] = 'rtl'
    XLSX.utils.book_append_sheet(wb, wsFullItems, 'Full Item Details')
  }

  // Sheet 3: Raw Text (for reference)
  const pages = extractedText.split(/---\s*Page\s+\d+\s*---/i).filter(p => p.trim())
  const rawTextData = [
    ['Source File', 'Page', 'Extracted Text'],
  ]

  pages.forEach((pageText, index) => {
    rawTextData.push([originalFilename, index + 1, pageText.trim()])
  })

  // If no pages were parsed, just put all text in one row
  if (pages.length === 0 && extractedText.trim()) {
    rawTextData.push([originalFilename, 1, extractedText.trim()])
  }

  const wsRaw = XLSX.utils.aoa_to_sheet(rawTextData)
  wsRaw['!cols'] = [
    { wch: 40 },  // Source File
    { wch: 8 },   // Page
    { wch: 100 }, // Text
  ]
  XLSX.utils.book_append_sheet(wb, wsRaw, 'Raw Text')

  // Generate Excel file as array buffer
  const xlsxBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })

  return new Blob([xlsxBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

async function uploadToDrive(blob, filename, parentFolderId, mimeType = null) {
  // Use blob's type if mimeType not provided, fallback to application/pdf
  const fileMimeType = mimeType || blob.type || 'application/pdf'

  // First, check if file already exists in the folder
  const escapedFilename = filename.replace(/'/g, "\\'")
  // Use spaces in query (they get encoded properly by encodeURIComponent), not + signs
  // Also add trashed=false to avoid finding trashed files
  const searchQuery = `'${parentFolderId}' in parents and name='${escapedFilename}' and trashed=false`

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(searchQuery)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  let existingFileId = null
  if (searchResponse.ok) {
    const searchData = await searchResponse.json()
    if (searchData.files && searchData.files.length > 0) {
      existingFileId = searchData.files[0].id
    }
  }

  if (existingFileId) {
    // Update existing file
    const form = new FormData()
    form.append('metadata', new Blob([JSON.stringify({ mimeType: fileMimeType })], { type: 'application/json' }))
    form.append('file', blob)

    const response = await fetch(
      `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=multipart&supportsAllDrives=true`,
      {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${accessToken.value}` },
        body: form
      }
    )

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}))
      throw new Error(errorData.error?.message || 'Failed to update file on Drive')
    }

    return await response.json()
  } else {
    // Create new file
    const metadata = {
      name: filename,
      parents: [parentFolderId],
      mimeType: fileMimeType
    }

    const form = new FormData()
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }))
    form.append('file', blob)

    const response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken.value}` },
        body: form
      }
    )

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}))
      throw new Error(errorData.error?.message || 'Failed to upload file to Drive')
    }

    return await response.json()
  }
}

async function downloadAll() {
  if (processedFiles.value.length === 0) return

  const zip = new JSZip()

  for (const file of processedFiles.value) {
    if (file.blob) {
      zip.file(file.name, file.blob)
    }
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(zipBlob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'normalized_pdfs.zip'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

// Helper to clear all auth data
function clearAuthData() {
  accessToken.value = ''
  isAuthenticated.value = false
  userEmail.value = ''
  localStorage.removeItem('google_access_token')
  localStorage.removeItem('google_token_expiry')
  localStorage.removeItem('google_user_email')
}

// Initialize - restore auth from localStorage if valid
onMounted(async () => {
  const savedToken = localStorage.getItem('google_access_token')
  const tokenExpiry = localStorage.getItem('google_token_expiry')
  const savedEmail = localStorage.getItem('google_user_email')

  if (!savedToken) return

  // Check if token has expired locally
  if (tokenExpiry && Date.now() > parseInt(tokenExpiry)) {
    clearAuthData()
    return
  }

  // Verify token with Google
  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${savedToken}` }
    })

    if (response.ok) {
      const data = await response.json()
      accessToken.value = savedToken
      isAuthenticated.value = true
      userEmail.value = data.email || savedEmail || ''

      // Update saved email if we got a fresh one
      if (data.email && data.email !== savedEmail) {
        localStorage.setItem('google_user_email', data.email)
      }
    } else {
      clearAuthData()
    }
  } catch (err) {
    console.error('Token validation error:', err)
    // On network error, use cached auth data if available
    if (savedEmail) {
      userEmail.value = savedEmail
      accessToken.value = savedToken
      isAuthenticated.value = true
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
