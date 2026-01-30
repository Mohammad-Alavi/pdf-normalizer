<template>
  <v-card class="mx-auto" max-width="900">
    <v-card-title class="text-h5 pa-6 pb-2">
      <v-icon color="primary" class="mr-2">mdi-google-drive</v-icon>
      Google Drive PDF Processor
    </v-card-title>

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
          <span>Processed files uploaded to: </span>
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
import { PDFDocument, StandardFonts } from 'pdf-lib'
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

// Options
const options = ref({
  ocr: true,
  standardize: true,
  compress: true,
  pageSize: 'ORIGINAL'
})

const pageSizes = [
  { title: 'Letter (8.5" x 11")', value: 'LETTER' },
  { title: 'A4 (210mm x 297mm)', value: 'A4' },
  { title: 'Legal (8.5" x 14")', value: 'LEGAL' },
  { title: 'Keep Original', value: 'ORIGINAL' }
]

// Google OAuth Config
const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || ''
const API_KEY = import.meta.env.VITE_GOOGLE_API_KEY || ''
const SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile'

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
        console.log('OAuth callback received, token expires_in:', response.expires_in)

        accessToken.value = response.access_token
        // Save token to localStorage for persistence across refreshes
        localStorage.setItem('google_access_token', response.access_token)
        // Also save the expiration time (Google tokens expire in ~1 hour)
        const expiryTime = Date.now() + (response.expires_in * 1000)
        localStorage.setItem('google_token_expiry', String(expiryTime))
        console.log('Token saved, expires at:', new Date(expiryTime).toISOString())

        isAuthenticated.value = true

        // Fetch and save user info - await to ensure it completes
        try {
          await fetchUserInfo()
          console.log('User email fetched:', userEmail.value)
          if (userEmail.value) {
            localStorage.setItem('google_user_email', userEmail.value)
            console.log('Email saved to localStorage:', userEmail.value)
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
    console.log('=== Fetching user info ===')
    console.log('Using access token (first 20 chars):', accessToken.value.substring(0, 20) + '...')

    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    })

    console.log('User info response status:', response.status)

    if (!response.ok) {
      const errorText = await response.text()
      console.error('Failed to fetch user info, status:', response.status, 'error:', errorText)
      return null
    }

    const data = await response.json()
    console.log('User info response data:', JSON.stringify(data))

    const email = data.email || ''
    userEmail.value = email
    console.log('Set userEmail to:', email)

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

    // Get or create the "Normalized" subfolder
    addLog('info', 'Checking for Normalized folder...')
    const { folderId: normalizedFolderId, textsFolderId, sheetsFolderId, created } = await getOrCreateNormalizedFolder(folderId)

    if (created) {
      addLog('success', 'Created new Normalized folder structure', `Main: ${normalizedFolderId}, Texts: ${textsFolderId}, Sheets: ${sheetsFolderId}`)
    } else {
      addLog('success', 'Found existing Normalized folder structure', `Main: ${normalizedFolderId}, Texts: ${textsFolderId}, Sheets: ${sheetsFolderId}`)
    }

    // Set the link to open the folder
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

// Parse Persian invoice text and extract structured data
// Specialized for Rastakhiz invoice format
function parseInvoiceText(text) {
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

  // 1. Extract شماره (Invoice Number)
  // Pattern: "شماره: 3427" or "شماره:3427"
  const invoiceNumMatch = normalizedText.match(/شماره[:\s]*(\d+)/)
  if (invoiceNumMatch) {
    result.invoiceNumber = invoiceNumMatch[1]
  }

  // 2. Extract تاریخ (Date)
  // Pattern: "1403/04/02 :تاریخ" (RTL format) or "تاریخ: 1403/04/02"
  const dateMatch = normalizedText.match(/(\d{4}\/\d{2}\/\d{2})\s*:?\s*تاریخ|تاریخ\s*:?\s*(\d{4}\/\d{2}\/\d{2})/)
  if (dateMatch) {
    result.date = dateMatch[1] || dateMatch[2]
  }

  // 3. Extract buyer's نام شخص حقیقی / حقوقی
  // This appears in the "مشخصات خریدار" section (buyer details)
  // We need to find the نام شخص حقیقی / حقوقی that comes AFTER مشخصات خریدار
  const buyerSectionStart = normalizedText.indexOf('مشخصات خریدار')
  if (buyerSectionStart !== -1) {
    const buyerSection = normalizedText.substring(buyerSectionStart)
    // Find نام شخص حقیقی / حقوقی: value in the buyer section
    const buyerNameMatch = buyerSection.match(/نام شخص حقیقی \/ حقوقی[:\s]*([^]+?)(?:فکس|تلفن|کد پستی|نشانی کامل|$)/)
    if (buyerNameMatch) {
      let buyerName = buyerNameMatch[1].trim()
      // Clean up - remove any trailing colons or separators
      buyerName = buyerName.replace(/[:\s،,]+$/, '').trim()
      // Remove any text that looks like the next field label
      buyerName = buyerName.split(/\s*(?:فکس|تلفن|کد پستی|نشانی)/)[0].trim()
      result.buyerName = buyerName
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
  const invoiceData = parseInvoiceText(extractedText)

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

  console.log('Checking for existing file with query:', searchQuery)
  console.log('File mimeType:', fileMimeType)

  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(searchQuery)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  let existingFileId = null
  if (searchResponse.ok) {
    const searchData = await searchResponse.json()
    console.log('File search response:', searchData)
    if (searchData.files && searchData.files.length > 0) {
      existingFileId = searchData.files[0].id
      console.log('Found existing file to update:', existingFileId)
    }
  }

  if (existingFileId) {
    // Update existing file
    console.log('Updating existing file:', existingFileId)
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
      console.error('Upload PATCH error:', errorData)
      throw new Error(errorData.error?.message || 'Failed to update file on Drive')
    }

    const result = await response.json()
    console.log('File updated successfully:', result)
    return result
  } else {
    // Create new file
    console.log('Creating new file in folder:', parentFolderId)
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
      console.error('Upload POST error:', errorData)
      throw new Error(errorData.error?.message || 'Failed to upload file to Drive')
    }

    const result = await response.json()
    console.log('File created successfully:', result)
    return result
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

// Initialize
onMounted(async () => {
  console.log('=== PdfProcessor mounted, checking for saved auth ===')

  // Check for existing token in localStorage
  const savedToken = localStorage.getItem('google_access_token')
  const tokenExpiry = localStorage.getItem('google_token_expiry')
  const savedEmail = localStorage.getItem('google_user_email')

  console.log('Saved token exists:', !!savedToken)
  console.log('Saved token (first 20 chars):', savedToken ? savedToken.substring(0, 20) + '...' : 'none')
  console.log('Saved email:', savedEmail)
  console.log('Token expiry:', tokenExpiry ? new Date(parseInt(tokenExpiry)).toISOString() : 'not set')
  console.log('Current time:', new Date().toISOString())
  console.log('Time until expiry:', tokenExpiry ? `${Math.round((parseInt(tokenExpiry) - Date.now()) / 1000 / 60)} minutes` : 'N/A')

  if (!savedToken) {
    console.log('No saved token found, user needs to sign in')
    return
  }

  // Check if token has expired locally first
  if (tokenExpiry && Date.now() > parseInt(tokenExpiry)) {
    console.log('Token expired locally, clearing auth data...')
    clearAuthData()
    return
  }

  // Token exists and hasn't expired locally - verify with Google
  console.log('Verifying token with Google API...')

  try {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${savedToken}` }
    })

    console.log('Google API response status:', response.status)

    if (response.ok) {
      const data = await response.json()
      console.log('Token VALID! User data:', JSON.stringify(data))

      // Token is valid - set all auth state
      accessToken.value = savedToken
      isAuthenticated.value = true

      // Use email from API response, fallback to saved email
      const email = data.email || savedEmail || ''
      userEmail.value = email
      console.log('Setting userEmail to:', email)

      // Update saved email if we got a fresh one from API
      if (data.email && data.email !== savedEmail) {
        localStorage.setItem('google_user_email', data.email)
        console.log('Updated saved email to:', data.email)
      }

      console.log('=== Auth restored successfully ===')
      console.log('isAuthenticated:', isAuthenticated.value)
      console.log('userEmail:', userEmail.value)
    } else {
      // Token is invalid
      const errorText = await response.text()
      console.log('Token INVALID! Status:', response.status, 'Error:', errorText)
      clearAuthData()
    }
  } catch (err) {
    // Network or other error
    console.error('Token validation network error:', err)
    // Don't clear auth on network errors - might be temporary
    // But still set state from localStorage for offline-ish UX
    if (savedEmail) {
      userEmail.value = savedEmail
      accessToken.value = savedToken
      isAuthenticated.value = true
      console.log('Network error but using cached auth data')
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
