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
                  label="OCR (Text Recognition)"
                  hint="Extract text from scanned PDFs"
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
            <v-btn
              variant="text"
              size="small"
              color="error"
              class="mt-2"
              @click="clearLogs"
            >
              <v-icon size="small" class="mr-1">mdi-delete</v-icon>
              Clear Logs
            </v-btn>
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
const SCOPES = 'https://www.googleapis.com/auth/drive'

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
    console.log('Fetching user info...')
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    })
    if (!response.ok) {
      console.error('Failed to fetch user info, status:', response.status)
      return
    }
    const data = await response.json()
    console.log('User info response:', data)
    userEmail.value = data.email || ''
    console.log('Set userEmail to:', userEmail.value)
  } catch (err) {
    console.error('Failed to fetch user info:', err)
  }
}

function signOut() {
  const token = accessToken.value
  accessToken.value = ''
  isAuthenticated.value = false
  userEmail.value = ''
  localStorage.removeItem('google_access_token')
  localStorage.removeItem('google_token_expiry')
  localStorage.removeItem('google_user_email')
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
  clearLogs()

  try {
    const folderId = extractFolderId(driveLink.value)
    addLog('info', 'Starting PDF processing', `Folder ID: ${folderId}`)

    // Get or create the "Normalized" subfolder
    addLog('info', 'Checking for Normalized folder...')
    const normalizedFolderId = await getOrCreateNormalizedFolder(folderId)
    addLog('success', 'Normalized folder ready', `Folder ID: ${normalizedFolderId}`)

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
        const processedBytes = await processPdf(pdfBytes, file.name)
        file.processedSize = processedBytes.byteLength
        file.blob = new Blob([processedBytes], { type: 'application/pdf' })

        const sizeChange = file.originalSize - file.processedSize
        const sizeChangeText = sizeChange > 0 ? `Reduced by ${formatBytes(sizeChange)}` : sizeChange < 0 ? `Increased by ${formatBytes(Math.abs(sizeChange))}` : 'No size change'
        addLog('success', `Processed: ${file.name}`, `${formatBytes(file.originalSize)} → ${formatBytes(file.processedSize)} (${sizeChangeText})`)

        // Upload back to Drive (to Normalized folder)
        file.status = 'uploading'
        file.statusText = 'Uploading to Drive...'
        addLog('info', `Uploading to Drive: ${file.name}`)
        await uploadToDrive(file.blob, file.name, normalizedFolderId)
        addLog('success', `Uploaded to Drive: ${file.name}`)

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
  const query = `'${folderId}' in parents and mimeType='application/pdf'`

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

async function getOrCreateNormalizedFolder(parentFolderId) {
  // First, check if "Normalized" folder already exists
  const folderQuery = `'${parentFolderId}' in parents and name='Normalized' and mimeType='application/vnd.google-apps.folder'`
  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(folderQuery)}&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    {
      headers: { Authorization: `Bearer ${accessToken.value}` }
    }
  )

  if (searchResponse.ok) {
    const searchData = await searchResponse.json()
    if (searchData.files && searchData.files.length > 0) {
      // Folder exists, return its ID
      return searchData.files[0].id
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
        name: 'Normalized',
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentFolderId]
      })
    }
  )

  if (!createResponse.ok) {
    const errorData = await createResponse.json().catch(() => ({}))
    throw new Error(`Failed to create Normalized folder: ${errorData.error?.message || 'Unknown error'}`)
  }

  const createData = await createResponse.json()
  return createData.id
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

async function processPdf(pdfBytes, filename) {
  try {
    // Load the PDF
    const pdfDoc = await PDFDocument.load(pdfBytes, {
      ignoreEncryption: true,
      updateMetadata: false
    })

    const pages = pdfDoc.getPages()

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

    return await pdfDoc.save(saveOptions)

  } catch (err) {
    console.error('PDF processing error:', err)
    // If processing fails, return original
    return pdfBytes
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

async function uploadToDrive(blob, filename, parentFolderId) {
  // First, check if file already exists in the folder
  const escapedFilename = filename.replace(/'/g, "\\'").replace(/\\/g, "\\\\")
  // Use spaces in query (they get encoded properly by encodeURIComponent), not + signs
  const searchQuery = `'${parentFolderId}' in parents and name='${escapedFilename}'`
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
    form.append('metadata', new Blob([JSON.stringify({ mimeType: 'application/pdf' })], { type: 'application/json' }))
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
    console.log('File updated successfully:', result.id)
    return result
  } else {
    // Create new file
    const metadata = {
      name: filename,
      parents: [parentFolderId],
      mimeType: 'application/pdf'
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
    console.log('File created successfully:', result.id)
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

// Initialize
onMounted(async () => {
  console.log('PdfProcessor mounted, checking for saved auth...')

  // Check for existing token in localStorage
  const savedToken = localStorage.getItem('google_access_token')
  const tokenExpiry = localStorage.getItem('google_token_expiry')
  const savedEmail = localStorage.getItem('google_user_email')

  console.log('Saved token exists:', !!savedToken)
  console.log('Saved email:', savedEmail)
  console.log('Token expiry:', tokenExpiry ? new Date(parseInt(tokenExpiry)).toISOString() : 'not set')

  if (savedToken) {
    // Check if token has expired locally first
    if (tokenExpiry && Date.now() > parseInt(tokenExpiry)) {
      console.log('Token expired locally, clearing...')
      // Token expired, clear it
      localStorage.removeItem('google_access_token')
      localStorage.removeItem('google_token_expiry')
      localStorage.removeItem('google_user_email')
      return
    }

    // Set the saved email immediately for better UX
    if (savedEmail) {
      userEmail.value = savedEmail
      console.log('Set email from localStorage:', savedEmail)
    }

    // Verify token is still valid with Google
    try {
      console.log('Verifying token with Google...')
      const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { Authorization: `Bearer ${savedToken}` }
      })
      console.log('Token verification response status:', response.status)

      if (response.ok) {
        const data = await response.json()
        console.log('Token valid, user email from API:', data.email)
        accessToken.value = savedToken
        isAuthenticated.value = true
        userEmail.value = data.email || savedEmail || ''
        // Update saved email if it changed
        if (data.email) {
          localStorage.setItem('google_user_email', data.email)
        }
      } else {
        // Token invalid, clear it
        console.log('Token invalid (response not ok), clearing...')
        localStorage.removeItem('google_access_token')
        localStorage.removeItem('google_token_expiry')
        localStorage.removeItem('google_user_email')
        userEmail.value = ''
      }
    } catch (err) {
      // Token validation failed, clear it
      console.error('Token validation error:', err)
      localStorage.removeItem('google_access_token')
      localStorage.removeItem('google_token_expiry')
      localStorage.removeItem('google_user_email')
      userEmail.value = ''
    }
  } else {
    console.log('No saved token found')
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
