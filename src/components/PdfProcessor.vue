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
      >
        <div class="d-flex align-center">
          <v-icon class="mr-2">mdi-check-circle</v-icon>
          <div class="flex-grow-1">
            Signed in as {{ userEmail }}
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
      <div class="d-flex gap-3 mb-4">
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
                  color="success"
                  variant="flat"
                >
                  {{ getSizeReduction(file.originalSize, file.processedSize) }}
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

function getSizeReduction(original, processed) {
  const reduction = ((original - processed) / original * 100).toFixed(0)
  return reduction > 0 ? `-${reduction}%` : `+${Math.abs(reduction)}%`
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
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
      callback: (response) => {
        if (response.error) {
          error.value = 'Authentication failed: ' + response.error
          isAuthenticating.value = false
          return
        }
        accessToken.value = response.access_token
        // Save token to sessionStorage for persistence
        sessionStorage.setItem('google_access_token', response.access_token)
        isAuthenticated.value = true
        fetchUserInfo()
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
    const data = await response.json()
    userEmail.value = data.email
  } catch (err) {
    console.error('Failed to fetch user info:', err)
  }
}

function signOut() {
  const token = accessToken.value
  accessToken.value = ''
  isAuthenticated.value = false
  userEmail.value = ''
  sessionStorage.removeItem('google_access_token')
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

  try {
    const folderId = extractFolderId(driveLink.value)

    // Get or create the "Normalized" subfolder
    const normalizedFolderId = await getOrCreateNormalizedFolder(folderId)

    // Fetch files from Google Drive
    const pdfFiles = await fetchDriveFiles(folderId)

    if (pdfFiles.length === 0) {
      error.value = 'No PDF files found in the specified folder'
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
      try {
        // Download
        file.status = 'downloading'
        file.statusText = 'Downloading...'
        const pdfBytes = await downloadFile(file.id)
        file.originalSize = pdfBytes.byteLength

        // Process
        file.status = 'processing'
        file.statusText = 'Processing...'
        const processedBytes = await processPdf(pdfBytes, file.name)
        file.processedSize = processedBytes.byteLength
        file.blob = new Blob([processedBytes], { type: 'application/pdf' })

        // Upload back to Drive (to Normalized folder)
        file.status = 'uploading'
        file.statusText = 'Uploading to Drive...'
        await uploadToDrive(file.blob, file.name, normalizedFolderId)

        file.status = 'done'
        file.statusText = `Complete (${formatBytes(file.originalSize)} \u2192 ${formatBytes(file.processedSize)})`
        processedFiles.value.push(file)

      } catch (err) {
        file.status = 'error'
        file.statusText = err.message || 'Processing failed'
        console.error(`Error processing ${file.name}:`, err)
      }

      progress.value = ((i + 1) / files.value.length) * 100
    }

    const successCount = files.value.filter(f => f.status === 'done').length
    if (successCount > 0) {
      successMessage.value = `Successfully processed ${successCount} files and uploaded to Google Drive!`
    }

  } catch (err) {
    error.value = err.message || 'An error occurred while processing files'
  } finally {
    isProcessing.value = false
  }
}

async function fetchDriveFiles(folderId) {
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?q='${folderId}'+in+parents+and+mimeType='application/pdf'&fields=files(id,name,size)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
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
  return allFiles.filter(file => {
    const name = file.name.toLowerCase()
    // Skip temp files
    if (name.startsWith('temp_') || name.startsWith('temp-')) return false
    // Skip already normalized files
    if (name.startsWith('normalized_')) return false
    // Only include files with .pdf extension
    if (!name.endsWith('.pdf')) return false
    return true
  })
}

async function getOrCreateNormalizedFolder(parentFolderId) {
  // First, check if "Normalized" folder already exists
  const searchResponse = await fetch(
    `https://www.googleapis.com/drive/v3/files?q='${parentFolderId}'+in+parents+and+name='Normalized'+and+mimeType='application/vnd.google-apps.folder'&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true`,
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
  const metadata = {
    name: filename,
    parents: [parentFolderId],
    mimeType: 'application/pdf'
  }

  const form = new FormData()
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }))
  form.append('file', blob)

  const response = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
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
  // Check for existing token in session storage
  const savedToken = sessionStorage.getItem('google_access_token')
  if (savedToken) {
    // Verify token is still valid
    try {
      const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { Authorization: `Bearer ${savedToken}` }
      })
      if (response.ok) {
        const data = await response.json()
        accessToken.value = savedToken
        isAuthenticated.value = true
        userEmail.value = data.email
      } else {
        // Token expired, clear it
        sessionStorage.removeItem('google_access_token')
      }
    } catch (err) {
      // Token invalid, clear it
      sessionStorage.removeItem('google_access_token')
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
