<template>
  <v-app>
    <v-app-bar color="primary" density="compact">
      <v-app-bar-title>
        <v-icon class="mr-2">mdi-file-pdf-box</v-icon>
        PDF Normalizer
      </v-app-bar-title>
    </v-app-bar>

    <v-main>
      <v-container class="py-8">
        <PdfProcessor />
      </v-container>
    </v-main>

    <v-footer app class="text-center d-flex flex-column py-3">
      <div class="text-caption">
        Process and normalize PDFs from Google Drive
      </div>
    </v-footer>
  </v-app>
</template>

<script setup>
import PdfProcessor from './components/PdfProcessor.vue'
</script>
