<template>
  <v-app>
    <v-app-bar color="primary" density="compact">
      <v-spacer></v-spacer>
      <v-app-bar-title style="direction: rtl; text-align: right;">
        پردازشگر فاکتورهای PDF
        <v-icon class="ml-2">mdi-file-pdf-box</v-icon>
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
