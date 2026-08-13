/// <reference types="vite/client" />

declare interface ImportMetaEnv {
  readonly VITE_AZURE_CLIENT_ID: string;
  readonly VITE_AZURE_TENANT_ID: string;
  readonly VITE_REDIRECT_URI: string;
  readonly VITE_API_BASE_URL: string;
}

declare interface ImportMeta {
  readonly env: ImportMetaEnv;
}
