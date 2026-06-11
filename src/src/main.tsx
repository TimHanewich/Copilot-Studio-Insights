import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import {
  BrowserCacheLocation,
  PublicClientApplication,
  type Configuration,
} from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import './index.css'
import App from './App.tsx'

const msalConfig: Configuration = {
  auth: {
    clientId: '537c9fd5-1887-4786-9fc2-05863149de86',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: BrowserCacheLocation.SessionStorage,
  },
}

const msalInstance = new PublicClientApplication(msalConfig)

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <MsalProvider instance={msalInstance}>
      <App />
    </MsalProvider>
  </StrictMode>,
)
