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

function decodeBase64Url(value: string) {
  const base64 = value.replace(/-/g, '+').replace(/_/g, '/')
  const paddedBase64 = base64.padEnd(
    base64.length + ((4 - (base64.length % 4)) % 4),
    '=',
  )

  return window.atob(paddedBase64)
}

function getMsalStateId(hash: string) {
  const state = new URLSearchParams(hash.slice(1)).get('state')
  if (!state) {
    return null
  }

  try {
    const parsedState: unknown = JSON.parse(decodeBase64Url(state))

    if (
      parsedState &&
      typeof parsedState === 'object' &&
      'id' in parsedState &&
      typeof parsedState.id === 'string'
    ) {
      return parsedState.id
    }
  } catch (error) {
    console.error('Unable to parse MSAL callback state.', error)
  }

  return null
}

const isMsalPopupCallback =
  window.opener &&
  window.location.hash.includes('code=') &&
  window.location.hash.includes('state=')

if (isMsalPopupCallback) {
  document.body.textContent = 'Completing sign-in...'
  const callbackHash = window.location.hash
  const channelId = getMsalStateId(callbackHash)

  if (channelId) {
    const channel = new BroadcastChannel(channelId)
    let messageCount = 0
    const messageTimer = window.setInterval(() => {
      channel.postMessage({ payload: callbackHash, v: 1 })
      messageCount += 1

      if (messageCount >= 10) {
        window.clearInterval(messageTimer)
        channel.close()
        window.close()
      }
    }, 100)
  } else {
    document.body.textContent =
      'Unable to complete sign-in because the authentication state could not be read.'
  }
} else {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>,
  )
}
