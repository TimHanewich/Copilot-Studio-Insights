import { useState, type FormEvent } from 'react'
import { InteractionStatus } from '@azure/msal-browser'
import { useMsal } from '@azure/msal-react'
import './App.css'

const DATAVERSE_DELEGATED_SCOPE = 'user_impersonation'
const CONVERSATION_TRANSCRIPT_TABLE = 'conversationtranscripts'

type ConversationTranscriptCountResponse = {
  '@odata.count'?: number
}

function getEnvironmentOrigin(environmentUrl: string) {
  const url = new URL(environmentUrl)

  return url.origin
}

function buildDataverseScope(environmentOrigin: string) {
  return `${environmentOrigin}/${DATAVERSE_DELEGATED_SCOPE}`
}

async function fetchConversationTranscriptCount(
  environmentOrigin: string,
  accessToken: string,
) {
  const endpoint = new URL(
    `/api/data/v9.2/${CONVERSATION_TRANSCRIPT_TABLE}`,
    environmentOrigin,
  )
  endpoint.searchParams.set('$select', 'conversationtranscriptid')
  endpoint.searchParams.set('$count', 'true')
  endpoint.searchParams.set('$top', '1')

  const response = await fetch(endpoint, {
    headers: {
      Accept: 'application/json',
      Authorization: `Bearer ${accessToken}`,
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
  })

  if (!response.ok) {
    throw new Error(
      `Dataverse request failed: ${response.status} ${response.statusText}`,
    )
  }

  const data = (await response.json()) as ConversationTranscriptCountResponse
  if (typeof data['@odata.count'] !== 'number') {
    throw new Error('Dataverse did not return a conversation transcript count.')
  }

  return data['@odata.count']
}

function App() {
  const { instance, accounts, inProgress } = useMsal()
  const [environmentUrl, setEnvironmentUrl] = useState('')
  const [errorMessage, setErrorMessage] = useState('')
  const [environmentOrigin, setEnvironmentOrigin] = useState('')
  const [conversationTranscriptCount, setConversationTranscriptCount] = useState<
    number | null
  >(null)
  const [isLoadingCount, setIsLoadingCount] = useState(false)
  const isLoggingIn = inProgress !== InteractionStatus.None

  async function handleLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setErrorMessage('')

    const trimmedEnvironmentUrl = environmentUrl.trim()
    if (!trimmedEnvironmentUrl) {
      setErrorMessage('Enter an environment URL before logging in.')
      return
    }

    let scope: string
    let parsedEnvironmentOrigin: string
    try {
      parsedEnvironmentOrigin = getEnvironmentOrigin(trimmedEnvironmentUrl)
      scope = buildDataverseScope(parsedEnvironmentOrigin)
    } catch {
      setErrorMessage(
        'Enter a valid environment URL, such as https://org53194471.crm.dynamics.com.',
      )
      return
    }

    try {
      const result = await instance.loginPopup({
        scopes: [scope],
      })

      if (result.account) {
        instance.setActiveAccount(result.account)
      }

      const accessToken =
        result.accessToken ||
        (
          await instance.acquireTokenSilent({
            scopes: [scope],
            account: result.account ?? undefined,
          })
        ).accessToken

      setEnvironmentOrigin(parsedEnvironmentOrigin)
      setIsLoadingCount(true)

      const count = await fetchConversationTranscriptCount(
        parsedEnvironmentOrigin,
        accessToken,
      )
      setConversationTranscriptCount(count)
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Unable to sign in with the provided environment URL.',
      )
    } finally {
      setIsLoadingCount(false)
    }
  }

  return (
    <main className="landing-page">
      <p className="credit">
        Made by{' '}
        <a href="https://timh.ai" target="_blank" rel="noreferrer">
          Tim Hanewich
        </a>
      </p>
      {environmentOrigin ? (
        <section className="insights-card" aria-labelledby="page-title">
          <h1 id="page-title">Copilot Studio Insights</h1>
          <p className="environment-name">{environmentOrigin}</p>
          {isLoadingCount ? (
            <p className="loading-message">Loading conversation transcripts...</p>
          ) : (
            <div className="metric">
              <span className="metric-value">
                {conversationTranscriptCount?.toLocaleString() ?? '--'}
              </span>
              <span className="metric-label">ConversationTranscript records</span>
            </div>
          )}
          {accounts.length > 0 && (
            <p className="status-message">
              Signed in as {accounts[0].username || accounts[0].name}
            </p>
          )}
          {errorMessage && <p className="error-message">{errorMessage}</p>}
        </section>
      ) : (
        <form
          className="login-card"
          aria-labelledby="page-title"
          onSubmit={handleLogin}
        >
          <h1 id="page-title">Copilot Studio Insights</h1>
          <label htmlFor="environment-url">Environment URL</label>
          <input
            id="environment-url"
            name="environmentUrl"
            type="url"
            placeholder="https://your-environment.crm.dynamics.com"
            value={environmentUrl}
            onChange={(event) => setEnvironmentUrl(event.target.value)}
            required
          />
          <button type="submit" disabled={isLoggingIn}>
            {isLoggingIn ? 'Logging in...' : 'Login'}
          </button>
          {errorMessage && <p className="error-message">{errorMessage}</p>}
        </form>
      )}
    </main>
  )
}

export default App
