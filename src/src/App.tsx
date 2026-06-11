import { useState, type FormEvent } from 'react'
import { InteractionStatus } from '@azure/msal-browser'
import { useMsal } from '@azure/msal-react'
import './App.css'

const DATAVERSE_DELEGATED_SCOPE = 'user_impersonation'

function buildDataverseScope(environmentUrl: string) {
  const url = new URL(environmentUrl)

  return `${url.origin}/${DATAVERSE_DELEGATED_SCOPE}`
}

function App() {
  const { instance, accounts, inProgress } = useMsal()
  const [environmentUrl, setEnvironmentUrl] = useState('')
  const [errorMessage, setErrorMessage] = useState('')
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
    try {
      scope = buildDataverseScope(trimmedEnvironmentUrl)
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
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Unable to sign in with the provided environment URL.',
      )
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
      <form className="login-card" aria-labelledby="page-title" onSubmit={handleLogin}>
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
        {accounts.length > 0 && (
          <p className="status-message">
            Signed in as {accounts[0].username || accounts[0].name}
          </p>
        )}
        {errorMessage && <p className="error-message">{errorMessage}</p>}
      </form>
    </main>
  )
}

export default App
