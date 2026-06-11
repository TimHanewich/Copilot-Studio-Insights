import { useState, type FormEvent } from 'react'
import { InteractionStatus } from '@azure/msal-browser'
import { useMsal } from '@azure/msal-react'
import './App.css'

const DATAVERSE_DELEGATED_SCOPE = 'user_impersonation'
const CONVERSATION_TRANSCRIPT_TABLE = 'conversationtranscripts'
const BOTS_TABLE = 'bots'
const SYSTEM_USERS_TABLE = 'systemusers'

type DataverseRecord = Record<string, unknown>

type DataverseCollectionResponse = {
  value: DataverseRecord[]
  '@odata.count'?: number
  '@odata.nextLink'?: string
}

type DataverseCollectionResult = {
  records: DataverseRecord[]
  count: number
}

type TranscriptMetadata = {
  BotId?: string
  BotName?: string
}

type BotSummary = {
  id: string
  name: string
  transcriptCount: number
  mostRecentConversation: Date | null
}

function getEnvironmentOrigin(environmentUrl: string) {
  const url = new URL(environmentUrl)

  return url.origin
}

function buildDataverseScope(environmentOrigin: string) {
  return `${environmentOrigin}/${DATAVERSE_DELEGATED_SCOPE}`
}

async function fetchDataverseCollection(
  environmentOrigin: string,
  accessToken: string,
  collectionName: string,
): Promise<DataverseCollectionResult> {
  const endpoint = new URL(`/api/data/v9.2/${collectionName}`, environmentOrigin)
  endpoint.searchParams.set('$count', 'true')
  let requestUrl = endpoint.toString()
  const records: DataverseRecord[] = []
  let count: number | null = null

  while (requestUrl) {
    const response = await fetch(requestUrl, {
      headers: {
        Accept: 'application/json',
        Authorization: `Bearer ${accessToken}`,
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.maxpagesize=5000',
      },
    })

    if (!response.ok) {
      throw new Error(
        `Dataverse ${collectionName} request failed: ${response.status} ${response.statusText}`,
      )
    }

    const data = (await response.json()) as DataverseCollectionResponse
    records.push(...data.value)

    if (typeof data['@odata.count'] === 'number') {
      count = data['@odata.count']
    }

    requestUrl = data['@odata.nextLink'] ?? ''
  }

  return {
    records,
    count: count ?? records.length,
  }
}

function getStringValue(record: DataverseRecord, fieldNames: string[]) {
  for (const fieldName of fieldNames) {
    const value = record[fieldName]

    if (typeof value === 'string' && value.trim()) {
      return value.trim()
    }
  }

  return null
}

function parseTranscriptMetadata(record: DataverseRecord): TranscriptMetadata {
  const metadata = record.metadata

  if (typeof metadata !== 'string') {
    return {}
  }

  try {
    const parsedMetadata: unknown = JSON.parse(metadata)

    if (parsedMetadata && typeof parsedMetadata === 'object') {
      return parsedMetadata as TranscriptMetadata
    }
  } catch {
    return {}
  }

  return {}
}

function normalizeLookupKey(value: string) {
  return value.trim().toLowerCase()
}

function getBotId(bot: DataverseRecord, index: number) {
  return (
    getStringValue(bot, ['botid', 'BotId', 'id']) ??
    `bot-${index}`
  )
}

function getBotName(bot: DataverseRecord) {
  return (
    getStringValue(bot, [
      'name',
      'botname',
      'BotName',
      'displayname',
      'schemaname',
    ]) ?? 'Untitled bot'
  )
}

function getBotLookupKeys(bot: DataverseRecord) {
  return [
    getStringValue(bot, ['botid', 'BotId', 'id']),
    getStringValue(bot, ['name', 'botname', 'BotName', 'displayname', 'schemaname']),
  ]
    .filter((value): value is string => Boolean(value))
    .map(normalizeLookupKey)
}

function getTranscriptLookupKeys(transcript: DataverseRecord) {
  const metadata = parseTranscriptMetadata(transcript)

  return [
    getStringValue(transcript, ['_bot_conversationtranscriptid_value']),
    metadata.BotId,
    metadata.BotName,
  ]
    .filter((value): value is string => Boolean(value))
    .map(normalizeLookupKey)
}

function getConversationDate(transcript: DataverseRecord) {
  const rawDate = getStringValue(transcript, [
    'conversationstarttime',
    'createdon',
    'modifiedon',
  ])

  if (!rawDate) {
    return null
  }

  const date = new Date(rawDate)
  return Number.isNaN(date.getTime()) ? null : date
}

function buildBotSummaries(
  bots: DataverseRecord[],
  conversationTranscripts: DataverseRecord[],
): BotSummary[] {
  const summaries = bots.map((bot, index) => ({
    id: getBotId(bot, index),
    name: getBotName(bot),
    transcriptCount: 0,
    mostRecentConversation: null as Date | null,
  }))
  const botIndexByLookupKey = new Map<string, number>()

  bots.forEach((bot, index) => {
    getBotLookupKeys(bot).forEach((lookupKey) => {
      botIndexByLookupKey.set(lookupKey, index)
    })
  })

  conversationTranscripts.forEach((transcript) => {
    const matchingBotIndex = getTranscriptLookupKeys(transcript)
      .map((lookupKey) => botIndexByLookupKey.get(lookupKey))
      .find((index): index is number => typeof index === 'number')

    if (matchingBotIndex === undefined) {
      return
    }

    const summary = summaries[matchingBotIndex]
    const conversationDate = getConversationDate(transcript)
    summary.transcriptCount += 1

    if (
      conversationDate &&
      (!summary.mostRecentConversation ||
        conversationDate > summary.mostRecentConversation)
    ) {
      summary.mostRecentConversation = conversationDate
    }
  })

  return summaries.sort((first, second) => {
    const firstTime = first.mostRecentConversation?.getTime() ?? 0
    const secondTime = second.mostRecentConversation?.getTime() ?? 0

    return (
      secondTime - firstTime ||
      second.transcriptCount - first.transcriptCount ||
      first.name.localeCompare(second.name)
    )
  })
}

function formatDate(date: Date | null) {
  if (!date) {
    return 'No conversations yet'
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(date)
}

function App() {
  const { instance, inProgress } = useMsal()
  const [environmentUrl, setEnvironmentUrl] = useState('')
  const [errorMessage, setErrorMessage] = useState('')
  const [environmentOrigin, setEnvironmentOrigin] = useState('')
  const [conversationTranscripts, setConversationTranscripts] = useState<
    DataverseRecord[]
  >([])
  const [bots, setBots] = useState<DataverseRecord[]>([])
  const [systemUsers, setSystemUsers] = useState<DataverseRecord[]>([])
  const [conversationTranscriptCount, setConversationTranscriptCount] =
    useState<number | null>(null)
  const [isLoadingDataverseData, setIsLoadingDataverseData] = useState(false)
  const [isLoginPaneFading, setIsLoginPaneFading] = useState(false)
  const [showBotPanel, setShowBotPanel] = useState(false)
  const isLoggingIn = inProgress !== InteractionStatus.None
  const botSummaries = buildBotSummaries(bots, conversationTranscripts)

  async function handleLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setErrorMessage('')
    setIsLoginPaneFading(false)
    setShowBotPanel(false)

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
      setIsLoadingDataverseData(true)

      const [conversationTranscriptResult, botResult, systemUserResult] =
        await Promise.all([
          fetchDataverseCollection(
            parsedEnvironmentOrigin,
            accessToken,
            CONVERSATION_TRANSCRIPT_TABLE,
          ),
          fetchDataverseCollection(parsedEnvironmentOrigin, accessToken, BOTS_TABLE),
          fetchDataverseCollection(
            parsedEnvironmentOrigin,
            accessToken,
            SYSTEM_USERS_TABLE,
          ),
        ])

      setConversationTranscripts(conversationTranscriptResult.records)
      setConversationTranscriptCount(conversationTranscriptResult.count)
      setBots(botResult.records)
      setSystemUsers(systemUserResult.records)
      setIsLoginPaneFading(true)
      window.setTimeout(() => setShowBotPanel(true), 450)
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Unable to sign in with the provided environment URL.',
      )
    } finally {
      setIsLoadingDataverseData(false)
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
      {showBotPanel ? (
        <section className="bot-panel" aria-labelledby="bot-panel-title">
          <div className="panel-heading">
            <p className="eyebrow">Environment overview</p>
            <h1 id="bot-panel-title">Bots</h1>
            <p className="environment-name">{environmentOrigin}</p>
          </div>
          <div className="panel-summary">
            <span>{bots.length.toLocaleString()} bots</span>
            <span>
              {conversationTranscriptCount?.toLocaleString() ??
                conversationTranscripts.length.toLocaleString()}{' '}
              transcripts
            </span>
            <span>{systemUsers.length.toLocaleString()} system users</span>
          </div>
          <div className="bot-list" aria-label="Bots in this environment">
            {botSummaries.length > 0 ? (
              botSummaries.map((bot) => (
                <article className="bot-row" key={bot.id}>
                  <div>
                    <h2>{bot.name}</h2>
                    <p>
                      Most recent interaction:{' '}
                      {formatDate(bot.mostRecentConversation)}
                    </p>
                  </div>
                  <div className="bot-stat">
                    <span>{bot.transcriptCount.toLocaleString()}</span>
                    <small>ConversationTranscripts</small>
                  </div>
                </article>
              ))
            ) : (
              <p className="empty-message">No bots were found in this environment.</p>
            )}
          </div>
        </section>
      ) : (
        <form
          className={`login-card${isLoginPaneFading ? ' is-fading' : ''}`}
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
            {isLoggingIn || isLoadingDataverseData ? 'Loading...' : 'Login'}
          </button>
          {isLoadingDataverseData && (
            <p className="loading-message">Loading Dataverse data...</p>
          )}
          {errorMessage && <p className="error-message">{errorMessage}</p>}
        </form>
      )}
    </main>
  )
}

export default App
