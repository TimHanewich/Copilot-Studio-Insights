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
  iconSource: string | null
  initials: string
  record: DataverseRecord
  lookupKeys: string[]
  transcriptCount: number
  mostRecentConversation: Date | null
}

type TranscriptDetail = {
  id: string
  date: Date | null
  messageCount: number
  record: DataverseRecord
}

type ChatMessage = {
  id: string
  author: 'user' | 'agent'
  text: string
  date: Date | null
  feedback: MessageFeedback | null
}

type MessageFeedback = {
  reaction: string
  text: string | null
}

type TranscriptReviewDetails = {
  channel: string
  outcome: string
  turnCount: number | null
  duration: string
  knowledgeSources: string[]
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

function getBotIconSource(bot: DataverseRecord) {
  const iconValue = getStringValue(bot, [
    'iconbase64',
    'iconBase64',
    'IconBase64',
    'boticon',
    'botIcon',
    'BotIcon',
    'icon',
    'Icon',
    'image',
    'Image',
    'entityimage',
    'picture',
    'avatar',
  ])

  if (!iconValue) {
    return null
  }

  if (
    iconValue.startsWith('data:image/') ||
    iconValue.startsWith('http://') ||
    iconValue.startsWith('https://')
  ) {
    return iconValue
  }

  if (iconValue.trim().startsWith('<svg')) {
    return `data:image/svg+xml;utf8,${encodeURIComponent(iconValue)}`
  }

  return `data:image/png;base64,${iconValue.replace(/\s/g, '')}`
}

function getBotInitials(botName: string) {
  return botName
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((word) => word[0]?.toUpperCase())
    .join('')
}

function getBotLookupKeys(bot: DataverseRecord) {
  return [
    getStringValue(bot, ['botid', 'BotId', 'id']),
    getStringValue(bot, ['name', 'botname', 'BotName', 'displayname', 'schemaname']),
  ]
    .filter((value): value is string => Boolean(value))
    .map(normalizeLookupKey)
}

function hasMatchingLookupKey(firstKeys: string[], secondKeys: string[]) {
  const lookupKeySet = new Set(firstKeys)

  return secondKeys.some((lookupKey) => lookupKeySet.has(lookupKey))
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

function getRecordDate(record: DataverseRecord, fieldNames: string[]) {
  const rawDate = getStringValue(record, fieldNames)

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
    iconSource: getBotIconSource(bot),
    initials: getBotInitials(getBotName(bot)),
    record: bot,
    lookupKeys: getBotLookupKeys(bot),
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

function formatDateOnly(date: Date | null) {
  if (!date) {
    return 'Unknown'
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
  }).format(date)
}

function getSystemUserDisplayName(systemUsers: DataverseRecord[], userId: string | null) {
  if (!userId) {
    return 'Unknown'
  }

  const normalizedUserId = normalizeLookupKey(userId)
  const user = systemUsers.find((systemUser) =>
    [
      getStringValue(systemUser, ['systemuserid', 'ownerid', 'id']),
      getStringValue(systemUser, ['azureactivedirectoryobjectid']),
    ]
      .filter((value): value is string => Boolean(value))
      .map(normalizeLookupKey)
      .includes(normalizedUserId),
  )

  return (
    user &&
    (getStringValue(user, ['fullname', 'name', 'domainname', 'internalemailaddress']) ??
      userId)
  ) || userId
}

function getTranscriptMessageCount(transcript: DataverseRecord) {
  const content = transcript.content

  if (typeof content !== 'string') {
    return 0
  }

  try {
    const parsedContent: unknown = JSON.parse(content)

    if (
      parsedContent &&
      typeof parsedContent === 'object' &&
      'activities' in parsedContent &&
      Array.isArray(parsedContent.activities)
    ) {
      return parsedContent.activities.filter(
        (activity: unknown) =>
          activity &&
          typeof activity === 'object' &&
          'type' in activity &&
          activity.type === 'message',
      ).length
    }
  } catch {
    return 0
  }

  return 0
}

function getTranscriptActivities(transcript: DataverseRecord) {
  const content = transcript.content

  if (typeof content !== 'string') {
    return []
  }

  try {
    const parsedContent: unknown = JSON.parse(content)

    if (
      parsedContent &&
      typeof parsedContent === 'object' &&
      'activities' in parsedContent &&
      Array.isArray(parsedContent.activities)
    ) {
      return parsedContent.activities.filter(
        (activity): activity is Record<string, unknown> =>
          Boolean(activity) && typeof activity === 'object',
      )
    }
  } catch {
    return []
  }

  return []
}

function getActivityDate(activity: Record<string, unknown>) {
  const timestampMs = activity.timestampMs

  if (typeof timestampMs === 'number') {
    const date = new Date(timestampMs)
    return Number.isNaN(date.getTime()) ? null : date
  }

  const timestamp = activity.timestamp
  if (typeof timestamp === 'number') {
    const date = new Date(timestamp * 1000)
    return Number.isNaN(date.getTime()) ? null : date
  }

  if (typeof timestamp === 'string') {
    const date = new Date(timestamp)
    return Number.isNaN(date.getTime()) ? null : date
  }

  return null
}

function getActivityAuthor(activity: Record<string, unknown>): ChatMessage['author'] {
  const from = activity.from

  if (from && typeof from === 'object' && 'role' in from) {
    return from.role === 1 ? 'user' : 'agent'
  }

  return 'agent'
}

function getRecordObject(value: unknown) {
  return value && typeof value === 'object' ? (value as Record<string, unknown>) : null
}

function addTextCitationTitles(value: unknown, sources: Set<string>) {
  if (!value || typeof value !== 'object') {
    return
  }

  const record = value as Record<string, unknown>
  const textCitations = record.textCitations

  if (Array.isArray(textCitations)) {
    textCitations.forEach((citation) => {
      const citationRecord = getRecordObject(citation)
      const title = citationRecord?.title

      if (typeof title === 'string' && title.trim()) {
        sources.add(title.trim())
      }
    })
  }

  Object.values(record).forEach((childValue) => addTextCitationTitles(childValue, sources))
}

function getTranscriptChannel(transcript: DataverseRecord) {
  const channelId = getTranscriptActivities(transcript)
    .map((activity) => activity.channelId)
    .find((channel): channel is string => typeof channel === 'string' && Boolean(channel))

  if (!channelId) {
    return 'Unknown'
  }

  return channelId
    .split(/[-_\s]+/)
    .filter(Boolean)
    .map((word) => `${word[0].toUpperCase()}${word.slice(1)}`)
    .join(' ')
}

function getSessionInfo(transcript: DataverseRecord) {
  const sessionActivity = getTranscriptActivities(transcript).find(
    (activity) => activity.valueType === 'SessionInfo',
  )

  return getRecordObject(sessionActivity?.value)
}

function formatDurationFromMs(durationMs: number) {
  if (!Number.isFinite(durationMs) || durationMs <= 0) {
    return 'Unknown'
  }

  const totalSeconds = Math.round(durationMs / 1000)
  const minutes = Math.floor(totalSeconds / 60)
  const seconds = totalSeconds % 60

  if (minutes === 0) {
    return `${seconds}s`
  }

  return `${minutes}m ${seconds}s`
}

function getTranscriptDuration(transcript: DataverseRecord) {
  const sessionInfo = getSessionInfo(transcript)
  const startTime =
    typeof sessionInfo?.startTimeUtc === 'string'
      ? new Date(sessionInfo.startTimeUtc)
      : null
  const endTime =
    typeof sessionInfo?.endTimeUtc === 'string' ? new Date(sessionInfo.endTimeUtc) : null

  if (
    startTime &&
    endTime &&
    !Number.isNaN(startTime.getTime()) &&
    !Number.isNaN(endTime.getTime())
  ) {
    return formatDurationFromMs(endTime.getTime() - startTime.getTime())
  }

  const activityDates = getTranscriptActivities(transcript)
    .map(getActivityDate)
    .filter((date): date is Date => Boolean(date))
    .sort((first, second) => first.getTime() - second.getTime())

  if (activityDates.length < 2) {
    return 'Unknown'
  }

  return formatDurationFromMs(
    activityDates[activityDates.length - 1].getTime() - activityDates[0].getTime(),
  )
}

function getTranscriptKnowledgeSources(transcript: DataverseRecord) {
  const sources = new Set<string>()

  getTranscriptActivities(transcript).forEach((activity) => {
    addTextCitationTitles(activity, sources)
  })

  return [...sources]
}

function getFeedbackByReplyToId(transcript: DataverseRecord) {
  const feedbackByReplyToId = new Map<string, MessageFeedback>()

  getTranscriptActivities(transcript).forEach((activity) => {
    if (activity.type !== 'invoke' || activity.name !== 'message/submitAction') {
      return
    }

    const value = getRecordObject(activity.value)
    const actionValue = getRecordObject(value?.actionValue)
    const feedback = getRecordObject(actionValue?.feedback)
    const replyToId = typeof activity.replyToId === 'string' ? activity.replyToId : null
    const reaction = typeof actionValue?.reaction === 'string' ? actionValue.reaction : null

    if (!replyToId || !reaction) {
      return
    }

    feedbackByReplyToId.set(replyToId, {
      reaction,
      text:
        typeof feedback?.feedbackText === 'string' && feedback.feedbackText.trim()
          ? feedback.feedbackText.trim()
          : null,
    })
  })

  return feedbackByReplyToId
}

function getTranscriptReviewDetails(transcript: DataverseRecord): TranscriptReviewDetails {
  const sessionInfo = getSessionInfo(transcript)

  return {
    channel: getTranscriptChannel(transcript),
    outcome:
      (typeof sessionInfo?.outcome === 'string' && sessionInfo.outcome) || 'Unknown',
    turnCount:
      typeof sessionInfo?.turnCount === 'number' ? sessionInfo.turnCount : null,
    duration: getTranscriptDuration(transcript),
    knowledgeSources: getTranscriptKnowledgeSources(transcript),
  }
}

function buildChatMessages(transcript: DataverseRecord): ChatMessage[] {
  const fallbackDate = getConversationDate(transcript)
  const feedbackByReplyToId = getFeedbackByReplyToId(transcript)

  return getTranscriptActivities(transcript)
    .filter((activity) => activity.type === 'message')
    .map((activity, index) => {
      const text =
        typeof activity.text === 'string' && activity.text.trim()
          ? activity.text.trim()
          : typeof activity.speak === 'string' && activity.speak.trim()
            ? activity.speak.trim()
            : ''

      if (!text) {
        return null
      }

      return {
        id: typeof activity.id === 'string' ? activity.id : `message-${index}`,
        author: getActivityAuthor(activity),
        text,
        date: getActivityDate(activity) ?? fallbackDate,
        feedback:
          typeof activity.id === 'string'
            ? feedbackByReplyToId.get(activity.id) ?? null
            : null,
      }
    })
    .filter((message): message is ChatMessage => Boolean(message))
}

function buildTranscriptDetailsForBot(
  bot: BotSummary,
  conversationTranscripts: DataverseRecord[],
): TranscriptDetail[] {
  return conversationTranscripts
    .filter((transcript) =>
      hasMatchingLookupKey(bot.lookupKeys, getTranscriptLookupKeys(transcript)),
    )
    .map((transcript, index) => ({
      id:
        getStringValue(transcript, ['conversationtranscriptid', 'activityid', 'id']) ??
        `transcript-${index}`,
      date: getConversationDate(transcript),
      messageCount: getTranscriptMessageCount(transcript),
      record: transcript,
    }))
    .sort((first, second) => {
      const firstTime = first.date?.getTime() ?? 0
      const secondTime = second.date?.getTime() ?? 0

      return secondTime - firstTime
    })
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
  const [selectedBotId, setSelectedBotId] = useState<string | null>(null)
  const [selectedTranscriptId, setSelectedTranscriptId] = useState<string | null>(
    null,
  )
  const isLoggingIn = inProgress !== InteractionStatus.None
  const botSummaries = buildBotSummaries(bots, conversationTranscripts)
  const selectedBot =
    selectedBotId && botSummaries.find((bot) => bot.id === selectedBotId)
  const selectedBotTranscripts = selectedBot
    ? buildTranscriptDetailsForBot(selectedBot, conversationTranscripts)
    : []
  const selectedTranscript =
    selectedTranscriptId &&
    selectedBotTranscripts.find((transcript) => transcript.id === selectedTranscriptId)
  const selectedChatMessages = selectedTranscript
    ? buildChatMessages(selectedTranscript.record)
    : []
  const selectedTranscriptReviewDetails = selectedTranscript
    ? getTranscriptReviewDetails(selectedTranscript.record)
    : null

  async function handleLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setErrorMessage('')
    setIsLoginPaneFading(false)
    setShowBotPanel(false)
    setSelectedBotId(null)
    setSelectedTranscriptId(null)

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
        <section
          className="bot-panel"
          aria-labelledby={
            selectedTranscript
              ? 'conversation-title'
              : selectedBot
                ? 'agent-title'
                : 'bot-panel-title'
          }
        >
          {selectedBot && selectedTranscript ? (
            <>
              <button
                type="button"
                className="back-button"
                onClick={() => setSelectedTranscriptId(null)}
              >
                Back to agent
              </button>
              <section
                className="conversation-pane"
                aria-labelledby="conversation-title"
              >
                <div className="conversation-header">
                  {selectedBot.iconSource ? (
                    <img className="agent-icon" src={selectedBot.iconSource} alt="" />
                  ) : (
                    <span className="agent-icon bot-icon-fallback" aria-hidden="true">
                      {selectedBot.initials}
                    </span>
                  )}
                  <div>
                    <p className="eyebrow">Recorded session</p>
                    <h1 id="conversation-title">{selectedBot.name}</h1>
                    <p>
                      {formatDate(selectedTranscript.date)} ·{' '}
                      {selectedTranscript.messageCount.toLocaleString()} messages
                    </p>
                  </div>
                </div>
                <div className="conversation-review-layout">
                  <aside className="conversation-meta">
                    {selectedTranscriptReviewDetails && (
                      <dl className="review-summary" aria-label="Transcript details">
                        <div>
                          <dt>Channel</dt>
                          <dd>{selectedTranscriptReviewDetails.channel}</dd>
                        </div>
                        <div>
                          <dt>Outcome</dt>
                          <dd>{selectedTranscriptReviewDetails.outcome}</dd>
                        </div>
                        <div>
                          <dt>Turn count</dt>
                          <dd>
                            {selectedTranscriptReviewDetails.turnCount?.toLocaleString() ??
                              'Unknown'}
                          </dd>
                        </div>
                        <div>
                          <dt>Duration</dt>
                          <dd>{selectedTranscriptReviewDetails.duration}</dd>
                        </div>
                        <div className="knowledge-summary">
                          <dt>Knowledge sources used</dt>
                          <dd>
                            {selectedTranscriptReviewDetails.knowledgeSources.length >
                            0 ? (
                              selectedTranscriptReviewDetails.knowledgeSources.map(
                                (source) => <span key={source}>{source}</span>,
                              )
                            ) : (
                              <span>None detected</span>
                            )}
                          </dd>
                        </div>
                      </dl>
                    )}
                  </aside>
                  <div className="chat-thread" aria-label="Conversation transcript">
                    {selectedChatMessages.length > 0 ? (
                      selectedChatMessages.map((message) => (
                        <article
                          className={`chat-message is-${message.author}`}
                          key={message.id}
                        >
                          <span className="chat-author">
                            {message.author === 'user' ? 'User' : selectedBot.name}
                          </span>
                          <p>{message.text}</p>
                          {message.feedback && (
                            <div className="message-feedback">
                              <strong>
                                Feedback:{' '}
                                {message.feedback.reaction === 'like'
                                  ? 'Liked'
                                  : message.feedback.reaction === 'dislike'
                                    ? 'Disliked'
                                    : message.feedback.reaction}
                              </strong>
                              {message.feedback.text && (
                                <span>{message.feedback.text}</span>
                              )}
                            </div>
                          )}
                          {message.date && (
                            <time dateTime={message.date.toISOString()}>
                              {formatDate(message.date)}
                            </time>
                          )}
                        </article>
                      ))
                    ) : (
                      <p className="empty-message">
                        No chat messages were found in this transcript.
                      </p>
                    )}
                  </div>
                </div>
              </section>
            </>
          ) : selectedBot ? (
            <>
              <button
                type="button"
                className="back-button"
                onClick={() => {
                  setSelectedTranscriptId(null)
                  setSelectedBotId(null)
                }}
              >
                Back to all bots
              </button>
              <div className="agent-detail">
                <section className="agent-profile" aria-labelledby="agent-title">
                  <div className="agent-hero">
                    {selectedBot.iconSource ? (
                      <img className="agent-icon" src={selectedBot.iconSource} alt="" />
                    ) : (
                      <span className="agent-icon bot-icon-fallback" aria-hidden="true">
                        {selectedBot.initials}
                      </span>
                    )}
                    <div>
                      <h1 id="agent-title">{selectedBot.name}</h1>
                      <p>{getStringValue(selectedBot.record, ['schemaname'])}</p>
                    </div>
                  </div>
                  <dl className="detail-grid">
                    <div>
                      <dt>Owner</dt>
                      <dd>
                        {getSystemUserDisplayName(
                          systemUsers,
                          getStringValue(selectedBot.record, [
                            '_ownerid_value',
                            '_owninguser_value',
                          ]),
                        )}
                      </dd>
                    </div>
                    <div>
                      <dt>Created by</dt>
                      <dd>
                        {getSystemUserDisplayName(
                          systemUsers,
                          getStringValue(selectedBot.record, ['_createdby_value']),
                        )}
                      </dd>
                    </div>
                    <div>
                      <dt>Created</dt>
                      <dd>
                        {formatDateOnly(
                          getRecordDate(selectedBot.record, ['createdon']),
                        )}
                      </dd>
                    </div>
                    <div>
                      <dt>Last modified</dt>
                      <dd>
                        {formatDate(
                          getRecordDate(selectedBot.record, ['modifiedon']),
                        )}
                      </dd>
                    </div>
                  </dl>
                </section>
                <section
                  className="transcript-panel"
                  aria-labelledby="transcript-panel-title"
                >
                  <div>
                    <h2 id="transcript-panel-title">
                      {selectedBotTranscripts.length.toLocaleString()} Recorded Sessions
                    </h2>
                  </div>
                  <div className="transcript-list">
                    {selectedBotTranscripts.length > 0 ? (
                      selectedBotTranscripts.map((transcript) => (
                        <button
                          className="transcript-row"
                          key={transcript.id}
                          type="button"
                          onClick={() => setSelectedTranscriptId(transcript.id)}
                        >
                          <span>{formatDate(transcript.date)}</span>
                          <strong>
                            {transcript.messageCount.toLocaleString()} messages
                          </strong>
                        </button>
                      ))
                    ) : (
                      <p className="empty-message">
                        No ConversationTranscripts were found for this bot.
                      </p>
                    )}
                  </div>
                </section>
              </div>
            </>
          ) : (
            <>
              <div className="panel-heading">
                <h1 id="bot-panel-title">Copilot Studio Activity in Your Environment</h1>
                <p className="environment-name">{environmentOrigin}</p>
              </div>
              <div className="panel-summary">
                <span>{bots.length.toLocaleString()} agents</span>
                <span>
                  {conversationTranscriptCount?.toLocaleString() ??
                    conversationTranscripts.length.toLocaleString()}{' '}
                  sessions
                </span>
              </div>
              <div className="bot-list" aria-label="Bots in this environment">
                {botSummaries.length > 0 ? (
                  botSummaries.map((bot) => (
                    <button
                      className="bot-row"
                      key={bot.id}
                      type="button"
                      onClick={() => setSelectedBotId(bot.id)}
                    >
                      <div className="bot-main">
                        {bot.iconSource ? (
                          <img className="bot-icon" src={bot.iconSource} alt="" />
                        ) : (
                          <span
                            className="bot-icon bot-icon-fallback"
                            aria-hidden="true"
                          >
                            {bot.initials}
                          </span>
                        )}
                        <div>
                          <h2>{bot.name}</h2>
                          <p>
                            Most recent interaction:{' '}
                            {formatDate(bot.mostRecentConversation)}
                          </p>
                        </div>
                      </div>
                      <div className="bot-stat">
                        <span>{bot.transcriptCount.toLocaleString()}</span>
                        <small>recorded sessions</small>
                      </div>
                    </button>
                  ))
                ) : (
                  <p className="empty-message">
                    No bots were found in this environment.
                  </p>
                )}
              </div>
            </>
          )}
        </section>
      ) : (
        <form
          className={`login-card${isLoginPaneFading ? ' is-fading' : ''}`}
          aria-labelledby="page-title"
          onSubmit={handleLogin}
        >
          <h1 id="page-title">Copilot Studio Insights</h1>
          <input
            id="environment-url"
            name="environmentUrl"
            type="url"
            placeholder="Your environment URL (e.g. https://your-environment.crm.dynamics.com)"
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
