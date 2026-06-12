import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from 'react'
import { InteractionStatus } from '@azure/msal-browser'
import { useMsal } from '@azure/msal-react'
import './App.css'

const DATAVERSE_DELEGATED_SCOPE = 'user_impersonation'
const CONVERSATION_TRANSCRIPT_TABLE = 'conversationtranscripts'
const BOTS_TABLE = 'bots'
const SYSTEM_USERS_TABLE = 'systemusers'
const LAST_ENVIRONMENT_URL_STORAGE_KEY =
  'copilot-studio-insights:last-environment-url'
const DIRECT_LINE_BASE_URL = 'https://directline.botframework.com/v3/directline'
const DIRECT_LINE_USER_ID = 'copilot-studio-insights-user'

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

type MakerSummary = {
  id: string
  name: string
  agents: BotSummary[]
}

type KnowledgeSourceAgentUsage = {
  agent: BotSummary
  count: number
}

type KnowledgeSourceSummary = {
  id: string
  name: string
  totalUsageCount: number
  agents: KnowledgeSourceAgentUsage[]
}

type FeedbackSummary = {
  id: string
  agent: BotSummary
  transcript: TranscriptDetail
  message: ChatMessage
  feedback: MessageFeedback
}

type LandingMode = 'choice' | 'telemetry' | 'direct-line-setup' | 'direct-line-chat'

type AppView = 'dashboard' | 'agents' | 'makers' | 'knowledge' | 'feedback'

type DirectLineTokenResponse = {
  token?: string
  expires_in?: number
  conversationId?: string
}

type DirectLineConversationResponse = {
  token?: string
  expires_in?: number
  conversationId?: string
}

type DirectLineEntity = {
  type?: string
  title?: string
  text?: string
}

type DirectLineActivity = {
  type?: string
  id?: string
  channelData?: {
    clientActivityId?: string
  }
  entities?: DirectLineEntity[]
  text?: string
  speak?: string
  timestamp?: string
  from?: {
    id?: string
    name?: string
    role?: string
  }
}

type DirectLineActivitiesResponse = {
  activities?: DirectLineActivity[]
  watermark?: string
}

type DirectLineSession = {
  token: string
  conversationId: string
  expiresAt: Date | null
}

type LiveChatMessage = {
  id: string
  author: 'user' | 'agent'
  kind: 'message' | 'thought'
  title: string | null
  text: string
  date: Date | null
  displayName: string
}

type DirectLinePostActivity = {
  type: string
  text?: string
  from: {
    id: string
    role: 'user'
  }
  channelData?: {
    clientActivityId: string
  }
}

const APP_NAV_ITEMS: { id: AppView; label: string }[] = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'agents', label: 'Agents' },
  { id: 'makers', label: 'Makers' },
  { id: 'knowledge', label: 'Knowledge' },
  { id: 'feedback', label: 'Feedback' },
]

function getEnvironmentOrigin(environmentUrl: string) {
  const url = new URL(environmentUrl)

  return url.origin
}

function buildDataverseScope(environmentOrigin: string) {
  return `${environmentOrigin}/${DATAVERSE_DELEGATED_SCOPE}`
}

function getStoredEnvironmentUrl() {
  if (typeof window === 'undefined') {
    return ''
  }

  try {
    return window.localStorage.getItem(LAST_ENVIRONMENT_URL_STORAGE_KEY) ?? ''
  } catch (error) {
    console.warn('Unable to read the saved environment URL.', error)
    return ''
  }
}

function storeLastEnvironmentUrl(environmentUrl: string) {
  try {
    window.localStorage.setItem(LAST_ENVIRONMENT_URL_STORAGE_KEY, environmentUrl)
  } catch (error) {
    console.warn('Unable to save the environment URL.', error)
  }
}

async function getResponseErrorMessage(response: Response, fallbackMessage: string) {
  const responseBody = await response.text()
  let detail = responseBody.trim()

  if (detail) {
    try {
      const parsedBody: unknown = JSON.parse(detail)

      if (
        parsedBody &&
        typeof parsedBody === 'object' &&
        'error' in parsedBody &&
        typeof parsedBody.error === 'object' &&
        parsedBody.error &&
        'message' in parsedBody.error &&
        typeof parsedBody.error.message === 'string'
      ) {
        detail = parsedBody.error.message
      } else if (
        parsedBody &&
        typeof parsedBody === 'object' &&
        'message' in parsedBody &&
        typeof parsedBody.message === 'string'
      ) {
        detail = parsedBody.message
      }
    } catch {
      detail = responseBody.trim()
    }
  }

  const statusText = `${response.status} ${response.statusText}`.trim()

  return detail
    ? `${fallbackMessage} (${statusText}): ${detail}`
    : `${fallbackMessage} (${statusText}).`
}

async function requestDirectLineToken(tokenEndpoint: string) {
  const response = await fetch(tokenEndpoint)

  if (!response.ok) {
    throw new Error(
      await getResponseErrorMessage(
        response,
        'Unable to request a Direct Line token from the token endpoint',
      ),
    )
  }

  const tokenResponse = (await response.json()) as DirectLineTokenResponse

  if (!tokenResponse.token) {
    throw new Error('The token endpoint did not return a Direct Line token.')
  }

  return tokenResponse
}

async function startDirectLineConversation(token: string) {
  const response = await fetch(`${DIRECT_LINE_BASE_URL}/conversations`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
    },
  })

  if (!response.ok) {
    throw new Error(
      await getResponseErrorMessage(
        response,
        'Unable to start a Direct Line conversation',
      ),
    )
  }

  return (await response.json()) as DirectLineConversationResponse
}

async function fetchDirectLineActivities(
  session: DirectLineSession,
  watermark: string | null,
) {
  const endpoint = new URL(
    `${DIRECT_LINE_BASE_URL}/conversations/${session.conversationId}/activities`,
  )

  if (watermark) {
    endpoint.searchParams.set('watermark', watermark)
  }

  const response = await fetch(endpoint, {
    headers: {
      Authorization: `Bearer ${session.token}`,
    },
  })

  if (!response.ok) {
    throw new Error(
      await getResponseErrorMessage(response, 'Unable to retrieve agent activities'),
    )
  }

  return (await response.json()) as DirectLineActivitiesResponse
}

async function postDirectLineActivity(
  session: DirectLineSession,
  activity: DirectLinePostActivity,
) {
  const response = await fetch(
    `${DIRECT_LINE_BASE_URL}/conversations/${session.conversationId}/activities`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${session.token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(activity),
    },
  )

  if (!response.ok) {
    throw new Error(
      await getResponseErrorMessage(response, 'Unable to send the Direct Line activity'),
    )
  }

  return response
}

function getDirectLineActivityText(activity: DirectLineActivity) {
  if (typeof activity.text === 'string' && activity.text.trim()) {
    return activity.text.trim()
  }

  if (typeof activity.speak === 'string' && activity.speak.trim()) {
    return activity.speak.trim()
  }

  return null
}

function getDirectLineThoughtEntities(activity: DirectLineActivity) {
  return (activity.entities ?? [])
    .filter((entity) => entity.type === 'thought')
    .map((entity) => {
      const title =
        typeof entity.title === 'string' && entity.title.trim()
          ? entity.title.trim()
          : null
      const text =
        typeof entity.text === 'string' && entity.text.trim()
          ? entity.text.trim()
          : null

      return title || text ? { title, text } : null
    })
    .filter(
      (thought): thought is { title: string | null; text: string | null } =>
        Boolean(thought),
    )
}

function getDirectLineActivityDate(activity: DirectLineActivity) {
  if (!activity.timestamp) {
    return null
  }

  const date = new Date(activity.timestamp)

  return Number.isNaN(date.getTime()) ? null : date
}

function getDirectLineErrorText(error: unknown) {
  return error instanceof Error
    ? error.message
    : 'Unable to connect to the agent through Direct Line.'
}

function getDirectLineClientActivityId(activity: DirectLineActivity) {
  return typeof activity.channelData?.clientActivityId === 'string'
    ? activity.channelData.clientActivityId
    : null
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

function buildMakerSummaries(
  bots: BotSummary[],
  systemUsers: DataverseRecord[],
): MakerSummary[] {
  const makers = new Map<string, MakerSummary>()

  bots.forEach((bot) => {
    const makerId =
      getStringValue(bot.record, ['_createdby_value', 'createdby']) ?? 'unknown-maker'
    const makerKey = normalizeLookupKey(makerId)
    const existingMaker = makers.get(makerKey)

    if (existingMaker) {
      existingMaker.agents.push(bot)
      return
    }

    makers.set(makerKey, {
      id: makerKey,
      name:
        makerId === 'unknown-maker'
          ? 'Unknown maker'
          : getSystemUserDisplayName(systemUsers, makerId),
      agents: [bot],
    })
  })

  return [...makers.values()]
    .map((maker) => ({
      ...maker,
      agents: [...maker.agents].sort((first, second) =>
        first.name.localeCompare(second.name),
      ),
    }))
    .sort(
      (first, second) =>
        second.agents.length - first.agents.length || first.name.localeCompare(second.name),
    )
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

function formatKnowledgeSourceName(source: string) {
  const trimmedSource = source.trim()

  if (trimmedSource === 'BingUnscopedSearchKnowledge') {
    return 'Bing Unscoped Search'
  }

  const fileSourceParts = trimmedSource.split('.file.')
  const fileSource = fileSourceParts[fileSourceParts.length - 1]
  const fileName =
    fileSource?.match(/^(.+\.(?:csv|docx?|html?|md|pdf|pptx?|txt|xlsx?))(?:[_-].*)?$/i)?.[1] ??
    fileSource

  return fileName
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function getKnowledgeSourceName(source: unknown) {
  if (typeof source === 'string' && source.trim()) {
    return formatKnowledgeSourceName(source)
  }

  const sourceRecord = getRecordObject(source)
  const sourceName =
    sourceRecord &&
    ['title', 'name', 'Name', 'displayName', 'sourceName']
      .map((fieldName) => sourceRecord[fieldName])
      .find(
        (fieldValue): fieldValue is string =>
          typeof fieldValue === 'string' && Boolean(fieldValue.trim()),
      )

  return sourceName ? formatKnowledgeSourceName(sourceName) : null
}

function getTranscriptKnowledgeSourceUsages(transcript: DataverseRecord) {
  const sources: string[] = []

  getTranscriptActivities(transcript).forEach((activity) => {
    const value = getRecordObject(activity.value)
    const outputKnowledgeSources = value?.outputKnowledgeSources

    if (!Array.isArray(outputKnowledgeSources)) {
      return
    }

    outputKnowledgeSources.forEach((source) => {
      const sourceName = getKnowledgeSourceName(source)

      if (sourceName) {
        sources.push(sourceName)
      }
    })
  })

  return sources
}

function getTranscriptKnowledgeSources(transcript: DataverseRecord) {
  const sources = new Set(getTranscriptKnowledgeSourceUsages(transcript))

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

function getFeedbackReactionLabel(reaction: string) {
  if (reaction === 'like') {
    return 'Liked'
  }

  if (reaction === 'dislike') {
    return 'Disliked'
  }

  return reaction
}

function getFeedbackReactionIcon(reaction: string) {
  return reaction === 'dislike' ? '👎' : '👍'
}

function getFeedbackTooltip(feedback: MessageFeedback) {
  const label = getFeedbackReactionLabel(feedback.reaction)

  return feedback.text ? `${label}: ${feedback.text}` : label
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

function buildKnowledgeSourceSummaries(
  bots: BotSummary[],
  conversationTranscripts: DataverseRecord[],
): KnowledgeSourceSummary[] {
  const knowledgeSources = new Map<
    string,
    {
      id: string
      name: string
      totalUsageCount: number
      agentUsageById: Map<string, KnowledgeSourceAgentUsage>
    }
  >()

  conversationTranscripts.forEach((transcript) => {
    const transcriptLookupKeys = getTranscriptLookupKeys(transcript)
    const matchingBot = bots.find((bot) =>
      hasMatchingLookupKey(bot.lookupKeys, transcriptLookupKeys),
    )

    if (!matchingBot) {
      return
    }

    getTranscriptKnowledgeSourceUsages(transcript).forEach((sourceName) => {
      const sourceKey = normalizeLookupKey(sourceName)
      const sourceSummary =
        knowledgeSources.get(sourceKey) ??
        {
          id: sourceKey,
          name: sourceName,
          totalUsageCount: 0,
          agentUsageById: new Map<string, KnowledgeSourceAgentUsage>(),
        }
      const agentUsage =
        sourceSummary.agentUsageById.get(matchingBot.id) ??
        {
          agent: matchingBot,
          count: 0,
        }

      sourceSummary.totalUsageCount += 1
      sourceSummary.agentUsageById.set(matchingBot.id, {
        ...agentUsage,
        count: agentUsage.count + 1,
      })
      knowledgeSources.set(sourceKey, sourceSummary)
    })
  })

  return [...knowledgeSources.values()]
    .map((source) => ({
      id: source.id,
      name: source.name,
      totalUsageCount: source.totalUsageCount,
      agents: [...source.agentUsageById.values()].sort(
        (first, second) =>
          second.count - first.count || first.agent.name.localeCompare(second.agent.name),
      ),
    }))
    .sort(
      (first, second) =>
        second.totalUsageCount - first.totalUsageCount ||
        first.name.localeCompare(second.name),
    )
}

function buildFeedbackSummaries(
  bots: BotSummary[],
  conversationTranscripts: DataverseRecord[],
): FeedbackSummary[] {
  return bots
    .flatMap((bot) =>
      buildTranscriptDetailsForBot(bot, conversationTranscripts).flatMap((transcript) =>
        buildChatMessages(transcript.record)
          .filter(
            (message): message is ChatMessage & { feedback: MessageFeedback } =>
              Boolean(message.feedback),
          )
          .map((message) => ({
            id: `${bot.id}-${transcript.id}-${message.id}`,
            agent: bot,
            transcript,
            message,
            feedback: message.feedback,
          })),
      ),
    )
    .sort((first, second) => {
      const firstTime = first.message.date?.getTime() ?? 0
      const secondTime = second.message.date?.getTime() ?? 0

      return secondTime - firstTime
    })
}

function App() {
  const { instance, inProgress } = useMsal()
  const [landingMode, setLandingMode] = useState<LandingMode>('choice')
  const [environmentUrl, setEnvironmentUrl] = useState(getStoredEnvironmentUrl)
  const [errorMessage, setErrorMessage] = useState('')
  const [environmentOrigin, setEnvironmentOrigin] = useState('')
  const [directLineTokenEndpoint, setDirectLineTokenEndpoint] = useState('')
  const [directLineErrorMessage, setDirectLineErrorMessage] = useState('')
  const [directLineSession, setDirectLineSession] = useState<DirectLineSession | null>(
    null,
  )
  const [directLineMessages, setDirectLineMessages] = useState<LiveChatMessage[]>([])
  const [directLineInput, setDirectLineInput] = useState('')
  const [isConnectingDirectLine, setIsConnectingDirectLine] = useState(false)
  const [isSendingDirectLine, setIsSendingDirectLine] = useState(false)
  const [directLineIsAwaitingAgent, setDirectLineIsAwaitingAgent] = useState(false)
  const directLineWatermarkRef = useRef<string | null>(null)
  const directLineIsPollingRef = useRef(false)
  const directLineUserActivityIdsRef = useRef<Set<string>>(new Set())
  const liveChatThreadRef = useRef<HTMLDivElement | null>(null)
  const [conversationTranscripts, setConversationTranscripts] = useState<
    DataverseRecord[]
  >([])
  const [bots, setBots] = useState<DataverseRecord[]>([])
  const [systemUsers, setSystemUsers] = useState<DataverseRecord[]>([])
  const [conversationTranscriptCount, setConversationTranscriptCount] =
    useState<number | null>(null)
  const [isLoadingDataverseData, setIsLoadingDataverseData] = useState(false)
  const [showBotPanel, setShowBotPanel] = useState(false)
  const [activeView, setActiveView] = useState<AppView>('dashboard')
  const [selectedBotId, setSelectedBotId] = useState<string | null>(null)
  const [selectedTranscriptId, setSelectedTranscriptId] = useState<string | null>(
    null,
  )
  const isLoggingIn = inProgress !== InteractionStatus.None
  const botSummaries = useMemo(
    () => buildBotSummaries(bots, conversationTranscripts),
    [bots, conversationTranscripts],
  )
  const makerSummaries = useMemo(
    () => buildMakerSummaries(botSummaries, systemUsers),
    [botSummaries, systemUsers],
  )
  const knowledgeSourceSummaries = useMemo(
    () => buildKnowledgeSourceSummaries(botSummaries, conversationTranscripts),
    [botSummaries, conversationTranscripts],
  )
  const feedbackSummaries = useMemo(
    () => buildFeedbackSummaries(botSummaries, conversationTranscripts),
    [botSummaries, conversationTranscripts],
  )
  const totalSessionCount = conversationTranscriptCount ?? conversationTranscripts.length
  const totalMakerCount = useMemo(
    () =>
      new Set(
        bots
          .map((bot) => getStringValue(bot, ['_createdby_value', 'createdby']))
          .filter((makerId): makerId is string => Boolean(makerId))
          .map(normalizeLookupKey),
      ).size,
    [bots],
  )
  const totalExchangedMessageCount = useMemo(
    () =>
      conversationTranscripts.reduce(
        (total, transcript) => total + getTranscriptMessageCount(transcript),
        0,
      ),
    [conversationTranscripts],
  )
  const totalKnowledgeSourceCount = knowledgeSourceSummaries.length
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

  const retrieveDirectLineActivities = useCallback(
    async (session: DirectLineSession) => {
      const activitiesResponse = await fetchDirectLineActivities(
        session,
        directLineWatermarkRef.current,
      )

      if (typeof activitiesResponse.watermark === 'string') {
        directLineWatermarkRef.current = activitiesResponse.watermark
      }

      const agentMessages: LiveChatMessage[] = []
      let hasAgentAnswer = false
      const activities = activitiesResponse.activities ?? []

      activities.forEach((activity, index) => {
        const clientActivityId = getDirectLineClientActivityId(activity)
        const isUserEcho =
          activity.from?.id === DIRECT_LINE_USER_ID ||
          activity.from?.role === 'user' ||
          (clientActivityId
            ? directLineUserActivityIdsRef.current.has(clientActivityId)
            : false)

        if (activity.type !== 'message' || isUserEcho) {
          return
        }

        const activityDate = getDirectLineActivityDate(activity)
        const activityId =
          typeof activity.id === 'string'
            ? activity.id
            : `direct-line-message-${Date.now()}-${index}`
        const displayName = activity.from?.name || 'Agent'
        const thoughtEntities = getDirectLineThoughtEntities(activity)
        const text = getDirectLineActivityText(activity)

        if (thoughtEntities.length === 0 && !text) {
          return
        }

        thoughtEntities.forEach((thought, thoughtIndex) => {
          const thoughtText = [thought.title, thought.text].filter(Boolean).join('\n')

          agentMessages.push({
            id: `${activityId}-thought-${thoughtIndex}`,
            author: 'agent',
            kind: 'thought',
            title: null,
            text: thoughtText,
            date: activityDate,
            displayName: 'Agent Thinking',
          })
        })

        if (!text) {
          return
        }

        hasAgentAnswer = true
        agentMessages.push({
          id: activityId,
          author: 'agent',
          kind: 'message',
          title: null,
          text,
          date: activityDate,
          displayName,
        })
      })

      if (hasAgentAnswer) {
        setDirectLineIsAwaitingAgent(false)
      }

      if (agentMessages.length > 0) {
        setDirectLineMessages((currentMessages) => {
          const currentMessageIds = new Set(
            currentMessages.map((message) => message.id),
          )
          const newMessages = agentMessages.filter(
            (message) => !currentMessageIds.has(message.id),
          )

          return newMessages.length > 0
            ? [...currentMessages, ...newMessages]
            : currentMessages
        })
      }

      return agentMessages.length
    },
    [],
  )

  useEffect(() => {
    const thread = liveChatThreadRef.current

    if (thread) {
      thread.scrollTop = thread.scrollHeight
    }
  }, [directLineMessages, directLineIsAwaitingAgent])

  useEffect(() => {
    if (!directLineSession || landingMode !== 'direct-line-chat') {
      return
    }

    let isCancelled = false

    async function pollActivities() {
      if (isCancelled || directLineIsPollingRef.current || !directLineSession) {
        return
      }

      directLineIsPollingRef.current = true

      try {
        await retrieveDirectLineActivities(directLineSession)
      } catch (error) {
        if (!isCancelled) {
          setDirectLineErrorMessage(getDirectLineErrorText(error))
          setDirectLineIsAwaitingAgent(false)
        }
      } finally {
        directLineIsPollingRef.current = false
      }
    }

    pollActivities()
    const pollingIntervalId = window.setInterval(pollActivities, 2500)

    return () => {
      isCancelled = true
      window.clearInterval(pollingIntervalId)
    }
  }, [directLineSession, landingMode, retrieveDirectLineActivities])

  function handleLandingModeChange(nextLandingMode: LandingMode) {
    setLandingMode(nextLandingMode)
    setErrorMessage('')
    setDirectLineErrorMessage('')
  }

  async function handleDirectLineConnect(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setDirectLineErrorMessage('')

    const trimmedTokenEndpoint = directLineTokenEndpoint.trim()

    if (!trimmedTokenEndpoint) {
      setDirectLineErrorMessage('Enter a Token Endpoint before starting the demo.')
      return
    }

    let parsedTokenEndpoint: string

    try {
      parsedTokenEndpoint = new URL(trimmedTokenEndpoint).toString()
    } catch {
      setDirectLineErrorMessage('Enter a valid Token Endpoint URL.')
      return
    }

    try {
      setIsConnectingDirectLine(true)
      directLineWatermarkRef.current = null
      directLineUserActivityIdsRef.current.clear()
      setDirectLineSession(null)
      setDirectLineMessages([])
      setDirectLineIsAwaitingAgent(false)
      setDirectLineTokenEndpoint(parsedTokenEndpoint)

      const tokenResponse = await requestDirectLineToken(parsedTokenEndpoint)
      const directLineToken = tokenResponse.token

      if (!directLineToken) {
        throw new Error('The token endpoint did not return a Direct Line token.')
      }

      const conversationResponse = await startDirectLineConversation(directLineToken)
      const conversationId =
        conversationResponse.conversationId ?? tokenResponse.conversationId

      if (!conversationId) {
        throw new Error('Direct Line did not return a conversation ID.')
      }

      const expiresIn = conversationResponse.expires_in ?? tokenResponse.expires_in

      setDirectLineSession({
        token: conversationResponse.token ?? directLineToken,
        conversationId,
        expiresAt:
          typeof expiresIn === 'number' && Number.isFinite(expiresIn)
            ? new Date(Date.now() + expiresIn * 1000)
            : null,
      })
      setLandingMode('direct-line-chat')
    } catch (error) {
      setDirectLineErrorMessage(getDirectLineErrorText(error))
    } finally {
      setIsConnectingDirectLine(false)
    }
  }

  async function handleDirectLineMessageSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setDirectLineErrorMessage('')

    const session = directLineSession
    const messageText = directLineInput.trim()

    if (!session) {
      setDirectLineErrorMessage('Start a Direct Line conversation before chatting.')
      return
    }

    if (!messageText) {
      setDirectLineErrorMessage('Enter a message before sending.')
      return
    }

    const userMessageId = `user-message-${Date.now()}`

    setDirectLineInput('')
    setIsSendingDirectLine(true)
    setDirectLineIsAwaitingAgent(true)
    directLineUserActivityIdsRef.current.add(userMessageId)
    setDirectLineMessages((currentMessages) => [
      ...currentMessages,
      {
        id: userMessageId,
        author: 'user',
        kind: 'message',
        title: null,
        text: messageText,
        date: new Date(),
        displayName: 'You',
      },
    ])

    try {
      await postDirectLineActivity(session, {
        type: 'message',
        text: messageText,
        from: {
          id: DIRECT_LINE_USER_ID,
          role: 'user',
        },
        channelData: {
          clientActivityId: userMessageId,
        },
      })
      await retrieveDirectLineActivities(session)
    } catch (error) {
      setDirectLineErrorMessage(getDirectLineErrorText(error))
      setDirectLineIsAwaitingAgent(false)
    } finally {
      setIsSendingDirectLine(false)
    }
  }

  async function handleLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setErrorMessage('')
    setShowBotPanel(false)
    setActiveView('dashboard')
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
      storeLastEnvironmentUrl(trimmedEnvironmentUrl)
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
      setActiveView('dashboard')
      setShowBotPanel(true)
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
    <main className={`landing-page${showBotPanel ? ' has-app-shell' : ''}`}>
      {!showBotPanel && (
        <p className="credit">
          Made by{' '}
          <a href="https://timh.ai" target="_blank" rel="noreferrer">
            Tim Hanewich
          </a>
        </p>
      )}
      {showBotPanel ? (
        <div className="app-shell">
          <aside className="app-sidebar" aria-label="Main navigation">
            <div className="app-brand">
              <strong>Copilot Studio Insights</strong>
            </div>
            <nav className="app-nav">
              {APP_NAV_ITEMS.map((item) => (
                <button
                  className={activeView === item.id ? 'is-active' : ''}
                  key={item.id}
                  type="button"
                  onClick={() => {
                    setActiveView(item.id)
                    setSelectedBotId(null)
                    setSelectedTranscriptId(null)
                  }}
                >
                  {item.label}
                </button>
              ))}
            </nav>
            <p className="app-sidebar-credit">
              Made by{' '}
              <a href="https://timh.ai" target="_blank" rel="noreferrer">
                Tim Hanewich
              </a>
            </p>
          </aside>
          <section className="app-content" aria-label="Application content">
            {activeView === 'dashboard' ? (
              <div className="dashboard-page">
                <div className="dashboard-heading">
                  <h1 id="dashboard-title">Copilot Studio Telemetry Dashboard</h1>
                  <p>{environmentOrigin}</p>
                </div>
                <div className="dashboard-metrics" aria-label="Environment totals">
                  <article>
                    <span>{bots.length.toLocaleString()}</span>
                    <strong>Total agents</strong>
                    <p>Agents found in this Dataverse environment.</p>
                  </article>
                  <article>
                    <span>{totalSessionCount.toLocaleString()}</span>
                    <strong>Total sessions</strong>
                    <p>ConversationTranscript records across all agents.</p>
                  </article>
                  <article>
                    <span>{totalMakerCount.toLocaleString()}</span>
                    <strong>Total makers</strong>
                    <p>Unique people who created agents in this environment.</p>
                  </article>
                  <article>
                    <span>{totalExchangedMessageCount.toLocaleString()}</span>
                    <strong>Total exchanged messages</strong>
                    <p>Messages exchanged across all recorded sessions.</p>
                  </article>
                  <article>
                    <span>{totalKnowledgeSourceCount.toLocaleString()}</span>
                    <strong>Knowledge sources used</strong>
                    <p>Unique knowledge sources referenced by all sessions.</p>
                  </article>
                </div>
              </div>
            ) : activeView === 'agents' && selectedBot && selectedTranscript ? (
            <>
              <section
                className="conversation-pane"
                aria-labelledby="conversation-title"
              >
                <div className="conversation-topbar">
                  <button
                    type="button"
                    className="back-button"
                    onClick={() => setSelectedTranscriptId(null)}
                  >
                    Back to agent
                  </button>
                  <div className="conversation-header">
                    {selectedBot.iconSource ? (
                      <img className="agent-icon" src={selectedBot.iconSource} alt="" />
                    ) : (
                      <span className="agent-icon bot-icon-fallback" aria-hidden="true">
                        {selectedBot.initials}
                      </span>
                    )}
                    <h1 id="conversation-title">
                      Recorded session · {selectedBot.name} ·{' '}
                      {formatDate(selectedTranscript.date)} ·{' '}
                      {selectedTranscript.messageCount.toLocaleString()} messages
                    </h1>
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
                          <dt>Knowledge sources used in this session</dt>
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
                            <span
                              className={`message-feedback-badge is-${message.feedback.reaction}`}
                              title={getFeedbackTooltip(message.feedback)}
                              aria-label={getFeedbackTooltip(message.feedback)}
                            >
                              {getFeedbackReactionIcon(message.feedback.reaction)}
                            </span>
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
            ) : activeView === 'agents' && selectedBot ? (
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
            ) : activeView === 'agents' ? (
            <>
              <div className="panel-heading">
                <h1 id="bot-panel-title">Agents in Your Environment</h1>
                <p className="environment-name">{environmentOrigin}</p>
              </div>
              <div className="panel-summary">
                <span>{bots.length.toLocaleString()} agents</span>
                <span>
                  {totalSessionCount.toLocaleString()}{' '}
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
            ) : activeView === 'makers' ? (
              <div className="makers-page">
                <div className="dashboard-heading">
                  <p className="eyebrow">Makers</p>
                  <h1>Agent makers</h1>
                  <p>
                    {makerSummaries.length.toLocaleString()} makers have created agents in
                    this environment.
                  </p>
                </div>
                <div className="maker-list" aria-label="Agent makers">
                  {makerSummaries.length > 0 ? (
                    makerSummaries.map((maker) => (
                      <article className="maker-row" key={maker.id}>
                        <div className="maker-profile">
                          <span className="maker-avatar" aria-hidden="true">
                            {getBotInitials(maker.name)}
                          </span>
                          <div>
                            <h2>{maker.name}</h2>
                            <p>
                              {maker.agents.length.toLocaleString()}{' '}
                              {maker.agents.length === 1 ? 'agent' : 'agents'} created
                            </p>
                          </div>
                        </div>
                        <div className="maker-agent-list">
                          {maker.agents.map((agent) => (
                            <button
                              className="maker-agent-pill"
                              key={agent.id}
                              type="button"
                              onClick={() => {
                                setActiveView('agents')
                                setSelectedBotId(agent.id)
                                setSelectedTranscriptId(null)
                              }}
                            >
                              {agent.iconSource ? (
                                <img src={agent.iconSource} alt="" />
                              ) : (
                                <span aria-hidden="true">{agent.initials}</span>
                              )}
                              <strong>{agent.name}</strong>
                              <small>
                                {agent.transcriptCount.toLocaleString()} sessions
                              </small>
                            </button>
                          ))}
                        </div>
                      </article>
                    ))
                  ) : (
                    <p className="empty-message">
                      No agent makers were found in this environment.
                    </p>
                  )}
                </div>
              </div>
            ) : activeView === 'knowledge' ? (
              <div className="knowledge-page">
                <div className="dashboard-heading">
                  <p className="eyebrow">Knowledge</p>
                  <h1>Knowledge sources used</h1>
                  <p>
                    {knowledgeSourceSummaries.length.toLocaleString()} sources used
                    across {bots.length.toLocaleString()} agents.
                  </p>
                </div>
                <div className="knowledge-list" aria-label="Knowledge source usage">
                  {knowledgeSourceSummaries.length > 0 ? (
                    knowledgeSourceSummaries.map((source) => (
                      <article className="knowledge-row" key={source.id}>
                        <div className="knowledge-source-main">
                          <span>{source.totalUsageCount.toLocaleString()}</span>
                          <div>
                            <h2>{source.name}</h2>
                            <p>
                              Used by {source.agents.length.toLocaleString()}{' '}
                              {source.agents.length === 1 ? 'agent' : 'agents'}
                            </p>
                          </div>
                        </div>
                        <div className="knowledge-agent-list">
                          {source.agents.map(({ agent, count }) => (
                            <button
                              className="knowledge-agent-row"
                              key={agent.id}
                              type="button"
                              onClick={() => {
                                setActiveView('agents')
                                setSelectedBotId(agent.id)
                                setSelectedTranscriptId(null)
                              }}
                            >
                              {agent.iconSource ? (
                                <img src={agent.iconSource} alt="" />
                              ) : (
                                <span aria-hidden="true">{agent.initials}</span>
                              )}
                              <strong>{agent.name}</strong>
                              <small>{count.toLocaleString()} uses</small>
                            </button>
                          ))}
                        </div>
                      </article>
                    ))
                  ) : (
                    <p className="empty-message">
                      No knowledge source usage was found in these transcripts.
                    </p>
                  )}
                </div>
              </div>
            ) : activeView === 'feedback' ? (
              <div className="feedback-page">
                <div className="dashboard-heading">
                  <p className="eyebrow">Feedback</p>
                  <h1>Collected feedback</h1>
                  <p>
                    {feedbackSummaries.length.toLocaleString()} feedback items across
                    all recorded sessions.
                  </p>
                </div>
                <div className="feedback-list" aria-label="Collected feedback">
                  {feedbackSummaries.length > 0 ? (
                    feedbackSummaries.map((item) => (
                      <button
                        className={`feedback-row is-${item.feedback.reaction}`}
                        key={item.id}
                        type="button"
                        onClick={() => {
                          setActiveView('agents')
                          setSelectedBotId(item.agent.id)
                          setSelectedTranscriptId(item.transcript.id)
                        }}
                      >
                        <span
                          className="feedback-reaction"
                          aria-label={getFeedbackReactionLabel(item.feedback.reaction)}
                        >
                          {getFeedbackReactionIcon(item.feedback.reaction)}
                        </span>
                        <span className="feedback-comment">
                          <strong>
                            {item.feedback.text || 'No comment provided'}
                          </strong>
                          <small>
                            {formatDate(item.message.date)}
                          </small>
                        </span>
                        <span className="feedback-agent">
                          {item.agent.iconSource ? (
                            <img src={item.agent.iconSource} alt="" />
                          ) : (
                            <span aria-hidden="true">{item.agent.initials}</span>
                          )}
                          <strong>{item.agent.name}</strong>
                        </span>
                      </button>
                    ))
                  ) : (
                    <p className="empty-message">
                      No message feedback was found in these transcripts.
                    </p>
                  )}
                </div>
              </div>
            ) : (
              <div className="placeholder-page">
                <p className="eyebrow">
                  {APP_NAV_ITEMS.find((item) => item.id === activeView)?.label}
                </p>
                <h1>
                  {APP_NAV_ITEMS.find((item) => item.id === activeView)?.label} insights
                </h1>
                <p>This section is ready for the next layer of analysis.</p>
              </div>
            )}
          </section>
        </div>
      ) : landingMode === 'choice' ? (
        <section className="landing-stage" aria-labelledby="page-title">
          <h1 id="page-title">Copilot Studio Insights</h1>
          <div className="choice-grid" aria-label="Copilot Studio app options">
            <button
              className="choice-card"
              type="button"
              onClick={() => handleLandingModeChange('telemetry')}
            >
              <strong>Review Copilot Studio Telemetry in my environment</strong>
            </button>
            <button
              className="choice-card"
              type="button"
              onClick={() => handleLandingModeChange('direct-line-setup')}
            >
              <strong>Demo Interface with an agent via Direct Line API</strong>
            </button>
          </div>
        </section>
      ) : landingMode === 'telemetry' ? (
        <form
          className="login-card"
          aria-labelledby="telemetry-page-title"
          onSubmit={handleLogin}
        >
          <h1 id="telemetry-page-title">Copilot Studio Insights</h1>
          <input
            id="environment-url"
            name="environmentUrl"
            type="url"
            placeholder="Your environment URL (e.g. https://your-environment.crm.dynamics.com)"
            value={environmentUrl}
            onChange={(event) => setEnvironmentUrl(event.target.value)}
            required
          />
          <button type="submit" disabled={isLoggingIn || isLoadingDataverseData}>
            {isLoggingIn || isLoadingDataverseData ? 'Loading...' : 'Login'}
          </button>
          {isLoadingDataverseData && (
            <p className="loading-message">Loading Dataverse data...</p>
          )}
          {errorMessage && <p className="error-message">{errorMessage}</p>}
        </form>
      ) : landingMode === 'direct-line-setup' ? (
        <form
          className="login-card direct-line-setup-card"
          aria-labelledby="direct-line-page-title"
          onSubmit={handleDirectLineConnect}
        >
          <p className="eyebrow">Direct Line API</p>
          <h1 id="direct-line-page-title">Chat with your agent.</h1>
          <p className="login-card-copy">
            Paste the Token Endpoint from your published Copilot Studio agent.
            The app will request a Direct Line token, start the conversation, and
            poll for agent responses.
          </p>
          <input
            id="direct-line-token-endpoint"
            name="directLineTokenEndpoint"
            type="url"
            placeholder="Token Endpoint"
            value={directLineTokenEndpoint}
            onChange={(event) => setDirectLineTokenEndpoint(event.target.value)}
            required
          />
          <button type="submit" disabled={isConnectingDirectLine}>
            {isConnectingDirectLine ? 'Starting chat...' : 'Start chat'}
          </button>
          {directLineErrorMessage && (
            <p className="error-message">{directLineErrorMessage}</p>
          )}
        </form>
      ) : (
        <section className="live-chat-shell" aria-labelledby="live-chat-title">
          <div className="live-chat-header">
            <div>
              <h1 id="live-chat-title">Direct Line API Demo Interface</h1>
            </div>
          </div>
          <div className="live-chat-panel">
            <div
              className="chat-thread live-chat-thread"
              ref={liveChatThreadRef}
              aria-label="Live Direct Line conversation"
            >
              {directLineMessages.length > 0 ? (
                directLineMessages.map((message) => (
                  <article
                    className={`chat-message is-${message.author} is-${message.kind}`}
                    key={message.id}
                  >
                    <span className="chat-author">{message.displayName}</span>
                    <p>{message.text}</p>
                    {message.date && (
                      <time dateTime={message.date.toISOString()}>
                        {formatDate(message.date)}
                      </time>
                    )}
                  </article>
                ))
              ) : (
                <p className="empty-message">
                  Send a message to begin the Direct Line conversation.
                </p>
              )}
              {directLineIsAwaitingAgent && (
                <article className="chat-message is-agent is-typing">
                  <span className="chat-author">Agent</span>
                  <p>Thinking...</p>
                </article>
              )}
            </div>
            <form className="live-chat-form" onSubmit={handleDirectLineMessageSubmit}>
              <input
                aria-label="Message"
                type="text"
                placeholder="Ask your agent anything..."
                value={directLineInput}
                onChange={(event) => setDirectLineInput(event.target.value)}
              />
              <button
                type="submit"
                disabled={!directLineInput.trim() || isSendingDirectLine}
              >
                {isSendingDirectLine ? 'Sending...' : 'Send'}
              </button>
            </form>
            {directLineErrorMessage && (
              <p className="error-message">{directLineErrorMessage}</p>
            )}
          </div>
        </section>
      )}
    </main>
  )
}

export default App
