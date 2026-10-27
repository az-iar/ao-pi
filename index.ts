// Agent Office extension for Pi / PiG (https://github.com/az-iar/agent-office).
//
//   /office [name]     connects: OAuth in the browser, Azri approves, the token
//                      is stored under the agent name, and the Office MCP tools
//                      appear in Pi. The name pre-fills a new agent.
//   while running      `working` while this Pi session streams or a dispatched
//                      worker runs, a heartbeat every 90 s, `idle` once both
//                      settle; `idle` never overrides `blocked`. New assigned
//                      tasks go to a worker Pi in a worktree, up to two tasks
//                      at a time per Office agent. Headless children do not
//                      poll, dispatch, or report status.
//
//   AGENT_OFFICE_URL                dashboard origin, default https://office.byma.tech
//   AGENT_OFFICE_TOKEN              an agent API token instead of OAuth (optional)
//   AGENT_OFFICE_PI_CREDENTIALS     exact credentials file, overrides agent-scoped storage
//   AGENT_OFFICE_PI_BIN             the pig/pi executable for dispatch, default `pig`
//   AGENT_OFFICE_CALLBACK_PORT      loopback port for the OAuth callback, default 41997
//   AGENT_OFFICE_POLL_MS            inbox poll interval, default 60000
//   AGENT_OFFICE_AUTODISPATCH       on (default) or off
//   AGENT_OFFICE_CLIENT_NAME        OAuth client name; defaults to folder@host
//
// Tokens are never logged.

import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent"
import { Type } from "@earendil-works/pi-ai"
import { spawn } from "node:child_process"
import { createHash, randomBytes } from "node:crypto"
import { existsSync, unlinkSync } from "node:fs"
import { createServer } from "node:http"
import { hostname } from "node:os"
import { basename, dirname, join, sep } from "node:path"
import { fileURLToPath } from "node:url"
import {
  type Credentials,
  acquireDispatch,
  addWorktree,
  agentCredentialsPath,
  agentStateDir,
  CREDENTIALS_OVERRIDE,
  claimNotice,
  claimTask,
  createBearer,
  createClient,
  cronLine,
  fromOthers,
  inboxPrompt,
  inHerdrWorktree,
  itemKey,
  loadAgentPointer,
  loadCredentials,
  MAX_DISPATCHES,
  OFFICE,
  onTaskBranch,
  REFRESH_EARLY_MS,
  releaseDispatch,
  releaseNotice,
  resolveCredentials,
  saveAgentPointers,
  saveCredentials,
  supervisorPrompt,
  taskID,
  tokenRequest,
  workspaceKey,
} from "./lib/office.ts"

const configuredCallbackPort = Number(process.env.AGENT_OFFICE_CALLBACK_PORT ?? 41997)
const CALLBACK_BASE_PORT =
  Number.isInteger(configuredCallbackPort) && configuredCallbackPort >= 1 && configuredCallbackPort <= 65535 ? configuredCallbackPort : 41997
const CALLBACK_TRIES = 20
const HERE = dirname(fileURLToPath(import.meta.url))
const TOKEN_SCRIPT = join(HERE, "lib", "token.mts")
const POLL_SCRIPT = join(HERE, "poll.ts")
const PI_BIN = process.env.AGENT_OFFICE_PI_BIN || "pig"
const callbackUrl = (port: number) => `http://127.0.0.1:${port}/callback`
const SCOPES = "office:read office:write office:events"
const HEARTBEAT_MS = 90_000
const APPROVAL_TIMEOUT_MS = 5 * 60_000
const DEFAULT_POLL_MS = 60_000
const DISPATCH_STALE_MS = 6 * 60 * 60_000

type PendingCallback = { resolve: (code: string) => void; reject: (err: Error) => void; close: () => void; cancel: (err: Error) => void }

const escapeHtml = (value: string) =>
  value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char] ?? char)

// A pasted callback URL can finish an approval without the browser reaching
// the loopback listener, which matters on remote machines and headless hosts.
const pendingCallbacks = new Map<string, PendingCallback>()

function completeCallback(raw: string): { handled: boolean; ok: boolean; message: string } {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return { handled: false, ok: false, message: "that is not a valid URL" }
  }
  const state = url.searchParams.get("state")
  if (!state) return { handled: false, ok: false, message: "the URL has no state" }
  const pending = pendingCallbacks.get(state)
  if (!pending) return { handled: false, ok: false, message: "no pending approval matches this URL; run /office again" }

  const code = url.searchParams.get("code")
  const error = url.searchParams.get("error_description") ?? url.searchParams.get("error")
  pendingCallbacks.delete(state)
  pending.close()
  if (error || !code) {
    pending.reject(new Error(`Office authorization failed: ${error ?? "no code"}`))
    return { handled: true, ok: false, message: `approval failed: ${error ?? "no code"}` }
  }
  pending.resolve(code)
  return { handled: true, ok: true, message: "callback accepted; finishing the Office connection" }
}

const MAX_CHILD_OUTPUT = 64_000
function run(command: string, args: string[], cwd?: string, timeoutMs = 30_000, env: NodeJS.ProcessEnv = process.env): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, env, stdio: ["ignore", "pipe", "pipe"] })
    let stdout = ""
    let stderr = ""
    const keep = (current: string, chunk: unknown) =>
      current.length >= MAX_CHILD_OUTPUT ? current : current + String(chunk).slice(0, MAX_CHILD_OUTPUT - current.length)
    let timer: ReturnType<typeof setTimeout> | undefined
    if (timeoutMs > 0) {
      timer = setTimeout(() => {
        child.kill("SIGTERM")
        setTimeout(() => child.kill("SIGKILL"), 5_000).unref()
      }, timeoutMs)
      timer.unref()
    }
    child.stdout.on("data", (chunk) => (stdout = keep(stdout, chunk)))
    child.stderr.on("data", (chunk) => (stderr = keep(stderr, chunk)))
    child.on("error", (err) => {
      clearTimeout(timer)
      reject(err)
    })
    child.on("close", (code) => {
      clearTimeout(timer)
      resolve({ code, stdout, stderr })
    })
  })
}

async function registerClient(callback: string, clientName: string): Promise<string> {
  const res = await fetch(`${OFFICE}/oauth/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      client_name: clientName,
      redirect_uris: [callback],
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      token_endpoint_auth_method: "none",
      scope: SCOPES,
    }),
    signal: AbortSignal.timeout(15_000),
  })
  const body = (await res.json()) as { client_id: string; error_description?: string; error?: string }
  if (!res.ok) throw new Error(`Office client registration failed: ${body.error_description ?? body.error ?? res.status}`)
  return body.client_id
}

function openBrowser(url: string) {
  try {
    const [cmd, args] = process.platform === "win32" ? ["cmd", ["/c", "start", "", url]] : process.platform === "darwin" ? ["open", [url]] : ["xdg-open", [url]]
    spawn(cmd, args, { detached: true, stdio: "ignore" }).on("error", () => {}).unref()
  } catch {}
}

// Binds a loopback callback for this state, falling back to the next free port
// when another Pi extension instance already holds the preferred one.
function listenForCode(state: string): Promise<{ port: number; code: Promise<string>; close: () => void }> {
  return new Promise((resolve, reject) => {
    const attempt = (offset: number) => {
      if (offset >= CALLBACK_TRIES) {
        reject(new Error(`cannot listen on ports ${CALLBACK_BASE_PORT}-${CALLBACK_BASE_PORT + CALLBACK_TRIES - 1}`))
        return
      }
      const port = CALLBACK_BASE_PORT + offset
      const callback = callbackUrl(port)
      let resolveCode!: (code: string) => void
      let rejectCode!: (err: Error) => void
      const code = new Promise<string>((resolve, reject) => {
        resolveCode = resolve
        rejectCode = reject
      })
      let settled = false
      let timer: ReturnType<typeof setTimeout> | undefined
      const server = createServer((req, res) => {
        const result = completeCallback(`http://127.0.0.1:${port}${req.url ?? "/"}`)
        if (!result.handled) {
          res.writeHead(404).end()
          return
        }
        res.writeHead(result.ok ? 200 : 400, { "Content-Type": "text/html; charset=utf-8" })
        res.end(`<!doctype html><title>Agent Office</title><p>${result.ok ? "Pi is connected to Agent Office. You can close this tab." : `Not connected: ${escapeHtml(result.message)}`}</p>`)
      })
      const cancel = (err: Error) => {
        if (settled) return
        settled = true
        pendingCallbacks.delete(state)
        clearTimeout(timer)
        server.close()
        rejectCode(err)
      }
      server.once("error", (err: NodeJS.ErrnoException) => {
        if (!server.listening && err.code === "EADDRINUSE") {
          attempt(offset + 1)
          return
        }
        cancel(new Error(`cannot listen on ${callback}: ${err.message}`))
      })
      server.listen(port, "127.0.0.1", () => {
        if (settled) return
        timer = setTimeout(() => cancel(new Error("timed out waiting for the approval in the browser")), APPROVAL_TIMEOUT_MS)
        pendingCallbacks.set(state, {
          resolve: resolveCode,
          reject: rejectCode,
          close: () => {
            settled = true
            clearTimeout(timer)
            server.close()
          },
          cancel,
        })
        resolve({ port, code, close: () => cancel(new Error("approval cancelled")) })
      })
    }
    attempt(0)
  })
}

// Starts an approval and returns its URL at once. The tokens follow when the
// browser reaches the loopback callback or the callback URL is pasted back,
// so a remote machine never blocks on a browser it cannot open.
async function authorize(clientName: string): Promise<{ url: string; state: string; credentials: Promise<Credentials>; close: () => void }> {
  const verifier = randomBytes(32).toString("base64url")
  const state = randomBytes(16).toString("base64url")
  const listener = await listenForCode(state)
  const callback = callbackUrl(listener.port)
  let client_id: string
  try {
    client_id = await registerClient(callback, clientName)
  } catch (err) {
    listener.close()
    throw err
  }
  const url = new URL(`${OFFICE}/oauth/authorize`)
  url.search = new URLSearchParams({
    response_type: "code",
    client_id,
    redirect_uri: callback,
    scope: SCOPES,
    state,
    code_challenge: createHash("sha256").update(verifier).digest("base64url"),
    code_challenge_method: "S256",
    resource: `${OFFICE}/mcp`,
  }).toString()
  openBrowser(url.href)
  const credentials = listener.code.then((code) =>
    tokenRequest(client_id, { grant_type: "authorization_code", code, redirect_uri: callback, code_verifier: verifier }),
  )
  return { url: url.href, state, credentials, close: listener.close }
}

// A pasted callback, tolerating the quotes or angle brackets a terminal or
// chat client may add around it.
const pastedCallback = (text: string) => {
  const url = text.trim().replace(/^[<"'`]+|[>"'`]+$/g, "")
  return /^http:\/\/(127\.0\.0\.1|localhost):\d+\/callback\?\S*$/.test(url) ? url : undefined
}

function approvalText(name: string, url: string) {
  return [
    `Open this link and approve "${name}" as the Office admin:`,
    url,
    "",
    "On this machine the browser opens it by itself and the connection finishes on its own.",
    "On a remote machine, open the link in any browser. After you approve, the browser lands on a page that does not load: copy that page's full address (http://127.0.0.1:.../callback?...) and paste it here, or run /office followed by the address.",
    `The link works for ${APPROVAL_TIMEOUT_MS / 60_000} minutes.`,
  ].join("\n")
}

// A single quoted shell word for the `!command` MCP Authorization header.
const shq = (value: string) => `'${value.replaceAll("'", `'\\''`)}'`

export default function agentOffice(pi: ExtensionAPI) {
  const warn = (message: string, extra?: Record<string, unknown>) => {
    console.warn(`[agent-office] ${message}${extra ? ` ${JSON.stringify(extra)}` : ""}`)
  }

  const directory = process.cwd()
  const canonical = canonicalRoot(directory)
  const locationKey = workspaceKey(directory)
  const clientName = process.env.AGENT_OFFICE_CLIENT_NAME ?? `${basename(directory) || "Pi"}@${hostname().split(".")[0]}`

  const resolved = resolveCredentials(directory, canonical)
  const { legacyCredentialsFile, pointerFile, repoPointerFile } = resolved
  let credentialsFile = resolved.credentialsFile
  const usingCredentialsOverride = Boolean(CREDENTIALS_OVERRIDE)
  const auth = createBearer(() => credentialsFile, resolved.credentials)
  const { call, checkin } = createClient(auth.bearer)

  const sameCredentials = (a: Credentials, b: Credentials) =>
    a.office === b.office && a.client_id === b.client_id && a.access_token === b.access_token && a.refresh_token === b.refresh_token && a.expires_at === b.expires_at
  const adoptAgent = (slug?: string) => {
    if (usingCredentialsOverride || !slug) return
    const target = agentCredentialsPath(slug)
    if (auth.credentials) {
      const current = loadCredentials(target)
      if (!current || !sameCredentials(current, auth.credentials)) saveCredentials(target, auth.credentials)
    }
    credentialsFile = target
    saveAgentPointers(directory, canonical, slug)
  }
  const dispatchDir = () => agentStateDir(legacyCredentialsFile, "dispatch", ownAgent?.slug ?? locationKey)
  const noticeDir = () => agentStateDir(legacyCredentialsFile, "notices", ownAgent?.slug ?? locationKey)
  const configuredDispatch = String(process.env.AGENT_OFFICE_AUTODISPATCH ?? "on").trim().toLowerCase()
  const dispatchMode = configuredDispatch === "off" || configuredDispatch === "false" ? "off" : "session"

  // The office MCP server authenticates through a self-refreshing bearer
  // command: Pi resolves it once when it creates the server transport, so the
  // registration is refreshed before the access token expires (a static header
  // on a live connection is not re-run per request).
  let mcpRegistered = false
  let lastMcpError: string | undefined
  let mcpTimer: ReturnType<typeof setTimeout> | undefined
  let mcpRefreshAt: number | undefined
  // Pi can run as a standalone binary whose process.execPath is not a JS
  // runtime, so fall back to node or bun from PATH for the token script. The
  // result is absolute: Pi resolves the header once, and a bare name would
  // depend on the PATH of whichever shell ends up resolving it.
  const whichOnPath = (name: string) => {
    for (const dir of (process.env.PATH ?? "").split(":")) {
      if (!dir) continue
      const candidate = join(dir, name)
      if (existsSync(candidate)) return candidate
    }
    return undefined
  }
  const tokenRuntime = () => {
    const name = basename(process.execPath)
    if (name === "node" || name === "bun") return process.execPath
    if (TOKEN_SCRIPT.includes(`${sep}node_modules${sep}`)) {
      const bun = whichOnPath("bun")
      if (bun) return bun
    }
    return whichOnPath("node") ?? whichOnPath("bun") ?? process.execPath
  }
  const bearerCommand = () => `!${shq(tokenRuntime())} ${shq(TOKEN_SCRIPT)} --dir ${shq(directory)}`
  const clearMcpTimer = () => {
    if (mcpTimer) clearTimeout(mcpTimer)
    mcpTimer = undefined
    mcpRefreshAt = undefined
  }
  const scheduleMcpRefresh = () => {
    clearMcpTimer()
    if (process.env.AGENT_OFFICE_TOKEN || !auth.credentials) return
    const expiresAt = auth.credentials.expires_at
    if (!Number.isFinite(expiresAt)) return
    const delay = Math.max(expiresAt - Date.now() - REFRESH_EARLY_MS, 10_000)
    mcpRefreshAt = Date.now() + delay
    mcpTimer = setTimeout(() => void connectMcp(true).catch((err) => warn("could not refresh the office MCP connection", { error: String(err) })), delay)
    mcpTimer.unref()
  }
  const connectMcp = async (force = false) => {
    if (mcpRegistered && !force) return
    let token: string | undefined
    try {
      token = await auth.bearer()
    } catch (err) {
      lastMcpError = String(err)
      warn("could not read the Office token", { error: String(err) })
      return
    }
    if (!token) return
    try {
      if (typeof pi.registerMcpServer !== "function") throw new Error("this Pi build has no registerMcpServer; Pi 0.99 or newer is required")
      if (force && mcpRegistered) {
        mcpRegistered = false
        if (typeof pi.unregisterMcpServer === "function") {
          try {
            pi.unregisterMcpServer("office")
          } catch {}
        }
      }
      pi.registerMcpServer("office", {
        url: `${OFFICE}/mcp`,
        headers: { Authorization: bearerCommand() },
        description: "Agent Office: check in, read tasks and messages, post journal entries and decisions",
      })
      mcpRegistered = true
      lastMcpError = undefined
      scheduleMcpRefresh()
    } catch (err) {
      lastMcpError = String(err)
      warn("could not register the office MCP server", { error: String(err) })
    }
  }

  const note = `pi: ${basename(directory)}`
  let streaming = false
  let working = false
  let heartbeat: ReturnType<typeof setInterval> | undefined
  let pollTimer: ReturnType<typeof setInterval> | undefined
  let queue = Promise.resolve()
  let lastPollAt: number | undefined
  let lastPollError: string | undefined
  let lastPromptAt: number | undefined
  let lastPending: Record<string, number> | undefined
  let ownAgent: { id?: string; slug?: string } | undefined
  let lastCtx: ExtensionContext | undefined
  const backlog = new Set<string>()
  const dispatching = new Set<string>()
  const assignedKeys = new Map<string, string>()
  const repo = canonical ?? directory
  const isChild = process.env.AGENT_OFFICE_CHILD === "1"

  const notice = (text: string, level: "info" | "warning" | "error" = "info") => {
    try {
      lastCtx?.ui.notify(text, level)
    } catch {}
  }

  // Wake the model with a user turn: now when idle, queued as a follow-up when
  // it is already streaming.
  const wake = (text: string) => {
    const idle = lastCtx ? lastCtx.isIdle() : !streaming
    try {
      pi.sendUserMessage(text, idle ? undefined : { deliverAs: "followUp" })
      lastPromptAt = Date.now()
      return true
    } catch (err) {
      warn("could not wake the session", { error: String(err) })
      return false
    }
  }

  // Reports run one after another so a short turn cannot land idle before working.
  const report = (status: "working" | "idle") => {
    queue = queue.then(async () => {
      try {
        if (!(await auth.bearer())) return
        if (status === "idle" && (await checkin()).agent?.work_status === "blocked") return
        await checkin({ status, note })
      } catch (err) {
        warn(`could not report ${status}`, { error: String(err) })
      }
    })
    return queue
  }

  // A dispatch worker is activity too: keep the agent working until both this
  // session and its workers settle. Child processes never report status, so a
  // worker cannot flip the shared agent idle under the parent session.
  const refreshWorking = () => {
    const next = !isChild && (streaming || dispatching.size > 0)
    if (next === working) return
    working = next
    if (next) {
      heartbeat = setInterval(() => checkin().catch(() => {}), HEARTBEAT_MS)
      heartbeat.unref()
      void report("working")
    } else {
      if (heartbeat) clearInterval(heartbeat)
      heartbeat = undefined
      void report("idle")
    }
  }
  const startWorking = () => {
    if (streaming) return
    streaming = true
    refreshWorking()
  }
  const stopWorking = () => {
    if (!streaming) return
    streaming = false
    refreshWorking()
  }

  const connectedSummary = async (reconnectMcp = false) => {
    // Check in before saving under an agent name: the approval token is in
    // memory (and a pending file), so a switch to a different agent cannot
    // overwrite the previous agent's credentials.
    const me = await checkin()
    adoptAgent(me.agent?.slug)
    ownAgent = me.agent
    await connectMcp(reconnectMcp)
    await observePending(me.pending, false)
    const pending = me.pending ?? {}
    const mcpLine = mcpRegistered
      ? "The office MCP tools are available; call checkin first."
      : `The office MCP server is not registered${lastMcpError ? `: ${lastMcpError}` : ""}.`
    const summaryLine = `Connected to Agent Office as @${me.agent?.slug} (${me.agent?.name}). Open tasks: ${pending.open_tasks?.length ?? 0}, unclaimed team tasks: ${pending.unclaimed_team_tasks?.length ?? 0}, unread messages: ${pending.unread_messages?.length ?? 0}. ${mcpLine}`
    if (!existsSync(POLL_SCRIPT)) return summaryLine
    return `${summaryLine}\nTo keep handling Office work while Pi is closed, add this line to the crontab (crontab -e), or ask to have it added:\n${cronLine(directory, POLL_SCRIPT)}`
  }

  // The approval in flight. A loopback callback finishes it in the background
  // and tells the session; a pasted URL reports it in reply.
  let approval: { state: string; done: Promise<string>; pasted: boolean; cancel: () => void } | undefined
  const startApproval = async (approvalName: string) => {
    approval?.cancel()
    const started = await authorize(approvalName)
    const current = {
      state: started.state,
      cancel: started.close,
      pasted: false,
      done: started.credentials.then(async (next) => {
        auth.credentials = next
        const previousFile = credentialsFile
        const pendingFile = usingCredentialsOverride ? undefined : join(dirname(legacyCredentialsFile), `pending-${started.state}.json`)
        if (pendingFile) {
          saveCredentials(pendingFile, next)
          credentialsFile = pendingFile
        } else {
          saveCredentials(credentialsFile, next)
        }
        try {
          return await connectedSummary(true)
        } finally {
          if (pendingFile) {
            try {
              unlinkSync(pendingFile)
            } catch {}
            if (credentialsFile === pendingFile) credentialsFile = previousFile
          }
        }
      }),
    }
    approval = current
    current.done.then(
      (summaryText) => current.pasted || wake(`Agent Office connected.\n${summaryText}`),
      (err) => current.pasted || notice(`Agent Office approval did not finish: ${err instanceof Error ? err.message : String(err)}. Run /office again.`, "error"),
    )
    return started.url
  }
  const finishWithCallback = async (raw: string) => {
    let state = ""
    try {
      state = new URL(raw.trim()).searchParams.get("state") ?? ""
    } catch {}
    const current = approval?.state === state ? approval : undefined
    if (current) current.pasted = true
    const result = completeCallback(raw.trim())
    if (!result.handled || !result.ok) {
      if (current) current.pasted = false
      return result.message
    }
    if (!current) return `${result.message} in the Pi that started it.`
    try {
      return await current.done
    } catch (err) {
      return `Agent Office approval did not finish: ${err instanceof Error ? err.message : String(err)}. Run /office again.`
    }
  }

  // Auto-dispatch: new assigned todo tasks run in their own worktree as a
  // headless worker Pi, up to MAX_DISPATCHES at once across every process.
  const dispatchTask = async (id: string) => {
    const assignedKey = assignedKeys.get(id)
    let workerStarted = false
    try {
      const detail = await call("get_task", { task_id: id }).catch(() => undefined)
      const task = (detail as any)?.task ?? detail
      await claimTask(call, id)
      const { path, branch } = await addWorktree(repo, id)
      notice(`Agent Office: dispatching task ${id} to a worker in ${path} (${branch}).`)
      const prompt = supervisorPrompt(id, task, path, branch)
      const result = await run(PI_BIN, ["-e", HERE, "--print", prompt], path, DISPATCH_STALE_MS, { ...process.env, AGENT_OFFICE_CHILD: "1" }).catch((err) => {
        warn("could not start the worker", { task: id, error: String(err) })
        return undefined
      })
      if (result === undefined) {
        // The worker never started: a missing binary or environment problem is
        // retryable, unlike a worker that ran and exited non-zero.
        if (assignedKey) seenAssigned.delete(assignedKey)
        return
      }
      workerStarted = true
      notice(`Agent Office task ${id} worker exited (${result?.code ?? "error"}). Review ${path}; it opens the PR and updates the task.`)
    } catch (err) {
      if (!workerStarted && assignedKey) seenAssigned.delete(assignedKey)
      warn("could not dispatch Office task", { task: id, error: String(err) })
      notice(`Agent Office: could not dispatch task ${id}: ${String(err)}`, "error")
    } finally {
      releaseDispatch(dispatchDir(), id)
      dispatching.delete(id)
      assignedKeys.delete(id)
      pumpDispatch()
      refreshWorking()
    }
  }

  const pumpDispatch = () => {
    if (dispatchMode === "off") return
    for (const id of [...backlog]) {
      if (dispatching.size >= MAX_DISPATCHES) return
      const acquired = acquireDispatch(dispatchDir(), id)
      if (acquired === "full") return
      backlog.delete(id)
      if (acquired === "busy") {
        assignedKeys.delete(id)
        continue
      }
      dispatching.add(id)
      refreshWorking()
      void dispatchTask(id)
    }
  }

  const seenMessages = new Set<string>()
  const seenTasks = new Set<string>()
  const seenAssigned = new Set<string>()
  let pollPrimed = false

  // Polling reports only new items. The first successful checkin is a
  // baseline: existing unread mail is already visible through checkin.
  const observePending = async (pending: any, live: boolean) => {
    const messages: any[] = (Array.isArray(pending?.unread_messages) ? pending.unread_messages : []).filter((item: any) => fromOthers(item, ownAgent))
    const tasks: any[] = Array.isArray(pending?.unclaimed_team_tasks) ? pending.unclaimed_team_tasks : []
    const assigned: any[] = Array.isArray(pending?.open_tasks) ? pending.open_tasks : []
    const freshMessages = messages.filter((item) => !seenMessages.has(itemKey(item, "message")))
    const freshTasks = tasks.filter((item) => !seenTasks.has(itemKey(item, "task")))
    const freshAssigned = assigned.filter((item) => !seenAssigned.has(itemKey(item, "task")))
    lastPending = { messages: messages.length, unclaimed: tasks.length, assigned: assigned.length }

    if (!pollPrimed) {
      for (const item of tasks) seenTasks.add(itemKey(item, "task"))
      for (const item of assigned) seenAssigned.add(itemKey(item, "task"))
      freshTasks.length = 0
      freshAssigned.length = 0
      pollPrimed = true
    }
    if (!live || (freshMessages.length === 0 && freshTasks.length === 0 && freshAssigned.length === 0)) return

    for (const item of tasks) seenTasks.add(itemKey(item, "task"))
    for (const item of assigned) seenAssigned.add(itemKey(item, "task"))

    for (const item of freshAssigned) {
      const id = taskID(item)
      if (String(item?.status ?? "todo").toLowerCase() === "todo" && dispatchMode !== "off") {
        assignedKeys.set(id, itemKey(item, "task"))
        backlog.add(id)
      }
    }

    if (freshMessages.length > 0) {
      const ids = freshMessages.map((item) => itemKey(item, "message"))
      if (claimNotice(noticeDir(), ids)) {
        if (wake(inboxPrompt(freshMessages))) {
          for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
        } else {
          releaseNotice(noticeDir(), ids)
        }
      }
      // A failed claim means another process owns the batch. Leave it unseen:
      // if that process dies, the claim expires and this poll retries it.
    }

    if (freshTasks.length > 0 || freshAssigned.length > 0) {
      notice(`Agent Office poll: ${freshTasks.length} new unclaimed, ${freshAssigned.length} new assigned task(s). Use list_tasks in the office MCP server.`)
    }
    if (backlog.size > 0) pumpDispatch()
  }

  const poll = async () => {
    try {
      if (!(await auth.bearer())) return
      await connectMcp()
      const me = await checkin()
      adoptAgent(me.agent?.slug)
      ownAgent = me.agent
      lastPollAt = Date.now()
      lastPollError = undefined
      await observePending(me.pending, true)
    } catch (err) {
      lastPollError = String(err)
      warn("could not poll the Office inbox", { error: String(err) })
    }
  }

  const configuredPoll = process.env.AGENT_OFFICE_POLL_MS?.trim()
  const configuredPollNumber = configuredPoll ? Number(configuredPoll) : NaN
  const pollInterval = Number.isFinite(configuredPollNumber) ? Math.max(configuredPollNumber, 1_000) : DEFAULT_POLL_MS
  const pollingAllowed = !isChild && !onTaskBranch(directory) && !inHerdrWorktree(directory)

  // --- status reporting on this session's own streaming ---
  pi.on("session_start", async (_event, ctx) => {
    lastCtx = ctx
    await connectMcp()
    if (pollingAllowed && !pollTimer) {
      pollTimer = setInterval(() => void poll(), pollInterval)
      pollTimer.unref()
      void poll()
    }
  })
  pi.on("agent_start", async (_event, ctx) => {
    lastCtx = ctx
    startWorking()
  })
  pi.on("agent_settled", async (_event, ctx) => {
    lastCtx = ctx
    stopWorking()
  })
  pi.on("session_shutdown", async () => {
    clearInterval(heartbeat)
    heartbeat = undefined
    clearInterval(pollTimer)
    pollTimer = undefined
    clearMcpTimer()
    approval?.cancel()
    if (working) {
      working = false
      streaming = false
      await report("idle")
    }
  })

  // --- /office command ---
  pi.registerCommand("office", {
    description: "Connect to Agent Office: register an agent, check status, or reconnect",
    handler: async (args, ctx) => {
      lastCtx = ctx
      const input = args.trim()
      const callback = pastedCallback(input)
      if (callback) {
        ctx.ui.notify(await finishWithCallback(callback), "info")
        return
      }
      const [verb, ...rest] = input.split(/\s+/)
      const argument = rest.join(" ").trim()
      try {
        if (verb === "status") {
          const report = statusReport()
          ctx.ui.notify(report, "info")
          wake(`Agent Office status:\n${report}`)
          return
        }
        if (!input && (await auth.bearer())) {
          ctx.ui.notify("Checking Agent Office…", "info")
          wake(await connectedSummary())
          return
        }
        const name = verb === "register" ? argument : input && verb !== "reconnect" ? input : ""
        const approvalName = name ? name.slice(0, 80) : clientName
        const url = await startApproval(approvalName)
        ctx.ui.notify("Agent Office approval started; see the link in the chat.", "info")
        if (!wake(approvalText(approvalName, url))) ctx.ui.notify(`Open this link and approve: ${url}`, "info")
      } catch (err) {
        ctx.ui.notify(`Agent Office: ${err instanceof Error ? err.message : String(err)}`, "error")
      }
    },
  })

  // --- tools the model can call directly ---
  pi.registerTool({
    name: "office_connect",
    label: "Office Connect",
    description:
      "Connect Pi to Agent Office. Starts an approval and returns its link at once; show the user the link. When the Office admin approves, the tokens are stored, the office MCP tools (checkin, tasks, messages, journal, decisions) are added and the session is told. Already connected: reports who you are. `reconnect` starts a fresh approval. `agent_name` pre-fills a new Office agent name. When the user pastes a http://127.0.0.1:.../callback?... URL, pass it as `callback_url`.",
    parameters: Type.Object({
      reconnect: Type.Optional(Type.Boolean({ description: "force a new OAuth approval" })),
      agent_name: Type.Optional(Type.String({ description: "Office agent name to pre-fill when approving a new agent" })),
      callback_url: Type.Optional(Type.String({ description: "the full http://127.0.0.1:.../callback?... URL from the browser" })),
    }),
    async execute(_id, params) {
      const { reconnect, callback_url, agent_name } = params as { reconnect?: boolean; callback_url?: string; agent_name?: string }
      const text = async () => {
        if (callback_url) return finishWithCallback(pastedCallback(callback_url) ?? callback_url)
        if (agent_name?.trim() || reconnect || !(await auth.bearer())) {
          const approvalName = agent_name?.trim() ? agent_name.trim().slice(0, 80) : clientName
          return approvalText(approvalName, await startApproval(approvalName))
        }
        return connectedSummary()
      }
      return { content: [{ type: "text", text: await text() }], details: undefined }
    },
  })

  pi.registerTool({
    name: "office_status",
    label: "Office Status",
    description: "Report the Agent Office extension's poll, session, and credential state.",
    parameters: Type.Object({}),
    async execute() {
      return { content: [{ type: "text", text: statusReport() }], details: undefined }
    },
  })

  function statusReport(): string {
    return JSON.stringify(
      {
        location: directory,
        canonical,
        credentialsFile,
        tokenStorage: usingCredentialsOverride ? "override" : "agent",
        agentPointer: pointerFile,
        repoPointer: repoPointerFile,
        clientName,
        mode: dispatchMode,
        pollInterval,
        pollingAllowed,
        mcpRegistered,
        working,
        streaming,
        isChild,
        lastMcpError,
        mcpRefreshAt: mcpRefreshAt ? new Date(mcpRefreshAt).toISOString() : undefined,
        lastPollAt: lastPollAt ? new Date(lastPollAt).toISOString() : undefined,
        lastPollError,
        lastPromptAt: lastPromptAt ? new Date(lastPromptAt).toISOString() : undefined,
        lastPending,
        agent: ownAgent,
        seen: { messages: seenMessages.size, unclaimed: seenTasks.size, assigned: seenAssigned.size },
        backlog: [...backlog],
        dispatching: [...dispatching],
      },
      null,
      2,
    )
  }
}

// The repo root a worktree belongs to, as the canonical project path.
function canonicalRoot(directory: string): string {
  try {
    const { execFileSync } = require("node:child_process") as typeof import("node:child_process")
    const common = execFileSync("git", ["-C", directory, "rev-parse", "--path-format=absolute", "--git-common-dir"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    }).trim()
    return common.endsWith(`${sep}.git`) || common.endsWith("/.git") ? dirname(common) : directory
  } catch {
    return directory
  }
}

export { agentCredentialsPath, agentPointerPath, credentialsPath } from "./lib/office.ts"
