// The token endpoint does not only answer with JSON. A gateway in front of the
// Office answers a 502 with an HTML page, and the status used to be read *after*
// res.json(): the failure arrived as `SyntaxError: Unexpected token '<'`, which
// says nothing about the Office being unreachable and made every failed refresh
// look like a bug in the extension. The status now comes first, with a short
// excerpt of the body, and JSON is parsed only when the body is JSON.
//
// Its own file and its own server: lib/office.ts reads AGENT_OFFICE_URL when it
// is imported, so the URL has to be settled before the import (the harness in
// own-identity-harness.mts does the same for the extension tests).
import assert from "node:assert/strict"
import { createServer } from "node:http"
import type { AddressInfo } from "node:net"
import { after, test } from "node:test"

type Reply = { status: number; type: string; body: string }

// One reply per request, in order: a test decides exactly what the endpoint says.
const replies: Reply[] = []
const server = createServer((req, res) => {
  // Drain the form body: the point of the test is the reply, and an unread stream
  // only keeps the connection open.
  req.resume()
  const reply = replies.shift()
  if (!reply) {
    res.writeHead(500, { "Content-Type": "text/plain" }).end("no reply queued")
    return
  }
  res.writeHead(reply.status, { "Content-Type": reply.type }).end(reply.body)
})
await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve))
process.env.AGENT_OFFICE_URL = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
const { OFFICE, tokenRequest, tokenRequestRefused } = await import("../lib/office.ts")
after(() => {
  // fetch keeps its socket alive; a close that waits for it would outlive the run.
  server.closeAllConnections()
  server.close()
})

const failure = async (promise: Promise<unknown>) => {
  try {
    await promise
  } catch (err) {
    return err as Error
  }
  assert.fail("the request resolved instead of failing")
}

test("a gateway's HTML error page reports the HTTP status instead of a SyntaxError", async () => {
  replies.push({
    status: 502,
    type: "text/html",
    body: "<!doctype html><html><head><title>502 Bad Gateway</title></head><body><h1>502 Bad Gateway</h1><hr><center>nginx</center></body></html>",
  })
  const error = await failure(tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-secret-1" }))
  assert.ok(!(error instanceof SyntaxError), `the failure is a ${error.name}: ${error.message}`)
  assert.match(error.message, /HTTP 502/, `the status is missing from: ${error.message}`)
  assert.match(error.message, /Bad Gateway/, `the page's own words are missing from: ${error.message}`)
  assert.doesNotMatch(error.message, /[<>]/, `the page's markup leaked into the message: ${error.message}`)
  assert.doesNotMatch(error.message, /refresh-secret-1/, `the request's refresh token leaked into the message: ${error.message}`)
})

test("a JSON token response is still parsed", async () => {
  replies.push({
    status: 200,
    type: "application/json",
    body: JSON.stringify({ access_token: "access-9", refresh_token: "refresh-9", expires_in: 3600, token_type: "Bearer", scope: "office:read office:write office:events" }),
  })
  const before = Date.now()
  const credentials = await tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-8" })
  assert.equal(credentials.office, OFFICE)
  assert.equal(credentials.client_id, "client-1")
  assert.equal(credentials.access_token, "access-9")
  assert.equal(credentials.refresh_token, "refresh-9")
  // expires_in is seconds, and the deadline is the point of storing it.
  assert.ok(credentials.expires_at >= before + 3_600_000, `expires_at ${credentials.expires_at - before}ms from now`)
})

test("a refused grant is told apart from a failure that passes", async () => {
  // 400 and 401 are the Office refusing the grant: the same refresh token would
  // be refused again, so the extension stops retrying. A 5xx is the Office (or a
  // gateway) failing to answer and is worth another attempt.
  replies.push({ status: 400, type: "application/json", body: JSON.stringify({ error: "invalid_grant", error_description: "expired" }) })
  assert.equal(tokenRequestRefused(await failure(tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-6" }))), true)
  replies.push({ status: 401, type: "application/json", body: JSON.stringify({ error: "invalid_client" }) })
  assert.equal(tokenRequestRefused(await failure(tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-5" }))), true)
  replies.push({ status: 503, type: "text/html", body: "<html><body><h1>503 Service Unavailable</h1></body></html>" })
  assert.equal(tokenRequestRefused(await failure(tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-4" }))), false)
})

test("an Office JSON error still reports its own description", async () => {
  replies.push({ status: 400, type: "application/json", body: JSON.stringify({ error: "invalid_grant", error_description: "the refresh token was already used" }) })
  const error = await failure(tokenRequest("client-1", { grant_type: "refresh_token", refresh_token: "refresh-7" }))
  assert.match(error.message, /HTTP 400/, `the status is missing from: ${error.message}`)
  assert.match(error.message, /the refresh token was already used/, `the Office's description is missing from: ${error.message}`)
})
