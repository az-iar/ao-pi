// Choosing a previous session by name: `/office session bruno` continues the PiG
// conversation that was named `bruno`, which is how a resumed session keeps the
// Office agent it was working as (the agent pointer is per directory, and the
// resumed chat carries its own history).
//
// The rule is deliberately predictable, because a wrong guess opens the wrong
// conversation and the model then answers out of somebody else's context:
// an exact name wins; a prefix wins when only one session has it; a substring
// wins when only one session has it; anything else is reported back as a list
// instead of a guess. When several sessions carry the same exact name the most
// recently modified one is resumed - `pig -r` shows the same order - and the
// command says which one it opened.

import type { SessionInfo } from "@earendil-works/pi-coding-agent"

export type SessionChoice =
  | { kind: "resume"; session: SessionInfo; via: "exact" | "prefix" | "substring" }
  | { kind: "ambiguous"; matches: SessionInfo[]; query: string; total: number }
  | { kind: "none"; named: SessionInfo[]; query: string }

const named = (infos: SessionInfo[]) => infos.filter((info) => (info.name ?? "").trim() !== "")

const newestFirst = (infos: SessionInfo[]) => [...infos].sort((a, b) => b.modified.getTime() - a.modified.getTime())

export function pickSession(infos: SessionInfo[], query: string): SessionChoice {
  const wanted = query.trim().toLowerCase()
  const candidates = named(infos)
  if (wanted === "") return { kind: "none", named: newestFirst(candidates).slice(0, 10), query }

  const byName = (test: (name: string) => boolean) => candidates.filter((info) => test((info.name ?? "").trim().toLowerCase()))

  for (const [via, test] of [
    ["exact", (name: string) => name === wanted],
    ["prefix", (name: string) => name.startsWith(wanted)],
    ["substring", (name: string) => name.includes(wanted)],
  ] as const) {
    const matches = newestFirst(byName(test))
    if (matches.length === 1) return { kind: "resume", session: matches[0], via }
    if (matches.length > 1) {
      // An exact name shared by several sessions is still a resume: the newest.
      return via === "exact"
        ? { kind: "resume", session: matches[0], via }
        : { kind: "ambiguous", matches: matches.slice(0, 5), query, total: matches.length }
    }
  }
  return { kind: "none", named: newestFirst(candidates).slice(0, 10), query }
}

// message renders a choice the way the command tells the user about it.
export function message(choice: SessionChoice, current?: string): string {
  const stamp = (info: SessionInfo) => info.modified.toISOString().slice(0, 16).replace("T", " ")
  const label = (info: SessionInfo) => `${info.name ?? "(unnamed)"} — ${stamp(info)}, ${info.messageCount} messages`
  switch (choice.kind) {
    case "resume":
      if (current && current === choice.session.path) return `Already in the session named ${choice.session.name ?? "?"}.`
      return `Continuing "${choice.session.name}" (${stamp(choice.session)}, ${choice.session.messageCount} messages).`
    case "ambiguous":
      return [
        // The "+" only when the list was cut: "3+" for three sessions that all fit would be a lie.
        `"${choice.query}" names ${choice.total}${choice.total > choice.matches.length ? "+" : ""} sessions; say which:`,
        ...choice.matches.map((info) => `  /office session "${info.name}"  (${label(info)})`),
      ].join("\n")
    case "none":
      return choice.named.length
        ? [
            choice.query
              ? `No session is named "${choice.query}" here. Named sessions in this folder:`
              : "Named sessions in this folder:",
            ...choice.named.map((info) => `  ${label(info)}`),
            "Name a session with /name, then /office session <name>.",
          ].join("\n")
        : `No named sessions in this folder yet. Name one with /name, then /office session <name>.`
  }
}
