# Nana — hackathon submission copy

Recovered and updated from the current product, deck, WDK implementation, and public demo on September 2, 2026. The original form text was not saved verbatim, so this is the canonical copy to use going forward.

## Project name

Nana

## Emoji

👵

Alternative if we want to emphasize the interaction rather than the initial audience: 🗣️

## Demonstration

https://nana-wallet-hybrid.vercel.app/app

## Short description

85 characters, including spaces.

> A voice-first agentic wallet that restores independence with protection by exception.

## Description

> Nana is a voice-first agentic wallet built for older adults and people with limited mobility. Users ask for a financial action in everyday language, review exactly what will happen, and explicitly confirm before money moves. The agent handles wallet complexity—addresses, networks, fees, and transaction preparation—while spending limits and approved recipients create a clear safety floor. Normal activity remains between the user and Nana; trusted family support is designed to appear only by exception, when an action is unusual or crosses a policy. Nana turns a crypto wallet into a calm conversation: more independence for the user and more peace of mind for the people who care for them.

## How it’s made

> Nana combines a React 19 and TypeScript frontend built with TanStack Start, Tailwind CSS, and shadcn/ui with a Node.js and Fastify backend. A ToolLoopAgent interprets natural-language payment requests and accesses Tether WDK through its bundled `wdk-mcp` process. Every transfer is first prepared with `send_token` in dry-run mode; the interface then presents the network, token, recipient, amount, and fee for explicit human confirmation. Only after confirmation can the backend repeat the same WDK call in live mode. The Sepolia testnet path fails closed behind a maximum transfer amount and recipient allowlist, uses idempotent confirmation handling, and verifies the chain ID, transaction hash, receipt, and final status after broadcast. A safe fixture mode keeps the flow reproducible without credentials or funds, while Capacitor packages the same frontend for iOS and Android. Nana also uses recipient memory to resolve familiar phrases such as “Lucas, my grandson” without exposing wallet addresses in the conversation, and asks for clarification whenever a match is ambiguous.

## GitHub repository

https://github.com/rober8b/aleph-hackathon

## Track

Aleph Hackathon 2026 — Track 1: Build with the WDK CLI

## Evidence notes

- Public landing and web demo: https://nana-wallet-hybrid.vercel.app
- WDK integration and submission README: https://github.com/rober8b/aleph-hackathon#aleph-hackathon-2026--wdk-track
- Reference network: Ethereum Sepolia
- Demo token alias: `usdt-test` (test USD₮)
- Safe fixture mode is the default; live mode requires an explicitly unlocked, limited test wallet.
