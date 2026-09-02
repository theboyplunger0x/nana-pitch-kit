# Nana Wallet — master product, brand and pitch context

This is the canonical context document for anyone refining the pitch, landing, product narrative, application copy or roadmap.

It is not a pitch deck. It is the source from which those artifacts should be made.

## 1. What Nana is

Nana is a voice-first agentic financial account, initially designed for seniors and people with limited mobility.

The core belief is:

> People shouldn't have to adapt to financial technology. Financial technology should adapt to them.

Instead of navigating menus, typing addresses, choosing networks or understanding wallet mechanics, a user expresses intent naturally:

> “Hey Nana, send $20 to my grandson.”

Nana understands the intent, resolves the person, checks relevant policies, prepares the action, explains what matters and executes after the appropriate authorization.

The user should not need to understand stablecoins, blockchain networks, gas, addresses, on/off-ramps or settlement. Those are infrastructure beneath the product.

## 2. Why the product exists

### Marcos' grandmothers

As Marcos' grandmothers got older, the family's way of protecting them gradually meant taking responsibilities away.

Someone else paid bills. Someone else managed money. Someone else handled financial decisions.

That reduced some risks but also reduced independence. One of his grandmothers was also the victim of a financial scam, making the security problem concrete.

The insight was:

> Protection often becomes control.

Nana asks whether technology can protect someone without requiring the loss of financial autonomy.

### Susi

A second experience came from teaching technology to Susi, an older woman with Parkinson's.

Susi understood what she wanted to do. The problem was not understanding money; it was physically interacting with technology.

Reducing five taps to three does not solve the underlying problem when tapping, typing and navigating are themselves the barrier.

That produced a more useful question:

> Instead of reducing interactions, which interactions actually deserve to exist?

## 3. How Nana emerged

Nana was discovered during the Aleph Hackathon rather than brought in as a predetermined solution.

The team arrived exploring other payment ideas. Marcos had even brought a physical POS. After several pivots and roughly five hours of searching, the pieces connected:

- older adults;
- Susi;
- voice;
- wallets;
- agents;
- family protection;
- payments.

The product emerged at the intersection of a real problem and newly available agentic wallet infrastructure.

## 4. The original product insight

Traditional fintech optimizes:

`5 interactions → 3 interactions → 2 interactions`

Nana asks which interaction should remain.

The intended loop is:

`intent → agent prepares → human understands and confirms → action executes`

The agent handles complexity. The human retains meaningful control.

## 5. Voice is the interface. The agent is the wallet.

Nana is not:

`voice → speech-to-text → traditional wallet UI`

Nana is:

`voice / intent → agent → financial action`

The agent does not merely explain which button to press. It performs the work within user-defined boundaries.

## 6. What the first prototype proved

The Aleph prototype demonstrated:

`user intent → agent → prepared transaction → confirmation → wallet infrastructure → onchain execution`

It used Tether WDK and executed seven confirmed transactions on Ethereum Sepolia, ranging from 1 to 20 test USD₮.

Nana won second place in the Tether WDK track.

This did not validate product-market fit or the business model. It did provide:

- a functioning prototype;
- technical feasibility for the core loop;
- external product signal;
- ecosystem interest;
- a reason to continue beyond the hackathon.

## 7. The two users

### End user

Initially a senior or person with limited mobility.

They want independence. They express intent, use money, understand consequences and confirm meaningful actions.

### Responsible user

A son, daughter, family member or designated trusted person.

They want peace of mind without becoming somebody else's financial interface. They may assist onboarding, configure relationships and policies, receive alerts and intervene when necessary.

Nana is therefore a dual-user consumer financial product.

## 8. The central rule

> The responsible intervenes by exception, not by default.

If every payment requires family approval, Nana has not restored independence. It has digitized dependency.

### Normal activity

Known recipient + normal amount + normal behavior + within policy → user confirms → execute.

The responsible person does not participate.

### Exceptional activity

New recipient + unusual amount or behavior + outside policy → hold → escalate → approve or block.

Protection should behave like an exception engine, not parental control.

## 9. The harder security problem

An obviously suspicious transaction is easy to flag. A sophisticated scam can look like a series of individually plausible payments.

Nana's long-term protection architecture may combine:

- deterministic policies;
- spending limits;
- trusted recipients;
- behavioral analysis;
- anomaly detection;
- contextual reasoning;
- escalation;
- human oversight.

This remains product and technical research, not a solved component.

## 10. Delegated agent execution

The agent should never possess unrestricted authority over user funds.

The direction under exploration is:

`Smart Account + programmable permissions + delegated execution/session keys + multisig escalation`

Authority must be:

- bounded;
- explicit;
- transparent;
- revocable;
- constrained by policy.

## 11. Family Graph

Nana needs more than a contact list. It needs to understand financial relationships.

Family Graph means:

`people + relationships + roles + financial permissions`

Example edges can represent responsible/admin, alerts-only, trusted recipient or other delegated roles.

The strategic hypothesis is:

> Banks have accounts. Nana understands families.

Family Graph may become more defensible than voice alone and can eventually support seniors, minors, guardians, shared family responsibilities and other delegated accounts.

## 12. The initial wedge

Nana starts with seniors and people with limited mobility.

For a typical young user, voice-controlled finance may be convenience. For someone with Parkinson's or severe mobility limitations, it can be accessibility.

The wedge provides:

- a specific user;
- a real and urgent problem;
- strong differentiation;
- an emotional narrative;
- a reason voice and agents are necessary rather than gimmicks.

Expansion should not dilute this positioning prematurely.

## 13. Go-to-market hypothesis

The current hypothesis is responsible-led acquisition:

`content → son/daughter/responsible → “I need this for my mom” → setup together → end user adopts daily voice use`

Potential message:

> Give your parents financial independence without giving up peace of mind.

This is a hypothesis to validate, not a proven funnel.

## 14. The onboarding tension

If the responsible person performs KYC and owns or funds the wallet, the money may technically belong to them. That undermines the independence thesis.

The likely direction is not “the responsible person does everything instead.” It is “the responsible person accompanies difficult setup once so the end user gets simplicity every day.”

Wallet ownership, KYC and legal responsibility remain open product/legal questions.

## 15. Nana Free and Nana Family

Core protection cannot be premium. A free Nana account must demonstrate independence plus protection.

Nana Free should include the voice agent, payments, a basic responsible relationship and essential protection.

Nana Family appears when one relationship becomes a network: multiple protected people, multiple responsible people, roles, permissions, shared coordination and a family dashboard.

The thesis is not “pay to be protected.” It is:

> Pay when financial coordination across your family becomes more complex.

The exact pricing unit remains unvalidated.

## 16. Business model hypotheses

Nana does not yet have a validated business model. Three paths are being explored:

1. Financial activity and revenue share from embedded rails or services.
2. Nana Family subscription for expanded coordination and Family Graph complexity.
3. B2B2C licensing or distribution through banks, fintechs, wallets, insurers or care networks.

These are hypotheses, not conclusions.

## 17. Why Nana can be a company, not a feature

Voice alone is not defensible.

The proprietary product layer is the combination of:

`Agent + Family Graph + Protection Engine + delegated permissions + accessible UX + user relationship`

## 18. What Nana owns

Nana should own:

- product experience;
- brand;
- voice UX;
- agent;
- intent orchestration;
- Family Graph;
- Protection Engine;
- policy experience;
- user relationship.

Partners can provide:

- wallet infrastructure;
- stablecoins;
- blockchain settlement;
- fiat rails and liquidity;
- KYC infrastructure;
- QR and bill payments;
- local payment networks.

Principle:

> Own the product and user relationship. Do not reinvent every financial rail.

## 19. Infrastructure paths

Nana should remain above multiple providers.

### Proven path

Tether WDK + USDT + local providers. The prototype already used WDK successfully.

### Evaluation path

Circle / Arc + USDC + local providers. This remains an exploration path rather than a committed product dependency.

Potential local providers being explored include Twin, Belo, Ripio, Manteca and others. The purpose is to determine the correct division between Nana's product layer and existing financial infrastructure—not to collect logos.

## 20. Mission, vision and promises

### Mission

Give people financial independence without asking them to adapt to financial technology.

### Vision

A world where anyone can manage their financial life naturally, while the people they trust can protect them without taking control away.

### Functional promise

Manage money by expressing intent, not navigating interfaces.

### Emotional promise — end user

I can still manage my own money.

### Emotional promise — responsible

I don't have to control everything to know they're protected.

### Brand promise

Independence through voice. Protection when it matters.

## 21. What Nana is not

Nana is not:

- a crypto wallet for crypto users;
- a chatbot embedded in a banking app;
- a parental-control wallet;
- a system where a family member approves every transaction;
- an LLM holding unrestricted private keys;
- another interface exposing networks, gas and addresses.

Nana is an agentic financial account that turns intent into action while preserving bounded human control.

## 22. North Star

Before MAU, volume or TVL, the conceptual North Star is:

> Financial actions completed independently that previously required assistance.

Every time a person completes an action through Nana that previously required someone else, Nana has created independence.

## 23. Expansion sequence

`Seniors + limited mobility → family finance → delegated / agentic finance`

The initial wedge stays focused while the architecture supports a broader future.

## 24. Current facts, hypotheses and open questions

### Built or observed

- Working agent-to-wallet transfer loop.
- Seven test transactions reported by the team.
- Second place in the Tether WDK track.
- Voice intent, explicit confirmation and receipt flow.
- Public landing and web demo.

### Hypotheses

- Responsible-led acquisition.
- Family Graph as a defensible product layer.
- Nana Family subscription.
- Financial-activity revenue share.
- B2B2C distribution/licensing.
- Infrastructure abstraction across providers.

### Open questions

- Wallet ownership, KYC and the legal role of the responsible person.
- Exact policy and anomaly thresholds.
- Safe delegated execution architecture.
- Best infrastructure mix per market.
- Nana Family willingness to pay and pricing unit.
- Acquisition economics and retention.

## 25. The final primitive

> An agentic financial account that allows people to act independently within a trusted, programmable network of relationships and permissions.

For the end user: **I can manage my own money.**

For the responsible: **I don't need to control everything to know they're protected.**

For Nana: **Independence with protection.**

## Provenance

This document synthesizes the current product-model PDFs, Mission Board, current pitch deck, hackathon submission copy, landing implementation and the team's extended GPT context export.

The unedited GPT export is preserved as `master-context-source-gpt.txt` for traceability. Its final section is truncated in the original attachment, so this document completes the architecture using the newer product-model update and Mission Board.

