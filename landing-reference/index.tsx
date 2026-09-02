import { Capacitor } from "@capacitor/core";
import { Link, Navigate, createFileRoute } from "@tanstack/react-router";
import {
  ArrowRight,
  Check,
  Download,
  ExternalLink,
  LockKeyhole,
  Mic2,
  ShieldCheck,
  Sparkles,
  Users,
} from "lucide-react";

import agentCommand from "@/assets/landing/agent-command.jpg";
import agentConfirmation from "@/assets/landing/agent-confirmation.jpg";
import agentConfirmed from "@/assets/landing/agent-confirmed.jpg";
import agentHome from "@/assets/landing/agent-home.jpg";
import naniEscuchando from "@/assets/nani/nani-escuchando.png";
import naniLista from "@/assets/nani/nani-lista.png";
import { HybridLandingIntro } from "@/components/HybridLandingIntro";
import { MinimalLanding } from "@/components/MinimalLanding";
import { publicSiteOrigin, socialPreviewUrl } from "@/lib/brand";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Nana Wallet — Independence through voice. Protection when it matters." },
      {
        name: "description",
        content:
          "A voice-first agentic wallet that gives older adults financial independence and gives their family peace of mind.",
      },
      {
        property: "og:title",
        content: "Nana Wallet — Independence through voice. Protection when it matters.",
      },
      {
        property: "og:description",
        content:
          "Financial independence for the person using it. Peace of mind for the person who cares for them.",
      },
      { property: "og:type", content: "website" },
      { property: "og:url", content: publicSiteOrigin },
      { property: "og:site_name", content: "Nana Wallet" },
      { property: "og:image", content: socialPreviewUrl },
      { property: "og:image:secure_url", content: socialPreviewUrl },
      { property: "og:image:type", content: "image/png" },
      { property: "og:image:width", content: "1280" },
      { property: "og:image:height", content: "640" },
      { property: "og:image:alt", content: "Nana Wallet with Nani, the wallet assistant" },
      { name: "twitter:card", content: "summary_large_image" },
      { name: "twitter:image", content: socialPreviewUrl },
      { name: "twitter:image:alt", content: "Nana Wallet with Nani, the wallet assistant" },
    ],
  }),
  component: LandingPage,
});

const productSteps = [
  {
    number: "01",
    title: "Say it naturally",
    description: "No menus. No forms. Ask Nana in the words you already use.",
    image: agentCommand,
    alt: "Nana Wallet receiving a transfer request in ordinary language",
  },
  {
    number: "02",
    title: "Review what matters",
    description: "Nana prepares the action and shows the important details in plain language.",
    image: agentConfirmation,
    alt: "Nana Wallet showing a prepared transfer for review",
  },
  {
    number: "03",
    title: "Confirm once",
    description: "You make the decision. Nana handles the complexity and shows the receipt.",
    image: agentConfirmed,
    alt: "Nana Wallet showing a confirmed onchain transaction",
  },
];

function LandingPage() {
  if (Capacitor.isNativePlatform()) {
    return <Navigate to="/app" replace />;
  }

  if (import.meta.env["VITE_LANDING_VARIANT"] === "minimal") {
    return <MinimalLanding />;
  }

  const isHybrid = import.meta.env["VITE_LANDING_VARIANT"] === "hybrid";

  return (
    <main
      id="top"
      lang="en"
      className={isHybrid ? "landing-page landing-page-hybrid" : "landing-page"}
    >
      <header className="landing-header">
        <div className="landing-container landing-header-inner">
          {isHybrid ? (
            <a href="#top" className="minimal-wordmark" aria-label="Nana Wallet home">
              nana
            </a>
          ) : (
            <Link to="/" className="landing-brand" aria-label="Nana Wallet home">
              <span className="landing-brand-mark" aria-hidden="true">
                <Sparkles />
              </span>
              <span>Nana Wallet</span>
            </Link>
          )}

          <nav className="landing-nav" aria-label="Landing navigation">
            <a href="#how-it-works">How it works</a>
            <a href="#protection">Protection</a>
            <Link to="/app" className="landing-nav-cta">
              Launch app <ArrowRight aria-hidden="true" />
            </Link>
          </nav>
        </div>
      </header>

      {isHybrid ? (
        <HybridLandingIntro />
      ) : (
        <section className="landing-hero" aria-labelledby="landing-title">
          <div className="landing-container landing-hero-grid">
            <div className="landing-hero-copy">
              <div className="landing-eyebrow">
                <span className="landing-eyebrow-dot" />A voice-first agentic wallet
              </div>

              <h1 id="landing-title">
                Independence through voice.
                <span>Protection when it matters.</span>
              </h1>

              <p className="landing-hero-description">
                Nana gives older adults their financial independence back, without asking their
                family to give up peace of mind.
              </p>

              <div className="landing-hero-actions">
                <Link to="/app" className="landing-button landing-button-primary">
                  Launch app <ArrowRight aria-hidden="true" />
                </Link>
                <a href="#mobile-app" className="landing-button landing-button-secondary">
                  <Download aria-hidden="true" /> Download the app
                </a>
              </div>
            </div>

            <div className="landing-hero-art" aria-label="Nana Wallet product preview">
              <div className="landing-hero-orbit landing-hero-orbit-one" />
              <div className="landing-hero-orbit landing-hero-orbit-two" />

              <img
                className="landing-hero-nani"
                src={naniEscuchando}
                alt="Nani, the Nana Wallet assistant, waving"
                fetchPriority="high"
              />

              <div className="landing-phone landing-hero-phone">
                <div className="landing-phone-speaker" aria-hidden="true" />
                <img src={agentHome} alt="Nana Wallet agent home screen" fetchPriority="high" />
              </div>

              <div className="landing-float-card landing-float-card-talk">
                <Mic2 aria-hidden="true" />
                <span>
                  <small>You ask</small>
                  In ordinary words
                </span>
              </div>

              <div className="landing-float-card landing-float-card-safe">
                <ShieldCheck aria-hidden="true" />
                <span>
                  <small>You decide</small>
                  Before money moves
                </span>
              </div>
            </div>
          </div>
        </section>
      )}

      <section className="landing-trust-strip" aria-label="Nana principles">
        <div className="landing-container landing-trust-items">
          <span>
            <Check aria-hidden="true" /> Plain language
          </span>
          <span>
            <Check aria-hidden="true" /> Human confirmation
          </span>
          <span>
            <Check aria-hidden="true" /> Protection by default
          </span>
          <span>
            <Check aria-hidden="true" /> Complexity stays invisible
          </span>
        </div>
      </section>

      <section id="how-it-works" className="landing-section landing-product-section">
        <div className="landing-container">
          <div className="landing-section-heading">
            <span className="landing-kicker">The product</span>
            <h2>The interface becomes a conversation.</h2>
            <p>
              The agent does the work. The user makes the one decision that carries consequence.
            </p>
          </div>

          <div className="landing-steps">
            {productSteps.map((step) => (
              <article className="landing-step" key={step.number}>
                <div className="landing-step-copy">
                  <span>{step.number}</span>
                  <h3>{step.title}</h3>
                  <p>{step.description}</p>
                </div>
                <div className="landing-phone landing-step-phone">
                  <div className="landing-phone-speaker" aria-hidden="true" />
                  <img src={step.image} alt={step.alt} loading="lazy" />
                </div>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="landing-section landing-dual-section">
        <div className="landing-container landing-dual-grid">
          <div className="landing-dual-copy">
            <span className="landing-kicker">The insight</span>
            <h2>One wallet. Two people it frees.</h2>
            <p>
              The responsible person absorbs complexity once, so the person using Nana gets
              simplicity every day.
            </p>

            <div className="landing-dual-cards">
              <article>
                <Mic2 aria-hidden="true" />
                <span>The person using it</span>
                <h3>Gets independence</h3>
                <p>Speaks, confirms and pays their own bills again. Nobody takes over.</p>
              </article>
              <article>
                <Users aria-hidden="true" />
                <span>The person who cares</span>
                <h3>Gets peace of mind</h3>
                <p>Configures it once and is told only when something genuinely matters.</p>
              </article>
            </div>
          </div>

          <div className="landing-dual-art">
            <div className="landing-dual-disc" aria-hidden="true" />
            <img src={naniLista} alt="Nani ready to help" loading="lazy" />
            <div className="landing-dual-caption">
              Independence for one.
              <span>Freedom for the other.</span>
            </div>
          </div>
        </div>
      </section>

      <section className="landing-manifesto" aria-label="Nana product principle">
        <div className="landing-container">
          <ShieldCheck aria-hidden="true" />
          <p>
            The family must not approve every payment.
            <span>The moment they do, independence was not returned. It was transferred.</span>
          </p>
        </div>
      </section>

      <section id="protection" className="landing-section landing-protection-section">
        <div className="landing-container">
          <div className="landing-section-heading landing-section-heading-left">
            <span className="landing-kicker">How protection works</span>
            <h2>Normal activity stays simple. Unusual activity gets a second look.</h2>
            <p>
              Confirmation, spending limits and trusted recipients are part of the product — not a
              safety upsell.
            </p>
          </div>

          <div className="landing-protection-grid">
            <article className="landing-protection-card landing-protection-ordinary">
              <div className="landing-protection-label">
                <Check aria-hidden="true" /> Ordinary
              </div>
              <blockquote>“Send $15,000 to Sofi.”</blockquote>
              <ul>
                <li>Known recipient</li>
                <li>Usual amount</li>
                <li>Normal behaviour</li>
              </ul>
              <div className="landing-protection-result">
                User confirms <ArrowRight aria-hidden="true" /> Done
              </div>
            </article>

            <article className="landing-protection-card landing-protection-unusual">
              <div className="landing-protection-label">
                <LockKeyhole aria-hidden="true" /> Unusual
              </div>
              <blockquote>“Send $900,000 to this account.”</blockquote>
              <ul>
                <li>New recipient</li>
                <li>Unusual amount</li>
                <li>Over the configured limit</li>
              </ul>
              <div className="landing-protection-result">
                Transaction held <ArrowRight aria-hidden="true" /> Family alerted
              </div>
            </article>
          </div>

          <p className="landing-protection-rule">
            <strong>Normal:</strong> user → Nana.
            <strong> Unusual only:</strong> user → Nana → responsible person.
          </p>
        </div>
      </section>

      <section className="landing-section landing-built-section">
        <div className="landing-container landing-built-grid">
          <div>
            <span className="landing-kicker">It is not a mockup</span>
            <h2>Nana already executes real transactions.</h2>
            <p>
              The full path works on testnet: voice, agent, prepared transaction, human
              confirmation, signature, settlement and an onchain receipt.
            </p>
            <a
              href="https://github.com/rober8b/aleph-hackathon"
              target="_blank"
              rel="noreferrer"
              className="landing-text-link"
            >
              View the project on GitHub <ExternalLink aria-hidden="true" />
            </a>
          </div>

          <div className="landing-built-stats">
            <article>
              <strong>1</strong>
              <span>decision left to the user: confirm what matters</span>
            </article>
          </div>
        </div>
      </section>

      <section id="mobile-app" className="landing-mobile-section">
        <div className="landing-container landing-mobile-card">
          <div className="landing-mobile-copy">
            <span className="landing-kicker">Web, iOS and Android</span>
            <h2>Use Nana wherever it feels easiest.</h2>
            <p>
              Launch the working web experience now. Native mobile builds are ready for iOS and
              Android; public store downloads are coming next.
            </p>
            <div className="landing-mobile-actions">
              <Link to="/app" className="landing-button landing-button-light">
                Launch web app <ArrowRight aria-hidden="true" />
              </Link>
              <span className="landing-store-status">App Store · Coming soon</span>
              <span className="landing-store-status">Google Play · Coming soon</span>
            </div>
          </div>

          <div className="landing-mobile-visual" aria-hidden="true">
            <div className="landing-phone landing-mobile-phone landing-mobile-phone-back">
              <div className="landing-phone-speaker" />
              <img src={agentConfirmation} alt="" loading="lazy" />
            </div>
            <div className="landing-phone landing-mobile-phone landing-mobile-phone-front">
              <div className="landing-phone-speaker" />
              <img src={agentHome} alt="" loading="lazy" />
            </div>
          </div>
        </div>
      </section>

      <footer className="landing-footer">
        <div className="landing-container landing-footer-inner">
          {isHybrid ? (
            <span className="minimal-wordmark">nana</span>
          ) : (
            <div className="landing-brand">
              <span className="landing-brand-mark" aria-hidden="true">
                <Sparkles />
              </span>
              <span>Nana Wallet</span>
            </div>
          )}
          <p>Independence through voice. Protection when it matters.</p>
          <a href="#top">Back to top</a>
        </div>
      </footer>
    </main>
  );
}
