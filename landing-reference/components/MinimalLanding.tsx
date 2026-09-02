import { Link } from "@tanstack/react-router";
import { Apple, ArrowRight, Play, Sparkles } from "lucide-react";

import agentConfirmation from "@/assets/landing/agent-confirmation.jpg";
import agentHome from "@/assets/landing/agent-home.jpg";
import naniLista from "@/assets/nani/nani-lista.png";

export function MinimalLanding() {
  return (
    <div className="minimal-landing" lang="en">
      <header className="minimal-header">
        <a href="#minimal-top" className="minimal-wordmark" aria-label="Nana Wallet home">
          nana
        </a>

        <div className="minimal-header-actions" id="download">
          <span className="minimal-store-button">
            <Apple aria-hidden="true" />
            <span>
              <small>Coming soon on</small>
              App Store
            </span>
          </span>
          <span className="minimal-store-button minimal-store-play">
            <Play aria-hidden="true" />
            <span>
              <small>Coming soon on</small>
              Google Play
            </span>
          </span>
          <Link to="/app" className="minimal-header-launch">
            Launch app
          </Link>
        </div>
      </header>

      <main id="minimal-top" className="minimal-hero">
        <div className="minimal-orb minimal-orb-main" aria-hidden="true" />
        <div className="minimal-orb minimal-orb-small" aria-hidden="true" />
        <div className="minimal-grid" aria-hidden="true" />

        <div className="minimal-copy">
          <div className="minimal-pill">
            <Sparkles aria-hidden="true" />A wallet that speaks your language
          </div>

          <h1>nana</h1>
          <h2>Your money, in the words you already use.</h2>
          <p>Independence through voice. Protection when it matters.</p>

          <div className="minimal-actions">
            <Link to="/app" className="minimal-button minimal-button-primary">
              Launch app <ArrowRight aria-hidden="true" />
            </Link>
            <a href="#download-note" className="minimal-button minimal-button-secondary">
              Download app
            </a>
          </div>

          <div id="download-note" className="minimal-download-note">
            Native builds for iOS and Android · Store release coming soon
          </div>
        </div>

        <div className="minimal-product-art" aria-label="Nana Wallet mobile experience">
          <div className="minimal-phone minimal-phone-left">
            <span aria-hidden="true" />
            <img src={agentHome} alt="Nana Wallet home screen" />
          </div>

          <img className="minimal-nani" src={naniLista} alt="Nani, the Nana Wallet assistant" />

          <div className="minimal-phone minimal-phone-right">
            <span aria-hidden="true" />
            <img src={agentConfirmation} alt="Nana Wallet transfer confirmation screen" />
          </div>
        </div>
      </main>
    </div>
  );
}
