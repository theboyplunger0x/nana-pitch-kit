import { Link } from "@tanstack/react-router";
import { ArrowRight } from "lucide-react";

import agentConfirmation from "@/assets/landing/agent-confirmation.jpg";
import agentHome from "@/assets/landing/agent-home.jpg";
import naniLista from "@/assets/nani/nani-lista.png";

export function HybridLandingIntro() {
  return (
    <section className="minimal-hero hybrid-hero" aria-labelledby="hybrid-title">
      <div className="minimal-orb minimal-orb-main" aria-hidden="true" />
      <div className="minimal-orb minimal-orb-small" aria-hidden="true" />
      <div className="minimal-grid" aria-hidden="true" />

      <div className="minimal-copy">
        <h1 id="hybrid-title">nana</h1>
        <h2>The agentic wallet for seniors and people with limited mobility.</h2>
        <p>Independence through voice. Protection when it matters.</p>

        <div className="minimal-actions">
          <Link to="/app" className="minimal-button minimal-button-primary">
            Launch app <ArrowRight aria-hidden="true" />
          </Link>
          <a href="#mobile-app" className="minimal-button minimal-button-secondary">
            Download app
          </a>
        </div>

        <div className="minimal-download-note">
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
    </section>
  );
}
