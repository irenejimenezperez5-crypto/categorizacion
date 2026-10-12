import React from "react";
import essityLogo from "../assets/essity-logo.svg";
import "./essity-brand.css";

/** Official artwork, bundled locally and kept on a white clear-space plate. */
export function EssityBrandLogo() {
  return <span className="essity-brand-logo">
    <img src={essityLogo} alt="Essity" width="112" height="28"
         decoding="async" fetchPriority="high" draggable={false} />
  </span>;
}
