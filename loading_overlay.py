"""Inline printer animation shown only while generating a report."""

from contextlib import contextmanager
from html import escape
from time import monotonic, sleep

import streamlit as st


MIN_DISPLAY_SECONDS = 4.0


_STYLE = """
<style>
.jay-loading-inline {
    display: grid;
    place-items: center;
    width: 100%;
    min-height: 310px;
    box-sizing: border-box;
    margin: 8px 0 16px;
    padding: 24px;
    border: 1px solid rgba(244, 214, 106, .22);
    border-radius: 22px;
    background: linear-gradient(145deg, #222326, #141517);
    box-shadow: inset 0 1px rgba(255, 255, 255, .06);
}
.jay-loading-content {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0;
    max-width: 460px;
    color: #f6f2e8;
    font-family: inherit;
    text-align: center;
}
.jay-printer {
    position: relative;
    width: 150px;
    height: 178px;
    margin-bottom: 4px;
}
.jay-printer-paper {
    position: absolute;
    top: 77px;
    left: 43px;
    width: 64px;
    height: 65px;
    border-radius: 3px;
    background: #f3f0e8;
    box-shadow: 0 5px 10px rgba(0, 0, 0, .35);
    animation: jay-paper-feed 2s ease-in-out infinite;
}
.jay-printer-paper::before,
.jay-printer-paper::after {
    content: "";
    position: absolute;
    left: 14px;
    right: 14px;
    height: 2px;
    background: #a9acb0;
}
.jay-printer-paper::before { top: 41px; }
.jay-printer-paper::after { top: 49px; }
.jay-printer-feeder {
    position: absolute;
    top: 12px;
    left: 39px;
    width: 72px;
    height: 58px;
    border: 3px solid #d7dade;
    border-radius: 6px 6px 0 0;
    background: #35393f;
}
.jay-printer-body {
    position: absolute;
    z-index: 1;
    top: 61px;
    left: 10px;
    width: 130px;
    height: 65px;
    border: 3px solid #d7dade;
    border-radius: 13px;
    background: #34383e;
    box-shadow: 0 12px 24px rgba(0, 0, 0, .32);
}
.jay-printer-body::before {
    content: "";
    position: absolute;
    top: 18px;
    right: 17px;
    width: 10px;
    height: 10px;
    border-radius: 50%;
    background: #f0cc5b;
    box-shadow: 0 0 12px rgba(240, 204, 91, .4);
    animation: jay-printer-light 2s ease-in-out infinite;
}
.jay-loading-title {
    margin: 0 0 8px;
    font-size: 23px;
    font-weight: 700;
    line-height: 1.25;
}
.jay-loading-dots::after {
    content: "";
    animation: jay-loading-dots 1.5s steps(4, end) infinite;
}
.jay-loading-detail {
    margin: 0;
    color: #c2c5c9;
    font-size: 15px;
    line-height: 1.45;
}
@keyframes jay-paper-feed {
    0%, 100% { transform: translateY(-19px); }
    50% { transform: translateY(23px); }
}
@keyframes jay-printer-light {
    0%, 100% { opacity: .4; }
    50% { opacity: 1; }
}
@keyframes jay-loading-dots {
    0% { content: ""; }
    25% { content: "."; }
    50% { content: ".."; }
    75%, 100% { content: "..."; }
}
@media (prefers-reduced-motion: reduce) {
    .jay-printer-paper, .jay-printer-body::before, .jay-loading-dots::after {
        animation: none;
    }
    .jay-loading-dots::after { content: "..."; }
}
</style>
"""


def _markup(detail: str) -> str:
    return f"""{_STYLE}
<div class="jay-loading-inline" role="status" aria-live="polite" aria-label="Preparing report">
  <div class="jay-loading-content">
    <div class="jay-printer" aria-hidden="true">
      <div class="jay-printer-feeder"></div>
      <div class="jay-printer-paper"></div>
      <div class="jay-printer-body"></div>
    </div>
    <div>
      <p class="jay-loading-title">Preparing report<span class="jay-loading-dots"></span></p>
      <p class="jay-loading-detail">{escape(detail)}</p>
    </div>
  </div>
</div>"""


@contextmanager
def printing_overlay(detail: str):
    """Show the inline printer for at least four seconds while preparing a report."""
    slot = st.empty()
    slot.html(_markup(detail))
    started = monotonic()
    try:
        yield
    finally:
        remaining = MIN_DISPLAY_SECONDS - (monotonic() - started)
        if remaining > 0:
            sleep(remaining)
        slot.empty()
