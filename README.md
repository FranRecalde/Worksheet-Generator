# Pictorio - Visual Vocabulary Generator

Pictorio (formerly VibeWord) is a single-file, no-build-step React flashcard generator designed for language teachers. It transforms vocabulary prompts into customizable visual flashcard grids with multiple image sources, quiz scaffolding modes, printable handouts, and PowerPoint (.pptx) export.

## Key Features

- **Multi-Provider Image Fallback Chain**:
  1. **Pixabay Vector**: High-quality vector graphics (via Pixabay API).
  2. **Wikimedia Commons**: Free media search in the File namespace (Namespace 6), prioritizing SVGs and rendering via thumbnail URLs with author and license metadata.
  3. **Openverse API**: Openly licensed images with no API key needed, throttled to ~1 request/second to respect anonymous rate limits.
  4. **Pixabay Illustration**: Hand-drawn and digital illustrations.
  5. **Pexels Photos**: High-resolution photography (Authorization header with API key) as a last resort.
- **Image Source Filter**:
  - **Anything**: Searches all providers in the full fallback chain.
  - **Clipart only**: Restricts results to vector graphics and illustrations (Pixabay vector, Wikimedia SVG, Openverse SVG, Pixabay illustration).
  - **Photos only**: Restricts results to photographs (Pexels, Wikimedia photos, Openverse photos).
- **Parallel Fetching & Safe Concurrency**:
  - Cards resolve in parallel via `Promise.all`.
  - Built-in semaphore concurrency cap (max 3 concurrent image resolution pipelines).
  - Dedicated rate limiter for Openverse enforcing ≥1100ms between calls.
- **Provider Attribution & Badges**:
  - Each card features a color-coded provider badge (Pixabay, Wikimedia Commons, Openverse, Pexels, Upload).
  - Visible author and license strings with tooltip details.
  - Comprehensive **Credits & Attributions Slide** automatically appended to PowerPoint exports.
- **PowerPoint (.pptx) Export with CORS Protection**:
  - Individual 16:9 flashcard slides with contain-fit imagery and bold labels.
  - Dynamic summary grid slide matching selected columns and rows.
  - Final image credits slide detailing word, provider, author, and license.
  - Resilient blob/base64 image fetching with graceful fallback if remote servers block CORS.
- **Configuration & Error Surfacing**:
  - Simple `CONFIG` object at the top of `index.html` with clear placeholders for `PIXABAY_KEY` and `PEXELS_KEY`.
  - Friendly dismissible banner when API keys are unconfigured.
  - UI error log drawer surfacing fetch status and provider notices.
- **Interactive Flashcard Features**:
  - Grid size selector (Small 3x2, Standard 3x3, Large 4x3).
  - Mode selector: Full labels, Scaffolded (`J**** a* f*****`), and Blank (`________________`).
  - Per-card image refresh, manual search override, local image file upload, and inline editable text.
  - Print-ready stylesheet for physical classroom handouts.

## Quick Start

Open `index.html` directly in any web browser. No compilation, node modules, or build step required.
