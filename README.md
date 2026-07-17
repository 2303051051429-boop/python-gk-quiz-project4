# 🐍 Python GK Quiz – Industrial Training Kit

**Project 4** by DecodeLabs (Batch 2026)

This repository contains a dual-component educational kit designed to teach beginner programming students about core Python concepts (Input processing, control flows, and defensive programming) through the lens of a General Knowledge Quiz.

## 📦 What's Inside?

This project features two main components:

### 1. Presentation Deck Generator (`generate-deck.js`)
A Node.js script utilizing `pptxgenjs` to automatically generate a high-quality, professional 20-slide `.pptx` presentation. 
- **Features:** Dark mode aesthetic, glass-panel cards, code blocks, comparison tables, and defensive programming workflows.
- **Content:** Covers IPO (Input-Process-Output) models, string sanitization (`.strip()`, `.lower()`), comparison operators, if/else flows, and real-world logic applications.

### 2. Interactive Web Application (`quiz-app/index.html`)
A live, responsive 4-option MCQ web app that implements the exact logic taught in the presentation.
- **Features:** 
  - 🎨 **Modern UI:** Aurora mesh backgrounds, glassmorphism cards, and particle networks.
  - ⚡ **Instant Feedback:** Color-coded reveals (green for correct, red/shake for wrong) and score XP bumping.
  - ⌨️ **Keyboard Navigation:** Use `1`, `2`, `3`, `4` to select options and `Enter` to advance.
  - 🎉 **Gamified Results:** Score ring animation, confetti bursts for perfect scores, and a full question breakdown.

## 🚀 Quick Start

### Generating the Presentation
1. Ensure you have Node.js installed.
2. Install the required dependencies:
   ```bash
   npm install
   ```
3. Run the generator script:
   ```bash
   node generate-deck.js
   ```
4. A file named `Python-GK-Quiz-Training-Deck.pptx` will be created in the root directory.

### Running the Web App
1. You can open `quiz-app/index.html` directly in any modern web browser.
2. Alternatively, serve it via a local HTTP server for the best experience:
   ```bash
   npx http-server -p 8080
   ```
   Then visit `http://localhost:8080/quiz-app/index.html`.

## 🛠️ Tech Stack
- **Web App:** Vanilla HTML, CSS, JavaScript (No external UI frameworks)
- **Deck Generator:** Node.js, `pptxgenjs`

## 💡 Design Philosophy
Built with a focus on modern, premium aesthetics. We avoided generic templates and instead opted for dynamic micro-animations, curated typography (Space Grotesk & JetBrains Mono), and fluid state transitions to maximize learner engagement.
