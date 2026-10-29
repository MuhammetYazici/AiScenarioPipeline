🤖 Autonomous AI-Agent Test Automation Orchestrator

An end-to-end, LLM-orchestrated test generation, live inspection, and self-healing pipeline that bridges the gap between raw business requirements and enterprise Playwright Java test automation frameworks.

📌 Architecture Overview

This project is the orchestration engine of a dual-repo autonomous testing ecosystem. It processes unstructured business requirements, converts them into standard Gherkin features (Shift-Left), inspects live web applications using Playwright Accessibility Object Models (AOM), and dynamically synthesizes contract-compliant Java test classes with automated compile-time healing.

Target Automation Core: The generated tests, Page Object Models, and Step Definitions are synthesized directly into the aiStagerPlaywright (Java Test Core) repository.

[ Business Docs (.docx / .xlsx) ]
               │
               ▼  (Phase 1: Shift-Left / LLM)
     [ Gherkin .feature File ]
               │
  ═════════════╪═════════════════════════════════════════════════
               │  (Phase 2: Live Orchestration)
               ▼
     [ PlaywrightPlanner ] ──► Live Web Inspection (DOM / AOM Tree)
               │
               ▼
     [ Cucumber CLI Dry-Run ] ──► Official Binding Contract (Snippets)
               │
               ▼
     [ JavaCodeGenerator ] ──► Synthesizes Runner, Page & Step Classes
               │
               ▼
     [ SelfHealer Engine ] ──► Cucumber Step Verification & Maven Log Healing
               │
               ▼
     [ Git Push / CI-CD ] ──► Automatic Deployment to Remote Core


🚀 Key Architectural Features

1. Shift-Left BDD Synthesis (--mode design-only)

Ingests requirements from Word (.docx) and Excel (.xlsx) matrices.

Employs Gemini LLM to extract clean Business Rules and generate structured Cucumber .feature scenarios covering positive, negative, and edge cases before a single line of application frontend is deployed.

Sanitizes syntax deterministically via zero-token Python Regex engines.

2. Live AOM/DOM Discovery (PlaywrightPlanner)

Launches a headless Chromium instance to inspect the live target application.

Extracts accessibility nodes (AriaRole.BUTTON, TEXTBOX, LINK, checkbox) to map semantic UI trees.

Generates an actionable, hallucination-free Execution Plan (JSON) correlating Gherkin steps with verified DOM selectors.

3. Contract-Driven Code Generation (extract_official_cucumber_snippets)

Executes an automated mvn test -Dcucumber.execution.dry-run=true pass prior to code generation.

Uses the Cucumber CLI output as an unalterable contract to eliminate parameter mismatching, regex errors, and signature drift in generated Java steps.

Adheres strictly to PicoContainer Constructor Injection for thread-safe state sharing (TestContext) and avoids static pollution.

4. Autonomous Compiler Healing Loop (SelfHealer)

Step Auditing: Detects missing or drifted Cucumber step implementations and patches them directly into existing Java classes.

Compiler Debugger: Triggers mvn test-compile in a sub-process. If compiler errors occur (method mismatch, invalid Locator casting, missing imports), it parses compiler stack traces and iteratively repairs Pages and StepDefinations until BUILD SUCCESS is attained.

Auto-Commit: Automatically pushes validated test suites to GitHub with timestamped telemetry commits.

🛠️ Project Structure

AiScenarioPipeline/
├── scenario_pipeline.py     # Core agent orchestration engine
├── input_docs/              # Raw requirements (.docx, .xlsx)
├── requirements.txt         # Python dependencies
├── .env.example             # Environment variable template
└── README.md                # Project documentation


⚙️ Installation & Setup

Prerequisites

Python 3.10+

Java JDK 17+ & Apache Maven configured in PATH

Google Gemini API Key

Installation

Clone the repository:

git clone https://github.com/MuhammetYazici/AiScenarioPipeline.git
cd AiScenarioPipeline


Create and activate a virtual environment:

python -m venv .venv
# Windows (PowerShell):
.venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate


Install dependencies:

pip install -r requirements.txt
playwright install chromium


Configure environment variables:
Create a .env file in the root directory:

GEMINI_API_KEY=your_gemini_api_key_here
TARGET_URL=https://aistager.ai/tr


💻 Execution Modes

The engine provides a decoupled, two-phase execution lifecycle:

Phase 1: Requirement Analysis & Feature Design

Used during early sprint phases when web pages are not yet deployed.

python scenario_pipeline.py --mode design-only


Parses documents in input_docs/, structures user stories, and generates .feature files directly into the Java test directory.

Phase 2: Live Automation, Synthesis & Self-Healing

Triggered once target environments (Staging/Dev) are live.

python scenario_pipeline.py --mode full-automation


Scans DOM/AOM elements, extracts Cucumber CLI contracts, synthesizes Page/Step classes, triggers Maven compile-healer loops, and commits verified suites to Git.

🔗 Related Repositories

Execution & Test Core: aiStagerPlaywright – Enterprise Playwright-Java framework utilizing TestNG, Cucumber BDD, PicoContainer, and Page Object Model (POM).

👤 Author

Muhammet Mustafa Yazıcı
Software Development Engineer in Test (SDET)
LinkedIn: https://www.linkedin.com/in/muhammet-yazici-556b57314/
GitHub: https://github.com/MuhammetYazici
