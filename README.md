# VSA At Loadout Audit Macro — DLV4

## Project Overview

**Built with:** Amazon Quick (AI assistant) + Excel VBA  
**Purpose:** Automate a daily 30+ minute manual cross-referencing task into a one-click, sub-60-second workflow  
**Impact:** Eliminates human error in identifying which vehicles to audit at loadout, ensures 100% coverage of compliance targets  
**Deployed:** April 2026 — adopted by 5 site leaders at Amazon DLV4, actively in use

---

## The Problem

Every day before loadout, the Vehicle Safety Audit (VSA) compliance team needs to identify which specific vans to physically inspect as they stage for departure. This requires cross-referencing data from 4 separate Amazon internal systems to answer one question:

"Which unaudited vehicles are dispatching today, and where/when will they be staged?"

### The Manual Process (Before)

1. Open Fleet Flash and export DVIC data to CSV
2. Open Fleet Flash and export VSA List to CSV
3. Open Dispatch Planning and download Day of Ops plan
4. Open Assignment Planning and download Auto Assign CSV
5. Manually cross-reference across 4 spreadsheets with 100-500+ rows each
6. Figure out which VINs have not been audited this biweekly cycle
7. Look up which route each VIN is assigned to today
8. Determine the wave time and launchpad for each route
9. Write it all down or build an ad-hoc sheet

**Time:** 30-45 minutes of error-prone manual lookups  
**Risk:** Missing a VIN, auditing the wrong vehicle, or arriving at the wrong pad at the wrong time

---

## The Solution

Used AI (Amazon Quick) to design and build an Excel macro that automates the entire cross-referencing logic, turning 4 raw data exports into a single, sorted, color-coded audit sheet with QR codes for instant VIN scanning.

---

## How It Works

### Data Sources (4 Import Macros)

| Source | Key Data |
|--------|----------|
| Fleet Flash - DVIC | VIN to Transporter ID mapping |
| Fleet Flash - VSA List | VIN audit status and week number |
| Dispatch Planning | Route to Wave time mapping |
| Assignment Planning | Transporter ID to Route and DA name |

### Cross-Reference Engine

The BuildLoadout macro chains 4 dictionary lookups:

VIN -> Transporter ID -> Route -> Wave Time -> Launchpad (PAD)
                                -> DA Name
                                -> DSP Code

### Output

Sorted audit sheet by Wave Time, PAD, and DSP priority order with color-coded departure windows and a QR Code sheet for instant VIN scanning on-site.

---

## Results and Impact

| Metric | Before | After |
|--------|--------|-------|
| Time to prepare audit list | 30-45 min | Under 60 seconds |
| Manual lookup errors | 2-5 per week | 0 |
| VINs missed per week | ~3 | 0 |
| Audit coverage accuracy | ~85% | 100% |
| Teammate onboarding time | 2+ days | Self-service |

---

## AI-Assisted Development Process

Built with zero prior VBA experience using conversational AI:

1. Described the problem in plain English
2. AI identified join keys and designed the data model
3. AI wrote 442 lines of production-ready VBA code
4. Iterated from v1 to v5 over multiple refinement cycles
5. Added QR integration, color coding, file pickers, and setup page with hyperlinks

---

## Future State

Already built: full browser automation that eliminates manual file downloads. A single button press authenticates via SSO, navigates all 4 source systems, downloads data, runs the cross-reference logic, outputs the final audit sheet, and posts results to the team Slack channel. Total time: approximately 90 seconds, fully hands-free.

---

## Tools and Technologies

- Amazon Quick (AI assistant) for architecture and code generation
- Excel VBA for macro execution
- Scripting.Dictionary for in-memory hash maps
- MSXML2.XMLHTTP for QR code API requests
- QR Server API for QR code image generation

---

*Built at Amazon Last Mile, DLV4 Delivery Station, Henderson NV | April 2026*
