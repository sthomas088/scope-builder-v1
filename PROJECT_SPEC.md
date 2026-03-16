# Scope Builder V1

## Goal
Build a simple browser-based tool for generating biology consulting scopes of work.

## V1 only
This version is biology only.

## Inputs
- Project Name
- Client
- Location
- Prepared By
- Date

## Tasks
The app should include these initial biology tasks:
- Biological Constraints Analysis
- Habitat Assessment
- Focused Burrowing Owl Survey
- Focused Botanical Survey
- Jurisdictional Delineation

## Output
Generate a Microsoft Word document (.docx), not a .txt file.

## File naming
Use this exact format:

[Project Name]_BioScope_[YYYY-MM-DD].docx

Example:
Rancho Vista Industrial_BioScope_2026-03-16.docx

## UI
The tool should be clean and simple:
- Project information section
- Biology task section
- Generate Scope button
- Scope preview area before export

## Important
- Do not use React or frameworks
- Use simple HTML, CSS, and JavaScript only
- Store task content in a structured data file so more tasks can be added later

## Out of scope for V1
Do not include these yet:
- CEQA section
- Cultural section
- Area plan filtering
- Fee calculator
- Save/reopen project logic
