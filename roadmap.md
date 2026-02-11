# Gephyra - Development Roadmap

This document outlines the planned features and improvements for Gephyra in future phases.

## Phase 2: Advanced Data Handling & Flexibility
**Goal:** Support complex real-world scenarios and richer content.

- [ ] **Conditional Logic:** Support `{% if %}` statements in Word templates for dynamic content (e.g., show paragraph only if `Age > 18`).
- [ ] **Image Support:** 
  - Insert images into Word templates dynamically from file paths in Excel.
  - Resize and position images automatically.
- [ ] **Multiple Tables:** Support populating multiple different tables in a single document from different Excel sheets.
- [ ] **Data Filtering (GUI):** Allow users to filter rows in the GUI before generating (e.g., "Select all rows where Status = 'Pending'").
- [ ] **Output format:** Support exporting to PDF (requires Word installation or third-party libs).

## Phase 3: Performance & Enterprise Features
**Goal:** Optimization for large-scale compatibility and speed.

- [ ] **Multithreading:** Process files in parallel to speed up generation (started in Phase 1.3, needs refinement).
- [ ] **Database Integration:** Connect directly to SQL databases (SQLite/PostgreSQL) instead of just Excel.
- [ ] **API Integration:** Fetch data from REST APIs (e.g., CRM systems) to generate documents.
- [ ] **Email Automation:** Automatically email the generated files to recipients specified in the Excel row.

## Phase 4: Cloud & Deployment
**Goal:** Seamless distribution and cloud capabilities.

- [ ] **Auto-Update:** Mechanism to check for updates and upgrade the application automatically.
- [ ] **Cloud Storage:** Direct upload to Google Drive / OneDrive after generation.
- [ ] **Web Interface:** Port key features to a Web App (Streamlit/Django) for remote access.
- [ ] **License Management:** Simple license key system for Pro version distribution.

## Backlog / Nice-to-have
- [ ] Dark Mode toggle in GUI.
- [ ] Drag & Drop support for Config/Excel files.
- [ ] Recent Files history.
- [ ] "Watch Folder" mode: Auto-generate when a new Excel file drops into a folder.
