# Faculty Weekly Performance System

> A role-based web application for structured faculty performance reporting, multi-tier review, and institutional oversight — built entirely on **Google Apps Script**.

---

## Overview

The **Faculty Weekly Performance System** digitises the full weekly reporting cycle. Faculty submit structured timesheets and self-assessments through a guided interface. Each report flows through a three-tier review chain with automated email alerts and in-app notifications at every stage.

---

## Stats at a Glance

| 9 | 5 | 15 | 3 |
|:---:|:---:|:---:|:---:|
| Data Sheets | User Roles | Activity Types | Review Tiers |

---

## Tech Stack

| Layer | Details |
|---|---|
| **Backend** | Google Apps Script — server-side JS, SpreadsheetApp API |
| **Frontend** | Single HTML file SPA — HTML5, CSS3, Vanilla JS (no framework) |
| **Database** | Google Sheets — 9 named sheets as relational tables |
| **Auth** | Custom password login; hash stored in Sheets |
| **Email** | MailApp — built-in Apps Script service |
| **Scheduling** | Time-based trigger — automated Friday reminder emails |
| **Config** | PropertiesService — email routing, registration codes |
| **Hosting** | Apps Script Web App deployment |

---

## How It Works

```
Faculty  ──›  HOD  ──›  HOI  ──›  IMO
Submit       Approve/   Second    Final
timesheet    Revise     review    sign-off
```

Each stage fires an email to the next reviewer. A **"Needs Revision"** decision returns the report to faculty with the reviewer's specific remark.

---

## User Roles

| Role | Responsibilities |
|---|---|
| **Faculty** | Submit timesheets · view status · resubmit after revision |
| **HOD** | Review department reports · approve or request changes |
| **HOI** | Second-tier review across all departments |
| **IMO** | Final oversight · manage faculty roster · system configuration |
| **Staff** | Support accounts for HOD / HOI / IMO tiers |

---

## Key Features

### Weekly Submission
- 5-day grid (Mon–Fri), 7 time slots per day — 8:30 AM to 3:30 PM
- 15 activity categories: teaching, research, LMS, mentoring, FDP, ERP, and more
- Self-assessment section: week outcomes and next-week targets
- Auto-generated Submission ID (`SUB-YYYYMMDD-XXXX`)

### Notifications & Reminders
- In-app notification badge per role with read/unread tracking
- Friday time-trigger emails every pending faculty — skips those who already submitted

### Enrollment & Auth
- IMO pre-enrolls faculty; a unique Faculty ID is auto-generated and shared
- Faculty self-activates by entering their ID and setting a password
- Self-registration also available on the public form

### UI Design
- Palette: Deep Navy · Antique Gold · Ivory
- Fonts: Playfair Display · DM Sans · DM Mono
- Split-panel login, role-tab selector, fade-up animations, sidebar navigation

---

## Data Schema

```
Staff_Master
  StaffID · StaffName · Email · Role · Department · PasswordHash · GoogleEmail · Status

Faculty_Master
  FacultyID · FacultyName · Email · Department · Campus · Institution
  Designation · PasswordHash · GoogleEmail · Status

Weekly_Submission
  SubmissionID · FacultyID · AcademicYearSemester · ReportingFrom · ReportingTo
  Declaration · SubmittedDateTime

Timesheet_Entries
  SubmissionID · Day · TimeSlot · ActivityType · ActivityDetails

Self_Assessment
  SubmissionID · OutcomeOfWeek · TargetPlanNextWeek

HOD_Remarks / HOI_Remarks / IMO_Monitoring
  SubmissionID · Remark · Status · DateTime

Notifications
  NotifID · ForRole · Type · Title · Body · SubmissionID · FacultyName · IsRead · CreatedAt
```

> `SubmissionID` is the primary foreign key linking all transaction tables.

---

## Server-Side Functions

<details>
<summary><strong>Auth & Registration</strong></summary>

| Function | Description |
|---|---|
| `login()` | Validates credentials; returns session token + role profile |
| `facultyRegister()` | Self-registration; auto-generates Faculty ID |
| `staffRegister()` | HOD / HOI / IMO registration with code validation |
| `activateFaculty()` | Faculty activates pre-enrolled account |
| `preEnrollFaculty()` | IMO creates faculty record in pending state |

</details>

<details>
<summary><strong>Submission</strong></summary>

| Function | Description |
|---|---|
| `getSubmitConfig()` | Returns dropdown lists for the submission form |
| `submitWeeklyReport()` | Writes to 3 sheets; fires HOD notification + email |
| `getMySubmissions()` | Full submission history for logged-in faculty |
| `getSubmissionDetail()` | Complete detail with timesheet rows and all remarks |

</details>

<details>
<summary><strong>Review</strong></summary>

| Function | Description |
|---|---|
| `getHODQueue()` | Pending submissions for the HOD's department |
| `submitHODReview()` | Writes HOD remark; triggers HOI on approval |
| `getHOIQueue()` | HOD-approved submissions pending HOI |
| `submitHOIReview()` | Writes HOI remark; triggers IMO on approval |
| `getIMOQueue()` | All submissions with full pipeline status |
| `submitIMOMonitoring()` | Writes IMO note; notifies faculty of final status |

</details>

<details>
<summary><strong>Utility</strong></summary>

| Function | Description |
|---|---|
| `getNotifications()` | Unread notifications for current role |
| `markNotifsRead()` | Marks all read for current role |
| `sendFridayReminders()` | Time-triggered; emails faculty with no submission this week |
| `initializeSystem()` | One-time setup: sheets, headers, validations, triggers |

</details>

---

## Deployment

```bash
# 1. Open Google Sheet → Extensions → Apps Script
# 2. Paste Code.gs content; create Index.html and paste frontend
# 3. Run initializeSystem() once — creates all sheets and triggers
# 4. Deploy → New Deployment → Web App
#    Execute as: Me  |  Access: Anyone
# 5. Set HOD_DEFAULT, HOI_DEFAULT, IMO_EMAIL in Script Properties
# 6. Share the Web App URL
```

### Default Registration Codes

| Role | Default Code | Script Property |
|---|---|---|
| HOD | `HOD@VMRF` | `REGCODE_HOD` |
| HOI | `HOI@VMRF` | `REGCODE_HOI` |
| IMO | `IMO@VMRF` | `REGCODE_IMO` |

> Change these in **Extensions → Apps Script → Project Settings → Script Properties**.

---

## Planned Enhancements

- [ ] Analytics dashboard — submission rates, activity mix, approval turnaround
- [ ] PDF export of approved faculty reports
- [ ] Google Workspace SSO — auto-login via Google accounts
- [ ] Bulk approve / export for the IMO
- [ ] Mobile-responsive layout
- [ ] Leave and absence tagging in the timesheet grid

---

## Architecture

```
Browser (Index.html SPA)
        │
        │  google.script.run  (async RPC)
        ▼
Apps Script Server (Code.gs)
        │
        │  SpreadsheetApp API
        ▼
Google Sheets (9 named sheets)
        │
        │  MailApp / Time Triggers
        ▼
Email (Gmail / Google Workspace)
```

---

*Built with Google Apps Script · 2026*
