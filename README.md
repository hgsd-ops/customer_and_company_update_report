# Customer & Competitor Updates Report

Generates automated monthly reports from Monday.com boards, tracking updates for both customers and competitors over the last 31 days.

## Features

- **Dual Board Support**: Generates separate reports for both Companies and Competitors boards
- **31-Day Rolling Window**: Automatically pulls updates from the last 31 days
- **PDF Generation**: Creates professional PDF reports with keyword highlighting
- **HTML Preview**: Opens browser preview of the company report
- **User Activity Heatmap**: Visual representation of team member contributions
- **Smart Text Processing**: Cleans and truncates email threads, highlights keywords (English & Norwegian)
- **Link Extraction**: Automatically extracts and displays URLs from updates

## Generated Reports

The script produces:
- `Company updates – [date range].pdf`
- `Competitor updates – [date range].pdf`
- `company_updates_preview.html`
- `competitor_updates_preview.html`

## Usage

```bash
python customer_monthly_report.py
```

The script will:
1. Fetch updates from both boards (last 31 days)
2. Generate HTML previews and PDFs
3. Display progress and update counts
4. Open the company preview in your default browser

## Configuration

Board IDs are configured in the script:
- **Companies Board**: `3401154685`
- **Competitors Board**: `18362600897`

## Requirements

- Python 3.x
- Chrome or Microsoft Edge (for PDF generation)
- Monday.com API access
