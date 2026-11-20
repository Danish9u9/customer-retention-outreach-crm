# Customer Retention & Outreach CRM

Python (`pandas`, `Flask`) customer segmentation and reorder-prediction pipeline paired with an interactive browser dashboard for scheduling and logging personalized WhatsApp and Email follow-ups[cite: 27, 28, 30].

## Key Features

- **Automated Customer Segmentation:** Classifies buyers from Shopify CSV exports into **Priority** (3+ orders), **Re-order** (2 orders), and **Regular** segments while filtering out cancelled and re-attempt orders.
- **Smart Cadence & Timing Prediction:** Calculates each repeat customer's average days between orders to forecast their next reorder date, and identifies their median ordering hour to suggest the best outreach time[cite: 28].
- **Interactive Outreach Dashboard (`Flask`):** Single-page web interface featuring live search (`/` shortcut), timing/segment filters, light/dark themes, and CSV export.
- **One-Click WhatsApp & Email Templates:** Generates pre-filled `wa.me` and `mailto:` links using customizable message templates with `{name}` and `{product}` placeholders[cite: 27].
- **Persistent Follow-Up Tracking:** Logs contacted customers in `contacted.json` so follow-ups remain marked complete until the customer places a new order[cite: 27].

## Project Structure

- `customer_scheduler.py` — ETL script that cleans Shopify order exports, normalizes phone numbers, computes reorder metrics, and generates `Customer_Outreach_Schedule.xlsx`[cite: 28, 30].
- `app.py` — Local Flask server and responsive frontend dashboard for managing daily customer outreach[cite: 27, 30].
- `requirements.txt` — Python dependencies (`pandas`, `flask`, `openpyxl`)[cite: 30].

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
