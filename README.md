# 🚀 Multi-Tab Document & BOQ Engine (v5.7)

An enterprise-grade Google Apps Script solution designed for businesses migrating from Airtable to Google Workspace. This engine handles relational data across multiple tabs to generate complex Customer Quotes and Contractor Offers.

## 🏗️ Technical Architecture
- **Relational Lookup:** Searches across `Quote gen`, `Bill of quantities`, and `Ground works` tabs using a unified Site Name key.
- **Dynamic Table Generation:** Automatically compiles multiple rows of "Install Items" into a single formatted list (`{{Full_BOQ_Table}}`).
- **Dual-Mode Templates:** Supports separate Google Doc templates for Customer-facing and Contractor-facing documents.
- **UK Formatting:** Standardized for GB localization (£ currency and dd/mm/yyyy dates).
- **Admin Dispatch:** Automatically archives the final PDF to Drive and dispatches a copy to a central Admin email for review.

## 📁 Repository Structure
- `Engine.gs`: The core automation logic.
- `README.md`: System documentation.

## 👨‍💻 Developed By
**Legacy Microsoft Partner (Since 2012)** Specializing in bridging the gap between field operations (Construction/Electrical) and high-performance digital systems.
