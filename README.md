# 📊 DCMA XER Schedule Health Dashboard

## 📋 Description

The DCMA XER Schedule Health Dashboard is a web-based analysis tool for assessing the quality and health of project schedules in XER format (Primavera P6 file format). This application implements the DCMA (Defense Contract Management Agency) 14-point assessment to provide schedule quality metrics and specific improvement suggestions.

## ✨ Features

- **📤 XER File Analysis**: Upload and process XER files from Primavera P6.
- **📈 KPI Visualisation**: Visually displays 10 key performance indicators based on DCMA methodology.
- **🔧 Threshold Adjustment**: Allows customisation of compliance thresholds for each KPI.
- **🔍 Issue Identification**: Identifies non-compliant elements affecting schedule quality.
- **📥 Downloadable Reports**: Export non-compliant activities to Excel for easier correction.
- **📑 Executive Summary**: Provides an overall summary of the schedule status and improvement recommendations.

## 🎯 Implemented KPIs

1. **🔄 Logic**: Percentage of incomplete tasks without predecessors and/or successors.
2. **⏪ Leads**: Percentage of relationships with negative lag.
3. **⏩ Lags**: Percentage of relationships with positive lag.
4. **🔗 Relationship Types**: Percentage of Finish-to-Start (FS) relationships.
5. **🔒 Hard Constraints**: Percentage of activities with hard constraints.
6. **↔️ High Float**: Percentage of activities with float > 44 days.
7. **↘️ Negative Float**: Percentage of activities with float < 0 days.
8. **⏳ High Duration**: Percentage of activities with duration > 44 days.
9. **📅 Invalid Dates**: Number of activities with invalid dates.
10. **🔓 Soft Constraints**: Percentage of activities with soft constraints.

## 💻 Technologies Used

- HTML5
- CSS3 (with custom variables)
- JavaScript (ES6+)
- [SheetJS](https://sheetjs.com/) for Excel data export

## 🚀 Installation and Usage

1. Clone or download this repository.
2. No installation or server required - simply open the `index.html` file in any modern web browser.
3. Select an XER file and set the data date to begin analysis.

## 📁 File Structure

```
├── index.html         # HTML structure and interface components
├── styles.css         # Styles and visual themes
├── script.js          # JavaScript logic for analysis and visualisation
└── README.md          # This file
```

## 📝 How to Use

1. Click "Select XER File" to select a Primavera P6 schedule file (XER format).
2. Set the Data Date in the date selector.
3. Click "Process Data" to begin the analysis.
4. Review KPIs and executive summary.
5. Use threshold adjustments to customise evaluation criteria based on your project needs.
6. Download the non-compliant activities report by clicking "Download Non-Compliant Activities".

## 📜 Licence

This project is licensed under the terms of the Apache 2.0 licence - see the [LICENCE](https://github.com/project2080/xer-schedule-health-dashboard/blob/main/LICENSE) file for details.

## 👥 Contributing

Contributions are welcome. For major changes, please open an issue first to discuss what you would like to change.

## 🖼️ Screenshots

*(Add screenshots of the interface here if desired)*

---

**Note**: ⚠️ This tool is for analysis purposes and does not modify the original XER files.
