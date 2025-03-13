# Greenwin Expense Report System

An AI-powered expense report system that automatically extracts information from PDF receipts, categorizes expenses, and generates professional reports in Excel and PDF formats.

## Features

- **AI-Powered Receipt Processing**: Upload PDF receipts and automatically extract expense details
- **G/L Code Classification**: Automatically categorize expenses with appropriate G/L codes
- **Excel Export**: Generate professionally formatted Excel reports
- **PDF Merging**: Combine multiple receipt PDFs into a single document
- **Session Management**: Keep track of uploads and ensure privacy by automatic cleanup

## Tech Stack

- Node.js
- Express.js
- OpenAI API
- ExcelJS
- PDF-Merger-JS
- Vanilla JavaScript (Frontend)
- CSS3

## Setup Instructions

### Prerequisites

- Node.js (v14 or higher)
- npm (v6 or higher)
- OpenAI API key

### Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/expense-report.git
   cd expense-report
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file based on the example:
   ```
   cp .env-example .env
   ```

4. Update the `.env` file with your OpenAI API key:
   ```
   PORT=3005
   OPENAI_API_KEY=your_openai_api_key
   ```

5. Start the server:
   ```
   npm start
   ```

6. Open your browser and navigate to `http://localhost:3005`

## Usage

1. Enter your name and department
2. Upload PDF receipts
3. Review extracted expenses
4. Export to Excel or merged PDF

## Deployment

### Deploying to Render.com

1. Create a new Web Service in Render
2. Connect your GitHub repository
3. Configure the build settings:
   - Build Command: `npm install`
   - Start Command: `node server.js`
4. Add environment variables in the Render dashboard

## License

MIT

## Acknowledgements

- OpenAI for their powerful API
- ExcelJS for Excel generation capabilities
- PDF-Merger-JS for PDF handling 