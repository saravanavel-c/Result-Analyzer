# Result Analyzer

A full-stack application for managing and analyzing student results, built with React (frontend) and Node.js/Express (backend) with MongoDB database.

## Features

- **Course Management**: Add, view, and delete courses with details like code, name, faculty, type, and credits
- **Student Management**: View student details and performance across different exams
- **Excel Integration**: Upload Excel files containing student marks for multiple exams (UT1, UT2, SEMESTER)
- **Result Analysis**: Generate comprehensive analysis reports with pass/fail statistics
- **GPA Calculation**: Automatic GPA calculation for semester results
- **Data Visualization**: View student performance data in an organized, expandable format

## Tech Stack

- **Frontend**: React with Vite
- **Backend**: Node.js with Express.js
- **Database**: MongoDB with Mongoose ODM
- **Excel Processing**: ExcelJS library
- **File Upload**: Multer middleware
- **Frontend Framework**: React with custom CSS

## Installation

1. Clone the repository:
```
git clone <repository-url>
cd result-analyzer
```

2. Install backend dependencies:
```
cd backend
npm install
```

3. Install frontend dependencies:
```
cd ../frontend/result-analysis-system
npm install
```

4. Go back to the root directory:
```
cd ../../
```

## Setup

1. Make sure MongoDB is running on your system
2. The application expects MongoDB to be available at `mongodb://127.0.0.1:27017/resultDB`

## Running the Application

1. Start the backend server:
```
cd backend
node server.js
```

2. In a new terminal, start the frontend development server:
```
cd frontend/result-analysis-system
npm run dev
```

3. Open your browser and navigate to `http://localhost:5173`

## Usage

### Adding Courses
1. Click on "Add Course" button
2. Fill in course details (code, name, faculty, type, credits)
3. Click "Add Course"

### Uploading Results
1. Prepare an Excel file with the correct format (see EXCEL_FORMAT_GUIDE.txt below)
2. Click "Choose Excel File" and select your file
3. The system will process the file and update student data

### Viewing Analysis
1. Click "Download Result Analysis Sheet" to get comprehensive reports
2. Reports include UT1, UT2, and SEMESTER analysis with pass/fail statistics

### Managing Students
1. Switch to the "Students" tab
2. View the list of students and click on any student to see detailed performance data

## Excel File Format

The Excel file should have the following structure:

- **Sheets**: "UT1", "UT2", and "SEMESTER" (or "UT 1", "UT 2", "SEM")
- **Columns**: 
  - Column A: SNO (Serial Number)
  - Column B: REG NO (Student Registration Number)
  - Column C: NAME (Student Name)
  - Columns D+: Course columns with course codes in headers (e.g., "CS432 - DATA STRUCTURES")

- **Course Codes**: The system recognizes course codes in formats like:
  - CS432, EC284, MAT264 (letters followed by numbers)
  - 22SVA4XX (numbers followed by letters and alphanumeric characters)

## API Endpoints

- `GET /courses` - Get all courses
- `POST /courses` - Add a new course
- `DELETE /courses/:id` - Delete a specific course
- `DELETE /courses` - Delete all courses
- `POST /upload` - Upload and process Excel file
- `GET /export` - Download analysis report
- `GET /students` - Get all students
- `GET /students/:regNo` - Get specific student details
- `DELETE /students` - Delete all students
- `GET /stats` - Get statistics

## Project Structure

```
Result Analyzer/
├── backend/
│   ├── server.js           # Main backend server
│   └── package.json
├── frontend/
│   └── result-analysis-system/
│       ├── src/
│       │   ├── App.jsx     # Main React component
│       │   ├── App.css     # Styles
│       │   └── main.jsx    # Entry point
│       └── package.json
├── README.md
└── package.json
```

## Troubleshooting

- If you get MongoDB connection errors, ensure MongoDB is running
- If Excel files aren't processed correctly, verify the format matches the required structure
- Make sure course codes in your Excel file match those added to the system

## License

This project is open source and available under the MIT License.