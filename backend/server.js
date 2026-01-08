const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const multer = require("multer");
const path = require("path");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(bodyParser.json());

if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads");
}

/* MONGODB CONNECTION */
mongoose.connect("mongodb://127.0.0.1:27017/resultDB", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log("MongoDB connected"))
.catch(err => console.error("MongoDB connection error:", err));

/* SCHEMAS */
const StudentSchema = new mongoose.Schema({
  regNo: String,
  name: String,
  courses: {
    type: Map,
    of: {
      UT1: { type: Number, default: null },
      UT2: { type: Number, default: null },
      SEMESTER: { type: Number, default: null }
    }
  },
  failedCourses: {
    UT1: { type: Number, default: 0 },
    UT2: { type: Number, default: 0 },
    SEMESTER: { type: Number, default: 0 }
  },
  GPA: { type: Number, default: 0 }
}, { minimize: false });

const Student = mongoose.model("Student", StudentSchema);

const CourseSchema = new mongoose.Schema({
  code: String,    
  name: String,     
  faculty: String,  
  type: { type: String, enum: ["theory", "practical", "extra"] },
  credits: { type: Number, default: 3 },
  failedStudents: {
    UT1: { type: Number, default: 0 },
    UT2: { type: Number, default: 0 },
    SEMESTER: { type: Number, default: 0 }
  }
}, { timestamps: true });
const Course = mongoose.model("Course", CourseSchema);

/* MULTER CONFIG */
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) => cb(null, Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage });

/* HELPER FUNCTIONS */
function extractCourseCode(cellValue) {
  if (!cellValue) return null;
  const str = cellValue.toString().trim();
  // Updated regex to match patterns like CS432 or 22SVA4XX (letters/numbers + alphanumeric mix)
  const match = str.match(/\b[A-Z]{2,4}\d{2,4}[A-Z0-9]*\b|\b\d{2}[A-Z]{3}[A-Z0-9]{2,4}\b/);
  return match ? match[0] : null;
}

function calculateGPA(student, courses) {
  let totalPoints = 0;
  let totalCredits = 0;
  
  for (const [code, marks] of student.courses.entries()) {
    const course = courses.find(c => c.code === code);
    if (course && marks.SEMESTER !== null && marks.SEMESTER !== undefined) {
      totalPoints += marks.SEMESTER * course.credits;
      totalCredits += course.credits;
    }
  }
  
  return totalCredits > 0 ? parseFloat((totalPoints / totalCredits).toFixed(2)) : 0;
}

/* ROUTES */
app.post("/courses", async (req, res) => {
  try {
    const { code, name, faculty, type, credits } = req.body;
    
    if (!code || !name || !faculty || !type) {
      return res.status(400).json({ error: "All fields are required" });
    }
    
    // Validate credits - allow decimal values up to one decimal place
    let creditsValue = 3; // default value
    if (credits !== undefined && credits !== null) {
      creditsValue = parseFloat(credits);
      if (isNaN(creditsValue) || creditsValue < 0) {
        return res.status(400).json({ error: "Credits must be a valid positive number" });
      }
      // Round to one decimal place
      creditsValue = Math.round(creditsValue * 10) / 10;
    }
    
    const existing = await Course.findOne({ code });
    if (existing) {
      return res.status(400).json({ error: "Course code already exists" });
    }
    
    const course = new Course({ code, name, faculty, type, credits: creditsValue });
    await course.save();
    res.json({ message: "Course added successfully!", course });
  } catch (err) {
    console.error("Error adding course:", err);
    res.status(500).json({ error: "Failed to add course." });
  }
});

app.get("/courses", async (req, res) => {
  try {
    const courses = await Course.find();
    res.json(courses);
  } catch (err) {
    console.error("Error fetching courses:", err);
    res.status(500).json({ error: "Failed to fetch courses" });
  }
});

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const filePath = req.file.path;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const courses = await Course.find();
    if (courses.length === 0) {
      return res.status(400).json({ error: "No courses found. Please add courses first." });
    }

    const sheetMapping = {
      "UT 1": "UT1", "UT 2": "UT2", "UT1": "UT1", "UT2": "UT2",
      "SEMESTER": "SEMESTER", "SEM": "SEMESTER"
    };

    for (const worksheet of workbook.worksheets) {
      const sheetName = worksheet.name.trim();
      const examType = sheetMapping[sheetName];
      
      if (!examType) {
        console.log(`Skipping sheet: ${sheetName}`);
        continue;
      }

      console.log(`Processing sheet: ${sheetName} as ${examType}`);

      const headerRow = worksheet.getRow(1);
      const courseColumns = {};
      
      headerRow.eachCell((cell, colNumber) => {
        if (colNumber <= 3) return;
        
        const courseCode = extractCourseCode(cell.value);
        if (courseCode) {
          const course = courses.find(c => c.code === courseCode);
          if (course) {
            if ((examType === "UT1" || examType === "UT2") && (course.type === "theory" || course.type === "extra")) {
              courseColumns[colNumber] = courseCode;
            } else if (examType === "SEMESTER") {
              courseColumns[colNumber] = courseCode;
            }
          }
        }
      });

      console.log(`Found ${Object.keys(courseColumns).length} courses in ${examType}`);

      for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        
        const regNoCell = row.getCell(2).value;
        const nameCell = row.getCell(3).value;
        
        if (!regNoCell || !nameCell) continue;
        
        const regNo = regNoCell.toString().trim();
        const name = nameCell.toString().trim();
        
        if (!regNo || !name || regNo === 'REG NO' || name === 'NAME') continue;

        let student = await Student.findOne({ regNo });
        if (!student) {
          student = new Student({
            regNo, name, courses: new Map(),
            failedCourses: { UT1: 0, UT2: 0, SEMESTER: 0 }, GPA: 0
          });
        }

        // First, store all marks for this exam
        for (const [col, code] of Object.entries(courseColumns)) {
          const markCell = row.getCell(parseInt(col)).value;
          
          let mark = null;
          
          // Check if the cell has hyphen (student not enrolled)
          if (typeof markCell === 'string' && markCell.trim() === '-') {
            // Hyphen means not enrolled, skip this course
            continue;
          }
          
          // Check if student is absent (marked as 'a' or 'A')
          if (typeof markCell === 'string' && (markCell.toLowerCase().trim() === 'a' || markCell.toLowerCase().trim() === 'absent')) {
            mark = 0; // Absent student gets 0
          } 
          // Check if the cell is empty (treat as 0)
          else if (markCell === null || markCell === undefined || markCell === '') {
            mark = 0;
          } 
          // Process numeric values
          else {
            if (typeof markCell === 'object' && markCell.result !== undefined) {
              mark = Number(markCell.result);
            } else {
              mark = Number(markCell);
            }
            
            if (isNaN(mark) || !isFinite(mark)) {
              console.log(`Invalid mark for ${regNo}, course ${code}, exam ${examType}`);
              mark = null;
            }
          }
          
          if (mark === null) continue;

          let courseData = student.courses.get(code);
          if (!courseData) {
            courseData = { UT1: null, UT2: null, SEMESTER: null };
          }
          
          courseData[examType] = mark;
          student.courses.set(code, courseData);
        }

        // NOW count failures for this exam type ONLY
        let failedCount = 0;
        for (const [code, marks] of student.courses.entries()) {
          const course = courses.find(c => c.code === code);
          const mark = marks[examType];
          
          if (mark !== null && mark !== undefined) {
            // For UT exams, only count theory and extra courses where student is enrolled
            if (examType === "UT1" || examType === "UT2") {
              if (course && (course.type === "theory" || course.type === "extra") && mark < 30) {
                failedCount++;
              }
            } 
            // For SEMESTER, count all courses where student is enrolled
            else if (examType === "SEMESTER") {
              if (mark < 5) failedCount++;
            }
          }
        }
        
        student.failedCourses[examType] = failedCount;

        if (examType === "SEMESTER") {
          student.GPA = calculateGPA(student, courses);
        }

        try {
          await student.save();
          console.log(`Saved ${regNo} - ${name} for ${examType}, failed: ${failedCount}`);
        } catch (saveErr) {
          console.error(`Error saving ${regNo}:`, saveErr.message);
          throw saveErr;
        }
      }

      // After processing all students for this exam, count failures per course
      console.log(`Calculating failure counts for each course in ${examType}...`);
      
      for (const courseCode of Object.values(courseColumns)) {
        const course = courses.find(c => c.code === courseCode);
        if (!course) continue;

        // Get all students who have this course
        const allStudents = await Student.find({});
        let failedCount = 0;

        for (const student of allStudents) {
          const courseData = student.courses.get(courseCode);
          if (!courseData) continue;

          const mark = courseData[examType];
          if (mark === null || mark === undefined) continue;

          // Check if student failed based on exam type
          if (examType === "UT1" || examType === "UT2") {
            // For UT exams: mark < 30 is fail (for theory and extra courses)
            if ((course.type === "theory" || course.type === "extra") && mark < 30) {
              failedCount++;
            }
          } else if (examType === "SEMESTER") {
            // For SEMESTER: grade point < 5 is fail (all courses)
            if (mark < 5) {
              failedCount++;
            }
          }
        }

        // Update the course's fail count for this exam type
        if (!course.failedStudents) {
          course.failedStudents = { UT1: 0, UT2: 0, SEMESTER: 0 };
        }
        course.failedStudents[examType] = failedCount;
        await course.save();
        
        console.log(`Course ${courseCode} - ${examType}: ${failedCount} students failed`);
      }
    }

    fs.unlinkSync(filePath);
    res.json({ message: "Excel marks uploaded and processed successfully!" });

  } catch (err) {
    console.error("Error processing Excel:", err);
    res.status(500).json({ error: `Failed to process Excel file: ${err.message}` });
  }
});

app.get("/export", async (req, res) => {
  try {
    const students = await Student.find().lean();
    const thinBorders = {
          border: {
              top: { style: 'thin' }, left: { style: 'thin' },
              bottom: { style: 'thin' }, right: { style: 'thin' }
            },
             alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
             font: {name: 'Times New Roman'}
    };

    const mediumBorders = {
          border: {
              top: { style: 'medium' }, left: { style: 'medium' },
              bottom: { style: 'medium' }, right: { style: 'medium' }
            },
             alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
             font: {name: 'Times New Roman'}
        };

    const headerStyle = {
      font: { name:'Times New Roman' ,bold: true, size: 8, color: { argb: 'FF000000' } },
      alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
      border: {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      }
    };

    top3style = {
      font: {name:'Times New Roman' ,bold: true, size:  10},
      alignment: {vertical: 'middle', horizontal: 'center'},
      border: {
        top: { style: 'medium' }, left: { style: 'medium' },
        bottom: { style: 'medium' }, right: { style: 'medium' }
      }
    };

    students.sort((a, b) => {
      const lastThreeA = parseInt(a.regNo.slice(-3));
      const lastThreeB = parseInt(b.regNo.slice(-3));
      return lastThreeA - lastThreeB;
    });

    const courses = await Course.find().sort({ createdAt: 1 });
    
    if (!students.length || !courses.length) {
      return res.status(400).json({ error: "No data found" });
    }

    students.forEach(student => {
      student.courses = student.courses && typeof student.courses === 'object' 
        ? new Map(Object.entries(student.courses)) 
        : new Map();
    });

    const workbook = new ExcelJS.Workbook();
    const passingMarks = 30;
    const passingGP = 5;

    // Sort courses: theory first, then practical, then extra (in insertion order within each type)
    const theoryCourses = courses.filter(c => c.type === "theory");
    const practicalCourses = courses.filter(c => c.type === "practical");
    const extraCourses = courses.filter(c => c.type === "extra");
    const sortedCourses = [...theoryCourses, ...practicalCourses, ...extraCourses];

    const addUTSheet = (examName) => {
      const sheet = workbook.addWorksheet(examName);
      
      // For UT sheets: only theory and extra courses (in insertion order within each type)
      const theoryCourses = courses.filter(c => c.type === "theory");
      const extraCourses = courses.filter(c => c.type === "extra");
      const theoryAndExtraCourses = [...theoryCourses, ...extraCourses];

      const totalDataColumns = 3 + theoryAndExtraCourses.length + 1;
      const finalColLetter = sheet.getColumn(totalDataColumns).letter;

      const failedColIndex = 3 + theoryAndExtraCourses.length + 1;
      const failedColLetter = sheet.getColumn(failedColIndex).letter;

      const startRow = 5;
      const endRow = startRow + students.length - 1;
      const failedRange = `${failedColLetter}${startRow}:${failedColLetter}${endRow}`;
      const maxFails = theoryAndExtraCourses.length;

      const redShades = [
          'FFFFDDDD', 'FFFFCCCC', 'FFFFBBBB', 'FFFF9999','FFFF7777', 
          'FFFF5555', 'FFFF3333', 'FFFF1111', 'FFCC0000', 'FF990000'
      ];
      
      for (let count = 1; count <= maxFails; count++) {
          const colorIndex = Math.min(count - 1, redShades.length - 1);
          const colorARGB = redShades[colorIndex];
          
          sheet.addConditionalFormatting({
              ref: failedRange,
              rules: [{
                  type: 'cellIs',
                  operator: 'equal',
                  priority: maxFails - count + 2,
                  formulae: [String(count)],
                  style: {
                      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: colorARGB } },
                      font: { color: { argb: 'FF000000' } }
                  }
              }]
          });
      }

      sheet.addRow(["GOVERNMENT COLLEGE OF TECHNOLOGY,CBE-13"]);
      sheet.addRow(["DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"]);
      sheet.addRow([`UNIT TEST ${examName.slice(-1)} ANALYSIS`]);

      sheet.mergeCells(`A1:${finalColLetter}1`);
      sheet.mergeCells(`A2:${finalColLetter}2`);
      sheet.mergeCells(`A3:${finalColLetter}3`);

      sheet.getCell('A1').style = top3style;
      sheet.getCell('A2').style = top3style;
      sheet.getCell('A3').style = top3style;

      const dataHeaderRow = sheet.addRow([
        "SNO", "REG NO", "NAME",
        ...theoryAndExtraCourses.map(c => `${c.code} - ${c.name}`),
        "No Of Subjects Failed"
      ]);

      dataHeaderRow.eachCell(cell => {
        cell.style = headerStyle;
      });

      sheet.getColumn(1).width = 10;
      sheet.getColumn(2).width = 25;
      sheet.getColumn(3).width = 30;

      students.forEach((stu, index) => {
        const marks = theoryAndExtraCourses.map(c => {
          const courseData = stu.courses.get(c.code);
          // If course data doesn't exist for extra courses, return empty string (not enrolled)
          if (!courseData || courseData[examName] === null || courseData[examName] === undefined) {
            return c.type === "extra" ? "" : 0;
          }
          return courseData[examName];
        });
        
        // Count failures only for enrolled courses
        const failedCount = marks.filter((m, idx) => {
          const course = theoryAndExtraCourses[idx];
          // Only count if student is enrolled (mark is not empty string)
          return m !== "" && m < passingMarks;
        }).length;
        
        sheet.addRow([index + 1, stu.regNo, stu.name, ...marks, failedCount]).eachCell(cell=>{
          cell.style = thinBorders;
        });
      });

      // Add conditional formatting for fail marks (marks < 30 shown in red)
      const firstMarkCol = 4; // Marks start from column 4 (after SNO, REG NO, NAME)
      const lastMarkCol = 3 + theoryAndExtraCourses.length;
      for (let col = firstMarkCol; col <= lastMarkCol; col++) {
        const colLetter = sheet.getColumn(col).letter;
        const markRange = `${colLetter}${startRow}:${colLetter}${endRow}`;
        
        sheet.addConditionalFormatting({
          ref: markRange,
          rules: [{
            type: 'cellIs',
            operator: 'lessThan',
            priority: 1,
            formulae: [String(passingMarks)],
            style: {
              font: { color: { argb: 'FFFF0000' }, bold: true }
            }
          }]
        });
      }

      sheet.addRow([]);
      sheet.addRow([]);
      theoryAndExtraCourses.forEach(c => sheet.addRow([c.code, c.name, c.faculty]).eachCell(cell=>{
          cell.style = thinBorders;
        }));
      sheet.addRow([]);
      sheet.addRow([]);
      sheet.addRow(["SUB CODE", "SUB NAME", "NO OF STUDENTS PASSED", "PASS %"]).eachCell(cell => {
          cell.style = mediumBorders;
        });

      theoryAndExtraCourses.forEach(c => {
        const studentsWithMarks = students.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData && courseData[examName] !== null && courseData[examName] !== undefined;
        });
        
        const passedCount = studentsWithMarks.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData[examName] >= passingMarks;
        }).length;
        
        const attended = studentsWithMarks.length;
        const passPercent = attended > 0 ? Math.round((passedCount / attended) * 100) : 0;
        const pcRows = sheet.addRow([c.code, c.name, passedCount, passPercent]);
        
        pcRows.eachCell(cell => {
          cell.style = mediumBorders;
        });
      });
    };
    
    addUTSheet("UT1");
    addUTSheet("UT2");

    
    const semSheet = workbook.addWorksheet("SEMESTER");

    const totalDataColumns = 3 + sortedCourses.length + 2;
    const finalColLetter = semSheet.getColumn(totalDataColumns).letter;
    
    semSheet.addRow(["GOVERNMENT COLLEGE OF TECHNOLOGY,CBE-13"]);
    semSheet.addRow(["DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"]);
    semSheet.addRow(["SEMESTER EXAM ANALYSIS"]);

    semSheet.mergeCells(`A1:${finalColLetter}1`);
    semSheet.mergeCells(`A2:${finalColLetter}2`);
    semSheet.mergeCells(`A3:${finalColLetter}3`);

    semSheet.getCell('A1').style = top3style;
    semSheet.getCell('A2').style = top3style;
    semSheet.getCell('A3').style = top3style;

    const dataHeaderRow = semSheet.addRow([
      "SNO","REG NO.", "NAME",
      ...sortedCourses.map(c => `${c.code} - ${c.name}`),
      "No. of Arrears", "GPA"
    ]);

    dataHeaderRow.eachCell(cell => {
      cell.style = headerStyle;
    });

    semSheet.getColumn(1).width = 10;
    semSheet.getColumn(2).width = 25;
    semSheet.getColumn(3).width = 30;

    const semStartRow = 5;
    students.forEach((stu, index) => {
      const marks = sortedCourses.map(c => {
        const courseData = stu.courses.get(c.code);
        // If course data doesn't exist for extra courses, return empty string (not enrolled)
        if (!courseData || courseData.SEMESTER === null || courseData.SEMESTER === undefined) {
          return c.type === "extra" ? "" : 0;
        }
        return courseData.SEMESTER;
      });
      
      // Count arrears only for enrolled courses
      const arrears = marks.filter((m, idx) => {
        const course = sortedCourses[idx];
        // Only count if student is enrolled (mark is not empty string)
        return m !== "" && m < passingGP;
      }).length;
      
      const gpa = stu.GPA || 0;
      semSheet.addRow([index+1, stu.regNo, stu.name, ...marks, arrears, parseFloat(gpa.toFixed(2))]).eachCell(cell=>{
          cell.style = thinBorders;
        });
    });

    // Add conditional formatting for failed grade points (< 5 shown in red)
    const semEndRow = semStartRow + students.length - 1;
    const firstGPCol = 4; // Grade points start from column 4 (after SNO, REG NO, NAME)
    const lastGPCol = 3 + sortedCourses.length;
    for (let col = firstGPCol; col <= lastGPCol; col++) {
      const colLetter = semSheet.getColumn(col).letter;
      const gpRange = `${colLetter}${semStartRow}:${colLetter}${semEndRow}`;
      
      semSheet.addConditionalFormatting({
        ref: gpRange,
        rules: [{
          type: 'cellIs',
          operator: 'lessThan',
          priority: 1,
          formulae: [String(passingGP)],
          style: {
            font: { color: { argb: 'FFFF0000' }, bold: true }
          }
        }]
      });
    }

    semSheet.addRow([]);
    semSheet.addRow([]);
    sortedCourses.forEach(c => semSheet.addRow([c.code, c.name, c.faculty]).eachCell(cell=>{
          cell.style = thinBorders;
        }));
    semSheet.addRow([]);
    semSheet.addRow([]);
    semSheet.addRow(["SUB CODE", "SUB NAME", "NO OF STUDENT ATTENDED", "NO OF STUDENTS PASSED", "NO OF STUDENTS FAILED", "PASS %"]).eachCell(cell=>{
          cell.style = headerStyle;
        });

    sortedCourses.forEach(c => {
      const studentsWithMarks = students.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData && courseData.SEMESTER !== null && courseData.SEMESTER !== undefined;
      });
      
      const attended = studentsWithMarks.length;
      const passed = studentsWithMarks.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData.SEMESTER >= passingGP;
      }).length;
      const failed = attended - passed;
      const passPercent = attended > 0 ? parseFloat(((passed / attended) * 100).toFixed(2)) : 0;
      
      semSheet.addRow([c.code, c.name, attended, passed, failed, passPercent]).eachCell(cell=>{
          cell.style = thinBorders;
        });
    });

    // ANALYSIS SHEET
    const analysisSheet = workbook.addWorksheet("ANALYSIS");
    
    analysisSheet.addRow(["GOVERNMENT COLLEGE OF TECHNOLOGY,CBE-13"]);
    analysisSheet.addRow(["DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"]);
    analysisSheet.addRow(["PASS PERCENTAGE ANALYSIS"]);
    
    analysisSheet.mergeCells('A1:F1');
    analysisSheet.mergeCells('A2:F2');
    analysisSheet.mergeCells('A3:F3');
    
    analysisSheet.getCell('A1').style = top3style;
    analysisSheet.getCell('A2').style = top3style;
    analysisSheet.getCell('A3').style = top3style;
    
    const analysisHeaderRow = analysisSheet.addRow([
      "SUB CODE", "SUB NAME", "UNIT TEST 1 PASS %", "UNIT TEST 2 PASS %", "END SEMESTER PASS %"
    ]);
    
    analysisHeaderRow.eachCell(cell => {
      cell.style = headerStyle;
    });
    
    analysisSheet.getColumn(1).width = 15;
    analysisSheet.getColumn(2).width = 40;
    analysisSheet.getColumn(3).width = 20;
    analysisSheet.getColumn(4).width = 20;
    analysisSheet.getColumn(5).width = 20;
    
    // Calculate pass percentages for all courses
    sortedCourses.forEach(c => {
      let ut1PassPercent, ut2PassPercent;
      
      // For practical courses, show "-" for UT pass percentages
      if (c.type === "practical") {
        ut1PassPercent = "-";
        ut2PassPercent = "-";
      } else {
        // UT1 Pass %
        const ut1Students = students.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData && courseData.UT1 !== null && courseData.UT1 !== undefined;
        });
        const ut1Passed = ut1Students.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData.UT1 >= passingMarks;
        }).length;
        ut1PassPercent = ut1Students.length > 0 ? parseFloat(((ut1Passed / ut1Students.length) * 100).toFixed(2)) : 0;
        
        // UT2 Pass %
        const ut2Students = students.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData && courseData.UT2 !== null && courseData.UT2 !== undefined;
        });
        const ut2Passed = ut2Students.filter(stu => {
          const courseData = stu.courses.get(c.code);
          return courseData.UT2 >= passingMarks;
        }).length;
        ut2PassPercent = ut2Students.length > 0 ? parseFloat(((ut2Passed / ut2Students.length) * 100).toFixed(2)) : 0;
      }
      
      // Semester Pass % (for all course types)
      const semStudents = students.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData && courseData.SEMESTER !== null && courseData.SEMESTER !== undefined;
      });
      const semPassed = semStudents.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData.SEMESTER >= passingGP;
      }).length;
      const semPassPercent = semStudents.length > 0 ? parseFloat(((semPassed / semStudents.length) * 100).toFixed(2)) : 0;
      
      analysisSheet.addRow([c.code, c.name, ut1PassPercent, ut2PassPercent, semPassPercent]).eachCell(cell => {
        cell.style = thinBorders;
      });
    });

    // ARREAR COUNT SHEET
    const arrearSheet = workbook.addWorksheet("ARREAR COUNT");
    
    arrearSheet.addRow(["GOVERNMENT COLLEGE OF TECHNOLOGY,CBE-13"]);
    arrearSheet.addRow(["DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"]);
    arrearSheet.addRow(["ARREAR COUNT ANALYSIS"]);
    
    arrearSheet.mergeCells('A1:C1');
    arrearSheet.mergeCells('A2:C2');
    arrearSheet.mergeCells('A3:C3');
    
    arrearSheet.getCell('A1').style = top3style;
    arrearSheet.getCell('A2').style = top3style;
    arrearSheet.getCell('A3').style = top3style;
    
    const arrearHeaderRow = arrearSheet.addRow([
      "SUB CODE", "SUB NAME", "ARREAR COUNT"
    ]);
    
    arrearHeaderRow.eachCell(cell => {
      cell.style = headerStyle;
    });
    
    arrearSheet.getColumn(1).width = 15;
    arrearSheet.getColumn(2).width = 40;
    arrearSheet.getColumn(3).width = 20;
    
    // Calculate arrear count for all courses (students who failed in semester exam)
    sortedCourses.forEach(c => {
      const studentsWithMarks = students.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData && courseData.SEMESTER !== null && courseData.SEMESTER !== undefined;
      });
      
      const arrearCount = studentsWithMarks.filter(stu => {
        const courseData = stu.courses.get(c.code);
        return courseData.SEMESTER < passingGP;
      }).length;
      
      arrearSheet.addRow([c.code, c.name, arrearCount]).eachCell(cell => {
        cell.style = thinBorders;
      });
    });

    // CURRENT ARREARS SHEET
    const currentArrearsSheet = workbook.addWorksheet("CURRENT ARREARS");
    
    currentArrearsSheet.addRow(["GOVERNMENT COLLEGE OF TECHNOLOGY,CBE-13"]);
    currentArrearsSheet.addRow(["DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING"]);
    currentArrearsSheet.addRow(["CURRENT ARREARS LIST"]);
    
    currentArrearsSheet.mergeCells('A1:D1');
    currentArrearsSheet.mergeCells('A2:D2');
    currentArrearsSheet.mergeCells('A3:D3');
    
    currentArrearsSheet.getCell('A1').style = top3style;
    currentArrearsSheet.getCell('A2').style = top3style;
    currentArrearsSheet.getCell('A3').style = top3style;
    
    const currentArrearsHeaderRow = currentArrearsSheet.addRow([
      "S.NO", "REGISTER NO", "NAME", "SUBJECT NAME"
    ]);
    
    currentArrearsHeaderRow.eachCell(cell => {
      cell.style = headerStyle;
    });
    
    currentArrearsSheet.getColumn(1).width = 10;
    currentArrearsSheet.getColumn(2).width = 25;
    currentArrearsSheet.getColumn(3).width = 30;
    currentArrearsSheet.getColumn(4).width = 50;
    
    // Find all students with arrears and list their failed subjects
    let serialNo = 1;
    students.forEach(stu => {
      const failedSubjects = [];
      
      // Check each course for failures (grade point < 5)
      sortedCourses.forEach(c => {
        const courseData = stu.courses.get(c.code);
        if (courseData && courseData.SEMESTER !== null && courseData.SEMESTER !== undefined) {
          if (courseData.SEMESTER < passingGP) {
            failedSubjects.push(c.name);
          }
        }
      });
      
      // If student has any failed subjects, add them to the sheet
      if (failedSubjects.length > 0) {
        failedSubjects.forEach((subjectName, index) => {
          if (index === 0) {
            // First row with S.NO, Register No, and Name
            currentArrearsSheet.addRow([serialNo, stu.regNo, stu.name, subjectName]).eachCell(cell => {
              cell.style = thinBorders;
            });
            serialNo++;
          } else {
            // Subsequent rows with only subject name
            currentArrearsSheet.addRow(['', '', '', subjectName]).eachCell(cell => {
              cell.style = thinBorders;
            });
          }
        });
      }
    });


    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=results.xlsx");
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error("Error exporting Excel:", err);
    res.status(500).json({ error: `Failed to export Excel: ${err.message}` });
  }
});

app.get("/students", async (req, res) => {
  try {
    const students = await Student.find();
    res.json({
      count: students.length,
      students: students.map(s => ({
        regNo: s.regNo, name: s.name,
        coursesCount: s.courses ? s.courses.size : 0,
        courses: s.courses ? Object.fromEntries(s.courses) : {},
        failedCourses: s.failedCourses,
        GPA: parseFloat(s.GPA.toFixed(2))
      }))
    });
  } catch (err) {
    console.error("Error fetching students:", err);
    res.status(500).json({ error: "Failed to fetch students" });
  }
});

app.get("/students/:regNo", async (req, res) => {
  try {
    const student = await Student.findOne({ regNo: req.params.regNo });
    if (!student) return res.status(404).json({ error: "Student not found" });
    
    // Convert Map to object for proper JSON serialization
    const coursesObject = {};
    if (student.courses) {
      for (const [code, marks] of student.courses.entries()) {
        coursesObject[code] = {
          UT1: marks.UT1 !== undefined ? marks.UT1 : null,
          UT2: marks.UT2 !== undefined ? marks.UT2 : null,
          SEMESTER: marks.SEMESTER !== undefined ? marks.SEMESTER : null
        };
      }
    }
    
    res.json({
      regNo: student.regNo, name: student.name,
      courses: coursesObject,
      failedCourses: student.failedCourses,
      GPA: parseFloat(student.GPA.toFixed(2))
    });
  } catch (err) {
    console.error("Error fetching student:", err);
    res.status(500).json({ error: "Failed to fetch student" });
  }
});

app.delete("/students", async (req, res) => {
  try {
    // Delete all students from database
    const result = await Student.deleteMany({});
    
    // Delete all uploaded files
    const uploadsDir = "uploads";
    if (fs.existsSync(uploadsDir)) {
      const files = fs.readdirSync(uploadsDir);
      for (const file of files) {
        fs.unlinkSync(path.join(uploadsDir, file));
      }
    }
    
    res.json({ message: "All students deleted", count: result.deletedCount });
  } catch (err) {
    res.status(500).json({ error: "Failed to delete students" });
  }
});

// Update the courses delete route to handle both single and all courses
app.delete("/courses/:id", async (req, res) => {
  try {
    const courseId = req.params.id;
    const result = await Course.findByIdAndDelete(courseId);
    
    if (!result) {
      return res.status(404).json({ error: "Course not found" });
    }
    
    res.json({ message: "Course deleted successfully", course: result });
  } catch (err) {
    console.error("Error deleting course:", err);
    res.status(500).json({ error: "Failed to delete course" });
  }
});

app.delete("/courses", async (req, res) => {
  try {
    const result = await Course.deleteMany({});
    res.json({ message: "All courses deleted", count: result.deletedCount });
  } catch (err) {
    res.status(500).json({ error: "Failed to delete courses" });
  }
});

app.get("/stats", async (req, res) => {
  try {
    const studentCount = await Student.countDocuments();
    const courseCount = await Course.countDocuments();
    const students = await Student.find().limit(5);
    const courses = await Course.find();
    
    res.json({
      studentCount, courseCount,
      courses: courses.map(c => ({ 
        code: c.code, 
        name: c.name, 
        type: c.type, 
        credits: c.credits, 
        faculty: c.faculty,
        failedStudents: c.failedStudents || { UT1: 0, UT2: 0, SEMESTER: 0 }
      })),
      sampleStudents: students.map(s => ({
        regNo: s.regNo, name: s.name,
        coursesCount: s.courses ? s.courses.size : 0,
        failedCourses: s.failedCourses,
        GPA: parseFloat(s.GPA.toFixed(2))
      }))
    });
  } catch (err) {
    res.status(500).json({ error: "Failed to fetch stats" });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Backend running on http://localhost:${PORT}`));