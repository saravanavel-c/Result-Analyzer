import { useState, useEffect } from 'react';
import './App.css';

function App() {
  const [courses, setCourses] = useState([]);
  const [showAddCourseDialog, setShowAddCourseDialog] = useState(false);
  const [newCourse, setNewCourse] = useState({
    code: '',
    name: '',
    faculty: '',
    type: 'theory',
    credits: 3
  });
  const [uploadedFileName, setUploadedFileName] = useState('');
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [activeView, setActiveView] = useState('courses'); // 'courses' or 'students'

  // Fetch courses from backend
  useEffect(() => {
    fetchCourses();
  }, []);

  const fetchCourses = async () => {
    try {
      const response = await fetch('http://localhost:5000/courses');
      const data = await response.json();
      setCourses(data);
    } catch (error) {
      console.error('Error fetching courses:', error);
    }
  };

  const handleAddCourse = async (e) => {
    e.preventDefault();
    try {
      const response = await fetch('http://localhost:5000/courses', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(newCourse),
      });
      
      const result = await response.json();
      if (response.ok) {
        fetchCourses(); // Refresh the course list
        setShowAddCourseDialog(false);
        setNewCourse({
          code: '',
          name: '',
          faculty: '',
          type: 'theory',
          credits: 3
        });
      } else {
        alert(result.error || 'Failed to add course');
      }
    } catch (error) {
      console.error('Error adding course:', error);
      alert('Failed to add course');
    }
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setNewCourse({
      ...newCourse,
      [name]: name === 'credits' ? parseFloat(value) || 0 : value
    });
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    // Set the uploaded filename
    setUploadedFileName(file.name);

    const formData = new FormData();
    formData.append('file', file);

    fetch('http://localhost:5000/upload', {
      method: 'POST',
      body: formData,
    })
    .then(response => response.json())
    .then(data => {
      alert(data.message);
      // Refresh courses to update fail counts
      fetchCourses();
    })
    .catch(error => {
      console.error('Error uploading file:', error);
      alert('Failed to upload file');
    });
  };

  const handleDownload = () => {
    window.location.href = 'http://localhost:5000/export';
  };

  const handleDeleteStudents = async () => {
    try {
      const response = await fetch('http://localhost:5000/students', {
        method: 'DELETE',
      });
      
      const result = await response.json();
      if (response.ok) {
        alert(`Deleted ${result.count} students`);
        // Refresh courses to update fail counts
        fetchCourses();
      } else {
        alert(result.error || 'Failed to delete students');
      }
    } catch (error) {
      console.error('Error deleting students:', error);
      alert('Failed to delete students');
    }
    setShowDeleteConfirm(false);
  };

  const handleDeleteCourse = async (courseId) => {
    if (!window.confirm('Are you sure you want to delete this course?')) {
      return;
    }
    
    try {
      const response = await fetch(`http://localhost:5000/courses/${courseId}`, {
        method: 'DELETE',
      });
      
      if (response.ok) {
        const result = await response.json();
        alert('Course deleted successfully');
        fetchCourses();
      } else {
        const errorResult = await response.json();
        alert(errorResult.error || 'Failed to delete course');
      }
    } catch (error) {
      console.error('Error deleting course:', error);
      alert('Failed to delete course: ' + error.message);
    }
  };

  return (
    <div className="app">
      {/* Navigation */}
      <div className="navigation">
        <button 
          className={activeView === 'courses' ? 'active' : ''}
          onClick={() => setActiveView('courses')}
        >
          Courses
        </button>
        <button 
          className={activeView === 'students' ? 'active' : ''}
          onClick={() => setActiveView('students')}
        >
          Students
        </button>
      </div>

      {activeView === 'courses' ? (
        <>
          {/* Upper Section - File Upload */}
          <div className="upload-section">
            <h2>Upload Excel File</h2>
            <div className="upload-area">
              <input 
                type="file" 
                accept=".xlsx,.xls" 
                onChange={handleFileUpload} 
                id="file-upload"
              />
              <label htmlFor="file-upload" className="upload-label">
                Choose Excel File
              </label>
              {uploadedFileName && (
                <p className="uploaded-file">Uploaded: {uploadedFileName}</p>
              )}
              <p>Drag & drop files here or click to browse</p>
            </div>
            
            <div className="upload-actions">
              <button className="download-btn" onClick={handleDownload}>
                Download Result Analysis Sheet
              </button>
              <button 
                className="delete-btn" 
                onClick={() => setShowDeleteConfirm(true)}
              >
                Delete Student Data
              </button>
            </div>
          </div>

          {/* Lower Section - Course Details */}
          <div className="course-section">
            <div className="course-header">
              <h2>Available Courses</h2>
              <button 
                className="add-course-btn" 
                onClick={() => setShowAddCourseDialog(true)}
              >
                Add Course
              </button>
            </div>

            <div className="course-list">
              {courses.length > 0 ? (
                courses.map((course) => (
                  <div key={course._id} className="course-card">
                    <div className="course-header-row">
                      <h3>{course.name}</h3>
                      <button 
                        className="delete-course-btn"
                        onClick={() => handleDeleteCourse(course._id)}
                      >
                        ×
                      </button>
                    </div>
                    <p><strong>Code:</strong> {course.code}</p>
                    <p><strong>Faculty:</strong> {course.faculty}</p>
                    <p><strong>Type:</strong> {course.type}</p>
                    <p><strong>Credits:</strong> {course.credits}</p>
                    <div className="failed-students">
                      <p><strong>Failed Students:</strong></p>
                      <ul>
                        <li>SEMESTER: {course.failedStudents?.SEMESTER || 0}</li>
                        <li>UT1: {course.failedStudents?.UT1 || 0}</li>
                        <li>UT2: {course.failedStudents?.UT2 || 0}</li>
                      </ul>
                    </div>
                  </div>
                ))
              ) : (
                <p>No courses available. Add a course to get started.</p>
              )}
            </div>
          </div>
        </>
      ) : (
        <StudentList />
      )}

      {/* Add Course Dialog */}
      {showAddCourseDialog && (
        <div className="dialog-overlay">
          <div className="dialog">
            <h2>Add New Course</h2>
            <form onSubmit={handleAddCourse}>
              <div className="form-group">
                <label>Course Code:</label>
                <input
                  type="text"
                  name="code"
                  value={newCourse.code}
                  onChange={handleInputChange}
                  required
                />
              </div>
              
              <div className="form-group">
                <label>Course Name:</label>
                <input
                  type="text"
                  name="name"
                  value={newCourse.name}
                  onChange={handleInputChange}
                  required
                />
              </div>
              
              <div className="form-group">
                <label>Faculty:</label>
                <input
                  type="text"
                  name="faculty"
                  value={newCourse.faculty}
                  onChange={handleInputChange}
                  required
                />
              </div>
              
              <div className="form-group">
                <label>Type:</label>
                <select
                  name="type"
                  value={newCourse.type}
                  onChange={handleInputChange}
                >
                  <option value="theory">Theory</option>
                  <option value="practical">Practical</option>
                  <option value="extra">Extra</option>
                </select>
              </div>
              
              <div className="form-group">
                <label>Credits:</label>
                <input
                  type="number"
                  name="credits"
                  min="0"
                  step="0.1"
                  value={newCourse.credits}
                  onChange={handleInputChange}
                  required
                />
              </div>
              
              <div className="dialog-actions">
                <button type="button" onClick={() => setShowAddCourseDialog(false)}>
                  Cancel
                </button>
                <button type="submit">Add Course</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Confirmation Dialog */}
      {showDeleteConfirm && (
        <div className="dialog-overlay">
          <div className="dialog">
            <h2>Confirm Deletion</h2>
            <p>Are you sure you want to delete all student data? This action cannot be undone.</p>
            <div className="dialog-actions">
              <button onClick={() => setShowDeleteConfirm(false)}>
                Cancel
              </button>
              <button className="danger" onClick={handleDeleteStudents}>
                Delete All Student Data
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StudentList() {
  const [students, setStudents] = useState([]);
  const [expandedStudent, setExpandedStudent] = useState(null);

  useEffect(() => {
    fetchStudents();
  }, []);

  const fetchStudents = async () => {
    try {
      const response = await fetch('http://localhost:5000/students');
      const data = await response.json();
      setStudents(data.students || []);
    } catch (error) {
      console.error('Error fetching students:', error);
    }
  };

  const fetchStudentDetails = async (regNo) => {
    try {
      const response = await fetch(`http://localhost:5000/students/${regNo}`);
      const data = await response.json();
      return data;
    } catch (error) {
      console.error('Error fetching student details:', error);
      return null;
    }
  };

  const handleStudentClick = async (regNo) => {
    if (expandedStudent === regNo) {
      setExpandedStudent(null);
      return;
    }
    
    const details = await fetchStudentDetails(regNo);
    if (details) {
      setExpandedStudent(regNo);
    }
  };

  return (
    <div className="student-section">
      <h2>Student List</h2>
      <div className="student-list-vertical">
        {students.length > 0 ? (
          students.map((student) => (
            <div key={student.regNo} className="student-list-item">
              <div 
                className="student-header"
                onClick={() => handleStudentClick(student.regNo)}
              >
                <div className="student-basic-info">
                  <h3>{student.name}</h3>
                  <p><strong>Reg No:</strong> {student.regNo}</p>
                </div>
                <div className="expand-indicator">
                  {expandedStudent === student.regNo ? '▲' : '▼'}
                </div>
              </div>
              
              {expandedStudent === student.regNo && (
                <div className="student-details">
                  <div className="student-stats">
                    <p><strong>GPA:</strong> {student.GPA}</p>
                    <p><strong>Courses:</strong> {student.coursesCount}</p>
                  </div>
                  <h4>Course Marks</h4>
                  {student.courses && Object.keys(student.courses).length > 0 ? (
                    <table>
                      <thead>
                        <tr>
                          <th>Course Code</th>
                          <th>UT1</th>
                          <th>UT2</th>
                          <th>Semester</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(student.courses).map(([code, marks]) => (
                          <tr key={code}>
                            <td>{code}</td>
                            <td>{marks.UT1 !== null ? marks.UT1 : '-'}</td>
                            <td>{marks.UT2 !== null ? marks.UT2 : '-'}</td>
                            <td>{marks.SEMESTER !== null ? marks.SEMESTER : '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <p>No course data available</p>
                  )}
                  
                  <h4>Failed Courses</h4>
                  <ul>
                    <li>UT1: {student.failedCourses?.UT1 || 0}</li>
                    <li>UT2: {student.failedCourses?.UT2 || 0}</li>
                    <li>SEMESTER: {student.failedCourses?.SEMESTER || 0}</li>
                  </ul>
                </div>
              )}
            </div>
          ))
        ) : (
          <p>No students available.</p>
        )}
      </div>
    </div>
  );
}

export default App;