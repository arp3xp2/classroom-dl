/**
 * Simple Google Classroom Assignment Downloader
 * 
 * This script downloads student assignments from Google Classroom,
 * organized by topic and student.
 */

/**
 * Shows a simple interface to select course and download assignments
 */
function doGet() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Google Classroom Assignment Downloader</h2>
    
    <div>
      <h3>Step 1: Select Course</h3>
      <select id="courseSelect" style="width: 100%; margin: 10px 0;" onchange="loadTopics()">
        <option value="">Loading courses...</option>
      </select>
    </div>

    <div id="topicsContainer" style="display:none; margin-top: 20px;">
      <h3>Step 2: Select Topics</h3>
      <div id="topicsList"></div>
      <label style="margin-top: 10px; display: block;">
        <input type="checkbox" id="selectAll" onclick="toggleAllTopics()"> Select All Topics
      </label>
    </div>

    <div style="margin-top: 20px; display: flex; gap: 10px;">
      <button id="downloadBtn" onclick="downloadSelected()" style="padding: 10px; flex: 1;" disabled>
        Download Selected Assignments
      </button>
      <button id="cancelBtn" onclick="cancelDownload()" style="padding: 10px; flex: 1; display: none; background-color: #f44336; color: white;">
        Cancel Download
      </button>
    </div>

    <div id="status" style="margin-top: 20px; color: #666;"></div>
    <div id="progress" style="margin-top: 10px; display: none;">
      <div style="width: 100%; background-color: #e0e0e0; border-radius: 4px;">
        <div id="progressBar" style="height: 20px; width: 0%; background-color: #4CAF50; border-radius: 4px; transition: width 0.3s;"></div>
      </div>
    </div>

    <script>
      let downloadInProgress = false;
      
      // Load courses on page load
      google.script.run
        .withSuccessHandler(showCourses)
        .withFailureHandler(showError)
        .listCourses();

      function showCourses(courses) {
        const select = document.getElementById('courseSelect');
        select.innerHTML = '<option value="">Select a course...</option>';
        courses.forEach(course => {
          select.innerHTML += \`<option value="\${course.id}">\${course.name}</option>\`;
        });
      }

      function loadTopics() {
        const courseId = document.getElementById('courseSelect').value;
        if (!courseId) return;

        document.getElementById('status').innerHTML = 'Loading topics...';
        document.getElementById('topicsContainer').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(showTopics)
          .withFailureHandler(showError)
          .getTopics(courseId);
      }

      function showTopics(topics) {
        const div = document.getElementById('topicsList');
        if (topics.length === 0) {
          div.innerHTML = '<p>No topics found in this course</p>';
          document.getElementById('downloadBtn').disabled = false;
          return;
        }

        div.innerHTML = topics.map(topic => \`
          <div style="margin: 5px 0;">
            <label>
              <input type="checkbox" name="topic" value="\${topic.id}">
              \${topic.name}
            </label>
          </div>
        \`).join('');

        document.getElementById('downloadBtn').disabled = false;
        document.getElementById('status').innerHTML = '';
      }

      function toggleAllTopics() {
        const checked = document.getElementById('selectAll').checked;
        document.querySelectorAll('input[name="topic"]')
          .forEach(box => box.checked = checked);
      }

      function downloadSelected() {
        const courseId = document.getElementById('courseSelect').value;
        const topicCheckboxes = document.querySelectorAll('input[name="topic"]:checked');
        const topicIds = Array.from(topicCheckboxes).map(cb => cb.value);

        document.getElementById('status').innerHTML = 'Downloading assignments...';
        document.getElementById('downloadBtn').disabled = true;
        document.getElementById('cancelBtn').style.display = 'block';
        document.getElementById('progress').style.display = 'block';
        downloadInProgress = true;
        
        google.script.run
          .withSuccessHandler(showSuccess)
          .withFailureHandler(showError)
          .downloadAssignments(courseId, topicIds);
      }
      
      function cancelDownload() {
        if (!downloadInProgress) return;
        
        document.getElementById('status').innerHTML = 'Cancelling download...';
        
        google.script.run
          .withSuccessHandler(function() {
            downloadInProgress = false;
            document.getElementById('status').innerHTML = 'Download cancelled';
            document.getElementById('downloadBtn').disabled = false;
            document.getElementById('cancelBtn').style.display = 'none';
            document.getElementById('progress').style.display = 'none';
          })
          .withFailureHandler(showError)
          .cancelDownload();
      }

      function showSuccess(message) {
        downloadInProgress = false;
        document.getElementById('status').innerHTML = message;
        document.getElementById('downloadBtn').disabled = false;
        document.getElementById('cancelBtn').style.display = 'none';
        document.getElementById('progressBar').style.width = '100%';
        
        // Reset progress bar after 2 seconds
        setTimeout(() => {
          document.getElementById('progress').style.display = 'none';
          document.getElementById('progressBar').style.width = '0%';
        }, 2000);
      }

      function showError(error) {
        downloadInProgress = false;
        document.getElementById('status').innerHTML = 'Error: ' + error;
        document.getElementById('downloadBtn').disabled = false;
        document.getElementById('cancelBtn').style.display = 'none';
        document.getElementById('progress').style.display = 'none';
      }
    </script>
  `);
  
  return html.setTitle('Classroom Assignment Downloader');
}

/**
 * Lists all available courses
 */
function listCourses() {
  try {
    const response = Classroom.Courses.list();
    if (!response || !response.courses) return [];
    
    return response.courses.map(c => ({
      id: c.id,
      name: c.name
    }));
  } catch (error) {
    Logger.log("Error listing courses: " + error);
    throw new Error("Failed to load courses. Please check permissions.");
  }
}

/**
 * Gets topics for a course
 */
function getTopics(courseId) {
  try {
    // Try direct topics API first
    try {
      const response = Classroom.Courses.Topics.list(courseId);
      if (response && response.topic) {
        return response.topic.map(t => ({
          id: t.topicId,
          name: t.name
        }));
      }
    } catch (e) {
      Logger.log("Topics API failed, falling back to coursework");
    }

    // Fallback: Get topics from coursework
    const courseWork = getAllCourseWork(courseId);
    const topicMap = new Map();
    
    courseWork.forEach(work => {
      if (work.topicId && !topicMap.has(work.topicId)) {
        topicMap.set(work.topicId, {
          id: work.topicId,
          name: work.topic || 'Topic ' + work.topicId
        });
      }
    });

    return Array.from(topicMap.values());
  } catch (error) {
    Logger.log("Error getting topics: " + error);
    throw new Error("Failed to load topics");
  }
}

// Add this at the top of your script
let shouldCancelDownload = false;

/**
 * Cancels the current download
 */
function cancelDownload() {
  shouldCancelDownload = true;
  return "Cancellation requested";
}

/**
 * Main function to download assignments
 */
function downloadAssignments(courseId, topicIds) {
  try {
    // Reset cancel flag
    shouldCancelDownload = false;
    
    // Verify course exists
    const courseDetails = Classroom.Courses.get(courseId);
    const courseName = courseDetails.name;
    
    // Create main folder
    const rootFolder = DriveApp.createFolder('Classroom Downloads - ' + courseName);
    
    // Get students and course work
    const students = getAllStudents(courseId);
    const courseWork = getAllCourseWork(courseId);
    
    // Create topic folders
    const topicFolders = {};
    
    // If no topics selected, download all
    if (!topicIds || topicIds.length === 0) {
      topicFolders['no-topic'] = rootFolder.createFolder('All Assignments');
    } else {
      // Get topic names
      const allTopics = getTopics(courseId);
      const topicMap = new Map(allTopics.map(t => [t.id, t.name]));
      
      // Create folders for selected topics
      topicIds.forEach(id => {
        const name = topicMap.get(id) || `Topic ${id}`;
        topicFolders[id] = rootFolder.createFolder(name);
      });
      
      // Add uncategorized folder
      topicFolders['no-topic'] = rootFolder.createFolder('Uncategorized');
    }
    
    // Create a map to track student folders by topic
    const studentFoldersByTopic = {};
    
    // Process each assignment
    let downloadCount = 0;
    
    for (const assignment of courseWork) {
      // Check if download was cancelled
      if (shouldCancelDownload) {
        return `Download cancelled. ${downloadCount} files were downloaded to "${rootFolder.getName()}" before cancellation. Open folder: ${rootFolder.getUrl()}`;
      }
      
      const topicId = assignment.topicId || 'no-topic';
      
      // Skip if not in selected topics
      if (topicIds && topicIds.length > 0 && !topicIds.includes(topicId) && topicId !== 'no-topic') {
        continue;
      }
      
      const topicFolder = topicFolders[topicId] || topicFolders['no-topic'];
      
      // Initialize student folders map for this topic if needed
      if (!studentFoldersByTopic[topicId]) {
        studentFoldersByTopic[topicId] = {};
      }
      
      const submissions = getAllSubmissions(courseId, assignment.id);
      
      // Process each submission
      for (const submission of submissions) {
        // Check if download was cancelled
        if (shouldCancelDownload) {
          return `Download cancelled. ${downloadCount} files were downloaded to "${rootFolder.getName()}" before cancellation. Open folder: ${rootFolder.getUrl()}`;
        }
        
        const studentId = submission.userId;
        const student = students.find(s => s.userId === studentId);
        
        if (!student || !submission.assignmentSubmission || !submission.assignmentSubmission.attachments) {
          continue;
        }
        
        // Get clean student name
        const studentName = student.profile.name || student.profile.emailAddress || studentId.toString();
        // Remove fullName, givenName, familyName prefixes and other non-alphanumeric characters
        const cleanStudentName = String(studentName)
          .replace(/fullName|givenName|familyName/g, '')
          .replace(/[^\w\s-]/g, '')
          .trim();
        
        // Use existing student folder or create a new one for this topic
        let studentFolder;
        if (!studentFoldersByTopic[topicId][studentId]) {
          studentFolder = topicFolder.createFolder(cleanStudentName);
          studentFoldersByTopic[topicId][studentId] = studentFolder;
        } else {
          studentFolder = studentFoldersByTopic[topicId][studentId];
        }
        
        // Process attachments
        const attachments = submission.assignmentSubmission.attachments;
        for (let i = 0; i < attachments.length; i++) {
          // Check if download was cancelled
          if (shouldCancelDownload) {
            return `Download cancelled. ${downloadCount} files were downloaded to "${rootFolder.getName()}" before cancellation. Open folder: ${rootFolder.getUrl()}`;
          }
          
          const attachment = attachments[i];
          const safeAssignmentTitle = String(assignment.title)
            .replace(/[^\w\s-]/g, '')
            .replace(/\s+/g, '-');
          const baseFilename = safeAssignmentTitle;
          
          if (attachment.driveFile) {
            downloadFile(attachment.driveFile, studentFolder, baseFilename);
            downloadCount++;
          } else if (attachment.link) {
            createLinkFile(attachment.link, studentFolder, baseFilename);
            downloadCount++;
          }
        }
      }
    }
    
    return `Downloaded ${downloadCount} files to "${rootFolder.getName()}". Open folder: ${rootFolder.getUrl()}`;
  } catch (error) {
    Logger.log("Error: " + error);
    throw new Error("Download failed: " + error.message);
  }
}

/**
 * Helper functions
 */
function getAllStudents(courseId) {
  const students = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.Students.list(courseId, { pageToken: pageToken });
    if (response.students) students.push(...response.students);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return students;
}

function getAllCourseWork(courseId) {
  const courseWork = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.CourseWork.list(courseId, { pageToken: pageToken });
    if (response.courseWork) courseWork.push(...response.courseWork);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return courseWork;
}

function getAllSubmissions(courseId, courseWorkId) {
  const submissions = [];
  let pageToken = null;
  
  do {
    const response = Classroom.Courses.CourseWork.StudentSubmissions.list(
      courseId, courseWorkId, { pageToken: pageToken }
    );
    if (response.studentSubmissions) submissions.push(...response.studentSubmissions);
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return submissions;
}

function downloadFile(driveFile, folder, baseFilename) {
  try {
    const file = DriveApp.getFileById(driveFile.id);
    const fileName = file.getName();
    const ext = fileName.split('.').pop();
    const newFilename = baseFilename + '.' + ext;
    
    file.makeCopy(newFilename, folder);
  } catch (error) {
    Logger.log('Error downloading file: ' + error);
  }
}

function createLinkFile(link, folder, baseFilename) {
  const content = 'URL: ' + link.url;
  const newFilename = baseFilename + '_link.txt';
  folder.createFile(newFilename, content);
}